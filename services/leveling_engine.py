"""
services/leveling_engine.py
Motor central de nivelación de appointments.
Entrada: lista de dicts (de Metabase o Excel).
Salida: dict JSON con resumen, alertas, sugerencias, carga por técnico y franja.
Sin dependencias de openpyxl ni generación de archivos Excel.
"""
import logging
from datetime import datetime, timedelta
from config import (
    MAX_IDEAL_LOAD, MIN_IDEAL_LOAD, MAX_ABSOLUTE_LOAD, MAX_ORDERS_PER_SLOT,
    MAX_DUPLICATED_SLOTS, MIN_IMBALANCE_TO_MOVE, ORDER_DURATION_HOURS,
    MAX_ORDER_DURATION_HOURS, MAX_ALLOWED_DISTANCE_KM, FRAGMENTATION_PENALTY,
    INTERZONE_DISTANCE_PENALTY, ZONE_ONLY_NO_COORDS_PENALTY,
    MAX_INTERZONE_ASSIGNMENTS_PER_TECH, ZONE_ADJACENCY, MOVABLE_STATUSES,
    ONSITE_ALERT_MINUTES, INICIADO_ALERT_MINUTES, OVERLOAD_PER_SLOT, EFFICIENT_TECH_PROTECTION_SCORE,
    ACTIVE_SLOT_NO_PROGRESS_MINUTES, SLOT_RISK_MINUTES_BEFORE_END, now_bogota,
)
from services.normalization import (
    normalize_order, haversine, get_centroid, is_same_unit, order_has_coords,
    parse_franja_hours, get_status_progress, status_effective_weight,
    status_completion_credit, norm_zone, is_movable, is_blocked,
    norm_status,
)

logger = logging.getLogger(__name__)

# ─── Utilidades internas ──────────────────────

def _count_duplicated_slots(franja_counts: dict) -> int:
    return sum(1 for c in franja_counts.values() if c >= 2)


def _estimate_remaining_hours(order: dict, current_hour: float) -> float:
    progress = order.get("progress", 0)
    base = ORDER_DURATION_HOURS
    if progress == 3:
        base = 0.6
    elif progress == 4:
        base = 0.3
    elif progress >= 5:
        base = 0.1
    onsite_hour = order.get("onsite_hour")
    if onsite_hour and current_hour:
        elapsed = current_hour - onsite_hour
        if elapsed > 0:
            return max(0.1, base - elapsed)
    return base


def _parse_updated_at(updated_at_str: str):
    """Intenta parsear updated_at a datetime naive (sin timezone). Devuelve None si falla."""
    if not updated_at_str:
        return None
    # Si es ya un datetime (de openpyxl), quitar timezone si tiene
    if hasattr(updated_at_str, 'replace'):
        try:
            return updated_at_str.replace(tzinfo=None)
        except Exception:
            pass
    s = str(updated_at_str).strip()[:19]
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except (ValueError, AttributeError):
            pass
    return None


def _tech_reference_point(tech: str, tech_orders: dict, tech_locs: dict):
    for o in tech_orders.get(tech, []):
        if 1 <= o.get("progress", 0) < 6 and o.get("lat") and o.get("lon"):
            return (o["lat"], o["lon"])
    upcoming = [
        (parse_franja_hours(o.get("franja", ""))[0] or 99, o)
        for o in tech_orders.get(tech, [])
        if o.get("movible") and o.get("lat") and o.get("lon")
    ]
    if upcoming:
        upcoming.sort(key=lambda x: x[0])
        od = upcoming[0][1]
        return (od["lat"], od["lon"])
    locs = tech_locs.get(tech, [])
    return get_centroid(locs) if locs else (0.0, 0.0)


def _dist_to_tech(order: dict, tech: str, tech_orders: dict, tech_locs: dict):
    if not order.get("lat") or not order.get("lon"):
        return None
    ref = _tech_reference_point(tech, tech_orders, tech_locs)
    if ref == (0.0, 0.0):
        return None
    return haversine(order["lat"], order["lon"], ref[0], ref[1])


def _tech_load_score(tech: str, tech_pending: dict, tech_total: dict,
                     tech_eff_load: dict, tech_credit: dict,
                     tech_franja: dict, tech_subzones: dict, tech_orders: dict) -> float:
    duplicated = _count_duplicated_slots(tech_franja.get(tech, {}))
    subzone_count = len(tech_subzones.get(tech, set()))
    fragmentation = (subzone_count * 1.2) + (duplicated * 1.8)
    return (
        tech_eff_load.get(tech, 0.0) * 1350
        + tech_pending.get(tech, 0) * 750
        + tech_total.get(tech, 0) * 150
        + duplicated * 650
        + max(0, subzone_count - 1) * 180
        + fragmentation * FRAGMENTATION_PENALTY
        - tech_credit.get(tech, 0.0) * 220
    )


def _is_tech_efficient(tech: str, tech_orders: dict, tech_franja: dict,
                        tech_subzones: dict, tech_credit: dict,
                        tech_pending: dict, tech_eff_load: dict) -> bool:
    if tech == "SIN_ASIGNAR":
        return False
    pending = tech_pending.get(tech, 0)
    if pending <= 0:
        return False
    completed = tech_credit.get(tech, 0.0)
    duplicated = _count_duplicated_slots(tech_franja.get(tech, {}))
    subzone_count = len(tech_subzones.get(tech, set()))
    active = sum(1 for o in tech_orders.get(tech, []) if 1 <= o.get("progress", 0) < 6)
    eff = tech_eff_load.get(tech, 0.0)
    return (
        completed >= EFFICIENT_TECH_PROTECTION_SCORE
        and duplicated == 0
        and subzone_count <= 2
        and active <= 1
        and eff <= max(2.2, pending + 0.5)
    )


def _can_add_to_franja(tech: str, franja: str, tech_franja: dict,
                        tech_orders: dict, current_hour: float,
                        same_unit: bool = False) -> tuple:
    franja_start, franja_end = parse_franja_hours(franja)
    current_in_slot = tech_franja.get(tech, {}).get(franja, 0)

    if current_in_slot >= MAX_ORDERS_PER_SLOT:
        return False, "Ya tiene 2 órdenes en esta franja"

    if current_in_slot >= 1:
        existing_dups = _count_duplicated_slots(tech_franja.get(tech, {}))
        if existing_dups >= MAX_DUPLICATED_SLOTS:
            if same_unit:
                return True, "OK (excepción misma unidad)"
            return False, "Ya usó su franja duplicada permitida"

    if franja_start is not None and franja_start >= 14.5:
        tarde_count = sum(
            c for f, c in tech_franja.get(tech, {}).items()
            if (parse_franja_hours(f)[0] or 0) >= 14.5
        )
        if tarde_count >= 2:
            return False, "Ya tiene 2 órdenes en franjas de tarde"

    return True, "OK"


# ─── Construcción de estructuras de datos ─────

def _build_indexes(orders: list) -> dict:
    """Construye todos los índices necesarios para el motor a partir de la lista normalizada."""
    tech_orders = {}          # tech -> [orders]
    tech_franja = {}          # tech -> {franja: count_total}
    tech_franja_active = {}   # tech -> {franja: count_no_finalizada_no_cancelada}
    tech_subzones = {}        # tech -> set(subzonas)
    tech_locs = {}            # tech -> [(lat, lon)]
    tech_zone = {}            # tech -> zona principal
    zone_techs = {}           # zona -> [techs]

    for o in orders:
        tech = o["tecnico"]
        franja = o["franja"]
        zona = o["zona"]
        subzona = o["subzona"]

        # tech_orders
        tech_orders.setdefault(tech, []).append(o)

        # tech_franja (conteo total, incluye finalizadas)
        tech_franja.setdefault(tech, {})
        tech_franja[tech][franja] = tech_franja[tech].get(franja, 0) + 1

        # tech_franja_active: solo órdenes que aún consumen capacidad
        # Excluye finalizadas (progress >= 6) y canceladas
        is_done = (o.get("progress", 0) >= 6) or ("cancel" in o.get("estado", "").lower())
        if not is_done:
            tech_franja_active.setdefault(tech, {})
            tech_franja_active[tech][franja] = tech_franja_active[tech].get(franja, 0) + 1

        # tech_subzones
        tech_subzones.setdefault(tech, set()).add(subzona)

        # tech_locs
        if o.get("lat") and o.get("lon"):
            tech_locs.setdefault(tech, []).append((o["lat"], o["lon"]))

        # tech_zone (zona con más órdenes gana)
        tech_zone.setdefault(tech, {})
        tech_zone[tech][zona] = tech_zone[tech].get(zona, 0) + 1

        # zone_techs
        zone_techs.setdefault(zona, set()).add(tech)

    # Resolver zona principal de cada técnico
    tech_main_zone = {
        t: max(zones.items(), key=lambda x: x[1])[0]
        for t, zones in tech_zone.items()
    }

    # Calcular cargas
    tech_total = {t: len(o_list) for t, o_list in tech_orders.items()}
    tech_pending = {
        t: sum(1 for o in o_list if o["movible"])
        for t, o_list in tech_orders.items()
    }
    tech_eff_load = {
        t: sum(o["effective_weight"] for o in o_list)
        for t, o_list in tech_orders.items()
    }
    tech_credit = {
        t: sum(o["completion_credit"] for o in o_list)
        for t, o_list in tech_orders.items()
    }

    return {
        "tech_orders":         tech_orders,
        "tech_franja":         tech_franja,
        "tech_franja_active":  tech_franja_active,   # sin finalizadas ni canceladas
        "tech_subzones":       tech_subzones,
        "tech_locs":           tech_locs,
        "tech_main_zone":      tech_main_zone,
        "zone_techs":          {z: list(ts) for z, ts in zone_techs.items()},
        "tech_total":          tech_total,
        "tech_pending":        tech_pending,
        "tech_eff_load":       tech_eff_load,
        "tech_credit":         tech_credit,
    }


# ─── Generación de alertas ────────────────────

def _generate_alerts(orders: list, idx: dict, now_dt) -> list:
    alerts = []
    now_hour = now_dt.hour + now_dt.minute / 60.0
    tech_orders = idx["tech_orders"]
    tech_franja = idx["tech_franja"]
    tech_pending = idx["tech_pending"]

    # 1. Alertas por estado prolongado: "En sitio" y "Iniciado/a"
    # ── "En sitio": técnico llegó pero NO inició la orden.
    #    Tiempo esperado: ≤ 30 min. Más de eso = bloqueo de acceso, cliente ausente,
    #    o técnico sin reportar. Umbral: ONSITE_ALERT_MINUTES (30 min default).
    # ── "Iniciado/a": orden en ejecución activa.
    #    Tiempo esperado: ≤ 90 min. Más de eso = problema técnico complejo o
    #    abandono sin reportar. Umbral: INICIADO_ALERT_MINUTES (90 min default).
    ALERTAS_PROLONGADO = {
        "en sitio":  (ONSITE_ALERT_MINUTES,   "EN_SITIO_PROLONGADO",   "en sitio (sin iniciar)"),
        "iniciado":  (INICIADO_ALERT_MINUTES,  "INICIADO_PROLONGADO",   "en ejecución (iniciado)"),
        "iniciada":  (INICIADO_ALERT_MINUTES,  "INICIADO_PROLONGADO",   "en ejecución (iniciada)"),
    }
    for o in orders:
        estado_norm = norm_status(o["estado"])
        if estado_norm not in ALERTAS_PROLONGADO:
            continue
        umbral_min, tipo_alerta, desc_estado = ALERTAS_PROLONGADO[estado_norm]
        updated = _parse_updated_at(o.get("updated_at", ""))
        if not updated:
            continue
        try:
            now_naive = now_dt.replace(tzinfo=None)
            minutes_elapsed = (now_naive - updated).total_seconds() / 60
            if minutes_elapsed > umbral_min:
                alerts.append({
                    "tipo":      tipo_alerta,
                    "severidad": "critica" if minutes_elapsed > umbral_min * 2 else "alta",
                    "orden":     o["id"],
                    "tecnico":   o["tecnico"],
                    "franja":    o["franja"],
                    "zona":      o["zona"],
                    "minutos":   int(minutes_elapsed),
                    "umbral":    umbral_min,
                    "estado":    estado_norm,
                    "detalle": (
                        f"Orden {o['id']} {desc_estado} hace {int(minutes_elapsed)} min "
                        f"(umbral: {umbral_min} min). "
                        f"Verificar bloqueo, novedad técnica o reporte en campo."
                    ),
                })
        except Exception:
            pass

    # 2. Técnico sobrecargado en una franja
    # Usa tech_franja_active: solo cuenta órdenes no finalizadas ni canceladas.
    # Esto evita falsas alertas cuando el técnico ya completó órdenes del slot.
    # Severidad:
    #   media  → 2 órdenes activas simultáneas (gestionable, seguimiento)
    #   alta   → 3+ órdenes activas reales simultáneas (riesgo de incumplimiento)
    tech_franja_active = idx.get("tech_franja_active", {})
    for tech, franja_map in tech_franja_active.items():
        if tech == "SIN_ASIGNAR":
            continue
        for franja, active_count in franja_map.items():
            if active_count < OVERLOAD_PER_SLOT:
                continue
            # 2 activas = media; 3+ activas = alta
            sev = "alta" if active_count >= 3 else "media"
            # Contar el total (incluye finalizadas) solo para contexto
            total_count = tech_franja.get(tech, {}).get(franja, active_count)
            completadas = total_count - active_count
            detalle = (
                f"{tech} tiene {active_count} orden(es) activa(s) en franja {franja}"
            )
            if completadas > 0:
                detalle += f" ({completadas} ya finalizada/cancelada)"
            alerts.append({
                "tipo":             "SOBRECARGA_FRANJA",
                "severidad":        sev,
                "tecnico":          tech,
                "franja":           franja,
                "count":            active_count,
                "count_total":      total_count,
                "count_completadas":completadas,
                "detalle":          detalle,
            })

    # 3a. Técnico con carga total > máximo ideal (sobrecargado)
    for tech, total in idx["tech_total"].items():
        if tech == "SIN_ASIGNAR":
            continue
        if total > MAX_ABSOLUTE_LOAD:
            alerts.append({
                "tipo":      "SOBRECARGA_TOTAL",
                "severidad": "alta",
                "tecnico":   tech,
                "total":     total,
                "detalle":   f"{tech} tiene {total} órdenes (máx recomendado: {MAX_IDEAL_LOAD})",
            })
        elif total > MAX_IDEAL_LOAD:
            alerts.append({
                "tipo":      "SOBRECARGA_TOTAL",
                "severidad": "media",
                "tecnico":   tech,
                "total":     total,
                "detalle":   f"{tech} tiene {total} órdenes (supera el techo de {MAX_IDEAL_LOAD})",
            })

    # 3b. Técnico con pocas órdenes (por debajo del mínimo ideal)
    for tech, total in idx["tech_total"].items():
        if tech == "SIN_ASIGNAR":
            continue
        pending = idx["tech_pending"].get(tech, 0)
        if total < MIN_IDEAL_LOAD and pending > 0:
            alerts.append({
                "tipo":      "CARGA_BAJA",
                "severidad": "media",
                "tecnico":   tech,
                "total":     total,
                "detalle":   f"{tech} tiene solo {total} órdenes (mínimo recomendado: {MIN_IDEAL_LOAD})",
            })

    # 4. Órdenes sin técnico o sin franja (por programar huérfanas)
    for o in orders:
        if o["movible"]:
            if o["tecnico"] == "SIN_ASIGNAR":
                alerts.append({
                    "tipo":      "SIN_TECNICO",
                    "severidad": "media",
                    "orden":     o["id"],
                    "franja":    o["franja"],
                    "zona":      o["zona"],
                    "detalle":   f"Orden {o['id']} sin técnico asignado",
                })
            elif o["franja"] == "Sin Franja":
                alerts.append({
                    "tipo":      "SIN_FRANJA",
                    "severidad": "media",
                    "orden":     o["id"],
                    "tecnico":   o["tecnico"],
                    "zona":      o["zona"],
                    "detalle":   f"Orden {o['id']} sin franja horaria ({o['tecnico']})",
                })

    # 5. Órdenes programadas en franja ya pasada
    for o in orders:
        if o["movible"] and o["franja"] != "Sin Franja":
            _, franja_end = parse_franja_hours(o["franja"])
            if franja_end is not None and franja_end < now_hour - 0.5:
                alerts.append({
                    "tipo":      "FRANJA_VENCIDA",
                    "severidad": "alta",
                    "orden":     o["id"],
                    "tecnico":   o["tecnico"],
                    "franja":    o["franja"],
                    "detalle":   f"Orden {o['id']} en franja ya vencida ({o['franja']})",
                })

    # 6. Técnico con franja activa iniciada y cero marcaciones
    # Escenario: franja arrancó hace ≥ ACTIVE_SLOT_NO_PROGRESS_MINUTES minutos,
    # el técnico tiene órdenes pendientes en esa franja y NO ha marcado NINGUNA orden
    # (progreso = 0 en TODAS sus órdenes). Indica posible ausencia, accidente o
    # desconexión en campo — requiere verificación inmediata del supervisor.
    threshold_hours = ACTIVE_SLOT_NO_PROGRESS_MINUTES / 60.0
    # Franjas que ya NO están completamente vencidas (evitar duplicar con FRANJA_VENCIDA)
    for tech, orders_list in tech_orders.items():
        if tech == "SIN_ASIGNAR":
            continue

        # ¿Ya tiene alguna orden marcada (en camino, en sitio, iniciado, etc.)?
        any_in_progress = any(o.get("progress", 0) >= 1 for o in orders_list)
        if any_in_progress:
            continue  # Técnico ya está trabajando — no aplica

        # Revisar cada franja en la que el técnico tiene órdenes
        for franja, total_in_franja in tech_franja.get(tech, {}).items():
            if franja == "Sin Franja":
                continue
            franja_start, franja_end = parse_franja_hours(franja)
            if franja_start is None:
                continue

            # ¿La franja está actualmente en curso (no completamente pasada)?
            if franja_end is not None and now_hour > franja_end:
                continue  # Ya venció — cubierto por FRANJA_VENCIDA

            # ¿La franja ya inició hace más del umbral mínimo?
            minutes_elapsed = (now_hour - franja_start) * 60
            if minutes_elapsed < ACTIVE_SLOT_NO_PROGRESS_MINUTES:
                continue  # Aún dentro del tiempo de tolerancia

            # Contar sólo las pendientes (movibles) en esa franja
            pending_in_slot = sum(
                1 for o in orders_list
                if o.get("franja") == franja and o.get("movible")
            )
            if pending_in_slot == 0:
                continue  # Sin órdenes activas en este slot

            # Severidad: muy_alta en cualquier caso; crítica si lleva >45 min
            sev = "critica" if minutes_elapsed >= 45 else "muy_alta"

            alerts.append({
                "tipo":               "FRANJA_ACTIVA_SIN_MARCACION",
                "severidad":          sev,
                "tecnico":            tech,
                "supervisor":         None,  # Campo disponible para mapeo futuro
                "franja":             franja,
                "zona":               idx["tech_main_zone"].get(tech, "SIN_ZONA"),
                "count_ordenes":      pending_in_slot,
                "minutos_sin_marcar": int(minutes_elapsed),
                "detalle": (
                    f"{tech}: {pending_in_slot} orden(es) en franja {franja} "
                    f"iniciada hace {int(minutes_elapsed)} min — "
                    f"SIN ninguna marcación. Verificar presencia, "
                    f"desplazamiento o novedad en campo."
                ),
            })

    # 7. Alerta preventiva: técnico cerca del fin de franja sin marcaciones
    # Diferente a Alert 6: esta es PROACTIVA — aún hay tiempo de reaccionar.
    # Dispara cuando faltan ≤ SLOT_RISK_MINUTES_BEFORE_END minutos para cerrar
    # la franja Y el técnico no ha marcado ninguna orden.
    # Permite al supervisor actuar antes del incumplimiento, no después.
    for tech, orders_list in tech_orders.items():
        if tech == "SIN_ASIGNAR":
            continue

        # Si ya tiene alguna marcación, no hay riesgo
        any_in_progress = any(o.get("progress", 0) >= 1 for o in orders_list)
        if any_in_progress:
            continue

        for franja in tech_franja.get(tech, {}):
            if franja == "Sin Franja":
                continue
            franja_start, franja_end = parse_franja_hours(franja)
            if franja_start is None or franja_end is None:
                continue

            # Solo franjas que aún no terminaron
            if now_hour >= franja_end:
                continue

            # ¿Estamos dentro de la ventana de riesgo (últimos N min)?
            minutes_to_end = (franja_end - now_hour) * 60
            if minutes_to_end > SLOT_RISK_MINUTES_BEFORE_END:
                continue

            # ¿Y la franja ya comenzó? (evitar alarma antes de que inicie)
            if now_hour < franja_start:
                continue

            pending_in_slot = sum(
                1 for o in orders_list
                if o.get("franja") == franja and o.get("movible")
            )
            if pending_in_slot == 0:
                continue

            minutes_elapsed = (now_hour - franja_start) * 60
            posible_impacto = (
                f"{pending_in_slot} orden(es) en riesgo de no ejecutarse. "
                f"Si no se marca en los próximos {int(minutes_to_end)} min, "
                f"quedarán como incumplimiento de franja."
            )

            alerts.append({
                "tipo":               "RIESGO_INCUMPLIMIENTO_FRANJA",
                "severidad":          "alta",
                "tecnico":            tech,
                "supervisor":         None,
                "franja":             franja,
                "zona":               idx["tech_main_zone"].get(tech, "SIN_ZONA"),
                "count_ordenes":      pending_in_slot,
                "tiempo_restante_min":int(minutes_to_end),
                "minutos_sin_marcar": int(minutes_elapsed),
                "posible_impacto":    posible_impacto,
                "detalle": (
                    f"⚠ RIESGO: {tech} tiene {pending_in_slot} orden(es) en franja "
                    f"{franja} y faltan solo {int(minutes_to_end)} min para cerrar — "
                    f"SIN marcación. Contactar y gestionar ahora."
                ),
            })

    return alerts


# ─── Motor de sugerencias ─────────────────────

def _score_suggestion(order: dict, donor: str, receiver: str,
                       idx: dict, current_hour: float) -> float:
    """
    Score de beneficio de mover 'order' de 'donor' a 'receiver'.
    Usa TOTALES de órdenes como métrica principal.
    Objetivo operativo: cada técnico con 4-5 órdenes.
    El coordinador decide si acepta o rechaza — el sistema sugiere, no bloquea.
    Las sugerencias que superan el ideal se muestran con riesgo='alto' para que
    el coordinador las evalúe conscientemente.
    """
    tech_orders  = idx["tech_orders"]
    tech_franja  = idx["tech_franja"]
    tech_subzones= idx["tech_subzones"]
    tech_locs    = idx["tech_locs"]
    tech_total   = idx["tech_total"]

    donor_total  = tech_total.get(donor, 0)
    recv_total   = tech_total.get(receiver, 0)

    # ─── Hard limit: nunca superar el límite absoluto ───
    if recv_total >= MAX_ABSOLUTE_LOAD:
        return -9999.0  # Límite duro: >6 no se sugiere nunca

    # ─── Penalización (no bloqueo) si receptor ya está en el ideal (5) ───
    # El coordinador puede decidir igualmente
    sobrecarga_receptor = recv_total >= MAX_IDEAL_LOAD

    # ─── Desequilibrio mínimo para que valga la sugerencia ───
    total_gap = donor_total - recv_total
    if total_gap < MIN_IMBALANCE_TO_MOVE and donor not in ("SIN_ASIGNAR", None):
        return -9999.0

    # ─── Score base: lógica separada según tipo de movimiento ───
    if donor in ("SIN_ASIGNAR", None):
        # Asignación de orden nueva (no hay desequilibrio que medir, solo capacidad + zona)
        # Base plana positiva para que el bono de zona pueda decidir correctamente
        score = 600
        if recv_total < MIN_IDEAL_LOAD:
            score += (MIN_IDEAL_LOAD - recv_total) * 500   # Prioridad a técnicos con déficit
        elif recv_total < MAX_IDEAL_LOAD:
            score += (MAX_IDEAL_LOAD - recv_total) * 200   # Bonus menor si tiene capacidad
        if sobrecarga_receptor:
            score -= 1200  # Penalizar si ya está en el ideal (visible pero baja prioridad)
    else:
        # Rebalanceo entre técnicos existentes
        score = total_gap * 800
        if sobrecarga_receptor:
            score -= 1500
        if recv_total < MIN_IDEAL_LOAD:
            score += (MIN_IDEAL_LOAD - recv_total) * 600
        if donor_total > MAX_IDEAL_LOAD:
            score += (donor_total - MAX_IDEAL_LOAD) * 500

    # ─── Verificar capacidad de franja ───
    can_add, _ = _can_add_to_franja(receiver, order["franja"], tech_franja,
                                     tech_orders, current_hour)
    if not can_add:
        return -9999.0
    score += 300

    # ─── Distancia (criterio secundario) ───
    dist_recv  = _dist_to_tech(order, receiver, tech_orders, tech_locs)
    dist_donor = _dist_to_tech(order, donor, tech_orders, tech_locs) if donor not in ("SIN_ASIGNAR", None) else None

    if dist_recv is not None and dist_recv > MAX_ALLOWED_DISTANCE_KM:
        score -= INTERZONE_DISTANCE_PENALTY
    if dist_recv is None:
        score -= 200

    if dist_donor is not None and dist_recv is not None:
        savings = dist_donor - dist_recv
        score += savings * 300  # Mayor peso al ahorro de distancia

    # ─── Bono directo por proximidad física al técnico receptor ───
    # Prioriza órdenes que caen dentro de la ruta actual del técnico
    if dist_recv is not None:
        if dist_recv < 2.0:
            score += 900  # Misma cuadra / muy cercano: altísimo valor operativo
        elif dist_recv < 5.0:
            score += 500  # Cercano: mejora ruta sin desvío significativo
        elif dist_recv < 10.0:
            score += 150  # Aceptable: leve desvío

    # ─── Bonus por zona: priorizar técnico de la misma zona que la orden ───
    recv_zone_main = idx["tech_main_zone"].get(receiver, "SIN_ZONA")
    order_zone_val = order.get("zona", "SIN_ZONA")
    if recv_zone_main == order_zone_val:
        score += 1200  # Mismo técnico de zona: prioridad clara
    elif order_zone_val in ZONE_ADJACENCY.get(recv_zone_main, []):
        score += 350   # Zona adyacente: también viable

    # Penalización por fragmentación de subzona del receptor
    recv_subzones = len(tech_subzones.get(receiver, set()))
    if order["subzona"] not in tech_subzones.get(receiver, set()) and recv_subzones >= 3:
        score -= FRAGMENTATION_PENALTY * 0.3

    return score


def _generate_suggestions(orders: list, idx: dict, current_hour: float) -> list:
    suggestions = []
    movable_orders = [o for o in orders if o["movible"]]
    techs = [t for t in idx["tech_orders"] if t != "SIN_ASIGNAR"]

    tech_orders  = idx["tech_orders"]
    tech_franja  = idx["tech_franja"]
    tech_total   = idx["tech_total"]
    tech_subzones= idx["tech_subzones"]
    tech_locs    = idx["tech_locs"]

    interzone_count = {}
    sugs_por_receptor = {}   # evitar monopolio: max 3 sugerencias por técnico receptor
    MAX_SUGS_RECEPTOR = 3

    for order in movable_orders:
        donor = order["tecnico"]
        donor_total = tech_total.get(donor, 0)

        # Solo procesar órdenes de técnicos sobrecargados O sin asignar
        if donor not in ("SIN_ASIGNAR", None) and donor_total <= MAX_IDEAL_LOAD:
            # ¿Hay algún técnico por debajo del mínimo que pueda recibir?
            hay_deficitario = any(
                tech_total.get(t, 0) < MIN_IDEAL_LOAD
                for t in techs if t != donor
            )
            if not hay_deficitario:
                continue  # No hay desequilibrio que justifique mover

        # Proteger técnico eficiente: si cumple criterios de rendimiento
        # no moverle órdenes salvo que su receptor sea muy deficitario
        if donor not in ("SIN_ASIGNAR", None) and _is_tech_efficient(
            donor, tech_orders, tech_franja, tech_subzones,
            idx["tech_credit"], idx["tech_pending"], idx["tech_eff_load"]
        ):
            hay_muy_deficitario = any(
                tech_total.get(t, 0) < MIN_IDEAL_LOAD - 1
                for t in techs if t != donor
            )
            if not hay_muy_deficitario:
                continue  # Técnico eficiente — no perturbar

        best_score = -9999.0
        best_receiver = None
        best_risk = "bajo"

        for receiver in techs:
            if receiver == donor:
                continue
            if tech_total.get(receiver, 0) >= MAX_ABSOLUTE_LOAD:
                continue
            if sugs_por_receptor.get(receiver, 0) >= MAX_SUGS_RECEPTOR:
                continue  # Ya acumuló suficientes sugerencias individuales

            score = _score_suggestion(order, donor, receiver, idx, current_hour)
            if score > best_score:
                best_score = score
                best_receiver = receiver

        if best_receiver is None or best_score < 0:
            continue

        donor_total_v = tech_total.get(donor, 0)
        recv_total_v  = tech_total.get(best_receiver, 0)
        donor_zone = idx["tech_main_zone"].get(donor, "SIN_ZONA")
        recv_zone  = idx["tech_main_zone"].get(best_receiver, "SIN_ZONA")
        order_zone = order["zona"]

        is_interzone = (recv_zone != order_zone and
                        order_zone not in ZONE_ADJACENCY.get(recv_zone, [recv_zone]))
        if is_interzone:
            best_risk = "alto"
            interzone_count[best_receiver] = interzone_count.get(best_receiver, 0) + 1
            if interzone_count[best_receiver] > MAX_INTERZONE_ASSIGNMENTS_PER_TECH:
                continue

        dist_recv  = _dist_to_tech(order, best_receiver, tech_orders, tech_locs)
        dist_donor = _dist_to_tech(order, donor, tech_orders, tech_locs) if donor not in ("SIN_ASIGNAR", None) else None

        # Motivo basado en totales (lógica correcta)
        if donor == "SIN_ASIGNAR":
            motivo = f"Orden sin técnico — asignar a {best_receiver} ({recv_total_v} órdenes, puede recibir más)"
        elif donor_total_v > MAX_IDEAL_LOAD:
            motivo = f"Sobrecarga: {donor} tiene {donor_total_v} órdenes (máx {MAX_IDEAL_LOAD}) → {best_receiver} tiene {recv_total_v}"
        elif recv_total_v < MIN_IDEAL_LOAD:
            motivo = f"Déficit: {best_receiver} tiene solo {recv_total_v} órdenes (mín {MIN_IDEAL_LOAD}) — necesita más"
        else:
            motivo = f"Balance: {donor} ({donor_total_v}) → {best_receiver} ({recv_total_v}) ords totales"

        beneficio = []
        if donor_total_v > MAX_IDEAL_LOAD:
            beneficio.append(f"Descarga a {donor} ({donor_total_v}→{donor_total_v-1})")
        if recv_total_v < MIN_IDEAL_LOAD:
            beneficio.append(f"Completa cuota de {best_receiver} ({recv_total_v}→{recv_total_v+1})")
        if dist_donor and dist_recv and dist_donor > dist_recv:
            beneficio.append(f"Ahorro ~{dist_donor - dist_recv:.1f}km")
        if not beneficio:
            beneficio.append("Mejora balance general de carga")

        # Advertencia si el receptor ya está en el ideal o lo supera
        aviso_sobrecarga = recv_total_v >= MAX_IDEAL_LOAD
        if aviso_sobrecarga:
            best_risk = "alto"
            motivo += f" ⚠ {best_receiver} quedará con {recv_total_v+1} órdenes (sobre el ideal de {MAX_IDEAL_LOAD})"

        suggestions.append({
            "orden":              order["id"],
            "tecnico_actual":     donor,
            "tecnico_sugerido":   best_receiver,
            "franja_actual":      order["franja"],
            "franja_sugerida":    order["franja"],
            "tipo":               order.get("tipo", ""),
            "estado":             order["estado"],
            "zona":               order_zone,
            "motivo":             motivo,
            "riesgo":             best_risk,
            "beneficio":          " / ".join(beneficio),
            "score":              round(best_score, 1),
            "interzona":          is_interzone,
            "aviso_sobrecarga":   aviso_sobrecarga,  # coordinador decide
            "total_receptor":     recv_total_v,
            "total_donante":      donor_total_v,
            "dist_receptor_km":   round(dist_recv, 2) if dist_recv else None,
        })
        sugs_por_receptor[best_receiver] = sugs_por_receptor.get(best_receiver, 0) + 1

    # Ordenar por score descendente
    suggestions.sort(key=lambda x: x["score"], reverse=True)
    return suggestions[:50]  # máximo 50 sugerencias



# ─── Rutas completas para técnicos con capacidad ─────────────────────

def _generate_route_suggestions(orders: list, idx: dict) -> list:
    """
    Para técnicos con capacidad (total < MIN_IDEAL_LOAD), agrupa las órdenes
    disponibles de su zona en una propuesta de ruta completa.
    En lugar de 5 sugerencias individuales para Wilson (1 orden),
    devuelve UNA propuesta: "Wilson puede tomar esta ruta de 4 órdenes en SABANETA".
    Solo incluye rutas que realmente mejoren la operación (calidad > cantidad).
    """
    rutas = []
    tech_total     = idx["tech_total"]
    tech_franja    = idx["tech_franja"]
    tech_main_zone = idx["tech_main_zone"]

    techs_con_capacidad = [
        t for t in idx["tech_orders"]
        if t != "SIN_ASIGNAR" and tech_total.get(t, 0) < MIN_IDEAL_LOAD
    ]
    if not techs_con_capacidad:
        return rutas

    # Órdenes candidatas: sin técnico O de técnicos sobrecargados
    candidatas = [
        o for o in orders
        if o["movible"] and (
            o["tecnico"] == "SIN_ASIGNAR"
            or tech_total.get(o["tecnico"], 0) > MAX_IDEAL_LOAD
        )
    ]
    if not candidatas:
        return rutas

    # Agrupar por zona
    ordenes_por_zona: dict = {}
    for o in candidatas:
        ordenes_por_zona.setdefault(o.get("zona", "SIN_ZONA"), []).append(o)

    vistas: set = set()

    for tech in techs_con_capacidad:
        cuota_actual = tech_total.get(tech, 0)
        capacidad    = MAX_IDEAL_LOAD - cuota_actual
        if capacidad <= 0:
            continue

        tech_zona = tech_main_zone.get(tech, "SIN_ZONA")
        zonas_a_revisar = [tech_zona] + list(ZONE_ADJACENCY.get(tech_zona, []))

        for zona in zonas_a_revisar:
            if (tech, zona) in vistas:
                continue
            disponibles = ordenes_por_zona.get(zona, [])
            if not disponibles:
                continue

            # Ordenar candidatas por proximidad al técnico (ruta más eficiente)
            tech_orders_local = idx["tech_orders"]
            tech_locs_local   = idx["tech_locs"]
            tech_ref = _tech_reference_point(tech, tech_orders_local, tech_locs_local)
            if tech_ref != (0.0, 0.0):
                def dist_to_ref(o):
                    if o.get("lat") and o.get("lon"):
                        return haversine(o["lat"], o["lon"], tech_ref[0], tech_ref[1])
                    return 999.0
                disponibles = sorted(disponibles, key=dist_to_ref)

            # Respetar límite de franja al armar la ruta
            franjas_ocupadas = dict(tech_franja.get(tech, {}))
            ordenes_validas  = []
            for o in disponibles:
                if len(ordenes_validas) >= capacidad:
                    break
                f = o.get("franja", "Sin Franja")
                if f != "Sin Franja" and franjas_ocupadas.get(f, 0) >= MAX_ORDERS_PER_SLOT:
                    continue
                if f != "Sin Franja":
                    franjas_ocupadas[f] = franjas_ocupadas.get(f, 0) + 1
                ordenes_validas.append(o)

            # Solo generar ruta si aporta al menos 2 órdenes (una sola no es "ruta")
            if len(ordenes_validas) < 2:
                continue

            vistas.add((tech, zona))
            es_zona_propia  = (zona == tech_zona)
            cuota_propuesta = cuota_actual + len(ordenes_validas)
            franjas_ruta    = sorted(set(o["franja"] for o in ordenes_validas if o["franja"] != "Sin Franja"))
            tipos_ruta      = list(set(o.get("tipo", "") for o in ordenes_validas))

            score_ruta = (
                len(ordenes_validas) * 900
                + (1200 if es_zona_propia else 350)
                + (MIN_IDEAL_LOAD - cuota_actual) * 600
            )

            motivo = (
                f"{tech} tiene {cuota_actual} orden(es) — puede completar cuota con "
                f"{len(ordenes_validas)} orden(es) disponibles en {zona}"
                f"{'  (su zona)' if es_zona_propia else ' (zona adyacente)'}. "
                f"Propuesta: {cuota_actual}\u2192{cuota_propuesta} órdenes."
            )

            rutas.append({
                "tipo_sugerencia":  "RUTA_COMPLETA",
                "tecnico_sugerido": tech,
                "zona":             zona,
                "es_zona_propia":   es_zona_propia,
                "ordenes":          [o["id"] for o in ordenes_validas],
                "ordenes_detalle":  [
                    {
                        "id":             o["id"],
                        "tipo":           o.get("tipo", ""),
                        "franja":         o["franja"],
                        "tecnico_actual": o["tecnico"],
                        "estado":         o["estado"],
                    }
                    for o in ordenes_validas
                ],
                "total_ordenes":    len(ordenes_validas),
                "cuota_actual":     cuota_actual,
                "cuota_propuesta":  cuota_propuesta,
                "franjas":          franjas_ruta,
                "tipos":            tipos_ruta,
                "motivo":           motivo,
                "riesgo":           "bajo" if cuota_propuesta <= MAX_IDEAL_LOAD else "medio",
                "score":            round(score_ruta, 1),
            })

    rutas.sort(key=lambda x: x["score"], reverse=True)
    return rutas[:10]

# ─── Punto de entrada principal ──────────────

def run_leveling(raw_orders: list) -> dict:
    """
    Recibe una lista de dicts (de Metabase o Excel),
    ejecuta el motor completo y devuelve el JSON de nivelación.
    """
    now_dt    = now_bogota()
    now_hour  = now_dt.hour + now_dt.minute / 60.0

    if not raw_orders:
        return _empty_result("Sin datos. Configura Metabase o sube un archivo.")

    # 1. Normalizar
    orders = [normalize_order(o) for o in raw_orders]

    # 2. Construir índices
    idx = _build_indexes(orders)

    # 3. Clasificar órdenes
    movibles   = [o for o in orders if o["movible"]]
    bloqueadas = [o for o in orders if not o["movible"]]

    # 4. Alertas
    alerts = _generate_alerts(orders, idx, now_dt)

    # 5. Sugerencias individuales
    suggestions = _generate_suggestions(orders, idx, now_hour)

    # 5b. Rutas completas para técnicos con capacidad disponible
    route_suggestions = _generate_route_suggestions(orders, idx)

    # 6. Carga por técnico
    carga_por_tecnico = []
    for tech, t_orders in sorted(idx["tech_orders"].items()):
        total = len(t_orders)
        movibles_t = sum(1 for o in t_orders if o["movible"])
        bloq_t     = total - movibles_t
        activas_t  = sum(1 for o in t_orders if 1 <= o.get("progress", 0) < 6)
        fin_t      = sum(1 for o in t_orders if o.get("progress", 0) >= 6)
        sobrecarga = total > MAX_IDEAL_LOAD
        franja_map = idx["tech_franja"].get(tech, {})
        subzones   = list(idx["tech_subzones"].get(tech, set()))
        carga_por_tecnico.append({
            "tecnico":    tech,
            "zona":       idx["tech_main_zone"].get(tech, "SIN_ZONA"),
            "total":      total,
            "movibles":   movibles_t,
            "bloqueadas": bloq_t,
            "activas":    activas_t,
            "finalizadas":fin_t,
            "sobrecarga": sobrecarga,
            "franjas":    franja_map,
            "subzonas":   subzones,
        })

    # 7. Carga por franja
    from config import FRANJAS
    franja_data = {f: {"total": 0, "movibles": 0, "bloqueadas": 0, "tecnicos": set()}
                   for f in FRANJAS}
    franja_data["Sin Franja"] = {"total": 0, "movibles": 0, "bloqueadas": 0, "tecnicos": set()}
    for o in orders:
        f = o["franja"]
        if f not in franja_data:
            franja_data[f] = {"total": 0, "movibles": 0, "bloqueadas": 0, "tecnicos": set()}
        franja_data[f]["total"] += 1
        if o["movible"]:
            franja_data[f]["movibles"] += 1
        else:
            franja_data[f]["bloqueadas"] += 1
        if o["tecnico"] != "SIN_ASIGNAR":
            franja_data[f]["tecnicos"].add(o["tecnico"])

    carga_por_franja = [
        {
            "franja":     f,
            "total":      d["total"],
            "movibles":   d["movibles"],
            "bloqueadas": d["bloqueadas"],
            "tecnicos":   len(d["tecnicos"]),
            "sobrecarga": d["total"] > MAX_ORDERS_PER_SLOT * len(d["tecnicos"]) if d["tecnicos"] else False,
        }
        for f, d in franja_data.items() if d["total"] > 0
    ]

    # 8. Enriquecer lista de órdenes para frontend
    def enrich_order(o):
        return {
            "id":          o["id"],
            "tecnico":     o["tecnico"],
            "estado":      o["estado"],
            "estado_clase":o["estado_clase"],
            "franja":      o["franja"],
            "tipo":        o["tipo"],
            "zona":        o["zona"],
            "subzona":     o["subzona"],
            "direccion":   o.get("direccion", ""),
            "movible":     o["movible"],
            "progress":    o["progress"],
            "lat":         o.get("lat", 0),
            "lon":         o.get("lon", 0),
        }

    # 9. Resumen
    alerta_critica      = sum(1 for a in alerts if a.get("severidad") in ("critica", "alta"))
    alerta_muy_alta     = sum(1 for a in alerts if a.get("severidad") == "muy_alta")
    techs_sobrecargados = sum(1 for c in carga_por_tecnico if c["sobrecarga"] and c["tecnico"] != "SIN_ASIGNAR")
    techs_con_capacidad = sum(1 for t, total in idx["tech_total"].items() if t != "SIN_ASIGNAR" and total < MAX_IDEAL_LOAD)
    techs_deficitarios  = sum(1 for t, total in idx["tech_total"].items() if t != "SIN_ASIGNAR" and total < MIN_IDEAL_LOAD)
    # Técnicos sin marcar en franja activa (para resumen ejecutivo)
    techs_sin_marcacion = len(set(
        a["tecnico"] for a in alerts
        if a.get("tipo") in ("FRANJA_ACTIVA_SIN_MARCACION", "RIESGO_INCUMPLIMIENTO_FRANJA")
    ))
    alertas_riesgo_franja = sum(
        1 for a in alerts if a.get("tipo") == "RIESGO_INCUMPLIMIENTO_FRANJA"
    )

    return {
        "generado_en": now_dt.strftime("%Y-%m-%d %H:%M:%S"),
        "resumen": {
            "total_ordenes":          len(orders),
            "movibles":               len(movibles),
            "bloqueadas":             len(bloqueadas),
            "alertas":                len(alerts),
            "alertas_criticas":       alerta_critica,
            "alertas_muy_altas":      alerta_muy_alta,
            "alertas_riesgo_franja":  alertas_riesgo_franja,
            "tecnicos_sin_marcacion": techs_sin_marcacion,
            "sugerencias":            len(suggestions),
            "rutas_sugeridas":         len(route_suggestions),
            "tecnicos_total":         len([t for t in idx["tech_orders"] if t != "SIN_ASIGNAR"]),
            "tecnicos_sobrecargados": techs_sobrecargados,
            "tecnicos_con_capacidad": techs_con_capacidad,
            "tecnicos_deficitarios":  techs_deficitarios,
            "sin_tecnico":            sum(1 for o in movibles if o["tecnico"] == "SIN_ASIGNAR"),
            "sin_franja":             sum(1 for o in movibles if o["franja"] == "Sin Franja"),
            "sin_franja":             sum(1 for o in movibles if o["franja"] == "Sin Franja"),
            "objetivo_por_tecnico":   f"{MIN_IDEAL_LOAD}-{MAX_IDEAL_LOAD} ordenes",
        },
        "carga_por_tecnico": carga_por_tecnico,
        "carga_por_franja":  carga_por_franja,
        "ordenes_movibles":  [enrich_order(o) for o in movibles],
        "ordenes_bloqueadas":[enrich_order(o) for o in bloqueadas],
        "alertas":           alerts,
        "sugerencias":       suggestions,
        "rutas_sugeridas":    route_suggestions,
    }


def _empty_result(msg: str) -> dict:
    return {
        "generado_en": now_bogota().strftime("%Y-%m-%d %H:%M:%S"),
        "mensaje":     msg,
        "resumen": {
            "total_ordenes": 0, "movibles": 0, "bloqueadas": 0,
            "alertas": 0, "alertas_criticas": 0, "sugerencias": 0,
            "alertas_muy_altas": 0, "alertas_riesgo_franja": 0,
            "tecnicos_sin_marcacion": 0,
            "tecnicos_total": 0, "tecnicos_sobrecargados": 0,
            "tecnicos_con_capacidad": 0, "tecnicos_deficitarios": 0,
            "sin_tecnico": 0, "sin_franja": 0,
            "objetivo_por_tecnico": "4-5 ordenes",
        },
        "carga_por_tecnico": [], "carga_por_franja": [],
        "ordenes_movibles": [], "ordenes_bloqueadas": [],
        "alertas": [], "sugerencias": [], "rutas_sugeridas": [],
    }
