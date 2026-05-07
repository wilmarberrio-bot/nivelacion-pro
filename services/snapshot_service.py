"""
services/snapshot_service.py
Registro evolutivo de ordenes por dia para el apartado de reportes.

Reglas del reporte:
  - Cada foto/corte guarda la hora exacta de captura.
  - Se permiten varios cortes en el mismo dia.
  - Solo se conserva informacion del dia actual.
  - El Excel descargable corresponde unicamente al dia actual.
  - Se comparan cortes para detectar diferencias por estado, franja y orden.
"""
import io
import json
import logging
import os
import tempfile
from datetime import datetime

logger = logging.getLogger(__name__)

# Almacenamiento en memoria y respaldo temporal del dia actual.
# En Render esto evita perder cortes durante el mismo proceso, sin conservar dias anteriores.
_store: dict = {}
_STORE_FILE = os.environ.get(
    "REPORT_SNAPSHOT_FILE",
    os.path.join(tempfile.gettempdir(), "nivelacion_pro_reporte_hoy.json"),
)

_GRUPOS = {
    "Programadas":    {"programado", "programada"},
    "Por Programar":  {"por programar"},
    "En Proceso":     {"en camino", "en sitio", "iniciado", "iniciada",
                       "mac enviada", "mac principal enviada",
                       "dispositivos subidos", "dispositivos cargados"},
    "Finalizadas":    {"finalizado", "finalizada", "cerrado", "cerrada", "completado", "completada"},
    "Por Auditar":    {"por auditar"},
    "Canceladas":     {"cancelado", "cancelada", "cancelada por cliente", "cancelado por cliente"},
}

CAMPOS = ["Total", "Programadas", "Por Programar", "Reprogramadas",
          "En Proceso", "Finalizadas", "Por Auditar", "Canceladas"]


def _now_naive() -> datetime:
    """Datetime naive en hora Bogota."""
    from config import now_bogota
    dt = now_bogota()
    return dt.replace(tzinfo=None) if getattr(dt, "tzinfo", None) else dt


def _today() -> str:
    return _now_naive().strftime("%Y-%m-%d")


def _norm_estado(value) -> str:
    """Normaliza estado para que el reporte no dependa de mayusculas o tildes."""
    try:
        from services.normalization import norm_status
        return norm_status(value)
    except Exception:
        s = str(value or "").strip().lower()
        for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u")):
            s = s.replace(a, b)
        return s


def _estado_grupo(estado: str) -> str:
    s = _norm_estado(estado)
    if "cancelad" in s:
        return "Canceladas"
    if "por auditar" in s:
        return "Por Auditar"
    if any(x in s for x in ("finaliz", "cerrad", "completad")):
        return "Finalizadas"
    if "por programar" in s:
        return "Por Programar"
    if any(x in s for x in ("programad", "programado")):
        return "Programadas"
    if any(x in s for x in ("en camino", "en sitio", "iniciad", "mac", "dispositivos")):
        return "En Proceso"
    return "Otros"


def _hora_label(dt: datetime) -> str:
    """Etiqueta descriptiva por hora del dia."""
    h = dt.hour
    if 5 <= h < 10:
        return f"Manana {dt.strftime('%H:%M:%S')}"
    if 10 <= h < 14:
        return f"Mediodia {dt.strftime('%H:%M:%S')}"
    if 14 <= h < 19:
        return f"Tarde {dt.strftime('%H:%M:%S')}"
    return f"Cierre {dt.strftime('%H:%M:%S')}"


def _purge_old_days() -> None:
    """Mantiene solamente cortes del dia actual y borra respaldos viejos."""
    global _store
    hoy = _today()
    if _store and set(_store.keys()) != {hoy}:
        _store = {hoy: _store.get(hoy, [])}
    try:
        if os.path.exists(_STORE_FILE):
            with open(_STORE_FILE, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
            if payload.get("fecha") != hoy:
                os.remove(_STORE_FILE)
    except Exception as exc:
        logger.warning("No se pudo depurar reporte diario: %s", exc)


def _load_store() -> None:
    """Carga el respaldo temporal del dia actual, si existe."""
    global _store
    hoy = _today()
    _purge_old_days()
    if _store.get(hoy):
        return
    try:
        if not os.path.exists(_STORE_FILE):
            _store.setdefault(hoy, [])
            return
        with open(_STORE_FILE, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
        if payload.get("fecha") == hoy:
            _store = {hoy: payload.get("cortes", []) or []}
        else:
            _store = {hoy: []}
    except Exception as exc:
        logger.warning("No se pudo cargar reporte diario: %s", exc)
        _store.setdefault(hoy, [])


def _save_store() -> None:
    """Guarda solo el reporte del dia actual."""
    hoy = _today()
    os.makedirs(os.path.dirname(_STORE_FILE), exist_ok=True)
    payload = {"fecha": hoy, "cortes": _store.get(hoy, [])}
    tmp = f"{_STORE_FILE}.tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False)
    os.replace(tmp, _STORE_FILE)


def _order_id(order: dict, pos: int) -> str:
    oid = str(order.get("id") or order.get("orden") or order.get("appointment_id") or "").strip()
    return oid if oid and oid.lower() not in ("none", "nan") else f"row_{pos}"


def _clasificar(orders: list) -> dict:
    """Clasifica ordenes y conserva estado por ID para comparar cortes."""
    por_estado = {g: 0 for g in _GRUPOS}
    por_estado.setdefault("Otros", 0)
    por_franja: dict = {}
    por_tipo: dict = {}
    order_state: dict = {}

    for pos, o in enumerate(orders):
        if not isinstance(o, dict):
            continue
        oid = _order_id(o, pos)
        estado_raw = o.get("estado", "")
        estado_norm = _norm_estado(estado_raw)
        grupo = _estado_grupo(estado_raw)
        por_estado[grupo] = por_estado.get(grupo, 0) + 1

        franja = str(o.get("franja", "Sin Franja") or "Sin Franja").strip() or "Sin Franja"
        tipo = str(o.get("tipo", "otro") or "otro").strip().lower() or "otro"
        tecnico = str(o.get("tecnico", "SIN_ASIGNAR") or "SIN_ASIGNAR").strip() or "SIN_ASIGNAR"

        por_franja[franja] = por_franja.get(franja, 0) + 1
        por_tipo[tipo] = por_tipo.get(tipo, 0) + 1
        order_state[oid] = {
            "estado": estado_norm,
            "grupo": grupo,
            "franja": franja,
            "tecnico": tecnico,
            "tipo": tipo,
        }

    return {
        "total": len(order_state),
        "por_estado": por_estado,
        "por_franja": dict(sorted(por_franja.items())),
        "por_tipo": dict(sorted(por_tipo.items())),
        "_order_state": order_state,
        "reprogramadas": 0,
        "nuevas": 0,
        "salieron": 0,
        "cambios_estado": 0,
        "cambios_franja": 0,
        "canceladas_nuevas": 0,
        "detalle_cambios": [],
    }


def _comparar(anterior: dict, actual: dict) -> None:
    """Calcula diferencias entre el corte anterior y el actual."""
    prev = anterior.get("_order_state", {}) or {}
    curr = actual.get("_order_state", {}) or {}
    prev_ids = set(prev.keys())
    curr_ids = set(curr.keys())

    salieron = sorted(prev_ids - curr_ids)
    nuevas = sorted(curr_ids - prev_ids)
    comunes = sorted(prev_ids & curr_ids)

    detalle = []
    cambios_estado = 0
    cambios_franja = 0
    canceladas_nuevas = 0

    for oid in comunes:
        p = prev[oid]
        c = curr[oid]
        if p.get("grupo") != c.get("grupo"):
            cambios_estado += 1
            detalle.append({
                "orden": oid,
                "tipo": "estado",
                "antes": p.get("grupo"),
                "despues": c.get("grupo"),
            })
        if p.get("franja") != c.get("franja"):
            cambios_franja += 1
            detalle.append({
                "orden": oid,
                "tipo": "franja",
                "antes": p.get("franja"),
                "despues": c.get("franja"),
            })
        if c.get("grupo") == "Canceladas" and p.get("grupo") != "Canceladas":
            canceladas_nuevas += 1

    for oid in salieron[:100]:
        detalle.append({"orden": oid, "tipo": "salio_del_dia", "antes": prev[oid].get("grupo"), "despues": "No aparece"})
    for oid in nuevas[:100]:
        detalle.append({"orden": oid, "tipo": "nueva", "antes": "No aparecia", "despues": curr[oid].get("grupo")})

    actual["salieron"] = len(salieron)
    actual["nuevas"] = len(nuevas)
    actual["reprogramadas"] = len(salieron)
    actual["cambios_estado"] = cambios_estado
    actual["cambios_franja"] = cambios_franja
    actual["canceladas_nuevas"] = canceladas_nuevas
    actual["detalle_cambios"] = detalle[:300]


def registrar_corte(orders: list, label: str = None) -> dict:
    """Registra una foto del dia con hora exacta y comparacion contra el corte anterior."""
    _load_store()
    now = _now_naive()
    fecha = now.strftime("%Y-%m-%d")
    hora_exacta = now.strftime("%H:%M:%S")
    etiqueta = label or _hora_label(now)

    stats = _clasificar(orders)
    dia = _store.setdefault(fecha, [])
    if dia:
        _comparar(dia[-1], stats)

    corte = {
        "id": f"{fecha}_{hora_exacta}_{len(dia) + 1}",
        "label": etiqueta,
        "fecha": fecha,
        "hora": hora_exacta,
        "hora_exacta": hora_exacta,
        "timestamp": now.isoformat(sep=" ", timespec="seconds"),
        "total": stats["total"],
        "por_estado": stats["por_estado"],
        "por_franja": stats["por_franja"],
        "por_tipo": stats["por_tipo"],
        "reprogramadas": stats["reprogramadas"],
        "salieron": stats["salieron"],
        "nuevas": stats["nuevas"],
        "cambios_estado": stats["cambios_estado"],
        "cambios_franja": stats["cambios_franja"],
        "canceladas_nuevas": stats["canceladas_nuevas"],
        "detalle_cambios": stats["detalle_cambios"],
        "_order_state": stats["_order_state"],
    }
    dia.append(corte)
    _save_store()

    logger.info(
        "Corte %s %s | total=%s cancel=%s reprog=%s estado=%s franja=%s",
        fecha, hora_exacta, corte["total"], corte["por_estado"].get("Canceladas", 0),
        corte["reprogramadas"], corte["cambios_estado"], corte["cambios_franja"],
    )
    return corte


def _public(corte: dict) -> dict:
    return {k: v for k, v in corte.items() if not k.startswith("_")}


def get_cortes(fecha: str = None) -> list:
    """Devuelve cortes solo del dia actual. Fechas anteriores vencen automaticamente."""
    _load_store()
    hoy = _today()
    if fecha and fecha != hoy:
        return []
    return [_public(c) for c in _store.get(hoy, [])]


def get_fechas() -> list:
    """El selector solo muestra el dia vigente para no mezclar reportes."""
    _load_store()
    hoy = _today()
    return [hoy] if _store.get(hoy) else []


def _delta_color(campo: str, delta: int) -> str:
    if delta == 0:
        return "00000000"
    bien_subir = {"Finalizadas", "Por Auditar"}
    bien_bajar = {"Canceladas", "Por Programar", "Reprogramadas"}
    if campo in bien_subir:
        return "D5E8D4" if delta > 0 else "F8CECC"
    if campo in bien_bajar:
        return "D5E8D4" if delta < 0 else "F8CECC"
    return "FFF2CC"


def generar_excel(fecha: str = None) -> bytes:
    """Genera el unico Excel descargable del dia actual."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise RuntimeError("openpyxl no instalado")

    hoy = _today()
    if fecha and fecha != hoy:
        raise ValueError("El reporte solicitado ya vencio. Solo se puede descargar el informe del dia actual.")
    fecha = hoy
    cortes = get_cortes(fecha)

    wb = Workbook()

    def F(bold=False, color="000000", sz=10, italic=False):
        return Font(bold=bold, color=color, size=sz, italic=italic)

    def Fill(hex_):
        return PatternFill("solid", fgColor=hex_) if hex_ and hex_ != "00000000" else PatternFill()

    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center")
    THIN = Border(*[Side(style="thin")] * 4)
    HDR_BG = "1F3864"
    COL_BG = "2E5FA3"
    FRANJA_BG = "1A5276"
    TITLE_BG = "D6E4F0"

    cols_total = max(2, 1 + len(cortes) + max(len(cortes) - 1, 0))

    ws1 = wb.active
    ws1.title = "Por Estado"
    ws1.merge_cells(f"A1:{get_column_letter(cols_total)}1")
    ws1["A1"] = f"Evolucion de Ordenes por Estado - {fecha}"
    ws1["A1"].font = F(bold=True, sz=13, color=HDR_BG)
    ws1["A1"].fill = Fill(TITLE_BG)
    ws1["A1"].alignment = CENTER

    if not cortes:
        ws1["A3"] = "Sin cortes registrados para esta fecha."
        ws1["A3"].font = F(color="888888", italic=True)
    else:
        def hdr(cell, txt, bg=HDR_BG):
            cell.value = txt
            cell.font = F(bold=True, color="FFFFFF", sz=10)
            cell.fill = Fill(bg)
            cell.alignment = CENTER
            cell.border = THIN

        r = 3
        hdr(ws1.cell(r, 1), "Estado")
        col = 2
        slot_cols, delta_cols = [], []
        for c in cortes:
            hdr(ws1.cell(r, col), f"{c['label']}\n{c.get('hora_exacta') or c['hora']}", bg=COL_BG)
            slot_cols.append(col)
            col += 1
        for i in range(len(cortes) - 1):
            hdr(ws1.cell(r, col), f"Delta {cortes[i]['hora']} -> {cortes[i+1]['hora']}", bg="4A235A")
            delta_cols.append(col)
            col += 1
        ws1.row_dimensions[r].height = 38

        campos = ["Total", "Programadas", "Por Programar", "Reprogramadas", "En Proceso",
                  "Finalizadas", "Por Auditar", "Canceladas", "Nuevas", "Cambios Estado", "Cambios Franja"]

        def val(corte, campo):
            if campo == "Total": return corte.get("total", 0)
            if campo == "Reprogramadas": return corte.get("reprogramadas", 0)
            if campo == "Nuevas": return corte.get("nuevas", 0)
            if campo == "Cambios Estado": return corte.get("cambios_estado", 0)
            if campo == "Cambios Franja": return corte.get("cambios_franja", 0)
            return (corte.get("por_estado") or {}).get(campo, 0)

        for campo in campos:
            r += 1
            ws1.cell(r, 1, campo).border = THIN
            ws1.cell(r, 1).alignment = LEFT
            if campo == "Total":
                ws1.cell(r, 1).font = F(bold=True)
            for i, c in enumerate(cortes):
                cell = ws1.cell(r, slot_cols[i], val(c, campo))
                cell.alignment = CENTER
                cell.border = THIN
                if campo == "Total":
                    cell.font = F(bold=True, color=HDR_BG)
            for i, dcol in enumerate(delta_cols):
                d = val(cortes[i + 1], campo) - val(cortes[i], campo)
                arrow = "UP" if d > 0 else "DOWN" if d < 0 else "="
                cell = ws1.cell(r, dcol, f"{arrow} {d:+d}" if d else "= 0")
                cell.alignment = CENTER
                cell.border = THIN
                cell.fill = Fill(_delta_color(campo, d))
                cell.font = F(bold=True, sz=10)

        ws1.column_dimensions["A"].width = 24
        for c in range(2, col):
            ws1.column_dimensions[get_column_letter(c)].width = 18
        ws1.freeze_panes = "B4"

    ws2 = wb.create_sheet("Por Franja")
    all_franjas = sorted({f for c in cortes for f in (c.get("por_franja") or {})})
    ws2.merge_cells(f"A1:{get_column_letter(max(2, cols_total))}1")
    ws2["A1"] = f"Evolucion por Franja Horaria - {fecha}"
    ws2["A1"].font = F(bold=True, sz=13, color=HDR_BG)
    ws2["A1"].fill = Fill(TITLE_BG)
    ws2["A1"].alignment = CENTER

    if cortes and all_franjas:
        def hdr2(cell, txt, bg=FRANJA_BG):
            cell.value = txt
            cell.font = F(bold=True, color="FFFFFF", sz=10)
            cell.fill = Fill(bg)
            cell.alignment = CENTER
            cell.border = THIN
        r2 = 3
        hdr2(ws2.cell(r2, 1), "Franja Horaria")
        col2 = 2
        s_cols, d_cols = [], []
        for c in cortes:
            hdr2(ws2.cell(r2, col2), f"{c['label']}\n{c['hora']}")
            s_cols.append(col2)
            col2 += 1
        for _ in range(len(cortes) - 1):
            hdr2(ws2.cell(r2, col2), "Delta", bg="1A3A4A")
            d_cols.append(col2)
            col2 += 1
        for franja in all_franjas:
            r2 += 1
            ws2.cell(r2, 1, franja).alignment = LEFT
            ws2.cell(r2, 1).border = THIN
            for i, c in enumerate(cortes):
                cell = ws2.cell(r2, s_cols[i], (c.get("por_franja") or {}).get(franja, 0))
                cell.alignment = CENTER
                cell.border = THIN
            for i, dcol in enumerate(d_cols):
                d = (cortes[i + 1].get("por_franja") or {}).get(franja, 0) - (cortes[i].get("por_franja") or {}).get(franja, 0)
                cell = ws2.cell(r2, dcol, f"{d:+d}" if d else "0")
                cell.alignment = CENTER
                cell.border = THIN
        ws2.column_dimensions["A"].width = 22
        for c in range(2, col2):
            ws2.column_dimensions[get_column_letter(c)].width = 16
        ws2.freeze_panes = "B4"

    ws3 = wb.create_sheet("Cambios")
    ws3.append(["Corte", "Hora", "Orden", "Tipo cambio", "Antes", "Despues"])
    for cell in ws3[1]:
        cell.font = F(bold=True, color="FFFFFF")
        cell.fill = Fill(HDR_BG)
        cell.alignment = CENTER
        cell.border = THIN
    for corte in cortes:
        for d in corte.get("detalle_cambios", []) or []:
            ws3.append([corte.get("label"), corte.get("hora"), d.get("orden"), d.get("tipo"), d.get("antes"), d.get("despues")])
    for col, width in zip("ABCDEF", [22, 12, 18, 18, 24, 24]):
        ws3.column_dimensions[col].width = width

    for corte in cortes:
        safe = (corte["label"][:24] + " " + corte["hora"].replace(":", ""))[:31]
        ws = wb.create_sheet(safe)
        ws.merge_cells("A1:C1")
        ws["A1"] = f"{corte['label']} - {corte['fecha']} {corte['hora']}"
        ws["A1"].font = F(bold=True, sz=12, color=HDR_BG)
        ws["A1"].fill = Fill(TITLE_BG)
        ws["A1"].alignment = CENTER
        ws.append([])
        ws.append(["Estado", "Cantidad"])
        for cell in ws[3]:
            cell.font = F(bold=True, color="FFFFFF")
            cell.fill = Fill(HDR_BG)
            cell.border = THIN
            cell.alignment = CENTER
        filas = [("Total", corte.get("total", 0)), ("Reprogramadas", corte.get("reprogramadas", 0)),
                 ("Nuevas", corte.get("nuevas", 0)), ("Cambios Estado", corte.get("cambios_estado", 0)),
                 ("Cambios Franja", corte.get("cambios_franja", 0))]
        filas += [(k, v) for k, v in (corte.get("por_estado") or {}).items()]
        for k, v in filas:
            ws.append([k, v])
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 14

    wsi = wb.create_sheet("Instrucciones")
    lines = [
        "Reporte diario de seguimiento fotografico",
        "Cada corte guarda la hora exacta en que se tomo la foto.",
        "El archivo contiene solamente cortes del dia actual.",
        "Al cambiar de dia, los cortes anteriores vencen y no quedan disponibles.",
        "Reprogramadas = ordenes que estaban en el corte anterior y ya no aparecen en el corte actual.",
        "Cambios Estado y Cambios Franja se calculan comparando las mismas ordenes entre cortes.",
    ]
    for i, txt in enumerate(lines, start=1):
        wsi.cell(i, 1, txt).font = F(bold=(i == 1), sz=12 if i == 1 else 10)
    wsi.column_dimensions["A"].width = 110

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    logger.info("Excel de reporte generado: %s cortes del %s", len(cortes), fecha)
    return buf.getvalue()
