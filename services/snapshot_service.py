"""
services/snapshot_service.py
Registro evolutivo de órdenes por día.
Cada vez que se sube el Excel se guarda un corte con:
  - Total y conteo por estado
  - Conteo por franja horaria
  - IDs para detectar cambios entre cortes (reprogramadas)
Sin datos por técnico. Compatible con Render free tier.
"""
import io
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

# Almacenamiento en memoria: {fecha_str: [corte, ...]}
_store: dict = {}

# Estados agrupados para el reporte
_GRUPOS = {
    "Programadas":    {"programado", "programada"},
    "Por Programar":  {"por programar"},
    "En Proceso":     {"en camino", "en sitio", "iniciado", "iniciada",
                       "mac enviada", "mac principal enviada",
                       "dispositivos subidos", "dispositivos cargados"},
    "Finalizadas":    {"finalizado", "finalizada", "cerrado", "cerrada"},
    "Por Auditar":    {"por auditar"},
    "Canceladas":     {"cancelado", "cancelada"},
}

# Orden de visualización en la tabla
CAMPOS = ["Total", "Programadas", "Por Programar", "Reprogramadas",
          "En Proceso", "Finalizadas", "Por Auditar", "Canceladas"]


def _now_naive() -> datetime:
    """Datetime naive en hora Bogotá."""
    from config import now_bogota
    dt = now_bogota()
    return dt.replace(tzinfo=None) if getattr(dt, "tzinfo", None) else dt


def _hora_label(dt: datetime) -> str:
    """Etiqueta descriptiva por hora del día."""
    h = dt.hour
    if 5 <= h < 10:   return f"Mañana {dt.strftime('%H:%M')}"
    if 10 <= h < 14:  return f"Mediodía {dt.strftime('%H:%M')}"
    if 14 <= h < 19:  return f"Tarde {dt.strftime('%H:%M')}"
    return f"Cierre {dt.strftime('%H:%M')}"


def _clasificar(orders: list) -> dict:
    """Clasifica ordenes por estado, franja y tipo de cita."""
    por_estado = {g: 0 for g in _GRUPOS}
    por_franja: dict = {}
    por_tipo:   dict = {}
    ids_programadas = set()
    ids_todos = set()
    for o in orders:
        if not isinstance(o, dict):
            continue
        est    = str(o.get("estado", "")).strip().lower()
        franja = str(o.get("franja", "Sin Franja")).strip() or "Sin Franja"
        tipo   = str(o.get("tipo", "otro")).strip().lower() or "otro"
        oid    = str(o.get("id", ""))
        ids_todos.add(oid)
        for grupo, estados in _GRUPOS.items():
            if est in estados:
                por_estado[grupo] += 1
                break
        if est in {"programado", "programada"}:
            ids_programadas.add(oid)
        por_franja[franja] = por_franja.get(franja, 0) + 1
        por_tipo[tipo] = por_tipo.get(tipo, 0) + 1
    return {"total":len(orders),"por_estado":por_estado,
            "por_franja":dict(sorted(por_franja.items())),
            "por_tipo":dict(sorted(por_tipo.items())),
            "_ids_prog":list(ids_programadas),"_ids_todos":list(ids_todos),"reprogramadas":0}


def registrar_corte(orders: list, label: str = None) -> dict:
    """
    Registra un corte con el estado actual de las órdenes.
    Compara automáticamente con el corte anterior para calcular reprogramadas.
    Devuelve el corte registrado.
    """
    now   = _now_naive()
    fecha = now.strftime("%Y-%m-%d")
    hora  = now.strftime("%H:%M:%S")   # Único por segundo
    etiq  = label or _hora_label(now)

    stats = _clasificar(orders)

    # Detectar reprogramadas: programadas en corte anterior que ya no están
    dia = _store.get(fecha, [])
    if dia:
        anterior = dia[-1]
        ids_prog_ant = set(anterior.get("_ids_prog", []))
        ids_ahora    = set(stats["_ids_todos"])
        stats["reprogramadas"] = len(ids_prog_ant - ids_ahora)

    corte = {
        "id":           f"{fecha}_{hora}",
        "label":        etiq,
        "fecha":        fecha,
        "hora":         now.strftime("%H:%M"),
        "timestamp":    now.isoformat(sep=" ", timespec="seconds"),
        "total":        stats["total"],
        "por_estado":   stats["por_estado"],
        "por_franja":   stats["por_franja"],
        "por_tipo":     stats["por_tipo"],
        "reprogramadas": stats["reprogramadas"],
        "_ids_prog":    stats["_ids_prog"],
        "_ids_todos":   stats["_ids_todos"],
    }

    if fecha not in _store:
        _store[fecha] = []
    _store[fecha].append(corte)

    logger.info(
        f"Corte '{etiq}' | total={corte['total']} "
        f"prog={stats['por_estado'].get('Programadas',0)} "
        f"fin={stats['por_estado'].get('Finalizadas',0)} "
        f"cancel={stats['por_estado'].get('Canceladas',0)} "
        f"reprog={corte['reprogramadas']}"
    )
    return corte


def get_cortes(fecha: str = None) -> list:
    """Devuelve los cortes del día (hoy por defecto)."""
    if not fecha:
        fecha = _now_naive().strftime("%Y-%m-%d")
    return list(_store.get(fecha, []))


def get_fechas() -> list:
    return sorted(_store.keys(), reverse=True)


def _delta_color(campo: str, delta: int) -> str:
    """Color semántico: verde=bueno, rojo=atención."""
    if delta == 0:
        return "00000000"   # Transparente
    bien_subir  = {"Finalizadas", "Por Auditar"}
    bien_bajar  = {"Canceladas", "Por Programar", "Reprogramadas"}
    if campo in bien_subir:
        return "D5E8D4" if delta > 0 else "F8CECC"
    if campo in bien_bajar:
        return "D5E8D4" if delta < 0 else "F8CECC"
    return "FFF2CC"   # Amarillo suave para neutros


def generar_excel(fecha: str = None) -> bytes:
    """
    Genera Excel con la evolución del día.
    Hojas:
      1. Evolución por Estado (comparativa entre cortes)
      2. Evolución por Franja (comparativa entre cortes)
      3. Detalle completo de cada corte
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise RuntimeError("openpyxl no instalado")

    if not fecha:
        fecha = _now_naive().strftime("%Y-%m-%d")

    cortes = get_cortes(fecha)
    wb = Workbook()

    # ── Helpers de estilo ──────────────────────────────────────────────
    def F(bold=False, color="000000", sz=10, italic=False):
        return Font(bold=bold, color=color, size=sz, italic=italic)

    def Fill(hex_):
        return PatternFill("solid", fgColor=hex_) if hex_ and hex_ != "00000000" else PatternFill()

    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT   = Alignment(horizontal="left",   vertical="center")
    THIN   = Border(*[Side(style="thin")] * 4)

    HDR_BG    = "1F3864"
    COL_BG    = "2E5FA3"
    FRANJA_BG = "1A5276"
    TITLE_BG  = "D6E4F0"

    # ── Hoja 1: Evolución por Estado ──────────────────────────────────
    ws1 = wb.active
    ws1.title = "Por Estado"

    # Título
    cols_total = 1 + len(cortes) + max(len(cortes) - 1, 0)
    ws1.merge_cells(f"A1:{get_column_letter(cols_total + 1)}1")
    ws1["A1"] = f"Evolución de Órdenes por Estado — {fecha}"
    ws1["A1"].font = F(bold=True, sz=13, color=HDR_BG)
    ws1["A1"].fill = Fill(TITLE_BG)
    ws1["A1"].alignment = CENTER
    ws1.row_dimensions[1].height = 28

    if not cortes:
        ws1["A3"] = "Sin cortes registrados para esta fecha."
        ws1["A3"].font = F(color="888888", italic=True)
    else:
        # Encabezados
        r = 3
        def hdr(cell, txt, bg=HDR_BG):
            cell.value = txt
            cell.font = F(bold=True, color="FFFFFF", sz=10)
            cell.fill = Fill(bg)
            cell.alignment = CENTER
            cell.border = THIN

        hdr(ws1.cell(r, 1), "Estado")
        col = 2
        slot_cols  = []
        delta_cols = []

        for c in cortes:
            hdr(ws1.cell(r, col), f"{c['label']}\n{c['hora']}", bg=COL_BG)
            slot_cols.append(col)
            col += 1
        for i in range(len(cortes) - 1):
            hdr(ws1.cell(r, col), f"Δ {cortes[i]['hora']} → {cortes[i+1]['hora']}", bg="4A235A")
            delta_cols.append(col)
            col += 1

        ws1.row_dimensions[r].height = 36

        # Fila Total
        r += 1
        ws1.cell(r, 1, "Total").font = F(bold=True)
        ws1.cell(r, 1).border = THIN
        for i, c in enumerate(cortes):
            cell = ws1.cell(r, slot_cols[i], c["total"])
            cell.font = F(bold=True, color=HDR_BG, sz=11)
            cell.alignment = CENTER
            cell.border = THIN
        for i, dcol in enumerate(delta_cols):
            d = cortes[i+1]["total"] - cortes[i]["total"]
            arrow = "▲" if d>0 else "▼" if d<0 else "→"
            cell = ws1.cell(r, dcol, f"{arrow} {d:+d}")
            cell.alignment = CENTER
            cell.border = THIN
            cell.font = F(bold=True, sz=10)

        # Fila Reprogramadas
        r += 1
        ws1.cell(r, 1, "Reprogramadas (salieron del día)").font = F(italic=True)
        ws1.cell(r, 1).border = THIN
        for i, c in enumerate(cortes):
            v = c.get("reprogramadas", 0)
            cell = ws1.cell(r, slot_cols[i], v)
            cell.alignment = CENTER
            cell.border = THIN
            if v > 0:
                cell.fill = Fill("F8CECC")
        for i, dcol in enumerate(delta_cols):
            d = cortes[i+1].get("reprogramadas", 0) - cortes[i].get("reprogramadas", 0)
            cell = ws1.cell(r, dcol, f"{'▲' if d>0 else '▼' if d<0 else '→'} {d:+d}" if d != 0 else "→ 0")
            cell.alignment = CENTER
            cell.border = THIN
            cell.fill = Fill(_delta_color("Reprogramadas", d))

        # Filas por estado
        for campo in list(_GRUPOS.keys()):
            r += 1
            ws1.cell(r, 1, campo).border = THIN
            ws1.cell(r, 1).alignment = LEFT
            for i, c in enumerate(cortes):
                v = c["por_estado"].get(campo, 0)
                cell = ws1.cell(r, slot_cols[i], v)
                cell.alignment = CENTER
                cell.border = THIN
            for i, dcol in enumerate(delta_cols):
                va = cortes[i]["por_estado"].get(campo, 0)
                vb = cortes[i+1]["por_estado"].get(campo, 0)
                d  = vb - va
                arrow = "▲" if d>0 else "▼" if d<0 else "→"
                cell = ws1.cell(r, dcol, f"{arrow} {d:+d}")
                cell.alignment = CENTER
                cell.border = THIN
                cell.fill = Fill(_delta_color(campo, d))
                cell.font = F(bold=True, sz=10)

        # Anchos
        ws1.column_dimensions["A"].width = 32
        for c in range(2, col):
            ws1.column_dimensions[get_column_letter(c)].width = 16
        ws1.freeze_panes = "B4"

    # ── Hoja 2: Evolución por Franja ──────────────────────────────────
    ws2 = wb.create_sheet("Por Franja")

    # Recoger todas las franjas únicas
    all_franjas = []
    for c in cortes:
        for f in c.get("por_franja", {}):
            if f not in all_franjas:
                all_franjas.append(f)
    all_franjas = sorted(all_franjas)

    ws2.merge_cells(f"A1:{get_column_letter(max(2, cols_total + 1))}1")
    ws2["A1"] = f"Evolución de Visitas por Franja Horaria — {fecha}"
    ws2["A1"].font = F(bold=True, sz=13, color=HDR_BG)
    ws2["A1"].fill = Fill(TITLE_BG)
    ws2["A1"].alignment = CENTER
    ws2.row_dimensions[1].height = 28

    if cortes and all_franjas:
        r2 = 3
        def hdr2(cell, txt, bg=FRANJA_BG):
            cell.value = txt
            cell.font = F(bold=True, color="FFFFFF", sz=10)
            cell.fill = Fill(bg)
            cell.alignment = CENTER
            cell.border = THIN

        hdr2(ws2.cell(r2, 1), "Franja Horaria")
        col2 = 2
        s_cols, d_cols = [], []
        for c in cortes:
            hdr2(ws2.cell(r2, col2), f"{c['label']}\n{c['hora']}", bg=FRANJA_BG)
            s_cols.append(col2)
            col2 += 1
        for i in range(len(cortes) - 1):
            hdr2(ws2.cell(r2, col2), f"Δ {cortes[i]['hora']} → {cortes[i+1]['hora']}", bg="1A3A4A")
            d_cols.append(col2)
            col2 += 1
        ws2.row_dimensions[r2].height = 36

        for franja in all_franjas:
            r2 += 1
            ws2.cell(r2, 1, franja).border = THIN
            ws2.cell(r2, 1).alignment = LEFT
            es_sin = franja.lower() == "sin franja"
            if es_sin:
                ws2.cell(r2, 1).font = F(italic=True, color="888888")

            for i, c in enumerate(cortes):
                v = c["por_franja"].get(franja, 0)
                cell = ws2.cell(r2, s_cols[i], v)
                cell.alignment = CENTER
                cell.border = THIN
                if es_sin and v > 0:
                    cell.fill = Fill("FFF2CC")

            for i, dcol in enumerate(d_cols):
                va = cortes[i]["por_franja"].get(franja, 0)
                vb = cortes[i+1]["por_franja"].get(franja, 0)
                d  = vb - va
                arrow = "▲" if d>0 else "▼" if d<0 else "→"
                cell = ws2.cell(r2, dcol, f"{arrow} {d:+d}" if d != 0 else "→ 0")
                cell.alignment = CENTER
                cell.border = THIN
                # Para franjas: más visitas no es necesariamente malo
                cell.fill = Fill("FFF2CC") if abs(d) > 0 else PatternFill()

        ws2.column_dimensions["A"].width = 22
        for c in range(2, col2):
            ws2.column_dimensions[get_column_letter(c)].width = 16
        ws2.freeze_panes = "B4"

    # ── Hoja 3: Detalle de cada corte ────────────────────────────────
    for corte in cortes:
        safe = corte["label"][:28]
        ws3 = wb.create_sheet(safe)

        ws3.merge_cells("A1:C1")
        ws3["A1"] = f"{corte['label']}  —  {corte['fecha']} {corte['hora']}"
        ws3["A1"].font = F(bold=True, sz=12, color=HDR_BG)
        ws3["A1"].fill = Fill(TITLE_BG)
        ws3["A1"].alignment = CENTER
        ws3.row_dimensions[1].height = 24

        # Estados
        ws3.cell(3, 1, "Estado").font = F(bold=True, color="FFFFFF")
        ws3.cell(3, 1).fill = Fill(HDR_BG)
        ws3.cell(3, 2, "Cantidad").font = F(bold=True, color="FFFFFF")
        ws3.cell(3, 2).fill = Fill(HDR_BG)
        for c in [ws3.cell(3, 1), ws3.cell(3, 2)]:
            c.alignment = CENTER
            c.border = THIN

        filas = [("Total", corte["total"]),
                 ("Reprogramadas (salieron del día)", corte.get("reprogramadas", 0))]
        filas += [(k, v) for k, v in corte["por_estado"].items()]

        for ri, (k, v) in enumerate(filas, start=4):
            ws3.cell(ri, 1, k).border = THIN
            ws3.cell(ri, 1).alignment = LEFT
            cell = ws3.cell(ri, 2, v)
            cell.alignment = CENTER
            cell.border = THIN
            if k == "Total":
                ws3.cell(ri, 1).font = F(bold=True)
                cell.font = F(bold=True, color=HDR_BG)
            if k == "Reprogramadas (salieron del día)" and v > 0:
                cell.fill = Fill("F8CECC")

        # Franjas
        r3 = len(filas) + 6
        ws3.cell(r3, 1, "Franja Horaria").font = F(bold=True, color="FFFFFF")
        ws3.cell(r3, 1).fill = Fill(FRANJA_BG)
        ws3.cell(r3, 1).alignment = CENTER
        ws3.cell(r3, 2, "Visitas").font = F(bold=True, color="FFFFFF")
        ws3.cell(r3, 2).fill = Fill(FRANJA_BG)
        ws3.cell(r3, 2).alignment = CENTER
        for c in [ws3.cell(r3, 1), ws3.cell(r3, 2)]:
            c.border = THIN

        for franja, v in sorted(corte["por_franja"].items()):
            r3 += 1
            ws3.cell(r3, 1, franja).border = THIN
            ws3.cell(r3, 1).alignment = LEFT
            cell = ws3.cell(r3, 2, v)
            cell.alignment = CENTER
            cell.border = THIN
            if franja.lower() == "sin franja" and v > 0:
                cell.fill = Fill("FFF2CC")

        ws3.column_dimensions["A"].width = 32
        ws3.column_dimensions["B"].width = 12

    # ── Hoja de instrucciones ─────────────────────────────────────────
    wsi = wb.create_sheet("Instrucciones")
    lines = [
        ("Reporte Evolutivo — Nivelación Pro", True, 13),
        ("", False, 10),
        ("Cómo funciona:", True, 11),
        ("• Cada vez que subes el Excel al dashboard se guarda un corte automáticamente.", False, 10),
        ("• El corte registra: estado de las órdenes + visitas por franja horaria.", False, 10),
        ("• 'Reprogramadas' = órdenes que estaban como Programadas y ya no aparecen en el siguiente corte.", False, 10),
        ("  Se entiende que fueron movidas a otro día o canceladas entre cortes.", False, 10),
        ("", False, 10),
        ("Lectura de los deltas (columnas Δ):", True, 11),
        ("  ▲ con fondo verde = mejora operativa (más finalizadas, menos canceladas)", False, 10),
        ("  ▲ con fondo rojo  = requiere atención (más canceladas, más sin franja)", False, 10),
        ("  → sin cambio", False, 10),
        ("", False, 10),
        ("Hojas incluidas:", True, 11),
        ("• 'Por Estado' — evolución de cada estado entre cortes del día", False, 10),
        ("• 'Por Franja' — evolución de visitas por franja horaria", False, 10),
        ("• Un detalle por cada corte registrado", False, 10),
    ]
    for ri, (txt, bold, sz) in enumerate(lines, start=1):
        c = wsi.cell(ri, 1, txt)
        c.font = F(bold=bold, sz=sz)
    wsi.column_dimensions["A"].width = 85

    # ── Guardar ───────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    logger.info(f"Excel generado: {len(cortes)} cortes del {fecha}")
    return buf.getvalue()
