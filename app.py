from flask import Flask, request, send_file, render_template_string, redirect, url_for, flash
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from copy import copy
from io import BytesIO
import tempfile
import os
import re

# PDF
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

app = Flask(__name__)
app.secret_key = "top5_secret_key"

HTML = '''
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>TOP MERMAS - Generador</title>
  <style>
    body { font-family: Arial, Helvetica, sans-serif; max-width: 960px; margin: 30px auto; padding: 0 20px; color: #1f1f1f; background: #f5f7fa; }
    h2 { color: #1F4E78; }
    h3 { color: #2E75B6; margin-top: 0; }
    .card { background: #fff; border-radius: 8px; padding: 24px 28px; margin-bottom: 24px; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
    label { display: inline-block; margin-bottom: 8px; font-size: 0.95em; }
    input[type=file] { display: block; margin: 8px 0 14px; }
    .btn-row { display: flex; gap: 12px; flex-wrap: wrap; margin-top: 4px; }
    button { padding: 9px 20px; border: none; border-radius: 5px; cursor: pointer; font-size: 0.95em; font-weight: bold; }
    .btn-excel { background: #1F7A3C; color: #fff; }
    .btn-excel:hover { background: #165c2d; }
    .btn-pdf   { background: #C0392B; color: #fff; }
    .btn-pdf:hover   { background: #962d22; }
    hr { border: none; border-top: 1px solid #dce3ec; margin: 8px 0 24px; }
    p.hint { color: #666; font-size: 0.88em; margin-top: 4px; }
  </style>
</head>
<body>
  <h2>TOP MERMAS — Generador</h2>

  <div class="card">
    <h3>Generar histórico (sube varios TOP MERMAS)</h3>
    <p class="hint">Sube los ficheros TOP en orden cronológico. Puedes descargar el resultado como Excel o PDF.</p>
    <form action="/historico" method="post" enctype="multipart/form-data" id="form-hist">
      <label>Ficheros TOP MERMAS (xlsx):</label>
      <input type="file" name="files" multiple accept=".xlsx,.xlsm,.xls" required>
      <div class="btn-row">
        <button type="submit" class="btn-excel" formaction="/historico">⬇ Excel</button>
        <button type="submit" class="btn-pdf"   formaction="/historico_pdf">⬇ PDF</button>
      </div>
    </form>
  </div>

  <div class="card">
    <h3>Actualizar histórico acumulado</h3>
    <p class="hint">Sube tu histórico acumulado anterior y añade los TOP nuevos que quieras incorporar.</p>
    <form action="/historico_acumulado" method="post" enctype="multipart/form-data" id="form-acum">
      <label>Histórico acumulado anterior (.xlsx):</label>
      <input type="file" name="master" accept=".xlsx" required>
      <label>Nuevos ficheros TOP a incorporar:</label>
      <input type="file" name="files" multiple accept=".xlsx,.xlsm,.xls" required>
      <div class="btn-row">
        <button type="submit" class="btn-excel" formaction="/historico_acumulado">⬇ Excel</button>
        <button type="submit" class="btn-pdf"   formaction="/historico_acumulado_pdf">⬇ PDF</button>
      </div>
    </form>
  </div>

  <p class="hint">Si necesitas subir tus propias plantillas, contacta para activar la opción.</p>
</body>
</html>
'''


# ─────────────────────────────────────────────
#  UTILIDADES COMUNES
# ─────────────────────────────────────────────

def copy_sheet_exact(src_ws, tgt_wb, title):
    tgt = tgt_wb.create_sheet(title=title)
    try:
        for col, dim in src_ws.column_dimensions.items():
            if getattr(dim, "width", None) is not None:
                tgt.column_dimensions[col].width = dim.width
    except Exception:
        pass
    try:
        for r, dim in src_ws.row_dimensions.items():
            if getattr(dim, "height", None) is not None:
                tgt.row_dimensions[r].height = dim.height
    except Exception:
        pass
    try:
        for merged in src_ws.merged_cells.ranges:
            tgt.merge_cells(str(merged))
    except Exception:
        pass
    for r in range(1, src_ws.max_row + 1):
        for c in range(1, src_ws.max_column + 1):
            s = src_ws.cell(row=r, column=c)
            t = tgt.cell(row=r, column=c, value=s.value)
            try:
                if s.has_style:
                    t.font = copy(s.font)
                    t.border = copy(s.border)
                    t.fill = copy(s.fill)
                    t.number_format = copy(s.number_format)
                    t.protection = copy(s.protection)
                    t.alignment = copy(s.alignment)
            except Exception:
                pass
    try:
        tgt.sheet_view = copy(src_ws.sheet_view)
    except Exception:
        pass
    return tgt


def extract_sem_ciclo_from_name(fname):
    sem = cic = None
    m = re.search(r"[sS][eE][mM][^0-9]*([0-9]{1,2})", fname)
    if m:
        sem = m.group(1)
    m2 = re.search(r"[cC]\s*[_\-]?\s*([0-9]{1,2})", fname)
    if m2:
        cic = m2.group(1)
    if not sem:
        m3 = re.search(r"\bS[_\s\-]?([0-9]{1,2})\b", fname, flags=re.IGNORECASE)
        if m3:
            sem = m3.group(1)
    if not cic:
        m4 = re.search(r"\bC[_\s\-]?([0-9]{1,2})\b", fname, flags=re.IGNORECASE)
        if m4:
            cic = m4.group(1)
    return sem, cic


def sanitize_sheet_name(name):
    s = re.sub(r'[:\\/*?\[\]]', '-', str(name))
    s = s.replace("\n", " ").strip()
    return s[:31]


def _extract_top_models_and_label(file_storage):
    file_bytes = BytesIO(file_storage.read())
    wb = load_workbook(file_bytes, data_only=True)

    top_name = None
    for name in wb.sheetnames:
        if 'top' in name.lower() or 'calidad' in name.lower():
            top_name = name
            break
    if not top_name:
        top_name = wb.sheetnames[0]

    ws = wb[top_name]
    headers = [(ws.cell(row=1, column=c).value or "") for c in range(1, ws.max_column + 1)]
    low = [str(h).lower() for h in headers]

    def find_idx(keys):
        for k in keys:
            for i, h in enumerate(low):
                if k in h:
                    return i + 1
        return None

    mc_idx  = find_idx(['mc', 'modelo', 'model', 'codigo', 'articulo']) or 1
    fam_idx = find_idx(['familia', 'family', 'grupo'])

    sem, cic = extract_sem_ciclo_from_name(file_storage.filename)
    if not sem or not cic:
        top_sem = ws["B2"].value
        top_cic = ws["D2"].value
        if not sem:
            m = re.search(r'\d+', str(top_sem) if top_sem is not None else "")
            if m: sem = m.group(0)
        if not cic:
            m = re.search(r'\d+', str(top_cic) if top_cic is not None else "")
            if m: cic = m.group(0)

    label = f"SEM {sem} C{cic}" if sem or cic else file_storage.filename

    models, families = set(), {}
    for r in range(2, ws.max_row + 1):
        mc = ws.cell(row=r, column=mc_idx).value
        if mc is None:
            continue
        mc = str(mc).strip()
        if not mc:
            continue
        models.add(mc)
        if fam_idx:
            fam = ws.cell(row=r, column=fam_idx).value
            if fam is not None and mc not in families:
                fam_txt = str(fam).strip()
                if fam_txt:
                    families[mc] = fam_txt

    return label, models, families


def _seed_records_from_master_workbook(master_wb):
    records, order = {}, []

    if '_ORDEN' in master_wb.sheetnames:
        ws_order = master_wb['_ORDEN']
        for r in range(2, ws_order.max_row + 1):
            label = ws_order.cell(row=r, column=1).value
            if label:
                label = str(label).strip()
                if label:
                    order.append(label)

    if 'DETALLE' not in master_wb.sheetnames:
        return records, order

    ws = master_wb['DETALLE']
    for r in range(2, ws.max_row + 1):
        mc = ws.cell(row=r, column=1).value
        if mc is None:
            continue
        mc = str(mc).strip()
        if not mc:
            continue

        count = ws.cell(row=r, column=2).value
        try: count = int(count)
        except Exception: count = 0

        ciclos_raw = ws.cell(row=r, column=3).value or ""
        labels = [c.strip() for c in str(ciclos_raw).split(",") if c.strip()]

        first_label = str(ws.cell(row=r, column=4).value or "").strip()
        last_label  = str(ws.cell(row=r, column=5).value or "").strip()
        family      = str(ws.cell(row=r, column=9).value or "").strip() if ws.max_column >= 9 else ""

        records[mc] = {
            "count": count if count else len(labels),
            "labels": labels,
            "first_label": first_label or (labels[0] if labels else "HISTORICO ACUMULADO"),
            "last_label":  last_label  or (labels[-1] if labels else "HISTORICO ACUMULADO"),
            "family": family,
        }

    return records, order


def _sequence_stats(order, labels):
    order_map = {label: idx for idx, label in enumerate(order)}
    positions = sorted(set(order_map[l] for l in labels if l in order_map))
    if not positions:
        return 0
    longest = current = 1
    for i in range(1, len(positions)):
        if positions[i] == positions[i-1] + 1:
            current += 1
            longest = max(longest, current)
        else:
            current = 1
    return longest


def _build_historico_rows(records, order, last_seen_models, last_label):
    rows = []
    family_totals, family_models = {}, {}
    persistent_total_count = persistent_consecutive_count = 0

    for mc, rec in records.items():
        labels = rec.get("labels", [])
        family = rec.get("family") or "Sin familia"
        appears = mc in last_seen_models
        is_recurrent = len(labels) > 1

        racha_max = _sequence_stats(order, labels)
        persistent_total = len(labels) >= 3
        persistent_consecutive = racha_max >= 3

        if appears and is_recurrent:
            prioridad, estado = "ALTA", "Reincidente"
        elif appears:
            prioridad, estado = "MEDIA", "Nuevo"
        else:
            prioridad, estado = "BAJA", "Histórico"

        if persistent_total:      persistent_total_count += 1
        if persistent_consecutive: persistent_consecutive_count += 1

        family_totals[family] = family_totals.get(family, 0) + len(labels)
        family_models.setdefault(family, set()).add(mc)

        rows.append({
            "mc": mc, "count": len(labels), "ciclos": ", ".join(labels),
            "first_label": rec.get("first_label", labels[0] if labels else ""),
            "last_label":  rec.get("last_label",  labels[-1] if labels else ""),
            "appears_in_last_top": "Sí" if appears else "No",
            "estado": estado, "prioridad": prioridad, "family": family,
            "persistencia_total": "Sí" if persistent_total else "No",
            "persistencia_consecutiva": "Sí" if persistent_consecutive else "No",
            "racha_max": racha_max,
        })

    priority_rank = {"ALTA": 0, "MEDIA": 1, "BAJA": 2}
    estado_rank   = {"Reincidente": 0, "Nuevo": 1, "Histórico": 2}
    rows.sort(key=lambda r: (
        priority_rank.get(r["prioridad"], 99),
        estado_rank.get(r["estado"], 99),
        -r["count"], r["first_label"], r["mc"]
    ))

    total_models    = len(rows)
    new_models      = sum(1 for r in rows if r["estado"] == "Nuevo")
    repeated_models = sum(1 for r in rows if r["count"] > 1)

    family_rows = sorted([
        {"family": f, "apariciones": t, "modelos_unicos": len(family_models.get(f, set()))}
        for f, t in family_totals.items()
    ], key=lambda x: (-x["apariciones"], -x["modelos_unicos"], x["family"]))

    alerts = [r for r in rows if r["prioridad"] == "ALTA"][:10]

    return (rows, total_models, new_models, repeated_models,
            persistent_total_count, persistent_consecutive_count,
            family_rows, alerts)


# ─────────────────────────────────────────────
#  RENDER EXCEL
# ─────────────────────────────────────────────

def _render_historico_visual(rows, total_models, new_models, repeated_models,
                              persistent_total_count, persistent_consecutive_count,
                              family_rows, alerts, order, sheet_title="HISTORICO"):

    wb_out = Workbook()
    ws_res = wb_out.active
    ws_res.title = "RESUMEN"
    ws_res.sheet_view.showGridLines = False

    title_font  = Font(bold=True, size=14, color="FFFFFF")
    header_font = Font(bold=True, color="FFFFFF")
    body_font   = Font(bold=False, color="1F1F1F")
    bold_font   = Font(bold=True, color="1F1F1F")

    fill_title     = PatternFill("solid", fgColor="1F4E78")
    fill_header    = PatternFill("solid", fgColor="4F81BD")
    fill_card      = PatternFill("solid", fgColor="D9EAF7")
    fill_green     = PatternFill("solid", fgColor="E2F0D9")
    fill_amber     = PatternFill("solid", fgColor="FFF2CC")
    fill_red       = PatternFill("solid", fgColor="F4CCCC")
    fill_gray      = PatternFill("solid", fgColor="E7E6E6")
    fill_blue_soft = PatternFill("solid", fgColor="DDEBF7")

    thin   = Side(style="thin",   color="7F7F7F")
    medium = Side(style="medium", color="404040")
    border_thin   = Border(left=thin,   right=thin,   top=thin,   bottom=thin)
    border_medium = Border(left=medium, right=medium, top=medium, bottom=medium)

    ws_res.merge_cells("A1:E1")
    ws_res["A1"] = f"{sheet_title} - RESUMEN HISTÓRICO TOP MERMAS"
    ws_res["A1"].font = title_font
    ws_res["A1"].fill = fill_title
    ws_res["A1"].alignment = Alignment(horizontal="center", vertical="center")

    summary_cards = [
        ("Total de modelos",           total_models,               fill_card),
        ("Cuántos son nuevos",         new_models,                 fill_green if new_models else fill_gray),
        ("Cuántos se repiten",         repeated_models,            fill_amber if repeated_models else fill_gray),
        ("Persistentes 3+ apariciones",persistent_total_count,     fill_blue_soft if persistent_total_count else fill_gray),
        ("Persistentes 3+ seguidos",   persistent_consecutive_count,fill_blue_soft if persistent_consecutive_count else fill_gray),
        ("Alertas ALTA",               len(alerts),                fill_red if alerts else fill_gray),
    ]

    for i, (label, value, fill) in enumerate(summary_cards, start=3):
        ws_res.merge_cells(start_row=i, start_column=1, end_row=i, end_column=2)
        ws_res.merge_cells(start_row=i, start_column=3, end_row=i, end_column=5)
        c1 = ws_res.cell(row=i, column=1, value=label)
        c2 = ws_res.cell(row=i, column=3, value=value)
        for c, align in ((c1, "left"), (c2, "center")):
            c.font = bold_font; c.fill = fill; c.border = border_medium
            c.alignment = Alignment(horizontal=align, vertical="center")

    # Alertas
    alert_start = 11
    ws_res.merge_cells(start_row=alert_start, start_column=1, end_row=alert_start, end_column=5)
    c = ws_res.cell(row=alert_start, column=1, value="ALERTAS AUTOMÁTICAS (PRIORIDAD ALTA)")
    c.font = Font(bold=True, color="FFFFFF"); c.fill = fill_title
    c.alignment = Alignment(horizontal="center", vertical="center")

    for col, h in enumerate(["MC","Familia","Veces","Estado","Prioridad"], 1):
        cell = ws_res.cell(row=alert_start+1, column=col, value=h)
        cell.font = header_font; cell.fill = fill_header
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center", vertical="center")

    if alerts:
        for i, r in enumerate(alerts, start=alert_start+2):
            for col, v in enumerate([r["mc"],r["family"],r["count"],r["estado"],r["prioridad"]], 1):
                cell = ws_res.cell(row=i, column=col, value=v)
                cell.border = border_thin; cell.font = body_font; cell.fill = fill_red
                cell.alignment = Alignment(horizontal="left" if col in (1,2,4,5) else "center", vertical="center")
    else:
        ws_res.merge_cells(start_row=alert_start+2, start_column=1, end_row=alert_start+2, end_column=5)
        c = ws_res.cell(row=alert_start+2, column=1, value="No hay modelos con prioridad ALTA.")
        c.border = border_thin; c.fill = fill_gray
        c.alignment = Alignment(horizontal="center", vertical="center")

    # Familias
    family_start = alert_start + 7
    ws_res.merge_cells(start_row=family_start, start_column=1, end_row=family_start, end_column=5)
    c = ws_res.cell(row=family_start, column=1, value="FAMILIAS CON MÁS MERMAS")
    c.font = Font(bold=True, color="FFFFFF"); c.fill = fill_title
    c.alignment = Alignment(horizontal="center", vertical="center")

    for col, h in enumerate(["Familia","Apariciones","Modelos únicos"], 1):
        cell = ws_res.cell(row=family_start+1, column=col, value=h)
        cell.font = header_font; cell.fill = fill_header
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, fam in enumerate(family_rows[:10], start=family_start+2):
        for col, v in enumerate([fam["family"], fam["apariciones"], fam["modelos_unicos"]], 1):
            cell = ws_res.cell(row=i, column=col, value=v)
            cell.border = border_thin; cell.font = body_font
            cell.alignment = Alignment(horizontal="left" if col==1 else "center", vertical="center")
            cell.fill = fill_blue_soft if i % 2 == 0 else fill_card

    # Top reincidentes
    top_start = family_start + 14
    ws_res.merge_cells(start_row=top_start, start_column=1, end_row=top_start, end_column=6)
    c = ws_res.cell(row=top_start, column=1, value="TOP 10 REINCIDENTES")
    c.font = Font(bold=True, color="FFFFFF"); c.fill = fill_title
    c.alignment = Alignment(horizontal="center", vertical="center")

    for col, h in enumerate(["MC","Familia","Veces","Ciclos","Estado","Prioridad"], 1):
        cell = ws_res.cell(row=top_start+1, column=col, value=h)
        cell.font = header_font; cell.fill = fill_header
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center", vertical="center")

    top_10 = sorted([r for r in rows if r["count"]>1], key=lambda x: (-x["count"], x["mc"]))[:10]
    for i, r in enumerate(top_10, start=top_start+2):
        for col, v in enumerate([r["mc"],r["family"],r["count"],r["ciclos"],r["estado"],r["prioridad"]], 1):
            cell = ws_res.cell(row=i, column=col, value=v)
            cell.border = border_thin; cell.font = body_font
            cell.alignment = Alignment(horizontal="left" if col in (1,2,4,5,6) else "center", vertical="center")
            if r["prioridad"]=="ALTA":      cell.fill = fill_red
            elif r["prioridad"]=="MEDIA":   cell.fill = fill_amber
            else: cell.fill = fill_green if r["estado"]=="Nuevo" else fill_gray

    for col, w in {1:24,2:18,3:12,4:56,5:16,6:14}.items():
        ws_res.column_dimensions[get_column_letter(col)].width = w
    ws_res.freeze_panes = "A3"

    # Hoja DETALLE
    ws_det = wb_out.create_sheet("DETALLE")
    ws_det.sheet_view.showGridLines = False
    det_headers = [
        "MC","Veces","Ciclos","Primera aparición","Última aparición",
        "¿Aparece en el último TOP?","Estado","Prioridad","Familia",
        "Persistencia 3+ apariciones","Persistencia 3+ seguidos","Racha máx.",
    ]
    for col, h in enumerate(det_headers, 1):
        cell = ws_det.cell(row=1, column=col, value=h)
        cell.font = header_font; cell.fill = fill_header
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, r in enumerate(rows, start=2):
        vals = [r["mc"],r["count"],r["ciclos"],r["first_label"],r["last_label"],
                r["appears_in_last_top"],r["estado"],r["prioridad"],r["family"],
                r["persistencia_total"],r["persistencia_consecutiva"],r["racha_max"]]
        for col, v in enumerate(vals, 1):
            cell = ws_det.cell(row=row_idx, column=col, value=v)
            cell.border = border_thin; cell.font = body_font
            cell.alignment = Alignment(horizontal="left" if col not in (2,12) else "center", vertical="center")
            if r["prioridad"]=="ALTA":    cell.fill = fill_red
            elif r["prioridad"]=="MEDIA": cell.fill = fill_amber
            else: cell.fill = fill_green if r["estado"]=="Nuevo" else fill_gray

    for col, w in {1:18,2:10,3:44,4:18,5:18,6:24,7:16,8:14,9:18,10:20,11:20,12:12}.items():
        ws_det.column_dimensions[get_column_letter(col)].width = w
    ws_det.freeze_panes = "A2"
    ws_det.auto_filter.ref = ws_det.dimensions

    # Hoja _ORDEN (oculta)
    ws_ord = wb_out.create_sheet("_ORDEN")
    ws_ord.sheet_state = "hidden"
    ws_ord["A1"] = "Etiqueta"
    for i, label in enumerate(order, start=2):
        ws_ord.cell(row=i, column=1, value=label)

    output = BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output


# ─────────────────────────────────────────────
#  RENDER PDF  (NUEVO)
# ─────────────────────────────────────────────

def _render_historico_pdf(rows, total_models, new_models, repeated_models,
                           persistent_total_count, persistent_consecutive_count,
                           family_rows, alerts, order, sheet_title="HISTORICO"):

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm,
    )

    # ── Colores corporativos ──
    C_DARK  = colors.HexColor("#1F4E78")
    C_MID   = colors.HexColor("#2E75B6")
    C_LIGHT = colors.HexColor("#DDEBF7")
    C_GREEN = colors.HexColor("#E2F0D9")
    C_AMBER = colors.HexColor("#FFF2CC")
    C_RED   = colors.HexColor("#F4CCCC")
    C_GRAY  = colors.HexColor("#E7E6E6")
    C_WHITE = colors.white
    C_BLACK = colors.HexColor("#1F1F1F")

    styles = getSampleStyleSheet()

    style_title = ParagraphStyle("title", fontSize=18, textColor=C_WHITE,
                                  fontName="Helvetica-Bold", alignment=TA_CENTER, spaceAfter=4)
    style_subtitle = ParagraphStyle("sub", fontSize=10, textColor=C_WHITE,
                                     fontName="Helvetica", alignment=TA_CENTER)
    style_section = ParagraphStyle("section", fontSize=12, textColor=C_WHITE,
                                    fontName="Helvetica-Bold", alignment=TA_LEFT,
                                    spaceBefore=14, spaceAfter=4)
    style_body = ParagraphStyle("body", fontSize=8, textColor=C_BLACK,
                                 fontName="Helvetica", leading=11)
    style_bold = ParagraphStyle("bold", fontSize=8, textColor=C_BLACK,
                                 fontName="Helvetica-Bold", leading=11)
    style_center = ParagraphStyle("center", fontSize=8, textColor=C_BLACK,
                                   fontName="Helvetica", alignment=TA_CENTER, leading=11)

    # helper: encabezado de sección con banda de color
    def section_header(text):
        tbl = Table([[Paragraph(text, style_section)]], colWidths=[doc.width])
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), C_DARK),
            ("TOPPADDING",    (0,0),(-1,-1), 6),
            ("BOTTOMPADDING", (0,0),(-1,-1), 6),
            ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ]))
        return tbl

    def hdr_cell(text):
        return Paragraph(text, ParagraphStyle("hdr", fontSize=8, textColor=C_WHITE,
                                               fontName="Helvetica-Bold", alignment=TA_CENTER, leading=10))
    def body_cell(text, align=TA_LEFT):
        return Paragraph(str(text), ParagraphStyle("bc", fontSize=7.5, textColor=C_BLACK,
                                                    fontName="Helvetica", alignment=align, leading=10))

    story = []

    # ── Portada / Cabecera ──
    header_tbl = Table(
        [[Paragraph(f"{sheet_title} — TOP MERMAS", style_title),
          Paragraph(f"Total modelos: {total_models}", style_subtitle)]],
        colWidths=[doc.width*0.7, doc.width*0.3]
    )
    header_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), C_DARK),
        ("TOPPADDING",    (0,0),(-1,-1), 12),
        ("BOTTOMPADDING", (0,0),(-1,-1), 12),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("RIGHTPADDING",  (0,0),(-1,-1), 10),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(header_tbl)
    story.append(Spacer(1, 10))

    # ── KPIs ──
    kpi_data = [
        [hdr_cell("Total modelos"), hdr_cell("Nuevos"), hdr_cell("Se repiten"),
         hdr_cell("Persist. 3+"), hdr_cell("Consec. 3+"), hdr_cell("Alertas ALTA")],
        [body_cell(str(total_models),  TA_CENTER),
         body_cell(str(new_models),     TA_CENTER),
         body_cell(str(repeated_models),TA_CENTER),
         body_cell(str(persistent_total_count),      TA_CENTER),
         body_cell(str(persistent_consecutive_count),TA_CENTER),
         body_cell(str(len(alerts)),    TA_CENTER)],
    ]
    kpi_fills = [C_LIGHT, C_GREEN, C_AMBER, C_LIGHT, C_LIGHT, C_RED if alerts else C_GRAY]
    col_w = doc.width / 6
    kpi_tbl = Table(kpi_data, colWidths=[col_w]*6)
    kpi_style = [
        ("GRID",          (0,0),(-1,-1), 0.5, C_MID),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("ALIGN",         (0,0),(-1,-1), "CENTER"),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
    ]
    for col_i, fill in enumerate(kpi_fills):
        kpi_style.append(("BACKGROUND", (col_i,0),(col_i,0), C_MID))
        kpi_style.append(("BACKGROUND", (col_i,1),(col_i,1), fill))
    kpi_tbl.setStyle(TableStyle(kpi_style))
    story.append(kpi_tbl)
    story.append(Spacer(1, 12))

    # ── Alertas ALTA ──
    story.append(section_header("ALERTAS AUTOMÁTICAS — PRIORIDAD ALTA"))
    story.append(Spacer(1, 4))
    if alerts:
        alert_headers = [["MC","Familia","Veces","Primera aparición","Última aparición","Estado","Prioridad"]]
        alert_rows = [[
            body_cell(r["mc"]), body_cell(r["family"]),
            body_cell(str(r["count"]), TA_CENTER),
            body_cell(r["first_label"]), body_cell(r["last_label"]),
            body_cell(r["estado"]), body_cell(r["prioridad"], TA_CENTER),
        ] for r in alerts]
        alert_data = [[hdr_cell(h) for h in alert_headers[0]]] + alert_rows
        col_ws = [3.5*cm, 4*cm, 1.8*cm, 3.5*cm, 3.5*cm, 2.5*cm, 2.5*cm]
        atbl = Table(alert_data, colWidths=col_ws)
        atbl_style = [
            ("BACKGROUND",    (0,0),(-1,0), C_MID),
            ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#7F7F7F")),
            ("TOPPADDING",    (0,0),(-1,-1), 4),
            ("BOTTOMPADDING", (0,0),(-1,-1), 4),
            ("LEFTPADDING",   (0,0),(-1,-1), 4),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ]
        for i in range(1, len(alert_data)):
            atbl_style.append(("BACKGROUND", (0,i),(-1,i), C_RED))
        atbl.setStyle(TableStyle(atbl_style))
        story.append(atbl)
    else:
        story.append(Paragraph("No hay modelos con prioridad ALTA.", style_body))

    story.append(Spacer(1, 14))

    # ── Top 10 Reincidentes ──
    story.append(section_header("TOP 10 REINCIDENTES"))
    story.append(Spacer(1, 4))
    top_10 = sorted([r for r in rows if r["count"]>1], key=lambda x:(-x["count"],x["mc"]))[:10]
    if top_10:
        t10_data = [[hdr_cell(h) for h in ["MC","Familia","Veces","Ciclos","Estado","Prioridad"]]]
        for r in top_10:
            fill_row = C_RED if r["prioridad"]=="ALTA" else (C_AMBER if r["prioridad"]=="MEDIA" else C_GREEN)
            t10_data.append([
                body_cell(r["mc"]), body_cell(r["family"]),
                body_cell(str(r["count"]), TA_CENTER),
                body_cell(r["ciclos"]), body_cell(r["estado"]),
                body_cell(r["prioridad"], TA_CENTER),
            ])
        t10_col_ws = [3*cm, 3.5*cm, 1.8*cm, 9*cm, 2.5*cm, 2.5*cm]
        t10_tbl = Table(t10_data, colWidths=t10_col_ws)
        t10_style = [
            ("BACKGROUND",    (0,0),(-1,0), C_MID),
            ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#7F7F7F")),
            ("TOPPADDING",    (0,0),(-1,-1), 4),
            ("BOTTOMPADDING", (0,0),(-1,-1), 4),
            ("LEFTPADDING",   (0,0),(-1,-1), 4),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ]
        for i, r in enumerate(top_10, start=1):
            fill_row = C_RED if r["prioridad"]=="ALTA" else (C_AMBER if r["prioridad"]=="MEDIA" else C_GREEN)
            t10_style.append(("BACKGROUND",(0,i),(-1,i), fill_row))
        t10_tbl.setStyle(TableStyle(t10_style))
        story.append(t10_tbl)
    else:
        story.append(Paragraph("No hay modelos reincidentes.", style_body))

    story.append(Spacer(1, 14))

    # ── Familias ──
    story.append(section_header("FAMILIAS CON MÁS MERMAS (TOP 10)"))
    story.append(Spacer(1, 4))
    fam_data = [[hdr_cell(h) for h in ["Familia","Apariciones","Modelos únicos"]]]
    for i, fam in enumerate(family_rows[:10]):
        fill_row = C_LIGHT if i % 2 == 0 else C_WHITE
        fam_data.append([
            body_cell(fam["family"]),
            body_cell(str(fam["apariciones"]),    TA_CENTER),
            body_cell(str(fam["modelos_unicos"]), TA_CENTER),
        ])
    fam_tbl = Table(fam_data, colWidths=[8*cm, 3*cm, 3*cm])
    fam_style = [
        ("BACKGROUND",    (0,0),(-1,0), C_MID),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#7F7F7F")),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("LEFTPADDING",   (0,0),(-1,-1), 4),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
    ]
    for i in range(1, len(fam_data)):
        fam_style.append(("BACKGROUND",(0,i),(-1,i), C_LIGHT if i%2==1 else C_WHITE))
    fam_tbl.setStyle(TableStyle(fam_style))
    story.append(fam_tbl)

    # ── Detalle completo (página nueva) ──
    story.append(PageBreak())
    story.append(section_header("DETALLE COMPLETO"))
    story.append(Spacer(1, 4))

    det_hdrs = ["MC","Veces","Primera","Última","En último TOP","Estado","Prioridad","Familia","Pers. 3+","Consec. 3+","Racha"]
    det_data = [[hdr_cell(h) for h in det_hdrs]]
    det_col_ws = [3*cm,1.5*cm,3*cm,3*cm,2.2*cm,2.2*cm,2.2*cm,3*cm,1.8*cm,1.8*cm,1.5*cm]

    for r in rows:
        fill_row = C_RED if r["prioridad"]=="ALTA" else (C_AMBER if r["prioridad"]=="MEDIA" else C_GREEN if r["estado"]=="Nuevo" else C_GRAY)
        det_data.append([
            body_cell(r["mc"]),
            body_cell(str(r["count"]), TA_CENTER),
            body_cell(r["first_label"]),
            body_cell(r["last_label"]),
            body_cell(r["appears_in_last_top"], TA_CENTER),
            body_cell(r["estado"]),
            body_cell(r["prioridad"], TA_CENTER),
            body_cell(r["family"]),
            body_cell(r["persistencia_total"],       TA_CENTER),
            body_cell(r["persistencia_consecutiva"], TA_CENTER),
            body_cell(str(r["racha_max"]),           TA_CENTER),
        ])

    det_tbl = Table(det_data, colWidths=det_col_ws, repeatRows=1)
    det_style = [
        ("BACKGROUND",    (0,0),(-1,0), C_MID),
        ("GRID",          (0,0),(-1,-1), 0.3, colors.HexColor("#7F7F7F")),
        ("TOPPADDING",    (0,0),(-1,-1), 3),
        ("BOTTOMPADDING", (0,0),(-1,-1), 3),
        ("LEFTPADDING",   (0,0),(-1,-1), 3),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
    ]
    for i, r in enumerate(rows, start=1):
        fill_row = C_RED if r["prioridad"]=="ALTA" else (C_AMBER if r["prioridad"]=="MEDIA" else C_GREEN if r["estado"]=="Nuevo" else C_GRAY)
        det_style.append(("BACKGROUND",(0,i),(-1,i), fill_row))
    det_tbl.setStyle(TableStyle(det_style))
    story.append(det_tbl)

    doc.build(story)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
#  HELPERS INTERNOS COMPARTIDOS
# ─────────────────────────────────────────────

def _process_files(uploaded_files, records=None, order=None):
    if records is None: records = {}
    if order is None:   order = []
    last_seen_models, last_label = set(), None

    for f in uploaded_files:
        if not (f and f.filename and f.filename.lower().endswith(('.xlsx','.xlsm','.xls'))):
            continue
        try:
            label, models, families = _extract_top_models_and_label(f)
            order.append(label)
            last_label = label
            last_seen_models = models
            for mc in models:
                if mc not in records:
                    records[mc] = {"count":0,"labels":[],"first_label":label,"last_label":label,"family":families.get(mc,"")}
                rec = records[mc]
                rec["count"] += 1
                if label not in rec["labels"]:
                    rec["labels"].append(label)
                if not rec.get("family") and families.get(mc):
                    rec["family"] = families.get(mc)
                rec["last_label"] = label
        except Exception:
            continue

    return records, order, last_seen_models, last_label


# ─────────────────────────────────────────────
#  RUTAS
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/historico', methods=['POST'])
def generar_historico():
    uploaded_files = request.files.getlist('files')
    records, order, last_seen_models, last_label = _process_files(uploaded_files)
    if not records:
        return "No se han podido leer datos válidos de los ficheros subidos", 400
    result = _build_historico_rows(records, order, last_seen_models, last_label)
    output = _render_historico_visual(*result, order, sheet_title="HISTORICO")
    return send_file(output, as_attachment=True, download_name="historico.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route('/historico_pdf', methods=['POST'])
def generar_historico_pdf():
    uploaded_files = request.files.getlist('files')
    records, order, last_seen_models, last_label = _process_files(uploaded_files)
    if not records:
        return "No se han podido leer datos válidos de los ficheros subidos", 400
    result = _build_historico_rows(records, order, last_seen_models, last_label)
    output = _render_historico_pdf(*result, order, sheet_title="HISTORICO")
    return send_file(output, as_attachment=True, download_name="historico.pdf",
                     mimetype="application/pdf")


@app.route('/historico_acumulado', methods=['POST'])
def generar_historico_acumulado():
    master_file    = request.files.get('master')
    uploaded_files = request.files.getlist('files')

    records, order = {}, []
    if master_file and master_file.filename and master_file.filename.lower().endswith('.xlsx'):
        try:
            master_wb = load_workbook(BytesIO(master_file.read()), data_only=True)
            records, order = _seed_records_from_master_workbook(master_wb)
        except Exception:
            pass

    records, order, last_seen_models, last_label = _process_files(uploaded_files, records, order)
    if not records:
        return "No se han podido consolidar datos para el histórico acumulado", 400
    result = _build_historico_rows(records, order, last_seen_models, last_label)
    output = _render_historico_visual(*result, order, sheet_title="HISTORICO ACUMULADO")
    return send_file(output, as_attachment=True, download_name="historico_acumulado.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route('/historico_acumulado_pdf', methods=['POST'])
def generar_historico_acumulado_pdf():
    master_file    = request.files.get('master')
    uploaded_files = request.files.getlist('files')

    records, order = {}, []
    if master_file and master_file.filename and master_file.filename.lower().endswith('.xlsx'):
        try:
            master_wb = load_workbook(BytesIO(master_file.read()), data_only=True)
            records, order = _seed_records_from_master_workbook(master_wb)
        except Exception:
            pass

    records, order, last_seen_models, last_label = _process_files(uploaded_files, records, order)
    if not records:
        return "No se han podido consolidar datos para el histórico acumulado", 400
    result = _build_historico_rows(records, order, last_seen_models, last_label)
    output = _render_historico_pdf(*result, order, sheet_title="HISTORICO ACUMULADO")
    return send_file(output, as_attachment=True, download_name="historico_acumulado.pdf",
                     mimetype="application/pdf")


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
