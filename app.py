from flask import Flask, request, send_file, render_template_string, redirect, url_for, flash
from openpyxl import load_workbook, Workbook
from copy import copy
import tempfile, os, re

app = Flask(__name__)
app.secret_key = "top5_secret_key"

HTML = """
<!doctype html>
<html lang='es'>
<head><meta charset='utf-8'><title>TOP MERMAS - Generador</title></head>
<body style="font-family:Arial,Helvetica,sans-serif;max-width:900px;margin:20px;">
  <h2>Generador TOP (sube solo el fichero MERMAS)</h2>
  <p>Sube únicamente el fichero MERMAS (.xlsx). Las plantillas ya están integradas en la app.</p>
  <form action="/generate" method="post" enctype="multipart/form-data">
    <label>Fichero MERMAS (xlsx): <input type="file" name="mermas" accept=".xls,.xlsx,.xlsm" required></label><br><br>
    <button type="submit">Generar Excel</button>
    </form>
    
    <br><br>
<form action="/historico" method="get">
  <button type="submit">Generar histórico</button>
</form>
  <p style="color:gray;font-size:0.9em">Si necesitas subir tus propias plantillas, contacta para activar la opción.</p>
</body>
</html>
"""

def copy_sheet_exact(src_ws, tgt_wb, title):
    # creates a sheet in tgt_wb and attempts to copy styles, merges, dims.
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
    max_row = src_ws.max_row
    max_col = src_ws.max_column
    for r in range(1, max_row+1):
        for c in range(1, max_col+1):
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
                if s.comment:
                    t.comment = copy(s.comment)
            except Exception:
                pass
    try:
        tgt.sheet_view = copy(src_ws.sheet_view)
    except Exception:
        pass
    try:
        tgt.page_setup = copy(src_ws.page_setup)
    except Exception:
        pass
    return tgt

def extract_sem_ciclo_from_name(fname):
    sem = None; cic = None
    m = re.search(r"[sS][eE][mM][^0-9]*([0-9]{1,2})", fname)
    if m:
        sem = m.group(1)
    m2 = re.search(r"[cC]\s*[_\-]?\s*([0-9]{1,2})", fname)
    if m2:
        cic = m2.group(1)
    # fallback patterns
    if not sem:
        m3 = re.search(r"\bS[_\s\-]?([0-9]{1,2})\b", fname, flags=re.IGNORECASE)
        if m3: sem = m3.group(1)
    if not cic:
        m4 = re.search(r"\bC[_\s\-]?([0-9]{1,2})\b", fname, flags=re.IGNORECASE)
        if m4: cic = m4.group(1)
    return sem, cic

def sanitize_sheet_name(name):
    # remove invalid chars and trim to 31 chars
    s = re.sub(r'[:\\/*?\[\]]', '-', name)
    s = s.replace("\n"," ").strip()
    return s[:31] if len(s)>31 else s

def convert_pct_cell_to_number(cell):
    # If value is 0<x<1 treat as fraction and multiply by 100.
    v = cell.value
    if v is None:
        return
    try:
        if isinstance(v, (int, float)) and abs(v) < 1.0:
            cell.value = round(float(v)*100.0, 6)
        elif isinstance(v, str) and '%' in v:
            s = v.replace('%','').strip()
            try:
                cell.value = float(s)
            except:
                pass
        # set number format to 2 decimals
        try:
            cell.number_format = '0.00'
        except:
            pass
    except Exception:
        pass

@app.route('/', methods=['GET'])
def index():
    return render_template_string(HTML)

@app.route('/generate', methods=['POST'])
def generate():
    if 'mermas' not in request.files:
        flash('Falta fichero MERMAS'); return redirect(url_for('index'))
    m = request.files['mermas']
    original_filename = m.filename or ''
    tmpdir = tempfile.mkdtemp()
    mer_path = os.path.join(tmpdir, 'mermas.xlsx')
    m.save(mer_path)
    # templates bundled with app
    tpl_art = os.path.join(os.path.dirname(__file__), 'PLANTILLA Artículos.xlsx')
    tpl_chk = os.path.join(os.path.dirname(__file__), 'CHECKLIST CALIDAD DE REPARTO.xlsx')
    if not os.path.exists(tpl_art) or not os.path.exists(tpl_chk):
        return "Faltan plantillas en la app. Contacta para añadirlas.", 500
    out_path = os.path.join(tmpdir, 'TOP_GENERADO.xlsx')
    try:
        generate_from_mermas(mer_path, tpl_art, tpl_chk, out_path, original_filename)
    except Exception as e:
        return f"Error durante generación: {e}", 500
    return send_file(out_path, as_attachment=True, download_name='TOP_GENERADO.xlsx')

def generate_from_mermas(mermas_path, tpl_art_path, tpl_chk_path, out_path, original_filename=''):
    wb_mermas = load_workbook(mermas_path, data_only=True)
    # detect top sheet
    top_name = None
    for name in wb_mermas.sheetnames:
        if 'top' in name.lower() or 'calidad' in name.lower():
            top_name = name; break
    if not top_name:
        top_name = wb_mermas.sheetnames[0]
    ws_top = wb_mermas[top_name]

    # extract sem/ciclo from the uploaded file name (the real input name)
    sem, cic = extract_sem_ciclo_from_name(original_filename)

    def parse_week_value(value):
        if value is None:
            return None
        if isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            return int(value) if float(value).is_integer() else value
        m = re.search(r'\d+', str(value))
        if m:
            return int(m.group())
        return str(value).strip()

    top_week = parse_week_value(sem)
    if top_week is None or str(top_week).strip() == '':
        raise ValueError("No se pudo extraer el número de semana del nombre del fichero importado (TOP MERMAS Sx Cx).")

    # detect mc and familia columns by header row
    headers = [ (ws_top.cell(row=1, column=c).value or "") for c in range(1, ws_top.max_column+1) ]
    low = [str(h).lower() for h in headers]
    def find_idx(keys):
        for k in keys:
            for i,h in enumerate(low):
                if k in h:
                    return i+1
        return None
    mc_idx = find_idx(['mc','modelo','model','codigo','articulo']) or 1
    fam_idx = find_idx(['familia','family','grupo'])

    # mapping columns to A48:H48 (map_cols = [2,3,4,5,6,7,9,10])
    map_cols = [2,3,4,5,6,7,9,10]

    # build mc map from top sheet
    mc_map = {}
    for r in range(2, ws_top.max_row+1):
        val = ws_top.cell(row=r, column=mc_idx).value
        if val is None: continue
        key = str(val).strip()
        mapped = []
        for col in map_cols:
            mapped.append(ws_top.cell(row=r, column=col).value if col <= ws_top.max_column else None)
        fam = ws_top.cell(row=r, column=fam_idx).value if fam_idx and fam_idx <= ws_top.max_column else None
        mc_map[key] = {'mapped': mapped, 'fam': fam, 'row': r}

    wb_art = load_workbook(tpl_art_path, data_only=False)
    wb_chk = load_workbook(tpl_chk_path, data_only=False)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    # copy TOP sheet exact
    copy_sheet_exact(ws_top, wb_out, "TOP-CALIDAD-2")

    tpl_art_sheet = wb_art[wb_art.sheetnames[0]]
    tpl_chk_sheet = wb_chk[wb_chk.sheetnames[0]]

    existing = set(wb_out.sheetnames)
    pairs = []

    for key in mc_map:
        safe = sanitize_sheet_name(key)
        base = safe
        i = 1
        while safe in existing:
            suf = f'-{i}'; safe = base[:31-len(suf)] + suf; i += 1
        existing.add(safe)
        # copy article template
        new_mc = copy_sheet_exact(tpl_art_sheet, wb_out, safe)
        # fill MC info
        new_mc.cell(row=4, column=2).value = key
        if mc_map[key]['fam'] is not None:
            new_mc.cell(row=5, column=2).value = mc_map[key]['fam']
        # B2 = numeric week extracted from TOP-CALIDAD-2!B2, D2 = cycle from filename when available
        new_mc.cell(row=2, column=2).value = top_week
        try:
            new_mc.cell(row=2, column=4).value = int(cic) if cic!='' else None
        except:
            new_mc.cell(row=2, column=4).value = cic
        # fill A48:H48 from mapped columns
        for idx, val in enumerate(mc_map[key]['mapped'], start=1):
            new_mc.cell(row=48, column=idx).value = val
        new_mc.cell(row=48, column=9).value = None
        # convert F48 and H48 to numeric percent without sign and format 2 decimals
        for col in (6,8):
            cell = new_mc.cell(row=48, column=col)
            convert_pct_cell_to_number(cell)

        # create checklist sheet
        chk_name = f'CHECKLIST-{safe}'
        chk = copy_sheet_exact(tpl_chk_sheet, wb_out, chk_name)
        # update A1: insert "Artículo: <model>" after colon
        a1 = chk.cell(row=1, column=1).value or ""
        # try to preserve existing text and replace after 'articulo:'
        if re.search(r"art[ií]culo\s*:", str(a1), flags=re.IGNORECASE):
            new_a1 = re.sub(r"(art[ií]culo\s*:\s*).*", lambda m: m.group(1) + str(key), str(a1), flags=re.IGNORECASE)
        else:
            new_a1 = str(a1).strip() + " Artículo: " + str(key)
        chk.cell(row=1, column=1).value = new_a1
        # update D1: Semana and Ciclo
        chk.cell(row=1, column=4).value = f"SEMANA: {sem}   CICLO: {cic}"
        # clear B4/B5 and A48:H48
        chk.cell(row=4, column=2).value = None
        chk.cell(row=5, column=2).value = None
        for c in range(1,9):
            chk.cell(row=48, column=c).value = None

        pairs.append((safe, chk_name))

    # order sheets TOP, MC, CHECKLIST-MC ...
    ordered = ['TOP-CALIDAD-2']
    for a,b in pairs:
        ordered.append(a); ordered.append(b)
    # append any other sheets (unlikely)
    for s in wb_out.sheetnames:
        if s not in ordered:
            ordered.append(s)
    wb_out._sheets = [wb_out[s] for s in ordered]
    wb_out.save(out_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
