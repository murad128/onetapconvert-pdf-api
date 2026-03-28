import os, io, base64, json, tempfile, threading, time, urllib.request
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

ILOVEPDF_PUBLIC  = 'project_public_ed073f8c4b49df5133977adf95d2b164__ccI3705d8cb2e8f956d14fe45a1065075c0b'
ILOVEPDF_SECRET  = 'secret_key_bd6896753a236ed482e7314193927b10_mRf1h55f86e331e672c4db911f9b38a6f2145'

# ── iLovePDF convert ──────────────────────────────────────────────────────────
def convert_via_ilovepdf(pdf_bytes, file_name):
    import requests

    # 1. Authenticate
    auth = requests.post('https://api.ilovepdf.com/v1/auth',
        json={'public_key': ILOVEPDF_PUBLIC, 'secret_key': ILOVEPDF_SECRET},
        timeout=20)
    auth.raise_for_status()
    token = auth.json()['token']
    headers = {'Authorization': f'Bearer {token}'}

    # 2. Start task
    task = requests.get('https://api.ilovepdf.com/v1/start/pdftoxlsx', headers=headers, timeout=20)
    task.raise_for_status()
    task_data = task.json()
    server   = task_data['server']
    task_id  = task_data['task']

    # 3. Upload file
    upload = requests.post(
        f'https://{server}/v1/upload',
        headers=headers,
        data={'task': task_id},
        files={'file': (file_name, io.BytesIO(pdf_bytes), 'application/pdf')},
        timeout=60
    )
    upload.raise_for_status()
    server_filename = upload.json()['server_filename']

    # 4. Process
    process = requests.post(
        f'https://{server}/v1/process',
        headers=headers,
        json={
            'task': task_id,
            'tool': 'pdftoxlsx',
            'files': [{'server_filename': server_filename, 'filename': file_name}]
        },
        timeout=120
    )
    process.raise_for_status()

    # 5. Download result
    download = requests.get(
        f'https://{server}/v1/download/{task_id}',
        headers=headers,
        timeout=60
    )
    download.raise_for_status()
    return download.content

# ── Fallback: pdfplumber ──────────────────────────────────────────────────────
def detect_pdf_type(pdf_bytes):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_chars = sum(len((p.extract_text() or '').strip()) for p in pdf.pages[:3])
    return 'text' if total_chars > 80 else 'scanned'

def extract_with_pdfplumber(pdf_bytes):
    import pdfplumber
    all_tables = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table in (tables or []):
                clean = [[str(c or '').strip() for c in row] for row in table if any(c for c in row)]
                if len(clean) > 1:
                    all_tables.append({'data': clean, 'page': page_num})
            if not tables:
                words = page.extract_words(x_tolerance=3, y_tolerance=3)
                if not words:
                    continue
                rows_map = {}
                for w in words:
                    y_key = round(w['top'] / 5) * 5
                    rows_map.setdefault(y_key, []).append(w)
                all_x = [w['x0'] for w in words]
                col_centers = []
                for x in sorted(all_x):
                    found = next((c for c in col_centers if abs(c['center'] - x) < 20), None)
                    if found:
                        found['xs'].append(x); found['center'] = sum(found['xs']) / len(found['xs'])
                    else:
                        col_centers.append({'center': x, 'xs': [x]})
                col_centers.sort(key=lambda c: c['center'])
                cols = [c for c in col_centers if len(c['xs']) >= max(1, len(rows_map) // 4)] or col_centers
                result_rows = []
                for y_key in sorted(rows_map.keys()):
                    row_words = sorted(rows_map[y_key], key=lambda w: w['x0'])
                    row = [''] * len(cols)
                    for w in row_words:
                        best_col = min(range(len(cols)), key=lambda i: abs(cols[i]['center'] - w['x0']))
                        row[best_col] = (row[best_col] + ' ' + w['text']).strip()
                    if any(row):
                        result_rows.append(row)
                if result_rows:
                    all_tables.append({'data': result_rows, 'page': page_num, 'fallback': True})
    return all_tables

def tables_to_xlsx(all_tables):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook(); ws = wb.active; ws.title = 'Sheet1'
    hdr_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    alt_fill = PatternFill(start_color='EBF3FA', end_color='EBF3FA', fill_type='solid')
    hdr_font = Font(bold=True, color='FFFFFF', size=10)
    norm_font = Font(size=10)
    thin = Side(style='thin', color='CCCCCC')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    current_row = 1; col_max = {}
    for t_idx, tbl in enumerate(all_tables):
        data = tbl.get('data', [])
        if not data: continue
        if t_idx > 0: current_row += 1
        max_cols = max(len(r) for r in data)
        for r_idx, row in enumerate(data):
            is_hdr = r_idx == 0 and not tbl.get('fallback')
            for c_idx in range(max_cols):
                val = str(row[c_idx]).strip() if c_idx < len(row) else ''
                cell = ws.cell(row=current_row, column=c_idx+1, value=val)
                cell.border = bdr
                col_max[c_idx+1] = max(col_max.get(c_idx+1, 8), len(val))
                if is_hdr:
                    cell.fill = hdr_fill; cell.font = hdr_font
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                elif r_idx % 2 == 0:
                    cell.fill = alt_fill; cell.font = norm_font
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                else:
                    cell.font = norm_font
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            current_row += 1
    for col_idx, max_len in col_max.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 50)
    ws.freeze_panes = 'A2'
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# ── Main route ────────────────────────────────────────────────────────────────
@app.route('/convert', methods=['POST'])
def convert():
    try:
        data = request.get_json()
        file_b64 = data.get('fileBase64', '')
        file_name = data.get('fileName', 'input.pdf')
        if not file_b64:
            return jsonify({'error': 'No file provided'}), 400

        pdf_bytes = base64.b64decode(file_b64)
        out_name = file_name.rsplit('.', 1)[0] + '.xlsx'

        # Try iLovePDF first
        try:
            xlsx_bytes = convert_via_ilovepdf(pdf_bytes, file_name)
            xlsx_b64 = base64.b64encode(xlsx_bytes).decode('utf-8')
            return jsonify({'base64': xlsx_b64, 'fileName': out_name, 'method': 'ilovepdf', 'warning': None})
        except Exception as e:
            print(f'iLovePDF failed: {e}, falling back to pdfplumber')

        # Fallback: pdfplumber
        tables = extract_with_pdfplumber(pdf_bytes)
        if not tables:
            return jsonify({'error': 'Could not extract content from this PDF.'}), 422

        xlsx_bytes = tables_to_xlsx(tables)
        xlsx_b64 = base64.b64encode(xlsx_bytes).decode('utf-8')
        has_fallback = any(t.get('fallback') for t in tables)
        warning = 'No clear table structure detected. Text extracted in approximate layout.' if has_fallback else None
        return jsonify({'base64': xlsx_b64, 'fileName': out_name, 'method': 'pdfplumber', 'warning': warning})

    except Exception as e:
        import traceback
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

# ── Keep-alive ────────────────────────────────────────────────────────────────
def keep_alive():
    time.sleep(60)
    while True:
        try: urllib.request.urlopen('https://onetapconvert-pdf-api.onrender.com/health', timeout=10)
        except: pass
        time.sleep(600)

threading.Thread(target=keep_alive, daemon=True).start()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
