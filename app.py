import os, io, base64, threading, time, urllib.request
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ── PDF type detection ────────────────────────────────────────────────────────
def detect_pdf_type(pdf_bytes):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_chars = sum(len((p.extract_text() or '').strip()) for p in pdf.pages[:3])
    return 'text' if total_chars > 80 else 'scanned'

# ── pdfplumber extraction (primary) ──────────────────────────────────────────
def extract_with_pdfplumber(pdf_bytes):
    import pdfplumber
    all_tables = []  # list of (table_rows)

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in (tables or []):
                rows = []
                for row in table:
                    clean = [str(c or '').replace('\n', ' ').strip() for c in row]
                    non_empty = [c for c in clean if c]
                    if not non_empty:
                        continue
                    # Skip rows where all content is in a single cell (header/address blocks)
                    if len(non_empty) == 1 and len(clean) > 2:
                        continue
                    rows.append(clean)
                if rows:
                    all_tables.append(rows)

    if not all_tables:
        return []

    # Determine the dominant column count (most common across all tables)
    from collections import Counter
    col_counts = Counter(max(len(r) for r in t) for t in all_tables)
    dominant_cols = col_counts.most_common(1)[0][0]

    # Merge all tables that match dominant column count into one list
    merged = []
    for table in all_tables:
        # Normalize each row to dominant_cols
        for row in table:
            normalized = (row + [''] * dominant_cols)[:dominant_cols]
            merged.append(normalized)

    return merged

# ── XLSX export ───────────────────────────────────────────────────────────────
def rows_to_xlsx(all_rows):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    thin = Side(style='thin', color='CCCCCC')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    alt_fill = PatternFill(start_color='EBF3FA', end_color='EBF3FA', fill_type='solid')

    if not all_rows:
        return None

    # Trim trailing empty columns
    real_max = 0
    for row in all_rows:
        for i in range(len(row)-1, -1, -1):
            if row[i].strip():
                real_max = max(real_max, i+1)
                break
    max_cols = real_max or max(len(r) for r in all_rows)
    col_max = {}

    for r_idx, row in enumerate(all_rows):
        is_hdr = r_idx == 0
        for c_idx in range(max_cols):
            val = row[c_idx] if c_idx < len(row) else ''
            cell = ws.cell(row=r_idx + 1, column=c_idx + 1, value=val)
            cell.border = bdr
            cell.font = Font(bold=True, color='FFFFFF', size=10) if is_hdr else Font(size=10)
            cell.fill = hdr_fill if is_hdr else (alt_fill if r_idx % 2 == 0 else PatternFill())
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
            col_max[c_idx + 1] = max(col_max.get(c_idx + 1, 8), len(val))

    for ci, ml in col_max.items():
        ws.column_dimensions[get_column_letter(ci)].width = min(ml + 3, 50)
    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
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
        pdf_type = detect_pdf_type(pdf_bytes)
        out_name = file_name.rsplit('.', 1)[0] + '.xlsx'
        warning = None

        if pdf_type == 'text':
            rows = extract_with_pdfplumber(pdf_bytes)
            if not rows:
                return jsonify({'error': 'No content could be extracted from this PDF.'}), 422
            xlsx_bytes = rows_to_xlsx(rows)
            method = 'pdfplumber'
        else:
            # Scanned PDF — warn user, try pdfplumber anyway
            rows = extract_with_pdfplumber(pdf_bytes)
            warning = 'Scanned PDF detected. Results may be lower accuracy.'
            if not rows:
                return jsonify({'error': 'No content could be extracted. This appears to be a scanned image PDF.'}), 422
            xlsx_bytes = rows_to_xlsx(rows)
            method = 'pdfplumber-scanned'

        if not xlsx_bytes:
            return jsonify({'error': 'Failed to generate Excel file.'}), 500

        xlsx_b64 = base64.b64encode(xlsx_bytes).decode('utf-8')
        return jsonify({
            'base64': xlsx_b64,
            'fileName': out_name,
            'pdfType': pdf_type,
            'method': method,
            'warning': warning
        })

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

# ── Keep-alive ────────────────────────────────────────────────────────────────
def keep_alive():
    time.sleep(60)
    while True:
        try:
            host = os.environ.get('RAILWAY_PUBLIC_DOMAIN') or 'onetapconvert-pdf-api.onrender.com'
            urllib.request.urlopen(f'https://{host}/health', timeout=10)
        except:
            pass
        time.sleep(600)

threading.Thread(target=keep_alive, daemon=True).start()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
