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
        # Fallback: extract all text as single-column rows
        fallback_rows = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ''
                for line in text.splitlines():
                    line = line.strip()
                    if line:
                        fallback_rows.append([line])
        return [fallback_rows] if fallback_rows else []

    # Normalize each table's rows to its own max col count
    result = []
    for table in all_tables:
        max_cols = max(len(r) for r in table)
        normalized = [(r + [''] * max_cols)[:max_cols] for r in table]
        result.append(normalized)

    return result

# ── XLSX export ───────────────────────────────────────────────────────────────
def write_sheet(ws, rows):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    thin = Side(style='thin', color='CCCCCC')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    alt_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    no_fill = PatternFill()

    max_cols = max(len(r) for r in rows) if rows else 0
    col_max = {}

    # Detect key-value table (2 cols, no obvious header)
    is_kv_table = max_cols == 2

    # Detect if first row is a real header (all caps or common header words)
    def is_header_row(row):
        vals = [str(v).strip() for v in row if str(v).strip()]
        if not vals: return False
        header_words = {'no', 'oem', 'fmsi', 'description', 'qty', 'price', 'amount', 'item', '#'}
        return any(v.lower() in header_words or v.isupper() for v in vals)

    has_header = not is_kv_table and is_header_row(rows[0]) if rows else False

    for r_idx, row in enumerate(rows):
        is_hdr = has_header and r_idx == 0
        for c_idx in range(max_cols):
            val = row[c_idx] if c_idx < len(row) else ''
            cell = ws.cell(row=r_idx + 1, column=c_idx + 1, value=val)
            cell.border = bdr
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
            col_max[c_idx + 1] = max(col_max.get(c_idx + 1, 8), len(str(val)))

            if is_hdr:
                cell.font = Font(bold=True, color='FFFFFF', size=10)
                cell.fill = hdr_fill
            elif is_kv_table:
                # Key-value: left col normal, right col bold
                cell.font = Font(bold=(c_idx == 1), size=10)
                cell.fill = alt_fill if r_idx % 2 == 0 else no_fill
            else:
                cell.font = Font(size=10)
                cell.fill = alt_fill if r_idx % 2 == 0 else no_fill

    for ci, ml in col_max.items():
        ws.column_dimensions[get_column_letter(ci)].width = min(ml + 3, 50)

    if has_header and rows:
        ws.freeze_panes = 'A2'

def tables_to_xlsx(all_tables):
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    for i, rows in enumerate(all_tables, 1):
        ws = wb.create_sheet(title=f'Table {i}')
        write_sheet(ws, rows)

    if not wb.sheetnames:
        return None

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

def rows_to_xlsx(all_rows):
    """Single-sheet export (kept for compatibility)."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    write_sheet(ws, all_rows)
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
            all_tables = extract_with_pdfplumber(pdf_bytes)
            if not all_tables:
                return jsonify({'error': 'No content could be extracted from this PDF.'}), 422
            xlsx_bytes = tables_to_xlsx(all_tables)
            method = 'pdfplumber'
        else:
            all_tables = extract_with_pdfplumber(pdf_bytes)
            warning = 'Scanned PDF detected. Results may be lower accuracy.'
            if not all_tables:
                return jsonify({'error': 'No content could be extracted. This appears to be a scanned image PDF.'}), 422
            xlsx_bytes = tables_to_xlsx(all_tables)
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
