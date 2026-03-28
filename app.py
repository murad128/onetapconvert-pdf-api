import os, io, base64, json, tempfile
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def detect_pdf_type(pdf_bytes):
    """Returns 'text' or 'scanned'."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_chars = sum(len((p.extract_text() or '').strip()) for p in pdf.pages[:3])
    return 'text' if total_chars > 80 else 'scanned'

def extract_with_camelot(pdf_path):
    """Try Camelot lattice (bordered tables) then stream (borderless)."""
    import camelot
    all_tables = []

    # Try lattice mode first (best for invoices/price lists with borders)
    try:
        tables = camelot.read_pdf(pdf_path, pages='all', flavor='lattice')
        if tables and tables.n > 0:
            for t in tables:
                df = t.df
                # Remove empty rows/cols
                df = df.replace('', None).dropna(how='all').fillna('')
                data = df.values.tolist()
                if len(data) > 1:
                    all_tables.append({'data': data, 'method': 'camelot-lattice', 'accuracy': t.parsing_report.get('accuracy', 0)})
    except Exception:
        pass

    # If lattice found nothing, try stream mode
    if not all_tables:
        try:
            tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream',
                                       edge_tol=50, row_tol=10, column_tol=0)
            if tables and tables.n > 0:
                for t in tables:
                    df = t.df
                    df = df.replace('', None).dropna(how='all').fillna('')
                    data = df.values.tolist()
                    if len(data) > 1:
                        all_tables.append({'data': data, 'method': 'camelot-stream', 'accuracy': t.parsing_report.get('accuracy', 0)})
        except Exception:
            pass

    return all_tables

def extract_with_tabula(pdf_path):
    """Tabula fallback."""
    import tabula
    all_tables = []
    try:
        dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True,
                               lattice=True, stream=False)
        if not dfs:
            dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True,
                                   lattice=False, stream=True)
        for df in dfs:
            df = df.fillna('')
            data = [list(df.columns)] + df.values.tolist()
            data = [[str(c) for c in row] for row in data]
            if len(data) > 1:
                all_tables.append({'data': data, 'method': 'tabula'})
    except Exception:
        pass
    return all_tables

def extract_with_pdfplumber(pdf_bytes):
    """pdfplumber fallback with improved column detection."""
    import pdfplumber
    all_tables = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # Try structured table first
            tables = page.extract_tables({
                'vertical_strategy': 'lines_strict',
                'horizontal_strategy': 'lines_strict',
                'intersection_tolerance': 5,
            })
            if not tables:
                tables = page.extract_tables({
                    'vertical_strategy': 'text',
                    'horizontal_strategy': 'text',
                    'min_words_vertical': 3,
                    'min_words_horizontal': 1,
                })

            for table in (tables or []):
                clean = [[str(c or '').strip() for c in row] for row in table if any(c for c in row)]
                if len(clean) > 1:
                    all_tables.append({'data': clean, 'method': 'pdfplumber', 'page': page_num})

            # If still nothing, use text with X-position column clustering
            if not tables:
                words = page.extract_words(x_tolerance=3, y_tolerance=3)
                if not words:
                    continue
                # cluster rows by Y
                rows_map = {}
                for w in words:
                    y_key = round(w['top'] / 5) * 5
                    rows_map.setdefault(y_key, []).append(w)

                sorted_ys = sorted(rows_map.keys())
                all_x = [w['x0'] for w in words]
                all_x.sort()

                # cluster columns by X
                col_centers = []
                for x in all_x:
                    found = next((c for c in col_centers if abs(c['center'] - x) < 20), None)
                    if found:
                        found['xs'].append(x)
                        found['center'] = sum(found['xs']) / len(found['xs'])
                    else:
                        col_centers.append({'center': x, 'xs': [x]})
                col_centers.sort(key=lambda c: c['center'])
                min_seen = max(1, len(sorted_ys) // 4)
                cols = [c for c in col_centers if len(c['xs']) >= min_seen]
                if not cols:
                    cols = col_centers

                result_rows = []
                for y_key in sorted_ys:
                    row_words = sorted(rows_map[y_key], key=lambda w: w['x0'])
                    row = [''] * len(cols)
                    for w in row_words:
                        best_col = min(range(len(cols)), key=lambda i: abs(cols[i]['center'] - w['x0']))
                        row[best_col] = (row[best_col] + ' ' + w['text']).strip()
                    if any(row):
                        result_rows.append(row)

                if result_rows:
                    all_tables.append({'data': result_rows, 'method': 'pdfplumber-text', 'page': page_num, 'fallback': True})

    return all_tables

def extract_with_ocr(pdf_bytes):
    """OCR for scanned PDFs."""
    from pdf2image import convert_from_bytes
    import pytesseract
    import re

    pages = convert_from_bytes(pdf_bytes, dpi=300)
    all_tables = []

    for page_num, img in enumerate(pages, 1):
        # Get word-level data with bounding boxes
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT, config='--psm 6')
        words = []
        for i, txt in enumerate(data['text']):
            if txt.strip() and int(data['conf'][i]) > 30:
                words.append({
                    'text': txt.strip(),
                    'x': data['left'][i],
                    'y': data['top'][i],
                    'h': data['height'][i]
                })
        if not words:
            continue

        # cluster rows
        rows_map = {}
        for w in words:
            y_key = round(w['y'] / 10) * 10
            rows_map.setdefault(y_key, []).append(w)

        # cluster columns
        all_x = [w['x'] for w in words]
        col_centers = []
        for x in sorted(all_x):
            found = next((c for c in col_centers if abs(c['center'] - x) < 30), None)
            if found:
                found['xs'].append(x)
                found['center'] = sum(found['xs']) / len(found['xs'])
            else:
                col_centers.append({'center': x, 'xs': [x]})
        col_centers.sort(key=lambda c: c['center'])
        min_seen = max(1, len(rows_map) // 5)
        cols = [c for c in col_centers if len(c['xs']) >= min_seen] or col_centers

        result_rows = []
        for y_key in sorted(rows_map.keys()):
            row_words = sorted(rows_map[y_key], key=lambda w: w['x'])
            row = [''] * len(cols)
            for w in row_words:
                best_col = min(range(len(cols)), key=lambda i: abs(cols[i]['center'] - w['x']))
                row[best_col] = (row[best_col] + ' ' + w['text']).strip()
            if any(row):
                result_rows.append(row)

        if result_rows:
            all_tables.append({'data': result_rows, 'method': 'ocr', 'page': page_num})

    return all_tables

def tables_to_xlsx(all_tables):
    """Convert tables to styled XLSX."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    hdr_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    alt_fill = PatternFill(start_color='EBF3FA', end_color='EBF3FA', fill_type='solid')
    hdr_font = Font(bold=True, color='FFFFFF', size=10)
    norm_font = Font(size=10)
    thin = Side(style='thin', color='CCCCCC')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    c_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    l_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    current_row = 1
    col_max_len = {}

    for t_idx, tbl in enumerate(all_tables):
        data = tbl.get('data', [])
        is_fallback = tbl.get('fallback', False)
        if not data:
            continue

        if t_idx > 0:
            current_row += 1  # blank row between tables

        max_cols = max(len(row) for row in data)

        for r_idx, row in enumerate(data):
            is_header = r_idx == 0 and not is_fallback
            for c_idx in range(max_cols):
                val = str(row[c_idx]).strip() if c_idx < len(row) else ''
                cell = ws.cell(row=current_row, column=c_idx + 1, value=val)
                cell.border = bdr
                col_key = c_idx + 1
                col_max_len[col_key] = max(col_max_len.get(col_key, 8), len(val))
                if is_header:
                    cell.fill = hdr_fill
                    cell.font = hdr_font
                    cell.alignment = c_align
                elif r_idx % 2 == 0:
                    cell.fill = alt_fill
                    cell.font = norm_font
                    cell.alignment = l_align
                else:
                    cell.font = norm_font
                    cell.alignment = l_align
            current_row += 1

    # Set column widths
    for col_idx, max_len in col_max_len.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 50)

    ws.freeze_panes = 'A2'
    ws.row_dimensions[1].height = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

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
        tables = []
        method_used = ''
        warning = None

        if pdf_type == 'text':
            # Write to temp file for Camelot/Tabula
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                tmp.write(pdf_bytes)
                tmp_path = tmp.name

            # Try Camelot first
            tables = extract_with_camelot(tmp_path)
            if tables:
                method_used = tables[0].get('method', 'camelot')
                # Check accuracy
                low_acc = [t for t in tables if t.get('accuracy', 100) < 80]
                if low_acc:
                    warning = 'Some tables extracted with lower accuracy. Review the output.'
            else:
                # Try Tabula
                tables = extract_with_tabula(tmp_path)
                if tables:
                    method_used = 'tabula'
                else:
                    # pdfplumber fallback
                    tables = extract_with_pdfplumber(pdf_bytes)
                    method_used = 'pdfplumber'
                    if tables and any(t.get('fallback') for t in tables):
                        warning = 'No clear table structure detected. Text extracted in approximate column layout.'

            try:
                os.unlink(tmp_path)
            except Exception:
                pass
        else:
            # Scanned PDF — OCR
            tables = extract_with_ocr(pdf_bytes)
            method_used = 'ocr'
            warning = 'Scanned PDF detected. OCR used — results may vary by image quality.'

        if not tables:
            return jsonify({'error': 'No content could be extracted from this PDF. The file may be empty, encrypted, or unsupported.'}), 422

        xlsx_bytes = tables_to_xlsx(tables)
        xlsx_b64 = base64.b64encode(xlsx_bytes).decode('utf-8')
        out_name = file_name.rsplit('.', 1)[0] + '.xlsx'

        return jsonify({
            'base64': xlsx_b64,
            'fileName': out_name,
            'pdfType': pdf_type,
            'method': method_used,
            'warning': warning
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()[-500:]}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
