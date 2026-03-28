import os, io, base64, json
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def detect_pdf_type(pdf_bytes):
    """Returns 'text' or 'scanned' based on extractable text content."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_chars = 0
        for page in pdf.pages[:3]:  # check first 3 pages
            text = page.extract_text() or ''
            total_chars += len(text.strip())
    return 'text' if total_chars > 50 else 'scanned'

def extract_tables_text_pdf(pdf_bytes):
    """Use pdfplumber to extract tables from text-based PDF."""
    import pdfplumber
    all_tables = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    all_tables.append({'page': page_num, 'data': table})
            else:
                # fallback: extract text as single-column rows
                text = page.extract_text() or ''
                if text.strip():
                    rows = [[line] for line in text.splitlines() if line.strip()]
                    all_tables.append({'page': page_num, 'data': rows, 'fallback': True})
    return all_tables

def extract_tables_scanned_pdf(pdf_bytes):
    """Use OCR (Tesseract) to extract text, then parse table-like structure."""
    from pdf2image import convert_from_bytes
    import pytesseract
    import re

    pages = convert_from_bytes(pdf_bytes, dpi=300)
    all_tables = []

    for page_num, img in enumerate(pages, 1):
        text = pytesseract.image_to_string(img, config='--psm 6')
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            continue

        # Detect columns by consistent whitespace gaps
        rows = []
        for line in lines:
            # Split by 2+ spaces (column separator heuristic)
            cells = re.split(r'  +', line)
            cells = [c.strip() for c in cells if c.strip()]
            if cells:
                rows.append(cells)

        if rows:
            all_tables.append({'page': page_num, 'data': rows, 'ocr': True})

    return all_tables

def tables_to_xlsx(all_tables, pdf_type):
    """Convert extracted tables to styled XLSX."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    alt_fill = PatternFill(start_color='EBF3FA', end_color='EBF3FA', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    normal_font = Font(size=10)
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    current_row = 1

    for t_idx, table_info in enumerate(all_tables):
        data = table_info.get('data', [])
        page = table_info.get('page', 1)
        is_fallback = table_info.get('fallback', False)
        is_ocr = table_info.get('ocr', False)

        if not data:
            continue

        # Section header
        if len(all_tables) > 1:
            ws.cell(row=current_row, column=1, value=f'Page {page}')
            ws.cell(row=current_row, column=1).font = Font(bold=True, size=10, color='666666')
            current_row += 1

        max_cols = max((len(row) for row in data), default=1)

        for r_idx, row in enumerate(data):
            if row is None:
                current_row += 1
                continue

            is_header_row = (r_idx == 0 and not is_fallback and not is_ocr)

            for c_idx in range(max_cols):
                cell_val = row[c_idx] if c_idx < len(row) else ''
                if cell_val is None:
                    cell_val = ''

                cell = ws.cell(row=current_row, column=c_idx + 1, value=str(cell_val).strip())
                cell.border = border

                if is_header_row:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center_align
                elif r_idx % 2 == 0:
                    cell.fill = alt_fill
                    cell.font = normal_font
                    cell.alignment = left_align
                else:
                    cell.font = normal_font
                    cell.alignment = left_align

            current_row += 1

        # Auto column widths
        for col_idx in range(1, max_cols + 1):
            max_len = 10
            col_letter = get_column_letter(col_idx)
            for row_data in data:
                if col_idx <= len(row_data) and row_data[col_idx-1]:
                    max_len = max(max_len, len(str(row_data[col_idx-1])))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

        current_row += 1  # blank row between tables

    # Freeze top row
    ws.freeze_panes = 'A2'

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

        # Detect PDF type
        pdf_type = detect_pdf_type(pdf_bytes)

        # Extract tables
        warning = None
        if pdf_type == 'text':
            tables = extract_tables_text_pdf(pdf_bytes)
            has_fallback = any(t.get('fallback') for t in tables)
            if has_fallback:
                warning = 'Some pages had no tables — text was extracted as rows.'
        else:
            tables = extract_tables_scanned_pdf(pdf_bytes)
            warning = 'Scanned PDF detected — OCR used. Results may vary.'

        if not tables:
            return jsonify({'error': 'No content could be extracted from this PDF.'}), 422

        # Generate XLSX
        xlsx_bytes = tables_to_xlsx(tables, pdf_type)
        xlsx_b64 = base64.b64encode(xlsx_bytes).decode('utf-8')

        out_name = file_name.rsplit('.', 1)[0] + '.xlsx'
        return jsonify({
            'base64': xlsx_b64,
            'fileName': out_name,
            'pdfType': pdf_type,
            'warning': warning
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
