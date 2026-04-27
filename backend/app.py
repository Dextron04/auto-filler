import os
import zipfile
import tempfile
from io import BytesIO
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from core_logic import (
    read_excel_mappings,
    read_excel_records,
    fill_document,
    safe_filename_part,
)

app = Flask(__name__)
# Allow CORS for development (Vite typically runs on 5173)
CORS(app)

app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB limit

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'healthy', 'message': 'Auto-Filler API is operational'}), 200

@app.route('/api/process', methods=['POST'])
def process_files():
    if 'excel' not in request.files or 'word' not in request.files:
        return jsonify({'error': 'Missing files (excel or word)'}), 400

    excel_file = request.files['excel']
    word_files = request.files.getlist('word')

    if not excel_file.filename or not word_files or not word_files[0].filename:
        return jsonify({'error': 'No files selected'}), 400

    try:
        # Create a temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = os.path.join(temp_dir, excel_file.filename)
            excel_file.save(excel_path)

            # Read mappings from Excel
            try:
                mappings, skipped = read_excel_mappings(excel_path)
            except Exception as e:
                return jsonify({'error': f'Excel processing error: {str(e)}'}), 400

            # If only one Word file is uploaded, return it directly
            if len(word_files) == 1:
                word_file = word_files[0]
                doc, total = fill_document(word_file, mappings)
                
                output = BytesIO()
                doc.save(output)
                output.seek(0)
                
                filename = f"{os.path.splitext(word_file.filename)[0]}_filled.docx"
                return send_file(
                    output,
                    mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    as_attachment=True,
                    download_name=filename
                )
            else:
                # Multiple files -> ZIP
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for word_file in word_files:
                        if not word_file.filename:
                            continue
                        doc, total = fill_document(word_file, mappings)
                        
                        file_buffer = BytesIO()
                        doc.save(file_buffer)
                        file_buffer.seek(0)
                        
                        filename = f"{os.path.splitext(word_file.filename)[0]}_filled.docx"
                        zf.writestr(filename, file_buffer.getvalue())
                
                zip_buffer.seek(0)
                return send_file(
                    zip_buffer,
                    mimetype='application/zip',
                    as_attachment=True,
                    download_name='filled_documents.zip'
                )

    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

@app.route('/api/bulk', methods=['POST'])
def bulk_process():
    """Bulk fill: one Excel (tabular export) + one Word template ->
    one filled docx per data row, returned as a ZIP.
    Each row is filled from a fresh copy of the template.
    """
    if 'excel' not in request.files or 'word' not in request.files:
        return jsonify({'error': 'Missing files (excel or word)'}), 400

    excel_file = request.files['excel']
    word_files = request.files.getlist('word')

    if not excel_file.filename or not word_files or not word_files[0].filename:
        return jsonify({'error': 'No files selected'}), 400

    if len(word_files) != 1:
        return jsonify({'error': 'Bulk mode expects exactly one Word template'}), 400

    word_file = word_files[0]
    template_bytes = word_file.read()
    if not template_bytes:
        return jsonify({'error': 'Word template is empty'}), 400

    sheet_name = request.form.get('sheet') or None

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = os.path.join(temp_dir, excel_file.filename)
            excel_file.save(excel_path)

            try:
                records, placeholder_columns, header_row_vals = read_excel_records(
                    excel_path, sheet_name=sheet_name
                )
            except Exception as e:
                return jsonify({'error': f'Excel processing error: {str(e)}'}), 400

            if not records:
                return jsonify({'error': 'No data rows found to fill'}), 400

            template_stem = os.path.splitext(word_file.filename)[0]

            # Pick a label column for filenames (Patient Name / dispute ID / first non-empty)
            label_col = None
            for candidate in ('Patient Name', 'PID', 'refid'):
                for i, h in enumerate(header_row_vals):
                    if h and str(h).strip().lower() == candidate.lower():
                        label_col = i
                        break
                if label_col is not None:
                    break

            # Find dispute ID column for a stable secondary tag
            dispute_col = None
            for i, h in enumerate(header_row_vals):
                if h and 'refid' in str(h).strip().lower():
                    dispute_col = i
                    break

            zip_buffer = BytesIO()
            used_names = set()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for idx, record in enumerate(records, start=1):
                    raw_row = record['row']
                    mappings = record['mappings']

                    # Fresh copy of template per row
                    doc, _ = fill_document(BytesIO(template_bytes), mappings)

                    label = None
                    if label_col is not None and label_col < len(raw_row):
                        label = raw_row[label_col]
                    dispute = None
                    if dispute_col is not None and dispute_col < len(raw_row):
                        dispute = raw_row[dispute_col]

                    parts = [template_stem]
                    if label:
                        parts.append(safe_filename_part(label))
                    if dispute:
                        parts.append(safe_filename_part(dispute))
                    if not label and not dispute:
                        parts.append(f"row{idx:03d}")

                    base = "_".join(parts)
                    name = f"{base}.docx"
                    n = 2
                    while name in used_names:
                        name = f"{base}_{n}.docx"
                        n += 1
                    used_names.add(name)

                    file_buffer = BytesIO()
                    doc.save(file_buffer)
                    file_buffer.seek(0)
                    zf.writestr(name, file_buffer.getvalue())

            zip_buffer.seek(0)
            zip_name = f"{template_stem}_bulk_filled.zip"
            return send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name=zip_name,
            )

    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

if __name__ == '__main__':
    # Running on 0.0.0.0 to make it accessible in all environments
    app.run(host='0.0.0.0', port=5001, debug=True)
