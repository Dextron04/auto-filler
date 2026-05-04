import os
import zipfile
import tempfile
from io import BytesIO
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from core_logic import (
    read_excel_mappings,
    read_excel_records,
    read_excel_records_column_oriented,
    is_scs_procedure,
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

@app.route('/api/bulk-multi', methods=['POST'])
def bulk_multi_process():
    """Multi-template bulk fill. One Excel (column-oriented "Fields to Replace"
    sheet) + two Word templates (SCS + default). Per record, pick template by
    detecting Spinal Cord Stimulator procedures, fill, and ZIP all docs.
    """
    if 'excel' not in request.files:
        return jsonify({'error': 'Missing excel file'}), 400
    if 'word_scs' not in request.files or 'word_default' not in request.files:
        return jsonify({'error': 'Missing template (need word_scs and word_default)'}), 400

    excel_file = request.files['excel']
    scs_file = request.files['word_scs']
    default_file = request.files['word_default']

    if not excel_file.filename or not scs_file.filename or not default_file.filename:
        return jsonify({'error': 'No files selected'}), 400

    scs_bytes = scs_file.read()
    default_bytes = default_file.read()
    if not scs_bytes or not default_bytes:
        return jsonify({'error': 'Word template is empty'}), 400

    sheet_name = request.form.get('sheet') or None

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = os.path.join(temp_dir, excel_file.filename)
            excel_file.save(excel_path)

            try:
                records = read_excel_records_column_oriented(
                    excel_path, sheet_name=sheet_name
                )
            except Exception as e:
                return jsonify({'error': f'Excel processing error: {str(e)}'}), 400

            if not records:
                return jsonify({'error': 'No data records found to fill'}), 400

            scs_stem = os.path.splitext(scs_file.filename)[0]
            default_stem = os.path.splitext(default_file.filename)[0]

            zip_buffer = BytesIO()
            used_names = set()
            scs_count = 0
            default_count = 0
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for idx, record in enumerate(records, start=1):
                    is_scs = is_scs_procedure(record.get('procedure'))
                    template_bytes = scs_bytes if is_scs else default_bytes
                    template_stem = scs_stem if is_scs else default_stem
                    if is_scs:
                        scs_count += 1
                    else:
                        default_count += 1

                    doc, _ = fill_document(BytesIO(template_bytes), record['mappings'])

                    parts = [template_stem]
                    if record.get('patient_name'):
                        parts.append(safe_filename_part(record['patient_name']))
                    if record.get('dispute_id'):
                        parts.append(safe_filename_part(record['dispute_id']))
                    if not record.get('patient_name') and not record.get('dispute_id'):
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
            response = send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name='multi_template_bulk_filled.zip',
            )
            response.headers['X-SCS-Count'] = str(scs_count)
            response.headers['X-Default-Count'] = str(default_count)
            response.headers['Access-Control-Expose-Headers'] = 'X-SCS-Count, X-Default-Count, Content-Disposition'
            return response

    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

def detect_ps_template_key(filename: str) -> str:
    """Returns a routing key from a PS template filename.
    Keys: '1', '2', '3', 'no_pain', 'no'
    """
    name = filename.lower()
    has_no = 'no carrier' in name or 'no_carrier' in name
    if has_no:
        return 'no_pain' if 'pain' in name else 'no'
    for n in ('1', '2', '3'):
        if f'{n} carrier' in name or f'{n}_carrier' in name:
            return n
    return 'no'


def select_ps_template(num_comps, procedure_type, templates: dict):
    """Pick template bytes for a record.
    templates: dict[key -> bytes], keys from detect_ps_template_key
    """
    if num_comps is None or num_comps == 0:
        if procedure_type and 'pain' in procedure_type.lower():
            return templates.get('no_pain') or templates.get('no')
        return templates.get('no') or templates.get('no_pain')
    key = str(num_comps)
    return templates.get(key) or templates.get('no')


@app.route('/api/bulk-ps', methods=['POST'])
def bulk_ps_process():
    """Position-Statement bulk fill.
    Accepts one Excel (column-oriented 'Field to Fill' sheet) and multiple
    Word templates. Template selection per record uses:
      - 'Number of Comps in Position Statement' row  -> 0/1/2/3 carriers
      - 'Procedure Type' row                         -> pain vs. spine (for 0-comp case)
    Template routing keys are auto-detected from filenames:
      '1 carriers' -> key '1', '2 carriers' -> '2', '3 carriers' -> '3',
      'no carriers' + 'pain' -> 'no_pain', 'no carriers' -> 'no'
    """
    if 'excel' not in request.files:
        return jsonify({'error': 'Missing excel file'}), 400
    word_files = request.files.getlist('word')
    if not word_files:
        return jsonify({'error': 'No Word templates provided'}), 400

    excel_file = request.files['excel']
    if not excel_file.filename:
        return jsonify({'error': 'No excel file selected'}), 400

    templates = {}
    for wf in word_files:
        if not wf.filename:
            continue
        key = detect_ps_template_key(wf.filename)
        templates[key] = wf.read()

    if not templates:
        return jsonify({'error': 'No valid Word templates uploaded'}), 400

    sheet_name = request.form.get('sheet') or None

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = os.path.join(temp_dir, excel_file.filename)
            excel_file.save(excel_path)

            try:
                records = read_excel_records_column_oriented(
                    excel_path, sheet_name=sheet_name
                )
            except Exception as e:
                return jsonify({'error': f'Excel processing error: {str(e)}'}), 400

            if not records:
                return jsonify({'error': 'No data records found to fill'}), 400

            counts = {k: 0 for k in templates}
            zip_buffer = BytesIO()
            used_names = set()

            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for idx, record in enumerate(records, start=1):
                    num_comps = record.get('num_comps')
                    procedure_type = record.get('procedure_type') or ''
                    tmpl_bytes = select_ps_template(num_comps, procedure_type, templates)

                    if tmpl_bytes is None:
                        continue

                    doc, _ = fill_document(BytesIO(tmpl_bytes), record['mappings'])

                    # Build filename
                    parts = []
                    if record.get('patient_name'):
                        parts.append(safe_filename_part(record['patient_name']))
                    if record.get('dispute_id'):
                        parts.append(safe_filename_part(record['dispute_id']))
                    if not parts:
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
                    zf.writestr(name, file_buffer.getvalue())

                    # Track counts by comp count
                    comp_key = str(num_comps) if num_comps else 'no'
                    if comp_key not in counts:
                        counts[comp_key] = 0
                    counts[comp_key] += 1

            zip_buffer.seek(0)
            response = send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name='position_statements_filled.zip',
            )
            response.headers['X-Record-Count'] = str(len(records))
            response.headers['X-Template-Counts'] = str(counts)
            response.headers['Access-Control-Expose-Headers'] = (
                'X-Record-Count, X-Template-Counts, Content-Disposition'
            )
            return response

    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500


if __name__ == '__main__':
    # Running on 0.0.0.0 to make it accessible in all environments
    app.run(host='0.0.0.0', port=5001, debug=True)
