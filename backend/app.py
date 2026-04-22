import os
import zipfile
import tempfile
from io import BytesIO
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from core_logic import read_excel_mappings, fill_document

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

if __name__ == '__main__':
    # Running on 0.0.0.0 to make it accessible in all environments
    app.run(host='0.0.0.0', port=5000, debug=True)
