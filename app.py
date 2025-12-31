"""
Workers Compensation Processor - Web Application
Flask backend for processing workers' comp reports
"""

import os
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file, session
from werkzeug.utils import secure_filename
import pandas as pd

# Import processing modules
from processing.wage_processor import load_and_process_wage_report
from processing.report_combiner import combine_reports
from processing.excel_exporter import process_csv_data, import_formatted_data_to_excel

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')

# Configuration
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
TEMPLATE_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates_excel')
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_session_folder(folder_type='upload'):
    """Get or create session-specific folder"""
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())

    base_folder = UPLOAD_FOLDER if folder_type == 'upload' else OUTPUT_FOLDER
    session_folder = os.path.join(base_folder, session['session_id'])
    os.makedirs(session_folder, exist_ok=True)
    return session_folder


@app.route('/')
def index():
    """Render main application page"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    file_type = request.form.get('type', 'asr')  # 'asr' or 'armorpro'

    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Allowed: CSV, XLSX, XLS'}), 400

    try:
        session_folder = get_session_folder('upload')
        filename = secure_filename(file.filename)

        # Add prefix to distinguish file types
        prefix = 'asr_' if file_type == 'asr' else 'armorpro_'
        saved_filename = f"{prefix}{filename}"
        file_path = os.path.join(session_folder, saved_filename)

        file.save(file_path)

        # Get file info
        file_size = os.path.getsize(file_path)

        return jsonify({
            'success': True,
            'filename': saved_filename,
            'original_name': file.filename,
            'size': file_size,
            'type': file_type
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/process', methods=['POST'])
def process_reports():
    """Process uploaded reports"""
    try:
        data = request.get_json()
        pay_period = data.get('pay_period', '')
        asr_file = data.get('asr_file')
        armorpro_file = data.get('armorpro_file')

        if not asr_file:
            return jsonify({'error': 'ASR report is required'}), 400

        if not pay_period:
            return jsonify({'error': 'Pay period is required'}), 400

        # Validate pay period format
        try:
            datetime.strptime(pay_period, "%Y%m%d")
        except ValueError:
            return jsonify({'error': 'Invalid pay period format. Use YYYYMMDD'}), 400

        session_upload = get_session_folder('upload')
        session_output = get_session_folder('output')

        asr_path = os.path.join(session_upload, asr_file)

        if not os.path.exists(asr_path):
            return jsonify({'error': 'ASR file not found'}), 400

        results = {
            'steps': [],
            'files': [],
            'summary': {}
        }

        # Step 1: Process ASR report
        results['steps'].append({'step': 1, 'name': 'Loading ASR Report', 'status': 'processing'})

        asr_df, asr_output_path = load_and_process_wage_report(
            asr_path,
            session_output,
            "ASRWorkersCompReport.csv"
        )
        results['steps'][-1]['status'] = 'complete'
        results['files'].append({
            'name': os.path.basename(asr_output_path),
            'path': asr_output_path,
            'type': 'csv'
        })

        # Step 2: Process ArmorPro report (if provided)
        armorpro_output_path = None
        if armorpro_file:
            results['steps'].append({'step': 2, 'name': 'Loading ArmorPro Report', 'status': 'processing'})
            armorpro_path = os.path.join(session_upload, armorpro_file)

            if os.path.exists(armorpro_path):
                armorpro_df, armorpro_output_path = load_and_process_wage_report(
                    armorpro_path,
                    session_output,
                    "ArmorProWorkersCompReport.csv"
                )
                results['steps'][-1]['status'] = 'complete'
                results['files'].append({
                    'name': os.path.basename(armorpro_output_path),
                    'path': armorpro_output_path,
                    'type': 'csv'
                })
        else:
            results['steps'].append({'step': 2, 'name': 'ArmorPro Report', 'status': 'skipped'})

        # Step 3: Combine reports (if ArmorPro was provided)
        if armorpro_output_path:
            results['steps'].append({'step': 3, 'name': 'Combining Reports', 'status': 'processing'})
            combined_path = combine_reports(asr_output_path, armorpro_output_path, session_output)
            results['steps'][-1]['status'] = 'complete'
            results['files'].append({
                'name': os.path.basename(combined_path),
                'path': combined_path,
                'type': 'csv'
            })
            csv_for_excel = combined_path
        else:
            results['steps'].append({'step': 3, 'name': 'Combining Reports', 'status': 'skipped'})
            csv_for_excel = asr_output_path

        # Step 4: Process data for Excel
        results['steps'].append({'step': 4, 'name': 'Processing Data', 'status': 'processing'})
        processed_df, total_source_earnings = process_csv_data(csv_for_excel)
        results['steps'][-1]['status'] = 'complete'

        # Step 5: Check for Excel template
        results['steps'].append({'step': 5, 'name': 'Preparing Excel Template', 'status': 'processing'})

        # Look for template in multiple locations
        template_locations = [
            os.path.join(TEMPLATE_FOLDER, '2025 ASR - WC SHOP Spreadsheet 08.2025.xlsx'),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), '2025 ASR - WC SHOP Spreadsheet 08.2025.xlsx'),
        ]

        excel_template = None
        for loc in template_locations:
            if os.path.exists(loc):
                excel_template = loc
                break

        if not excel_template:
            # If no template, we'll generate a standalone Excel file
            results['steps'][-1]['status'] = 'warning'
            results['steps'][-1]['message'] = 'No Excel template found. Generating standalone report.'

            # Create a simple Excel export
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_output = os.path.join(session_output, f"Workers_Comp_{pay_period}_{timestamp}.xlsx")

            # Add summary columns
            processed_df['Total Wages'] = processed_df['REGULAR'] + processed_df['OVERTIME'] + processed_df['DOUBLETIME']
            processed_df.to_excel(excel_output, index=False, sheet_name='Payroll Data')

            results['files'].append({
                'name': os.path.basename(excel_output),
                'path': excel_output,
                'type': 'xlsx'
            })
        else:
            results['steps'][-1]['status'] = 'complete'

            # Step 6: Generate Excel report
            results['steps'].append({'step': 6, 'name': 'Generating Excel Report', 'status': 'processing'})

            excel_output, totals = import_formatted_data_to_excel(
                processed_df,
                excel_template,
                session_output,
                total_source_earnings,
                pay_period=pay_period
            )

            results['steps'][-1]['status'] = 'complete'
            results['files'].append({
                'name': os.path.basename(excel_output),
                'path': excel_output,
                'type': 'xlsx'
            })

            results['summary'] = {
                'regular_wages': round(totals['regular'], 2),
                'overtime_wages': round(totals['overtime'], 2),
                'doubletime_wages': round(totals['doubletime'], 2),
                'grand_total': round(totals['grand_total'], 2),
                'record_count': totals['record_count'],
                'source_total': round(total_source_earnings, 2)
            }

        # Store file paths in session for download
        session['output_files'] = results['files']

        return jsonify({
            'success': True,
            'results': results
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/download/<filename>')
def download_file(filename):
    """Download generated file"""
    try:
        session_output = get_session_folder('output')
        file_path = os.path.join(session_output, secure_filename(filename))

        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404

        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/cleanup', methods=['POST'])
def cleanup_session():
    """Clean up session files"""
    try:
        if 'session_id' in session:
            import shutil

            upload_folder = os.path.join(UPLOAD_FOLDER, session['session_id'])
            output_folder = os.path.join(OUTPUT_FOLDER, session['session_id'])

            if os.path.exists(upload_folder):
                shutil.rmtree(upload_folder)
            if os.path.exists(output_folder):
                shutil.rmtree(output_folder)

            session.clear()

        return jsonify({'success': True})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/health')
def health_check():
    """Health check endpoint for Render"""
    return jsonify({'status': 'healthy', 'timestamp': datetime.now().isoformat()})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug)
