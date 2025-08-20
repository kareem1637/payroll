import pandas as pd
import os
from flask import Flask, request, jsonify, render_template, send_file
from werkzeug.utils import secure_filename
import sys
import threading
import webbrowser
import re
from rapidfuzz import process, fuzz
import numpy as np
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.dml import MSO_FILL
from pptx.oxml.xmlchemy import OxmlElement


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS  # PyInstaller extracts to this temp dir
    return os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__,
            template_folder=os.path.join(get_base_dir(), 'templates'))

BASE_DIR = get_base_dir()
print("Base Directory:", BASE_DIR)
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


from flask import send_from_directory

# Serve files from the uploads directory
@app.route('/uploads/<path:filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

# Function to open the default web browser
import os
import importlib.util
from importlib.machinery import SourceFileLoader
def open_browser():
    webbrowser.open_new("http://localhost:8000")

# --- Load feature modules (without importing their Flask apps) ---
BASE_DIR = get_base_dir()

def _load_module_from_path(mod_name: str, path: str):
    spec = importlib.util.spec_from_file_location(mod_name, path)
    module = importlib.util.module_from_spec(spec)
    if spec and spec.loader:
        spec.loader.exec_module(module)
        return module
    raise ImportError(f"Could not load module {mod_name} from {path}")

# Margin Report module (file name has space, so load via path)
MARGIN_MOD = None
margin_path = os.path.join(BASE_DIR, 'Margin Report.py')
if os.path.exists(margin_path):
    try:
        MARGIN_MOD = _load_module_from_path('margin_report_mod', margin_path)
    except Exception as e:
        print('Warning: could not load Margin Report module:', e)

# PBJ/PPG module
PBJ_MOD = None
pbj_path = os.path.join(BASE_DIR, 'PBJ_report.py')
if os.path.exists(pbj_path):
    try:
        PBJ_MOD = _load_module_from_path('pbj_report_mod', pbj_path)
    except Exception as e:
        print('Warning: could not load PBJ Report module:', e)

# Productivity module
PROD_MOD = None
prod_path = os.path.join(BASE_DIR, 'Productivity.py')
if os.path.exists(prod_path):
    try:
        PROD_MOD = _load_module_from_path('productivity_mod', prod_path)
    except Exception as e:
        print('Warning: could not load Productivity module:', e)

# Payroll module
PAYROLL_MOD = None
payroll_path = os.path.join(BASE_DIR, 'payroll.py')
if os.path.exists(payroll_path):
    try:
        PAYROLL_MOD = _load_module_from_path('payroll_mod', payroll_path)
    except Exception as e:
        print('Warning: could not load Payroll module:', e)

# --- Routes: Home and Pages ---
@app.route('/')
def home():
    return render_template('home.html')

@app.route('/margin')
def page_margin():
    return render_template('Margin_report.html')

@app.route('/ppg')
def page_ppg():
    return render_template('PPG_report.html')

@app.route('/weekly')
def page_weekly():
    return render_template('Weekly_report.html')

@app.route('/dashboard')
def page_dashboard():
    return render_template('dashboard.html')

# --- Namespaced APIs: Margin ---
@app.route('/api/margin/list_pr_sheets', methods=['POST'])
def api_margin_list_pr_sheets():
    if 'PRFile' not in request.files:
        return jsonify({'error': 'No PRFile provided'}), 400
    pr_file = request.files['PRFile']
    filename = secure_filename(pr_file.filename or '')
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ('.xls', '.xlsx'):
        return jsonify({'error': 'Invalid PR file type; expected .xls or .xlsx'}), 400
    try:
        xls = pd.ExcelFile(pr_file)
        return jsonify({'sheets': xls.sheet_names})
    except Exception:
        try:
            tmp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"tmp_pr_{threading.get_ident()}{ext}")
            pr_file.seek(0)
            pr_file.save(tmp_path)
            xls = pd.ExcelFile(tmp_path)
            sheets = xls.sheet_names
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return jsonify({'sheets': sheets})
        except Exception as e2:
            return jsonify({'error': str(e2)}), 500

@app.route('/api/margin/upload_data', methods=['POST'])
def api_margin_upload_data():
    if MARGIN_MOD is None:
        return jsonify({'error': 'Margin module not available'}), 500
    if 'Roster' not in request.files or 'captureFile' not in request.files:
        return jsonify({'error': 'Missing files'}), 400

    Roster_file = request.files['Roster']
    capture_file = request.files['captureFile']
    PR_file = request.files.get('PRFile')
    FDR_file = request.files.get('FDRFile')
    PR_sheet_name = request.form.get('PR_sheet_name', 'Payroll Register')
    day = request.form.get('day')
    date = request.form.get('date')

    capture_file_filename = secure_filename(capture_file.filename)
    Roster_file_filename = secure_filename(Roster_file.filename)
    PR_file_filename = secure_filename(PR_file.filename) if PR_file and PR_file.filename else None
    FDR_file_filename = secure_filename(FDR_file.filename) if FDR_file and FDR_file.filename else None

    captuer_ext = os.path.splitext(capture_file_filename)[1].lower()
    roster_ext = os.path.splitext(Roster_file_filename)[1].lower()
    PR_ext = os.path.splitext(PR_file_filename)[1].lower() if PR_file_filename else None
    FDR_ext = os.path.splitext(FDR_file_filename)[1].lower() if FDR_file_filename else None
    if captuer_ext not in ['.csv', '.xlsx'] or roster_ext not in ['.csv', '.xlsx']:
        return jsonify({'error': 'Invalid file type'}), 400
    # Require PR and FDR for this workflow
    if not PR_file_filename or not FDR_file_filename:
        return jsonify({'error': 'PRFile (.xls) and FDRFile (.xlsx) are required'}), 400
    if not PR_ext.endswith('.xls'):
        return jsonify({'error': 'Invalid PR file type, please upload with .xls extension'}), 400
    if not FDR_ext.endswith('.xlsx'):
        return jsonify({'error': 'Invalid FDR file type, please upload with .xlsx extension'}), 400

    capture_temp_path = os.path.join(app.config['UPLOAD_FOLDER'], capture_file_filename)
    roster_temp_path = os.path.join(app.config['UPLOAD_FOLDER'], Roster_file_filename)
    PR_temp_path = os.path.join(app.config['UPLOAD_FOLDER'], PR_file_filename)
    FDR_temp_path = os.path.join(app.config['UPLOAD_FOLDER'], FDR_file_filename)
    capture_file.save(capture_temp_path)
    Roster_file.save(roster_temp_path)
    PR_file.save(PR_temp_path)
    FDR_file.save(FDR_temp_path)

    # Read capture & roster
    charge_capture_df = pd.read_csv(capture_temp_path) if capture_file_filename.endswith('.csv') else pd.read_excel(capture_temp_path)
    company_roaster = pd.read_csv(roster_temp_path) if Roster_file_filename.endswith('.csv') else pd.read_excel(roster_temp_path, skiprows=2)
    # Read PR & FDR
    PR_df = pd.read_excel(PR_temp_path, sheet_name=PR_sheet_name, header=None)
    FDR_df = pd.read_excel(FDR_temp_path)

    # Process via margin module
    PR_df, FDR_df = MARGIN_MOD.preprocess_PR_FDR(PR_df, FDR_df)
    Regional_Dashboard, unmatched_providers, matched_providers, grouped_CC_ByRegion, Margin_df = MARGIN_MOD.build_metadata(
        charge_capture_df, company_roaster, PR_df, FDR_df
    )
    prs = MARGIN_MOD.load_editable_presentation(os.path.join(BASE_DIR, 'static', 'Margin Report.pptx'), day=day, Date=date)
    MARGIN_MOD.generate_pbj_presentation(prs, Regional_Dashboard, grouped_CC_ByRegion, day, date, Add_region=True)
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Margin_Report.pptx')
    workBook_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Work_book.xlsx')
    MARGIN_MOD.save_workbook(Margin_df, workBook_path)
    prs.save(output_path)

    return jsonify({
        'success': True,
        'pptx_path': f"uploads/Margin_Report.pptx",
        'matched_providers': matched_providers,
        'unmatched_providers': unmatched_providers
    })

# --- Namespaced APIs: PBJ/PPG ---
@app.route('/api/ppg/upload_data', methods=['POST'])
def api_ppg_upload_data():
    if PBJ_MOD is None:
        return jsonify({'error': 'PPG/PBJ module not available'}), 500
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    filename = secure_filename(file.filename or '')
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ['.csv', '.xlsx']:
        return jsonify({'error': 'Invalid file type'}), 400

    # Save uploaded file
    temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename or f"charge_{threading.get_ident()}{ext}")
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    file.save(temp_path)

    # Load DataFrame
    if ext == '.csv':
        charge_capture_df = pd.read_csv(temp_path)
    else:
        charge_capture_df = pd.read_excel(temp_path)

    # Get CPT dict
    import json
    cpt_time_dict = request.form.get('cpt_time_dict')
    if cpt_time_dict:
        try:
            cpt_time_dict = json.loads(cpt_time_dict)
        except Exception:
            cpt_time_dict = None
    # Month and report type
    month_index_raw = request.form.get('month')
    try:
        month_index = int(month_index_raw)
    except (TypeError, ValueError):
        month_index = 4
    months = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]
    month_name = months[month_index] if 0 <= month_index < 12 else "May"
    report_type = request.form.get('report_type', 'Both')

    # Build metadata and presentation
    corporate_groups, single_facilities = PBJ_MOD.group_facilities(charge_capture_df)
    metadata = PBJ_MOD.build_metadata(charge_capture_df, corporate_groups, single_facilities)
    prs = PBJ_MOD.load_editable_presentation(os.path.join(BASE_DIR, 'static', 'reference_slide.pptx'), month_index=month_index)
    PBJ_MOD.generate_pbj_presentation(prs, metadata, month=month_name, report_type=report_type)
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'generated_pbj_report.pptx')
    prs.save(output_path)
    return jsonify({'success': True, 'pptx_path': f"uploads/generated_pbj_report.pptx", 'month': month_name, 'report_type': report_type})

@app.route('/api/ppg/upload_logo', methods=['POST'])
def api_ppg_upload_logo():
    if PBJ_MOD is None:
        return jsonify({'error': 'PPG/PBJ module not available'}), 500
    if 'logo' not in request.files or 'corp_name' not in request.form:
        return jsonify({'error': 'Logo file and corp_name are required'}), 400
    logo_file = request.files['logo']
    corp_name = request.form['corp_name'].strip().lower().replace(' ', '_')
    ext = os.path.splitext(logo_file.filename or '')[1].lower()
    if ext not in ['.jpg', '.jpeg', '.png']:
        return jsonify({'error': 'Invalid file type'}), 400
    logos_dir = os.path.join(BASE_DIR, 'static', 'logos')
    os.makedirs(logos_dir, exist_ok=True)
    logo_filename = f"{corp_name}_logo{ext}"
    logo_path = os.path.join(logos_dir, logo_filename)
    logo_file.save(logo_path)
    # Return relative static path for browser
    return jsonify({'success': True, 'logo_path': f"static/logos/{logo_filename}"})

# --- Namespaced APIs: Productivity ---
@app.route('/api/productivity/upload_data', methods=['POST'])
def api_productivity_upload_data():
    if PROD_MOD is None:
        return jsonify({'error': 'Productivity module not available'}), 500
    if 'Roster' not in request.files or 'captureFile' not in request.files:
        return jsonify({'error': 'Missing files'}), 400
    roster_file = request.files['Roster']
    capture_file = request.files['captureFile']

    roster_filename = secure_filename(roster_file.filename or '')
    capture_filename = secure_filename(capture_file.filename or '')
    r_ext = os.path.splitext(roster_filename)[1].lower()
    c_ext = os.path.splitext(capture_filename)[1].lower()
    if r_ext not in ['.csv', '.xlsx'] or c_ext not in ['.csv', '.xlsx']:
        return jsonify({'error': 'Invalid file type'}), 400

    roster_path = os.path.join(app.config['UPLOAD_FOLDER'], roster_filename or f"roster_{threading.get_ident()}{r_ext}")
    capture_path = os.path.join(app.config['UPLOAD_FOLDER'], capture_filename or f"capture_{threading.get_ident()}{c_ext}")
    roster_file.save(roster_path)
    capture_file.save(capture_path)

    # Load DataFrames
    company_roster_df = pd.read_csv(roster_path) if r_ext == '.csv' else pd.read_excel(roster_path, skiprows=2)
    charge_capture_df = pd.read_csv(capture_path) if c_ext == '.csv' else pd.read_excel(capture_path)

    day = request.form.get('day')
    date = request.form.get('date')
    include_regions = request.form.get('include_regions', True)

    Regional_Dashboard, unmatched_providers, matched_providers, gouped_CC_ByRegion, _ = PROD_MOD.build_metadata(charge_capture_df, company_roster_df)
    prs = PROD_MOD.load_editable_presentation(os.path.join(BASE_DIR, 'static', 'Weekly Report Example.pptx'), day=day, Date=date)
    PROD_MOD.generate_pbj_presentation(prs, Regional_Dashboard, gouped_CC_ByRegion, day, date, Add_region=include_regions)
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Weekly_report_report.pptx')
    prs.save(output_path)
    # Map keys to match Weekly_report.html expectations
    def _map_rows(rows):
        out = []
        for r in rows or []:
            out.append({
                'company_roster': r.get('providers') or r.get('company_roster') or r.get('Provider') or r.get('Clinician'),
                'charge_capture_name': r.get('org_providers_name') or r.get('charge_capture_name') or r.get('Provider'),
                'score': r.get('score')
            })
        return out
    return jsonify({'success': True, 'pptx_path': f"uploads/Weekly_report_report.pptx", 'matched_providers': _map_rows(matched_providers), 'unmatched_providers': _map_rows(unmatched_providers)})

# --- Namespaced APIs: Payroll ---
@app.route('/api/payroll/process', methods=['POST'])
def api_payroll_process():
    if PAYROLL_MOD is None:
        return jsonify({'error': 'Payroll module not available'}), 500
    if 'payrollFile' not in request.files or 'captureFile' not in request.files:
        return jsonify({'error': 'Missing files'}), 400
    payroll_file = request.files['payrollFile']
    capture_file = request.files['captureFile']
    payroll_sheet = request.form.get('payrollSheet')
    output_filename = request.form.get('outputFileName', 'processed_payroll.xlsx')

    payroll_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(payroll_file.filename or f"payroll_{threading.get_ident()}.xlsx"))
    capture_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(capture_file.filename or f"capture_{threading.get_ident()}.xlsx"))
    payroll_file.save(payroll_path)
    capture_file.save(capture_path)

    try:
        results = PAYROLL_MOD.process_files(payroll_path, capture_path, payroll_sheet, output_filename)
        invalid_cpts_array = [
            {"name": codes.get("name"), "found": {k: v for k, v in codes.items() if k != "name"}}
            for _, codes in results.get('invalid_cpts', {}).items()
        ]
        return jsonify({
            'invalidCPTs': invalid_cpts_array,
            'missingPositions': results.get('missing_positions', []),
            'missingProviders': results.get('missing_providers', []),
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/payroll/add_missing_cpt', methods=['POST'])
def api_payroll_add_missing_cpt():
    if PAYROLL_MOD is None:
        return jsonify({'error': 'Payroll module not available'}), 500
    data = request.get_json(silent=True) or {}
    provider = data.get('provider')
    cpt = data.get('cpt')
    dos = data.get('dos')
    if not provider or not cpt:
        return jsonify({'error': 'Missing provider or CPT'}), 400
    try:
        # Update in-memory dict and weekly counts in the module
        PAYROLL_MOD.update_cpt_counts(PAYROLL_MOD.provider_cpt_dict, provider, cpt)
        PAYROLL_MOD.weekly_counts = PAYROLL_MOD.increment_encounter(provider, dos, PAYROLL_MOD.weekly_counts, PAYROLL_MOD.date_range, cpt)
        PAYROLL_MOD.manual_cpt_updates.append({"provider": provider, "cpt": cpt})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/payroll/save_changes', methods=['POST'])
def api_payroll_save_changes():
    if PAYROLL_MOD is None:
        return jsonify({'error': 'Payroll module not available'}), 500
    # Ensure workbook and data present
    if not getattr(PAYROLL_MOD, 'current_workbook', None) or not getattr(PAYROLL_MOD, 'sheet', None):
        return jsonify({'error': 'No workbook loaded. Process files first.'}), 400
    output_filename = PAYROLL_MOD.output_filename or 'processed_payroll_download.xlsx'
    try:
        # Write latest data into sheet
        PAYROLL_MOD.write_provider_cpt_data_to_sheet(
            payroll_df=PAYROLL_MOD.payroll_df,
            common_providers=PAYROLL_MOD.common_providers,
            practitioner_list=PAYROLL_MOD.practitioner_list,
            provider_cpt_dict=PAYROLL_MOD.provider_cpt_dict,
            cpt_positions=PAYROLL_MOD.cpt_positions,
            output_filename=output_filename
        )
        # Save workbook to bytes
        from io import BytesIO
        file_stream = BytesIO()
        PAYROLL_MOD.current_workbook.save(file_stream)
        file_stream.seek(0)
        return send_file(file_stream, as_attachment=True, download_name=output_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': f"Failed to prepare file for download: {str(e)}"}), 500

@app.route('/api/payroll/add_new_cpt_column', methods=['POST'])
def api_payroll_add_new_cpt_column():
    if PAYROLL_MOD is None:
        return jsonify({'error': 'Payroll module not available'}), 500
    data = request.get_json(silent=True) or {}
    cpt_code = (data.get('cpt') or '').strip()
    if not cpt_code:
        return jsonify({'success': False, 'error': 'CPT code is required.'}), 400
    if cpt_code in PAYROLL_MOD.cpt_positions:
        return jsonify({'success': False, 'error': 'CPT code already exists.'}), 400
    try:
        last_key = list(PAYROLL_MOD.cpt_positions)[-1]
        last_cpt_index = PAYROLL_MOD.cpt_positions[last_key]
        new_cpt_index = last_cpt_index + 6  # not zero based in their sheet logic
        PAYROLL_MOD.add_new_cpt(cpt_code, new_cpt_index)
        # Update indices and mapping
        PAYROLL_MOD.cpt_positions[cpt_code] = new_cpt_index - 1  # zero-based
        PAYROLL_MOD.Gross_encounters_col = (PAYROLL_MOD.Gross_encounters_col or 0) + 5
        if PAYROLL_MOD.week2_encounters_col_idx is not None:
            PAYROLL_MOD.week2_encounters_col_idx += 5
        if PAYROLL_MOD.week1_encounters_col_idx is not None:
            PAYROLL_MOD.week1_encounters_col_idx += 5
        # Recompute CPT counts from capture_df if available
        if getattr(PAYROLL_MOD, 'capture_df', None) is not None:
            cpt_counts_df = PAYROLL_MOD.capture_df.groupby(['Provider', 'CPT Codes']).size().reset_index(name='Counts')
            PAYROLL_MOD.process_cpt_counts(cpt_counts_df, PAYROLL_MOD.cpt_positions)
        PAYROLL_MOD.apply_manual_cpt_updates()
        PAYROLL_MOD.write_provider_cpt_data_to_sheet(
            payroll_df=PAYROLL_MOD.payroll_df,
            common_providers=PAYROLL_MOD.common_providers,
            practitioner_list=PAYROLL_MOD.practitioner_list,
            provider_cpt_dict=PAYROLL_MOD.provider_cpt_dict,
            cpt_positions=PAYROLL_MOD.cpt_positions,
            output_filename=PAYROLL_MOD.output_filename or 'processed_payroll.xlsx'
        )
        return jsonify({'success': True, 'invalid_cpts': PAYROLL_MOD.not_recognized}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    # Only open browser in local development
    if os.environ.get('RENDER') is None:
        threading.Timer(1, open_browser).start()  # Open the browser after 1 second
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)

