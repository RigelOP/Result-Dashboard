import os
import threading
import time
import uuid
from queue import Queue
from flask import Flask, request, render_template, send_from_directory, jsonify, redirect, url_for
import pandas as pd
from datetime import datetime

# Import functions from adapter module for clarity
from minimal_browser_extractor import process_single_student, save_results_to_excel
from Result_Downloader import create_grade_distribution_analysis
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

APP_ROOT = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(APP_ROOT, 'uploads')
OUTPUT_FOLDER = os.path.join(APP_ROOT, 'outputs')
SAMPLE_FILE = os.path.join(APP_ROOT, 'sample.xlsx')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Global processing state
STATE = {
    'students': [],
    'results': [],
    'current_index': 0,
    'running': False,
    'successful': 0,
    'failed': 0,
    'filename': None,
    'subject_info': {},
}
STATE_LOCK = threading.Lock()

# File expiry tracking (timestamps are epoch seconds)
UPLOAD_EXPIRY = {}   # upload filename -> expiry_ts
RESULT_EXPIRY = {}   # result filename -> expiry_ts
EXPIRY_FILE = os.path.join(APP_ROOT, '.expiries.json')
EXPIRY_LOCK = threading.Lock()

def load_expiries():
    try:
        if os.path.exists(EXPIRY_FILE):
            import json
            with open(EXPIRY_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # Only keep entries whose timestamps are still in the future
            now = time.time()
            with EXPIRY_LOCK:
                UPLOAD_EXPIRY.clear()
                RESULT_EXPIRY.clear()
                for k, v in (data.get('uploads') or {}).items():
                    if v and v > now:
                        UPLOAD_EXPIRY[k] = v
                for k, v in (data.get('results') or {}).items():
                    if v and v > now:
                        RESULT_EXPIRY[k] = v
    except Exception:
        pass

def save_expiries():
    try:
        import json
        with EXPIRY_LOCK:
            data = {'uploads': UPLOAD_EXPIRY, 'results': RESULT_EXPIRY}
        with open(EXPIRY_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f)
    except Exception:
        pass

def set_upload_expiry(filename, seconds_from_now=10*60):
    try:
        with EXPIRY_LOCK:
            UPLOAD_EXPIRY[filename] = time.time() + seconds_from_now
        save_expiries()
    except Exception:
        pass

def set_result_expiry(filename, seconds_from_now=20*60):
    try:
        with EXPIRY_LOCK:
            RESULT_EXPIRY[filename] = time.time() + seconds_from_now
        save_expiries()
    except Exception:
        pass

# Cleaner thread: deletes expired uploaded files (10m) and result files (20m)
def cleaner_loop():
    while True:
        now = time.time()
        # uploads
        for fname, exp in list(UPLOAD_EXPIRY.items()):
            if now >= exp:
                path = os.path.join(UPLOAD_FOLDER, fname)
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except:
                    pass
                UPLOAD_EXPIRY.pop(fname, None)
        # results
        for fname, exp in list(RESULT_EXPIRY.items()):
            if now >= exp:
                path = os.path.join(OUTPUT_FOLDER, fname)
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except:
                    pass
                RESULT_EXPIRY.pop(fname, None)
                # if the deleted file was the current STATE filename, clear it
                with STATE_LOCK:
                    if STATE.get('filename') == fname:
                        STATE['filename'] = None
        time.sleep(60)

# start cleaner thread
threading.Thread(target=cleaner_loop, daemon=True).start()

def create_driver_headless():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--log-level=3')
    options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)
    # Speed: don't load images
    prefs = {"profile.managed_default_content_settings.images": 2}
    options.add_experimental_option("prefs", prefs)
    # Use eager page load to speed navigation (don't wait for all resources)
    try:
        options.page_load_strategy = 'eager'
    except Exception:
        pass
    os.environ.setdefault('WDM_LOG_LEVEL', '0')
    service = Service(ChromeDriverManager().install(), log_path=os.devnull)
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def worker_process(upload_path, class_name, semester):
    with STATE_LOCK:
        STATE['running'] = True
        STATE['current_index'] = 0
        STATE['results'] = []
        STATE['successful'] = 0
        STATE['failed'] = 0

    df = pd.read_excel(upload_path)
    students = []
    for i in range(len(df)):
        row = df.iloc[i]
        if not (pd.isna(row['Roll Number']) or pd.isna(row['Registration Number']) or pd.isna(row['Student Name'])):
            students.append({
                'name': str(row['Student Name']).strip(),
                'roll': str(int(row['Roll Number'])),
                'reg': str(int(row['Registration Number']))
            })

    with STATE_LOCK:
        STATE['students'] = students

    driver = create_driver_headless()
    # open base result page once
    try:
        driver.get('https://result.mdurtk.in/postexam/result.aspx')
    except Exception:
        pass

    for i, student in enumerate(students, 1):
        with STATE_LOCK:
            STATE['current_index'] = i
        # Clear cookies and reload page between students to avoid stale session state
        try:
            driver.delete_all_cookies()
            driver.get('https://result.mdurtk.in/postexam/result.aspx')
            time.sleep(0.1)
        except Exception:
            pass
        try:
            result = process_single_student(driver, student['name'], student['roll'], student['reg'])
        except Exception as e:
            result = {
                'name': student['name'],
                'roll': student['roll'],
                'reg': student['reg'],
                'status': f'Failed - {e}',
                'marks': {}
            }

        with STATE_LOCK:
            STATE['results'].append(result)
            if result['status'] == 'Success':
                STATE['successful'] += 1
            else:
                STATE['failed'] += 1

        # Save intermediate results occasionally to reduce IO (every 5 students and final)
        if i % 5 == 0 or i == len(students):
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            # Use class and semester if provided for filename
            safe_class = (class_name or 'Class').strip()
            safe_sem = (semester or 'Semester').strip()
            base_name = f"{safe_class} {safe_sem} result.xlsx"
            outname = base_name
            outpath = os.path.join(OUTPUT_FOLDER, outname)
            # If file exists, append timestamp to avoid overwrite
            if os.path.exists(outpath):
                outname = f"{safe_class} {safe_sem} result {timestamp}.xlsx"
                outpath = os.path.join(OUTPUT_FOLDER, outname)
            save_results_to_excel(STATE['results'], outpath)
            # record result expiry (20 minutes) and persist
            set_result_expiry(outname, seconds_from_now=20*60)
            with STATE_LOCK:
                STATE['filename'] = outname

    try:
        driver.quit()
    except:
        pass

    with STATE_LOCK:
        STATE['running'] = False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download-sample')
def download_sample():
    return send_from_directory(APP_ROOT, 'sample.xlsx', as_attachment=True)

@app.route('/upload', methods=['POST'])
def upload():
    class_name = request.form.get('class_name', '')
    semester = request.form.get('semester', '')
    file = request.files.get('file')
    if not file:
        return 'No file', 400
    filename = f"upload_{uuid.uuid4().hex}.xlsx"
    upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(upload_path)
    # record upload expiry (10 minutes) and persist
    set_upload_expiry(filename, seconds_from_now=10*60)

    # start background worker
    t = threading.Thread(target=worker_process, args=(upload_path, class_name, semester), daemon=True)
    t.start()

    return redirect(url_for('progress'))

@app.route('/progress')
def progress():
    return render_template('progress.html')

@app.route('/status')
def status():
    with STATE_LOCK:
        total = len(STATE['students'])
        current = STATE['current_index']
        running = STATE['running']
        successful = STATE['successful']
        failed = STATE['failed']
        filename = STATE.get('filename')
        curr_student = None
        if 0 < current <= total:
            curr_student = STATE['students'][current-1]['name']
        failed_list = [r for r in STATE['results'] if r['status'] != 'Success']

    # compute expires_in and expires_at for client convenience
    expires_in = None
    expires_at = None
    if filename and filename in RESULT_EXPIRY:
        expires_at = int(RESULT_EXPIRY.get(filename, 0))
        expires_in = max(0, int(expires_at - time.time()))

    return jsonify({
        'total': total,
        'current': current,
        'running': running,
        'successful': successful,
        'failed': failed,
        'current_student': curr_student,
        'failed_list': failed_list,
        'result_file': filename,
        # seconds until result file auto-deletes (null if unknown)
        'result_expires_in': expires_in,
        # absolute epoch seconds when file will expire (null if unknown)
        'result_expires_at': expires_at
    })

@app.route('/failed/<int:idx>')
def failed_detail(idx):
    with STATE_LOCK:
        if idx < 0 or idx >= len(STATE['results']):
            return 'Not found', 404
        result = STATE['results'][idx]
    return render_template('analysis.html', idx=idx, result=result)

@app.route('/failed/<int:idx>/submit', methods=['POST'])
def failed_submit(idx):
    # Expect a textarea 'manual_marks' where each line: CODE,NAME,TOTAL,OBTAINED
    manual = request.form.get('manual_marks', '')
    marks = {}
    for line in manual.splitlines():
        parts = [p.strip() for p in line.split(',')]
        if len(parts) >= 4:
            code = parts[0]
            obtained = parts[3]
            marks[code] = obtained

    with STATE_LOCK:
        if idx < 0 or idx >= len(STATE['results']):
            return 'Not found', 404
        STATE['results'][idx]['marks'] = marks
        STATE['results'][idx]['status'] = 'Success (manual)'

        # regenerate results file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        outname = f"Dashboard_Results_{timestamp}.xlsx"
        outpath = os.path.join(OUTPUT_FOLDER, outname)
        save_results_to_excel(STATE['results'], outpath)
        set_result_expiry(outname, seconds_from_now=20*60)
        STATE['filename'] = outname

    return redirect(url_for('progress'))

@app.route('/download-results')
def download_results():
    with STATE_LOCK:
        fname = STATE.get('filename')
    if not fname:
        # return JSON for fetch-based clients
        if request.headers.get('Accept', '').lower().find('application/json') != -1:
            return jsonify({'error': 'no_results'}), 404
        return 'No results yet', 404
    return send_from_directory(OUTPUT_FOLDER, fname, as_attachment=True)


@app.route('/detailed')
def detailed():
    # Show subject codes collected from results and ask for names and totals
    with STATE_LOCK:
        all_subjects = set()
        for r in STATE['results']:
            if r.get('marks'):
                all_subjects.update(r['marks'].keys())
    all_subjects = sorted(list(all_subjects))
    if not all_subjects:
        return redirect(url_for('progress'))
    return render_template('detailed.html', subjects=all_subjects)


@app.route('/detailed/submit', methods=['POST'])
def detailed_submit():
    # Expect fields subject_{code}_name and subject_{code}_total for each subject code
    subject_info = {}
    for key, val in request.form.items():
        if key.startswith('code_'):
            code = key[len('code_'):]
            name = val.strip()
            total_key = f'total_{code}'
            total_val = request.form.get(total_key, '').strip()
            try:
                total_marks = int(total_val)
            except:
                total_marks = None
            if name and total_marks:
                subject_info[code] = {'name': name, 'total_marks': total_marks}

    # regenerate results file with subject_info and store subject_info in STATE
    with STATE_LOCK:
        results_copy = list(STATE['results'])
        STATE['subject_info'] = subject_info
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    outname = f"Dashboard_Results_{timestamp}.xlsx"
    outpath = os.path.join(OUTPUT_FOLDER, outname)
    save_results_to_excel(results_copy, outpath, subject_info)
    set_result_expiry(outname, seconds_from_now=20*60)
    with STATE_LOCK:
        STATE['filename'] = outname

    return redirect(url_for('results'))


@app.route('/results')
def results():
    with STATE_LOCK:
        results_copy = list(STATE['results'])
        subject_info = STATE.get('subject_info', {})

    # Build columns and rows for table display
    all_subjects = sorted({s for r in results_copy for s in (r.get('marks') or {}).keys()})
    columns = ['Roll_Number', 'Student_Name', 'Registration_Number', 'Status'] + all_subjects + ['Total', 'Percentage']
    rows = []

    # Precompute max total from subject_info if available
    max_total_by_subject = {code: info.get('total_marks', 0) for code, info in subject_info.items()} if subject_info else {}

    for r in results_copy:
        row = {
            'Student_Name': r['name'],
            'Roll_Number': r['roll'],
            'Registration_Number': r['reg'],
            'Status': r['status']
        }
        total = 0
        max_total = 0
        for s in all_subjects:
            raw = r.get('marks', {}).get(s, 'N/A')
            row[s] = raw
            # extract leading digits if present (handles '29F')
            digits = ''.join([c for c in str(raw) if c.isdigit()])
            try:
                val = int(digits) if digits else 0
            except:
                val = 0
            total += val
            max_total += max_total_by_subject.get(s, 0)

        pct = None
        if max_total > 0:
            pct = round((total / max_total) * 100, 1)

        row['Total'] = total
        row['Percentage'] = f"{pct}%" if pct is not None else ''
        rows.append(row)

    # Prepare analysis (grade distribution) if subject_info available
    analysis_data = []
    analysis_columns = []
    analysis_rows = []
    with STATE_LOCK:
        subject_info = STATE.get('subject_info') or {}
    if subject_info:
        analysis_data = create_grade_distribution_analysis(results_copy, subject_info)
        if analysis_data:
            analysis_columns = list(analysis_data[0].keys())
            analysis_rows = analysis_data

    return render_template('results.html', columns=columns, rows=rows, analysis_columns=analysis_columns, analysis_rows=analysis_rows)


@app.route('/analysis')
def analysis():
    with STATE_LOCK:
        results_copy = list(STATE['results'])
        subject_info = STATE.get('subject_info') or {}

    if not subject_info:
        return redirect(url_for('results'))

    analysis_data = create_grade_distribution_analysis(results_copy, subject_info)
    if not analysis_data:
        return render_template('analysis.html', columns=[], rows=[])
    columns = list(analysis_data[0].keys())
    rows = analysis_data
    return render_template('analysis.html', columns=columns, rows=rows)

if __name__ == '__main__':
    app.run(port=5000, debug=False)
#app.run(host="0.0.0.0", port=5000, debug=True)
