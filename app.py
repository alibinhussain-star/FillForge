from flask import Flask, request, jsonify, send_file, render_template, session, redirect, url_for
import pandas as pd, re, io, tempfile, os, json, smtplib, ssl, random, string, time
from email.mime.text import MIMEText
from functools import wraps
from datetime import datetime
from openpyxl import load_workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.secret_key = os.environ.get('FLASK_SECRET', 'change-me-in-prod-' + ''.join(random.choices(string.ascii_letters, k=24)))

# ── SMTP config (set these env vars on your machine) ───────────
SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.gmail.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '587'))
SMTP_USER = os.environ.get('SMTP_USER', '')        # e.g. youraddr@gmail.com
SMTP_PASS = os.environ.get('SMTP_PASS', '')        # gmail app-password
SMTP_FROM = os.environ.get('SMTP_FROM', SMTP_USER)

# Optional allow-list (comma-separated emails / domains). Empty = allow any.
ALLOWED_EMAILS = [e.strip().lower() for e in os.environ.get('ALLOWED_EMAILS', '').split(',') if e.strip()]

# In-memory OTP store: { email: {'otp': '123456', 'exp': 1234567890, 'attempts': 0} }
OTP_STORE = {}
OTP_TTL = 5 * 60          # 5 minutes
OTP_MAX_ATTEMPTS = 5

# Logs (JSON-line file)
LOG_PATH = os.path.join(os.path.dirname(__file__), 'activity.log')

def write_log(email, action, details=''):
    entry = {
        'ts': datetime.utcnow().isoformat() + 'Z',
        'email': email or 'anonymous',
        'action': action,
        'details': details,
    }
    try:
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(json.dumps(entry) + '\n')
    except Exception as e:
        print('log err', e)

def read_logs(limit=200):
    if not os.path.exists(LOG_PATH): return []
    out = []
    try:
        with open(LOG_PATH, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line: continue
                try: out.append(json.loads(line))
                except: pass
    except: pass
    return list(reversed(out))[:limit]

def login_required(f):
    @wraps(f)
    def w(*a, **kw):
        if not session.get('email'):
            if request.path.startswith('/api') or request.method == 'POST':
                return jsonify({'error': 'auth_required'}), 401
            return redirect(url_for('login_page'))
        return f(*a, **kw)
    return w

def send_otp_email(to_email, otp):
    if not SMTP_USER or not SMTP_PASS:
        print(f'[DEV MODE] OTP for {to_email} = {otp}')
        return True, 'dev'
    msg = MIMEText(
        f"Your FillForge login code is: {otp}\n\n"
        f"This code expires in 5 minutes.\n"
        f"If you did not request this, ignore this email.\n\n— FillForge",
        'plain'
    )
    msg['Subject'] = f'FillForge login code: {otp}'
    msg['From']    = SMTP_FROM
    msg['To']      = to_email
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=15) as srv:
            srv.starttls(context=ctx)
            srv.login(SMTP_USER, SMTP_PASS)
            srv.sendmail(SMTP_FROM, [to_email], msg.as_string())
        return True, 'sent'
    except Exception as e:
        return False, str(e)

EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')

# ── Auth routes ────────────────────────────────────────────────
@app.route('/login')
def login_page():
    if session.get('email'): return redirect(url_for('index'))
    return render_template('login.html')

@app.route('/auth/send_otp', methods=['POST'])
def auth_send_otp():
    email = (request.json or {}).get('email', '').strip().lower()
    if not EMAIL_RE.match(email):
        return jsonify({'error': 'Invalid email address'}), 400
    if ALLOWED_EMAILS:
        domain = email.split('@')[-1]
        if email not in ALLOWED_EMAILS and domain not in ALLOWED_EMAILS:
            return jsonify({'error': 'This email is not authorized to access FillForge'}), 403

    otp = ''.join(random.choices(string.digits, k=6))
    OTP_STORE[email] = {'otp': otp, 'exp': time.time() + OTP_TTL, 'attempts': 0}
    ok, info = send_otp_email(email, otp)
    if not ok:
        return jsonify({'error': f'Could not send email: {info}'}), 500
    write_log(email, 'otp_requested')
    return jsonify({'status': 'ok', 'dev': info == 'dev'})

@app.route('/auth/verify_otp', methods=['POST'])
def auth_verify_otp():
    data  = request.json or {}
    email = data.get('email', '').strip().lower()
    code  = data.get('otp', '').strip()
    rec   = OTP_STORE.get(email)
    if not rec:
        return jsonify({'error': 'No OTP requested for this email'}), 400
    if time.time() > rec['exp']:
        OTP_STORE.pop(email, None)
        return jsonify({'error': 'Code expired, request a new one'}), 400
    rec['attempts'] += 1
    if rec['attempts'] > OTP_MAX_ATTEMPTS:
        OTP_STORE.pop(email, None)
        return jsonify({'error': 'Too many attempts'}), 400
    if code != rec['otp']:
        return jsonify({'error': 'Incorrect code'}), 400
    OTP_STORE.pop(email, None)
    session['email']     = email
    session['logged_at'] = datetime.utcnow().isoformat() + 'Z'
    write_log(email, 'login_success')
    return jsonify({'status': 'ok', 'email': email})

@app.route('/auth/logout', methods=['POST'])
def auth_logout():
    email = session.get('email')
    session.clear()
    if email: write_log(email, 'logout')
    return jsonify({'status': 'ok'})

@app.route('/auth/me')
def auth_me():
    return jsonify({'email': session.get('email')})

# ── Template constants (unchanged from your original) ──────────
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'Logic___Template_File.xlsx')

def _build_header_row_map():
    wb = load_workbook(TEMPLATE_PATH); ws = wb['PV Template']
    hdr_map, static_map = {}, {}
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == 'Category *':
            for r2 in range(r + 1, min(r + 5, ws.max_row + 1)):
                d2 = ws.cell(r2, 4).value
                if d2 and str(d2).strip() not in ('SubType','Category *','nan',''):
                    st = str(d2).strip()
                    if st not in hdr_map:
                        hdr_map[st] = r
                        hdrs = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                        entry = {}
                        for ci, col in enumerate(hdrs):
                            if col in ('Category *','SubCategory *','CategoryType *','SubType','PVID *','discoveryCategoryIds'):
                                v = str(ws.cell(r2, ci + 1).value or '').strip()
                                if v and v not in ('nan','NaN','None'):
                                    entry[col] = v
                        static_map[st] = entry
                    break
    return hdr_map, static_map

SUBTYPE_HEADER_ROW, SUBTYPE_MAP = _build_header_row_map()

def load_pv_list():
    pv = pd.read_excel(TEMPLATE_PATH, sheet_name='PV List')
    col = pv.columns[0]
    return [str(v).strip() for v in pv[col].dropna() if str(v).strip() not in ('nan','SubType')]

PV_LIST = load_pv_list()

def get_template_wb_for_subtype(subtype):
    from openpyxl import Workbook
    wb_src = load_workbook(TEMPLATE_PATH); ws_src = wb_src['PV Template']
    hdr_row = SUBTYPE_HEADER_ROW.get(subtype, 1)
    headers = [ws_src.cell(hdr_row, c).value for c in range(1, ws_src.max_column + 1)]
    while headers and headers[-1] is None: headers.pop()
    wb_new = Workbook(); ws_new = wb_new.active; ws_new.title = 'PV Template'
    for ci, h in enumerate(headers, 1):
        ws_new.cell(1, ci).value = h
    return wb_new, headers

# Default config
DEFAULT_CONFIG = {
    "brand_name":"", "brand_id":"", "biz_cat_id":"BCAT-139461", "biz_cat_name":"Footwear",
    "relationship":"Parent", "catalog_status":"ACTIVE", "status_remark":"Ready to Launch",
    "tax_master_status":"active", "gst_cgst":50, "gst_sgst":50, "gst_igst":0,
    "country_of_origin":"India", "product_condition":"Fresh", "manufacturing_year":"2026",
    "discovery_cat":"DISCAT-135542",
}
config = {k: v for k, v in DEFAULT_CONFIG.items()}

# ─── (your existing helpers — safe / detect_col / build_col_map / title_case_color /
#      merge_colors / extract_article / expand_size_range / build_set_details /
#      parse_lbh / derive_gender / make_title / make_internal_title / make_description /
#      DUMP_COL_HINTS / BASE_COL_HINTS / fill_template — all unchanged ) ───
# Paste your originals here verbatim. I'm omitting them for brevity.

# ── Routes ─────────────────────────────────────────────────────
@app.route('/')
@login_required
def index():
    return render_template('index.html', user_email=session.get('email'))

@app.route('/subtypes')
@login_required
def get_subtypes():
    return jsonify({'subtypes': PV_LIST})

@app.route('/config', methods=['GET'])
@login_required
def get_config(): return jsonify(config)

@app.route('/config', methods=['POST'])
@login_required
def update_config():
    global config
    config.update(request.json)
    write_log(session.get('email'), 'config_updated')
    return jsonify({'status': 'ok'})

@app.route('/logs')
@login_required
def get_logs():
    return jsonify({'logs': read_logs(500)})

@app.route('/download_template/<vertical>')
@login_required
def download_template(vertical):
    # Only Footwear unlocked
    if vertical.lower() != 'footwear':
        return jsonify({'error': 'This vertical is under development'}), 403
    write_log(session.get('email'), 'template_downloaded', vertical)
    return send_file(TEMPLATE_PATH, as_attachment=True,
                     download_name='Footwear_Template.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/detect_verticals', methods=['POST'])
@login_required
def detect_verticals():
    try:
        dump_file = request.files.get('dump')
        if not dump_file: return jsonify({'verticals': []})
        xl = pd.ExcelFile(io.BytesIO(dump_file.read()))
        frames = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        col_map = build_col_map(all_dump, DUMP_COL_HINTS)
        vert_col = col_map.get('vertical')
        if vert_col and vert_col in all_dump.columns:
            found = [str(v).strip() for v in all_dump[vert_col].dropna().unique()
                     if str(v).strip() not in ('nan','None','')]
            matched = [v for v in found if v in SUBTYPE_MAP]
            return jsonify({'verticals': matched, 'all_found': found})
        return jsonify({'verticals': [], 'all_found': []})
    except Exception as e:
        return jsonify({'verticals': [], 'error': str(e)})

@app.route('/process', methods=['POST'])
@login_required
def process():
    # ← keep your existing /process implementation verbatim,
    #    and add one line at the very start of the try block:
    #      user_email = session.get('email')
    #    and one line after grand_filled is known:
    #      write_log(user_email, 'catalog_generated',
    #                f'subtypes={subtypes} filled={grand_filled} skipped={len(all_skipped)}')
    ...

@app.route('/download/<token>')
@login_required
def download(token):
    if '..' in token or '/' in token or '\\' in token: return 'Invalid', 400
    path = os.path.join(tempfile.gettempdir(), token)
    if not os.path.exists(path): return 'File not found', 404
    fname = request.args.get('filename', 'filled_template.xlsx')
    write_log(session.get('email'), 'file_downloaded', fname)
    return send_file(path, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=False, port=5050)
