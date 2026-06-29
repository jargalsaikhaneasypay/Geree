from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, session
import psycopg2
from psycopg2.extras import RealDictCursor
import os
import tempfile
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
from io import BytesIO
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'contract-registry-secret-2026')

# -----------------------------------------------------------
# Users
# -----------------------------------------------------------
USERS = {
    'bolor-erdene@easypay.mn': generate_password_hash('Bolor2026!'),
    'uyanga@easypay.mn':       generate_password_hash('Uyanga2026!'),
    'myagmarsuren_lkh@easypay.mn': generate_password_hash('Myagmar2026!'),
}

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated

DEPARTMENTS = ['Борлуулалт', 'ХҮА', 'Салбар', 'Хөдөө орон нутаг', 'Томилолт']
STATUSES = ['Гэрээ', 'Зарагдсан', 'Нэр шилжүүлэг', 'Түрээс']
INSPECTION_RESULTS = [
    'Бүрэн',
    'Бүрдэл дутуу',
    'Салбар дээр архивлагдсан/бүрэн',
    'Салбар дээр архивлагдсан/дутуу',
    'Бүртгэл буруу',
    'Бүртгэл буруу/дутуу'
]

# -----------------------------------------------------------
# Database helpers
# -----------------------------------------------------------
def get_db():
    conn = psycopg2.connect(os.environ.get('DATABASE_URL', ''), sslmode='require')
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS contracts (
            id               SERIAL PRIMARY KEY,
            dd               INTEGER,
            merchant_name    TEXT,
            pos_serial       TEXT,
            merchant_number  TEXT,
            terminal_number  TEXT,
            status           TEXT,
            pos_issue_date   TEXT,
            phone            TEXT,
            merchant_type    TEXT,
            issued_by        TEXT,
            department       TEXT,
            expected_date    TEXT,
            received_date    TEXT,
            overdue_days     INTEGER,
            time_category    TEXT,
            first_inspection TEXT,
            description      TEXT,
            return_date      TEXT,
            last_inspection_date TEXT,
            created_at       TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    for new_col in ['status', 'return_date_2', 'return_date_3', 'return_date_4', 'return_date_5',
                    'is_inactive', 'inactive_reason', 'modified_by', 'scanned']:
        cur.execute(f"ALTER TABLE contracts ADD COLUMN IF NOT EXISTS {new_col} TEXT")
    conn.commit()
    cur.close()
    conn.close()


def calculate_overdue(expected_date, received_date):
    if not expected_date or not received_date:
        return None
    try:
        exp = datetime.strptime(expected_date, '%Y-%m-%d').date()
        rec = datetime.strptime(received_date, '%Y-%m-%d').date()
        return (rec - exp).days
    except Exception:
        return None


def get_time_category(overdue_days):
    if overdue_days is None:
        return 'Хугацаа хэтэрсэн'
    return 'Хугацаандаа' if overdue_days <= 0 else 'Хугацаа хэтэрсэн'


def row_to_dict(row):
    return dict(row) if row else None


# -----------------------------------------------------------
# Auth routes
# -----------------------------------------------------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user' in session:
        return redirect(url_for('index'))
    if request.method == 'POST':
        email    = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        if email in USERS and check_password_hash(USERS[email], password):
            session['user'] = email
            next_url = request.args.get('next') or url_for('index')
            return redirect('/' + next_url.lstrip('/'))
        flash('И-мэйл эсвэл нууц үг буруу байна.', 'error')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login'))


# -----------------------------------------------------------
# Routes – main list
# -----------------------------------------------------------
@app.route('/')
def root():
    return redirect(url_for('dashboard'))


@app.route('/list')
@login_required
def index():
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)

    search = request.args.get('search', '').strip()
    dept   = request.args.get('dept', '').strip()
    cat    = request.args.get('cat', '').strip()

    query  = 'SELECT * FROM contracts WHERE 1=1'
    params = []
    if search:
        query += ' AND (merchant_name LIKE %s OR pos_serial LIKE %s OR merchant_number LIKE %s OR terminal_number LIKE %s)'
        like = f'%{search}%'
        params.extend([like, like, like, like])
    if dept:
        query += ' AND department = %s'
        params.append(dept)
    if cat:
        query += ' AND time_category = %s'
        params.append(cat)

    query += ' ORDER BY dd'
    cur.execute(query, params)
    contracts = cur.fetchall()

    cur.execute('SELECT COUNT(*) FROM contracts')
    total = cur.fetchone()['count']
    cur.execute("SELECT COUNT(*) FROM contracts WHERE time_category='Хугацаандаа'")
    on_time = cur.fetchone()['count']
    cur.execute("SELECT COUNT(*) FROM contracts WHERE time_category='Хугацаа хэтэрсэн'")
    overdue = cur.fetchone()['count']

    cur.close()
    conn.close()

    return render_template(
        'index.html',
        contracts=contracts,
        departments=DEPARTMENTS,
        search=search, dept=dept, cat=cat,
        total=total, on_time=on_time, overdue=overdue
    )


# -----------------------------------------------------------
# Add new contract (manual)
# -----------------------------------------------------------
@app.route('/add', methods=['GET', 'POST'])
@login_required
def add():
    if request.method == 'POST':
        conn = get_db()
        cur  = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute('SELECT MAX(dd) AS m FROM contracts')
        last = cur.fetchone()
        dd   = (last['m'] or 0) + 1

        merchant_name        = request.form.get('merchant_name', '').strip()
        pos_serial           = request.form.get('pos_serial', '').strip()
        merchant_number      = request.form.get('merchant_number', '').strip()
        terminal_number      = request.form.get('terminal_number', '').strip()
        status               = request.form.get('status', '').strip()
        pos_issue_date       = request.form.get('pos_issue_date', '').strip()
        phone                = request.form.get('phone', '').strip()
        merchant_type        = request.form.get('merchant_type', '').strip()
        issued_by            = request.form.get('issued_by', '').strip()
        department           = request.form.get('department', '').strip()
        expected_date        = request.form.get('expected_date', '').strip()
        received_date        = request.form.get('received_date', '').strip()
        first_inspection     = request.form.get('first_inspection', '').strip()
        description          = request.form.get('description', '').strip()
        return_date          = request.form.get('return_date', '').strip()
        return_date_2        = request.form.get('return_date_2', '').strip()
        return_date_3        = request.form.get('return_date_3', '').strip()
        return_date_4        = request.form.get('return_date_4', '').strip()
        return_date_5        = request.form.get('return_date_5', '').strip()
        last_inspection_date = request.form.get('last_inspection_date', '').strip()
        scanned              = '1' if request.form.get('scanned') else ''

        overdue_days  = calculate_overdue(expected_date, received_date)
        time_category = get_time_category(overdue_days)

        if terminal_number:
            cur.execute('SELECT id FROM contracts WHERE terminal_number = %s', (terminal_number,))
            if cur.fetchone():
                flash(f'Терминалын дугаар давхацсан: {terminal_number}', 'error')
                cur.close()
                conn.close()
                return redirect(url_for('add'))

        cur.execute('''
            INSERT INTO contracts
            (dd, merchant_name, pos_serial, merchant_number, terminal_number,
             status, pos_issue_date, phone, merchant_type, issued_by, department,
             expected_date, received_date, overdue_days, time_category,
             first_inspection, description,
             return_date, return_date_2, return_date_3, return_date_4, return_date_5,
             last_inspection_date, scanned)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        ''', (dd, merchant_name, pos_serial, merchant_number, terminal_number,
              status, pos_issue_date, phone, merchant_type, issued_by, department,
              expected_date, received_date, overdue_days, time_category,
              first_inspection, description,
              return_date, return_date_2, return_date_3, return_date_4, return_date_5,
              last_inspection_date, scanned))
        conn.commit()
        cur.close()
        conn.close()

        flash('Бүртгэл амжилттай хийгдлээ!', 'success')
        return redirect(url_for('index'))

    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT MAX(dd) AS m FROM contracts')
    last    = cur.fetchone()
    next_dd = (last['m'] or 0) + 1
    cur.close()
    conn.close()
    return render_template('add.html', next_dd=next_dd,
                           departments=DEPARTMENTS,
                           statuses=STATUSES,
                           inspection_results=INSPECTION_RESULTS)


# -----------------------------------------------------------
# Edit existing contract
# -----------------------------------------------------------
@app.route('/edit/<int:cid>', methods=['GET', 'POST'])
@login_required
def edit(cid):
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT * FROM contracts WHERE id=%s', (cid,))
    contract = cur.fetchone()
    if not contract:
        flash('Бүртгэл олдсонгүй!', 'error')
        cur.close()
        conn.close()
        return redirect(url_for('index'))

    if request.method == 'POST':
        merchant_name        = request.form.get('merchant_name', '').strip()
        pos_serial           = request.form.get('pos_serial', '').strip()
        merchant_number      = request.form.get('merchant_number', '').strip()
        terminal_number      = request.form.get('terminal_number', '').strip()
        status               = request.form.get('status', '').strip()
        pos_issue_date       = request.form.get('pos_issue_date', '').strip()
        phone                = request.form.get('phone', '').strip()
        merchant_type        = request.form.get('merchant_type', '').strip()
        issued_by            = request.form.get('issued_by', '').strip()
        department           = request.form.get('department', '').strip()
        expected_date        = request.form.get('expected_date', '').strip()
        received_date        = request.form.get('received_date', '').strip()
        first_inspection     = request.form.get('first_inspection', '').strip()
        description          = request.form.get('description', '').strip()
        return_date          = request.form.get('return_date', '').strip()
        return_date_2        = request.form.get('return_date_2', '').strip()
        return_date_3        = request.form.get('return_date_3', '').strip()
        return_date_4        = request.form.get('return_date_4', '').strip()
        return_date_5        = request.form.get('return_date_5', '').strip()
        last_inspection_date = request.form.get('last_inspection_date', '').strip()
        is_inactive          = '1' if request.form.get('is_inactive') else ''
        scanned              = '1' if request.form.get('scanned') else ''
        inactive_reason      = request.form.get('inactive_reason', '').strip()

        overdue_days  = calculate_overdue(expected_date, received_date)
        time_category = get_time_category(overdue_days)

        if terminal_number:
            cur.execute('SELECT id FROM contracts WHERE terminal_number = %s AND id != %s', (terminal_number, cid))
            if cur.fetchone():
                flash(f'Терминалын дугаар давхацсан: {terminal_number}', 'error')
                cur.close()
                conn.close()
                return render_template('edit.html', c=contract,
                                       departments=DEPARTMENTS,
                                       statuses=STATUSES,
                                       inspection_results=INSPECTION_RESULTS)

        cur.execute('''
            UPDATE contracts SET
                merchant_name=%s, pos_serial=%s, merchant_number=%s, terminal_number=%s,
                status=%s, pos_issue_date=%s, phone=%s, merchant_type=%s, issued_by=%s,
                department=%s, expected_date=%s, received_date=%s,
                overdue_days=%s, time_category=%s,
                first_inspection=%s, description=%s,
                return_date=%s, return_date_2=%s, return_date_3=%s, return_date_4=%s, return_date_5=%s,
                last_inspection_date=%s, is_inactive=%s, inactive_reason=%s,
                scanned=%s, modified_by=%s
            WHERE id=%s
        ''', (merchant_name, pos_serial, merchant_number, terminal_number,
              status, pos_issue_date, phone, merchant_type, issued_by,
              department, expected_date, received_date,
              overdue_days, time_category,
              first_inspection, description,
              return_date, return_date_2, return_date_3, return_date_4, return_date_5,
              last_inspection_date, is_inactive, inactive_reason,
              scanned, session.get('user', ''), cid))
        conn.commit()
        cur.close()
        conn.close()

        flash('Бүртгэл амжилттай шинэчлэгдлээ!', 'success')
        return redirect(url_for('index'))

    cur.close()
    conn.close()
    return render_template('edit.html', c=contract,
                           departments=DEPARTMENTS,
                           statuses=STATUSES,
                           inspection_results=INSPECTION_RESULTS)


# -----------------------------------------------------------
# Delete
# -----------------------------------------------------------
@app.route('/delete/<int:cid>', methods=['POST'])
@login_required
def delete(cid):
    conn = get_db()
    cur  = conn.cursor()
    cur.execute('DELETE FROM contracts WHERE id=%s', (cid,))
    conn.commit()
    cur.close()
    conn.close()
    flash('Бүртгэл устгагдлаа!', 'info')
    return redirect(url_for('index'))


# -----------------------------------------------------------
# Excel import page (GET) + process (POST)
# -----------------------------------------------------------

HEADER_MAP = {
    # identity
    'д/д':                              'dd_col',
    'дугаар':                           'dd_col',
    '№':                                'dd_col',
    # section 1
    'посын мерчантын нэр':              'merchant_name',
    'мерчантын нэр':                    'merchant_name',
    'нэр':                              'merchant_name',
    'посын сериал':                     'pos_serial',
    'сериал':                           'pos_serial',
    'мерчантын дугаар':                 'merchant_number',
    'терминалын дугаар':                'terminal_number',
    'терминал':                         'terminal_number',
    'төлөв':                            'status',
    'гэрээний төлөв':                   'status',
    'пос гаргасан огноо':               'pos_issue_date',
    'утас':                             'phone',
    'мерчантын хэлбэр':                 'merchant_type',
    'хэлбэр':                           'merchant_type',
    'мерчант гаргасан ажилтан':         'issued_by',
    'ажилтан':                          'issued_by',
    # section 2
    'хэлтэс':                           'department',
    'гэрээ ирсэн байх ёстой огноо':     'expected_date',
    'гэрээ хүлээн авсан огноо':         'received_date',
    'хугацаа хэтэрсэн хоног':          'overdue_days',
    'хугацааны ангилал':                'time_category',
    'эхний хяналтаарх үр дүн':         'first_inspection',
    'тайлбар':                          'description',
    'буцаасан огноо':                   'return_date',
    'сүүлийн хяналтаар хүлээн авсан огноо': 'last_inspection_date',
}

FIELD_LABELS = {
    'merchant_name':        'Посын мерчантын нэр',
    'pos_serial':           'Посын сериал',
    'merchant_number':      'Мерчантын дугаар',
    'terminal_number':      'Терминалын дугаар',
    'status':               'Төлөв',
    'pos_issue_date':       'Пос гаргасан огноо',
    'phone':                'Утас',
    'merchant_type':        'Мерчантын хэлбэр',
    'issued_by':            'Мерчант гаргасан ажилтан',
    'department':           'Хэлтэс',
    'expected_date':        'Гэрээ ирсэн байх ёстой огноо',
    'received_date':        'Гэрээ хүлээн авсан огноо',
    'overdue_days':         'Хугацаа хэтэрсэн хоног',
    'time_category':        'Хугацааны ангилал',
    'first_inspection':     'Эхний хяналтаарх үр дүн',
    'description':          'Тайлбар',
    'return_date':          'Буцаасан огноо',
    'last_inspection_date': 'Сүүлийн хяналтаар хүлээн авсан огноо',
}

POSITIONAL_MAP = {
    0: 'dd_col',
    1: 'merchant_name',
    2: 'pos_serial',
    3: 'merchant_number',
    4: 'terminal_number',
    5: 'status',
    6: 'pos_issue_date',
    7: 'phone',
    8: 'merchant_type',
    9: 'issued_by',
}


def cell_val(v):
    if v is None:
        return ''
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, date):
        return v.isoformat()
    return str(v).strip()


@app.route('/import-page')
@login_required
def import_page():
    return render_template('import_page.html')


@app.route('/import', methods=['POST'])
@login_required
def import_excel():
    if 'file' not in request.files:
        flash('Файл сонгоогүй байна!', 'error')
        return redirect(url_for('import_page'))

    file = request.files['file']
    if not file or file.filename == '':
        flash('Файл сонгоогүй байна!', 'error')
        return redirect(url_for('import_page'))

    if not file.filename.lower().endswith('.xlsx'):
        flash('Зөвхөн Excel .xlsx файл оруулна уу!', 'error')
        return redirect(url_for('import_page'))

    try:
        # Save to a temp file so we can re-read it on confirm
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        file.save(tmp.name)
        tmp.close()

        wb = load_workbook(tmp.name, data_only=True)
        ws = wb.active
        first_row = [cell_val(c) for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        wb.close()

        has_header = any(h.lower().strip() in HEADER_MAP for h in first_row)

        # Build preview rows: one entry per Excel column
        preview = []
        for h in first_row:
            key = h.lower().strip()
            field = HEADER_MAP.get(key)
            if field == 'dd_col':
                field = None
            preview.append({
                'excel_col': h,
                'db_field':  field,
                'label':     FIELD_LABELS.get(field, '') if field else '',
            })

        session['import_tmp']        = tmp.name
        session['import_has_header'] = has_header
        session['import_filename']   = file.filename

        return render_template('import_preview.html',
                               preview=preview,
                               filename=file.filename)

    except Exception as e:
        flash(f'Файл уншихад алдаа гарлаа: {str(e)}', 'error')
        return redirect(url_for('import_page'))


def _run_import(tmp_path, has_header):
    wb = load_workbook(tmp_path, data_only=True)
    ws = wb.active

    first_row = [cell_val(c) for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    col_map = {}
    for idx, h in enumerate(first_row):
        key = h.lower().strip()
        if key in HEADER_MAP:
            field = HEADER_MAP[key]
            if field != 'dd_col':
                col_map[field] = idx

    if not has_header:
        for idx, field in POSITIONAL_MAP.items():
            if field != 'dd_col':
                col_map[field] = idx

    data_start_row = 2 if has_header else 1

    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT MAX(dd) AS m FROM contracts')
    dd = (cur.fetchone()['m'] or 0) + 1

    imported = 0
    skipped  = 0
    err_rows = []

    def get_field(row, field):
        idx = col_map.get(field)
        if idx is None or idx >= len(row):
            return ''
        return cell_val(row[idx])

    for r_idx, row in enumerate(ws.iter_rows(min_row=data_start_row, values_only=True), start=data_start_row):
        if not any(v for v in row if v not in (None, '')):
            continue
        try:
            merchant_name        = get_field(row, 'merchant_name')
            pos_serial           = get_field(row, 'pos_serial')
            merchant_number      = get_field(row, 'merchant_number')
            terminal_number      = get_field(row, 'terminal_number')
            status               = get_field(row, 'status')
            pos_issue_date       = get_field(row, 'pos_issue_date')
            phone                = get_field(row, 'phone')
            merchant_type        = get_field(row, 'merchant_type')
            issued_by            = get_field(row, 'issued_by')
            department           = get_field(row, 'department')
            expected_date        = get_field(row, 'expected_date')
            received_date        = get_field(row, 'received_date')
            first_inspection     = get_field(row, 'first_inspection')
            description          = get_field(row, 'description')
            return_date          = get_field(row, 'return_date')
            last_inspection_date = get_field(row, 'last_inspection_date')

            overdue_days_raw  = get_field(row, 'overdue_days')
            time_category_raw = get_field(row, 'time_category')

            if expected_date and received_date:
                try:
                    diff = (date.fromisoformat(received_date) - date.fromisoformat(expected_date)).days
                    overdue_days  = diff
                    time_category = 'Хугацаандаа' if diff <= 0 else 'Хугацаа хэтэрсэн'
                except Exception:
                    overdue_days  = overdue_days_raw or None
                    time_category = time_category_raw or 'Хугацаа хэтэрсэн'
            else:
                overdue_days  = overdue_days_raw or None
                time_category = time_category_raw or 'Хугацаа хэтэрсэн'

            if not merchant_name:
                skipped += 1
                continue

            if terminal_number:
                cur.execute('SELECT id FROM contracts WHERE terminal_number = %s', (terminal_number,))
                if cur.fetchone():
                    err_rows.append(f'Мөр {r_idx}: Терминалын дугаар давхацсан ({terminal_number})')
                    skipped += 1
                    continue

            cur.execute('''
                INSERT INTO contracts
                (dd, merchant_name, pos_serial, merchant_number, terminal_number,
                 status, pos_issue_date, phone, merchant_type, issued_by,
                 department, expected_date, received_date, overdue_days, time_category,
                 first_inspection, description, return_date, last_inspection_date)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ''', (dd, merchant_name, pos_serial, merchant_number, terminal_number,
                  status, pos_issue_date, phone, merchant_type, issued_by,
                  department, expected_date, received_date, overdue_days, time_category,
                  first_inspection, description, return_date, last_inspection_date))
            dd       += 1
            imported += 1

        except Exception as ex:
            err_rows.append(f'Мөр {r_idx}: {ex}')

    conn.commit()
    cur.close()
    conn.close()
    wb.close()
    return imported, skipped, err_rows


@app.route('/import-confirm', methods=['POST'])
@login_required
def import_confirm():
    tmp_path   = session.pop('import_tmp', None)
    has_header = session.pop('import_has_header', True)
    session.pop('import_filename', None)

    if not tmp_path or not os.path.exists(tmp_path):
        flash('Сесс дууссан байна. Файлыг дахин оруулна уу.', 'error')
        return redirect(url_for('import_page'))

    try:
        imported, _, err_rows = _run_import(tmp_path, has_header)
    except Exception as e:
        flash(f'Импортлоход алдаа гарлаа: {str(e)}', 'error')
        return redirect(url_for('import_page'))
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    if imported:
        flash(f'{imported} бүртгэл амжилттай импортлогдлоо!', 'success')
    for err in err_rows:
        flash(err, 'error')
    if not imported and not err_rows:
        flash('Импортлох мэдээлэл олдсонгүй. Файлын формат зөв эсэхийг шалгана уу.', 'error')

    return redirect(url_for('index'))


# -----------------------------------------------------------
# Download Excel template
# -----------------------------------------------------------
@app.route('/template')
def download_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Мерчант бүртгэл'

    headers = [
        'Д/д',
        'Посын мерчантын нэр',
        'Посын сериал',
        'Мерчантын дугаар',
        'Терминалын дугаар',
        'Төлөв',
        'Пос гаргасан огноо',
        'Утас',
        'Мерчантын хэлбэр',
        'Мерчант гаргасан ажилтан',
    ]
    col_widths = [6, 30, 18, 18, 18, 18, 20, 16, 22, 26]

    header_font  = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    header_fill  = PatternFill('solid', fgColor='1E3A5F')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    sample_fill  = PatternFill('solid', fgColor='F0F4FA')
    thin         = Side(style='thin', color='CCCCCC')
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.row_dimensions[1].height = 32
    for i, (h, w) in enumerate(zip(headers, col_widths), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = header_font; c.fill = header_fill
        c.alignment = header_align; c.border = border
        ws.column_dimensions[get_column_letter(i)].width = w

    samples = [
        [1, 'Дэлгүүрийн нэр ХХК', 'SN123456', 'M001234', 'T009876', 'Гэрээ',    '2026-01-15', '99001122', 'Бизнес',    'Болд Б'],
        [2, 'Жишээ ХХК',           'SN654321', 'M005678', 'T005432', 'Зарагдсан', '2026-02-20', '88112233', 'Хувиараа', 'Сарнай Д'],
    ]
    for r_idx, row in enumerate(samples, start=2):
        ws.row_dimensions[r_idx].height = 18
        for c_idx, val in enumerate(row, start=1):
            c = ws.cell(row=r_idx, column=c_idx, value=val)
            c.fill = sample_fill; c.border = border
            c.alignment = Alignment(horizontal='center', vertical='center')

    ws.freeze_panes = 'A2'

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name='merchant_template.xlsx')


# -----------------------------------------------------------
# Admin page
# -----------------------------------------------------------
@app.route('/admin')
@login_required
def admin():
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT * FROM contracts ORDER BY dd')
    contracts = cur.fetchall()
    cur.close()
    conn.close()

    total   = len(contracts)
    on_time = sum(1 for c in contracts if c['time_category'] == 'Хугацаандаа')
    overdue = sum(1 for c in contracts if c['time_category'] == 'Хугацаа хэтэрсэн')

    by_dept = {}
    for c in contracts:
        d = c['department'] or 'Тодорхойгүй'
        by_dept[d] = by_dept.get(d, 0) + 1

    by_result = {}
    for c in contracts:
        r = c['first_inspection'] or 'Тодорхойгүй'
        by_result[r] = by_result.get(r, 0) + 1

    return render_template('admin.html',
                           contracts=contracts,
                           total=total, on_time=on_time, overdue=overdue,
                           by_dept=by_dept, by_result=by_result)


# -----------------------------------------------------------
# Dashboard
# -----------------------------------------------------------
@app.route('/dashboard')
def dashboard():
    import calendar
    from datetime import date as dt_date
    today = dt_date.today()

    MONTHS_MN = ['1-р сар','2-р сар','3-р сар','4-р сар','5-р сар','6-р сар',
                 '7-р сар','8-р сар','9-р сар','10-р сар','11-р сар','12-р сар']

    period    = request.args.get('period', 'month')
    sel_month = request.args.get('sel_month', '')
    sel_q     = request.args.get('sel_q', '')
    sel_half  = request.args.get('sel_half', '')
    sel_year  = request.args.get('sel_year', '')
    dept      = request.args.get('dept', '')
    employee  = request.args.get('employee', '')
    date_from = request.args.get('date_from', '')
    date_to   = request.args.get('date_to', '')

    if period == 'day':
        d_from = today.isoformat()
        d_to   = today.isoformat()

    elif period == 'month':
        if sel_month:
            try:
                y, m = map(int, sel_month.split('-'))
                d_from = dt_date(y, m, 1).isoformat()
                d_to   = dt_date(y, m, calendar.monthrange(y, m)[1]).isoformat()
            except Exception:
                d_from = today.replace(day=1).isoformat()
                d_to   = today.isoformat()
        else:
            d_from = today.replace(day=1).isoformat()
            d_to   = today.isoformat()

    elif period == 'quarter':
        if sel_q:
            try:
                parts = sel_q.split('-Q')
                y = int(parts[0]); q = int(parts[1])
                m_start = (q - 1) * 3 + 1
                m_end   = m_start + 2
                d_from  = dt_date(y, m_start, 1).isoformat()
                d_to    = dt_date(y, m_end, calendar.monthrange(y, m_end)[1]).isoformat()
            except Exception:
                q_start = ((today.month - 1) // 3) * 3 + 1
                d_from  = today.replace(month=q_start, day=1).isoformat()
                d_to    = today.isoformat()
        else:
            q_start = ((today.month - 1) // 3) * 3 + 1
            d_from  = today.replace(month=q_start, day=1).isoformat()
            d_to    = today.isoformat()

    elif period == 'halfyear':
        if sel_half:
            try:
                parts = sel_half.split('-H')
                y = int(parts[0]); h = int(parts[1])
                m_start = 1 if h == 1 else 7
                m_end   = 6 if h == 1 else 12
                d_from  = dt_date(y, m_start, 1).isoformat()
                d_to    = dt_date(y, m_end, calendar.monthrange(y, m_end)[1]).isoformat()
            except Exception:
                h_start = 1 if today.month <= 6 else 7
                d_from  = today.replace(month=h_start, day=1).isoformat()
                d_to    = today.isoformat()
        else:
            h_start = 1 if today.month <= 6 else 7
            d_from  = today.replace(month=h_start, day=1).isoformat()
            d_to    = today.isoformat()

    elif period == 'year':
        y = int(sel_year) if sel_year else today.year
        d_from = dt_date(y, 1, 1).isoformat()
        d_to   = dt_date(y, 12, 31).isoformat()

    elif period == 'custom' and date_from and date_to:
        d_from = date_from
        d_to   = date_to

    else:
        d_from = today.replace(day=1).isoformat()
        d_to   = today.isoformat()
        period = 'month'

    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)

    cur.execute("SELECT DISTINCT issued_by FROM contracts WHERE issued_by IS NOT NULL AND issued_by != '' ORDER BY issued_by")
    all_employees = [r['issued_by'] for r in cur.fetchall()]

    base_cond   = "WHERE pos_issue_date BETWEEN %s AND %s"
    base_params = [d_from, d_to]
    if dept:
        base_cond  += " AND department = %s"
        base_params.append(dept)
    if employee:
        base_cond  += " AND issued_by = %s"
        base_params.append(employee)

    cur.execute(f"SELECT * FROM contracts {base_cond}", base_params)
    all_rows = cur.fetchall()
    inactive_count = sum(1 for r in all_rows if r['is_inactive'] == '1')
    rows = [r for r in all_rows if r['is_inactive'] != '1']

    total   = len(rows)
    on_time = sum(1 for r in rows if r['time_category'] == 'Хугацаандаа')
    overdue = total - on_time

    on_time_pct = round(on_time / total * 100, 1) if total else 0
    overdue_pct = round(overdue / total * 100, 1) if total else 0

    overdue_days_list = [r['overdue_days'] for r in rows
                         if r['overdue_days'] is not None and r['overdue_days'] > 0]
    avg_overdue = round(sum(overdue_days_list) / len(overdue_days_list), 1) if overdue_days_list else 0
    max_overdue = max(overdue_days_list) if overdue_days_list else 0

    COMPLETE_VALS = {'Бүрэн', 'Салбар дээр архивлагдсан/бүрэн'}
    complete      = sum(1 for r in rows if r['first_inspection'] in COMPLETE_VALS)
    returned      = sum(1 for r in rows if r['return_date'] or r['return_date_2'] or r['return_date_3'] or r['return_date_4'] or r['return_date_5'])
    not_received_rows = [r for r in rows if not r['received_date']]
    not_received = len(not_received_rows)

    nr_overdue_days = []
    for r in not_received_rows:
        if r['expected_date']:
            try:
                exp = datetime.strptime(r['expected_date'], '%Y-%m-%d').date()
                diff = (today - exp).days
                if diff > 0:
                    nr_overdue_days.append(diff)
            except Exception:
                pass

    nr_avg_overdue = round(sum(nr_overdue_days) / len(nr_overdue_days), 1) if nr_overdue_days else 0
    nr_max_overdue = max(nr_overdue_days) if nr_overdue_days else 0

    incomplete = total - complete
    complete_pct = round(complete / total * 100, 1) if total else 0
    returned_pct = round(returned / total * 100, 1) if total else 0

    insp_counts = {}
    for r in rows:
        k = r['first_inspection'] or 'Тодорхойгүй'
        insp_counts[k] = insp_counts.get(k, 0) + 1

    dept_stats = {}
    for r in rows:
        d = r['department'] or 'Тодорхойгүй'
        if d not in dept_stats:
            dept_stats[d] = {'total': 0, 'on_time': 0, 'overdue': 0, 'returned': 0}
        dept_stats[d]['total'] += 1
        if r['time_category'] == 'Хугацаандаа':
            dept_stats[d]['on_time'] += 1
        else:
            dept_stats[d]['overdue'] += 1
        if r['return_date'] or r['return_date_2'] or r['return_date_3'] or r['return_date_4'] or r['return_date_5']:
            dept_stats[d]['returned'] += 1

    trend = []
    for i in range(5, -1, -1):
        month = today.month - i
        year  = today.year
        while month <= 0:
            month += 12; year -= 1
        m_from = dt_date(year, month, 1).isoformat()
        m_to   = dt_date(year, month, calendar.monthrange(year, month)[1]).isoformat()
        cur.execute("SELECT COUNT(*) FROM contracts WHERE pos_issue_date BETWEEN %s AND %s", [m_from, m_to])
        t_total = cur.fetchone()['count']
        cur.execute("SELECT COUNT(*) FROM contracts WHERE pos_issue_date BETWEEN %s AND %s AND time_category='Хугацаандаа'", [m_from, m_to])
        t_on = cur.fetchone()['count']
        trend.append({'label': MONTHS_MN[month - 1], 'total': t_total, 'on_time': t_on, 'overdue': t_total - t_on})

    cur.execute("SELECT MIN(SUBSTRING(pos_issue_date, 1, 4)) AS min_year FROM contracts WHERE pos_issue_date IS NOT NULL AND pos_issue_date != ''")
    min_year_row = cur.fetchone()['min_year']
    min_year = int(min_year_row) if min_year_row else today.year
    years = list(range(today.year, min_year - 1, -1))

    cur.close()
    conn.close()

    return render_template('dashboard.html',
        period=period, dept=dept, employee=employee,
        sel_month=sel_month, sel_q=sel_q, sel_half=sel_half, sel_year=sel_year,
        date_from=date_from, date_to=date_to,
        d_from=d_from, d_to=d_to,
        departments=DEPARTMENTS, all_employees=all_employees,
        years=years, today=today.isoformat(),
        months_mn=MONTHS_MN,
        total=total, on_time=on_time, overdue=overdue,
        on_time_pct=on_time_pct, overdue_pct=overdue_pct,
        avg_overdue=avg_overdue, max_overdue=max_overdue,
        complete=complete, incomplete=incomplete, returned=returned,
        not_received=not_received, nr_avg_overdue=nr_avg_overdue, nr_max_overdue=nr_max_overdue,
        complete_pct=complete_pct, returned_pct=returned_pct,
        insp_counts=insp_counts, dept_stats=dept_stats, trend=trend,
        inactive_count=inactive_count,
    )


# -----------------------------------------------------------
# Export to Excel
# -----------------------------------------------------------
@app.route('/export')
@login_required
def export():
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT * FROM contracts ORDER BY dd')
    contracts = cur.fetchall()
    cur.close()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Гэрээ бүртгэл'

    header_font    = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    header_fill    = PatternFill('solid', fgColor='2C3E50')
    header_align   = Alignment(horizontal='center', vertical='center', wrap_text=True)
    center_align   = Alignment(horizontal='center', vertical='center')
    wrap_align     = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    on_time_fill   = PatternFill('solid', fgColor='D5F5E3')
    overdue_fill   = PatternFill('solid', fgColor='FADBD8')
    thin           = Side(style='thin', color='CCCCCC')
    border         = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = [
        'Д/д', 'Посын мерчантын нэр', 'Посын сериал', 'Мерчантын дугаар',
        'Терминалын дугаар', 'Пос гаргасан огноо', 'Утас', 'Мерчантын хэлбэр',
        'Мерчант гаргасан ажилтан', 'Хэлтэс',
        'Гэрээ ирсэн байх ёстой огноо', 'Гэрээ хүлээн авсан огноо',
        'Хугацаа хэтэрсэн хоног', 'Хугацааны ангилал',
        'Эхний хяналтаарх үр дүн', 'Тайлбар',
        'Буцаасан огноо', 'Сүүлийн хяналтаар хүлээн авсан огноо'
    ]
    col_widths = [6, 28, 18, 18, 18, 18, 16, 20, 24, 18,
                  22, 22, 20, 20, 30, 30, 18, 30]

    ws.row_dimensions[1].height = 36
    for i, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = border
        ws.column_dimensions[get_column_letter(i)].width = w

    for r_idx, row in enumerate(contracts, start=2):
        ws.row_dimensions[r_idx].height = 20
        values = [
            row['dd'], row['merchant_name'], row['pos_serial'],
            row['merchant_number'], row['terminal_number'], row['pos_issue_date'],
            row['phone'], row['merchant_type'], row['issued_by'],
            row['department'], row['expected_date'], row['received_date'],
            row['overdue_days'], row['time_category'],
            row['first_inspection'], row['description'],
            row['return_date'], row['last_inspection_date']
        ]
        is_overdue = row['time_category'] == 'Хугацаа хэтэрсэн'
        row_fill   = overdue_fill if is_overdue else on_time_fill

        for c_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border    = border
            cell.alignment = wrap_align if c_idx in (2, 15, 16) else center_align
            if c_idx == 14:
                cell.fill = row_fill

    ws.freeze_panes = 'A2'

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"contract_registry_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(output,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name=filename)


# -----------------------------------------------------------
# API – autocomplete last-used values
# -----------------------------------------------------------
@app.route('/api/autocomplete')
def autocomplete():
    field   = request.args.get('field', '')
    q       = request.args.get('q', '')
    allowed = {'merchant_name', 'merchant_type', 'issued_by', 'merchant_number', 'terminal_number'}
    if field not in allowed:
        return jsonify([])
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute(
        f"SELECT DISTINCT {field} FROM contracts WHERE {field} LIKE %s ORDER BY {field} LIMIT 10",
        (f'%{q}%',)
    )
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return jsonify([r[field] for r in rows if r[field]])


_db_ready = False
@app.before_request
def ensure_db():
    global _db_ready
    if not _db_ready:
        if not os.environ.get('DATABASE_URL'):
            raise RuntimeError('DATABASE_URL environment variable is not set in Render!')
        init_db()
        _db_ready = True

if __name__ == '__main__':
    init_db()
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    print('=' * 55)
    print('  Гэрээ бүртгэл хяналтын систем')
    print('  http://127.0.0.1:5000')
    print('=' * 55)
    app.run(debug=True, port=5000)
