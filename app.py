from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, session, g
import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2 import pool as _pg_pool
import os
import tempfile
import json
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


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated

DEPARTMENTS = ['Борлуулалт', 'ХҮТ', 'Салбар', 'Хө/орон нутаг']
STATUSES = ['Гэрээ', 'Зарагдсан', 'Нэр шилжүүлэг', 'Түрээс']
INSPECTION_RESULTS = [
    '1.Бүрэн',
    '2.Бүрдэл дутуу',
    '3.Салбар дээр архивлагдсан/бүрэн',
    '4.Бүртгэл буруу/дутуу',
    '5.Салбар дээр архивлагдсан/дутуу',
    '6.Татагдсан',
    '7.Татагдсан',
]

FIELD_LABELS = {
    'merchant_name':        'Мерчантын нэр',
    'pos_serial':           'Посын сериал',
    'merchant_number':      'Мерч. дугаар',
    'terminal_number':      'Терминал',
    'status':               'Төлөв',
    'pos_issue_date':       'Пос огноо',
    'phone':                'Утас',
    'merchant_type':        'Хэлбэр',
    'issued_by':            'Ажилтан',
    'department':           'Хэлтэс',
    'expected_date':        'Ирэх ёстой огноо',
    'received_date':        'Хүлээн авсан огноо',
    'first_inspection':     'Эхний хяналт',
    'description':          'Тайлбар',
    'return_date':          'Буцаасан огноо',
    'return_date_2':        'Буцаасан огноо 2',
    'return_date_3':        'Буцаасан огноо 3',
    'return_date_4':        'Буцаасан огноо 4',
    'return_date_5':        'Буцаасан огноо 5',
    'last_inspection_date': 'Сүүлийн хяналт',
    'is_inactive':          'Идэвхжил',
    'inactive_reason':      'Шалтгаан',
    'scanned':              'Scanned',
    'is_repaired':          'Засварлагдсан эсэх',
    'repaired_date':        'Засварлагдсан огноо',
}

ACTION_LABELS = {
    'create':        'Үүсгэсэн',
    'edit':          'Засварласан',
    'toggle_active': 'Идэвхжил',
    'delete':        'Устгасан',
}


def _log_history(cur, contract_id, user, action, merchant_name='', changes=None):
    if not changes:
        changes = [(None, None, None)]
    for field_name, old_val, new_val in changes:
        cur.execute('''
            INSERT INTO contract_history
                (contract_id, merchant_name, user_email, action, field_name, old_value, new_value)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        ''', (contract_id, merchant_name or '', user or '', action,
              field_name, old_val, new_val))


# -----------------------------------------------------------
# Database helpers  – per-worker connection pool
# -----------------------------------------------------------
_pool = None

def _db_url():
    url = os.environ.get('DATABASE_URL', '')
    if url.startswith('postgres://'):
        url = 'postgresql://' + url[len('postgres://'):]
    return url

def _get_pool():
    global _pool
    if _pool is None:
        url = _db_url()
        if url:
            _pool = _pg_pool.SimpleConnectionPool(
                1, 3, url,
                sslmode='require', connect_timeout=10,
                keepalives=1, keepalives_idle=30,
                keepalives_interval=10, keepalives_count=5,
            )
    return _pool


class _DBConn:
    """Thin wrapper so routes can call conn.close() without destroying
    the pooled connection — teardown_appcontext handles the real cleanup."""
    __slots__ = ('_c',)
    def __init__(self, c): self._c = c
    def cursor(self, **kw): return self._c.cursor(**kw)
    def commit(self):       return self._c.commit()
    def rollback(self):     return self._c.rollback()
    def close(self):        pass   # intentional no-op
    @property
    def closed(self):  return self._c.closed
    @property
    def status(self):  return self._c.status


def get_db():
    if 'db' not in g:
        pool = _get_pool()
        raw = None
        if pool:
            for _ in range(2):   # retry once if we get a dead connection
                raw = pool.getconn()
                if raw.closed == 0:
                    try:
                        raw.poll()
                        if raw.status != psycopg2.extensions.STATUS_READY:
                            raw.rollback()
                        break
                    except Exception:
                        pool.putconn(raw, close=True)
                        raw = None
                else:
                    pool.putconn(raw, close=True)
                    raw = None
        if raw is None:          # pool unavailable — fall back to direct
            raw = psycopg2.connect(_db_url(), sslmode='require',
                                   connect_timeout=10)
        g.db = raw
        g.db_wrapper = _DBConn(raw)
    return g.db_wrapper


@app.teardown_appcontext
def _teardown_db(_exc):
    raw = g.pop('db', None)
    if raw is None:
        return
    pool = _pool
    try:
        if raw.closed == 0 and raw.status != psycopg2.extensions.STATUS_READY:
            raw.rollback()
        if pool:
            pool.putconn(raw)   # return to pool — keeps connection alive
        else:
            raw.close()
    except Exception:
        try: raw.close()
        except Exception: pass


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
                    'is_inactive', 'inactive_reason', 'modified_by', 'scanned',
                    'is_repaired', 'repaired_date', 'uploaded_by', 'uploaded_date']:
        cur.execute(f"ALTER TABLE contracts ADD COLUMN IF NOT EXISTS {new_col} TEXT")
    conn.commit()

    cur.execute('''
        CREATE TABLE IF NOT EXISTS managers (
            id         SERIAL PRIMARY KEY,
            name       TEXT NOT NULL,
            department TEXT NOT NULL,
            start_date DATE NOT NULL DEFAULT '2000-01-01'
        )
    ''')
    # Migrate: add start_date column if it doesn't exist yet
    cur.execute("ALTER TABLE managers ADD COLUMN IF NOT EXISTS start_date DATE NOT NULL DEFAULT '2000-01-01'")
    # Migrate: drop old single-name unique constraint, add (name, start_date) unique instead
    cur.execute("ALTER TABLE managers DROP CONSTRAINT IF EXISTS managers_name_key")
    cur.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM pg_constraint WHERE conname = 'managers_name_start_key'
            ) THEN
                ALTER TABLE managers ADD CONSTRAINT managers_name_start_key UNIQUE (name, start_date);
            END IF;
        END $$
    """)
    conn.commit()

    # Run each migration in its own savepoint so a failure doesn't abort the whole transaction
    migrations = [
        "UPDATE contracts SET time_category='Хоосон', overdue_days=NULL WHERE (expected_date IS NULL OR expected_date='') AND (received_date IS NULL OR received_date='')",
        "UPDATE contracts SET time_category='Хоосон', overdue_days=0 WHERE expected_date IS NOT NULL AND expected_date!='' AND (received_date IS NULL OR received_date='') AND expected_date ~ '^\\d{4}-\\d{2}-\\d{2}$' AND expected_date::date >= CURRENT_DATE",
        "UPDATE contracts SET department='ХҮТ' WHERE department IN ('ХҮА','ХҮАжилтан')",
        "UPDATE managers  SET department='ХҮТ' WHERE department IN ('ХҮА','ХҮАжилтан')",
        "UPDATE contracts SET department='Хө/орон нутаг' WHERE department='Хөдөө орон нутаг'",
        "UPDATE managers  SET department='Хө/орон нутаг' WHERE department='Хөдөө орон нутаг'",
    ]
    for sql in migrations:
        try:
            cur.execute("SAVEPOINT mig")
            cur.execute(sql)
            cur.execute("RELEASE SAVEPOINT mig")
        except Exception:
            cur.execute("ROLLBACK TO SAVEPOINT mig")
    conn.commit()

    cur.execute('''
        CREATE TABLE IF NOT EXISTS contract_history (
            id            SERIAL PRIMARY KEY,
            contract_id   INTEGER,
            merchant_name TEXT,
            user_email    TEXT,
            changed_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            action        TEXT,
            field_name    TEXT,
            old_value     TEXT,
            new_value     TEXT
        )
    ''')
    cur.execute("CREATE INDEX IF NOT EXISTS idx_ch_contract ON contract_history(contract_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_ch_time ON contract_history(changed_at DESC)")
    conn.commit()

    cur.execute('''
        CREATE TABLE IF NOT EXISTS app_users (
            id            SERIAL PRIMARY KEY,
            email         TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL
        )
    ''')
    cur.execute('SELECT COUNT(*) FROM app_users')
    if cur.fetchone()[0] == 0:
        seed = [
            ('bolor-erdene@easypay.mn', 'Bolor2026!'),
            ('uyanga@easypay.mn',        'Uyanga2026!'),
            ('myagmarsuren_lkh@easypay.mn', 'Myagmar2026!'),
        ]
        for em, pw in seed:
            cur.execute(
                'INSERT INTO app_users (email, password_hash) VALUES (%s, %s) ON CONFLICT DO NOTHING',
                (em, generate_password_hash(pw, method='pbkdf2:sha256'))
            )
    conn.commit()
    cur.close()
    conn.close()


def calc_status(expected_date, received_date):
    """Return (overdue_days, time_category) based on the date combination.

    Both dates present:
      received <= expected → Хугацаандаа (on time or early)
      received >  expected → Хугацаа хэтэрсэн (late)
    expected only (received blank):
      expected >= today    → Хоосон (not yet due, not received)
      expected <  today    → Хугацаа хэтэрсэн (past due, never received)
    Neither date:          → Хоосон
    """
    today = date.today()
    if expected_date and received_date:
        try:
            exp = datetime.strptime(expected_date, '%Y-%m-%d').date()
            rec = datetime.strptime(received_date, '%Y-%m-%d').date()
            diff = (rec - exp).days
            cat  = 'Хугацаандаа' if diff <= 0 else 'Хугацаа хэтэрсэн'
            return diff, cat
        except Exception:
            return None, 'Хоосон'
    elif expected_date:
        try:
            exp  = datetime.strptime(expected_date, '%Y-%m-%d').date()
            diff = (today - exp).days
            if exp >= today:
                return 0, 'Хоосон'
            else:
                return diff, 'Хугацаа хэтэрсэн'
        except Exception:
            return None, 'Хоосон'
    else:
        return None, 'Хоосон'


def row_to_dict(row):
    return dict(row) if row else None


def sync_contracts_dept(cur, name):
    """Recalculate department for all contracts where issued_by = name.
    Single SQL UPDATE: best date-match, fallback to earliest entry."""
    cur.execute('''
        WITH dept_choice AS (
            SELECT
                c.id,
                COALESCE(
                    (SELECT m.department
                     FROM managers m
                     WHERE m.name = %(n)s
                       AND m.start_date IS NOT NULL
                       AND c.pos_issue_date ~ '^\\d{4}-\\d{2}-\\d{2}'
                       AND m.start_date <= c.pos_issue_date::date
                     ORDER BY m.start_date DESC
                     LIMIT 1),
                    (SELECT m.department
                     FROM managers m
                     WHERE m.name = %(n)s
                     ORDER BY m.start_date ASC NULLS LAST
                     LIMIT 1)
                ) AS new_dept
            FROM contracts c
            WHERE c.issued_by = %(n)s
        )
        UPDATE contracts
        SET department = dept_choice.new_dept
        FROM dept_choice
        WHERE contracts.id = dept_choice.id
          AND dept_choice.new_dept IS NOT NULL
    ''', {'n': name})
    return cur.rowcount


# -----------------------------------------------------------
# Auth routes
# -----------------------------------------------------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user' in session:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        email    = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        conn = get_db()
        cur  = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute('SELECT password_hash FROM app_users WHERE email = %s', (email,))
        user = cur.fetchone()
        cur.close()
        conn.close()
        if user and check_password_hash(user['password_hash'], password):
            session['user'] = email
            next_url = request.args.get('next') or url_for('dashboard')
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
@app.route('/health')
def health():
    return 'ok', 200

@app.route('/ping')
def ping():
    try:
        conn = get_db()
        cur  = conn.cursor()
        cur.execute('SELECT 1')
        cur.close()
        return 'OK', 200
    except Exception:
        return 'DB error', 500


@app.route('/db-status')
def db_status():
    import html as _html
    lines = []
    lines.append(f'_db_ready: {_db_ready}')
    lines.append(f'_db_init_error: {_db_init_error or "none"}')
    db_url = os.environ.get('DATABASE_URL', '')
    lines.append(f'DATABASE_URL set: {bool(db_url)}')
    if db_url:
        lines.append(f'DATABASE_URL prefix: {db_url[:20]}...')
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute('SELECT COUNT(*) FROM contracts')
        cnt = cur.fetchone()[0]
        cur.close()
        conn.close()
        lines.append(f'DB connection: OK  (contracts rows: {cnt})')
    except Exception as ex:
        lines.append(f'DB connection: FAILED — {ex}')
    body = _html.escape('\n'.join(lines))
    return f'<pre style="padding:20px;font-size:13px">{body}</pre>', 200

@app.route('/')
def root():
    return redirect(url_for('dashboard'))


@app.route('/list')
@login_required
def index():
    import calendar as _cal
    today = date.today()

    search    = request.args.get('search', '').strip()
    dept      = request.args.get('dept', '').strip()
    cat       = request.args.get('cat', '').strip()
    period    = request.args.get('period', '').strip()
    sel_month = request.args.get('sel_month', '').strip()
    sel_q     = request.args.get('sel_q', '').strip()
    sel_half  = request.args.get('sel_half', '').strip()
    sel_year  = request.args.get('sel_year', '').strip()
    date_from = request.args.get('date_from', '').strip()
    date_to   = request.args.get('date_to', '').strip()

    d_from = d_to = None
    if period == 'day':
        d_from = d_to = today.isoformat()
    elif period == 'month':
        if sel_month:
            try:
                y, m = map(int, sel_month.split('-'))
                d_from = date(y, m, 1).isoformat()
                d_to   = date(y, m, _cal.monthrange(y, m)[1]).isoformat()
            except Exception:
                d_from = today.replace(day=1).isoformat(); d_to = today.isoformat()
        else:
            d_from = today.replace(day=1).isoformat(); d_to = today.isoformat()
    elif period == 'quarter':
        if sel_q:
            try:
                y, q = int(sel_q.split('-Q')[0]), int(sel_q.split('-Q')[1])
                ms = (q - 1) * 3 + 1; me = ms + 2
                d_from = date(y, ms, 1).isoformat()
                d_to   = date(y, me, _cal.monthrange(y, me)[1]).isoformat()
            except Exception:
                qs = ((today.month - 1) // 3) * 3 + 1
                d_from = today.replace(month=qs, day=1).isoformat(); d_to = today.isoformat()
        else:
            qs = ((today.month - 1) // 3) * 3 + 1
            d_from = today.replace(month=qs, day=1).isoformat(); d_to = today.isoformat()
    elif period == 'halfyear':
        if sel_half:
            try:
                y, h = int(sel_half.split('-H')[0]), int(sel_half.split('-H')[1])
                ms = 1 if h == 1 else 7; me = 6 if h == 1 else 12
                d_from = date(y, ms, 1).isoformat()
                d_to   = date(y, me, _cal.monthrange(y, me)[1]).isoformat()
            except Exception:
                hs = 1 if today.month <= 6 else 7
                d_from = today.replace(month=hs, day=1).isoformat(); d_to = today.isoformat()
        else:
            hs = 1 if today.month <= 6 else 7
            d_from = today.replace(month=hs, day=1).isoformat(); d_to = today.isoformat()
    elif period == 'year':
        y = int(sel_year) if sel_year else today.year
        d_from = date(y, 1, 1).isoformat(); d_to = date(y, 12, 31).isoformat()
    elif period == 'custom' and date_from and date_to:
        d_from = date_from; d_to = date_to

    MONTHS_MN = ['1-р сар','2-р сар','3-р сар','4-р сар','5-р сар','6-р сар',
                 '7-р сар','8-р сар','9-р сар','10-р сар','11-р сар','12-р сар']

    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)

    # Year range + global stats in one query
    cur.execute("""
        SELECT
            MIN(SUBSTRING(pos_issue_date,1,4)) AS min_yr,
            COUNT(*) AS total,
            SUM(CASE WHEN time_category='Хугацаандаа' THEN 1 ELSE 0 END) AS on_time,
            SUM(CASE WHEN time_category='Хугацаа хэтэрсэн' THEN 1 ELSE 0 END) AS overdue
        FROM contracts
        WHERE is_inactive IS DISTINCT FROM '1'
    """)
    stats = cur.fetchone()
    min_yr   = stats['min_yr']
    total    = int(stats['total'] or 0)
    on_time  = int(stats['on_time'] or 0)
    overdue  = int(stats['overdue'] or 0)
    min_year = int(min_yr) if min_yr else today.year
    years = list(range(today.year, min_year - 1, -1))

    query  = 'SELECT * FROM contracts WHERE 1=1'
    params = []
    if d_from and d_to:
        query += ' AND pos_issue_date BETWEEN %s AND %s'
        params.extend([d_from, d_to])
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

    cur.close()
    conn.close()

    return render_template(
        'index.html',
        contracts=contracts,
        departments=DEPARTMENTS,
        search=search, dept=dept, cat=cat,
        period=period, sel_month=sel_month, sel_q=sel_q,
        sel_half=sel_half, sel_year=sel_year,
        date_from=date_from, date_to=date_to,
        d_from=d_from or '', d_to=d_to or '',
        years=years, months_mn=MONTHS_MN,
        today=today.isoformat(),
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
        is_repaired          = 'Тийм' if request.form.get('is_repaired') else ''
        repaired_date        = date.today().isoformat() if is_repaired == 'Тийм' else ''

        overdue_days, time_category = calc_status(expected_date, received_date)

        if terminal_number:
            cur.execute("SELECT id FROM contracts WHERE terminal_number = %s AND COALESCE(status,'') = %s", (terminal_number, status or ''))
            if cur.fetchone():
                flash(f'Терминалын дугаар давхацсан: {terminal_number} ({status})', 'error')
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
             last_inspection_date, scanned, is_repaired, repaired_date)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        ''', (dd, merchant_name, pos_serial, merchant_number, terminal_number,
              status, pos_issue_date, phone, merchant_type, issued_by, department,
              expected_date, received_date, overdue_days, time_category,
              first_inspection, description,
              return_date, return_date_2, return_date_3, return_date_4, return_date_5,
              last_inspection_date, scanned, is_repaired, repaired_date))
        cur.execute("SELECT lastval()")
        new_cid = cur.fetchone()['lastval']
        _log_history(cur, new_cid, session.get('user', ''), 'create', merchant_name)
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
        is_repaired          = 'Тийм' if request.form.get('is_repaired') else ''
        existing_repaired    = (contract.get('repaired_date') or '').strip()
        if is_repaired == 'Тийм':
            repaired_date = existing_repaired if existing_repaired else date.today().isoformat()
        else:
            repaired_date = ''

        overdue_days, time_category = calc_status(expected_date, received_date)

        if terminal_number:
            cur.execute("SELECT id FROM contracts WHERE terminal_number = %s AND COALESCE(status,'') = %s AND id != %s", (terminal_number, status or '', cid))
            if cur.fetchone():
                flash(f'Терминалын дугаар давхацсан: {terminal_number} ({status})', 'error')
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
                scanned=%s, modified_by=%s,
                is_repaired=%s, repaired_date=%s
            WHERE id=%s
        ''', (merchant_name, pos_serial, merchant_number, terminal_number,
              status, pos_issue_date, phone, merchant_type, issued_by,
              department, expected_date, received_date,
              overdue_days, time_category,
              first_inspection, description,
              return_date, return_date_2, return_date_3, return_date_4, return_date_5,
              last_inspection_date, is_inactive, inactive_reason,
              scanned, session.get('user', ''),
              is_repaired, repaired_date, cid))

        _h_fields = [
            ('merchant_name', merchant_name), ('pos_serial', pos_serial),
            ('merchant_number', merchant_number), ('terminal_number', terminal_number),
            ('status', status), ('pos_issue_date', pos_issue_date),
            ('phone', phone), ('merchant_type', merchant_type),
            ('issued_by', issued_by), ('department', department),
            ('expected_date', expected_date), ('received_date', received_date),
            ('first_inspection', first_inspection), ('description', description),
            ('return_date', return_date), ('return_date_2', return_date_2),
            ('return_date_3', return_date_3), ('return_date_4', return_date_4),
            ('return_date_5', return_date_5), ('last_inspection_date', last_inspection_date),
            ('is_inactive', is_inactive), ('inactive_reason', inactive_reason),
            ('scanned', scanned), ('is_repaired', is_repaired), ('repaired_date', repaired_date),
        ]
        _h_changes = [(f, str(contract.get(f) or ''), str(v or ''))
                      for f, v in _h_fields if str(contract.get(f) or '') != str(v or '')]
        if _h_changes:
            _log_history(cur, cid, session.get('user', ''), 'edit',
                         contract.get('merchant_name', ''), _h_changes)
        conn.commit()
        cur.close()
        conn.close()

        flash('Бүртгэл амжилттай шинэчлэгдлээ!', 'success')
        next_url = request.form.get('next', '').strip()
        return redirect(next_url if next_url else url_for('index'))

    cur.close()
    conn.close()
    return render_template('edit.html', c=contract,
                           departments=DEPARTMENTS,
                           statuses=STATUSES,
                           inspection_results=INSPECTION_RESULTS)


# -----------------------------------------------------------
# Delete
# -----------------------------------------------------------
@app.route('/toggle-active/<int:cid>', methods=['POST'])
@login_required
def toggle_active(cid):
    from flask import jsonify
    conn = get_db()
    cur  = conn.cursor()
    cur.execute("SELECT is_inactive, merchant_name FROM contracts WHERE id=%s", (cid,))
    row = cur.fetchone()
    if not row:
        cur.close()
        return jsonify(ok=False, error='not found'), 404
    currently_inactive = (row[0] == '1')
    mname = (row[1] or '') if len(row) > 1 else ''
    new_val = '' if currently_inactive else '1'
    cur.execute("UPDATE contracts SET is_inactive=%s WHERE id=%s", (new_val, cid))
    _log_history(cur, cid, session.get('user', ''), 'toggle_active', mname, [
        ('is_inactive', '1' if currently_inactive else '', new_val),
    ])
    conn.commit()
    cur.close()
    return jsonify(ok=True, is_active=(new_val != '1'))


@app.route('/delete/<int:cid>', methods=['POST'])
@login_required
def delete(cid):
    conn = get_db()
    cur  = conn.cursor()
    cur.execute('SELECT merchant_name FROM contracts WHERE id=%s', (cid,))
    mname_row = cur.fetchone()
    mname = (mname_row[0] if mname_row else '') or ''
    cur.execute('DELETE FROM contracts WHERE id=%s', (cid,))
    _log_history(cur, cid, session.get('user', ''), 'delete', mname)
    conn.commit()
    cur.close()
    conn.close()
    flash('Бүртгэл устгагдлаа!', 'info')
    return redirect(url_for('index'))


# -----------------------------------------------------------
# Change history
# -----------------------------------------------------------

@app.route('/history')
@login_required
def history_view():
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    limit = min(int(request.args.get('limit', 300)), 2000)
    cur.execute('''
        SELECT id, contract_id, merchant_name, user_email, changed_at,
               action, field_name, old_value, new_value
        FROM contract_history
        ORDER BY changed_at DESC, id DESC
        LIMIT %s
    ''', (limit,))
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('history.html', rows=rows,
                           field_labels=FIELD_LABELS, action_labels=ACTION_LABELS,
                           limit=limit, contract_id=None, merchant_name=None)


@app.route('/history/<int:cid>')
@login_required
def contract_history(cid):
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('''
        SELECT id, contract_id, merchant_name, user_email, changed_at,
               action, field_name, old_value, new_value
        FROM contract_history
        WHERE contract_id = %s
        ORDER BY changed_at DESC, id DESC
    ''', (cid,))
    rows = cur.fetchall()
    mname = rows[0]['merchant_name'] if rows else f'#{cid}'
    cur.close()
    conn.close()
    return render_template('history.html', rows=rows,
                           field_labels=FIELD_LABELS, action_labels=ACTION_LABELS,
                           limit=None, contract_id=cid, merchant_name=mname)


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


@app.route('/import-confirm', methods=['POST'])
@login_required
def import_confirm():
    # Mapping sent as JSON: {"col_idx": "db_field", ...}
    try:
        col_map_raw = json.loads(request.form.get('mapping', '{}'))
        col_map = {v: int(k) for k, v in col_map_raw.items() if v}
    except Exception:
        flash('Баганын харгалзаа алдаатай байна.', 'error')
        return redirect(url_for('import_page'))

    if 'file' not in request.files or request.files['file'].filename == '':
        flash('Файл олдсонгүй.', 'error')
        return redirect(url_for('import_page'))

    file = request.files['file']
    sheet_name = request.form.get('sheet_name', '').strip()
    try:
        wb  = load_workbook(file, data_only=True)
        ws  = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        conn = get_db()
        cur  = conn.cursor(cursor_factory=RealDictCursor)

        # Pre-load max dd and all existing terminal numbers in one go
        cur.execute('SELECT MAX(dd) AS m FROM contracts')
        dd = (cur.fetchone()['m'] or 0) + 1
        cur.execute("SELECT terminal_number, COALESCE(status,'') AS status FROM contracts WHERE terminal_number IS NOT NULL AND terminal_number != ''")
        existing_terminals = {(r['terminal_number'], r['status']) for r in cur.fetchall()}

        # Pre-load manager history for date-based department lookup
        cur.execute('SELECT name, department, start_date FROM managers ORDER BY name, start_date')
        mgr_history = {}
        for r in cur.fetchall():
            mgr_history.setdefault(r['name'], []).append((r['start_date'], r['department']))

        def get_dept_for_date(manager_name, pos_date_str):
            if not manager_name or manager_name not in mgr_history:
                return ''
            try:
                pos_d = datetime.strptime(pos_date_str[:10], '%Y-%m-%d').date() if pos_date_str else date.today()
            except Exception:
                pos_d = date.today()
            best_dept, best_start = '', None
            for start_d, dept in mgr_history[manager_name]:
                if start_d <= pos_d and (best_start is None or start_d > best_start):
                    best_start, best_dept = start_d, dept
            return best_dept

        imported = 0
        skipped  = 0
        err_rows = []
        batch    = []

        def gf(row, field):
            idx = col_map.get(field)
            if idx is None or idx >= len(row):
                return ''
            return cell_val(row[idx])

        for r_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not any(v for v in row if v not in (None, '')):
                continue
            try:
                merchant_name        = gf(row, 'merchant_name')
                if not merchant_name:
                    skipped += 1
                    continue

                pos_serial           = gf(row, 'pos_serial')
                merchant_number      = gf(row, 'merchant_number')
                terminal_number      = gf(row, 'terminal_number')
                status               = gf(row, 'status')
                pos_issue_date       = gf(row, 'pos_issue_date')
                phone                = gf(row, 'phone')
                merchant_type        = gf(row, 'merchant_type')
                issued_by            = gf(row, 'issued_by')
                department           = gf(row, 'department')

                # Auto-fill department from manager history based on pos_issue_date
                if issued_by and not department:
                    department = get_dept_for_date(issued_by, pos_issue_date)

                expected_date        = gf(row, 'expected_date')
                received_date        = gf(row, 'received_date')
                first_inspection     = gf(row, 'first_inspection')
                description          = gf(row, 'description')
                return_date          = gf(row, 'return_date')
                last_inspection_date = gf(row, 'last_inspection_date')

                overdue_days, time_category = calc_status(expected_date, received_date)

                if terminal_number and (terminal_number, status or '') in existing_terminals:
                    err_rows.append(f'Мөр {r_idx}: Терминалын дугаар давхацсан ({terminal_number}, {status})')
                    skipped += 1
                    continue

                batch.append((dd, merchant_name, pos_serial, merchant_number, terminal_number,
                              status, pos_issue_date, phone, merchant_type, issued_by,
                              department, expected_date, received_date, overdue_days, time_category,
                              first_inspection, description, return_date, last_inspection_date,
                              session.get('user', ''), date.today().isoformat()))
                if terminal_number:
                    existing_terminals.add((terminal_number, status or ''))
                dd += 1
                imported += 1

            except Exception as ex:
                err_rows.append(f'Мөр {r_idx}: {ex}')

        wb.close()

        if batch:
            from psycopg2.extras import execute_values
            execute_values(cur, '''
                INSERT INTO contracts
                (dd, merchant_name, pos_serial, merchant_number, terminal_number,
                 status, pos_issue_date, phone, merchant_type, issued_by,
                 department, expected_date, received_date, overdue_days, time_category,
                 first_inspection, description, return_date, last_inspection_date,
                 uploaded_by, uploaded_date)
                VALUES %s
            ''', batch)

        conn.commit()
        cur.close()
        conn.close()

    except Exception as e:
        flash(f'Файл уншихад алдаа гарлаа: {str(e)}', 'error')
        return redirect(url_for('import_page'))

    if imported:
        flash(f'{imported} бүртгэл амжилттай импортлогдлоо!', 'success')
    # Cap error flashes to avoid overflowing the session cookie
    for err in err_rows[:20]:
        flash(err, 'error')
    if len(err_rows) > 20:
        flash(f'... болон өөр {len(err_rows) - 20} алдаа', 'error')
    if not imported and not err_rows:
        flash('Импортлох мэдээлэл олдсонгүй.', 'error')

    return redirect(url_for('index'))

# -----------------------------------------------------------
# Manager → Department table (CRUD + API)
# -----------------------------------------------------------
@app.route('/managers', methods=['GET'])
@login_required
def managers_list():
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT * FROM managers ORDER BY name, start_date')
    mgrs = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('managers.html', managers=mgrs, departments=DEPARTMENTS)


@app.route('/managers/add', methods=['POST'])
@login_required
def managers_add():
    name = request.form.get('name', '').strip()
    dept = request.form.get('department', '').strip()
    start_date = request.form.get('start_date', '').strip() or '2000-01-01'
    if not name or not dept:
        flash('Нэр болон хэлтэс шаардлагатай.', 'error')
        return redirect(url_for('managers_list'))
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute('INSERT INTO managers (name, department, start_date) VALUES (%s, %s, %s)', (name, dept, start_date))
        n = sync_contracts_dept(cur, name)
        conn.commit()
        cur.close()
        conn.close()
        flash(f'"{name}" нэмэгдлээ.', 'success')
        if n:
            flash(f'{n} гэрээний хэлтэс автоматаар шинэчлэгдлээ.', 'success')
    except Exception as e:
        flash(f'Алдаа: {e}', 'error')
    return redirect(url_for('managers_list'))


@app.route('/managers/edit/<int:mid>', methods=['POST'])
@login_required
def managers_edit(mid):
    name = request.form.get('name', '').strip()
    dept = request.form.get('department', '').strip()
    start_date = request.form.get('start_date', '').strip() or '2000-01-01'
    if not name or not dept:
        flash('Нэр болон хэлтэс шаардлагатай.', 'error')
        return redirect(url_for('managers_list'))
    try:
        conn = get_db()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute('SELECT name FROM managers WHERE id = %s', (mid,))
        old = cur.fetchone()
        old_name = old['name'] if old else name
        cur.execute('UPDATE managers SET name = %s, department = %s, start_date = %s WHERE id = %s',
                    (name, dept, start_date, mid))
        names = set([name])
        if old_name != name:
            names.add(old_name)
        n = sum(sync_contracts_dept(cur, nm) for nm in names)
        conn.commit()
        cur.close()
        conn.close()
        flash('Өөрчлөлт хадгалагдлаа.', 'success')
        if n:
            flash(f'{n} гэрээний хэлтэс автоматаар шинэчлэгдлээ.', 'success')
    except Exception as e:
        flash(f'Алдаа: {e}', 'error')
    return redirect(url_for('managers_list'))


@app.route('/managers/delete/<int:mid>', methods=['POST'])
@login_required
def managers_delete(mid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('DELETE FROM managers WHERE id = %s', (mid,))
    conn.commit()
    cur.close()
    conn.close()
    flash('Устгагдлаа.', 'success')
    return redirect(url_for('managers_list'))


# -----------------------------------------------------------
# Routes – app user management (login accounts)
# -----------------------------------------------------------
@app.route('/users')
@login_required
def users_list():
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT id, email FROM app_users ORDER BY email')
    users = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('users.html', users=users)


@app.route('/users/add', methods=['POST'])
@login_required
def users_add():
    email    = request.form.get('email', '').strip().lower()
    password = request.form.get('password', '').strip()
    if not email or not password:
        flash('И-мэйл болон нууц үг шаардлагатай.', 'error')
        return redirect(url_for('users_list'))
    if len(password) < 6:
        flash('Нууц үг хамгийн багадаа 6 тэмдэгт байх ёстой.', 'error')
        return redirect(url_for('users_list'))
    try:
        conn = get_db()
        cur  = conn.cursor()
        cur.execute(
            'INSERT INTO app_users (email, password_hash) VALUES (%s, %s)',
            (email, generate_password_hash(password, method='pbkdf2:sha256'))
        )
        conn.commit()
        cur.close()
        conn.close()
        flash(f'"{email}" нэмэгдлээ.', 'success')
    except Exception as e:
        flash(f'Алдаа: {e}', 'error')
    return redirect(url_for('users_list'))


@app.route('/users/change-password/<int:uid>', methods=['POST'])
@login_required
def users_change_password(uid):
    password = request.form.get('password', '').strip()
    if len(password) < 6:
        flash('Нууц үг хамгийн багадаа 6 тэмдэгт байх ёстой.', 'error')
        return redirect(url_for('users_list'))
    conn = get_db()
    cur  = conn.cursor()
    cur.execute(
        'UPDATE app_users SET password_hash = %s WHERE id = %s',
        (generate_password_hash(password, method='pbkdf2:sha256'), uid)
    )
    conn.commit()
    cur.close()
    conn.close()
    flash('Нууц үг шинэчлэгдлээ.', 'success')
    return redirect(url_for('users_list'))


@app.route('/users/delete/<int:uid>', methods=['POST'])
@login_required
def users_delete(uid):
    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute('SELECT COUNT(*) AS cnt FROM app_users')
    if cur.fetchone()['cnt'] <= 1:
        cur.close()
        conn.close()
        flash('Хамгийн багадаа нэг хэрэглэгч байх ёстой.', 'error')
        return redirect(url_for('users_list'))
    cur.execute('SELECT email FROM app_users WHERE id = %s', (uid,))
    row = cur.fetchone()
    if row and row['email'] == session.get('user'):
        cur.close()
        conn.close()
        flash('Өөрийгөө устгах боломжгүй.', 'error')
        return redirect(url_for('users_list'))
    cur2 = conn.cursor()
    cur2.execute('DELETE FROM app_users WHERE id = %s', (uid,))
    conn.commit()
    cur.close()
    cur2.close()
    conn.close()
    flash('Хэрэглэгч устгагдлаа.', 'success')
    return redirect(url_for('users_list'))


@app.route('/managers/sync-all-depts', methods=['POST'])
@login_required
def managers_sync_all_depts():
    try:
        conn = get_db()
        cur = conn.cursor()
        # Single bulk UPDATE for all managers at once — no Python loop, no timeout risk
        cur.execute('''
            WITH dept_choice AS (
                SELECT
                    c.id,
                    COALESCE(
                        (SELECT m.department
                         FROM managers m
                         WHERE m.name = c.issued_by
                           AND m.start_date IS NOT NULL
                           AND c.pos_issue_date ~ '^\\d{4}-\\d{2}-\\d{2}'
                           AND m.start_date <= c.pos_issue_date::date
                         ORDER BY m.start_date DESC
                         LIMIT 1),
                        (SELECT m.department
                         FROM managers m
                         WHERE m.name = c.issued_by
                         ORDER BY m.start_date ASC NULLS LAST
                         LIMIT 1)
                    ) AS new_dept
                FROM contracts c
                WHERE c.issued_by IN (
                    SELECT DISTINCT name FROM managers WHERE name IS NOT NULL
                )
            )
            UPDATE contracts
            SET department = dept_choice.new_dept
            FROM dept_choice
            WHERE contracts.id = dept_choice.id
              AND dept_choice.new_dept IS NOT NULL
        ''')
        total = cur.rowcount
        conn.commit()
        cur.close()
        conn.close()
        flash(f'Нийт {total} гэрээний хэлтэс шинэчлэгдлээ.', 'success')
    except Exception as e:
        import traceback
        flash(f'Алдаа: {e} — {traceback.format_exc()[-300:]}', 'error')
    return redirect(url_for('managers_list'))


@app.route('/api/manager-dept')
@login_required
def api_manager_dept():
    name = request.args.get('name', '').strip()
    pos_date = request.args.get('pos_date', '').strip()
    if not name:
        return jsonify({'department': ''})
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    if pos_date:
        cur.execute('''
            SELECT department FROM managers
            WHERE name = %s AND start_date <= %s
            ORDER BY start_date DESC LIMIT 1
        ''', (name, pos_date))
    else:
        cur.execute('SELECT department FROM managers WHERE name = %s ORDER BY start_date DESC LIMIT 1', (name,))
    row = cur.fetchone()
    cur.close()
    conn.close()
    return jsonify({'department': row['department'] if row else ''})


#------------------------------------------------------------
# -----------------------------------------------------------
# Download Excel template------------------------------------
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
    import calendar as _cal
    today = date.today()

    MONTHS_MN = ['1-р сар','2-р сар','3-р сар','4-р сар','5-р сар','6-р сар',
                 '7-р сар','8-р сар','9-р сар','10-р сар','11-р сар','12-р сар']

    period    = request.args.get('period', '')
    sel_month = request.args.get('sel_month', '').strip()
    sel_q     = request.args.get('sel_q', '').strip()
    sel_half  = request.args.get('sel_half', '').strip()
    sel_year  = request.args.get('sel_year', '').strip()
    date_from = request.args.get('date_from', '').strip()
    date_to   = request.args.get('date_to', '').strip()

    d_from = d_to = None
    if period == 'day':
        d_from = d_to = today.isoformat()
    elif period == 'month':
        if sel_month:
            try:
                y, m = map(int, sel_month.split('-'))
            except Exception:
                y, m = today.year, today.month
        else:
            y, m = today.year, today.month
        d_from = date(y, m, 1).isoformat()
        d_to   = date(y, m, _cal.monthrange(y, m)[1]).isoformat()
    elif period == 'quarter':
        if sel_q:
            try:
                parts = sel_q.split('-Q'); y = int(parts[0]); q = int(parts[1])
            except Exception:
                y = today.year; q = (today.month - 1) // 3 + 1
        else:
            y = today.year; q = (today.month - 1) // 3 + 1
        ms = (q - 1) * 3 + 1; me = ms + 2
        d_from = date(y, ms, 1).isoformat()
        d_to   = date(y, me, _cal.monthrange(y, me)[1]).isoformat()
    elif period == 'halfyear':
        if sel_half:
            try:
                parts = sel_half.split('-H'); y = int(parts[0]); h = int(parts[1])
            except Exception:
                y = today.year; h = 1 if today.month <= 6 else 2
        else:
            y = today.year; h = 1 if today.month <= 6 else 2
        ms = 1 if h == 1 else 7; me = 6 if h == 1 else 12
        d_from = date(y, ms, 1).isoformat()
        d_to   = date(y, me, _cal.monthrange(y, me)[1]).isoformat()
    elif period == 'year':
        y = int(sel_year) if sel_year else today.year
        d_from = date(y, 1, 1).isoformat()
        d_to   = date(y, 12, 31).isoformat()
    elif period == 'custom' and date_from and date_to:
        d_from = date_from; d_to = date_to

    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)

    cur.execute("SELECT MIN(SUBSTRING(pos_issue_date,1,4)) AS m FROM contracts WHERE pos_issue_date IS NOT NULL AND pos_issue_date != ''")
    min_yr = cur.fetchone()['m']
    min_year = int(min_yr) if min_yr else today.year
    years = list(range(today.year, min_year - 1, -1))

    if d_from and d_to:
        cur.execute('SELECT * FROM contracts WHERE pos_issue_date BETWEEN %s AND %s ORDER BY dd', [d_from, d_to])
    else:
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
                           by_dept=by_dept, by_result=by_result,
                           period=period, d_from=d_from or '', d_to=d_to or '',
                           sel_month=sel_month, sel_q=sel_q, sel_half=sel_half,
                           sel_year=sel_year, date_from=date_from, date_to=date_to,
                           years=years, today=today.isoformat(), months_mn=MONTHS_MN)


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
    depts          = request.args.getlist('dept')
    employees      = request.args.getlist('employee')
    status_filters = request.args.getlist('status_filter')
    date_from     = request.args.get('date_from', '')
    date_to       = request.args.get('date_to', '')

    if period == 'day':
        d_from = today.isoformat()
        d_to   = today.isoformat()

    elif period == 'month':
        if sel_month:
            try:
                y, m = map(int, sel_month.split('-'))
            except Exception:
                y, m = today.year, today.month
        else:
            y, m = today.year, today.month
        d_from = dt_date(y, m, 1).isoformat()
        d_to   = dt_date(y, m, calendar.monthrange(y, m)[1]).isoformat()

    elif period == 'quarter':
        if sel_q:
            try:
                parts = sel_q.split('-Q')
                y = int(parts[0]); q = int(parts[1])
            except Exception:
                y = today.year; q = (today.month - 1) // 3 + 1
        else:
            y = today.year; q = (today.month - 1) // 3 + 1
        m_start = (q - 1) * 3 + 1
        m_end   = m_start + 2
        d_from  = dt_date(y, m_start, 1).isoformat()
        d_to    = dt_date(y, m_end, calendar.monthrange(y, m_end)[1]).isoformat()

    elif period == 'halfyear':
        if sel_half:
            try:
                parts = sel_half.split('-H')
                y = int(parts[0]); h = int(parts[1])
            except Exception:
                y = today.year; h = 1 if today.month <= 6 else 2
        else:
            y = today.year; h = 1 if today.month <= 6 else 2
        m_start = 1 if h == 1 else 7
        m_end   = 6 if h == 1 else 12
        d_from  = dt_date(y, m_start, 1).isoformat()
        d_to    = dt_date(y, m_end, calendar.monthrange(y, m_end)[1]).isoformat()

    elif period == 'year':
        y = int(sel_year) if sel_year else today.year
        d_from = dt_date(y, 1, 1).isoformat()
        d_to   = dt_date(y, 12, 31).isoformat()

    elif period == 'custom' and date_from and date_to:
        d_from = date_from
        d_to   = date_to

    else:
        d_from = dt_date(today.year, today.month, 1).isoformat()
        d_to   = dt_date(today.year, today.month, calendar.monthrange(today.year, today.month)[1]).isoformat()
        period = 'month'

    conn = get_db()
    cur  = conn.cursor(cursor_factory=RealDictCursor)

    cur.execute("SELECT DISTINCT issued_by FROM contracts WHERE issued_by IS NOT NULL AND issued_by != '' ORDER BY issued_by")
    all_employees = [r['issued_by'] for r in cur.fetchall()]

    cur.execute("SELECT DISTINCT department FROM managers WHERE department IS NOT NULL AND department != '' ORDER BY department")
    dash_departments = [r['department'] for r in cur.fetchall()]

    cur.execute("SELECT name, department FROM managers ORDER BY name")
    mgr_dept_map = {r['name']: r['department'] for r in cur.fetchall()}

    # Recalculate time_category once per day (skip if already done today)
    today_str = today.isoformat()
    if session.get('_tc_updated') != today_str:
        tc_queries = [
            """UPDATE contracts
               SET time_category = 'Хугацаа хэтэрсэн',
                   overdue_days  = (CURRENT_DATE - expected_date::date)
               WHERE expected_date IS NOT NULL AND expected_date != ''
                 AND expected_date ~ '^\\d{4}-\\d{2}-\\d{2}'
                 AND (received_date IS NULL OR received_date = '')
                 AND expected_date::date < CURRENT_DATE
                 AND (is_inactive IS NULL OR is_inactive != '1')""",
            """UPDATE contracts
               SET time_category = 'Хоосон', overdue_days = 0
               WHERE expected_date IS NOT NULL AND expected_date != ''
                 AND expected_date ~ '^\\d{4}-\\d{2}-\\d{2}'
                 AND (received_date IS NULL OR received_date = '')
                 AND expected_date::date >= CURRENT_DATE
                 AND (is_inactive IS NULL OR is_inactive != '1')""",
            """UPDATE contracts
               SET overdue_days  = (received_date::date - expected_date::date),
                   time_category = CASE
                     WHEN (received_date::date - expected_date::date) <= 0 THEN 'Хугацаандаа'
                     ELSE 'Хугацаа хэтэрсэн'
                   END
               WHERE expected_date IS NOT NULL AND expected_date != ''
                 AND expected_date ~ '^\\d{4}-\\d{2}-\\d{2}'
                 AND received_date  IS NOT NULL AND received_date  != ''
                 AND received_date  ~ '^\\d{4}-\\d{2}-\\d{2}'
                 AND (is_inactive IS NULL OR is_inactive != '1')""",
        ]
        for sql in tc_queries:
            try:
                cur.execute("SAVEPOINT tc")
                cur.execute(sql)
                cur.execute("RELEASE SAVEPOINT tc")
            except Exception:
                cur.execute("ROLLBACK TO SAVEPOINT tc")
        conn.commit()
        session['_tc_updated'] = today_str

    # For year period use SUBSTRING match so dates like '2026/01/05' (non-ISO) are included
    if period == 'year':
        year_str  = d_from[:4]
        base_cond = "WHERE (pos_issue_date IS NULL OR pos_issue_date = '' OR SUBSTRING(pos_issue_date, 1, 4) = %s)"
        base_params = [year_str]
    else:
        base_cond   = "WHERE (pos_issue_date BETWEEN %s AND %s OR pos_issue_date IS NULL OR pos_issue_date = '')"
        base_params = [d_from, d_to]
    if depts:
        base_cond += " AND department IN ({})".format(','.join(['%s']*len(depts)))
        base_params.extend(depts)
    if employees:
        base_cond += " AND issued_by IN ({})".format(','.join(['%s']*len(employees)))
        base_params.extend(employees)
    if status_filters:
        base_cond += " AND COALESCE(status,'') IN ({})".format(','.join(['%s']*len(status_filters)))
        base_params.extend(status_filters)

    cur.execute(f"SELECT * FROM contracts {base_cond}", base_params)
    EXCLUDE_INSP = {'7.Татагдсан', '6.Татагдсан'}
    all_rows = [r for r in cur.fetchall() if r['first_inspection'] not in EXCLUDE_INSP]
    inactive_count = sum(1 for r in all_rows if r['is_inactive'] == '1')
    active_rows    = [r for r in all_rows if r['is_inactive'] != '1']
    hoosoon_count  = sum(1 for r in active_rows if r['time_category'] == 'Хоосон')
    rows = [r for r in active_rows if r['time_category'] != 'Хоосон']

    total     = len(rows)
    total_all = total + hoosoon_count + inactive_count  # all contracts in the period
    on_time = sum(1 for r in rows if r['time_category'] == 'Хугацаандаа')
    overdue = total - on_time

    on_time_pct = round(on_time / total * 100, 1) if total else 0
    overdue_pct = round(overdue / total * 100, 1) if total else 0

    overdue_days_list = [r['overdue_days'] for r in rows
                         if r['overdue_days'] is not None and r['overdue_days'] > 0]
    avg_overdue = round(sum(overdue_days_list) / len(overdue_days_list), 1) if overdue_days_list else 0
    max_overdue = max(overdue_days_list) if overdue_days_list else 0

    complete      = sum(1 for r in rows if r['first_inspection'] in ('1.Бүрэн', '3.Салбар дээр архивлагдсан/бүрэн'))
    returned      = sum(1 for r in rows if r['first_inspection'] in ('2.Бүрдэл дутуу', '4.Бүртгэл буруу/дутуу', '5.Салбар дээр архивлагдсан/дутуу'))
    returned_repaired     = sum(1 for r in rows if r['first_inspection'] in ('2.Бүрдэл дутуу', '4.Бүртгэл буруу/дутуу', '5.Салбар дээр архивлагдсан/дутуу') and r.get('is_repaired') == 'Тийм')
    returned_not_repaired = returned - returned_repaired
    not_received  = sum(1 for r in rows if not r['first_inspection'])
    not_received_rows = [r for r in rows if not r['received_date']]

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
    performance_pct = round((on_time_pct + complete_pct) / 2, 1)

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

    # Build 6-month range list (oldest first)
    month_ranges = []
    for i in range(5, -1, -1):
        mo = today.month - i
        yr = today.year
        while mo <= 0:
            mo += 12; yr -= 1
        month_ranges.append((yr, mo))
    trend_from = dt_date(month_ranges[0][0], month_ranges[0][1], 1).isoformat()
    trend_to   = dt_date(month_ranges[-1][0], month_ranges[-1][1],
                         calendar.monthrange(month_ranges[-1][0], month_ranges[-1][1])[1]).isoformat()

    # Single query replaces 18 separate COUNT queries
    cur.execute("""
        SELECT
            SUBSTRING(pos_issue_date, 1, 7) AS ym,
            COUNT(*) AS total,
            SUM(CASE WHEN time_category = 'Хоосон'      THEN 1 ELSE 0 END) AS hoosoon,
            SUM(CASE WHEN time_category = 'Хугацаандаа' THEN 1 ELSE 0 END) AS on_time
        FROM contracts
        WHERE (is_inactive IS NULL OR is_inactive != '1')
          AND pos_issue_date BETWEEN %s AND %s
        GROUP BY 1
    """, [trend_from, trend_to])
    trend_data = {r['ym']: r for r in cur.fetchall()}

    trend = []
    for yr, mo in month_ranges:
        ym  = f'{yr}-{mo:02d}'
        row = trend_data.get(ym, {})
        t_total   = int(row.get('total')   or 0)
        t_hoosoon = int(row.get('hoosoon') or 0)
        t_on      = int(row.get('on_time') or 0)
        t_timed   = t_total - t_hoosoon
        trend.append({'label': MONTHS_MN[mo - 1], 'total': t_total,
                      'on_time': t_on, 'overdue': t_timed - t_on, 'hoosoon': t_hoosoon})

    cur.execute("SELECT MIN(SUBSTRING(pos_issue_date, 1, 4)) AS min_year FROM contracts WHERE pos_issue_date IS NOT NULL AND pos_issue_date != ''")
    min_year_row = cur.fetchone()['min_year']
    min_year = int(min_year_row) if min_year_row else today.year
    years = list(range(today.year, min_year - 1, -1))

    cur.close()
    conn.close()

    return render_template('dashboard.html',
        period=period, depts=depts, employees=employees,
        sel_month=sel_month, sel_q=sel_q, sel_half=sel_half, sel_year=sel_year,
        date_from=date_from, date_to=date_to,
        d_from=d_from, d_to=d_to,
        departments=dash_departments, statuses=STATUSES, all_employees=all_employees,
        status_filters=status_filters,
        years=years, today=today.isoformat(),
        months_mn=MONTHS_MN,
        total=total, total_all=total_all, hoosoon_count=hoosoon_count,
        on_time=on_time, overdue=overdue,
        on_time_pct=on_time_pct, overdue_pct=overdue_pct,
        avg_overdue=avg_overdue, max_overdue=max_overdue,
        complete=complete, incomplete=incomplete, returned=returned,
        returned_repaired=returned_repaired, returned_not_repaired=returned_not_repaired,
        not_received=not_received, nr_avg_overdue=nr_avg_overdue, nr_max_overdue=nr_max_overdue,
        complete_pct=complete_pct, returned_pct=returned_pct,
        performance_pct=performance_pct,
        insp_counts=insp_counts, dept_stats=dept_stats, trend=trend,
        inactive_count=inactive_count,
        contracts=sorted(
            [r for r in all_rows if r['first_inspection'] in
             {'2.Бүрдэл дутуу', '4.Бүртгэл буруу/дутуу', '5.Салбар дээр архивлагдсан/дутуу'}],
            key=lambda r: (r['dd'] or 0)
        ),
        mgr_dept_map=mgr_dept_map,
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
        'Д/д', 'Мерчантын нэр', 'Посын сериал', 'Мерч. дугаар', 'Терминал',
        'Төлөв', 'Пос огноо', 'Утас', 'Хэлбэр', 'Ажилтан', 'Хэлтэс',
        'Ирэх ёстой огноо', 'Хүлээн авсан огноо', 'Хэтэрсэн хоног', 'Ангилал',
        'Эхний хяналтын үр дүн', 'Тайлбар', 'Буцаасан огноо', 'Сүүлийн хяналт',
        'Зассан', 'Scanned', 'Засварлагдсан эсэх', 'Засварлагдсан огноо',
        'Оруулсан ажилтан', 'Оруулсан огноо', 'Идэвхжил'
    ]
    col_widths = [6, 28, 18, 16, 16, 12, 14, 14, 14, 16, 14,
                  18, 18, 14, 16, 26, 30, 22, 18,
                  14, 10, 18, 18, 18, 18, 12]

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
        rdates = [row.get('return_date') or '', row.get('return_date_2') or '',
                  row.get('return_date_3') or '', row.get('return_date_4') or '',
                  row.get('return_date_5') or '']
        rdates_str = ', '.join(d for d in rdates if d)
        values = [
            row['dd'], row['merchant_name'], row['pos_serial'],
            row['merchant_number'], row['terminal_number'],
            row['status'], row['pos_issue_date'], row['phone'],
            row['merchant_type'], row['issued_by'], row['department'],
            row['expected_date'], row['received_date'],
            row['overdue_days'], row['time_category'],
            row['first_inspection'], row['description'],
            rdates_str, row['last_inspection_date'],
            (row.get('modified_by') or '').split('@')[0] or '',
            '✓' if row.get('scanned') == '1' else '',
            row.get('is_repaired') or '',
            row.get('repaired_date') or '',
            (row.get('uploaded_by') or '').split('@')[0] or '',
            row.get('uploaded_date') or '',
            'Идэвхгүй' if row.get('is_inactive') == '1' else 'Идэвхитэй',
        ]
        is_overdue = row['time_category'] == 'Хугацаа хэтэрсэн'
        row_fill   = overdue_fill if is_overdue else on_time_fill

        for c_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border    = border
            cell.alignment = wrap_align if c_idx in (2, 16, 17) else center_align
            if c_idx == 15:
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
_db_init_error = None

@app.before_request
def ensure_db():
    global _db_ready, _db_init_error
    if not _db_ready:
        _db_ready = True  # set first so a crash doesn't retry on every request
        if not os.environ.get('DATABASE_URL'):
            _db_init_error = 'DATABASE_URL environment variable is not set!'
            return
        try:
            init_db()
        except Exception as e:
            import traceback
            _db_init_error = traceback.format_exc()

@app.errorhandler(Exception)
def handle_any_error(e):
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e  # pass 404 / 405 etc. through normally
    import traceback, html as _html
    try:
        tb = ''.join(traceback.format_exception(type(e), e, e.__traceback__))
        body = _html.escape(tb)
    except Exception:
        body = _html.escape(str(e))
    return f'<pre style="padding:20px;font-size:13px">{body}</pre>', 500

if __name__ == '__main__':
    init_db()
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    print('=' * 55)
    print('  Гэрээ бүртгэл хяналтын систем')
    print('  http://127.0.0.1:5000')
    print('=' * 55)
    app.run(debug=True, port=5000)
