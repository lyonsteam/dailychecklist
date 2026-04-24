from flask import Flask, request, jsonify, render_template
import psycopg2
import psycopg2.extras
import time
import uuid
import os
import hashlib
import json

# ── Optional imports for email pipeline (Windows only) ───────────────────────
try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False

try:
    from apscheduler.schedulers.background import BackgroundScheduler
    SCHEDULER_AVAILABLE = True
except ImportError:
    SCHEDULER_AVAILABLE = False

try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# ── Config ────────────────────────────────────────────────────────────────────
app = Flask(__name__)

# Render injects DATABASE_URL automatically when a Postgres instance is attached.
DATABASE_URL = os.environ.get('DATABASE_URL', '')

# Render's Postgres URLs start with "postgres://" but psycopg2 requires
# "postgresql://" — fix it silently if needed.
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)

OPENAI_API_KEY        = os.environ.get('OPENAI_API_KEY', 'sk-REPLACE_WITH_YOUR_OPENAI_API_KEY')
ADMIN_PASSWORD        = os.environ.get('ADMIN_PASSWORD', 'ADMIN321')
AUTO_ARCHIVE_CHECKED_SECS = 300
URGENT_HOURS          = 48
ARCHIVE_HOURS         = 24
EMAIL_LOOKBACK_MINS   = 180

# ── Mailbox Map ───────────────────────────────────────────────────────────────
MAILBOX_MAP = {
    "jlyons@cmgfi.com": {
        "outlook_name": "Jill Lyons",
        "mailbox_key":  "jlyons",
        "display_name": "Jill Lyons",
    },
    # Future user example:
    # "jreed@gmail.com": {
    #     "outlook_name": "John Reed",
    #     "mailbox_key":  "jreed",
    #     "display_name": "John Reed",
    # },
}


# ── Database ──────────────────────────────────────────────────────────────────

def get_db():
    """Return a new psycopg2 connection using DATABASE_URL."""
    if not DATABASE_URL:
        raise RuntimeError(
            "DATABASE_URL environment variable is not set. "
            "Add it in Render → your web service → Environment."
        )
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    return conn


def make_user_id(email: str) -> str:
    return hashlib.sha256(email.strip().lower().encode()).hexdigest()


def init_db():
    """Create tables if they don't exist and seed approved users."""
    conn = get_db()
    try:
        with conn:
            cur = conn.cursor()

            # Users table
            cur.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    user_id      TEXT PRIMARY KEY,
                    email        TEXT UNIQUE NOT NULL,
                    approved     INTEGER NOT NULL DEFAULT 0,
                    display_name TEXT,
                    mailbox_key  TEXT
                )
            ''')

            # Tasks table
            cur.execute('''
                CREATE TABLE IF NOT EXISTS tasks (
                    id          TEXT PRIMARY KEY,
                    text        TEXT NOT NULL,
                    created_at  DOUBLE PRECISION NOT NULL,
                    checked     INTEGER NOT NULL DEFAULT 0,
                    checked_at  DOUBLE PRECISION,
                    archived    INTEGER NOT NULL DEFAULT 0,
                    user_id     TEXT
                )
            ''')

            # Email tasks table
            cur.execute('''
                CREATE TABLE IF NOT EXISTS email_tasks (
                    id           TEXT PRIMARY KEY,
                    sender       TEXT,
                    email        TEXT,
                    subject      TEXT,
                    summary      TEXT,
                    reason       TEXT,
                    timestamp    DOUBLE PRECISION,
                    mailbox_key  TEXT
                )
            ''')

            # Seed / sync approved users from MAILBOX_MAP
            for email, cfg in MAILBOX_MAP.items():
                uid = make_user_id(email)
                cur.execute('''
                    INSERT INTO users (user_id, email, approved, display_name, mailbox_key)
                    VALUES (%s, %s, 1, %s, %s)
                    ON CONFLICT (email) DO UPDATE SET
                        approved     = 1,
                        display_name = EXCLUDED.display_name,
                        mailbox_key  = EXCLUDED.mailbox_key
                ''', (uid, email, cfg['display_name'], cfg['mailbox_key']))

        print("[db] Tables ready.")
    finally:
        conn.close()


# ── Helpers ───────────────────────────────────────────────────────────────────

def row_to_dict(row):
    """Convert a RealDictRow to a plain dict with normalised types."""
    return {
        'id':         row['id'],
        'text':       row['text'],
        'created_at': float(row['created_at']),
        'checked':    bool(row['checked']),
        'checked_at': float(row['checked_at']) if row['checked_at'] is not None else None,
        'archived':   bool(row['archived']),
        'user_id':    row.get('user_id'),
    }


def local_day_start(ts: float, tz_offset_mins: int) -> float:
    local_offset = -tz_offset_mins * 60
    local_ts     = ts + local_offset
    midnight     = (local_ts // 86400) * 86400
    return midnight - local_offset


def sweep_auto_archive(now: float):
    """Promote tasks that have been checked for 5+ minutes to archived."""
    cutoff = now - AUTO_ARCHIVE_CHECKED_SECS
    conn = get_db()
    try:
        with conn:
            cur = conn.cursor()
            cur.execute('''
                UPDATE tasks SET archived = 1
                WHERE checked = 1 AND archived = 0
                  AND checked_at IS NOT NULL AND checked_at <= %s
            ''', (cutoff,))
    finally:
        conn.close()


def get_approved_user(email: str):
    conn = get_db()
    try:
        cur = conn.cursor()
        cur.execute(
            'SELECT * FROM users WHERE email = %s AND approved = 1',
            (email.strip().lower(),)
        )
        row = cur.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


# ── Auth / Login ──────────────────────────────────────────────────────────────

@app.route('/login', methods=['POST'])
def login():
    data  = request.get_json(force=True)
    email = (data.get('email') or '').strip().lower()
    if not email:
        return jsonify({'error': 'email required'}), 400

    user = get_approved_user(email)
    if not user:
        return jsonify({'error': 'Access denied. This email has not been approved.'}), 403

    return jsonify({
        'user_id':      user['user_id'],
        'email':        user['email'],
        'display_name': user['display_name'] or user['email'],
        'mailbox_key':  user['mailbox_key'],
    })


# ── Admin endpoints ───────────────────────────────────────────────────────────

@app.route('/admin/users', methods=['POST'])
def admin_list_users():
    data = request.get_json(force=True)
    if data.get('admin_password') != ADMIN_PASSWORD:
        return jsonify({'error': 'Unauthorized'}), 401
    conn = get_db()
    try:
        cur = conn.cursor()
        cur.execute(
            'SELECT user_id, email, approved, display_name, mailbox_key '
            'FROM users ORDER BY email'
        )
        return jsonify({'users': [dict(r) for r in cur.fetchall()]})
    finally:
        conn.close()


@app.route('/admin/add_user', methods=['POST'])
def admin_add_user():
    data = request.get_json(force=True)
    if data.get('admin_password') != ADMIN_PASSWORD:
        return jsonify({'error': 'Unauthorized'}), 401

    email = (data.get('email') or '').strip().lower()
    if not email or '@' not in email:
        return jsonify({'error': 'Valid email required'}), 400

    display_name = (data.get('display_name') or email).strip()
    mailbox_key  = (data.get('mailbox_key') or email.split('@')[0]).strip()
    uid          = make_user_id(email)

    conn = get_db()
    try:
        with conn:
            cur = conn.cursor()
            cur.execute('''
                INSERT INTO users (user_id, email, approved, display_name, mailbox_key)
                VALUES (%s, %s, 1, %s, %s)
                ON CONFLICT (email) DO UPDATE SET
                    approved     = 1,
                    display_name = EXCLUDED.display_name,
                    mailbox_key  = EXCLUDED.mailbox_key
            ''', (uid, email, display_name, mailbox_key))
    finally:
        conn.close()

    return jsonify({
        'added':        True,
        'user_id':      uid,
        'email':        email,
        'display_name': display_name,
        'mailbox_key':  mailbox_key,
    })


@app.route('/admin/remove_user', methods=['POST'])
def admin_remove_user():
    data = request.get_json(force=True)
    if data.get('admin_password') != ADMIN_PASSWORD:
        return jsonify({'error': 'Unauthorized'}), 401
    email = (data.get('email') or '').strip().lower()
    if not email:
        return jsonify({'error': 'email required'}), 400
    conn = get_db()
    try:
        with conn:
            cur = conn.cursor()
            cur.execute('UPDATE users SET approved = 0 WHERE email = %s', (email,))
    finally:
        conn.close()
    return jsonify({'revoked': email})


# ── Main app routes ───────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/get', methods=['GET'])
def get_tasks():
    user_id = request.args.get('user_id', '').strip()
    if not user_id:
        return jsonify({'error': 'user_id required'}), 400

    server_now = time.time()
    sweep_auto_archive(server_now)

    try:
        client_now = float(request.args.get('now', server_now))
    except (TypeError, ValueError):
        client_now = server_now

    try:
        tz_offset = int(request.args.get('tz', 0))
    except (TypeError, ValueError):
        tz_offset = 0

    today_start = local_day_start(client_now, tz_offset)
    archive_cut = server_now - (ARCHIVE_HOURS * 3600)

    conn = get_db()
    try:
        cur = conn.cursor()
        cur.execute(
            'SELECT * FROM tasks WHERE user_id = %s ORDER BY created_at ASC',
            (user_id,)
        )
        rows = cur.fetchall()
    finally:
        conn.close()

    all_tasks = [row_to_dict(r) for r in rows]

    local_offset_secs = -tz_offset * 60

    def task_local_day_start(ts: float) -> float:
        local_ts = ts + local_offset_secs
        midnight  = (local_ts // 86400) * 86400
        return midnight - local_offset_secs

    day_starts_with_tasks: set = set()
    for t in all_tasks:
        ds = task_local_day_start(t['created_at'])
        if ds < today_start:
            day_starts_with_tasks.add(ds)

    active_day_starts = sorted(day_starts_with_tasks, reverse=True)[:3]
    while len(active_day_starts) < 3:
        active_day_starts.append(None)

    panel_day_starts = active_day_starts

    buckets = {
        'today': [],
        'day1': [], 'day2': [], 'day3': [],
        'archive': [], 'unchecked_archive': [],
    }

    panel_map = {}
    for i, ds in enumerate(panel_day_starts):
        if ds is not None:
            panel_map[ds] = f'day{i+1}'

    for t in all_tasks:
        ca    = t['created_at']
        age_h = (server_now - ca) / 3600

        if t['archived']:
            buckets['archive'].append(t)
            continue

        if ca < archive_cut:
            if t['checked']:
                buckets['archive'].append(t)
            else:
                t['urgent'] = True
                buckets['unchecked_archive'].append(t)
            continue

        if ca >= today_start:
            buckets['today'].append(t)
        else:
            ds     = task_local_day_start(ca)
            bucket = panel_map.get(ds)
            if bucket:
                if not t['checked'] and age_h >= URGENT_HOURS:
                    t['urgent'] = True
                buckets[bucket].append(t)
            else:
                if t['checked']:
                    buckets['archive'].append(t)
                else:
                    t['urgent'] = True
                    buckets['unchecked_archive'].append(t)

    return jsonify({
        'buckets': buckets,
        'panel_dates': {
            'today': today_start,
            'day1':  panel_day_starts[0],
            'day2':  panel_day_starts[1],
            'day3':  panel_day_starts[2],
        },
        'server_now': server_now,
    })


@app.route('/add', methods=['POST'])
def add_task():
    data    = request.get_json(force=True)
    user_id = (data.get('user_id') or '').strip()
    if not user_id:
        return jsonify({'error': 'user_id required'}), 400

    texts = data.get('tasks', [])
    if isinstance(texts, str):
        texts = [texts]

    now     = time.time()
    created = []
    conn    = get_db()
    try:
        with conn:
            cur = conn.cursor()
            for text in texts:
                text = text.strip()
                if not text:
                    continue
                tid = str(uuid.uuid4())
                cur.execute(
                    'INSERT INTO tasks '
                    '(id, text, created_at, checked, checked_at, archived, user_id) '
                    'VALUES (%s, %s, %s, 0, NULL, 0, %s)',
                    (tid, text, now, user_id)
                )
                created.append({
                    'id': tid, 'text': text, 'created_at': now,
                    'checked': False, 'checked_at': None,
                    'archived': False, 'user_id': user_id,
                })
    finally:
        conn.close()

    return jsonify({'created': created}), 201


@app.route('/toggle', methods=['POST'])
def toggle_task():
    data    = request.get_json(force=True)
    task_id = data.get('id')
    user_id = (data.get('user_id') or '').strip()
    if not task_id:
        return jsonify({'error': 'id required'}), 400
    if not user_id:
        return jsonify({'error': 'user_id required'}), 400

    now  = time.time()
    conn = get_db()
    try:
        with conn:
            cur = conn.cursor()
            cur.execute(
                'SELECT * FROM tasks WHERE id = %s AND user_id = %s',
                (task_id, user_id)
            )
            row = cur.fetchone()
            if not row:
                return jsonify({'error': 'not found'}), 404
            task           = row_to_dict(row)
            new_checked    = not task['checked']
            new_checked_at = now if new_checked else None
            cur.execute(
                'UPDATE tasks SET checked = %s, checked_at = %s, archived = 0 '
                'WHERE id = %s AND user_id = %s',
                (int(new_checked), new_checked_at, task_id, user_id)
            )
            task.update(checked=new_checked, checked_at=new_checked_at, archived=False)
    finally:
        conn.close()

    return jsonify({'task': task})


@app.route('/delete', methods=['POST'])
def delete_task():
    data    = request.get_json(force=True)
    task_id = data.get('id')
    user_id = (data.get('user_id') or '').strip()
    if not task_id:
        return jsonify({'error': 'id required'}), 400
    conn = get_db()
    try:
        with conn:
            cur = conn.cursor()
            cur.execute(
                'DELETE FROM tasks WHERE id = %s AND user_id = %s',
                (task_id, user_id)
            )
    finally:
        conn.close()
    return jsonify({'deleted': task_id})


@app.route('/archive-now', methods=['POST'])
def archive_now():
    """Immediately archive a task — skips the 5-minute countdown."""
    data    = request.get_json(force=True)
    task_id = data.get('id')
    user_id = (data.get('user_id') or '').strip()
    if not task_id:
        return jsonify({'error': 'id required'}), 400
    if not user_id:
        return jsonify({'error': 'user_id required'}), 400
    conn = get_db()
    try:
        with conn:
            cur = conn.cursor()
            cur.execute(
                'SELECT id FROM tasks WHERE id = %s AND user_id = %s',
                (task_id, user_id)
            )
            if not cur.fetchone():
                return jsonify({'error': 'not found'}), 404
            cur.execute(
                'UPDATE tasks SET archived = 1, checked = 1, checked_at = %s '
                'WHERE id = %s AND user_id = %s',
                (time.time(), task_id, user_id)
            )
    finally:
        conn.close()
    return jsonify({'archived': task_id})


# ── Email pipeline ────────────────────────────────────────────────────────────

def read_outlook_emails_for(outlook_name: str) -> list:
    if not OUTLOOK_AVAILABLE:
        print(f"[email_pipeline] pywin32 not available — skipping {outlook_name}")
        return []
    try:
        outlook      = win32com.client.Dispatch("Outlook.Application")
        namespace    = outlook.GetNamespace("MAPI")
        target_store = None

        for store in namespace.Stores:
            if store.DisplayName == outlook_name:
                target_store = store
                break
        if target_store is None:
            for folder in namespace.Folders:
                if folder.Name == outlook_name:
                    target_store = folder
                    break
        if target_store is None:
            print(f"[email_pipeline] Mailbox '{outlook_name}' not found")
            return []

        try:
            inbox = target_store.GetRootFolder().Folders["Inbox"]
        except Exception:
            inbox = target_store.GetDefaultFolder(6)

        cutoff_time = time.time() - (EMAIL_LOOKBACK_MINS * 60)
        emails = []
        for msg in inbox.Items:
            try:
                received_ts = msg.ReceivedTime.timestamp()
                if received_ts < cutoff_time:
                    continue
                emails.append({
                    'sender':    msg.SenderName,
                    'email':     msg.SenderEmailAddress,
                    'subject':   msg.Subject,
                    'body':      msg.Body[:2000],
                    'timestamp': received_ts,
                })
            except Exception:
                continue
        return emails
    except Exception as e:
        print(f"[email_pipeline] Outlook error for '{outlook_name}': {e}")
        return []


def triage_emails_with_openai(emails: list) -> list:
    if not OPENAI_AVAILABLE or not emails:
        return []
    client     = OpenAI(api_key=OPENAI_API_KEY)
    email_text = json.dumps(emails, default=str, indent=2)
    prompt = f"""You are an email triage assistant. Review these emails and return ONLY the important or actionable ones.

Ignore: newsletters, promotions, automated notifications, spam.
Include: emails requiring action, urgent requests, important updates, replies needing a response.

Return STRICT JSON ONLY — no markdown, no preamble, no explanation.
Format:
[
  {{
    "sender": "Name",
    "email": "address@example.com",
    "subject": "Subject line",
    "summary": "One-sentence summary of what action is needed",
    "reason": "Why this email is important",
    "timestamp": 1234567890.0
  }}
]

If no emails are important, return an empty array: []

EMAILS:
{email_text}"""
    try:
        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = response.choices[0].message.content.strip()
        raw = raw.replace('```json', '').replace('```', '').strip()
        return json.loads(raw)
    except Exception as e:
        print(f"[email_pipeline] OpenAI error: {e}")
        return []


def store_email_tasks(important_emails: list, mailbox_key: str, user_id: str):
    if not important_emails:
        return
    now  = time.time()
    conn = get_db()
    try:
        with conn:
            cur = conn.cursor()
            cur.execute(
                'SELECT subject FROM email_tasks WHERE mailbox_key = %s',
                (mailbox_key,)
            )
            existing = {r['subject'] for r in cur.fetchall()}

            for em in important_emails:
                subject = em.get('subject', '')
                if subject in existing:
                    continue

                em_id = str(uuid.uuid4())
                cur.execute('''
                    INSERT INTO email_tasks
                    (id, sender, email, subject, summary, reason, timestamp, mailbox_key)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                ''', (
                    em_id,
                    em.get('sender', ''),
                    em.get('email', ''),
                    subject,
                    em.get('summary', ''),
                    em.get('reason', ''),
                    em.get('timestamp', now),
                    mailbox_key,
                ))

                task_text = f"[Email] {subject} — {em.get('summary', '')}"
                task_id   = str(uuid.uuid4())
                cur.execute(
                    'INSERT INTO tasks '
                    '(id, text, created_at, checked, checked_at, archived, user_id) '
                    'VALUES (%s, %s, %s, 0, NULL, 0, %s)',
                    (task_id, task_text, now, user_id)
                )
                existing.add(subject)

    finally:
        conn.close()
    print(f"[email_pipeline:{mailbox_key}] Stored {len(important_emails)} email task(s)")


def run_pipeline_for(login_email: str):
    cfg = MAILBOX_MAP.get(login_email)
    if not cfg:
        print(f"[email_pipeline] No MAILBOX_MAP entry for {login_email}")
        return
    outlook_name = cfg['outlook_name']
    mailbox_key  = cfg['mailbox_key']
    user_id      = make_user_id(login_email)
    print(f"[email_pipeline] Running for {outlook_name} ({mailbox_key})…")
    emails    = read_outlook_emails_for(outlook_name)
    important = triage_emails_with_openai(emails)
    store_email_tasks(important, mailbox_key, user_id)
    print(f"[email_pipeline:{mailbox_key}] Done. {len(emails)} read, {len(important)} important.")


def email_pipeline():
    for login_email in MAILBOX_MAP:
        run_pipeline_for(login_email)


@app.route('/get_emails', methods=['GET'])
def get_emails():
    user_id = request.args.get('user_id', '').strip()
    if not user_id:
        return jsonify({'error': 'user_id required'}), 400
    conn = get_db()
    try:
        cur = conn.cursor()
        cur.execute(
            'SELECT mailbox_key FROM users WHERE user_id = %s AND approved = 1',
            (user_id,)
        )
        user = cur.fetchone()
        if not user or not user['mailbox_key']:
            return jsonify([])
        cur.execute(
            'SELECT * FROM email_tasks WHERE mailbox_key = %s ORDER BY timestamp DESC',
            (user['mailbox_key'],)
        )
        return jsonify([dict(r) for r in cur.fetchall()])
    finally:
        conn.close()


# ── Scheduler ─────────────────────────────────────────────────────────────────
_pipeline_last_run: dict = {'ts': None, 'status': 'never run'}


def _tracked_email_pipeline():
    _pipeline_last_run['ts']     = time.time()
    _pipeline_last_run['status'] = 'running'
    try:
        email_pipeline()
        _pipeline_last_run['status'] = 'ok'
    except Exception as e:
        _pipeline_last_run['status'] = f'error: {e}'
        print(f"[scheduler] Pipeline error: {e}")


def _start_scheduler():
    if not SCHEDULER_AVAILABLE:
        print("[scheduler] APScheduler not installed — pipeline will not run automatically.")
        return

    # Avoid double-start under Flask debug reloader
    if os.environ.get('WERKZEUG_RUN_MAIN') == 'false':
        print("[scheduler] Reloader parent — skipping scheduler start.")
        return

    scheduler = BackgroundScheduler(
        job_defaults={
            'misfire_grace_time': 600,
            'coalesce':           True,
            'max_instances':      1,
        }
    )
    scheduler.add_job(
        _tracked_email_pipeline,
        trigger='interval',
        hours=3,
        id='email_pipeline',
        next_run_time=None,
    )
    scheduler.start()
    print("[scheduler] APScheduler started — email pipeline runs every 3 hours.")

    import atexit
    atexit.register(lambda: scheduler.shutdown(wait=False))


# ── Manual trigger & status ───────────────────────────────────────────────────

@app.route('/run-now', methods=['POST'])
def run_now():
    data = request.get_json(force=True) or {}
    if data.get('admin_password') != ADMIN_PASSWORD:
        return jsonify({'error': 'Unauthorized'}), 401
    try:
        _tracked_email_pipeline()
        return jsonify({
            'status':  'ok',
            'ran_at':  _pipeline_last_run['ts'],
            'message': 'Pipeline completed successfully.',
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/pipeline-status', methods=['GET'])
def pipeline_status():
    return jsonify({
        'last_run_ts': _pipeline_last_run['ts'],
        'last_run_at': (
            time.strftime('%Y-%m-%d %H:%M:%S UTC', time.gmtime(_pipeline_last_run['ts']))
            if _pipeline_last_run['ts'] else None
        ),
        'status':     _pipeline_last_run['status'],
        'interval_h': 3,
    })


# ── Boot ──────────────────────────────────────────────────────────────────────
# Use Flask's before_request to init lazily on the first real request.
# This avoids the DATABASE_URL-not-set crash that happens when gunicorn
# imports the module before environment variables are fully injected.

_db_initialised = False

@app.before_request
def _lazy_init():
    global _db_initialised
    if not _db_initialised:
        db_host = DATABASE_URL.split('@')[-1].split('/')[0] if DATABASE_URL else 'NOT SET'
        print(f"[db] Connecting to: {db_host}")
        init_db()
        _start_scheduler()
        _db_initialised = True

if __name__ == '__main__':
    app.run(debug=True, port=5000)
