from flask import Flask, request, jsonify, render_template
import sqlite3
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
DB_PATH = os.path.join(os.path.dirname(__file__), 'tasks.db')

OPENAI_API_KEY            = "sk-REPLACE_WITH_YOUR_OPENAI_API_KEY"   # <-- fill this in
ADMIN_PASSWORD            = "ADMIN321"
AUTO_ARCHIVE_CHECKED_SECS = 300
URGENT_HOURS              = 48
ARCHIVE_HOURS             = 96
EMAIL_LOOKBACK_MINS       = 180

# ── Mailbox Map ───────────────────────────────────────────────────────────────
# Maps a login email → their Outlook mailbox display name + a short scope key.
# To add a future user: add an entry here AND approve them via the admin panel.
MAILBOX_MAP = {
    "jlyons@cmgfi.com": {
        "outlook_name": "Jill Lyons",   # must match Outlook store DisplayName exactly
        "mailbox_key":  "jlyons",
        "display_name": "Jill Lyons",
    },
    # Future user example — uncomment + fill in when ready:
    # "jreed@gmail.com": {
    #     "outlook_name": "John Reed",
    #     "mailbox_key":  "jreed",
    #     "display_name": "John Reed",
    # },
}


# ── Database ──────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def make_user_id(email: str) -> str:
    return hashlib.sha256(email.strip().lower().encode()).hexdigest()


def init_db():
    with get_db() as conn:
        # Users table — approved=1 required to log in
        conn.execute('''
            CREATE TABLE IF NOT EXISTS users (
                user_id      TEXT PRIMARY KEY,
                email        TEXT UNIQUE,
                approved     INTEGER NOT NULL DEFAULT 0,
                display_name TEXT,
                mailbox_key  TEXT
            )
        ''')

        # Tasks table
        conn.execute('''
            CREATE TABLE IF NOT EXISTS tasks (
                id          TEXT PRIMARY KEY,
                text        TEXT NOT NULL,
                created_at  REAL NOT NULL,
                checked     INTEGER NOT NULL DEFAULT 0,
                checked_at  REAL,
                archived    INTEGER NOT NULL DEFAULT 0,
                user_id     TEXT
            )
        ''')

        # Email tasks — scoped per mailbox_key
        conn.execute('''
            CREATE TABLE IF NOT EXISTS email_tasks (
                id           TEXT PRIMARY KEY,
                sender       TEXT,
                email        TEXT,
                subject      TEXT,
                summary      TEXT,
                reason       TEXT,
                timestamp    REAL,
                mailbox_key  TEXT
            )
        ''')

        # Non-destructive migrations for existing DBs
        migrations = [
            ('tasks',       'archived',     'INTEGER NOT NULL DEFAULT 0'),
            ('tasks',       'user_id',      'TEXT'),
            ('users',       'approved',     'INTEGER NOT NULL DEFAULT 0'),
            ('users',       'display_name', 'TEXT'),
            ('users',       'mailbox_key',  'TEXT'),
            ('email_tasks', 'mailbox_key',  'TEXT'),
        ]
        for table, col, defn in migrations:
            try:
                conn.execute(f'ALTER TABLE {table} ADD COLUMN {col} {defn}')
            except Exception:
                pass

        # Seed / sync approved users from MAILBOX_MAP
        for email, cfg in MAILBOX_MAP.items():
            uid = make_user_id(email)
            conn.execute(
                '''INSERT INTO users (user_id, email, approved, display_name, mailbox_key)
                   VALUES (?,?,1,?,?)
                   ON CONFLICT(email) DO UPDATE SET
                     approved=1,
                     display_name=excluded.display_name,
                     mailbox_key=excluded.mailbox_key''',
                (uid, email, cfg['display_name'], cfg['mailbox_key'])
            )

        conn.commit()


# ── Helpers ───────────────────────────────────────────────────────────────────

def row_to_dict(row):
    keys = row.keys()
    return {
        'id':         row['id'],
        'text':       row['text'],
        'created_at': row['created_at'],
        'checked':    bool(row['checked']),
        'checked_at': row['checked_at'],
        'archived':   bool(row['archived']),
        'user_id':    row['user_id'] if 'user_id' in keys else None,
    }


def local_day_start(ts: float, tz_offset_mins: int) -> float:
    local_offset = -tz_offset_mins * 60
    local_ts     = ts + local_offset
    midnight     = (local_ts // 86400) * 86400
    return midnight - local_offset


def sweep_auto_archive(now: float):
    cutoff = now - AUTO_ARCHIVE_CHECKED_SECS
    with get_db() as conn:
        conn.execute(
            '''UPDATE tasks SET archived = 1
               WHERE checked = 1 AND archived = 0
                 AND checked_at IS NOT NULL AND checked_at <= ?''',
            (cutoff,)
        )
        conn.commit()


def get_approved_user(email: str):
    with get_db() as conn:
        row = conn.execute(
            'SELECT * FROM users WHERE email=? AND approved=1',
            (email.strip().lower(),)
        ).fetchone()
    return dict(row) if row else None


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
    with get_db() as conn:
        rows = conn.execute(
            'SELECT user_id, email, approved, display_name, mailbox_key FROM users ORDER BY email'
        ).fetchall()
    return jsonify({'users': [dict(r) for r in rows]})


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

    with get_db() as conn:
        conn.execute(
            '''INSERT INTO users (user_id, email, approved, display_name, mailbox_key)
               VALUES (?,?,1,?,?)
               ON CONFLICT(email) DO UPDATE SET
                 approved=1,
                 display_name=excluded.display_name,
                 mailbox_key=excluded.mailbox_key''',
            (uid, email, display_name, mailbox_key)
        )
        conn.commit()

    return jsonify({
        'added': True,
        'user_id': uid,
        'email': email,
        'display_name': display_name,
        'mailbox_key': mailbox_key,
    })


@app.route('/admin/remove_user', methods=['POST'])
def admin_remove_user():
    data = request.get_json(force=True)
    if data.get('admin_password') != ADMIN_PASSWORD:
        return jsonify({'error': 'Unauthorized'}), 401
    email = (data.get('email') or '').strip().lower()
    if not email:
        return jsonify({'error': 'email required'}), 400
    with get_db() as conn:
        conn.execute('UPDATE users SET approved=0 WHERE email=?', (email,))
        conn.commit()
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
    day1_start  = today_start - 86400
    day2_start  = today_start - 172800
    day3_start  = today_start - 259200
    archive_cut = server_now - (ARCHIVE_HOURS * 3600)

    with get_db() as conn:
        rows = conn.execute(
            'SELECT * FROM tasks WHERE user_id = ? ORDER BY created_at ASC',
            (user_id,)
        ).fetchall()

    buckets = {
        'today': [], 'day1': [], 'day2': [], 'day3': [],
        'archive': [], 'unchecked_archive': [],
    }

    for row in rows:
        t     = row_to_dict(row)
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
        elif ca >= day1_start:
            if not t['checked'] and age_h >= URGENT_HOURS:
                t['urgent'] = True
            buckets['day1'].append(t)
        elif ca >= day2_start:
            if not t['checked']:
                t['urgent'] = True
            buckets['day2'].append(t)
        else:
            if not t['checked']:
                t['urgent'] = True
            buckets['day3'].append(t)

    return jsonify({
        'buckets': buckets,
        'panel_dates': {
            'today': today_start,
            'day1':  day1_start,
            'day2':  day2_start,
            'day3':  day3_start,
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
    with get_db() as conn:
        for text in texts:
            text = text.strip()
            if not text:
                continue
            tid = str(uuid.uuid4())
            conn.execute(
                'INSERT INTO tasks (id, text, created_at, checked, checked_at, archived, user_id) '
                'VALUES (?,?,?,0,NULL,0,?)',
                (tid, text, now, user_id)
            )
            created.append({
                'id': tid, 'text': text, 'created_at': now,
                'checked': False, 'checked_at': None, 'archived': False,
                'user_id': user_id,
            })
        conn.commit()

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

    now = time.time()
    with get_db() as conn:
        row = conn.execute(
            'SELECT * FROM tasks WHERE id=? AND user_id=?',
            (task_id, user_id)
        ).fetchone()
        if not row:
            return jsonify({'error': 'not found'}), 404
        task           = row_to_dict(row)
        new_checked    = not task['checked']
        new_checked_at = now if new_checked else None
        conn.execute(
            'UPDATE tasks SET checked=?, checked_at=?, archived=0 WHERE id=? AND user_id=?',
            (int(new_checked), new_checked_at, task_id, user_id)
        )
        conn.commit()
        task.update(checked=new_checked, checked_at=new_checked_at, archived=False)

    return jsonify({'task': task})


@app.route('/delete', methods=['POST'])
def delete_task():
    data    = request.get_json(force=True)
    task_id = data.get('id')
    user_id = (data.get('user_id') or '').strip()
    if not task_id:
        return jsonify({'error': 'id required'}), 400
    with get_db() as conn:
        conn.execute('DELETE FROM tasks WHERE id=? AND user_id=?', (task_id, user_id))
        conn.commit()
    return jsonify({'deleted': task_id})


# ── Email pipeline ────────────────────────────────────────────────────────────

def read_outlook_emails_for(outlook_name: str) -> list:
    if not OUTLOOK_AVAILABLE:
        print(f"[email_pipeline] pywin32 not available — skipping {outlook_name}")
        return []

    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
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
            print(f"[email_pipeline] Mailbox '{outlook_name}' not found in Outlook")
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

    now = time.time()
    with get_db() as conn:
        existing = {
            row[0] for row in conn.execute(
                'SELECT subject FROM email_tasks WHERE mailbox_key=?',
                (mailbox_key,)
            ).fetchall()
        }

        for em in important_emails:
            subject = em.get('subject', '')
            if subject in existing:
                continue

            em_id = str(uuid.uuid4())
            conn.execute(
                '''INSERT INTO email_tasks
                   (id, sender, email, subject, summary, reason, timestamp, mailbox_key)
                   VALUES (?,?,?,?,?,?,?,?)''',
                (
                    em_id,
                    em.get('sender', ''),
                    em.get('email', ''),
                    subject,
                    em.get('summary', ''),
                    em.get('reason', ''),
                    em.get('timestamp', now),
                    mailbox_key,
                )
            )

            # Inject into tasks scoped to this user
            task_text = f"[Email] {subject} — {em.get('summary', '')}"
            task_id   = str(uuid.uuid4())
            conn.execute(
                'INSERT INTO tasks (id, text, created_at, checked, checked_at, archived, user_id) '
                'VALUES (?,?,?,0,NULL,0,?)',
                (task_id, task_text, now, user_id)
            )

            existing.add(subject)

        conn.commit()
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
    """Run pipeline for every entry defined in MAILBOX_MAP."""
    for login_email in MAILBOX_MAP:
        run_pipeline_for(login_email)


@app.route('/get_emails', methods=['GET'])
def get_emails():
    user_id = request.args.get('user_id', '').strip()
    if not user_id:
        return jsonify({'error': 'user_id required'}), 400

    with get_db() as conn:
        user = conn.execute(
            'SELECT mailbox_key FROM users WHERE user_id=? AND approved=1',
            (user_id,)
        ).fetchone()

    if not user or not user['mailbox_key']:
        return jsonify([])

    with get_db() as conn:
        rows = conn.execute(
            'SELECT * FROM email_tasks WHERE mailbox_key=? ORDER BY timestamp DESC',
            (user['mailbox_key'],)
        ).fetchall()

    return jsonify([dict(r) for r in rows])


# ── Scheduler ─────────────────────────────────────────────────────────────────

if SCHEDULER_AVAILABLE:
    scheduler = BackgroundScheduler()
    scheduler.add_job(email_pipeline, 'interval', hours=3)
    scheduler.start()
    print("[scheduler] Email pipeline scheduled every 3 hours")
else:
    print("[scheduler] APScheduler not available — email pipeline will not run automatically")


# ── Boot ──────────────────────────────────────────────────────────────────────

init_db()

if __name__ == '__main__':
    app.run(debug=True, port=5000)