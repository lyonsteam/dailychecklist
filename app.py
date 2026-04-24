"""
Daily Tracker — Flask backend
• Postgres  when DATABASE_URL env var is present  (Render production)
• SQLite    fallback when DATABASE_URL is absent   (local development)
"""

import os, time, uuid, hashlib, json
from contextlib import contextmanager
from flask import Flask, request, jsonify, render_template

# ── optional heavy imports ────────────────────────────────────────────────────
try:
    import psycopg2, psycopg2.extras
    _PSYCOPG2 = True
except ImportError:
    _PSYCOPG2 = False

try:
    import win32com.client
    _OUTLOOK = True
except ImportError:
    _OUTLOOK = False

try:
    from apscheduler.schedulers.background import BackgroundScheduler
    _SCHEDULER = True
except ImportError:
    _SCHEDULER = False

try:
    from openai import OpenAI
    _OPENAI = True
except ImportError:
    _OPENAI = False


# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════════════════════════
app = Flask(__name__)

# Read DATABASE_URL.  Render still emits "postgres://" — fix it.
_RAW_DB_URL = os.environ.get("DATABASE_URL", "")
DATABASE_URL = (
    _RAW_DB_URL.replace("postgres://", "postgresql://", 1)
    if _RAW_DB_URL.startswith("postgres://")
    else _RAW_DB_URL
)
USE_PG = bool(DATABASE_URL and _PSYCOPG2)

# Secrets — set as environment variables in Render dashboard.
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "sk-REPLACE_WITH_YOUR_OPENAI_API_KEY")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "ADMIN321")

# Tuning
AUTO_ARCHIVE_SECS  = 300   # 5 min after checked  → archived
ARCHIVE_HOURS      = 24    # tasks older than 24 h → archive bucket
URGENT_HOURS       = 48
EMAIL_LOOKBACK_MIN = 180   # read emails from past 3 h

# SQLite path (local only — never used on Render)
_HERE          = os.path.dirname(os.path.abspath(__file__))
SQLITE_PATH    = os.path.join(_HERE, "tasks.db")

# Mailbox map — login email → Outlook store name + scope key
MAILBOX_MAP = {
    "jlyons@cmgfi.com": {
        "outlook_name": "Jill Lyons",
        "mailbox_key":  "jlyons",
        "display_name": "Jill Lyons",
    },
    # Add future users here when their local Outlook is configured:
    # "jreed@gmail.com": {
    #     "outlook_name": "John Reed",
    #     "mailbox_key":  "jreed",
    #     "display_name": "John Reed",
    # },
}


# ═══════════════════════════════════════════════════════════════════════════════
#  DATABASE LAYER
#  All SQL goes through get_conn() / run() / fetch_one() / fetch_all().
#  Placeholder is %s for Postgres, ? for SQLite — _P handles it.
# ═══════════════════════════════════════════════════════════════════════════════
_P = "%s" if USE_PG else "?"   # SQL placeholder


@contextmanager
def get_conn():
    """Yield an open DB connection, commit on success, close always."""
    if USE_PG:
        conn = psycopg2.connect(DATABASE_URL,
                                cursor_factory=psycopg2.extras.RealDictCursor)
    else:
        import sqlite3
        conn = sqlite3.connect(SQLITE_PATH)
        conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _cur(conn):
    """Return a cursor (Postgres needs an explicit cursor; SQLite is its own cursor)."""
    return conn.cursor() if USE_PG else conn


def run(conn, sql, params=()):
    _cur(conn).execute(sql, params)


def fetch_one(conn, sql, params=()):
    cur = _cur(conn)
    cur.execute(sql, params)
    row = cur.fetchone()
    return dict(row) if row else None


def fetch_all(conn, sql, params=()):
    cur = _cur(conn)
    cur.execute(sql, params)
    return [dict(r) for r in cur.fetchall()]


# ═══════════════════════════════════════════════════════════════════════════════
#  SCHEMA
# ═══════════════════════════════════════════════════════════════════════════════
_FLOAT = "DOUBLE PRECISION" if USE_PG else "REAL"


def init_db():
    with get_conn() as conn:
        run(conn, f"""
            CREATE TABLE IF NOT EXISTS users (
                user_id      TEXT PRIMARY KEY,
                email        TEXT UNIQUE NOT NULL,
                approved     INTEGER NOT NULL DEFAULT 0,
                display_name TEXT,
                mailbox_key  TEXT
            )
        """)
        run(conn, f"""
            CREATE TABLE IF NOT EXISTS tasks (
                id          TEXT PRIMARY KEY,
                text        TEXT NOT NULL,
                created_at  {_FLOAT} NOT NULL,
                checked     INTEGER NOT NULL DEFAULT 0,
                checked_at  {_FLOAT},
                archived    INTEGER NOT NULL DEFAULT 0,
                user_id     TEXT
            )
        """)
        run(conn, f"""
            CREATE TABLE IF NOT EXISTS email_tasks (
                id           TEXT PRIMARY KEY,
                sender       TEXT,
                email        TEXT,
                subject      TEXT,
                summary      TEXT,
                reason       TEXT,
                timestamp    {_FLOAT},
                mailbox_key  TEXT
            )
        """)

        # SQLite-only: non-destructive column migrations for older local DBs
        if not USE_PG:
            for tbl, col, defn in [
                ("tasks",       "archived",     "INTEGER NOT NULL DEFAULT 0"),
                ("tasks",       "user_id",      "TEXT"),
                ("users",       "approved",     "INTEGER NOT NULL DEFAULT 0"),
                ("users",       "display_name", "TEXT"),
                ("users",       "mailbox_key",  "TEXT"),
                ("email_tasks", "mailbox_key",  "TEXT"),
            ]:
                try:
                    run(conn, f"ALTER TABLE {tbl} ADD COLUMN {col} {defn}")
                except Exception:
                    pass

        # Seed approved users from MAILBOX_MAP
        for email, cfg in MAILBOX_MAP.items():
            uid = _hash(email)
            run(conn, f"""
                INSERT INTO users (user_id, email, approved, display_name, mailbox_key)
                VALUES ({_P},{_P},1,{_P},{_P})
                ON CONFLICT(email) DO UPDATE SET
                    approved=1,
                    display_name=EXCLUDED.display_name,
                    mailbox_key=EXCLUDED.mailbox_key
            """, (uid, email, cfg["display_name"], cfg["mailbox_key"]))

    print(f"[db] ready  ({'postgres' if USE_PG else 'sqlite'})")


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def _hash(email: str) -> str:
    return hashlib.sha256(email.strip().lower().encode()).hexdigest()


def _task(row: dict) -> dict:
    return {
        "id":         row["id"],
        "text":       row["text"],
        "created_at": float(row["created_at"]),
        "checked":    bool(row["checked"]),
        "checked_at": float(row["checked_at"]) if row.get("checked_at") else None,
        "archived":   bool(row["archived"]),
        "user_id":    row.get("user_id"),
    }


def _day_start(ts: float, tz_offset_mins: int) -> float:
    """UTC epoch of local midnight containing ts."""
    offset    = -tz_offset_mins * 60
    local_ts  = ts + offset
    midnight  = (local_ts // 86400) * 86400
    return midnight - offset


def sweep_archive(now: float):
    cutoff = now - AUTO_ARCHIVE_SECS
    with get_conn() as conn:
        run(conn, f"""
            UPDATE tasks SET archived=1
            WHERE checked=1 AND archived=0
              AND checked_at IS NOT NULL AND checked_at<={_P}
        """, (cutoff,))


# ═══════════════════════════════════════════════════════════════════════════════
#  LAZY INITIALISATION
#  gunicorn imports this module before env vars are fully injected into the
#  worker.  We defer DB init + scheduler start to the first real HTTP request.
# ═══════════════════════════════════════════════════════════════════════════════
_ready = False


@app.before_request
def _boot():
    global _ready
    if _ready:
        return
    _ready = True   # set before init so re-entrant calls don't double-init
    if USE_PG:
        host = DATABASE_URL.split("@")[-1].split("/")[0]
        print(f"[db] postgres host: {host}")
    else:
        print(f"[db] sqlite: {SQLITE_PATH}")
    init_db()
    _start_scheduler()


# ═══════════════════════════════════════════════════════════════════════════════
#  AUTH
# ═══════════════════════════════════════════════════════════════════════════════
@app.route("/login", methods=["POST"])
def login():
    data  = request.get_json(force=True)
    email = (data.get("email") or "").strip().lower()
    if not email:
        return jsonify({"error": "email required"}), 400
    with get_conn() as conn:
        user = fetch_one(conn,
            f"SELECT * FROM users WHERE email={_P} AND approved=1", (email,))
    if not user:
        return jsonify({"error": "Access denied. This email has not been approved."}), 403
    return jsonify({
        "user_id":      user["user_id"],
        "email":        user["email"],
        "display_name": user["display_name"] or user["email"],
        "mailbox_key":  user["mailbox_key"],
    })


# ═══════════════════════════════════════════════════════════════════════════════
#  ADMIN
# ═══════════════════════════════════════════════════════════════════════════════
def _admin_check(data):
    if data.get("admin_password") != ADMIN_PASSWORD:
        return jsonify({"error": "Unauthorized"}), 401
    return None


@app.route("/admin/users", methods=["POST"])
def admin_list_users():
    data = request.get_json(force=True)
    err  = _admin_check(data)
    if err: return err
    with get_conn() as conn:
        rows = fetch_all(conn,
            "SELECT user_id,email,approved,display_name,mailbox_key "
            "FROM users ORDER BY email")
    return jsonify({"users": rows})


@app.route("/admin/add_user", methods=["POST"])
def admin_add_user():
    data = request.get_json(force=True)
    err  = _admin_check(data)
    if err: return err
    email = (data.get("email") or "").strip().lower()
    if not email or "@" not in email:
        return jsonify({"error": "Valid email required"}), 400
    display = (data.get("display_name") or email).strip()
    mkey    = (data.get("mailbox_key")  or email.split("@")[0]).strip()
    uid     = _hash(email)
    with get_conn() as conn:
        run(conn, f"""
            INSERT INTO users (user_id,email,approved,display_name,mailbox_key)
            VALUES ({_P},{_P},1,{_P},{_P})
            ON CONFLICT(email) DO UPDATE SET
                approved=1,
                display_name=EXCLUDED.display_name,
                mailbox_key=EXCLUDED.mailbox_key
        """, (uid, email, display, mkey))
    return jsonify({"added": True, "user_id": uid, "email": email,
                    "display_name": display, "mailbox_key": mkey})


@app.route("/admin/remove_user", methods=["POST"])
def admin_remove_user():
    data  = request.get_json(force=True)
    err   = _admin_check(data)
    if err: return err
    email = (data.get("email") or "").strip().lower()
    if not email:
        return jsonify({"error": "email required"}), 400
    with get_conn() as conn:
        run(conn, f"UPDATE users SET approved=0 WHERE email={_P}", (email,))
    return jsonify({"revoked": email})


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN ROUTES
# ═══════════════════════════════════════════════════════════════════════════════
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/get")
def get_tasks():
    uid = request.args.get("user_id", "").strip()
    if not uid:
        return jsonify({"error": "user_id required"}), 400

    now = time.time()
    sweep_archive(now)

    try:    client_now = float(request.args.get("now", now))
    except: client_now = now
    try:    tz = int(request.args.get("tz", 0))
    except: tz = 0

    today_start = _day_start(client_now, tz)
    archive_cut = now - ARCHIVE_HOURS * 3600
    local_off   = -tz * 60

    def day_of(ts):
        lt = ts + local_off
        return (lt // 86400) * 86400 - local_off   # local midnight in UTC epoch

    with get_conn() as conn:
        rows = fetch_all(conn,
            f"SELECT * FROM tasks WHERE user_id={_P} ORDER BY created_at ASC",
            (uid,))

    all_tasks = [_task(r) for r in rows]

    # Find 3 most-recent days with activity (excluding today)
    past_days = sorted(
        {day_of(t["created_at"]) for t in all_tasks if day_of(t["created_at"]) < today_start},
        reverse=True
    )[:3]
    while len(past_days) < 3:
        past_days.append(None)

    panel_map = {ds: f"day{i+1}" for i, ds in enumerate(past_days) if ds is not None}

    buckets = {"today": [], "day1": [], "day2": [], "day3": [],
               "archive": [], "unchecked_archive": []}

    for t in all_tasks:
        ca    = t["created_at"]
        age_h = (now - ca) / 3600

        if t["archived"]:
            buckets["archive"].append(t); continue

        if ca < archive_cut:
            if t["checked"]: buckets["archive"].append(t)
            else:             t["urgent"] = True; buckets["unchecked_archive"].append(t)
            continue

        if ca >= today_start:
            buckets["today"].append(t)
        else:
            ds = day_of(ca)
            bk = panel_map.get(ds)
            if bk:
                if not t["checked"] and age_h >= URGENT_HOURS:
                    t["urgent"] = True
                buckets[bk].append(t)
            else:
                if t["checked"]: buckets["archive"].append(t)
                else:             t["urgent"] = True; buckets["unchecked_archive"].append(t)

    return jsonify({
        "buckets": buckets,
        "panel_dates": {
            "today": today_start,
            "day1":  past_days[0],
            "day2":  past_days[1],
            "day3":  past_days[2],
        },
        "server_now": now,
    })


@app.route("/add", methods=["POST"])
def add_task():
    data = request.get_json(force=True)
    uid  = (data.get("user_id") or "").strip()
    if not uid:
        return jsonify({"error": "user_id required"}), 400
    texts = data.get("tasks", [])
    if isinstance(texts, str): texts = [texts]
    now     = time.time()
    created = []
    with get_conn() as conn:
        for text in texts:
            text = text.strip()
            if not text: continue
            tid = str(uuid.uuid4())
            run(conn,
                f"INSERT INTO tasks(id,text,created_at,checked,checked_at,archived,user_id)"
                f" VALUES({_P},{_P},{_P},0,NULL,0,{_P})",
                (tid, text, now, uid))
            created.append({"id": tid, "text": text, "created_at": now,
                             "checked": False, "checked_at": None,
                             "archived": False, "user_id": uid})
    return jsonify({"created": created}), 201


@app.route("/toggle", methods=["POST"])
def toggle_task():
    data = request.get_json(force=True)
    tid  = data.get("id")
    uid  = (data.get("user_id") or "").strip()
    if not tid: return jsonify({"error": "id required"}), 400
    if not uid: return jsonify({"error": "user_id required"}), 400
    now = time.time()
    with get_conn() as conn:
        row = fetch_one(conn,
            f"SELECT * FROM tasks WHERE id={_P} AND user_id={_P}", (tid, uid))
        if not row: return jsonify({"error": "not found"}), 404
        task = _task(row)
        new_checked    = not task["checked"]
        new_checked_at = now if new_checked else None
        run(conn,
            f"UPDATE tasks SET checked={_P},checked_at={_P},archived=0"
            f" WHERE id={_P} AND user_id={_P}",
            (int(new_checked), new_checked_at, tid, uid))
    task.update(checked=new_checked, checked_at=new_checked_at, archived=False)
    return jsonify({"task": task})


@app.route("/delete", methods=["POST"])
def delete_task():
    data = request.get_json(force=True)
    tid  = data.get("id")
    uid  = (data.get("user_id") or "").strip()
    if not tid: return jsonify({"error": "id required"}), 400
    with get_conn() as conn:
        run(conn, f"DELETE FROM tasks WHERE id={_P} AND user_id={_P}", (tid, uid))
    return jsonify({"deleted": tid})


@app.route("/archive-now", methods=["POST"])
def archive_now():
    data = request.get_json(force=True)
    tid  = data.get("id")
    uid  = (data.get("user_id") or "").strip()
    if not tid: return jsonify({"error": "id required"}), 400
    if not uid: return jsonify({"error": "user_id required"}), 400
    with get_conn() as conn:
        row = fetch_one(conn,
            f"SELECT id FROM tasks WHERE id={_P} AND user_id={_P}", (tid, uid))
        if not row: return jsonify({"error": "not found"}), 404
        run(conn,
            f"UPDATE tasks SET archived=1,checked=1,checked_at={_P}"
            f" WHERE id={_P} AND user_id={_P}",
            (time.time(), tid, uid))
    return jsonify({"archived": tid})


# ═══════════════════════════════════════════════════════════════════════════════
#  EMAIL PIPELINE
# ═══════════════════════════════════════════════════════════════════════════════
def _read_outlook(outlook_name: str) -> list:
    if not _OUTLOOK:
        print(f"[pipeline] pywin32 not available — skipping {outlook_name}")
        return []
    try:
        ns    = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        store = next((s for s in ns.Stores if s.DisplayName == outlook_name), None)
        if not store:
            store = next((f for f in ns.Folders if f.Name == outlook_name), None)
        if not store:
            print(f"[pipeline] mailbox '{outlook_name}' not found"); return []
        try:    inbox = store.GetRootFolder().Folders["Inbox"]
        except: inbox = store.GetDefaultFolder(6)
        cutoff = time.time() - EMAIL_LOOKBACK_MIN * 60
        out    = []
        for msg in inbox.Items:
            try:
                ts = msg.ReceivedTime.timestamp()
                if ts < cutoff: continue
                out.append({"sender": msg.SenderName,
                            "email":  msg.SenderEmailAddress,
                            "subject": msg.Subject,
                            "body":   msg.Body[:2000],
                            "timestamp": ts})
            except: pass
        return out
    except Exception as e:
        print(f"[pipeline] outlook error: {e}"); return []


def _triage(emails: list) -> list:
    if not _OPENAI or not emails: return []
    client = OpenAI(api_key=OPENAI_API_KEY)
    prompt = (
        "You are an email triage assistant. Return ONLY important/actionable emails.\n"
        "Ignore newsletters, promotions, automated alerts, spam.\n"
        "Return STRICT JSON array only — no markdown, no preamble.\n"
        'Schema: [{"sender","email","subject","summary","reason","timestamp"}]\n'
        "Return [] if nothing is important.\n\nEMAILS:\n"
        + json.dumps(emails, default=str, indent=2)
    )
    try:
        r   = client.chat.completions.create(model="gpt-4.1-mini", max_tokens=1000,
                  messages=[{"role": "user", "content": prompt}])
        raw = r.choices[0].message.content.strip().lstrip("```json").rstrip("```").strip()
        return json.loads(raw)
    except Exception as e:
        print(f"[pipeline] openai error: {e}"); return []


def _store_emails(important: list, mailbox_key: str, user_id: str):
    if not important: return
    now = time.time()
    with get_conn() as conn:
        existing = {r["subject"] for r in fetch_all(conn,
            f"SELECT subject FROM email_tasks WHERE mailbox_key={_P}", (mailbox_key,))}
        for em in important:
            subj = em.get("subject", "")
            if subj in existing: continue
            eid = str(uuid.uuid4())
            run(conn,
                f"INSERT INTO email_tasks(id,sender,email,subject,summary,reason,timestamp,mailbox_key)"
                f" VALUES({_P},{_P},{_P},{_P},{_P},{_P},{_P},{_P})",
                (eid, em.get("sender",""), em.get("email",""), subj,
                 em.get("summary",""), em.get("reason",""), em.get("timestamp", now), mailbox_key))
            tid = str(uuid.uuid4())
            run(conn,
                f"INSERT INTO tasks(id,text,created_at,checked,checked_at,archived,user_id)"
                f" VALUES({_P},{_P},{_P},0,NULL,0,{_P})",
                (tid, f"[Email] {subj} — {em.get('summary','')}", now, user_id))
            existing.add(subj)
    print(f"[pipeline:{mailbox_key}] stored {len(important)} email(s)")


def email_pipeline():
    for login_email, cfg in MAILBOX_MAP.items():
        print(f"[pipeline] running for {cfg['outlook_name']}…")
        emails    = _read_outlook(cfg["outlook_name"])
        important = _triage(emails)
        _store_emails(important, cfg["mailbox_key"], _hash(login_email))
        print(f"[pipeline:{cfg['mailbox_key']}] {len(emails)} read, {len(important)} important")


@app.route("/get_emails")
def get_emails():
    uid = request.args.get("user_id", "").strip()
    if not uid: return jsonify({"error": "user_id required"}), 400
    with get_conn() as conn:
        user = fetch_one(conn,
            f"SELECT mailbox_key FROM users WHERE user_id={_P} AND approved=1", (uid,))
        if not user or not user.get("mailbox_key"): return jsonify([])
        rows = fetch_all(conn,
            f"SELECT * FROM email_tasks WHERE mailbox_key={_P} ORDER BY timestamp DESC",
            (user["mailbox_key"],))
    return jsonify(rows)


# ═══════════════════════════════════════════════════════════════════════════════
#  SCHEDULER
# ═══════════════════════════════════════════════════════════════════════════════
_last_run: dict = {"ts": None, "status": "never run"}


def _run_pipeline():
    _last_run["ts"]     = time.time()
    _last_run["status"] = "running"
    try:
        email_pipeline()
        _last_run["status"] = "ok"
    except Exception as e:
        _last_run["status"] = f"error: {e}"
        print(f"[scheduler] {e}")


def _start_scheduler():
    if not _SCHEDULER:
        print("[scheduler] apscheduler not installed — pipeline will not auto-run"); return
    # Skip if we are the Flask reloader parent (not the real worker)
    if os.environ.get("WERKZEUG_RUN_MAIN") == "false": return
    s = BackgroundScheduler(job_defaults={
        "misfire_grace_time": 600, "coalesce": True, "max_instances": 1})
    s.add_job(_run_pipeline, "interval", hours=3, id="email", next_run_time=None)
    s.start()
    print("[scheduler] email pipeline every 3 h")
    import atexit
    atexit.register(lambda: s.shutdown(wait=False))


@app.route("/run-now", methods=["POST"])
def run_now():
    data = request.get_json(force=True) or {}
    if data.get("admin_password") != ADMIN_PASSWORD:
        return jsonify({"error": "Unauthorized"}), 401
    try:
        _run_pipeline()
        return jsonify({"status": "ok", "ran_at": _last_run["ts"]})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/pipeline-status")
def pipeline_status():
    return jsonify({
        "status":     _last_run["status"],
        "last_run_at": (time.strftime("%Y-%m-%d %H:%M UTC", time.gmtime(_last_run["ts"]))
                        if _last_run["ts"] else None),
        "interval_h": 3,
    })


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT  (gunicorn calls app directly — __main__ only for local dev)
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app.run(debug=True, port=5000)
