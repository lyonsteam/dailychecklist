from flask import Flask, request, jsonify, render_template
import sqlite3
import time
import uuid
import os

app = Flask(__name__)
DB_PATH = os.path.join(os.path.dirname(__file__), 'tasks.db')

AUTO_ARCHIVE_CHECKED_SECS = 300   # 5 minutes after checked → archived
URGENT_HOURS              = 48
ARCHIVE_HOURS             = 96


# ── Database ──────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS tasks (
                id          TEXT PRIMARY KEY,
                text        TEXT NOT NULL,
                created_at  REAL NOT NULL,
                checked     INTEGER NOT NULL DEFAULT 0,
                checked_at  REAL,
                archived    INTEGER NOT NULL DEFAULT 0
            )
        ''')
        # Non-destructive migration for existing DBs
        try:
            conn.execute('ALTER TABLE tasks ADD COLUMN archived INTEGER NOT NULL DEFAULT 0')
        except Exception:
            pass
        conn.commit()


init_db()


# ── Helpers ───────────────────────────────────────────────────────────────────

def row_to_dict(row):
    return {
        'id':         row['id'],
        'text':       row['text'],
        'created_at': row['created_at'],
        'checked':    bool(row['checked']),
        'checked_at': row['checked_at'],
        'archived':   bool(row['archived']),
    }


def local_day_start(ts: float, tz_offset_mins: int) -> float:
    """
    Return UTC epoch for local midnight of the day containing `ts`.

    JS getTimezoneOffset() is sign-inverted vs UTC offset:
      EST (UTC-5)  → +300 minutes
      IST (UTC+5:30) → -330 minutes
    local_offset_secs = -tz_offset_mins * 60
    """
    local_offset = -tz_offset_mins * 60
    local_ts     = ts + local_offset
    midnight     = (local_ts // 86400) * 86400
    return midnight - local_offset


def sweep_auto_archive(now: float):
    """Promote checked items that have been checked for 5+ minutes to archived."""
    cutoff = now - AUTO_ARCHIVE_CHECKED_SECS
    with get_db() as conn:
        conn.execute(
            '''UPDATE tasks SET archived = 1
               WHERE checked = 1 AND archived = 0
                 AND checked_at IS NOT NULL AND checked_at <= ?''',
            (cutoff,)
        )
        conn.commit()


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/get', methods=['GET'])
def get_tasks():
    """
    Query params:
      tz  – JS getTimezoneOffset() integer (minutes, sign-inverted vs UTC)
      now – client Unix timestamp in seconds

    Buckets returned:
      today / day1 / day2 / day3  – live calendar-day panels
      archive                      – checked + auto-archived items
      unchecked_archive            – unchecked items > 96 h old
    panel_dates keys map to UTC midnight timestamps for each panel.
    """
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
        rows = conn.execute('SELECT * FROM tasks ORDER BY created_at ASC').fetchall()

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
    data  = request.get_json(force=True)
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
                'INSERT INTO tasks (id, text, created_at, checked, checked_at, archived) '
                'VALUES (?,?,?,0,NULL,0)',
                (tid, text, now)
            )
            created.append({
                'id': tid, 'text': text, 'created_at': now,
                'checked': False, 'checked_at': None, 'archived': False,
            })
        conn.commit()

    return jsonify({'created': created}), 201


@app.route('/toggle', methods=['POST'])
def toggle_task():
    data    = request.get_json(force=True)
    task_id = data.get('id')
    if not task_id:
        return jsonify({'error': 'id required'}), 400

    now = time.time()
    with get_db() as conn:
        row = conn.execute('SELECT * FROM tasks WHERE id=?', (task_id,)).fetchone()
        if not row:
            return jsonify({'error': 'not found'}), 404
        task           = row_to_dict(row)
        new_checked    = not task['checked']
        new_checked_at = now if new_checked else None
        conn.execute(
            'UPDATE tasks SET checked=?, checked_at=?, archived=0 WHERE id=?',
            (int(new_checked), new_checked_at, task_id)
        )
        conn.commit()
        task.update(checked=new_checked, checked_at=new_checked_at, archived=False)

    return jsonify({'task': task})


@app.route('/delete', methods=['POST'])
def delete_task():
    data    = request.get_json(force=True)
    task_id = data.get('id')
    if not task_id:
        return jsonify({'error': 'id required'}), 400
    with get_db() as conn:
        conn.execute('DELETE FROM tasks WHERE id=?', (task_id,))
        conn.commit()
    return jsonify({'deleted': task_id})


if __name__ == '__main__':
    app.run(debug=True, port=5000)
