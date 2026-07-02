"""SQLite storage for scraped jobs and scrape runs."""
import sqlite3
import os
import json
from datetime import datetime, timezone

DB_PATH = os.path.join(os.path.dirname(__file__), "jobs.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            keywords TEXT,
            location TEXT,
            limit_per_actor INTEGER,
            sources TEXT,
            total_jobs INTEGER DEFAULT 0,
            status TEXT DEFAULT 'running',
            log TEXT
        );

        CREATE TABLE IF NOT EXISTS jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            source TEXT,
            title TEXT,
            company TEXT,
            location TEXT,
            salary TEXT,
            posted_date TEXT,
            url TEXT,
            description TEXT,
            scraped_at TEXT,
            dedupe_key TEXT,
            UNIQUE(dedupe_key) ON CONFLICT IGNORE,
            FOREIGN KEY(run_id) REFERENCES runs(id)
        );

        CREATE INDEX IF NOT EXISTS idx_jobs_run ON jobs(run_id);
        CREATE INDEX IF NOT EXISTS idx_jobs_source ON jobs(source);

        CREATE TABLE IF NOT EXISTS resume_queue (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            position INTEGER,
            profile_id TEXT,
            profile_name TEXT,
            label TEXT,
            company TEXT,
            title TEXT,
            job_description TEXT,
            status TEXT DEFAULT 'queued',
            ats INTEGER,
            saved_path TEXT,
            preview TEXT,
            data_json TEXT,
            error TEXT,
            model TEXT,
            job_id INTEGER
        );
        CREATE INDEX IF NOT EXISTS idx_rq_status ON resume_queue(status, position);

        CREATE TABLE IF NOT EXISTS outreach (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            resume_queue_id INTEGER,
            company TEXT,
            job_title TEXT,
            contact_name TEXT,
            contact_title TEXT,
            contact_email TEXT,
            subject TEXT,
            body TEXT,
            status TEXT DEFAULT 'draft',
            error TEXT,
            sent_at TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_outreach_status ON outreach(status);
        """
    )
    # Migrations for pre-existing databases (CREATE TABLE IF NOT EXISTS won't add columns).
    for stmt in (
        "ALTER TABLE jobs ADD COLUMN seen INTEGER DEFAULT 0",
        "ALTER TABLE resume_queue ADD COLUMN model TEXT",
        "ALTER TABLE resume_queue ADD COLUMN job_id INTEGER",
        "ALTER TABLE resume_queue ADD COLUMN started_at TEXT",   # when generation began
        "ALTER TABLE resume_queue ADD COLUMN finished_at TEXT",  # when it finished/failed
    ):
        try:
            conn.execute(stmt)
        except sqlite3.OperationalError:
            pass  # column already exists
    conn.commit()
    conn.close()


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def create_run(keywords, location, limit_per_actor, sources):
    conn = get_conn()
    cur = conn.execute(
        "INSERT INTO runs (created_at, keywords, location, limit_per_actor, sources, status) "
        "VALUES (?, ?, ?, ?, ?, 'running')",
        (now_iso(), keywords, location, limit_per_actor, json.dumps(sources)),
    )
    conn.commit()
    run_id = cur.lastrowid
    conn.close()
    return run_id


def finish_run(run_id, total_jobs, status, log):
    conn = get_conn()
    conn.execute(
        "UPDATE runs SET total_jobs = ?, status = ?, log = ? WHERE id = ?",
        (total_jobs, status, json.dumps(log), run_id),
    )
    conn.commit()
    conn.close()


def insert_jobs(run_id, jobs):
    """Insert normalized job dicts. Returns number of newly inserted rows."""
    conn = get_conn()
    inserted = 0
    for j in jobs:
        dedupe_key = "|".join(
            [
                (j.get("title") or "").strip().lower(),
                (j.get("company") or "").strip().lower(),
                (j.get("location") or "").strip().lower(),
                (j.get("url") or "").strip().lower(),
            ]
        )
        cur = conn.execute(
            "INSERT OR IGNORE INTO jobs "
            "(run_id, source, title, company, location, salary, posted_date, url, description, scraped_at, dedupe_key) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                run_id,
                j.get("source"),
                j.get("title"),
                j.get("company"),
                j.get("location"),
                j.get("salary"),
                j.get("posted_date"),
                j.get("url"),
                j.get("description"),
                now_iso(),
                dedupe_key,
            ),
        )
        inserted += cur.rowcount
    conn.commit()
    conn.close()
    return inserted


def query_jobs(run_id=None, source=None, search=None, sort="scraped_at", order="desc"):
    allowed_sort = {
        "title", "company", "location", "source",
        "posted_date", "salary", "scraped_at",
    }
    if sort not in allowed_sort:
        sort = "scraped_at"
    order = "ASC" if str(order).lower() == "asc" else "DESC"

    sql = "SELECT * FROM jobs WHERE 1=1"
    params = []
    if run_id:
        sql += " AND run_id = ?"
        params.append(run_id)
    if source:
        sql += " AND source = ?"
        params.append(source)
    if search:
        sql += " AND (title LIKE ? OR company LIKE ? OR location LIKE ?)"
        like = f"%{search}%"
        params += [like, like, like]
    sql += f" ORDER BY {sort} {order}"

    conn = get_conn()
    rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
    conn.close()
    return rows


def mark_seen(job_id):
    conn = get_conn()
    conn.execute("UPDATE jobs SET seen = 1 WHERE id = ?", (job_id,))
    conn.commit()
    conn.close()


def list_runs():
    conn = get_conn()
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM runs ORDER BY id DESC LIMIT 50").fetchall()]
    conn.close()
    return rows


def get_stats(run_id=None):
    conn = get_conn()
    sql = "SELECT source, COUNT(*) c FROM jobs"
    params = []
    if run_id:
        sql += " WHERE run_id = ?"
        params.append(run_id)
    sql += " GROUP BY source"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return {r["source"]: r["c"] for r in rows}


# ---------------------------------------------------------------------------
# Resume generation queue (server-side, shared by dashboard + resume page)
# ---------------------------------------------------------------------------
def enqueue_resume(item):
    """Insert a queued item at the end. `item` keys: profile_id, profile_name,
    label, company, title, job_description. Returns the new row id."""
    conn = get_conn()
    pos = (conn.execute("SELECT COALESCE(MAX(position), 0) FROM resume_queue").fetchone()[0]) + 1
    cur = conn.execute(
        "INSERT INTO resume_queue (created_at, position, profile_id, profile_name, "
        "label, company, title, job_description, model, job_id, status) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'queued')",
        (now_iso(), pos, item.get("profile_id"), item.get("profile_name"),
         item.get("label"), item.get("company"), item.get("title"),
         item.get("job_description"), item.get("model"), item.get("job_id")),
    )
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid


def list_resume_queue():
    """All queue items ordered by position (excludes the heavy data_json)."""
    conn = get_conn()
    rows = [dict(r) for r in conn.execute(
        "SELECT id, created_at, position, profile_id, profile_name, label, company, "
        "title, status, ats, saved_path, preview, error, model, started_at, finished_at "
        "FROM resume_queue ORDER BY position ASC, id ASC").fetchall()]
    conn.close()
    return rows


def claim_next_queued():
    """Atomically mark the lowest-position queued item as generating and return it.
    Stamps started_at so the queue can report per-item generation time."""
    conn = get_conn()
    row = conn.execute(
        "UPDATE resume_queue SET status='generating', started_at=?, finished_at=NULL "
        "WHERE id = (SELECT id FROM resume_queue WHERE status='queued' "
        "            ORDER BY position ASC, id ASC LIMIT 1) "
        "RETURNING *",
        (now_iso(),),
    ).fetchone()
    conn.commit()
    result = dict(row) if row else None
    conn.close()
    return result


def update_queue_item(item_id, **fields):
    if not fields:
        return
    # Stamp finish time when an item reaches a terminal state (for timing stats).
    if fields.get("status") in ("done", "error") and "finished_at" not in fields:
        fields["finished_at"] = now_iso()
    cols = ", ".join(f"{k} = ?" for k in fields)
    conn = get_conn()
    conn.execute(f"UPDATE resume_queue SET {cols} WHERE id = ?",
                 (*fields.values(), item_id))
    conn.commit()
    conn.close()


def get_queue_item(item_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM resume_queue WHERE id = ?", (item_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def delete_queue_item(item_id):
    """Delete any item that isn't currently generating. Returns True if removed."""
    conn = get_conn()
    cur = conn.execute(
        "DELETE FROM resume_queue WHERE id = ? AND status != 'generating'", (item_id,))
    conn.commit()
    removed = cur.rowcount > 0
    conn.close()
    return removed


def retry_queue_item(item_id):
    """Re-queue a failed (or done) item: clear its result and put it at the back of the
    queue so the worker regenerates it. Won't touch an item that's currently generating."""
    conn = get_conn()
    pos = conn.execute("SELECT COALESCE(MAX(position), 0) FROM resume_queue").fetchone()[0] + 1
    cur = conn.execute(
        "UPDATE resume_queue SET status='queued', error=NULL, ats=NULL, saved_path=NULL, "
        "preview=NULL, data_json=NULL, started_at=NULL, finished_at=NULL, position=? "
        "WHERE id=? AND status!='generating'",
        (pos, item_id),
    )
    conn.commit()
    ok = cur.rowcount > 0
    conn.close()
    return ok


def requeue_stuck():
    """On startup, return any items left 'generating' (from a crash/restart) to the
    queue so the worker re-processes them. Returns how many were recovered."""
    conn = get_conn()
    cur = conn.execute("UPDATE resume_queue SET status='queued', started_at=NULL "
                       "WHERE status='generating'")
    conn.commit()
    n = cur.rowcount
    conn.close()
    return n


def clear_queue():
    """Remove everything that isn't currently generating."""
    conn = get_conn()
    conn.execute("DELETE FROM resume_queue WHERE status != 'generating'")
    conn.commit()
    conn.close()


def reorder_queue(ordered_ids):
    """Rearrange only the queued items, within the slots they already occupy, leaving
    done/generating items in place; then renumber every row's position to a clean 1..N
    so the list never collapses (and any pre-existing position corruption self-heals)."""
    conn = get_conn()
    rows = [dict(r) for r in conn.execute(
        "SELECT id, status FROM resume_queue ORDER BY position, id").fetchall()]
    full = [r["id"] for r in rows]                       # current display order (all items)
    queued_ids = [r["id"] for r in rows if r["status"] == "queued"]
    queued_slots = [i for i, r in enumerate(rows) if r["status"] == "queued"]

    valid = set(queued_ids)
    new_order = [i for i in ordered_ids if i in valid]   # requested order, queued only
    for q in queued_ids:                                 # append any missing (race-safety)
        if q not in new_order:
            new_order.append(q)

    for slot, qid in zip(queued_slots, new_order):       # queued items back into their slots
        full[slot] = qid

    for pos, iid in enumerate(full, start=1):            # clean, collision-free positions
        conn.execute("UPDATE resume_queue SET position = ? WHERE id = ?", (pos, iid))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Outreach (Apollo contacts + AI-drafted emails, reviewed then sent via Gmail)
# ---------------------------------------------------------------------------
def add_outreach(item):
    conn = get_conn()
    cur = conn.execute(
        "INSERT INTO outreach (created_at, resume_queue_id, company, job_title, "
        "contact_name, contact_title, contact_email, subject, body, status) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft')",
        (now_iso(), item.get("resume_queue_id"), item.get("company"), item.get("job_title"),
         item.get("contact_name"), item.get("contact_title"), item.get("contact_email"),
         item.get("subject"), item.get("body")),
    )
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid


def list_outreach():
    conn = get_conn()
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM outreach ORDER BY id DESC").fetchall()]
    conn.close()
    return rows


def get_outreach(item_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM outreach WHERE id = ?", (item_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def update_outreach(item_id, **fields):
    if not fields:
        return
    cols = ", ".join(f"{k} = ?" for k in fields)
    conn = get_conn()
    conn.execute(f"UPDATE outreach SET {cols} WHERE id = ?", (*fields.values(), item_id))
    conn.commit()
    conn.close()


def delete_outreach(item_id):
    conn = get_conn()
    conn.execute("DELETE FROM outreach WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()


def outreach_exists(resume_queue_id, contact_email):
    """Avoid duplicate drafts for the same résumé+contact."""
    conn = get_conn()
    row = conn.execute(
        "SELECT 1 FROM outreach WHERE resume_queue_id = ? AND contact_email = ?",
        (resume_queue_id, contact_email)).fetchone()
    conn.close()
    return row is not None
