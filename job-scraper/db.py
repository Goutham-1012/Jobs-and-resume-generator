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
        """
    )
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
