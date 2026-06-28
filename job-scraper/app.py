"""Flask backend: scrape orchestration + dashboard API."""
import os
import json
import uuid
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from dotenv import load_dotenv
from flask import Flask, request, jsonify, render_template

# Load .env BEFORE importing modules that read env vars at import time
# (resume_gen captures OPENAI_MODEL into DEFAULT_MODEL when it is imported).
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

import db
import apify_client
import resume_gen
import apollo_client
import outreach as outreach_mod
import email_send

app = Flask(__name__)
db.init_db()

# Every generated resume is auto-saved here, regardless of download.
GENERATED_DIR = os.path.join(os.path.dirname(__file__), "generated_resumes")
os.makedirs(GENERATED_DIR, exist_ok=True)
# Excel-openable log of every generated resume and the JD it was tailored for.
RESUME_LOG = os.path.join(GENERATED_DIR, "resumes_log.csv")


def _sanitize(text):
    """Make a string safe for a filename: keep alphanumerics, collapse the rest to '_'."""
    import re
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", (text or "").strip()).strip("_")
    return cleaned[:40]  # cap length


def _log_generation(resume_filename, company, title, job_description, ats=""):
    """Append a row to the CSV log; write a header once. Best-effort: if the CSV is
    locked (e.g. open in Excel), skip logging rather than failing the generated resume."""
    import csv
    try:
        new_file = not os.path.exists(RESUME_LOG)
        with open(RESUME_LOG, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            if new_file:
                writer.writerow(["Resume File", "Company", "Title", "ATS %", "Job Description"])
            writer.writerow([resume_filename, (company or "").strip(), (title or "").strip(),
                             str(ats), (job_description or "").strip()])
    except OSError as e:
        print(f"[warn] could not write resumes_log.csv (is it open in Excel?): {e}")


def _save_generated(resume_data, job_description="", company="", title=""):
    """Render the resume to a .docx in GENERATED_DIR, log it, return its path.

    Filename: <Name>_Resume[_<Company>][_<Title>]_<datetime>.docx
    """
    contact_name = (resume_data.get("_contact") or {}).get("name")
    name = (contact_name or resume_data.get("name") or "Resume").title().replace(" ", "_")
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    parts = [name, "Resume"]
    if _sanitize(company):
        parts.append(_sanitize(company))
    if _sanitize(title):
        parts.append(_sanitize(title))
    parts.append(stamp)
    base = "_".join(parts)
    filename = base + ".docx"
    path = os.path.join(GENERATED_DIR, filename)
    counter = 2
    while os.path.exists(path):  # avoid collision on rapid back-to-back saves
        filename = f"{base}_{counter}.docx"
        path = os.path.join(GENERATED_DIR, filename)
        counter += 1
    resume_gen.render_docx(resume_data, path)
    _log_generation(filename, company, title, job_description,
                    resume_data.get("ats_score", ""))
    return path


# ---------------------------------------------------------------------------
# Profiles: multiple résumé identities, stored in a gitignored JSON file.
# ---------------------------------------------------------------------------
PROFILES_PATH = os.path.join(os.path.dirname(__file__), "profiles.json")

# Section field -> heading used when assembling the full résumé text for the model.
PROFILE_SECTIONS = [
    ("summary", "PROFESSIONAL SUMMARY"),
    ("skills", "TECHNICAL SKILLS"),
    ("experience", "PROFESSIONAL EXPERIENCE"),
    ("projects", "PROJECT HIGHLIGHTS"),
    ("education", "EDUCATION"),
    ("certifications", "CERTIFICATIONS"),
]
PROFILE_REQUIRED = ["name", "email", "phone", "linkedin", "github", "portfolio",
                    "summary", "skills", "experience", "education", "employers"]


def load_profiles():
    if os.path.exists(PROFILES_PATH):
        with open(PROFILES_PATH, encoding="utf-8") as f:
            return json.load(f)
    return []


def save_profiles(profiles):
    with open(PROFILES_PATH, "w", encoding="utf-8") as f:
        json.dump(profiles, f, indent=2)


def get_profile(profile_id):
    for p in load_profiles():
        if p.get("id") == profile_id:
            return p
    return None


def seed_default_profile():
    """Create the default (Goutham) profile from the existing base résumé + constants
    on first run, so behavior is unchanged. Stores the full base text as `resume_text`."""
    if load_profiles():
        return
    save_profiles([{
        "id": "default",
        "name": "GOUTHAM REDDY GUNNALA",
        "email": resume_gen.EMAIL,
        "phone": resume_gen.PHONE,
        "linkedin": resume_gen.LINKS["LinkedIn"],
        "github": resume_gen.LINKS["GitHub"],
        "portfolio": resume_gen.LINKS["Portfolio"],
        "employers": list(resume_gen.EXPECTED_COMPANIES),
        "resume_text": resume_gen.load_base_resume(),
    }])


def assemble_resume_text(profile):
    """Full plain-text résumé for the model: raw `resume_text` if present (default
    profile), else assembled from the section fields with their headings."""
    if profile.get("resume_text"):
        return profile["resume_text"]
    parts = [
        profile.get("name", ""),
        f"{profile.get('email','')} | {profile.get('phone','')} | LinkedIn | GitHub | Portfolio",
    ]
    for key, heading in PROFILE_SECTIONS:
        val = (profile.get(key) or "").strip()
        if val:
            parts.append(f"{heading}\n{val}")
    return "\n\n".join(parts)


def profile_contact(profile):
    return {
        "name": profile.get("name", ""),
        "email": profile.get("email", ""),
        "phone": profile.get("phone", ""),
        "linkedin": profile.get("linkedin", ""),
        "github": profile.get("github", ""),
        "portfolio": profile.get("portfolio", ""),
    }


seed_default_profile()


# ---------------------------------------------------------------------------
# Background worker: processes the resume_queue one item at a time.
# ---------------------------------------------------------------------------
import threading
import time


def _process_queue_item(item):
    profile = get_profile(item.get("profile_id"))
    if not profile:
        db.update_queue_item(item["id"], status="error", error="Profile not found")
        return
    try:
        result = resume_gen.generate_resume(
            assemble_resume_text(profile),
            item.get("job_description", ""),
            model=item.get("model") or None,
            expected_companies=profile.get("employers") or None,
        )
        result["_contact"] = profile_contact(profile)
        saved_path = _save_generated(result, item.get("job_description", ""),
                                     item.get("company", ""), item.get("title", ""))
        db.update_queue_item(
            item["id"], status="done", ats=result.get("ats_score"),
            saved_path=os.path.basename(saved_path),
            preview=resume_gen.data_to_text(result),
            data_json=json.dumps(result),
        )
    except Exception as e:  # noqa: BLE001 - surface any failure on the item
        db.update_queue_item(item["id"], status="error", error=str(e)[:300])


def _queue_worker():
    while True:
        item = db.claim_next_queued()
        if not item:
            time.sleep(2)
            continue
        _process_queue_item(item)


_worker_started = False


def _start_worker():
    """Start the single background worker (idempotent). The app runs without the
    Flask reloader (see __main__), so this is one process / one worker; the atomic
    claim in db.claim_next_queued is the safety net if ever run twice."""
    global _worker_started
    if _worker_started:
        return
    _worker_started = True
    db.requeue_stuck()  # recover items left mid-generation by a previous run
    threading.Thread(target=_queue_worker, daemon=True).start()


@app.route("/")
def index():
    return render_template("index.html", sources=apify_client.ACTORS)


@app.route("/api/scrape", methods=["POST"])
def scrape():
    data = request.get_json(force=True) or {}
    keywords = (data.get("keywords") or "").strip()
    location = (data.get("location") or "").strip()
    limit = int(data.get("limit") or 25)
    max_age = data.get("maxAge")  # days; None/empty = any
    max_age = int(max_age) if str(max_age or "").strip().isdigit() else None
    sources = data.get("sources") or list(apify_client.ACTORS.keys())
    career_urls = [u.strip() for u in (data.get("careerUrls") or []) if u.strip()]
    profile = get_profile(data.get("profileId")) if data.get("profileId") else None

    sources = [s for s in sources if s in apify_client.ACTORS]
    if not sources:
        return jsonify({"error": "No valid sources selected"}), 400
    if limit < 1:
        return jsonify({"error": "Limit must be >= 1"}), 400

    keyword_list = [k.strip() for k in keywords.split(",") if k.strip()]
    params = {"keywords": keywords, "keywordList": keyword_list,
              "location": location, "careerUrls": career_urls, "maxAge": max_age}
    run_id = db.create_run(keywords, location, limit, sources)

    log = []
    all_inserted = 0

    def work(src):
        return src, apify_client.run_actor(src, limit, params)

    with ThreadPoolExecutor(max_workers=len(sources)) as pool:
        for src, (jobs, info) in pool.map(work, sources):
            log.append(info)
            if jobs:
                all_inserted += db.insert_jobs(run_id, jobs)

    db.finish_run(run_id, all_inserted, "done", log)

    # Auto-enqueue a tailored resume for each newly-scraped job under the chosen profile.
    queued = 0
    if profile:
        for job in db.query_jobs(run_id=run_id):
            jd = "\n".join(filter(None, [
                job.get("title"),
                f"Company: {job.get('company','')}".strip(),
                f"Location: {job.get('location','')}".strip(),
                "",
                job.get("description") or "",
            ]))
            db.enqueue_resume({
                "profile_id": profile["id"],
                "profile_name": profile.get("name"),
                "label": job.get("title") or "Untitled",
                "company": job.get("company") or "",
                "title": job.get("title") or "",
                "job_id": job.get("id"),
                "job_description": jd,
            })
            queued += 1

    return jsonify({"run_id": run_id, "inserted": all_inserted, "log": log, "queued": queued})


@app.route("/api/jobs")
def jobs():
    run_id = request.args.get("run_id", type=int)
    return jsonify({
        "jobs": db.query_jobs(
            run_id=run_id,
            source=request.args.get("source"),
            search=request.args.get("search"),
            sort=request.args.get("sort", "posted_date"),
            order=request.args.get("order", "desc"),
        ),
        "stats": db.get_stats(run_id),
    })


@app.route("/api/jobs/<int:job_id>/seen", methods=["POST"])
def mark_job_seen(job_id):
    db.mark_seen(job_id)
    return jsonify({"ok": True})


@app.route("/api/runs")
def runs():
    return jsonify({"runs": db.list_runs()})


@app.route("/resume")
def resume_page():
    return render_template("resume.html")


@app.route("/api/base-resume")
def base_resume():
    return jsonify({"resume": resume_gen.load_base_resume()})


@app.route("/api/profiles", methods=["GET"])
def list_profiles():
    return jsonify({"profiles": load_profiles()})


@app.route("/api/profiles", methods=["POST"])
def save_profile():
    p = request.get_json(force=True) or {}
    # Normalize employers (accept comma-separated string or list).
    emp = p.get("employers", "")
    if isinstance(emp, str):
        p["employers"] = [e.strip() for e in emp.split(",") if e.strip()]
    missing = [f for f in PROFILE_REQUIRED if not (p.get(f) if f != "employers" else p.get(f))]
    if missing:
        return jsonify({"error": "Missing required fields: " + ", ".join(missing)}), 400
    p.pop("resume_text", None)  # form profiles use section fields, not raw text
    profiles = load_profiles()
    if p.get("id"):
        profiles = [x for x in profiles if x.get("id") != p["id"]]
    else:
        p["id"] = uuid.uuid4().hex[:12]
    profiles.append(p)
    save_profiles(profiles)
    return jsonify({"profile": p})


@app.route("/api/profiles/<profile_id>", methods=["DELETE"])
def delete_profile(profile_id):
    profiles = load_profiles()
    if len(profiles) <= 1:
        return jsonify({"error": "Cannot delete the last profile."}), 400
    profiles = [x for x in profiles if x.get("id") != profile_id]
    save_profiles(profiles)
    return jsonify({"ok": True})


@app.route("/api/resume", methods=["POST"])
def resume_api():
    data = request.get_json(force=True) or {}
    profile = get_profile(data.get("profileId")) if data.get("profileId") else None
    if profile:
        resume_text = assemble_resume_text(profile)
        expected = profile.get("employers") or None
        contact = profile_contact(profile)
    else:  # backward compatible: no profile -> default base résumé + constants
        resume_text = data.get("resume", "")
        expected = None
        contact = None
    try:
        result = resume_gen.generate_resume(
            resume_text,
            data.get("jobDescription", ""),
            model=data.get("model") or None,
            expected_companies=expected,
        )
        if contact:
            result["_contact"] = contact  # rides along for render/download
        saved_path = _save_generated(result, data.get("jobDescription", ""),
                                     data.get("company", ""), data.get("title", ""))
        return jsonify({
            "data": result,
            "preview": resume_gen.data_to_text(result),
            "savedPath": saved_path,
            "ats": result.get("ats_score"),
        })
    except (ValueError, RuntimeError) as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/resume/download", methods=["POST"])
def resume_download():
    data = request.get_json(force=True) or {}
    resume_data = data.get("data")
    if not resume_data:
        return jsonify({"error": "No resume data"}), 400
    import tempfile
    from flask import send_file
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    resume_gen.render_docx(resume_data, tmp.name)
    name = (resume_data.get("name") or "resume").replace(" ", "_")
    return send_file(tmp.name, as_attachment=True,
                     download_name=f"{name}_tailored.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ---------------------------------------------------------------------------
# Unified resume queue endpoints (background worker processes items)
# ---------------------------------------------------------------------------
@app.route("/api/resume/enqueue", methods=["POST"])
def resume_enqueue():
    data = request.get_json(force=True) or {}
    profile = get_profile(data.get("profileId"))
    if not profile:
        return jsonify({"error": "Select a profile first."}), 400
    jd = (data.get("jobDescription") or "").strip()
    if not jd:
        return jsonify({"error": "Job description is required."}), 400
    rid = db.enqueue_resume({
        "profile_id": profile["id"],
        "profile_name": profile.get("name"),
        "label": (data.get("title") or "").strip() or "Manual",
        "company": (data.get("company") or "").strip(),
        "title": (data.get("title") or "").strip(),
        "job_description": jd,
        "model": (data.get("model") or "").strip() or None,
    })
    return jsonify({"id": rid})


@app.route("/api/resume/queue")
def resume_queue():
    return jsonify({"queue": db.list_resume_queue()})


@app.route("/api/resume/queue/reorder", methods=["POST"])
def resume_queue_reorder():
    ids = (request.get_json(force=True) or {}).get("ids") or []
    db.reorder_queue([int(i) for i in ids])
    return jsonify({"ok": True})


@app.route("/api/resume/queue/<int:item_id>", methods=["DELETE"])
def resume_queue_delete(item_id):
    return jsonify({"removed": db.delete_queue_item(item_id)})


@app.route("/api/resume/queue/<int:item_id>/retry", methods=["POST"])
def resume_queue_retry(item_id):
    return jsonify({"requeued": db.retry_queue_item(item_id)})


@app.route("/api/resume/queue/clear", methods=["POST"])
def resume_queue_clear():
    db.clear_queue()
    return jsonify({"ok": True})


@app.route("/api/resume/queue/<int:item_id>/download")
def resume_queue_download(item_id):
    item = db.get_queue_item(item_id)
    if not item or not item.get("data_json"):
        return jsonify({"error": "Not ready"}), 404
    import tempfile
    from flask import send_file
    resume_data = json.loads(item["data_json"])
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    resume_gen.render_docx(resume_data, tmp.name)
    name = (resume_data.get("name") or "resume").replace(" ", "_")
    return send_file(tmp.name, as_attachment=True,
                     download_name=f"{name}_tailored.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ---------------------------------------------------------------------------
# Outreach: Apollo contacts -> AI draft -> review -> Gmail send (résumé attached)
# ---------------------------------------------------------------------------
@app.route("/outreach")
def outreach_page():
    return render_template("outreach.html")


def _render_resume_attachment(rq_item):
    """Render a résumé-queue item's stored data_json to a temp .docx; return (path, filename)."""
    import tempfile
    data = json.loads(rq_item["data_json"])
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    resume_gen.render_docx(data, tmp.name)
    nm = (data.get("name") or "resume").replace(" ", "_")
    co = _sanitize(rq_item.get("company") or "")
    return tmp.name, f"{nm}_{co}_Resume.docx" if co else f"{nm}_Resume.docx"


@app.route("/api/outreach/candidates")
def outreach_candidates():
    """Completed résumés (each = a job) that can have outreach drafted."""
    done = [r for r in db.list_resume_queue() if r["status"] == "done"]
    drafted = {o["resume_queue_id"] for o in db.list_outreach()}
    for r in done:
        r["has_outreach"] = r["id"] in drafted
    return jsonify({"candidates": done})


@app.route("/api/outreach/draft", methods=["POST"])
def outreach_draft():
    data = request.get_json(force=True) or {}
    rq = db.get_queue_item(data.get("resume_queue_id"))
    if not rq or rq.get("status") != "done":
        return jsonify({"error": "Generate the résumé for this job first."}), 400
    profile = get_profile(rq.get("profile_id")) or (load_profiles() or [None])[0]
    if not profile:
        return jsonify({"error": "No profile found."}), 400

    try:
        contacts, note = apollo_client.find_contacts(rq.get("company"), limit=3)
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 400
    if not contacts:
        return jsonify({"drafts": [], "note": note})

    job = {"company": rq.get("company"), "job_title": rq.get("title"),
           "job_description": rq.get("job_description")}
    created = []
    for c in contacts:
        if db.outreach_exists(rq["id"], c["email"]):
            continue
        try:
            email = outreach_mod.draft_email(profile, job, c)
        except (RuntimeError, ValueError) as e:
            return jsonify({"error": str(e)}), 400
        oid = db.add_outreach({
            "resume_queue_id": rq["id"], "company": rq.get("company"),
            "job_title": rq.get("title"), "contact_name": c["name"],
            "contact_title": c["title"], "contact_email": c["email"],
            "subject": email["subject"], "body": email["body"],
        })
        created.append(oid)
    return jsonify({"drafted": len(created), "note": note})


@app.route("/api/outreach")
def outreach_list():
    return jsonify({"outreach": db.list_outreach()})


@app.route("/api/outreach/<int:item_id>", methods=["PUT"])
def outreach_update(item_id):
    data = request.get_json(force=True) or {}
    fields = {k: data[k] for k in ("subject", "body") if k in data}
    if fields:
        db.update_outreach(item_id, **fields)
    return jsonify({"ok": True})


@app.route("/api/outreach/<int:item_id>", methods=["DELETE"])
def outreach_delete(item_id):
    db.delete_outreach(item_id)
    return jsonify({"ok": True})


def _send_one_outreach(item):
    rq = db.get_queue_item(item.get("resume_queue_id"))
    if not rq or not rq.get("data_json"):
        db.update_outreach(item["id"], status="failed", error="résumé not available")
        return False, "résumé not available"
    path = None
    try:
        path, fname = _render_resume_attachment(rq)
        email_send.send_gmail(item["contact_email"], item["subject"], item["body"], path, fname)
        db.update_outreach(item["id"], status="sent", sent_at=db.now_iso(), error=None)
        return True, "sent"
    except Exception as e:  # noqa: BLE001
        db.update_outreach(item["id"], status="failed", error=str(e)[:300])
        return False, str(e)[:200]
    finally:
        if path and os.path.exists(path):
            os.remove(path)


@app.route("/api/outreach/<int:item_id>/send", methods=["POST"])
def outreach_send(item_id):
    item = db.get_outreach(item_id)
    if not item:
        return jsonify({"error": "not found"}), 404
    ok, msg = _send_one_outreach(item)
    return (jsonify({"sent": True}) if ok else jsonify({"error": msg}), 200 if ok else 400)


@app.route("/api/outreach/send-approved", methods=["POST"])
def outreach_send_approved():
    sent = 0
    for item in db.list_outreach():
        if item.get("status") == "draft":
            ok, _ = _send_one_outreach(item)
            sent += 1 if ok else 0
    return jsonify({"sent": sent})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    _start_worker()
    # Reloader disabled so the single background worker isn't duplicated/orphaned.
    app.run(host="127.0.0.1", port=port, debug=True, use_reloader=False)
