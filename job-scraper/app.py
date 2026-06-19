"""Flask backend: scrape orchestration + dashboard API."""
import os
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from dotenv import load_dotenv
from flask import Flask, request, jsonify, render_template

import db
import apify_client
import resume_gen

load_dotenv()

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
    """Append a row to the CSV log; write a header once."""
    import csv
    new_file = not os.path.exists(RESUME_LOG)
    with open(RESUME_LOG, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(["Resume File", "Company", "Title", "ATS %", "Job Description"])
        writer.writerow([resume_filename, (company or "").strip(), (title or "").strip(),
                         str(ats), (job_description or "").strip()])


def _save_generated(resume_data, job_description="", company="", title=""):
    """Render the resume to a .docx in GENERATED_DIR, log it, return its path.

    Filename: <Name>_Resume[_<Company>][_<Title>]_<datetime>.docx
    """
    name = (resume_data.get("name") or "Resume").title().replace(" ", "_")
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
    return jsonify({"run_id": run_id, "inserted": all_inserted, "log": log})


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


@app.route("/api/resume", methods=["POST"])
def resume_api():
    data = request.get_json(force=True) or {}
    try:
        result = resume_gen.generate_resume(
            data.get("resume", ""),
            data.get("jobDescription", ""),
            model=data.get("model") or None,
        )
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="127.0.0.1", port=port, debug=True)
