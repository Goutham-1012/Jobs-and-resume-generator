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


def _save_generated(resume_data):
    """Render the resume to a timestamped .docx in GENERATED_DIR and return its path."""
    name = (resume_data.get("name") or "Resume").title().replace(" ", "_")
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    path = os.path.join(GENERATED_DIR, f"{name}_Resume_{stamp}.docx")
    resume_gen.render_docx(resume_data, path)
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
        saved_path = _save_generated(result)
        return jsonify({
            "data": result,
            "preview": resume_gen.data_to_text(result),
            "savedPath": saved_path,
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
