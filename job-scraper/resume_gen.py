"""Resume generator — replicates the Claude /resume-generator skill via the OpenAI API,
but enforces the user's EXACT resume format and renders a matching .docx."""
import os
import json
import requests
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

# Known personal links (rendered as real hyperlinks in the contact line).
LINKS = {
    "LinkedIn": "https://www.linkedin.com/in/gouthamgunnala/",
    "GitHub": "https://github.com/Goutham-1012",
    "Portfolio": "https://goutham-1012.github.io/My-Portfolio/",
}
EMAIL = "gunnalagouthamreddy0@gmail.com"
PHONE = "913-406-5191"
HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

OPENAI_URL = "https://api.openai.com/v1/chat/completions"
DEFAULT_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o")

FONT = "Century Gothic"
BASE_RESUME_PATH = os.path.join(os.path.dirname(__file__), "base_resume.txt")

SYSTEM_PROMPT = """ROLE
You are a top 10 percent resume strategist specializing in Data Scientist and AI/ML Engineer resumes. You understand ATS systems, recruiter screening behavior, and hiring manager expectations. Transform the resume so it strongly aligns with the target job description while staying authentic and credible.

Do NOT change: company names, locations, education, employment dates, total years of experience.
You MAY modify: professional summary, job titles (only when realistic), responsibilities, skills, technologies/tools, metrics, KPIs, dashboards.

STRATEGY
- Analyze the job description first: extract core technical skills, languages, frameworks, data/cloud platforms, ETL/pipeline tools, visualization tools, databases, domain knowledge, responsibilities, expected outcomes. Identify mandatory vs preferred skills and technical keywords.
- Every extracted technology/tool/platform/framework/language must appear in the Skills section and be reflected in experience bullets when realistic.
- Achieve greater than 90 percent ATS keyword alignment. Never copy sentences from the job description.
- Most recent role must include 100 percent of mandatory skills; preferred skills there too when realistic. Earlier roles reflect ~80-90 percent of relevant skills when plausible.
- Every experience/project bullet must include a technical action, the tools/technologies used, and a measurable outcome or operational impact.
- Maintain domain realism and a believable technology progression (earlier roles simpler, recent roles deeper).
- No EM dashes. No generic adjectives (dedicated, hardworking, adaptable, motivated, team player). No fluff. No keyword stuffing.

PRECEDENCE RULE (read carefully)
- For CONTENT QUALITY, follow ALL skill rules above EXACTLY: >90 percent ATS keyword alignment, mandatory skills present in the most recent role at 100 percent, every bullet contains a technical action + tools/technologies + a measurable outcome, skill distribution across roles, domain consistency, believable technology progression, no EM dashes, no generic adjectives, no copied JD sentences, post-generation keyword audit, and the final ATS analysis.
- For STRUCTURE AND FORMAT, follow the user's existing resume EXACTLY (this overrides the skill's "max 7 skills" and any other layout limit).

STRUCTURE REQUIREMENTS (from the user's resume — these override skill layout limits)
- "summary": 3 to 4 dense paragraph-style lines (NOT bullets), each one or two sentences, keyword rich.
- "skills": keep the SAME categorized style as the original, one entry per category formatted as "Category Name: item, item, item". KEEP roughly 12 to 15 categories like the original (do NOT condense to 7). Reorder and augment categories/items to surface the job description's mandatory and preferred keywords first, while staying realistic.
- "experience": keep every company, location, and date EXACTLY as in the original. Rewrite all bullets (keep a similar count per role, typically 6 to 8 for recent roles). APPLY THE SKILL'S ROLE POSITIONING to titles based on what the JD emphasizes: SQL / reporting / dashboards / KPIs => Data Analyst; ETL / Spark / Airflow pipelines => Data Engineer; data modeling / warehouse / transformations => Analytics Engineer; otherwise keep an AI/ML title. Reposition titles when the JD justifies it; never change the company, location, or dates.
- "projects": keep the Project Highlights section with project names and rewritten bullets.
- "education" and "certifications": keep exactly as provided.

HARD REQUIREMENTS (non-negotiable — verify before returning)
1. EVERY bullet in EVERY experience role AND every project MUST contain at least one quantified metric (a number, percent, time saved, throughput, dataset size, accuracy, latency, or cost). No bullet may be metric-free.
1b. EVERY experience bullet must be detailed and substantial: roughly 22 to 38 words (about two full lines), combining the technical action, the specific tools/technologies used, AND the quantified business or operational outcome. Do NOT write short single-line bullets. Project bullets should be at least 18 words.
2. EVERY tool, technology, framework, language, platform, or orchestration tool named explicitly in the job description MUST appear verbatim in the "skills" section AND must appear in the MOST RECENT role's bullets.
3. The most recent role must contain 100 percent of the JD's mandatory skills, named explicitly.
4. Use the literal token "LLM" or "LLMs" somewhere in the skills section.
5. Zero EM dashes (—) and zero EN dashes used as separators in any rewritten text. Zero generic adjectives.
6. Do not invent new companies, locations, or dates. Keep education and certifications verbatim.

OUTPUT
Return ONLY a JSON object (no markdown) with this exact shape:
{
  "name": "string",
  "contact": "string (single line: email | phone | LinkedIn | GitHub | Portfolio)",
  "summary": ["line 1", "line 2", "line 3", "line 4"],
  "skills": ["Category: items", "Category: items", ...],
  "experience": [
    {"title": "string", "dates": "string", "company_location": "string", "bullets": ["...", "..."]}
  ],
  "projects": [
    {"name": "string", "bullets": ["...", "..."]}
  ],
  "education": ["line", "line"],
  "certifications": ["line", "line"],
  "analysis": "Plain-text ATS analysis: JD keyword breakdown, mandatory skills, preferred skills, before vs after keywords added, ATS alignment estimate, and how the resume was improved."
}
"""


def get_key():
    key = os.environ.get("OPENAI_API_KEY")
    if not key or key.startswith("sk-xxxx"):
        raise RuntimeError("OPENAI_API_KEY not set. Add it to your .env file.")
    return key


def load_base_resume():
    if os.path.exists(BASE_RESUME_PATH):
        with open(BASE_RESUME_PATH, encoding="utf-8") as f:
            return f.read()
    return ""


def generate_resume(resume_text, job_description, model=None):
    """Call OpenAI (JSON mode) and return the structured resume dict."""
    resume_text = (resume_text or "").strip() or load_base_resume()
    if not resume_text:
        raise ValueError("Original resume is required.")
    if not (job_description or "").strip():
        raise ValueError("Target job description is required.")

    user_msg = (
        "ORIGINAL RESUME (preserve this exact structure and formatting):\n"
        f"{resume_text}\n\n"
        "TARGET JOB DESCRIPTION:\n"
        f"{job_description.strip()}\n\n"
        "Rewrite the resume to align with the job description and return the JSON object."
    )

    payload = {
        "model": model or DEFAULT_MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_msg},
        ],
        "temperature": 0.4,
        "response_format": {"type": "json_object"},
    }

    resp = requests.post(
        OPENAI_URL,
        headers={"Authorization": f"Bearer {get_key()}",
                 "Content-Type": "application/json"},
        json=payload,
        timeout=180,
    )
    if resp.status_code >= 400:
        raise RuntimeError(f"OpenAI HTTP {resp.status_code}: {resp.text[:300]}")
    content = resp.json()["choices"][0]["message"]["content"]
    data = json.loads(content)

    # Automatic audit + repair loop: iterate until the draft passes the skill's
    # hard rules, or up to 3 passes.
    best = data
    for _ in range(6):
        problems = _audit(data, job_description)
        if not problems:
            best = data
            break
        # Keep the draft with the fewest outstanding problems as a fallback.
        if len(problems) <= len(_audit(best, job_description)):
            best = data
        data = _repair(data, job_description, problems, model or DEFAULT_MODEL)
    else:
        # Loop exhausted: return whichever draft had the fewest problems.
        if len(_audit(data, job_description)) <= len(_audit(best, job_description)):
            best = data
    return best


# Concrete named technologies that, if present in the JD, must surface in the resume.
_RECENT_TECH = {"python", "sql", "airflow", "prefect", "llm", "llms", "agentic",
                "pytorch", "tensorflow", "spark", "kafka", "docker", "kubernetes"}
_SKILLS_TECH = _RECENT_TECH | {"forecasting", "classification", "optimization",
                               "statistics", "feature engineering", "orchestration",
                               "pipelines", "rag", "fine-tuning"}


def _jd_terms(jd, vocab):
    low = jd.lower()
    return sorted({t for t in vocab if t in low})


def _audit(d, jd):
    """Return concrete, fixable problems (with the exact offending content)."""
    import re
    issues = []

    no_metric = [b for e in d.get("experience", []) for b in e.get("bullets", [])
                 if not re.search(r"\d", b)]
    no_metric += [b for p in d.get("projects", []) for b in p.get("bullets", [])
                  if not re.search(r"\d", b)]
    if no_metric:
        issues.append("Rewrite these EXACT bullets to include a realistic quantified metric "
                      "(percent, time, volume, accuracy), keeping their meaning:\n   - "
                      + "\n   - ".join(no_metric))

    short = [b for e in d.get("experience", []) for b in e.get("bullets", [])
             if len(b.split()) < 22]
    if short:
        issues.append("Expand these EXACT experience bullets to 22-38 words by adding the tools "
                      "used and the business/operational outcome (keep the metric):\n   - "
                      + "\n   - ".join(short))

    skills_txt = json.dumps(d.get("skills", [])).lower()
    recent_txt = json.dumps(d["experience"][0]).lower() if d.get("experience") else ""
    miss_skills = [t for t in _jd_terms(jd, _SKILLS_TECH) if t not in skills_txt]
    if miss_skills:
        issues.append("Add these JD terms verbatim into the skills section: " + ", ".join(miss_skills))

    # Recent-role check. Interchangeable tool groups only need ONE present
    # (the JD lists them as alternatives, e.g. "Airflow, Prefect").
    groups = [{"airflow", "prefect"}, {"pytorch", "tensorflow"}]
    standalone = _RECENT_TECH - set().union(*groups)
    miss_recent = [t for t in _jd_terms(jd, standalone) if t not in recent_txt]
    for g in groups:
        in_jd = _jd_terms(jd, g)
        if in_jd and not any(t in recent_txt for t in in_jd):
            miss_recent.append("/".join(in_jd) + " (at least one)")
    if miss_recent:
        issues.append("The most recent role (first experience entry) must mention these JD "
                      "terms naturally in its bullets: " + ", ".join(miss_recent))

    if "—" in json.dumps(d):
        issues.append("Remove all EM dashes.")
    full = json.dumps(d).lower()
    bad_adj = [w for w in ("dedicated", "hardworking", "adaptable", "motivated", "team player") if w in full]
    if bad_adj:
        issues.append("Remove generic adjectives: " + ", ".join(bad_adj))
    return issues


def _repair(d, jd, problems, model):
    """Send the draft back to the model with a targeted fix list."""
    msg = (
        "Here is a draft resume JSON. Make MINIMAL edits to fix ONLY the listed problems. "
        "Keep EVERYTHING ELSE byte-for-byte identical: same JSON shape, same number of skill "
        "categories, same companies, locations, dates, titles, education, and certifications. "
        "Do not drop or reword any bullet or skill that is not in the problem list. "
        "Do NOT add new bullets or filler bullets; keep the same number of bullets per role. "
        "When a required term must appear, weave it naturally into an EXISTING strong bullet "
        "that already has a metric, never as a new short bullet. "
        "Never shorten a bullet; experience bullets must stay 22-38 words. "
        "Every existing quantified metric must be preserved. Return the corrected JSON only.\n\n"
        "PROBLEMS TO FIX:\n" + "\n".join(problems) +
        "\n\nJOB DESCRIPTION:\n" + jd.strip() +
        "\n\nDRAFT JSON:\n" + json.dumps(d)
    )
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": msg},
        ],
        "temperature": 0.3,
        "response_format": {"type": "json_object"},
    }
    resp = requests.post(
        OPENAI_URL,
        headers={"Authorization": f"Bearer {get_key()}",
                 "Content-Type": "application/json"},
        json=payload, timeout=180,
    )
    if resp.status_code >= 400:
        return d  # keep the draft if repair call fails
    try:
        return json.loads(resp.json()["choices"][0]["message"]["content"])
    except (KeyError, ValueError):
        return d


def data_to_text(d):
    """Plain-text preview mirroring the resume structure."""
    out = [d.get("name", ""), d.get("contact", ""), "", "PROFESSIONAL SUMMARY"]
    out += d.get("summary", [])
    out += ["", "TECHNICAL SKILLS"] + d.get("skills", [])
    out += ["", "PROFESSIONAL EXPERIENCE"]
    for e in d.get("experience", []):
        out += ["", f"{e.get('title','')}    {e.get('dates','')}", e.get("company_location", "")]
        out += [f"- {b}" for b in e.get("bullets", [])]
    if d.get("projects"):
        out += ["", "PROJECT HIGHLIGHTS"]
        for p in d["projects"]:
            out += ["", p.get("name", "")] + [f"- {b}" for b in p.get("bullets", [])]
    out += ["", "EDUCATION"] + d.get("education", [])
    if d.get("certifications"):
        out += ["", "Certifications"] + d.get("certifications", [])
    if d.get("analysis"):
        out += ["", "=" * 60, "ATS ANALYSIS", "=" * 60, d["analysis"]]
    return "\n".join(out)


# ---------------------------------------------------------------------------
# DOCX rendering — reproduces the original styling (Century Gothic, sizes, etc.)
# ---------------------------------------------------------------------------
def _set_font(run, size, bold=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold


def _add_bottom_border(p):
    """Draw a horizontal divider line under a paragraph (section heading)."""
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pbdr.append(bottom)
    pPr.append(pbdr)


def _add_hyperlink(paragraph, url, text, size=10.0):
    r_id = paragraph.part.relate_to(url, HYPERLINK_REL, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), FONT)
    fonts.set(qn("w:hAnsi"), FONT)
    rpr.append(fonts)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2))); rpr.append(sz)
    color = OxmlElement("w:color"); color.set(qn("w:val"), "0563C1"); rpr.append(color)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rpr.append(u)
    run.append(rpr)
    t = OxmlElement("w:t"); t.text = text; run.append(t)
    link.append(run)
    paragraph._p.append(link)


def _heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2)
    _set_font(p.add_run(text.upper()), 7.5, bold=True)
    _add_bottom_border(p)
    return p


def _body(doc, text, justify=False, bold=False, size=7.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_font(p.add_run(text), size, bold=bold)
    return p


def _bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(1)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_font(p.add_run(text), 7.5)
    return p


def _xml_safe(s):
    """Replace Windows-1252 control-char dashes/quotes and strip other control chars."""
    if not isinstance(s, str):
        return s
    repl = {"\x91": "'", "\x92": "'", "\x93": '"', "\x94": '"',
            "\x96": "-", "\x97": "-", "\x95": "-", "�": "-", "\x85": "..."}
    for bad, good in repl.items():
        s = s.replace(bad, good)
    return "".join(c for c in s if c == "\n" or ord(c) >= 32)


def _clean(obj):
    if isinstance(obj, str):
        return _xml_safe(obj)
    if isinstance(obj, list):
        return [_clean(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    return obj


def render_docx(d, out_path):
    d = _clean(d)
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Inches(1)
        s.left_margin = s.right_margin = Inches(1)
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(7.5)

    # Header
    name_p = doc.add_paragraph()
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_p.paragraph_format.space_after = Pt(0)
    _set_font(name_p.add_run(d.get("name", "")), 14, bold=True)

    # Contact line with real hyperlinks: email | phone | LinkedIn | GitHub | Portfolio
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_font(cp.add_run(EMAIL), 10)
    _set_font(cp.add_run(f"  |  {PHONE}  |  "), 10)
    for i, (label, url) in enumerate(LINKS.items()):
        if i:
            _set_font(cp.add_run("  |  "), 10)
        _add_hyperlink(cp, url, label, 10)

    _heading(doc, "Professional Summary")
    for line in d.get("summary", []):
        _bullet(doc, line)

    _heading(doc, "Technical Skills")
    for line in d.get("skills", []):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if ":" in line:
            cat, rest = line.split(":", 1)
            _set_font(p.add_run(cat + ":"), 7.5, bold=True)
            _set_font(p.add_run(rest), 7.5)
        else:
            _set_font(p.add_run(line), 7.5)

    _heading(doc, "Professional Experience")
    for e in d.get("experience", []):
        head = doc.add_paragraph()
        head.paragraph_format.space_before = Pt(3)
        head.paragraph_format.space_after = Pt(0)
        # title left, dates right via a tab stop
        from docx.enum.text import WD_TAB_ALIGNMENT
        head.paragraph_format.tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)
        r1 = head.add_run(e.get("title", "")); _set_font(r1, 7.5, bold=True)
        r1.add_tab()
        r2 = head.add_run(e.get("dates", "")); _set_font(r2, 7.5, bold=True)
        _body(doc, e.get("company_location", ""), bold=True)
        for b in e.get("bullets", []):
            _bullet(doc, b)

    if d.get("projects"):
        _heading(doc, "Project Highlights")
        for p in d["projects"]:
            _body(doc, p.get("name", ""), bold=True)
            for b in p.get("bullets", []):
                _bullet(doc, b)

    _heading(doc, "Education")
    for line in d.get("education", []):
        _body(doc, line)

    if d.get("certifications"):
        _heading(doc, "Certifications")
        for line in d.get("certifications", []):
            _body(doc, line, bold=True)

    doc.save(out_path)
    return out_path
