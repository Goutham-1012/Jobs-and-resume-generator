#!/usr/bin/env python
"""Standalone CLI around resume_gen — lets career-ops (or any external tool)
generate the user's EXACT .docx resume for a given job description, without the
Flask app or a browser.

This is the bridge that lets career-ops delegate CV generation to this project
instead of its own Playwright HTML->PDF flow. It reuses the same functions the
web UI uses (resume_gen.generate_resume + render_docx), so the .docx it produces
is byte-for-byte the format you already send to recruiters.

Usage:
  python cli.py --jd-file jd.txt --out output.docx [--company Acme] [--title "ML Engineer"]
  echo "<job description text>" | python cli.py --out output.docx
  python cli.py --jd "<inline jd text>" --out output.docx --model gpt-5-mini

On success prints ONE JSON line to stdout:
  {"ok": true, "path": "...abs.docx", "company": "...", "title": "...", "ats_score": 87}
On failure prints {"ok": false, "error": "..."} and exits non-zero.
"""
import argparse
import json
import os
import sys

from dotenv import load_dotenv

# Load .env from this file's directory so OPENAI_API_KEY / OPENAI_MODEL are honored
# regardless of the caller's working directory (career-ops runs from its own dir).
_HERE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(_HERE, ".env"))

import resume_gen  # noqa: E402  (import after load_dotenv so model/env resolve correctly)


def _read_jd(args):
    if args.jd_file:
        with open(args.jd_file, encoding="utf-8") as fh:
            return fh.read()
    if args.jd:
        return args.jd
    if not sys.stdin.isatty():  # piped input
        return sys.stdin.read()
    return ""


def main():
    parser = argparse.ArgumentParser(description="Generate the user's exact .docx resume for a JD.")
    parser.add_argument("--jd-file", help="Path to a file containing the job description.")
    parser.add_argument("--jd", help="Job description text passed inline.")
    parser.add_argument("--out", required=True, help="Output .docx path.")
    parser.add_argument("--company", default="", help="Company name (metadata only).")
    parser.add_argument("--title", default="", help="Role title (metadata only).")
    parser.add_argument("--model", help="Override OPENAI_MODEL for this run.")
    parser.add_argument("--base-resume", help="Override the base resume text file (default: base_resume.txt).")
    args = parser.parse_args()

    jd = (_read_jd(args) or "").strip()
    if not jd:
        print(json.dumps({"ok": False, "error": "No job description provided (use --jd-file, --jd, or stdin)."}))
        sys.exit(2)

    if args.base_resume:
        with open(args.base_resume, encoding="utf-8") as fh:
            base = fh.read()
    else:
        base = resume_gen.load_base_resume()

    try:
        data = resume_gen.generate_resume(
            base, jd, model=args.model, expected_companies=resume_gen.EXPECTED_COMPANIES
        )
    except Exception as exc:  # surface a clean JSON error to the caller
        print(json.dumps({"ok": False, "error": f"generation failed: {exc}"}))
        sys.exit(1)

    out_path = os.path.abspath(args.out)
    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    try:
        resume_gen.render_docx(data, out_path)
    except Exception as exc:
        print(json.dumps({"ok": False, "error": f"render failed: {exc}"}))
        sys.exit(1)

    result = {"ok": True, "path": out_path, "company": args.company, "title": args.title}
    try:
        result["ats_score"] = resume_gen.ats_score(data, jd)
    except Exception:
        pass  # score is advisory; never fail the run over it
    print(json.dumps(result))


if __name__ == "__main__":
    main()
