"""Draft personalized outreach emails (cheapest model) grounded in the JD + why-fit.
The candidate's links signature is appended deterministically so it's always correct."""
import json
import requests
import resume_gen

DRAFT_MODEL = "gpt-4o-mini"  # cheapest — drafting is simple


def _signature(profile):
    name = profile.get("name") or "Candidate"
    email = profile.get("email") or resume_gen.EMAIL
    phone = profile.get("phone") or resume_gen.PHONE
    links = {
        "Portfolio": profile.get("portfolio") or resume_gen.LINKS.get("Portfolio"),
        "LinkedIn": profile.get("linkedin") or resume_gen.LINKS.get("LinkedIn"),
        "GitHub": profile.get("github") or resume_gen.LINKS.get("GitHub"),
    }
    lines = ["Best regards,", name]
    if email:
        lines.append(email)
    if phone:
        lines.append(phone)
    lines += [f"{k}: {v}" for k, v in links.items() if v]
    return "\n".join(lines)


def draft_email(profile, job, contact):
    """Return {subject, body}. `job` = {company, job_title, job_description}, `contact` =
    {name, title, email}. Body = AI message + appended links signature."""
    name = profile.get("name") or "the candidate"
    blurb = (profile.get("summary") or "").strip()
    if not blurb and profile.get("resume_text"):
        blurb = profile["resume_text"][:1200]

    system = (
        "You write concise, professional job-outreach emails for an applicant. "
        "Return ONLY JSON {\"subject\": ..., \"body\": ...}. "
        "The body is 120-160 words, addressed to the contact by first name, names the specific role "
        "and company, and explains in 2-3 sentences WHY the candidate is a strong fit by mapping the "
        "job description's key requirements to the candidate's real strengths. Mention that a tailored "
        "resume is attached. Be honest and specific; no fabrication, no buzzword stuffing, no EM dashes. "
        "End the body with a one-line opt-out such as 'If now isn't the right time, just let me know and "
        "I won't follow up.' Do NOT add a signature or contact links in the body — those are appended "
        "separately."
    )
    user = (
        f"CANDIDATE: {name}\nCANDIDATE BACKGROUND:\n{blurb[:900]}\n\n"
        f"ROLE: {job.get('job_title')} at {job.get('company')}\n"
        f"JOB DESCRIPTION:\n{(job.get('job_description') or '')[:2500]}\n\n"
        f"CONTACT: {contact.get('name')} — {contact.get('title')}\n"
        "Write the outreach email."
    )
    payload = resume_gen._chat_payload(
        DRAFT_MODEL,
        [{"role": "system", "content": system}, {"role": "user", "content": user}],
        0.6,
    )
    resp = requests.post(
        resume_gen.OPENAI_URL,
        headers={"Authorization": f"Bearer {resume_gen.get_key()}", "Content-Type": "application/json"},
        json=payload, timeout=120,
    )
    if resp.status_code >= 400:
        raise RuntimeError(f"OpenAI HTTP {resp.status_code}: {resp.text[:200]}")
    data = json.loads(resp.json()["choices"][0]["message"]["content"])
    subject = (data.get("subject") or
               f"Application for {job.get('job_title')} at {job.get('company')}").strip()
    body = (data.get("body") or "").strip() + "\n\n" + _signature(profile)
    return {"subject": subject, "body": body}
