"""Apollo.io people search: find recruiters / hiring managers at a company, with emails."""
import os
import requests

APOLLO_URL = "https://api.apollo.io/v1/mixed_people/search"
TIMEOUT = 60

# Titles we target for job outreach, in priority order.
TARGET_TITLES = [
    "technical recruiter", "recruiter", "talent acquisition", "talent partner",
    "hiring manager", "engineering manager", "head of talent", "people operations",
    "human resources",
]


def get_key():
    key = os.environ.get("APOLLO_API_KEY")
    if not key or key.startswith("xxxx"):
        raise RuntimeError("APOLLO_API_KEY not set. Add it to your .env file.")
    return key


def _usable_email(person):
    """Return a real, unlocked email or None (Apollo locks emails behind credits)."""
    email = (person.get("email") or "").strip()
    if not email:
        return None
    low = email.lower()
    if "email_not_unlocked" in low or "domain.com" in low or low in ("email_not_unlocked@domain.com",):
        return None
    if (person.get("email_status") or "").lower() == "unavailable":
        return None
    return email


def find_contacts(company, domain=None, limit=3):
    """Return up to `limit` contacts [{name, title, email}] at `company` who match a
    recruiter/hiring-manager title and have a usable email. Returns (contacts, note)."""
    if not company:
        return [], "no company"
    payload = {
        "q_organization_name": company,
        "person_titles": TARGET_TITLES,
        "page": 1,
        "per_page": 25,
    }
    if domain:
        payload["q_organization_domains"] = domain
    try:
        resp = requests.post(
            APOLLO_URL,
            headers={"Content-Type": "application/json", "Cache-Control": "no-cache",
                     "X-Api-Key": get_key()},
            json=payload, timeout=TIMEOUT,
        )
    except requests.exceptions.RequestException as e:
        return [], f"request error — {e}"
    if resp.status_code >= 400:
        return [], f"Apollo HTTP {resp.status_code} — {resp.text[:160]}"
    try:
        people = resp.json().get("people", [])
    except ValueError:
        return [], "invalid JSON from Apollo"

    contacts = []
    for p in people:
        email = _usable_email(p)
        if not email:
            continue
        name = (p.get("name")
                or f"{p.get('first_name','')} {p.get('last_name','')}".strip()
                or "there")
        contacts.append({"name": name, "title": p.get("title") or "", "email": email})
        if len(contacts) >= limit:
            break

    if not contacts:
        return [], (f"{len(people)} people found but no unlockable emails "
                    "(Apollo free tier limits email reveals)")
    return contacts, f"{len(contacts)} contact(s)"
