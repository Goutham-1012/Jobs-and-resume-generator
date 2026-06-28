"""Apify integration: actor registry, input builders, and output normalizers.

Every actor receives the SAME shared `limit` (number of jobs to scrape per actor),
mapped to whatever input field that particular actor expects.
"""
import os
import threading
from datetime import datetime, timedelta, timezone
import requests

APIFY_BASE = "https://api.apify.com/v2"
TIMEOUT = 300  # seconds per actor (run-sync)

# Token rotation across multiple Apify accounts (each free tier ~$5/month). When one
# token is exhausted/invalid we fall back to the next and remember the working one.
_token_lock = threading.Lock()
_token_start = 0  # index of the token to try first


def _first(d, *keys):
    """Return the first present, non-empty value among keys (supports nested via dotted path)."""
    for k in keys:
        if "." in k:
            cur = d
            ok = True
            for part in k.split("."):
                if isinstance(cur, dict) and part in cur:
                    cur = cur[part]
                else:
                    ok = False
                    break
            if ok and cur not in (None, "", []):
                return cur
        elif k in d and d[k] not in (None, "", []):
            return d[k]
    return None


def _fmt_pay(cur, mn, mx, unit):
    cur = cur or ""
    unit = f"/{unit.lower()}" if unit else ""
    if mn and mx and mn != mx:
        return f"{cur} {mn:,}-{mx:,}{unit}".strip()
    one = mn or mx
    return f"{cur} {one:,}{unit}".strip() if one else None


def _salary(item):
    # 1) fantastic-jobs (LinkedIn) AI-enriched fields
    if item.get("ai_salary_min_value") or item.get("ai_salary_max_value"):
        return _fmt_pay(item.get("ai_salary_currency"),
                        item.get("ai_salary_min_value"),
                        item.get("ai_salary_max_value"),
                        item.get("ai_salary_unit_text"))
    # 2) schema.org MonetaryAmount nested under salary/salary_raw (LinkedIn)
    for key in ("salary", "salary_raw"):
        v = item.get(key)
        if isinstance(v, dict) and isinstance(v.get("value"), dict):
            val = v["value"]
            r = _fmt_pay(v.get("currency"), val.get("minValue"),
                         val.get("maxValue"), val.get("unitText"))
            if r:
                return r
    # 3) Glassdoor-style "pay" dict
    p = item.get("pay")
    if isinstance(p, dict) and (p.get("min") or p.get("max")):
        return _fmt_pay(p.get("currency"), p.get("min"), p.get("max"), p.get("period"))
    # 4) Indeed-style flat salary dict
    v = item.get("salary")
    if isinstance(v, dict):
        if v.get("text"):
            return v["text"]
        r = _fmt_pay(v.get("currency"), v.get("min"), v.get("max"), v.get("period"))
        if r:
            return r
        return None
    # 4) plain string
    return v if isinstance(v, str) else None


# ---------------------------------------------------------------------------
# Normalizers: convert each actor's raw item into our common schema.
# Actors vary in field names, so we probe several likely keys.
# ---------------------------------------------------------------------------
def _age_to_date(age_days):
    """Glassdoor gives ageInDays instead of a timestamp — convert to a date."""
    if age_days is None:
        return None
    try:
        d = datetime.now(timezone.utc).date() - timedelta(days=int(age_days))
        return d.isoformat()
    except (ValueError, TypeError):
        return None


def _loc(val):
    """Locations can be a string or a list (e.g. fantastic-jobs locations_derived)."""
    if isinstance(val, list):
        return ", ".join(str(v) for v in val if v)
    return val


def _scalar(v):
    """SQLite can only bind str/num/None — flatten dicts/lists to readable text."""
    if v is None or isinstance(v, (str, int, float)):
        return v
    if isinstance(v, list):
        return ", ".join(str(x) for x in v if x not in (None, ""))
    if isinstance(v, dict):
        for k in ("text", "name", "label", "value", "display"):
            if v.get(k):
                return str(v[k])
        return ", ".join(f"{k}: {val}" for k, val in v.items() if val not in (None, "", [], {}))
    return str(v)


def _normalize_generic(item, source):
    company = _first(item, "company.name", "company", "companyName",
                     "company_name", "employer", "organization", "companyInfo.name")
    raw = {
        "source": source,
        "title": _first(item, "title", "jobTitle", "position", "name"),
        "company": company,
        "location": _loc(_first(item, "locations_derived", "location.formattedShort",
                                "location.formatted", "formattedLocation",
                                "location", "jobLocation", "city", "place")),
        "salary": _salary(item),
        "posted_date": _first(item, "date_posted", "dates.posted", "postedDate",
                             "posted_at", "publishedAt", "datePosted",
                             "postingDate", "listedAt", "date")
                       or _age_to_date(item.get("ageInDays")),
        "url": _first(item, "url", "urls.indeed", "apply.url", "jobUrl", "link",
                     "applyUrl", "jobPostingUrl", "detailsUrl"),
        "description": _first(item, "description", "descriptionText",
                             "jobDescription", "snippet", "description_text"),
    }
    return {k: _scalar(v) for k, v in raw.items()}


# Map shared "posted within N days" to each actor's own recency field.
def _days_to_fromdays(max_age):
    """Indeed/Glassdoor style: nearest allowed bucket of 1/3/7/14."""
    if not max_age:
        return None
    n = int(max_age)
    for bucket in (1, 3, 7, 14):
        if n <= bucket:
            return str(bucket)
    return "14"


def _days_to_timerange(max_age):
    """fantastic-jobs LinkedIn timeRange: 1h / 24h / 7d / 6m."""
    if not max_age:
        return "6m"
    n = int(max_age)
    if n <= 1:
        return "24h"
    if n <= 7:
        return "7d"
    return "7d"


# ---------------------------------------------------------------------------
# Actor registry. `build_input(limit, params)` returns the actor input dict.
# `params` carries keywords / location / careerUrls from the dashboard.
# ---------------------------------------------------------------------------
def _keyword_list(p):
    return p.get("keywordList") or ([p["keywords"]] if p.get("keywords") else [])


def _keyword_or(p, cap=6):
    """Boolean-OR query string for actors that take a single keyword field
    (Indeed, Glassdoor). Capped to `cap` titles — these actors return ZERO results
    when the OR query gets too long (e.g. 44 titles -> ~1,100 chars). The full title
    list is still used for LinkedIn / Career Sites, which take a titleSearch array."""
    terms = _keyword_list(p)[:cap]
    return " OR ".join(terms) if terms else ""


def _linkedin_input(limit, p):
    body = {
        "timeRange": _days_to_timerange(p.get("maxAge")),
        "limit": max(int(limit), 10),  # actor minimum is 10
        "descriptionType": "text",
        "removeAgency": False,
    }
    titles = _keyword_list(p)
    if titles:
        body["titleSearch"] = titles  # array = OR matching across all titles
    if p.get("location"):
        body["locationSearch"] = [p["location"]]
    return body


ACTORS = {
    "linkedin": {
        # Real-time LinkedIn jobs database — freshest source (down to last 24h).
        "label": "LinkedIn",
        "actor_id": "fantastic-jobs~advanced-linkedin-job-search-api",
        "build_input": _linkedin_input,
        "normalize": lambda item: _normalize_generic(item, "LinkedIn"),
    },
    "indeed": {
        # Cheapest Indeed actor with native sort-by-date + posted-within filter.
        "label": "Indeed",
        "actor_id": "kaix~indeed-scraper",
        "build_input": lambda limit, p: {
            k: v for k, v in {
                "keyword": _keyword_or(p),
                "location": p.get("location", ""),
                "maxItems": int(limit),
                "sort": "date",  # newest first
                "fromDays": _days_to_fromdays(p.get("maxAge")),
            }.items() if v not in (None, "")
        },
        "normalize": lambda item: _normalize_generic(item, "Indeed"),
    },
    "glassdoor": {
        "label": "Glassdoor",
        "actor_id": "valig~glassdoor-jobs-scraper",
        "build_input": lambda limit, p: {
            k: v for k, v in {
                "keywords": _keyword_or(p) or "Data Scientist",
                "location": p.get("location", "Remote"),
                "limit": int(limit),
                "sortBy": "date_desc",  # newest first
                "daysOld": int(p["maxAge"]) if p.get("maxAge") else None,
            }.items() if v is not None
        },
        "normalize": lambda item: _normalize_generic(item, "Glassdoor"),
    },
    "careersites": {
        # Searchable ATS database (Greenhouse, Lever, Workday, Ashby, ...) — no URLs needed.
        "label": "Career Sites",
        "actor_id": "fantastic-jobs~career-site-job-listing-api",
        "build_input": _linkedin_input,  # same title/location/timeRange schema
        "normalize": lambda item: _normalize_generic(item, "Career Sites"),
    },
}


def get_tokens():
    """All Apify tokens to rotate through. Supports (merged):
      APIFY_TOKENS=tok1,tok2,tok3   (comma-separated — easiest for many accounts)
      APIFY_TOKEN=tok               (single — backward compatible)
      APIFY_TOKEN_1, APIFY_TOKEN_2, ...  (numbered)
    Either variable may hold multiple tokens separated by commas, spaces, semicolons,
    or newlines — they are all split apart. Placeholders/duplicates dropped, order kept."""
    import re
    raw = []
    for var in ("APIFY_TOKENS", "APIFY_TOKEN"):
        raw += re.split(r"[\s,;]+", os.environ.get(var) or "")
    i = 1
    while os.environ.get(f"APIFY_TOKEN_{i}"):
        raw += re.split(r"[\s,;]+", os.environ[f"APIFY_TOKEN_{i}"])
        i += 1
    seen, out = set(), []
    for t in (x.strip() for x in raw):
        if t and not t.startswith("apify_api_xxxx") and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def get_token():
    """First token (backward-compatible helper)."""
    tokens = get_tokens()
    if not tokens:
        raise RuntimeError("No Apify token set. Add APIFY_TOKENS (comma-separated) or APIFY_TOKEN to .env.")
    return tokens[0]


def _token_exhausted(status, text):
    """True when the failure means this token is invalid or out of credit/quota, so
    we should try the next token (vs a real actor/input error that all tokens share)."""
    if status in (401, 402, 403, 429):
        return True
    t = (text or "").lower()
    return any(k in t for k in (
        "usage limit", "monthly usage", "insufficient", "exceeded", "payment required",
        "not enough", "credit", "quota", "user-or-token-not-found", "token is not valid"))


def run_actor(source_key, limit, params):
    """Run one actor synchronously and return (normalized_jobs, info_string).
    Rotates across all configured Apify tokens, falling back to the next whenever a
    token is exhausted/invalid, and remembering the working one for later calls."""
    global _token_start
    cfg = ACTORS[source_key]
    actor_id = cfg["actor_id"]
    body = cfg["build_input"](limit, params)
    url = f"{APIFY_BASE}/acts/{actor_id}/run-sync-get-dataset-items"

    tokens = get_tokens()
    if not tokens:
        return [], f"{cfg['label']}: no Apify token set (APIFY_TOKENS / APIFY_TOKEN)."

    n = len(tokens)
    start = _token_start % n
    last_err = ""
    for offset in range(n):
        idx = (start + offset) % n
        try:
            resp = requests.post(url, params={"token": tokens[idx]}, json=body, timeout=TIMEOUT)
        except requests.exceptions.RequestException as e:
            last_err = f"request error — {e}"
            continue
        if resp.status_code >= 400:
            last_err = f"HTTP {resp.status_code} — {resp.text[:160]}"
            if _token_exhausted(resp.status_code, resp.text):
                with _token_lock:          # this token is dead; skip it next time
                    _token_start = idx + 1
                continue
            return [], f"{cfg['label']}: {last_err}"  # real actor/input error
        try:
            items = resp.json()
        except ValueError:
            return [], f"{cfg['label']}: invalid JSON response"
        if not isinstance(items, list):
            return [], f"{cfg['label']}: unexpected response shape"
        if offset > 0:                     # advanced past an exhausted token; stick here
            with _token_lock:
                _token_start = idx
        jobs = [cfg["normalize"](it)
                for it in items[: int(limit) if int(limit) > 0 else None] if isinstance(it, dict)]
        tag = f" (token #{idx + 1}/{n})" if n > 1 else ""
        return jobs, f"{cfg['label']}: {len(jobs)} jobs{tag}"

    return [], f"{cfg['label']}: all {n} Apify token(s) exhausted/failed — last: {last_err}"
