# Job Scraper Dashboard

Scrapes jobs from **LinkedIn, Indeed, Glassdoor, and company career sites** using
[Apify](https://apify.com) actors, stores them in SQLite, and shows them in a
sortable/filterable web dashboard. Every source is scraped with the **same shared
"jobs per source" limit**.

## Actors used (recency-optimized)

Actors were chosen to surface **the most recently posted jobs**. Each one sorts by
date and/or honors the shared "Posted within" window.

| Source       | Apify actor                                       | Limit field        | Recency |
|--------------|---------------------------------------------------|--------------------|---------|
| LinkedIn     | `fantastic-jobs/advanced-linkedin-job-search-api` | `limit` (min 10)   | real-time DB, `timeRange` 24h/7d |
| Indeed       | `kaix/indeed-scraper`                             | `maxItems`         | `sort=date` + `fromDays` |
| Glassdoor    | `valig/glassdoor-jobs-scraper`                    | `limit`            | `sortBy=date_desc` + `daysOld` |
| Career Sites | `santamaria-automations/career-site-jobs-scraper` | `maxJobsPerCompany`| no date filter (returns current openings) |

The shared **Posted within** dropdown (default: last 24h) is mapped to each actor's
own recency field. Swap any actor in [`apify_client.py`](apify_client.py) — the
registry maps both the shared limit and the recency window to each actor's input.

> **LinkedIn note:** this actor is a real-time jobs *database* (not a live page
> crawl). Search is by `titleSearch` / `locationSearch`. It returns the freshest
> postings but costs more (~$5/1k on the free tier).

> These are paid (pay-per-result) actors. Each run costs a small amount on your
> Apify account. Keep the limit low while testing.

## Setup (Windows / PowerShell)

```powershell
cd "job-scraper"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env      # then edit .env and paste your APIFY_TOKEN
python app.py
```

Open http://127.0.0.1:5000

## Usage

1. Enter keywords, location, and the **jobs-per-source** number.
2. Tick the sources you want. For **Career Sites**, paste career-page URLs
   (Lever, Greenhouse, Workday, Ashby, SmartRecruiters, etc.), one per line.
3. Click **Scrape Jobs**. Actors run in parallel; results are deduped and saved.
4. Sort by clicking column headers; filter by text, source, or run.

## Notes

- Get a free token at https://console.apify.com/account/integrations
- Data persists in `jobs.db`. Delete the file to reset.
- Duplicate jobs (same title+company+location+url) are ignored on insert.

cd "job-scraper"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env      # paste your APIFY_TOKEN
python app.py