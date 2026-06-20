let state = { sort: "posted_date", order: "desc", run_id: "", latestRunId: null };
let allJobs = [];

const $ = (id) => document.getElementById(id);

async function scrape() {
  const btn = $("scrapeBtn");
  const sources = [...document.querySelectorAll(".source:checked")].map((c) => c.value);
  if (!sources.length) { setStatus("Select at least one source."); return; }

  const payload = {
    keywords: $("keywords").value,
    location: $("location").value,
    limit: parseInt($("limit").value, 10) || 25,
    maxAge: $("maxAge").value,
    sources,
  };

  btn.disabled = true;
  setStatus("Scraping… running " + sources.length + " actor(s), this can take a minute.");
  try {
    const r = await fetch("/api/scrape", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await r.json();
    if (!r.ok) { setStatus("Error: " + (data.error || r.status)); return; }
    setStatus(`Done. ${data.inserted} new jobs saved.\n` + (data.log || []).join("\n"));
    await loadRuns();
    await loadJobs();
  } catch (e) {
    setStatus("Request failed: " + e.message);
  } finally {
    btn.disabled = false;
  }
}

function setStatus(t) { $("status").textContent = t; }

async function loadRuns() {
  const r = await fetch("/api/runs");
  const { runs } = await r.json();
  state.latestRunId = runs.length ? runs[0].id : null;  // runs are ordered newest-first
  const sel = $("runFilter");
  const current = sel.value;
  sel.innerHTML = '<option value="">All runs</option>' +
    runs.map((run) => `<option value="${run.id}">#${run.id} · ${(run.keywords || "—")} · ${run.total_jobs} jobs</option>`).join("");
  sel.value = current;
}

async function loadJobs() {
  const params = new URLSearchParams();
  if (state.run_id) params.set("run_id", state.run_id);
  const src = $("sourceFilter").value;
  if (src) params.set("source", src);
  const search = $("search").value.trim();
  if (search) params.set("search", search);
  params.set("sort", state.sort);
  params.set("order", state.order);

  const r = await fetch("/api/jobs?" + params.toString());
  const { jobs, stats } = await r.json();
  allJobs = jobs;
  renderJobs(allJobs);
  renderCounts(allJobs.length, stats);
}

function renderCounts(total, stats) {
  const parts = Object.entries(stats || {}).map(([s, c]) => `${s}: ${c}`);
  $("counts").textContent = `${total} shown` + (parts.length ? ` · ${parts.join(" · ")}` : "");
}

function esc(s) {
  return (s == null ? "" : String(s)).replace(/[&<>"]/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

function rowHtml(j) {
  const unseen = !j.seen;
  const link = j.url
    ? `<a href="${esc(j.url)}" target="_blank" rel="noopener" data-job-id="${j.id}">Open ↗</a>` +
      (unseen ? ` <span class="badge unseen-badge">Unseen</span>` : "")
    : "";
  return `
    <tr class="${unseen ? "unseen" : ""}">
      <td>${unseen ? '<span class="unseen-dot" title="Unseen"></span>' : ""}${esc(j.title)}</td>
      <td>${esc(j.company)}</td>
      <td>${esc(j.location)}</td>
      <td>${esc(j.salary)}</td>
      <td><span class="badge">${esc(j.source)}</span></td>
      <td>${esc((j.posted_date || "").slice(0, 10))}</td>
      <td>${link}</td>
    </tr>`;
}

function groupHeader(text, n, isNew) {
  return `<tr class="group-header${isNew ? " group-new" : ""}"><td colspan="7">${esc(text)} · ${n}</td></tr>`;
}

function renderJobs(jobs) {
  const body = $("jobsBody");
  $("empty").style.display = jobs.length ? "none" : "block";
  const isNew = (j) => state.latestRunId != null && j.run_id === state.latestRunId;
  const fresh = jobs.filter(isNew);
  const older = jobs.filter((j) => !isNew(j));
  let html = "";
  if (fresh.length)
    html += groupHeader(`Newly scraped · run #${state.latestRunId}`, fresh.length, true) +
            fresh.map(rowHtml).join("");
  if (older.length)
    html += groupHeader("Previously scraped", older.length, false) + older.map(rowHtml).join("");
  body.innerHTML = html;
}

async function markSeen(id) {
  const j = allJobs.find((x) => x.id == id);
  if (j && !j.seen) { j.seen = 1; renderJobs(allJobs); }
  try { await fetch(`/api/jobs/${id}/seen`, { method: "POST" }); } catch (_) { /* ignore */ }
}

document.querySelectorAll("th[data-sort]").forEach((th) => {
  th.addEventListener("click", () => {
    const col = th.dataset.sort;
    if (state.sort === col) state.order = state.order === "asc" ? "desc" : "asc";
    else { state.sort = col; state.order = "asc"; }
    document.querySelectorAll("th[data-sort]").forEach((h) =>
      (h.textContent = h.textContent.replace(/[ ▾▴]/g, "")));
    th.textContent += state.order === "asc" ? " ▴" : " ▾";
    loadJobs();
  });
});

$("jobsBody").addEventListener("click", (e) => {
  const a = e.target.closest("a[data-job-id]");
  if (a) markSeen(a.dataset.jobId);  // link still opens in a new tab
});
$("scrapeBtn").addEventListener("click", scrape);
$("search").addEventListener("input", () => { clearTimeout(window._t); window._t = setTimeout(loadJobs, 250); });
$("sourceFilter").addEventListener("change", loadJobs);
$("runFilter").addEventListener("change", (e) => { state.run_id = e.target.value; loadJobs(); });

loadRuns();
loadJobs();
