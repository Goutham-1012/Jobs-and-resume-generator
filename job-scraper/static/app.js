let state = { sort: "posted_date", order: "desc", run_id: "" };

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
  renderJobs(jobs);
  renderCounts(jobs.length, stats);
}

function renderCounts(total, stats) {
  const parts = Object.entries(stats || {}).map(([s, c]) => `${s}: ${c}`);
  $("counts").textContent = `${total} shown` + (parts.length ? ` · ${parts.join(" · ")}` : "");
}

function esc(s) {
  return (s == null ? "" : String(s)).replace(/[&<>"]/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

function renderJobs(jobs) {
  const body = $("jobsBody");
  $("empty").style.display = jobs.length ? "none" : "block";
  body.innerHTML = jobs.map((j) => `
    <tr>
      <td>${esc(j.title)}</td>
      <td>${esc(j.company)}</td>
      <td>${esc(j.location)}</td>
      <td>${esc(j.salary)}</td>
      <td><span class="badge">${esc(j.source)}</span></td>
      <td>${esc((j.posted_date || "").slice(0, 10))}</td>
      <td>${j.url ? `<a href="${esc(j.url)}" target="_blank" rel="noopener">Open ↗</a>` : ""}</td>
    </tr>`).join("");
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

$("scrapeBtn").addEventListener("click", scrape);
$("search").addEventListener("input", () => { clearTimeout(window._t); window._t = setTimeout(loadJobs, 250); });
$("sourceFilter").addEventListener("change", loadJobs);
$("runFilter").addEventListener("change", (e) => { state.run_id = e.target.value; loadJobs(); });

loadRuns();
loadJobs();
