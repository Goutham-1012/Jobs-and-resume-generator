const $ = (id) => document.getElementById(id);
let queue = [];        // [{id, profileId, jobDescription, company, title, status, data, ...}]
let queueTiming = {};  // {avg_seconds, queued, generating, done, eta_seconds}
let processing = false;
let nextId = 1;
let profiles = [];

function fmtDur(s) {
  if (s == null) return "";
  s = Math.round(s);
  if (s < 60) return s + "s";
  const m = Math.floor(s / 60), sec = s % 60;
  return sec ? `${m}m ${sec}s` : `${m}m`;
}

// ---- Profiles ----
const PF_FIELDS = ["name", "email", "phone", "linkedin", "github", "portfolio",
                   "employers", "summary", "skills", "experience", "education",
                   "projects", "certifications"];
const PF_REQUIRED = ["name", "email", "phone", "linkedin", "github", "portfolio",
                     "employers", "summary", "skills", "experience", "education"];

async function loadProfiles(selectId) {
  const r = await fetch("/api/profiles");
  profiles = (await r.json()).profiles || [];
  const sel = $("profileSelect");
  sel.innerHTML = profiles.map((p) => `<option value="${p.id}">${esc(p.name)}</option>`).join("");
  if (selectId) sel.value = selectId;
}

function activeProfile() {
  return profiles.find((p) => p.id === $("profileSelect").value);
}

function openEditor(profile) {
  $("editorTitle").textContent = profile ? "Edit profile" : "New profile";
  $("profileEditor").dataset.editId = profile ? profile.id : "";
  PF_FIELDS.forEach((f) => {
    let v = profile ? profile[f] : "";
    if (f === "employers" && Array.isArray(v)) v = v.join(", ");
    $("pf_" + f).value = v || "";
  });
  $("editorStatus").textContent = "";
  $("profileEditor").style.display = "block";
  $("profileEditor").scrollIntoView({ behavior: "smooth" });
}

async function saveProfile() {
  const body = {};
  PF_FIELDS.forEach((f) => (body[f] = $("pf_" + f).value.trim()));
  const editId = $("profileEditor").dataset.editId;
  if (editId) body.id = editId;
  const missing = PF_REQUIRED.filter((f) => !body[f]);
  if (missing.length) { $("editorStatus").textContent = "Required: " + missing.join(", "); return; }
  const r = await fetch("/api/profiles", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  });
  const data = await r.json();
  if (!r.ok) { $("editorStatus").textContent = data.error || "Save failed."; return; }
  await loadProfiles(data.profile.id);
  $("profileEditor").style.display = "none";
  setStatus("Profile saved.");
}

async function deleteActiveProfile() {
  const p = activeProfile();
  if (!p) return;
  if (!confirm(`Delete profile "${p.name}"?`)) return;
  const r = await fetch("/api/profiles/" + p.id, { method: "DELETE" });
  const data = await r.json();
  if (!r.ok) { setStatus(data.error || "Delete failed."); return; }
  await loadProfiles();
  setStatus("Profile deleted.");
}

// Populate the "load from scraped job" picker with jobs that have descriptions.
async function loadJobPicker() {
  try {
    const r = await fetch("/api/jobs?sort=posted_date&order=desc");
    const { jobs } = await r.json();
    const withDesc = jobs.filter((j) => (j.description || "").trim().length > 40);
    const sel = $("jobPicker");
    withDesc.slice(0, 100).forEach((j, i) => {
      const o = document.createElement("option");
      o.value = i;
      o.textContent = `[${j.source}] ${j.title} @ ${j.company || "?"}`;
      sel._jobs = sel._jobs || [];
      sel._jobs[i] = j;
      sel.appendChild(o);
    });
  } catch (_) { /* dashboard may be empty; ignore */ }
}

$("jobPicker").addEventListener("change", (e) => {
  const jobs = e.target._jobs || [];
  const j = jobs[e.target.value];
  if (j) {
    $("jobDescription").value =
      `${j.title}\nCompany: ${j.company || ""}\nLocation: ${j.location || ""}\n\n${j.description || ""}`;
    if (!$("company").value && j.company) $("company").value = j.company;
    if (!$("title").value && j.title) $("title").value = j.title;
  }
});

function setStatus(t) { $("status").textContent = t; }
function esc(s) {
  return (s == null ? "" : String(s)).replace(/[&<>"]/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

// ---- Server-backed queue (shared with the dashboard auto-enqueue) ----
// "Generate" adds an item to the server queue; a background worker processes it.
async function enqueue() {
  const profile = activeProfile();
  if (!profile) { setStatus("Select or create a profile first."); return; }
  const jd = $("jobDescription").value.trim();
  if (!jd) { setStatus("Add a job description first."); return; }
  const r = await fetch("/api/resume/enqueue", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      profileId: profile.id, jobDescription: jd,
      company: $("company").value.trim(), title: $("title").value.trim(),
      model: $("modelSelect").value,
    }),
  });
  const data = await r.json();
  if (!r.ok) { setStatus("Error: " + (data.error || r.status)); return; }
  $("jobDescription").value = ""; $("company").value = ""; $("title").value = ""; $("jobPicker").value = "";
  setStatus("Added to the queue.");
  refreshQueue();
}

function label(item) {
  const parts = [item.company, item.title].filter(Boolean).join(" · ") || item.label || "Untitled";
  return item.profile_name ? `${item.profile_name} — ${parts}` : parts;
}

const ICON = { queued: "⏳", generating: "⚙️", done: "✅", error: "❌", canceling: "🛑" };

function renderQueue() {
  $("queuePanel").style.display = queue.length ? "block" : "none";
  $("queueCount").textContent = queue.length;

  // Queue-level timing summary: average generation time + ETA for what's still waiting.
  const t = queueTiming || {};
  const bits = [];
  if (t.avg_seconds) bits.push(`avg ${fmtDur(t.avg_seconds)}/resume`);
  if (t.queued) bits.push(`${t.queued} waiting`);
  if (t.eta_seconds && (t.queued || t.generating)) bits.push(`~${fmtDur(t.eta_seconds)} left`);
  $("queueTiming").textContent = bits.join(" · ");

  $("queueList").innerHTML = queue.map((item) => {
    const m = item.model ? ` · ${item.model}` : "";
    const t2 = item.elapsed_seconds;
    const note = item.status === "error" ? (item.error || "error")
      : item.status === "done"
        ? `ATS ${item.ats}%${t2 != null ? ` · took ${fmtDur(t2)}` : ""}${item.saved_path ? ` · ${item.saved_path}` : ""}`
      : item.status === "generating"
        ? `generating${m}${t2 != null ? ` · ${fmtDur(t2)}…` : ""}`
      : item.status === "canceling"
        ? "stopping…"
        : item.status + m;
    const draggable = item.status === "queued";
    return `
    <li class="qitem${draggable ? " draggable" : ""}" data-id="${item.id}" ${draggable ? 'draggable="true"' : ""}>
      <span class="qhandle">${draggable ? "⠿" : ""}</span>
      <span class="qstatus">${ICON[item.status] || ""}</span>
      <span class="qlabel">${esc(label(item))}</span>
      <span class="qnote">${esc(note)}</span>
      <span class="qactions">
        ${item.status === "done" ? `<button class="mini" data-dl="${item.id}">Download</button> <button class="mini" data-view="${item.id}">Preview</button>` : ""}
        ${item.status === "error" ? `<button class="mini" data-retry="${item.id}">Regenerate</button> <button class="mini" data-rm="${item.id}">Remove</button>` : ""}
        ${item.status === "queued" ? `<button class="mini" data-rm="${item.id}">Remove</button>` : ""}
        ${item.status === "generating" ? `<button class="mini" data-stop="${item.id}">Stop</button>` : ""}
        ${item.status === "canceling" ? `<button class="mini" disabled>Stopping…</button>` : ""}
      </span>
    </li>`;
  }).join("");
}

async function refreshQueue() {
  if (isDragging) return;  // don't rebuild the list under the pointer mid-drag
  try {
    const r = await fetch("/api/resume/queue");
    const d = await r.json();
    queue = d.queue || [];
    queueTiming = d.timing || {};
    renderQueue();
  } catch (_) { /* ignore transient poll errors */ }
}

// ---- Drag to reorder (queued items only): true insert-between ----
let dragId = null;
let isDragging = false;

function clearDropMarks() {
  $("queueList").querySelectorAll(".drop-before, .drop-after")
    .forEach((el) => el.classList.remove("drop-before", "drop-after"));
}

// Among queued items (excluding the one being dragged), find the one whose midpoint
// is just below the cursor — the dragged item goes BEFORE it. null => append at end.
function dropAfterElement(y) {
  const els = [...$("queueList").querySelectorAll("li.draggable:not(.dragging)")];
  let best = { offset: -Infinity, el: null };
  for (const el of els) {
    const box = el.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > best.offset) best = { offset, el };
  }
  return best.el;
}

$("queueList").addEventListener("dragstart", (e) => {
  const li = e.target.closest("li.draggable[data-id]");
  if (!li) return;
  dragId = li.dataset.id; isDragging = true;
  li.classList.add("dragging");
});
$("queueList").addEventListener("dragend", () => {
  clearDropMarks();
  const d = $("queueList").querySelector(".dragging");
  if (d) d.classList.remove("dragging");
  dragId = null; isDragging = false;
  refreshQueue();
});
$("queueList").addEventListener("dragover", (e) => {
  if (!dragId) return;
  e.preventDefault();
  clearDropMarks();
  const after = dropAfterElement(e.clientY);
  if (after) after.classList.add("drop-before");
  else {
    const queuedEls = $("queueList").querySelectorAll("li.draggable:not(.dragging)");
    if (queuedEls.length) queuedEls[queuedEls.length - 1].classList.add("drop-after");
  }
});
$("queueList").addEventListener("drop", async (e) => {
  e.preventDefault();
  if (!dragId) return;
  const after = dropAfterElement(e.clientY);
  let ids = queue.filter((i) => i.status === "queued").map((i) => String(i.id))
    .filter((id) => id !== String(dragId));
  const idx = after ? ids.indexOf(after.dataset.id) : ids.length;
  ids.splice(idx < 0 ? ids.length : idx, 0, String(dragId));
  clearDropMarks();
  dragId = null; isDragging = false;
  await fetch("/api/resume/queue/reorder", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids }),
  });
  refreshQueue();
});

// ---- Download / preview ----
async function downloadQueueItem(id) {
  setStatus("Building .docx…");
  const r = await fetch(`/api/resume/queue/${id}/download`);
  if (!r.ok) { setStatus("Not ready to download yet."); return; }
  const blob = await r.blob();
  const filename = "resume_tailored.docx";
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [{ description: "Word Document",
          accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] } }],
      });
      const w = await handle.createWritable(); await w.write(blob); await w.close();
      setStatus("Saved."); return;
    } catch (e) { if (e.name === "AbortError") { setStatus("Save cancelled."); return; } }
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = filename; a.click();
  setStatus("Downloaded (check your browser's download folder).");
}

function showPreview(item) {
  $("output").textContent = item.preview || "";
  $("previewAts").textContent = item.ats != null ? `ATS ${item.ats}%` : "";
  $("outputPanel").style.display = "block";
  $("outputPanel").scrollIntoView({ behavior: "smooth" });
}

// ---- Events ----
$("genBtn").addEventListener("click", enqueue);
$("clearBtn").addEventListener("click", async () => {
  await fetch("/api/resume/queue/clear", { method: "POST" });
  setStatus("Queue cleared.");
  refreshQueue();
});
$("queueList").addEventListener("click", async (e) => {
  const t = e.target;
  if (t.dataset.retry) {
    await fetch("/api/resume/queue/" + t.dataset.retry + "/retry", { method: "POST" });
    setStatus("Re-queued for regeneration.");
    refreshQueue();
  } else if (t.dataset.rm) {
    await fetch("/api/resume/queue/" + t.dataset.rm, { method: "DELETE" });
    refreshQueue();
  } else if (t.dataset.stop) {
    t.disabled = true;
    await fetch("/api/resume/queue/" + t.dataset.stop + "/cancel", { method: "POST" });
    setStatus("Stopping…");
    refreshQueue();
  } else if (t.dataset.dl) {
    downloadQueueItem(t.dataset.dl);
  } else if (t.dataset.view) {
    const it = queue.find((i) => String(i.id) === t.dataset.view); if (it) showPreview(it);
  }
});
$("copyBtn").addEventListener("click", () => {
  navigator.clipboard.writeText($("output").textContent);
  setStatus("Copied to clipboard.");
});
$("dlBtn").addEventListener("click", () => {
  const lastDone = [...queue].reverse().find((i) => i.status === "done");
  if (lastDone) downloadQueueItem(lastDone.id);
  else setStatus("No completed resume yet.");
});

// Profile editor events
$("newProfileBtn").addEventListener("click", () => openEditor(null));
$("editProfileBtn").addEventListener("click", () => { const p = activeProfile(); if (p) openEditor(p); });
$("delProfileBtn").addEventListener("click", deleteActiveProfile);
$("saveProfileBtn").addEventListener("click", saveProfile);
$("cancelProfileBtn").addEventListener("click", () => { $("profileEditor").style.display = "none"; });

loadProfiles();
loadJobPicker();
refreshQueue();
setInterval(refreshQueue, 3000);  // poll the shared server queue
