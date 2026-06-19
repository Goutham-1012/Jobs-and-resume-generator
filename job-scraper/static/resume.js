const $ = (id) => document.getElementById(id);
let queue = [];        // [{id, jobDescription, company, title, status, data, savedPath, error}]
let processing = false;
let nextId = 1;

// Auto-load the user's base resume from the repo so they don't have to paste it.
async function loadBaseResume() {
  try {
    const r = await fetch("/api/base-resume");
    const { resume } = await r.json();
    if (resume && !$("resume").value.trim()) $("resume").value = resume;
  } catch (_) { /* ignore */ }
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

// ---- Queue management ----
function addToQueue() {
  const jd = $("jobDescription").value.trim();
  if (!jd) { setStatus("Add a job description before queueing."); return; }
  queue.push({
    id: nextId++,
    jobDescription: jd,
    company: $("company").value.trim(),
    title: $("title").value.trim(),
    status: "queued",
  });
  // clear the per-item fields, keep the resume
  $("jobDescription").value = "";
  $("company").value = "";
  $("title").value = "";
  $("jobPicker").value = "";
  renderQueue();
  setStatus("Added to queue.");
}

function label(item) {
  const parts = [item.company, item.title].filter(Boolean).join(" · ");
  return parts || `Untitled #${item.id}`;
}

const ICON = { queued: "⏳", generating: "⚙️", done: "✅", error: "❌" };

function renderQueue() {
  $("queuePanel").style.display = queue.length ? "block" : "none";
  $("queueCount").textContent = queue.length;
  $("queueList").innerHTML = queue.map((item) => `
    <li class="qitem">
      <span class="qstatus">${ICON[item.status] || ""}</span>
      <span class="qlabel">${esc(label(item))}</span>
      <span class="qnote">${esc(item.status === "error" ? item.error : (item.status === "done" ? `ATS ${item.ats}% · ${item.savedPath.split(/[\\\\/]/).pop()}` : item.status))}</span>
      <span class="qactions">
        ${item.status === "done" ? `<button class="mini" data-dl="${item.id}">Download</button> <button class="mini" data-view="${item.id}">Preview</button>` : ""}
        ${item.status === "queued" ? `<button class="mini" data-rm="${item.id}">Remove</button>` : ""}
      </span>
    </li>`).join("");
}

async function generateOne(item) {
  item.status = "generating"; renderQueue();
  try {
    const r = await fetch("/api/resume", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        resume: $("resume").value.trim(),
        jobDescription: item.jobDescription,
        company: item.company,
        title: item.title,
      }),
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data.error || ("HTTP " + r.status));
    item.data = data.data;
    item.preview = data.preview;
    item.savedPath = data.savedPath;
    item.ats = data.ats;
    item.status = "done";
  } catch (e) {
    item.status = "error";
    item.error = e.message;
  }
  renderQueue();
}

async function generateAll() {
  if (processing) return;
  if (!$("resume").value.trim()) { setStatus("Paste your resume first."); return; }
  const pending = queue.filter((i) => i.status === "queued" || i.status === "error");
  if (!pending.length) { setStatus("Nothing queued to generate."); return; }
  processing = true;
  $("genAllBtn").disabled = true;
  $("addBtn").disabled = true;
  let done = 0;
  for (const item of pending) {
    setStatus(`Generating ${++done}/${pending.length}: ${label(item)}… (keep this tab open)`);
    await generateOne(item);
  }
  processing = false;
  $("genAllBtn").disabled = false;
  $("addBtn").disabled = false;
  const ok = queue.filter((i) => i.status === "done").length;
  setStatus(`Finished. ${ok} resume(s) generated and auto-saved to generated_resumes/.`);
}

// ---- Download / preview ----
async function downloadDocx(data) {
  setStatus("Building .docx…");
  const r = await fetch("/api/resume/download", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ data }),
  });
  if (!r.ok) { setStatus("Download failed."); return; }
  const blob = await r.blob();
  const filename = (data.name || "resume").replace(/ /g, "_") + "_tailored.docx";
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [{
          description: "Word Document",
          accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] },
        }],
      });
      const w = await handle.createWritable();
      await w.write(blob); await w.close();
      setStatus("Saved.");
      return;
    } catch (e) {
      if (e.name === "AbortError") { setStatus("Save cancelled."); return; }
    }
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setStatus("Downloaded (check your browser's download folder).");
}

function showPreview(item) {
  $("output").textContent = item.preview || "";
  $("outputPanel").style.display = "block";
  $("outputPanel").scrollIntoView({ behavior: "smooth" });
}

// ---- Events ----
$("addBtn").addEventListener("click", addToQueue);
$("genAllBtn").addEventListener("click", generateAll);
$("clearBtn").addEventListener("click", () => {
  queue = queue.filter((i) => i.status === "generating");
  renderQueue();
  setStatus("Queue cleared.");
});
$("queueList").addEventListener("click", (e) => {
  const t = e.target;
  if (t.dataset.rm) { queue = queue.filter((i) => i.id != t.dataset.rm); renderQueue(); }
  else if (t.dataset.dl) { const it = queue.find((i) => i.id == t.dataset.dl); if (it) downloadDocx(it.data); }
  else if (t.dataset.view) { const it = queue.find((i) => i.id == t.dataset.view); if (it) showPreview(it); }
});
$("copyBtn").addEventListener("click", () => {
  navigator.clipboard.writeText($("output").textContent);
  setStatus("Copied to clipboard.");
});
$("dlBtn").addEventListener("click", () => {
  const lastDone = [...queue].reverse().find((i) => i.status === "done");
  if (lastDone) downloadDocx(lastDone.data);
  else setStatus("Generate a resume first.");
});

loadBaseResume();
loadJobPicker();
