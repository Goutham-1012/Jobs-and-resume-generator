const $ = (id) => document.getElementById(id);
let lastData = null;

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
  }
});

async function generate() {
  const btn = $("genBtn");
  const resume = $("resume").value.trim();
  const jd = $("jobDescription").value.trim();
  if (!resume) { setStatus("Paste your resume first."); return; }
  if (!jd) { setStatus("Add a target job description first."); return; }

  btn.disabled = true;
  setStatus("Generating with OpenAI… this can take 20-40 seconds.");
  try {
    const r = await fetch("/api/resume", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ resume, jobDescription: jd }),
    });
    const data = await r.json();
    if (!r.ok) { setStatus("Error: " + (data.error || r.status)); return; }
    lastData = data.data;
    $("output").textContent = data.preview;
    $("outputPanel").style.display = "block";
    setStatus("Done. Download the .docx to get your exact format.");
    $("outputPanel").scrollIntoView({ behavior: "smooth" });
  } catch (e) {
    setStatus("Request failed: " + e.message);
  } finally {
    btn.disabled = false;
  }
}

function setStatus(t) { $("status").textContent = t; }

$("genBtn").addEventListener("click", generate);
$("copyBtn").addEventListener("click", () => {
  navigator.clipboard.writeText($("output").textContent);
  setStatus("Copied to clipboard.");
});
$("dlBtn").addEventListener("click", async () => {
  if (!lastData) { setStatus("Generate a resume first."); return; }
  setStatus("Building .docx…");
  const r = await fetch("/api/resume/download", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ data: lastData }),
  });
  if (!r.ok) { setStatus("Download failed."); return; }
  const blob = await r.blob();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = (lastData.name || "resume").replace(/ /g, "_") + "_tailored.docx";
  a.click();
  setStatus("Downloaded.");
});

loadBaseResume();
loadJobPicker();
