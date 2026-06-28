const $ = (id) => document.getElementById(id);
function setStatus(t) { $("status").textContent = t; }
function esc(s) {
  return (s == null ? "" : String(s)).replace(/[&<>"]/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

async function loadCandidates() {
  const r = await fetch("/api/outreach/candidates");
  const { candidates } = await r.json();
  $("candEmpty").style.display = candidates.length ? "none" : "block";
  $("candList").innerHTML = candidates.map((c) => `
    <li class="qitem">
      <span class="qlabel">${esc(c.company || "?")} · ${esc(c.title || "")}</span>
      <span class="qnote">${c.has_outreach ? "drafted" : "résumé: " + esc(c.saved_path || "")}</span>
      <span class="qactions">
        <button class="mini" data-draft="${c.id}">${c.has_outreach ? "Draft more" : "Find contacts & draft"}</button>
      </span>
    </li>`).join("");
}

const ICON = { draft: "✉️", sent: "✅", failed: "❌", skipped: "⏭️" };

async function loadOutreach() {
  const r = await fetch("/api/outreach");
  const { outreach } = await r.json();
  $("outreachPanel").style.display = outreach.length ? "block" : "none";
  $("draftCount").textContent = outreach.filter((o) => o.status === "draft").length;
  $("outreachList").innerHTML = outreach.map((o) => `
    <div class="ocard" data-id="${o.id}">
      <div class="ohead">
        <span>${ICON[o.status] || ""} <strong>${esc(o.contact_name)}</strong>
          <span class="opt">${esc(o.contact_title)} · ${esc(o.contact_email)}</span></span>
        <span class="counts">${esc(o.company)} · ${esc(o.job_title)}${o.status === "failed" ? ' · <span style="color:#ff6b6b">' + esc(o.error || "failed") + "</span>" : ""}</span>
      </div>
      ${o.status === "draft" ? `
        <input class="osubject" value="${esc(o.subject)}" />
        <textarea class="obody" rows="9">${esc(o.body)}</textarea>
        <div class="resume-actions">
          <button class="mini" data-send="${o.id}">Send (with résumé)</button>
          <button class="mini" data-skip="${o.id}">Delete</button>
          <span class="opt">résumé for this job is attached automatically</span>
        </div>` : `
        <div class="opt">Subject: ${esc(o.subject)}</div>`}
    </div>`).join("");
}

async function refresh() { await loadCandidates(); await loadOutreach(); }

$("candList").addEventListener("click", async (e) => {
  const b = e.target.closest("button[data-draft]");
  if (!b) return;
  b.disabled = true; setStatus("Finding contacts via Apollo and drafting…");
  const r = await fetch("/api/outreach/draft", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ resume_queue_id: Number(b.dataset.draft) }),
  });
  const d = await r.json();
  if (!r.ok) setStatus("Error: " + (d.error || r.status));
  else setStatus(d.drafted ? `Drafted ${d.drafted} email(s).` : ("No drafts: " + (d.note || "no contacts")));
  refresh();
});

$("outreachList").addEventListener("click", async (e) => {
  const t = e.target;
  if (t.dataset.send) {
    const card = t.closest(".ocard");
    // persist any edits first
    await fetch("/api/outreach/" + t.dataset.send, {
      method: "PUT", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ subject: card.querySelector(".osubject").value,
                             body: card.querySelector(".obody").value }),
    });
    t.disabled = true; setStatus("Sending…");
    const r = await fetch("/api/outreach/" + t.dataset.send + "/send", { method: "POST" });
    const d = await r.json();
    setStatus(r.ok ? "Sent ✅" : "Send failed: " + (d.error || r.status));
    refresh();
  } else if (t.dataset.skip) {
    await fetch("/api/outreach/" + t.dataset.skip, { method: "DELETE" });
    refresh();
  }
});

$("sendAllBtn").addEventListener("click", async () => {
  if (!confirm("Send all draft emails now (each with its résumé attached)?")) return;
  setStatus("Sending all drafts…");
  const r = await fetch("/api/outreach/send-approved", { method: "POST" });
  const d = await r.json();
  setStatus(`Sent ${d.sent} email(s).`);
  refresh();
});
$("refreshBtn").addEventListener("click", refresh);

refresh();
