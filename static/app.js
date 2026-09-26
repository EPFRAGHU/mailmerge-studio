/* MailMerge Studio — front-end */
const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);

let COLUMNS = [], ROWS = [], checked = new Set(), selectedProvider = null;
let pollTimer = null;

function toast(msg, bad = false) {
  const t = $("#toast");
  t.textContent = msg;
  t.style.background = bad ? "#d6455b" : "#1d1d2b";
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 3000);
}

/* ---------- navigation ---------- */
const STEPS = ["connect", "contacts", "compose", "attachments", "send"];
let step = 0;
function goTo(i) {
  step = Math.max(0, Math.min(STEPS.length - 1, i));
  $$(".nav-item").forEach((b, k) => b.classList.toggle("active", k === step));
  $$(".step").forEach((s) => s.classList.remove("active"));
  $("#step-" + STEPS[step]).classList.add("active");
  if (STEPS[step] === "attachments") loadAttachments();
  if (STEPS[step] === "compose") renderMergeTags();
}
$$(".nav-item").forEach((b) => b.addEventListener("click", () => goTo(+b.dataset.step === undefined ? STEPS.indexOf(b.dataset.step) : STEPS.indexOf(b.dataset.step))));
$$(".nav-item").forEach((b) => b.addEventListener("click", () => goTo(STEPS.indexOf(b.dataset.step))));
$("#prevStep").onclick = () => goTo(step - 1);
$("#nextStep").onclick = () => goTo(step + 1);

/* ---------- connect ---------- */
const PRESET_INFO = {
  gmail: { label: "Gmail", icon: "📧", hint: "Gmail requires a 16-character <b>App Password</b>: Google Account → Security → 2-Step Verification → App passwords. Regular passwords won't work." },
  outlook: { label: "Outlook / M365", icon: "📨", hint: "Use your Microsoft account. If 2FA is on, create an app password at account.live.com/proofs." },
  yahoo: { label: "Yahoo", icon: "🟪", hint: "Generate an app password in Yahoo Account Security." },
  zoho: { label: "Zoho", icon: "🟩", hint: "Use an application-specific password from Zoho Mail settings." },
  icloud: { label: "iCloud", icon: "🍎", hint: "Create an app-specific password at appleid.apple.com." },
  custom: { label: "Custom SMTP", icon: "⚙️", hint: "Enter any SMTP server your provider gives you." },
};
function renderProviders() {
  $("#providerGrid").innerHTML = Object.entries(PRESET_INFO)
    .map(([k, v]) => `<div class="provider" data-p="${k}">${v.icon}<br>${v.label}</div>`).join("");
  $$(".provider").forEach((el) => el.onclick = () => {
    $$(".provider").forEach((p) => p.classList.remove("active"));
    el.classList.add("active");
    selectedProvider = el.dataset.p;
    fetch("/api/presets").then(r => r.json()).then(p => {
      const pre = p[selectedProvider];
      $("#smtpHost").value = pre.host; $("#smtpPort").value = pre.port; $("#smtpSecurity").value = pre.security;
      $("#providerHint").innerHTML = PRESET_INFO[selectedProvider].hint;
    });
  });
}
renderProviders();

$("#saveSmtp").onclick = async () => {
  const body = {
    smtp: {
      host: $("#smtpHost").value, port: +$("#smtpPort").value, security: $("#smtpSecurity").value,
      user: $("#smtpUser").value, replyTo: $("#smtpReplyTo").value, fromName: $("#smtpFromName").value,
    },
  };
  const pass = $("#smtpPass").value;
  await fetch("/api/config", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  if (pass && pass !== "********") await fetch("/api/config/save-password", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ password: pass }) });
  toast("Connection saved ✓");
};

$("#sendTest").onclick = async () => {
  if (!$("#smtpHost").value) return toast("Save your connection first", true);
  const to = prompt("Send a test email to:", $("#smtpUser").value);
  if (!to) return;
  await fetch("/api/send", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ subject: "Test from MailMerge Studio", html: "<h2>It works! 🎉</h2><p>Your SMTP settings are correct.</p>", testMode: false, testRecipient: to, delaySeconds: 0.5, batchSize: 1, batchPauseSeconds: 30 }),
  });
  toast("Test sending — check Send & Track for the result");
};

async function loadConfig() {
  const cfg = await (await fetch("/api/config")).json();
  const s = cfg.smtp || {};
  $("#smtpHost").value = s.host || ""; $("#smtpPort").value = s.port || 587;
  $("#smtpSecurity").value = s.security || "starttls"; $("#smtpUser").value = s.user || "";
  $("#smtpReplyTo").value = s.replyTo || ""; $("#smtpFromName").value = s.fromName || "";
}
loadConfig();

/* ---------- contacts ---------- */
const dz = $("#dropzone");
dz.onclick = (e) => { if (e.target.tagName !== "INPUT" && e.target.tagName !== "LABEL") $("#fileInput").click(); };
["dragover", "dragenter"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("over"); }));
["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("over"); }));
dz.addEventListener("drop", (e) => { if (e.dataTransfer.files[0]) uploadFile(e.dataTransfer.files[0]); });
$("#fileInput").onchange = (e) => { if (e.target.files[0]) uploadFile(e.target.files[0]); };

async function uploadFile(f) {
  const fd = new FormData(); fd.append("file", f);
  const r = await (await fetch("/api/import", { method: "POST", body: fd })).json();
  if (r.error) return toast(r.error, true);
  COLUMNS = r.columns; ROWS = r.rows; checked = new Set();
  $("#dataCard").style.display = "";
  $("#dataSummary").textContent = `${r.source} — ${r.total} contacts, ${r.columns.length} columns`;
  $("#importInfo").textContent = "";
  renderTable();
  renderMergeTags();
  fillAttachColumn();
  toast(`Loaded ${r.total} contacts ✓`);
}

function renderTable() {
  const head = `<tr><th><input type="checkbox" id="checkAll" checked></th>` + COLUMNS.map((c) => `<th>${esc(c)}</th>`).join("") + "</tr>";
  const body = ROWS.map((r, i) => `<tr><td><input type="checkbox" class="rowChk" data-i="${i}" checked></td>` +
    COLUMNS.map((c) => `<td>${esc(String(r[c] ?? ""))}</td>`).join("") + "</tr>").join("");
  $("#dataTable").innerHTML = head + body;
  $("#checkAll").onchange = (e) => { $$(".rowChk").forEach((c) => (c.checked = e.target.checked)); syncChecked(); };
  $$(".rowChk").forEach((c) => (c.onchange = syncChecked));
}
function syncChecked() {
  checked = new Set([...$$(".rowChk")].filter((c) => c.checked).map((c) => +c.dataset.i));
}
$("#clearData").onclick = () => { $("#dataCard").style.display = "none"; ROWS = []; };

function esc(s) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;"); }

/* ---------- compose ---------- */
$("#toolbar [data-cmd]").forEach((b) => b.onmousedown = (e) => { e.preventDefault(); document.execCommand(b.dataset.cmd); });
$("#fontSize").onchange = () => { document.execCommand("fontSize", false, $("#fontSize").value); $("#fontSize").value = ""; };
$("#textColor").oninput = () => document.execCommand("foreColor", false, $("#textColor").value);
$("#addLink").onclick = () => { const u = prompt("Link URL:", "https://"); if (u) document.execCommand("createLink", false, u); };
$("#addImage").onclick = () => { const u = prompt("Image URL:", "https://"); if (u) document.execCommand("insertImage", false, u); };
$("#addSignature").onclick = () => {
  editorHTML += "";
  document.execCommand("insertHTML", false, `<br><p style="color:#555">— <b>Your Name</b><br>Your Title · Company<br>📞 +91 00000 00000</p>`);
};

function renderMergeTags() {
  $("#mergeTags").innerHTML = COLUMNS.slice(0, 20).map((c) => `<button data-t="${esc(c)}">{{${esc(c)}}}</button>`).join("");
  $$("#mergeTags button").forEach((b) => b.onclick = () => document.execCommand("insertText", false, `{{${b.dataset.t}}}`));
}

$("#btnPreview").onclick = () => {
  if (!ROWS.length) return toast("Import contacts first", true);
  const row = ROWS[0];
  const html = mergeHTML(editorHTML(), row);
  $("#previewSubject").innerHTML = "<b>Subject:</b> " + esc(merge($("#subject").value, row));
  $("#previewFrame").srcdoc = html;
  $("#previewCard").style.display = "";
};

function editorHTML() { return $("#editor").innerHTML; }
function merge(text, row) {
  return (text || "").replace(/\{\{\s*([^}]+?)\s*\}\}/g, (m, k) => row[k] ?? "");
}
function mergeHTML(html, row) { return merge(html, row); }

$("#saveDraft").onclick = () => {
  localStorage.setItem("mm_draft", JSON.stringify({ subject: $("#subject").value, html: editorHTML() }));
  $("#draftInfo").textContent = "Draft saved in this browser ✓";
  toast("Draft saved ✓");
};
try {
  const d = JSON.parse(localStorage.getItem("mm_draft") || "null");
  if (d) { $("#subject").value = d.subject; $("#editor").innerHTML = d.html; }
} catch (e) {}

/* ---------- attachments ---------- */
$("#attachInput").onchange = async (e) => {
  const fd = new FormData();
  [...e.target.files].forEach((f) => fd.append("files", f));
  await fetch("/api/attachments", { method: "POST", body: fd });
  e.target.value = ""; loadAttachments(); toast("Files uploaded ✓");
};
async function loadAttachments() {
  const r = await (await fetch("/api/attachments")).json();
  $("#attachList").innerHTML = r.files.map((f) => `<span class="attach-item">📄 ${esc(f)} <button data-f="${esc(f)}">×</button></span>`).join("");
  $$("#attachList button").forEach((b) => b.onclick = async () => {
    await fetch("/api/attachments/delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: b.dataset.f }) });
    loadAttachments();
  });
}
function fillAttachColumn() {
  $("#attachColumn").innerHTML = `<option value="">— none —</option>` + COLUMNS.map((c) => `<option>${esc(c)}</option>`).join("");
}

/* ---------- send & track ---------- */
$("#btnLaunch").onclick = async () => {
  const mode = document.querySelector('input[name="mode"]:checked').value;
  const payload = {
    subject: $("#subject").value,
    html: editorHTML(),
    attachments: (await (await fetch("/api/attachments")).json()).files,
    attachmentColumn: $("#attachColumn").value,
    delaySeconds: +$("#delaySeconds").value,
    batchSize: +$("#batchSize").value,
    batchPauseSeconds: +$("#batchPauseSeconds").value,
    testMode: $("#testMode").checked,
    testRecipient: $("#testRecipient").value,
    selectedRows: mode === "selected" ? [...checked] : null,
  };
  if (mode === "test" && !payload.testRecipient) return toast("Enter a test recipient", true);
  if (mode !== "test" && !ROWS.length) return toast("Import contacts first", true);
  const r = await (await fetch("/api/send", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })).json();
  if (r.error) return toast(r.error, true);
  $("#btnStop").disabled = false;
  toast(mode === "test" ? "Test send started" : "Campaign started 🚀");
  startPolling();
};

$("#btnStop").onclick = async () => { await fetch("/api/stop", { method: "POST" }); toast("Stopping…"); };

function startPolling() {
  clearInterval(pollTimer);
  pollTimer = setInterval(updateProgress, 1200);
  updateProgress();
}
async function updateProgress() {
  const p = await (await fetch("/api/progress")).json();
  $("#statSent").textContent = `${p.sent} sent`;
  $("#statFailed").textContent = `${p.failed} failed`;
  $("#statTotal").textContent = `${p.total} total`;
  const pct = p.total ? Math.round(((p.sent + p.failed) / p.total) * 100) : 0;
  $("#progressBar").style.width = pct + "%";
  const pill = $("#statusPill");
  pill.textContent = p.status === "running" ? "Sending…" : p.status === "stopped" ? "Stopped" : p.status === "done" ? "Finished" : "Idle";
  pill.classList.toggle("live", p.status === "running");
  $("#logTable").innerHTML =
    `<tr><th>Email</th><th>Status</th><th>Detail</th></tr>` +
    [...(p.log || [])].reverse().slice(0, 100).map((e) =>
      `<tr><td>${esc(e.to || "")}</td><td style="color:${e.status === "sent" ? "#1f9d6b" : "#d6455b"};font-weight:700">${e.status}</td><td class="muted">${esc(e.error || "")}</td></tr>`).join("");
  if (p.status !== "running") {
    $("#btnStop").disabled = true;
    if (p.status === "done") { clearInterval(pollTimer); toast("Campaign finished ✓"); }
  }
}
$("#exportLog").onclick = () => (location.href = "/api/log/export");
