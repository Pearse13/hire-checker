(function () {
  const E = window.HireEngine;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const state = { acs: null, suppliers: [], active: 0, errors: [], cross: null, queue: [], pos: 0, decisions: new Map(), ran: false, step: 1, second: [], spos: 0, again: false, jobs: [], parked: new Set() };
  let downloads = null;

  // ---------- formatting ----------
  const fmtDate = (k) => k ? `${k.slice(8, 10)}/${k.slice(5, 7)}/${k.slice(2, 4)}` : "";
  const fmtRate = (v) => typeof v === "number" ? (v ? "£" + v.toFixed(2) : "–") : (v ? esc(v) : "–");
  const plural = (n, one, many) => `${n} ${n === 1 ? one : (many || one + "s")}`;

  // ---------- reading files ----------
  async function readBook(file) {
    if (/\.xls$/i.test(file.name)) throw new Error(`${file.name} is an old .xls file. Open it in Excel, Save As .xlsx, then add it again.`);
    let buf = await file.arrayBuffer();
    try { buf = await E.cleanXlsx(buf, JSZip); } catch (e) { /* not a zip: ExcelJS reports it below */ }
    const wb = new ExcelJS.Workbook();
    try { await wb.xlsx.load(buf.slice(0)); }
    catch (e) { throw new Error(`${file.name} could not be read as an Excel workbook. Open it in Excel, Save As .xlsx, then add it again.`); }
    const sheets = wb.worksheets.map((ws) => {
      const rows = [];
      for (let r = 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r), out = [];
        for (let c = 1; c <= ws.columnCount; c++) out.push(E.plain(row.getCell(c).value));
        rows.push(out);
      }
      return { title: ws.name, rows };
    });
    return { buf, book: { name: file.name, sheets } };
  }

  // New files mean a new check: nothing shows until Run check is pressed
  const ready = () => state.ran && state.acs && state.suppliers.length;
  function notRun() { state.ran = false; state.cross = null; state.queue = []; }
  function runCheck() { rerun(); state.ran = true; state.focusReview = true; render(); }

  function rerun() {
    state.acsSave = null;
    for (const s of state.suppliers) s.saveState = null, s.result = E.checkSupplier(s.book, state.acs, s.override || null);
    crossCheck();
  }
  // Runs after every supplier report is checked: suggestions for unknown plant codes, then the £0 end pass.
  // Step 1 is the pairings, step 2 the second look, step 3 finished.
  function crossCheck() {
    state.cross = state.acs && state.suppliers.length ? E.crossCheck(state.acs, state.suppliers.map((s) => s.result)) : null;
    state.queue = state.cross ? E.pairings(state.acs, state.suppliers.map((s) => s.result), state.cross) : [];
    // One job at a time, in the ACS report's sheet order
    state.jobs = state.cross ? [...new Set([...state.acs.items.map((a) => a.sheet), ...state.queue.map(mainJob)])] : [];
    state.queue.sort((a, b) => state.jobs.indexOf(mainJob(a)) - state.jobs.indexOf(mainJob(b)));
    state.parked = new Set(); state.lineCache = new Map();
    state.decisions = loadDecisions();
    state.pos = nextWaiting(0);
    state.step = 1; state.second = []; state.spos = 0; state.again = false;
    if (state.cross && state.queue.every(decisionOf)) enterSecond();
  }

  // ---------- approving pairings ----------
  // Decisions are kept in this browser per ACS report, so a reload doesn't lose the clicking done so far.
  const storeKey = () => "hirecheck:" + state.acs.name;
  function loadDecisions() {
    try { return new Map(Object.entries(JSON.parse(localStorage.getItem(storeKey()) || "{}"))); } catch (e) { return new Map(); }
  }
  function saveDecisions() {
    try { localStorage.setItem(storeKey(), JSON.stringify(Object.fromEntries(state.decisions))); } catch (e) { /* not kept */ }
  }
  const decisionOf = (p) => state.decisions.get(p.id) || null;
  const approved = () => state.queue.filter((p) => decisionOf(p) === "approved");
  // First pairing at or after i still waiting for an answer, or the end of the queue
  function nextWaiting(i) {
    for (let j = i; j < state.queue.length; j++) if (!decisionOf(state.queue[j]) && !state.parked.has(state.queue[j].id)) return j;
    return state.queue.length;
  }
  function changed() { saveDecisions(); state.acsSave = null; for (const s of state.suppliers) s.saveState = null; }
  // An answer covers everything that's the same thing: the same code charged again, and the other assets on an ACS line listing several
  function setAnswer(p, v) {
    const cells = new Set(p.items.filter((a) => a.codes.length > 1).map((a) => a.cell));
    state.decisions.set(p.id, v);
    for (const q of state.queue) {
      if (decisionOf(q) || stillCharged(q)) continue;
      if ((q.line.supplier === p.line.supplier && q.line.code === p.line.code) || q.items.some((a) => cells.has(a.cell))) state.decisions.set(q.id, v);
    }
  }
  function decide(v) {
    const p = state.queue[state.pos];
    if (!p) return;
    setAnswer(p, v);
    changed();
    state.pos = state.again ? state.pos + 1 : nextWaiting(state.pos + 1);
    if (state.pos >= state.queue.length) state.pos = Math.min(nextWaiting(0), state.queue.length);
    if (state.pos >= state.queue.length) enterSecond(true);
    show();
  }
  // For the job on the card: approves every waiting pairing where the code and cost agree.
  // Anything else stays waiting and comes up in the second look. Then on to the next job.
  function checkAll() {
    const job = mainJob(state.queue[state.pos]);
    for (const p of state.queue) if (mainJob(p) === job && !decisionOf(p)) {
      const c = checks(p, true);
      if (c.rows[0].ok && !c.bad) state.decisions.set(p.id, "approved");
      else state.parked.add(p.id);
    }
    changed();
    state.pos = nextWaiting(0);
    if (state.pos >= state.queue.length) enterSecond(true);
    show();
  }
  function move(by) {
    state.pos = Math.max(0, Math.min(state.queue.length, state.pos + by));
    if (state.pos >= state.queue.length && nextWaiting(0) >= state.queue.length) enterSecond(true);
    show();
  }
  function show() {
    state.focusReview = true;
    render();
    // Everything checked: download both marked-up copies straight away
    if (state.autoSave) { state.autoSave = false; saveAll(); }
  }

  // ---------- second look: one at a time, once the pairings are done ----------
  // Pairings still waiting for an answer, then everything nothing could be paired with. An answered pairing never comes back.
  const lineKey = (f) => `${f.supplier}|${f.sheet}|${f.row}`;
  function buildSecond() {
    const out = [];
    const cells = new Set(state.queue.flatMap((p) => p.items.map((a) => a.cell)));
    const lines = new Set(state.queue.flatMap((p) => p.lines));
    const asks = state.queue.filter((p) => !decisionOf(p) || state.decisions.has("seen:p:" + p.id));
    for (const p of [...asks.filter(stillCharged), ...asks.filter((p) => !stillCharged(p))]) out.push({ id: "seen:p:" + p.id, type: "pairing", p });
    for (const s of state.suppliers) for (const f of s.result.flagged) if (f.charged && !lines.has(f)) out.push({ id: "seen:c:" + lineKey(f), type: "charged", f, company: s.result.company });
    for (const s of state.suppliers) for (const a of s.result.notCharged) if (!cells.has(a.cell)) { cells.add(a.cell); out.push({ id: "seen:n:" + a.cell, type: "notCharged", a }); }
    for (const z of state.cross.zeroPass) if (!z.line && !cells.has(z.acs.cell)) out.push({ id: "seen:z:" + z.acs.cell, type: "zero", a: z.acs });
    return out;
  }
  function nextUnseen(i) {
    for (let j = i; j < state.second.length; j++) {
      const it = state.second[j];
      if (!state.decisions.get(it.id) && !(it.p && decisionOf(it.p))) return j;
    }
    return state.second.length;
  }
  function enterSecond(click) {
    state.second = buildSecond();
    state.step = 2; state.again = false;
    state.spos = nextUnseen(0);
    if (state.spos >= state.second.length) { state.step = 3; state.autoSave = !!click; }
  }
  // v: approve or reject a pairing. outcome: what happened to it, kept with the check and left off the spreadsheets.
  function answer(v, outcome) {
    const it = state.second[state.spos];
    if (!it) return;
    if (it.type === "pairing" && v) setAnswer(it.p, v);
    if (outcome) state.decisions.set("outcome:" + it.id, outcome);
    state.otherOpen = false;
    state.decisions.set(it.id, "seen");
    changed();
    state.spos = nextUnseen(state.spos + 1);
    if (state.spos >= state.second.length) { state.step = 3; state.autoSave = true; }
    show();
  }
  function secondBack() {
    state.otherOpen = false;
    if (state.spos > 0) state.spos--;
    else { state.step = 1; state.pos = Math.max(0, state.queue.length - 1); state.again = true; }
    show();
  }
  function goAgain() {
    for (const k of [...state.decisions.keys()]) if (k.startsWith("seen:")) state.decisions.delete(k);
    saveDecisions();
    state.step = 1; state.pos = 0; state.again = true;
    show();
  }

  async function addACS(file) {
    state.errors = [];
    try {
      const { book, buf } = await readBook(file);
      state.acsSave = null;
      state.acs = { name: file.name, buf, sheetNames: book.sheets.map((s) => s.title), reportDate: reportDate(file.name), ...E.readACS(book) };
      notRun();
    } catch (e) { state.errors.push(e.message); }
    render();
  }

  async function addSuppliers(files) {
    state.errors = [];
    for (const file of files) {
      if (/_checked\.xlsx$/i.test(file.name)) { state.errors.push(`${file.name} is a checked copy, so it was skipped. Add the original supplier report instead.`); continue; }
      try {
        const { buf, book } = await readBook(file);
        const existing = state.suppliers.findIndex((s) => s.name === file.name);
        const entry = { name: file.name, book, buf, override: "", result: null };
        if (existing >= 0) state.suppliers[existing] = entry; else state.suppliers.push(entry);
        state.active = existing >= 0 ? existing : state.suppliers.length - 1;
      } catch (e) { state.errors.push(e.message); }
    }
    notRun();
    render();
  }

  // ---------- download ----------
  async function buildChecked(s) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(s.buf.slice(0));
    E.markWorkbook(wb, s.result, supplierMarks(s), lineOutcomes().filter((o) => o.supplier === s.name));
    E.markMissing(wb, supplierMissing(s));
    const out = await wb.xlsx.writeBuffer();
    return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }
  const checkedName = (s) => s.name.replace(/\.xlsx?m?$/i, "") + "_checked.xlsx";

  // Only approved pairings are marked: the supplier line gets the ACS cell(s) it was paired with
  const supplierMarks = (s) => approved().flatMap((p) => p.lines.filter((f) => f.supplier === s.name).map((f) => ({ sheet: f.sheet, row: f.row, acsAt: f.acsAt || p.acsAt })));
  const acsLines = () => E.acsApprovedLines(state.acs, approved());
  // Second-look outcomes, for the Outcome column on each copy
  const withOutcome = () => state.second.map((it) => ({ it, text: state.decisions.get("outcome:" + it.id) })).filter((o) => o.text);
  const lineOutcomes = () => withOutcome().flatMap(({ it, text }) => (it.p ? it.p.lines : it.f ? [it.f] : []).map((f) => ({ supplier: f.supplier, sheet: f.sheet, row: f.row, text })));
  const itemOutcomes = () => withOutcome().flatMap(({ it, text }) => (it.p ? it.p.items : it.a ? [it.a] : []).map((item) => ({ item, text })));
  const costText = (v) => `£${(v || 0).toFixed(2)}`;
  // Charged lines not on the ACS report, unless an approved pairing covers them
  function supplierMissing(s) {
    const paired = new Set(approved().flatMap((p) => p.lines));
    return s.result.flagged.filter((f) => f.charged && !paired.has(f)).map((f) => ({ sheet: f.sheet, row: f.row, width: s.result.sheets.find((x) => x.title === f.sheet).width,
      text: `Row ${f.row}: ${f.raw} ${f.desc}, ${costText(E.money(f.rate))}: not on the ACS report.` }))
      .concat(state.queue.filter((p) => stillCharged(p) && decisionOf(p) !== "approved").flatMap((p) => p.lines).filter((f) => f.supplier === s.name && f.charged)
        .map((f) => ({ sheet: f.sheet, row: f.row, width: s.result.sheets.find((x) => x.title === f.sheet).width, text: `Row ${f.row}: ${f.raw} ${f.desc}, ${costText(E.money(f.rate))}: off-hire on ACS, still being charged.` })));
  }
  // ACS lines not on their supplier's report, and £0 lines not found on any report, unless they've been marked
  function acsMissing() {
    const marked = new Set(acsLines().map((a) => a.cell)), lines = new Map();
    const add = (a, codes, why) => { if (!marked.has(a.cell) && !lines.has(a.cell)) lines.set(a.cell, { sheet: a.sheet, row: a.row, width: a.width, text: `Row ${a.row}: ${codes} ${a.desc}, ${costText(a.rate)}: ${why}.` }); };
    for (const s of state.suppliers) for (const a of s.result.notCharged) add(a, a.missing.length === a.codes.length ? a.assetText : a.missing.join(", "), `not on the ${a.company} report`);
    for (const p of state.queue) if (stillCharged(p) && decisionOf(p) !== "approved") for (const a of p.items) add(a, a.assetText, "off-hire, but still being charged");
    for (const z of state.cross.zeroPass) if (!z.line) add(z.acs, z.acs.assetText, "not found on a supplier report");
    return [...lines.values()];
  }
  async function buildACS() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(state.acs.buf.slice(0));
    E.markACS(wb, acsLines(), itemOutcomes());
    E.markMissing(wb, acsMissing());
    return new Blob([await wb.xlsx.writeBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  // Saves one marked-up copy: i = supplier index, or "acs". Resolves "saved", "declined" or "failed".
  async function saveChecked(i) {
    const isAcs = i === "acs";
    const s = isAcs ? null : state.suppliers[i];
    const name = isAcs ? checkedName(state.acs) : checkedName(s);
    const set = (v) => { if (isAcs) state.acsSave = v; else s.saveState = v; };
    if (!downloads || !(isAcs ? state.acs.buf : s.buf)) return "failed";
    set("saving"); renderMarked();
    let result;
    try {
      await downloads.save({ filename: name, data: isAcs ? await buildACS() : await buildChecked(s) });
      result = "saved";
    } catch (e) {
      result = e && e.code === "declined" ? "declined" : "failed";
      if (result === "failed") {
        state.errors = [e && e.code === "rate_limited" ? "A save is already waiting for your answer. Finish that one first." : `${name} could not be saved here. Run hire_check.py on your computer to make it instead.`];
        renderNotices();
      }
    }
    set(result); renderMarked();
    return result;
  }

  async function saveAll() {
    if (state.acsSave !== "saved" && await saveChecked("acs") === "failed") return;
    for (let i = 0; i < state.suppliers.length; i++) {
      if (state.suppliers[i].saveState === "saved") continue;
      if (await saveChecked(i) === "failed") break;
    }
  }

  // ---------- Excel Online ----------
  // From the Microsoft app registration. The Excel Online buttons stay hidden until clientId is filled in.
  const MS = { clientId: "", tenantId: "consumers" };
  const MSAL_JS = "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.30.0/lib/msal-browser.min.js";
  const MS_SCOPES = ["Files.ReadWrite"];
  const excelSaving = new Set();
  let ms = null, msAccount = null;

  async function startMicrosoft() {
    if (!MS.clientId) return;
    if (!window.msal) await new Promise((ok, no) => document.head.appendChild(Object.assign(document.createElement("script"), { src: MSAL_JS, onload: ok, onerror: no })));
    const app = new msal.PublicClientApplication({
      auth: { clientId: MS.clientId, authority: `https://login.microsoftonline.com/${MS.tenantId}`, redirectUri: location.origin + "/auth.html" },
      cache: { cacheLocation: "sessionStorage" },
    });
    await app.initialize();
    ms = app; msAccount = ms.getAllAccounts()[0] || null;
    renderMarked();
  }

  async function msSignIn() {
    try { msAccount = (await ms.loginPopup({ scopes: MS_SCOPES })).account; }
    catch (e) { state.errors = ["Microsoft sign-in didn't finish. Try again."]; renderNotices(); }
    renderMarked();
  }

  const oneDriveName = (t) => t.replace(/["*:<>?/\\|]/g, "-").trim();

  // Saves one marked-up copy to Hire Checks/<ACS report> in the viewer's OneDrive and opens it in Excel Online.
  // Pressing it again replaces that copy.
  async function openInExcel(i) {
    const s = i === "acs" ? null : state.suppliers[i];
    const name = checkedName(s || state.acs);
    // Opened here: browsers block a new tab once the upload has finished
    const tab = window.open("", "_blank");
    if (!tab) { state.errors = ["Your browser blocked the new tab. Allow pop-ups for this site, then press it again."]; renderNotices(); return; }
    tab.document.title = name;
    tab.document.body.textContent = "Saving to OneDrive…";
    excelSaving.add(i); renderMarked();
    try {
      const { accessToken } = await ms.acquireTokenSilent({ scopes: MS_SCOPES, account: msAccount }).catch((e) => { msAccount = null; throw e; });
      const data = s ? await buildChecked(s) : await buildACS();
      const folder = state.acs.name.replace(/\.xlsx?m?$/i, "");
      const path = ["Hire Checks", folder, name].map((p) => encodeURIComponent(oneDriveName(p))).join("/");
      const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${path}:/content`, { method: "PUT", headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": data.type }, body: data });
      if (!res.ok) throw { status: res.status };
      const { webUrl } = await res.json();
      tab.location.href = webUrl;
    } catch (e) {
      tab.close();
      state.errors = [!msAccount ? "Sign in to Microsoft again, then press it again." : e.status === 423 ? `${name} is open in Excel. Close it, then press it again.` : `${name} could not be saved to OneDrive.`];
      renderNotices();
    }
    excelSaving.delete(i); renderMarked();
  }

  // ---------- rendering ----------
  function renderFiles() {
    const acsBox = $("acsFiles"), supBox = $("supFiles");
    acsBox.innerHTML = state.acs ? `<div class="file">${esc(state.acs.name)} <small>${plural(state.acs.sheetNames.length, "sheet")}: ${esc(state.acs.sheetNames.join(", "))}</small></div>` : "";
    supBox.innerHTML = state.suppliers.map((s) => `<div class="file">${esc(s.name)}</div>`).join("");
    $("acsDrop").classList.toggle("filled", !!state.acs);
    $("supDrop").classList.toggle("filled", state.suppliers.length > 0);
  }

  function renderNotices() {
    const parts = [];
    if (!state.acs && !state.suppliers.length) parts.push(`<div class="notice info"><b>Add this week's reports</b> in steps 1 and 2.</div>`);
    else if (!state.acs) parts.push(`<div class="notice info"><b>Add the ACS on-hire report</b> in step 1.</div>`);
    else if (!state.suppliers.length) parts.push(`<div class="notice info"><b>Add supplier hire reports</b> in step 2.</div>`);
    for (const e of state.errors) parts.push(`<div class="notice error"><b>Not added.</b> ${esc(e)}</div>`);
    $("notices").innerHTML = parts.join("");
  }

  function counts(r) {
    const charged = r.flagged.filter((f) => f.charged);
    return { matched: r.matched.length, charged: charged.length, free: r.flagged.length - charged.length, notCharged: r.notCharged.length };
  }

  function renderTabs() {
    const tabs = $("tabs");
    if (!ready()) { tabs.innerHTML = ""; return; }
    tabs.innerHTML = state.suppliers.map((s, i) => {
      const c = counts(s.result);
      const status = c.charged ? `<span class="pill flag">${plural(c.charged, "charge")} not on ACS</span>` : `<span class="pill ok">All charges on ACS</span>`;
      return `<button class="sup-tab" role="tab" id="tab${i}" aria-selected="${i === state.active}" data-i="${i}"><span class="n">${esc(s.name)}</span><span class="s">${status} ${c.notCharged ? `<span class="pill warn">${c.notCharged} not charged</span>` : ""}</span></button>`;
    }).join("");
    tabs.querySelectorAll(".sup-tab").forEach((b) => b.addEventListener("click", () => { state.active = +b.dataset.i; render(); }));
  }

  function columnOptions(s) {
    const sh = s.book.sheets[0];
    const width = Math.max(0, ...sh.rows.map((r) => r.length));
    const auto = s.result.sheets[0];
    const head = (c) => { for (const row of sh.rows.slice(0, 40)) { const v = row[c]; if (typeof v === "string" && v.trim() && !E.looksLikeCode(v)) return v.trim(); } return ""; };
    let html = `<option value="">Auto${auto && !s.override ? ` (${auto.colLetter}${auto.header ? " · " + esc(auto.header) : ""})` : ""}</option>`;
    for (let c = 0; c < width; c++) {
      const L = E.colLetter(c), h = head(c);
      html += `<option value="${L}" ${s.override === L ? "selected" : ""}>${L}${h ? " · " + esc(h.slice(0, 28)) : ""}</option>`;
    }
    return html;
  }

  function renderResult() {
    const box = $("result");
    if (!ready()) { box.innerHTML = ""; return; }
    const s = state.suppliers[state.active];
    const r = s.result, c = counts(r);
    const free = r.flagged.filter((f) => !f.charged);
    const cols = r.sheets.map((x) => `${x.colLetter}${x.header ? " (" + esc(x.header) + ")" : ""}`).join(", ");

    const byJob = new Map();
    for (const m of [...r.matched, ...r.flagged.filter((f) => pairingOf(f))]) {
      const k = (m.jobs || pairingOf(m).items.map((a) => a.sheet)).join(", ");
      if (!byJob.has(k)) byJob.set(k, []); byJob.get(k).push(m);
    }
    const jobs = [...byJob.entries()].sort((a, b) => b[1].length - a[1].length);
    const matchedTable = r.matched.length ? `<div class="table-wrap"><table>
      <thead><tr><th>Row</th><th>Plant code</th><th>Plant</th><th>On hire</th><th class="num">Rate</th><th>ACS cell</th><th>Your check</th></tr></thead>
      <tbody>${jobs.map(([job, ms]) => `<tr class="group"><td colspan="7">${esc(job)} · ${ms.length}</td></tr>` + ms.map((m) => { const d = lineDecision(m); return `<tr class="${rowClass(m)}"><td class="num">${m.row}</td><td class="code">${esc(m.raw)}</td><td class="desc">${esc(m.desc)}</td><td class="num">${fmtDate(m.date)}</td><td class="num">${fmtRate(m.rate)}</td><td class="code">${(m.acsAt || pairingOf(m).acsAt).map(esc).join("<br>")}</td><td>${decisionPill(d)}</td></tr>`; }).join("")).join("")}</tbody></table></div>`
      : `<p class="empty">No lines on this report match the ACS report. Check the plant code column above.</p>`;

    const freeTable = free.length ? `<details><summary>${plural(free.length, "line")} not on the ACS report with no charge (accessories, notes)</summary><div class="table-wrap"><table>
      <thead><tr><th>Row</th><th>Plant code</th><th>Plant</th><th>On hire</th></tr></thead>
      <tbody>${free.map((f) => `<tr><td class="num">${f.row}</td><td class="code">${esc(f.raw)}</td><td class="desc">${esc(f.desc)}</td><td class="num">${fmtDate(f.date)}</td></tr>`).join("")}</tbody></table></div></details>` : "";

    box.innerHTML = `<div class="result" role="tabpanel" aria-labelledby="tab${state.active}">
      <div class="result-head">
        <div><h2>${esc(s.name)}</h2><div class="meta">${r.company ? `Listed as <b>${esc(r.company)}</b> on the ACS report · ` : ""}plant codes read from column ${cols || "none found"}</div></div>
        <div class="controls">
          <label for="codeCol">Plant code column<select id="codeCol">${columnOptions(s)}</select></label>
        </div>
      </div>
      <div class="totals">
        <div class="total ok"><b>${c.matched}</b><span>on both reports</span></div>
        <div class="total ${c.charged ? "flag" : "ok"}"><b>${c.charged}</b><span>charged, not on ACS</span></div>
        <div class="total ${c.notCharged ? "warn" : "ok"}"><b>${c.notCharged}</b><span>on ACS, not charged</span></div>
        <div class="total"><b>${c.free}</b><span>no-charge lines not on ACS</span></div>
      </div>
      <section class="block">
        <div class="block-head"><h3>On both reports</h3><p>Grouped by ACS job sheet.</p></div>
        ${matchedTable}
      </section>
      ${free.length ? `<section class="block">${freeTable}</section>` : ""}
    </div>`;

    $("codeCol").addEventListener("change", (e) => { s.override = e.target.value; s.saveState = null; state.acsSave = null; s.result = E.checkSupplier(s.book, state.acs, s.override || null); crossCheck(); render(); });
  }

  function renderMarked() {
    const box = $("marked");
    const real = ready();
    if (!real) { box.innerHTML = ""; return; }
    const label = { saving: "Saving…", saved: "Saved", declined: "Not saved", failed: "Not saved" };
    const pill = (st) => st ? `<span class="pill ${st === "saved" ? "ok" : st === "saving" ? "warn" : "flag"}">${label[st]}</span>` : "";
    const lines = acsLines(), acsTotal = state.acs.items.length;
    const waiting = state.queue.filter((p) => !decisionOf(p)).length;
    const excelBtn = (i) => msAccount ? `<button class="btn ghost" data-excel="${i}" ${excelSaving.has(i) ? "disabled" : ""}>${excelSaving.has(i) ? "Saving…" : "Open in Excel Online"}</button>` : "";
    const acsRow = `<li class="mk-row">
        <span class="mk-name"><span class="code">${esc(checkedName(state.acs))}</span><small>ACS report · ${plural(lines.length, "line")} marked</small></span>
        ${pill(state.acsSave)}
        <button class="btn ghost" data-save="acs" ${!downloads || state.acsSave === "saving" ? "disabled" : ""}>${state.acsSave === "saved" ? "Save again" : "Download"}</button>
        ${excelBtn("acs")}
      </li>`;
    const rows = acsRow + state.suppliers.map((s, i) => {
      const n = supplierMarks(s).length, total = s.result.matched.length + s.result.flagged.length;
      return `<li class="mk-row">
        <span class="mk-name"><span class="code">${esc(checkedName(s))}</span><small>${plural(n, "line")} marked</small></span>
        ${pill(s.saveState)}
        <button class="btn ghost" data-save="${i}" ${!downloads || s.saveState === "saving" ? "disabled" : ""}>${s.saveState === "saved" ? "Save again" : "Download"}</button>
        ${excelBtn(i)}
      </li>`;
    }).join("");
    const pending = state.acsSave === "saving" || state.suppliers.some((s) => s.saveState === "saving");
    box.innerHTML = `<section class="marked">
      <div class="mk-head">
        <div><span class="step">Step 4</span><h2>Marked-up spreadsheets</h2>
        ${waiting ? `<p><b>${plural(waiting, "pairing")} skipped.</b></p>` : ""}</div>
        <div class="mk-actions">
          ${ms && !msAccount ? `<button class="btn" id="msSignIn">Sign in to Microsoft</button>` : ""}
          <button class="btn" id="saveAll" ${!downloads || pending ? "disabled" : ""}>Download all (${state.suppliers.length + 1})</button>
        </div>
      </div>
      <ul class="mk-list">${rows}</ul>
      ${downloads ? "" : `<p class="empty">Downloads aren't available in this view. Open the page on claude.ai, or run hire_check.py on your computer to make the same copies.</p>`}
    </section>`;
    box.querySelectorAll("[data-save]").forEach((b) => b.addEventListener("click", () => saveChecked(b.dataset.save === "acs" ? "acs" : +b.dataset.save)));
    box.querySelectorAll("[data-excel]").forEach((b) => b.addEventListener("click", () => openInExcel(b.dataset.excel === "acs" ? "acs" : +b.dataset.excel)));
    const all = $("saveAll"); if (all) all.addEventListener("click", saveAll);
    const signIn = $("msSignIn"); if (signIn) signIn.addEventListener("click", msSignIn);
  }

  const afterReport = (f) => f.date && state.acs.reportDate && f.date > state.acs.reportDate;

  // ---------- step 3: one pairing at a time ----------
  const pairingOf = (m) => state.queue.find((q) => q.lines.includes(m));
  // The supplier lines on the card showing now
  function currentLines() {
    if (state.step === 1) { const p = state.queue[state.pos]; return p ? p.lines : []; }
    const it = state.second[state.spos];
    return !it || state.step !== 2 ? [] : it.p ? it.p.lines : it.f ? [it.f] : [];
  }
  const rowClass = (m) => [lineDecision(m) === "approved" ? "matched" : "", currentLines().includes(m) ? "current" : ""].join(" ");
  const lineDecision = (m) => { const p = pairingOf(m); return p ? decisionOf(p) : null; };
  const decisionPill = (d) => d === "approved" ? `<span class="pill ok">Approved</span>` : d === "rejected" ? `<span class="pill flag">Rejected</span>` : `<span class="pill warn">Waiting</span>`;
  const reviewing = () => !!state.cross && (state.step === 1 ? state.pos < state.queue.length : state.step === 2);
  const near = (a, b) => Math.abs(a - b) < 0.005;
  const shortName = (name) => name.replace(/\.xlsx?m?$/i, "");
  const missing = `<span class="pc-none">Not given</span>`;

  // Off-hire on the ACS report but still being charged: never goes through by itself
  const stillCharged = (p) => p.items.length > 0 && p.items.every((a) => a.offHire) && p.lines.some((f) => f.charged);
  const lineTotal = (f) => (E.money(f.rate) || 0) * (f.qty || 1);
  // Every supplier line charging for an ACS line: its codes, read past typing slips, and lines suggested for it
  function linesFor(a) {
    if (!state.lineCache.has(a.cell)) {
      const lines = state.suppliers.flatMap((s) => [...s.result.matched, ...s.result.flagged]).filter((f) => E.codeOnLine(f.code, a));
      lines.push(...state.queue.filter((q) => q.items.some((i) => i.cell === a.cell)).flatMap((q) => q.stand));
      state.lineCache.set(a.cell, [...new Set(lines)].filter((f) => lineTotal(f) > 0).slice(0, 14));
    }
    return state.lineCache.get(a.cell);
  }
  // The fewest of those lines that add up to the ACS cost exactly, and include must if given.
  // Rates are per item (rate × quantity), but some suppliers' rate is already the line's cost.
  function addsUp(a, ar, must) { return sumsTo(a, ar, must, lineTotal) || sumsTo(a, ar, must, (f) => E.money(f.rate) || 0); }
  function sumsTo(a, ar, must, value) {
    const lines = linesFor(a);
    let best = null;
    for (let m = 1; m < 1 << lines.length; m++) {
      const pick = lines.filter((_, i) => m & (1 << i));
      if ((!must || pick.includes(must)) && (!best || pick.length < best.length) && near(pick.reduce((t, f) => t + value(f), 0), ar)) best = pick;
    }
    return best;
  }
  // What every pairing shows: item, cost, company, on-hire date. Only the code and cost are checked: ok is true when they agree.
  // Descriptions are never compared, and a code a slip away from the ACS one counts as the same.
  // bad: the cost doesn't agree. Only the second look (hl) highlights it.
  function checks(p, hl) {
    const ls = p.lines;
    const sup = state.suppliers.find((s) => s.name === ls[0].supplier);
    const company = sup && sup.result.company;
    const rates = ls.map((f) => E.money(f.rate));
    const fr = rates.some((r) => r == null) ? null : rates.reduce((t, r) => t + r, 0);
    // The supplier lines this cost agrees through, or null. 2 items in one is 2 costs.
    const same = (i) => {
      const ar = i.zero ? 0 : i.rate, n = Math.max(i.qty || 1, i.codes.length);
      if (fr == null || ar == null) return null;
      if (near(fr, ar) || near(fr * n, ar) || near(fr * (ls[0].qty || 1), ar)) return [ls[0]];
      return addsUp(i, ar, fr > 0 ? ls[0] : null);
    };
    const a = p.items.find(same) || p.items[0] || {}, ar = a.zero ? 0 : a.rate;
    const hit = a.cell ? same(a) : null, others = (hit || []).filter((f) => f !== ls[0]);
    const parts = Math.max(a.qty || 1, (a.codes || []).length);
    const cost = (v, qty) => v == null ? missing : `£${v.toFixed(2)}${qty > 1 ? ` <small>× ${qty}</small>` : ""}`;
    const name = (n, code) => `${n ? esc(n) : missing}${code ? `<span class="pc-code">${esc(code)}</span>` : ""}`;
    const dates = [...new Set(ls.map((f) => f.date))];
    const rest = [
      { k: "Company", s: esc(company || shortName(ls[0].supplier)), a: a.company ? esc(a.company) : missing },
      { k: "On hire", s: dates.map((d) => d ? fmtDate(d) : missing).join("<br>"), a: a.date ? fmtDate(a.date) : missing },
    ];
    const rows = [
      { k: "Item", s: name(ls[0].desc, ls[0].raw), a: name(a.desc, a.assetText || p.acsAt.join(", ")),
        ok: p.kind === "match" || (a.codes || []).some((c) => E.nearCode(ls[0].code, c)) || E.codeOnLine(ls[0].code, a) },
      { k: "Cost", s: `${cost(fr, ls[0].qty)}${others.length ? `<span class="pc-code">${fr ? "With" : "Cost on"} row${others.length > 1 ? "s" : ""} ${others.map((f) => f.row).join(", ")}: £${ar.toFixed(2)}</span>` : ""}`, a: `${cost(ar, 1)}${parts > 1 ? ` <small>for ${parts}</small>` : ""}`, ok: !!hit },
      ...rest,
    ];
    return { rows, bad: !rows[1].ok || stillCharged(p) };
  }

  const jobOf = (p) => [...new Set(p.items.length ? p.items.map((a) => a.sheet) : p.acsAt.map((c) => c.split("!")[0]))].join(", ");
  const mainJob = (p) => p.items.length ? p.items[0].sheet : p.acsAt[0].split("!")[0];
  const bar = (at, n) => `<div class="rv-bar" aria-hidden="true"><span style="width:${n ? 100 * at / n : 0}%"></span></div>`;

  // The card for one pairing. hl: the second look, where what doesn't agree is highlighted
  function pairingCard(p, count, progress, hl, back, skip, extra) {
    const d = decisionOf(p), c = checks(p, hl);
    const rows = c.rows.map((r) => `<tr class="${!hl || r.ok === undefined ? "" : r.ok ? "pc-ok" : "pc-bad"}"><th>${r.k}</th><td>${r.s}</td><td>${r.a}</td><td class="pc-mark">${!hl || r.ok === undefined ? "" : r.ok ? "✓" : "✗"}</td></tr>`).join("");
    return `<section class="review">
      <div class="rv-top"><h2 class="rv-job">${esc(jobOf(p))}</h2><span class="rv-pos">${count}</span>${p.kind !== "match" ? `<span class="rv-kind">Different code</span>` : ""}${d ? decisionPill(d) : ""}${extra ? extra.pill : ""}</div>
      ${progress}
      ${hl && stillCharged(p) ? `<p class="rv-note">Off-hire on ACS, still being charged.</p>` : ""}
      <table class="pc"><thead><tr><th></th><th>Supplier</th><th>ACS</th><th></th></tr></thead><tbody>${rows}</tbody></table>
      <div class="rv-actions">
        <button class="btn big" id="rvYes">Approve <kbd>A</kbd></button>
        <button class="btn big" id="rvNo">Reject <kbd>R</kbd></button>
        <button class="btn" id="rvBack" ${back ? "" : "disabled"}>Back</button>
        ${skip ? `<button class="btn" id="rvSkip">Skip</button><button class="btn" id="rvAll">Check all ${esc(mainJob(p))}</button>` : ""}
      </div>${extra ? extra.buttons : ""}</section>`;
  }

  // What happened to a second-look item. Keys 1 to 5 press the buttons in order.
  const outcomesFor = (it) => [...(it.type === "charged" ? ["New plant"] : []), "Off-hired", "Moved site", "Loss charge applied", "Query with supplier"];
  function outcomeBits(it) {
    const was = state.decisions.get("outcome:" + it.id);
    return {
      pill: was ? `<span class="pill ok">${esc(was)}</span>` : "",
      buttons: `<div class="rv-outcome-head">What happened?</div>
        <div class="rv-outcomes">${outcomesFor(it).map((o, i) => `<button class="btn" data-out="${esc(o)}" aria-pressed="${was === o}"><span>${esc(o)}</span><kbd>${i + 1}</kbd></button>`).join("")}
        <button class="btn" id="rvOther" aria-pressed="${!!was && !outcomesFor(it).includes(was)}"><span>Other</span></button></div>
        ${state.otherOpen ? `<form class="rv-other" id="rvOtherForm"><input id="rvOtherText" aria-label="What happened" placeholder="What happened?" maxlength="120"><button class="btn">Save</button></form>` : ""}`,
    };
  }

  // Second-look card for something with no pairing: reported, not marked
  function leftoverCard(it, count, progress) {
    const f = it.f, a = it.a, bits = outcomeBits(it);
    const gbp = (v) => v == null ? missing : `£${v.toFixed(2)}`;
    const row = (k, v) => `<tr><th>${k}</th><td>${v}</td></tr>`;
    const notes = [];
    if (f && f.date && state.acs.reportDate && f.date > state.acs.reportDate) notes.push("On hire after the ACS report date: probably new plant.");
    if (it.type === "zero" && a.company && !state.suppliers.some((s) => (s.result.company || "").toLowerCase() === a.company.toLowerCase())) notes.push(`No ${a.company} report has been checked.`);
    const label = { charged: "Charged, not on ACS", notCharged: "On ACS, not charged", zero: "£0 on ACS, not on a supplier report" }[it.type];
    const codes = a && (a.missing ? a.missing.join(", ") : a.assetText);
    const rows = f
      ? row("Item", `${f.desc ? esc(f.desc) : missing}<span class="pc-code">${esc(f.raw)}</span>`) + row("Cost", gbp(E.money(f.rate))) + row("Company", esc(it.company || shortName(f.supplier))) + row("On hire", f.date ? fmtDate(f.date) : missing)
      : row("Item", `${a.desc ? esc(a.desc) : missing}<span class="pc-code">${esc(codes)}</span>`) + row("Cost", gbp(a.zero ? 0 : a.rate)) + row("Company", a.company ? esc(a.company) : missing) + row("On hire", a.date ? fmtDate(a.date) : missing);
    return `<section class="review">
      <div class="rv-top"><h2 class="rv-job">${esc(f ? shortName(f.supplier) : a.sheet)}<small>${label}</small></h2><span class="rv-pos">${count}</span>${bits.pill}</div>
      ${progress}
      <table class="pc pc-one"><tbody>${rows}</tbody></table>
      ${notes.map((n) => `<p class="rv-note">${esc(n)}</p>`).join("")}
      ${bits.buttons}
      <div class="rv-actions">
        <button class="btn" id="rvBack">Back</button>
        <button class="btn" id="rvSkip">Skip</button>
      </div></section>`;
  }

  function renderReview() {
    const box = $("review");
    if (!state.cross) { box.innerHTML = ""; return; }
    const q = state.queue, n = q.length;
    const done = q.filter((p) => decisionOf(p)).length;
    const wire = (id, fn) => { if ($(id)) $(id).addEventListener("click", fn); };

    if (state.step === 3) {
      const yes = approved().length, tally = new Map();
      for (const it of state.second) { const o = state.decisions.get("outcome:" + it.id); if (o) tally.set(o, (tally.get(o) || 0) + 1); }
      box.innerHTML = `<section class="review rv-done"><p><b>All ${n} checked.</b> ${yes} approved, ${n - yes} not.</p>
        ${tally.size ? `<p>${[...tally].map(([o, c]) => `${esc(o)}: ${c}`).join(" · ")}</p>` : ""}
        <div class="rv-actions"><button class="btn" id="rvAgain">Go through again</button></div></section>`;
      wire("rvAgain", goAgain);
      return;
    }
    if (state.step === 2) {
      const it = state.second[state.spos], m = state.second.length;
      const count = `<b>${state.spos + 1}</b> of ${m} second look`, progress = bar(state.spos, m);
      box.innerHTML = it.p ? pairingCard(it.p, count, progress, true, true, false, outcomeBits(it)) : leftoverCard(it, count, progress);
      wire("rvYes", () => answer("approved")); wire("rvNo", () => answer("rejected")); wire("rvSkip", () => answer());
      wire("rvBack", secondBack);
      box.querySelectorAll("[data-out]").forEach((b) => b.addEventListener("click", () => answer(null, b.dataset.out)));
      wire("rvOther", () => { state.otherOpen = true; render(); $("rvOtherText").focus(); });
      if ($("rvOtherForm")) $("rvOtherForm").addEventListener("submit", (e) => { e.preventDefault(); const t = $("rvOtherText").value.trim(); if (t) answer(null, t); });
      return;
    }
    if (state.pos >= n) {
      const first = nextWaiting(0);
      box.innerHTML = `<section class="review rv-done"><p><b>${plural(n - done, "pairing")} skipped.</b></p>
        <div class="rv-actions"><button class="btn" id="rvFirst">Finish the skipped ones</button><button class="btn" id="rvOn">Carry on</button></div></section>`;
      wire("rvFirst", () => { state.pos = first; show(); });
      wire("rvOn", () => { enterSecond(true); show(); });
      return;
    }
    const job = mainJob(q[state.pos]), inJob = q.filter((p) => mainJob(p) === job);
    box.innerHTML = pairingCard(q[state.pos], `<b>${inJob.indexOf(q[state.pos]) + 1}</b> of ${inJob.length} · job ${state.jobs.indexOf(job) + 1} of ${state.jobs.length}`, bar(done, n), false, state.pos > 0, true);
    wire("rvYes", () => decide("approved")); wire("rvNo", () => decide("rejected"));
    wire("rvSkip", () => move(1)); wire("rvBack", () => move(-1)); wire("rvAll", checkAll);
  }

  function render() {
    document.body.classList.toggle("empty", !(state.acs && state.suppliers.length));
    $("runBox").innerHTML = state.acs && state.suppliers.length && !state.ran ? `<button class="btn run" id="runCheck">Run check</button>` : "";
    if ($("runCheck")) $("runCheck").addEventListener("click", runCheck);
    renderFiles(); renderNotices(); renderReview();
    document.body.classList.toggle("checking", !!reviewing());
    // While checking: the card, then the current line's supplier tab, updating as each one is approved
    const now = currentLines()[0];
    if (now) state.active = Math.max(0, state.suppliers.findIndex((s) => s.name === now.supplier));
    if (state.step === 3) renderMarked(); else $("marked").innerHTML = "";
    renderTabs(); renderResult();
    if (state.focusReview) {
      state.focusReview = false;
      const b = $("rvYes") || document.querySelector("[data-out]") || $("rvFirst") || $("rvAgain");
      if (b) b.focus({ preventScroll: true });
    }
  }

  // ACS report date from its filename (e.g. "01.09.26"), used to explain recent hires
  function reportDate(name) {
    const m = /(\d{2})[.\-_](\d{2})[.\-_](\d{2,4})/.exec(name || "");
    if (!m) return null;
    const y = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${y}-${m[2]}-${m[1]}`;
  }

  // ---------- wiring ----------
  function wireDrop(zoneId, inputId, handler) {
    const zone = $(zoneId), input = $(inputId);
    input.addEventListener("change", () => { if (input.files.length) handler([...input.files]); input.value = ""; });
    ["dragenter", "dragover"].forEach((ev) => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add("over"); }));
    ["dragleave", "drop"].forEach((ev) => zone.addEventListener(ev, () => zone.classList.remove("over")));
    zone.addEventListener("drop", (e) => { e.preventDefault(); const f = [...e.dataTransfer.files]; if (f.length) handler(f); });
  }
  wireDrop("acsDrop", "acsInput", (files) => addACS(files[0]));
  wireDrop("supDrop", "supInput", (files) => addSuppliers(files));
  document.addEventListener("keydown", (e) => {
    if (!state.cross || !reviewing() || e.ctrlKey || e.metaKey || e.altKey || /^(INPUT|SELECT|TEXTAREA)$/.test(e.target.tagName)) return;
    const k = e.key.toLowerCase(), second = state.step === 2;
    const it = second ? state.second[state.spos] : null;
    const out = second && /^[1-9]$/.test(k) ? document.querySelectorAll("[data-out]")[+k - 1] : null;
    if (out) out.click();
    else if (k === "a" && (!second || it.p)) second ? answer("approved") : decide("approved");
    else if (k === "r" && (!second || it.p)) second ? answer("rejected") : decide("rejected");
    else if (k === "arrowright") second ? answer() : move(1);
    else if (k === "arrowleft") { if (second) secondBack(); else if (state.pos > 0) move(-1); }
    else return;
    e.preventDefault();
  });

  render();

  if (window.claude && window.claude.use) {
    window.claude.use("downloads").then((d) => { downloads = d; renderMarked(); }).catch(() => {});
  } else {
    // Standalone site: an ordinary browser download
    downloads = { save: async ({ filename, data }) => {
      const url = URL.createObjectURL(data);
      const a = Object.assign(document.createElement("a"), { href: url, download: filename });
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      await new Promise((r) => setTimeout(r, 400));   // let the browser start one download before the next
      return { status: "saved" };
    } };
    renderMarked();
    startMicrosoft().catch(() => {});
  }
})();
