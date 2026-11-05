(function () {
  const E = window.HireEngine;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const state = { acs: null, suppliers: [], active: 0, errors: [], cross: null, suggestFor: new Map(), queue: [], pos: 0, decisions: new Map(), ran: false };
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
  // Runs after every supplier report is checked: suggestions for unknown plant codes, then the £0 end pass
  function crossCheck() {
    state.cross = state.acs && state.suppliers.length ? E.crossCheck(state.acs, state.suppliers.map((s) => s.result)) : null;
    state.suggestFor = new Map(state.cross ? state.cross.suggestions.map((g) => [g.line, g]) : []);
    state.queue = state.cross ? E.pairings(state.acs, state.suppliers.map((s) => s.result), state.cross) : [];
    state.decisions = loadDecisions();
    state.pos = nextWaiting(0);
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
    for (let j = i; j < state.queue.length; j++) if (!decisionOf(state.queue[j])) return j;
    return state.queue.length;
  }
  function decide(v) {
    const p = state.queue[state.pos];
    if (!p) return;
    const wasDone = state.queue.every(decisionOf);
    state.decisions.set(p.id, v);
    saveDecisions();
    state.acsSave = null; for (const s of state.suppliers) s.saveState = null;
    state.pos = nextWaiting(state.pos + 1);
    if (state.pos >= state.queue.length) state.pos = Math.min(nextWaiting(0), state.queue.length);
    state.focusReview = true;
    render();
    // Last pairing checked: download both marked-up copies straight away
    if (!wasDone && state.queue.every(decisionOf)) saveAll();
  }
  function move(by) {
    state.pos = Math.max(0, Math.min(state.queue.length, state.pos + by));
    state.focusReview = true;
    render();
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
    E.markWorkbook(wb, s.result, supplierMarks(s));
    const out = await wb.xlsx.writeBuffer();
    return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }
  const checkedName = (s) => s.name.replace(/\.xlsx?m?$/i, "") + "_checked.xlsx";

  // Only approved pairings are marked: the supplier line gets the ACS cell(s) it was paired with
  const supplierMarks = (s) => approved().flatMap((p) => p.lines.filter((f) => f.supplier === s.name).map((f) => ({ sheet: f.sheet, row: f.row, acsAt: f.acsAt || p.acsAt })));
  const acsLines = () => E.acsApprovedLines(state.acs, approved());
  async function buildACS() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(state.acs.buf.slice(0));
    E.markACS(wb, acsLines());
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
    const charged = r.flagged.filter((f) => f.charged), free = r.flagged.filter((f) => !f.charged);
    const cols = r.sheets.map((x) => `${x.colLetter}${x.header ? " (" + esc(x.header) + ")" : ""}`).join(", ");

    const flaggedTable = charged.length ? `<div class="table-wrap"><table>
      <thead><tr><th>Row</th><th>Plant code</th><th>Plant</th><th>On hire</th><th class="num">Rate</th><th>Possible ACS match</th></tr></thead>
      <tbody>${charged.map((f) => `<tr class="${rowClass(f)}">
        <td class="num">${f.row}</td><td class="code">${esc(f.raw)}</td><td class="desc">${esc(f.desc)}</td>
        <td class="num">${fmtDate(f.date)}</td><td class="num">${fmtRate(f.rate)}</td>
        <td class="hint">${possible(f)}</td>
      </tr>`).join("")}</tbody></table></div>` : `<p class="empty">Every charged line on this report is on the ACS report.</p>`;

    const notChargedTable = !r.company ? `<p class="empty">No ACS plant matched this supplier, so its ACS lines can't be identified.</p>` :
      r.notCharged.length ? `<div class="table-wrap"><table>
      <thead><tr><th>Job</th><th>Plant</th><th>Asset no.</th><th>On hire</th><th class="num">ACS rate</th></tr></thead>
      <tbody>${r.notCharged.map((n) => `<tr><td>${esc(n.sheet)}</td><td class="desc">${esc(n.desc)}</td><td class="code">${n.missing.map(esc).join(", ")}${n.missing.length < n.codes.length ? ` <em class="hint">of ${esc(n.assetText)}</em>` : ""}</td><td class="num">${fmtDate(n.date)}</td><td class="num">${fmtRate(n.rate)}</td></tr>`).join("")}</tbody></table></div>`
      : `<p class="empty">Everything listed under ${esc(r.company)} on the ACS report is on this supplier report.</p>`;

    const byJob = new Map();
    for (const m of r.matched) { const k = m.jobs.join(", "); if (!byJob.has(k)) byJob.set(k, []); byJob.get(k).push(m); }
    const jobs = [...byJob.entries()].sort((a, b) => b[1].length - a[1].length);
    const matchedTable = r.matched.length ? `<div class="table-wrap"><table>
      <thead><tr><th>Row</th><th>Plant code</th><th>Plant</th><th>On hire</th><th class="num">Rate</th><th>ACS cell</th><th>Your check</th></tr></thead>
      <tbody>${jobs.map(([job, ms]) => `<tr class="group"><td colspan="7">${esc(job)} · ${ms.length}</td></tr>` + ms.map((m) => { const d = lineDecision(m); return `<tr class="${rowClass(m)}"><td class="num">${m.row}</td><td class="code">${esc(m.raw)}</td><td class="desc">${esc(m.desc)}</td><td class="num">${fmtDate(m.date)}</td><td class="num">${fmtRate(m.rate)}</td><td class="code">${m.acsAt.map(esc).join("<br>")}</td><td>${decisionPill(d)}</td></tr>`; }).join("")).join("")}</tbody></table></div>`
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
        <div class="block-head"><h3>Charged but not on the ACS report</h3><p>Plant codes on no ACS job sheet.</p></div>
        ${flaggedTable}
      </section>
      <section class="block">
        <div class="block-head"><h3>On the ACS report but not charged</h3><p>${r.company ? "Probably off-hired or swapped." : ""}</p></div>
        ${notChargedTable}
      </section>
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
    const acsRow = `<li class="mk-row">
        <span class="mk-name"><span class="code">${esc(checkedName(state.acs))}</span><small>ACS report · ${plural(lines.length, "line")} marked</small></span>
        ${pill(state.acsSave)}
        <button class="btn ghost" data-save="acs" ${!downloads || state.acsSave === "saving" ? "disabled" : ""}>${state.acsSave === "saved" ? "Save again" : "Download"}</button>
      </li>`;
    const rows = acsRow + state.suppliers.map((s, i) => {
      const n = supplierMarks(s).length, total = s.result.matched.length + s.result.flagged.length;
      return `<li class="mk-row">
        <span class="mk-name"><span class="code">${esc(checkedName(s))}</span><small>${plural(n, "line")} marked</small></span>
        ${pill(s.saveState)}
        <button class="btn ghost" data-save="${i}" ${!downloads || s.saveState === "saving" ? "disabled" : ""}>${s.saveState === "saved" ? "Save again" : "Download"}</button>
      </li>`;
    }).join("");
    const pending = state.acsSave === "saving" || state.suppliers.some((s) => s.saveState === "saving");
    box.innerHTML = `<section class="marked">
      <div class="mk-head">
        <div><span class="step">Step 4</span><h2>Marked-up spreadsheets</h2>
        <p>${waiting ? `<b>${plural(waiting, "pairing")} still to check.</b> ` : ""}Downloads by themselves once every pairing is checked.</p></div>
        <button class="btn" id="saveAll" ${!downloads || pending ? "disabled" : ""}>Download all (${state.suppliers.length + 1})</button>
      </div>
      <ul class="mk-list">${rows}</ul>
      ${downloads ? "" : `<p class="empty">Downloads aren't available in this view. Open the page on claude.ai, or run hire_check.py on your computer to make the same copies.</p>`}
    </section>`;
    box.querySelectorAll("[data-save]").forEach((b) => b.addEventListener("click", () => saveChecked(b.dataset.save === "acs" ? "acs" : +b.dataset.save)));
    const all = $("saveAll"); if (all) all.addEventListener("click", saveAll);
  }

  const afterReport = (f) => f.date && state.acs.reportDate && f.date > state.acs.reportDate;
  const supplierLabel = (name) => state.suppliers.length > 1 ? `${name.replace(/\.xlsx?m?$/i, "")}, ` : "";
  const sentence = (list) => list.length > 1 ? list.slice(0, -1).join(", ") + " and " + list[list.length - 1] : (list[0] || "");
  const money = (v) => { const n = E.money(v); return n ? "£" + n.toFixed(2) : "£0"; };

  function possible(f) {
    const g = state.suggestFor.get(f);
    if (g && g.acs) return `<span class="code">${esc(g.acs.cell)}</span> <em>${esc(g.acs.desc)}. See Suggestions.</em>`;
    return `<em>${afterReport(f) ? "Hired after the ACS report date" : "None found"}</em>`;
  }

  function renderSuggestions() {
    const box = $("suggest");
    if (!state.cross) { box.innerHTML = ""; return; }
    const { suggestions, zeroPass } = state.cross;
    const companies = new Set(state.suppliers.map((s) => (s.result.company || "").toLowerCase()).filter(Boolean));

    const unknown = suggestions.map((g) => {
      const f = g.line;
      const what = `<b>${esc(supplierLabel(f.supplier))}row ${f.row}: <span class="code">${esc(f.raw)}</span> ${esc(f.desc)}</b>, ${money(f.rate)}${f.date ? `, on hire ${fmtDate(f.date)}` : ""}.`;
      if (!g.acs) return `<li>${what} No likely ACS line.${afterReport(f) ? " On hire after the ACS report date: probably new plant." : ""}</li>`;
      const a = g.acs;
      return `<li>${what} Probably <span class="code">${esc(a.cell)}</span> ${esc(a.desc)} (asset <span class="code">${esc(a.free.join(", "))}</span>, ${money(a.rate)}${a.offHire ? ", off-hire" : ""}): ${esc(sentence(g.reasons))}.</li>`;
    }).join("");

    const zero = zeroPass.map((z) => {
      const a = z.acs;
      const what = `<b><span class="code">${esc(a.cell)}</span> ${esc(a.desc)}</b> (${esc(a.company || "no company")}, £0).`;
      if (z.line) return `<li>${what} Found on ${esc(z.line.supplier)} row ${z.line.row}: <span class="code">${esc(z.line.raw)}</span> ${esc(z.line.desc)}, ${money(z.line.rate)}. Reason: ${esc(sentence(z.reasons))}.</li>`;
      const noReport = a.company && !companies.has(a.company.toLowerCase());
      return `<li>${what} Not found on the supplier reports checked.${noReport ? ` No ${esc(a.company)} report has been checked yet.` : ""}</li>`;
    }).join("");

    box.innerHTML = `<section class="suggest">
      <div><span class="step">Suggestions</span><h2>Plant that needs a second look</h2>
</div>
      <div class="sg-part"><h3>Plant codes not on the ACS report</h3>
        ${unknown ? `<ul class="sg-list">${unknown}</ul>` : `<p class="empty">Every charged plant code is on the ACS report.</p>`}</div>
      <div class="sg-part"><h3>£0 items on the ACS report</h3>
        ${zero ? `<ul class="sg-list">${zero}</ul>` : `<p class="empty">No unpaired £0 items on the ACS report.</p>`}</div>
    </section>`;
  }

  // ---------- step 3: one pairing at a time ----------
  const rowClass = (m) => [lineDecision(m) === "approved" ? "matched" : "", reviewing() && state.queue[state.pos].lines.includes(m) ? "current" : ""].join(" ");
  const lineDecision = (m) => { const p = state.queue.find((q) => q.lines.includes(m)); return p ? decisionOf(p) : null; };
  const decisionPill = (d) => d === "approved" ? `<span class="pill ok">Approved</span>` : d === "rejected" ? `<span class="pill flag">Rejected</span>` : `<span class="pill warn">Waiting</span>`;
  const reviewing = () => state.cross && state.queue.length && state.pos < state.queue.length;
  const near = (a, b) => Math.abs(a - b) < 0.005;
  const shortName = (name) => name.replace(/\.xlsx?m?$/i, "");
  const missing = `<span class="pc-none">Not given</span>`;

  // The four things checked on every pairing: item, cost, company, on-hire date. ok: true agree, false don't (or missing).
  // An ACS line listing several assets is checked as one: every supplier line for those assets, costs added up.
  function checks(p) {
    const ls = p.lines, a = p.items[0] || {};
    const sup = state.suppliers.find((s) => s.name === ls[0].supplier);
    const company = sup && sup.result.company;
    const rates = ls.map((f) => E.money(f.rate));
    const fr = rates.some((r) => r == null) ? null : rates.reduce((t, r) => t + r, 0);
    const ar = a.zero ? 0 : a.rate;
    const fw = new Set(ls.flatMap((f) => [...E.words(f.desc)]));
    const sameWord = [...E.words(a.desc || "")].some((w) => fw.has(w));
    const cost = (v, qty) => v == null ? missing : `£${v.toFixed(2)}${qty > 1 ? ` <small>× ${qty}</small>` : ""}`;
    const name = (n, code) => `${n ? esc(n) : missing}${code ? `<span class="pc-code">${esc(code)}</span>` : ""}`;
    const group = ls.length > 1 || (a.codes && a.codes.length > 1);
    // Assets on the ACS line that this supplier hasn't charged for
    const left = group && p.kind === "match" ? a.codes.filter((c) => !p.codes.includes(c)).slice(p.stand.length) : [];
    const dates = [...new Set(ls.map((f) => f.date))];
    const rest = [
      { k: "Company", s: esc(company || shortName(ls[0].supplier)), a: a.company ? esc(a.company) : missing,
        ok: !!(company && a.company && company.toLowerCase() === a.company.toLowerCase()) },
      { k: "On hire", s: dates.map((d) => d ? fmtDate(d) : missing).join("<br>"), a: a.date ? fmtDate(a.date) : missing,
        ok: !!a.date && dates.length === 1 && dates[0] === a.date },
    ];
    if (group) {
      // Running tally: each supplier line adds its cost, ending on the ACS line's cost
      let run = 0;
      const lines = ls.map((f, i) => {
        if (rates[i] != null) run += rates[i];
        return `<tr><td>${name(f.desc, f.raw)}${p.stand.includes(f) && p.kind === "match" ? `<span class="rv-kind">Different code</span>` : ""}</td><td class="num">${cost(rates[i], 1)}</td><td class="num">£${run.toFixed(2)}</td></tr>`;
      }).concat(left.map((c) => `<tr class="pc-left"><td>Not charged<span class="pc-code">${esc(c)}</span></td><td class="num">–</td><td class="num">£${run.toFixed(2)}</td></tr>`));
      const ok = fr != null && ar != null && near(fr, ar) && !left.length;
      return { group: true, title: a.desc, rows: rest, tally: `<table class="tally"><tbody>${lines.join("")}</tbody>
        <tfoot><tr class="${ok ? "pc-ok" : "pc-bad"}"><th>Tally</th><td colspan="2" class="num"><b>£${run.toFixed(2)}</b> ${ok ? "=" : "≠"} ACS <b>${cost(ar, 1)}</b> <span class="pc-mark">${ok ? "✓" : "✗"}</span></td></tr></tfoot></table>` };
    }
    return { group: false, rows: [
      { k: "Item", s: name(ls[0].desc, ls[0].raw), a: name(a.desc, a.assetText || p.acsAt.join(", ")),
        ok: !!ls[0].desc && !!a.desc && (p.kind === "match" || sameWord) },
      { k: "Cost", s: cost(fr, ls[0].qty), a: cost(ar, a.qty),
        ok: fr != null && ar != null && (near(fr, ar) || near(fr * (a.qty || 1), ar) || near(fr * (ls[0].qty || 1), ar)) },
      ...rest,
    ] };
  }

  function renderReview() {
    const box = $("review");
    if (!state.cross) { box.innerHTML = ""; return; }
    const q = state.queue, n = q.length;
    const done = q.filter((p) => decisionOf(p)).length;
    if (!n) { box.innerHTML = `<section class="review"><p class="empty">No pairings found between the supplier reports and the ACS report.</p></section>`; return; }

    if (state.pos >= n) {
      const first = nextWaiting(0), yes = approved().length;
      box.innerHTML = `<section class="review rv-done">
        <p>${done === n ? `<b>All ${n} checked.</b> ${yes} approved, ${done - yes} rejected.` : `<b>${plural(n - done, "pairing")} skipped.</b>`}</p>
        <div class="rv-actions">${first < n ? `<button class="btn" id="rvFirst">Finish the skipped ones</button>` : ""}<button class="btn" id="rvAgain">Go through again</button></div></section>`;
      if ($("rvFirst")) $("rvFirst").addEventListener("click", () => { state.pos = first; state.focusReview = true; render(); });
      $("rvAgain").addEventListener("click", () => { state.pos = 0; state.focusReview = true; render(); });
      return;
    }

    const p = q[state.pos], d = decisionOf(p);
    // The ACS job sheet this pairing is on, as the card's title
    const job = [...new Set(p.items.length ? p.items.map((a) => a.sheet) : p.acsAt.map((c) => c.split("!")[0]))].join(", ");
    const c = checks(p);
    const rows = c.rows.map((c) => `<tr class="${c.ok ? "pc-ok" : "pc-bad"}"><th>${c.k}</th><td>${c.s}</td><td>${c.a}</td><td class="pc-mark">${c.ok ? "✓" : "✗"}</td></tr>`).join("");
    box.innerHTML = `<section class="review">
      <div class="rv-top"><h2 class="rv-job">${esc(job)}${c.group ? `<small>${esc(c.title)}</small>` : ""}</h2><span class="rv-pos"><b>${state.pos + 1}</b> of ${n}</span>${p.kind !== "match" ? `<span class="rv-kind">Different code</span>` : ""}${d ? decisionPill(d) : ""}</div>
      <div class="rv-bar" aria-hidden="true"><span style="width:${100 * done / n}%"></span></div>
      ${c.group ? c.tally : ""}
      <table class="pc"><thead><tr><th></th><th>Supplier</th><th>ACS</th><th></th></tr></thead><tbody>${rows}</tbody></table>
      <div class="rv-actions">
        <button class="btn big" id="rvYes">Approve <kbd>A</kbd></button>
        <button class="btn big" id="rvNo">Reject <kbd>R</kbd></button>
        <button class="btn" id="rvBack" ${state.pos ? "" : "disabled"}>Back</button>
        <button class="btn" id="rvSkip">Skip</button>
      </div></section>`;
    $("rvYes").addEventListener("click", () => decide("approved"));
    $("rvNo").addEventListener("click", () => decide("rejected"));
    $("rvSkip").addEventListener("click", () => move(1));
    $("rvBack").addEventListener("click", () => move(-1));
  }

  function render() {
    document.body.classList.toggle("empty", !(state.acs && state.suppliers.length));
    $("runBox").innerHTML = state.acs && state.suppliers.length && !state.ran ? `<button class="btn run" id="runCheck">Run check</button>` : "";
    if ($("runCheck")) $("runCheck").addEventListener("click", runCheck);
    renderFiles(); renderNotices(); renderReview();
    document.body.classList.toggle("checking", !!reviewing());
    // While checking: the card, then the current pairing's supplier tab, updating as each one is approved
    if (reviewing()) {
      const i = state.suppliers.findIndex((s) => s.name === state.queue[state.pos].line.supplier);
      if (i >= 0) state.active = i;
      $("suggest").innerHTML = $("marked").innerHTML = "";
    } else { renderSuggestions(); renderMarked(); }
    renderTabs(); renderResult();
    if (state.focusReview) {
      state.focusReview = false;
      const b = $("rvYes") || $("rvFirst") || $("rvAgain");
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
    if (!state.cross || !state.queue.length || e.ctrlKey || e.metaKey || e.altKey || /^(INPUT|SELECT|TEXTAREA)$/.test(e.target.tagName)) return;
    const k = e.key.toLowerCase();
    if (k === "a" && state.pos < state.queue.length) decide("approved");
    else if (k === "r" && state.pos < state.queue.length) decide("rejected");
    else if (k === "arrowright" && state.pos < state.queue.length) move(1);
    else if (k === "arrowleft" && state.pos > 0) move(-1);
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
  }
})();
