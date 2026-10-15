(function () {
  const E = window.HireEngine;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const state = { acs: null, suppliers: [], active: 0, errors: [], cross: null, suggestFor: new Map() };
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

  function rerun() {
    state.acsSave = null;
    for (const s of state.suppliers) s.saveState = null, s.result = E.checkSupplier(s.book, state.acs, s.override || null);
    crossCheck();
  }
  // Runs after every supplier report is checked: suggestions for unknown plant codes, then the £0 end pass
  function crossCheck() {
    state.cross = state.acs && state.suppliers.length ? E.crossCheck(state.acs, state.suppliers.map((s) => s.result)) : null;
    state.suggestFor = new Map(state.cross ? state.cross.suggestions.map((g) => [g.line, g]) : []);
  }

  async function addACS(file) {
    state.errors = [];
    try {
      const { book, buf } = await readBook(file);
      state.acsSave = null;
      state.acs = { name: file.name, buf, sheetNames: book.sheets.map((s) => s.title), reportDate: reportDate(file.name), ...E.readACS(book) };
      rerun();
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
    if (state.acs) rerun();
    render();
  }

  // ---------- download ----------
  async function buildChecked(s) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(s.buf.slice(0));
    E.markWorkbook(wb, s.result);
    const out = await wb.xlsx.writeBuffer();
    return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }
  const checkedName = (s) => s.name.replace(/\.xlsx?m?$/i, "") + "_checked.xlsx";

  const acsLines = () => E.acsChargedLines(state.acs, state.suppliers.map((s) => s.result));
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
    if (!state.acs && !state.suppliers.length) parts.push(`<div class="notice info"><b>Nothing checked yet.</b> Add this week's ACS on-hire report in step 1 and the supplier hire reports in step 2. The check runs on the files you add, every time.</div>`);
    else if (!state.acs) parts.push(`<div class="notice info"><b>Add the ACS on-hire report</b> in step 1 to check ${state.suppliers.length ? "these supplier reports" : "supplier reports"} against it.</div>`);
    else if (!state.suppliers.length) parts.push(`<div class="notice info"><b>Add supplier hire reports</b> in step 2. ${state.acs.codes.size} plant codes found across ${plural(state.acs.sheetNames.length, "job sheet")}.</div>`);
    for (const e of state.errors) parts.push(`<div class="notice error"><b>Not added.</b> ${esc(e)}</div>`);
    $("notices").innerHTML = parts.join("");
  }

  function counts(r) {
    const charged = r.flagged.filter((f) => f.charged);
    return { matched: r.matched.length, charged: charged.length, free: r.flagged.length - charged.length, notCharged: r.notCharged.length };
  }

  function renderTabs() {
    const tabs = $("tabs");
    if (!state.acs || !state.suppliers.length) { tabs.innerHTML = ""; return; }
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
    if (!state.acs || !state.suppliers.length) { box.innerHTML = ""; return; }
    const s = state.suppliers[state.active];
    const r = s.result, c = counts(r);
    const charged = r.flagged.filter((f) => f.charged), free = r.flagged.filter((f) => !f.charged);
    const cols = r.sheets.map((x) => `${x.colLetter}${x.header ? " (" + esc(x.header) + ")" : ""}`).join(", ");

    const flaggedTable = charged.length ? `<div class="table-wrap"><table>
      <thead><tr><th>Row</th><th>Plant code</th><th>Plant</th><th>On hire</th><th class="num">Rate</th><th>Possible ACS match</th></tr></thead>
      <tbody>${charged.map((f) => `<tr>
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
      <thead><tr><th>Row</th><th>Plant code</th><th>Plant</th><th>On hire</th><th class="num">Rate</th><th>ACS cell</th></tr></thead>
      <tbody>${jobs.map(([job, ms]) => `<tr class="group"><td colspan="6">${esc(job)} · ${ms.length}</td></tr>` + ms.map((m) => `<tr class="matched"><td class="num">${m.row}</td><td class="code">${esc(m.raw)}</td><td class="desc">${esc(m.desc)}</td><td class="num">${fmtDate(m.date)}</td><td class="num">${fmtRate(m.rate)}</td><td class="code">${m.acsAt.map(esc).join("<br>")}</td></tr>`).join("")).join("")}</tbody></table></div>`
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
        <div class="block-head"><h3>Charged but not on the ACS report</h3><p>Lines with a rate whose plant code is on none of the ACS job sheets. The possible match comes from the Suggestions above.</p></div>
        ${flaggedTable}
      </section>
      <section class="block">
        <div class="block-head"><h3>On the ACS report but not charged</h3><p>${r.company ? `Plant listed under ${esc(r.company)} on the ACS report that isn't on this supplier report. Probably off-hired, or swapped for another unit.` : ""}</p></div>
        ${notChargedTable}
      </section>
      <section class="block">
        <div class="block-head"><h3>On both reports</h3><p>Grouped by ACS job sheet. These lines are green and bold in the marked-up spreadsheet, with the ACS cell in a new "ACS match" column.</p></div>
        ${matchedTable}
      </section>
      ${free.length ? `<section class="block">${freeTable}</section>` : ""}
    </div>`;

    $("codeCol").addEventListener("change", (e) => { s.override = e.target.value; s.saveState = null; state.acsSave = null; s.result = E.checkSupplier(s.book, state.acs, s.override || null); crossCheck(); render(); });
  }

  function renderMarked() {
    const box = $("marked");
    const real = state.acs && state.suppliers.length;
    if (!real) { box.innerHTML = ""; return; }
    const label = { saving: "Saving…", saved: "Saved", declined: "Not saved", failed: "Not saved" };
    const pill = (st) => st ? `<span class="pill ${st === "saved" ? "ok" : st === "saving" ? "warn" : "flag"}">${label[st]}</span>` : "";
    const lines = acsLines(), acsTotal = state.acs.items.length;
    const acsRow = `<li class="mk-row">
        <span class="mk-name"><span class="code">${esc(checkedName(state.acs))}</span><small>ACS report · ${plural(lines.length, "line")} charged on the supplier reports in bold with a thick border, your colours kept · ${acsTotal - lines.length} of ${acsTotal} plant lines left unmarked</small></span>
        ${pill(state.acsSave)}
        <button class="btn ghost" data-save="acs" ${!downloads || state.acsSave === "saving" ? "disabled" : ""}>${state.acsSave === "saved" ? "Save again" : "Download"}</button>
      </li>`;
    const rows = acsRow + state.suppliers.map((s, i) => {
      const c = counts(s.result);
      return `<li class="mk-row">
        <span class="mk-name"><span class="code">${esc(checkedName(s))}</span><small>${plural(c.matched, "line")} green and bold with their ACS cell in an "ACS match" column · ${c.charged + c.free} not on ACS left unmarked</small></span>
        ${pill(s.saveState)}
        <button class="btn ghost" data-save="${i}" ${!downloads || s.saveState === "saving" ? "disabled" : ""}>${s.saveState === "saved" ? "Save again" : "Download"}</button>
      </li>`;
    }).join("");
    const pending = state.acsSave === "saving" || state.suppliers.some((s) => s.saveState === "saving");
    box.innerHTML = `<section class="marked">
      <div class="mk-head">
        <div><span class="step">Step 3</span><h2>Marked-up spreadsheets</h2>
        <p>The ACS report with plant that suppliers are charging for in bold with a thick border, keeping your own colours. Each supplier report with lines found on the ACS report in green and bold, plus the ACS sheet and cell each one matched. Your original files are not changed.</p></div>
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
      if (!g.acs) return `<li>${what} No likely ACS line found.${afterReport(f) ? " It went on hire after the ACS report date, so it may be new plant to add to the ACS report." : " Check whether it should be on the ACS report."}</li>`;
      const a = g.acs;
      return `<li>${what} Probably the ACS line at <span class="code">${esc(a.cell)}</span>: ${esc(a.desc)}, asset <span class="code">${esc(a.free.join(", "))}</span>, ${money(a.rate)}${a.offHire ? ", marked off-hire on the ACS report" : ""}. Reason: ${esc(sentence(g.reasons))}. If it's the same plant, update the asset number on the ACS report.</li>`;
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
      <p class="sg-note">Worked out after every report was checked. These are suggestions only, and nothing here is marked in the spreadsheets.</p></div>
      <div class="sg-part"><h3>Plant codes not on the ACS report</h3>
        <p class="sg-note">Each charged line whose plant code isn't on the ACS report, compared with the ACS lines no supplier has charged for, by on-hire date, rate, quantity and description.</p>
        ${unknown ? `<ul class="sg-list">${unknown}</ul>` : `<p class="empty">Every charged plant code is on the ACS report.</p>`}</div>
      <div class="sg-part"><h3>£0 items on the ACS report</h3>
        <p class="sg-note">Checked last: ACS lines at £0 that nothing else paired with, looked for on the supplier reports.</p>
        ${zero ? `<ul class="sg-list">${zero}</ul>` : `<p class="empty">No unpaired £0 items on the ACS report.</p>`}</div>
    </section>`;
  }

  function render() { renderFiles(); renderNotices(); renderSuggestions(); renderMarked(); renderTabs(); renderResult(); }

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
