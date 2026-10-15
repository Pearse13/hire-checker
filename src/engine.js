// Hire check engine: same matching rules as hire_check.py, plus the two-way report and suggestions.
(function (root) {
  const NOT_CODES = new Set(["CODE", "PLANTCODE", "PLANT", "ITEM", "ITEMCODE", "PRODUCT", "PRODUCTCODE",
    "REF", "REFERENCE", "STOCK", "STOCKCODE", "DESCRIPTION", "QTY", "QUANTITY", "DAYS", "RATE", "CHARGE",
    "AMOUNT", "TOTAL", "SUBTOTAL", "VAT", "NET", "GROSS"]);
  const SPLIT = /[+,;&]/;
  const STOP = new Set(["the", "and", "with", "for", "inc", "incl", "including", "donotuse", "set", "per", "week", "day",
    "standard", "heavy", "duty", "unit", "type", "new", "only", "from", "use"]);
  const isDate = (v) => Object.prototype.toString.call(v) === "[object Date]";

  // ExcelJS cell value -> plain value (string | number | Date | null)
  function plain(v) {
    if (v === null || v === undefined) return null;
    if (isDate(v) || typeof v === "number" || typeof v === "string") return v;
    if (typeof v === "boolean") return v;
    if (typeof v === "object") {
      if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join("");
      if ("result" in v) return plain(v.result);
      if ("formula" in v || "sharedFormula" in v) return null;
      if ("text" in v) return plain(v.text);
      if ("error" in v) return null;
    }
    return String(v);
  }

  function norm(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, "").toUpperCase();
  }

  function looksLikeCode(v) {
    if (v === null || v === undefined || isDate(v) || typeof v === "boolean") return false;
    const s = String(v).trim();
    if (!s || s.includes(" ") || s.length > 30) return false;
    const n = norm(s);
    if (NOT_CODES.has(n)) return false;
    return /[A-Z]/.test(n) || (/^\d+$/.test(n) && n.length >= 4);
  }

  // Codes inside one cell: the cell itself, or each part of "A1 + A2, A3"
  function cellCodes(v) {
    if (looksLikeCode(v)) return [norm(v)];
    if (typeof v === "string" && SPLIT.test(v)) {
      return v.split(SPLIT).filter(p => looksLikeCode(p) && /\d/.test(p)).map(norm);
    }
    return [];
  }

  const colLetter = (i) => { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const colIndex = (l) => l.toUpperCase().split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const dayKey = (v) => isDate(v) && !isNaN(v) ? v.toISOString().slice(0, 10) : null;
  const text = (v) => v === null || v === undefined ? "" : (isDate(v) ? dayKey(v) : String(v).replace(/\s+/g, " ").trim());

  // ---------- description helpers ----------
  // Words worth comparing: 3+ letters, plurals folded ("ladders" = "ladder")
  function words(s) {
    const out = new Set();
    for (let w of String(s || "").toLowerCase().match(/[a-z]{3,}/g) || []) {
      if (STOP.has(w)) continue;
      if (w.length > 4 && w.endsWith("s") && !w.endsWith("ss")) w = w.slice(0, -1);
      out.add(w);
    }
    return out;
  }
  // Sizes: "4m", "350mm", "3.0 kva", "20T", "8 cfm" -> {"4m", "350mm", "3kva", "20t", "8cfm"}
  function measures(s) {
    const out = new Map();
    const re = /(\d+(?:\.\d+)?)\s*\/?\s*(mm|mtrs?|metres?|meters?|m|tonnes?|t|kva|cfm|ft)(?![a-z])/gi;
    let m;
    while ((m = re.exec(String(s || "")))) {
      let unit = m[2].toLowerCase();
      if (/^(mtr|metre|meter)/.test(unit)) unit = "m";
      if (/^tonne/.test(unit)) unit = "t";
      out.set(parseFloat(m[1]) + unit, unit);
    }
    return out;
  }
  // Quantity written into a description: "2no", "2 no.", "x3", "3 x"
  function qtyFrom(s) {
    const m = /(\d+)\s*no\b\.?/i.exec(s || "") || /\bx\s*(\d+)\b/i.exec(s || "") || /^\s*(\d+)\s*x\b/i.exec(s || "");
    return m ? parseInt(m[1], 10) : 1;
  }
  function money(v) {
    if (typeof v === "number") return v;
    if (typeof v === "string") { const n = parseFloat(v.replace(/[£,\s]/g, "")); return isNaN(n) ? null : n; }
    return null;
  }
  const close = (a, b) => Math.abs(a - b) < 0.005;
  const days = (a, b) => Math.round(Math.abs(new Date(a) - new Date(b)) / 86400000);
  const gbp = (n) => "£" + n.toFixed(2);

  // book = {name, sheets:[{title, rows:[[plain values]]}]}
  function readACS(book) {
    const codes = new Map();   // code -> Set(sheet titles)
    const locs = new Map();    // code -> ["Sheet!C16", ...] every cell it appears in
    const items = [];          // plant lines, for the reverse check and suggestions
    for (const sh of book.sheets) {
      let cols = null;
      sh.rows.forEach((row, r) => {
        row.forEach((v, c) => cellCodes(v).forEach(code => {
          if (!codes.has(code)) { codes.set(code, new Set()); locs.set(code, []); }
          codes.get(code).add(sh.title);
          locs.get(code).push(`${sh.title}!${colLetter(c)}${r + 1}`);
        }));
        // header row: names the company and asset columns (can repeat mid-sheet)
        const hdr = row.map(v => typeof v === "string" ? v.trim().toLowerCase() : "");
        const asset = hdr.findIndex(h => /asset|plant\s*code|^code$/.test(h));
        const company = hdr.findIndex(h => /company|supplier|hire\s*co/.test(h));
        if (asset >= 0 && company >= 0) {
          cols = { asset, company, width: hdr.reduce((w, h, i) => h ? i + 1 : w, 0),
            desc: hdr.findIndex(h => /detail|desc|plant hired/.test(h)),
            date: hdr.findIndex(h => /date/.test(h)),
            rate: hdr.findIndex(h => /rate|price|cost/.test(h)),
            status: hdr.findIndex(h => /status/.test(h)) };
          return;
        }
        if (!cols) return;
        const assetCell = row[cols.asset];
        let parts = cellCodes(assetCell);
        if (!parts.length && typeof assetCell === "string" && /[A-Za-z]/.test(assetCell) && /\d/.test(assetCell) && assetCell.trim().length <= 30) parts = [norm(assetCell)];
        const desc = cols.desc >= 0 ? text(row[cols.desc]) : "";
        const comp = text(row[cols.company]);
        if (!parts.length || (!desc && !comp)) return;   // blank lines, colour keys
        const status = cols.status >= 0 ? text(row[cols.status]) : "";
        const rate = cols.rate >= 0 ? money(row[cols.rate]) : null;
        items.push({
          sheet: sh.title, row: r + 1, codes: parts, width: cols.width,
          cell: `${sh.title}!${colLetter(cols.asset)}${r + 1}`,
          assetText: text(row[cols.asset]), company: comp, desc,
          date: cols.date >= 0 ? dayKey(row[cols.date]) : null,
          rate, zero: !rate, qty: qtyFrom(desc),
          offHire: /off/i.test(status),
        });
      });
    }
    return { codes, locs, items };
  }

  function findCodeColumn(rows, codes) {
    const width = Math.max(0, ...rows.map(r => r.length));
    let best = -1, bestHits = 0;
    for (let c = 0; c < width; c++) {
      let hits = 0;
      for (const row of rows) { const v = row[c]; if (looksLikeCode(v) && codes.has(norm(v))) hits++; }
      if (hits > bestHits) { best = c; bestHits = hits; }
    }
    if (best >= 0) return best;
    for (const row of rows.slice(0, 30)) {
      const c = row.findIndex(v => typeof v === "string" && /code|plant|item|product/i.test(v));
      if (c >= 0) return c;
    }
    return -1;
  }

  // Header row above the first charge line: the nearest row whose code-column cell is text but not a code
  function supplierHeader(rows, col, firstRow) {
    for (let r = firstRow - 1; r >= 0; r--) {
      const v = rows[r][col];
      if (typeof v === "string" && v.trim() && !looksLikeCode(v)) {
        const hdr = rows[r].map(h => typeof h === "string" ? h.trim().toLowerCase() : "");
        const pick = (...res) => { for (const re of res) { const i = hdr.findIndex(h => re.test(h)); if (i >= 0) return i; } return -1; };
        return {
          row: r, labels: rows[r].map(text),
          desc: hdr.map((h, i) => /desc/.test(h) ? i : -1).filter(i => i >= 0),
          rate: pick(/^rate$/, /\brate\b/, /price|amount|cost|charge$/),
          date: pick(/start/, /on.?hire/, /^date$/, /date/),
          qty: pick(/^qty$/, /^quantity$/, /\bqty\b/, /quantity/),
        };
      }
    }
    return null;
  }

  function checkSupplier(book, acs, codeColOverride) {
    const sheets = [];
    const matched = [], flagged = [];
    const seen = new Set();
    for (const sh of book.sheets) {
      const col = codeColOverride ? colIndex(codeColOverride) : findCodeColumn(sh.rows, acs.codes);
      if (col < 0) continue;
      const lineRows = [];
      sh.rows.forEach((row, r) => { if (looksLikeCode(row[col])) lineRows.push(r); });
      if (!lineRows.length) continue;
      const hdr = supplierHeader(sh.rows, col, lineRows[0]);
      sheets.push({ title: sh.title, col, colLetter: colLetter(col), header: hdr ? hdr.labels[col] : "",
        headerRow: hdr ? hdr.row + 1 : null, width: Math.max(0, ...sh.rows.map(r => r.length)) });
      for (const r of lineRows) {
        const row = sh.rows[r];
        const code = norm(row[col]);
        seen.add(code);
        const line = {
          supplier: book.name, sheet: sh.title, row: r + 1, code, raw: text(row[col]),
          desc: hdr && hdr.desc.length ? [...new Set(hdr.desc.map(i => text(row[i])).filter(Boolean))].join(" ") : "",
          rate: hdr && hdr.rate >= 0 ? row[hdr.rate] : null,
          date: hdr && hdr.date >= 0 ? dayKey(row[hdr.date]) : null,
          qty: hdr && hdr.qty >= 0 ? (money(row[hdr.qty]) || 1) : 1,
          full: row.map(text).filter(x => x && x !== text(row[col])),
        };
        if (!line.desc) line.desc = line.full.slice(0, 3).join(" · ");
        line.charged = !hdr || hdr.rate < 0 ? true : !!money(line.rate);
        if (acs.codes.has(code)) {
          line.jobs = [...acs.codes.get(code)].sort();
          line.acsAt = acs.locs.get(code);
          matched.push(line);
        } else flagged.push(line);
      }
    }

    // Which supplier on the ACS report this is: the company behind most of the matches
    const matchedCodes = new Set(matched.map(m => m.code));
    const tally = new Map();
    for (const it of acs.items) if (it.codes.some(c => matchedCodes.has(c))) {
      const k = it.company.toLowerCase(); tally.set(k, (tally.get(k) || 0) + 1);
    }
    const companyKey = [...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || null;
    const company = companyKey ? acs.items.find(i => i.company.toLowerCase() === companyKey).company : null;
    for (const f of flagged) f.company = company;

    // On the ACS report under that company, but not on the supplier's report
    const notCharged = companyKey ? acs.items
      .filter(i => i.company.toLowerCase() === companyKey && !i.offHire)
      .map(i => ({ ...i, missing: i.codes.filter(c => !seen.has(c)) }))
      .filter(i => i.missing.length) : [];

    return { name: book.name, sheets, matched, flagged, company, notCharged };
  }

  // ---------- suggestions: run once every supplier report has been checked ----------
  // Compares a supplier line with an ACS line. Returns {score, reasons, strong} or null when they can't be the same plant.
  function compare(f, a, useRate) {
    let score = 0; const reasons = []; let dateEq = false, rateEq = false;
    // sizes that disagree (4m vs 5m ladder) rule the pair out
    const fm = measures(f.desc), am = measures(a.desc);
    const sameSize = [...fm.keys()].filter(k => am.has(k));
    const units = new Set([...fm.values()].filter(u => [...am.values()].includes(u)));
    if (units.size && !sameSize.length) return null;
    if (sameSize.length) { score += 2 * Math.min(sameSize.length, 2); reasons.push(`both ${sameSize.join(", ")}`); }

    const w = [...words(f.desc)].filter(x => words(a.desc).has(x));
    if (w.length) { score += Math.min(w.length, 3); reasons.push(`both "${w.slice(0, 3).join(" ")}"`); }

    if (f.date && a.date) {
      const d = days(f.date, a.date);
      if (d === 0) { score += 3; reasons.push("same on-hire date"); dateEq = true; }
      else if (d <= 3) { score += 1; reasons.push(`on-hire dates ${d} day${d > 1 ? "s" : ""} apart`); }
    }
    const fr = money(f.rate), ar = a.rate;
    if (useRate && fr > 0 && ar > 0) {
      if (close(fr, ar)) { score += 3; reasons.push(`same rate ${gbp(fr)}`); rateEq = true; }
      else if (a.qty > 1 && close(fr * a.qty, ar)) { score += 2; reasons.push(`${gbp(fr)} × ${a.qty} = ACS ${gbp(ar)}`); rateEq = true; }
    }
    if (f.qty > 1 && f.qty === a.qty) { score += 1; reasons.push(`both quantity ${a.qty}`); }
    if (f.company && a.company && f.company.toLowerCase() === a.company.toLowerCase()) { score += 1; reasons.push(`ACS lists it under ${a.company}`); }
    return { score, reasons, dateEq, rateEq, words: w.length, sizes: sameSize.length };
  }

  // Same plant, worded differently: the description agrees AND the date or rate agrees, or both date and rate agree.
  const likelySame = (c) => c.score >= 4 && (((c.words || c.sizes) && (c.dateEq || c.rateEq)) || (c.dateEq && c.rateEq));
  // £0 lines have no rate to compare: the description must share a word (a ladder is a ladder).
  const likelyZero = (c) => c.words > 0 && c.score >= 2;

  // Pairs highest scores first, each line used once (an ACS line once per unpaired asset on it).
  function pairUp(lines, items, useRate, accept) {
    const pairs = [];
    for (const f of lines) for (const a of items) {
      const c = compare(f, a, useRate);
      if (c && accept(c)) pairs.push({ f, a, ...c });
    }
    pairs.sort((x, y) => y.score - x.score);
    const used = new Set(), out = [];
    for (const p of pairs) {
      if (used.has(p.f) || p.a.left <= 0) continue;
      used.add(p.f); p.a.left--;
      out.push(p);
    }
    return out;
  }

  function crossCheck(acs, results) {
    const paired = new Set();
    for (const r of results) for (const m of r.matched) paired.add(m.code);
    // ACS lines with assets no supplier has charged for
    const open = acs.items
      .map(a => ({ ...a, free: a.codes.filter(c => !paired.has(c)) }))
      .filter(a => a.free.length)
      .map(a => ({ ...a, left: a.free.length }));
    const unmatched = results.flatMap(r => r.flagged);

    // 1. Plant code not on the ACS report: best unpaired ACS line by date, rate, quantity and description
    const found = pairUp(unmatched, open, true, likelySame);
    const bySupplierLine = new Map(found.map(p => [p.f, p]));
    const suggestions = unmatched
      .filter(f => f.charged || bySupplierLine.has(f))
      .map(f => {
        const p = bySupplierLine.get(f);
        return { line: f, acs: p ? p.a : null, reasons: p ? p.reasons : [], score: p ? p.score : 0 };
      });

    // 2. End pass: £0 ACS lines still unpaired, looked for among the supplier lines nothing else claimed
    const claimed = new Set(found.map(p => p.f));
    const zeroOpen = open.filter(a => a.zero && a.left === a.free.length);
    const spare = unmatched.filter(f => !claimed.has(f));
    const zeroFound = pairUp(spare, zeroOpen, false, likelyZero);
    const zeroByItem = new Map(zeroFound.map(p => [p.a.cell, p]));
    const zeroPass = zeroOpen.map(a => {
      const p = zeroByItem.get(a.cell);
      return { acs: a, line: p ? p.f : null, reasons: p ? p.reasons : [] };
    });
    return { suggestions, zeroPass };
  }

  // ---------- marked-up copies ----------
  const GREEN = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC6EFCE" }, bgColor: { argb: "FFC6EFCE" } };
  // cells share style objects in ExcelJS: always give a cell its own copy
  const restyle = (cell, extra) => {
    if (cell.isMerged && cell.master !== cell) return;
    cell.style = Object.assign({}, cell.style, extra(cell.style || {}));
  };

  // Supplier copy: matched lines green + bold, and an "ACS match" column naming the ACS sheet and cell.
  function markWorkbook(wb, result) {
    for (const sh of result.sheets) {
      const ws = wb.getWorksheet(sh.title);
      if (!ws) continue;
      const lines = result.matched.filter(m => m.sheet === sh.title);
      const extra = sh.width + 1;
      if (sh.headerRow) {
        const h = ws.getRow(sh.headerRow).getCell(extra);
        h.value = "ACS match";
        restyle(h, s => ({ font: Object.assign({}, s.font, { bold: true }) }));
      }
      for (const m of lines) {
        const row = ws.getRow(m.row);
        row.getCell(extra).value = m.acsAt.join("; ");
        for (let c = 1; c <= extra; c++) restyle(row.getCell(c), s => ({ fill: GREEN, font: Object.assign({}, s.font, { bold: true }) }));
      }
    }
  }

  // ACS lines charged on the supplier reports. A line listing several assets needs all of them charged.
  function acsChargedLines(acs, results) {
    const charged = new Set();
    for (const r of results) for (const m of r.matched) charged.add(m.code);
    return acs.items.filter(i => i.codes.every(c => charged.has(c)));
  }

  // ACS copy: charged lines bold with a thick border round the line. Existing colours are kept.
  function markACS(wb, lines) {
    const thick = { style: "thick", color: { argb: "FF000000" } };
    for (const l of lines) {
      const ws = wb.getWorksheet(l.sheet);
      if (!ws) continue;
      const row = ws.getRow(l.row);
      for (let c = 1; c <= l.width; c++) {
        restyle(row.getCell(c), s => ({
          font: Object.assign({}, s.font, { bold: true }),
          border: Object.assign({}, s.border, { top: thick, bottom: thick },
            c === 1 ? { left: thick } : {}, c === l.width ? { right: thick } : {}),
        }));
      }
    }
  }

  // ExcelJS reads <b val="0"/> ("bold off") as bold, so a plain open and save can turn whole sheets bold.
  // An "off" switch means the same as no switch: strip them from styles.xml before ExcelJS sees the file.
  async function cleanXlsx(buf, JSZip) {
    const zip = await JSZip.loadAsync(buf);
    const f = zip.file("xl/styles.xml");
    if (!f) return buf;
    const xml = await f.async("string");
    const fixed = xml.replace(/<(b|i|strike|outline|shadow|condense|extend)\s+val="(?:0|false)"\s*(?:\/>|>\s*<\/\1>)/g, "");
    if (fixed === xml) return buf;
    zip.file("xl/styles.xml", fixed);
    return zip.generateAsync({ type: "arraybuffer" });
  }

  const api = { cleanXlsx, crossCheck, acsChargedLines, markACS, markWorkbook, plain, norm, looksLikeCode, cellCodes,
    readACS, checkSupplier, colLetter, colIndex, words, measures, qtyFrom, money };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.HireEngine = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
