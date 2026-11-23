/* IDCO BRS Tool - the general statement reader.
 *
 * It does not know any bank's layout. It works from four ideas:
 *   1. The table is found by its column headings (Date / Debit or Withdrawal / Credit or Deposit / Balance).
 *   2. A transaction starts wherever a date sits in the date column.
 *   3. Columns are found from the data itself (amounts line up on their right edge, narration on its left edge).
 *   4. The amount of every line is PROVED from the change in the running balance and compared with what is printed.
 *
 * Input : pages = [{no, width, height, tokens:[{s,x0,x1,y,h}]}]   (see pdftext.js)
 * Output: { meta, opening, closing, rows, totals, printed, checks, warnings, layout }
 * All money is kept in integer paise so that nothing is ever lost to rounding. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementReader = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------- small helpers
  const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const DASH = /[‐-―−]/g;

  function parseDate(str) {
    const s = String(str).replace(DASH, '-').trim();
    let d, mo, y;
    let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
    if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1990 || y > 2100) return null; return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0'); }
    m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4}|\d{2})$/);
    if (m) { d = +m[1]; mo = +m[2]; y = +m[3]; }
    else {
      m = s.match(/^(\d{1,2})[-\/.\s]?([A-Za-z]{3})[A-Za-z]*[-\/.,\s]\s?(\d{4}|\d{2})$/);
      if (m) { d = +m[1]; mo = MONTHS[m[2].toLowerCase()]; y = +m[3]; }
      else {
        m = s.match(/^([A-Za-z]{3})[A-Za-z]*\.?\s+(\d{1,2}),?\s+(\d{4})$/);
        if (!m) return null;
        mo = MONTHS[m[1].toLowerCase()]; d = +m[2]; y = +m[3];
      }
    }
    if (!mo || d < 1 || d > 31 || mo > 12) return null;
    if (y < 100) y += 2000;
    if (y < 1990 || y > 2100) return null;
    return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  }

  // 1,234.00   -1,234.00   (-) 1,234.00   1,234.00-   (1,234.00)   1,234.00 Cr   1,234.00DR
  const RE_MONEY = /^(\(-\)|-|\()?\s*(\d[\d,]*\.\d{2})\s*(-|\))?\s*(cr|dr)?\.?$/i;
  const RE_FRAG = /^[\d,]*\.?[\d,]*\s*(cr|dr)?\.?$/i;      // a piece of a wrapped amount: "48,71,24,009.3" + "6CR"
  const RE_SIGN = /^(cr|dr|\(-\)|-)\.?$/i;

  function isMoney(s) { return RE_MONEY.test(s); }
  /** "1,23,456.78" -> 12345678 (paise). Returns null when it is not a clean amount. */
  function toPaise(digits) {
    const m = String(digits).match(/^(\d[\d,]*)\.(\d{2})$/);
    if (!m) return null;
    const whole = m[1].replace(/,/g, '');
    if (whole.length > 15) return null;
    return Number(whole) * 100 + Number(m[2]);
  }

  function median(a) { if (!a.length) return 0; const b = a.slice().sort((x, y) => x - y); return b[Math.floor(b.length / 2)]; }
  function norm(s) { return String(s).toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim(); }

  // ---------------------------------------------------------------- column headings
  const LABELS = [
    ['valuedate', /^value ?(date|dt)$/],
    ['date', /^((tran|trans|transaction|txn|post|posting|tr) ?)?(date|dt)$/],
    ['narr', /^(particulars?|description|narration|remarks|transaction (details|particulars|remarks)|details)$/],
    ['chq', /^(chq|cheque)( ?(no|number))?( ref(erence)?( no)?)?$|^(ref|reference)( ?(no|number))?$|^no reference$|^chq ref no$|^instrument( no)?$/],
    ['debit', /^(debits?|withdrawals?|withdrawal ?(amt|amount)( inr)?|debit ?(amt|amount)|dr|dr amount|withdrawal dr)$/],
    ['credit', /^(credits?|deposits?|deposit ?(amt|amount)( inr)?|credit ?(amt|amount)|cr|cr amount|deposit cr)$/],
    ['amount', /^((transaction|txn|tran) )?(amount|amt)( inr| rs)?$/],
    ['balance', /^(closing |running |available |ledger |account )?balance( inr| rs| amount| amt)?$/],
    ['sweepout', /^auto ?sweep$/],
    ['sweepin', /^reverse( sweep)?$/],
    ['mode', /^mode$/],
    ['branch', /^(init|init br|br|branch|branch code|tran br)$/]
  ];
  function labelKind(s) {
    const n = norm(s);
    if (!n || n.length > 32) return null;
    for (const [k, re] of LABELS) if (re.test(n)) return k;
    return null;
  }

  /** Find the transaction-table heading on a page. Returns null, or {y, yBottom, labels, foreign}. */
  function findHeader(page) {
    const toks = page.tokens;
    let best = null; let foreign = null;
    for (const b of toks) {
      if (labelKind(b.s) !== 'balance' && !/^balance\b/i.test(b.s.trim())) continue;
      const band = toks.filter(t => Math.abs(t.y - b.y) <= 10 && t.s.length <= 34);
      const labels = [];
      for (const t of band) { const k = labelKind(t.s); if (k) labels.push({ kind: k, x0: t.x0, x1: t.x1, y: t.y }); }
      const has = k => labels.some(l => l.kind === k);
      const yBottom = Math.max.apply(null, band.map(t => t.y)) + 1.5;
      if (labelKind(b.s) === 'balance' && has('date') && (has('debit') || has('credit') || has('amount'))) {
        if (!best || b.y < best.y) best = { y: Math.min.apply(null, band.map(t => t.y)), yBottom, labels };
      } else if (band.length >= 3) {
        if (!foreign || b.y < foreign.y) foreign = { y: b.y };
      }
    }
    return { header: best, foreign };
  }

  // Lines that are table furniture, not transactions.
  const RE_END = /^((page|grand|sub)\s*)?total\s*:?$|^transaction total$|^closing balance\s*:?$|^statement summary|end of statement|^page\s*(no\.?)?\s*:?\s*\d+(\s*(of|\/)\s*\d+)?\b|^legends?\s*:?$|^carried forward$|^c\/f$|^balance c\/f$|^\*+\s*end/i;
  const RE_FINAL = /^transaction total$|^closing balance\s*:?$|^statement summary|end of statement|^(grand\s*)?total$/i;
  const RE_TOTAL = /^((page|grand|sub)\s*)?total\s*:?$|^transaction total$/i;
  const RE_CLOSING = /^closing balance\s*:?$/i;
  const RE_OPENING = /^(opening balance|brought forward|b\/f|balance b\/f|balance brought forward|opening bal\.?)\s*:?$/i;

  // ---------------------------------------------------------------- un-merging cells
  // The PDF reader sometimes hands over two neighbouring cells as one piece of text
  // ("06-04-2026 NET BANKING", "0.00 3,67,634.00"). Split them back, estimating positions from the character widths.
  function piece(t, from, to) {
    const cw = t.s.length ? (t.x1 - t.x0) / t.s.length : 0;
    return { s: t.s.slice(from, to), x0: t.x0 + from * cw, x1: t.x0 + to * cw, y: t.y, h: t.h };
  }
  function splitMoneyRun(t) {
    if (isMoney(t.s) || !/\s/.test(t.s)) return null;
    const parts = []; const re = /\S+/g; let m;
    while ((m = re.exec(t.s))) parts.push({ from: m.index, to: m.index + m[0].length, s: m[0] });
    const groups = [];
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i];
      if (/^\(-\)$/.test(p.s) && parts[i + 1]) { groups.push({ from: p.from, to: parts[i + 1].to }); i++; continue; }
      if (/^(cr|dr)\.?$/i.test(p.s) && groups.length) { groups[groups.length - 1].to = p.to; continue; }
      groups.push({ from: p.from, to: p.to });
    }
    if (groups.length < 2) return null;
    const out = groups.map(g => piece(t, g.from, g.to));
    return out.every(o => isMoney(o.s)) ? out : null;
  }
  function unmerge(pages) {
    // where do dates start? (pure dates and date-led pieces together)
    const xs = [];
    const lead = /^(\d{4}[-\/.]\d{1,2}[-\/.]\d{1,2}|\d{1,2}[-\/.\u2010-\u2015]\d{1,2}[-\/.\u2010-\u2015]\d{2,4}|\d{1,2}[- ][A-Za-z]{3}[- ]\d{2,4})\s+\S/;
    for (const p of pages) for (const t of p.tokens) if (parseDate(t.s) || (lead.test(t.s) && parseDate(t.s.match(lead)[1]))) xs.push(t.x0);
    xs.sort((a, b) => a - b);
    const cl = [];
    for (const x of xs) { const c = cl[cl.length - 1]; if (c && x - c.max <= 4) { c.max = x; c.n++; } else cl.push({ min: x, max: x, n: 1 }); }
    const mx = cl.length ? Math.max.apply(null, cl.map(c => c.n)) : 0;
    const col = cl.find(c => c.n >= 0.5 * mx);
    return pages.map(p0 => {
      // a number the PDF hands over in pieces ("1,23," + "456.00") is put back together
      const srt = p0.tokens.slice().sort((a, b) => (Math.abs(a.y - b.y) <= 0.8 ? a.x0 - b.x0 : a.y - b.y));
      const glued = [];
      for (const t of srt) {
        const q = glued[glued.length - 1];
        if (q && Math.abs(q.y - t.y) <= 0.8 && t.x0 - q.x1 >= -0.6 && t.x0 - q.x1 <= 1.2 && /^[\d,.]+$/.test(q.s) && /^[\d,.]+$/.test(t.s) && !/\.\d{2}$/.test(q.s))
          glued[glued.length - 1] = { s: q.s + t.s, x0: q.x0, x1: t.x1, y: q.y, h: q.h };
        else glued.push(t);
      }
      const p = { no: p0.no, width: p0.width, height: p0.height, tokens: glued };
      const tokens = [];
      for (const t of p.tokens) {
        const m = col && t.x0 >= col.min - 3 && t.x0 <= col.max + 3 ? t.s.match(lead) : null;
        if (m && parseDate(m[1])) {
          const rest = t.s.slice(m[1].length); const off = m[1].length + (rest.length - rest.replace(/^\s+/, '').length);
          tokens.push(piece(t, 0, m[1].length));
          const r = piece(t, off, t.s.length);
          const sm = splitMoneyRun(r); if (sm) tokens.push.apply(tokens, sm); else tokens.push(r);
          continue;
        }
        const sm = splitMoneyRun(t);
        if (sm) tokens.push.apply(tokens, sm); else tokens.push(t);
      }
      return { no: p.no, width: p.width, height: p.height, tokens };
    });
  }

  // ---------------------------------------------------------------- main
  function parseStatement(pages, opts) {
    opts = opts || {};
    const first = parseOnce(pages, opts);
    if (opts.anchorBy || !first.rows || first.rows.length < 3) return first;
    const bad = first.rows.filter(r => !r.ok).length;
    if (bad <= 0.3 * first.rows.length) return first;
    const second = parseOnce(pages, Object.assign({}, opts, { anchorBy: 'balance' }));
    const bad2 = second.rows ? second.rows.filter(r => !r.ok).length : Infinity;
    if (second.rows && second.rows.length >= first.rows.length && bad2 < 0.1 * bad + 1) { second.layout.rowsFrom = 'balance column (dates are not repeated on every line)'; return second; }
    return first;
  }
  function parseOnce(pages, opts) {
    pages = unmerge(pages);
    const warnings = [];
    const out = { meta: {}, opening: null, closing: null, rows: [], totals: {}, printed: {}, checks: [], warnings, layout: {} };

    // ---- 1. which pages carry the transaction table
    let open = false, ended = false;
    const tpages = [];
    for (const p of pages) {
      const { header, foreign } = findHeader(p);
      let top = null;
      if (header) { top = header.yBottom; open = true; ended = false; }
      else if (foreign) { open = false; ended = true; }
      else if (open && !ended) { top = 0; }
      if (top === null) continue;
      // where does the table stop on this page
      let bottom = p.height + 1; let finalSeen = false;
      for (const t of p.tokens) {
        if (t.y <= top) continue;
        if (RE_END.test(t.s.trim()) && t.y < bottom) { bottom = t.y - 1.5; }
        if (RE_FINAL.test(t.s.trim())) finalSeen = true;
      }
      if (foreign && header && foreign.y > header.yBottom && foreign.y - 12 < bottom) { bottom = foreign.y - 12; finalSeen = true; }
      tpages.push({ page: p, header, top, bottom, furniture: p.tokens.filter(t => t.y > top && t.y >= bottom) });
      if (finalSeen) { ended = true; }
    }
    if (!tpages.length) {
      let noBal = false;
      for (const p of pages) { const ks = new Set(p.tokens.map(t => labelKind(t.s)).filter(Boolean)); if (ks.has('date') && (ks.has('debit') || ks.has('credit') || ks.has('amount')) && !ks.has('balance')) noBal = true; }
      if (noBal) { out.noBalanceColumn = true; warnings.push('This statement has no running Balance column, so its lines cannot be proved. Ask the bank for a statement that shows the balance after every transaction.'); }
      else warnings.push('No transaction table was found. The file may be a scanned image, or not a bank statement.');
      out.checks.push({ name: 'Transaction table found', ok: false, detail: 'No headings like Date / Withdrawal / Deposit / Balance were found.' });
      out.pass = false;
      return out;
    }
    const firstHeader = (tpages.find(tp => tp.header) || {}).header;

    // tokens inside the table area
    for (const tp of tpages) tp.toks = tp.page.tokens.filter(t => t.y > tp.top && t.y < tp.bottom);

    // ---- 2. the column model, learned from the data of the whole document
    // 2a. amount columns: money lines up on its right edge
    const moneyX1 = [];
    for (const tp of tpages) for (const t of tp.toks) if (isMoney(t.s)) moneyX1.push(t.x1);
    moneyX1.sort((a, b) => a - b);
    let clusters = [];
    for (const x of moneyX1) {
      const c = clusters[clusters.length - 1];
      if (c && x - c.max <= 3) { c.max = x; c.n++; c.sum += x; } else clusters.push({ min: x, max: x, n: 1, sum: x });
    }
    clusters.forEach(c => { c.x1 = c.sum / c.n; });

    // 2b. date column: the left-most column of pure dates
    const dateX0 = [];
    for (const tp of tpages) for (const t of tp.toks) if (parseDate(t.s)) dateX0.push(t.x0);
    dateX0.sort((a, b) => a - b);
    const dclusters = [];
    for (const x of dateX0) {
      const c = dclusters[dclusters.length - 1];
      if (c && x - c.max <= 4) { c.max = x; c.n++; } else dclusters.push({ min: x, max: x, n: 1 });
    }
    if (!dclusters.length || !clusters.length) {
      warnings.push('The table headings were found but no dated lines with amounts could be read.');
      out.checks.push({ name: 'Transaction lines found', ok: false, detail: 'No dated lines with amounts.' });
      out.pass = false;
      return out;
    }
    const maxDates = Math.max.apply(null, dclusters.map(c => c.n));
    const dateCol = opts.anchorBy === 'balance' ? dclusters.find(c => c.n >= Math.max(2, 0.02 * maxDates)) : dclusters.find(c => c.n >= 0.5 * maxDates);
    const inDateCol = t => t.x0 >= dateCol.min - 3 && t.x0 <= dateCol.max + 3 && parseDate(t.s);
    const otherDateCols = dclusters.filter(c => c !== dateCol && c.n >= 0.3 * maxDates);
    const inOtherDateCol = t => parseDate(t.s) && otherDateCols.some(c => t.x0 >= c.min - 3 && t.x0 <= c.max + 3);

    // amount clusters must lie to the right of the date column
    clusters = clusters.filter(c => c.x1 > dateCol.max + 40);
    // balance column: the right-most amount column that appears on most lines
    const anchorsTotal = dateX0.filter(x => x >= dateCol.min - 3 && x <= dateCol.max + 3).length;
    let balCluster = null;
    for (let i = clusters.length - 1; i >= 0; i--) if (clusters[i].n >= 0.5 * anchorsTotal) { balCluster = clusters[i]; break; }
    if (!balCluster) balCluster = clusters[clusters.length - 1];
    clusters = clusters.filter(c => c.x1 <= balCluster.x1 + 1);
    balCluster.kind = 'balance';
    // name the other amount columns from the nearest heading
    const amtLabels = firstHeader ? firstHeader.labels.filter(l => ['debit', 'credit', 'sweepout', 'sweepin'].includes(l.kind)) : [];
    const pairs = [];
    for (const c of clusters) if (c !== balCluster) for (const l of amtLabels) pairs.push({ c, l, d: Math.abs(c.x1 - l.x1) });
    pairs.sort((a, b) => a.d - b.d);
    const usedL = new Set();
    for (const pr of pairs) {
      if (pr.c.kind || usedL.has(pr.l) || pr.d > 60) continue;
      pr.c.kind = pr.l.kind; usedL.add(pr.l);
    }
    clusters.forEach(c => { if (!c.kind) c.kind = 'other'; });
    const clusterOf = (t, slack) => {
      let best = null, bd = slack || 3.2;
      for (const c of clusters) { const d = Math.abs(t.x1 - c.x1); if (d <= bd) { bd = d; best = c; } }
      // a lone Cr / Dr printed in its own narrow column just to the right of the balance
      if (!best && /^(cr|dr)\.?$/i.test(t.s.trim()) && t.x0 >= balCluster.x1 - 1 && t.x0 <= balCluster.x1 + 28) return balCluster;
      return best;
    };
    const firstMoneyLeft = Math.min.apply(null, clusters.map(c => c.x1)) - 30;

    // 2c. narration column: the left edge that carries the most text
    const hist = new Map();
    for (const tp of tpages) for (const t of tp.toks) {
      if (parseDate(t.s) || isMoney(t.s) || RE_SIGN.test(t.s)) continue;
      if (t.x0 <= dateCol.max + 3) continue;
      if (RE_FRAG.test(t.s) && clusterOf(t)) continue;
      const k = Math.round(t.x0);
      hist.set(k, (hist.get(k) || 0) + t.s.length);
    }
    let narrX0 = null, bestChars = -1;
    for (const [k, v] of hist) {
      const tot = v + (hist.get(k - 1) || 0) + (hist.get(k + 1) || 0);
      if (tot > bestChars) { bestChars = tot; narrX0 = k; }
    }
    const isNarrStart = t => narrX0 !== null && Math.abs(t.x0 - narrX0) <= 2.2 && !inDateCol(t);

    const sideLabelsAll = (firstHeader ? firstHeader.labels : []).filter(l => ['chq', 'mode', 'branch', 'valuedate'].includes(l.kind));
    const underSideLabel = t => { const cx = (t.x0 + t.x1) / 2; return sideLabelsAll.some(l => cx >= l.x0 - 12 && cx <= l.x1 + 12); };
    // ---- 3. per page: anchors (dates), narration lines, other tokens
    let narrMaxX1 = 0; const narrMinX0 = narrX0 === null ? 0 : narrX0 - 1;
    for (const tp of tpages) {
      const toks = tp.toks.slice().sort((a, b) => a.y - b.y || a.x0 - b.x0);
      const dateToks = toks.filter(inDateCol);
      if (opts.anchorBy === 'balance') {
        // some banks print the date only on the first transaction of a day: then every balance figure starts a row
        const ys = [];
        // pieces of the balance column, top to bottom; a piece that only continues a wrapped figure does not start a row
        const bits = toks.filter(t => /\d/.test(t.s) && (isMoney(t.s) || (RE_FRAG.test(t.s) && /[.,]/.test(t.s)) || (RE_FRAG.test(t.s) && t.s.length <= 5)) && clusterOf(t) === balCluster).sort((a, b) => a.y - b.y);
        let open = null;
        for (const t of bits) {
          const cont = open && t.y - open.y > 1.5 && t.y - open.y <= 13 && !isMoney(open.s);
          const startsRow = isMoney(t.s) || /[.,]/.test(t.s);      // a bare digit can only be the tail of a wrapped figure
          if (!cont && startsRow && !ys.some(y => Math.abs(y - t.y) <= 1.5)) ys.push(t.y);
          open = cont ? null : t;
        }
        tp.anchors = ys.map(y => { const dt = dateToks.find(t => Math.abs(t.y - y) <= 2.5); return { y, date: dt ? parseDate(dt.s) : null, tok: dt || null, lines: [], extra: [], page: tp.page.no }; });
      } else tp.anchors = dateToks.map(t => ({ y: t.y, date: parseDate(t.s), tok: t, lines: [], extra: [], page: tp.page.no }));
      const used = new Set(dateToks);
      // narration lines: start at the narration edge, plus anything that continues on the same line
      const lines = [];
      for (const t of toks) {
        if (used.has(t) || !isNarrStart(t)) continue;
        if (isMoney(t.s) && clusterOf(t)) continue;
        const line = { y: t.y, x0: t.x0, x1: t.x1, text: t.s, toks: [t] };
        used.add(t);
        let grew = true;
        while (grew) {
          grew = false;
          for (const u of toks) {
            if (used.has(u) || Math.abs(u.y - line.y) > 1.3) continue;
            if (u.x0 < line.x1 - 1 || u.x0 > line.x1 + 70 || u.x0 > firstMoneyLeft) continue;
            if (isMoney(u.s) && clusterOf(u)) continue;
            if (underSideLabel(u) && u.x0 > line.x1 + 6) continue;      // a cheque no. or branch code in its own column
            line.text += (u.x0 - line.x1 > 1 ? ' ' : '') + u.s; line.x1 = u.x1; line.toks.push(u); used.add(u); grew = true;
          }
        }
        lines.push(line);
      }
      // opening / total / closing lines are not narration
      tp.special = [];
      tp.lines = [];
      for (const l of lines) {
        const s = l.text.trim();
        if (RE_OPENING.test(s) || RE_CLOSING.test(s) || RE_TOTAL.test(s)) tp.special.push(l); else { tp.lines.push(l); if (l.x1 > narrMaxX1) narrMaxX1 = l.x1; }
      }
      tp.rest = toks.filter(t => !used.has(t));
    }

    if (opts.anchorBy === 'balance') { let lastDate = null; for (const tp of tpages) for (const a of tp.anchors) { if (a.date) lastDate = a.date; else { a.date = lastDate; a.dateCarried = true; } } }
    // typical line spacing
    const gaps = [];
    for (const tp of tpages) for (let i = 1; i < tp.lines.length; i++) { const g = tp.lines[i].y - tp.lines[i - 1].y; if (g > 3 && g < 30) gaps.push(g); }
    const gsorted = gaps.slice().sort((a, b) => a - b);
    const s = gsorted.length ? gsorted[Math.floor(gsorted.length * 0.15)] : 9;   // the tight (within-row) spacing
    const tol = Math.max(2, 0.38 * s);

    // ---- 4. which narration lines belong to which date: top / centre / bottom aligned
    function assign(mode, commit) {
      let v = 0; let prevLast = null; let carry = [];
      for (const tp of tpages) {
        const A = tp.anchors, L = tp.lines;
        const sets = A.map(() => []);
        if (mode === 'top') {
          let li = 0;
          const pre = [];
          while (li < L.length && A.length && L[li].y < A[0].y - tol) pre.push(L[li++]);
          if (!A.length) { pre.push.apply(pre, L.slice(li)); li = L.length; }
          v += pre.length;
          if (commit && prevLast && pre.length) prevLast.lines.push.apply(prevLast.lines, pre);
          for (let i = 0; i < A.length; i++) {
            const lim = i + 1 < A.length ? A[i + 1].y - tol : Infinity;
            let prevY = null;
            while (li < L.length && L[li].y < lim) {
              if (i + 1 >= A.length && prevY !== null && L[li].y - prevY > 2.6 * s) { li = L.length; break; }
              sets[i].push(L[li]); prevY = L[li].y; li++;
            }
            if (!sets[i].some(l => Math.abs(l.y - A[i].y) <= tol + 0.5)) v++;
          }
        } else if (mode === 'bottom') {
          let li = 0;
          for (let i = 0; i < A.length; i++) {
            if (i === 0 && carry.length) { sets[0].push.apply(sets[0], carry); }
            while (li < L.length && L[li].y <= A[i].y + tol) { sets[i].push(L[li]); li++; }
            const last = sets[i][sets[i].length - 1];
            if (!last || Math.abs(last.y - A[i].y) > tol + 0.5) v++;
          }
          carry = L.slice(li).filter((l, k, arr) => true);
          // lines after the last date on a page should not exist in a bottom-aligned layout
          const trailing = L.slice(li).filter(l => !A.length || l.y - A[A.length - 1].y < 3 * s);
          v += trailing.length; carry = trailing;
        } else { // centre
          let li = 0;
          for (let i = 0; i < A.length; i++) {
            const above = [];
            while (li < L.length && L[li].y < A[i].y - 2.0) above.push(L[li++]);
            const on = [];
            while (li < L.length && Math.abs(L[li].y - A[i].y) <= 2.0) on.push(L[li++]);
            const below = [];
            while (below.length < above.length && li < L.length && (i + 1 >= A.length || L[li].y < A[i + 1].y + 2.0)) below.push(L[li++]);
            sets[i] = above.concat(on, below);
            if (below.length < above.length) v++;
            else if (sets[i].length) {
              const mean = sets[i].reduce((q, l) => q + l.y, 0) / sets[i].length;
              if (Math.abs(mean - A[i].y) > 2.4) v++;
            }
          }
        }
        if (commit) for (let i = 0; i < A.length; i++) A[i].lines = sets[i];
        if (A.length) prevLast = A[A.length - 1];
      }
      return v;
    }
    const score = { centre: assign('centre'), top: assign('top'), bottom: assign('bottom') };
    let mode = 'centre';
    if (score.top < score[mode]) mode = 'top';
    if (score.bottom < score[mode]) mode = 'bottom';
    // A statement with very few wrapped narrations can leave the three readings level.
    // Then let the wrapped lines themselves decide: a line sitting just under one date and far above the next belongs to the one above.
    const best = Math.min(score.centre, score.top, score.bottom);
    if (['centre', 'top', 'bottom'].filter(k => score[k] === best).length > 1) {
      let up = 0, down = 0;
      for (const tp of tpages) for (const l of tp.lines) {
        let prevA = null, nextA = null;
        for (const a of tp.anchors) { if (a.y <= l.y + 0.01) prevA = a; else if (!nextA) nextA = a; }
        if (prevA && Math.abs(prevA.y - l.y) <= tol) continue;
        if (nextA && Math.abs(nextA.y - l.y) <= tol) continue;
        const dPrev = prevA ? l.y - prevA.y : Infinity, dNext = nextA ? nextA.y - l.y : Infinity;
        if (dPrev < 0.8 * dNext) up++; else if (dNext < 0.8 * dPrev) down++;
      }
      const tied = k => score[k] === best;
      if (up > 2 * down && tied('top')) mode = 'top';
      else if (down > 2 * up && tied('bottom')) mode = 'bottom';
      else if (tied('centre')) mode = 'centre';
      out.layout.alignmentVotes = { linesBelongingAbove: up, linesBelongingBelow: down };
    }
    tpages.forEach(tp => tp.anchors.forEach(a => { a.lines = []; }));
    assign(mode, true);
    const allAnchors = [];
    tpages.forEach(tp => allAnchors.push.apply(allAnchors, tp.anchors));
    out.layout = { alignment: mode, alignmentScores: score, lineSpacing: +s.toFixed(2), narrationX: narrX0, dateX: +dateCol.min.toFixed(1),
      amountColumns: clusters.map(c => ({ x1: +c.x1.toFixed(1), kind: c.kind, n: c.n })), pagesWithTable: tpages.map(tp => tp.page.no) };
    if (allAnchors.length && score[mode] > Math.max(2, 0.02 * allAnchors.length))
      warnings.push('The narration of some lines may be attached to the neighbouring line (layout not fully understood). Amounts and balances are not affected.');

    // ---- 5. attach the remaining tokens (amounts, value date, cheque no, etc.) to their row
    const margin = Math.min(0.45 * s, 4);
    for (const tp of tpages) {
      const A = tp.anchors;
      A.forEach(a => {
        const ys = a.lines.filter(l => l.toks[0] && tp.toks.includes(l.toks[0])).map(l => l.y).concat([a.y]);
        a.yTop = Math.min.apply(null, ys) - margin; a.yBot = Math.max.apply(null, ys) + margin;
      });
      tp.orphans = [];
      for (const t of tp.rest) {
        let row = A.find(a => t.y >= a.yTop && t.y <= a.yBot);
        if (!row) {
          let bd = 1.25 * s;
          for (const a of A) { const d = Math.abs(t.y - a.y); if (d < bd) { bd = d; row = a; } }
        }
        if (row) row.extra.push(t); else tp.orphans.push(t);
      }
    }

    function readCell(frags) {
      // frags: tokens of one amount cell, possibly wrapped over lines: digits, "Cr"/"Dr", "(-)"
      if (!frags.length) return null;
      frags = frags.slice().sort((a, b) => (Math.abs(a.y - b.y) <= 1.3 ? a.x0 - b.x0 : a.y - b.y));
      let neg = false, digits = '';
      for (const f of frags) {
        let t = f.s.replace(/\s+/g, '');
        if (/\(-\)/.test(t) || /^-/.test(t) || /-$/.test(t) || /^\(.*\)$/.test(t)) neg = true;
        if (/dr\.?$/i.test(t)) neg = true;
        t = t.replace(/\(-\)/g, '').replace(/(cr|dr)\.?$/i, '').replace(/^[-(]/, '').replace(/[-)]$/, '');
        digits += t;
      }
      if (!digits) return null;
      const p = toPaise(digits);
      if (p === null) return { bad: true, raw: frags.map(f => f.s).join(' ') };
      return { paise: neg ? -p : p };
    }

    const hdrLabels = firstHeader ? firstHeader.labels : [];
    const sideLabels = hdrLabels.filter(l => ['chq', 'mode', 'branch'].includes(l.kind));
    const rowsRaw = [];
    for (const tp of tpages) {
      // opening balance printed on a line of its own (no date)
      for (const sp of tp.special) {
        const cellToks = tp.rest.concat(tp.furniture).filter(t => Math.abs(t.y - sp.y) <= 3.5 && (isMoney(t.s)) && clusterOf(t, 14));
        sp.cells = {};
        for (const t of cellToks) { const c = clusterOf(t, 14); const v = readCell([t]); if (v && !v.bad) sp.cells[c.kind] = v.paise; }
        sp.kindOf = RE_OPENING.test(sp.text.trim()) ? 'opening' : RE_CLOSING.test(sp.text.trim()) ? 'closing' : 'total';
        sp.page = tp.page.no;
        if (sp.kindOf === 'opening') {
          // remove the opening figure from any row that may have picked it up
          for (const a of tp.anchors) {
            a.extra = a.extra.filter(t => !cellToks.includes(t));
            if (Math.abs(a.y - sp.y) <= 3.5) a.skip = true;      // a dated "B/F" line is the opening balance itself
          }
          rowsRaw.push({ special: 'opening', y: sp.y - 0.01, page: tp.page.no, bal: sp.cells.balance });
        }
      }
      for (const a of tp.anchors) if (!a.skip) rowsRaw.push(a);
    }
    rowsRaw.sort((p, q) => p.page - q.page || p.y - q.y);

    // totals / closing lines printed by the bank (they sit outside the table area)
    const printedTotals = []; let printedClosing = null;
    for (const tp of tpages) {
      const zone = tp.page.tokens.filter(t => t.y > tp.top);
      for (const t of zone) {
        const txt = t.s.trim();
        const isTot = RE_TOTAL.test(txt), isClo = RE_CLOSING.test(txt);
        if (!isTot && !isClo) continue;
        const cells = {};
        for (const u of zone) {
          if (u === t || Math.abs(u.y - t.y) > 3.5 || !isMoney(u.s)) continue;
          const c = clusterOf(u, 14); if (!c) continue;
          const v = readCell([u]); if (v && !v.bad) cells[c.kind] = v.paise;
        }
        if (isClo && cells.balance !== undefined) printedClosing = cells.balance;
        if (isTot && Object.keys(cells).length) printedTotals.push({ page: tp.page.no, cells });
      }
    }

    // ---- 6. build the rows and prove every one with the running balance
    function joinLines(lines) {
      let text = '';
      lines.forEach((l, i) => {
        const t = l.text.trim();
        if (i === 0) { text = t; return; }
        const p = lines[i - 1]; const pt = p.text.trim();
        // a line that is one unbroken word and nearly as wide as the column was cut in mid-word: join without a space
        // (a space inside a reference number is the worse mistake, so a number continuing on the next line is always joined)
        const share = (p.x1 - p.x0) / Math.max(1, narrMaxX1 - narrMinX0);
        const glue = /[-\/*_:~]$/.test(pt) || (!/\s/.test(pt) && (share >= 0.975 || (share >= 0.85 && /^\d/.test(t))));
        text += (glue ? '' : ' ') + t;
      });
      return text.replace(/\s+/g, ' ').trim();
    }

    function prove(list) {
    const warnings = [];
    let prev = null; let opening = null;
    const rows = [];
    const sums = { wd: 0, dep: 0, nWd: 0, nDep: 0, byKind: {} };
    for (const r of list) {
      if (r.special === 'opening') {
        if (r.bal === undefined) continue;
        if (prev === null) { prev = r.bal; opening = r.bal; }
        else if (r.bal !== prev) warnings.push('Page ' + r.page + ': the brought-forward balance printed there does not equal the balance of the line before it.');
        continue;
      }
      // cells of this row
      const byCluster = new Map();
      const extras = [];
      for (const t of r.extra) {
        const numeric = isMoney(t.s) || RE_SIGN.test(t.s) || (RE_FRAG.test(t.s) && /\d/.test(t.s));
        const c = numeric ? clusterOf(t) : null;
        if (c) { if (!byCluster.has(c)) byCluster.set(c, []); byCluster.get(c).push(t); }
        else extras.push(t);
      }
      const cells = []; let balCell = null; let badCell = null;
      for (const [c, fr] of byCluster) {
        const v = readCell(fr);
        if (!v) continue;
        if (v.bad) { badCell = v.raw; continue; }
        if (c === balCluster) balCell = v.paise; else if (v.paise !== 0) cells.push({ kind: c.kind, paise: Math.abs(v.paise) });
      }
      const narration = joinLines(r.lines);
      // a dated "B/F" line is an opening balance, not a transaction
      if (RE_OPENING.test(narration) && !cells.length && balCell !== null) {
        if (prev === null) { prev = balCell; opening = balCell; }
        else if (balCell !== prev) warnings.push('Page ' + r.page + ': the brought-forward balance printed there does not equal the balance of the line before it.');
        continue;
      }
      const row = { date: r.date, valueDate: null, narration, lines: r.lines.map(l => l.text.trim()), chq: '', other: '', page: r.page,
        withdrawal: 0, deposit: 0, balance: balCell, ok: true, note: '' };
      // value date / cheque no / other columns
      extras.sort((a, b) => a.y - b.y || a.x0 - b.x0);
      for (const t of extras) {
        if (inOtherDateCol(t) && !row.valueDate) { row.valueDate = parseDate(t.s); continue; }
        let lab = null, bd = 45;
        const cx = (t.x0 + t.x1) / 2;
        for (const l of sideLabels) { const d = cx < l.x0 ? l.x0 - cx : cx > l.x1 ? cx - l.x1 : 0; if (d < bd) { bd = d; lab = l; } }
        if (lab && lab.kind === 'chq') row.chq += (row.chq ? ' ' : '') + t.s;
        else row.other += (row.other ? ' ' : '') + t.s;
      }
      // the proof
      const prefer = ['debit', 'credit', 'sweepout', 'sweepin', 'other'];
      cells.sort((a, b) => prefer.indexOf(a.kind) - prefer.indexOf(b.kind));
      let delta = null;
      if (balCell === null) {
        const dcell = cells.find(c => c.kind === 'debit' || c.kind === 'sweepout'), ccell = cells.find(c => c.kind === 'credit' || c.kind === 'sweepin');
        if (prev !== null && (dcell || ccell) && !(dcell && ccell)) {
          delta = dcell ? -dcell.paise : ccell.paise; row.balance = prev + delta;
          row.ok = false; row.note = 'Balance not readable on this line; worked out from the amount.';
        } else { row.ok = false; row.note = 'Balance not readable on this line' + (badCell ? ' ("' + badCell + '")' : '') + '.'; }
      } else if (prev === null) {
        // no opening balance was printed: work it out from the first line
        const dcell = cells.find(c => c.kind === 'debit' || c.kind === 'sweepout'), ccell = cells.find(c => c.kind === 'credit' || c.kind === 'sweepin');
        if ((dcell || ccell) && !(dcell && ccell)) { delta = dcell ? -dcell.paise : ccell.paise; opening = balCell - delta; }
        else { row.ok = false; row.note = 'No opening balance printed and the first amount is unclear.'; opening = balCell; delta = 0; }
      } else {
        delta = balCell - prev;
        const hit = cells.find(c => c.paise === Math.abs(delta));
        if (delta === 0 && !cells.length) { /* a nil line */ }
        else if (!hit) {
          row.ok = false;
          row.note = cells.length ? 'Printed amount ' + cells.map(c => fmt(c.paise)).join(' / ') + ' does not equal the change in balance ' + fmt(Math.abs(delta)) + '.'
            : 'No amount could be read on this line; the change in balance is ' + fmt(Math.abs(delta)) + '.' + (badCell ? ' Unclear figure: "' + badCell + '".' : '');
        } else {
          const debitSide = hit.kind === 'debit' || hit.kind === 'sweepout', creditSide = hit.kind === 'credit' || hit.kind === 'sweepin';
          if ((debitSide && delta > 0) || (creditSide && delta < 0)) { row.ok = false; row.note = 'The amount is printed in the ' + (debitSide ? 'withdrawal' : 'deposit') + ' column but the balance moved the other way.'; }
          row.colKind = hit.kind;
        }
      }
      if (delta !== null) {
        if (delta < 0) { row.withdrawal = -delta; sums.wd += -delta; sums.nWd++; } else if (delta > 0) { row.deposit = delta; sums.dep += delta; sums.nDep++; }
        if (row.colKind) sums.byKind[row.colKind] = (sums.byKind[row.colKind] || 0) + Math.abs(delta);
      }
      if (row.balance !== null) prev = row.balance;
      rows.push(row);
    }
    return { rows, opening, closing: prev, sums, warnings, bad: rows.filter(x => !x.ok).length };
    }
    if (opts._reverseForTest) rowsRaw.reverse();
    let proved = prove(rowsRaw);
    if (proved.rows.length >= 3 && proved.bad > 0.3 * proved.rows.length) {
      // some banks print the newest transaction first: try reading from the bottom up
      const back = prove(rowsRaw.slice().reverse());
      if (back.bad < 0.1 * proved.bad + 1) { proved = back; out.layout.order = 'newest first (re-ordered oldest first)'; }
    }
    const rows = proved.rows, opening = proved.opening, prev = proved.closing, sums = proved.sums;
    proved.warnings.forEach(w => warnings.push(w));
    function fmt(p) { const neg = p < 0; p = Math.abs(p); const w = Math.floor(p / 100), f = String(p % 100).padStart(2, '0'); return (neg ? '-' : '') + w.toLocaleString('en-IN') + '.' + f; }

    out.rows = rows; out.opening = opening; out.closing = prev;
    out.totals = { withdrawals: sums.wd, deposits: sums.dep, nWithdrawals: sums.nWd, nDeposits: sums.nDep, byColumn: sums.byKind };

    // ---- 7. checks
    const bad = rows.filter(r => !r.ok);
    out.checks.push({ name: 'Every line agrees with the running balance', ok: bad.length === 0 && rows.length > 0,
      detail: rows.length ? (bad.length ? bad.length + ' of ' + rows.length + ' lines need a look (first: page ' + bad[0].page + ', ' + bad[0].date + ').' : 'All ' + rows.length + ' lines checked.') : 'No lines read.' });
    if (opening !== null && prev !== null)
      out.checks.push({ name: 'Opening + deposits - withdrawals = closing', ok: opening + sums.dep - sums.wd === prev, detail: fmt(opening) + ' + ' + fmt(sums.dep) + ' - ' + fmt(sums.wd) + ' = ' + fmt(prev) });

    // totals printed by the bank
    out.printed = { closing: printedClosing, totals: printedTotals };
    if (printedTotals.length) {
      const res = {};
      for (const side of ['debit', 'credit']) {
        const vals = printedTotals.filter(t => t.cells[side] !== undefined).map(t => t.cells[side]);
        if (!vals.length) continue;
        const sumAll = vals.reduce((a, b) => a + b, 0), last = vals[vals.length - 1];
        const all = side === 'debit' ? sums.wd : sums.dep, prim = sums.byKind[side] || 0;
        res[side] = { printed: (sumAll === all || sumAll === prim) ? sumAll : last, ok: [all, prim].includes(sumAll) || [all, prim].includes(last), how: (sumAll === all || sumAll === prim) && vals.length > 1 ? 'sum of page totals' : 'total line' };
      }
      out.printed.totalCheck = res;
      for (const side of Object.keys(res)) out.checks.push({ name: "Bank's printed total of " + (side === 'debit' ? 'withdrawals' : 'deposits'), ok: res[side].ok,
        detail: 'Printed ' + fmt(res[side].printed) + ' (' + res[side].how + '); read ' + fmt(side === 'debit' ? ((sums.byKind.debit === res[side].printed) ? sums.byKind.debit : sums.wd) : ((sums.byKind.credit === res[side].printed) ? sums.byKind.credit : sums.dep)) + '.' });
    }
    if (printedClosing !== null) out.checks.push({ name: "Bank's printed closing balance", ok: printedClosing === prev, detail: 'Printed ' + fmt(printedClosing) + '; read ' + fmt(prev) + '.' });

    out.fmt = fmt;
    return out;
  }

  // ---------------------------------------------------------------- who / which account / which period
  const IFSC_BANK = { ICIC: 'ICICI Bank', SBIN: 'State Bank of India', UTIB: 'Axis Bank', HDFC: 'HDFC Bank', PUNB: 'Punjab National Bank', CNRB: 'Canara Bank',
    IDIB: 'Indian Bank', UCBA: 'UCO Bank', BARB: 'Bank of Baroda', UBIN: 'Union Bank of India', BKID: 'Bank of India', IOBA: 'Indian Overseas Bank',
    KKBK: 'Kotak Mahindra Bank', YESB: 'Yes Bank', IBKL: 'IDBI Bank', CBIN: 'Central Bank of India', MAHB: 'Bank of Maharashtra', INDB: 'IndusInd Bank', FDRL: 'Federal Bank' };
  const BANK_NAMES = [[/ICICI\s*BANK/i, 'ICICI Bank'], [/STATE BANK OF INDIA/i, 'State Bank of India'], [/AXIS\s*BANK/i, 'Axis Bank'], [/HDFC\s*BANK/i, 'HDFC Bank'],
    [/PUNJAB NATIONAL BANK/i, 'Punjab National Bank'], [/CANARA BANK/i, 'Canara Bank'], [/INDIAN OVERSEAS BANK/i, 'Indian Overseas Bank'], [/INDIAN BANK/i, 'Indian Bank'],
    [/UCO BANK/i, 'UCO Bank'], [/BANK OF BARODA/i, 'Bank of Baroda'], [/UNION BANK/i, 'Union Bank of India'], [/BANK OF INDIA/i, 'Bank of India'],
    [/KOTAK/i, 'Kotak Mahindra Bank'], [/YES BANK/i, 'Yes Bank'], [/IDBI/i, 'IDBI Bank'], [/CENTRAL BANK/i, 'Central Bank of India']];

  /** Join a page's tokens into reading-order text lines. */
  function textLines(page) {
    const toks = page.tokens.slice().sort((a, b) => a.y - b.y || a.x0 - b.x0);
    const lines = []; let cur = null;
    for (const t of toks) {
      if (cur && Math.abs(t.y - cur.y) <= 2.2) { cur.parts.push(t); }
      else { cur = { y: t.y, parts: [t] }; lines.push(cur); }
    }
    return lines.map(l => { l.parts.sort((a, b) => a.x0 - b.x0); return { y: l.y, text: l.parts.map(p => p.s).join('  ') }; });
  }

  function readMeta(pages) {
    const meta = { bank: '', ifsc: '', account: '', accountMasked: false, from: null, to: null, holder: '' };
    const all = [];
    pages.forEach(p => textLines(p).forEach(l => all.push({ page: p.no, y: l.y, text: l.text })));
    // IFSC: only where it is labelled, or in the letterhead of page 1 (narrations are full of other banks' codes)
    for (const l of all) {
      const m = l.text.match(/IFS\s*C?\s*(?:CODE)?\s*:?\s*([A-Z]{4}0[A-Z0-9]{6})\b/i);
      if (m) { meta.ifsc = m[1].toUpperCase(); break; }
    }
    if (!meta.ifsc) for (const l of all) {
      if (!/IFS/i.test(l.text)) continue;
      const idx = all.indexOf(l);
      for (let k = idx; k < Math.min(all.length, idx + 4); k++) { const m = all[k].text.match(/\b([A-Z]{4}0[A-Z0-9]{6})\b/); if (m) { meta.ifsc = m[1]; break; } }
      if (meta.ifsc) break;
    }
    const head = all.filter(l => l.page === 1).slice(0, 40).map(l => l.text).join('\n');
    for (const [re, name] of BANK_NAMES) if (re.test(head)) { meta.bank = name; break; }
    if (!meta.bank && meta.ifsc && IFSC_BANK[meta.ifsc.slice(0, 4)]) meta.bank = IFSC_BANK[meta.ifsc.slice(0, 4)];
    if (!meta.bank) { const whole = all.map(l => l.text).join('\n'); for (const [re, name] of BANK_NAMES) if (re.test(whole)) { meta.bank = name; break; } }

    // account number: prefer a full number over a masked one
    const cands = [];
    const res = [/Statement of Account No\.?\s*:?\s*([X\d]{8,18})/i, /Account\s*(?:No|Number)\.?\s*:?\s*([X\d]{8,18})/i, /Account Number\s*:\s*([X\d]{8,18})/i,
      /(?:Current|Savings|Saving|CC|OD)\s*A\/c\s+([X\d]{8,18})/i, /in (?:Current|Savings)? ?Account(?: Number| No\.?)?\s*:?\s*([X\d]{8,18})/i,
      /Account number:\s*([X\d]{8,18})/i, /^\s*(?:Current|Savings)\s+(\d{9,18})\s/i];
    for (const l of all) for (const re of res) { const m = l.text.replace(/\s{2,}/g, ' ').match(re); if (m) cands.push(m[1]); }
    const full = cands.find(c => /^\d+$/.test(c));
    meta.account = full || cands[0] || '';
    meta.accountMasked = !!meta.account && !/^\d+$/.test(meta.account);
    const tail = x => String(x).replace(/\D/g, '').slice(-4);
    meta.otherAccounts = Array.from(new Set(cands.filter(c => tail(c) && tail(c) !== tail(meta.account))));

    // period
    const DATE_ANY = /(\d{1,2}[-\/.]\d{1,2}[-\/.]\d{2,4}|[A-Za-z]{3,9}\.? \d{1,2}, ?\d{4}|\d{1,2}[- ][A-Za-z]{3,9}[- ,]+\d{2,4})/g;
    for (const l of all) {
      if (!/period|statement from|from\s*:|between/i.test(l.text)) continue;
      const ds = (l.text.replace(DASH, '-').replace(/\s{2,}/g, ' ').match(DATE_ANY) || []).map(parseDate).filter(Boolean);
      if (ds.length >= 2) { meta.from = ds[0]; meta.to = ds[1]; break; }
    }
    return meta;
  }

  /** Figures the bank prints in its own summary boxes (not in the transaction table). */
  function readPrintedSummary(pages) {
    const out = {};
    const all = [];
    pages.forEach(p => textLines(p).forEach(l => all.push(l.text.replace(/\s{2,}/g, ' ').trim())));
    const M = '(\\d[\\d,]*\\.\\d{2})';
    const val = (d, sfx, pre) => { const p = toPaise(d); return p === null ? null : ((sfx && /dr/i.test(sfx)) || pre ? -p : p); };
    // a summary box with headings and the figures under them (headings may sit a few points apart)
    for (const p of pages) {
      const heads = {};
      for (const t of p.tokens) {
        const n = norm(t.s);
        const k = /^dr\.? ?count$/.test(n) ? 'nDebits' : /^cr\.? ?count$/.test(n) ? 'nCredits' : /^total debits?$/.test(n) ? 'totalDebits'
          : /^total credits?$/.test(n) ? 'totalCredits' : /^closing balance$/.test(n) ? 'closing' : /^brought forward$|^opening balance$/.test(n) ? 'opening' : null;
        if (k) heads[k] = t;
      }
      if (!(heads.totalDebits && heads.totalCredits)) continue;
      for (const k of Object.keys(heads)) {
        const h = heads[k]; let best = null, bd = 1e9;
        for (const t of p.tokens) {
          if (t.y <= h.y + 4 || t.y > h.y + 30) continue;
          if (t.x1 < h.x0 - 12 || t.x0 > h.x1 + 12) continue;
          const d = (t.y - h.y) + Math.abs((t.x0 + t.x1) / 2 - (h.x0 + h.x1) / 2) * 0.2;
          if (d < bd) { bd = d; best = t; }
        }
        if (!best) continue;
        const m = best.s.replace(/\s+/g, '').match(/^(\d[\d,]*\.\d{2})(CR|DR)?$/i);
        if (k === 'nDebits' || k === 'nCredits') { if (/^\d+$/.test(best.s.trim())) out[k] = +best.s.trim(); }
        else if (m) out[k] = val(m[1], m[2]);
      }
      if (out.totalDebits === undefined) { delete out.nDebits; delete out.nCredits; }
    }
    for (let i = 0; i < all.length; i++) {
      if (out.accountBalance === undefined) {
        const m = all[i].match(new RegExp('^(?:Current|Savings|Saving)(?: A\\/c)? [X\\d]{8,18} (\\(-\\) ?|-)?' + M + ' ?(Cr|Dr)?\\b', 'i'));
        if (m) out.accountBalance = val(m[2], m[3], m[1]);
      }
    }
    return out;
  }

  function read(pages, opts) {
    pages = unmerge(pages);
    const r = parseStatement(pages, opts);
    r.meta = readMeta(pages);
    if (r.rows && r.rows.length && r.fmt) {
      const ps = readPrintedSummary(pages); const f = r.fmt;
      r.printed.summary = ps;
      if (ps.totalDebits !== undefined) {
        if (ps.nDebits !== undefined && ps.nCredits !== undefined) r.checks.push({ name: "Bank's summary: number of withdrawals and deposits", ok: ps.nDebits === r.totals.nWithdrawals && ps.nCredits === r.totals.nDeposits,
          detail: 'Printed ' + ps.nDebits + ' and ' + ps.nCredits + '; read ' + r.totals.nWithdrawals + ' and ' + r.totals.nDeposits + '.' });
        r.checks.push({ name: "Bank's summary: total withdrawals and deposits", ok: ps.totalDebits === r.totals.withdrawals && ps.totalCredits === r.totals.deposits,
          detail: 'Printed ' + f(ps.totalDebits) + ' and ' + f(ps.totalCredits) + '; read ' + f(r.totals.withdrawals) + ' and ' + f(r.totals.deposits) + '.' });
        if (ps.closing !== undefined && ps.closing !== null && r.printed.closing === null) r.checks.push({ name: "Bank's printed closing balance", ok: ps.closing === r.closing, detail: 'Printed ' + f(ps.closing) + '; read ' + f(r.closing) + '.' });
      }
      if (ps.accountBalance !== undefined && ps.accountBalance !== null && r.printed.closing === null && ps.closing === undefined)
        r.checks.push({ name: "Account balance printed on the statement's first page", ok: ps.accountBalance === r.closing, detail: 'Printed ' + f(ps.accountBalance) + '; closing balance read ' + f(r.closing) + '.' });
    }
    if (r.rows && r.rows.length) {
      const ds = r.rows.map(x => x.date).filter(Boolean).sort();
      r.meta.firstDate = ds[0]; r.meta.lastDate = ds[ds.length - 1];
      if (!r.meta.from) { r.meta.from = ds[0]; r.meta.to = ds[ds.length - 1]; r.meta.periodFromLines = true; }
    }
    if (r.rows && r.rows.length && r.meta.to && !r.meta.periodFromLines) {
      const gap = (Date.parse(r.meta.to) - Date.parse(r.meta.lastDate)) / 86400000, gap0 = (Date.parse(r.meta.firstDate) - Date.parse(r.meta.from)) / 86400000;
      const nd = x => x.slice(8, 10) + '-' + x.slice(5, 7) + '-' + x.slice(0, 4);
      if (gap > 10) r.warnings.push('The heading says the statement runs to ' + nd(r.meta.to) + ', but its last transaction is on ' + nd(r.meta.lastDate) + '. If the account had transactions after that date, this file is incomplete.');
      if (gap0 > 10) r.warnings.push('The heading says the statement starts on ' + nd(r.meta.from) + ', but its first transaction is on ' + nd(r.meta.firstDate) + '.');
    }
    if (r.meta.otherAccounts && r.meta.otherAccounts.length && r.rows && r.rows.some(x => !x.ok))
      r.warnings.push('More than one account number appears in this file (' + [r.meta.account].concat(r.meta.otherAccounts).join(', ') + '). A statement holding several accounts cannot be read as one; please use one PDF per account.');
    r.pass = r.checks.length > 0 && r.checks.every(c => c.ok);
    return r;
  }

  return { read, parseStatement, readMeta, readPrintedSummary, parseDate, toPaise, textLines };
});
