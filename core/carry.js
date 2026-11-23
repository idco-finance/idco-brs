/* IDCO Bank Statement to Excel - carry forward.
 * The BRS workbook keeps, in a hidden sheet, the inputs that do not change once a month is over: the bank lines
 * (already proved against the running balance), the Eazypay payments and the opening position of the last signed BRS.
 * Next month only the new statement, the new Eazypay reports and the Bank Book from 1 April are needed.
 * Only inputs are carried, never results: everything is matched again from 1 April each time. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementCarry = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const FORMAT = 1, SHEET = 'Carry forward', CHUNK = 30000;
  // statements read by a tool version below this one are not carried (set when a release corrects how statements are read)
  const REREAD_BELOW = '0.0.0';

  const nice = iso => (iso ? iso.slice(8, 10) + '-' + iso.slice(5, 7) + '-' + iso.slice(0, 4) : '');
  const fyStartOf = iso => (+iso.slice(5, 7) >= 4 ? iso.slice(0, 4) : String(+iso.slice(0, 4) - 1)) + '-04-01';
  const last4 = s => String(s || '').replace(/\D/g, '').slice(-4);
  const verNum = v => { const m = String(v || '').match(/(\d+)\.(\d+)\.(\d+)/); return m ? +m[1] * 1e6 + +m[2] * 1e3 + +m[3] : 0; };
  function fmt(p) { const neg = p < 0; p = Math.abs(p); return (neg ? '-' : '') + Math.floor(p / 100).toLocaleString('en-IN') + '.' + String(p % 100).padStart(2, '0'); }

  // a check value over the carried text (cyrb53): notices a cell that was changed, cut or lost
  function hash(str) {
    let h1 = 0xdeadbeef ^ str.length, h2 = 0x41c6ce57 ^ str.length;
    for (let i = 0; i < str.length; i++) { const ch = str.charCodeAt(i); h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677); }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507); h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507); h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (h2 >>> 0).toString(16).padStart(8, '0') + (h1 >>> 0).toString(16).padStart(8, '0');
  }
  // plain ASCII text that a spreadsheet program has no reason to touch (amounts stay whole paise, dates stay yyyy-mm-dd)
  const enc = obj => JSON.stringify(obj).replace(/[\u007f-￿]/g, c => '\\u' + ('0000' + c.charCodeAt(0).toString(16)).slice(-4)).replace(/_x/g, '_\\u0078');

  /** What to carry, from a finished run. rec.src = { acc, mis, op }. Null when there is nothing safe to carry. */
  function pack(rec, version, when) {
    const src = rec.src || {}, acc = src.acc, op = src.op;
    if (!acc || !op || !acc.pass || acc.gap || !acc.rows || !acc.rows.length || !acc.files || !op.items) return null;
    const asOn = rec.brs && rec.brs.asOn ? rec.brs.asOn : '';
    const old = acc.carry || null;
    const files = acc.files.map(f => Object.assign({}, f, { readBy: f.readBy || version, firstUsed: f.firstUsed || asOn }));
    files.forEach(f => { delete f.carried; });
    const hasEz = rec.bankLines.some(b => b.type === 'Eazypay settlement') || rec.payments.some(m => m.bankSl !== null && m.bankSl !== undefined);
    let ez = null;
    if (hasEz) {
      const seen = new Set(), rows = [], info = [];
      for (const f of src.mis || []) {
        let n = 0, from = '', to = '';
        for (const r of f.rows) {
          if (seen.has(r.tran)) continue; seen.add(r.tran); rows.push(Object.assign({}, r, { file: f.fileName })); n++;
          if (r.settle) { if (!from || r.settle < from) from = r.settle; if (!to || r.settle > to) to = r.settle; }
        }
        const was = old && old.ez ? old.ez.files.find(x => x.fileName === f.fileName) : null;
        info.push({ fileName: f.fileName, rows: n, from, to, firstSettle: f.firstSettle || from, firstUsed: was ? was.firstUsed : asOn });
      }
      ez = { files: info, rows };
    }
    const bookTo = rec.summary.lastBookFile || rec.summary.lastBook || '';
    return { format: FORMAT, tool: version, made: (when || new Date()).toISOString(), bank: acc.bank, account: acc.account, ifsc: acc.ifsc || '', asOn, fy: fyStartOf(acc.from || acc.firstDate),
      st: { from: acc.from, to: acc.to, firstDate: acc.firstDate, lastDate: acc.lastDate, opening: acc.opening, closing: acc.closing, files, rows: acc.rows.map(r => { const o = Object.assign({}, r); delete o.gapBefore; return o; }) },
      ez, op: { fileName: op.fileName, account: op.account, gl: op.gl, name: op.name, asOn: op.asOn, book: op.book, bank: op.bank, heads: op.heads, items: op.items, checks: op.checks, sheet: op.sheet },
      bookTo, v: rec.vouchers.map(v => [v.blockId, v.vchDate, v.net, v.instrDate || '', v.provNo || '', v.party || '']) };
  }

  /** Writes the hidden sheet. */
  function write(wb, p) {
    if (!p) return false;
    const text = enc(p);
    const ws = wb.addWorksheet(SHEET); ws.state = 'veryHidden';
    ws.getColumn(1).width = 60;
    ws.addRow(['IDCO BRS carry forward - kept by the tool for next month. Please do not change this sheet.']);
    ws.addRow(['format', String(FORMAT)]); ws.addRow(['check', hash(text)]); ws.addRow(['length', String(text.length)]);
    const n = Math.ceil(text.length / CHUNK); ws.addRow(['parts', String(n)]);
    for (let i = 0; i < n; i++) ws.addRow(['#' + text.slice(i * CHUNK, (i + 1) * CHUNK) + '#']);
    return true;
  }

  const cell = v => (v === null || v === undefined ? '' : typeof v === 'object' ? (v.text !== undefined ? String(v.text) : v.richText ? v.richText.map(x => x.text).join('') : v.result !== undefined ? String(v.result) : '') : String(v));

  /** Reads the carried data of a BRS workbook. null = the workbook has none; { problem } = it cannot be trusted; else the data. */
  function read(wb, fileName) {
    const ws = wb.getWorksheet(SHEET); if (!ws) return null;
    const bad = why => ({ problem: 'The data carried in ' + fileName + ' cannot be used: ' + why + ' Drop the original bank statements and Eazypay reports from 1 April instead; the decisions in the file are still read.' });
    try {
      const get = r => cell(ws.getRow(r).getCell(2).value).trim();
      if (get(2) !== String(FORMAT)) return bad('it was written by a newer version of the tool.');
      const n = +get(5), len = +get(4); let text = '';
      for (let i = 0; i < n; i++) { const c = cell(ws.getRow(6 + i).getCell(1).value); if (c.charAt(0) !== '#' || c.charAt(c.length - 1) !== '#') return bad('part of it is missing or was changed.'); text += c.slice(1, -1); }
      if (text.length !== len || hash(text) !== get(3)) return bad('it was changed after the tool wrote it.');
      const p = JSON.parse(text);
      if (!p.st || !p.st.rows || !p.st.rows.length || !p.op) return bad('it is incomplete.');
      const stale = p.st.files.filter(f => verNum(f.readBy) < verNum(REREAD_BELOW));
      if (stale.length) return bad('its statements were read by an older version of the tool (v' + stale[0].readBy + '), and the way statements are read has been corrected since.');
      p.workbook = fileName;
      return p;
    } catch (e) { return bad('it could not be read.'); }
  }

  const pseudoName = c => 'Statements carried in ' + c.workbook;
  /** The carried bank lines as one statement, to be joined with the new ones. */
  function statementOf(c) {
    return { fileName: pseudoName(c), carry: c, meta: { bank: c.bank, account: c.account, ifsc: c.ifsc, from: c.st.from, to: c.st.to, firstDate: c.st.firstDate, lastDate: c.st.lastDate },
      opening: c.st.opening, closing: c.st.closing, rows: c.st.rows.map(r => Object.assign({}, r, { _file: r.file })), pass: true, checks: [], warnings: [], totals: null };
  }
  /** Of several BRS workbooks of one account, the one that carries the most recent lines. */
  function choose(carries) {
    const by = new Map();
    for (const c of carries) { const k = c.bank + '|' + last4(c.account); const o = by.get(k); if (!o || c.st.lastDate > o.st.lastDate || (c.st.lastDate === o.st.lastDate && c.made > o.made)) by.set(k, c); }
    return Array.from(by.values());
  }
  /** After StatementMerge.groupAndJoin: puts the carried statements back under their own names. statements = the real ones read this time. */
  function afterJoin(accounts, statements) {
    for (const a of accounts) {
      const i = a.files.findIndex(f => /^Statements carried in /.test(f.fileName)); if (i < 0) continue;
      const c = carriesOf(a, statements); if (!c) continue;
      a.carry = c;
      const info = a.files[i];
      a.files.splice.apply(a.files, [i, 1].concat(c.st.files.map(f => Object.assign({}, f, { carried: c.workbook }))));
      for (const r of a.rows) if (r._file !== undefined) { r.file = r._file; delete r._file; }
      a.carryNote = c.st.files.length + ' statement(s), ' + c.st.rows.length + ' lines up to ' + nice(c.st.lastDate) + ', carried in ' + c.workbook + (info.used < info.lines ? ' (' + (info.lines - info.used) + ' of them were also in a statement dropped now)' : '') + '.';
      a.joins.unshift({ ok: true, text: a.carryNote });
      // original files dropped together with the workbook: do the carried lines agree with them?
      const at = new Map(); c.st.rows.forEach(r => { const k = r.date + '|' + r.balance + '|' + r.deposit + '|' + r.withdrawal; if (!at.has(k)) at.set(k, r); });
      let same = 0, differ = 0; const names = [];
      for (const st of statements) {
        if (st.carry || (st.meta.bank || '') !== c.bank || last4(st.meta.account) !== last4(c.account)) continue;
        let hit = 0;
        for (const r of st.rows) { const o = at.get(r.date + '|' + r.balance + '|' + r.deposit + '|' + r.withdrawal); if (!o) continue; hit++; if (o.narration === r.narration) same++; else differ++; }
        if (hit) names.push(st.fileName);
      }
      a.carryVerify = same + differ ? { same, differ, files: names } : null;
      if (differ) a.joins.push({ ok: false, text: differ + ' carried line(s) have a different narration in the statement file(s) dropped now (' + names.join(', ') + '). Drop all the original statements from 1 April without the carried workbook.' });
      else if (same) a.joins.push({ ok: true, text: same + ' carried line(s) were checked against the original statement(s) dropped now (' + names.join(', ') + '): all agree.' });
      if (differ) { a.pass = false; }
    }
    return accounts;
  }
  function carriesOf(a, statements) { for (const st of statements) if (st.carry && st.meta.bank === a.bank && last4(st.meta.account) === last4(a.account)) return st.carry; return null; }

  /** Eazypay: the carried payments plus the reports dropped now. A payment is counted once. */
  function mergeMis(c, newMis) {
    if (!c || !c.ez) return { mis: newMis, problems: [], repeated: 0 };
    const have = new Map(); c.ez.rows.forEach(r => have.set(r.tran, r));
    const files = c.ez.files.map(f => ({ fileName: f.fileName, carried: c.workbook, firstUsed: f.firstUsed, firstSettle: f.firstSettle || '', rows: [] }));
    const byName = new Map(files.map(f => [f.fileName, f]));
    for (const r of c.ez.rows) { let f = byName.get(r.file); if (!f) { f = { fileName: r.file, carried: c.workbook, rows: [] }; byName.set(r.file, f); files.push(f); } f.rows.push(r); }
    let repeated = 0; const clash = [];
    for (const f of newMis) for (const r of f.rows) { const o = have.get(r.tran); if (!o) continue; repeated++; if (o.amt !== r.amt || o.settle !== r.settle) clash.push(r.tran + ' (' + fmt(o.amt) + ' settled ' + nice(o.settle) + ' in the carried data; ' + fmt(r.amt) + ' settled ' + nice(r.settle) + ' in ' + f.fileName + ')'); }
    const problems = clash.length ? [clash.length + ' Eazypay payment(s) are in the carried data and in a report dropped now with a different amount or settlement date: ' + clash.slice(0, 5).join('; ') + (clash.length > 5 ? '; …' : '') + '. Check the reports; then drop all Eazypay reports from 1 April without the carried workbook.'] : [];
    return { mis: files.concat(newMis.filter(f => !byName.has(f.fileName))), problems, repeated };
  }

  /** What changed in the books, for vouchers dated up to the end of the Bank Book used last time. */
  function bookChanges(c, rec) {
    if (!c || !c.v) return null;
    const nowTo = rec.summary.lastBookFile || rec.summary.lastBook || '';
    const upto = !c.bookTo || (nowTo && nowTo < c.bookTo) ? nowTo : c.bookTo;
    const old = new Map(); c.v.forEach(x => { if (x[1] <= upto) old.set(x[0] + '|' + x[1], x); });
    const cur = new Map(); rec.vouchers.forEach(v => { if (v.vchDate <= upto) cur.set(v.blockId + '|' + v.vchDate, [v.blockId, v.vchDate, v.net, v.instrDate || '', v.provNo || '', v.party || '']); });
    const out = []; const row = (what, x, was, now) => out.push({ what, voucher: x[0], date: x[1], party: x[5], amt: x[2], was: was || '', now: now || '' });
    const gone = [], came = [];
    old.forEach((o, k) => { const n = cur.get(k); if (!n) { gone.push(o); return; }
      if (o[2] !== n[2]) row('Amount changed', n, fmt(o[2]), fmt(n[2]));
      if (o[3] !== n[3]) row('Instrument date changed', n, nice(o[3]) || 'none', nice(n[3]) || 'none'); });
    cur.forEach((n, k) => { if (!old.has(k)) came.push(n); });
    for (const o of gone.slice()) {
      let j = came.findIndex(n => n[0] === o[0]);
      if (j >= 0) { row('Voucher date changed', came[j], nice(o[1]), nice(came[j][1])); if (o[2] !== came[j][2]) row('Amount changed', came[j], fmt(o[2]), fmt(came[j][2])); came.splice(j, 1); gone.splice(gone.indexOf(o), 1); continue; }
      if (/PROV/i.test(o[0]) && o[4]) { j = came.findIndex(n => n[4] === o[4] && n[2] === o[2]); if (j >= 0) { row('Voucher number given', came[j], o[0], came[j][0]); came.splice(j, 1); gone.splice(gone.indexOf(o), 1); } }
    }
    gone.forEach(o => row('No longer in the Bank Book', o, fmt(o[2]), 'not there'));
    came.forEach(n => row('New in the Bank Book (dated in a period already reconciled)', n, 'not there', fmt(n[2])));
    out.sort((a, b) => a.date.localeCompare(b.date) || a.voucher.localeCompare(b.voucher, undefined, { numeric: true }));
    return { upto, since: c.asOn || c.st.lastDate, workbook: c.workbook, rows: out };
  }

  return { FORMAT, SHEET, pack, write, read, statementOf, choose, afterJoin, mergeMis, bookChanges, hash, verNum };
});
