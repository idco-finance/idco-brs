/* IDCO Bank Statement to Excel - the reconciliation workbook (Step 1: review in Excel). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementReconBook = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const TEAL = 'FF0F766E', GREEN = 'FFD1FAE5', AMBER = 'FFFEF3C7', RED = 'FFFEE2E2', GREY = 'FFE5E7EB', BLUE = 'FFDBEAFE';
  const MONEY = '#,##0.00;[Red]-#,##0.00';
  const COLOUR = { 'Matched': GREEN, 'Suggested': AMBER, 'Needs review': RED, 'Not in books': RED, 'Not in bank': RED, 'Bank only': AMBER, 'Sweep': BLUE, 'Earlier period': GREY, 'After statement period': GREY,
    'Suspense entry': GREY, 'Cancelled receipt': GREY, 'No bank effect': GREY, 'Booked': GREEN, 'Booked in part': AMBER };
  const HOWTO = [
    '1. Open the "To review" sheet. Each row says what the tool found, what to do and where to look.',
    '2. In the Decision column choose from the list:  Accept = the suggested match is right.  Reject = it is wrong.  Keep open = looked at; it is a genuine open item.',
    '3. If you know the other side yourself, leave Decision empty and type it in "Match with": a voucher number (CRV-123), a bank line (line 45), or several separated by commas. Two bank lines that cancel each other can be matched with each other in the same way.',
    '4. Remark is free: write anything you want kept with the item.',
    '5. Save this file. Next time, drop it on the page together with the statements, the Bank Books and the Eazypay files. The decisions are applied again and listed in the "Decisions" sheet; only what is still undecided comes back in "To review".',
    '6. To take a decision back, choose Undo in the "Decisions" sheet (or clear the Decision on the Bank lines / Vouchers sheet), save, and drop the file in again.',
    'The tool applies what you decide and nothing more. It never books, changes or deletes anything in the ERP or at the bank.'
  ];
  const LIST_ALL = '"Accept,Reject,Keep open,Undo"';
  // drop-down lists are written as whole ranges (one rule per run of rows), which keeps the file small and valid
  function pick(cell, formula) { const ws = cell.worksheet; (ws._dv = ws._dv || []).push([cell.row, cell.col, formula]); cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFBEB' } }; }
  function writeLists(ws) {
    const list = (ws._dv || []).sort((a, b) => a[1] - b[1] || a[0] - b[0]); let i = 0;
    while (i < list.length) {
      let j = i; while (j + 1 < list.length && list[j + 1][1] === list[i][1] && list[j + 1][2] === list[i][2] && list[j + 1][0] === list[j][0] + 1) j++;
      const L = ws.getColumn(list[i][1]).letter;
      ws.dataValidations.model[L + list[i][0] + ':' + L + list[j][0]] = { type: 'list', allowBlank: true, formulae: [list[i][2]], showErrorMessage: true, errorTitle: 'Decision', error: 'Please choose from the list.' };
      i = j + 1;
    }
  }
  const input = cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFBEB' } }; };
  const MEANING = [
    ['Matched', 'Proved by a reference (Eazypay TranID or UTR) or by amount, date and party name together.'],
    ['Suggested', 'A likely match. Please confirm it.'],
    ['Needs review', 'Could not be settled by the tool. The Note says where to look and why.'],
    ['Not in books', 'In the bank statement; no voucher found in the Bank Book files given. (BRS: credited / debited by bank only.)'],
    ['Not in bank', 'In the Bank Book; no line found in the statements given. (BRS: deposited but not credited / issued but not presented.)'],
    ['Bank only', 'Entries made by the bank itself: interest and closure of linked deposits, chargebacks.'],
    ['Sweep', 'Money moved by the bank between the account and its linked deposits. Not in the Bank Book.'],
    ['Earlier period', 'A receipt dated before the first statement given. No bank line can be expected here.'],
    ['Suspense entry', 'A suspense booking or its reversal. It has no bank line of its own.'],
    ['Cancelled receipt', 'A money receipt and the voucher that cancels it. Neither has a bank line.']
  ];
  const d = iso => (iso ? new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10))) : null);
  const rs = p => (p === null || p === undefined ? null : p / 100);

  function sheet(wb, name, cols, opt) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = cols.map(c => ({ width: c[1] }));
    const h = ws.addRow(cols.map(c => c[0])); h.height = 22;
    h.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } }; c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; });
    cols.forEach((c, i) => { if (c[2] === 'm') ws.getColumn(i + 1).numFmt = MONEY; if (c[2] === 'd') { ws.getColumn(i + 1).numFmt = 'dd-mm-yyyy'; ws.getColumn(i + 1).alignment = { horizontal: 'center' }; } if (c[2] === 't') ws.getColumn(i + 1).numFmt = '@'; });
    h.eachCell(c => { c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; });
    ws._statusCol = cols.findIndex(c => c[0] === 'Status') + 1;
    ws._n = cols.length;
    return ws;
  }
  function add(ws, vals, status) {
    const r = ws.addRow(vals);
    if (ws._statusCol && status && COLOUR[status]) r.getCell(ws._statusCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOUR[status] } };
    return r;
  }
  const finish = ws => { writeLists(ws); ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, ws.rowCount), column: ws._n } }; };

  /** The items still waiting for a person: the rows of the To review sheet. */
  function pendingOf(rec) {
    const SHOW = new Set(['Suggested', 'Needs review', 'Not in books', 'Not in bank', 'Bank only']);
    const rankOf = { 'Needs review': 0, 'Suggested': 1, 'Not in books': 2, 'Not in bank': 3, 'Bank only': 4 };
    const pending = [];
    rec.bankLines.filter(b => SHOW.has(b.status) && b.found && !b.reviewed).forEach(b => pending.push({ x: b, st: b.status, date: b.date, v: ['Bank', d(b.date), b.party || b.type, rs(b.amt)] }));
    rec.vouchers.filter(v => SHOW.has(v.status) && v.found && !v.reviewed && !(v.status === 'Suggested' && v.bank.length)).forEach(v => pending.push({ x: v, st: v.status, date: v.vchDate, v: ['ERP books', d(v.vchDate), v.paidBy || v.party, rs(v.net)] }));
    rec.payments.filter(m => m.found && !m.reviewed).forEach(m => { const st = m.status === 'Booked in part' ? 'Needs review' : 'Not in books'; pending.push({ x: m, st, date: m.settle, v: ['Eazypay', d(m.settle), (m.partyCode ? m.partyCode + ' ' : '') + m.party, rs(m.amt)] }); });
    pending.sort((a, b) => (rankOf[a.st] - rankOf[b.st]) || a.date.localeCompare(b.date));
    return pending;
  }

  function build(ExcelJS, rec, when, version) {
    const wb = new ExcelJS.Workbook(); wb.creator = 'IDCO Bank Statement to Excel v' + version; wb.created = when || new Date();
    const s = rec.summary, fmt = rec.fmt, nice = rec.nice;
    const dc = s.decisions, describe = rec.describe || (k => k), linksOf = rec.linksOf || (() => []);
    const pending = pendingOf(rec);

    // ---------------- Summary
    const w0 = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
    w0.columns = [{ width: 44 }, { width: 14 }, { width: 24 }, { width: 24 }, { width: 70 }];
    const title = w0.addRow(['Bank reconciliation - matching sheet']); title.font = { bold: true, size: 14, color: { argb: TEAL } };
    w0.addRow([s.bank + '  A/c ' + s.account]).font = { bold: true };
    w0.addRow(['Bank statements: ' + nice(s.firstBank) + ' to ' + nice(s.lastBank) + ' (' + s.bankLines + ' lines).   Bank Book: vouchers dated ' + nice(s.firstBook) + ' to ' + nice(s.lastBook) + ' (' + s.vouchers + ' vouchers).']);
    w0.addRow(['This sheet only matches and lists. It is not the signed BRS. Every Suggested and Needs review item is for a person to decide.']).font = { italic: true, color: { argb: 'FF4B5563' } };
    w0.addRow([]);
    const sec = t => { const r = w0.addRow(t); r.font = { bold: true, color: { argb: 'FFFFFFFF' } }; for (let i = 1; i <= 5; i++) r.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } }; };
    const order = ['Matched', 'Suggested', 'Needs review', 'Not in books', 'Not in bank', 'Bank only', 'Sweep', 'Earlier period', 'After statement period', 'Suspense entry', 'Cancelled receipt', 'No bank effect'];
    const mean = Object.fromEntries(MEANING);
    sec(['Bank statement lines', 'Lines', 'Deposits', 'Withdrawals', 'What it means']);
    const bStart = w0.rowCount + 1;
    for (const k of order) if (s.bankBy[k]) { const v = s.bankBy[k]; const r = w0.addRow([k, v.n, rs(v.dep), rs(v.wd), mean[k] || '']); r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOUR[k] || GREY } }; r.getCell(5).alignment = { wrapText: true, vertical: 'top' }; }
    const bEnd = w0.rowCount;
    const bt = w0.addRow(['Total', { formula: 'SUM(B' + bStart + ':B' + bEnd + ')', result: s.bankLines }, { formula: 'SUM(C' + bStart + ':C' + bEnd + ')', result: rs(Object.values(s.bankBy).reduce((a, v) => a + v.dep, 0)) }, { formula: 'SUM(D' + bStart + ':D' + bEnd + ')', result: rs(Object.values(s.bankBy).reduce((a, v) => a + v.wd, 0)) }]); bt.font = { bold: true };
    w0.addRow([]);
    sec(['Bank Book vouchers', 'Vouchers', 'Receipts (bank debit)', 'Payments (bank credit)', 'What it means']);
    const vStart = w0.rowCount + 1;
    for (const k of order) if (s.vBy[k]) { const v = s.vBy[k]; const r = w0.addRow([k, v.n, rs(v.dr), rs(v.cr), mean[k] || '']); r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOUR[k] || GREY } }; r.getCell(5).alignment = { wrapText: true, vertical: 'top' }; }
    const vEnd = w0.rowCount;
    const vt = w0.addRow(['Total', { formula: 'SUM(B' + vStart + ':B' + vEnd + ')', result: s.vouchers }, { formula: 'SUM(C' + vStart + ':C' + vEnd + ')', result: rs(Object.values(s.vBy).reduce((a, v) => a + v.dr, 0)) }, { formula: 'SUM(D' + vStart + ':D' + vEnd + ')', result: rs(Object.values(s.vBy).reduce((a, v) => a + v.cr, 0)) }]); vt.font = { bold: true };
    w0.addRow([]);
    sec(['Eazypay and sweeps', '', 'Amount', '', '']);
    w0.addRow(['Eazypay payments inside the settlement credits of the statement', s.eazypay.inStatement]);
    w0.addRow(['   of these, not in the Bank Book yet', s.eazypay.notInBooks, rs(s.eazypay.notInBooksAmt), '', 'Listed in the Eazypay payments sheet, for the section that books them.']);
    w0.addRow(['   of these, only partly booked', s.eazypay.partly]);
    w0.addRow(['Swept to linked deposits', '', rs(s.sweepOut)]);
    w0.addRow(['Swept back from linked deposits', '', rs(s.sweepIn)]);
    w0.addRow(['Net amount moved to linked deposits in this period', '', { formula: 'C' + (w0.rowCount - 1) + '-C' + w0.rowCount, result: rs(s.sweepOut - s.sweepIn) }, '', 'A negative figure means more came back than went out.']);
    w0.addRow([]);
    sec(['Files used', 'Vouchers / rows', 'Period', 'Its own check', '']);
    for (const f of s.bookFiles) w0.addRow(['Bank Book: ' + f.fileName, f.blocks, nice(f.from) + ' to ' + nice(f.to), f.tie || '']);
    for (const f of s.misFiles) w0.addRow(['Eazypay: ' + f.fileName, f.rows]);
    if (s.notes.length) { w0.addRow([]); sec(['Notes', '', '', '', '']); for (const n of s.notes) { const r = w0.addRow([n]); w0.mergeCells(r.number, 1, r.number, 5); r.getCell(1).alignment = { wrapText: true, vertical: 'top' }; r.height = 30; r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } }; } }
    w0.addRow([]);
    sec(['Decisions', 'Items', '', '', '']);
    w0.addRow(['Still waiting for a decision (rows of the To review sheet)', pending.length]);
    if (dc && dc.total) {
      w0.addRow(['Decisions read from the saved workbook', dc.total, '', '', (rec.reviewFiles || []).join(', ')]);
      w0.addRow(['   applied', dc.applied]);
      if (dc.notApplied) { const r = w0.addRow(['   not applied', dc.notApplied, '', '', 'The Decisions sheet says why, row by row.']); r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } }; }
      if (dc.absent) w0.addRow(['   kept for later (the item is not in the files given this time)', dc.absent]);
      if (dc.diff) w0.addRow(['Differences left in matches made by the reviewer', '', rs(dc.diff), '', 'Bank and book amounts differ in these matches. See the Note column.']);
    }
    for (const t of (rec.reviewProblems || [])) { const r = w0.addRow([t]); w0.mergeCells(r.number, 1, r.number, 5); r.getCell(1).alignment = { wrapText: true, vertical: 'top' }; r.height = 30; r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } }; }
    w0.addRow([]);
    sec(['How to decide', '', '', '', '']);
    for (const t of HOWTO) { const r = w0.addRow([t]); w0.mergeCells(r.number, 1, r.number, 5); r.getCell(1).alignment = { wrapText: true, vertical: 'top' }; r.height = t.length > 150 ? 30 : 16; }
    w0.getColumn(3).numFmt = MONEY; w0.getColumn(4).numFmt = MONEY; w0.getColumn(2).alignment = { horizontal: 'center' };
    w0.addRow([]);
    w0.addRow(['Made by IDCO Bank Statement to Excel v' + version + ' on ' + (when || new Date()).toLocaleString('en-IN') + '. Nothing was uploaded.']).font = { italic: true, color: { argb: 'FF6B7280' } };

    // ---------------- To review (written for someone who is not an accountant)
    const SEP = ' ¦ ';
    const sig = (a, b, c) => (a || b || c ? [a || '', b || '', c || ''].join(SEP).trim() : '');
    const mwText = x => (x.matchWith ? x.matchWith.map(describe).join(', ') : '');
    const hide = (ws, names) => { ws.getRow(1).eachCell((c, i) => { if (names.includes(c.value)) ws.getColumn(i).hidden = true; }); };
    const w1 = sheet(wb, 'To review', [['No.', 6], ['Status', 15], ['Kind of item', 30], ['Found in', 10], ['Date', 12, 'd'], ['Who', 34], ['Amount', 17, 'm'], ['Days open', 8], ['What the tool found', 70], ['What to do', 62], ['Where to look', 44], ['Decision', 13], ['Match with', 22, 't'], ['Remark', 30, 't'], ['Item key', 10, 't'], ['Carried', 10, 't']]);
    pending.forEach((it, i) => {
      const x = it.x;
      const r = add(w1, [i + 1, it.st, x.kind || ''].concat(it.v, [x.age, x.found, x.todo, x.where, '', '', x.remark || '', x.ikey, sig('', '', x.remark)]), it.st);
      [3, 6, 9, 10, 11, 14].forEach(c => { r.getCell(c).alignment = { wrapText: true, vertical: 'top' }; });
      [1, 2, 4, 5, 7, 8, 12, 13].forEach(c => { r.getCell(c).alignment = Object.assign({}, r.getCell(c).alignment, { vertical: 'top' }); });
      pick(r.getCell(12), it.st === 'Suggested' ? '"Accept,Reject"' : '"Keep open"'); input(r.getCell(13)); input(r.getCell(14));
    });
    hide(w1, ['Item key', 'Carried']); finish(w1);

    // ---------------- Bank lines
    const w2 = sheet(wb, 'Bank lines', [['Line', 7], ['Date', 12, 'd'], ['Narration', 62], ['Type', 20], ['Withdrawal', 18, 'm'], ['Deposit', 18, 'm'], ['Status', 16], ['Kind of item', 30], ['How it was matched', 46], ['Voucher(s)', 26], ['Paid by (Eazypay credits)', 44], ['Days from bank date to voucher', 12], ['Note', 80], ['Page', 7], ['Statement file', 28], ['Decision', 13], ['Match with', 22, 't'], ['Remark', 30, 't'], ['Item key', 10, 't'], ['Linked keys', 10, 't'], ['Carried', 10, 't']]);
    rec.bankLines.forEach(b => {
      const dec = b.decision || '', mw = mwText(b);
      const r = add(w2, [b.sl, d(b.date), b.narration, b.type, b.amt < 0 ? rs(-b.amt) : null, b.amt > 0 ? rs(b.amt) : null, b.status, b.kind || '', b.how, b.vouchers.join(', '), b.paidBy || '', b.lag, b.note, b.page, b.file || '', dec, mw, b.remark || '', b.ikey, linksOf(b).join(' '), sig(dec, mw, b.remark)], b.status);
      pick(r.getCell(16), LIST_ALL); input(r.getCell(17)); input(r.getCell(18));
    });
    hide(w2, ['Item key', 'Linked keys', 'Carried']); finish(w2);

    // ---------------- Vouchers
    const w3 = sheet(wb, 'Vouchers', [['Voucher', 14], ['Voucher date', 12, 'd'], ['Party in the voucher', 32], ['Paid by (from the Eazypay URN)', 44], ['Receipt (bank debit)', 18, 'm'], ['Payment (bank credit)', 18, 'm'], ['Eazypay URN(s)', 34, 't'], ['Instrument date', 12, 'd'], ['UTR / cheque in ERP', 26, 't'], ['Status', 18], ['Kind of item', 30], ['How it was matched', 46], ['Bank line(s)', 14, 't'], ['Bank date', 12, 'd'], ['Note', 80], ['MR No', 9, 't'], ['Prov No', 9, 't'], ['Bank Book file', 30], ['Decision', 13], ['Match with', 22, 't'], ['Remark', 30, 't'], ['Item key', 10, 't'], ['Linked keys', 10, 't'], ['Carried', 10, 't']]);
    rec.vouchers.forEach(v => {
      const dec = v.decision || '', mw = mwText(v);
      const r = add(w3, [v.blockId, d(v.vchDate), v.party, v.paidBy || '', v.net > 0 ? rs(v.net) : null, v.net < 0 ? rs(-v.net) : null, v.urnText || '', d(v.instrDate), v.utr, v.status, v.kind || '', v.how, v.bank.join(', '), v.bank.length ? d(rec.bankLines[v.bank[0] - 1].date) : null, v.note, v.mrNo, v.provNo, v.file, dec, mw, v.remark || '', v.ikey, linksOf(v).join(' '), sig(dec, mw, v.remark)], v.status);
      pick(r.getCell(19), LIST_ALL); input(r.getCell(20)); input(r.getCell(21));
    });
    hide(w3, ['Item key', 'Linked keys', 'Carried']); finish(w3);

    // ---------------- Eazypay payments
    const w4 = sheet(wb, 'Eazypay payments', [['Settlement date', 12, 'd'], ['Bank line', 9], ['Bank date', 12, 'd'], ['TranID', 18, 't'], ['Party code', 10], ['Party', 40], ['Mode', 16], ['Amount', 16, 'm'], ['Booked in Bank Book', 16, 'm'], ['Voucher(s)', 24], ['Status', 16], ['Kind of item', 30], ['Reviewed', 12], ['Eazypay file', 24]]);
    const kv = new Map(rec.vouchers.map(v => [v.key, v.blockId]));
    rec.payments.slice().sort((a, b) => a.settle.localeCompare(b.settle) || a.tran.localeCompare(b.tran)).forEach(m => add(w4, [d(m.settle), m.bankSl, d(m.bankDate), m.tran, m.partyCode, m.party, m.mode, rs(m.amt), m.vouchers.length ? rs(m.booked) : null, m.vouchers.map(k => kv.get(k)).join(', '), m.bankSl === null ? 'Not in statements given' : m.status, m.kind || '', m.reviewed || '', m.file], m.status));
    finish(w4);

    // ---------------- Decisions (the running record of what the reviewer decided)
    const w5 = sheet(wb, 'Decisions', [['No.', 6], ['Decided on', 12, 't'], ['Item', 20], ['Date', 12, 'd'], ['Who', 34], ['Amount', 17, 'm'], ['Decision', 13], ['Match with', 30, 't'], ['Remark', 34, 't'], ['Result', 60], ['Item key', 10, 't'], ['Linked keys', 10, 't'], ['Carried', 10, 't']]);
    (rec.decisions || []).forEach((x, i) => {
      const L = x.label || {};
      const r = w5.addRow([i + 1, x.on || '', L.what || '(not in these files)', L.date ? d(String(L.date).slice(0, 10)) : null, L.who || '', L.amt === null || L.amt === undefined ? null : rs(L.amt), x.decision, x.decision === 'Match' || x.decision === 'Accept' || x.decision === 'Reject' ? x.with.map(describe).join(', ') : '', x.remark || '', x.result, x.key, x.with.join(' '), sig(x.decision, '', x.remark)]);
      r.getCell(10).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: x.ok ? GREEN : AMBER } };
      [5, 8, 9, 10].forEach(c => { r.getCell(c).alignment = { wrapText: true, vertical: 'top' }; });
      pick(r.getCell(7), LIST_ALL); input(r.getCell(9));
    });
    hide(w5, ['Item key', 'Linked keys', 'Carried']); finish(w5);

    // ---------------- BRS headings (the library of standard situations, with what this file holds of each)
    const w6 = sheet(wb, 'BRS headings', [['Kind of item', 44], ['Usual place in a BRS (from the balance as per Bank Book)', 52], ['What it means', 80], ['Items in this file', 10], ['Amount', 20, 'm'], ['How it is recognised', 44]]);
    (rec.kinds || []).forEach(k => { const r = w6.addRow([k.name, k.brs, k.meaning, k.n || null, k.n ? rs(k.amt) : null, k.seen ? 'Seen and tested on IDCO files' : 'From the usual wording of banks; not yet seen in IDCO files. Please confirm each one.']); [2, 3, 6].forEach(c => { r.getCell(c).alignment = { wrapText: true, vertical: 'top' }; }); if (k.n) r.getCell(1).font = { bold: true }; });
    finish(w6);
    return wb;
  }
  return { build, pendingOf };
});
