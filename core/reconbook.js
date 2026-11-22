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
    '2. In the Decision column choose from the list:  Accept = the suggested match is right.  Reject = it is wrong.  Keep open = looked at; it is a genuine open item. The "Your choices" column says what is possible on each row. A row can stand for several entries that are one question; the decision then applies to all of them.',
    '3. If you know the other side yourself, leave Decision empty and type it in "Match with": a voucher number (CRV-123), a bank line (line 45), or several separated by commas. Two bank lines that cancel each other can be matched with each other in the same way.',
    '4. Remark is free: write anything you want kept with the item.',
    'Items you keep open move to the sheet "Kept open (follow up)", which keeps what to do and where to look until they are settled.',
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
      const f = list[i][2];
      const tip = f === '"Keep open"' ? 'No match was found. Choose Keep open if it is a genuine open item - or leave this empty and type the voucher no. / bank line in Match with.'
        : f === '"Accept,Reject"' ? 'The tool suggests a match (see What the tool found). Accept if it is right, Reject if it is not.'
        : 'Accept or Reject a match, Keep open an open item, Undo to take a decision back.';
      ws.dataValidations.model[L + list[i][0] + ':' + L + list[j][0]] = { type: 'list', allowBlank: true, formulae: [f], showErrorMessage: true, errorTitle: 'Decision', error: 'Please choose from the list.', showInputMessage: true, promptTitle: 'Decision', prompt: tip };
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

  /** The items still waiting for a person: the rows of the To review sheet.
   *  Entries that are one question for the reviewer (a partly booked Eazypay collection with its payments and
   *  vouchers, the deposit interest credits, everything waiting on one missing file, the Re 1 test credits)
   *  are one row; a decision on that row applies to every entry in it. */
  function pendingOf(rec, kept) {
    const want = x => !!x.reviewed === !!kept;
    const SHOW = new Set(['Suggested', 'Needs review', 'Not in books', 'Not in bank', 'Bank only']);
    const fmt = rec.fmt, nice = rec.nice;
    const asOn = rec.brs && rec.brs.ready ? rec.brs.asOn : '';
    const upto = asOn || rec.summary.lastBank || '';
    const age = iso => (iso && upto ? Math.round((Date.parse(upto) - Date.parse(iso)) / 86400000) : null);
    const short = (t, n) => { t = String(t || ''); return t.length > n ? t.slice(0, n - 1) + '…' : t; };
    const list = (a, n) => a.slice(0, n).join(', ') + (a.length > n ? ' and ' + (a.length - n) + ' more' : '');
    const KEEP = 'Keep open if it is a genuine open item, or type the voucher no. / bank line in Match with.';
    const one = [];
    rec.bankLines.filter(b => SHOW.has(b.status) && b.found && want(b)).forEach(b => one.push({ x: b, t: 'B', st: b.status, date: b.date, from: 'Bank', who: b.party || (b.type === 'Other' ? short(b.narration, 40) : b.type), amt: b.amt, narr: b.narration }));
    rec.vouchers.filter(v => SHOW.has(v.status) && v.found && want(v) && !(v.status === 'Suggested' && v.bank.length)).forEach(v => one.push({ x: v, t: 'V', st: v.status, date: v.vchDate, from: 'ERP books', who: v.paidBy || v.party, amt: v.net, narr: '' }));
    rec.payments.filter(m => m.found && want(m)).forEach(m => one.push({ x: m, t: 'E', st: m.status === 'Booked in part' ? 'Needs review' : 'Not in books', date: m.settle, from: 'Eazypay', who: (m.partyCode ? m.partyCode + ' ' : '') + m.party, amt: m.amt, narr: '' }));
    const live = asOn ? one.filter(p => p.date <= asOn) : one;

    // which entries belong together
    const parent = new Map(); const find = k => { while (parent.get(k) !== k) { parent.set(k, parent.get(parent.get(k))); k = parent.get(k); } return k; };
    const join = (a, b) => { [a, b].forEach(k => { if (!parent.has(k)) parent.set(k, k); }); const ra = find(a), rb = find(b); if (ra !== rb) parent.set(ra, rb); };
    const vByKey = new Map(rec.vouchers.map(v => [v.key, v]));
    const bucket = new Map(); const put = (k, it) => { if (!bucket.has(k)) bucket.set(k, []); bucket.get(k).push(it); };
    for (const it of live) {
      const x = it.x;
      if (it.t === 'V' && x.status === 'Needs review' && x.ez && x.ez.diff) (x.ez.trans || []).forEach(tr => join(x.ikey, 'E|' + tr));
      else if (it.t === 'E' && x.status === 'Booked in part') x.vouchers.forEach(k => { const v = vByKey.get(k); if (v) join(x.ikey, v.ikey); });
    }
    const plain = (it, choices) => { const x = it.x; it.keys = [x.ikey]; it.kind = x.kind || ''; it.found = x.found; it.todo = x.todo; it.where = x.where; it.remark = x.remark || ''; it.age = age(it.t === 'V' ? (x.instruments.map(i => i.date).filter(Boolean).sort()[0] || x.vchDate) : it.date); it.choices = choices; return it; };
    const out = [];
    for (const it of live) {
      const x = it.x;
      if (parent.has(x.ikey)) put('EZ|' + find(x.ikey), it);
      else if (x.kindCode === 'FILE_MISSING') put('FILE', it);
      else if (it.t === 'B' && x.status === 'Bank only' && x.type === 'FD closure / interest') put('FD', it);
      else if (x.kindCode === 'TOKEN') put('TOKEN', it);
      else out.push(plain(it, it.st === 'Suggested' ? 'Accept if the match is right. Reject if it is not.' : it.t === 'E' ? 'Keep open until it is booked. Once a voucher carries this URN in the ERP, the tool matches it by itself.' : KEEP));
    }
    const ALL = n => 'Keep open (it applies to all ' + n + ' entries of this row). ';
    for (const [k, mem] of bucket) {
      if (mem.length === 1 && !/^EZ/.test(k)) { out.push(plain(mem[0], k === 'TOKEN' ? KEEP : 'Keep open.')); continue; }
      mem.sort((a, b) => a.date.localeCompare(b.date));
      const g = { x: null, members: mem, keys: mem.map(m => m.x.ikey), date: mem[0].date, remark: (mem.find(m => m.x.remark) || { x: {} }).x.remark || '', age: age(mem[0].date), narr: '' };
      const lines = mem.filter(m => m.t === 'B').map(m => m.x.sl), sum = mem.filter(m => m.t === 'B').reduce((a, m) => a + m.amt, 0);
      if (/^EZ/.test(k)) {
        const vs = mem.filter(m => m.t === 'V').map(m => m.x), ps = mem.filter(m => m.t === 'E').map(m => m.x);
        const v0 = vs[0] || vByKey.get((ps[0].vouchers || [])[0]); const ez = v0 && v0.ez && v0.ez.diff ? v0.ez : null;
        if (!ez) { mem.forEach(it => out.push(plain(it, 'Keep open.'))); continue; }
        const less = ez.diff < 0, ids = ez.vs, urns = ez.trans;
        const sls = Array.from(new Set(ps.map(m => m.bankSl).filter(Boolean)));
        Object.assign(g, { st: 'Needs review', from: 'ERP books + Eazypay', who: short(v0.paidBy || v0.party, 60), amt: Math.abs(ez.diff), kind: less ? 'Collection only partly booked' : 'Booked more than received',
          found: urns.length + ' Eazypay payment' + (urns.length > 1 ? 's totalling ' : ' of ') + '' + fmt(ez.pay) + ' (URN ' + list(urns, 4) + ') ' + (urns.length > 1 ? 'belong' : 'belongs') + ' to voucher' + (ids.length > 1 ? 's ' : ' ') + list(ids, 6) + '. ' + (ids.length > 1 ? 'The vouchers together book ' : 'The voucher books ') + fmt(ez.net) + (less ? ', so ' + fmt(-ez.diff) + ' of the money received is not booked.' : ', which is ' + fmt(ez.diff) + ' more than was received.'),
          todo: less ? 'Open ' + (ids.length > 1 ? 'these vouchers' : 'the voucher') + ' in the ERP and find where the remaining ' + fmt(-ez.diff) + ' should go. Often another division or wing has not yet booked its share.' : 'Open ' + (ids.length > 1 ? 'these vouchers' : 'the voucher') + ' in the ERP and check the amounts. Part of a payment may have been booked twice, or a URN may be missing from one voucher.',
          where: 'ERP voucher' + (ids.length > 1 ? 's ' : ' ') + list(ids, 6) + ' (first dated ' + nice(v0.vchDate) + '). Bank Book file: ' + v0.file + '.' + (sls.length ? ' In the bank statement the money is inside line' + (sls.length > 1 ? 's ' : ' ') + sls.join(', ') + '.' : ''),
          choices: (mem.length > 1 ? ALL(mem.length) : 'Keep open. ') + 'Nothing else is possible here: once it is ' + (less ? 'booked' : 'corrected') + ' in the ERP, the tool matches it by itself.' });
      } else if (k === 'FD') {
        Object.assign(g, { st: 'Bank only', from: 'Bank', who: 'Linked fixed deposits', amt: sum, kind: 'Interest / deposit closure credits',
          found: 'The bank added ' + mem.length + ' credits totalling ' + fmt(sum) + ' between ' + nice(mem[0].date) + ' and ' + nice(mem[mem.length - 1].date) + ' from fixed deposits linked to this account. This is interest or closure money, put in by the bank itself. None of it is in the ERP Bank Book.',
          todo: 'Tell the person who books bank interest. The interest and TDS figures are on the last pages of the bank statement.',
          where: 'Bank lines sheet: filter the Type column on "FD closure / interest" (lines ' + list(lines, 8) + ').', choices: ALL(mem.length) + 'Once booked in the ERP, type the voucher no. on the Bank lines sheet.' });
      } else if (k === 'TOKEN') {
        Object.assign(g, { st: 'Not in books', from: 'Bank', who: short(list(mem.map(m => m.x.party).filter(Boolean), 3), 60), amt: sum, kind: 'Token / test credit of Re 1',
          found: mem.length + ' credits of Re 1 were received; parties usually send Re 1 to test the account number before paying (' + list(mem.map(m => (m.x.party || 'party not named') + ' on ' + nice(m.date)), 6) + '). No receipt voucher was found for these.',
          todo: 'Check in the ERP whether these have been booked. If not, they still have to be booked.',
          where: 'Bank lines sheet, lines ' + list(lines, 10) + '.', choices: ALL(mem.length) + 'To match one of them with a voucher, use the Bank lines sheet.' });
      } else {
        const bm = mem.filter(m => m.t === 'B'), vm = mem.filter(m => m.t === 'V');
        Object.assign(g, { st: 'Needs review', from: bm.length ? 'Bank + ERP books' : 'ERP books', who: 'Eazypay', amt: bm.length ? sum : vm.reduce((a, m) => a + m.amt, 0), kind: 'Cannot be checked - a file is missing',
          found: (bm.length ? bm.length + ' Eazypay credit' + (bm.length > 1 ? 's' : '') + ' in the bank (' + list(bm.map(m => fmt(m.amt) + ' on ' + nice(m.date)), 4) + ')' : '') + (bm.length && vm.length ? ' and ' : '') + (vm.length ? vm.length + ' voucher' + (vm.length > 1 ? 's' : '') + ' received through Eazypay (' + list(vm.map(m => m.x.blockId), 8) + ')' : '') + ' could not be checked, because the Eazypay report rows for those days were not given.',
          todo: bm.length ? 'Download from the bank the Eazypay report that covers ' + list(bm.map(m => nice(m.date)), 4) + ' and the days just before, and drop it in with the other files. Nothing is wrong with the vouchers; they will clear once the report is added.' : vm[0].x.todo,
          where: (bm.length ? 'Bank lines sheet, line' + (lines.length > 1 ? 's ' : ' ') + list(lines, 6) + '. ' : '') + (vm.length ? 'Vouchers sheet: ' + list(vm.map(m => m.x.blockId), 8) + '.' : ''),
          choices: 'Usually nothing to decide: add the missing report next time. Keep open only if the report cannot be had (it applies to all ' + mem.length + ' entries of this row).' });
      }
      out.push(g);
    }
    // quick confirmations first, then the largest amounts
    out.sort((a, b) => ((a.st === 'Suggested' ? 0 : 1) - (b.st === 'Suggested' ? 0 : 1)) || (Math.abs(b.amt) - Math.abs(a.amt)) || a.date.localeCompare(b.date));
    return out;
  }

  const nd = iso => (iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.' + iso.slice(0, 4) : '');
  // IDCO's own BRS layout: four headings only. The tool's extra lines are folded in by their sign and shown below.
  const IDCO = { ISSUED: 'ISSUED', BANK_CR: 'BANK_CR', DIFF_PLUS: 'BANK_CR', BOOKNET_PLUS: 'ISSUED', DEPOSITED: 'DEPOSITED', BANK_DR: 'BANK_DR', DIFF_MINUS: 'DEPOSITED', BOOKNET_MINUS: 'DEPOSITED' };
  const IDCO_LABEL = { ISSUED: 'Add Cheque issued but not presented for payment', BANK_CR: 'Add Amount Credited by Bank Only', DEPOSITED: 'Less Cheque deposited but not credited by Bank', BANK_DR: 'Less Debited by Bank Only' };
  function idcoFormat(wb, rec, R) {
    const ws = wb.addWorksheet('BRS (IDCO format)', { views: [{ showGridLines: false }] });
    ws.columns = [{ width: 52 }, { width: 30 }, { width: 70 }];
    const YEL = 'FFFFFF99', B = { style: 'thin', color: { argb: 'FF999999' } }, box = c => { c.border = { top: B, left: B, bottom: B, right: B }; };
    const t = ws.addRow(['BANK RECONCILIATION STATEMENT AS ON ' + nd(R.asOn)]); t.font = { bold: true, size: 13 };
    ws.addRow([]);
    const totals = { ISSUED: 0, BANK_CR: 0, DEPOSITED: 0, BANK_DR: 0 }, parts = { ISSUED: [], BANK_CR: [], DEPOSITED: [], BANK_DR: [] };
    for (const h of R.order) { const tt = R.totals[h]; if (!tt.n) continue; const k = IDCO[h]; totals[k] += tt.amt; parts[k].push([R.heads[h].replace(/^(Add|Less): /, ''), tt.n, tt.amt]); }
    if (R.sweepN) { const k = R.sweep >= 0 ? 'BANK_CR' : 'BANK_DR'; totals[k] += Math.abs(R.sweep); parts[k].push(['Net amount moved by the bank to / from linked deposits (sweep)', R.sweepN, Math.abs(R.sweep)]); }
    const head = ws.addRow(['Particulars', (R.opening.name || (R.bankName + ' A/C')).toUpperCase()]);
    head.font = { bold: true }; head.height = 34; [1, 2].forEach(c => { const x = head.getCell(c); x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: YEL } }; x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; box(x); });
    const line = (label, v, opt) => { opt = opt || {}; const r = ws.addRow([label, v]); [1, 2].forEach(c => box(r.getCell(c))); r.getCell(1).alignment = { wrapText: true, vertical: 'middle', horizontal: opt.center ? 'center' : 'left' }; r.getCell(2).alignment = { horizontal: opt.center ? 'center' : 'right', vertical: 'middle' }; if (typeof v === 'number' || (v && v.formula)) r.getCell(2).numFmt = MONEY; if (opt.bold) r.font = { bold: true }; if (opt.fill) [1, 2].forEach(c => { r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: YEL } }; }); r.height = opt.h || 30; return r; };
    line('GL CODE NO-', R.opening.gl || '', { center: true, bold: true, h: 20 });
    line('A/C-NO-', R.account, { center: true, bold: true, h: 20 }).getCell(2).font = { bold: true, color: { argb: 'FF1D4ED8' } };
    const rb = line('Balance as per Bank Book as on ' + nd(R.asOn), rs(R.book)); rb.getCell(1).font = { bold: true, color: { argb: 'FF7C2D5B' } };
    const rows = {}; for (const k of ['ISSUED', 'BANK_CR', 'DEPOSITED', 'BANK_DR']) { rows[k] = line(IDCO_LABEL[k], rs(totals[k])); rows[k].getCell(1).font = { bold: true }; }
    const n = k => 'B' + rows[k].number;
    const res = R.book + totals.ISSUED + totals.BANK_CR - totals.DEPOSITED - totals.BANK_DR;
    const rr = line('Balance as per Bank statement as on ' + nd(R.asOn), { formula: 'B' + rb.number + '+' + n('ISSUED') + '+' + n('BANK_CR') + '-' + n('DEPOSITED') + '-' + n('BANK_DR'), result: rs(res) }, { bold: true, fill: true });
    const pr = ws.addRow(['Balance printed by the bank on ' + nd(R.asOn), rs(R.bank)]); pr.getCell(2).numFmt = MONEY; pr.getCell(1).font = { color: { argb: 'FF6B7280' } };
    const df = ws.addRow(['Difference', { formula: 'B' + pr.number + '-B' + rr.number, result: rs(R.bank - res) }]); df.getCell(2).numFmt = MONEY; df.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: R.bank === res ? GREEN : RED } };
    ws.addRow([]);
    const sh = ws.addRow(['What each figure is made of', 'Amount (Rs.)', 'Where the items are']); sh.font = { bold: true, color: { argb: 'FFFFFFFF' } }; [1, 2, 3].forEach(c => { sh.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } }; });
    for (const k of ['ISSUED', 'BANK_CR', 'DEPOSITED', 'BANK_DR']) {
      const r = ws.addRow([IDCO_LABEL[k], rs(totals[k])]); r.font = { bold: true }; r.getCell(2).numFmt = MONEY;
      if (!parts[k].length) { ws.addRow(['     (nothing)']); continue; }
      for (const [lab, cnt, amt] of parts[k]) { const q = ws.addRow(['     ' + lab, rs(amt), cnt + ' item' + (cnt > 1 ? 's' : '') + (/sweep/.test(lab) ? ' (Bank lines sheet, Status = Sweep)' : ' (BRS items sheet)')]); q.getCell(2).numFmt = MONEY; q.getCell(1).alignment = { wrapText: true }; }
    }
    ws.addRow([]);
    const failed = R.checks.filter(c => !c.pass);
    const notes = (failed.length ? ['NOT READY TO SIGN - ' + failed.map(c => c.what + ': ' + c.detail).join(' ')] : []).concat(R.tried ? [] : ['NOT YET TRIED ON THIS ACCOUNT: the tool has so far been checked only on ' + R.triedList + '. Check every figure of this BRS before relying on it.']).concat(['Prepared by the tool from the files given. ' + (R.suggested ? R.suggested + ' item(s) are open only because the match the tool suggests for them is not yet confirmed; once accepted in To review, they drop out. ' : '') + 'The open items are listed in the sheet "BRS items"; check them before signing.',
      'The four headings are those of the IDCO BRS. The tool\'s other lines (differences in matched entries, book entries with no bank entry this year, sweeps) have no heading of their own in that format, so each is put under the heading of the same sign, as shown above. Move them if the accounts section shows them differently.']);
    for (const tx of notes) { const r = ws.addRow([tx]); ws.mergeCells(r.number, 1, r.number, 3); r.getCell(1).alignment = { wrapText: true, vertical: 'top' }; r.height = 44; }
    ws.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  }

  function brsStatement(wb, rec, R, when, version) {
    const fmt = rec.fmt; const ws = wb.addWorksheet('BRS', { views: [{ showGridLines: false }] });
    ws.columns = [{ width: 86 }, { width: 9 }, { width: 22 }, { width: 70 }];
    const T = ws.addRow(['BANK RECONCILIATION STATEMENT AS ON ' + nd(R.asOn)]); T.font = { bold: true, size: 14, color: { argb: TEAL } };
    ws.addRow([R.bankName + ' - A/c ' + R.account + (R.opening.gl ? ' - GL ' + R.opening.gl : '') + (R.opening.name ? ' - ' + R.opening.name : '')]).font = { bold: true };
    ws.addRow(['Period ' + nd(R.fy) + ' to ' + nd(R.asOn) + '. Made by IDCO Bank Statement to Excel v' + version + ' on ' + nd(when.toISOString().slice(0, 10)) + '. Opening position: BRS as on ' + nd(R.opening.asOn) + ' (' + R.opening.file + ').']).font = { color: { argb: 'FF6B7280' } };
    ws.addRow([]);
    const H = ws.addRow(['Particulars', 'Items', 'Amount (Rs.)', 'Note']); H.font = { bold: true, color: { argb: 'FFFFFFFF' } }; [1, 2, 3, 4].forEach(c => { H.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } }; });
    const money = r => { r.getCell(3).numFmt = MONEY; r.getCell(4).alignment = { wrapText: true, vertical: 'top' }; r.getCell(1).alignment = { wrapText: true, vertical: 'top' }; return r; };
    const bookRow = money(ws.addRow(['Balance as per Bank Book as on ' + nd(R.asOn), '', null, 'Balance of the last BRS + vouchers of this year up to ' + nd(R.asOn) + '.'])); bookRow.font = { bold: true };
    const r1 = money(ws.addRow(['     Balance as per Bank Book as on ' + nd(R.opening.asOn) + ' (last signed BRS)', '', rs(R.bookOpen)]));
    const r2 = money(ws.addRow(['     Net of all vouchers dated ' + nd(R.fy) + ' to ' + nd(R.asOn) + ' (receipts less payments)', '', rs(R.bookMove)]));
    bookRow.getCell(3).value = { formula: 'C' + r1.number + '+C' + r2.number, result: rs(R.book) };
    const erp = money(ws.addRow(['     Balance as per the ERP Bank Book report as on ' + nd(R.asOn) + ' - type it here to cross-check', '', null]));
    erp.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFBEB' } };
    erp.getCell(4).value = { formula: 'IF(C' + erp.number + '="","Type the closing balance shown by the ERP to cross-check.",IF(ROUND(C' + erp.number + '-C' + bookRow.number + ',2)=0,"Agrees with the ERP.","Differs from the ERP by "&TEXT(C' + erp.number + '-C' + bookRow.number + ',"#,##0.00")&". Please check."))', result: 'Type the closing balance shown by the ERP to cross-check.' };
    const lines = [];
    for (const h of R.order) {
      const t = R.totals[h]; if (!t.n && /^DIFF|^BOOKNET/.test(h)) continue;
      const lab = R.heads[h];
      const r = money(ws.addRow([lab, { formula: 'COUNTIFS(\'BRS items\'!B:B,A' + (ws.rowCount + 1) + ')', result: t.n }, { formula: 'SUMIFS(\'BRS items\'!F:F,\'BRS items\'!B:B,A' + (ws.rowCount + 1) + ')', result: rs(t.amt) }, t.n ? 'Listed in the sheet "BRS items" under this heading.' : '']));
      lines.push([r.number, R.sign[h]]);
    }
    if (R.sweepN) { const r = money(ws.addRow([(R.sweep >= 0 ? 'Add' : 'Less') + ': Net amount moved by the bank between this account and its linked deposits (sweep)', R.sweepN, rs(Math.abs(R.sweep)), 'Shown as one net line. Sweeps are not entered in the Bank Book.'])); lines.push([r.number, R.sweep >= 0 ? 1 : -1]); }
    const calc = money(ws.addRow(['Balance as per bank statement as on ' + nd(R.asOn) + ' (worked out)', '', { formula: 'C' + bookRow.number + lines.map(([n, sg]) => (sg > 0 ? '+' : '-') + 'C' + n).join(''), result: rs(R.computed) }])); calc.font = { bold: true };
    const real = money(ws.addRow(['Balance as per bank statement as on ' + nd(R.asOn) + ' (printed by the bank)', '', rs(R.bank), 'From the bank statements given.'])); real.font = { bold: true };
    const diff = money(ws.addRow(['Difference', '', { formula: 'C' + real.number + '-C' + calc.number, result: rs(R.diff) }, R.diff === 0 ? 'Nil: the statement agrees.' : 'Not nil: see the checks below.']));
    diff.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: R.diff === 0 ? GREEN : RED } };
    ws.addRow([]);
    const sec = t => { const r = ws.addRow([t]); r.font = { bold: true, color: { argb: 'FFFFFFFF' } }; [1, 2, 3, 4].forEach(c => { r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } }; }); };
    sec('Read this first');
    const notes = (R.tried ? [] : ['NOT YET TRIED ON THIS ACCOUNT: the tool has so far been checked only on ' + R.triedList + '. Check every figure of this BRS before relying on it.']).concat([
      'Prepared by the tool from the bank statements, the ERP Bank Book and the last signed BRS. It lists and adds up; it does not decide. Check the items before signing.',
      'An item counts as settled only if both sides are dated on or before ' + nd(R.asOn) + '. Items settled later still appear here; the column "Settled after ' + nd(R.asOn) + '" in "BRS items" says when and how.',
      R.suggested ? R.suggested + ' item(s) are listed as open only because the match the tool suggests for them is not yet confirmed. Accept the right matches in the "To review" sheet and those items drop out next time. The BRS never counts a match nobody has confirmed.' : 'Every match used in this BRS is confirmed (no suggested matches).',
      'Last signed BRS (' + nd(R.opening.asOn) + '): ' + R.opening.n + ' items; ' + R.opening.cleared + ' settled by ' + nd(R.asOn) + ', ' + R.opening.stillOpen + ' still open. See the sheet "Last BRS".'
    ]);
    if (R.bookNet) notes.push('Receipts booked this year for money that came into the bank before 1 April, and reversals of suspense entries of the earlier year, have no bank entry this year. They are shown as one net line (' + fmt(Math.abs(R.bookNet.net)) + '); the vouchers are in the sheet "BRS book entries". Compare them with the suspense details of the earlier year.');
    for (const t of notes) { const r = ws.addRow([t]); ws.mergeCells(r.number, 1, r.number, 4); r.getCell(1).alignment = { wrapText: true, vertical: 'top' }; r.height = t.length > 120 ? 32 : 18; }
    ws.addRow([]); sec('Checks');
    for (const c of R.checks) { const r = ws.addRow([c.what, c.pass ? 'PASS' : 'CHECK', '', c.detail]); r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: c.pass ? GREEN : RED } }; r.getCell(1).alignment = { wrapText: true, vertical: 'top' }; r.getCell(4).alignment = { wrapText: true, vertical: 'top' }; }
    ws.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  }

  function brsSchedules(wb, rec, R, h) {
    const nice = rec.nice;
    const w = sheet(wb, 'BRS items', [['No.', 6], ['Heading', 46], ['Date', 12, 'd'], ['Who', 34], ['Reference', 26, 't'], ['Amount', 18, 'm'], ['Particulars', 70], ['From', 22], ['Kind of item', 30], ['Days open', 9], ['Settled after ' + nd(R.asOn), 44], ['Status in matching', 16], ['Heading in the IDCO format', 40], ['Reviewed', 11], ['Remark', 30], ['What to do', 56], ['Where to look', 44]]);
    const byKey = new Map(); rec.bankLines.forEach(b => byKey.set(b.ikey || b.key, b)); rec.vouchers.forEach(v => byKey.set(v.ikey, v)); rec.payments.forEach(m => byKey.set(m.ikey, m)); (rec.opening || []).forEach(o => byKey.set(o.ikey, o));
    R.items.forEach(it => { const r = add(w, [it.no, R.heads[it.head], d(it.date), it.who, it.ref, rs(it.amt), it.what, it.source, it.kind, it.age, it.later, it.status, IDCO_LABEL[IDCO[it.head]], (g => g.reviewed || '')(it.keys.length === 1 && byKey.get(it.keys[0]) || {}), (g => g.remark || '')(it.keys.length === 1 && byKey.get(it.keys[0]) || {}), (g => g.todo || '')(it.keys.length === 1 && byKey.get(it.keys[0]) || {}), (g => g.where || '')(it.keys.length === 1 && byKey.get(it.keys[0]) || {})], it.later ? 'Matched' : (it.status === 'Suggested' ? 'Suggested' : 'Not in books')); [2, 4, 7, 11, 15, 16, 17].forEach(c => { r.getCell(c).alignment = { wrapText: true, vertical: 'top' }; }); });
    finish(w);
    if (R.bookNet) {
      const wb2 = sheet(wb, 'BRS book entries', [['Kind', 30], ['Date', 12, 'd'], ['Voucher', 14], ['Who', 34], ['Amount (+ receipt, - payment)', 20, 'm'], ['Instrument date', 13, 'd'], ['Narration', 90]]);
      R.bookNet.members.slice().sort((a, b) => a.date.localeCompare(b.date)).forEach(m => { const r = add(wb2, [m.head, d(m.date), m.id, m.who, rs(m.amt), d(m.instr), m.what], m.head.indexOf('Suspense') === 0 ? 'Suspense entry' : 'Earlier period'); r.getCell(7).alignment = { wrapText: true, vertical: 'top' }; });
      const tr = wb2.addRow(['Net', null, '', '', { formula: 'SUM(E2:E' + wb2.rowCount + ')', result: rs(R.bookNet.net) }]); tr.font = { bold: true }; tr.getCell(5).numFmt = MONEY;
      finish(wb2);
    }
    const ops = rec.opening || []; if (!ops.length) return;
    const cleared = new Set(R.openingCleared || []);
    const w3 = sheet(wb, 'Last BRS', [['No.', 6], ['Heading in the BRS of ' + nd(R.opening.asOn), 40], ['Date', 12, 't'], ['Party', 34], ['Reference', 26, 't'], ['Amount', 18, 'm'], ['As on ' + nd(R.asOn), 16], ['Settled by', 40], ['Decision', 13], ['Match with', 22, 't'], ['Remark', 30, 't'], ['Item key', 10, 't'], ['Linked keys', 10, 't'], ['Carried', 10, 't']]);
    const NAMES = { BANK_CR: 'Amount credited by bank only', BANK_DR: 'Amount debited by bank only', DEPOSITED: 'Cheque deposited but not credited by bank', ISSUED: 'Cheque issued but not presented for payment' };
    ops.forEach(o => {
      const st = cleared.has(o.ikey) ? 'Settled' : 'Still open';
      const by = (o.openLinks || []).map(x => (x.sl !== undefined && x.narration !== undefined ? 'bank line ' + x.sl + ' (' + nice(x.date) + ')' : x.blockId ? x.blockId + ' (' + nice(x.vchDate) + ')' : '')).filter(Boolean).join(', ');
      const dec = o.decision || '', mw = o.matchWith ? o.matchWith.map(h.describe).join(', ') : '';
      const r = add(w3, [o.no, NAMES[o.head], o.dateText, o.party, o.ref, rs(o.amt), st + (o.status === 'Suggested' ? ' (suggested)' : ''), by + (by && o.how ? ' - ' + o.how : ''), dec, mw, o.remark || '', o.ikey, h.linksOf(o).join(' '), h.sig(dec, mw, o.remark)], st === 'Settled' ? (o.status === 'Suggested' ? 'Suggested' : 'Matched') : 'Not in bank');
      pick(r.getCell(9), '"Accept,Reject,Keep open,Undo"'); input(r.getCell(10)); input(r.getCell(11)); [4, 8].forEach(c => { r.getCell(c).alignment = { wrapText: true, vertical: 'top' }; });
    });
    h.hide(w3, ['Item key', 'Linked keys', 'Carried']); finish(w3);
  }

  function build(ExcelJS, rec, when, version) {
    const wb = new ExcelJS.Workbook(); wb.creator = 'IDCO Bank Statement to Excel v' + version; wb.created = when || new Date();
    const s = rec.summary, fmt = rec.fmt, nice = rec.nice;
    const R0 = rec.brs && rec.brs.ready ? rec.brs : null;
    if (R0) { brsStatement(wb, rec, R0, when || new Date(), version); idcoFormat(wb, rec, R0); }
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
    const SEP = ' ¦ ', GSEP = ' ;; ';
    const sig = (a, b, c) => (a || b || c ? [a || '', b || '', c || ''].join(SEP).trim() : '');
    const mwText = x => (x.matchWith ? x.matchWith.map(describe).join(', ') : '');
    const hide = (ws, names) => { ws.getRow(1).eachCell((c, i) => { if (names.includes(c.value)) ws.getColumn(i).hidden = true; }); };
    const w1 = sheet(wb, 'To review', [['No.', 6], ['Status', 14], ['Date', 12, 'd'], ['Who', 32], ['Amount', 17, 'm'], ['Decision', 13], ['Match with', 18, 't'], ['Remark', 26, 't'], ['Your choices', 44], ['Kind of item', 28], ['Found in', 11], ['Days open', 8], ['What the tool found', 70], ['What to do', 56], ['Bank narration', 44], ['Where to look', 44], ['Item key', 10, 't'], ['Carried', 10, 't']]);
    w1.views = [{ state: 'frozen', xSplit: 5, ySplit: 1 }];
    pending.forEach((it, i) => {
      const r = add(w1, [i + 1, it.st, d(it.date), it.who, rs(it.amt), '', '', it.remark || '', it.choices, it.kind || '', it.from, it.age, it.found, it.todo, it.narr || '', it.where, it.keys.join(GSEP), sig('', '', it.remark)], it.st);
      [4, 8, 9, 10, 11, 13, 14, 15, 16].forEach(c => { r.getCell(c).alignment = { wrapText: true, vertical: 'top' }; });
      [1, 2, 3, 5, 6, 7, 12].forEach(c => { r.getCell(c).alignment = Object.assign({}, r.getCell(c).alignment, { vertical: 'top' }); });
      pick(r.getCell(6), it.st === 'Suggested' ? '"Accept,Reject"' : '"Keep open"'); input(r.getCell(7)); input(r.getCell(8));
    });
    hide(w1, ['Item key', 'Carried']); finish(w1);

    // ---------------- Kept open (follow up): decided as genuinely open, still open, with the guidance kept in view
    {
      const kept = pendingOf(rec, true);
      const when = new Map((rec.decisions || []).filter(x => x.decision === 'Keep open').map(x => [x.key, x.on || '']));
      const wk = sheet(wb, 'Kept open (follow up)', [['No.', 6], ['Date', 12, 'd'], ['Who', 32], ['Amount', 17, 'm'], ['Days open', 8], ['Remark', 30], ['Kept open on', 12, 't'], ['Kind of item', 28], ['Found in', 11], ['What the tool found', 70], ['What to do', 56], ['Bank narration', 44], ['Where to look', 44]]);
      wk.views = [{ state: 'frozen', xSplit: 4, ySplit: 1 }];
      kept.forEach((it, i) => {
        const r = add(wk, [i + 1, d(it.date), it.who, rs(it.amt), it.age, it.remark || '', when.get(it.keys[0]) || '', it.kind || '', it.from, it.found, it.todo, it.narr || '', it.where], it.st);
        [3, 6, 8, 10, 11, 12, 13].forEach(c => { r.getCell(c).alignment = { wrapText: true, vertical: 'top' }; });
        [1, 2, 4, 5, 7, 9].forEach(c => { r.getCell(c).alignment = Object.assign({}, r.getCell(c).alignment, { vertical: 'top' }); });
      });
      if (!kept.length) wk.addRow(['', null, 'Nothing has been kept open yet. Items you mark Keep open in To review are listed here, with what to do and where to look, until they are settled.']);
      finish(wk);
    }

    // ---------------- To be booked: Eazypay money the bank has credited and no voucher carries, party by party
    {
      const lim = R0 ? R0.asOn : '';
      const un = rec.payments.filter(m => m.bankSl && m.status === 'Not in books' && (!lim || (m.bankDate || m.settle) <= lim));
      if (un.length) {
        const by = new Map(); un.forEach(m => { const k = (m.partyCode || '') + '|' + m.party; if (!by.has(k)) by.set(k, []); by.get(k).push(m); });
        const wq = sheet(wb, 'To be booked', [['Party code', 11], ['Party', 46], ['Payments', 9], ['Amount', 18, 'm'], ['First credited', 13, 'd'], ['Last credited', 13, 'd'], ['Eazypay URN(s)', 90, 't']]);
        Array.from(by.values()).map(a => ({ a, t: a.reduce((x, m) => x + m.amt, 0) })).sort((p, q) => q.t - p.t).forEach(({ a, t }) => {
          const ds = a.map(m => m.bankDate || m.settle).sort();
          const r = wq.addRow([a[0].partyCode || '', a[0].party, a.length, rs(t), d(ds[0]), d(ds[ds.length - 1]), a.map(m => m.tran).join(', ')]); r.getCell(7).alignment = { wrapText: true, vertical: 'top' }; r.getCell(2).alignment = { wrapText: true, vertical: 'top' };
        });
        const n = wq.rowCount;
        const tr = wq.addRow(['', 'Total', { formula: 'SUM(C2:C' + n + ')', result: un.length }, { formula: 'SUM(D2:D' + n + ')', result: rs(un.reduce((x, m) => x + m.amt, 0)) }]); tr.font = { bold: true }; tr.getCell(4).numFmt = MONEY;
        finish(wq);
      }
    }


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
    if (R0) brsSchedules(wb, rec, R0, { sig, describe, linksOf, hide });

    // ---------------- BRS headings (the library of standard situations, with what this file holds of each)
    const w6 = sheet(wb, 'BRS headings', [['Kind of item', 44], ['Usual place in a BRS (from the balance as per Bank Book)', 52], ['What it means', 80], ['Items in this file', 10], ['Amount', 20, 'm'], ['How it is recognised', 44]]);
    (rec.kinds || []).forEach(k => { const r = w6.addRow([k.name, k.brs, k.meaning, k.n || null, k.n ? rs(k.amt) : null, k.seen ? 'Seen and tested on IDCO files' : 'From the usual wording of banks; not yet seen in IDCO files. Please confirm each one.']); [2, 3, 6].forEach(c => { r.getCell(c).alignment = { wrapText: true, vertical: 'top' }; }); if (k.n) r.getCell(1).font = { bold: true }; });
    finish(w6);
    return wb;
  }
  return { build, pendingOf };
});
