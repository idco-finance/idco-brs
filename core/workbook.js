/* IDCO Bank Statement to Excel - the Excel workbook.
 * Same three sheets and the same columns for every bank:
 *   Statement - one row per transaction
 *   Control   - what was checked, with a clear PASS or CHECK
 *   By Type   - counts and totals by type of transaction */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementWorkbook = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const VERSION = '1.4.2';
  const TEAL = 'FF0F766E', PAPER = 'FFF7F3EA', INK = 'FF1F2937', GREEN = 'FFD1FAE5', RED = 'FFFEE2E2', AMBER = 'FFFEF3C7', GREY = 'FFE5E7EB';
  const MONEY = '#,##0.00;[Red]-#,##0.00';

  const rupees = p => (p === null || p === undefined ? null : p / 100);
  const toDate = iso => (iso ? new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10))) : null);
  const nice = iso => (iso ? iso.slice(8, 10) + '-' + iso.slice(5, 7) + '-' + iso.slice(0, 4) : '');

  function build(ExcelJS, acc, when) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'IDCO Bank Statement to Excel v' + VERSION;
    wb.created = when || new Date();

    // ------------------------------------------------------------ Statement
    const ws = wb.addWorksheet('Statement', { views: [{ state: 'frozen', ySplit: 1 }] });
    const cols = [
      ['Sl No', 7], ['Date', 12], ['Value Date', 12], ['Narration', 70], ['Cheque / Ref No', 16], ['Withdrawal', 18], ['Deposit', 18], ['Balance', 20],
      ['Type', 20], ['UTR / Reference', 28], ['Party (from narration)', 34], ['Other details', 16], ['Source file', 30], ['Page', 7], ['Check', 28]
    ];
    ws.columns = cols.map(c => ({ width: c[1] }));
    const head = ws.addRow(cols.map(c => c[0]));
    head.height = 22;
    head.eachCell(c => {
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } };
      c.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    acc.rows.forEach((r, i) => {
      const note = (r.gapBefore ? r.gapBefore + ' ' : '') + (r.ok ? '' : r.note);
      const row = ws.addRow([
        i + 1, toDate(r.date), toDate(r.valueDate), r.narration, r.chq || '', r.withdrawal ? rupees(r.withdrawal) : null, r.deposit ? rupees(r.deposit) : null, rupees(r.balance),
        r.type || '', r.ref || '', r.party || '', r.other || '', r.file || '', r.page, note ? note.trim() : 'OK'
      ]);
      if (note) row.getCell(15).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
    });
    const last = acc.rows.length + 1;
    ws.getColumn(2).numFmt = 'dd-mm-yyyy'; ws.getColumn(3).numFmt = 'dd-mm-yyyy';
    [6, 7, 8].forEach(c => { ws.getColumn(c).numFmt = MONEY; });
    [5, 10].forEach(c => { ws.getColumn(c).numFmt = '@'; });
    [1, 2, 3, 14].forEach(c => { ws.getColumn(c).alignment = { horizontal: 'center' }; });
    head.eachCell(c => { c.alignment = { vertical: 'middle', horizontal: 'center' }; });
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: last, column: cols.length } };

    // ------------------------------------------------------------ Control
    const wc = wb.addWorksheet('Control', { views: [{ showGridLines: false }] });
    wc.columns = [{ width: 46 }, { width: 26 }, { width: 80 }];
    const t = wc.addRow(['Bank statement converted to Excel']); t.font = { bold: true, size: 14, color: { argb: TEAL } };
    wc.addRow([]);
    const res = wc.addRow(['Result', acc.pass ? 'PASS' : 'CHECK',
      acc.pass ? 'Every line agrees with the running balance and with the figures printed by the bank.' : 'One or more checks did not agree. See the Checks list below and the Check column of the Statement sheet.']);
    res.getCell(1).font = { bold: true, size: 12 };
    res.getCell(2).font = { bold: true, size: 12 }; res.getCell(2).alignment = { horizontal: 'center' };
    res.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: acc.pass ? GREEN : RED } };
    res.getCell(3).alignment = { wrapText: true, vertical: 'top' };
    wc.addRow([]);
    const kv = (k, v, fmt) => { const r = wc.addRow([k, v]); r.getCell(1).font = { color: { argb: 'FF4B5563' } }; r.getCell(2).alignment = { horizontal: 'right' }; if (fmt) r.getCell(2).numFmt = fmt; return r; };
    kv('Bank', acc.bank);
    kv('Account number (as printed)', String(acc.account || 'not found'));
    kv('IFSC', acc.ifsc || '');
    kv('Period covered', nice(acc.firstDate) + ' to ' + nice(acc.lastDate));
    kv('Number of lines', acc.rows.length);
    const rOpen = kv('Opening balance', rupees(acc.opening), MONEY);
    const rWd = kv('Total withdrawals (' + acc.totals.nWithdrawals + ' lines)', rupees(acc.totals.withdrawals), MONEY);
    const rDep = kv('Total deposits (' + acc.totals.nDeposits + ' lines)', rupees(acc.totals.deposits), MONEY);
    const rClose = kv('Closing balance', rupees(acc.closing), MONEY);
    wc.addRow([]);
    const sec = name => { const r = wc.addRow([name]); r.font = { bold: true, color: { argb: 'FFFFFFFF' } }; [1, 2, 3].forEach(c => { r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } }; }); return r; };

    sec('Live check (recalculates if the Statement sheet is edited)');
    const f1 = wc.addRow(['Sum of the Withdrawal column', { formula: 'SUM(Statement!F2:F' + last + ')', result: rupees(acc.totals.withdrawals) }]);
    const f2 = wc.addRow(['Sum of the Deposit column', { formula: 'SUM(Statement!G2:G' + last + ')', result: rupees(acc.totals.deposits) }]);
    const calc = (acc.opening === null ? 0 : acc.opening) + acc.totals.deposits - acc.totals.withdrawals;
    const f3 = wc.addRow(['Opening + deposits - withdrawals', { formula: 'B' + rOpen.number + '+B' + f2.number + '-B' + f1.number, result: rupees(calc) }]);
    const f4 = wc.addRow(['Difference from the closing balance (should be 0.00)', { formula: 'ROUND(B' + f3.number + '-B' + rClose.number + ',2)', result: rupees(calc - (acc.closing || 0)) }]);
    [f1, f2, f3, f4].forEach(r => { r.getCell(2).numFmt = MONEY; r.getCell(2).alignment = { horizontal: 'right' }; });
    wc.addRow([]);

    sec('Checks');
    const hdr = wc.addRow(['What was checked', 'Result', 'Detail']); hdr.font = { bold: true };
    const addCheck = (name, ok, detail) => {
      const r = wc.addRow([name, ok ? 'PASS' : 'CHECK', detail || '']);
      r.getCell(2).alignment = { horizontal: 'center' };
      r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ok ? GREEN : RED } };
      r.getCell(3).alignment = { wrapText: true, vertical: 'top' }; r.getCell(1).alignment = { wrapText: true, vertical: 'top' };
    };
    const many = acc.files.length > 1;
    for (const f of acc.files) for (const c of f.checks) addCheck((many ? f.fileName + ': ' : '') + c.name, c.ok, c.detail);
    if (many) {
      for (const j of acc.joins) addCheck('Joining the statements', j.ok, j.text);
      addCheck('All statements together: opening + deposits - withdrawals = closing', acc.arithmeticOk, acc.gap ? 'Not meaningful while a statement is missing.' : '');
    }
    wc.addRow([]);

    sec('Statement files used');
    const fh = wc.addRow(['File', 'Period', 'Opening balance  /  closing balance  /  lines  /  note']); fh.font = { bold: true };
    for (const f of acc.files) {
      const r = wc.addRow([f.fileName, nice(f.firstDate) + ' to ' + nice(f.lastDate),
        fmtIN(f.opening) + '  /  ' + fmtIN(f.closing) + '  /  ' + f.lines + ' lines' + (f.used !== f.lines ? ' (' + f.used + ' used)' : '') + (f.note ? '  /  ' + f.note : '')]);
      r.getCell(3).alignment = { wrapText: true, vertical: 'top' }; r.getCell(1).alignment = { wrapText: true, vertical: 'top' };
    }
    const warns = []; acc.files.forEach(f => (f.warnings || []).forEach(w => warns.push((many ? f.fileName + ': ' : '') + w)));
    if (warns.length) {
      wc.addRow([]); sec('Notes');
      for (const w of warns) { const r = wc.addRow([w]); wc.mergeCells(r.number, 1, r.number, 3); r.getCell(1).alignment = { wrapText: true, vertical: 'top' }; r.height = 30; r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } }; }
    }
    wc.addRow([]);
    const n1 = wc.addRow(['Dates, amounts and balances are read from the statement and proved against the running balance. Type, UTR / Reference and Party are helper columns picked out of the narration and should be treated as a guide.']);
    wc.mergeCells(n1.number, 1, n1.number, 3); n1.getCell(1).alignment = { wrapText: true, vertical: 'top' }; n1.height = 32; n1.getCell(1).font = { italic: true, color: { argb: 'FF4B5563' } };
    const n2 = wc.addRow(['Made by IDCO Bank Statement to Excel v' + VERSION + ' on ' + (when || new Date()).toLocaleString('en-IN') + '. The statement was read on this computer; nothing was uploaded.']);
    wc.mergeCells(n2.number, 1, n2.number, 3); n2.getCell(1).font = { italic: true, color: { argb: 'FF6B7280' } };

    // ------------------------------------------------------------ By Type
    const wt = wb.addWorksheet('By Type', { views: [{ state: 'frozen', ySplit: 1 }] });
    wt.columns = [{ width: 28 }, { width: 10 }, { width: 22 }, { width: 22 }];
    const th = wt.addRow(['Type', 'Lines', 'Withdrawals', 'Deposits']);
    th.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } }; c.alignment = { horizontal: 'center' }; });
    const by = new Map();
    for (const r of acc.rows) { const k = r.type || 'Other'; const o = by.get(k) || { n: 0, wd: 0, dep: 0 }; o.n++; o.wd += r.withdrawal || 0; o.dep += r.deposit || 0; by.set(k, o); }
    const types = Array.from(by.entries()).sort((a, b) => b[1].n - a[1].n);
    types.forEach(([k, o], i) => {
      const n = i + 2;
      wt.addRow([k, { formula: 'COUNTIF(Statement!$I$2:$I$' + last + ',A' + n + ')', result: o.n },
        { formula: 'SUMIF(Statement!$I$2:$I$' + last + ',A' + n + ',Statement!$F$2:$F$' + last + ')', result: rupees(o.wd) },
        { formula: 'SUMIF(Statement!$I$2:$I$' + last + ',A' + n + ',Statement!$G$2:$G$' + last + ')', result: rupees(o.dep) }]);
    });
    const end = types.length + 1;
    const tot = wt.addRow(['Total', { formula: 'SUM(B2:B' + end + ')', result: acc.rows.length }, { formula: 'SUM(C2:C' + end + ')', result: rupees(acc.totals.withdrawals) }, { formula: 'SUM(D2:D' + end + ')', result: rupees(acc.totals.deposits) }]);
    tot.font = { bold: true }; tot.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREY } }; });
    wt.getColumn(3).numFmt = MONEY; wt.getColumn(4).numFmt = MONEY; wt.getColumn(2).alignment = { horizontal: 'center' };

    return wb;
  }

  function fmtIN(p) {
    if (p === null || p === undefined) return '';
    const neg = p < 0; p = Math.abs(p);
    return (neg ? '-' : '') + Math.floor(p / 100).toLocaleString('en-IN') + '.' + String(p % 100).padStart(2, '0');
  }

  return { build, VERSION };
});
