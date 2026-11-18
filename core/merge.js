/* IDCO Bank Statement to Excel - joining several statements of one account.
 * Files are grouped by bank + account number, put in date order and joined.
 * The closing balance of one statement must equal the opening balance of the next; overlaps are removed, gaps are reported. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementMerge = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const SHORT = { 'ICICI Bank': 'ICICI', 'State Bank of India': 'SBI', 'Axis Bank': 'Axis', 'HDFC Bank': 'HDFC', 'Punjab National Bank': 'PNB', 'Canara Bank': 'Canara',
    'Indian Bank': 'IndianBank', 'UCO Bank': 'UCO', 'Bank of Baroda': 'BoB', 'Union Bank of India': 'UBI', 'Bank of India': 'BoI', 'Indian Overseas Bank': 'IOB',
    'Kotak Mahindra Bank': 'Kotak', 'Yes Bank': 'Yes', 'IDBI Bank': 'IDBI', 'Central Bank of India': 'CBI' };

  function fmt(p) {
    if (p === null || p === undefined) return '';
    const neg = p < 0; p = Math.abs(p);
    return (neg ? '-' : '') + Math.floor(p / 100).toLocaleString('en-IN') + '.' + String(p % 100).padStart(2, '0');
  }
  function niceDate(iso) { return iso ? iso.slice(8, 10) + '-' + iso.slice(5, 7) + '-' + iso.slice(0, 4) : ''; }

  function keyOf(st) {
    const digits = String(st.meta.account || '').replace(/\D/g, '');
    return (st.meta.bank || 'Bank not recognised') + '|' + (digits ? digits.slice(-4) : 'no-number');
  }

  /** statements: [{fileName, meta, opening, closing, rows, totals, checks, warnings, pass}] (already enriched) */
  function groupAndJoin(statements) {
    const groups = new Map();
    for (const st of statements) {
      if (!st.rows || !st.rows.length) continue;
      const k = keyOf(st);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(st);
    }
    const accounts = [];
    for (const [key, list] of groups) {
      list.sort((a, b) => (a.meta.firstDate || '').localeCompare(b.meta.firstDate || '') || (b.meta.lastDate || '').localeCompare(a.meta.lastDate || '') || (b.rows.length - a.rows.length));
      const acc = { key, bank: list[0].meta.bank || 'Bank not recognised', ifsc: '', account: '', files: [], rows: [], joins: [], warnings: [] };
      acc.short = SHORT[acc.bank] || acc.bank.replace(/[^A-Za-z]/g, '') || 'Bank';
      for (const st of list) {
        if (!acc.ifsc && st.meta.ifsc) acc.ifsc = st.meta.ifsc;
        if (st.meta.account && (!acc.account || (/\D/.test(acc.account) && /^\d+$/.test(st.meta.account)))) acc.account = st.meta.account;
      }
      let closing = null; let gap = false;
      for (const st of list) {
        const info = { fileName: st.fileName, from: st.meta.from, to: st.meta.to, firstDate: st.meta.firstDate, lastDate: st.meta.lastDate, opening: st.opening, closing: st.closing,
          lines: st.rows.length, pass: st.pass, checks: st.checks, warnings: st.warnings, used: st.rows.length, note: '' };
        let rows = st.rows;
        if (closing === null) { acc.opening = st.opening; }
        else if (st.opening === closing) { info.note = 'Continues exactly from the previous statement.'; acc.joins.push({ ok: true, text: st.fileName + ': opening balance ' + fmt(st.opening) + ' equals the previous closing balance.' }); }
        else {
          // does this statement overlap what we already have?
          let k = -1;
          for (let i = acc.rows.length - 1; i >= 0; i--) if (acc.rows[i].balance === st.opening && sameRun(acc.rows, i + 1, rows)) { k = i; break; }
          if (k < 0 && acc.opening === st.opening && sameRun(acc.rows, 0, rows)) k = -2;      // starts at the very same point
          if (k !== -1) {
            const start = k === -2 ? 0 : k + 1;
            const dup = Math.min(acc.rows.length - start, rows.length);
            rows = rows.slice(dup); info.used = rows.length;
            info.note = dup + ' line(s) were already in an earlier statement and were not repeated.';
            acc.joins.push({ ok: true, text: st.fileName + ': overlaps the previous statement; ' + dup + ' repeated line(s) removed' + (rows.length ? '.' : ' (nothing new in this file).') });
          } else {
            gap = true;
            info.note = 'GAP: previous closing balance ' + fmt(closing) + ', this statement opens with ' + fmt(st.opening) + '.';
            acc.joins.push({ ok: false, text: 'GAP before ' + st.fileName + ': the earlier statement closes at ' + fmt(closing) + ' but this one opens at ' + fmt(st.opening) + '. A statement in between is probably missing.' });
            if (rows.length) rows[0] = Object.assign({}, rows[0], { gapBefore: info.note });
          }
        }
        for (const r of rows) acc.rows.push(Object.assign({}, r, { file: st.fileName }));
        if (rows.length) closing = rows[rows.length - 1].balance;
        acc.files.push(info);
      }
      acc.closing = closing;
      let wd = 0, dep = 0, nWd = 0, nDep = 0;
      for (const r of acc.rows) { if (r.withdrawal) { wd += r.withdrawal; nWd++; } if (r.deposit) { dep += r.deposit; nDep++; } }
      acc.totals = { withdrawals: wd, deposits: dep, nWithdrawals: nWd, nDeposits: nDep };
      const ds = acc.rows.map(r => r.date).filter(Boolean).sort();
      acc.from = list.map(s => s.meta.from).filter(Boolean).sort()[0] || ds[0];
      acc.to = list.map(s => s.meta.to).filter(Boolean).sort().pop() || ds[ds.length - 1];
      acc.firstDate = ds[0]; acc.lastDate = ds[ds.length - 1];
      acc.badRows = acc.rows.filter(r => !r.ok).length;
      acc.arithmeticOk = !gap && acc.opening !== null && acc.opening + dep - wd === acc.closing;
      acc.gap = gap;
      acc.pass = acc.files.every(f => f.pass) && !gap && acc.arithmeticOk && acc.badRows === 0;
      const acctDigits = String(acc.account || '').replace(/\s+/g, '');
      acc.fileName = 'IDCO_Statement_' + acc.short + '_' + (acctDigits || 'account') + '_' + (acc.firstDate || 'from') + '_to_' + (acc.lastDate || 'to') + '.xlsx';
      accounts.push(acc);
    }
    return accounts;
  }

  function sameRun(have, start, rows) {
    const n = Math.min(have.length - start, rows.length);
    if (n <= 0) return start === have.length ? false : false;
    for (let i = 0; i < n; i++) {
      const a = have[start + i], b = rows[i];
      if (a.date !== b.date || a.balance !== b.balance || a.withdrawal !== b.withdrawal || a.deposit !== b.deposit) return false;
    }
    return true;
  }

  return { groupAndJoin, fmt, niceDate };
});
