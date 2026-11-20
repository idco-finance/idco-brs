/* IDCO Bank Statement to Excel - the last signed BRS, read as the opening position.
 *
 * IDCO's yearly BRS workbook has one summary sheet ("BRS") with nine accounts side by side per block
 * (GL code, A/c no, balance as per Bank Book, the four headings, balance as per bank statement), and one
 * detail sheet per account listing the items under each heading.
 * read(grid, account) finds the account by its number, reads its figures and its items, and checks that
 *   book + adds - lesses = bank, and that every heading's items add up to the summary figure.
 * A grid is { names: [sheet names], rows(name) -> array of rows of cell values (Date for dates) }.
 * All money in paise. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementOpening = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const HEAD = {
    ISSUED: { label: 'Add: Cheque issued but not presented for payment', sign: +1, side: 'book' },
    BANK_CR: { label: 'Add: Amount credited by bank only', sign: +1, side: 'bank' },
    DEPOSITED: { label: 'Less: Cheque deposited but not credited by bank', sign: -1, side: 'book' },
    BANK_DR: { label: 'Less: Amount debited by bank only', sign: -1, side: 'bank' }
  };
  const ORDER = ['ISSUED', 'BANK_CR', 'DEPOSITED', 'BANK_DR'];

  const s = v => (v === null || v === undefined ? '' : v instanceof Date ? v.toISOString().slice(0, 10) : String(v)).replace(/\s+/g, ' ').trim();
  const digits = v => s(v).replace(/\.0+$/, '').replace(/\D/g, '');
  const paise = v => { if (typeof v === 'number') return Math.round(v * 100); const t = s(v).replace(/,/g, ''); return /^-?\d+(\.\d+)?$/.test(t) ? Math.round(parseFloat(t) * 100) : null; };
  function headOf(text) {
    const t = text.toLowerCase();
    if (/issued\s+but\s+not\s+(p\w*sented|debited)/.test(t)) return 'ISSUED';
    if (/deposited\s+but\s+not\s+credited/.test(t)) return 'DEPOSITED';
    if (/debited\s+by\s+bank\s+only/.test(t)) return 'BANK_DR';
    if (/cre?dited\s+by\s+bank\s+only/.test(t)) return 'BANK_CR';
    return '';
  }
  function dateOf(v) {
    if (v instanceof Date && !isNaN(v)) return new Date(v.getTime() + 12 * 3600e3).toISOString().slice(0, 10);   // noon, so time zones cannot shift the day
    const m = s(v).match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
    if (!m) return '';
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    return y + '-' + String(+m[2]).padStart(2, '0') + '-' + String(+m[1]).padStart(2, '0');
  }

  /** Is this workbook an IDCO BRS workbook? */
  function isBRS(grid) {
    const name = grid.names.find(n => /^\s*brs\s*$/i.test(n)) || grid.names[0];
    const rows = grid.rows(name).slice(0, 12);
    return rows.some(r => r.some(c => /bank\s+reconciliation\s+statement/i.test(s(c)))) && rows.some(r => /^a\/c/i.test(s(r[0])));
  }

  /** All accounts in the summary sheet: [{name, gl, account, col, block}]. */
  function accounts(grid) {
    const name = grid.names.find(n => /^\s*brs\s*$/i.test(n)) || grid.names[0];
    const rows = grid.rows(name); const out = [];
    for (let r = 0; r < rows.length; r++) {
      if (!/^particulars/i.test(s(rows[r][0]))) continue;
      const lab = {};
      for (let q = r + 1; q < rows.length && !/^particulars|^bank reconciliation/i.test(s(rows[q][0])); q++) lab[q] = s(rows[q][0]);
      const find = re => +(Object.keys(lab).find(q => re.test(lab[q])) || -1);
      const rGl = find(/^gl/i), rAc = find(/^a\/c/i), rBook = find(/^balance as per bank book/i), rBank = find(/^balance as per bank statement/i);
      for (let c = 1; c < rows[r].length; c++) {
        const nm = s(rows[r][c]); if (!nm) continue;
        const heads = {};
        Object.keys(lab).forEach(q => { const h = headOf(lab[q]); if (h) heads[h] = paise(rows[q][c]) || 0; });
        out.push({ name: nm.replace(/\s*\(\d+\)\s*$/, ''), gl: rGl >= 0 ? s(rows[rGl][c]).replace(/\.0+$/, '') : '', account: rAc >= 0 ? digits(rows[rAc][c]) : '',
          book: rBook >= 0 ? paise(rows[rBook][c]) : null, bank: rBank >= 0 ? paise(rows[rBank][c]) : null, heads,
          asOn: rBook >= 0 ? dateOf((lab[rBook].match(/(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})/) || [])[1]) : '' });
      }
    }
    return out;
  }

  const last4 = a => String(a || '').replace(/\D/g, '').slice(-4);

  /** The opening position of one account. */
  function read(grid, account, fileName) {
    const all = accounts(grid);
    const acc = String(account || '').replace(/\D/g, '');
    let a = all.find(x => x.account && x.account === acc);
    if (!a) { const m = all.filter(x => x.account && acc && x.account.slice(-4) === last4(acc) && (x.account.length < 9 || acc.indexOf(x.account) >= 0 || x.account.indexOf(acc) >= 0)); if (m.length === 1) a = m[0]; }
    if (!a) return null;
    const out = { fileName, account: a.account, gl: a.gl, name: a.name, asOn: a.asOn, book: a.book, bank: a.bank, heads: a.heads, items: [], checks: [], sheet: '' };
    const ok = (pass, what, detail) => out.checks.push({ pass, what, detail });
    const f = p => (p < 0 ? '-' : '') + (Math.abs(p) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d)\.)/g, ',').replace(/(\d)(?=(\d\d)+\d{3}\.)/g, '$1,');
    // detail sheet: named after the GL code, else the one whose text carries the account number
    const glSheet = acc2 => acc2.gl ? grid.names.find(n => n.replace(/\D/g, ' ').split(' ').includes(acc2.gl)) : '';
    let sheet = glSheet(a);
    if (!sheet) sheet = grid.names.find(n => !/^\s*brs\s*$/i.test(n) && grid.rows(n).slice(0, 4).some(r => r.some(c => digits(c) === a.account)));
    if (!sheet) {                       // sheets named after the bank and kind of account ("ICICI PAYMENT", "AXIS PAYMENT", "SBI IDCO TOWER")
      const words = t => s(t).toUpperCase().replace(/S\.B\.I/g, 'SBI').replace(/[^A-Z0-9]+/g, ' ').split(' ').filter(w => w.length > 1 && !/^(BANK|AC|A|C|BR|BRANCH|THE|OF|FOR|COLL)$/.test(w));
      const taken = new Set(all.map(glSheet).filter(Boolean));
      const fits = (n, x) => { const w = words(n); const nm = words(x.name); return w.length && w.every(t => nm.includes(t) || (t === 'COLLECTION' && nm.includes('COLLECTION'))) ? nm.length - w.length : -1; };
      const cands = grid.names.filter(n => !/^\s*brs\s*$/i.test(n) && !taken.has(n) && fits(n, a) >= 0);
      // a sheet belongs to the account it fits most closely
      sheet = cands.find(n => { const mine = fits(n, a); return !all.some(x => x !== a && !glSheet(x) && fits(n, x) >= 0 && fits(n, x) < mine); }) || '';
    }
    const sumHeads = ORDER.reduce((t, h) => t + HEAD[h].sign * (a.heads[h] || 0), 0);
    ok(a.book !== null && a.bank !== null && a.book + sumHeads === a.bank, 'Book balance + additions - deductions = bank balance (summary sheet)',
      'Book ' + f(a.book || 0) + ', headings ' + f(sumHeads) + ', bank ' + f(a.bank || 0) + '.');
    const need = ORDER.filter(h => a.heads[h]);
    if (!sheet) { ok(!need.length, 'Detail sheet of this account found', need.length ? 'No sheet named after GL ' + a.gl + ' was found, so the items behind the headings cannot be listed.' : 'Nothing to list.'); return out; }
    out.sheet = sheet;
    const rows = grid.rows(sheet); let head = '', col = null; const total = {};
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r]; const texts = row.map(s);
      const joined = texts.filter(Boolean).join(' ');
      if (!joined) continue;
      const h = headOf(texts.find(t => headOf(t)) || '');
      if (h && !texts.some(t => /^total/i.test(t))) { head = h; col = null; continue; }
      if (texts.some(t => /amount/i.test(t)) && texts.some(t => /date|cheque|name/i.test(t))) {
        col = {}; texts.forEach((t, i) => { const l = t.toLowerCase();
          if (/balance/.test(l) && /amt|amount/.test(l)) col.amt = i;
          else if (/amount/.test(l) && col.amt === undefined) col.amt = i; else if (/date/.test(l) && col.date === undefined) col.date = i;
          else if (/name|party/.test(l) && col.party === undefined) col.party = i; else if (/urn/.test(l)) col.urn = i;
          else if (/cheque|req|d\.d/.test(l) && col.chq === undefined) col.chq = i; else if (/remark/.test(l)) col.remark = i; });
        continue;
      }
      if (!head || !col || col.amt === undefined) continue;
      if (texts.some(t => /^total/i.test(t))) { total[head] = (total[head] || 0) + (paise(row[col.amt]) || 0); continue; }
      const amt = paise(row[col.amt]); if (!amt) continue;
      const ref = [col.urn, col.chq, col.remark].filter(i => i !== undefined).map(i => s(row[i])).filter(Boolean).join(' ').replace(/\.0\b/g, '');
      const refs = (ref.match(/\d{9,}/g) || []).concat((s(row[col.party]).match(/[A-Z]{4}[A-Z0-9]?\d{8,}|\d{12,}/g) || []));
      out.items.push({ head, amt, date: col.date !== undefined ? dateOf(row[col.date]) : '', dateText: col.date !== undefined ? s(row[col.date]) : '',
        party: col.party !== undefined ? s(row[col.party]) : '', ref, refs: Array.from(new Set(refs)), row: r + 1, sheet, chq: col.chq !== undefined ? s(row[col.chq]).replace(/\.0$/, '') : '' });
    }
    for (const h of ORDER) {
      const items = out.items.filter(i => i.head === h), sum = items.reduce((t, i) => t + i.amt, 0), want = a.heads[h] || 0;
      if (!want && !items.length) continue;
      ok(sum === want, HEAD[h].label.replace(/^(Add|Less): /, '') + ': items add up to the summary figure',
        items.length + ' items, ' + f(sum) + '; summary sheet ' + f(want) + (total[h] !== undefined && total[h] !== sum ? '; TOTAL row ' + f(total[h]) : '') + '.');
    }
    out.items.forEach((it, i) => { it.no = i + 1; it.ikey = ('O|' + out.asOn + '|' + it.head + '|' + it.amt + '|' + (it.refs[0] || it.chq || it.party.slice(0, 20)) + '|' + it.dateText).replace(/\s+/g, '_'); });
    { const seen = new Map(); out.items.forEach(it => { const n = (seen.get(it.ikey) || 0) + 1; seen.set(it.ikey, n); if (n > 1) it.ikey += '|' + n; }); }
    return out;
  }

  /** A grid from a SheetJS workbook (read with cellDates: true). */
  function fromSheetJS(XLSX, wb) {
    return { names: wb.SheetNames.slice(), rows: n => XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }) };
  }
  /** A grid from an ExcelJS workbook. */
  function fromExcelJS(wb) {
    const val = c => (c === null || c === undefined ? '' : c instanceof Date ? c : typeof c === 'object' ? (c.result !== undefined ? c.result : c.text !== undefined ? c.text : c.richText ? c.richText.map(x => x.text).join('') : '') : c);
    return { names: wb.worksheets.map(w => w.name), rows: n => { const out = []; wb.getWorksheet(n).eachRow({ includeEmpty: true }, (row, i) => { const a = []; row.eachCell({ includeEmpty: true }, (cell, c) => { a[c - 1] = val(cell.value); }); out[i - 1] = a; }); for (let i = 0; i < out.length; i++) out[i] = out[i] || []; return out; } };
  }

  return { HEAD, ORDER, isBRS, accounts, read, fromSheetJS, fromExcelJS };
});
