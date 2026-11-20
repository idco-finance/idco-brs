/* IDCO Bank Statement to Excel - the BRS as on a chosen date.
 *
 * matchOpening(rec, op)  ties the items of the last signed BRS (31 March) to the bank lines and vouchers of the
 *                        new year that clear them (same amount; reference found = Matched, else Suggested).
 * build(rec, op, acc, asOn)  the BRS as on asOn for the period 1 April .. asOn:
 *   - every bank line, voucher, Eazypay payment and opening item is put in a group with whatever it is matched to;
 *   - only the members dated on or before asOn count, so a receipt credited in June and booked in July is open
 *     as on 30 June and cleared as on 31 July;
 *   - what a group does not cancel is an open item, under IDCO's four headings (or as a difference in amounts);
 *   - sweeps are one net line;
 *   - balance as per Bank Book + additions - deductions must equal the balance as per bank statement. By
 *     construction it does, as long as the opening BRS agrees with the statement's opening balance; any
 *     difference is shown, never hidden.
 * All money in paise. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementBRS = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DAY = 86400000;
  const t = iso => Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10));
  const days = (a, b) => Math.round((t(a) - t(b)) / DAY);
  const fyStartOf = iso => (+iso.slice(5, 7) >= 4 ? iso.slice(0, 4) : String(+iso.slice(0, 4) - 1)) + '-04-01';
  const monthEnd = ym => { const d = new Date(Date.UTC(+ym.slice(0, 4), +ym.slice(5, 7), 0)); return d.toISOString().slice(0, 10); };
  const compact = x => String(x || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const HEADS = {
    ISSUED: 'Add: Cheque issued but not presented for payment',
    BANK_CR: 'Add: Amount credited by bank only',
    DIFF_PLUS: 'Add: Difference in matched entries (bank more than books)',
    DEPOSITED: 'Less: Cheque deposited but not credited by bank',
    BANK_DR: 'Less: Amount debited by bank only',
    DIFF_MINUS: 'Less: Difference in matched entries (books more than bank)',
    BOOKNET_PLUS: 'Add: Book entries with no bank entry this year - earlier-period receipts and suspense reversals (net)',
    BOOKNET_MINUS: 'Less: Book entries with no bank entry this year - earlier-period receipts and suspense reversals (net)'
  };
  const ORDER = ['ISSUED', 'BANK_CR', 'DIFF_PLUS', 'BOOKNET_PLUS', 'DEPOSITED', 'BANK_DR', 'DIFF_MINUS', 'BOOKNET_MINUS'];
  const SIGN = { ISSUED: 1, BANK_CR: 1, DIFF_PLUS: 1, BOOKNET_PLUS: 1, DEPOSITED: -1, BANK_DR: -1, DIFF_MINUS: -1, BOOKNET_MINUS: -1 };
  const OP_EFFECT = { BANK_CR: ['bank', 1], BANK_DR: ['bank', -1], DEPOSITED: ['book', 1], ISSUED: ['book', -1] };
  const OP_NAME = { BANK_CR: 'Amount credited by bank only', BANK_DR: 'Amount debited by bank only', DEPOSITED: 'Cheque deposited but not credited by bank', ISSUED: 'Cheque issued but not presented for payment' };

  /** Months for which a BRS can be made: April of the statement's year up to the last month covered by both sides. */
  function months(acc, rec) {
    const fy = fyStartOf(acc.from || acc.firstDate);
    const lastBank = acc.to || acc.lastDate, lastBook = rec.summary.lastBookFile || rec.summary.lastBook || '';
    const end = [lastBank, lastBook].filter(Boolean).sort()[0] || lastBank;
    const out = [];
    for (let y = +fy.slice(0, 4), m = 4; ; m++) { if (m > 12) { m = 1; y++; } const ym = y + '-' + String(m).padStart(2, '0'); const e = monthEnd(ym); if (e > end) break; out.push(e); if (out.length > 24) break; }
    return out;
  }

  // ------------------------------------------------------------------------------------- opening items
  function matchOpening(rec, op) {
    const B = rec.bankLines, V = rec.vouchers, fmt = rec.fmt, nice = rec.nice;
    const asOn = op.asOn || '';
    rec.opening = op.items.map(it => Object.assign({}, it, { status: 'Open', openLinks: [], how: '', note: '' }));
    const O = rec.opening;
    const freeB = b => (b.status === 'Not in books' || (b.status === 'Needs review' && /^EZY_/.test(b.code || ''))) && !(b.openLinks && b.openLinks.length);
    const freeV = v => (v.status === 'Not in bank' || v.status === 'Earlier period') && !(v.openLinks && v.openLinks.length);
    const tie = (o, x, status, how) => {
      o.status = status; o.openLinks.push(x); o.how = how;
      x.openLinks = (x.openLinks || []).concat([o]);
      if (x.status !== 'Matched') x.status = status === 'Matched' ? 'Matched' : 'Suggested';
      x.how = how; x.opMatch = true;
    };
    const desc = o => fmt(o.amt) + (o.party ? ' (' + o.party.slice(0, 60) + (o.dateText ? ', ' + o.dateText : '') + ')' : o.dateText ? ' (' + o.dateText + ')' : '');
    const refIn = (o, text) => o.refs.concat(o.chq && o.chq.length >= 5 ? [o.chq] : []).some(r => r.length >= 5 && compact(text).indexOf(compact(r)) >= 0);
    const whereB = b => 'Line ' + b.sl + ' of the Bank lines sheet. In the bank statement: ' + (b.file ? b.file + ', ' : '') + 'page ' + b.page + '. In the last BRS: sheet "' + op.sheet + '".';
    const whereV = v => 'ERP voucher ' + v.blockId + ' dated ' + nice(v.vchDate) + '. Bank Book file: ' + v.file + '. In the last BRS: sheet "' + op.sheet + '".';
    const brsName = 'the BRS as on ' + nice(asOn);

    // 1. one entry clears one item: same amount, right side, reference found = Matched, else the nearest date = Suggested
    for (const o of O) {
      const [side, sg] = OP_EFFECT[o.head];
      let cands;
      if (o.head === 'DEPOSITED') cands = B.filter(b => freeB(b) && b.amt === o.amt && (!o.date || b.date >= o.date));
      else if (o.head === 'ISSUED') cands = B.filter(b => freeB(b) && b.amt === -o.amt && (!o.date || b.date >= o.date));
      else if (o.head === 'BANK_CR') cands = V.filter(v => freeV(v) && v.net === o.amt).concat(B.filter(b => freeB(b) && b.amt === -o.amt));
      else cands = V.filter(v => freeV(v) && v.net === -o.amt).concat(B.filter(b => freeB(b) && b.amt === o.amt));
      if (!cands.length) continue;
      const isB = x => x.sl !== undefined;
      const text = x => (isB(x) ? x.narration : [x.utr, x.narration, x.urnText, x.party].join(' '));
      const strong = cands.filter(x => refIn(o, text(x)));
      const dateOf = x => (isB(x) ? x.date : x.vchDate);
      const pick = (strong.length ? strong : cands).slice().sort((a, b) => dateOf(a).localeCompare(dateOf(b)))[0];
      const how = 'Clears an item of ' + brsName + ': ' + OP_NAME[o.head] + (strong.length ? ' (reference agrees)' : ' (same amount)');
      tie(o, pick, strong.length ? 'Matched' : 'Suggested', how);
      if (!strong.length) {
        const many = cands.length > 1 ? ' ' + cands.length + ' entries of this amount were found; the earliest is shown.' : '';
        if (isB(pick)) {
          pick.found = 'The bank ' + (pick.amt > 0 ? 'credited ' : 'debited ') + fmt(Math.abs(pick.amt)) + ' on ' + nice(pick.date) + (pick.party ? ' (' + pick.party + ')' : '') + '. ' + brsName.replace(/^t/, 'T') + ' lists ' + desc(o) + ' under "' + OP_NAME[o.head] + '". The amounts agree.' + many;
          pick.todo = 'If this bank line is that item of the last BRS, choose Accept: the item is then cleared. If not, choose Reject.';
          pick.where = whereB(pick);
        } else {
          pick.found = 'Voucher ' + pick.blockId + ' of ' + nice(pick.vchDate) + ' shows ' + fmt(Math.abs(pick.net)) + '. ' + brsName.replace(/^t/, 'T') + ' lists ' + desc(o) + ' under "' + OP_NAME[o.head] + '". The amounts agree.' + many;
          pick.todo = 'If this voucher books that item of the last BRS, choose Accept: the item is then cleared. If not, choose Reject.';
          pick.where = whereV(pick);
        }
      }
    }
    // 2. Eazypay credits of early April whose report rows were not given: March-end receipts listed in the last BRS
    //    (URN dated a few days before) that add up exactly to the part of the credit not explained by a report
    for (const b of B.filter(b => b.type === 'Eazypay settlement' && /^EZY_/.test(b.code || '') && !(b.openLinks && b.openLinks.length))) {
      const known = rec.payments.filter(m => m.bankSl === b.sl).reduce((a, m) => a + m.amt, 0);
      const target = b.amt - known; if (target <= 0) continue;
      const urnDate = o => { const u = o.refs.find(r => /^\d{15}$/.test(r)); if (!u) return ''; return '20' + u.slice(0, 2) + '-' + u.slice(2, 4) + '-' + u.slice(4, 6); };
      const cands = O.filter(o => o.status === 'Open' && o.head === 'DEPOSITED' && urnDate(o) && days(b.date, urnDate(o)) >= 0 && days(b.date, urnDate(o)) <= 6);
      const pickSet = subsetSum(cands, target, 22);
      if (!pickSet) continue;
      const how = 'Clears ' + pickSet.length + ' item(s) of ' + brsName + ' (Eazypay receipts of late March that add up to this credit)';
      pickSet.forEach(o => { o.status = 'Suggested'; o.openLinks.push(b); o.how = how; });
      b.openLinks = (b.openLinks || []).concat(pickSet); b.opMatch = true;
      b.status = 'Suggested'; b.how = how; b.code = '';
      b.found = 'The bank credited ' + fmt(b.amt) + ' on ' + nice(b.date) + ' as one Eazypay amount.' + (known ? ' The Eazypay report explains ' + fmt(known) + ' of it.' : ' Its Eazypay report was not given.') + ' ' + brsName.replace(/^t/, 'T') + ' lists ' + pickSet.length + ' Eazypay receipts of late March (URN ' + pickSet.map(o => o.refs[0]).join(', ') + ') that add up to exactly ' + fmt(target) + '.';
      b.todo = 'These are probably the March-end payments settled by the bank on this day. If you can, check with the Eazypay report of those dates; then choose Accept.';
      b.where = whereB(b);
    }
    return rec;
  }

  // exact subset of items whose amounts add up to target (meet in the middle; n <= max)
  function subsetSum(items, target, max) {
    if (!items.length || items.length > max) return null;
    const exact = items.filter(o => o.amt === target); if (exact.length === 1) return exact;
    const h = items.length >> 1, A = items.slice(0, h), Bs = items.slice(h);
    const left = new Map();
    for (let m = 0; m < (1 << A.length); m++) { let s = 0; for (let i = 0; i < A.length; i++) if (m & (1 << i)) s += A[i].amt; if (s <= target && !left.has(s)) left.set(s, m); }
    for (let m = 0; m < (1 << Bs.length); m++) {
      let s = 0; for (let i = 0; i < Bs.length; i++) if (m & (1 << i)) s += Bs[i].amt;
      if (s > target || !left.has(target - s)) continue;
      const lm = left.get(target - s); const out = [];
      A.forEach((o, i) => { if (lm & (1 << i)) out.push(o); }); Bs.forEach((o, i) => { if (m & (1 << i)) out.push(o); });
      if (out.length) return out;
    }
    return null;
  }

  // ------------------------------------------------------------------------------------------------- the BRS
  function build(rec, op, acc, asOn) {
    const B = rec.bankLines, V = rec.vouchers, P = rec.payments, O = rec.opening || [], fmt = rec.fmt, nice = rec.nice;
    const fy = fyStartOf(asOn);
    const checks = []; const ok = (pass, what, detail) => checks.push({ pass, what, detail });
    // what has to be true before a BRS can be stated
    const stFrom = acc.from || acc.firstDate, stTo = acc.to || acc.lastDate;
    ok(stFrom <= fy, 'Bank statements start on ' + nice(fy), 'The statements given start on ' + nice(stFrom) + '.');
    ok(stTo >= asOn, 'Bank statements reach ' + nice(asOn), 'The statements given end on ' + nice(stTo) + '.');
    ok(!acc.gap && acc.pass, 'Bank statements are complete and proved', acc.pass ? 'Every line agrees with the running balance; no month is missing.' : 'The statement check is CHECK; see the statement Excel.');
    const lastBook = rec.summary.lastBookFile || rec.summary.lastBook || '';
    ok(lastBook >= asOn, 'Bank Book reaches ' + nice(asOn), 'The Bank Book files given end on ' + nice(lastBook || '-') + '.');
    ok(!!op, 'Last signed BRS given', op ? 'BRS as on ' + nice(op.asOn) + ' (' + op.fileName + ', sheet "' + op.sheet + '").' : 'Drop the last signed BRS workbook (for example the 31.03.2026 BRS) so that the opening items and the book balance are known.');
    if (op) {
      ok(op.asOn === new Date(t(fy) - DAY).toISOString().slice(0, 10), 'The last BRS is as on the day before the year starts', 'It is as on ' + nice(op.asOn) + '.');
      ok(op.bank === acc.opening, 'Bank balance in the last BRS = opening balance of the statement', 'Last BRS ' + fmt(op.bank) + ', statement opening on ' + nice(stFrom) + ' ' + fmt(acc.opening) + '.');
      op.checks.forEach(c => checks.push(c));
    }
    if (checks.some(c => !c.pass) && (!op || stFrom > fy || stTo < asOn)) return { asOn, fy, checks, ready: false };

    // groups
    const parent = new Map(); const find = x => { while (parent.get(x) !== x) { parent.set(x, parent.get(parent.get(x))); x = parent.get(x); } return x; };
    const add = x => { if (!parent.has(x)) parent.set(x, x); };
    const join = (a, b) => { add(a); add(b); const ra = find(a), rb = find(b); if (ra !== rb) parent.set(ra, rb); };
    const vByKey = new Map(V.map(v => [v.key, v]));
    B.forEach(add); V.forEach(add); P.forEach(add); O.forEach(add);
    // a match only counts once it is confirmed: a pair the tool merely suggests stays open on both sides until accepted
    const sure = (x, y) => x.status !== 'Suggested' && y.status !== 'Suggested';
    V.forEach(v => { v.bank.forEach(sl => { if (sure(v, B[sl - 1])) join(v, B[sl - 1]); }); (v.setOff || []).forEach(x => join(v, x)); (v.openLinks || []).forEach(o => { if (sure(v, o)) join(v, o); }); });
    B.forEach(b => { (b.setOff || []).forEach(x => join(b, x)); (b.openLinks || []).forEach(o => { if (sure(b, o)) join(b, o); }); });
    P.forEach(m => { if (m.bankSl && B[m.bankSl - 1].status !== 'Suggested') join(m, B[m.bankSl - 1]); m.vouchers.forEach(k => { const v = vByKey.get(k); if (v) join(m, v); }); });
    // receipts whose money came in before 1 April (booked this year) and reversals of a suspense entry of an earlier year:
    // book entries with no bank entry of their own this year. They are shown together, as one net line with a schedule.
    const ids = new Set(V.map(v => v.blockId));
    const BOOKNET = { bookNet: true };
    const lone = v => !v.bank.length && !(v.openLinks && v.openLinks.length) && !(v.setOff && v.setOff.length) && !P.some(m => m.vouchers.includes(v.key));
    const netMembers = V.filter(v => lone(v) && (v.status === 'Earlier period' || (v.status === 'Suspense entry' && !(v.suspenseRef && ids.has(v.suspenseRef.replace(/\s/g, '')))))) ;
    if (netMembers.length) { add(BOOKNET); netMembers.forEach(v => join(v, BOOKNET)); }
    const groups = new Map(); for (const x of parent.keys()) { const r = find(x); if (!groups.has(r)) groups.set(r, []); groups.get(r).push(x); }

    const isB = x => x.sl !== undefined && x.narration !== undefined, isV = x => x.blockId !== undefined, isO = x => x.head !== undefined, isP = x => x.tran !== undefined;
    const inB = b => b.date >= fy && b.date <= asOn, inV = v => v.vchDate >= fy && v.vchDate <= asOn;
    const items = []; let sweep = 0, sweepN = 0, suggested = 0, bookNet = null; const clearedO = new Set();
    const later = g => {
      const out = [];
      g.filter(x => isB(x) && x.date > asOn).forEach(b => out.push((b.amt > 0 ? 'Credited' : 'Debited') + ' by the bank on ' + nice(b.date) + ' (line ' + b.sl + ')'));
      g.filter(x => isV(x) && x.vchDate > asOn).forEach(v => out.push('Booked on ' + nice(v.vchDate) + ' in ' + v.blockId));
      return out.slice(0, 4).join('; ') + (out.length > 4 ? '; and ' + (out.length - 4) + ' more' : '');
    };
    const pushB = (b, g, kind) => items.push({ head: b.amt > 0 ? 'BANK_CR' : 'BANK_DR', amt: Math.abs(b.amt), date: b.date, who: b.party || b.type, ref: b.ref || '', what: b.narration, source: 'Bank line ' + b.sl, kind: kind || b.kind || '', status: b.status, later: later(g), keys: [b.ikey || b.key] });
    const pushV = (v, g) => items.push({ head: v.net > 0 ? 'DEPOSITED' : 'ISSUED', amt: Math.abs(v.net), date: v.vchDate, who: v.paidBy || v.party, ref: v.urnText || v.utr || '', what: v.blockId + (v.narration ? ' - ' + v.narration : ''), source: 'Voucher ' + v.blockId, kind: v.kind || '', status: v.status, later: later(g), keys: [v.ikey] });
    const pushO = (o, g) => items.push({ head: o.head, amt: o.amt, date: o.date, who: o.party, ref: o.ref, what: 'Item ' + o.no + ' of the BRS as on ' + nice(op.asOn), source: 'Last BRS', kind: 'Open since the last BRS', status: o.status === 'Open' ? 'Open' : o.status, later: later(g), keys: [o.ikey] });

    for (const g of groups.values()) {
      if (g.length === 1 && isB(g[0]) && g[0].status === 'Sweep') { if (inB(g[0])) { sweep += g[0].amt; sweepN++; } continue; }
      if (g.includes(BOOKNET)) {
        const mem = g.filter(x => isV(x) && inV(x)); const net = mem.reduce((a, v) => a + v.net, 0);
        bookNet = { members: mem.map(v => ({ head: v.status === 'Earlier period' ? 'Receipt of an earlier period' : 'Suspense entry / its reversal', date: v.vchDate, id: v.blockId, who: v.paidBy || v.party, amt: v.net, instr: v.instrDate, what: v.narration, key: v.ikey })), net };
        if (net) items.push({ head: net > 0 ? 'BOOKNET_MINUS' : 'BOOKNET_PLUS', amt: Math.abs(net), date: mem.map(v => v.vchDate).sort()[0], who: mem.length + ' vouchers', ref: '', what: 'Receipts booked this year for money that came into the bank before 1 April (' + mem.filter(v => v.status === 'Earlier period').length + ' vouchers), and reversals of suspense entries of an earlier year (' + mem.filter(v => v.status !== 'Earlier period').map(v => v.blockId).join(', ') + '). See the sheet "BRS book entries".', source: 'Book entries', kind: 'Earlier period receipts and suspense reversals', status: '', later: '', keys: mem.map(v => v.ikey) });
        continue;
      }
      const bIn = g.filter(x => isB(x) && inB(x)), vIn = g.filter(x => isV(x) && inV(x)), oIn = g.filter(isO);
      const bankPart = bIn.reduce((a, b) => a + b.amt, 0) + oIn.filter(o => OP_EFFECT[o.head][0] === 'bank').reduce((a, o) => a + OP_EFFECT[o.head][1] * o.amt, 0);
      const bookPart = vIn.reduce((a, v) => a + v.net, 0) + oIn.filter(o => OP_EFFECT[o.head][0] === 'book').reduce((a, o) => a + OP_EFFECT[o.head][1] * o.amt, 0);
      const diff = bankPart - bookPart;
      const anyBank = bIn.length || oIn.some(o => OP_EFFECT[o.head][0] === 'bank'), anyBook = vIn.length || oIn.some(o => OP_EFFECT[o.head][0] === 'book');
      if (diff === 0 && (anyBank || anyBook)) { oIn.forEach(o => clearedO.add(o)); continue; }
      if (!anyBank && !anyBook) continue;
      if (!anyBook) { bIn.forEach(b => pushB(b, g)); oIn.forEach(o => pushO(o, g)); continue; }
      if (!anyBank) { vIn.forEach(v => pushV(v, g)); oIn.forEach(o => pushO(o, g)); continue; }
      // both sides present: first the Eazypay payments the bank has credited but nobody has booked yet
      let rest = diff;
      for (const m of g.filter(isP)) {
        const lump = m.bankSl ? B[m.bankSl - 1] : null; if (!lump || !inB(lump)) continue;
        const booked = m.vouchers.map(k => vByKey.get(k)).filter(v => v && inV(v));
        if (booked.length) continue;
        const bookedLater = m.vouchers.map(k => vByKey.get(k)).filter(v => v && v.vchDate > asOn);
        items.push({ head: 'BANK_CR', amt: m.amt, date: lump.date, who: (m.partyCode ? m.partyCode + ' ' : '') + m.party, ref: m.tran, what: 'Eazypay payment ' + m.tran + ' (' + (m.mode || '') + '), inside the bank credit of ' + nice(lump.date) + ' (line ' + lump.sl + ')', source: 'Eazypay ' + m.tran, kind: m.kind || 'Online collection (Eazypay), not in books', status: m.status, later: bookedLater.map(v => 'Booked on ' + nice(v.vchDate) + ' in ' + v.blockId).join('; '), keys: [m.ikey] });
        rest -= m.amt;
      }
      if (rest !== 0) {
        const mem = g.filter(x => (isB(x) && inB(x)) || (isV(x) && inV(x)) || isO(x));
        const label = mem.slice(0, 6).map(x => isB(x) ? 'line ' + x.sl : isV(x) ? x.blockId : 'last-BRS item ' + x.no).join(', ') + (mem.length > 6 ? ' and ' + (mem.length - 6) + ' more' : '');
        const first = mem.map(x => isB(x) ? x.date : isV(x) ? x.vchDate : x.date || '').filter(Boolean).sort()[0] || asOn;
        items.push({ head: rest > 0 ? 'DIFF_PLUS' : 'DIFF_MINUS', amt: Math.abs(rest), date: first, who: (vIn[0] && (vIn[0].paidBy || vIn[0].party)) || (bIn[0] && bIn[0].party) || '', ref: '', what: 'Matched entries do not agree in amount: ' + label + '. Bank ' + fmt(bankPart) + ', books ' + fmt(bookPart) + '.', source: 'Matched group', kind: 'Amount differs between bank and books', status: 'Matched', later: later(g), keys: mem.map(x => x.ikey || x.key) });
      }
    }
    items.forEach(it => { it.age = days(asOn, it.date || asOn); });
    suggested = items.filter(it => it.status === 'Suggested').length;
    items.sort((a, b) => ORDER.indexOf(a.head) - ORDER.indexOf(b.head) || String(a.date).localeCompare(String(b.date)) || b.amt - a.amt);
    items.forEach((it, i) => { it.no = i + 1; });

    const totals = {}; ORDER.forEach(h => { totals[h] = { n: 0, amt: 0 }; }); items.forEach(it => { totals[it.head].n++; totals[it.head].amt += it.amt; });
    const bookMove = V.filter(inV).reduce((a, v) => a + v.net, 0);
    const book = op.book + bookMove;
    let bankBal = acc.opening; for (const r of acc.rows) if (r.date <= asOn) bankBal = r.balance;
    const computed = book + ORDER.reduce((a, h) => a + SIGN[h] * totals[h].amt, 0) + sweep;
    ok(computed === bankBal, 'Balance as per Bank Book + additions - deductions = balance as per bank statement',
      'Worked out ' + fmt(computed) + ', bank statement ' + fmt(bankBal) + (computed === bankBal ? '.' : '. Difference ' + fmt(bankBal - computed) + '.'));
    return { asOn, fy, ready: true, checks, items, totals, sweep, sweepN, book, bookOpen: op.book, bookMove, bank: bankBal, computed, diff: bankBal - computed, suggested, bookNet,
      openingCleared: Array.from(clearedO).map(o => o.ikey),
      opening: { asOn: op.asOn, n: O.length, cleared: clearedO.size, stillOpen: O.length - clearedO.size, file: op.fileName, sheet: op.sheet, gl: op.gl, name: op.name }, heads: HEADS, order: ORDER, sign: SIGN,
      account: acc.account, bankName: acc.bank };
  }

  return { HEADS, ORDER, SIGN, months, matchOpening, build, monthEnd, fyStartOf };
});
