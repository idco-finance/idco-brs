/* IDCO Bank Statement to Excel - the matching step of the BRS.
 *
 * Inputs : bank lines of one account (from the statement reader), Bank Book files (Linearizer), Eazypay files.
 * Output : every bank line and every voucher with a status, the evidence and a plain reason.
 * The tool makes no accounting judgement. It puts each item in one of three lists:
 *   Matched       - proved by a reference or by amount + date + party
 *   Suggested     - a likely match for a person to confirm
 *   Needs review  - with where to look and why
 * and it labels the things that are known not to match one-to-one (sweeps, earlier-period receipts, suspense reversals).
 * All money in paise. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementReconcile = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DAY = 86400000;
  const t = iso => (iso ? Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) : NaN);
  const days = (a, b) => Math.round((t(a) - t(b)) / DAY);
  const addDays = (iso, n) => new Date(t(iso) + n * DAY).toISOString().slice(0, 10);
  const compact = s => String(s || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  const nice = iso => (iso ? iso.slice(8, 10) + '-' + iso.slice(5, 7) + '-' + iso.slice(0, 4) : '');
  function fmt(p) { if (p === null || p === undefined) return ''; const n = p < 0; p = Math.abs(p); return (n ? '-' : '') + Math.floor(p / 100).toLocaleString('en-IN') + '.' + String(p % 100).padStart(2, '0'); }

  function realUtr(u) {
    const c = compact(u);
    if (c.length < 9) return false;
    if ((c.match(/\d/g) || []).length < 5) return false;
    if (/^(RTGS|NEFT|INSTR|IMPS|UTR)/.test(c) && c.length < 14) return false;
    return true;
  }
  const STOP = new Set(['PVT', 'LTD', 'LIMITED', 'PRIVATE', 'INDIA', 'AND', 'THE', 'IDCO', 'NEFT', 'RTGS', 'ODISHA', 'INDUSTRIES', 'INDUSTRIAL', 'COMPANY', 'CORPORATION', 'IPICOL', 'RECM', 'FOR', 'BANK', 'PAYMENT', 'ENTERPRISES', 'ENTERPRISE', 'PRODUCTS', 'UNIT', 'DIVISION', 'URGENT', 'ATTN', 'INB', 'TRANSFER', 'FROM', 'DEP', 'TFR']);
  const tokens = s => { const o = new Set(); (String(s || '').toUpperCase().match(/[A-Z]{3,}/g) || []).forEach(w => { if (!STOP.has(w)) o.add(w); }); return o; };
  function nameMatch(a, b) {
    const A = tokens(a), B = tokens(b); let n = 0;
    // names run together in the bank narration ("TATASTEELLIMITED"), and initials ("JTRFPL")
    const words = (String(a || '').toUpperCase().match(/[A-Z]{2,}/g) || []);
    const cb = compact(b);
    const sig = words.filter(w => !STOP.has(w));
    if (sig.length >= 2 && (sig[0] + sig[1]).length >= 7 && cb.indexOf(sig[0] + sig[1]) >= 0) n++;
    if (sig.length === 1 && sig[0].length >= 7 && cb.indexOf(sig[0]) >= 0) n++;
    const ini = words.map(w => w[0]).join('');
    if (ini.length >= 4 && (String(b || '').toUpperCase().match(/[A-Z]{4,}/g) || []).some(w => w === ini || w === ini.replace(/PL$/, '') + 'PL')) n++;
    for (const w of A) { if (B.has(w)) n++; else for (const x of B) if (w.length >= 5 && x.length >= 5 && (w.startsWith(x.slice(0, 5)) || x.startsWith(w.slice(0, 5)))) { n++; break; } }
    return n;
  }

  /**
   * bank  : { rows:[{date, narration, withdrawal, deposit, balance, type, ref, party, page, file}], account, bank }
   * books : [ readBankBook() results ]      mis : [ readEazypay() results ]
   */
  function reconcile(bank, books, mis, opts) {
    opts = opts || {};
    const WINDOW = opts.dateWindow === undefined ? 2 : opts.dateWindow;
    const notes = [];

    // ---------------------------------------------------------------- the three inputs, cleaned
    const B = bank.rows.map((r, i) => ({ sl: i + 1, date: r.date, narration: r.narration, amt: r.deposit ? r.deposit : -r.withdrawal, type: r.type || 'Other', ref: r.ref || '', party: r.party || '',
      page: r.page, file: r.file, c: compact(r.narration), status: '', how: '', vouchers: [], note: '', lag: null, bal: r.balance,
      key: 'B|' + r.date + '|' + (r.deposit ? r.deposit : -r.withdrawal) + '|' + r.balance }));
    { const seenK = new Map(); for (const b of B) { const n = (seenK.get(b.key) || 0) + 1; seenK.set(b.key, n); if (n > 1) b.key += '|' + n; } }
    const firstBank = B.length ? B[0].date : '', lastBank = B.length ? B[B.length - 1].date : '';

    // Bank Book: same block from two overlapping downloads is kept once (the later file wins)
    const blockMap = new Map(); const urnAll = [];
    const bookFiles = books.slice().sort((a, b) => (a.from || '').localeCompare(b.from || '') || (a.to || '').localeCompare(b.to || ''));
    for (const f of bookFiles) {
      const ids = new Set(f.blocks.map(b => b.blockId + '|' + b.vchDate));
      for (const k of ids) if (blockMap.has(k)) { const old = blockMap.get(k); old.dropped = true; }
      for (const b of f.blocks) blockMap.set(b.blockId + '|' + b.vchDate, Object.assign({}, b, { key: b.blockId + '|' + b.vchDate, status: '', how: '', bank: [], note: '', instruments: b.instruments.map(x => Object.assign({ used: null }, x)) }));
      for (const u of f.urns) urnAll.push(Object.assign({ key: u.blockId + '|' + u.vchDate, file: f.fileName }, u));
      if (f.tie && !/EXACT/i.test(f.tie)) notes.push('Bank Book ' + f.fileName + ': its own Control sheet says "' + f.tie + '", so some vouchers may be missing from it. Download it again with the latest Linearizer.');
    }
    const V = Array.from(blockMap.values()).sort((a, b) => (a.vchDate || '').localeCompare(b.vchDate || '') || a.blockId.localeCompare(b.blockId, undefined, { numeric: true }));
    const VK = new Map(V.map(v => [v.key, v]));
    const seenU = new Set(); const URN = [];
    for (const u of urnAll) { const k = u.key + '|' + u.tran + '|' + u.amt + '|' + u.partyCode; if (!VK.has(u.key)) continue; URN.push(u); }
    // (urn lines of a block that came in two files would double: keep those of the surviving file only)
    const urnByFile = new Map(); for (const u of URN) { const v = VK.get(u.key); if (v.file === u.file) { if (!urnByFile.has(u.key)) urnByFile.set(u.key, []); urnByFile.get(u.key).push(u); } }
    const urnLines = []; urnByFile.forEach(a => urnLines.push.apply(urnLines, a));
    const firstBook = V.length ? V[0].vchDate : '', lastBook = V.length ? V[V.length - 1].vchDate : '';

    // Eazypay: one row per TranID
    const M = new Map();
    for (const f of mis) for (const r of f.rows) if (!M.has(r.tran)) M.set(r.tran, Object.assign({ bankSl: null, vouchers: [], booked: 0, status: '' }, r));
    const bySettle = new Map();
    for (const m of M.values()) { if (!bySettle.has(m.settle)) bySettle.set(m.settle, []); bySettle.get(m.settle).push(m); }
    const urnByTran = new Map();
    for (const u of urnLines) { if (!urnByTran.has(u.tran)) urnByTran.set(u.tran, []); urnByTran.get(u.tran).push(u); }

    // ---------------------------------------------------------------- A. Eazypay lumps
    const usedSettle = new Set();
    const lumps = B.filter(b => b.type === 'Eazypay settlement');
    const lumpDate = b => { const m = b.c.match(/EZY(\d{2})(\d{2})(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : b.date; };
    const sumOf = d => (bySettle.get(d) || []).reduce((a, m) => a + m.amt, 0);
    for (const b of lumps) { const d = lumpDate(b); b.settle = d; if (bySettle.has(d) && sumOf(d) === b.amt) { usedSettle.add(d); b.ezy = d; } }
    for (const b of lumps) {
      if (b.ezy) continue;
      for (const k of [1, -1, 2, -2]) { const d = addDays(b.settle, k); if (!usedSettle.has(d) && bySettle.has(d) && sumOf(d) === b.amt) { usedSettle.add(d); b.ezy = d; b.ezyShift = k; break; } }
    }
    for (const b of lumps) {
      if (!b.ezy) {
        const have = bySettle.has(b.settle);
        b.status = 'Needs review';
        b.how = '';
        b.code = have ? 'EZY_PARTFILE' : 'EZY_NOFILE';
        b.note = have ? 'Eazypay file shows ' + fmt(sumOf(b.settle)) + ' for settlement date ' + nice(b.settle) + ', the bank credited ' + fmt(b.amt) + '. A neighbouring Eazypay file is probably missing.'
          : 'No Eazypay rows for settlement date ' + nice(b.settle) + '. Add the Eazypay file that covers this date.';
        continue;
      }
      const pays = bySettle.get(b.ezy); pays.forEach(m => { m.bankSl = b.sl; m.bankDate = b.date; });
      b.pays = pays;
    }
    // each payment against the URN lines of the Bank Book
    for (const m of M.values()) {
      const lines = urnByTran.get(m.tran) || [];
      m.booked = lines.reduce((a, u) => a + u.amt, 0);
      m.vouchers = Array.from(new Set(lines.map(u => u.key)));
      m.status = !lines.length ? 'Not in books' : 'Booked';
    }
    // groups of vouchers and payments that belong together (one payment split over wings, one voucher holding many payments)
    const par = new Map(); const find = x => { while (par.get(x) !== x) { par.set(x, par.get(par.get(x))); x = par.get(x); } return x; };
    const uni = (a, c) => { if (!par.has(a)) par.set(a, a); if (!par.has(c)) par.set(c, c); par.set(find(a), find(c)); };
    for (const u of urnLines) uni('V' + u.key, 'T' + u.tran);
    const sib = new Map();
    for (const v of V) { const k = v.voucherId + '|' + v.vchDate; if (!sib.has(k)) sib.set(k, []); sib.get(k).push(v); }
    for (const list of sib.values()) if (list.length > 1 && list.some(v => par.has('V' + v.key))) { const head = list.find(v => par.has('V' + v.key)); list.forEach(v => uni('V' + v.key, 'V' + head.key)); }
    const groups = new Map();
    for (const x of par.keys()) { const r = find(x); if (!groups.has(r)) groups.set(r, { v: [], t: [] }); groups.get(r)[x[0] === 'V' ? 'v' : 't'].push(x.slice(1)); }
    for (const g of groups.values()) {
      const vs = g.v.map(k => VK.get(k)); const known = g.t.every(x => M.has(x));
      const net = vs.reduce((a, v) => a + v.net, 0);
      if (!known) {
        const miss = g.t.filter(x => !M.has(x));
        vs.forEach(v => { v.ez = { missing: miss.slice(), trans: g.t.slice() }; v.status = 'Needs review'; v.how = 'Eazypay TranID'; v.note = 'URN ' + miss.slice(0, 3).join(', ') + (miss.length > 3 ? '…' : '') + ' is not in the Eazypay files given' + (v.vchDate <= addDays(firstBank || v.vchDate, 3) ? ' (probably paid before ' + nice(firstBank) + ').' : '. Add the Eazypay file for that period.'); });
        continue;
      }
      const pay = g.t.reduce((a, x) => a + M.get(x).amt, 0);
      const lumpsOk = g.t.every(x => M.get(x).bankSl !== null);
      const diff = net - pay;
      const sls = Array.from(new Set(g.t.map(x => M.get(x).bankSl).filter(x => x !== null)));
      for (const v of vs) {
        v.ez = { pay, net, diff, trans: g.t.slice(), vs: vs.map(x => x.blockId) };
        v.bank = sls; v.how = 'Eazypay TranID';
        if (!lumpsOk && g.t.some(x => M.get(x).bankSl === null && lumps.some(b => !b.ezy && b.settle === M.get(x).settle))) { v.status = 'Needs review'; v.note = 'Its Eazypay settlement credit is in the statement but could not be tallied, because the Eazypay file for that settlement date is incomplete.'; }
        else if (!lumpsOk) { v.status = 'Needs review'; v.note = 'The Eazypay payment is in the file but its settlement credit is not in the statements loaded (settled ' + g.t.map(x => nice(M.get(x).settle)).filter((x, i, a) => a.indexOf(x) === i).join(', ') + ').'; }
        else if (diff === 0) { v.status = 'Matched'; }
        else if (Math.abs(diff) < 100) { v.status = 'Matched'; v.note = 'Differs from the Eazypay amount by paise only (' + fmt(diff) + ').'; }
        else if (diff < 0) { v.status = 'Needs review'; v.note = 'Books show ' + fmt(net) + ' against ' + fmt(pay) + ' received through Eazypay: ' + fmt(-diff) + ' of this payment is not booked yet (another wing\'s share?).'; }
        else { v.status = 'Needs review'; v.note = 'Books show ' + fmt(net) + ' against ' + fmt(pay) + ' received through Eazypay: booked ' + fmt(diff) + ' more than received.'; }
      }
      g.t.forEach(x => { const m = M.get(x); if (diff < 0 && Math.abs(diff) >= 100) m.status = 'Booked in part'; });
    }
    for (const b of lumps) {
      if (!b.pays) continue;
      const un = b.pays.filter(m => m.status === 'Not in books'), part = b.pays.filter(m => m.status === 'Booked in part');
      b.status = b.ezyShift ? 'Suggested' : 'Matched';
      b.how = 'Eazypay file: ' + b.pays.length + ' payment(s) settled ' + nice(b.ezy) + ' add up to this credit';
      b.vouchers = Array.from(new Set([].concat.apply([], b.pays.map(m => m.vouchers)))).map(k => VK.get(k).blockId);
      const bits = [];
      if (b.ezyShift) bits.push('Matched to the Eazypay total of ' + nice(b.ezy) + ', not ' + nice(b.settle) + '. Please confirm.');
      if (un.length) bits.push(un.length + ' payment(s) of ' + fmt(un.reduce((a, m) => a + m.amt, 0)) + ' are not in the Bank Book yet (see the Eazypay sheet).');
      if (part.length) bits.push(part.length + ' payment(s) are only partly booked.');
      b.note = bits.join(' '); b.unbooked = un.reduce((a, m) => a + m.amt, 0);
      const vd = b.pays.map(m => m.vouchers.map(k => VK.get(k).vchDate)).reduce((a, x) => a.concat(x), []).sort();
      if (vd.length) b.lag = days(vd[vd.length - 1], b.date);
    }

    // ---------------------------------------------------------------- B. direct credits against receipt vouchers
    const INS = [];
    for (const v of V) {
      if (v.hasUrn || v.net <= 0 || v.status) continue;
      v.instruments.forEach((x, i) => { if (x.amt > 0) INS.push({ v, i, x, amt: x.amt, date: x.date, utr: x.utr, real: realUtr(x.utr), cu: compact(x.utr) }); });
    }
    const NOT_DIRECT = new Set(['Eazypay settlement', 'Sweep in', 'Sweep out', 'FD closure / interest', 'Chargeback']);
    const credits = B.filter(b => b.amt > 0 && !NOT_DIRECT.has(b.type));
    const take = (b, list, status, how, note) => {
      list.forEach(n => { n.x.used = b.sl; n.usedBy = b; });
      b.status = status; b.how = how; b.note = note || '';
      b.vouchers = Array.from(new Set(list.map(n => n.v.blockId)));
      b.lag = Math.max.apply(null, list.map(n => days(n.v.vchDate, b.date)));
      b._ins = list;
    };
    // B1. a real UTR that appears in the bank narration
    const byUtr = new Map();
    for (const n of INS) if (n.real) { if (!byUtr.has(n.cu)) byUtr.set(n.cu, []); byUtr.get(n.cu).push(n); }
    for (const b of credits) {
      for (const [u, list] of byUtr) {
        if (b.c.indexOf(u) < 0) continue;
        const free = list.filter(n => !n.usedBy); if (!free.length) continue;
        const sum = free.reduce((a, n) => a + n.amt, 0);
        if (sum === b.amt) take(b, free, 'Matched', 'UTR ' + free[0].utr + (free.length > 1 ? ' (' + free.length + ' vouchers share it and add up)' : ''));
        else { const one = free.find(n => n.amt === b.amt); if (one) take(b, [one], 'Matched', 'UTR ' + one.utr); else take(b, free, 'Suggested', 'UTR ' + free[0].utr, 'Same UTR, but the bank credited ' + fmt(b.amt) + ' and the voucher(s) show ' + fmt(sum) + ' (TDS deducted, or part booked?).'); }
        break;
      }
    }
    // B2. amount + instrument date, with the party name as the decider
    const near = (b, n) => n.date && Math.abs(days(n.date, b.date)) <= WINDOW;
    const cands = b => INS.filter(n => !n.usedBy && n.amt === b.amt && near(b, n));
    const bankName = b => b.party + ' ' + b.narration;
    for (const b of credits) {
      if (b.status) continue;
      const c = cands(b).filter(n => nameMatch(n.v.party + ' ' + n.v.narration.slice(0, 0), bankName(b)) > 0);
      if (c.length === 1) take(b, c, 'Matched', 'Amount, date and party name' + (c[0].utr && !c[0].real ? ' (the ERP\'s UTR "' + c[0].utr + '" is a placeholder)' : ''));
    }
    for (const b of credits) {
      if (b.status) continue;
      const c = cands(b);
      const cn = c.filter(n => nameMatch(n.v.party, bankName(b)) > 0);
      const one = l => new Set(l.map(n => n.v.key)).size === 1;
      const nearest = l => l.slice().sort((x, y) => Math.abs(days(x.date, b.date)) - Math.abs(days(y.date, b.date)))[0];
      if (cn.length === 1 || (cn.length > 1 && one(cn))) take(b, [nearest(cn)], 'Matched', 'Amount, date and party name');
      else if (c.length > 1 && one(c)) take(b, [nearest(c)], 'Suggested', 'Amount and date', 'Party name differs: bank "' + (b.party || b.narration.slice(0, 40)) + '", books "' + c[0].v.party + '".');
      else if (c.length === 1) take(b, c, 'Suggested', 'Amount and date', 'Only voucher with this amount and date, but the party name differs: bank "' + (b.party || b.narration.slice(0, 40)) + '", books "' + c[0].v.party + '".');
      else if (c.length > 1) { b.status = 'Needs review'; b.note = 'Several vouchers fit this amount and date: ' + Array.from(new Set(c.map(n => n.v.blockId))).slice(0, 6).join(', ') + '. Pick the right one.'; b._many = c; }
    }
    // B2b. the same, when bank and books differ only in paise (the ERP rounded, or the paise were left out)
    for (const b of credits) {
      if (b.status) continue;
      const c = INS.filter(n => !n.usedBy && n.amt !== b.amt && Math.abs(n.amt - b.amt) < 100 && near(b, n) && nameMatch(n.v.party, bankName(b)) > 0);
      if (new Set(c.map(n => n.v.key)).size === 1) take(b, [c[0]], 'Suggested', 'Amount (paise apart), date and party name', 'The bank credited ' + fmt(b.amt) + ' and the voucher shows ' + fmt(c[0].amt) + ': they differ by ' + fmt(Math.abs(b.amt - c[0].amt)) + '. Please confirm.');
    }
    // B2c. same amount and same party, but the payment date in the voucher is a few days away from the bank date
    for (const b of credits) {
      if (b.status) continue;
      const c = INS.filter(n => !n.usedBy && n.amt === b.amt && n.date && Math.abs(Math.round((Date.parse(n.date) - Date.parse(b.date)) / 864e5)) <= 10 && nameMatch(n.v.party, bankName(b)) > 0);
      if (new Set(c.map(n => n.v.key)).size === 1) { const dd = Math.abs(Math.round((Date.parse(c[0].date) - Date.parse(b.date)) / 864e5)); take(b, [c[0]], 'Suggested', 'Amount and party name (payment date ' + dd + ' days away)', 'The voucher gives ' + nice(c[0].date) + ' as the payment date; the bank credited it on ' + nice(b.date) + '. Please confirm.'); }
    }
    // B3. one bank credit booked as several vouchers of the same party
    for (const b of credits) {
      if (b.status) continue;
      const pool = INS.filter(n => !n.usedBy && near(b, n) && n.amt < b.amt && nameMatch(n.v.party, bankName(b)) > 0).slice(0, 12);
      if (pool.length < 2) continue;
      let hit = null; const N = pool.length;
      for (let mask = 3; mask < (1 << N) && !hit; mask++) {
        let s = 0, k = 0; for (let i = 0; i < N; i++) if (mask & (1 << i)) { s += pool[i].amt; k++; }
        if (k >= 2 && s === b.amt) { let other = false; for (let m2 = mask + 1; m2 < (1 << N); m2++) { let s2 = 0; for (let i = 0; i < N; i++) if (m2 & (1 << i)) s2 += pool[i].amt; if (s2 === b.amt) { other = true; break; } } hit = { mask, other }; }
      }
      if (hit) { const list = pool.filter((n, i) => hit.mask & (1 << i)); take(b, list, 'Suggested', list.length + ' vouchers of the same party add up to this credit', hit.other ? 'Another combination also adds up. Please check.' : 'Please confirm.'); }
    }
    // B4. amount alone (voucher has no usable instrument date)
    for (const b of credits) {
      if (b.status) continue;
      const c = INS.filter(n => !n.usedBy && n.amt === b.amt && !n.date && days(n.v.vchDate, b.date) >= -3 && days(n.v.vchDate, b.date) <= 200);
      if (c.length === 1) take(b, c, 'Suggested', 'Amount only', 'The voucher shows no instrument date. Please confirm.');
    }
    for (const b of credits) if (!b.status) { b.status = 'Not in books'; b.note = 'No receipt voucher found for this credit in the Bank Book files given (up to ' + nice(lastBook) + ').'; }

    // ---------------------------------------------------------------- C. withdrawals against payment vouchers
    const outs = V.filter(v => v.net < 0);
    const isSusp = v => !!v.suspenseRef || /suspense/i.test(v.narration);
    const isCancel = v => /cancellation of mr/i.test(v.narration);
    const debits = B.filter(b => b.amt < 0 && !NOT_DIRECT.has(b.type));
    const freeOuts = () => outs.filter(v => !v.status && !isSusp(v) && !isCancel(v));
    // equal amounts paid several times (the regular transfers): pair them in date order when the counts agree
    const amtsSeen = new Set(debits.map(b => b.amt));
    for (const a of amtsSeen) {
      const bl = debits.filter(b => b.amt === a && !b.status).sort((x, y) => x.date.localeCompare(y.date) || x.sl - y.sl);
      const vl = freeOuts().filter(v => v.net === a).sort((x, y) => x.vchDate.localeCompare(y.vchDate));
      if (bl.length < 2 || bl.length !== vl.length) continue;
      if (!bl.every((b, i) => days(vl[i].vchDate, b.date) >= -10 && days(vl[i].vchDate, b.date) <= 45)) continue;
      bl.forEach((b, i) => { const v = vl[i]; v.status = 'Matched'; v.how = 'Amount, in date order (' + bl.length + ' payments of this amount, ' + bl.length + ' vouchers)'; v.bank = [b.sl]; b.status = 'Matched'; b.how = v.how; b.vouchers = [v.blockId]; b.lag = days(v.vchDate, b.date); });
    }
    for (const b of debits) {
      if (b.status) continue;
      const c = freeOuts().filter(v => v.net === b.amt && days(v.vchDate, b.date) >= -10 && days(v.vchDate, b.date) <= 45).sort((x, y) => Math.abs(days(x.vchDate, b.date)) - Math.abs(days(y.vchDate, b.date)));
      if (!c.length) continue;
      const same = debits.filter(o => o.amt === b.amt && !o.status && Math.abs(days(o.date, b.date)) <= 45).length;
      const v = c[0];
      v.status = (c.length === 1 && same === 1) ? 'Matched' : 'Suggested'; v.how = 'Amount and nearest date'; v.bank = [b.sl];
      v.note = v.status === 'Suggested' ? 'Several payments of the same amount fall close together; paired in date order. Please confirm.' : '';
      b.status = v.status; b.how = 'Amount and nearest date'; b.vouchers = [v.blockId]; b.lag = days(v.vchDate, b.date); b.note = v.note;
    }
    for (const v of freeOuts().sort((x, y) => x.net - y.net)) {
      const pool = debits.filter(b => !b.status && days(b.date, v.vchDate) >= -15 && days(b.date, v.vchDate) <= 15 && b.amt > v.net).slice(0, 14);
      if (pool.length < 2) continue;
      let hit = null; const N = pool.length, want = -v.net;
      // first: all withdrawals of one day
      const byDay = new Map(); pool.forEach(b => { if (!byDay.has(b.date)) byDay.set(b.date, []); byDay.get(b.date).push(b); });
      for (const list of byDay.values()) if (list.length >= 2 && list.reduce((a, b) => a - b.amt, 0) === want) { hit = list; break; }
      for (let mask = 3; mask < (1 << N) && !hit; mask++) { let sum = 0, k = 0; for (let i = 0; i < N; i++) if (mask & (1 << i)) { sum -= pool[i].amt; k++; } if (k >= 2 && sum === want) hit = pool.filter((b, i) => mask & (1 << i)); }
      if (!hit) continue;
      const named = hit.every(b => nameMatch(v.party, b.party + ' ' + b.narration) > 0);
      v.status = named ? 'Matched' : 'Suggested'; v.how = hit.length + ' bank withdrawals add up to this voucher' + (named ? ' (same payee)' : ''); v.bank = hit.map(b => b.sl);
      v.note = named ? '' : 'Please confirm.';
      hit.forEach(b => { b.status = v.status; b.how = 'Part of ' + v.blockId + ' (' + hit.length + ' withdrawals add up to ' + fmt(want) + ')'; b.vouchers = [v.blockId]; b.lag = days(v.vchDate, b.date); b.note = v.note; });
    }
    for (const b of debits) if (!b.status) { b.status = 'Not in books'; b.note = 'No payment voucher found for this withdrawal in the Bank Book files given.'; }

    // ---------------------------------------------------------------- D. the things that are known not to match one-to-one
    for (const b of B) {
      if (b.status) continue;
      if (b.type === 'Sweep out' || b.type === 'Sweep in') { b.status = 'Sweep'; b.note = 'Movement between the account and its linked deposit (' + (b.ref || 'FD') + '). Not in the Bank Book.'; }
      else if (b.type === 'FD closure / interest') { b.status = 'Bank only'; b.note = 'Interest / closure proceeds of linked deposit ' + b.ref + ' credited by the bank. Needs a voucher.'; }
      else if (b.type === 'Chargeback') {
        b.status = 'Bank only';
        const m = Array.from(M.values()).find(x => x.amt === -b.amt && b.c.indexOf(x.initiated.slice(8, 10) + x.initiated.slice(5, 7) + x.initiated.slice(2, 4)) >= 0);
        b.note = 'Chargeback debited by the bank.' + (m ? ' Looks like the reversal of Eazypay payment ' + m.tran + ' of ' + nice(m.initiated) + ' (' + m.party + ').' : '');
      } else { b.status = 'Not in books'; }
    }
    // vouchers
    const mrOf = new Map(); V.forEach(v => { if (v.mrNo && v.net > 0) mrOf.set(v.mrNo, v); });
    for (const v of V) {
      if (v.net > 0 && !v.hasUrn) {
        const used = v.instruments.filter(x => x.used !== null), all = v.instruments.filter(x => x.amt > 0);
        if (used.length) {
          const lines = Array.from(new Set(used.map(x => x.used))); v.bank = lines;
          const sts = lines.map(sl => B[sl - 1].status); const how = B[lines[0] - 1].how;
          v.how = how;
          if (used.length < all.length) { v.status = 'Needs review'; v.note = used.length + ' of ' + all.length + ' instruments of this voucher were found in the bank statement; the rest were not.'; }
          else { v.status = sts.includes('Suggested') ? 'Suggested' : 'Matched'; v.note = B[lines[0] - 1].status === 'Suggested' ? B[lines[0] - 1].note : ''; }
          continue;
        }
      }
      if (v.status) continue;
      if (isSusp(v)) { v.status = 'Suspense entry'; v.note = 'Suspense entry' + (v.suspenseRef ? ' referring to ' + v.suspenseRef : '') + '. No bank line of its own; see the statement enclosed with the voucher.'; continue; }
      if (isCancel(v)) {
        const m = v.narration.match(/MR\s*No\.?\s*-?\s*(\d+)/i); const o = m ? mrOf.get(m[1]) : null;
        v.status = 'Cancelled receipt'; v.note = 'Cancels money receipt ' + (m ? m[1] : '') + (o ? ' (' + o.blockId + ' of ' + nice(o.vchDate) + ', ' + fmt(o.net) + ')' : '') + '. No bank line of its own.';
        if (o && !o.bank.length && o.net === -v.net) { o.status = 'Cancelled receipt'; o.note = 'Cancelled by ' + v.blockId + ' of ' + nice(v.vchDate) + '. No bank line expected.'; }
        continue;
      }
      if (v.hasUrn) { v.status = 'Needs review'; v.note = v.note || 'Has Eazypay URNs but none could be traced.'; continue; }
      const d = v.instruments.map(x => x.date).filter(Boolean).sort()[0] || '';
      if (v.net > 0 && d && firstBank && d < firstBank) { v.status = 'Earlier period'; v.note = 'Receipt dated ' + nice(d) + ', before the statements loaded (they start ' + nice(firstBank) + '). No bank line expected here.'; }
      else if (v.net > 0 && d && lastBank && d > lastBank) { v.status = 'After statement period'; v.note = 'Receipt dated ' + nice(d) + ', after the statements loaded (they end ' + nice(lastBank) + ').'; }
      else if (v.net > 0) { v.status = 'Not in bank'; v.note = d ? 'No bank credit of ' + fmt(v.net) + ' around ' + nice(d) + ' was found.' : 'The voucher shows no instrument date, and no unmatched bank credit of ' + fmt(v.net) + ' was found.'; }
      else if (v.net < 0) { v.status = 'Not in bank'; v.note = 'No bank withdrawal of ' + fmt(-v.net) + ' near ' + nice(v.vchDate) + ' was found.'; }
      else { v.status = 'No bank effect'; }
    }
    // a provisional voucher that later got its number: point it out
    for (const v of V) if (/PROV/.test(v.blockId)) v.note = (v.note ? v.note + ' ' : '') + 'Provisional voucher (Prov No ' + v.provNo + '): it has no voucher number yet and may reappear with one.';

    // ---------------------------------------------------------------- who paid, and which URNs (for the Vouchers sheet)
    const linesOf = new Map(); for (const u of urnLines) { if (!linesOf.has(u.key)) linesOf.set(u.key, []); linesOf.get(u.key).push(u); }
    for (const v of V) {
      let ls = linesOf.get(v.key) || [];
      if (!ls.length && v.ez && v.ez.trans) ls = urnLines.filter(u => v.ez.trans.includes(u.tran));
      const seenT = new Set(); v.urnList = [];
      for (const u of ls) { if (seenT.has(u.tran)) continue; seenT.add(u.tran); const m = M.get(u.tran); v.urnList.push({ raw: u.raw, tran: u.tran, party: u.party || (m ? m.party : ''), code: u.partyCode || (m ? m.partyCode : '') }); }
      v.urnText = v.urnList.map(u => u.raw).join(', ');
      v.paidBy = Array.from(new Set(v.urnList.map(u => (u.code ? u.code + ' ' : '') + u.party).filter(Boolean))).join('; ');
      if (v.how === 'Eazypay TranID') v.how = v.urnList.length === 1 ? 'Eazypay URN ' + v.urnList[0].raw : 'Eazypay URN (' + v.urnList.length + ' URNs, listed in the Eazypay URN column)';
    }
    for (const b of lumps) if (b.pays) b.paidBy = Array.from(new Set(b.pays.map(m => m.party))).join('; ');

    // ---------------------------------------------------------------- plain-language explanations for the "To review" sheet
    const who = v => (v.paidBy ? v.paidBy : v.party);
    const short = (x, n) => (x.length > n ? x.slice(0, n - 1) + '…' : x);
    const vNames = ids => ids.join(', ');
    for (const b of B) {
      const amt = fmt(Math.abs(b.amt)), on = nice(b.date), from = b.party ? ' from ' + b.party : '';
      const at = 'Line ' + b.sl + ' of the Bank lines sheet. In the bank statement: ' + (b.file ? b.file + ', ' : '') + 'page ' + b.page + '.';
      if (b.code === 'EZY_NOFILE') { b.found = 'The bank credited ' + amt + ' on ' + on + ' as one Eazypay amount. The Eazypay report for that day was not given, so the tool cannot tell which parties are inside it.'; b.todo = 'Download the Eazypay report that covers ' + nice(b.settle) + ' (and the day before) from the bank and drop it in with the other files.'; }
      else if (b.code === 'EZY_PARTFILE') { b.found = 'The bank credited ' + amt + ' on ' + on + ' as one Eazypay amount, but the Eazypay reports given add up to only ' + fmt(sumOf(b.settle)) + ' for that day.'; b.todo = 'Some payments are missing from the Eazypay reports, usually those paid on the last days of the month before. Download the report for those days and drop it in.'; }
      else if (b.status === 'Suggested' && b.ezyShift) { b.found = 'This Eazypay amount of ' + amt + ' equals the Eazypay report total of ' + nice(b.ezy) + ', one day away from the date in the bank narration.'; b.todo = 'Nothing to search for. Just confirm that this is acceptable.'; }
      else if (b.status === 'Suggested' && /^UTR/.test(b.how)) { b.found = 'The bank received ' + amt + ' on ' + on + from + '. Voucher ' + vNames(b.vouchers) + ' has the same UTR number but a different amount.'; b.todo = 'Open voucher ' + vNames(b.vouchers) + ' in the ERP and compare the amount. The party may have deducted TDS, or only part may be booked.'; }
      else if (b.status === 'Suggested' && /days away/.test(b.how)) { const v0 = b._ins[0]; b.found = 'The bank received ' + amt + ' on ' + on + from + '. Voucher ' + vNames(b.vouchers) + ' of the same party and amount gives ' + nice(v0.date) + ' as the payment date.'; b.todo = 'Open voucher ' + vNames(b.vouchers) + ' in the ERP. If it is this receipt (the date in the voucher may be wrong), choose Accept. If not, choose Reject.'; }
      else if (b.status === 'Suggested' && /paise apart/.test(b.how)) { const v0 = b._ins[0]; b.found = 'The bank received ' + amt + ' on ' + on + from + '. Voucher ' + vNames(b.vouchers) + ' of the same party and date shows ' + fmt(v0.amt) + ': ' + fmt(Math.abs(b.amt - v0.amt)) + ' ' + (b.amt > v0.amt ? 'less' : 'more') + ' than the bank.'; b.todo = 'Check voucher ' + vNames(b.vouchers) + ' in the ERP. If it is the same receipt and only the paise differ, choose Accept: the ' + fmt(Math.abs(b.amt - v0.amt)) + ' then shows in the BRS as a difference in matched entries. If it is not the same receipt, choose Reject.'; }
      else if (b.status === 'Suggested' && b.how === 'Amount and date') { b.found = 'The bank received ' + amt + ' on ' + on + from + '. Voucher ' + vNames(b.vouchers) + ' has the same amount and date, but it is in the name of ' + (b._ins ? b._ins[0].v.party : 'another party') + '.'; b.todo = 'Check whether these are the same party under two names (for example an owner paying for his firm). If yes, accept the match.'; }
      else if (b.status === 'Suggested' && /add up to this credit/.test(b.how)) { b.found = 'The bank received ' + amt + ' on ' + on + from + ' in one go. The ERP has ' + b.vouchers.length + ' vouchers of this party (' + vNames(b.vouchers) + ') that add up to it.'; b.todo = 'Confirm that these vouchers together are for this one payment.'; }
      else if (b.status === 'Suggested' && b.how === 'Amount only') { b.found = 'The bank received ' + amt + ' on ' + on + from + '. Voucher ' + vNames(b.vouchers) + ' is for the same amount but shows no payment date.'; b.todo = 'Open voucher ' + vNames(b.vouchers) + ' in the ERP and check that it is for this payment.'; }
      else if (b.status === 'Suggested') { b.found = 'The bank paid out ' + amt + ' on ' + on + '. Voucher ' + vNames(b.vouchers) + ' is for the same amount.'; b.todo = 'Confirm that voucher ' + vNames(b.vouchers) + ' is for this payment.'; }
      else if (b.status === 'Needs review' && b._many) { b.found = 'The bank received ' + amt + ' on ' + on + from + '. More than one voucher has this amount and date: ' + Array.from(new Set(b._many.map(n => n.v.blockId))).slice(0, 6).join(', ') + '.'; b.todo = 'Open these vouchers in the ERP and decide which one belongs to this payment.'; }
      else if (b.status === 'Not in books' && b.amt > 0) { b.found = 'The bank received ' + amt + ' on ' + on + from + '. No receipt voucher for this money was found in the ERP Bank Book (checked up to ' + nice(lastBook) + ').'; b.todo = 'Check in the ERP whether a receipt has been made for this money. If not, it still has to be booked.'; }
      else if (b.status === 'Not in books') { b.found = 'The bank paid out ' + amt + ' on ' + on + '. No payment voucher for it was found in the ERP Bank Book.'; b.todo = 'Check in the ERP whether a payment voucher exists for this amount.'; }
      else if (b.type === 'FD closure / interest') { b.found = 'The bank added ' + amt + ' on ' + on + ' from a fixed deposit linked to this account (deposit no. ' + b.ref + '). This is interest or closure money, put in by the bank itself.'; b.todo = 'This money is not in the ERP Bank Book yet. Tell the person who books bank interest. The interest and TDS figures are on the last pages of the bank statement.'; }
      else if (b.type === 'Chargeback') { b.found = 'The bank took back ' + amt + ' on ' + on + '. A card payment made earlier through Eazypay was reversed.' + (/reversal of Eazypay payment (\d+) of (\S+) \((.*)\)/.test(b.note) ? ' It is payment ' + RegExp.$1 + ' of ' + RegExp.$2 + ' by ' + RegExp.$3 + '.' : ''); b.todo = 'The money for that payment did not stay with IDCO. Tell the section that booked the receipt, so they can decide what to do with it.'; }
      if (b.found) b.where = at;
    }
    for (const v of V) {
      const amt = fmt(Math.abs(v.net)), on = nice(v.vchDate), p = who(v);
      const at = 'ERP voucher ' + v.blockId + ' dated ' + on + (v.provNo ? ' (Prov No ' + v.provNo + ')' : '') + '. Bank Book file: ' + v.file + '.';
      if (v.status === 'Needs review' && v.ez && v.ez.missing) { v.found = 'Voucher ' + v.blockId + ' (' + amt + ', ' + short(p, 70) + ') was received through Eazypay, but its URN ' + v.ez.missing.slice(0, 2).join(', ') + ' is not in any Eazypay report given.'; v.todo = 'Download the Eazypay report for the days just before ' + on + ' and drop it in.'; }
      else if (v.status === 'Needs review' && /could not be tallied/.test(v.note)) { v.found = 'Voucher ' + v.blockId + ' (' + amt + ', ' + short(p, 70) + ') was received through Eazypay. The bank credit that contains it could not be checked because the Eazypay report for that day is incomplete.'; v.todo = 'Nothing is wrong with this voucher. It will clear once the missing Eazypay report (last days of the earlier month) is added.'; }
      else if (v.status === 'Needs review' && /not in the statements loaded/.test(v.note)) { v.found = 'Voucher ' + v.blockId + ' (' + amt + ', ' + short(p, 70) + ') was received through Eazypay, but the bank statement for the day it was credited was not given.'; v.todo = 'Add the bank statement that covers the days after ' + on + '.'; }
      else if (v.status === 'Needs review' && v.ez && v.ez.diff) {
        const single = v.ez.vs.length === 1 && v.ez.trans.length === 1;
        const lead = single ? short(p, 80) + ' paid ' + fmt(v.ez.pay) + ' through Eazypay (URN ' + v.urnText + '). Voucher ' + v.blockId + ' books '
          : 'Voucher ' + v.blockId + ' is one of ' + v.ez.vs.length + ' linked voucher(s) (' + short(v.ez.vs.join(', '), 70) + ') that share ' + v.ez.trans.length + ' Eazypay payment(s). The payments total ' + fmt(v.ez.pay) + '. The vouchers together book ';
        if (v.ez.diff < 0) { v.found = lead + (single ? 'only ' : 'only ') + fmt(v.ez.net) + '. ' + fmt(-v.ez.diff) + ' of the money received is not booked.'; v.todo = 'Open ' + (single ? 'the voucher' : 'these vouchers') + ' in the ERP and find where the remaining ' + fmt(-v.ez.diff) + ' should go. Often another division or wing has not yet booked its share.'; }
        else { v.found = lead + fmt(v.ez.net) + ', which is ' + fmt(v.ez.diff) + ' more than was received.'; v.todo = 'Open ' + (single ? 'the voucher' : 'these vouchers') + ' in the ERP and check the amounts. Part of a payment may have been booked twice, or a URN may be missing from one voucher.'; }
      }
      else if (v.status === 'Needs review' && /instruments of this voucher/.test(v.note)) { v.found = 'Voucher ' + v.blockId + ' (' + amt + ', ' + short(p, 70) + ') lists several payments. Only some of them were found in the bank statement.'; v.todo = 'Open the voucher in the ERP, look at each payment listed in it, and check the ones not found against the bank statement.'; }
      else if (v.status === 'Needs review') { v.found = 'Voucher ' + v.blockId + ' (' + amt + ', ' + short(p, 70) + ') could not be traced to the bank statement.'; v.todo = 'Open the voucher in the ERP and check its payment details.'; }
      else if (v.status === 'Not in bank' && v.net > 0) { const dd = v.instruments.map(x => x.date).filter(Boolean).sort()[0]; v.found = 'The ERP shows ' + amt + ' received from ' + short(p, 70) + ' (voucher ' + v.blockId + ')' + (dd ? ', paid on ' + nice(dd) : '') + '. No such amount came into the bank around that date.'; v.todo = 'Open the voucher in the ERP and check the amount and payment date. The party may have paid a different amount (for example after deducting TDS), or paid into another IDCO bank account.'; }
      else if (v.status === 'Not in bank') { v.found = 'The ERP shows ' + amt + ' paid out (voucher ' + v.blockId + ', ' + short(p, 70) + '). No such amount left the bank around ' + on + '.'; v.todo = 'Check whether the payment was actually made, and from which bank account.'; }
      if (v.found) v.where = at;
    }
    for (const m of M.values()) {
      if (m.bankSl === null || m.status === 'Booked') continue;
      const p = (m.partyCode ? m.partyCode + ' ' : '') + m.party;
      m.where = 'Search URN ' + m.tran + ' in the ERP. In the bank statement it is inside line ' + m.bankSl + ' (' + nice(m.bankDate) + ').';
      if (m.status === 'Booked in part') { m.found = p + ' paid ' + fmt(m.amt) + ' through Eazypay on ' + nice(m.initiated) + '. Only part of it is booked in the ERP (' + m.vouchers.map(k => VK.get(k).blockId).join(', ') + ').'; m.todo = 'Find out which division or wing has not booked its share of this payment.'; }
      else { m.found = p + ' paid ' + fmt(m.amt) + ' through Eazypay on ' + nice(m.initiated) + ', and the bank has credited it. No voucher in the ERP carries this URN.'; m.todo = 'Check in the ERP whether a receipt exists for this URN. If not, pass it to the section that books this party\'s receipts.'; }
    }

    // ---------------------------------------------------------------- summary
    const sum = (list, f) => list.reduce((a, x) => a + f(x), 0);
    const bankBy = {}; for (const b of B) { const k = b.status; bankBy[k] = bankBy[k] || { n: 0, dep: 0, wd: 0 }; bankBy[k].n++; if (b.amt > 0) bankBy[k].dep += b.amt; else bankBy[k].wd -= b.amt; }
    const vBy = {}; for (const v of V) { const k = v.status; vBy[k] = vBy[k] || { n: 0, dr: 0, cr: 0 }; vBy[k].n++; if (v.net > 0) vBy[k].dr += v.net; else vBy[k].cr -= v.net; }
    const pays = Array.from(M.values());
    const inLumps = pays.filter(m => m.bankSl !== null);
    const summary = {
      bank: bank.bank, account: bank.account, firstBank, lastBank, firstBook, lastBook, lastBookFile: bookFiles.map(f => f.to).sort().pop() || lastBook, bankLines: B.length, vouchers: V.length,
      bankBy, vBy,
      eazypay: { payments: pays.length, inStatement: inLumps.length, notInBooks: inLumps.filter(m => m.status === 'Not in books').length, notInBooksAmt: sum(inLumps.filter(m => m.status === 'Not in books'), m => m.amt),
        partly: inLumps.filter(m => m.status === 'Booked in part').length },
      sweepOut: sum(B.filter(b => b.type === 'Sweep out'), b => -b.amt), sweepIn: sum(B.filter(b => b.type === 'Sweep in'), b => b.amt),
      bookFiles: bookFiles.map(f => ({ fileName: f.fileName, from: f.from, to: f.to, tie: f.tie, blocks: f.blocks.length, gl: f.gl, account: f.account })),
      misFiles: mis.map(f => ({ fileName: f.fileName, rows: f.rows.length })),
      notes
    };
    // months of Bank Book missing inside the statement period
    if (firstBank && bookFiles.length) {
      const have = new Set(); bookFiles.forEach(f => { if (!f.from || !f.to) return; for (let d = f.from.slice(0, 7); d <= f.to.slice(0, 7);) { have.add(d); const y = +d.slice(0, 4), m = +d.slice(5, 7); d = m === 12 ? (y + 1) + '-01' : y + '-' + String(m + 1).padStart(2, '0'); } });
      const miss = []; for (let d = firstBank.slice(0, 7); d <= lastBank.slice(0, 7);) { if (!have.has(d)) miss.push(d); const y = +d.slice(0, 4), m = +d.slice(5, 7); d = m === 12 ? (y + 1) + '-01' : y + '-' + String(m + 1).padStart(2, '0'); }
      if (miss.length) notes.push('No Bank Book file was given for: ' + miss.map(x => x.slice(5, 7) + '-' + x.slice(0, 4)).join(', ') + '. Bank lines of those months will show as not in books.');
    }
    for (const v of V) v.ikey = 'V|' + v.key;
    for (const m of pays) m.ikey = 'E|' + m.tran;
    return { bankLines: B, vouchers: V, payments: pays, summary, fmt, nice };
  }

  /** Counts again after a reviewer's decisions have changed some statuses. */
  function recount(rec) {
    const bankBy = {}, vBy = {};
    for (const b of rec.bankLines) { const k = b.status; bankBy[k] = bankBy[k] || { n: 0, dep: 0, wd: 0 }; bankBy[k].n++; if (b.amt > 0) bankBy[k].dep += b.amt; else bankBy[k].wd -= b.amt; }
    for (const v of rec.vouchers) { const k = v.status; vBy[k] = vBy[k] || { n: 0, dr: 0, cr: 0 }; vBy[k].n++; if (v.net > 0) vBy[k].dr += v.net; else vBy[k].cr -= v.net; }
    rec.summary.bankBy = bankBy; rec.summary.vBy = vBy;
    return rec;
  }
  return { reconcile, recount, realUtr, nameMatch, fmt };
});
