/* IDCO Bank Statement to Excel - the decision step.
 *
 * A person fills the Decision / Match with / Remark columns of the matching workbook, saves it, and drops it
 * back in together with the statements and Bank Books. This module
 *   readReview(wb)        reads those decisions out of the saved workbook, and
 *   apply(rec, decisions) carries them onto a fresh reconcile() result.
 * Decisions are tied to the item itself (bank line: date + amount + balance; voucher: number + date), not to a
 * row or line number, so they survive when more months are added later.
 *
 *   Accept     the tool's suggested match is right
 *   Reject     the tool's match is wrong; both sides become open again
 *   Keep open  looked at; it is a genuine open item
 *   Match with the person names the other side (voucher numbers and / or bank line numbers)
 *   Undo       take an earlier decision back
 * All money in paise. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementDecisions = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const OPTIONS = ['Accept', 'Reject', 'Keep open', 'Undo'];
  const SHEETS = ['To review', 'Bank lines', 'Vouchers', 'Last BRS', 'Decisions'];
  const SEP = ' ¦ ';

  function val(c) {
    if (c === null || c === undefined) return null;
    if (c instanceof Date) return c;
    if (typeof c === 'object') {
      if (c.result !== undefined) return val(c.result);
      if (c.text !== undefined) return String(c.text);
      if (c.richText) return c.richText.map(r => r.text).join('');
      return null;
    }
    return c;
  }
  const str = v => { v = val(v); return v === null ? '' : (v instanceof Date ? v.toISOString().slice(0, 10) : String(v)).replace(/ /g, ' ').trim(); };
  const sig = (d, m, r) => (d || m || r ? [d || '', m || '', r || ''].join(SEP).trim() : '');
  const normDecision = s => { const x = String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); return x === 'accept' ? 'Accept' : x === 'reject' ? 'Reject' : x === 'keep open' ? 'Keep open' : x === 'undo' ? 'Undo' : x === 'match' ? 'Match' : x === 'remark' ? 'Remark' : x ? '?' : ''; };
  const sameSet = (a, b) => a.length === b.length && a.every(x => b.includes(x));

  function rowsOf(ws) {
    const out = []; let head = null;
    ws.eachRow({ includeEmpty: false }, (row, n) => {
      const vals = []; row.eachCell({ includeEmpty: true }, (cell, col) => { vals[col - 1] = cell.value; });
      if (!head) { head = vals.map(str); return; }
      const o = { _row: n }; head.forEach((h, i) => { if (h) o[h] = vals[i]; });
      out.push(o);
    });
    return out;
  }

  /** Is this workbook a matching workbook made by this tool? */
  function isReview(wb) {
    const ws = wb.getWorksheet('To review'); if (!ws || !wb.getWorksheet('Bank lines') || !wb.getWorksheet('Vouchers')) return false;
    const first = []; ws.getRow(1).eachCell(c => first.push(str(c.value)));
    return first.includes('Decision');
  }
  /** A matching workbook made before the decision step existed (v1.1 / v1.2): it has no item keys. */
  function isOldReview(wb) {
    if (!isReview(wb)) return false;
    const first = []; wb.getWorksheet('To review').getRow(1).eachCell(c => first.push(str(c.value)));
    return !first.includes('Item key');
  }

  function readReview(wb, fileName) {
    const out = { fileName, account: '', made: wb.created instanceof Date ? wb.created.getTime() : 0, decisions: [], problems: [], seen: 0 };
    const s0 = wb.getWorksheet('Summary'); if (s0) { const m = str(s0.getRow(2).getCell(1).value).match(/A\/c\s+(\S+)/); if (m) out.account = m[1]; }
    const data = {}; SHEETS.forEach(n => { const ws = wb.getWorksheet(n); data[n] = ws ? rowsOf(ws) : []; });
    // what the saved workbook itself says about line numbers, voucher numbers and who is linked to whom
    const lineKey = new Map(), vchKey = new Map(), links = new Map(), opKey = new Map();
    for (const r of data['Last BRS']) { const k = str(r['Item key']); if (!k) continue; opKey.set(str(r['No.']), k); links.set(k, str(r['Linked keys']).split(' ').filter(Boolean)); }
    for (const r of data['Bank lines']) { const k = str(r['Item key']); if (!k) continue; lineKey.set(str(r.Line), k); links.set(k, str(r['Linked keys']).split(' ').filter(Boolean)); }
    for (const r of data.Vouchers) { const k = str(r['Item key']); if (!k) continue; const id = str(r.Voucher).toUpperCase().replace(/\s/g, ''); if (!vchKey.has(id)) vchKey.set(id, []); vchKey.get(id).push(k); links.set(k, str(r['Linked keys']).split(' ').filter(Boolean)); }

    function parseMatch(text, where) {
      const keys = []; let bad = '';
      for (let tok of String(text).split(/[,;\n]+|\s+and\s+|\s*&\s*|\s*\+\s*/i)) {
        tok = tok.trim(); if (!tok) continue;
        const om = tok.match(/^(?:last[\s-]*brs|opening|op)\.?\s*(?:item)?\s*(?:no\.?)?\s*#?\s*(\d{1,5})$/i);
        if (om) { const k = opKey.get(String(+om[1])); if (k) keys.push(k); else bad = bad || 'there is no item ' + om[1] + ' in the Last BRS sheet'; continue; }
        const m = tok.match(/^(?:bank\s*)?(?:line|l|sl|no)?\.?\s*#?\s*(\d{1,6})$/i);
        if (m) { const k = lineKey.get(String(+m[1])); if (k) keys.push(k); else bad = bad || 'there is no bank line ' + m[1]; continue; }
        const ks = vchKey.get(tok.toUpperCase().replace(/\s/g, ''));
        if (!ks) bad = bad || '"' + tok + '" is neither a voucher number nor a bank line number of this file';
        else if (ks.length > 1) bad = bad || 'voucher number ' + tok + ' occurs ' + ks.length + ' times (different dates), so it is not clear which one is meant';
        else keys.push(ks[0]);
      }
      if (bad) out.problems.push(where + ': "Match with" could not be understood - ' + bad + '.');
      return bad ? null : Array.from(new Set(keys));
    }

    const base = new Map(), edits = new Map();
    for (const r of data.Decisions) {
      const k = str(r['Item key']); if (!k) continue;
      base.set(k, { key: k, decision: normDecision(r.Decision), with: str(r['Linked keys']).split(' ').filter(Boolean), remark: str(r.Remark), on: str(r['Decided on']),
        label: { what: str(r.Item), date: str(r.Date), who: str(r.Who), amt: typeof val(r.Amount) === 'number' ? Math.round(val(r.Amount) * 100) : null }, _carried: str(r.Carried), _row: r._row, _raw: str(r.Decision) });
    }
    for (const name of SHEETS) for (const r of data[name]) {
      const k = str(r['Item key']); if (!k) continue;
      const dRaw = str(r.Decision), mw = name === 'Decisions' ? '' : str(r['Match with']), rem = str(r.Remark);
      const s = name === 'Decisions' ? sig(dRaw, '', rem) : sig(dRaw, mw, rem);
      const carried = str(r.Carried);
      if (name === 'Decisions' ? s === carried : (!s && !carried) || s === carried) continue;      // untouched row
      if (!edits.has(k)) edits.set(k, []);
      edits.get(k).push({ sheet: name, row: r._row, dRaw, mw, rem, s });
    }
    out.seen = edits.size;
    const today = new Date().toISOString().slice(0, 10);
    for (const [k, list] of edits) {
      const where = list[0].sheet + ' sheet, row ' + list[0].row;
      if (new Set(list.map(e => e.dRaw.toLowerCase() + '|' + e.mw.toLowerCase())).size > 1) { out.problems.push(where + ' and ' + list[1].sheet + ' sheet, row ' + list[1].row + ': two different decisions were written for the same item. Neither was used; write it in one place only.'); continue; }
      const e = list[0], b = base.get(k), d = normDecision(e.dRaw);
      const rem = list.map(x => x.rem).find(Boolean) || '';
      if (d === '?') { out.problems.push(where + ': "' + e.dRaw + '" is not one of Accept, Reject, Keep open, Undo. It was not used.'); continue; }
      if (e.sheet === 'Decisions') {                       // an earlier decision changed in the log
        if (d === 'Undo' || d === '') { base.delete(k); continue; }
        if (d === 'Remark' || b.decision === 'Remark') { b.remark = rem; if (d !== 'Remark') b.decision = d; b.on = today; continue; }
        b.decision = d === 'Accept' && b.decision === 'Match' ? 'Match' : d; b.remark = rem; b.on = today; if (d === 'Keep open') b.with = [];
        continue;
      }
      if (d === 'Undo' || (!d && !e.mw && !rem)) { base.delete(k); continue; }
      if (!d && !e.mw) { if (b) b.remark = rem; else base.set(k, { key: k, decision: 'Remark', with: [], remark: rem, on: today, label: null }); continue; }
      let rec = { key: k, decision: d, with: [], remark: rem, on: today, label: null };
      if (e.mw) {
        if (d === 'Keep open') { out.problems.push(where + ': both "Keep open" and "Match with" are filled. Use one of them.'); continue; }
        const keys = parseMatch(e.mw, where); if (!keys) continue;
        rec.decision = 'Match'; rec.with = keys.filter(x => x !== k);
        if (!rec.with.length) { out.problems.push(where + ': "Match with" names the item itself.'); continue; }
      } else if (d === 'Accept' || d === 'Reject') rec.with = (links.get(k) || []).slice();
      base.set(k, rec);
    }
    out.decisions = Array.from(base.values()).filter(d => d.decision && d.decision !== '?' && d.decision !== 'Undo').map(d => ({ key: d.key, decision: d.decision, with: d.with, remark: d.remark, on: d.on, label: d.label, file: fileName }));
    return out;
  }

  /** Several saved workbooks of one account: the one made last wins for an item. */
  function combine(reviews) {
    const m = new Map(); const problems = [];
    const said = s2 => s2.decision + ' ' + (s2.with || []).slice().sort().join(' ');
    reviews.slice().sort((a, b) => a.made - b.made).forEach(r => {
      r.decisions.forEach(d => { const o = m.get(d.key); if (o && o.file !== d.file && said(o) !== said(d) && o.decision !== 'Remark' && d.decision !== 'Remark') problems.push('The same item was decided differently in ' + o.file + ' (' + o.decision + ') and in ' + d.file + ' (' + d.decision + '). The later file (' + d.file + ') was used. Item: ' + (d.label && d.label.what ? d.label.what : d.key) + '.'); m.set(d.key, d); });
      r.problems.forEach(p => problems.push(r.fileName + ' - ' + p));
    });
    return { decisions: Array.from(m.values()), problems };
  }

  // ------------------------------------------------------------------------------------------------ applying
  function apply(rec, decisions, R) {
    const B = rec.bankLines, V = rec.vouchers, fmt = rec.fmt, nice = rec.nice;
    const idx = new Map(); B.forEach(b => { b.ikey = b.key; idx.set(b.key, b); }); V.forEach(v => idx.set(v.ikey, v)); rec.payments.forEach(m => idx.set(m.ikey, m)); (rec.opening || []).forEach(o => idx.set(o.ikey, o));
    const ol = x => x.openLinks || [];
    const opSide = o => (o.head === 'BANK_CR' || o.head === 'BANK_DR' ? 'bank' : 'book'), opAmt = o => (o.head === 'BANK_CR' || o.head === 'DEPOSITED' ? o.amt : -o.amt);
    const kind = x => x.ikey[0];
    const name = x => (kind(x) === 'B' ? 'bank line ' + x.sl : kind(x) === 'V' ? 'voucher ' + x.blockId : kind(x) === 'O' ? 'item ' + x.no + ' of the last BRS' : 'Eazypay URN ' + x.tran);
    const names = xs => xs.map(name).join(', ');
    const viaEazypay = x => (kind(x) === 'B' ? x.type === 'Eazypay settlement' && !ol(x).length : kind(x) === 'V' ? !!(x.hasUrn || x.ez) && !ol(x).length : kind(x) === 'O' ? false : true);
    const partners = x => (kind(x) === 'B' ? V.filter(v => v.bank.includes(x.sl)).concat(x.setOff || [], ol(x)) : kind(x) === 'V' ? x.bank.map(sl => B[sl - 1]).concat(x.setOff || [], ol(x)) : kind(x) === 'O' ? ol(x).slice() : []);
    const linksOf = x => (kind(x) === 'B' && x.type === 'Eazypay settlement' && !ol(x).length ? (x.ezyShift ? ['S|' + x.ezy] : []) : kind(x) === 'V' && x.ez && !ol(x).length ? [] : partners(x).map(p => p.ikey));
    function group(x) { const seen = new Set([x]); const q = [x]; while (q.length) { const y = q.pop(); for (const p of partners(y)) if (!seen.has(p) && !viaEazypay(p)) { seen.add(p); q.push(p); } } return Array.from(seen); }
    const whereB = b => 'Line ' + b.sl + ' of the Bank lines sheet. In the bank statement: ' + (b.file ? b.file + ', ' : '') + 'page ' + b.page + '.';
    const whereV = v => 'ERP voucher ' + v.blockId + ' dated ' + nice(v.vchDate) + (v.provNo ? ' (Prov No ' + v.provNo + ')' : '') + '. Bank Book file: ' + v.file + '.';
    const who = v => v.paidBy || v.party || '';
    const MW = 'type its number in the "Match with" column of this row';

    function unlink(items) {
      for (const x of items) {
        const others = items.filter(y => y !== x);
        if (kind(x) === 'O') { x.status = 'Open'; x.openLinks = []; x.how = 'The reviewer said the match with ' + names(others) + ' is not correct.'; x.decided = 'Reject'; continue; }
        x.openLinks = []; x.opMatch = false;
        if (kind(x) === 'B') {
          const cr = x.amt > 0, amt = fmt(Math.abs(x.amt));
          x.status = 'Not in books'; x.how = ''; x.vouchers = []; x.lag = null; x._ins = null; x._many = null; x.setOff = null;
          x.note = 'The tool had paired this with ' + names(others) + '. The reviewer said that is not correct.';
          x.found = 'The bank ' + (cr ? 'received ' : 'paid out ') + amt + ' on ' + nice(x.date) + (x.party ? (cr ? ' from ' : ' to ') + x.party : '') + '. The tool had paired it with ' + names(others) + ', and the reviewer said that is not correct.';
          x.todo = 'Find the right ' + (cr ? 'receipt' : 'payment') + ' voucher in the ERP and ' + MW + '. If there is none, it still has to be booked.';
          x.where = whereB(x);
        } else {
          x.status = 'Not in bank'; x.how = ''; x.bank = []; x.setOff = null; x.instruments.forEach(i => { i.used = null; });
          x.note = 'The tool had paired this with ' + names(others) + '. The reviewer said that is not correct.';
          x.found = 'Voucher ' + x.blockId + ' shows ' + fmt(Math.abs(x.net)) + (x.net > 0 ? ' received from ' : ' paid to ') + who(x) + '. The tool had paired it with ' + names(others) + ', and the reviewer said that is not correct.';
          x.todo = 'Find the right line in the bank statement and ' + MW + ' (for example: line 45).';
          x.where = whereV(x);
        }
        x.decided = 'Reject';
      }
    }
    function link(items, how) {
      const bs = items.filter(x => kind(x) === 'B'), vs = items.filter(x => kind(x) === 'V'), os = items.filter(x => kind(x) === 'O');
      const diff = bs.reduce((a, b) => a + b.amt, 0) + os.filter(o => opSide(o) === 'bank').reduce((a, o) => a + opAmt(o), 0) - vs.reduce((a, v) => a + v.net, 0) - os.filter(o => opSide(o) === 'book').reduce((a, o) => a + opAmt(o), 0);
      for (const o of os) { o.status = 'Matched'; o.openLinks = bs.concat(vs); o.how = how; o.decided = 'Match'; }
      for (const x of bs.concat(vs)) { x.openLinks = os.slice(); if (os.length) x.opMatch = true; }
      const note = diff === 0 ? '' : 'Matched by the reviewer although the amounts differ by ' + fmt(Math.abs(diff)) + ' (bank ' + (diff > 0 ? 'more' : 'less') + ' than books).';
      for (const b of bs) {
        b.status = 'Matched'; b.vouchers = vs.map(v => v.blockId); b.setOff = vs.length || os.length ? null : bs.filter(o => o !== b);
        b.how = how + (vs.length ? '' : os.length ? ': clears ' + names(os) : ': set off against ' + names(bs.filter(o => o !== b))); b.note = note; b.decided = 'Match';
        b.lag = vs.length ? Math.max.apply(null, vs.map(v => Math.round((Date.parse(v.vchDate) - Date.parse(b.date)) / 86400000))) : null;
      }
      for (const v of vs) {
        v.status = 'Matched'; v.bank = bs.map(b => b.sl); v.setOff = bs.length || os.length ? null : vs.filter(o => o !== v);
        v.how = how + (bs.length ? '' : os.length ? ': clears ' + names(os) : ': set off against ' + names(vs.filter(o => o !== v))); v.note = note; v.decided = 'Match';
      }
      return diff;
    }

    const out = decisions.map(d => Object.assign({}, d, { with: (d.with || []).slice(), ok: false, result: '' }));
    const item = d => idx.get(d.key);
    const fail = (d, why) => { d.ok = false; d.result = 'Not applied: ' + why; };
    const done = (d, text) => { d.ok = true; d.result = text || 'Applied'; };
    for (const d of out) {
      const x = item(d);
      if (!x) { fail(d, 'this item is not in the files given this time. The decision is kept for when it is.'); d.absent = true; continue; }
      d.label = { what: name(x).replace(/^./, c => c.toUpperCase()), date: kind(x) === 'B' ? x.date : kind(x) === 'V' ? x.vchDate : kind(x) === 'O' ? x.date : x.settle, who: kind(x) === 'B' ? (x.party || x.type) : kind(x) === 'V' ? who(x) : x.party, amt: kind(x) === 'B' ? x.amt : kind(x) === 'V' ? x.net : kind(x) === 'O' ? opAmt(x) : x.amt };
      if (d.remark) x.remark = d.remark;
    }
    const live = d => !d.result;
    const E_MSG = 'it is tied to the bank through the Eazypay report (by its URN), which cannot be changed by hand. If the URN in the voucher is wrong, it has to be corrected in the ERP.';

    // 1. Reject
    for (const d of out.filter(d => live(d) && d.decision === 'Reject')) {
      const x = item(d);
      if (kind(x) === 'E') { fail(d, 'an Eazypay payment has no match to reject.'); continue; }
      if (kind(x) === 'B' && x.type === 'Eazypay settlement' && !ol(x).length) {
        if (!x.ezyShift) { fail(d, E_MSG); continue; }
        if (d.with.length && d.with[0] !== 'S|' + x.ezy) { fail(d, 'the tool now pairs this credit with another Eazypay total. Please look at it again.'); continue; }
        x.status = 'Needs review'; x.how = ''; x.decided = 'Reject';
        x.note = 'The reviewer did not accept the Eazypay total of ' + nice(x.ezy) + ' for this credit.';
        x.found = 'The bank credited ' + fmt(x.amt) + ' on ' + nice(x.date) + ' as one Eazypay amount. It equals the Eazypay report total of ' + nice(x.ezy) + ', but the reviewer did not accept that.';
        x.todo = 'Ask the bank which payments are inside this credit.'; x.where = whereB(x); x.ezyShift = 0;
        done(d); continue;
      }
      if (viaEazypay(x)) { fail(d, E_MSG); continue; }
      const cur = linksOf(x);
      if (!cur.length) { done(d, 'Applied (the item is open)'); continue; }
      if (d.with.length && !sameSet(cur, d.with)) { fail(d, 'the tool now pairs this item with ' + names(partners(x)) + ', which is not what was rejected. Please look at it again.'); continue; }
      d.with = cur; unlink(group(x)); done(d);
    }
    // 2. Match with (and Accept of something the tool no longer suggests by itself)
    const manual = (d, x, how) => {
      const targets = d.with.map(k => idx.get(k));
      if (targets.some(y => !y)) { fail(d, 'the other side (' + d.with.filter(k => !idx.get(k)).length + ' item) is not in the files given this time. The decision is kept.'); d.absent = true; return; }
      const all = Array.from(new Set([x].concat(targets)));
      const bad = all.find(y => kind(y) === 'E' || viaEazypay(y));
      if (bad) { fail(d, name(bad) + ' cannot be matched by hand: ' + E_MSG); return; }
      for (const y of all) {
        const cur = group(y).filter(p => !all.includes(p));
        if (cur.length && y.status !== 'Suggested') { fail(d, name(y) + ' is already matched to ' + names(cur) + '. Reject that match first, then match again.'); return; }
      }
      for (const y of all) { const g = group(y); if (g.some(p => !all.includes(p))) unlink(g.filter(p => !all.includes(p) || p === y)); }
      const diff = link(all, how);
      if (all.every(y => kind(y) === all[0].ikey[0]) && diff !== 0) done(d, 'Applied. The entries do not cancel out: ' + fmt(Math.abs(diff)) + ' is left.');
      else done(d, diff === 0 ? 'Applied' : 'Applied. Amounts differ by ' + fmt(Math.abs(diff)) + '.');
      d.diff = diff;
    };
    for (const d of out.filter(d => live(d) && d.decision === 'Match')) manual(d, item(d), 'Matched by the reviewer');
    // 3. Accept
    for (const d of out.filter(d => live(d) && d.decision === 'Accept')) {
      const x = item(d);
      if (kind(x) === 'E') { fail(d, 'an Eazypay payment has no suggested match. Use Keep open.'); continue; }
      if (kind(x) === 'B' && x.type === 'Eazypay settlement' && !ol(x).length) {
        if (x.status === 'Matched') { done(d, 'Applied (already matched)'); continue; }
        if (!x.ezyShift || (d.with.length && d.with[0] !== 'S|' + x.ezy)) { fail(d, 'there is no suggested Eazypay total for this credit any more.'); continue; }
        x.status = 'Matched'; x.how += ' - date confirmed by the reviewer'; x.note = x.note.replace(/\s*Please confirm\./, ''); x.decided = 'Accept'; d.with = ['S|' + x.ezy]; done(d); continue;
      }
      const cur = linksOf(x);
      if (cur.length && (!d.with.length || sameSet(cur, d.with))) {
        const g = viaEazypay(x) ? [x] : group(x);
        if (g.every(y => y.status !== 'Suggested')) { done(d, 'Applied (already matched)'); continue; }
        g.forEach(y => { if (y.status === 'Suggested') { y.status = 'Matched'; y.how += ' - confirmed by the reviewer'; y.note = String(y.note || '').replace(/\s*Please (confirm|check)\./, ''); y.decided = 'Accept'; } });
        d.with = cur; done(d); continue;
      }
      if (d.with.length) { manual(d, x, 'Matched by the reviewer'); continue; }
      fail(d, 'there is no suggested match on this item now. Use "Match with" to name the other side.');
    }
    for (const d of out.filter(d => live(d) && d.decision === 'Remark')) done(d, 'Remark kept');
    // 4. Keep open
    for (const d of out.filter(d => live(d) && d.decision === 'Keep open')) {
      const x = item(d);
      if (kind(x) === 'E') { if (x.status === 'Booked') done(d, 'Not needed any more: this payment is now booked.'); else { x.reviewed = 'Keep open'; done(d); } continue; }
      if (kind(x) === 'O') { if (x.status === 'Open') { x.reviewed = 'Keep open'; done(d); } else if (x.status === 'Suggested') fail(d, 'this item has a suggested match. Choose Accept or Reject for it.'); else done(d, 'Not needed any more: it has been matched since (' + names(partners(x)) + ').'); continue; }
      if (x.status === 'Matched') { done(d, 'Not needed any more: the tool has matched this item since (' + names(partners(x)) + ').'); continue; }
      if (x.status === 'Suggested') { fail(d, 'this item has a suggested match. Choose Accept or Reject for it.'); continue; }
      x.reviewed = 'Keep open'; done(d);
    }
    for (const d of out) { const x = item(d); if (x && d.ok) { x.decision = d.decision === 'Match' || d.decision === 'Remark' ? '' : d.decision; x.matchWith = d.decision === 'Match' ? d.with.slice() : null; } }

    R.recount(rec);
    rec.decisions = out;
    rec.describe = k => { if (k.slice(0, 2) === 'S|') return 'Eazypay total of ' + nice(k.slice(2)); const y = idx.get(k); return y ? (kind(y) === 'B' ? 'line ' + y.sl : kind(y) === 'V' ? y.blockId : kind(y) === 'O' ? 'last BRS ' + y.no : y.tran) : '(not in these files)'; };
    rec.linksOf = linksOf;
    rec.summary.decisions = { total: out.length, applied: out.filter(d => d.ok).length, notApplied: out.filter(d => !d.ok && !d.absent).length, absent: out.filter(d => d.absent).length,
      diff: out.reduce((a, d) => a + Math.abs(d.diff || 0), 0) };
    return rec;
  }

  return { OPTIONS, SEP, sig, isReview, isOldReview, readReview, combine, apply };
});
