/* IDCO Bank Statement to Excel - the library of standard bank-reconciliation situations.
 *
 * Every item that is not plainly matched gets a "Kind of item": the name an accountant would use for it in a BRS.
 * The tool still makes no accounting judgement: a kind is a label that says where an item of this sort is
 * usually shown, and the person doing the BRS decides.
 * Kinds marked seen:false have not yet occurred in IDCO's own files; they are recognised from the usual wording
 * of banks and are always worded as "please confirm".
 * All money in paise. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementScenarios = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // where an item of this kind is usually shown, starting from the balance as per Bank Book
  const ADD_BANK = 'Add: credited by the bank, not in the books';
  const LESS_BANK = 'Less: debited by the bank, not in the books';
  const LESS_BOOK = 'Less: received in the books, not credited by the bank';
  const ADD_BOOK = 'Add: paid in the books, not debited by the bank';
  const NONE = 'No effect on the BRS once confirmed';
  const DECIDE = 'To be settled before the BRS is signed';

  const KINDS = [
    // ---- in the bank statement, not in the books
    ['DIRECT_CR', 'Direct credit by a party, not in books', ADD_BANK, 'A party paid straight into the account (RTGS, NEFT, IMPS, UPI, transfer) and no receipt voucher was found.', true],
    ['EZY_UNBOOKED', 'Online collection (Eazypay), not in books', ADD_BANK, 'A payment made through the Eazypay gateway that the bank has credited, with no voucher carrying its URN.', true],
    ['TOKEN', 'Token / test credit of Re 1', ADD_BANK, 'A party sent Re 1 to test the account number before paying.', true],
    ['INTEREST', 'Interest credited by the bank', ADD_BANK, 'Interest put in by the bank itself, on the account or on a linked deposit.', true],
    ['FD_PROCEEDS', 'Deposit closure / maturity proceeds', ADD_BANK, 'Money returned by the bank when a linked deposit was closed or matured.', true],
    ['CHEQUE_CR', 'Cheque collected by the bank, not in books', ADD_BANK, 'A cheque or clearing credit with no receipt voucher.', false],
    ['CASH_CR', 'Cash deposited, not in books', ADD_BANK, 'Cash paid in at the bank with no receipt voucher.', false],
    ['RETURN_CR', 'Payment returned to the account', ADD_BANK, 'A payment that failed and was credited back by the bank (wrong account number, account closed).', false],
    ['OTHER_CR', 'Credited by the bank - nature not clear', ADD_BANK, 'A credit whose narration does not say what it is.', true],
    ['CHARGES', 'Bank charges', LESS_BANK, 'Charges, commission or GST on charges taken by the bank.', false],
    ['TDS_BANK', 'TDS deducted by the bank', LESS_BANK, 'Tax deducted by the bank, usually on interest.', false],
    ['CHARGEBACK', 'Chargeback of a card collection', LESS_BANK, 'A card payment received earlier through the gateway was taken back by the bank.', true],
    ['RETURN_DR', 'Cheque / collection returned unpaid', LESS_BANK, 'A cheque or collection credited earlier was dishonoured and debited back.', false],
    ['DIRECT_DR', 'Direct debit / standing instruction', LESS_BANK, 'An automatic debit (NACH, ECS, standing instruction) with no payment voucher.', false],
    ['TAX_DR', 'Tax paid through the bank, not in books', LESS_BANK, 'A tax payment (income tax, GST) debited by the bank with no payment voucher.', false],
    ['CHEQUE_DR', 'Cheque paid by the bank, not in books', LESS_BANK, 'A cheque debited by the bank with no payment voucher.', false],
    ['CASH_DR', 'Cash withdrawn, not in books', LESS_BANK, 'Cash taken out with no payment voucher.', false],
    ['PAY_UNBOOKED', 'Payment made from the bank, not in books', LESS_BANK, 'Money sent out (RTGS, NEFT, transfer) with no payment voucher.', true],
    ['CONTRA', 'Transfer between own accounts', DECIDE, 'Money moved to or from another account of the same organisation; the other leg has to be booked.', true],
    ['OTHER_DR', 'Debited by the bank - nature not clear', LESS_BANK, 'A debit whose narration does not say what it is.', false],
    ['SWEEP', 'Sweep to / from linked deposit', 'Shown as one net line: balance lying in linked deposits', 'Money moved by the bank between the account and its linked deposits. Not in the Bank Book.', true],
    // ---- in the books, not in the bank statement
    ['RECEIPT_NOT_CR', 'Receipt in books, not credited by the bank', LESS_BOOK, 'A receipt voucher for which no bank credit was found.', true],
    ['CHEQUE_NOT_CLEARED', 'Cheque / DD deposited, not yet cleared', LESS_BOOK, 'A cheque or draft entered as received, not yet credited by the bank.', false],
    ['PAYMENT_NOT_DR', 'Payment in books, not debited by the bank', ADD_BOOK, 'A payment voucher for which no bank debit was found.', false],
    ['ISSUED_NOT_PRESENTED', 'Cheque issued, not yet presented', ADD_BOOK, 'A cheque given out that the payee has not yet put into his bank.', false],
    ['STALE', 'Stale cheque (issued over 3 months ago)', ADD_BOOK, 'A cheque given out more than three months ago and never presented; it can no longer be paid.', false],
    ['EARLIER', 'Receipt of an earlier period', 'Opening item: belongs to the BRS of the earlier period', 'The money came in before the first statement given; no bank line can be expected here.', true],
    ['AFTER', 'Dated after the statement period', LESS_BOOK, 'The payment date in the voucher is later than the last statement given.', false],
    ['SUSPENSE', 'Suspense entry / its reversal', NONE, 'A suspense booking or the voucher that reverses it. It has no bank line of its own.', true],
    ['CANCELLED', 'Receipt cancelled in the books', NONE, 'A money receipt and the voucher that cancels it. Neither has a bank line.', true],
    ['NO_EFFECT', 'No effect on the bank', NONE, 'The voucher has no net amount in this bank account.', false],
    // ---- on both sides, but something differs
    ['AMOUNT_DIFFERS', 'Amount differs between bank and books', DECIDE, 'The same reference number is on both sides, with different amounts.', true],
    ['SHORT_RECEIPT', 'Received less than booked (TDS or deduction?)', DECIDE, 'The same party paid a little less than the voucher shows, around the same date.', false],
    ['PART_BOOKED', 'Collection only partly booked', ADD_BANK, 'More was received than the vouchers show; a share is still to be booked.', true],
    ['OVER_BOOKED', 'Booked more than received', DECIDE, 'The vouchers show more than the payment received.', true],
    ['PART_FOUND', 'Voucher with several payments, only some found', DECIDE, 'The voucher lists several payments; not all were found in the bank statement.', true],
    ['TRANSPOSED', 'Figures possibly interchanged', DECIDE, 'A bank amount and a book amount near the same date have the same digits in a different order.', false],
    ['WRONG_SIDE', 'Possibly entered on the wrong side', DECIDE, 'The bank shows a receipt and the books a payment of the same amount (or the other way round).', false],
    ['DOUBLE_BOOKED', 'Possibly booked twice in the books', DECIDE, 'A voucher looks the same as another voucher that is already matched.', false],
    ['DOUBLE_BANK', 'Possibly entered twice by the bank', DECIDE, 'Two bank lines of the same date carry the same amount and reference.', false],
    ['REVERSED_PAIR', 'Entered and reversed by the bank', NONE, 'A debit and a credit of the same amount that look like one entry and its reversal.', false],
    // ---- a match that a person has to confirm
    ['NAME_DIFFERS', 'Match to confirm - party name differs', NONE, 'Amount and date agree; the names on the two sides are not the same.', true],
    ['COMBINED', 'Match to confirm - several entries add up', NONE, 'One entry on one side equals several entries on the other.', true],
    ['SAME_AMOUNT', 'Match to confirm - same amount paid several times', NONE, 'Several payments of the same amount lie close together; paired in date order.', true],
    ['NO_DATE', 'Match to confirm - voucher shows no payment date', NONE, 'Matched on the amount alone.', true],
    ['DATE_SHIFT', 'Match to confirm - Eazypay date one day away', NONE, 'The settlement credit equals the Eazypay total of the neighbouring day.', true],
    ['MANY_FIT', 'More than one voucher fits', DECIDE, 'Several vouchers have this amount and date; a person has to pick.', true],
    ['FILE_MISSING', 'Cannot be checked - a file is missing', DECIDE, 'The Eazypay report or the bank statement that covers this item was not given.', true],
    ['UNTRACED', 'Voucher not traced to the bank', DECIDE, 'The voucher could not be followed to the statement.', true],
    ['OPENING', 'Clears an item of the last signed BRS', NONE, 'This entry of the new year settles an item that was open in the last signed BRS (same amount).', true]
  ].map(a => ({ code: a[0], name: a[1], brs: a[2], meaning: a[3], seen: a[4] }));
  const BY = {}; KINDS.forEach(k => { BY[k.code] = k; });

  const DAY = 86400000;
  const t = iso => (iso ? Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) : NaN);
  const days = (a, b) => Math.round((t(a) - t(b)) / DAY);
  const digits = p => String(Math.abs(p)).split('').sort().join('');
  const chequeLike = u => /^\d{6}$/.test(String(u || '').replace(/\s/g, '')) || /\b(CHQ|CHEQUE|DD)\b/i.test(String(u || ''));

  function bankKind(b, own) {
    const u = ' ' + String(b.narration || '').toUpperCase() + ' ', cr = b.amt > 0;
    if (b.opMatch) return 'OPENING';
    if (b.status === 'Sweep') return 'SWEEP';
    if (b.type === 'Chargeback') return 'CHARGEBACK';
    if (b.type === 'FD closure / interest') return /INT/.test(u.replace(/[^A-Z]/g, '')) ? 'INTEREST' : 'FD_PROCEEDS';
    if (b.code === 'EZY_NOFILE' || b.code === 'EZY_PARTFILE') return 'FILE_MISSING';
    if (b.status === 'Suggested') {
      if (b.ezyShift) return 'DATE_SHIFT';
      if (/^UTR/.test(b.how)) return 'AMOUNT_DIFFERS';
      if (b.how === 'Amount and date') return 'NAME_DIFFERS';
      if (b.how === 'Amount only') return 'NO_DATE';
      if (/add up/.test(b.how)) return 'COMBINED';
      return 'SAME_AMOUNT';
    }
    if (b.status === 'Needs review') return b._many ? 'MANY_FIT' : 'UNTRACED';
    if (b.status !== 'Not in books') return '';
    const isOwn = own.some(n => (cr ? String(b.party || '').toUpperCase().indexOf(n) === 0 : u.indexOf(n) >= 0));
    if (cr) {
      if (b.amt === 100) return 'TOKEN';
      if (/\b(RETURN(ED)?|RTN|REVERSAL|REVERSED|REFUND|REJECT(ED)?|BOUNCE)\b/.test(u)) return 'RETURN_CR';
      if (/\bCASH\b/.test(u)) return 'CASH_CR';
      if (/\bINTEREST\b|\bINT\.? ?(PD|PAID|CR)\b/.test(u)) return 'INTEREST';
      if (isOwn) return 'CONTRA';
      if (b.type === 'Cheque') return 'CHEQUE_CR';
      if (b.type !== 'Other') return 'DIRECT_CR';
      return 'OTHER_CR';
    }
    if (/\bTDS\b|TAX DEDUCTED/.test(u)) return 'TDS_BANK';
    if (b.type === 'Bank charges' || /\b(CHARGES?|CHRGS?|CHGS|COMMISSION)\b/.test(u)) return 'CHARGES';
    if (/\b(RETURN(ED)?|RTN|DISHONOU?R(ED)?|UNPAID|BOUNCE|REJECT(ED)?)\b/.test(u)) return 'RETURN_DR';
    if (/\b(NACH|ECS|ACH|SI|MANDATE)\b|STANDING INSTR|AUTO ?DEBIT/.test(u)) return 'DIRECT_DR';
    if (b.type === 'Tax payment') return 'TAX_DR';
    if (isOwn) return 'CONTRA';
    if (/\b(CASH|ATM)\b/.test(u)) return 'CASH_DR';
    if (b.type === 'Cheque') return 'CHEQUE_DR';
    if (b.type !== 'Other') return 'PAY_UNBOOKED';
    return 'OTHER_DR';
  }

  function voucherKind(v, lastBank, bankOf) {
    if (v.opMatch) return 'OPENING';
    switch (v.status) {
      case 'Earlier period': return 'EARLIER';
      case 'After statement period': return 'AFTER';
      case 'Suspense entry': return 'SUSPENSE';
      case 'Cancelled receipt': return 'CANCELLED';
      case 'No bank effect': return 'NO_EFFECT';
      case 'Suggested': { const b = v.bank.length ? bankOf(v.bank[0]) : null; return b && b.kindCode ? b.kindCode : 'SAME_AMOUNT'; }
      case 'Needs review':
        if (v.ez && v.ez.missing) return 'FILE_MISSING';
        if (/could not be tallied|not in the statements loaded/.test(v.note)) return 'FILE_MISSING';
        if (v.ez && v.ez.diff < 0) return 'PART_BOOKED';
        if (v.ez && v.ez.diff > 0) return 'OVER_BOOKED';
        if (/instruments of this voucher|released/.test(v.note)) return 'PART_FOUND';
        return 'UNTRACED';
      case 'Not in bank': {
        const chq = chequeLike(v.utr);
        if (v.net > 0) return chq ? 'CHEQUE_NOT_CLEARED' : 'RECEIPT_NOT_CR';
        const d = v.instrDate || v.vchDate;
        if (chq && lastBank && d && days(lastBank, d) > 90) return 'STALE';
        return chq ? 'ISSUED_NOT_PRESENTED' : 'PAYMENT_NOT_DR';
      }
    }
    return '';
  }

  /** Labels every open item of a reconcile() result and adds the "look at these two together" hints. */
  function annotate(rec, opts) {
    opts = opts || {};
    const own = (opts.ownNames || ['IDCO', 'ODISHA INDUSTRIAL INFRA']).map(s => s.toUpperCase());
    const B = rec.bankLines, V = rec.vouchers, fmt = rec.fmt, nice = rec.nice, nameMatch = opts.nameMatch || (() => 0);
    const lastBank = rec.summary.lastBank;
    const set = (x, code) => { x.kindCode = code; x.kind = code ? BY[code].name : ''; };
    for (const b of B) { set(b, bankKind(b, own)); b.age = b.status === 'Matched' || !lastBank ? null : days(lastBank, b.date); }
    for (const v of V) {
      set(v, voucherKind(v, lastBank, sl => B[sl - 1]));
      const d = v.instruments.map(x => x.date).filter(Boolean).sort()[0] || v.vchDate;
      v.age = v.status === 'Matched' || !lastBank || !d ? null : days(lastBank, d);
    }
    for (const m of rec.payments) { set(m, m.bankSl === null || m.status === 'Booked' ? '' : (m.status === 'Booked in part' ? 'PART_BOOKED' : 'EZY_UNBOOKED')); m.age = m.kindCode && lastBank && m.bankDate ? days(lastBank, m.bankDate) : null; }

    const openB = B.filter(b => b.status === 'Not in books' && !b.reviewed), openV = V.filter(v => v.status === 'Not in bank' && !v.reviewed);
    const hint = (x, code, found, todo) => { set(x, code); x.found = found; x.todo = todo; x.hinted = true; };
    const lineOf = b => 'line ' + b.sl;
    const how = 'If that is so, type ';
    const mw = ' in the "Match with" column of this row.';

    // 1. an entry and its reversal, both inside the bank statement
    const RET = /\b(RETURN(ED)?|RTN|REVERSAL|REVERSED|REV|REFUND|REJECT(ED)?|BOUNCE|DISHONOU?R(ED)?|UNPAID)\b/i;
    for (const d of openB.filter(b => b.amt < 0 && !b.hinted)) {
      const c = openB.find(c => !c.hinted && c.amt === -d.amt && Math.abs(days(c.date, d.date)) <= 7 &&
        (RET.test(c.narration) || RET.test(d.narration) || (d.ref && d.ref.length >= 9 && c.c.indexOf(d.ref.replace(/[^A-Za-z0-9]/g, '').toUpperCase()) >= 0)));
      if (!c) continue;
      const amt = fmt(c.amt);
      hint(d, 'REVERSED_PAIR', 'The bank paid out ' + amt + ' on ' + nice(d.date) + ' and put the same amount back on ' + nice(c.date) + ' (' + lineOf(c) + '). This looks like one entry and its reversal.', 'Read the two narrations. ' + how + '"' + lineOf(c) + '"' + mw + ' Nothing then needs to be booked.');
      hint(c, 'REVERSED_PAIR', 'The bank put in ' + amt + ' on ' + nice(c.date) + ' and took out the same amount on ' + nice(d.date) + ' (' + lineOf(d) + '). This looks like one entry and its reversal.', 'Read the two narrations. ' + how + '"' + lineOf(d) + '"' + mw + ' Nothing then needs to be booked.');
    }
    // 2. the same entry twice in the bank statement
    for (const a of openB) {
      if (a.hinted || !a.ref) continue;
      const o = openB.find(o => o !== a && !o.hinted && o.date === a.date && o.amt === a.amt && o.ref === a.ref);
      if (!o) continue;
      const txt = x => 'The bank shows ' + fmt(Math.abs(a.amt)) + ' twice on ' + nice(a.date) + ' with the same reference ' + a.ref + ' (' + lineOf(a) + ' and ' + lineOf(o) + ').';
      hint(a, 'DOUBLE_BANK', txt(), 'Check with the bank whether this is one payment entered twice, or two separate payments.');
      hint(o, 'DOUBLE_BANK', txt(), 'Check with the bank whether this is one payment entered twice, or two separate payments.');
    }
    const vDate = v => v.instruments.map(x => x.date).filter(Boolean).sort()[0] || v.vchDate;
    const nearV = (b, v, n) => Math.abs(days(vDate(v), b.date)) <= n || (days(v.vchDate, b.date) >= -3 && days(v.vchDate, b.date) <= 45 && !v.instrDate);
    const who = v => v.paidBy || v.party || '';
    // 3. figures interchanged (12,345 against 12,354)
    for (const b of openB.filter(b => !b.hinted)) {
      const v = openV.find(v => !v.hinted && (v.net > 0) === (b.amt > 0) && v.net !== b.amt && Math.abs(v.net) >= 1000 && digits(v.net) === digits(b.amt) && nearV(b, v, 5) && Math.abs(v.net - b.amt) * 5 < Math.abs(b.amt));
      if (!v) continue;
      hint(b, 'TRANSPOSED', 'The bank shows ' + fmt(Math.abs(b.amt)) + ' on ' + nice(b.date) + '. Voucher ' + v.blockId + ' shows ' + fmt(Math.abs(v.net)) + ': the same digits in a different order.', 'Open voucher ' + v.blockId + ' in the ERP and compare the amount with the bank statement; a figure may have been typed wrongly. ' + how + '"' + v.blockId + '"' + mw);
      hint(v, 'TRANSPOSED', 'Voucher ' + v.blockId + ' shows ' + fmt(Math.abs(v.net)) + '. The bank shows ' + fmt(Math.abs(b.amt)) + ' on ' + nice(b.date) + ' (' + lineOf(b) + '): the same digits in a different order.', 'Compare the voucher with the bank statement; a figure may have been typed wrongly. ' + how + '"' + lineOf(b) + '"' + mw);
    }
    // 4. received a little less than booked (TDS or another deduction)
    for (const b of openB.filter(b => b.amt > 0 && !b.hinted)) {
      const v = openV.find(v => !v.hinted && v.net > b.amt && nearV(b, v, 3) && ((v.tds > 0 && v.net - b.amt === v.tds) || ((v.net - b.amt) * 1000 <= v.net * 105 && nameMatch(v.party, b.party + ' ' + b.narration) > 0)));
      if (!v) continue;
      const diff = fmt(v.net - b.amt);
      hint(b, 'SHORT_RECEIPT', 'The bank received ' + fmt(b.amt) + ' on ' + nice(b.date) + (b.party ? ' from ' + b.party : '') + '. Voucher ' + v.blockId + ' of the same party shows ' + fmt(v.net) + ', which is ' + diff + ' more.', 'The party may have deducted TDS or another amount before paying. Open voucher ' + v.blockId + ' in the ERP. ' + how + '"' + v.blockId + '"' + mw);
      hint(v, 'SHORT_RECEIPT', 'Voucher ' + v.blockId + ' shows ' + fmt(v.net) + ' received from ' + who(v) + '. The bank received ' + fmt(b.amt) + ' on ' + nice(b.date) + ' (' + lineOf(b) + '), which is ' + diff + ' less.', 'The party may have deducted TDS or another amount before paying. ' + how + '"' + lineOf(b) + '"' + mw);
    }
    // 5. entered on the wrong side
    for (const b of openB.filter(b => !b.hinted)) {
      const v = openV.find(v => !v.hinted && v.net === -b.amt && Math.abs(days(v.vchDate, b.date)) <= 10);
      if (!v) continue;
      const side = b.amt > 0 ? ['received', 'payment'] : ['paid out', 'receipt'];
      hint(b, 'WRONG_SIDE', 'The bank ' + side[0] + ' ' + fmt(Math.abs(b.amt)) + ' on ' + nice(b.date) + '. Voucher ' + v.blockId + ' of ' + nice(v.vchDate) + ' is a ' + side[1] + ' of exactly this amount.', 'Open voucher ' + v.blockId + ' in the ERP and check whether it was entered as a ' + side[1] + ' by mistake.');
      hint(v, 'WRONG_SIDE', 'Voucher ' + v.blockId + ' is a ' + side[1] + ' of ' + fmt(Math.abs(v.net)) + '. The bank ' + side[0] + ' exactly this amount on ' + nice(b.date) + ' (' + lineOf(b) + ').', 'Check whether the voucher was entered as a ' + side[1] + ' by mistake.');
    }
    // 6. booked twice in the books
    const matchedIns = new Map();
    for (const v of V) if (v.status === 'Matched' && v.net > 0 && !v.hasUrn) v.instruments.forEach(x => { if (x.amt > 0 && x.date) { const k = x.amt + '|' + x.date; if (!matchedIns.has(k)) matchedIns.set(k, []); matchedIns.get(k).push(v); } });
    for (const v of openV.filter(v => v.net > 0 && !v.hinted && !v.hasUrn)) {
      const d = vDate(v); const tw = (matchedIns.get(v.net + '|' + d) || []).find(o => o.key !== v.key && ((o.partyCode && o.partyCode === v.partyCode) || (!o.partyCode && o.party === v.party)));
      if (!tw) continue;
      hint(v, 'DOUBLE_BOOKED', 'Voucher ' + v.blockId + ' shows ' + fmt(v.net) + ' received from ' + who(v) + ' on ' + nice(d) + '. Voucher ' + tw.blockId + ' has the same party, amount and payment date, and is already matched to bank line ' + tw.bank.join(', ') + '. The bank shows this amount only once.', 'Open both vouchers in the ERP and check whether the same receipt has been booked twice.');
    }

    // 7. several vouchers of one party fit one bank credit: the receipt may have been booked twice
    for (const b of B) {
      if (b.status !== 'Needs review' || !b._many) continue;
      const ids = Array.from(new Set(b._many.map(n => n.v.blockId)));
      if (ids.length < 2 || new Set(b._many.map(n => n.v.partyCode || n.v.party)).size !== 1) continue;
      set(b, 'DOUBLE_BOOKED');
      b.found = 'The bank received ' + fmt(b.amt) + ' on ' + nice(b.date) + (b.party ? ' from ' + b.party : '') + ' once. The ERP has ' + ids.length + ' vouchers of the same party with this amount and payment date: ' + ids.slice(0, 6).join(', ') + '.';
      b.todo = 'Open these vouchers in the ERP and check whether the same receipt has been booked more than once. If they are separate receipts, type the right voucher number in the "Match with" column.';
    }
    // totals of the open items, kind by kind (for the BRS headings sheet)
    const tally = new Map(); KINDS.forEach(k => tally.set(k.code, { n: 0, amt: 0 }));
    const count = (code, amt) => { if (!code) return; const o = tally.get(code); o.n++; o.amt += Math.abs(amt); };
    B.forEach(b => { if (b.status !== 'Matched') count(b.kindCode, b.amt); });
    V.forEach(v => { if (v.status !== 'Matched' && !(v.status === 'Suggested' && v.bank.length)) count(v.kindCode, v.net); });
    rec.payments.forEach(m => count(m.kindCode, m.amt));
    rec.kinds = KINDS.map(k => Object.assign({}, k, tally.get(k.code)));
    return rec;
  }

  return { KINDS, annotate, bankKind, voucherKind };
});
