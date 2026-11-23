/* IDCO BRS Tool - helper columns.
 * Reads the narration of a line and fills three best-effort helpers: Type, Reference (UTR etc.) and Party.
 * These never change a date, an amount or a balance. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementClassify = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const clean = s => String(s || '').replace(/\s+/g, ' ').trim();
  const tidyParty = s => clean(String(s || '').replace(/^[\s\-\/*:.,]+|[\s\-\/*:.,]+$/g, ''));

  function enrich(row) {
    const n = clean(row.narration);
    const u = n.toUpperCase();
    const c = u.replace(/\s+/g, '');            // compact: wrapped reference numbers become whole again
    const isDebit = row.withdrawal > 0;
    let type = 'Other', ref = '', party = '';
    let m;

    if (/EZY~|EZY\/EZY/.test(c)) {
      type = 'Eazypay settlement';
      m = c.match(/EZY~(\d{8})(\d{4})?~(\d+)/);
      if (m) ref = 'EZY~' + m[1] + (m[2] || '') + '~' + m[3];
    } else if (/CHBK|CHARGEBACK|ACQBU\/ONUS\/CB|\/CB\/CC/.test(c)) {
      type = 'Chargeback';
      m = n.match(/(\d{9,})/); if (m) ref = m[1];
    } else if (/AUTOSWEEPTO|SWEEPTFRDR|SWEEPTRFDR|TRANSFERTOMOD|TOSWEEPDEPOSIT/.test(c)) {
      type = 'Sweep out';
      m = c.match(/(?:AUTOSWEEPTO|SWEEPTFRDR|SWEEPTRFDR)(\d{9,16})/) || c.match(/(\d{10,16})/); if (m) ref = m[1];
    } else if (/REVSWEEP|SWEEPTRFCREDT|SWEEPTRFCR|SWEEPDEPOSITBYTRANSFER|SWEEPFROM/.test(c)) {
      type = 'Sweep in';
      m = c.match(/^(\d{9,16}):?REVSWEEP/) || c.match(/(\d{10,16})/); if (m) ref = m[1];
    } else if (/CLOSUREPROCEEDS|CLOSUREINT|INT\.?PD|INTERESTPAID|INTERESTCREDIT|^INT\.?COLL/.test(c)) {
      type = 'FD closure / interest';
      m = c.match(/(\d{9,16})/); if (m) ref = m[1];
    } else if ((m = n.match(/^[\d.\-\s]{5,24}(?:to[\d.\-\s]{5,14})?\/(\d{6,12})\/(.+)$/i))) {
      type = 'Gateway receipt'; ref = m[1]; party = m[2];
    } else if (/(^|[^A-Z])ITD([^A-Z]|$)|CBDT|TAXE-?PAYMENT|GSTPAYMENT|\bTIN\b/.test(u) && isDebit) {
      type = 'Tax payment';
      m = c.match(/RTGS:?([A-Z]{4}[A-Z0-9]\d{12,18})/) || c.match(/(\d{9,})/); if (m) ref = m[1];
    } else if ((m = u.match(/\b(RTGS|NEFT)\b\s*(?:UTR\s*NO\s*:?\s*)?([-\/*:])?\s*(.*)$/))) {
      type = m[1];
      const delim = m[2]; const rest = n.slice(n.length - m[3].length);
      if (/UTR\s*NO/i.test(u)) {
        const k = c.match(/UTRNO:?([A-Z0-9]{8,25})/); if (k) ref = k[1];
        const o = n.match(/\bOF\s+(.+?)(?:\s+AT\s+\d|$)/i); if (o) party = o[1];
      } else if (delim) {
        const f = rest.split(delim === ':' ? '/' : delim).map(clean);
        if (delim === '*') {                       // NEFT*IFSC*reference*name
          let i = 0; if (/^[A-Z]{4}0[A-Z0-9]{6}$/i.test(f[0] || '')) i = 1;
          ref = (f[i] || '').replace(/\s+/g, ''); party = (f[i + 1] || '').replace(/\s+\d{9,}.*$/, '');
        } else {                                   // RTGS-reference-name-account-IFSC   or   NEFT/reference/name
          ref = (f[0] || '').replace(/\s+/g, ''); party = f[1] || '';
          if (!party && f.length > 2) party = f[2];
        }
      }
      if (!/^[A-Z0-9]{8,25}$/i.test(ref)) { const k = c.match(/(?:RTGS|NEFT)[-\/*:]?([A-Z0-9]{10,22})/); ref = k ? k[1] : ''; }
      if (type === 'RTGS') {
        // an RTGS UTR is always 22 characters: bank code + R + 17 digits
        const k = c.match(/[A-Z]{4}R[0-9A-Z]\d{16}/);
        if (k && k[0] !== ref) {
          ref = k[0];
          if (delim && delim !== '*') {
            const f = rest.split(delim === ':' ? '/' : delim).map(clean);
            const i = f.findIndex(x => x.replace(/\s+/g, '').toUpperCase().indexOf(ref) >= 0);
            const nx = f.slice(i + 1).find(x => /[A-Za-z]{3}/.test(x) && !/^[A-Z]{4}0[A-Z0-9]{6}$/i.test(x));
            if (i >= 0 && nx) party = nx;
          }
        }
      }
    } else if (/\bIMPS\b/.test(u)) {
      type = 'IMPS';
      const f = n.split('/').map(clean); const i = f.findIndex(x => /^IMPS$/i.test(x));
      if (i >= 0) { ref = (f[i + 1] || '').replace(/\s+/g, ''); party = f[i + 3] || f[i + 2] || ''; }
    } else if (/\bUPI\b/.test(u)) {
      type = 'UPI';
      const f = n.split('/').map(clean); const i = f.findIndex(x => /\bUPI$/i.test(x));
      const k = f.find(x => /^\d{12}$/.test(x)); if (k) ref = k;
      if (i >= 0) party = f.slice(i + 1).find(x => /[A-Za-z]{3}/.test(x) && !/^(CR|DR)$/i.test(x)) || '';
    } else if (/^INF\/|\/INFT\/|\bINFT\b/.test(u)) {
      type = 'Internal transfer';
      const f = n.split('/').map(clean); const i = f.findIndex(x => /^INFT$/i.test(x));
      if (i >= 0) { ref = (f[i + 1] || '').replace(/\s+/g, ''); party = f.length > i + 3 ? f[f.length - 1] : (f[i + 2] || ''); }
    } else if (/\bCMS\b|^CMP\b/.test(u)) {
      type = 'CMS collection';
      m = c.match(/CMS\/?(CMS\d{8,12})/) || c.match(/(CMS\d{8,12})/); if (m) ref = m[1];
      const f = n.split('/').map(clean); if (f.length > 1) party = f[f.length - 1];
      const o = n.match(/\bOF\s+(.+?)(?:\s+AT\s+\d|$)/i); if (o) party = o[1];
    } else if (/SBIVC|BILLPAY|\bBIL\//.test(c)) {
      type = 'Bill / e-payment';
      m = c.match(/SBIVC(\d{10,})/); if (m) ref = m[1];
      const k = n.match(/SBIVC\s*\d+\s+(.+)$/i); if (k) party = k[1];
    } else if (/TRFR ?TO|TRFR ?FROM|TRF ?TO|TRF ?FROM|^TRF\/|XFER|CHQ TRFR|TRANSFER|^(DEP|WDL) TFR\b/.test(u)) {
      type = 'Transfer';
      m = n.match(/TRFR?\s*(?:TO|FROM)\s*:?\s*(.+)$/i); if (m) party = m[1];
      const o = n.match(/\bOF\s+(.+?)(?:\s+AT\s+\d|$)/i); if (o) party = o[1];
      const k = n.match(/\b(\d{10,16})\b/); if (k) ref = k[1];
    } else if (/CHQ|CHEQUE|\bCLG\b|CLEARING/.test(u) || row.chq) {
      type = 'Cheque';
      const o = n.match(/\bOF\s+(.+?)(?:\s+AT\s+\d|$)/i); if (o) party = o[1];
    } else if (/CHARGES?|CHRG|CHGS|\bFEE\b|COMMISSION|\bGST\b/.test(u) && isDebit) {
      type = 'Bank charges';
    }
    if (!ref && row.chq) ref = row.chq;
    // a transfer or cheque that also shows a cheque number
    row.type = type; row.ref = clean(ref); row.party = tidyParty(party);
    return row;
  }

  return { enrich };
});
