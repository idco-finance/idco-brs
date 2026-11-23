/* IDCO BRS Tool - reading the two Excel inputs of the reconciliation:
 *   the Bank Book workbook made by the IDCO Bank Book Linearizer, and the Eazypay MIS report from the bank.
 * Works on an ExcelJS workbook that has already been loaded. Money is returned in paise. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementInputs = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function val(c) {
    if (c === null || c === undefined) return null;
    if (c instanceof Date) return c;
    if (typeof c === 'object') {
      if (c.result !== undefined) return val(c.result);
      if (c.text !== undefined) return String(c.text);
      if (c.richText) return c.richText.map(r => r.text).join('');
      if (c.error) return null;
      return null;
    }
    return c;
  }
  const str = v => { v = val(v); return v === null ? '' : (v instanceof Date ? iso(v) : String(v)).replace(/ /g, ' ').trim(); };
  function iso(v) {
    v = val(v);
    if (v === null || v === '') return '';
    if (v instanceof Date) return v.getUTCFullYear() + '-' + String(v.getUTCMonth() + 1).padStart(2, '0') + '-' + String(v.getUTCDate()).padStart(2, '0');
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[1] + '-' + m[2] + '-' + m[3];
    m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/); if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
    return '';
  }
  function paise(v) {
    v = val(v);
    if (v === null || v === '') return 0;
    if (typeof v === 'number') return Math.round(v * 100);
    const n = Number(String(v).replace(/,/g, '').trim());
    return isFinite(n) ? Math.round(n * 100) : 0;
  }
  function table(ws) {
    const rows = [];
    let head = null;
    ws.eachRow({ includeEmpty: false }, (row) => {
      const vals = []; row.eachCell({ includeEmpty: true }, (cell, col) => { vals[col - 1] = cell.value; });
      if (!head) { head = vals.map(v => str(v)); return; }
      const o = {}; head.forEach((h, i) => { if (h) o[h] = vals[i]; });
      rows.push(o);
    });
    return { head: head || [], rows };
  }

  /** What kind of workbook is this? 'bankbook' | 'eazypay' | null */
  function kindOf(wb) {
    const names = wb.worksheets.map(w => w.name);
    if (names.includes('1_Vouchers_Linear') && names.includes('5_Control_TieOut')) return 'bankbook';
    for (const ws of wb.worksheets) {
      const first = []; const r = ws.getRow(1); r.eachCell(c => first.push(str(c.value)));
      if (first.includes('TranID') && first.includes('Settlement Date') && first.includes('Transaction Amount')) return 'eazypay';
      if (first.includes('Vch. No & Date') || (ws.getRow(8) && str(ws.getRow(8).getCell(1).value) === 'Vch. No & Date')) return 'erp-raw';
    }
    return null;
  }

  function readBankBook(wb, fileName) {
    const ctl = {};
    table(wb.getWorksheet('5_Control_TieOut')).rows.forEach(r => { ctl[str(r.Metric)] = val(r.Value); });
    const out = { fileName, gl: str(ctl['Bank GL']), account: str(ctl['Bank Account No']), accountName: str(ctl['Bank Account Name']), division: str(ctl['Division']),
      from: iso(ctl['Period From']), to: iso(ctl['Period To']), tie: str(ctl['Arithmetic Check']), generated: str(ctl['Generated']),
      madeBy: (str(ctl['Generated']).match(/Linearizer v(\d+\.\d+\.\d+)/i) || [])[1] || '', debit: paise(ctl['ERP Summary Debit']), credit: paise(ctl['ERP Summary Credit']), blocks: [], urns: [] };
    for (const r of table(wb.getWorksheet('1_Vouchers_Linear')).rows) {
      const id = str(r['Block ID']); if (!id) continue;
      const b = { blockId: id, voucherId: str(r['Voucher ID']), type: str(r['Vch Type']), vchNo: str(r['Vch No']), vchDate: iso(r['Vch Date']), mrNo: str(r['MR No']), mrDate: iso(r['MR Date']),
        provNo: str(r['Prov No']), hasUrn: str(r['Has Eazypay URN']).toUpperCase() === 'YES', utr: str(r['UTR / Cheque No']), instrDate: iso(r['Actual Instr. / Bank Date']),
        instrCount: Number(val(r['Instr. Count'])) || 0, allDates: str(r['All Instr. Dates']), instrAmounts: str(r['Instr. Amounts']), utrCheck: str(r['UTR Check']),
        partyCode: str(r['Party Code']), party: str(r['Party Name']), glHead: str(r['Primary GL Head']), net: paise(r['Net Bank Impact (Dr - Cr)']), tds: paise(r['TDS Deducted (512001)']),
        suspenseRef: str(r['Ref CRV (Suspense)']), narration: str(r['Narration']), fileNo: str(r['File No']), file: fileName };
      // the instruments (one voucher can be paid in several parts)
      b.instruments = [];
      const utrs = b.utr ? b.utr.split('|').map(s => s.trim()) : [];
      if (b.instrCount > 1 && b.instrAmounts) {
        const amts = b.instrAmounts.split('|').map(s => paise(s)); const dates = b.allDates.split('|').map(s => iso(s.trim()));
        amts.forEach((a, i) => b.instruments.push({ amt: a, date: dates[i] || b.instrDate, utr: utrs[i] || (utrs.length === 1 ? utrs[0] : '') }));
      } else b.instruments.push({ amt: b.net, date: b.instrDate, utr: utrs[0] || '' });
      out.blocks.push(b);
    }
    const u = wb.getWorksheet('3_Eazypay_URNs');
    if (u) for (const r of table(u).rows) {
      const id = str(r['Block ID']); if (!id) continue;
      out.urns.push({ blockId: id, vchDate: iso(r['Vch Date']), tran: str(r['Normalized TranID (ICL->2)']), raw: str(r['URN (As Written)']), amt: paise(r['URN Line Amount']), party: str(r['Party Name']), partyCode: str(r['Party Code']) });
    }
    return out;
  }

  function readEazypay(wb, fileName) {
    const out = { fileName, rows: [] };
    for (const ws of wb.worksheets) {
      const t = table(ws);
      if (!t.head.includes('TranID')) continue;
      for (const r of t.rows) {
        const id = str(r.TranID); if (!id) continue;
        out.rows.push({ tran: id, amt: paise(r['Transaction Amount']), settle: iso(r['Settlement Date']), initiated: iso(r['Transaction Initiated On']), mode: str(r['Mode of Payment']),
          status: str(r.Status), partyCode: str(r['Party Code']), party: str(r['Party Name']), ref: str(r['Reference No']), uniqueRef: str(r['Unique Ref No']), van: str(r['VAN Number'] === undefined ? r['VAN Number '] : r['VAN Number']), file: fileName });
      }
    }
    return out;
  }

  return { kindOf, readBankBook, readEazypay, iso, paise, str };
});
