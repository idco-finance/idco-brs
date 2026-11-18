/* IDCO Bank Statement to Excel - PDF text layer.
 * Turns a PDF (through pdf.js) into plain "tokens": a piece of text with its position on the page.
 * Everything after this file works on tokens only and knows nothing about PDFs. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.StatementPdfText = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function layOut(items, vp) {
    const m = vp.transform; const tokens = [];
    for (const it of items) {
      if (!it.str || !it.str.trim()) continue;
      const t = it.transform;
      const a = m[0] * t[0] + m[2] * t[1], b = m[1] * t[0] + m[3] * t[1];
      const x = m[0] * t[4] + m[2] * t[5] + m[4], y = m[1] * t[4] + m[3] * t[5] + m[5];
      // text that runs at an angle (watermarks, side notes) is not part of a statement table
      if (Math.abs(b) > 0.01 * Math.max(1, Math.abs(a)) || a <= 0) continue;
      const trail = it.str.length - it.str.replace(/\s+$/, '').length;
      const cw = it.str.length ? it.width / it.str.length : 0;
      tokens.push({ s: it.str.trim(), x0: x, x1: x + it.width - trail * cw, y: y, h: it.height || Math.abs(t[3]) || 8 });
    }
    return tokens;
  }

  /** Read every page of an opened pdf.js document into tokens. */
  async function readPages(pdfDoc, onProgress) {
    const pages = [];
    for (let n = 1; n <= pdfDoc.numPages; n++) {
      const page = await pdfDoc.getPage(n);
      const tc = await page.getTextContent();
      // Read the page the way its text runs. Normally that is how the page is shown; if a page has been
      // turned in the PDF so that its text runs sideways, try the other three directions.
      let vp = null, tokens = [];
      const total = tc.items.filter(it => it.str && it.str.trim()).length;
      for (const rot of [undefined, 0, 90, 180, 270]) {
        const v = rot === undefined ? page.getViewport({ scale: 1 }) : page.getViewport({ scale: 1, rotation: rot });
        const got = layOut(tc.items, v);
        if (!vp || got.length > tokens.length) { vp = v; tokens = got; }
        if (!total || tokens.length >= 0.6 * total) break;
      }
      pages.push({ no: n, width: vp.width, height: vp.height, tokens });
      if (onProgress) onProgress(n, pdfDoc.numPages);
      page.cleanup();
    }
    return pages;
  }

  return { readPages };
});
