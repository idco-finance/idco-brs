/* IDCO Bank Statement to Excel - the page.
 * No chrome.* calls here, so the same page also works when opened straight from the folder. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const drop = $('drop'), fileInput = $('file'), progress = $('progress'), barfill = $('barfill'), progtext = $('progtext');
  const problemsEl = $('problems'), toolbar = $('toolbar'), accountsEl = $('accounts'), summaryline = $('summaryline');
  const fmt = StatementMerge.fmt, nice = StatementMerge.niceDate;

  // pdf.js: the worker script is already loaded on this page, so no separate worker file is fetched
  // (this is what lets the page work both inside the extension and when opened from a folder).
  pdfjsLib.GlobalWorkerOptions.workerSrc = '../lib/pdf.worker.min.js';

  $('ver').textContent = 'IDCO Bank Statement to Excel v' + StatementWorkbook.VERSION;

  let statements = [];        // every statement read so far
  let books = [], misFiles = [];   // Bank Book (Linearizer) and Eazypay files, for the matching step
  let reviews = [];           // matching workbooks saved with the reviewer's decisions
  let brsFiles = [];          // last signed BRS workbooks (opening position)
  const asOnOf = {};          // chosen BRS month-end, per account
  let problems = [];          // files that could not be read
  let accounts = [];
  let busy = false;

  // ------------------------------------------------------------ choosing files
  drop.addEventListener('click', () => { if (!busy) fileInput.click(); });
  drop.addEventListener('keydown', e => { if ((e.key === 'Enter' || e.key === ' ') && !busy) { e.preventDefault(); fileInput.click(); } });
  fileInput.addEventListener('change', () => { const f = Array.from(fileInput.files); fileInput.value = ''; handleFiles(f); });
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); if (ev === 'drop' || e.target === document.documentElement || e.relatedTarget === null) drop.classList.remove('over'); }));
  document.addEventListener('drop', e => { if (e.dataTransfer && e.dataTransfer.files.length) handleFiles(Array.from(e.dataTransfer.files)); });
  $('reset').addEventListener('click', () => { statements = []; problems = []; accounts = []; books = []; misFiles = []; reviews = []; brsFiles = []; render(); });
  $('dlall').addEventListener('click', async () => { for (const a of accounts) { await download(a); await sleep(350); } });

  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // ------------------------------------------------------------ password box
  function askPassword(fileName, wrong) {
    return new Promise(resolve => {
      const modal = $('pwmodal'), input = $('pwinput'), msg = $('pwmsg');
      $('pwfile').textContent = fileName;
      msg.textContent = wrong ? 'That password did not open the file. Please try again.' : 'Type the password given by the bank to open it.';
      msg.classList.toggle('wrong', !!wrong);
      input.value = ''; input.type = 'password'; $('pwshow').checked = false;
      modal.hidden = false; input.focus();
      const done = v => { modal.hidden = true; $('pwok').onclick = $('pwskip').onclick = input.onkeydown = null; resolve(v); };
      $('pwok').onclick = () => { if (input.value) done(input.value); else input.focus(); };
      $('pwskip').onclick = () => done(null);
      input.onkeydown = e => { if (e.key === 'Enter' && input.value) done(input.value); if (e.key === 'Escape') done(null); };
      $('pwshow').onchange = () => { input.type = $('pwshow').checked ? 'text' : 'password'; };
    });
  }

  async function openPdf(file) {
    const data = new Uint8Array(await file.arrayBuffer());
    const task = pdfjsLib.getDocument({ data, isEvalSupported: false, useSystemFonts: false, stopAtErrors: false });
    let skipped = false;
    task.onPassword = async (update, reason) => {
      const pw = await askPassword(file.name, reason === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD);
      if (pw === null) { skipped = true; task.destroy(); } else update(pw);
    };
    try { return await task.promise; }
    catch (e) { if (skipped) { const er = new Error('skipped'); er.skipped = true; throw er; } throw e; }
  }

  // ------------------------------------------------------------ reading
  async function handleFiles(files) {
    if (busy) return;
    const pdfs = files.filter(f => /\.pdf$/i.test(f.name) || f.type === 'application/pdf');
    const sheets = files.filter(f => /\.xlsx$/i.test(f.name));
    const oldSheets = files.filter(f => /\.xls$/i.test(f.name));
    const others = files.filter(f => !pdfs.includes(f) && !sheets.includes(f) && !oldSheets.includes(f));
    others.forEach(f => problems.push({ name: f.name, text: 'This is not a PDF or an Excel file. Drop bank statements as PDF, and the Bank Book and Eazypay files as .xlsx.' }));
    if (!pdfs.length && !sheets.length && !oldSheets.length) { render(); return; }
    busy = true; progress.hidden = false; setProgress(0, 'Starting…');
    for (const file of oldSheets) {
      setProgress(0, 'Opening ' + file.name + '…'); await sleep(0);
      try {
        const x = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array', cellDates: true });
        const grid = StatementOpening.fromSheetJS(XLSX, x);
        if (StatementOpening.isBRS(grid)) { brsFiles = brsFiles.filter(b => b.fileName !== file.name); brsFiles.push({ fileName: file.name, grid }); }
        else problems.push({ name: file.name, text: 'This old-style Excel file is not a BRS workbook. Bank Books and Eazypay reports have to be .xlsx files.' });
      } catch (e) { problems.push({ name: file.name, text: 'This Excel file could not be opened' + (e && e.message ? ' (' + e.message + ').' : '.') }); }
    }
    for (const file of sheets) {
      setProgress(0, 'Opening ' + file.name + '…'); await sleep(0);
      try {
        const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
        const kind = StatementDecisions.isReview(wb) ? 'review' : StatementInputs.kindOf(wb);
        if (kind === 'review' && StatementDecisions.isOldReview(wb)) problems.push({ name: file.name, text: 'This matching Excel was made by an older version of the tool, which had no place to keep decisions. Download a new matching Excel from this page and decide in that one.' });
        else if (kind === 'review') { const r = StatementDecisions.readReview(wb, file.name); reviews = reviews.filter(x => x.fileName !== r.fileName); reviews.push(r); }
        else if (kind === 'bankbook') { const b = StatementInputs.readBankBook(wb, file.name); books = books.filter(x => x.fileName !== b.fileName); books.push(b); }
        else if (kind === 'eazypay') { const m = StatementInputs.readEazypay(wb, file.name); misFiles = misFiles.filter(x => x.fileName !== m.fileName); misFiles.push(m); }
        else if (kind === 'erp-raw') problems.push({ name: file.name, text: 'This is the ERP\'s own Bank Book export. Please drop the clean workbook made by the Bank Book Linearizer instead (its name starts with IDCO_Clean_BankBook).' });
        else if (StatementOpening.isBRS(StatementOpening.fromExcelJS(wb))) { brsFiles = brsFiles.filter(b => b.fileName !== file.name); brsFiles.push({ fileName: file.name, grid: StatementOpening.fromExcelJS(wb) }); }
        else problems.push({ name: file.name, text: 'This Excel file is not a Linearizer Bank Book, an Eazypay report, a BRS workbook or a matching Excel made by this tool, so it was not used.' });
      } catch (e) { problems.push({ name: file.name, text: 'This Excel file could not be opened' + (e && e.message ? ' (' + e.message + ').' : '.') }); }
    }
    for (let i = 0; i < pdfs.length; i++) {
      const file = pdfs[i];
      const base = i / pdfs.length, span = 1 / pdfs.length;
      setProgress(base, 'Opening ' + file.name + '…');
      try {
        const doc = await openPdf(file);
        const pages = await StatementPdfText.readPages(doc, (n, total) => setProgress(base + span * 0.9 * n / total, file.name + ': reading page ' + n + ' of ' + total));
        await sleep(0);
        const st = StatementReader.read(pages);
        try { doc.destroy(); } catch (e) { /* ignore */ }
        st.fileName = file.name;
        if (!st.rows || !st.rows.length) {
          const letters = pages.reduce((n, p) => n + p.tokens.length, 0);
          problems.push({ name: file.name, text: letters < 20
            ? 'No text could be read. This looks like a scanned or photographed statement, which this tool cannot read.'
            : st.noBalanceColumn ? 'This statement has no running Balance column, so its lines cannot be proved. Ask the bank for a statement that shows the balance after every transaction.'
            : 'No transaction table was found. Is this a bank statement? (Headings like Date, Withdrawal, Deposit and Balance were not found.)' });
        } else {
          st.rows.forEach(StatementClassify.enrich);
          statements = statements.filter(s => s.fileName !== st.fileName || s.rows.length !== st.rows.length || s.opening !== st.opening);
          statements.push(st);
        }
      } catch (e) {
        if (e && e.skipped) problems.push({ name: file.name, text: 'Skipped: the password was not given.' });
        else problems.push({ name: file.name, text: 'This file could not be opened as a PDF' + (e && e.message ? ' (' + e.message + ').' : '.') });
      }
    }
    setProgress(1, 'Matching…'); await sleep(0);
    accounts = StatementMerge.groupAndJoin(statements);
    runMatching();
    busy = false; progress.hidden = true;
    render();
  }
  // ------------------------------------------------------------ matching (when a Bank Book of the same account is loaded)
  const last4 = s => String(s || '').replace(/\D/g, '').slice(-4);
  function runMatching(only) {
    for (const a of accounts) {
      if (only && a !== only) continue;
      const mine = books.filter(b => last4(b.account) && last4(b.account) === last4(a.account));
      a.recon = null;
      if (!mine.length) continue;
      try {
        const rec = StatementReconcile.reconcile(a, mine, misFiles);
        let op = null;
        for (const bf of brsFiles) { const o = StatementOpening.read(bf.grid, a.account, bf.fileName); if (o) { op = o; break; } }
        if (op) StatementBRS.matchOpening(rec, op);
        const rv = reviews.filter(r => last4(r.account) === last4(a.account)); const c = StatementDecisions.combine(rv);
        StatementDecisions.apply(rec, c.decisions, StatementReconcile); rec.reviewProblems = c.problems; rec.reviewFiles = rv.map(r => r.fileName);
        StatementScenarios.annotate(rec, { nameMatch: StatementReconcile.nameMatch });
        rec.months = StatementBRS.months(a, rec); rec.op = op;
        if (op && rec.months.length) {
          const want = asOnOf[a.key]; const asOn = rec.months.includes(want) ? want : rec.months[rec.months.length - 1];
          asOnOf[a.key] = asOn; rec.brs = StatementBRS.build(rec, op, a, asOn);
        }
        rec.pendingCount = StatementReconBook.pendingOf(rec).length;
        a.recon = rec;
        a.reconFile = rec.brs && rec.brs.ready ? 'IDCO_BRS_' + a.short + '_' + a.account + '_as_on_' + rec.brs.asOn + '.xlsx' : a.fileName.replace('IDCO_Statement_', 'IDCO_Reconciliation_');
      } catch (e) { problems.push({ name: 'Matching', text: 'The matching step failed for A/c ' + a.account + (e && e.message ? ' (' + e.message + ').' : '.') }); }
    }
  }
  function setProgress(frac, text) { barfill.style.width = Math.round(frac * 100) + '%'; progtext.textContent = text; }

  // ------------------------------------------------------------ showing
  function el(tag, cls, text) { const e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }

  function render() {
    problemsEl.innerHTML = ''; problemsEl.hidden = !problems.length;
    for (const p of problems) { const d = el('div', 'problem'); d.appendChild(el('b', '', p.name + ': ')); d.appendChild(document.createTextNode(p.text)); problemsEl.appendChild(d); }
    accountsEl.innerHTML = '';
    const has = accounts.length > 0;
    toolbar.hidden = !has && !problems.length && !books.length && !misFiles.length && !reviews.length;
    $('dlall').hidden = accounts.length < 2;
    drop.classList.toggle('small', has);
    drop.querySelector('.drop-title').textContent = has ? 'Add more statements' : 'Drop bank statement PDFs here';
    $('how').hidden = has;
    if (has) {
      const nPass = accounts.filter(a => a.pass).length;
      summaryline.textContent = accounts.length + (accounts.length === 1 ? ' account' : ' accounts') + ' read from ' + statements.length + (statements.length === 1 ? ' statement' : ' statements') +
        (nPass === accounts.length ? '. Everything agrees.' : '. ' + (accounts.length - nPass) + ' need' + (accounts.length - nPass === 1 ? 's' : '') + ' a look.');
    } else summaryline.textContent = '';
    for (const a of accounts) accountsEl.appendChild(card(a));
    const orphan = books.filter(b => !accounts.some(a => last4(a.account) === last4(b.account)));
    const lonely = reviews.filter(r => !accounts.some(a => a.recon && last4(a.account) === last4(r.account)));
    if (orphan.length || (misFiles.length && !books.length) || lonely.length) {
      const d = el('div', 'acc waiting');
      d.appendChild(el('b', '', 'Waiting for the other files'));
      if (orphan.length) d.appendChild(el('p', '', orphan.length + ' Bank Book file(s) for A/c ' + Array.from(new Set(orphan.map(b => b.account))).join(', ') + ' are loaded. Drop the bank statement PDFs of the same account to match them.'));
      if (misFiles.length && !books.length) d.appendChild(el('p', '', misFiles.length + ' Eazypay file(s) are loaded. Drop the Bank Book workbook(s) from the Linearizer to start matching.'));
      if (lonely.length) d.appendChild(el('p', '', 'Decisions were read from ' + lonely.map(r => r.fileName).join(', ') + ' (A/c ' + Array.from(new Set(lonely.map(r => r.account || 'not known'))).join(', ') + '). Drop the bank statements and the Bank Book workbooks of that account as well, so that the decisions can be applied.'));
      accountsEl.appendChild(d);
    }
  }

  function card(a) {
    const c = el('article', 'acc');
    const head = el('div', 'acc-head');
    const title = el('div', 'acc-title');
    const h = el('h2'); h.appendChild(document.createTextNode(a.bank + '  ·  A/c ' + (a.account || 'number not found') + '  ')); h.appendChild(el('span', 'pill ' + (a.pass ? 'pass' : 'check'), a.pass ? 'PASS' : 'CHECK'));
    title.appendChild(h);
    title.appendChild(el('div', 'sub', nice(a.firstDate) + ' to ' + nice(a.lastDate) + '  ·  ' + a.rows.length + ' lines  ·  ' + a.files.length + (a.files.length === 1 ? ' statement' : ' statements joined')));
    head.appendChild(title);
    const btn = el('button', 'btn primary', 'Download Excel'); btn.addEventListener('click', () => download(a)); head.appendChild(btn);
    c.appendChild(head);

    c.appendChild(el('div', 'verdict', a.pass
      ? 'Every line agrees with the running balance and with the figures printed by the bank.'
      : 'Something did not agree. The Excel file can still be downloaded, but check the points below first.'));

    const figs = el('div', 'figs');
    const fig = (k, v, n) => { const f = el('div', 'fig'); f.appendChild(el('div', 'k', k)); f.appendChild(el('div', 'v', v)); if (n) f.appendChild(el('div', 'n', n)); figs.appendChild(f); };
    fig('Opening balance', fmt(a.opening), nice(a.firstDate));
    fig('Withdrawals', fmt(a.totals.withdrawals), a.totals.nWithdrawals + ' lines');
    fig('Deposits', fmt(a.totals.deposits), a.totals.nDeposits + ' lines');
    fig('Closing balance', fmt(a.closing), nice(a.lastDate));
    c.appendChild(figs);

    if (!a.pass) {
      const look = el('div', 'look'); look.appendChild(el('b', '', 'Please look at:'));
      const ul = el('ul');
      a.joins.filter(j => !j.ok).forEach(j => ul.appendChild(el('li', '', j.text)));
      a.files.forEach(f => f.checks.filter(k => !k.ok).forEach(k => ul.appendChild(el('li', '', (a.files.length > 1 ? f.fileName + ': ' : '') + k.name + '. ' + k.detail))));
      a.rows.filter(r => !r.ok).slice(0, 8).forEach(r => ul.appendChild(el('li', '', 'Page ' + r.page + ', ' + nice(r.date) + ': ' + r.note)));
      const more = a.rows.filter(r => !r.ok).length - 8; if (more > 0) ul.appendChild(el('li', '', '…and ' + more + ' more, marked in the Check column of the Excel sheet.'));
      look.appendChild(ul); c.appendChild(look);
    }
    const warns = []; a.files.forEach(f => (f.warnings || []).forEach(w => warns.push((a.files.length > 1 ? f.fileName + ': ' : '') + w)));
    if (warns.length) { const look = el('div', 'look'); const ul = el('ul'); warns.forEach(w => ul.appendChild(el('li', '', w))); look.appendChild(el('b', '', 'Note:')); look.appendChild(ul); c.appendChild(look); }

    c.appendChild(reconPanel(a));
    const det = el('details'); det.appendChild(el('summary', '', 'What was checked'));
    const tb = el('table', 'checks');
    const addRow = (name, ok, detail) => { const tr = el('tr'); tr.appendChild(el('td', '', name)); const r = el('td', 'r'); r.appendChild(el('span', 'pill ' + (ok ? 'pass' : 'check'), ok ? 'PASS' : 'CHECK')); tr.appendChild(r); tr.appendChild(el('td', 'd', detail || '')); tb.appendChild(tr); };
    a.files.forEach(f => f.checks.forEach(k => addRow((a.files.length > 1 ? f.fileName + ': ' : '') + k.name, k.ok, k.detail)));
    a.joins.forEach(j => addRow('Joining the statements', j.ok, j.text));
    det.appendChild(tb); c.appendChild(det);

    if (a.files.length > 1) {
      const d2 = el('details'); d2.appendChild(el('summary', '', 'Statements used (' + a.files.length + ')'));
      const t2 = el('table', 'checks');
      a.files.forEach(f => { const tr = el('tr'); tr.appendChild(el('td', '', f.fileName)); tr.appendChild(el('td', 'd', nice(f.firstDate) + ' to ' + nice(f.lastDate) + '  ·  ' + f.lines + ' lines  ·  opening ' + fmt(f.opening) + '  ·  closing ' + fmt(f.closing) + (f.note ? '  ·  ' + f.note : ''))); t2.appendChild(tr); });
      d2.appendChild(t2); c.appendChild(d2);
    }
    return c;
  }

  function reconPanel(a) {
    const box = el('div', 'recon');
    if (!a.recon) {
      box.classList.add('empty');
      box.appendChild(el('b', '', 'Reconcile with the books'));
      box.appendChild(el('span', '', '  Drop the Bank Book workbook(s) from the Linearizer for this account, and the Eazypay files if it has any, on the box above.'));
      return box;
    }
    const s = a.recon.summary; const n = k => (s.bankBy[k] ? s.bankBy[k].n : 0), vn = k => (s.vBy[k] ? s.vBy[k].n : 0);
    const head = el('div', 'recon-head');
    const t = el('div'); t.appendChild(el('b', '', 'Matching with the Bank Book'));
    t.appendChild(el('div', 'sub', s.bookFiles.length + ' Bank Book file(s), vouchers ' + nice(s.firstBook) + ' to ' + nice(s.lastBook) + '  ·  ' + s.misFiles.length + ' Eazypay file(s), ' + s.eazypay.payments + ' payments'));
    head.appendChild(t);
    const btn = el('button', 'btn primary', a.recon.brs && a.recon.brs.ready ? 'Download BRS Excel' : 'Download matching Excel'); btn.addEventListener('click', () => downloadRecon(a)); head.appendChild(btn);
    box.appendChild(head);
    box.appendChild(brsPanel(a));
    const grid = el('div', 'figs');
    const fig = (k, v, note, cls) => { const f = el('div', 'fig ' + (cls || '')); f.appendChild(el('div', 'k', k)); f.appendChild(el('div', 'v', String(v))); if (note) f.appendChild(el('div', 'n', note)); grid.appendChild(f); };
    fig('Bank lines matched', n('Matched'), 'of ' + s.bankLines, 'good');
    fig('Suggested', n('Suggested'), 'please confirm', 'warn');
    fig('Needs review', n('Needs review') + vn('Needs review'), n('Needs review') + ' bank, ' + vn('Needs review') + ' vouchers', 'bad');
    fig('In bank, not in books', n('Not in books'), fmt((s.bankBy['Not in books'] || { dep: 0 }).dep) + ' credits', 'bad');
    fig('In books, not in bank', vn('Not in bank'), '', 'bad');
    fig('Eazypay payments not booked', s.eazypay.notInBooks, fmt(s.eazypay.notInBooksAmt), 'bad');
    fig('Sweeps to / from deposits', n('Sweep'), 'net out ' + fmt(s.sweepOut - s.sweepIn));
    fig('Vouchers of an earlier period', vn('Earlier period'), 'no bank line expected');
    fig('Waiting for a decision', a.recon.pendingCount, 'rows in the To review sheet', a.recon.pendingCount ? 'warn' : 'good');
    box.appendChild(grid);
    const dc = s.decisions;
    if (dc && dc.total) {
      const dl = el('div', 'look'); dl.appendChild(el('b', '', 'Decisions read from ' + a.recon.reviewFiles.join(', ') + ':'));
      const ul = el('ul');
      ul.appendChild(el('li', '', dc.applied + ' of ' + dc.total + ' applied.' + (dc.notApplied ? ' ' + dc.notApplied + ' could not be applied; the Decisions sheet of the new Excel says why.' : '') + (dc.absent ? ' ' + dc.absent + ' are kept for later, because their items are not in the files given this time.' : '')));
      a.recon.reviewProblems.forEach(w => ul.appendChild(el('li', '', w)));
      dl.appendChild(ul); box.appendChild(dl);
    } else if (a.recon.reviewProblems && a.recon.reviewProblems.length) { const dl = el('div', 'look'); const ul = el('ul'); a.recon.reviewProblems.forEach(w => ul.appendChild(el('li', '', w))); dl.appendChild(el('b', '', 'Decisions:')); dl.appendChild(ul); box.appendChild(dl); }
    else if (a.recon.reviewFiles.length) box.appendChild(el('div', 'fine', 'The saved Excel ' + a.recon.reviewFiles.join(', ') + ' was read. It holds no decisions yet.'));
    if (s.notes.length) { const look = el('div', 'look'); const ul = el('ul'); s.notes.forEach(w => ul.appendChild(el('li', '', w))); look.appendChild(el('b', '', 'Note:')); look.appendChild(ul); box.appendChild(look); }
    box.appendChild(el('div', 'fine', (a.recon.brs && a.recon.brs.ready ? 'The BRS is prepared by the tool from the files given; check its items before signing. ' : 'This is a matching sheet for review, not the signed BRS. ') + 'Open the "To review" sheet of the Excel file, fill the Decision column, save the file and drop it back here next time.'));
    return box;
  }

  function brsPanel(a) {
    const rec = a.recon, R = rec.brs, wrap = el('div', 'brs');
    if (!rec.op) {
      wrap.appendChild(el('b', '', 'BRS as on a month'));
      wrap.appendChild(el('div', 'sub', brsFiles.length ? 'The BRS workbook given (' + brsFiles.map(b => b.fileName).join(', ') + ') has no column for A/c ' + a.account + '.' : 'To get the BRS statement, also drop the last signed BRS workbook (the yearly BRS Excel as on 31 March).'));
      return wrap;
    }
    const top = el('div', 'brs-head');
    top.appendChild(el('b', '', 'BRS as on '));
    const sel = document.createElement('select'); sel.className = 'month';
    rec.months.forEach(m => { const o = document.createElement('option'); o.value = m; o.textContent = nice(m); if (R && R.asOn === m) o.selected = true; sel.appendChild(o); });
    sel.addEventListener('change', () => { asOnOf[a.key] = sel.value; runMatching(a); render(); });
    top.appendChild(sel);
    top.appendChild(el('span', 'sub', '  from 1 April. Last signed BRS: ' + rec.op.fileName + ' (as on ' + nice(rec.op.asOn) + ').'));
    wrap.appendChild(top);
    if (!R) { wrap.appendChild(el('div', 'sub', 'No complete month is covered by both the statements and the Bank Book yet.')); return wrap; }
    if (!R.ready) { const ul = el('ul'); R.checks.filter(c => !c.pass).forEach(c => ul.appendChild(el('li', '', c.what + ': ' + c.detail))); const lk = el('div', 'look'); lk.appendChild(el('b', '', 'The BRS cannot be made yet:')); lk.appendChild(ul); wrap.appendChild(lk); return wrap; }
    const tbl = el('table', 'brs-table');
    const row = (label, amt, cls, n) => { const tr = el('tr', cls || ''); tr.appendChild(el('td', '', label)); tr.appendChild(el('td', 'n', n === undefined ? '' : String(n))); tr.appendChild(el('td', 'amt', amt === null ? '' : fmt(amt))); tbl.appendChild(tr); };
    row('Balance as per Bank Book as on ' + nice(R.asOn), R.book, 'tot');
    R.order.forEach(h => { const t = R.totals[h]; if (!t.n && /^DIFF|^BOOKNET/.test(h)) return; row(R.heads[h], t.amt, '', t.n); });
    if (R.sweepN) row((R.sweep >= 0 ? 'Add' : 'Less') + ': Net amount moved between the account and linked deposits (sweep)', Math.abs(R.sweep), '', R.sweepN);
    row('Balance as per bank statement (worked out)', R.computed, 'tot');
    row('Balance as per bank statement (printed by the bank)', R.bank, 'tot');
    row('Difference', R.diff, R.diff === 0 ? 'tot good' : 'tot bad');
    wrap.appendChild(tbl);
    const notes = [];
    if (!R.tried) notes.push('NOT YET TRIED ON THIS ACCOUNT: the tool has so far been checked only on ' + R.triedList + '. Check every figure before relying on it.');
    if (R.suggested) notes.push(R.suggested + ' open item(s) have a match suggested by the tool. They stay open until accepted in the To review sheet.');
    notes.push('Last signed BRS: ' + R.opening.n + ' items, ' + R.opening.cleared + ' settled by ' + nice(R.asOn) + ', ' + R.opening.stillOpen + ' still open.');
    R.checks.filter(c => !c.pass).forEach(c => notes.push('CHECK - ' + c.what + ': ' + c.detail));
    const ul = el('ul'); notes.forEach(t => ul.appendChild(el('li', '', t))); const nb = el('div', R.checks.some(c => !c.pass) || !R.tried ? 'look' : 'fine'); nb.appendChild(ul); wrap.appendChild(nb);
    return wrap;
  }

  // ------------------------------------------------------------ Excel
  async function downloadRecon(a) {
    const wb = StatementReconBook.build(ExcelJS, a.recon, new Date(), StatementWorkbook.VERSION);
    const buf = await wb.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = document.createElement('a'); link.href = url; link.download = a.reconFile; document.body.appendChild(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
  async function download(a) {
    const wb = StatementWorkbook.build(ExcelJS, a, new Date());
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a'); link.href = url; link.download = a.fileName;
    document.body.appendChild(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  // for automated tests
  window.__statementTool = { handleFiles, getBooks: () => books, getReviews: () => reviews, getBrsFiles: () => brsFiles, setAsOn: (key, d) => { asOnOf[key] = d; runMatching(); render(); }, getMis: () => misFiles, getAccounts: () => accounts, getProblems: () => problems, build: a => StatementWorkbook.build(ExcelJS, a, new Date()) };
  render();
})();
