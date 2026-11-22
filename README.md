# IDCO Bank Statement to Excel and BRS

**Open the tool:** https://idco-finance.github.io/idco-brs/

Use it in Chrome or Microsoft Edge. Nothing to install.

## What it does

1. Turns bank statement PDFs (ICICI, SBI, Axis and others) into one clean Excel file per bank account.
   Every line is checked against the running balance and the totals printed by the bank.
2. Matches the statement with the ERP Bank Book (made by the IDCO Bank Book Linearizer) and the Eazypay reports.
3. Lets the person doing the BRS decide each open item in Excel (Accept, Reject, Keep open, Match with),
   and applies those decisions again the next time.

4. Prepares the Bank Reconciliation Statement as on any month-end from 1 April, also in IDCO's own
   four-heading format, from the last signed BRS, the statements and the Bank Book.

**Accounts tried so far:** the ICICI Bank Collection account. On any other account the tool shows a
"not yet tried on this account" notice; check every figure before relying on it.

## Your files stay on your computer

The page runs entirely inside your browser. Statements, Bank Books and Eazypay files are read on your own
computer and are **never uploaded** — not to GitHub and not anywhere else. This repository holds only the
tool's code; it contains no bank data.

## How to use

1. Open the link above.
2. Drop the statement PDFs on the page. If a PDF is locked, the tool asks for its password (used once, not saved).
3. To match with the books, also drop the Bank Book workbooks from the Linearizer and the Eazypay files.
4. Download the Excel files. In the matching Excel, open the **To review** sheet and fill the **Decision** column.
5. Save that Excel. Next time, drop it in together with the other files: your decisions are applied again.

The tool lists and matches. It never books, changes or deletes anything in the ERP or at the bank.
The signed BRS remains the responsibility of the person doing it.

## Updates

New versions appear at the same link. Reload the page (Ctrl + F5) to be sure you have the latest one;
the version number is shown at the bottom of the page.

## Also works as a Chrome extension

This folder is also an unpacked Chrome extension (`manifest.json`), for use without internet:
`chrome://extensions` → Developer mode → Load unpacked → choose this folder.

Libraries used: pdf.js (Apache 2.0) and ExcelJS (MIT) — see `lib/LICENSES.txt`.
