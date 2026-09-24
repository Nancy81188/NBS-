# Saber Accounting MVP

Saber Accounting is a Windows desktop accounting application with a central shared database for three users. This first version includes:

- English, Arabic, and French interface
- Automatic USD, EUR, LBP, and AED detection from currency and amount cells
- Missing currency defaults to USD; conflicting or unsupported currencies are flagged for review
- Currency filter for the dashboard, invoices, trial balance, and exported reports
- Lebanese VAT at 11%
- Purchase and sales invoice import from Excel
- Manual purchase and sales invoice entry with multiple items
- Automatic 11% VAT per item, with editable VAT rate and VAT amount
- Editable total before VAT per item with automatic invoice totals
- Original invoice number when supplied; otherwise original Excel row number
- Completely empty rows skipped; duplicate invoices retained
- Automatic double-entry journal posting
- Invoice register, dashboard, and trial balance
- User authentication, roles, audit log, customer/supplier, inventory, and stock database foundations

## Important status

This is an MVP for controlled testing. Before production use, add HTTPS, automatic encrypted backups, user-management screens, sequential journal controls, complete inventory valuation, period locking, exchange-rate revaluation, invoice editing/reversal workflows, and Lebanese tax-report validation by the firm's accountant.

## Quick start on one computer

Install Python 3.11 or newer. Open Command Prompt in this folder and run:

```bat
python -m pip install -r requirements.txt
python run_server.py --admin-password YourStrongPassword
```

Open a second Command Prompt in the same folder:

```bat
python run_desktop.py
```

Sign in with username `admin`, your chosen server password, and server address `http://127.0.0.1:8765`.

## Three synchronized computers

1. Choose one always-on office computer or Windows server to host the shared database.
2. Give that computer a fixed local IP address.
3. Allow TCP port `8765` only on the trusted office network.
4. Run `run_server.py` only on the server computer.
5. Run the desktop client on each of the three computers.
6. Enter `http://SERVER-IP:8765` on the sign-in screen.

For access outside the office, do not expose port 8765 directly to the internet. Use a professionally configured HTTPS reverse proxy or VPN.

## Excel import rules

The first worksheet is imported. Recognized English, Arabic, and French headings include Invoice Number, Date, Supplier/Customer Name, Total Before VAT, VAT, Total After VAT, Currency, and Type.

- If an invoice-number column contains a value, that value is used.
- Otherwise the original Excel row number is used. Excel row 24 becomes invoice 24.
- Completely empty rows are ignored.
- Duplicate invoices are deliberately retained.
- The original filename and Excel row number are stored for audit tracking.
- Missing VAT is calculated at 11% when the subtotal exists.
- Existing VAT values are preserved, even when they differ from 11%.
- Source totals that do not equal subtotal plus VAT are preserved, marked `review`, and posted against Import Variance so the ledger remains balanced.

## Build the Windows executable

On Windows, after installing the requirements:

```bat
pyinstaller --noconfirm --onefile --windowed --name SaberAccounting run_desktop.py
pyinstaller --noconfirm --onefile --name SaberAccountingServer run_server.py
```

The executables will be created in the `dist` folder. No GitHub account is required.

## Build a one-click installer online

Upload this project to a private GitHub repository. The included workflow runs the tests, creates the standalone client and server, and packages them as `SaberAccountingSetup.exe`. Open the repository's Actions tab, select **Build Saber Accounting Installer**, run the workflow, and download the **SaberAccountingSetup** artifact. End users do not need Python or GitHub.

## Version 0.7.1 fresh start

The default server database is `SaberAccounting/saber_accounting_v0_7.db`. This gives the upgraded application a completely fresh company file with only the default admin account. The previous `saber_accounting.db` is not loaded and remains available as a safety archive. New entries in v0.7.1 persist normally after the application is closed and reopened.

## Version 0.7.2 Excel currency formats

Excel imports detect USD, EUR, LBP, and AED from both cell contents and Excel Accounting/Custom number formats. This supports sheets where currency symbols are displayed beside numeric values without a separate Currency column. Imported dashboards, invoice lists, trial balances, and exports remain separated by currency.

## Version 0.7.3 currency page filter

The Import Excel preview includes an All/USD/EUR/LBP/AED selector and Apply button. Only rows assigned to the selected currency are displayed. Blank cells with leftover number formatting are ignored during detection; genuinely mixed rows are assigned using Total, Before VAT, and VAT evidence and remain flagged for review.

## Version 1.12.0 final release

### Payroll official reports (Payroll > Official Reports)
- **R10** quarterly salary tax withholding, **R5** annual employer declaration, **R6** individual annual statement.
- Any **month, quarter or year**; employees and managers in **separate sections**, plus a grand total.
- NSSF **employee 3%** and employer medical, family and end-of-service contributions.
- Rates and ceilings are **effective-dated (Date From to Date To)**: each month uses the rules in force on its own date. Saving a new Date From automatically ends the previous period the day before; overlapping periods are rejected.
- Separate columns for **retro salary (with its own retro tax)**, transport, schooling, bonus and 13th salary. R6 shows each retro period.
- Official amounts in LBP (other currencies converted at the payroll month's rate). Posted payroll only, with an optional draft preview.
- **Excel and PDF** export.

### Quarterly VAT (Quarterly VAT tab)
- **Q1-Q4** dates set automatically.
- Sales (output) VAT; deductible VAT on purchases, fixed assets, expenses and customs/imports; **non-deductible VAT** shown separately.
- VAT **payable or credit** carried forward to the next quarter (including into Q1 from the previous fiscal-year file).
- Totals **by currency with LBP equivalents**.
- **Manual adjustments** with a mandatory reason. Saving a return locks the quarter; only an administrator can reopen it.
- Marking VAT non-deductible (Uploaded Data > "VAT Deductible / Non-Deductible", or the expense checkbox) moves that VAT into cost, so the ledger always equals the return.
- **Excel and PDF** export.

### Security and polish
- New non-admin users are valid for **1 year** ("Renew 1 Year" in Security > Users). Expired users cannot sign in. Sessions end after 24 hours.
- Per-user **Payroll** and **VAT** access. At least one active administrator is always kept.
- **Legal document alerts** at sign-in and from the header button, for expired documents and documents expiring within 30 days.
- Backups use SQLite's online backup (safe while others work). Restore validates the file and creates a safety backup first.
- Clearer error messages. A page that fails to load no longer stops the others.

### Fixes
- Payroll saved from the desktop app now stores the period correctly (this previously affected payroll numbering and rate lookups).
- Exchange-rate lookups now pick the latest rate on or before the date, chronologically.
- Users created in Settings are saved to the sign-in database.

### To confirm with your accountant
The salary tax method is unchanged: transport and schooling are included in taxable salary, and one-off bonus / 13th salary are annualized ×12 in the month paid. Adjust in Tax & NSSF Settings or ask for a rule change if your practice differs.

### Build the installer (final workflow)
1. Upload this project to the GitHub repository (replace the old files).
2. Open **Actions**, select **Build Saber Accounting Installer**, click **Run workflow**. It also runs automatically on every push to `main`.
3. When it finishes, download the **SaberAccountingSetup** artifact and run `SaberAccountingSetup.exe`.

One installer only: no Python, no manual server. The data service starts automatically inside the app, and company data stays in the user's `SaberAccounting` folder across upgrades. First sign-in on a new computer: `admin` / `admin` (change it in Security > Users).


---

# Version 2.0.0 - complete release

## Modules
| Tab | What it does |
|---|---|
| Sales Invoice | Automatic number, editable lines (Item, description, qty, price, VAT), VAT treatment (taxable / zero-rated / exempt / out of scope), New / Save / Save & Post, reopen and edit |
| Journal Voucher | BRAINS-style voucher: multi-currency lines with LBP and USD amounts and rates, due date, reference, department, project, navigation |
| Import | Excel or PDF invoices as Purchases, Sales, Expenses or Assets. PDFs are read automatically and attached. Adds to existing data (replace is optional) |
| Customers / Suppliers | Account number first; type 4 digits and the full number fills in |
| Payment & Receipt | Customer receipts (RV-) and supplier payments (PV-) on the page, with balance, edit and delete |
| Purchases & Expenses | Purchases with PDF, cost on purchase (customs, freight, insurance, import VAT - manual, Excel or PDF), VAT use; expenses with PDF, Excel import, New / Edit / Delete / Search |
| Inventory | Items, warehouses, stock documents (opening, receipt, issue, adjustments, transfer), weighted average or FIFO, reports, stock variation |
| Payroll | Lebanese rules 2024-2026, R5 / R6 / R10 reports |
| Quarterly VAT | Lebanese periodic declaration with the partial deduction right |
| Trial Balance / Statement | Balance des Comptes options (BRAINS) |
| Profit & Loss | Fiscal-year dates, closing 6 & 7 as a Journal Voucher, automatic opening of the next year |
| Financial Reports | Ledger, balance sheet, cash flow, aging, comparative P&L, budget |
| Security / Backup / Rates | Users with 1-year validity and permissions, departments, projects, backups, exchange rates |

## Inventory (Lebanese periodic method)
- Purchases stay in 601. Stock quantities and costs come from the stock documents.
- Costing: weighted average (default) or FIFO. Stock can never go negative in a warehouse.
- A Sales Invoice line with an Item code issues the stock automatically (Stock Issue linked to the invoice; deleting or cancelling the invoice removes it).
- Reports: Stock Valuation (at any date, by warehouse, at cost and at sales price), Stock Card, Stock Movements, Sales Margin (COGS), Reorder, Slow-moving stock. Excel, PDF and print.
- Year end: the Stock Variation voucher (type 06) cancels account 37 against 6051 and books the closing stock (Dr 37 / Cr 6052). It is posted automatically when the year is closed, and the closing stock becomes the Opening Stock of the next year.

## Year-end order
1. Enter the last documents of the year and check the stock (Inventory > Reports > Stock Valuation at 31-12).
2. Profit & Loss > Preview Closing 6&7.
3. Close the year: stock variation, closing 6 & 7 (result to 121 / 125), and the opening of the next year (balances and stock) are made automatically.
4. "Delete Closing & Reopen Year" undoes everything if a correction is needed.

## Points to confirm with the accountant
- Employer NSSF sickness & maternity rate (8% per PwC; one source says 11%).
- Exact start dates of Jan-Feb 2024 ceilings and the 28M minimum wage.
- VAT declaration box numbers against the official MoF form; rounding of the deduction ratio.
