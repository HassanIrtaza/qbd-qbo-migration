# QBD → QuickBooks Online Migration

A customer-facing web tool for migrating financial data from **QuickBooks Desktop** to **QuickBooks Online**.

```
QuickBooks Desktop  ──►  Migration Agent  ──►  QuickBooks Online
  (Excel export)          (OAuth 2.0)            (REST API v3)
```

## What It Migrates

- Chart of Accounts
- Customers (with sub-customers / jobs)
- Vendors (incl. 1099 flag)
- Items (Service, Inventory, Non-Inventory)
- Payment Terms
- Open AR Invoices
- Open AP Bills
- Employees
- Trial Balance (opening journal)

## Screens

The web UI has three panels:

1. **Source** — upload QBD Excel exports or generate sample data
2. **Target** — Connect to QuickBooks (OAuth 2.0, Sandbox or Production)
3. **Run** — dry-run validation, then live migration with streaming logs

## Quick Start

### Prerequisites

- Python 3.9+
- An **Intuit Developer** app — grab your `Client ID` / `Client Secret` from https://developer.intuit.com
- A QBO **Sandbox company** for testing (free, included with Developer account)

### Install

```bash
git clone https://github.com/HassanIrtaza/qbd-qbo-migration.git
cd qbd-qbo-migration
pip install -r requirements.txt
cp .env.example .env
# edit .env with your Intuit app Client ID + Secret
```

### Run

```bash
python app.py
# open http://localhost:5050
```

Click **Connect to QuickBooks** → authorize your sandbox company → upload QBD Excel files → **Dry Run** → **Start Migration**.

## QBD Export Format

The app expects Excel files in `qbd_exports/`:

| File | Source in QBD |
|---|---|
| `QBD_ChartOfAccounts.xlsx` | Lists → Chart of Accounts → Excel |
| `QBD_Customers.xlsx` | Customer Center → Excel |
| `QBD_Vendors.xlsx` | Vendor Center → Excel |
| `QBD_Items.xlsx` | Lists → Item List → Excel |
| `QBD_Employees.xlsx` | Employee Center → Excel |
| `QBD_OpenInvoices.xlsx` | Reports → Customers → Open Invoices |
| `QBD_OpenBills.xlsx` | Reports → Vendors → Unpaid Bills |
| `QBD_TrialBalance.xlsx` | Reports → Accountant → Trial Balance |

Click **Generate Sample Data** in the UI if you just want to demo with dummy data.

## Architecture

```
Phase 1 — Extract from QBD Excel
Phase 2 — Transform
  ├── Map QB account types → QBO AccountType/AccountSubType
  ├── Convert QB jobs → QBO sub-customers (parent ref)
  ├── Map QB terms → QBO Term objects
  └── Validate before writing
Phase 3 — Load via QBO REST API v3
  ├── Dependency-ordered writes (Accounts → Terms → Customers → Jobs → Vendors → Items → Invoices → Bills)
  ├── OAuth 2.0 token refresh on 401
  └── 429 retry with exponential backoff
Phase 4 — Validate
  └── Query QBO and compare counts to source
```

## QBD → QBO Mapping Highlights

| QBD Concept | QBO Concept |
|---|---|
| Customer:Job | Customer with `Job=true`, `ParentRef` |
| Terms (e.g. Net 30) | Term object (global list) |
| 1099 Vendor | Vendor with `Vendor1099=true` |
| Service/Inventory/Non-Inv Item | Item with `Type` enum |
| Account (Bank/AR/AP/Income/Expense) | Account with `AccountType` + `AccountSubType` |

## Configuration

`.env`:

```
QBO_CLIENT_ID=your_client_id
QBO_CLIENT_SECRET=your_client_secret
QBO_REDIRECT_URI=http://localhost:5050/callback
QBO_ENVIRONMENT=sandbox        # or production
QBO_SCOPES=com.intuit.quickbooks.accounting
```

## Demo Mode

If no `.env` is present, the app runs in **Demo Mode** — OAuth is stubbed and migrations write to a local SQLite mirror so you can run the full UI flow for customer presentations without real QBO credentials.

## License

MIT
