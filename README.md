# Stack Wealth Analysis

Static dashboard + CSV-based reporting for:

- Recommendation performance
- Portfolio P&L
- MTF P&L

## Nightly SOP

### 1) Update input files

Place/update CSVs in:

- `input/recommendation.csv`
- `input/tradebook.csv`
- `input/all-ledger.csv`
- `input/mtf-ledger.csv`
- `input/contract_notes.xlsx` (required for exact Charges/Taxes/Others)

### 2) Run one command

From repo root:

```bash
./shell/update_data.sh
```

If you are not in repo root:

```bash
cd /Users/hy/Documents/finance/personal-portfolio
./shell/update_data.sh
```

### 3) Commit and push

The script refreshes generated files in:

- `reports/`
- `dashboard/data.js`

Commit these updated files and push to GitHub. GitHub Pages will reflect latest data.

## Manual equivalent (reference)

`./shell/update_data.sh` internally does:

```bash
cd /Users/hy/Documents/finance/personal-portfolio
python3 python/fetch_yahoo_prices.py
python3 python/build_dashboard_data_bundle.py
```

If Yahoo market fetch fails (network/SSL), the script now automatically falls back to:
- rebuild `reports/stock_pnl_summary.csv` and `reports/mtf_pnl_summary.csv` from local files
- keep contract-note charges integration (`input/contract_notes.xlsx`)
- still regenerate `dashboard/data.js`

## View Website Locally

From repo root:

```bash
cd /Users/hy/Documents/finance/personal-portfolio
open ./index.html
```

This opens directly from `file://` and works because `dashboard/data.js` contains the bundled report CSVs.

Optional (still supported):

```text
python3 -m http.server 8000
http://127.0.0.1:8000/
```

This auto-redirects to the dashboard page.
