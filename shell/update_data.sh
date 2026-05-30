#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

cd "${REPO_ROOT}"

echo "Running data refresh from: ${REPO_ROOT}"

if ! python3 -c "import openpyxl" >/dev/null 2>&1; then
  echo "Installing missing dependency: openpyxl"
  python3 -m pip install --user openpyxl
fi

if ! python3 python/fetch_yahoo_prices.py; then
  echo "Full refresh failed (likely market-data network/SSL issue)."
  echo "Rebuilding P&L and MTF reports from local files so charges/net realized stay updated."
  python3 - <<'PY'
import csv
from pathlib import Path
import sys

sys.path.insert(0, str(Path("python").resolve()))
import fetch_yahoo_prices as f  # noqa: E402

if not f.OUTPUT_FILE.exists():
    raise SystemExit(
        "Cannot rebuild local reports because reports/stock_closing_prices.csv is missing. "
        "Run full refresh once with working network."
    )

with f.OUTPUT_FILE.open("r", newline="", encoding="utf-8") as fh:
    report_rows = list(csv.DictReader(fh))

pnl_rows = f.build_pnl_summary(report_rows)
f.write_pnl_output(pnl_rows)
mtf_rows = f.build_mtf_pnl_summary(report_rows)
f.write_mtf_pnl_output(mtf_rows)

print(f"Rebuilt local reports. PNL rows: {len(pnl_rows)}, MTF rows: {len(mtf_rows)}")
PY
fi

python3 python/build_dashboard_data_bundle.py
echo "Done. Generated CSVs are updated in reports/ and dashboard/data.js is refreshed."
