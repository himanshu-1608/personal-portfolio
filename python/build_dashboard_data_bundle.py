#!/usr/bin/env python3
"""Build dashboard/data.js from generated CSV reports."""

from __future__ import annotations

import csv
import io
import json
from datetime import datetime, timezone
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
REPORTS_DIR = BASE_DIR / "reports"
DASHBOARD_DATA_FILE = BASE_DIR / "dashboard" / "data.js"

STOCK_REPORT_FILE = REPORTS_DIR / "stock_closing_prices.csv"
PNL_REPORT_FILE = REPORTS_DIR / "stock_pnl_summary.csv"
MTF_REPORT_FILE = REPORTS_DIR / "mtf_pnl_summary.csv"
ALL_LEDGER_FILE = BASE_DIR / "input" / "all-ledger.csv"
EXPORTED_LEDGER_FILE = REPORTS_DIR / "exported_all_ledger.csv"

LEDGER_EXPORT_COLUMNS = ["particulars", "posting_date", "voucher_type", "debit", "credit", "net_balance"]


def read_text_or_empty(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8")


def export_stripped_ledger() -> None:
    if not ALL_LEDGER_FILE.exists():
        return
    with ALL_LEDGER_FILE.open(encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = [
            {col: row.get(col, "") for col in LEDGER_EXPORT_COLUMNS}
            for row in reader
        ]
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=LEDGER_EXPORT_COLUMNS, lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    EXPORTED_LEDGER_FILE.write_text(out.getvalue(), encoding="utf-8")
    print(f"Exported stripped ledger: {EXPORTED_LEDGER_FILE} ({len(rows)} rows)")


def main() -> int:
    export_stripped_ledger()
    payload = {
        "generatedAtUtc": datetime.now(timezone.utc).isoformat(),
        "stockClosingPricesCsv": read_text_or_empty(STOCK_REPORT_FILE),
        "stockPnlSummaryCsv": read_text_or_empty(PNL_REPORT_FILE),
        "mtfPnlSummaryCsv": read_text_or_empty(MTF_REPORT_FILE),
        "allLedgerCsv": read_text_or_empty(EXPORTED_LEDGER_FILE),
    }
    output = f"window.__DASHBOARD_DATA__ = {json.dumps(payload, separators=(',', ':'))};\n"
    DASHBOARD_DATA_FILE.write_text(output, encoding="utf-8")
    print(f"Wrote dashboard data bundle: {DASHBOARD_DATA_FILE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
