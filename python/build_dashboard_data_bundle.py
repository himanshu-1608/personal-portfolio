#!/usr/bin/env python3
"""Build dashboard/data.js from generated CSV reports."""

from __future__ import annotations

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


def read_text_or_empty(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8")


def main() -> int:
    payload = {
        "generatedAtUtc": datetime.now(timezone.utc).isoformat(),
        "stockClosingPricesCsv": read_text_or_empty(STOCK_REPORT_FILE),
        "stockPnlSummaryCsv": read_text_or_empty(PNL_REPORT_FILE),
        "mtfPnlSummaryCsv": read_text_or_empty(MTF_REPORT_FILE),
        "allLedgerCsv": read_text_or_empty(ALL_LEDGER_FILE),
    }
    output = f"window.__DASHBOARD_DATA__ = {json.dumps(payload, separators=(',', ':'))};\n"
    DASHBOARD_DATA_FILE.write_text(output, encoding="utf-8")
    print(f"Wrote dashboard data bundle: {DASHBOARD_DATA_FILE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
