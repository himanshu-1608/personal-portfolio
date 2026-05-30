#!/usr/bin/env python3
"""Build dashboard/data.YYYYMMDD.js from generated CSV reports.

Each run writes a date-stamped data file (e.g. data.20260531.js) and
updates the <script> tag in index.html to reference it. Old date-stamped
files from previous days are deleted. This forces GitHub Pages CDN to
serve a new URL on every push, bypassing CDN-level caching.
"""

from __future__ import annotations

import csv
import io
import json
import re
from datetime import date, datetime, timezone
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
REPORTS_DIR = BASE_DIR / "reports"
DASHBOARD_DIR = BASE_DIR / "dashboard"
INDEX_HTML_FILE = DASHBOARD_DIR / "index.html"

STOCK_REPORT_FILE = REPORTS_DIR / "stock_closing_prices.csv"
PNL_REPORT_FILE = REPORTS_DIR / "stock_pnl_summary.csv"
MTF_REPORT_FILE = REPORTS_DIR / "mtf_pnl_summary.csv"
ALL_LEDGER_FILE = BASE_DIR / "input" / "all-ledger.csv"
EXPORTED_LEDGER_FILE = REPORTS_DIR / "exported_all_ledger.csv"
PORTFOLIO_TIMELINE_FILE = REPORTS_DIR / "portfolio_timeline.csv"

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


def cleanup_old_data_files(keep: Path) -> None:
    removed = []
    for old in DASHBOARD_DIR.glob("data.js"):
        old.unlink()
        removed.append(old.name)
    for old in DASHBOARD_DIR.glob("data.2*.js"):
        if old != keep:
            old.unlink()
            removed.append(old.name)
    if removed:
        print(f"Removed old data files: {', '.join(removed)}")


def update_index_html(new_filename: str) -> None:
    content = INDEX_HTML_FILE.read_text(encoding="utf-8")
    updated = re.sub(
        r'<script src="\.\/data[^"]*\.js[^"]*">',
        f'<script src="./{new_filename}">',
        content,
    )
    if updated != content:
        INDEX_HTML_FILE.write_text(updated, encoding="utf-8")
        print(f"Updated index.html: data script → {new_filename}")


def main() -> int:
    export_stripped_ledger()

    now = datetime.now(timezone.utc).astimezone()
    data_filename = f"data.{now.strftime('%Y%m%d-%H%M')}.js"
    data_file = DASHBOARD_DIR / data_filename

    payload = {
        "generatedAtUtc": datetime.now(timezone.utc).isoformat(),
        "stockClosingPricesCsv": read_text_or_empty(STOCK_REPORT_FILE),
        "stockPnlSummaryCsv": read_text_or_empty(PNL_REPORT_FILE),
        "mtfPnlSummaryCsv": read_text_or_empty(MTF_REPORT_FILE),
        "allLedgerCsv": read_text_or_empty(EXPORTED_LEDGER_FILE),
        "portfolioTimelineCsv": read_text_or_empty(PORTFOLIO_TIMELINE_FILE),
    }
    output = f"window.__DASHBOARD_DATA__ = {json.dumps(payload, separators=(',', ':'))};\n"
    data_file.write_text(output, encoding="utf-8")
    print(f"Wrote dashboard data bundle: {data_file}")

    cleanup_old_data_files(keep=data_file)
    update_index_html(data_filename)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
