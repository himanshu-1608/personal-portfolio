"""Incrementally merge freshly downloaded Zerodha daily reports into the canonical
`input/` CSVs that feed the dashboard pipeline.

Source of daily downloads: `zerodha-daily-reports/<YYYY-MM-DD>/` produced by
`zerodha_download.py`. For each report we:

  1. Drop opening/closing balance rows (per-download artifacts that would otherwise
     pile up mid-file on every merge).
  2. Append only rows not already present in the existing `input/` file (dedup by a
     stable key), preserving chronological order.

Existing `input/` files are also stripped of any balance rows on rewrite, so legacy
opening/closing rows present from manual downloads are cleaned out on first run.

Run:  python3 python/merge_zerodha_reports.py [--date-dir PATH] [--date YYYY-MM-DD]
"""

from __future__ import annotations

import argparse
import csv
from datetime import date
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import portfolio_core as core

DAILY_REPORTS_DIR = core.BASE_DIR / "zerodha-daily-reports"

# (glob prefix in the dated folder, target input file, dedup-key function)
Row = Dict[str, str]
KeyFn = Callable[[Row], Tuple]


def tradebook_key(row: Row) -> Tuple:
    tid = (row.get("trade_id") or "").strip()
    if tid:
        return ("trade_id", tid)
    # Fallback when trade_id is missing: identify by the trade's economics.
    return (
        "composite",
        (row.get("trade_date") or "").strip(),
        (row.get("symbol") or "").strip(),
        (row.get("trade_type") or "").strip(),
        (row.get("quantity") or "").strip(),
        (row.get("price") or "").strip(),
    )


def ledger_key(row: Row) -> Tuple:
    return (
        (row.get("particulars") or "").strip(),
        (row.get("posting_date") or "").strip(),
        (row.get("cost_center") or "").strip(),
        (row.get("voucher_type") or "").strip(),
        (row.get("debit") or "").strip(),
        (row.get("credit") or "").strip(),
    )


REPORTS = [
    ("tradebook-*.csv", core.TRADEBOOK_FILE, tradebook_key),
    ("all_ledger-*.csv", core.ALL_LEDGER_FILE, ledger_key),
    ("mtf_ledger-*.csv", core.MTF_LEDGER_FILE, ledger_key),
]


def is_balance_row(row: Row) -> bool:
    """Opening/closing balance (and bare summary) rows.

    Ledger reports prefix/suffix the statement with "Opening Balance" / "Closing
    Balance" rows that carry only net_balance and an empty posting_date. Tradebook
    rows have no `particulars`/`posting_date` columns, so this never trips for them.
    """
    particulars = (row.get("particulars") or "").strip().lower()
    if particulars in ("opening balance", "closing balance"):
        return True
    # Ledger entry rows always have a posting_date; rows missing it are summaries.
    if "posting_date" in row and not (row.get("posting_date") or "").strip():
        return True
    return False


def read_rows(path: Path) -> Tuple[List[str], List[Row]]:
    with path.open(newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        fieldnames = reader.fieldnames or []
        rows = [dict(r) for r in reader]
    return fieldnames, rows


def write_rows(path: Path, fieldnames: List[str], rows: List[Row]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(
            fh, fieldnames=fieldnames, lineterminator="\n", extrasaction="ignore"
        )
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in fieldnames})


def merge_one(daily_path: Path, target_path: Path, key_fn: KeyFn) -> int:
    daily_fields, daily_rows = read_rows(daily_path)
    daily_clean = [r for r in daily_rows if not is_balance_row(r)]

    if target_path.exists():
        fieldnames, existing_rows = read_rows(target_path)
    else:
        fieldnames, existing_rows = daily_fields, []
    if not fieldnames:
        fieldnames = daily_fields

    existing_clean = [r for r in existing_rows if not is_balance_row(r)]
    seen = {key_fn(r) for r in existing_clean}

    new_rows = []
    for row in daily_clean:
        k = key_fn(row)
        if k in seen:
            continue
        seen.add(k)
        new_rows.append(row)

    # Existing rows kept in original order; new rows appended chronologically after.
    merged = existing_clean + new_rows
    write_rows(target_path, fieldnames, merged)

    dropped = len(existing_rows) - len(existing_clean)
    note = f" (stripped {dropped} balance row(s) from existing)" if dropped else ""
    print(
        f"{target_path.name}: +{len(new_rows)} new row(s), "
        f"{len(merged)} total{note}"
    )
    return len(new_rows)


def resolve_date_dir(date_dir: Optional[str], day: Optional[str]) -> Path:
    if date_dir:
        return Path(date_dir).expanduser().resolve()
    day = day or date.today().isoformat()
    return DAILY_REPORTS_DIR / day


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--date-dir", help="explicit path to a dated downloads folder")
    parser.add_argument("--date", help="YYYY-MM-DD under zerodha-daily-reports/")
    args = parser.parse_args()

    folder = resolve_date_dir(args.date_dir, args.date)
    if not folder.is_dir():
        raise SystemExit(f"Downloads folder not found: {folder}")

    print(f"Merging from: {folder}")
    total_new = 0
    for prefix, target, key_fn in REPORTS:
        matches = sorted(folder.glob(prefix))
        if not matches:
            print(f"WARNING: no file matching {prefix} in {folder} — skipping")
            continue
        if len(matches) > 1:
            print(f"WARNING: multiple files match {prefix}; using {matches[-1].name}")
        total_new += merge_one(matches[-1], target, key_fn)

    print(f"Merge complete. {total_new} new row(s) added across all reports.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
