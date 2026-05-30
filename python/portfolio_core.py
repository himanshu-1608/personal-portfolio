"""Shared constants and small helpers used by both the price-fetching pipeline
(`fetch_yahoo_prices.py`) and the portfolio analytics builders
(`portfolio_reports.py`).

Keeping these here avoids a circular import between the two while giving the
report builders a single, focused place to find paths, column names, and the
generic number/symbol helpers.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Sequence


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input"
REPORTS_DIR = BASE_DIR / "reports"

TRADEBOOK_FILE = INPUT_DIR / "tradebook.csv"
ALL_LEDGER_FILE = INPUT_DIR / "all-ledger.csv"
MTF_LEDGER_FILE = INPUT_DIR / "mtf-ledger.csv"
CONTRACT_NOTES_FILE = INPUT_DIR / "contract_notes.xlsx"
PNL_OUTPUT_FILE = REPORTS_DIR / "stock_pnl_summary.csv"
MTF_PNL_OUTPUT_FILE = REPORTS_DIR / "mtf_pnl_summary.csv"
PORTFOLIO_TIMELINE_FILE = REPORTS_DIR / "portfolio_timeline.csv"

STOCK_COLUMN = "Stock Code/Name"
DATE_COLUMN = "Recommendation Date"
DAY_COUNT = 252

PNL_HEADERS = [
    "Stock Code/Name",
    "Trade Date",
    "Matched Report Symbol",
    "Buy Quantity",
    "Buy Value",
    "Sell Quantity",
    "Sell Value",
    "Net Quantity",
    "Average Buy Price",
    "Average Sell Price",
    "Open Average Cost",
    "Latest Market Price",
    "Latest Market Date",
    "Realized P&L",
    "Charges, Taxes, Others",
    "Net Realized P&L",
    "Unrealized P&L",
    "Total P&L",
    "Return %",
]
MTF_PNL_HEADERS = [
    "Stock Code/Name",
    "Trade Date",
    "Matched Report Symbol",
    "Buy Quantity",
    "Buy Value",
    "Your Funding",
    "Zerodha Funding",
    "Leverage X",
    "Sell Quantity",
    "Sell Value",
    "Net Quantity",
    "Latest Market Price",
    "Latest Market Date",
    "Holding P&L",
    "Holding Return %",
    "Funding Return %",
    "MTF Interest Cost",
    "MTF Pledge Charges",
    "Total Carrying Cost",
    "Net Funding P&L",
    "Net Funding Return %",
]
PORTFOLIO_TIMELINE_HEADERS = [
    "date",
    "r_invested",   # realized lots: full buy value of qty sold by this date
    "r_value",      # realized lots: actual sell proceeds locked in by this date
    "r_zerodha",    # realized lots: zerodha (MTF) funding portion of sold qty
    "u_invested",   # open lots: full buy value of qty still held on this date
    "u_value",      # open lots: market value (qty x forward-filled price) on this date
    "u_zerodha",    # open lots: zerodha (MTF) funding portion of held qty
]


def normalize_symbol(symbol: str) -> str:
    return symbol.strip().upper()


def parse_amount(value: str) -> float:
    cleaned = (value or "").replace(",", "").strip()
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def format_number(value: float) -> str:
    return f"{value:.4f}"


def format_quantity(value: float) -> str:
    return str(int(round(value)))


def load_latest_market_data(rows: Sequence[Dict[str, str]]) -> Dict[str, Dict[str, str]]:
    market_data: Dict[str, Dict[str, str]] = {}
    for row in rows:
        stock_code = (row.get(STOCK_COLUMN) or "").strip()
        base_symbol = normalize_symbol(stock_code.split(".")[0])

        latest_date = ""
        latest_price = ""
        for index in range(1, DAY_COUNT + 1):
            row_date = (row.get(f"Date {index}") or "").strip()
            row_price = (row.get(f"Day {index} Price") or "").strip()
            if row_date and row_price and row_date > latest_date:
                latest_date = row_date
                latest_price = row_price

        if latest_date and latest_price:
            market_data[base_symbol] = {
                "report_symbol": stock_code,
                "latest_market_date": latest_date,
                "latest_market_price": latest_price,
            }
    return market_data
