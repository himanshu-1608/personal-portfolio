#!/usr/bin/env python3
"""
Fetch closing prices from Yahoo Finance for stock recommendations, from the
recommendation date through the latest available trading day (no cap).

Input:
  /input/recommendation.csv

Expected input columns:
  - Stock Code/Name
  - Recommendation Date

Output:
  /reports/stock_closing_prices.csv

The script refreshes all recommendations on each run so the latest trading days
are reflected automatically in the output report.
"""

from __future__ import annotations

import csv
import io
import json
import random
import re
import sys
import time
import http.cookiejar
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, time as dt_time, timedelta, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import yfinance as yf
except ImportError:  # pragma: no cover
    yf = None  # type: ignore

try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    load_workbook = None  # type: ignore

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover
    ZoneInfo = None  # type: ignore


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input"
REPORTS_DIR = BASE_DIR / "reports"
INPUT_FILE = INPUT_DIR / "recommendation.csv"
TRADEBOOK_FILE = INPUT_DIR / "tradebook.csv"
ALL_LEDGER_FILE = INPUT_DIR / "all-ledger.csv"
MTF_LEDGER_FILE = INPUT_DIR / "mtf-ledger.csv"
CONTRACT_NOTES_FILE = INPUT_DIR / "contract_notes.xlsx"
OUTPUT_FILE = REPORTS_DIR / "stock_closing_prices.csv"
PNL_OUTPUT_FILE = REPORTS_DIR / "stock_pnl_summary.csv"
MTF_PNL_OUTPUT_FILE = REPORTS_DIR / "mtf_pnl_summary.csv"
PORTFOLIO_TIMELINE_FILE = REPORTS_DIR / "portfolio_timeline.csv"
PORTFOLIO_TIMELINE_HEADERS = [
    "date",
    "r_invested",   # realized lots: full buy value of qty sold by this date
    "r_value",      # realized lots: actual sell proceeds locked in by this date
    "r_zerodha",    # realized lots: zerodha (MTF) funding portion of sold qty
    "u_invested",   # open lots: full buy value of qty still held on this date
    "u_value",      # open lots: market value (qty x forward-filled price) on this date
    "u_zerodha",    # open lots: zerodha (MTF) funding portion of held qty
]

STOCK_COLUMN = "Stock Code/Name"
DATE_COLUMN = "Recommendation Date"
TARGET_1_COLUMN = "Target 1"
TARGET_2_COLUMN = "Target 2"
TARGET_3_COLUMN = "Target 3"
TARGET_4_COLUMN = "Target 4"
BUY_PRICE_RECOMMENDATION_COLUMN = "Buy Price Recommendation"
DAY_COUNT = 252
REQUEST_TIMEOUT_SECONDS = 30
MAX_RETRIES = 5
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
REQUEST_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "application/json,text/plain,*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Origin": "https://finance.yahoo.com",
    "Referer": "https://finance.yahoo.com/",
}
NSE_REQUEST_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "application/json,text/plain,*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Origin": "https://www.nseindia.com",
    "Referer": "https://www.nseindia.com/",
}

OUTPUT_HEADERS = [
    STOCK_COLUMN,
    DATE_COLUMN,
    TARGET_1_COLUMN,
    TARGET_2_COLUMN,
    TARGET_3_COLUMN,
    TARGET_4_COLUMN,
    BUY_PRICE_RECOMMENDATION_COLUMN,
    "Highest Price",
    "Hit Target 1",
    "Hit Target 2",
    "Hit Target 3",
    "Hit Target 4",
    "Target 1 Return %",
    "Target 2 Return %",
    "Target 3 Return %",
    "Target 4 Return %",
    *[
        column_name
        for index in range(1, DAY_COUNT + 1)
        for column_name in (f"Date {index}", f"Day {index} Price", f"Day {index} Return %")
    ],
]
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

COOKIE_JAR = http.cookiejar.CookieJar()
HTTP_OPENER = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(COOKIE_JAR))
NSE_COOKIE_JAR = http.cookiejar.CookieJar()
NSE_HTTP_OPENER = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(NSE_COOKIE_JAR))
NSE_BHAVCOPY_DAY_CACHE: Dict[str, Dict[str, Tuple[float, float]]] = {}


@dataclass(frozen=True)
class Recommendation:
    stock_code: str
    recommendation_date: str
    target_1: str = ""
    target_2: str = ""
    buy_price_recommendation: str = ""

    @property
    def key(self) -> Tuple[str, str]:
        return self.stock_code, self.recommendation_date


def ensure_directories() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)


def parse_recommendation_date(raw_value: str) -> date:
    value = raw_value.strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(
        f"Unsupported Recommendation Date '{raw_value}'. "
        "Use one of: YYYY-MM-DD, DD-MM-YYYY, MM/DD/YYYY, DD/MM/YYYY."
    )


def read_recommendations() -> List[Recommendation]:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE}. "
            f"Create it with columns '{STOCK_COLUMN}' and '{DATE_COLUMN}'."
        )

    with INPUT_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        if reader.fieldnames is None:
            raise ValueError(f"Input file is empty: {INPUT_FILE}")

        missing_columns = {
            column for column in (STOCK_COLUMN, DATE_COLUMN) if column not in reader.fieldnames
        }
        if missing_columns:
            raise ValueError(
                "Input file is missing required columns: "
                + ", ".join(sorted(missing_columns))
            )

        recommendations: List[Recommendation] = []
        seen_keys = set()
        for row_number, row in enumerate(reader, start=2):
            stock_code = (row.get(STOCK_COLUMN) or "").strip()
            recommendation_date_raw = (row.get(DATE_COLUMN) or "").strip()

            if not stock_code or not recommendation_date_raw:
                raise ValueError(
                    f"Row {row_number} has blank '{STOCK_COLUMN}' or '{DATE_COLUMN}'."
                )

            normalized_date = parse_recommendation_date(recommendation_date_raw).isoformat()
            recommendation = Recommendation(
                stock_code=stock_code,
                recommendation_date=normalized_date,
                target_1=(row.get(TARGET_1_COLUMN) or "").strip(),
                target_2=(row.get(TARGET_2_COLUMN) or "").strip(),
                buy_price_recommendation=(row.get(BUY_PRICE_RECOMMENDATION_COLUMN) or "").strip(),
            )

            if recommendation.key not in seen_keys:
                recommendations.append(recommendation)
                seen_keys.add(recommendation.key)

    return recommendations


def read_existing_output() -> Tuple[List[Dict[str, str]], set]:
    if not OUTPUT_FILE.exists():
        return [], set()

    with OUTPUT_FILE.open("r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        if reader.fieldnames is None:
            return [], set()

        rows = list(reader)
        existing_keys = {
            (
                (row.get(STOCK_COLUMN) or "").strip(),
                (row.get(DATE_COLUMN) or "").strip(),
            )
            for row in rows
        }
    return rows, existing_keys


def has_all_four_targets_hit(row: Dict[str, str]) -> bool:
    return all((row.get(f"Hit Target {index}") or "").strip().upper() == "TRUE" for index in range(1, 5))


def get_missing_recommendations(
    recommendations: Sequence[Recommendation], existing_keys: set
) -> List[Recommendation]:
    return [recommendation for recommendation in recommendations if recommendation.key not in existing_keys]


def build_chart_url(
    stock_code: str, start_date: date, end_date: date, host: str = "query1.finance.yahoo.com"
) -> str:
    period1 = int(datetime.combine(start_date, dt_time.min, tzinfo=timezone.utc).timestamp())
    period2 = int(datetime.combine(end_date, dt_time.max, tzinfo=timezone.utc).timestamp())
    query = urllib.parse.urlencode(
        {
            "period1": period1,
            "period2": period2,
            "interval": "1d",
            "includePrePost": "false",
            "events": "div,splits",
        }
    )
    encoded_stock = urllib.parse.quote(stock_code, safe=".^-")
    return f"https://{host}/v8/finance/chart/{encoded_stock}?{query}"


def fetch_json(url: str) -> dict:
    request = urllib.request.Request(url, headers=REQUEST_HEADERS)
    with HTTP_OPENER.open(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
        payload = response.read().decode("utf-8")
    return json.loads(payload)


def nse_symbol_from_stock_code(stock_code: str) -> str:
    return normalize_symbol(stock_code.split(".")[0])


def warm_up_nse_session(symbol: str) -> None:
    requests = [
        urllib.request.Request("https://www.nseindia.com/", headers=NSE_REQUEST_HEADERS),
        urllib.request.Request(
            f"https://www.nseindia.com/get-quotes/equity?symbol={urllib.parse.quote(symbol)}",
            headers=NSE_REQUEST_HEADERS,
        ),
    ]
    for request in requests:
        with NSE_HTTP_OPENER.open(request, timeout=REQUEST_TIMEOUT_SECONDS):
            continue


def build_nse_bhavcopy_url(trading_date: date) -> str:
    date_token = trading_date.strftime("%Y%m%d")
    return (
        "https://nsearchives.nseindia.com/content/cm/"
        f"BhavCopy_NSE_CM_0_0_0_{date_token}_F_0000.csv.zip"
    )


def fetch_nse_bhavcopy_day(trading_date: date) -> Dict[str, Tuple[float, float]]:
    cache_key = trading_date.isoformat()
    if cache_key in NSE_BHAVCOPY_DAY_CACHE:
        return NSE_BHAVCOPY_DAY_CACHE[cache_key]

    url = build_nse_bhavcopy_url(trading_date)
    request = urllib.request.Request(url, headers=NSE_REQUEST_HEADERS)
    try:
        with NSE_HTTP_OPENER.open(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            archive_blob = response.read()
    except urllib.error.HTTPError as exc:
        if exc.code in {403, 401}:
            try:
                warm_up_nse_session("RELIANCE")
                with NSE_HTTP_OPENER.open(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
                    archive_blob = response.read()
            except Exception:
                NSE_BHAVCOPY_DAY_CACHE[cache_key] = {}
                return {}
        elif exc.code == 404:
            NSE_BHAVCOPY_DAY_CACHE[cache_key] = {}
            return {}
        else:
            NSE_BHAVCOPY_DAY_CACHE[cache_key] = {}
            return {}
    except Exception:
        NSE_BHAVCOPY_DAY_CACHE[cache_key] = {}
        return {}

    day_data: Dict[str, Tuple[float, float]] = {}
    try:
        with zipfile.ZipFile(io.BytesIO(archive_blob)) as archive:
            names = archive.namelist()
            if not names:
                NSE_BHAVCOPY_DAY_CACHE[cache_key] = {}
                return {}
            csv_blob = archive.read(names[0]).decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(csv_blob))
        for row in reader:
            symbol = normalize_symbol((row.get("TckrSymb") or "").strip())
            if not symbol:
                continue
            try:
                close_price = float((row.get("ClsPric") or "").strip())
                high_price = float((row.get("HghPric") or "").strip())
            except ValueError:
                continue
            day_data[symbol] = (close_price, high_price)
    except Exception:
        NSE_BHAVCOPY_DAY_CACHE[cache_key] = {}
        return {}

    NSE_BHAVCOPY_DAY_CACHE[cache_key] = day_data
    return day_data


def iter_dates(start_date: date, end_date: date) -> Iterable[date]:
    current_date = start_date
    while current_date <= end_date:
        yield current_date
        current_date += timedelta(days=1)


def merge_nse_fallback_prices(
    stock_code: str,
    start_date: date,
    end_date: date,
    prices: Sequence[Tuple[str, str]],
) -> Tuple[List[Tuple[str, str]], float]:
    if not stock_code.upper().endswith(".NS"):
        return list(prices), 0.0

    nse_symbol = nse_symbol_from_stock_code(stock_code)
    if not nse_symbol:
        return list(prices), 0.0

    last_market_date = min(end_date, date.today())
    fallback_start = max(start_date, last_market_date - timedelta(days=14))
    if fallback_start > last_market_date:
        return list(prices), 0.0

    merged_by_date: Dict[str, float] = {}
    for trading_date, price in prices:
        if trading_date and price:
            try:
                merged_by_date[trading_date] = float(price)
            except ValueError:
                continue

    fallback_highest_price = 0.0
    for trading_date in iter_dates(fallback_start, last_market_date):
        day_data = fetch_nse_bhavcopy_day(trading_date)
        if not day_data:
            continue
        nse_prices = day_data.get(nse_symbol)
        if not nse_prices:
            continue
        close_price, high_price = nse_prices
        iso_date = trading_date.isoformat()
        if iso_date not in merged_by_date:
            merged_by_date[iso_date] = close_price
        fallback_highest_price = max(fallback_highest_price, high_price)

    merged_prices = [(day, f"{merged_by_date[day]:.4f}") for day in sorted(merged_by_date.keys())]
    merged_prices = merged_prices[:DAY_COUNT]
    while len(merged_prices) < DAY_COUNT:
        merged_prices.append(("", ""))

    return merged_prices, fallback_highest_price


def warm_up_yahoo_session() -> None:
    request = urllib.request.Request("https://finance.yahoo.com/", headers=REQUEST_HEADERS)
    with HTTP_OPENER.open(request, timeout=REQUEST_TIMEOUT_SECONDS):
        return


def fetch_chart_payload(stock_code: str, start_date: date, end_date: date) -> dict:
    last_error: Exception | None = None
    urls = [
        build_chart_url(stock_code, start_date, end_date, host="query1.finance.yahoo.com"),
        build_chart_url(stock_code, start_date, end_date, host="query2.finance.yahoo.com"),
    ]

    for attempt in range(1, MAX_RETRIES + 1):
        if attempt == 1:
            try:
                warm_up_yahoo_session()
            except Exception:
                pass

        for url in urls:
            try:
                return fetch_json(url)
            except urllib.error.HTTPError as exc:
                last_error = exc
                if exc.code != 429:
                    raise
            except urllib.error.URLError as exc:
                last_error = exc

        sleep_seconds = min(12, (2 ** (attempt - 1)) + random.uniform(0.2, 0.8))
        time.sleep(sleep_seconds)

    if isinstance(last_error, urllib.error.HTTPError) and last_error.code == 429:
        raise RuntimeError(
            "Yahoo Finance is rate-limiting requests right now (HTTP 429). "
            "Please wait a few minutes and try again. If this keeps happening, "
            "install 'yfinance' to use its more resilient Yahoo session handling."
        ) from last_error

    if last_error is not None:
        raise last_error

    raise RuntimeError("Unable to fetch data from Yahoo Finance.")


def extract_price_points(payload: dict, filter_from_date: date) -> List[Tuple[str, float, float]]:
    chart = payload.get("chart", {})
    error = chart.get("error")
    if error:
        description = error.get("description") or "Unknown Yahoo Finance API error."
        raise ValueError(description)

    results = chart.get("result") or []
    if not results:
        raise ValueError("No data returned from Yahoo Finance.")

    result = results[0]
    timestamps = result.get("timestamp") or []
    quote_list = (((result.get("indicators") or {}).get("quote")) or [])
    closes = (quote_list[0] or {}).get("close") if quote_list else None
    highs = (quote_list[0] or {}).get("high") if quote_list else None

    if not timestamps or not closes or not highs:
        return []

    exchange_timezone_name = ((result.get("meta") or {}).get("exchangeTimezoneName")) or "UTC"
    exchange_timezone = (
        ZoneInfo(exchange_timezone_name) if ZoneInfo is not None else timezone.utc
    )

    prices: List[Tuple[str, float, float]] = []
    for timestamp_value, close_value, high_value in zip(timestamps, closes, highs):
        if close_value is None or high_value is None:
            continue

        trading_date = datetime.fromtimestamp(timestamp_value, tz=exchange_timezone).date()
        if trading_date < filter_from_date:
            continue

        prices.append(
            (
                trading_date.isoformat(),
                round(float(close_value), 4),
                round(float(high_value), 4),
            )
        )
        if len(prices) == DAY_COUNT:
            break

    return prices


def fetch_30_day_prices_with_yfinance(
    stock_code: str, recommendation_date: str
) -> Tuple[List[Tuple[str, str]], str]:
    if yf is None:
        raise RuntimeError("yfinance is not installed.")

    start_date = parse_recommendation_date(recommendation_date)
    end_date = date.today() + timedelta(days=1)

    ticker = yf.Ticker(stock_code)
    history = ticker.history(
        start=start_date.isoformat(),
        end=(end_date + timedelta(days=1)).isoformat(),
        interval="1d",
        auto_adjust=False,
    )

    if history.empty:
        raise ValueError(f"No Yahoo Finance data returned for symbol '{stock_code}'.")

    close_series = history.get("Close")
    high_series = history.get("High")
    if close_series is None:
        raise ValueError(f"Yahoo Finance data for '{stock_code}' did not include Close prices.")
    if high_series is None:
        raise ValueError(f"Yahoo Finance data for '{stock_code}' did not include High prices.")

    # Some yfinance versions can still surface a single-column dataframe here.
    if hasattr(close_series, "ndim") and close_series.ndim > 1:
        if getattr(close_series, "shape", (0, 0))[1] != 1:
            raise ValueError(
                f"Yahoo Finance returned multiple Close columns for '{stock_code}', "
                "and the script could not determine which one to use."
            )
        close_series = close_series.iloc[:, 0]
    if hasattr(high_series, "ndim") and high_series.ndim > 1:
        if getattr(high_series, "shape", (0, 0))[1] != 1:
            raise ValueError(
                f"Yahoo Finance returned multiple High columns for '{stock_code}', "
                "and the script could not determine which one to use."
            )
        high_series = high_series.iloc[:, 0]

    prices: List[Tuple[str, str]] = []
    highest_price = 0.0
    for index_value, close_value in close_series.items():
        high_value = high_series.get(index_value)
        if close_value is None or high_value is None:
            continue
        if hasattr(close_value, "item") and getattr(close_value, "size", 1) == 1:
            close_value = close_value.item()
        if hasattr(high_value, "item") and getattr(high_value, "size", 1) == 1:
            high_value = high_value.item()
        if close_value != close_value:
            continue
        if high_value != high_value:
            continue

        trading_date = index_value.date() if hasattr(index_value, "date") else index_value
        if trading_date < start_date:
            continue

        highest_price = max(highest_price, float(high_value))
        prices.append((trading_date.isoformat(), f"{float(close_value):.4f}"))
        if len(prices) == DAY_COUNT:
            break

    while len(prices) < DAY_COUNT:
        prices.append(("", ""))

    return prices, (f"{highest_price:.4f}" if highest_price else "")


def fetch_30_day_prices(stock_code: str, recommendation_date: str) -> Tuple[List[Tuple[str, str]], str]:
    start_date = parse_recommendation_date(recommendation_date)
    end_date = date.today() + timedelta(days=1)

    if yf is not None:
        try:
            prices, highest_price = fetch_30_day_prices_with_yfinance(stock_code, recommendation_date)
            merged_prices, fallback_high = merge_nse_fallback_prices(
                stock_code=stock_code,
                start_date=start_date,
                end_date=end_date,
                prices=prices,
            )
            highest_value = max(parse_optional_float(highest_price) or 0.0, fallback_high)
            return merged_prices, (f"{highest_value:.4f}" if highest_value else "")
        except Exception as exc:
            print(
                f"Warning: yfinance fetch failed for {stock_code}. "
                f"Falling back to direct Yahoo API. Reason: {exc}",
                file=sys.stderr,
            )

    # Over-fetch a wider date range so we can collect 30 trading closes after
    # the recommendation date even across weekends and market holidays.
    payload = fetch_chart_payload(stock_code, start_date, end_date)
    prices = extract_price_points(payload, start_date)

    highest_price = max((high for _, _, high in prices), default=0.0)
    padded_prices = [(trading_date, f"{price:.4f}") for trading_date, price, _high in prices]
    while len(padded_prices) < DAY_COUNT:
        padded_prices.append(("", ""))

    merged_prices, fallback_high = merge_nse_fallback_prices(
        stock_code=stock_code,
        start_date=start_date,
        end_date=end_date,
        prices=padded_prices,
    )
    highest_value = max(highest_price, fallback_high)
    return merged_prices, (f"{highest_value:.4f}" if highest_value else "")


INCREMENTAL_LOOKBACK_DAYS = 5


def parse_optional_float(value: str) -> Optional[float]:
    stripped = value.strip()
    if not stripped:
        return None
    return float(stripped)


def get_next_empty_slot(row: Dict[str, str]) -> int:
    for index in range(1, DAY_COUNT + 1):
        if not (row.get(f"Date {index}") or "").strip():
            return index
    return DAY_COUNT + 1


def fetch_prices_since(stock_code: str, since_date: date) -> Tuple[List[Tuple[str, str]], str]:
    """Fetch closing prices from since_date to today. Returns only actual trading days (no padding)."""
    end_date = date.today() + timedelta(days=1)

    if yf is not None:
        try:
            ticker = yf.Ticker(stock_code)
            history = ticker.history(
                start=since_date.isoformat(),
                end=end_date.isoformat(),
                interval="1d",
                auto_adjust=False,
            )
            if not history.empty:
                close_series = history.get("Close")
                high_series = history.get("High")
                if close_series is not None and high_series is not None:
                    if hasattr(close_series, "ndim") and close_series.ndim > 1:
                        close_series = close_series.iloc[:, 0]
                    if hasattr(high_series, "ndim") and high_series.ndim > 1:
                        high_series = high_series.iloc[:, 0]

                    prices: List[Tuple[str, str]] = []
                    highest_price = 0.0
                    for index_value, close_value in close_series.items():
                        high_value = high_series.get(index_value)
                        if close_value is None or high_value is None:
                            continue
                        if hasattr(close_value, "item") and getattr(close_value, "size", 1) == 1:
                            close_value = close_value.item()
                        if hasattr(high_value, "item") and getattr(high_value, "size", 1) == 1:
                            high_value = high_value.item()
                        if close_value != close_value or high_value != high_value:
                            continue
                        trading_date = index_value.date() if hasattr(index_value, "date") else index_value
                        if trading_date < since_date:
                            continue
                        highest_price = max(highest_price, float(high_value))
                        prices.append((trading_date.isoformat(), f"{float(close_value):.4f}"))

                    merged_padded, fallback_high = merge_nse_fallback_prices(
                        stock_code=stock_code, start_date=since_date, end_date=end_date, prices=prices
                    )
                    merged = [(d, p) for d, p in merged_padded if d]
                    highest_value = max(highest_price, fallback_high)
                    return merged, (f"{highest_value:.4f}" if highest_value else "")
        except Exception as exc:
            print(
                f"Warning: yfinance incremental fetch failed for {stock_code}. "
                f"Falling back to Yahoo API. Reason: {exc}",
                file=sys.stderr,
            )

    payload = fetch_chart_payload(stock_code, since_date, end_date)
    raw_prices = extract_price_points(payload, since_date)
    highest_price = max((high for _, _, high in raw_prices), default=0.0)
    padded_prices = [(d, f"{p:.4f}") for d, p, _ in raw_prices]
    merged_padded, fallback_high = merge_nse_fallback_prices(
        stock_code=stock_code, start_date=since_date, end_date=end_date, prices=padded_prices
    )
    merged = [(d, p) for d, p in merged_padded if d]
    highest_value = max(highest_price, fallback_high)
    return merged, (f"{highest_value:.4f}" if highest_value else "")


def distribute_targets(
    target_1: Optional[float], target_2: Optional[float]
) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float]]:
    if target_1 is None and target_2 is None:
        return None, None, None, None
    if target_1 is None:
        return target_2, target_2, target_2, target_2
    if target_2 is None:
        return target_1, target_1, target_1, target_1

    step = (target_2 - target_1) / 3
    target_2_split = float(round(target_1 + step))
    target_3_split = float(round(target_1 + (2 * step)))
    return target_1, target_2_split, target_3_split, target_2


def _apply_hit_targets(row: Dict[str, str]) -> None:
    highest_price_val = parse_optional_float(row.get("Highest Price") or "")
    for i, target_col in enumerate(["Target 1", "Target 2", "Target 3", "Target 4"], 1):
        target_val = parse_optional_float(row.get(target_col) or "")
        if target_val is None:
            row[f"Hit Target {i}"] = ""
        else:
            row[f"Hit Target {i}"] = (
                "TRUE" if highest_price_val is not None and highest_price_val >= target_val else "FALSE"
            )


def build_output_row(
    recommendation: Recommendation, existing_row: Optional[Dict[str, str]] = None
) -> Dict[str, str]:
    # ── Incremental path: existing data present, only fetch last N days ──────
    if existing_row is not None:
        since_date = date.today() - timedelta(days=INCREMENTAL_LOOKBACK_DAYS)
        new_prices, new_highest_str = fetch_prices_since(recommendation.stock_code, since_date)

        rec_date = parse_recommendation_date(recommendation.recommendation_date)
        stored_dates = {
            (existing_row.get(f"Date {i}") or "").strip()
            for i in range(1, DAY_COUNT + 1)
        }
        fresh_prices = [
            (d, p) for d, p in new_prices
            if d and d not in stored_dates and date.fromisoformat(d) >= rec_date
        ]

        if not fresh_prices:
            return existing_row

        merged_row = dict(existing_row)

        old_high = parse_optional_float(existing_row.get("Highest Price") or "") or 0.0
        new_high = parse_optional_float(new_highest_str) or 0.0
        merged_high = max(old_high, new_high)
        if merged_high:
            merged_row["Highest Price"] = format_number(merged_high)

        base_price = (
            parse_optional_float(recommendation.buy_price_recommendation)
            if recommendation.buy_price_recommendation
            else None
        )
        if base_price is None:
            for i in range(1, DAY_COUNT + 1):
                day_price = (existing_row.get(f"Day {i} Price") or "").strip()
                if day_price:
                    base_price = parse_optional_float(day_price)
                    break

        next_slot = get_next_empty_slot(merged_row)
        for trading_date, price in fresh_prices:
            if next_slot > DAY_COUNT:
                break
            merged_row[f"Date {next_slot}"] = trading_date
            merged_row[f"Day {next_slot} Price"] = price
            if price and base_price not in (None, 0):
                return_pct = ((float(price) / base_price) - 1) * 100
                merged_row[f"Day {next_slot} Return %"] = f"{return_pct:.2f}%"
            else:
                merged_row[f"Day {next_slot} Return %"] = ""
            next_slot += 1

        _apply_hit_targets(merged_row)
        return merged_row

    # ── Full fetch path (first run / no existing data) ────────────────────────
    prices, highest_price = fetch_30_day_prices(
        recommendation.stock_code, recommendation.recommendation_date
    )
    input_target_1 = parse_optional_float(recommendation.target_1) if recommendation.target_1 else None
    input_target_2 = parse_optional_float(recommendation.target_2) if recommendation.target_2 else None
    target_1, target_2, target_3, target_4 = distribute_targets(input_target_1, input_target_2)

    buy_price_recommendation = (
        parse_optional_float(recommendation.buy_price_recommendation)
        if recommendation.buy_price_recommendation
        else None
    )
    highest_price_value = parse_optional_float(highest_price) if highest_price else None

    row: Dict[str, str] = {
        STOCK_COLUMN: recommendation.stock_code,
        DATE_COLUMN: recommendation.recommendation_date,
        TARGET_1_COLUMN: format_number(target_1) if target_1 is not None else "",
        TARGET_2_COLUMN: format_number(target_2) if target_2 is not None else "",
        TARGET_3_COLUMN: format_number(target_3) if target_3 is not None else "",
        TARGET_4_COLUMN: format_number(target_4) if target_4 is not None else "",
        BUY_PRICE_RECOMMENDATION_COLUMN: recommendation.buy_price_recommendation,
        "Highest Price": highest_price,
        "Hit Target 1": (
            ("TRUE" if highest_price_value is not None and highest_price_value >= target_1 else "FALSE")
            if target_1 is not None else ""
        ),
        "Hit Target 2": (
            ("TRUE" if highest_price_value is not None and highest_price_value >= target_2 else "FALSE")
            if target_2 is not None else ""
        ),
        "Hit Target 3": (
            ("TRUE" if highest_price_value is not None and highest_price_value >= target_3 else "FALSE")
            if target_3 is not None else ""
        ),
        "Hit Target 4": (
            ("TRUE" if highest_price_value is not None and highest_price_value >= target_4 else "FALSE")
            if target_4 is not None else ""
        ),
        "Target 1 Return %": (
            f"{(((target_1 / buy_price_recommendation) - 1) * 100):.2f}%"
            if target_1 is not None and buy_price_recommendation not in (None, 0) else ""
        ),
        "Target 2 Return %": (
            f"{(((target_2 / buy_price_recommendation) - 1) * 100):.2f}%"
            if target_2 is not None and buy_price_recommendation not in (None, 0) else ""
        ),
        "Target 3 Return %": (
            f"{(((target_3 / buy_price_recommendation) - 1) * 100):.2f}%"
            if target_3 is not None and buy_price_recommendation not in (None, 0) else ""
        ),
        "Target 4 Return %": (
            f"{(((target_4 / buy_price_recommendation) - 1) * 100):.2f}%"
            if target_4 is not None and buy_price_recommendation not in (None, 0) else ""
        ),
    }

    # Return % measured vs buy recommendation price to avoid Day 1 being 0.00%.
    base_price = buy_price_recommendation if buy_price_recommendation not in (None, 0) else None
    if base_price is None:
        for _trading_date, price in prices:
            if price:
                base_price = float(price)
                break

    for index, (trading_date, price) in enumerate(prices, start=1):
        row[f"Date {index}"] = trading_date
        row[f"Day {index} Price"] = price
        if price and base_price not in (None, 0):
            return_pct = ((float(price) / base_price) - 1) * 100
            row[f"Day {index} Return %"] = f"{return_pct:.2f}%"
        else:
            row[f"Day {index} Return %"] = ""
    return row


def write_output(rows: Iterable[Dict[str, str]]) -> None:
    sorted_rows = sorted(
        rows,
        key=lambda row: (
            (row.get(STOCK_COLUMN) or "").strip(),
            (row.get(DATE_COLUMN) or "").strip(),
        ),
    )

    with OUTPUT_FILE.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=OUTPUT_HEADERS)
        writer.writeheader()
        writer.writerows(sorted_rows)


def normalize_symbol(symbol: str) -> str:
    return symbol.strip().upper()


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


def format_number(value: float) -> str:
    return f"{value:.4f}"


def format_quantity(value: float) -> str:
    return str(int(round(value)))


def parse_contract_notes_charge_data() -> Dict[str, Dict[str, float]]:
    if load_workbook is None or not CONTRACT_NOTES_FILE.exists():
        return {}

    charge_labels = {
        "Taxable value of Supply (Brokerage) (₹)",
        "Exchange transaction charges (₹)",
        "Clearing charges (₹)",
        "CGST (@9% of Brok, SEBI, Trans & Clearing Charges) (₹)",
        "SGST (@9% of Brok, SEBI, Trans & Clearing Charges) (₹)",
        "IGST (@18% of Brok, SEBI, Trans & Clearing Charges) (₹)",
        "Securities transaction tax (₹)",
        "SEBI turnover fees (₹)",
        "Stamp duty (₹)",
        "NSE Investor Protection Fund Trust (₹)",
    }
    charge_data_by_date: Dict[str, Dict[str, float]] = {}

    workbook = load_workbook(CONTRACT_NOTES_FILE, read_only=True, data_only=False)
    for sheet_name in workbook.sheetnames:
        try:
            sheet_date = datetime.strptime(sheet_name.strip(), "%d-%m-%Y").date().isoformat()
        except ValueError:
            continue
        worksheet = workbook[sheet_name]
        symbol_sell_value: Dict[str, float] = {}
        total_charges = 0.0

        in_trade_rows = False
        for row_number in range(1, 600):
            first_cell = worksheet.cell(row=row_number, column=1).value
            first_text = str(first_cell or "").strip()
            if first_text == "Order No.":
                in_trade_rows = True
                continue

            if in_trade_rows:
                if first_text.startswith("Pay in/Pay out obligation"):
                    in_trade_rows = False
                else:
                    side = str(worksheet.cell(row=row_number, column=6).value or "").strip().upper()
                    contract_desc = str(worksheet.cell(row=row_number, column=5).value or "").strip()
                    net_total = worksheet.cell(row=row_number, column=13).value
                    if side == "S" and contract_desc:
                        symbol = normalize_symbol(contract_desc.split("-")[0])
                        if symbol:
                            try:
                                net_value = abs(float(net_total))
                            except (TypeError, ValueError):
                                net_value = 0.0
                            symbol_sell_value[symbol] = symbol_sell_value.get(symbol, 0.0) + net_value

            if first_text in charge_labels:
                # Contract note keeps charge values in "NCL-Cash" / "NET TOTAL" columns (H/K).
                raw_amount = worksheet.cell(row=row_number, column=8).value
                if raw_amount in (None, ""):
                    raw_amount = worksheet.cell(row=row_number, column=11).value
                try:
                    charge_amount = abs(float(raw_amount))
                except (TypeError, ValueError):
                    charge_amount = 0.0
                total_charges += charge_amount

        charge_data_by_date[sheet_date] = {
            "__total_charges__": total_charges,
            **symbol_sell_value,
        }

    return charge_data_by_date


def build_pnl_summary(report_rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    if not TRADEBOOK_FILE.exists():
        return []

    market_data = load_latest_market_data(report_rows)
    recommendation_dates: Dict[str, str] = {}
    for row in report_rows:
        stock_code = (row.get(STOCK_COLUMN) or "").strip()
        recommendation_date = (row.get(DATE_COLUMN) or "").strip()
        base_symbol = normalize_symbol(stock_code.split(".")[0])
        if base_symbol and recommendation_date:
            recommendation_dates[base_symbol] = recommendation_date
    lots_by_symbol: Dict[str, List[Dict[str, float | str]]] = {}
    trades: List[Tuple[str, str, float, float, str, int]] = []

    with TRADEBOOK_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        required_columns = {"symbol", "trade_type", "quantity", "price", "trade_date"}
        if reader.fieldnames is None:
            raise ValueError(f"Tradebook file is empty: {TRADEBOOK_FILE}")
        missing_columns = required_columns.difference(reader.fieldnames)
        if missing_columns:
            raise ValueError(
                "Tradebook file is missing required columns: "
                + ", ".join(sorted(missing_columns))
            )

        for row_number, row in enumerate(reader, start=2):
            symbol = normalize_symbol((row.get("symbol") or "").strip())
            trade_type = (row.get("trade_type") or "").strip().lower()
            trade_date = (row.get("trade_date") or "").strip()
            if not symbol or trade_type not in {"buy", "sell"}:
                raise ValueError(
                    f"Tradebook row {row_number} has invalid symbol or trade_type."
                )
            if not trade_date:
                raise ValueError(f"Tradebook row {row_number} is missing trade_date.")

            try:
                quantity = float((row.get("quantity") or "").strip())
                price = float((row.get("price") or "").strip())
            except ValueError as exc:
                raise ValueError(
                    f"Tradebook row {row_number} has invalid numeric quantity or price."
                ) from exc

            sort_timestamp = (row.get("order_execution_time") or "").strip() or trade_date
            trades.append((sort_timestamp, symbol, quantity, price, trade_type, row_number))

    trades.sort(key=lambda item: (item[0], item[5]))

    for _sort_time, symbol, quantity, price, trade_type, row_number in trades:
        symbol_lots = lots_by_symbol.setdefault(symbol, [])

        if trade_type == "buy":
            trade_date = _sort_time[:10]
            lot = next(
                (
                    entry
                    for entry in symbol_lots
                    if str(entry["trade_date"]) == trade_date
                ),
                None,
            )
            if lot is None:
                lot = {
                    "trade_date": trade_date,
                    "buy_quantity": 0.0,
                    "buy_value": 0.0,
                    "sell_quantity": 0.0,
                    "sell_value": 0.0,
                    "position_quantity": 0.0,
                    "position_cost": 0.0,
                    "realized_pnl": 0.0,
                    "charges_taxes_others": 0.0,
                    "realized_cost_basis": 0.0,
                }
                symbol_lots.append(lot)

            lot["buy_quantity"] = float(lot["buy_quantity"]) + quantity
            lot["buy_value"] = float(lot["buy_value"]) + (quantity * price)
            lot["position_quantity"] = float(lot["position_quantity"]) + quantity
            lot["position_cost"] = float(lot["position_cost"]) + (quantity * price)
            continue

        remaining_to_sell = quantity
        for lot in symbol_lots:
            lot_position_quantity = float(lot["position_quantity"])
            if lot_position_quantity <= 0:
                continue

            matched_quantity = min(remaining_to_sell, lot_position_quantity)
            average_cost = float(lot["position_cost"]) / lot_position_quantity if lot_position_quantity else 0.0

            lot["sell_quantity"] = float(lot["sell_quantity"]) + matched_quantity
            lot["sell_value"] = float(lot["sell_value"]) + (matched_quantity * price)
            lot["realized_pnl"] = float(lot["realized_pnl"]) + (matched_quantity * (price - average_cost))
            lot["realized_cost_basis"] = float(lot["realized_cost_basis"]) + (matched_quantity * average_cost)
            lot["position_quantity"] = lot_position_quantity - matched_quantity
            lot["position_cost"] = float(lot["position_cost"]) - (matched_quantity * average_cost)

            remaining_to_sell -= matched_quantity
            if remaining_to_sell <= 0:
                break

        # If sells exceed open lots, keep behavior non-short: ignore unmatched remainder.
        if remaining_to_sell > 0:
            print(
                f"Warning: Ignored unmatched sell quantity ({remaining_to_sell:.4f}) for {symbol} at row {row_number}.",
                file=sys.stderr,
            )

    contract_note_charges = parse_contract_notes_charge_data()
    if contract_note_charges:
        unallocated_contract_note_charges = 0.0
        for charge_date, payload in contract_note_charges.items():
            total_charges = float(payload.get("__total_charges__", 0.0))
            if total_charges <= 0:
                continue

            symbol_weights = {
                symbol: float(value)
                for symbol, value in payload.items()
                if symbol != "__total_charges__" and float(value) > 0
            }
            total_symbol_weight = sum(symbol_weights.values())

            date_lots_all_symbols: List[Tuple[str, Dict[str, float | str]]] = []
            for symbol, symbol_lots in lots_by_symbol.items():
                for lot in symbol_lots:
                    if str(lot["trade_date"]) != charge_date:
                        continue
                    if float(lot["sell_value"]) <= 0:
                        continue
                    date_lots_all_symbols.append((symbol, lot))

            if not date_lots_all_symbols:
                unallocated_contract_note_charges += total_charges
                continue

            per_symbol_charge: Dict[str, float] = {}
            if total_symbol_weight > 0:
                running_allocated = 0.0
                symbols = sorted(symbol_weights.keys())
                for index, symbol in enumerate(symbols):
                    symbol_weight = symbol_weights[symbol]
                    if index == len(symbols) - 1:
                        allocated_charge = total_charges - running_allocated
                    else:
                        allocated_charge = total_charges * (symbol_weight / total_symbol_weight)
                        running_allocated += allocated_charge
                    per_symbol_charge[symbol] = per_symbol_charge.get(symbol, 0.0) + allocated_charge
            else:
                realized_value_on_date = sum(float(lot["sell_value"]) for _, lot in date_lots_all_symbols)
                if realized_value_on_date <= 0:
                    continue
                running_allocated = 0.0
                for index, (symbol, lot) in enumerate(date_lots_all_symbols):
                    sell_value = float(lot["sell_value"])
                    if index == len(date_lots_all_symbols) - 1:
                        allocated_charge = total_charges - running_allocated
                    else:
                        allocated_charge = total_charges * (sell_value / realized_value_on_date)
                        running_allocated += allocated_charge
                    lot["charges_taxes_others"] = float(lot["charges_taxes_others"]) + allocated_charge
                continue

            for symbol, symbol_charge in per_symbol_charge.items():
                symbol_lots = [
                    lot
                    for lot_symbol, lot in date_lots_all_symbols
                    if lot_symbol == symbol
                ]
                if not symbol_lots:
                    unallocated_contract_note_charges += symbol_charge
                    continue
                symbol_sell_value = sum(float(lot["sell_value"]) for lot in symbol_lots)
                if symbol_sell_value <= 0:
                    unallocated_contract_note_charges += symbol_charge
                    continue
                running_allocated = 0.0
                for index, lot in enumerate(symbol_lots):
                    sell_value = float(lot["sell_value"])
                    if index == len(symbol_lots) - 1:
                        allocated_charge = symbol_charge - running_allocated
                    else:
                        allocated_charge = symbol_charge * (sell_value / symbol_sell_value)
                        running_allocated += allocated_charge
                    lot["charges_taxes_others"] = float(lot["charges_taxes_others"]) + allocated_charge

        if unallocated_contract_note_charges > 0:
            realized_lots_all_symbols = [
                lot
                for symbol_lots in lots_by_symbol.values()
                for lot in symbol_lots
                if float(lot["sell_value"]) > 0
            ]
            realized_total_sell_value = sum(float(lot["sell_value"]) for lot in realized_lots_all_symbols)
            if realized_total_sell_value > 0:
                allocated = 0.0
                for index, lot in enumerate(realized_lots_all_symbols):
                    sell_value = float(lot["sell_value"])
                    if index == len(realized_lots_all_symbols) - 1:
                        lot_charge = unallocated_contract_note_charges - allocated
                    else:
                        lot_charge = unallocated_contract_note_charges * (sell_value / realized_total_sell_value)
                        allocated += lot_charge
                    lot["charges_taxes_others"] = float(lot["charges_taxes_others"]) + lot_charge
    else:
        sale_charge_events: List[Tuple[str, float]] = []
        generic_charge_pool = 0.0
        if ALL_LEDGER_FILE.exists():
            with ALL_LEDGER_FILE.open("r", newline="", encoding="utf-8-sig") as file:
                reader = csv.DictReader(file)
                for row in reader:
                    particulars = (row.get("particulars") or "").strip()
                    particulars_lower = particulars.lower()
                    debit_amount = parse_amount(row.get("debit") or "")
                    if debit_amount <= 0:
                        continue

                    # Inclusion approach: only count debits that are explicitly known
                    # brokerage/exchange charges. Unknown entries are ignored so new
                    # ledger categories never silently inflate the charge pool.
                    CHARGE_INCLUDE_TOKENS = (
                        "dp charges",
                        "brokerage",
                        "stamp duty",
                        "stt",
                        "securities transaction tax",
                        "sebi",
                        "exchange transaction charge",
                        "clearing charge",
                        "cgst",
                        "sgst",
                        "igst",
                    )
                    is_known_charge = any(token in particulars_lower for token in CHARGE_INCLUDE_TOKENS)
                    if not is_known_charge:
                        continue

                    matched_sale_charge = re.search(r"sale of ([A-Za-z0-9._-]+)", particulars, flags=re.IGNORECASE)
                    if matched_sale_charge:
                        charge_symbol = normalize_symbol(matched_sale_charge.group(1))
                        if charge_symbol:
                            sale_charge_events.append((charge_symbol, debit_amount))
                            continue

                    generic_charge_pool += debit_amount

        for symbol, charge_amount in sale_charge_events:
            symbol_lots = lots_by_symbol.get(symbol, [])
            realized_lots = [lot for lot in symbol_lots if float(lot["sell_value"]) > 0]
            total_sell_value = sum(float(lot["sell_value"]) for lot in realized_lots)
            if total_sell_value <= 0:
                generic_charge_pool += charge_amount
                continue
            allocated = 0.0
            for index, lot in enumerate(realized_lots):
                sell_value = float(lot["sell_value"])
                if index == len(realized_lots) - 1:
                    lot_charge = charge_amount - allocated
                else:
                    lot_charge = charge_amount * (sell_value / total_sell_value)
                    allocated += lot_charge
                lot["charges_taxes_others"] = float(lot["charges_taxes_others"]) + lot_charge

        realized_lots_all_symbols = [
            lot
            for symbol_lots in lots_by_symbol.values()
            for lot in symbol_lots
            if float(lot["sell_value"]) > 0
        ]
        realized_total_sell_value = sum(float(lot["sell_value"]) for lot in realized_lots_all_symbols)
        if generic_charge_pool > 0 and realized_total_sell_value > 0:
            allocated = 0.0
            for index, lot in enumerate(realized_lots_all_symbols):
                sell_value = float(lot["sell_value"])
                if index == len(realized_lots_all_symbols) - 1:
                    lot_charge = generic_charge_pool - allocated
                else:
                    lot_charge = generic_charge_pool * (sell_value / realized_total_sell_value)
                    allocated += lot_charge
                lot["charges_taxes_others"] = float(lot["charges_taxes_others"]) + lot_charge

    sortable_rows: List[Tuple[str, str, Dict[str, str]]] = []
    for symbol, symbol_lots in lots_by_symbol.items():
        for lot in symbol_lots:
            position_quantity = float(lot["position_quantity"])
            position_cost = float(lot["position_cost"])
            buy_quantity = float(lot["buy_quantity"])
            buy_value = float(lot["buy_value"])
            sell_quantity = float(lot["sell_quantity"])
            sell_value = float(lot["sell_value"])
            realized_pnl = float(lot["realized_pnl"])
            charges_taxes_others = float(lot["charges_taxes_others"])
            net_realized_pnl = realized_pnl - charges_taxes_others

            average_buy_price = (buy_value / buy_quantity) if buy_quantity else 0.0
            average_sell_price = (sell_value / sell_quantity) if sell_quantity else 0.0
            open_average_cost = (position_cost / position_quantity) if position_quantity else 0.0

            market_entry = market_data.get(symbol, {})
            latest_market_price = float(market_entry.get("latest_market_price", "0") or 0.0)
            latest_market_date = str(market_entry.get("latest_market_date", "") or "")
            matched_report_symbol = str(market_entry.get("report_symbol", symbol) or symbol)
            unrealized_pnl = (
                position_quantity * (latest_market_price - open_average_cost)
                if position_quantity and latest_market_price
                else 0.0
            )
            total_pnl = net_realized_pnl + unrealized_pnl
            total_return_pct = ((total_pnl / buy_value) * 100) if buy_value else None

            row = {
                "Stock Code/Name": symbol,
                "Trade Date": str(lot["trade_date"]),
                "Matched Report Symbol": matched_report_symbol,
                "Buy Quantity": format_quantity(buy_quantity),
                "Buy Value": format_number(buy_value),
                "Sell Quantity": format_quantity(sell_quantity),
                "Sell Value": format_number(sell_value),
                "Net Quantity": format_quantity(position_quantity),
                "Average Buy Price": format_number(average_buy_price),
                "Average Sell Price": format_number(average_sell_price),
                "Open Average Cost": format_number(open_average_cost),
                "Latest Market Price": format_number(latest_market_price) if latest_market_price else "",
                "Latest Market Date": latest_market_date,
                "Realized P&L": format_number(realized_pnl),
                "Charges, Taxes, Others": format_number(charges_taxes_others),
                "Net Realized P&L": format_number(net_realized_pnl),
                "Unrealized P&L": format_number(unrealized_pnl),
                "Total P&L": format_number(total_pnl),
                "Return %": f"{total_return_pct:.2f}%" if total_return_pct is not None else "",
            }
            sort_date = recommendation_dates.get(symbol, "9999-12-31")
            trade_date = str(lot["trade_date"])
            sortable_rows.append((sort_date, symbol, trade_date, row))

    sortable_rows.sort(key=lambda item: (item[0], item[1], item[2]))
    return [row for _date, _symbol, _trade_date, row in sortable_rows]


def build_portfolio_timeline(report_rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    """Compute real day-by-day portfolio components, split by realized/unrealized
    and by full-value/MTF-funding, so the dashboard can render any filter combo.

    Per date we emit, for both realized (sold-by-date) and unrealized (still-held)
    buckets: full buy value, current value, and the zerodha (MTF) funding portion.
    The frontend derives the 6 filter combinations:
      - basis "all holdings"  -> use full value (r_invested / u_invested, r_value / u_value)
      - basis "personal"      -> subtract zerodha funding (full - zerodha)
      - filter realized/unrealized/all -> pick r_*, u_*, or their sum
    """
    import bisect

    if not TRADEBOOK_FILE.exists():
        return []

    # Per-buy-date MTF margin ratio. ratio = your_funding / buy_value.
    # Cash (non-MTF) buy dates are absent -> treated as ratio 1.0 (zero zerodha funding).
    margin_ratio_by_date = build_mtf_margin_ratio_by_date()

    # Build price lookup: {base_symbol: {iso_date: price}}
    price_by_symbol: Dict[str, Dict[str, float]] = {}
    for row in report_rows:
        stock_code = (row.get(STOCK_COLUMN) or "").strip()
        base_symbol = normalize_symbol(stock_code.split(".")[0])
        if not base_symbol:
            continue
        sym_prices = price_by_symbol.setdefault(base_symbol, {})
        for index in range(1, DAY_COUNT + 1):
            row_date = (row.get(f"Date {index}") or "").strip()
            row_price = (row.get(f"Day {index} Price") or "").strip()
            if row_date and row_price:
                try:
                    sym_prices[row_date] = float(row_price)
                except ValueError:
                    pass

    if not price_by_symbol:
        return []

    # Read tradebook into chronological trade events (mirror FIFO of build_pnl_summary).
    trades: List[Tuple[str, str, float, float, str]] = []  # (sort_time, symbol, qty, price, trade_type)
    with TRADEBOOK_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        if reader.fieldnames is None:
            return []
        for row in reader:
            symbol = normalize_symbol((row.get("symbol") or "").strip())
            trade_type = (row.get("trade_type") or "").strip().lower()
            quantity = parse_amount(row.get("quantity") or "")
            price = parse_amount(row.get("price") or "")
            trade_date = (row.get("trade_date") or "").strip()
            if not symbol or trade_type not in {"buy", "sell"} or not trade_date:
                continue
            if quantity <= 0 or price <= 0:
                continue
            sort_time = (row.get("order_execution_time") or "").strip() or trade_date
            trades.append((sort_time, symbol, quantity, price, trade_type))

    if not trades:
        return []

    trades.sort(key=lambda t: t[0])

    # Build lots aggregated by (symbol, buy_date). FIFO sells record per-lot sell events
    # with their dates so we can reconstruct realized/unrealized split at any past date.
    lots_by_symbol: Dict[str, List[Dict[str, object]]] = {}
    for sort_time, symbol, quantity, price, trade_type in trades:
        symbol_lots = lots_by_symbol.setdefault(symbol, [])
        if trade_type == "buy":
            trade_date = sort_time[:10]
            lot = next((l for l in symbol_lots if l["trade_date"] == trade_date), None)
            if lot is None:
                lot = {
                    "trade_date": trade_date,
                    "buy_quantity": 0.0,
                    "buy_value": 0.0,
                    "zerodha_value": 0.0,
                    "position_quantity": 0.0,
                    "sells": [],  # list of (sell_date, qty, sell_price)
                }
                symbol_lots.append(lot)
            ratio = margin_ratio_by_date.get(trade_date, 1.0)
            buy_value = quantity * price
            lot["buy_quantity"] = float(lot["buy_quantity"]) + quantity
            lot["buy_value"] = float(lot["buy_value"]) + buy_value
            lot["zerodha_value"] = float(lot["zerodha_value"]) + buy_value * (1.0 - ratio)
            lot["position_quantity"] = float(lot["position_quantity"]) + quantity
            continue

        # sell: FIFO across this symbol's lots in chronological order
        sell_date = sort_time[:10]
        remaining = quantity
        for lot in symbol_lots:
            open_qty = float(lot["position_quantity"])
            if open_qty <= 0:
                continue
            matched = min(remaining, open_qty)
            lot["position_quantity"] = open_qty - matched
            lot["sells"].append((sell_date, matched, price))
            remaining -= matched
            if remaining <= 0:
                break

    # All market dates across all stocks in the price data
    all_market_dates = sorted({d for sym_prices in price_by_symbol.values() for d in sym_prices})
    if not all_market_dates:
        return []

    sorted_dates_by_symbol: Dict[str, List[str]] = {
        sym: sorted(dates.keys()) for sym, dates in price_by_symbol.items()
    }

    def price_on(symbol: str, market_date: str, fallback: float) -> float:
        sym_dates = sorted_dates_by_symbol.get(symbol)
        sym_prices = price_by_symbol.get(symbol)
        if not sym_dates or not sym_prices:
            return fallback
        pos = bisect.bisect_right(sym_dates, market_date)
        if pos == 0:
            return fallback
        return sym_prices[sym_dates[pos - 1]]

    timeline: List[Dict[str, str]] = []
    for market_date in all_market_dates:
        r_invested = r_value = r_zerodha = 0.0
        u_invested = u_value = u_zerodha = 0.0

        for symbol, symbol_lots in lots_by_symbol.items():
            for lot in symbol_lots:
                if lot["trade_date"] > market_date:
                    continue  # lot not bought yet
                buy_qty = float(lot["buy_quantity"])
                if buy_qty <= 0:
                    continue
                buy_price_ps = float(lot["buy_value"]) / buy_qty
                zerodha_ps = float(lot["zerodha_value"]) / buy_qty

                sold_qty = 0.0
                realized_value = 0.0
                for sell_date, qty, sell_price in lot["sells"]:
                    if sell_date <= market_date:
                        sold_qty += qty
                        realized_value += qty * sell_price
                open_qty = buy_qty - sold_qty

                if sold_qty > 0:
                    r_invested += sold_qty * buy_price_ps
                    r_value += realized_value
                    r_zerodha += sold_qty * zerodha_ps
                if open_qty > 0:
                    u_invested += open_qty * buy_price_ps
                    u_value += open_qty * price_on(symbol, market_date, buy_price_ps)
                    u_zerodha += open_qty * zerodha_ps

        if r_invested <= 0 and u_invested <= 0:
            continue

        timeline.append({
            "date": market_date,
            "r_invested": f"{r_invested:.2f}",
            "r_value": f"{r_value:.2f}",
            "r_zerodha": f"{r_zerodha:.2f}",
            "u_invested": f"{u_invested:.2f}",
            "u_value": f"{u_value:.2f}",
            "u_zerodha": f"{u_zerodha:.2f}",
        })

    return timeline


def write_portfolio_timeline_output(rows: Sequence[Dict[str, str]]) -> None:
    with PORTFOLIO_TIMELINE_FILE.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=PORTFOLIO_TIMELINE_HEADERS)
        writer.writeheader()
        writer.writerows(rows)


def write_pnl_output(rows: Sequence[Dict[str, str]]) -> None:
    with PNL_OUTPUT_FILE.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=PNL_HEADERS)
        writer.writeheader()
        writer.writerows(rows)


def parse_amount(value: str) -> float:
    cleaned = (value or "").replace(",", "").strip()
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def build_mtf_margin_ratio_by_date() -> Dict[str, float]:
    if not ALL_LEDGER_FILE.exists() or not MTF_LEDGER_FILE.exists():
        return {}

    initial_margin_by_date: Dict[str, float] = {}
    with ALL_LEDGER_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            particulars = (row.get("particulars") or "").strip().lower()
            posting_date = (row.get("posting_date") or "").strip()
            if not posting_date:
                continue
            if particulars != "initial margin charged for mtf":
                continue
            initial_margin_by_date[posting_date] = initial_margin_by_date.get(posting_date, 0.0) + parse_amount(
                row.get("debit") or ""
            )

    gross_obligation_by_date: Dict[str, float] = {}
    with MTF_LEDGER_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            particulars = (row.get("particulars") or "").strip().lower()
            posting_date = (row.get("posting_date") or "").strip()
            if not posting_date:
                continue
            if particulars != "gross obligation for mtf":
                continue
            debit_value = parse_amount(row.get("debit") or "")
            if debit_value <= 0:
                continue
            gross_obligation_by_date[posting_date] = gross_obligation_by_date.get(posting_date, 0.0) + debit_value

    ratios: Dict[str, float] = {}
    for posting_date, margin_value in initial_margin_by_date.items():
        gross_value = gross_obligation_by_date.get(posting_date, 0.0)
        if gross_value <= 0:
            continue
        ratios[posting_date] = margin_value / gross_value
    return ratios


def build_mtf_pnl_summary(report_rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    if not TRADEBOOK_FILE.exists():
        return []

    margin_ratio_by_date = build_mtf_margin_ratio_by_date()
    if not margin_ratio_by_date:
        return []

    market_data = load_latest_market_data(report_rows)
    lots_by_symbol: Dict[str, List[Dict[str, float | str]]] = {}
    events: List[Tuple[str, str, Dict[str, float | str], int]] = []

    with TRADEBOOK_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        required_columns = {"symbol", "trade_type", "quantity", "price", "trade_date"}
        if reader.fieldnames is None:
            raise ValueError(f"Tradebook file is empty: {TRADEBOOK_FILE}")
        missing_columns = required_columns.difference(reader.fieldnames)
        if missing_columns:
            raise ValueError(
                "Tradebook file is missing required columns: "
                + ", ".join(sorted(missing_columns))
            )

        for row_number, row in enumerate(reader, start=2):
            symbol = normalize_symbol((row.get("symbol") or "").strip())
            trade_type = (row.get("trade_type") or "").strip().lower()
            trade_date = (row.get("trade_date") or "").strip()
            if not symbol or trade_type not in {"buy", "sell"} or not trade_date:
                continue

            quantity = parse_amount(row.get("quantity") or "")
            price = parse_amount(row.get("price") or "")
            if quantity <= 0 or price <= 0:
                continue

            sort_timestamp = (row.get("order_execution_time") or "").strip() or trade_date
            events.append(
                (
                    sort_timestamp,
                    "trade",
                    {
                        "symbol": symbol,
                        "quantity": quantity,
                        "price": price,
                        "trade_type": trade_type,
                        "trade_date": trade_date,
                    },
                    row_number,
                )
            )

    if ALL_LEDGER_FILE.exists():
        with ALL_LEDGER_FILE.open("r", newline="", encoding="utf-8-sig") as file:
            reader = csv.DictReader(file)
            for row_number, row in enumerate(reader, start=2):
                particulars = (row.get("particulars") or "").strip()
                particulars_lower = particulars.lower()
                posting_date = (row.get("posting_date") or "").strip()
                if not posting_date:
                    continue

                if "interest for mtf funded value on " in particulars_lower:
                    matched = re.search(r"on\s+(\d{4}-\d{2}-\d{2})", particulars)
                    accrual_date = matched.group(1) if matched else posting_date
                    debit_value = parse_amount(row.get("debit") or "")
                    if debit_value <= 0:
                        continue
                    events.append(
                        (
                            f"{accrual_date}T23:59:59",
                            "interest",
                            {"amount": debit_value},
                            row_number,
                        )
                    )
                    continue

                if "mtf pledge charges for " in particulars_lower or "mtf unpledge charges for " in particulars_lower:
                    debit_value = parse_amount(row.get("debit") or "")
                    if debit_value <= 0:
                        continue
                    matched_symbol = re.search(r"for\s+([A-Za-z0-9._-]+)", particulars, flags=re.IGNORECASE)
                    if not matched_symbol:
                        continue
                    events.append(
                        (
                            f"{posting_date}T23:59:59",
                            "pledge",
                            {
                                "symbol": normalize_symbol(matched_symbol.group(1)),
                                "amount": debit_value,
                            },
                            row_number,
                        )
                    )

    events.sort(key=lambda item: (item[0], item[3]))

    def allocate_amount_across_lots(target_lots: Sequence[Dict[str, float | str]], amount: float, field: str) -> None:
        total_open_funding = sum(float(lot["open_zerodha_funding"]) for lot in target_lots if float(lot["open_zerodha_funding"]) > 0)
        if amount <= 0 or total_open_funding <= 0:
            return

        running_allocated = 0.0
        open_lots = [lot for lot in target_lots if float(lot["open_zerodha_funding"]) > 0]
        for index, lot in enumerate(open_lots):
            open_funding = float(lot["open_zerodha_funding"])
            if index == len(open_lots) - 1:
                allocation = amount - running_allocated
            else:
                allocation = amount * (open_funding / total_open_funding)
                running_allocated += allocation
            lot[field] = float(lot[field]) + allocation

    for _event_time, event_type, payload, _row_number in events:
        if event_type == "trade":
            symbol = str(payload["symbol"])
            quantity = float(payload["quantity"])
            price = float(payload["price"])
            trade_type = str(payload["trade_type"])
            trade_date = str(payload["trade_date"])
            symbol_lots = lots_by_symbol.setdefault(symbol, [])

            if trade_type == "buy":
                # Cash (non-MTF) buys still create a lot so FIFO sell-matching stays
                # aligned with the P&L summary; they carry zero funding and are excluded
                # from the final MTF rows via the is_mtf flag. Skipping them here would
                # mis-assign sell prices to MTF lots for mixed cash+MTF symbols.
                ratio = margin_ratio_by_date.get(trade_date)
                is_mtf = ratio is not None
                effective_ratio = ratio if ratio is not None else 1.0
                lot = next((entry for entry in symbol_lots if str(entry["trade_date"]) == trade_date), None)
                if lot is None:
                    lot = {
                        "trade_date": trade_date,
                        "is_mtf": is_mtf,
                        "buy_quantity": 0.0,
                        "buy_value": 0.0,
                        "your_funding": 0.0,
                        "zerodha_funding": 0.0,
                        "sell_quantity": 0.0,
                        "sell_value": 0.0,
                        "position_quantity": 0.0,
                        "position_cost": 0.0,
                        "open_zerodha_funding": 0.0,
                        "realized_pnl": 0.0,
                        "mtf_interest_cost": 0.0,
                        "mtf_pledge_charges": 0.0,
                    }
                    symbol_lots.append(lot)

                buy_value = quantity * price
                your_funding = buy_value * effective_ratio
                zerodha_funding = buy_value - your_funding
                lot["buy_quantity"] = float(lot["buy_quantity"]) + quantity
                lot["buy_value"] = float(lot["buy_value"]) + buy_value
                lot["your_funding"] = float(lot["your_funding"]) + your_funding
                lot["zerodha_funding"] = float(lot["zerodha_funding"]) + zerodha_funding
                lot["position_quantity"] = float(lot["position_quantity"]) + quantity
                lot["position_cost"] = float(lot["position_cost"]) + buy_value
                lot["open_zerodha_funding"] = float(lot["open_zerodha_funding"]) + zerodha_funding
                continue

            remaining_to_sell = quantity
            for lot in symbol_lots:
                lot_position_quantity = float(lot["position_quantity"])
                if lot_position_quantity <= 0:
                    continue

                matched_quantity = min(remaining_to_sell, lot_position_quantity)
                average_cost = float(lot["position_cost"]) / lot_position_quantity if lot_position_quantity else 0.0
                matched_value = matched_quantity * price

                lot["sell_quantity"] = float(lot["sell_quantity"]) + matched_quantity
                lot["sell_value"] = float(lot["sell_value"]) + matched_value
                lot["realized_pnl"] = float(lot["realized_pnl"]) + (matched_quantity * (price - average_cost))
                lot["position_quantity"] = lot_position_quantity - matched_quantity
                lot["position_cost"] = float(lot["position_cost"]) - (matched_quantity * average_cost)

                open_funding = float(lot["open_zerodha_funding"])
                funding_per_share = open_funding / lot_position_quantity if lot_position_quantity else 0.0
                lot["open_zerodha_funding"] = max(0.0, open_funding - (matched_quantity * funding_per_share))

                remaining_to_sell -= matched_quantity
                if remaining_to_sell <= 0:
                    break
            continue

        if event_type == "interest":
            open_lots = [
                lot
                for symbol_lots in lots_by_symbol.values()
                for lot in symbol_lots
                if float(lot["open_zerodha_funding"]) > 0
            ]
            allocate_amount_across_lots(open_lots, float(payload["amount"]), "mtf_interest_cost")
            continue

        if event_type == "pledge":
            symbol = str(payload["symbol"])
            symbol_lots = lots_by_symbol.get(symbol, [])
            open_symbol_lots = [lot for lot in symbol_lots if float(lot["open_zerodha_funding"]) > 0]
            allocate_amount_across_lots(open_symbol_lots, float(payload["amount"]), "mtf_pledge_charges")

    sortable_rows: List[Tuple[str, str, Dict[str, str]]] = []
    for symbol, symbol_lots in lots_by_symbol.items():
        for lot in symbol_lots:
            buy_quantity = float(lot["buy_quantity"])
            buy_value = float(lot["buy_value"])
            your_funding = float(lot["your_funding"])
            zerodha_funding = float(lot["zerodha_funding"])
            sell_quantity = float(lot["sell_quantity"])
            sell_value = float(lot["sell_value"])
            position_quantity = float(lot["position_quantity"])
            position_cost = float(lot["position_cost"])
            realized_pnl = float(lot["realized_pnl"])
            mtf_interest_cost = float(lot["mtf_interest_cost"])
            mtf_pledge_charges = float(lot["mtf_pledge_charges"])

            if buy_quantity <= 0:
                continue
            if not lot.get("is_mtf"):
                continue  # cash lot — tracked only for FIFO alignment, not an MTF row

            market_entry = market_data.get(symbol, {})
            latest_market_price = float(market_entry.get("latest_market_price", "0") or 0.0)
            latest_market_date = str(market_entry.get("latest_market_date", "") or "")
            matched_report_symbol = str(market_entry.get("report_symbol", symbol) or symbol)
            open_average_cost = (position_cost / position_quantity) if position_quantity else 0.0
            unrealized_pnl = (
                position_quantity * (latest_market_price - open_average_cost)
                if position_quantity and latest_market_price
                else 0.0
            )
            holding_pnl = realized_pnl + unrealized_pnl
            holding_return_pct = ((holding_pnl / buy_value) * 100) if buy_value else None
            funding_return_pct = ((holding_pnl / your_funding) * 100) if your_funding else None
            carrying_cost = mtf_interest_cost + mtf_pledge_charges
            net_funding_pnl = holding_pnl - carrying_cost
            net_funding_return_pct = ((net_funding_pnl / your_funding) * 100) if your_funding else None
            leverage = (buy_value / your_funding) if your_funding else None

            row = {
                "Stock Code/Name": symbol,
                "Trade Date": str(lot["trade_date"]),
                "Matched Report Symbol": matched_report_symbol,
                "Buy Quantity": format_quantity(buy_quantity),
                "Buy Value": format_number(buy_value),
                "Your Funding": format_number(your_funding),
                "Zerodha Funding": format_number(zerodha_funding),
                "Leverage X": f"{leverage:.2f}x" if leverage is not None else "",
                "Sell Quantity": format_quantity(sell_quantity),
                "Sell Value": format_number(sell_value),
                "Net Quantity": format_quantity(position_quantity),
                "Latest Market Price": format_number(latest_market_price) if latest_market_price else "",
                "Latest Market Date": latest_market_date,
                "Holding P&L": format_number(holding_pnl),
                "Holding Return %": f"{holding_return_pct:.2f}%" if holding_return_pct is not None else "",
                "Funding Return %": f"{funding_return_pct:.2f}%" if funding_return_pct is not None else "",
                "MTF Interest Cost": format_number(mtf_interest_cost),
                "MTF Pledge Charges": format_number(mtf_pledge_charges),
                "Total Carrying Cost": format_number(carrying_cost),
                "Net Funding P&L": format_number(net_funding_pnl),
                "Net Funding Return %": f"{net_funding_return_pct:.2f}%" if net_funding_return_pct is not None else "",
            }
            trade_date = str(lot["trade_date"])
            sortable_rows.append((trade_date, symbol, row))

    sortable_rows.sort(key=lambda item: (item[0], item[1]))
    return [row for _trade_date, _symbol, row in sortable_rows]


def write_mtf_pnl_output(rows: Sequence[Dict[str, str]]) -> None:
    with MTF_PNL_OUTPUT_FILE.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=MTF_PNL_HEADERS)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    ensure_directories()

    try:
        recommendations = read_recommendations()
        existing_rows, _existing_keys = read_existing_output()
        existing_by_key = {
            (
                (row.get(STOCK_COLUMN) or "").strip(),
                (row.get(DATE_COLUMN) or "").strip(),
            ): row
            for row in existing_rows
        }
        refreshed_rows: List[Dict[str, str]] = []

        if recommendations:
            print(f"Refreshing price data for {len(recommendations)} recommendation(s).")
            total = len(recommendations)
            for index, recommendation in enumerate(recommendations, start=1):
                existing_row = existing_by_key.get(recommendation.key)
                if existing_row is not None and has_all_four_targets_hit(existing_row):
                    print(
                        f"[{index}/{total}] Skipping API fetch for "
                        f"{recommendation.stock_code} ({recommendation.recommendation_date}) - all 4 targets already hit."
                    )
                    refreshed_rows.append(existing_row)
                    continue

                mode = "incremental" if existing_row is not None else "full"
                print(
                    f"[{index}/{total}] Fetching Yahoo Finance data for "
                    f"{recommendation.stock_code} ({recommendation.recommendation_date}) [{mode}]"
                )
                try:
                    refreshed_rows.append(build_output_row(recommendation, existing_row))
                    time.sleep(0.2)
                except urllib.error.HTTPError as exc:
                    raise RuntimeError(
                        f"Yahoo Finance request failed for {recommendation.stock_code}: HTTP {exc.code}"
                    ) from exc
                except urllib.error.URLError as exc:
                    raise RuntimeError(
                        f"Network error while fetching {recommendation.stock_code}: {exc.reason}"
                    ) from exc
        else:
            print("No recommendations found in input. Nothing to refresh.")

        final_rows = refreshed_rows
        if refreshed_rows or not OUTPUT_FILE.exists():
            write_output(final_rows)
            print(f"Saved price report to: {OUTPUT_FILE}")

        pnl_rows = build_pnl_summary(final_rows)
        if pnl_rows:
            write_pnl_output(pnl_rows)
            print(f"Saved P&L report to: {PNL_OUTPUT_FILE}")
        else:
            print("No tradebook data found. Skipped P&L summary generation.")

        mtf_pnl_rows = build_mtf_pnl_summary(final_rows)
        if mtf_pnl_rows:
            write_mtf_pnl_output(mtf_pnl_rows)
            print(f"Saved MTF P&L report to: {MTF_PNL_OUTPUT_FILE}")
        else:
            print("No MTF summary data found. Skipped MTF P&L summary generation.")

        timeline_rows = build_portfolio_timeline(final_rows)
        if timeline_rows:
            write_portfolio_timeline_output(timeline_rows)
            print(f"Saved portfolio timeline to: {PORTFOLIO_TIMELINE_FILE}")
        else:
            print("No portfolio timeline data generated.")

        return 0

    except Exception as exc:  # pragma: no cover
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
