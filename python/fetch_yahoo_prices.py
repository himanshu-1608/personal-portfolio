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
import sys
import time
import http.cookiejar
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, time as dt_time, timedelta, timezone
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


from portfolio_core import (
    INPUT_DIR,
    REPORTS_DIR,
    STOCK_COLUMN,
    DATE_COLUMN,
    DAY_COUNT,
    PNL_OUTPUT_FILE,
    MTF_PNL_OUTPUT_FILE,
    PORTFOLIO_TIMELINE_FILE,
    format_number,
    normalize_symbol,
)
from portfolio_reports import (
    build_pnl_summary,
    build_mtf_pnl_summary,
    build_portfolio_timeline,
    write_pnl_output,
    write_mtf_pnl_output,
    write_portfolio_timeline_output,
)

INPUT_FILE = INPUT_DIR / "recommendation.csv"
OUTPUT_FILE = REPORTS_DIR / "stock_closing_prices.csv"

TARGET_1_COLUMN = "Target 1"
TARGET_2_COLUMN = "Target 2"
BUY_PRICE_RECOMMENDATION_COLUMN = "Buy Price Recommendation"
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
    BUY_PRICE_RECOMMENDATION_COLUMN,
    "Highest Price",
    "Hit Target 1",
    "Hit Target 2",
    "Target 1 Return %",
    "Target 2 Return %",
    *[
        column_name
        for index in range(1, DAY_COUNT + 1)
        for column_name in (f"Date {index}", f"Day {index} Price", f"Day {index} Return %")
    ],
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


def has_all_targets_hit(row: Dict[str, str]) -> bool:
    return all((row.get(f"Hit Target {index}") or "").strip().upper() == "TRUE" for index in range(1, 3))


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


def _apply_hit_targets(row: Dict[str, str]) -> None:
    highest_price_val = parse_optional_float(row.get("Highest Price") or "")
    for i, target_col in enumerate(["Target 1", "Target 2"], 1):
        target_val = parse_optional_float(row.get(target_col) or "")
        if target_val is None:
            row[f"Hit Target {i}"] = ""
        else:
            row[f"Hit Target {i}"] = (
                "TRUE" if highest_price_val is not None and highest_price_val >= target_val else "FALSE"
            )


def resequence_date_slots(row: Dict[str, str], rec_date: date) -> None:
    """Rewrite the Date/Day Price/Day Return slots so they are strictly increasing
    by date. Drops any date before the recommendation date and de-dupes. Makes the
    table correct regardless of the order prices were fetched or appended."""
    triples: List[Tuple[date, str, str, str]] = []
    seen: set = set()
    for i in range(1, DAY_COUNT + 1):
        d = (row.get(f"Date {i}") or "").strip()
        if not d or d in seen:
            continue
        try:
            parsed = date.fromisoformat(d)
        except ValueError:
            continue
        if parsed < rec_date:
            continue
        seen.add(d)
        triples.append((
            parsed,
            d,
            (row.get(f"Day {i} Price") or "").strip(),
            (row.get(f"Day {i} Return %") or "").strip(),
        ))

    triples.sort(key=lambda t: t[0])
    for i in range(1, DAY_COUNT + 1):
        if i <= len(triples):
            _, d, price, ret = triples[i - 1]
            row[f"Date {i}"] = d
            row[f"Day {i} Price"] = price
            row[f"Day {i} Return %"] = ret
        else:
            row[f"Date {i}"] = ""
            row[f"Day {i} Price"] = ""
            row[f"Day {i} Return %"] = ""


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

        resequence_date_slots(merged_row, rec_date)
        _apply_hit_targets(merged_row)
        return merged_row

    # ── Full fetch path (first run / no existing data) ────────────────────────
    prices, highest_price = fetch_30_day_prices(
        recommendation.stock_code, recommendation.recommendation_date
    )
    target_1 = parse_optional_float(recommendation.target_1) if recommendation.target_1 else None
    target_2 = parse_optional_float(recommendation.target_2) if recommendation.target_2 else None

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
        "Target 1 Return %": (
            f"{(((target_1 / buy_price_recommendation) - 1) * 100):.2f}%"
            if target_1 is not None and buy_price_recommendation not in (None, 0) else ""
        ),
        "Target 2 Return %": (
            f"{(((target_2 / buy_price_recommendation) - 1) * 100):.2f}%"
            if target_2 is not None and buy_price_recommendation not in (None, 0) else ""
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
        # extrasaction="ignore" drops any legacy columns (e.g. the old Target 3/4
        # carried in incremental rows from a previously-written CSV) instead of raising.
        writer = csv.DictWriter(file, fieldnames=OUTPUT_HEADERS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(sorted_rows)


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
                if existing_row is not None and has_all_targets_hit(existing_row):
                    print(
                        f"[{index}/{total}] Skipping API fetch for "
                        f"{recommendation.stock_code} ({recommendation.recommendation_date}) - all targets already hit."
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
                except Exception as exc:
                    # One symbol failing (rate-limit/delisted/network) must NOT abort the
                    # whole refresh. Keep any previously-stored row so the run still
                    # completes and writes every symbol that did succeed.
                    if isinstance(exc, urllib.error.HTTPError):
                        reason = f"HTTP {exc.code}"
                    elif isinstance(exc, urllib.error.URLError):
                        reason = f"network error: {exc.reason}"
                    else:
                        reason = str(exc)
                    print(
                        f"[{index}/{total}] WARNING: fetch failed for "
                        f"{recommendation.stock_code} ({reason}). "
                        f"{'Keeping previously-stored data.' if existing_row is not None else 'Skipping this symbol.'}",
                        file=sys.stderr,
                    )
                    if existing_row is not None:
                        refreshed_rows.append(existing_row)
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
