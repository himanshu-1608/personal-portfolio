"""Download the nightly Zerodha console reports via browser automation.

Logs into https://console.zerodha.com with user id + password + TOTP (generated from
the stored Google-Authenticator secret, so no phone is needed), then downloads three
reports into a dated archive folder:

    zerodha-daily-reports/<YYYY-MM-DD>/
        tradebook-<CLIENT_ID>.csv     (Report tab → Tradebook → segment: Equity)
        all_ledger-<CLIENT_ID>.csv    (Funds tab  → category: All segments)
        mtf_ledger-<CLIENT_ID>.csv    (Funds tab  → category: MTF)

`download.save_as(...)` controls the filename, so the all_ledger / mtf_ledger
same-name collision from Zerodha is avoided.

Credentials come from python/zerodha_secrets.env (gitignored). Set HEADLESS=0 in the
env file for the first run so you can watch / fix selectors.

NOTE: page URLs and selectors below are best-effort and MUST be confirmed against the
live UI during the first headed run — search for "CONFIRM" markers.
"""

from __future__ import annotations

import os
import time
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List

import pyotp
from dotenv import load_dotenv
from playwright.sync_api import Download, Page, sync_playwright

BASE_DIR = Path(__file__).resolve().parent.parent
ENV_FILE = Path(__file__).resolve().parent / "zerodha_secrets.env"
USER_DATA_DIR = BASE_DIR / ".zerodha-browser"  # persistent profile, gitignored
DAILY_REPORTS_DIR = BASE_DIR / "zerodha-daily-reports"

# Login happens on Kite; console pages share the session via SSO afterward.
LOGIN_URL = "https://kite.zerodha.com/"
TRADEBOOK_URL = "https://console.zerodha.com/reports/tradebook"
FUNDS_URL = "https://console.zerodha.com/funds/statement"


def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def click_first(page: Page, selectors: List[str], timeout: int = 2000) -> bool:
    """Click the first selector that resolves. Returns False if none match."""
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=timeout)
            loc.click()
            return True
        except Exception:
            continue
    return False


def fill_first(page: Page, selectors: List[str], value: str, timeout: int = 2000) -> bool:
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=timeout)
            loc.fill(value)
            return True
        except Exception:
            continue
    return False


def login(page: Page, user_id: str, password: str, totp_secret: str) -> None:
    page.goto(LOGIN_URL, wait_until="load")

    # Kite may remember the user id and show password-only (or jump straight to
    # TOTP). Fill whatever is present; require at least one credential field.
    got_userid = fill_first(
        page,
        ['input#userid', 'input[name="userid"]', 'input[placeholder*="User" i]'],
        user_id,
    )
    got_password = fill_first(
        page,
        ['input#password', 'input[name="password"]', 'input[type="password"]'],
        password,
    )
    if not (got_userid or got_password):
        raise RuntimeError("Could not find user-id or password field (CONFIRM login selectors)")
    click_first(page, ['button[type="submit"]', 'button:has-text("Login")'])

    # TOTP step appears after credentials are accepted. Wait for the field, then
    # generate the code so it's fresh (codes rotate every 30s).
    totp_selectors = [
        'input#userid', 'input#pin', 'input[name="totp"]',
        'input[placeholder*="TOTP" i]', 'input[label*="TOTP" i]',
        'input[maxlength="6"]', 'input[type="number"]', 'input[type="text"]',
    ]
    field = None
    for sel in totp_selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=4000)
            field = loc
            break
        except Exception:
            continue
    if field is None:
        raise RuntimeError("Could not find the TOTP field (CONFIRM 2FA selectors)")
    field.fill(pyotp.TOTP(totp_secret).now())
    click_first(page, ['button[type="submit"]', 'button:has-text("Continue")'])

    # After TOTP, Kite redirects off the login page (to dashboard/root). Wait only
    # for that redirect to commit — don't block on a specific path.
    try:
        page.wait_for_url(
            lambda u: "kite.zerodha.com" in u and "/connect/login" not in u,
            timeout=8000,
        )
    except Exception:
        pass
    url = page.url
    if "kite.zerodha.com" not in url or "/connect/login" in url:
        raise RuntimeError(f"Login did not complete; at {url}")


def select_option_by_label(page: Page, label: str) -> bool:
    """Select `label` in whichever <select> on the page offers it as an option."""
    selects = page.locator("select")
    for i in range(selects.count()):
        s = selects.nth(i)
        try:
            opts = [t.strip() for t in s.locator("option").all_inner_texts()]
        except Exception:
            continue
        if label in opts:
            s.select_option(label=label)
            return True
    return False


# The console report pages expose the export as a bare <a>CSV</a> (href="#",
# JS-triggered). It only appears once a report has rendered with data.
CSV_LINK_SELECTORS = ['a:text-is("CSV")', 'a:has-text("CSV")']


def capture_download(page: Page, trigger_selectors: List[str], save_path: Path) -> None:
    save_path.parent.mkdir(parents=True, exist_ok=True)
    with page.expect_download(timeout=30000) as dl_info:
        if not click_first(page, trigger_selectors, timeout=8000):
            raise RuntimeError(
                f"Could not find a download trigger for {save_path.name} "
                f"(CONFIRM the 'CSV' link selector)"
            )
    download: Download = dl_info.value
    download.save_as(str(save_path))
    print(f"Saved {save_path}")


def goto_console(page: Page, url: str) -> None:
    """Navigate to a console page; the first visit may bounce through the Kite
    OAuth handshake before settling back on console.zerodha.com."""
    # "commit" returns as soon as navigation starts (don't block on the SPA
    # loading every asset); readiness is gated below by the URL + control checks.
    page.goto(url, wait_until="commit")
    try:
        page.wait_for_url("**console.zerodha.com/**", timeout=10000)
    except Exception:
        pass
    if "console.zerodha.com" not in page.url:
        raise RuntimeError(f"Could not reach console page {url}; at {page.url}")
    # Wait for the report controls to mount.
    try:
        page.locator("select").first.wait_for(state="visible", timeout=8000)
    except Exception:
        pass


# Download window. MTF settlement rows (pledge/unpledge, obligation, interest)
# post ~T+1..T+2, and a run that misses a couple of nights must still recover
# them, so a 25-day lookback is used instead of the old "last 7 days". The
# merge dedups, so a wider window only ever re-fetches rows we already have.
# The mx-datepicker has no 25-day shortcut, so a custom start~end range is typed
# into the input; if that fails we fall back to the "last 30 days" shortcut.
LOOKBACK_DAYS = 25
FALLBACK_SHORTCUT = "last 30 days"


def pick_date_range(page: Page) -> None:
    """Open the mx-datepicker and set a LOOKBACK_DAYS window, then Apply. Used by
    both the tradebook and funds-statement pages (identical component)."""
    if not click_first(page, ['input.mx-input[name="date"]', 'input.mx-input']):
        raise RuntimeError("Could not open date picker")
    page.wait_for_timeout(300)

    if _type_custom_range(page):
        return

    # Fallback: nearest preset shortcut that still covers the lookback window.
    shortcut = page.locator("button.mx-shortcuts", has_text=FALLBACK_SHORTCUT).first
    try:
        shortcut.wait_for(state="visible", timeout=4000)
        shortcut.click()
    except Exception:
        raise RuntimeError(f"Could not set date range (custom type + '{FALLBACK_SHORTCUT}' both failed)")
    apply_btn = page.locator("button.mx-apply").first
    try:
        if apply_btn.is_visible():
            apply_btn.click()
    except Exception:
        pass


def _type_custom_range(page: Page) -> bool:
    """Type a "start ~ end" range covering the last LOOKBACK_DAYS into the
    mx-datepicker input. Returns True if the field accepted the value."""
    end = date.today()
    start = end - timedelta(days=LOOKBACK_DAYS)
    value = f"{start.isoformat()} ~ {end.isoformat()}"
    try:
        field = page.locator('input.mx-input[name="date"], input.mx-input').first
        field.fill("")
        field.type(value, delay=10)
        field.press("Enter")
        page.wait_for_timeout(300)
        apply_btn = page.locator("button.mx-apply").first
        try:
            if apply_btn.is_visible():
                apply_btn.click()
        except Exception:
            pass
        # Confirm the input actually holds a range (contains "~").
        return "~" in (field.input_value() or "")
    except Exception:
        return False


def build_report_and_download(page: Page, save_path: Path, attempts: int = 3) -> None:
    """Submit the report form, wait for the export link to appear, then download.

    The console occasionally takes longer than 20s to render the export link (slow
    backend), which used to fail the whole nightly run. Re-submit the form and wait
    again a few times before giving up."""
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            click_first(page, ['form button[type="submit"]', 'button.btn-blue', 'button[type="submit"]'])
            page.locator('a:text-is("CSV")').first.wait_for(state="visible", timeout=45000)
            capture_download(page, CSV_LINK_SELECTORS, save_path)
            return
        except Exception as exc:  # noqa: BLE001 - retry any render/download flakiness
            last_error = exc
            print(f"build_report_and_download attempt {attempt}/{attempts} failed: {exc}")
            if attempt < attempts:
                page.wait_for_timeout(3000)
    raise RuntimeError(
        f"Could not produce/download report for {save_path} after {attempts} attempts"
    ) from last_error


def download_tradebook(page: Page, save_path: Path) -> None:
    goto_console(page, TRADEBOOK_URL)
    if not select_option_by_label(page, "Equity"):
        raise RuntimeError("Could not select segment 'Equity' on tradebook")
    pick_date_range(page)
    build_report_and_download(page, save_path)


def download_ledger(page: Page, category_label: str, save_path: Path) -> None:
    goto_console(page, FUNDS_URL)
    if not select_option_by_label(page, category_label):
        raise RuntimeError(f"Could not select category '{category_label}'")
    pick_date_range(page)
    build_report_and_download(page, save_path)


def write_error(date_dir: Path, page=None) -> None:
    """Dump the current traceback (and a screenshot if possible) to
    zerodha-daily-reports/<date>/<millis>_error.txt for later debugging."""
    millis = int(time.time() * 1000)
    err_path = date_dir / f"{millis}_error.txt"
    try:
        date_dir.mkdir(parents=True, exist_ok=True)
        parts = [
            f"timestamp: {datetime.now().isoformat()}",
            f"url: {page.url if page else 'n/a'}",
            "",
            traceback.format_exc(),
        ]
        err_path.write_text("\n".join(parts), encoding="utf-8")
        print(f"ERROR details written to {err_path}")
        if page is not None:
            try:
                page.screenshot(path=str(date_dir / f"{millis}_error.png"), full_page=True)
                print(f"ERROR screenshot saved to {date_dir / f'{millis}_error.png'}")
            except Exception:
                pass
    except Exception as e:
        print(f"Failed to write error file {err_path}: {e}")


def main() -> int:
    load_dotenv(ENV_FILE)
    user_id = env("ZERODHA_USER_ID")
    password = env("ZERODHA_PASSWORD")
    totp_secret = env("ZERODHA_TOTP_SECRET")
    if not all([user_id, password, totp_secret]):
        raise SystemExit(
            f"Missing credentials. Fill {ENV_FILE} with ZERODHA_USER_ID, "
            f"ZERODHA_PASSWORD, ZERODHA_TOTP_SECRET."
        )
    headless = env("HEADLESS", "1") not in ("0", "false", "False", "no")
    client_id = user_id

    date_dir = DAILY_REPORTS_DIR / date.today().isoformat()
    date_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        # The default headless backend (chrome-headless-shell) makes Zerodha's
        # funds/statement page error out to chrome-error://chromewebdata/ (blank
        # page, CSV export never renders). Using the full Chromium build via
        # channel="chromium" runs the modern --headless=new mode, which behaves
        # like a real browser and loads the funds page correctly.
        launch_kwargs = dict(
            headless=headless,
            accept_downloads=True,
        )
        if headless:
            launch_kwargs["channel"] = "chromium"
        context = p.chromium.launch_persistent_context(
            str(USER_DATA_DIR),
            **launch_kwargs,
        )
        page = context.pages[0] if context.pages else context.new_page()
        try:
            print(f"Logging in as {user_id} (headless={headless})")
            login(page, user_id, password, totp_secret)

            download_tradebook(page, date_dir / f"tradebook-{client_id}.csv")
            download_ledger(page, "All segments", date_dir / f"all_ledger-{client_id}.csv")
            download_ledger(page, "MTF", date_dir / f"mtf_ledger-{client_id}.csv")
        except Exception:
            write_error(date_dir, page)
            raise
        finally:
            context.close()

    print(f"DOWNLOAD_DIR={date_dir}")
    print(f"Done at {datetime.now().isoformat()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
