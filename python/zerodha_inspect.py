"""One-off helper: dump the real controls on the console report pages so the
selectors in zerodha_download.py can be made exact.

Reuses the saved browser profile (.zerodha-browser); if the session is still
valid it won't re-login. Run headed once:  HEADLESS=0 python3 python/zerodha_inspect.py
Paste the output back.
"""

from __future__ import annotations

import os

from playwright.sync_api import sync_playwright

import zerodha_download as z


def dump_page(page, label, url):
    print(f"\n========== {label}: {url} ==========")
    z.goto_console(page, url)
    print("FINAL URL:", page.url)

    print("\n--- <select> elements + options ---")
    for s in page.locator("select").all():
        name = s.get_attribute("name") or s.get_attribute("id") or "?"
        opts = [o.inner_text().strip() for o in s.locator("option").all()]
        print(f"select[{name}] options={opts}")

    print("\n--- elements whose text mentions segment / MTF / All ---")
    for txt in ["All segments", "MTF", "All", "Equity"]:
        n = page.get_by_text(txt, exact=False).count()
        if n:
            print(f"text '{txt}': {n} match(es)")

    print("\n--- buttons / links with Download or CSV ---")
    for sel in ["button", "a"]:
        for el in page.locator(sel).all():
            try:
                t = (el.inner_text() or "").strip()
            except Exception:
                t = ""
            if t and ("download" in t.lower() or "csv" in t.lower()):
                cls = el.get_attribute("class") or ""
                href = el.get_attribute("href") or ""
                print(f"{sel} text={t!r} class={cls!r} href={href!r}")

    print("\n--- input fields (date pickers etc.) ---")
    for el in page.locator("input").all():
        attrs = {a: el.get_attribute(a) for a in ("name", "id", "type", "placeholder")}
        print(attrs)

    # If a date-range picker exists, open it and dump the preset choices.
    picker = page.locator('input[placeholder="Select range"], input[name="date"]').first
    if picker.count() if hasattr(picker, "count") else 0:
        try:
            picker.click(timeout=2000)
            page.wait_for_timeout(800)
            print("\n--- date picker contents (visible clickable text) ---")
            seen = set()
            for sel in ["button", "a", "li", "span", "div"]:
                for el in page.locator(sel).all():
                    try:
                        if not el.is_visible():
                            continue
                        t = (el.inner_text() or "").strip()
                    except Exception:
                        continue
                    if t and len(t) <= 30 and t not in seen:
                        seen.add(t)
                        print(f"{sel}: {t!r}")
        except Exception as e:
            print("could not open picker:", e)


def main():
    os.environ.setdefault("HEADLESS", "0")
    z.load_dotenv(z.ENV_FILE)
    user_id = z.env("ZERODHA_USER_ID")
    password = z.env("ZERODHA_PASSWORD")
    totp_secret = z.env("ZERODHA_TOTP_SECRET")
    headless = z.env("HEADLESS", "0") not in ("0", "false", "False", "no")

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(z.USER_DATA_DIR), headless=headless, accept_downloads=True
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            # Always start from a clean Kite login (session may have expired).
            z.login(page, user_id, password, totp_secret)
            dump_page(page, "FUNDS/LEDGER", z.FUNDS_URL)
            dump_page(page, "TRADEBOOK", z.TRADEBOOK_URL)
        finally:
            ctx.close()


if __name__ == "__main__":
    raise SystemExit(main())
