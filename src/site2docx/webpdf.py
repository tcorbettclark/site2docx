"""Web page -> PDF using headless system Chrome (via Playwright).

The Ada Computer Science site is a React SPA (empty shell HTML), so a real
browser is mandatory. Steps:

  1. load the page and wait for the React app to settle
  2. dismiss any cookie banner
  3. scroll through the page to trigger lazy-loaded content
  4. click every collapsed-section toggle (React state), then force open
     anything left hidden by direct class surgery, and verify
  5. print to PDF with Chrome's default headers/footers (date/title header,
     URL + "N/M" footer), backgrounds on, A4
"""

import re
import time

from playwright.sync_api import sync_playwright

# Elements that look like collapsed-section toggles and are safe to click.
TOGGLE_SELECTOR = '[aria-expanded="false"], button.collapsed, [class*="btn-collapsed"]'

# Chrome's default print margins: header/footer live in the top/bottom ones.
PRINT_MARGINS = {"top": "0.4in", "bottom": "0.56in", "left": "0.4in", "right": "0.4in"}


def _log(msg):
    print(msg, file=__import__("sys").stderr)


def _wait_stable(page, timeout_s=30.0):
    """Wait until the React app stops changing the DOM (text length + height)."""
    last_len, last_h, stable = -1, -1, 0
    t0 = time.monotonic()
    while time.monotonic() - t0 < timeout_s:
        try:
            txt_len = page.evaluate("document.body.innerText.length")
            height = page.evaluate("document.body.scrollHeight")
        except Exception:
            txt_len = height = 0
        if txt_len == last_len and height == last_h:
            stable += 1
            if stable >= 3:
                _log(f"  content stable after {time.monotonic() - t0:.1f}s "
                     f"({txt_len:,} chars, page height {height:,} px)")
                return txt_len, height
        else:
            stable = 0
        last_len, last_h = txt_len, height
        page.wait_for_timeout(400)
    _log(f"  warning: content never fully stabilised "
         f"({last_len:,} chars, height {last_h:,} px)")
    return last_len, last_h


def _dismiss_cookie_banner(page):
    """Click the first visible accept/allow/agree button, if one exists."""
    pattern = re.compile(r"accept|allow all|allow|agree|consent", re.I)
    for btn in page.query_selector_all("button, [role='button']"):
        try:
            text = btn.inner_text().strip()
            if text and pattern.search(text) and btn.is_visible():
                btn.click()
                _log(f"  dismissed banner by clicking '{text}'")
                page.wait_for_timeout(500)
                return
        except Exception:
            continue


def _scroll_through(page):
    """Scroll to the bottom and back to trigger lazy-loaded images."""
    page.evaluate("""async () => {
        const h = document.body.scrollHeight;
        for (let y = 0; y < h; y += window.innerHeight) {
            window.scrollTo(0, y);
            await new Promise(r => setTimeout(r, 150));
        }
        window.scrollTo(0, 0);
    }""")
    _log("  scrolled through page to trigger lazy-loaded content")


def _expand_all(page):
    """Click every collapsed toggle, then force-open leftovers, then verify."""
    stats = {"toggles_clicked": 0, "forced_open": 0, "rounds": 0}
    for round_no in range(1, 6):
        toggles = page.query_selector_all(TOGGLE_SELECTOR)
        if not toggles:
            break
        clicked = 0
        for t in toggles:
            try:
                if t.is_visible():
                    t.click(timeout=3000)
                    clicked += 1
            except Exception:
                continue
        stats["rounds"] = round_no
        stats["toggles_clicked"] += clicked
        _log(f"  expansion round {round_no}: clicked {clicked} toggle(s)")
        page.wait_for_timeout(400)
        if clicked == 0:
            break

    # Force open anything React did not open itself (class surgery + verify).
    stats["forced_open"] = page.evaluate("""() => {
        let n = 0;
        document.querySelectorAll('.collapse:not(.show)').forEach(el => {
            el.classList.add('show');
            el.style.removeProperty('display');
            el.style.removeProperty('height');
            el.style.removeProperty('overflow');
            n++;
        });
        return n;
    }""")
    if stats["forced_open"]:
        _log(f"  forced open {stats['forced_open']} collapsed element(s) via class surgery")
    hidden = page.evaluate("document.querySelectorAll('.collapse:not(.show)').length")
    stats["remaining_hidden"] = hidden
    _log(f"  collapsed sections remaining hidden: {hidden}")
    return stats


def render(url, pdf_path):
    """Render a URL to a PDF file. Returns (page_title, stats dict)."""
    stats = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page()
        try:
            _log(f"browser: launching headless system Chrome")
            _log(f"navigating: {url}")
            t0 = time.monotonic()
            page.goto(url, wait_until="load", timeout=90_000)
            _log(f"  page loaded in {time.monotonic() - t0:.1f}s")

            chars, height = _wait_stable(page)
            stats["text_chars"], stats["page_height_px"] = chars, height

            _dismiss_cookie_banner(page)
            _scroll_through(page)

            _log("expanding collapsed sections...")
            t0 = time.monotonic()
            stats.update(_expand_all(page))
            _log(f"  expansion finished in {time.monotonic() - t0:.1f}s "
                 f"({stats.get('toggles_clicked', 0)} toggles clicked, "
                 f"{stats.get('forced_open', 0)} forced open)")

            _wait_stable(page, timeout_s=15.0)

            title = page.title()
            stats["title"] = title
            _log(f"page title: {title}")

            _log("printing to PDF (A4, backgrounds on, Chrome default headers/footers)...")
            t0 = time.monotonic()
            page.pdf(
                path=str(pdf_path),
                format="A4",
                print_background=True,
                display_header_footer=True,  # Chrome's default templates
                margin=PRINT_MARGINS,
            )
            _log(f"  printed in {time.monotonic() - t0:.1f}s -> {pdf_path}")
        finally:
            browser.close()
    return stats.get("title", url), stats