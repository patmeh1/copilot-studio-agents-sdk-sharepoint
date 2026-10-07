"""Connect to an Edge instance started with --remote-debugging-port and reuse its signed-in session."""
import contextlib
import os

from playwright.sync_api import sync_playwright

from settings import SHOTS_DIR


@contextlib.contextmanager
def session(url_contains=None, new=False, port=9333):
    pw = sync_playwright().start()
    try:
        browser = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
        ctx = browser.contexts[0]
        page = None
        if new:
            page = ctx.new_page()
        elif url_contains:
            for p in ctx.pages:
                if url_contains.lower() in p.url.lower():
                    page = p
        if page is None:
            pages = [p for p in ctx.pages if not p.url.startswith(("edge://", "devtools://"))]
            page = pages[-1] if pages else ctx.new_page()
        page.bring_to_front()
        yield browser, ctx, page
    finally:
        pw.stop()


def shot(page, name, full=False):
    os.makedirs(SHOTS_DIR, exist_ok=True)
    path = os.path.join(SHOTS_DIR, f"{name}.png")
    page.screenshot(path=path, full_page=full)
    return path
