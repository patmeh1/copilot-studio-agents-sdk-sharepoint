"""Sign a test persona in (dedicated Edge profile), including first-run MFA registration with a software TOTP.

usage: python signin.py <a|b|c|d> [start_url|-]
Credentials: <SECRETS_DIR>/test-users.json   [{"persona": "User A", "upn": "...", "password": "..."}, ...]
TOTP seeds:  <SECRETS_DIR>/test-users-totp.json  (written on first registration)
"""
import json
import os
import re
import sys
import time

import pyotp

from cdp import session, shot
from settings import PERSONAS, SECRETS_DIR, SHAREPOINT_ROOT, SITES


def load_users():
    with open(os.path.join(SECRETS_DIR, "test-users.json"), encoding="utf-8-sig") as f:
        data = json.load(f)
    return {u["persona"][-1].lower(): u for u in data}


def _totp_path():
    return os.path.join(SECRETS_DIR, "test-users-totp.json")


def save_secret(letter, secret):
    data = {}
    if os.path.exists(_totp_path()):
        with open(_totp_path(), encoding="utf-8") as f:
            data = json.load(f)
    data[letter] = secret
    with open(_totp_path(), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1)


def load_secret(letter):
    if not os.path.exists(_totp_path()):
        return None
    with open(_totp_path(), encoding="utf-8") as f:
        return json.load(f).get(letter)


def _visible(page, selector):
    loc = page.locator(selector)
    return loc.count() > 0 and loc.first.is_visible()


def _text(page):
    try:
        return page.inner_text("body")
    except Exception:  # noqa: BLE001
        return ""


def step(page, letter, user):
    """Advance the Microsoft login flow by one screen; returns what was done."""
    url, text = page.url, _text(page)
    if not any(h in url for h in ("login.microsoftonline.com", "login.microsoft.com", "mysignins")):
        return "done"
    if "Pick an account" in text:
        tile = page.get_by_role("button", name=re.compile(re.escape(user["upn"]), re.I))
        (tile if tile.count() else page.get_by_role("button", name=re.compile("Use another account"))).first.click()
        return "pick-account"
    if "Enter password" in text and _visible(page, "input[name=passwd]"):
        page.fill("input[name=passwd]", user["password"])
        page.click("input[type=submit]")
        return "password"
    if _visible(page, "input[name=loginfmt]"):
        page.fill("input[name=loginfmt]", user["upn"])
        page.click("input[type=submit]")
        return "email"
    if "Stay signed in" in text:
        page.click("input[type=submit]")
        return "kmsi"
    for sel in ("input[name=otc]", "input[id=idTxtBx_SAOTCC_OTC]"):
        if _visible(page, sel):
            secret = load_secret(letter)
            if not secret:
                return "need-secret"
            page.fill(sel, pyotp.TOTP(secret).now())
            page.keyboard.press("Enter")
            return "totp"
    if re.search(r"keep your account secure|More information required", text, re.I):
        return "mfa-registration"
    return "unknown"


def _click(page, role, name, wait=3000):
    loc = page.get_by_role(role, name=name)
    if loc.count() and loc.first.is_visible() and loc.first.is_enabled():
        loc.first.click()
        page.wait_for_timeout(wait)
        return True
    return False


def register_totp(page, letter):
    """Walk the 'Keep your account secure' wizard choosing 'a different authentication app'."""
    secret = None
    for _ in range(25):
        if "login.microsoftonline.com" in page.url:
            if re.search(r"keep your account secure", _text(page), re.I):
                _click(page, "button", "Next", 9000)
                continue
            return
        if "mysignins.microsoft.com" not in page.url:
            return
        h = page.locator("main h2, main h1")
        heading = h.first.inner_text().strip() if h.count() else ""
        if heading.startswith("Install Microsoft Authenticator"):
            _click(page, "link", "Set up a different authentication app")
        elif heading.startswith("Set up your account"):
            _click(page, "button", "Next")
        elif heading.startswith("Scan the QR code"):
            _click(page, "button", "Can't scan the QR code?")
            m = re.search(r"Secret key:?\s*([A-Za-z2-7 ]{16,})", page.inner_text("main"))
            secret = m.group(1).replace(" ", "").strip()
            save_secret(letter, secret)
            _click(page, "button", "Next")
        elif heading.startswith("Enter the code"):
            if not secret:
                _click(page, "button", "Back")
                continue
            if 30 - (int(time.time()) % 30) < 5:
                time.sleep(6)
            page.get_by_role("textbox", name="Code").fill(pyotp.TOTP(secret).now())
            _click(page, "button", "Next", 6000)
        elif not (_click(page, "button", "Done", 9000) or _click(page, "button", "Next")):
            return


def sign_in(page, letter, start=None):
    user = load_users()[letter]
    if start:
        page.goto(start, wait_until="domcontentloaded")
    for _ in range(16):
        page.wait_for_timeout(3500)
        what = step(page, letter, user)
        if what == "mfa-registration":
            register_totp(page, letter)
            continue
        if what in ("done", "unknown", "need-secret"):
            return what
    return "incomplete"


if __name__ == "__main__":
    who = sys.argv[1].lower()
    start = sys.argv[2] if len(sys.argv) > 2 else SHAREPOINT_ROOT + SITES["A"]
    with session(port=PERSONAS[who]) as (b, ctx, page):
        print("result:", sign_in(page, who, None if start == "-" else start), "|", page.url[:120])
        shot(page, f"signin-{who}")
