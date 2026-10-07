"""Run the SharePoint x Copilot Studio permission matrix for one persona.

usage: python run_matrix.py <mario|a|b|c|d> [ABC]
Writes <OUTPUT_DIR>/results/<persona>.json and per-web-part screenshots in <OUTPUT_DIR>/shots.
"""
import json
import os
import re
import sys
import time

from cdp import session
from chat import WP, bot_texts, send, state, wait_ready
from settings import EXPECT_A, PAGE_NAME, PERSONAS, QUESTION_A, QUESTION_B, RESULTS_DIR, SHAREPOINT_ROOT, SHOTS_DIR, SITES
from signin import sign_in

UNSUPPORTED = "doesn't support agents built with the GitHub Copilot harness"
NO_ACCESS = "You don't have access to talk to this bot"


def _body(page):
    try:
        return re.sub(r"\s+", " ", page.inner_text("body"))
    except Exception:  # noqa: BLE001
        return ""


def open_site(page, who, letter):
    url = f"{SHAREPOINT_ROOT}{SITES[letter]}/SitePages/{PAGE_NAME}"
    page.goto(url, wait_until="domcontentloaded")
    end = time.time() + 60
    while time.time() < end:
        page.wait_for_timeout(2000)
        if who != "mario" and "login.microsoftonline.com" in page.url:
            sign_in(page, who)
        if page.locator(WP).count() >= 3:
            return "ok", url
        if "AccessDenied" in page.url or re.search(r"You need access|Access required|Request access", _body(page), re.I):
            return "denied", page.url
    return "unknown", page.url


def sdk_verdict(greeting, reply, status, st):
    joined = " ".join(greeting + reply)
    if UNSUPPORTED in joined:
        return "NOT SUPPORTED"
    if NO_ACCESS in joined:
        return "BLOCKED - not shared"
    if status == "needs-sign-in":
        return "SIGN-IN PROMPT"
    if status in ("error", "timeout") or (st and st.get("banner")):
        return "ERROR"
    if any(EXPECT_A in r for r in reply):
        return "WORKS"
    return "UNCLEAR"


def test_sdk(page, idx, question, tag):
    status, st = wait_ready(page, idx)
    greeting = bot_texts(st)
    reply = send(page, idx, question) if status == "chat" else []
    st = state(page, idx)
    page.locator(WP).nth(idx).screenshot(path=os.path.join(SHOTS_DIR, f"{tag}.png"))
    return {"status": status, "greeting": greeting[:1], "reply": [r[:400] for r in reply],
            "banner": (st or {}).get("banner", ""), "footer": (st or {}).get("footer", ""),
            "verdict": sdk_verdict(greeting, reply, status, st)}


def _frame_text(fr):
    try:
        return re.sub(r"\s+", " ", fr.locator("body").inner_text())
    except Exception:  # noqa: BLE001
        return ""


def test_embed(page, tag):
    wp = page.locator(WP).nth(2)
    wp.scroll_into_view_if_needed()
    fr, text = None, ""
    end = time.time() + 90
    while time.time() < end:
        page.wait_for_timeout(2000)
        frames = [f for f in page.frames if "copilotstudio" in f.url]
        if frames:
            fr, text = frames[0], _frame_text(frames[0])
            if " said:" in text:
                break
    reply = ""
    if fr and " said:" in text and fr.locator("textarea, input[type=text]").count():
        before = len(text)
        box = fr.locator("textarea, input[type=text]").first
        box.fill(QUESTION_B)
        box.press("Enter")
        last, stable, end = "", 0, time.time() + 150
        while time.time() < end:
            time.sleep(3)
            t = _frame_text(fr)
            stable = stable + 1 if (t == last and len(t) > before + 100) else 0
            last = t
            if stable >= 2:
                break
        reply = last[before:][:600]
    wp.screenshot(path=os.path.join(SHOTS_DIR, f"{tag}.png"))
    if fr is None:
        verdict = "NO EMBED FRAME"
    elif re.search(r"°F|forecast", reply, re.I):
        verdict = "ANSWERS (no auth enforced)"
    else:
        verdict = "CONNECTED, NO ANSWER" if " said:" in text else "BLOCKED"
    return {"initial": text[:400], "reply": reply, "verdict": verdict}


def main():
    who = sys.argv[1].lower()
    letters = sys.argv[2] if len(sys.argv) > 2 else "".join(SITES)
    os.makedirs(SHOTS_DIR, exist_ok=True)
    os.makedirs(RESULTS_DIR, exist_ok=True)
    results = {"persona": who, "started": time.strftime("%Y-%m-%d %H:%M:%S"), "sites": {}}
    with session(new=True, port=PERSONAS[who]) as (b, ctx, page):
        page.set_viewport_size({"width": 1500, "height": 950})
        for letter in letters:
            access, url = open_site(page, who, letter)
            entry = {"access": access, "url": url}
            print(f"[{who}] Site {letter}: {access}")
            if access == "ok":
                page.wait_for_timeout(3000)
                entry["agentA_sdk"] = test_sdk(page, 0, QUESTION_A, f"{who}-site{letter}-agentA-sdk")
                entry["agentB_sdk"] = test_sdk(page, 1, QUESTION_B, f"{who}-site{letter}-agentB-sdk")
                entry["agentB_embed"] = test_embed(page, f"{who}-site{letter}-agentB-embed")
                for k in ("agentA_sdk", "agentB_sdk", "agentB_embed"):
                    print(f"   {k}: {entry[k]['verdict']}")
            else:
                page.screenshot(path=os.path.join(SHOTS_DIR, f"{who}-site{letter}-denied.png"))
            results["sites"][letter] = entry
        page.close()
    path = os.path.join(RESULTS_DIR, f"{who}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=1, ensure_ascii=False)
    print("saved", path)


if __name__ == "__main__":
    main()
