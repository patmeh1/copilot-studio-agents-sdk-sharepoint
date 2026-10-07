"""Extra Agents SDK scenarios for one persona on one site page (Agent A web part = first web part on the page).

TC-03  Interactive sign-in fallback. The hidden-iframe SSO (MSAL ssoSilent) is blocked - the same effect as a browser
       or host that blocks third-party cookies, or a Conditional Access policy that needs interaction. The MSAL cache
       is cleared first so no cached/refresh token can hide the problem. Expected: the web part shows
       "Sign in to chat with this agent", the "Sign in with Microsoft" popup completes and the chat connects.
TC-22  "New chat" starts a fresh conversation: the transcript is cleared, the greeting is shown again and the agent
       answers a new question.

usage: python run_extra_scenarios.py <a|b|c|d|mario> [A|B|C]
Writes <OUTPUT_DIR>/results/extra-<persona>-site<X>.json and screenshots in <OUTPUT_DIR>/shots.
"""
import json
import os
import re
import sys
import time

from cdp import session
from chat import WP, bot_texts, send, state, wait_ready
from run_matrix import open_site
from settings import EXPECT_A, PERSONAS, QUESTION_A, RESULTS_DIR, SHOTS_DIR
from signin import load_users, step

LOGIN = re.compile(r"^https://login\.microsoftonline\.com/")
SIGN_IN_PROMPT = "Sign in to chat with this agent"

CLEAR_MSAL_CACHE = """() => {
  let removed = 0;
  for (const k of Object.keys(localStorage)) {
    if (/msal|login\\.windows\\.net|login\\.microsoftonline\\.com/i.test(k)) { localStorage.removeItem(k); removed++; }
  }
  return removed;
}"""


def block_silent_sso(page):
    """Abort Entra sign-in pages loaded inside iframes (MSAL ssoSilent). Top-level popups are not affected."""
    def handler(route):
        req = route.request
        if req.resource_type == "document" and req.frame != page.main_frame:
            route.abort()
        else:
            route.continue_()
    page.route(LOGIN, handler)
    return handler


def wait_for_text(page, idx, pattern, timeout=90):
    end = time.time() + timeout
    while time.time() < end:
        text = (state(page, idx) or {}).get("text", "")
        if re.search(pattern, text):
            return text
        time.sleep(1)
    return None


def complete_popup(popup, who, timeout=90):
    """Drive the Microsoft sign-in popup until MSAL closes it."""
    user = load_users()[who]
    steps, end = [], time.time() + timeout
    while time.time() < end and not popup.is_closed():
        try:
            popup.wait_for_timeout(2000)
            what = step(popup, who, user)
        except Exception:  # noqa: BLE001 - MSAL closed the popup while it was inspected
            what = "closed"
        if not steps or steps[-1] != what:
            steps.append(what)
        if what == "closed":
            break
    return steps


def loaded_bundle(page):
    """File name of the web part bundle the page loaded (shows which package version was tested)."""
    return page.evaluate("""() => performance.getEntriesByType('resource').map(e => e.name)
        .filter(n => /copilot-studio-agent-web-part_[0-9a-f]+\\.js/.test(n)).map(n => n.split('/').pop())""")


def interactive_fallback(page, who, tag):
    wp = page.locator(WP).nth(0)
    removed = page.evaluate(CLEAR_MSAL_CACHE)
    handler = block_silent_sso(page)
    page.reload(wait_until="domcontentloaded")
    page.wait_for_selector(WP, timeout=60000)
    prompt = wait_for_text(page, 0, SIGN_IN_PROMPT)
    wp.scroll_into_view_if_needed()
    wp.screenshot(path=os.path.join(SHOTS_DIR, f"{tag}-tc03-sign-in-prompt.png"))
    page.unroute(LOGIN, handler)
    result = {"bundle": loaded_bundle(page), "msal_cache_entries_removed": removed, "prompt_shown": bool(prompt),
              "prompt": re.sub(r"\s+", " ", prompt or "")[:300]}
    if not prompt:
        result["verdict"] = "FAIL - no sign-in prompt"
        return result
    started = time.time()
    with page.expect_popup(timeout=20000) as popup_info:
        wp.get_by_role("button", name="Sign in with Microsoft").click()
    popup = popup_info.value
    result["popup_host"] = re.sub(r"^(https://[^/]+).*", r"\1", popup.url or "")
    result["popup_steps"] = complete_popup(popup, who)
    status, st = wait_ready(page, 0, timeout=90)
    greeting = bot_texts(st)
    reply = send(page, 0, QUESTION_A) if status == "chat" else []
    result.update(status=status, greeting=greeting[:1], reply=[r[:300] for r in reply],
                  seconds_to_answer=round(time.time() - started, 1), footer=(state(page, 0) or {}).get("footer", ""))
    wp.screenshot(path=os.path.join(SHOTS_DIR, f"{tag}-tc03-after-popup.png"))
    # v1.0.2+: the other Agents SDK web part on the page retries silently after this sign-in.
    other = None
    if page.locator(WP).count() > 1:
        st_other, end = None, time.time() + 45
        while time.time() < end:
            st_other = state(page, 1) or {}
            if SIGN_IN_PROMPT not in st_other.get("text", "") and bot_texts(st_other):
                break
            time.sleep(1.5)
        other = {"still_prompting": SIGN_IN_PROMPT in (st_other or {}).get("text", ""),
                 "first_message": (bot_texts(st_other) or [""])[0][:200]}
        other["verdict"] = "FAIL" if other["still_prompting"] or not other["first_message"] else "PASS"
    result["other_web_part"] = other
    result["verdict"] = "PASS" if any(EXPECT_A in r for r in reply) else "FAIL"
    return result


def new_chat(page, tag):
    wp = page.locator(WP).nth(0)
    before = state(page, 0) or {}
    asked_before = QUESTION_A in before.get("text", "")
    wp.get_by_role("button", name="New chat").click()
    st, end = None, time.time() + 60
    while time.time() < end:
        time.sleep(1.5)
        st = state(page, 0)
        if st and QUESTION_A not in st["text"] and bot_texts(st) and not st["typing"]:
            break
    cleared = bool(st) and QUESTION_A not in st["text"]
    greeting = bot_texts(st)
    wp.screenshot(path=os.path.join(SHOTS_DIR, f"{tag}-tc22-new-chat.png"))
    reply = send(page, 0, QUESTION_A)
    ok = asked_before and cleared and bool(greeting) and any(EXPECT_A in r for r in reply)
    return {"question_in_transcript_before": asked_before, "transcript_cleared": cleared, "greeting_after": greeting[:1],
            "reply_after": [r[:300] for r in reply], "verdict": "PASS" if ok else "FAIL"}


def main():
    who = sys.argv[1].lower()
    letter = (sys.argv[2] if len(sys.argv) > 2 else "A").upper()
    os.makedirs(SHOTS_DIR, exist_ok=True)
    os.makedirs(RESULTS_DIR, exist_ok=True)
    tag = f"{who}-site{letter}"
    out = {"persona": who, "site": letter, "started": time.strftime("%Y-%m-%d %H:%M:%S")}
    logs = []
    with session(new=True, port=PERSONAS[who]) as (b, ctx, page):
        page.set_viewport_size({"width": 1500, "height": 950})
        page.on("console", lambda m: logs.append(m.text[:300]) if "CopilotAgent" in m.text else None)
        out["access"], out["url"] = open_site(page, who, letter)
        if out["access"] == "ok":
            page.wait_for_timeout(3000)
            out["tc03_interactive_fallback"] = interactive_fallback(page, who, tag)
            print("TC-03:", out["tc03_interactive_fallback"]["verdict"])
            other = out["tc03_interactive_fallback"].get("other_web_part")
            if other:
                print("TC-03 (other web part on the page):", other["verdict"])
            if out["tc03_interactive_fallback"].get("status") == "chat":
                out["tc22_new_chat"] = new_chat(page, tag)
                print("TC-22:", out["tc22_new_chat"]["verdict"])
        page.close()
    out["console"] = logs[:20]
    path = os.path.join(RESULTS_DIR, f"extra-{tag}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=1, ensure_ascii=False)
    print("saved", path)


if __name__ == "__main__":
    main()
