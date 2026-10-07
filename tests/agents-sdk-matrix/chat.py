"""Helpers to drive the 'Copilot Studio Agent (Agents SDK)' web part (Bot Framework Web Chat) in a SharePoint page."""
import re
import time

WP = "[data-sp-feature-tag*='CopilotStudioAgent']"

JS_STATE = """(idx) => {
  const wp = document.querySelectorAll("[data-sp-feature-tag*='CopilotStudioAgent']")[idx];
  if (!wp) return null;
  const acts = [...wp.querySelectorAll('.webchat__basic-transcript__activity')].map(a => {
    const t = (a.innerText || '').replace(/\\s+/g, ' ').trim();
    return { user: !!a.querySelector('.webchat__bubble--from-user') || / You said: /.test(' ' + t), text: t };
  });
  const typing = !!wp.querySelector('.webchat__typing-indicator, .webchat__typingIndicator');
  const banner = [...wp.querySelectorAll('[role=alert]')].map(e => e.innerText.trim()).filter(Boolean).join(' | ');
  const conn = [...wp.querySelectorAll('.webchat__connectivityStatus, [class*=connectivity]')].map(e => e.innerText.trim()).filter(Boolean).join(' | ');
  return {
    text: (wp.innerText || '').replace(/\\s+/g, ' ').trim(),
    acts, typing, banner, conn,
    hasSendBox: !!wp.querySelector('textarea, input[type=text]'),
    footer: ((wp.querySelector('footer') || {}).innerText || '').trim()
  };
}"""


def state(page, idx):
    return page.evaluate(JS_STATE, idx)


def clean(text):
    """Turn a Web Chat activity's innerText into just the message."""
    t = re.sub(r"\s+", " ", text or "").strip()
    m = re.search(r"said:\s*(.*?)(?:\s+(?:Like Dislike|Sent at|Message is interactive|One or more links)|$)", t)
    return (m.group(1) if m else t).replace("\u200b", "").strip()


def bot_texts(st):
    return [clean(a["text"]) for a in (st or {}).get("acts", []) if not a["user"]]


def wait_ready(page, idx, timeout=75):
    """Wait until Web Chat shows a bot activity, or a terminal error/sign-in state."""
    end = time.time() + timeout
    st = None
    while time.time() < end:
        st = state(page, idx)
        if st:
            t = st["text"]
            if bot_texts(st) and not st["typing"]:
                return "chat", st
            if "Sign in to chat" in t:
                return "needs-sign-in", st
            if st["banner"]:
                return "error", st
            if "Try again" in t and "Signing you in" not in t and "Connecting" not in t:
                return "error", st
        time.sleep(1.5)
    return "timeout", st


def send(page, idx, message, timeout=150):
    before = len(bot_texts(state(page, idx)))
    wp = page.locator(WP).nth(idx)
    box = wp.locator("textarea, input[type=text]").first
    box.click()
    box.fill(message)
    box.press("Enter")
    end = time.time() + timeout
    last, stable = None, 0
    while time.time() < end:
        time.sleep(2.5)
        st = state(page, idx)
        texts = bot_texts(st)
        if st and st["banner"] and len(texts) <= before:
            return [f"<error> {st['banner']}"]
        if len(texts) > before and not st["typing"]:
            cur = texts[-1]
            stable = stable + 1 if cur == last else 0
            last = cur
            if stable >= 2:
                return texts[before:]
    st = state(page, idx)
    return bot_texts(st)[before:] or [f"<no reply within {timeout}s> " + (st or {}).get('text', '')[-300:]]
