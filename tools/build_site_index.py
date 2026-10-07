"""Build docs/index.html - landing page of the GitHub Pages site (source: main branch, docs/ folder) - and docs/.nojekyll.

    python tools/build_site_index.py
.nojekyll makes GitHub Pages serve docs/ as static files (no Jekyll processing of the markdown files).
"""
import html
import os

import build_usecase_html as base
import sdk_content as c

DOCS = os.path.join(c.ROOT, "docs")
OUT = os.path.join(DOCS, "index.html")

CSS = """
.hero{padding:44px 0 8px}
.cards{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px;margin-top:22px}
@media(max-width:900px){.cards{grid-template-columns:1fr}}
a.tile{display:flex;flex-direction:column;gap:8px;text-decoration:none;color:var(--cp-text)}
a.tile:hover{border-color:var(--cp-accent)}
a.tile:hover .go{color:var(--cp-accent-hover)}
a.tile:focus-visible{outline:2px solid var(--cp-accent);outline-offset:2px}
.tile .kind{font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--cp-text-muted)}
.tile h3{font-size:19px;margin:0}
.tile p{margin:0;color:var(--cp-text-muted)}
.tile .go{margin-top:auto;padding-top:6px;font-weight:600;color:var(--cp-accent)}
.tile.primary{border-left:4px solid var(--cp-accent)}
.facts{display:flex;flex-wrap:wrap;gap:8px 22px;margin-top:18px;color:var(--cp-text-muted);font-size:14px}
.facts b{color:var(--cp-text)}
"""


def size(name):
    kb = os.path.getsize(os.path.join(DOCS, name)) / 1024
    return f"{kb / 1024:.1f}\u00a0MB" if kb >= 1024 else f"{kb:.0f}\u00a0KB"


def tile(href, kind, title, text, action, primary=False, download=False):
    attrs = ' download' if download else ''
    return (f'<a class="card tile{" primary" if primary else ""}" href="{href}"{attrs}><span class="kind">{kind}</span>'
            f'<h3>{html.escape(title)}</h3><p>{text}</p><span class="go">{action} &rarr;</span></a>')


def build():
    k = c.counts()
    validated = sum(u["status"] == "validated" for u in c.USE_CASES)
    blob = c.REPO_URL + "/blob/main/docs/"
    o = ['<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
         '<meta name="robots" content="noindex">',
         "<title>Copilot Studio agents on SharePoint with the Microsoft 365 Agents SDK</title>",
         '<script>(() => {const param = new URLSearchParams(window.location.search).get("scoutTheme");'
         'const theme = param || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");'
         'document.documentElement.setAttribute("data-theme", theme);})();</script>',
         f"<style>{base.CSS}{CSS}</style></head><body>",
         '<header class="top"><div class="topin"><div class="brand">Agents SDK &times; SharePoint</div><nav class="sections"></nav>'
         '<button class="iconbtn" id="themeToggle" title="Toggle light/dark">&#9680; Theme</button></div></header><main>',
         f'<section class="hero"><div class="eyebrow">Copilot Studio &times; SharePoint &middot; validated {c.TESTED_ON}</div>'
         '<h1>Authenticated Copilot Studio agents on any SharePoint site</h1>'
         '<p class="lead">How to publish a Copilot Studio agent to many SharePoint sites through the <b>Microsoft 365 Agents SDK</b> - one web '
         'part for the whole tenant, single sign-on with Microsoft Entra ID and Copilot Studio sharing as the access control - built and '
         'tested end to end in a live Microsoft demo tenant, and compared with the SharePoint channel + Agent link pattern.</p>'
         f'<div class="facts"><span><b>{len(c.USE_CASES)}</b> use cases ({validated} validated)</span>'
         f'<span><b>{len(c.TESTS)}</b> test scenarios</span><span><b>{k["pass"]}</b> passed in the demo tenant</span>'
         f'<span><b>{k["notrun"]}</b> to validate in your tenant</span><span>Web part <b>{c.PACKAGE_VERSION}</b></span></div>',
         '<div class="cards">',
         tile("Agents-SDK-SharePoint-Tutorial-and-Test-Scenarios.html", "Web page", "Tutorial, use cases and test scenarios",
              f"Build it in {sum(len(p['steps']) for p in c.TUTORIAL)} steps with checkpoints and screenshots, see what it enables, and "
              "review every test scenario with expected and actual results.", "Open the page", primary=True),
         tile("Agents-SDK-SharePoint-Use-Case.html", "Web page", "Customer pattern: SharePoint channel + Agent link",
              "The customer's setup reproduced in the same tenant: what works, the Site A dependency, concerns and recommendations - "
              "compared head-to-head with the Agents SDK web part.", "Open the page"),
         tile("Agents-SDK-SharePoint-Test-Plan.xlsx", "Download", "Test plan (Excel)",
              f"All {len(c.TESTS)} scenarios with preconditions, steps and expected results. Record your own results; the summary "
              f"sheet counts passes, failures and the P1 pass rate. ({size('Agents-SDK-SharePoint-Test-Plan.xlsx')})",
              "Download", download=True),
         tile("Copilot-Studio-Agents-SDK-SharePoint.pptx", "Download", "Presentation (PowerPoint)",
              f"16 slides: the use case, the answer at a glance, both patterns, test results and recommendations. "
              f"({size('Copilot-Studio-Agents-SDK-SharePoint.pptx')})", "Download", download=True),
         "</div>",
         '<div class="grid g2" style="margin-top:16px">'
         f'<div class="card"><h3>Source code and guide</h3><p><a href="{c.REPO_URL}">{html.escape(c.REPO_URL)}</a></p><ul class="checks">'
         f'<li><a href="{blob}Agents-SDK-SharePoint-Guide.md">Guide</a> - variables, findings, troubleshooting</li>'
         f'<li><a href="{blob}Agents-SDK-SharePoint-Tutorial.md">Tutorial</a> and '
         f'<a href="{blob}Agents-SDK-SharePoint-Use-Cases-and-Test-Scenarios.md">use cases and test scenarios</a> as markdown</li>'
         '<li>SPFx web part, setup scripts and the Edge + Playwright test harness</li></ul></div>'
         '<div class="card"><h3>About the test environment</h3><p>Results come from automated runs in a Microsoft demo tenant with four test users '
         '(Microsoft 365 E5, no Copilot licenses, MFA), three SharePoint sites and two agents - one on the standard harness and one on the '
         'GitHub Copilot harness. Tenant identifiers are demo values; no credentials are included.</p></div></div>',
         f'<footer>Microsoft 365 Agents SDK in SharePoint &middot; validated {c.TESTED_ON}. Preview features may change.</footer>',
         "</section></main>",
         '<script>document.getElementById("themeToggle").addEventListener("click",function(){var r=document.documentElement;'
         'r.setAttribute("data-theme",r.getAttribute("data-theme")==="dark"?"light":"dark");});</script></body></html>']
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("".join(o))
    with open(os.path.join(DOCS, ".nojekyll"), "w", encoding="utf-8"):
        pass
    print(f"wrote {OUT} and docs/.nojekyll")


if __name__ == "__main__":
    build()
