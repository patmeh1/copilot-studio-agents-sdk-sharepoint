"""Build docs/Agents-SDK-SharePoint-Use-Case.html - a single, self-contained (sandbox-safe) HTML page.

Screenshots are embedded as compressed WebP data URIs; results are embedded as JSON. No external resources,
no network calls, no storage. Re-run after updating screenshots or results:  python tools/build_usecase_html.py
"""
import base64
import html
import io
import json
import os

from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG = os.path.join(ROOT, "docs", "images")
RES = os.path.join(ROOT, "docs", "test-results")
OUT = os.path.join(ROOT, "docs", "Agents-SDK-SharePoint-Use-Case.html")
REPO_URL = os.environ.get("REPO_URL", "https://github.com/patmeh1/copilot-studio-agents-sdk-sharepoint")

_cache = {}

# Regions covered before a screenshot is embedded (original image coordinates). The App-channel embed code of the
# GitHub Copilot harness agent opens a chat without any sign-in (guide, section 10), so its URL is never published.
REDACT = {
    "04-app-channel-agents-sdk-connection-string.png": [(578, 391, 1280, 484, "Embed code removed - it opens the agent without sign-in")],
}


def _redact(im, name):
    from PIL import ImageDraw, ImageFont
    draw = ImageDraw.Draw(im)
    try:
        font = ImageFont.truetype("segoeui.ttf", 18)
    except OSError:
        font = ImageFont.load_default()
    for x1, y1, x2, y2, label in REDACT.get(name, []):
        draw.rectangle((x1, y1, x2, y2), fill=(236, 236, 236), outline=(190, 190, 190))
        draw.text(((x1 + x2 - draw.textlength(label, font=font)) / 2, (y1 + y2) / 2 - 12), label, fill=(80, 80, 80), font=font)
    return im


def img(name, max_w=1500, q=80, crop=None):
    key = (name, crop)
    if key not in _cache:
        im = _redact(Image.open(os.path.join(IMG, name)).convert("RGB"), name)
        if crop:
            im = im.crop(crop)
        if im.width > max_w:
            im = im.resize((max_w, int(im.height * max_w / im.width)), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, "WEBP", quality=q, method=6)
        _cache[key] = "data:image/webp;base64," + base64.b64encode(buf.getvalue()).decode()
    return _cache[key]


def figure(name, caption, cls="", crop=None):
    return (f'<figure class="shot {cls}"><img src="{img(name, crop=crop)}" alt="{html.escape(caption)}" '
            f'data-caption="{html.escape(caption)}"><figcaption>{html.escape(caption)}</figcaption></figure>')


def esc(s):
    return html.escape(s)


def load_results():
    with open(os.path.join(RES, "summary.json"), encoding="utf-8") as f:
        summary = json.load(f)
    labels = {
        "mario": ("MarioR (maker)", "Owner of A + B", {"A": "Owner", "B": "Owner", "C": "Owner"}),
        "a": ("User A", "Agent A + Agent B", {"A": "Visitor", "B": "Member", "C": "Owner"}),
        "b": ("User B", "Agent A only", {"A": "Member", "B": "Visitor", "C": "No access"}),
        "c": ("User C", "Agent B only", {"A": "No access", "B": "Member", "C": "Visitor"}),
        "d": ("User D", "No agent access", {"A": "Visitor", "B": "No access", "C": "Member"}),
    }
    rows = []
    for who in ("a", "b", "c", "d", "mario"):
        name, agents, perms = labels[who]
        for site in "ABC":
            cell = summary.get(who, {}).get(site)
            if not cell:
                continue
            denied = cell.get("site") != "OK"
            rows.append({
                "persona": name, "key": who, "agents": agents, "site": site, "perm": perms[site],
                "page": "ACCESS DENIED" if denied else "OK",
                "a_sdk": "-" if denied else cell["A_sdk"],
                "b_sdk": "-" if denied else cell["B_sdk"],
                "b_embed": "-" if denied else cell["B_embed"],
                "greeting": "" if denied else cell.get("A_greeting", ""),
            })
    return rows


STEPS = [
    ("Build a standard-harness agent with Microsoft authentication",
     "Copilot Studio home &gt; <b>Other ways to build</b> &gt; <b>Agent (Standard)</b>. Add instructions. Settings &gt; Security &gt; "
     "Authentication = <b>Authenticate with Microsoft</b> (default). Optional: greet the user with <code>{System.User.DisplayName}</code> "
     "in the Conversation Start topic to prove identity flows through. <b>Publish</b>.",
     [("07-create-standard-agent.png", "Home > Other ways to build > Agent (Standard)"),
      ("10-agentA-authenticate-with-microsoft.png", "Security > Authentication: Authenticate with Microsoft"),
      ("09-agentA-greeting-yaml.png", "Conversation Start greets {System.User.DisplayName}")], None),
    ("Copy the Microsoft 365 Agents SDK connection string",
     "Standard UI: <b>Channels &gt; Web app</b> (or Native app). New UI: right panel <b>Channels (+) &gt; App</b>. The string identifies the "
     "environment and agent schema - no secret inside.",
     [("11-agentA-web-app-connection-string.png", "Standard agent: Channels > Web app > connection string"),
      ("04-app-channel-agents-sdk-connection-string.png", "New UI: Channels > App > Microsoft 365 Agents SDK")],
     "https://afdeda3552fae1459bc07b43e75630.0a.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/new_AgentAStandardharness/conversations?api-version=2022-03-01-preview"),
    ("Share the agent with the people or security groups who may chat",
     "Agent &gt; <b>Options (...) &gt; Share</b> &gt; add users / groups with <b>End user access</b>. This list is enforced on every site "
     "(verified, including revoke and re-grant).",
     [("12-agentA-share.png", "Agent shared with User A and User B (Viewer / end-user access)")], None),
    ("Register one Entra SPA app for the whole tenant (scripted)",
     "Delegated <b>Power Platform API / CopilotStudio.Copilots.Invoke</b> + Graph basics, SPA redirect URI = the SharePoint origin "
     "(covers every site), tenant-wide consent. SPFx's built-in token provider cannot obtain this scope (sp-dev-docs #10775), so the "
     "web part uses MSAL.js with this app.",
     [], "az login --tenant $TenantDomain --allow-no-subscriptions\n.\\scripts\\New-AgentsSdkSpaApp.ps1 -TenantId $TenantDomain -SharePointRootUrl https://<tenant>.sharepoint.com\n# -> appClientId 1f06ba87-...  adminConsent: granted"),
    ("Build the SPFx package with tenant defaults (scripted)",
     "SPFx 1.21.1 + <code>@microsoft/agents-copilotstudio-client</code> 1.9.1 + MSAL 4.30 + Bot Framework Web Chat 4.19.1, all bundled "
     "(no CDN, no API-access approvals). Node 22 required.",
     [], ".\\scripts\\Build-SpfxPackage.ps1 -AppClientId <appClientId> -Version 1.0.1.0 -NodePath C:\\tools\\node-v22\n# -> spfx-copilot-agent\\sharepoint\\solution\\copilot-studio-agent-webpart.sppkg"),
    ("Deploy once, tenant-wide",
     "Add yourself as <b>site admin of the app catalog</b> (required even for SharePoint Administrators), upload the .sppkg, choose "
     "<b>Enable this app and add it to all sites</b>.",
     [("15-appcatalog-site-admin.png", "Admin center: add yourself as app catalog site admin"),
      ("16-enable-app.png", "Enable this app and add it to all sites"),
      ("17-app-enabled.png", "Enabled tenant-wide - available on all sites")], None),
    ("Add the web part to any page on any site",
     "Edit page &gt; <b>+</b> &gt; <b>Copilot Studio Agent (Agents SDK)</b> &gt; paste the connection string. App ID and tenant come from the "
     "package defaults. Repeat on as many sites as needed - or provision pages through the SharePoint pages REST API.",
     [("18-webpart-toolbox.png", "Toolbox: Copilot Studio Agent (Agents SDK)"),
      ("19-webpart-property-pane.png", "Property pane: paste the connection string"),
      ("20-site-page-overview.png", "Published page - Agent A works, Agent B (GitHub harness) rejected")], None),
    ("Grant normal SharePoint access",
     "Visitor (Read) on the hosting site is enough. No access to any other site is required.",
     [("25-site-access-denied.png", "Without site permission SharePoint blocks the page itself")], None),
    ("Validate",
     "Greeting shows the user's name, the footer shows the signed-in UPN, unshared users get <i>You don't have access to talk to this "
     "bot</i>, and an anonymous call to the connection string returns HTTP 401.",
     [("21-userA-agentA-sdk-works.png", "Shared user: personalised greeting + answer"),
      ("22-userD-agentA-sdk-not-shared.png", "Unshared user: blocked by Copilot Studio")], None),
]

CONCERNS = [
    ("high", "Single point of failure in Site A",
     "Every Agent link depends on the <code>.agent</code> file in the deployment site. Archiving/deleting Site A, changing its permissions, "
     "or a Site A member with Edit rights deleting or renaming the file breaks the agent on every site. Host it in a dedicated, locked-down "
     "&quot;agent host&quot; site or library."),
    ("high", "Over-permissioning",
     "&quot;Users must have access to Site A&quot; exposes Site A content to every agent user. Verified alternative: Read on <b>only the "
     "<code>.agent</code> file</b> (or a dedicated folder/library) - Site A stays Access Denied and the Agent link works."),
    ("medium", "Two permission systems to keep in sync",
     "Copilot Studio sharing <i>and</i> SharePoint file permissions. Drive both from one Entra security group."),
    ("medium", "Misleading error message",
     "Users who cannot read the <code>.agent</code> file see <i>You need a Microsoft 365 Copilot license to access this agent</i> - even "
     "though unlicensed users work. Brief the help desk."),
    ("medium", "Capacity, not licenses",
     "Copilot Studio agents in SharePoint &quot;consume Copilot Studio capacity&quot; - unlicensed users work, so plan and monitor capacity / "
     "pay-as-you-go."),
    ("low", "One deployment site at a time",
     "Copilot Studio: <i>Only one SharePoint site can be deployed to at a time.</i> Moving the host means undeploy/redeploy and re-pointing "
     "every Agent link."),
    ("low", "Harness support",
     "GitHub Copilot harness agents are not available for the SharePoint channel and not supported by the Agents SDK client today. "
     "Keep SharePoint-facing agents on the standard harness."),
]

H2H = [
    ("1", "User A", "Shared on Agent A · Read on Site A · no M365 Copilot license", "WORKS", "WORKS",
     "35-agent-link-userA-works.png", "Agent link: side pane greets User A and answers"),
    ("2", "User C", "Not shared · no Site A access", "BLOCKED", "BLOCKED",
     None, "Agent link shows a license message; SDK shows 'You don't have access to talk to this bot'"),
    ("3", "User C", "Shared on Agent A · no Site A access", "BLOCKED", "WORKS",
     "36-agent-link-userC-no-site-a.png", "Agent link: 'You need a Microsoft 365 Copilot license...' (misleading)"),
    ("4", "User C", "Shared · Read on only the .agent file (Site A still Access Denied)", "WORKS", "WORKS",
     "38-agent-link-userC-file-read-only-works.png", "Agent link works with file-only Read"),
]


def pill(v):
    v2 = v.upper()
    if v2.startswith("WORKS") or v2 == "OK":
        cls = "ok"
    elif "NOT SUPPORTED" in v2 or v2.startswith("BLOCKED") or "DENIED" in v2:
        cls = "bad"
    elif "ANSWERS" in v2:
        cls = "warn"
    else:
        cls = "muted"
    return f'<span class="pill {cls}">{esc(v)}</span>'


def diagram_customer():
    return """<svg viewBox="0 0 560 300" role="img" aria-label="SharePoint channel plus Agent link pattern" class="diagram">
  <defs><marker id="ah1" markerWidth="10" markerHeight="10" refX="9" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" style="fill:var(--cp-text-muted)"/></marker>
  <marker id="ah1d" markerWidth="10" markerHeight="10" refX="9" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" style="fill:var(--cp-danger)"/></marker></defs>
  <rect x="10" y="20" width="160" height="70" rx="10" style="fill:var(--cp-surface);stroke:var(--cp-border-strong)"/>
  <text x="90" y="48" text-anchor="middle" class="t1">Site B / C page</text><text x="90" y="70" text-anchor="middle" class="t2">Agent link web part</text>
  <rect x="10" y="190" width="160" height="80" rx="10" style="fill:var(--cp-highlight);stroke:var(--cp-danger);stroke-dasharray:5 4"/>
  <text x="90" y="218" text-anchor="middle" class="t1">Site A (host)</text><text x="90" y="238" text-anchor="middle" class="t2">Documents/Copilot Studio</text><text x="90" y="254" text-anchor="middle" class="t2">Agents/&lt;agent&gt;.agent</text>
  <rect x="215" y="20" width="150" height="70" rx="10" style="fill:var(--cp-surface);stroke:var(--cp-border-strong)"/>
  <text x="290" y="48" text-anchor="middle" class="t1">Copilot side pane</text><text x="290" y="70" text-anchor="middle" class="t2">in SharePoint</text>
  <rect x="405" y="20" width="145" height="70" rx="10" style="fill:var(--cp-surface);stroke:var(--cp-border-strong)"/>
  <text x="478" y="48" text-anchor="middle" class="t1">Copilot Studio</text><text x="478" y="70" text-anchor="middle" class="t2">api.powerplatform.com</text>
  <rect x="405" y="190" width="145" height="80" rx="10" style="fill:var(--cp-surface);stroke:var(--cp-border-strong)"/>
  <text x="478" y="222" text-anchor="middle" class="t1">Agent</text><text x="478" y="244" text-anchor="middle" class="t2">sharing check</text>
  <line x1="170" y1="55" x2="212" y2="55" style="stroke:var(--cp-text-muted);stroke-width:2" marker-end="url(#ah1)"/>
  <line x1="290" y1="90" x2="170" y2="205" style="stroke:var(--cp-danger);stroke-width:2;stroke-dasharray:6 4" marker-end="url(#ah1d)"/>
  <text x="250" y="160" class="t3 danger">reads .agent file</text><text x="250" y="176" class="t3 danger">(user needs Read)</text>
  <line x1="365" y1="55" x2="402" y2="55" style="stroke:var(--cp-text-muted);stroke-width:2" marker-end="url(#ah1)"/>
  <line x1="478" y1="90" x2="478" y2="187" style="stroke:var(--cp-text-muted);stroke-width:2" marker-end="url(#ah1)"/>
</svg>"""


def diagram_sdk():
    return """<svg viewBox="0 0 560 300" role="img" aria-label="Agents SDK web part pattern" class="diagram">
  <defs><marker id="ah2" markerWidth="10" markerHeight="10" refX="9" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" style="fill:var(--cp-text-muted)"/></marker>
  <marker id="ah2s" markerWidth="10" markerHeight="10" refX="9" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" style="fill:var(--cp-success)"/></marker></defs>
  <rect x="10" y="20" width="170" height="70" rx="10" style="fill:var(--cp-surface);stroke:var(--cp-border-strong)"/>
  <text x="95" y="48" text-anchor="middle" class="t1">Any site page</text><text x="95" y="70" text-anchor="middle" class="t2">SPFx web part (inline chat)</text>
  <rect x="10" y="190" width="170" height="80" rx="10" style="fill:var(--cp-surface);stroke:var(--cp-border-strong)"/>
  <text x="95" y="220" text-anchor="middle" class="t1">Microsoft Entra ID</text><text x="95" y="242" text-anchor="middle" class="t2">MSAL ssoSilent (SPA app)</text>
  <rect x="300" y="20" width="250" height="70" rx="10" style="fill:var(--cp-surface);stroke:var(--cp-border-strong)"/>
  <text x="425" y="48" text-anchor="middle" class="t1">Copilot Studio (Agents SDK)</text><text x="425" y="70" text-anchor="middle" class="t2">api.powerplatform.com</text>
  <rect x="300" y="190" width="250" height="80" rx="10" style="fill:var(--cp-surface);stroke:var(--cp-border-strong)"/>
  <text x="425" y="222" text-anchor="middle" class="t1">Agent</text><text x="425" y="244" text-anchor="middle" class="t2">sharing check</text>
  <line x1="70" y1="90" x2="70" y2="187" style="stroke:var(--cp-text-muted);stroke-width:2" marker-end="url(#ah2)"/>
  <line x1="120" y1="187" x2="120" y2="93" style="stroke:var(--cp-text-muted);stroke-width:2" marker-end="url(#ah2)"/>
  <text x="128" y="145" class="t3">token (CopilotStudio.Copilots.Invoke)</text>
  <line x1="180" y1="55" x2="297" y2="55" style="stroke:var(--cp-success);stroke-width:2.5" marker-end="url(#ah2s)"/>
  <text x="190" y="44" class="t3 success">connection string</text>
  <line x1="425" y1="90" x2="425" y2="187" style="stroke:var(--cp-text-muted);stroke-width:2" marker-end="url(#ah2)"/>
  <text x="280" y="292" text-anchor="middle" class="t3 success">No host site - nothing to read in any other site</text>
</svg>"""


CSS = """
:root{color-scheme:light;--cp-bg:#f7f4ef;--cp-bg-elevated:#fcfbf8;--cp-surface:#ffffff;--cp-surface-soft:#f5f5f5;--cp-border:#dedede;--cp-border-strong:#919191;--cp-text:#242424;--cp-text-muted:#5c5c5c;--cp-text-soft:#6f6f6f;--cp-accent:#b11f4b;--cp-accent-hover:#9a1a41;--cp-accent-soft:rgba(177,31,75,0.08);--cp-accent-fg:#ffffff;--cp-success:#16a34a;--cp-danger:#dc2626;--cp-warning:#f59e0b;--cp-link:#0078d4;--cp-shadow:0 18px 48px rgba(0,0,0,0.12);--cp-overlay:rgba(255,255,255,0.8);--cp-panel:rgba(255,255,255,0.86);--cp-panel-strong:rgba(255,255,255,0.96);--cp-sheen:rgba(255,255,255,0.55);--cp-highlight:rgba(177,31,75,0.12);}
html[data-theme="dark"]{color-scheme:dark;--cp-bg:#3d3b3a;--cp-bg-elevated:#343231;--cp-surface:#292929;--cp-surface-soft:#2e2e2e;--cp-border:#474747;--cp-border-strong:#5f5f5f;--cp-text:#dedede;--cp-text-muted:#919191;--cp-text-soft:#b0b0b0;--cp-accent:#fd8ea1;--cp-accent-hover:#fb7b91;--cp-accent-soft:rgba(253,142,161,0.14);--cp-accent-fg:#1a1a1a;--cp-success:#4ade80;--cp-danger:#f87171;--cp-warning:#fbbf24;--cp-link:#4da6ff;--cp-shadow:0 18px 48px rgba(0,0,0,0.32);--cp-overlay:rgba(41,41,41,0.88);--cp-panel:rgba(41,41,41,0.72);--cp-panel-strong:rgba(41,41,41,0.96);--cp-sheen:rgba(255,255,255,0.04);--cp-highlight:rgba(253,142,161,0.12);}
*{box-sizing:border-box}
body{margin:0;background:var(--cp-bg);color:var(--cp-text);font-family:"Segoe UI",Aptos,Calibri,-apple-system,BlinkMacSystemFont,sans-serif;line-height:1.55;font-size:15px}
a{color:var(--cp-link)}
code,pre{font-family:Consolas,"Courier New",Courier,monospace}
code{background:var(--cp-surface-soft);border:1px solid var(--cp-border);border-radius:6px;padding:1px 5px;font-size:.9em}
pre{background:var(--cp-surface-soft);border:1px solid var(--cp-border);border-radius:.625rem;padding:12px 14px;overflow:auto;font-size:13px;margin:12px 0;white-space:pre-wrap;word-break:break-all}
header.top{position:sticky;top:0;z-index:20;background:var(--cp-panel-strong);border-bottom:1px solid var(--cp-border);backdrop-filter:blur(6px)}
.topin{max-width:1200px;margin:0 auto;padding:10px 24px;display:flex;align-items:center;gap:16px}
.brand{font-weight:700;white-space:nowrap}
nav.sections{display:flex;gap:4px;overflow-x:auto;flex:1}
nav.sections button{background:none;border:0;color:var(--cp-text-muted);font:inherit;font-size:13px;padding:6px 10px;border-radius:.625rem;cursor:pointer;white-space:nowrap}
nav.sections button:hover{background:var(--cp-accent-soft);color:var(--cp-text)}
.iconbtn{border:1px solid var(--cp-border);background:var(--cp-surface);color:var(--cp-text);border-radius:.625rem;padding:6px 10px;cursor:pointer;font:inherit;font-size:13px}
main{max-width:1200px;margin:0 auto;padding:8px 24px 64px}
section{padding:36px 0 8px;scroll-margin-top:64px}
h1{font-size:34px;line-height:1.2;margin:28px 0 8px}
h2{font-size:24px;margin:0 0 6px}
h3{font-size:17px;margin:0 0 6px}
.lead{color:var(--cp-text-muted);font-size:17px;max-width:900px}
.eyebrow{color:var(--cp-accent);font-weight:700;letter-spacing:.06em;text-transform:uppercase;font-size:12px}
.sub{color:var(--cp-text-muted);margin:0 0 18px}
.card{min-width:0;background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:16px;padding:18px 20px;box-shadow:0 0 2px rgba(0,0,0,0.12),0 1px 2px rgba(0,0,0,0.14)}
.grid{display:grid;gap:16px}
.g2{grid-template-columns:repeat(2,minmax(0,1fr))}.g3{grid-template-columns:repeat(3,minmax(0,1fr))}.g4{grid-template-columns:repeat(4,minmax(0,1fr))}
@media(max-width:900px){.g2,.g3,.g4{grid-template-columns:1fr}}
.kpi .n{font-size:30px;font-weight:700;color:var(--cp-accent)}.kpi .l{color:var(--cp-text-muted);font-size:13px}
.verdict{display:flex;gap:10px;align-items:flex-start}
.verdict .dot{width:12px;height:12px;border-radius:50%;margin-top:6px;flex:none}
.dot.ok{background:var(--cp-success)}.dot.bad{background:var(--cp-danger)}.dot.warn{background:var(--cp-warning)}
.pill{display:inline-block;border-radius:999px;padding:2px 10px;font-size:12px;font-weight:600;white-space:nowrap;border:1px solid transparent}
.pill.ok{background:color-mix(in srgb,var(--cp-success) 14%,transparent);color:var(--cp-text);border-color:color-mix(in srgb,var(--cp-success) 50%,transparent)}
.pill.bad{background:color-mix(in srgb,var(--cp-danger) 14%,transparent);color:var(--cp-text);border-color:color-mix(in srgb,var(--cp-danger) 50%,transparent)}
.pill.warn{background:color-mix(in srgb,var(--cp-warning) 18%,transparent);color:var(--cp-text);border-color:color-mix(in srgb,var(--cp-warning) 55%,transparent)}
.pill.muted{background:var(--cp-surface-soft);color:var(--cp-text-muted);border-color:var(--cp-border)}
ol.steps{margin:0;padding-left:20px}ol.steps li{margin:4px 0}
ul.checks{list-style:none;margin:0;padding:0}ul.checks li{padding:5px 0 5px 26px;position:relative}
ul.checks li:before{content:"\\2714";position:absolute;left:4px;color:var(--cp-success);font-weight:700}
.quote{border-left:4px solid var(--cp-accent);background:var(--cp-accent-soft);border-radius:0 .625rem .625rem 0;padding:12px 16px;margin-top:12px}
table{width:100%;border-collapse:collapse;background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:16px;overflow:hidden;font-size:14px}
th,td{text-align:left;padding:9px 12px;border-bottom:1px solid var(--cp-border);vertical-align:top}
th{background:var(--cp-surface-soft);font-weight:600;color:var(--cp-text)}
tr:last-child td{border-bottom:0}
.tablewrap{overflow-x:auto;border-radius:16px}
.diagram{width:100%;height:auto;background:var(--cp-bg-elevated);border:1px solid var(--cp-border);border-radius:16px}
.diagram .t1{font-size:15px;font-weight:600;fill:var(--cp-text)}.diagram .t2{font-size:12px;fill:var(--cp-text-muted)}
.diagram .t3{font-size:12px;fill:var(--cp-text-muted)}.diagram .danger{fill:var(--cp-danger)}.diagram .success{fill:var(--cp-success)}
figure.shot{margin:0;background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:16px;overflow:hidden;display:flex;flex-direction:column}
figure.shot img{width:100%;display:block;cursor:zoom-in;background:var(--cp-surface-soft)}
figure.shot figcaption{font-size:12.5px;color:var(--cp-text-muted);padding:8px 12px;border-top:1px solid var(--cp-border)}
.step{display:grid;grid-template-columns:48px minmax(0,1fr);gap:14px;margin:0 0 18px}
.stepn{width:40px;height:40px;border-radius:50%;background:var(--cp-accent);color:var(--cp-accent-fg);display:flex;align-items:center;justify-content:center;font-weight:700}
.shots{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px;margin-top:12px}
.concern{display:grid;grid-template-columns:auto 1fr;gap:12px;align-items:start}
.sev{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;padding:3px 8px;border-radius:6px;border:1px solid var(--cp-border)}
.sev.high{color:var(--cp-danger);border-color:var(--cp-danger)}.sev.medium{color:var(--cp-warning);border-color:var(--cp-warning)}.sev.low{color:var(--cp-text-muted)}
.filters{display:flex;flex-wrap:wrap;gap:8px;margin:10px 0 14px}
.filters button{border:1px solid var(--cp-border);background:var(--cp-surface);color:var(--cp-text);border-radius:999px;padding:5px 12px;cursor:pointer;font:inherit;font-size:13px}
.filters button.on{background:var(--cp-accent);border-color:var(--cp-accent);color:var(--cp-accent-fg)}
.callout{border:1px solid var(--cp-border);border-left:4px solid var(--cp-warning);background:var(--cp-surface);border-radius:.625rem;padding:12px 16px;margin:12px 0}
.callout.danger{border-left-color:var(--cp-danger)}.callout.ok{border-left-color:var(--cp-success)}
#lightbox{position:fixed;inset:0;background:var(--cp-overlay);display:none;align-items:center;justify-content:center;z-index:50;padding:24px;cursor:zoom-out}
#lightbox.show{display:flex}
#lightbox figure{margin:0;max-width:96vw;max-height:92vh;background:var(--cp-surface);border:1px solid var(--cp-border);border-radius:16px;overflow:hidden;box-shadow:var(--cp-shadow)}
#lightbox img{max-width:96vw;max-height:84vh;display:block}
#lightbox figcaption{padding:8px 14px;color:var(--cp-text-muted);font-size:13px}
footer{color:var(--cp-text-soft);font-size:12.5px;border-top:1px solid var(--cp-border);margin-top:40px;padding-top:16px}
.muted{color:var(--cp-text-muted)}
"""

JS = """
(function(){
  var lb=document.getElementById('lightbox'), lbImg=lb.querySelector('img'), lbCap=lb.querySelector('figcaption');
  document.addEventListener('click',function(e){
    var t=e.target;
    if(t.tagName==='IMG' && t.closest('figure.shot')){lbImg.src=t.src;lbCap.textContent=t.getAttribute('data-caption')||'';lb.classList.add('show');}
    else if(t.closest('#lightbox')){lb.classList.remove('show');}
  });
  document.addEventListener('keydown',function(e){if(e.key==='Escape')lb.classList.remove('show');});
  document.querySelectorAll('nav.sections button').forEach(function(b){
    b.addEventListener('click',function(){var el=document.getElementById(b.getAttribute('data-target'));if(el)el.scrollIntoView({behavior:'smooth',block:'start'});});
  });
  var tb=document.getElementById('themeToggle');
  tb.addEventListener('click',function(){var r=document.documentElement;r.setAttribute('data-theme',r.getAttribute('data-theme')==='dark'?'light':'dark');});
  var rows=document.querySelectorAll('#matrix tbody tr');
  document.querySelectorAll('#matrixFilters button').forEach(function(b){
    b.addEventListener('click',function(){
      document.querySelectorAll('#matrixFilters button').forEach(function(x){x.classList.remove('on');});
      b.classList.add('on');var f=b.getAttribute('data-f');
      rows.forEach(function(r){r.style.display=(f==='all'||r.getAttribute('data-k')===f||(r.getAttribute('data-v')||'').indexOf(f)>=0)?'':'none';});
    });
  });
})();
"""


def build():
    rows = load_results()
    sdk_works = sum(1 for r in rows if r["a_sdk"] == "WORKS")
    sdk_blocked = sum(1 for r in rows if r["a_sdk"].startswith("BLOCKED"))
    denied = sum(1 for r in rows if r["page"] != "OK")
    cells = len(rows) * 3

    nav = [("usecase", "Use case"), ("patterns", "Two patterns"), ("h2h", "Head-to-head"), ("compare", "Comparison"),
           ("concerns", "Concerns"), ("steps", "Step-by-step"), ("matrix", "Test matrix"), ("ghcp", "GitHub Copilot harness"),
           ("resources", "Resources")]
    out = []
    out.append('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
               '<meta name="robots" content="noindex">')
    out.append('<title>Authenticated Copilot Studio agents on many SharePoint sites - Microsoft 365 Agents SDK use case</title>')
    out.append('<script>(() => {const param = new URLSearchParams(window.location.search).get("scoutTheme");'
               'const theme = param || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");'
               'document.documentElement.setAttribute("data-theme", theme);})();</script>')
    out.append(f"<style>{CSS}</style></head><body>")
    out.append('<header class="top"><div class="topin"><div class="brand">Agents SDK &times; SharePoint</div><nav class="sections">')
    out += [f'<button data-target="{t}">{esc(l)}</button>' for t, l in nav]
    out.append('</nav><button class="iconbtn" id="themeToggle" title="Toggle light/dark">&#9680; Theme</button></div></header><main>')

    # Hero
    out.append('<section id="top"><div class="eyebrow">Customer use case &middot; validated 7 Oct 2026</div>'
               '<h1>One authenticated Copilot Studio agent on every SharePoint site</h1>'
               '<p class="lead">How your customer\'s SharePoint-channel setup compares with publishing the same agent through the '
               '<b>Microsoft 365 Agents SDK</b> - built, deployed and tested end-to-end in a live tenant with four test users, three sites, '
               'two agents and two harnesses.</p>')
    out.append('<div class="grid g4" style="margin-top:20px">')
    for n, l in ((f"{len(rows)}", "persona &times; site runs (5 personas, 3 sites)"), (f"{cells}", "web-part test cells"),
                 (f"{sdk_works} / {sdk_blocked}", "Agent A via SDK: works for shared users / blocked for unshared - exactly as shared"),
                 ("0", "extra site permissions needed with the SDK web part")):
        out.append(f'<div class="card kpi"><div class="n">{n}</div><div class="l">{l}</div></div>')
    out.append('</div>')
    out.append('<div class="grid g3" style="margin-top:16px">'
               '<div class="card verdict"><span class="dot ok"></span><div><h3>Works</h3>Standard-harness agent + Agents SDK web part on any '
               'number of sites, silent Entra SSO, user identity reaches the agent, sharing enforced (incl. revoke &amp; re-grant).</div></div>'
               '<div class="card verdict"><span class="dot warn"></span><div><h3>Works, with a dependency</h3>Customer pattern (SharePoint '
               'channel + Agent link): users need Read on the <code>.agent</code> file in Site A - verified that file-only Read is enough.</div></div>'
               '<div class="card verdict"><span class="dot bad"></span><div><h3>Not today</h3>GitHub Copilot harness agents via the Agents SDK '
               '(<i>This action doesn\'t support agents built with the GitHub Copilot harness</i>); their embed code is not authenticated.</div></div>'
               '</div></section>')

    # Use case
    out.append('<section id="usecase"><div class="eyebrow">The use case</div><h2>What the customer built and observed</h2>'
               '<p class="sub">Goal: let authenticated employees chat with the same Copilot Studio agent from many SharePoint sites.</p>'
               '<div class="grid g2"><div class="card"><h3>Customer setup (paraphrased)</h3><ol class="steps">'
               '<li>Agent deployed to <b>Site A</b> with the Copilot Studio <b>SharePoint channel</b>.</li>'
               '<li>On Site B, Site C and others: an <b>Agent link</b> web part pointing to that agent.</li>'
               '<li>Users are added (shared) on the agent.</li>'
               '<li>Users get access to Site A, Site B and every site where the agent appears.</li></ol></div>'
               '<div class="card"><h3>Customer observations</h3><ul class="checks">'
               '<li>Users can chat from any site that has the Agent link web part.</li>'
               '<li>Removing a user from the agent blocks them (<i>You don\'t have access to this agent</i>); adding them back restores access.</li>'
               '<li>Users must have access to Site A even when they use the agent from another site.</li>'
               '<li>Same behaviour for users with and without a Microsoft 365 Copilot license.</li></ul>'
               '<div class="quote"><b>Question:</b> any concerns or potential issues with this approach?</div></div></div>'
               '<div class="callout ok" style="margin-top:16px"><b>Short answer:</b> the approach is valid and supported (no code). We reproduced '
               'every observation. The real cost is the <b>Site A dependency</b>: it is a single point of failure and pushes you to over-grant Site A. '
               'Narrow it to Read on just the <code>.agent</code> file, or use the <b>Agents SDK web part</b>, which has no host-site dependency at all.</div>'
               '</section>')

    # Patterns
    out.append('<section id="patterns"><div class="eyebrow">Two patterns, one runtime</div><h2>How each pattern reaches the agent</h2>'
               '<p class="sub">Both end up in the same Copilot Studio runtime on <code>api.powerplatform.com</code> - the SharePoint channel\'s '
               '<code>.agent</code> file literally contains <code>botSchema</code>, <code>environmentID</code> and <code>transport: "rest"</code>. '
               'That is why sharing, revocation and licensing behave the same; the difference is the extra hop through the file in Site A.</p>'
               '<div class="grid g2"><div class="card"><h3>Customer pattern - SharePoint channel + Agent link</h3>' + diagram_customer() +
               '</div><div class="card"><h3>This implementation - Agents SDK web part</h3>' + diagram_sdk() + '</div></div>'
               '<div class="shots">' + figure("30-sp-channel-one-site-only.png", "Copilot Studio SharePoint channel: only one site at a time") +
               figure("32-agent-file-in-site-a.png", "The deployment creates Site A/Documents/Copilot Studio Agents/<agent>.agent") +
               figure("34-agent-link-on-site-b.png", "Agent link web part on Site B pointing to Site A's .agent file") +
               figure("20-site-page-overview.png", "Agents SDK web part: inline chat, any site") + '</div></section>')

    # Head to head
    out.append('<section id="h2h"><div class="eyebrow">Head-to-head on Site B</div><h2>Does the user need Site A?</h2>'
               '<p class="sub">Same agent (Agent A, standard harness), same page site, same users. Only the integration pattern changes.</p>'
               '<div class="tablewrap"><table><thead><tr><th>#</th><th>User</th><th>State</th><th>Agent link (customer)</th><th>Agents SDK web part</th></tr></thead><tbody>')
    for n, u, state, al, sdk, _, note in H2H:
        out.append(f'<tr><td>{n}</td><td>{u}</td><td>{esc(state)}<div class="muted" style="font-size:12.5px">{esc(note)}</div></td><td>{pill(al)}</td><td>{pill(sdk)}</td></tr>')
    out.append('</tbody></table></div><div class="shots">' +
               figure("35-agent-link-userA-works.png", "Test 1 - Agent link works for User A (no M365 Copilot license)") +
               figure("36-agent-link-userC-no-site-a.png", "Test 3 - Shared user without Site A: misleading license message") +
               figure("37-sdk-userC-no-site-a-works.png", "Test 3 - Same user via Agents SDK web part: works") +
               figure("38-agent-link-userC-file-read-only-works.png", "Test 4 - Read on only the .agent file fixes Agent link") + '</div></section>')

    # Comparison
    comp = [
        ("Effort", "No code", "One-time SPFx package + Entra SPA app + app catalog deployment"),
        ("Host-site dependency", "Yes - Read on the .agent file in the deployment site; one deployment site at a time", "None - each page stores the connection string"),
        ("Agent sharing &amp; revocation", "Enforced (customer test)", "Enforced (verified: unshared blocked, revoke and re-grant)"),
        ("Microsoft 365 Copilot license", "Not required (verified); consumes Copilot Studio capacity", "Not required (verified); consumes Copilot Studio capacity"),
        ("User experience", "Card on the page -> Copilot side pane", "Inline chat on the page, themable, several agents per page"),
        ("Error when access is missing", "Misleading license message", "Agent's own access message / explicit web part errors"),
        ("Governance", "Site A editors can delete/rename the .agent file", "Central versioned package; per-page configuration"),
        ("GitHub Copilot harness", "SharePoint channel not available", "Not supported by the Agents SDK client"),
    ]
    out.append('<section id="compare"><div class="eyebrow">Comparison</div><h2>SharePoint channel + Agent link vs. Agents SDK web part</h2>'
               '<div class="tablewrap"><table><thead><tr><th></th><th>SharePoint channel + Agent link</th><th>Agents SDK web part</th></tr></thead><tbody>')
    out += [f"<tr><td><b>{a}</b></td><td>{b}</td><td>{c}</td></tr>" for a, b, c in comp]
    out.append('</tbody></table></div></section>')

    # Concerns
    out.append('<section id="concerns"><div class="eyebrow">Answering the customer</div><h2>Concerns and recommendations</h2><div class="grid g2">')
    for sev, title, body in CONCERNS:
        out.append(f'<div class="card concern"><span class="sev {sev}">{sev}</span><div><h3>{esc(title)}</h3><div>{body}</div></div></div>')
    out.append('</div><div class="callout ok"><b>Recommended setup for the SharePoint-channel pattern:</b> dedicated agent host site/library, one Entra '
               'security group used both for Copilot Studio sharing and for Read on the <code>.agent</code> file, minimal Edit rights, capacity monitoring. '
               '<b>Choose the Agents SDK web part</b> when you need no host-site dependency, inline/branded chat, several agents per page, '
               'programmatic control, or cannot grant access to the host site.</div></section>')

    # Steps
    out.append('<section id="steps"><div class="eyebrow">Build it</div><h2>Step-by-step: the Microsoft 365 Agents SDK implementation</h2>'
               '<p class="sub">Exactly what was executed in the test tenant. Scripts and the SPFx solution are in the repository.</p>')
    for i, (title, body, shots, code) in enumerate(STEPS, 1):
        out.append(f'<div class="step"><div class="stepn">{i}</div><div class="card"><h3>{title}</h3><div>{body}</div>')
        if code:
            out.append(f"<pre>{esc(code)}</pre>")
        if shots:
            out.append('<div class="shots">' + "".join(figure(s, cap) for s, cap in shots) + "</div>")
        out.append("</div></div>")
    out.append("</section>")

    # Matrix
    out.append('<section id="matrix"><div class="eyebrow">Permission test matrix</div><h2>4 users &times; 3 sites &times; 2 agents</h2>'
               '<p class="sub">Agent A (standard harness) shared with User A + User B; Agent B (GitHub Copilot harness, Weather) shared with User A + '
               'User C; User D has no agent. Site permissions vary per user. Automated with Microsoft Edge + Playwright, one browser profile per user '
               '(MFA via TOTP), greeting + question per chat, screenshot per web part.</p>'
               '<div class="filters" id="matrixFilters"><button class="on" data-f="all">All</button>'
               '<button data-f="a">User A</button><button data-f="b">User B</button><button data-f="c">User C</button><button data-f="d">User D</button>'
               '<button data-f="mario">MarioR</button><button data-f="WORKS">Works</button><button data-f="BLOCKED">Blocked</button>'
               '<button data-f="DENIED">Access denied</button></div>'
               '<div class="tablewrap"><table id="matrix"><thead><tr><th>Persona</th><th>Agent sharing</th><th>Site</th><th>Site permission</th>'
               '<th>Page</th><th>Agent A via Agents SDK</th><th>Agent B via Agents SDK</th><th>Agent B via hosted embed</th></tr></thead><tbody>')
    for r in rows:
        v = " ".join([r["page"], r["a_sdk"], r["b_sdk"]])
        out.append(f'<tr data-k="{r["key"]}" data-v="{esc(v)}"><td>{esc(r["persona"])}</td><td>{esc(r["agents"])}</td><td>Site {r["site"]}</td>'
                   f'<td>{esc(r["perm"])}</td><td>{pill(r["page"])}</td><td>{pill(r["a_sdk"])}</td><td>{pill(r["b_sdk"])}</td><td>{pill(r["b_embed"])}</td></tr>')
    out.append('</tbody></table></div>'
               '<div class="grid g2" style="margin-top:16px"><div class="card"><h3>Additional checks</h3><ul class="checks">'
               '<li>Revocation: User B removed from Agent A -> blocked on next load; re-shared -> works again.</li>'
               '<li>Anonymous POST to either connection string -> HTTP 401.</li>'
               '<li>MFA users: silent SSO, no prompts; popup fallback never needed.</li>'
               '<li>No tester had a Microsoft 365 Copilot or Copilot Studio license.</li></ul></div>'
               '<div class="card"><h3>Reading the matrix</h3><p><span class="pill ok">WORKS</span> personalised greeting + on-topic answer. '
               '<span class="pill bad">BLOCKED - not shared</span> the agent replied <i>You don\'t have access to talk to this bot</i>. '
               '<span class="pill bad">NOT SUPPORTED</span> GitHub Copilot harness via Agents SDK. '
               '<span class="pill warn">ANSWERS (no auth enforced)</span> hosted embed answered regardless of sharing. '
               '<span class="pill bad">ACCESS DENIED</span> SharePoint page not accessible.</p></div></div>'
               '<div class="shots">' + figure("21-userA-agentA-sdk-works.png", "User A: 'Hello Agents SDK User A!' + answer") +
               figure("22-userD-agentA-sdk-not-shared.png", "User D: blocked by Copilot Studio sharing") +
               figure("28-userB-after-revoke.png", "User B right after removal from the agent") +
               figure("25-site-access-denied.png", "No site permission: SharePoint blocks the page") + '</div></section>')

    # GHCP
    out.append('<section id="ghcp"><div class="eyebrow">GitHub Copilot harness</div><h2>What happened with the original Weather agent</h2>'
               '<div class="grid g2"><div class="card"><h3>Agents SDK: not supported (yet)</h3><p>Every request - including from the agent owner - '
               'returned <i>This action doesn\'t support agents built with the GitHub Copilot harness.</i> Microsoft Learn: the Copilot Studio client '
               'library supports standard-harness agents only.</p></div>'
               '<div class="card"><h3>App-channel embed: renders but is not secured</h3><p>The hosted web chat answered users the agent was not shared with '
               'and even a signed-out browser (anonymous Direct Line token). Treat the existing embed/button on the root site as public until this '
               'is enforced.</p></div></div><div class="shots">' +
               figure("23-agentB-sdk-not-supported.png", "GitHub Copilot harness via Agents SDK") +
               figure("24-userB-agentB-embed-answers.png", "Embed answering a user who is not shared") +
               figure("26-embed-anonymous-browser.png", "Embed answering a signed-out browser") + '</div></section>')

    # Resources (relative links work both in docs/ and on the published site, where the files sit side by side)
    import sdk_content
    out.append('<section id="resources"><div class="eyebrow">Resources</div><h2>Companion material and references</h2><div class="grid g2">'
               '<div class="card"><h3>Companion material</h3><ul class="checks">'
               '<li><a href="Agents-SDK-SharePoint-Tutorial-and-Test-Scenarios.html">Tutorial, use cases and test scenarios</a> - the Agents SDK '
               f'implementation step by step, {len(sdk_content.USE_CASES)} use cases and {len(sdk_content.TESTS)} test scenarios</li>'
               '<li><a href="Agents-SDK-SharePoint-Test-Plan.xlsx" download>Test plan (Excel)</a> - run the scenarios in your tenant</li></ul>'
               f'<h3 style="margin-top:14px">Source code and guide</h3><p><a href="{REPO_URL}">{esc(REPO_URL)}</a></p><ul class="checks">'
               '<li><code>docs/Agents-SDK-SharePoint-Guide.md</code> - parameterised guide</li>'
               '<li><code>scripts/New-AgentsSdkSpaApp.ps1</code>, <code>Build-SpfxPackage.ps1</code>, <code>variables.sample.ps1</code></li>'
               '<li><code>spfx-copilot-agent/</code> - the web part</li><li><code>tests/agents-sdk-matrix/</code> - the automated test harness</li></ul></div>'
               '<div class="card"><h3>Microsoft documentation</h3><ul class="checks">'
               '<li><a href="https://learn.microsoft.com/microsoft-copilot-studio/publication-integrate-web-or-native-app-m365-agents-sdk">Integrate with web or native apps using the Microsoft 365 Agents SDK</a></li>'
               '<li><a href="https://learn.microsoft.com/microsoft-copilot-studio/publication-add-bot-to-sharepoint">Publish an agent to SharePoint</a></li>'
               '<li><a href="https://support.microsoft.com/office/manage-agents-in-sharepoint-bcab837a-835c-4a1a-8ad4-d53a353c369f">Manage agents in SharePoint</a></li>'
               '<li><a href="https://learn.microsoft.com/microsoft-copilot-studio/harnesses-overview">Harnesses in Copilot Studio</a></li>'
               '<li><a href="https://github.com/SharePoint/sp-dev-docs/issues/10775">sp-dev-docs #10775 (SPFx token provider limitation)</a></li></ul></div></div>'
               '<footer>Built from the automated test run in a Microsoft demo tenant (M365CPI03719864), 7 Oct 2026. Tenant identifiers are demo values; '
               'test-user credentials are not included. Behaviour of preview features may change.</footer></section>')

    out.append('</main><div id="lightbox" role="dialog" aria-label="Screenshot"><figure><img alt=""><figcaption></figcaption></figure></div>')
    out.append(f"<script>{JS}</script></body></html>")
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("".join(out))
    print(f"wrote {OUT} ({os.path.getsize(OUT) / 1024 / 1024:.2f} MB, {len(_cache)} images)")


if __name__ == "__main__":
    build()
