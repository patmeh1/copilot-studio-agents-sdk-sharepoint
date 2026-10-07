"""Build docs/Agents-SDK-SharePoint-Tutorial-and-Test-Scenarios.html - self-contained and sandbox-safe
(no external resources, no network calls, no storage), so it works as an attachment and in OneDrive/SharePoint preview.

    python tools/build_sdk_html.py
Content comes from tools/sdk_content.py; screenshots are embedded as WebP data URIs.
"""
import html
import os
import re

import build_usecase_html as base
import sdk_content as c

OUT = os.path.join(c.ROOT, "docs", "Agents-SDK-SharePoint-Tutorial-and-Test-Scenarios.html")

# Crop boxes (left, top, right, bottom) so tutorial screenshots show the relevant region at a readable size.
CROPS = {
    "07-create-standard-agent.png": (700, 210, 1470, 540),
    "10-agentA-authenticate-with-microsoft.png": (545, 90, 1095, 376),
    "09-agentA-greeting-yaml.png": (256, 90, 1330, 400),
    "11-agentA-web-app-connection-string.png": (1268, 0, 1912, 360),
    "12-agentA-share.png": (530, 192, 1390, 695),
    "16-enable-app.png": (1320, 90, 1912, 680),
    "18-webpart-toolbox.png": (900, 280, 1530, 760),
    "19-webpart-property-pane.png": (1525, 70, 1840, 600),
    "20-site-page-overview.png": (80, 60, 1590, 1000),
    "43-app-catalog-v1.0.3-enabled.png": (220, 325, 1020, 432),
}


def shot(name, caption):
    """Figure with the step's crop; very wide strips span the whole row so their text stays readable."""
    crop = CROPS.get(name)
    w, h = (crop[2] - crop[0], crop[3] - crop[1]) if crop else base.Image.open(os.path.join(base.IMG, name)).size
    return base.figure(name, caption, cls="wide" if w / h > 3 else "", crop=crop)


def md(text):
    """Inline markup (**bold**, *italic*, `code`, [text](url)) -> HTML."""
    s = html.escape(text, quote=False)
    codes = []

    def keep(m):
        codes.append(f"<code>{m.group(1)}</code>")
        return f"\x00{len(codes) - 1}\x00"

    s = re.sub(r"`([^`]+)`", keep, s)
    s = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", r'<a href="\2">\1</a>', s)
    s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    s = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<i>\1</i>", s)
    return re.sub(r"\x00(\d+)\x00", lambda m: codes[int(m.group(1))], s)


def pill(label, cls):
    return f'<span class="pill {cls}">{html.escape(label)}</span>'


def code_block(code):
    return (f'<div class="codewrap"><button class="copy" type="button" title="Copy to clipboard">Copy</button>'
            f"<pre><code>{html.escape(code)}</code></pre></div>")


CAPTIONS = {
    "20-site-page-overview.png": "One page, several agents: Agent A answers, Agent B is rejected",
    "21-userA-agentA-sdk-works.png": "User A: greeting by name and answer; footer shows the signed-in account",
    "22-userD-agentA-sdk-not-shared.png": "User D (not shared): blocked by Copilot Studio",
    "23-agentB-sdk-not-supported.png": "GitHub Copilot harness agent through the Agents SDK",
    "24-userB-agentB-embed-answers.png": "Hosted embed answering a user it is not shared with",
    "25-site-access-denied.png": "No site permission: SharePoint blocks the page",
    "26-embed-anonymous-browser.png": "Hosted embed answering a signed-out browser",
    "28-userB-after-revoke.png": "User B right after removal from sharing",
    "37-sdk-userC-no-site-a-works.png": "User C without Site A access chats from Site B",
    "39-tc03-v1.0.1-popup-blocked.png": "Before the fix (1.0.1.0): first click blocked by the popup blocker",
    "40-tc03-sign-in-prompt.png": "Silent sign-in blocked: one-click sign-in offered",
    "41-tc03-after-popup-sign-in.png": "After one click: greeting and answer (1.0.3.0)",
    "42-tc22-new-chat.png": "New chat: transcript cleared, greeting shown again",
    "43-app-catalog-v1.0.3-enabled.png": "Version 1.0.3.0 enabled for all sites",
}


def evidence(paths):
    names = [p.split("/", 1)[1] for p in paths if p.startswith("images/")]
    shots = [shot(n, CAPTIONS.get(n, n)) for n in names]
    files = [f"<code>docs/{html.escape(p)}</code>" for p in paths if not p.startswith("images/")]
    out = ""
    if files:
        out += f'<p class="muted small">Raw results: {", ".join(files)}</p>'
    if shots:
        out += '<div class="shots">' + "".join(shots) + "</div>"
    return out or '<p class="muted small">No screenshot - see the result text.</p>'


EXTRA_CSS = """
.small{font-size:12.5px}
.kpi .n small{font-size:15px;color:var(--cp-text-muted);font-weight:600}
.uc{display:flex;flex-direction:column;gap:6px}
.uc p{margin:0}
.uchead{display:flex;justify-content:space-between;align-items:center;gap:8px}
.ucid{font-weight:700;color:var(--cp-accent);font-size:13px;letter-spacing:.04em}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:auto;padding-top:8px}
.chip{border:1px solid var(--cp-border);background:var(--cp-surface-soft);color:var(--cp-text);border-radius:999px;padding:2px 10px;font:inherit;font-size:12px;cursor:pointer}
.chip:hover{border-color:var(--cp-accent);color:var(--cp-accent)}
.part{margin:26px 0 12px;font-size:19px;display:flex;align-items:center;gap:10px}
.part .tag{font-size:12px;font-weight:700;color:var(--cp-accent-fg);background:var(--cp-accent);border-radius:999px;padding:2px 10px}
.codewrap{position:relative}
.codewrap pre{padding-right:70px}
.copy{position:absolute;top:8px;right:8px;border:1px solid var(--cp-border);background:var(--cp-surface);color:var(--cp-text-muted);border-radius:6px;font:inherit;font-size:12px;padding:2px 8px;cursor:pointer}
.copy:hover{color:var(--cp-accent);border-color:var(--cp-accent)}
.check{border:1px solid var(--cp-border);border-left:4px solid var(--cp-success);background:var(--cp-bg-elevated);border-radius:.625rem;padding:8px 14px;margin:12px 0 0}
.why{color:var(--cp-text-muted);margin:0 0 8px}
.stack{display:flex;height:14px;border-radius:999px;overflow:hidden;border:1px solid var(--cp-border);margin:10px 0 8px}
.stack span{display:block;height:100%}
.seg-ok{background:var(--cp-success)}.seg-bad{background:var(--cp-danger)}.seg-warn{background:var(--cp-warning)}.seg-muted{background:var(--cp-border-strong)}
.legend{display:flex;flex-wrap:wrap;gap:14px;font-size:13px;color:var(--cp-text-muted)}
.legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;vertical-align:middle}
.toolbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin:14px 0}
.toolbar .group{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.toolbar .lbl{font-size:12px;color:var(--cp-text-muted);font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-right:2px}
.toolbar button.f{border:1px solid var(--cp-border);background:var(--cp-surface);color:var(--cp-text);border-radius:999px;padding:4px 11px;cursor:pointer;font:inherit;font-size:13px}
.toolbar button.f.on{background:var(--cp-accent);border-color:var(--cp-accent);color:var(--cp-accent-fg)}
.toolbar select,.toolbar input[type=search]{border:1px solid var(--cp-border);background:var(--cp-surface);color:var(--cp-text);border-radius:.625rem;padding:5px 10px;font:inherit;font-size:13px}
.toolbar input[type=search]{min-width:220px}
.toolbar label{display:inline-flex;align-items:center;gap:6px;cursor:pointer}
.toolbar input[type=checkbox]{accent-color:var(--cp-accent);margin:0}
table.tc tbody tr.row{cursor:pointer}
table.tc tbody tr.row:hover td{background:var(--cp-accent-soft)}
table.tc tbody tr.row.open td{background:var(--cp-highlight)}
table.tc td.id{font-weight:700;white-space:nowrap;color:var(--cp-accent)}
table.tc tr.detail td{background:var(--cp-bg-elevated)}
.detailgrid{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,1fr);gap:18px}
@media(max-width:900px){.detailgrid{grid-template-columns:1fr}}
.detailgrid dl{margin:0;display:grid;grid-template-columns:130px minmax(0,1fr);gap:6px 12px}
.detailgrid dt{font-weight:600;color:var(--cp-text-muted);font-size:13px}
.detailgrid dd{margin:0}
.detailgrid ol{margin:0;padding-left:18px}
.caret{display:inline-block;width:12px;color:var(--cp-text-muted);transition:transform .15s}
tr.open .caret{transform:rotate(90deg)}
.flash td{animation:flash 1.6s ease-out}
@keyframes flash{from{background:var(--cp-highlight)}to{background:transparent}}
.empty{padding:18px;color:var(--cp-text-muted);text-align:center}
.limits li{margin:6px 0}
#tutorial .shots{grid-template-columns:repeat(auto-fit,minmax(280px,1fr))}
.shots.one{grid-template-columns:minmax(0,680px)}
#tutorial figure.shot img{max-height:460px;object-fit:contain}
.shots figure.wide{grid-column:1/-1}
.detailgrid .shots{grid-template-columns:repeat(auto-fill,minmax(200px,1fr))}
"""

JS = """
(function(){
  var lb=document.getElementById('lightbox'), lbImg=lb.querySelector('img'), lbCap=lb.querySelector('figcaption');
  document.addEventListener('click',function(e){
    var t=e.target;
    if(t.tagName==='IMG' && t.closest('figure.shot')){lbImg.src=t.src;lbCap.textContent=t.getAttribute('data-caption')||'';lb.classList.add('show');e.stopPropagation();}
    else if(t.closest('#lightbox')){lb.classList.remove('show');}
  });
  document.addEventListener('keydown',function(e){if(e.key==='Escape')lb.classList.remove('show');});
  function go(id){var el=document.getElementById(id);if(el)el.scrollIntoView({behavior:'smooth',block:'start'});}
  document.querySelectorAll('[data-target]').forEach(function(b){b.addEventListener('click',function(){go(b.getAttribute('data-target'));});});
  document.getElementById('themeToggle').addEventListener('click',function(){var r=document.documentElement;r.setAttribute('data-theme',r.getAttribute('data-theme')==='dark'?'light':'dark');});

  document.querySelectorAll('button.copy').forEach(function(b){
    b.addEventListener('click',function(){
      var text=b.parentNode.querySelector('code').textContent, done=function(){b.textContent='Copied';setTimeout(function(){b.textContent='Copy';},1500);};
      try{navigator.clipboard.writeText(text).then(done,function(){selectCode(b);});}catch(err){selectCode(b);}
    });
  });
  function selectCode(b){var r=document.createRange();r.selectNodeContents(b.parentNode.querySelector('code'));var s=window.getSelection();s.removeAllRanges();s.addRange(r);b.textContent='Press Ctrl+C';}

  var ucs=document.querySelectorAll('.uc');
  document.querySelectorAll('#ucFilters button').forEach(function(b){
    b.addEventListener('click',function(){
      document.querySelectorAll('#ucFilters button').forEach(function(x){x.classList.remove('on');});b.classList.add('on');
      var f=b.getAttribute('data-f');ucs.forEach(function(u){u.style.display=(f==='all'||u.getAttribute('data-status')===f)?'':'none';});
    });
  });

  var rows=[].slice.call(document.querySelectorAll('table.tc tr.row')), state={status:'all',priority:'all',area:'all',core:false,q:''};
  function apply(){
    var shown=0;
    rows.forEach(function(r){
      var ok=(state.status==='all'||r.dataset.status===state.status)&&(state.priority==='all'||r.dataset.priority===state.priority)&&
        (state.area==='all'||r.dataset.area===state.area)&&(!state.core||r.dataset.core==='1')&&(!state.q||r.dataset.text.indexOf(state.q)>=0);
      r.style.display=ok?'':'none';
      var d=r.nextElementSibling; if(!ok){d.style.display='none';r.classList.remove('open');}
      else{d.style.display=r.classList.contains('open')?'':'none';}
      if(ok)shown++;
    });
    document.getElementById('tcCount').textContent=shown+' of '+rows.length+' scenarios';
    document.getElementById('tcEmpty').style.display=shown?'none':'';
  }
  function group(id,key){
    document.querySelectorAll('#'+id+' button.f').forEach(function(b){
      b.addEventListener('click',function(){document.querySelectorAll('#'+id+' button.f').forEach(function(x){x.classList.remove('on');});b.classList.add('on');state[key]=b.getAttribute('data-v');apply();});
    });
  }
  group('fStatus','status');group('fPriority','priority');
  document.getElementById('fArea').addEventListener('change',function(e){state.area=e.target.value;apply();});
  document.getElementById('fCore').addEventListener('change',function(e){state.core=e.target.checked;apply();});
  document.getElementById('fSearch').addEventListener('input',function(e){state.q=e.target.value.trim().toLowerCase();apply();});
  function toggle(r,open){var d=r.nextElementSibling;var o=open===undefined?!r.classList.contains('open'):open;r.classList.toggle('open',o);r.setAttribute('aria-expanded',o?'true':'false');d.style.display=o?'':'none';}
  rows.forEach(function(r){
    r.addEventListener('click',function(){toggle(r);});
    r.addEventListener('keydown',function(e){if(e.key==='Enter'||e.key===' '){e.preventDefault();toggle(r);}});
  });
  document.getElementById('expandAll').addEventListener('click',function(){rows.forEach(function(r){if(r.style.display!=='none')toggle(r,true);});});
  document.getElementById('collapseAll').addEventListener('click',function(){rows.forEach(function(r){toggle(r,false);});});
  function reset(){
    state={status:'all',priority:'all',area:'all',core:false,q:''};
    ['fStatus','fPriority'].forEach(function(id){document.querySelectorAll('#'+id+' button.f').forEach(function(b,i){b.classList.toggle('on',i===0);});});
    document.getElementById('fArea').value='all';document.getElementById('fCore').checked=false;document.getElementById('fSearch').value='';
  }
  document.querySelectorAll('.chip[data-tc]').forEach(function(ch){
    ch.addEventListener('click',function(){
      var r=document.getElementById('row-'+ch.getAttribute('data-tc'));if(!r)return;
      reset();apply();toggle(r,true);r.scrollIntoView({behavior:'smooth',block:'center'});
      r.classList.remove('flash');void r.offsetWidth;r.classList.add('flash');
    });
  });
  apply();
})();
"""


def build():
    c.check()
    k = c.counts()
    total = len(c.TESTS)
    validated = sum(u["status"] == "validated" for u in c.USE_CASES)
    nav = [("overview", "Overview"), ("usecases", "Use cases"), ("tutorial", "Tutorial"), ("tests", "Test scenarios"),
           ("yourtenant", "Run it in your tenant"), ("limits", "Limits"), ("resources", "Resources")]
    o = ['<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
         '<meta name="robots" content="noindex">',
         "<title>Microsoft 365 Agents SDK in SharePoint - tutorial, use cases and test scenarios</title>",
         '<script>(() => {const param = new URLSearchParams(window.location.search).get("scoutTheme");'
         'const theme = param || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");'
         'document.documentElement.setAttribute("data-theme", theme);})();</script>',
         f"<style>{base.CSS}{EXTRA_CSS}</style></head><body>",
         '<header class="top"><div class="topin"><div class="brand">Agents SDK &times; SharePoint</div><nav class="sections">']
    o += [f'<button data-target="{t}">{html.escape(l)}</button>' for t, l in nav]
    o.append('</nav><button class="iconbtn" id="themeToggle" title="Toggle light/dark">&#9680; Theme</button></div></header><main>')

    # Hero
    o.append(f'<section id="top"><div class="eyebrow">Microsoft 365 Agents SDK &middot; validated {c.TESTED_ON}</div>'
             '<h1>Authenticated Copilot Studio agents on any SharePoint site</h1>'
             '<p class="lead">Tutorial, use cases and test scenarios for publishing a Copilot Studio agent through the <b>Microsoft 365 Agents '
             'SDK</b> - one SharePoint web part for the whole tenant, single sign-on with Microsoft Entra ID, and Copilot Studio sharing as '
             'the access control. Built and tested end to end in a live Microsoft demo tenant.</p><div class="grid g4" style="margin-top:20px">')
    kpis = [(f"{len(c.USE_CASES)}", f"use cases - {validated} validated in the demo tenant"),
            (f"{sum(len(p['steps']) for p in c.TUTORIAL)}", "tutorial steps with checkpoints and screenshots"),
            (f"{total}", "test scenarios with expected and actual results"),
            (f"{k['pass']}<small> / {total}</small>", f"passed; {k['limit']} limitation, {k['finding']} finding, {k['notrun']} to validate in your tenant")]
    o += [f'<div class="card kpi"><div class="n">{n}</div><div class="l">{l}</div></div>' for n, l in kpis]
    o.append("</div></section>")

    # Overview
    o.append('<section id="overview"><div class="eyebrow">Overview</div><h2>How the Agents SDK web part works</h2>'
             '<p class="sub">The page talks to Copilot Studio directly with the user\'s own Entra ID token - no host site, no hidden '
             'files, no anonymous path.</p><div class="grid g2"><div class="card">' + base.diagram_sdk() + '</div><div class="card">'
             '<h3>What you get</h3><ul class="checks">'
             '<li>One SPFx package and one Entra SPA app for every site in the tenant.</li>'
             '<li>Silent single sign-on with the SharePoint session, including MFA; one-click fallback when a browser blocks it.</li>'
             '<li>The agent knows who is chatting (<code>System.User.DisplayName</code> and more).</li>'
             '<li>Copilot Studio sharing enforced on every site; changes apply on the next page load.</li>'
             '<li>Inline, themable chat; several agents per page; no Microsoft 365 Copilot license needed.</li></ul>'
             '<h3 style="margin-top:14px">Three independent gates</h3><ol class="steps">'
             '<li><b>Microsoft Entra ID</b> - the user is signed in (token for <code>CopilotStudio.Copilots.Invoke</code>).</li>'
             '<li><b>Copilot Studio sharing</b> - the agent is shared with the user or their group.</li>'
             '<li><b>SharePoint permission</b> - the user can open the page (Visitor is enough).</li></ol></div></div></section>')

    # Use cases
    o.append('<section id="usecases"><div class="eyebrow">Use cases</div><h2>What the Agents SDK web part enables</h2>'
             '<p class="sub">Status reflects this engagement: <b>Validated</b> = executed in the demo tenant, <b>To validate</b> = run it in your '
             'tenant before relying on it, <b>By design</b> = supported pattern that was not tested here.</p>'
             '<div class="filters" id="ucFilters"><button class="on" data-f="all">All</button><button data-f="validated">Validated</button>'
             '<button data-f="validate">To validate</button><button data-f="design">By design</button></div><div class="grid g3">')
    for u in c.USE_CASES:
        label, cls = c.UC_STATUS[u["status"]]
        chips = "".join(f'<button class="chip" data-tc="{t}" title="Show test scenario {t}">{t}</button>' for t in u["tests"])
        o.append(f'<div class="card uc" data-status="{u["status"]}"><div class="uchead"><span class="ucid">{u["id"]}</span>{pill(label, cls)}</div>'
                 f'<h3>{html.escape(u["title"])}</h3><p class="muted">{md(u["scenario"])}</p><p><b>How:</b> {md(u["how"])}</p>'
                 f'<p class="small"><b>Evidence:</b> {md(u["evidence"])}</p>' + (f'<div class="chips">{chips}</div>' if chips else "") + "</div>")
    o.append("</div></section>")

    # Tutorial
    o.append('<section id="tutorial"><div class="eyebrow">Tutorial</div><h2>Build it step by step</h2>'
             '<p class="sub">About 60-90 minutes the first time; adding the agent to another site takes about 2 minutes. Scripts and the SPFx '
             'solution are in the repository. Every checkpoint below was observed in the demo tenant.</p>'
             '<div class="callout"><b>You need:</b> Copilot Studio maker access and capacity &middot; permission to register an Entra app and grant '
             'tenant-wide consent &middot; SharePoint Administrator + app catalog site admin &middot; Windows with PowerShell 7, Azure CLI and Node.js 22.</div>')
    n = 0
    for pi, part in enumerate(c.TUTORIAL, 1):
        o.append(f'<h3 class="part"><span class="tag">Part {pi}</span>{html.escape(part["part"])}</h3>')
        for st in part["steps"]:
            n += 1
            o.append(f'<div class="step"><div class="stepn">{n}</div><div class="card"><h3>{html.escape(st["title"])}</h3>'
                     f'<p class="why">{md(st["why"])}</p><ol class="steps">' + "".join(f"<li>{md(d)}</li>" for d in st["do"]) + "</ol>")
            if st.get("code"):
                o.append(code_block(st["code"][1]))
            o.append(f'<div class="check"><b>Checkpoint:</b> {md(st["check"])}</div>')
            if st["images"]:
                wide = " one" if len(st["images"]) == 1 else ""
                o.append(f'<div class="shots{wide}">' + "".join(shot(i, cap) for i, cap in st["images"]) + "</div>")
            o.append("</div></div>")
    o.append("</section>")

    # Test scenarios
    seg = [("pass", "ok"), ("limit", "bad"), ("finding", "warn"), ("notrun", "muted")]
    o.append(f'<section id="tests"><div class="eyebrow">Test scenarios</div><h2>{total} scenarios with expected and actual results</h2>'
             f'<p class="sub">Executed in the demo tenant on {c.TESTED_ON} with 4 test users, 3 sites, 2 agents and the maker as control '
             f'(web part {c.PACKAGE_VERSION}). Select a row for steps, expected result, actual result and evidence.</p>'
             '<div class="grid g2" style="align-items:start"><div class="card"><h3>Results in the demo tenant</h3><div class="stack">' +
             "".join(f'<span class="seg-{cls}" style="width:{k[key] * 100 / total:.2f}%" title="{c.STATUS[key][0]}: {k[key]}"></span>' for key, cls in seg) +
             '</div><div class="legend">' + "".join(f'<span><i class="seg-{cls}"></i>{c.STATUS[key][0].split(" - ")[0]}: <b>{k[key]}</b></span>' for key, cls in seg) +
             '</div><p class="small muted" style="margin-top:10px">Defect found and fixed during testing: TC-03 (sign-in popup blocked on pages with '
             'several agents) - fixed in web part 1.0.2.0/1.0.3.0 and re-tested.</p></div><div class="card"><h3>By area</h3>'
             '<div class="tablewrap"><table><thead><tr><th>Area</th><th>Scenarios</th><th>Pass</th><th>Limitation / finding</th><th>Not run</th></tr></thead><tbody>')
    for area in c.AREAS:
        items = [t for t in c.TESTS if t["area"] == area]
        o.append(f'<tr><td>{area}</td><td>{len(items)}</td><td>{sum(t["status"] == "pass" for t in items)}</td>'
                 f'<td>{sum(t["status"] in ("limit", "finding") for t in items)}</td><td>{sum(t["status"] == "notrun" for t in items)}</td></tr>')
    o.append("</tbody></table></div></div></div>")
    o.append('<div class="toolbar"><div class="group" id="fStatus"><span class="lbl">Status</span>'
             '<button class="f on" data-v="all">All</button>' +
             "".join(f'<button class="f" data-v="{key}">{c.STATUS[key][0].split(" - ")[0]}</button>' for key, _ in seg) +
             '</div><div class="group" id="fPriority"><span class="lbl">Priority</span><button class="f on" data-v="all">All</button>'
             '<button class="f" data-v="P1">P1</button><button class="f" data-v="P2">P2</button><button class="f" data-v="P3">P3</button></div>'
             '<div class="group"><select id="fArea" aria-label="Area"><option value="all">All areas</option>' +
             "".join(f'<option value="{html.escape(a)}">{html.escape(a)}</option>' for a in c.AREAS) +
             '</select><label class="small"><input type="checkbox" id="fCore"> Core set only</label>'
             '<input id="fSearch" type="search" placeholder="Search scenarios" aria-label="Search scenarios"></div>'
             '<div class="group"><button class="f" id="expandAll" type="button">Expand all</button><button class="f" id="collapseAll" type="button">Collapse all</button>'
             '<span class="small muted" id="tcCount"></span></div></div>')
    o.append('<div class="tablewrap"><table class="tc"><thead><tr><th>ID</th><th>Scenario</th><th>Area</th><th>Priority</th><th>Result in the demo tenant</th></tr></thead><tbody>')
    for t in c.TESTS:
        label, cls = c.STATUS[t["status"]]
        text = " ".join([t["id"], t["title"], t["area"], t["pre"], " ".join(t["steps"]), t["expected"], t["actual"]]).lower()
        o.append(f'<tr class="row" id="row-{t["id"]}" tabindex="0" aria-expanded="false" data-status="{t["status"]}" data-priority="{t["priority"]}" '
                 f'data-area="{html.escape(t["area"])}" data-core="{1 if t["id"] in c.CORE_SET else 0}" data-text="{html.escape(text)}">'
                 f'<td class="id"><span class="caret">&#9656;</span> {t["id"]}</td><td>{html.escape(t["title"])}</td><td>{t["area"]}</td>'
                 f'<td>{t["priority"]}</td><td>{pill(label.split(" - ")[0], cls)}</td></tr>')
        steps = "".join(f"<li>{md(s)}</li>" for s in t["steps"])
        o.append(f'<tr class="detail" style="display:none"><td colspan="5"><div class="detailgrid"><dl>'
                 f"<dt>Preconditions</dt><dd>{md(t['pre'])}</dd><dt>Steps</dt><dd><ol>{steps}</ol></dd>"
                 f"<dt>Expected</dt><dd>{md(t['expected'])}</dd><dt>Actual (demo)</dt><dd>{md(t['actual'])}</dd>"
                 f"<dt>Automation</dt><dd><code>{html.escape(t['auto'])}</code></dd></dl><div>{evidence(t['evidence'])}</div></div></td></tr>")
    o.append('</tbody></table></div><div class="empty" id="tcEmpty" style="display:none">No scenario matches the filters.</div>')
    o.append('<div class="grid g2" style="margin-top:16px"><div class="card"><h3>Test setup</h3>'
             '<p class="small">Agent A = standard harness, shared with User A and User B. Agent B (Weather) = GitHub Copilot harness, shared with '
             'User A and User C. User D has no agent. Every site page hosts Agent A and Agent B via the Agents SDK web part plus Agent B\'s hosted '
             'embed for comparison. Test users: Microsoft 365 E5 (no Teams), no Copilot licenses, MFA.</p><div class="tablewrap"><table>'
             '<thead><tr><th>Persona</th><th>Agent sharing</th><th>Site A</th><th>Site B</th><th>Site C</th></tr></thead><tbody>' +
             "".join("<tr>" + "".join(f"<td>{html.escape(v)}</td>" for v in p) + "</tr>" for p in c.PERSONAS) +
             '</tbody></table></div></div><div class="card"><h3>Reading the results</h3><p>' + pill("Pass", "ok") +
             ' behaved as expected. ' + pill("Confirmed limitation", "bad") + ' documented product limitation reproduced. ' +
             pill("Security finding", "warn") + ' behaviour to be aware of (not the Agents SDK path). ' + pill("Not run", "muted") +
             ' recommended for your tenant - the expected result is the requirement to verify.</p><p class="small muted">The <b>core set</b> '
             f'({", ".join(c.CORE_SET)}) is the quick regression after any agent or package change.</p></div></div></section>')

    # Your tenant
    o.append('<section id="yourtenant"><div class="eyebrow">Run it in your tenant</div><h2>From demo results to your sign-off</h2><div class="grid g2">'
             '<div class="card"><h3>1. Use the Excel test plan</h3><p>The <a href="Agents-SDK-SharePoint-Test-Plan.xlsx" download>Excel test plan</a> lists every scenario with '
             'preconditions, steps, expected result and the demo result. Fill in <b>Your result</b>, <b>Tester</b>, <b>Date</b> and <b>Notes</b>; the '
             'Summary sheet counts passes, failures and the P1 pass rate.</p></div>'
             '<div class="card"><h3>2. Start with P1 and the core set</h3><p>' +
             f'{sum(t["priority"] == "P1" for t in c.TESTS)} scenarios are P1 (must pass before go-live), including the ones this engagement could not run: '
             'security-group sharing (TC-09) and SharePoint knowledge permissions (TC-26, when used).</p></div>'
             '<div class="card"><h3>3. Automate the regression</h3><p>The Playwright harness drives real Edge windows, one profile per test user, '
             'and records greeting, answer, errors and a screenshot per web part.</p>' +
             code_block("cd tests\\agents-sdk-matrix\n.\\launch-edge.ps1\nforeach ($p in 'mario','a','b','c','d') { python run_matrix.py $p }\n"
                        "python run_extra_scenarios.py a A   # TC-03 + TC-22\npython summarize.py") + '</div>'
             '<div class="card"><h3>4. Keep test accounts safe</h3><p>Test-user passwords and MFA seeds stay in a secrets folder outside the repository '
             '(<code>$env:AGENTS_SDK_TEST_SECRETS_DIR</code>). Delete the test users and the SPA app registration when you are done '
             '(guide section 13).</p></div></div></section>')

    # Limits
    o.append('<section id="limits"><div class="eyebrow">Limits</div><h2>Not supported, and known design limits</h2><div class="grid g2">'
             '<div class="card"><h3>Not supported</h3><ul class="limits">' +
             "".join(f"<li><b>{md(a)}</b> - {md(b)}</li>" for a, b in c.NOT_SUPPORTED) + '</ul></div>'
             '<div class="card"><h3>Design limits of this web part</h3><ul class="limits">'
             '<li><b>Token lifetime</b> - the client receives an access token, not a token provider. If a long-open conversation stops answering, '
             '<b>New chat</b> signs in again silently (TC-23).</li>'
             '<li><b>Embedded hosts</b> - Teams, Viva Connections or the mobile app can block the hidden sign-in frame or popups; validate TC-15 '
             'before using these entry points.</li>'
             '<li><b>Page context</b> - not passed to the agent yet; the documented event pattern needs a small extension (UC-11).</li>'
             '<li><b>Capacity</b> - unlicensed users consume Copilot Studio capacity; monitor it (TC-17).</li></ul></div></div></section>')

    # Resources (relative links work both in docs/ and on the published site, where the files sit side by side)
    o.append('<section id="resources"><div class="eyebrow">Resources</div><h2>Companion material and references</h2><div class="grid g2">'
             '<div class="card"><h3>Companion material</h3><ul class="checks">'
             '<li><a href="Agents-SDK-SharePoint-Use-Case.html">Customer pattern: SharePoint channel + Agent link vs. Agents SDK</a> - the '
             'customer\'s setup reproduced in the same tenant and compared</li>'
             '<li><a href="Agents-SDK-SharePoint-Test-Plan.xlsx" download>Test plan (Excel)</a> - record your own results</li></ul>'
             f'<h3 style="margin-top:14px">Source code and guide</h3><p><a href="{c.REPO_URL}">{html.escape(c.REPO_URL)}</a></p><ul class="checks">'
             '<li><code>docs/Agents-SDK-SharePoint-Tutorial.md</code> and <code>docs/Agents-SDK-SharePoint-Use-Cases-and-Test-Scenarios.md</code> - '
             'this page as markdown</li>'
             '<li><code>docs/Agents-SDK-SharePoint-Guide.md</code> - full guide, findings and troubleshooting</li>'
             '<li><code>spfx-copilot-agent/</code>, <code>scripts/</code>, <code>tests/agents-sdk-matrix/</code> - web part, setup scripts, test harness</li></ul></div>'
             '<div class="card"><h3>Microsoft documentation</h3><ul class="checks">'
             '<li><a href="https://learn.microsoft.com/microsoft-copilot-studio/publication-integrate-web-or-native-app-m365-agents-sdk">Integrate with web or native apps using the Microsoft 365 Agents SDK</a></li>'
             '<li><a href="https://learn.microsoft.com/microsoft-365/agents-sdk/integrate-with-mcs">Integrate with Copilot Studio (Agents SDK)</a></li>'
             '<li><a href="https://learn.microsoft.com/microsoft-copilot-studio/configuration-end-user-authentication">Configure user authentication in Copilot Studio</a></li>'
             '<li><a href="https://learn.microsoft.com/microsoft-copilot-studio/knowledge-add-sharepoint">Add SharePoint as a knowledge source</a></li>'
             '<li><a href="https://learn.microsoft.com/microsoft-copilot-studio/harnesses-overview">Harnesses in Copilot Studio</a></li>'
             '<li><a href="https://github.com/SharePoint/sp-dev-docs/issues/10775">sp-dev-docs #10775 (SPFx token provider limitation)</a></li></ul></div></div>'
             f'<footer>Built from automated test runs in a Microsoft demo tenant (M365CPI03719864), {c.TESTED_ON}. Tenant identifiers are demo values; '
             'test-user credentials are not included. Preview features may change.</footer></section>')

    o.append('</main><div id="lightbox" role="dialog" aria-label="Screenshot"><figure><img alt=""><figcaption></figcaption></figure></div>')
    o.append(f"<script>{JS}</script></body></html>")
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("".join(o))
    print(f"wrote {OUT} ({os.path.getsize(OUT) / 1024 / 1024:.2f} MB, {len(base._cache)} images)")


if __name__ == "__main__":
    build()
