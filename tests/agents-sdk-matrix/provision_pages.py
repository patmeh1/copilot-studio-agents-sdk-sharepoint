"""Create or refresh the test page (three web part instances) on every test site, using MarioR's (owner) session.

usage: python provision_pages.py [ABC]
"""
import json
import sys
import uuid

from cdp import session
from settings import AGENT_A, AGENT_B, PAGE_NAME, PERSONAS, SHAREPOINT_ROOT, SITES, WEBPART_ID
from sp import must, rest


def webpart(title, section, zone, factor, mode="sdk", conn="", embed=""):
    iid = str(uuid.uuid4())
    return {
        "controlType": 3, "id": iid,
        "position": {"zoneIndex": zone, "sectionIndex": section, "controlIndex": 1, "sectionFactor": factor, "layoutIndex": 1},
        "webPartId": WEBPART_ID, "emphasis": {}, "zoneGroupMetadata": {"type": 0},
        "webPartData": {
            "id": WEBPART_ID, "instanceId": iid, "title": "Copilot Studio Agent (Agents SDK)", "description": "",
            "serverProcessedContent": {"htmlStrings": {}, "searchablePlainTexts": {}, "imageSources": {}, "links": {}},
            "dataVersion": "1.0",
            "properties": {"title": title, "mode": mode, "connectionString": conn, "embedUrl": embed,
                           "environmentId": "", "agentIdentifier": "", "appClientId": "", "tenantId": "",
                           "height": 600, "useFluentTheme": True, "showHeader": True},
        },
    }


def canvas(letter):
    intro = {
        "controlType": 4, "id": str(uuid.uuid4()),
        "position": {"zoneIndex": 1, "sectionIndex": 1, "controlIndex": 1, "sectionFactor": 12, "layoutIndex": 1},
        "addedFromPersistedData": True,
        "innerHTML": (f"<h2>Site {letter} - Copilot Studio agents via the Microsoft 365 Agents SDK</h2>"
                      "<p>All chats below use the same SPFx web part. In <strong>Agents SDK</strong> mode users are signed in silently "
                      "with Microsoft Entra ID (MSAL SSO) and the web part calls Copilot Studio through "
                      "<strong>@microsoft/agents-copilotstudio-client</strong>. The last chat shows Agent B through the Copilot Studio "
                      "hosted web chat (embed mode) for comparison.</p>"),
    }
    return [intro,
            webpart(AGENT_A["title"], 1, 2, 6, conn=AGENT_A["connectionString"]),
            webpart(AGENT_B["title"], 2, 2, 6, conn=AGENT_B["connectionString"]),
            webpart(AGENT_B["title"] + " via hosted web chat (embed)", 1, 3, 12, mode="embed", embed=AGENT_B["embedUrl"]),
            {"controlType": 0, "pageSettingsSlice": {"isDefaultDescription": True, "isDefaultThumbnail": True}}]


def provision(page, letter):
    site = SHAREPOINT_ROOT + SITES[letter]
    page.goto(site + "/_layouts/15/viewlsts.aspx", wait_until="domcontentloaded")
    page.wait_for_timeout(2500)
    _, data = rest(page, "GET", f"/_api/sitepages/pages?$select=Id,FileName&$filter=FileName eq '{PAGE_NAME}'", site=site)
    existing = data.get("value", []) if isinstance(data, dict) else []
    if existing:
        pid = existing[0]["Id"]
        must(page, "POST", f"/_api/sitepages/pages({pid})/checkoutpage", {}, site=site)
    else:
        pid = must(page, "POST", "/_api/sitepages/pages", {"PageLayoutType": "Article", "PromotedState": 0}, site=site)["Id"]
    title = PAGE_NAME.replace(".aspx", "").replace("-", " ")
    must(page, "POST", f"/_api/sitepages/pages({pid})/savepageasdraft",
         {"Title": title, "CanvasContent1": json.dumps(canvas(letter)), "BannerImageUrl": ""}, site=site)
    must(page, "POST", f"/_api/sitepages/pages({pid})/publish", {}, site=site)
    print(f"Site {letter}: published {site}/SitePages/{PAGE_NAME}")


if __name__ == "__main__":
    if not AGENT_B["embedUrl"]:
        print("warning: AGENT_B_EMBED_URL is not set - the hosted web chat (embed) web part is added without a URL")
    with session(new=True, port=PERSONAS["mario"]) as (b, ctx, page):
        for letter in (sys.argv[1] if len(sys.argv) > 1 else "".join(SITES)):
            provision(page, letter)
        page.close()
