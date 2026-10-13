"""Single source of truth for the Microsoft 365 Agents SDK tutorial, use cases and test scenarios.

Used by tools/build_sdk_docs.py (markdown, Excel test plan, JSON catalog) and tools/build_sdk_html.py (HTML page).
Counts in the actual results are computed from docs/test-results so the documents cannot drift from the evidence.
Inline markup in strings: **bold**, `code`, [text](url).
"""
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RES = os.path.join(ROOT, "docs", "test-results")
REPO_URL = os.environ.get("REPO_URL", "https://github.com/patmeh1/copilot-studio-agents-sdk-sharepoint")
# Public GitHub Pages site: GitHub Pages serves the docs/ folder of this repository (main branch)
SITE_URL = os.environ.get("SITE_URL", "https://patmeh1.github.io/copilot-studio-agents-sdk-sharepoint/")
TESTED_ON = "7 Oct 2026"
PACKAGE_VERSION = "1.0.3.0"
STACK = "SPFx 1.21.1, `@microsoft/agents-copilotstudio-client` 1.9.1, MSAL.js 4.30, Bot Framework Web Chat 4.19.1"
AUTH_VERIFIED_ON = "13 Oct 2026"

AUTH_CONFIG = [
    ("Application type", "Single-page application (public client)", "Supported"),
    ("OAuth flow", "Authorization code with Proof Key for Code Exchange (PKCE)", "Supported"),
    ("Implicit access-token issuance", "Disabled (`enableAccessTokenIssuance: false`)", "Verified"),
    ("Implicit ID-token issuance", "Disabled (`enableIdTokenIssuance: false`)", "Verified"),
    ("Sign-in audience", "Single tenant (`AzureADMyOrg`)", "Verified"),
    ("SPA redirect URI", "`https://<tenant>.sharepoint.com`", "Verified"),
    ("Copilot Studio permission", "Power Platform API / `CopilotStudio.Copilots.Invoke` (delegated)", "Verified"),
    ("Client credential", "None - a browser public client must not contain a secret", "Verified"),
    ("Authentication library", "MSAL Browser 4.30 (`PublicClientApplication`)", "Verified"),
    ("Current token cache", "`localStorage`; use `sessionStorage` or `memoryStorage` if policy requires less persistence", "Hardenable"),
]

AUTH_STEPS = [
    {
        "title": "SharePoint loads the tenant-wide SPFx web part",
        "body": "The page supplies the current tenant, signed-in SharePoint user and the agent's Microsoft 365 Agents SDK connection string.",
        "result": "No Copilot Studio credential or secret is stored in the page configuration.",
    },
    {
        "title": "The web part initializes MSAL as a public client",
        "body": "MSAL receives the single-tenant client ID, tenant authority and the SharePoint origin as its SPA redirect URI. A public client has no client secret.",
        "result": "Microsoft Entra ID recognizes the request as a browser SPA request.",
    },
    {
        "title": "MSAL first checks its token cache",
        "body": "`acquireTokenSilent` looks for the correct account (matched to the SharePoint UPN) and a usable delegated token.",
        "result": "A valid cached token avoids another round trip and no user interface appears.",
    },
    {
        "title": "MSAL reuses the existing Entra session",
        "body": "If the cache has no usable token, `ssoSilent` starts an authorization-code request in a hidden frame with `login_hint` set to the SharePoint user.",
        "result": "The existing SharePoint / Entra session normally satisfies sign-in and MFA without another prompt.",
    },
    {
        "title": "PKCE binds the authorization code to this browser request",
        "body": "MSAL creates a one-time code verifier and sends its derived challenge to Microsoft Entra ID. The returned authorization code can only be redeemed with that verifier.",
        "result": "A stolen authorization code cannot be redeemed by a different client.",
    },
    {
        "title": "MSAL redeems the code for a delegated Power Platform token",
        "body": "MSAL calls the Microsoft identity platform token endpoint and requests the scope derived from the connection string (for example `https://api.powerplatform.com/.default`).",
        "result": "The token represents the signed-in user and contains the consented `CopilotStudio.Copilots.Invoke` permission.",
    },
    {
        "title": "The Copilot Studio client starts the conversation",
        "body": "The web part gives the access token and connection settings to `CopilotStudioClient`, then bridges its activity stream to Bot Framework Web Chat.",
        "result": "The token is sent only to the Power Platform / Copilot Studio endpoint.",
    },
    {
        "title": "Copilot Studio checks both identity and sharing",
        "body": "Copilot Studio validates the token, identifies the user and enforces whether the agent is shared with that user or one of their groups.",
        "result": "Shared users chat; unshared users receive the agent's access-denied response.",
    },
    {
        "title": "Interactive fallback uses the same secure flow",
        "body": "If silent sign-in is blocked by browser privacy controls or Conditional Access, the user selects **Sign in with Microsoft**. `acquireTokenPopup` still uses authorization code + PKCE.",
        "result": "The popup does not switch to implicit grant. One completed sign-in reconnects the other agent web parts on the page.",
    },
]


def _load(name):
    with open(os.path.join(RES, name), encoding="utf-8") as f:
        return json.load(f)


def _stats():
    summary = _load("summary.json")
    runs = [(who, site, cell) for who in ("a", "b", "c", "d", "mario") for site, cell in summary[who].items()]
    ok = [r for r in runs if r[2]["site"] == "OK"]
    sessions = silent = 0
    for who in ("a", "b", "c", "d", "mario"):
        for site in _load(f"{who}.json")["sites"].values():
            for key in ("agentA_sdk", "agentB_sdk"):
                if key in site:
                    sessions += 1
                    silent += site[key]["status"] == "chat" and "Entra ID SSO · " in site[key]["footer"]
    extra = _load("tc03-tc22-userA-siteA-v1.0.3.json")
    return {
        "runs": len(runs), "pages_ok": len(ok), "pages_denied": len(runs) - len(ok), "sessions": sessions, "silent": silent,
        "a_works": sum(r[2]["A_sdk"] == "WORKS" for r in ok), "a_blocked": sum(r[2]["A_sdk"].startswith("BLOCKED") for r in ok),
        "b_unsupported": sum(r[2]["B_sdk"] == "NOT SUPPORTED" for r in ok),
        "embed_unshared": sum(r[2]["B_embed"].startswith("ANSWERS") for r in ok if r[0] in ("b", "d")),
        "fallback_seconds": extra["tc03_interactive_fallback"]["seconds_to_answer"],
    }


S = _stats()

STATUS = {
    "pass": ("Pass", "ok"),
    "limit": ("Confirmed limitation", "bad"),
    "finding": ("Security finding", "warn"),
    "notrun": ("Not run - validate in your tenant", "muted"),
}
UC_STATUS = {
    "validated": ("Validated", "ok"),
    "design": ("By design - not tested here", "muted"),
    "validate": ("To validate", "warn"),
}

PERSONAS = [
    ("User A", "Agent A + Agent B", "Visitor", "Member", "Owner"),
    ("User B", "Agent A", "Member", "Visitor", "No access"),
    ("User C", "Agent B", "No access", "Member", "Visitor"),
    ("User D", "None", "Visitor", "No access", "Member"),
    ("MarioR (maker)", "Owner of both", "Owner", "Owner", "Owner"),
]

USE_CASES = [
    {"id": "UC-01", "title": "One assistant on every intranet site", "status": "validated",
     "scenario": "Employees open the intranet home, HR, IT or a project site and chat with the same company agent, already signed in.",
     "how": "One tenant-wide SPFx package and one Entra SPA app serve every site. Each page stores only the agent's connection string.",
     "evidence": f"Agent A answered every shared tester on every site they could open ({S['a_works']} of {S['a_works']} runs across Sites A, B and C).",
     "tests": ["TC-11", "TC-12", "TC-05"]},
    {"id": "UC-02", "title": "Only approved people can chat", "status": "validated",
     "scenario": "HR shares the HR agent with the HR team. Anyone else who opens the page is refused, on every site.",
     "how": "Copilot Studio sharing (**End user access**) is enforced on every Agents SDK call. Use Entra security groups in production.",
     "evidence": "Unshared users got *You don't have access to talk to this bot* on every site. Revoking and re-sharing took effect on the next page load.",
     "tests": ["TC-06", "TC-07", "TC-08", "TC-09"]},
    {"id": "UC-03", "title": "The agent knows who is chatting", "status": "validated",
     "scenario": "The agent greets employees by name and can use their identity in topics, for example to look up their own requests.",
     "how": "With **Authenticate with Microsoft**, Copilot Studio receives the signed-in user, so `System.User.*` variables are filled in.",
     "evidence": "Every greeting used the tester's name (`Hello Agents SDK User A!`). The web part footer shows the signed-in account.",
     "tests": ["TC-02"]},
    {"id": "UC-04", "title": "Single sign-on, also with MFA", "status": "validated",
     "scenario": "People already signed in to SharePoint start chatting without another prompt. If the browser blocks silent sign-in, one click fixes it.",
     "how": "MSAL `acquireTokenSilent`, then `ssoSilent` with the SharePoint user as login hint. Fallback: **Sign in with Microsoft** popup; one sign-in serves every agent on the page.",
     "evidence": f"{S['silent']} of {S['sessions']} Agents SDK sessions signed in silently. With silent SSO blocked, one click on the sign-in button led to an answer in about {S['fallback_seconds']:.0f} s.",
     "tests": ["TC-01", "TC-03"]},
    {"id": "UC-05", "title": "Several agents on one page, or one agent per site", "status": "validated",
     "scenario": "A hub page shows an HR agent and an IT agent side by side; each department site shows its own agent.",
     "how": "Every web part instance has its own connection string and conversation.",
     "evidence": "Each test page hosted three agent web parts with independent results.",
     "tests": ["TC-14"]},
    {"id": "UC-06", "title": "No host site to open up", "status": "validated",
     "scenario": "A person may use the agent but must not see the agent owner's site. They still chat from their own site.",
     "how": "The web part calls Copilot Studio directly - there is no `.agent` file or host site in the path.",
     "evidence": "Shared User C with no access to Site A chatted from Site B. The Agent link (SharePoint channel) blocked the same user.",
     "tests": ["TC-13", "TC-10"]},
    {"id": "UC-07", "title": "Employees without Microsoft 365 Copilot", "status": "validated",
     "scenario": "Employees without a Microsoft 365 Copilot license use the agent; usage is billed to Copilot Studio capacity.",
     "how": "Standard-harness agents reached through the Agents SDK consume the tenant's Copilot Studio capacity (Copilot Credits).",
     "evidence": "Every test user had Microsoft 365 E5 (no Teams) only - no Microsoft 365 Copilot or Copilot Studio license.",
     "tests": ["TC-16", "TC-17"]},
    {"id": "UC-08", "title": "Never reachable anonymously", "status": "validated",
     "scenario": "Security requires that the agent cannot be used without an Entra sign-in, even if a URL leaks.",
     "how": "Agents SDK endpoints accept only user tokens for `CopilotStudio.Copilots.Invoke`.",
     "evidence": "Anonymous calls to both connection strings returned HTTP 401 (re-checked on 7 Oct 2026). The hosted embed of the GitHub Copilot harness agent did not enforce sign-in.",
     "tests": ["TC-04", "TC-19"]},
    {"id": "UC-09", "title": "Update once, every site follows", "status": "validated",
     "scenario": "IT ships a new web part version once and every site gets it; makers publish agent changes without touching pages.",
     "how": "Tenant app catalog with **Enable this app and add it to all sites**; agent changes go live when the agent is published.",
     "evidence": f"Package 1.0.1.0 -> {PACKAGE_VERSION} was rolled out tenant-wide without editing a page; the regression run on Site A passed.",
     "tests": ["TC-20", "TC-21"]},
    {"id": "UC-10", "title": "Answers that respect SharePoint permissions", "status": "validate",
     "scenario": "The agent answers from policy documents in SharePoint and only uses documents the user can read.",
     "how": "SharePoint knowledge with Microsoft authentication runs generative answers on behalf of the signed-in user (documented for Teams, Power Apps and Microsoft 365 Copilot).",
     "evidence": "Not tested through the Agents SDK in this engagement - validate before relying on it.",
     "tests": ["TC-26"]},
    {"id": "UC-11", "title": "Agent that knows the page it is on", "status": "design",
     "scenario": "On a project site the agent already knows the project and tailors its answers.",
     "how": "Post a custom event with site/page context after the first agent message and read it in an `OnEventActivity` topic (Microsoft Learn pattern). Needs a small web part extension.",
     "evidence": "Not implemented in this web part yet.",
     "tests": ["TC-25"]},
    {"id": "UC-12", "title": "Same agent in other web apps", "status": "design",
     "scenario": "The agent is also embedded in a line-of-business portal or a native app with the same sign-in.",
     "how": "Copilot Studio client library for JavaScript, .NET or Python with the same connection string and a user token (service principal tokens are not supported).",
     "evidence": "[Integrate with web or native apps](https://learn.microsoft.com/microsoft-copilot-studio/publication-integrate-web-or-native-app-m365-agents-sdk) - not tested here.",
     "tests": []},
]

NOT_SUPPORTED = [
    ("GitHub Copilot harness agents through the Agents SDK",
     "Copilot Studio replies *This action doesn't support agents built with the GitHub Copilot harness.* - verified for every user including the owner. Build SharePoint-facing agents on the standard harness."),
    ("SPFx built-in token provider (`AadTokenProvider`)",
     "Cannot obtain `CopilotStudio.Copilots.Invoke` ([sp-dev-docs #10775](https://github.com/SharePoint/sp-dev-docs/issues/10775)). The web part uses MSAL.js with its own SPA app instead."),
    ("App-only (service principal) tokens",
     "Not supported by the Agents SDK client (Microsoft Learn). Use Direct Line for server-to-server scenarios."),
    ("Hosted web chat embed as an access boundary",
     "Copilot Studio offers the embed code only for agents without authentication. The GitHub Copilot harness embed answered unshared and signed-out users in this test."),
]

TUTORIAL = [
    {"part": "Prepare the agent in Copilot Studio", "steps": [
        {"title": "Create a standard-harness agent",
         "why": "The Copilot Studio client library (Agents SDK) supports agents built on the standard harness.",
         "do": ["Copilot Studio > **Home** > **Other ways to build** > **Agent (Standard)**, enter a name, select **Create**.",
                "Add a description and instructions on the **Overview** page."],
         "check": "The agent opens with **Overview**, **Knowledge**, **Topics** and **Channels** tabs.",
         "images": [("07-create-standard-agent.png", "Home > Other ways to build > Agent (Standard)")]},
        {"title": "Require Microsoft sign-in",
         "why": "The Agents SDK only works with authenticated agents, and it is what makes the agent know the user.",
         "do": ["**Settings** > **Security** > **Authentication** > **Authenticate with Microsoft** (default for new agents) > **Save**."],
         "check": "Channels shows *You're using Microsoft authentication, so only Teams and Microsoft 365 and SharePoint channels are available.*",
         "images": [("10-agentA-authenticate-with-microsoft.png", "Settings > Security > Authentication")]},
        {"title": "Greet the user by name (recommended)",
         "why": "A personalised greeting proves on every page that the user's identity reaches the agent.",
         "do": ["**Topics** > **System** > **Conversation Start** > **...** > **Open code editor**.",
                "Replace only the text of the first message (keep the existing `id`), then **Save**."],
         "code": ("yaml", "    text:\n      - \"Hello {System.User.DisplayName}! I'm {System.Bot.Name}. How can I help?\""),
         "check": "The topic saves without errors. After you publish, the greeting shows the signed-in user's name (seen on every test page).",
         "images": [("09-agentA-greeting-yaml.png", "Conversation Start topic in the code editor")]},
        {"title": "Publish and copy the connection string",
         "why": "Pages connect to the published agent through this string. It contains no secret.",
         "do": ["Select **Publish**.",
                "**Channels** > **Web app** (or **Native app**) > **Microsoft 365 Agents SDK** > **Copy**."],
         "code": ("text", "https://<environment>.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/<schema-name>/conversations?api-version=2022-03-01-preview"),
         "check": "The string ends with `/bots/<agent schema name>/conversations?api-version=...`.",
         "images": [("11-agentA-web-app-connection-string.png", "Channels > Web app > Microsoft 365 Agents SDK")]},
        {"title": "Share the agent",
         "why": "Sharing decides who can chat - on every site and page.",
         "do": ["**...** > **Share** > add users or Entra **security groups** > keep **End user access** > **Share**."],
         "check": "The people or groups appear under **Existing users** as *Viewer*.",
         "images": [("12-agentA-share.png", "Share with users or security groups")]},
    ]},
    {"part": "One-time tenant setup", "steps": [
        {"title": "Register the Entra SPA app",
         "why": "SPFx's built-in token provider cannot get the Copilot Studio scope, so the web part signs in with its own app.",
         "do": ["Copy `scripts/variables.sample.ps1` to `scripts/variables.ps1` and fill it in.",
                "Run the script (idempotent). It adds the SPA redirect URI (your SharePoint origin), the delegated **Power Platform API > CopilotStudio.Copilots.Invoke** permission and tenant-wide consent."],
         "code": ("powershell", ". .\\scripts\\variables.ps1\naz login --tenant $TenantDomain --allow-no-subscriptions\n.\\scripts\\New-AgentsSdkSpaApp.ps1 -TenantId $TenantDomain -SharePointRootUrl $SharePointRootUrl"),
         "check": "The output shows `\"adminConsent\": \"granted\"`; note the `appClientId`.",
         "images": []},
        {"title": "Build the SharePoint package",
         "why": "The package carries the app ID as tenant default, so page authors only paste a connection string.",
         "do": ["Node 22 is required for SPFx 1.21.1. MSAL, the Agents SDK client and Web Chat are bundled - no CDN or API-access approvals."],
         "code": ("powershell", ".\\scripts\\Build-SpfxPackage.ps1 -AppClientId $AppClientId -Version $PackageVersion -NodePath $Node22Path\n# -> spfx-copilot-agent\\sharepoint\\solution\\copilot-studio-agent-webpart.sppkg"),
         "check": "`copilot-studio-agent-webpart.sppkg` is created.",
         "images": []},
        {"title": "Deploy it to every site",
         "why": "One deployment makes the web part available on all current and future sites.",
         "do": ["SharePoint admin center > **Active sites** > **Apps** > **Membership** > add yourself as **site admin** of the app catalog.",
                "App catalog > **Manage apps** > **Upload** the `.sppkg` (**Replace** for updates) > **Enable this app and add it to all sites** > **Enable app**."],
         "check": "Manage apps lists the package as *Enabled* and *Added to all sites: Yes*.",
         "images": [("16-enable-app.png", "Enable this app and add it to all sites"),
                    ("43-app-catalog-v1.0.3-enabled.png", "Version 1.0.3.0 enabled for all sites")]},
    ]},
    {"part": "Add the agent to your sites", "steps": [
        {"title": "Add the web part and paste the connection string",
         "why": "Each page decides which agent it shows; repeat on as many sites as you like.",
         "do": ["Edit the page > **+** > **Copilot Studio Agent (Agents SDK)**.",
                "**Edit properties** > Integration mode **Microsoft 365 Agents SDK** > paste the connection string. Leave app and tenant ID empty (package defaults).",
                "**Republish** the page."],
         "check": "The chat shows your personalised greeting, even in edit mode.",
         "images": [("18-webpart-toolbox.png", "Toolbox: Copilot Studio Agent (Agents SDK)"),
                    ("19-webpart-property-pane.png", "Property pane: paste the connection string")]},
        {"title": "Give people access to the site",
         "why": "SharePoint still decides who can open the page. Visitor (Read) is enough.",
         "do": ["Grant site or page access as usual. Agent sharing and site permissions are independent - people need both."],
         "check": "A test user opens the page and sees the chat.",
         "images": [("20-site-page-overview.png", "Published page with two agents")]},
    ]},
    {"part": "Verify", "steps": [
        {"title": "Run the core test scenarios",
         "why": "Five checks prove sign-in, identity, access control and the security boundary.",
         "do": ["Shared user: greeting by name and an answer (**TC-01**, **TC-02**, **TC-05**).",
                "Unshared user: *You don't have access to talk to this bot* (**TC-06**).",
                "Anonymous call to the connection string returns HTTP 401 (**TC-04**)."],
         "code": ("powershell", "Invoke-WebRequest -Method Post -Uri $AgentConnectionString -ContentType 'application/json' `\n  -Body '{\"emitStartConversationEvent\":true}' -SkipHttpErrorCheck | Select-Object StatusCode   # 401"),
         "check": "All five results match the expected results in the test scenarios.",
         "images": [("21-userA-agentA-sdk-works.png", "Shared user: personalised greeting and answer"),
                    ("22-userD-agentA-sdk-not-shared.png", "Unshared user: blocked by Copilot Studio")]},
        {"title": "Run the resilience checks",
         "why": "Confirms the experience when silent sign-in is blocked, access changes, or a user restarts.",
         "do": ["Silent sign-in blocked: **Sign in with Microsoft** opens the popup on the first click (**TC-03**).",
                "Remove and re-add a user on the agent (**TC-07**, **TC-08**); select **New chat** (**TC-22**)."],
         "check": "The chat connects after one click and access changes apply on the next page load.",
         "images": [("40-tc03-sign-in-prompt.png", "Silent sign-in blocked: one-click sign-in"),
                    ("41-tc03-after-popup-sign-in.png", "After the popup: greeting and answer")]},
        {"title": "Automate the regression run (optional)",
         "why": "Re-run the same scenarios after every agent or package change.",
         "do": ["Edit `tests/agents-sdk-matrix/settings.py`, then run the harness (Edge + Playwright, one profile per test user)."],
         "code": ("powershell", "cd tests\\agents-sdk-matrix\n.\\launch-edge.ps1\nforeach ($p in 'mario','a','b','c','d') { python run_matrix.py $p }\npython run_extra_scenarios.py a A      # TC-03 sign-in fallback + TC-22 New chat\npython summarize.py"),
         "check": "`output/results/summary.md` shows the expected result for every persona and site.",
         "images": []},
    ]},
]

AREAS = ["Sign-in and identity", "Access control", "Sites and pages", "Licensing and capacity", "Agent types",
         "Lifecycle and operations", "Extended scenarios"]

TESTS = [
    # Sign-in and identity
    {"id": "TC-01", "area": AREAS[0], "priority": "P1", "title": "Silent single sign-on for MFA-protected users",
     "pre": "User signed in to SharePoint (MFA satisfied); agent shared with the user.",
     "steps": ["Open the page with the web part.", "Watch the web part while it loads."],
     "expected": "The chat connects without a prompt or popup; the footer shows `Microsoft Entra ID SSO · <user>`.",
     "actual": f"{S['silent']} of {S['sessions']} Agents SDK sessions ({S['pages_ok']} page loads, 5 personas) connected silently with the user's account in the footer.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["images/21-userA-agentA-sdk-works.png", "test-results/a.json"]},
    {"id": "TC-02", "area": AREAS[0], "priority": "P2", "title": "The agent knows who is chatting",
     "pre": "Conversation Start greets with `{System.User.DisplayName}`.",
     "steps": ["Open the page.", "Read the greeting."],
     "expected": "The greeting contains the user's display name.",
     "actual": "`Hello Agents SDK User A! ...` - every shared tester and the maker were greeted by name on every site.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["images/21-userA-agentA-sdk-works.png"]},
    {"id": "TC-03", "area": AREAS[0], "priority": "P1", "title": "One-click sign-in when silent SSO is blocked",
     "pre": "Silent SSO cannot complete (third-party cookies blocked, Conditional Access needing interaction, or an embedding host). Simulated by clearing the MSAL cache and blocking the hidden sign-in frame; two Agents SDK web parts on the page.",
     "steps": ["Open the page and wait for *Sign in to chat with this agent*.", "Select **Sign in with Microsoft** right away.",
               "Complete the popup if prompted, then ask a question."],
     "expected": "The popup opens on the first click, the chat connects and answers, and the other agent on the page connects without a second sign-in.",
     "actual": f"Pass on {PACKAGE_VERSION}: answer {S['fallback_seconds']} s after the click; the second web part reconnected by itself. **Defect found and fixed:** on 1.0.1.0 the first click was blocked by the popup blocker while another web part was still trying silent sign-in.",
     "status": "pass", "auto": "run_extra_scenarios.py",
     "evidence": ["images/40-tc03-sign-in-prompt.png", "images/41-tc03-after-popup-sign-in.png", "images/39-tc03-v1.0.1-popup-blocked.png",
                  "test-results/tc03-tc22-userA-siteA-v1.0.3.json"]},
    {"id": "TC-04", "area": AREAS[0], "priority": "P1", "title": "No anonymous access to the agent endpoint",
     "pre": "The agent's connection string.",
     "steps": ["POST to the connection string without an `Authorization` header (command in tutorial step 11)."],
     "expected": "HTTP 401.",
     "actual": "HTTP 401 for Agent A and Agent B (re-checked on 7 Oct 2026).",
     "status": "pass", "auto": "Manual (one command)", "evidence": []},
    # Access control
    {"id": "TC-05", "area": AREAS[1], "priority": "P1", "title": "Shared users can chat",
     "pre": "Agent shared with the user (End user access), or the user owns it.",
     "steps": ["Open the page.", "Send *Give me one tip for running shorter meetings.*"],
     "expected": "Greeting and an on-topic answer (`Agent A (standard harness): ...`).",
     "actual": f"{S['a_works']} of {S['a_works']} runs: User A on Sites A-C, User B on Sites A-B, MarioR on Sites A-C.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["images/21-userA-agentA-sdk-works.png", "test-results/summary.md"]},
    {"id": "TC-06", "area": AREAS[1], "priority": "P1", "title": "Unshared users are blocked",
     "pre": "Agent not shared with the user; the user can open the page.",
     "steps": ["Open the page.", "Send a question."],
     "expected": "*You don't have access to talk to this bot, contact the owner.* and no answer.",
     "actual": f"{S['a_blocked']} of {S['a_blocked']} runs blocked: User C on Sites B-C, User D on Sites A and C.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["images/22-userD-agentA-sdk-not-shared.png"]},
    {"id": "TC-07", "area": AREAS[1], "priority": "P1", "title": "Removing a user from sharing blocks them",
     "pre": "User B is shared and can chat.",
     "steps": ["Remove User B from the agent's sharing.", "Reload the page about 2 minutes later."],
     "expected": "The blocked message instead of an answer.",
     "actual": "*You don't have access to talk to this bot, contact the owner.*",
     "status": "pass", "auto": "Manual change + run_matrix.py",
     "evidence": ["images/28-userB-after-revoke.png", "test-results/b-after-revocation.json"]},
    {"id": "TC-08", "area": AREAS[1], "priority": "P2", "title": "Re-sharing restores access",
     "pre": "User B was removed (TC-07).",
     "steps": ["Share the agent with User B again.", "Reload the page."],
     "expected": "Greeting by name and an answer.",
     "actual": "`Hello Agents SDK User B! ...` and an answer.",
     "status": "pass", "auto": "Manual change + run_matrix.py", "evidence": ["test-results/b-after-regrant.json"]},
    {"id": "TC-09", "area": AREAS[1], "priority": "P1", "title": "Access through an Entra security group",
     "pre": "Agent shared with a security group; the test user is a member and not shared individually.",
     "steps": ["Open the page as the member.", "Remove the user from the group, wait for the change to apply, reload."],
     "expected": "The member can chat; after removal the user is blocked.",
     "actual": "Not run - individual sharing was tested. Recommended before go-live because production access should be group-based.",
     "status": "notrun", "auto": "run_matrix.py (after the group change)", "evidence": []},
    {"id": "TC-10", "area": AREAS[1], "priority": "P1", "title": "Agent sharing and site permissions stay independent",
     "pre": "One shared user without site access; one unshared user with site access.",
     "steps": ["Open the page as the shared user without site access.", "Open it as the unshared site member."],
     "expected": "1: SharePoint *You need access* - the web part never loads. 2: The page opens and the agent refuses.",
     "actual": f"{S['pages_denied']} access-denied runs (User B Site C, User C Site A, User D Site B) and {S['a_blocked']} agent-blocked runs.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["images/25-site-access-denied.png", "test-results/summary.md"]},
    # Sites and pages
    {"id": "TC-11", "area": AREAS[2], "priority": "P1", "title": "One package and one Entra app serve many sites",
     "pre": "Web part deployed tenant-wide.",
     "steps": ["Add the web part with the same connection string on Sites A, B and C.", "Open each page as a shared user."],
     "expected": "The agent works on every site without per-site app registrations or packages.",
     "actual": "Agent A worked on all three sites with one package and one SPA app registration.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["images/20-site-page-overview.png"]},
    {"id": "TC-12", "area": AREAS[2], "priority": "P2", "title": "Visitor (Read) permission is enough",
     "pre": "Shared users with Visitor, Member or Owner permission.",
     "steps": ["Open the page with each permission level."],
     "expected": "The chat works for every level.",
     "actual": "User A: Visitor (Site A), Member (B), Owner (C). User B: Member (A), Visitor (B). All worked.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["test-results/summary.md"]},
    {"id": "TC-13", "area": AREAS[2], "priority": "P2", "title": "No host-site dependency",
     "pre": "Agent shared with User C, who has no access to Site A.",
     "steps": ["Open the Site B page as User C.", "Ask a question."],
     "expected": "The chat works; no access to any other site is needed.",
     "actual": "Greeting and answer. The Agent link (SharePoint channel) blocked the same user with *You need a Microsoft 365 Copilot license...*.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["images/37-sdk-userC-no-site-a-works.png", "test-results/c-sdk-shared-no-site-a.json"]},
    {"id": "TC-14", "area": AREAS[2], "priority": "P3", "title": "Several agents on one page",
     "pre": "Two or more web parts with different connection strings on one page.",
     "steps": ["Open the page."],
     "expected": "Each web part connects to its own agent and behaves independently.",
     "actual": "Every test page hosted three agent web parts; Agent A answered while Agent B returned its harness message.",
     "status": "pass", "auto": "run_matrix.py", "evidence": ["images/20-site-page-overview.png"]},
    {"id": "TC-15", "area": AREAS[2], "priority": "P2", "title": "Page opened in Teams, Viva Connections or the mobile app",
     "pre": "Applies when these entry points are used.",
     "steps": ["Open the page inside the host.", "Chat."],
     "expected": "The chat connects silently, or the one-click sign-in appears and its popup is allowed by the host.",
     "actual": "Not run. Embedded hosts can block the hidden sign-in frame or popups - validate before announcing these entry points.",
     "status": "notrun", "auto": "Manual", "evidence": []},
    # Licensing and capacity
    {"id": "TC-16", "area": AREAS[3], "priority": "P2", "title": "Users without Microsoft 365 Copilot licenses",
     "pre": "Test users without Microsoft 365 Copilot or Copilot Studio licenses.",
     "steps": ["Chat as these users."],
     "expected": "Chat works.",
     "actual": "Every test user had Microsoft 365 E5 (no Teams) only; all shared users chatted.",
     "status": "pass", "auto": "run_matrix.py", "evidence": []},
    {"id": "TC-17", "area": AREAS[3], "priority": "P2", "title": "Usage is billed to Copilot Studio capacity",
     "pre": "A few completed conversations.",
     "steps": ["Open the environment's Copilot Studio consumption (Copilot Credits) in the Power Platform admin center."],
     "expected": "Agents SDK conversations appear as Copilot Studio consumption.",
     "actual": "Not run.",
     "status": "notrun", "auto": "Manual", "evidence": []},
    # Agent types
    {"id": "TC-18", "area": AREAS[4], "priority": "P3", "title": "GitHub Copilot harness agent through the Agents SDK",
     "pre": "Agent built on the GitHub Copilot harness, shared with the user.",
     "steps": ["Open a page whose web part points to that agent."],
     "expected": "Not supported (Microsoft Learn: the client library supports standard-harness agents).",
     "actual": f"*This action doesn't support agents built with the GitHub Copilot harness.* - {S['b_unsupported']} of {S['b_unsupported']} page loads, including the owner.",
     "status": "limit", "auto": "run_matrix.py", "evidence": ["images/23-agentB-sdk-not-supported.png"]},
    {"id": "TC-19", "area": AREAS[4], "priority": "P1", "title": "Hosted web chat embed is not an access boundary (comparison)",
     "pre": "GitHub Copilot harness agent set to Authenticate with Microsoft, App-channel embed on a page.",
     "steps": ["Open the page as a user the agent is not shared with.", "Open the embed URL in a signed-out browser."],
     "expected": "Users without access are refused.",
     "actual": f"The embed answered unshared users ({S['embed_unshared']} page loads) and a signed-out browser. Use the Agents SDK path for access-controlled agents.",
     "status": "finding", "auto": "run_matrix.py", "evidence": ["images/24-userB-agentB-embed-answers.png", "images/26-embed-anonymous-browser.png"]},
    # Lifecycle and operations
    {"id": "TC-20", "area": AREAS[5], "priority": "P2", "title": "A web part update reaches every site",
     "pre": "Web part deployed tenant-wide.",
     "steps": ["Increase the version, rebuild and upload with **Replace** > **Enable app**.", "Reload a page on any site."],
     "expected": "The new version runs on every site without editing pages.",
     "actual": f"1.0.1.0 -> 1.0.2.0 -> {PACKAGE_VERSION} on 7 Oct 2026: pages loaded the new bundle right away and the regression run passed.",
     "status": "pass", "auto": "Manual + run_matrix.py", "evidence": ["images/43-app-catalog-v1.0.3-enabled.png", "test-results/regression-v1.0.3-userA-siteA.json"]},
    {"id": "TC-21", "area": AREAS[5], "priority": "P2", "title": "Published agent changes reach every site",
     "pre": "Agent in use on several pages.",
     "steps": ["Change the agent (for example the greeting) and publish.", "Select **New chat** on any page."],
     "expected": "The next conversation uses the change; no page or package change is needed.",
     "actual": "Not run.",
     "status": "notrun", "auto": "Manual", "evidence": []},
    {"id": "TC-22", "area": AREAS[5], "priority": "P3", "title": "New chat starts a fresh conversation",
     "pre": "An ongoing conversation.",
     "steps": ["Select **New chat**.", "Ask a question."],
     "expected": "Transcript cleared, greeting shown again, the agent answers.",
     "actual": f"Pass on {PACKAGE_VERSION}.",
     "status": "pass", "auto": "run_extra_scenarios.py", "evidence": ["images/42-tc22-new-chat.png", "test-results/tc03-tc22-userA-siteA-v1.0.3.json"]},
    {"id": "TC-23", "area": AREAS[5], "priority": "P3", "title": "Conversation longer than the token lifetime",
     "pre": "A conversation left open for more than about an hour.",
     "steps": ["Send a message after the access token has expired."],
     "expected": "If the message fails, **New chat** signs in again silently and the chat continues.",
     "actual": "Not run. Known design limit: the client receives a token, not a token provider.",
     "status": "notrun", "auto": "Manual", "evidence": []},
    # Extended scenarios
    {"id": "TC-24", "area": AREAS[6], "priority": "P3", "title": "Clear error for a wrong connection string",
     "pre": "Web part with a mistyped agent schema name or an unpublished agent.",
     "steps": ["Open the page."],
     "expected": "A readable error, for example *Agent not found (404). Verify the connection string / schema name and that the agent is published.*",
     "actual": "Not run.",
     "status": "notrun", "auto": "Manual", "evidence": []},
    {"id": "TC-25", "area": AREAS[6], "priority": "P3", "title": "Page context reaches the agent (extension)",
     "pre": "Web part extended to post a context event after the first message; `OnEventActivity` topic stores it.",
     "steps": ["Open the page on a project site.", "Ask a project question."],
     "expected": "The agent uses the site or page context.",
     "actual": "Not run - not implemented in this web part.",
     "status": "notrun", "auto": "Manual", "evidence": []},
    {"id": "TC-26", "area": AREAS[6], "priority": "P1", "title": "SharePoint knowledge respects the user's permissions",
     "pre": "Applies when the agent uses SharePoint knowledge. User X can read a document, User Y cannot.",
     "steps": ["Ask X and Y the same question about that document."],
     "expected": "X gets a grounded answer; Y gets no content from the document.",
     "actual": "Not run through the Agents SDK (documented by Microsoft for Teams, Power Apps and Microsoft 365 Copilot).",
     "status": "notrun", "auto": "Manual", "evidence": []},
    {"id": "TC-27", "area": AREAS[6], "priority": "P3", "title": "Guest (B2B) users",
     "pre": "Applies when guests use the sites. A guest with page access and agent sharing.",
     "steps": ["Open the page as the guest.", "Chat."],
     "expected": "The guest signs in (silently or with one click) and chats if shared; otherwise the agent refuses.",
     "actual": "Not run.",
     "status": "notrun", "auto": "Manual", "evidence": []},
]

CORE_SET = ["TC-01", "TC-02", "TC-04", "TC-05", "TC-06"]


def counts():
    out = {k: 0 for k in STATUS}
    for t in TESTS:
        out[t["status"]] += 1
    return out


def check():
    ids = [t["id"] for t in TESTS]
    assert len(ids) == len(set(ids)), "duplicate test id"
    for uc in USE_CASES:
        for ref in uc["tests"]:
            assert ref in ids, f"{uc['id']} references unknown {ref}"
    for t in TESTS:
        for ev in t["evidence"]:
            assert os.path.exists(os.path.join(ROOT, "docs", ev)), f"{t['id']} evidence missing: {ev}"
    for part in TUTORIAL:
        for st in part["steps"]:
            for img, _ in st["images"]:
                assert os.path.exists(os.path.join(ROOT, "docs", "images", img)), f"tutorial image missing: {img}"
    return True
