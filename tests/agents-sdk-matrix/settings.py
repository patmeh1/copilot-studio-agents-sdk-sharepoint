"""Tenant-specific settings for the Agents SDK x SharePoint test harness.

Edit the values below for your tenant. Secrets (test-user passwords and TOTP seeds) are NOT stored in the repo:
they are read from AGENTS_SDK_TEST_SECRETS_DIR (default: ~/.agents-sdk-test).
"""
import os

SHAREPOINT_ROOT = "https://m365cpi03719864.sharepoint.com"
SITES = {"A": "/sites/AgentsSDK-SiteA", "B": "/sites/AgentsSDK-SiteB", "C": "/sites/AgentsSDK-SiteC"}
PAGE_NAME = "Agents-SDK-Test.aspx"
WEBPART_ID = "0790f94f-bcf1-4fb1-8a0d-324acab20e32"  # Copilot Studio Agent (Agents SDK) web part manifest id

AGENT_A = {
    "title": "Agent A - Standard harness",
    "connectionString": "https://afdeda3552fae1459bc07b43e75630.0a.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/new_AgentAStandardharness/conversations?api-version=2022-03-01-preview",
}
AGENT_B = {
    "title": "Agent B - GitHub Copilot harness (Weather)",
    "connectionString": "https://afdeda3552fae1459bc07b43e75630.0a.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/cr9b0_weatheragent_zl4CqE/conversations?api-version=2022-03-01-preview",
    # iframe src of Copilot Studio > agent > Channels > App > Embed code. Not stored in the repository: for GitHub Copilot
    # harness agents this URL opens the agent without any sign-in (guide, section 10). Set AGENT_B_EMBED_URL to provision it.
    "embedUrl": os.environ.get("AGENT_B_EMBED_URL", ""),
}

QUESTION_A = "Give me one tip for running shorter meetings."
QUESTION_B = "What's the weather in Seattle today?"
EXPECT_A = "Agent A (standard harness)"

# persona -> Edge remote-debugging port (one Edge profile per persona, see launch-edge.ps1)
PERSONAS = {"mario": 9333, "a": 9341, "b": 9342, "c": 9343, "d": 9344}

# Labels used in the summary table: (name, agent sharing, {site: permission})
PERSONA_LABELS = {
    "mario": ("MarioR (maker/owner)", "Owner of A + B", {"A": "Owner", "B": "Owner", "C": "Owner"}),
    "a": ("User A", "Agent A + Agent B", {"A": "Visitor (Read)", "B": "Member (Edit)", "C": "Owner (Full control)"}),
    "b": ("User B", "Agent A only", {"A": "Member (Edit)", "B": "Visitor (Read)", "C": "No access"}),
    "c": ("User C", "Agent B only", {"A": "No access", "B": "Member (Edit)", "C": "Visitor (Read)"}),
    "d": ("User D", "No agent access", {"A": "Visitor (Read)", "B": "No access", "C": "Member (Edit)"}),
}

# Site permission matrix applied by grant_site_access.py: persona -> {site: Owner|Member|Visitor|None}
SITE_ACCESS = {
    "a": {"A": "Visitor", "B": "Member", "C": "Owner"},
    "b": {"A": "Member", "B": "Visitor", "C": None},
    "c": {"A": None, "B": "Member", "C": "Visitor"},
    "d": {"A": "Visitor", "B": None, "C": "Member"},
}

SECRETS_DIR = os.environ.get("AGENTS_SDK_TEST_SECRETS_DIR", os.path.join(os.path.expanduser("~"), ".agents-sdk-test"))
OUTPUT_DIR = os.environ.get("AGENTS_SDK_TEST_OUTPUT_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "output"))
SHOTS_DIR = os.path.join(OUTPUT_DIR, "shots")
RESULTS_DIR = os.path.join(OUTPUT_DIR, "results")
