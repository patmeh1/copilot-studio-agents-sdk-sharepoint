# Agents SDK x SharePoint permission test harness

Automates the test matrix described in [`docs/Agents-SDK-SharePoint-Guide.md`](../../docs/Agents-SDK-SharePoint-Guide.md#7-test-matrix-and-results):
4 personas (+ the maker as control) x 3 SharePoint sites x 3 web part instances
(Agent A via Agents SDK, Agent B via Agents SDK, Agent B via hosted embed).

It drives **real Microsoft Edge windows** (one profile per persona) through the Chrome DevTools Protocol with Playwright,
signs test users in (including first-run MFA registration with a software TOTP authenticator), opens every site page,
chats with each agent and records greeting / answer / errors plus a screenshot of every web part.

## Files

| File | Purpose |
|---|---|
| `settings.py` | **Edit first** - tenant URL, sites, agent connection strings/embed URL, personas, site-permission matrix |
| `New-AgentsSdkTestUsers.ps1` | Creates licensed users `agentsdk.usera..d@<tenant>` and writes `test-users.json` to the secrets folder |
| `launch-edge.ps1` | Starts one Edge per persona with `--remote-debugging-port` (9333 maker, 9341-9344 users) |
| `signin.py` | Signs a persona in; registers a TOTP authenticator on first sign-in and stores the seed in the secrets folder |
| `provision_pages.py` | (maker) Creates/republishes `SitePages/Agents-SDK-Test.aspx` with the three web parts on every site |
| `grant_site_access.py` | (maker) Applies the Owner/Member/Visitor matrix from `settings.SITE_ACCESS` |
| `run_matrix.py` | Runs the matrix for one persona -> `output/results/<persona>.json` + `output/shots/*.png` |
| `run_extra_scenarios.py` | TC-03 (silent SSO blocked -> one-click popup sign-in, second web part reconnects) and TC-22 (New chat) -> `output/results/extra-<persona>-site<X>.json` |
| `summarize.py` | Builds `output/results/summary.md` / `summary.json` |
| `cdp.py`, `sp.py`, `chat.py` | Helpers (CDP session, SharePoint REST with request digest, Web Chat driver) |

## Prerequisites

* Python 3.12 with `pip install playwright pyotp` (Playwright drives the installed Edge; no browser download needed)
* Azure CLI signed in to the tenant as a User Administrator (for `New-AgentsSdkTestUsers.ps1`)
* The web part deployed (guide steps 3-5), Agent A and Agent B published and shared as in the guide (step 2)
* Three communication sites owned by the maker account (the URLs in `settings.SITES`)
* A secrets folder **outside the repo**: `$env:AGENTS_SDK_TEST_SECRETS_DIR` (default `~/.agents-sdk-test`)

## Run

```powershell
$env:AGENTS_SDK_TEST_SECRETS_DIR = "$HOME\.agents-sdk-test"
$env:AGENT_B_EMBED_URL = '<iframe src from Copilot Studio > Agent B > Channels > App > Embed code>'   # only for the embed comparison
.\New-AgentsSdkTestUsers.ps1 -TenantDomain contoso.onmicrosoft.com      # once
.\launch-edge.ps1                                                       # sign the maker in manually in the 9333 window (MFA)
foreach ($p in 'a','b','c','d') { python signin.py $p }                 # first run registers TOTP for each user
python provision_pages.py                                               # maker: pages on sites A, B, C
python grant_site_access.py                                             # maker: site permission matrix
foreach ($p in 'mario','a','b','c','d') { python run_matrix.py $p }
python run_extra_scenarios.py a A                                      # TC-03 sign-in fallback + TC-22 New chat (User A, Site A)
python summarize.py                                                     # -> output\results\summary.md
```

The scenarios these scripts automate (and the ones to run by hand) are described in
[`docs/Agents-SDK-SharePoint-Use-Cases-and-Test-Scenarios.md`](../../docs/Agents-SDK-SharePoint-Use-Cases-and-Test-Scenarios.md).

`run_matrix.py <persona> AC` limits the run to sites A and C. Personas run independently (separate Edge profiles), so they can
run in parallel from separate terminals.

## Verdicts

| Verdict | Meaning |
|---|---|
| `WORKS` | Personalised greeting and an answer that contains `settings.EXPECT_A` |
| `BLOCKED - not shared` | The agent answered `You don't have access to talk to this bot, contact the owner.` |
| `NOT SUPPORTED` | `This action doesn't support agents built with the GitHub Copilot harness.` |
| `ANSWERS (no auth enforced)` | The hosted embed returned a forecast (it does so for any user, even unshared ones) |
| `NO EMBED FRAME` | The page has no hosted web chat frame (for example `AGENT_B_EMBED_URL` was not set when the pages were provisioned) |
| `ACCESS DENIED` | SharePoint *You need access* page - no web part loaded |
| `SIGN-IN PROMPT` / `ERROR` | Silent SSO failed / the web part showed an error banner (see `banner` in the JSON) |

> Security: `test-users.json` contains passwords and `test-users-totp.json` contains MFA seeds for the test accounts.
> Keep them in the secrets folder, never in source control, and delete the test users when you are done.
