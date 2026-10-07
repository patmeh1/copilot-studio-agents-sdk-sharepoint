# Microsoft 365 Agents SDK in SharePoint - use cases and test scenarios

> Tested on **7 Oct 2026** in a Microsoft demo tenant with 4 test users, 3 sites, 2 agents and the maker as control (web part 1.0.3.0). **27 test scenarios:** 16 passed, 1 confirmed limitation, 1 security finding, 9 to validate in your tenant.
> [Tutorial](Agents-SDK-SharePoint-Tutorial.md) · [Test plan (Excel)](Agents-SDK-SharePoint-Test-Plan.xlsx) · [Web page](https://patmeh1.github.io/copilot-studio-agents-sdk-sharepoint/Agents-SDK-SharePoint-Tutorial-and-Test-Scenarios.html) · [Full guide and findings](Agents-SDK-SharePoint-Guide.md)

## Use cases

| ID | Use case | Status | Test scenarios |
|---|---|---|---|
| [UC-01](#uc-01-one-assistant-on-every-intranet-site) | One assistant on every intranet site | ✅ Validated | [TC-11](#tc-11-one-package-and-one-entra-app-serve-many-sites), [TC-12](#tc-12-visitor-read-permission-is-enough), [TC-05](#tc-05-shared-users-can-chat) |
| [UC-02](#uc-02-only-approved-people-can-chat) | Only approved people can chat | ✅ Validated | [TC-06](#tc-06-unshared-users-are-blocked), [TC-07](#tc-07-removing-a-user-from-sharing-blocks-them), [TC-08](#tc-08-re-sharing-restores-access), [TC-09](#tc-09-access-through-an-entra-security-group) |
| [UC-03](#uc-03-the-agent-knows-who-is-chatting) | The agent knows who is chatting | ✅ Validated | [TC-02](#tc-02-the-agent-knows-who-is-chatting) |
| [UC-04](#uc-04-single-sign-on-also-with-mfa) | Single sign-on, also with MFA | ✅ Validated | [TC-01](#tc-01-silent-single-sign-on-for-mfa-protected-users), [TC-03](#tc-03-one-click-sign-in-when-silent-sso-is-blocked) |
| [UC-05](#uc-05-several-agents-on-one-page-or-one-agent-per-site) | Several agents on one page, or one agent per site | ✅ Validated | [TC-14](#tc-14-several-agents-on-one-page) |
| [UC-06](#uc-06-no-host-site-to-open-up) | No host site to open up | ✅ Validated | [TC-13](#tc-13-no-host-site-dependency), [TC-10](#tc-10-agent-sharing-and-site-permissions-stay-independent) |
| [UC-07](#uc-07-employees-without-microsoft-365-copilot) | Employees without Microsoft 365 Copilot | ✅ Validated | [TC-16](#tc-16-users-without-microsoft-365-copilot-licenses), [TC-17](#tc-17-usage-is-billed-to-copilot-studio-capacity) |
| [UC-08](#uc-08-never-reachable-anonymously) | Never reachable anonymously | ✅ Validated | [TC-04](#tc-04-no-anonymous-access-to-the-agent-endpoint), [TC-19](#tc-19-hosted-web-chat-embed-is-not-an-access-boundary-comparison) |
| [UC-09](#uc-09-update-once-every-site-follows) | Update once, every site follows | ✅ Validated | [TC-20](#tc-20-a-web-part-update-reaches-every-site), [TC-21](#tc-21-published-agent-changes-reach-every-site) |
| [UC-10](#uc-10-answers-that-respect-sharepoint-permissions) | Answers that respect SharePoint permissions | 🔎 To validate | [TC-26](#tc-26-sharepoint-knowledge-respects-the-users-permissions) |
| [UC-11](#uc-11-agent-that-knows-the-page-it-is-on) | Agent that knows the page it is on | 📘 By design - not tested here | [TC-25](#tc-25-page-context-reaches-the-agent-extension) |
| [UC-12](#uc-12-same-agent-in-other-web-apps) | Same agent in other web apps | 📘 By design - not tested here | - |

### UC-01 One assistant on every intranet site

* **Scenario:** Employees open the intranet home, HR, IT or a project site and chat with the same company agent, already signed in.
* **How the Agents SDK web part does it:** One tenant-wide SPFx package and one Entra SPA app serve every site. Each page stores only the agent's connection string.
* **Evidence:** Agent A answered every shared tester on every site they could open (8 of 8 runs across Sites A, B and C).
* **Status:** Validated

### UC-02 Only approved people can chat

* **Scenario:** HR shares the HR agent with the HR team. Anyone else who opens the page is refused, on every site.
* **How the Agents SDK web part does it:** Copilot Studio sharing (**End user access**) is enforced on every Agents SDK call. Use Entra security groups in production.
* **Evidence:** Unshared users got *You don't have access to talk to this bot* on every site. Revoking and re-sharing took effect on the next page load.
* **Status:** Validated

### UC-03 The agent knows who is chatting

* **Scenario:** The agent greets employees by name and can use their identity in topics, for example to look up their own requests.
* **How the Agents SDK web part does it:** With **Authenticate with Microsoft**, Copilot Studio receives the signed-in user, so `System.User.*` variables are filled in.
* **Evidence:** Every greeting used the tester's name (`Hello Agents SDK User A!`). The web part footer shows the signed-in account.
* **Status:** Validated

### UC-04 Single sign-on, also with MFA

* **Scenario:** People already signed in to SharePoint start chatting without another prompt. If the browser blocks silent sign-in, one click fixes it.
* **How the Agents SDK web part does it:** MSAL `acquireTokenSilent`, then `ssoSilent` with the SharePoint user as login hint. Fallback: **Sign in with Microsoft** popup; one sign-in serves every agent on the page.
* **Evidence:** 24 of 24 Agents SDK sessions signed in silently. With silent SSO blocked, one click on the sign-in button led to an answer in about 20 s.
* **Status:** Validated

### UC-05 Several agents on one page, or one agent per site

* **Scenario:** A hub page shows an HR agent and an IT agent side by side; each department site shows its own agent.
* **How the Agents SDK web part does it:** Every web part instance has its own connection string and conversation.
* **Evidence:** Each test page hosted three agent web parts with independent results.
* **Status:** Validated

### UC-06 No host site to open up

* **Scenario:** A person may use the agent but must not see the agent owner's site. They still chat from their own site.
* **How the Agents SDK web part does it:** The web part calls Copilot Studio directly - there is no `.agent` file or host site in the path.
* **Evidence:** Shared User C with no access to Site A chatted from Site B. The Agent link (SharePoint channel) blocked the same user.
* **Status:** Validated

### UC-07 Employees without Microsoft 365 Copilot

* **Scenario:** Employees without a Microsoft 365 Copilot license use the agent; usage is billed to Copilot Studio capacity.
* **How the Agents SDK web part does it:** Standard-harness agents reached through the Agents SDK consume the tenant's Copilot Studio capacity (Copilot Credits).
* **Evidence:** Every test user had Microsoft 365 E5 (no Teams) only - no Microsoft 365 Copilot or Copilot Studio license.
* **Status:** Validated

### UC-08 Never reachable anonymously

* **Scenario:** Security requires that the agent cannot be used without an Entra sign-in, even if a URL leaks.
* **How the Agents SDK web part does it:** Agents SDK endpoints accept only user tokens for `CopilotStudio.Copilots.Invoke`.
* **Evidence:** Anonymous calls to both connection strings returned HTTP 401 (re-checked on 7 Oct 2026). The hosted embed of the GitHub Copilot harness agent did not enforce sign-in.
* **Status:** Validated

### UC-09 Update once, every site follows

* **Scenario:** IT ships a new web part version once and every site gets it; makers publish agent changes without touching pages.
* **How the Agents SDK web part does it:** Tenant app catalog with **Enable this app and add it to all sites**; agent changes go live when the agent is published.
* **Evidence:** Package 1.0.1.0 -> 1.0.3.0 was rolled out tenant-wide without editing a page; the regression run on Site A passed.
* **Status:** Validated

### UC-10 Answers that respect SharePoint permissions

* **Scenario:** The agent answers from policy documents in SharePoint and only uses documents the user can read.
* **How the Agents SDK web part does it:** SharePoint knowledge with Microsoft authentication runs generative answers on behalf of the signed-in user (documented for Teams, Power Apps and Microsoft 365 Copilot).
* **Evidence:** Not tested through the Agents SDK in this engagement - validate before relying on it.
* **Status:** To validate

### UC-11 Agent that knows the page it is on

* **Scenario:** On a project site the agent already knows the project and tailors its answers.
* **How the Agents SDK web part does it:** Post a custom event with site/page context after the first agent message and read it in an `OnEventActivity` topic (Microsoft Learn pattern). Needs a small web part extension.
* **Evidence:** Not implemented in this web part yet.
* **Status:** By design - not tested here

### UC-12 Same agent in other web apps

* **Scenario:** The agent is also embedded in a line-of-business portal or a native app with the same sign-in.
* **How the Agents SDK web part does it:** Copilot Studio client library for JavaScript, .NET or Python with the same connection string and a user token (service principal tokens are not supported).
* **Evidence:** [Integrate with web or native apps](https://learn.microsoft.com/microsoft-copilot-studio/publication-integrate-web-or-native-app-m365-agents-sdk) - not tested here.
* **Status:** By design - not tested here

### Not supported

* **GitHub Copilot harness agents through the Agents SDK** - Copilot Studio replies *This action doesn't support agents built with the GitHub Copilot harness.* - verified for every user including the owner. Build SharePoint-facing agents on the standard harness.
* **SPFx built-in token provider (`AadTokenProvider`)** - Cannot obtain `CopilotStudio.Copilots.Invoke` ([sp-dev-docs #10775](https://github.com/SharePoint/sp-dev-docs/issues/10775)). The web part uses MSAL.js with its own SPA app instead.
* **App-only (service principal) tokens** - Not supported by the Agents SDK client (Microsoft Learn). Use Direct Line for server-to-server scenarios.
* **Hosted web chat embed as an access boundary** - Copilot Studio offers the embed code only for agents without authentication. The GitHub Copilot harness embed answered unshared and signed-out users in this test.

## Test setup

Agent A is a standard-harness agent shared with User A and User B. Agent B (Weather) is a GitHub Copilot harness agent shared with User A and User C. User D has no agent. Every site page hosts Agent A and Agent B through the Agents SDK web part plus Agent B's hosted embed for comparison. Test users have Microsoft 365 E5 (no Teams), no Copilot licenses, and MFA.

| Persona | Agent sharing | Site A | Site B | Site C |
|---|---|---|---|---|
| User A | Agent A + Agent B | Visitor | Member | Owner |
| User B | Agent A | Member | Visitor | No access |
| User C | Agent B | No access | Member | Visitor |
| User D | None | Visitor | No access | Member |
| MarioR (maker) | Owner of both | Owner | Owner | Owner |

## Test scenarios

Run the **core set** (TC-01, TC-02, TC-04, TC-05, TC-06) after every change. Priority P1 = must pass before go-live, P2 = should pass, P3 = optional. *Automation* names the script in `tests/agents-sdk-matrix`.

| ID | Area | Scenario | Priority | Result in the demo tenant |
|---|---|---|---|---|
| [TC-01](#tc-01-silent-single-sign-on-for-mfa-protected-users) | Sign-in and identity | Silent single sign-on for MFA-protected users | P1 | ✅ Pass |
| [TC-02](#tc-02-the-agent-knows-who-is-chatting) | Sign-in and identity | The agent knows who is chatting | P2 | ✅ Pass |
| [TC-03](#tc-03-one-click-sign-in-when-silent-sso-is-blocked) | Sign-in and identity | One-click sign-in when silent SSO is blocked | P1 | ✅ Pass |
| [TC-04](#tc-04-no-anonymous-access-to-the-agent-endpoint) | Sign-in and identity | No anonymous access to the agent endpoint | P1 | ✅ Pass |
| [TC-05](#tc-05-shared-users-can-chat) | Access control | Shared users can chat | P1 | ✅ Pass |
| [TC-06](#tc-06-unshared-users-are-blocked) | Access control | Unshared users are blocked | P1 | ✅ Pass |
| [TC-07](#tc-07-removing-a-user-from-sharing-blocks-them) | Access control | Removing a user from sharing blocks them | P1 | ✅ Pass |
| [TC-08](#tc-08-re-sharing-restores-access) | Access control | Re-sharing restores access | P2 | ✅ Pass |
| [TC-09](#tc-09-access-through-an-entra-security-group) | Access control | Access through an Entra security group | P1 | ⬜ Not run - validate in your tenant |
| [TC-10](#tc-10-agent-sharing-and-site-permissions-stay-independent) | Access control | Agent sharing and site permissions stay independent | P1 | ✅ Pass |
| [TC-11](#tc-11-one-package-and-one-entra-app-serve-many-sites) | Sites and pages | One package and one Entra app serve many sites | P1 | ✅ Pass |
| [TC-12](#tc-12-visitor-read-permission-is-enough) | Sites and pages | Visitor (Read) permission is enough | P2 | ✅ Pass |
| [TC-13](#tc-13-no-host-site-dependency) | Sites and pages | No host-site dependency | P2 | ✅ Pass |
| [TC-14](#tc-14-several-agents-on-one-page) | Sites and pages | Several agents on one page | P3 | ✅ Pass |
| [TC-15](#tc-15-page-opened-in-teams-viva-connections-or-the-mobile-app) | Sites and pages | Page opened in Teams, Viva Connections or the mobile app | P2 | ⬜ Not run - validate in your tenant |
| [TC-16](#tc-16-users-without-microsoft-365-copilot-licenses) | Licensing and capacity | Users without Microsoft 365 Copilot licenses | P2 | ✅ Pass |
| [TC-17](#tc-17-usage-is-billed-to-copilot-studio-capacity) | Licensing and capacity | Usage is billed to Copilot Studio capacity | P2 | ⬜ Not run - validate in your tenant |
| [TC-18](#tc-18-github-copilot-harness-agent-through-the-agents-sdk) | Agent types | GitHub Copilot harness agent through the Agents SDK | P3 | ⛔ Confirmed limitation |
| [TC-19](#tc-19-hosted-web-chat-embed-is-not-an-access-boundary-comparison) | Agent types | Hosted web chat embed is not an access boundary (comparison) | P1 | ⚠️ Security finding |
| [TC-20](#tc-20-a-web-part-update-reaches-every-site) | Lifecycle and operations | A web part update reaches every site | P2 | ✅ Pass |
| [TC-21](#tc-21-published-agent-changes-reach-every-site) | Lifecycle and operations | Published agent changes reach every site | P2 | ⬜ Not run - validate in your tenant |
| [TC-22](#tc-22-new-chat-starts-a-fresh-conversation) | Lifecycle and operations | New chat starts a fresh conversation | P3 | ✅ Pass |
| [TC-23](#tc-23-conversation-longer-than-the-token-lifetime) | Lifecycle and operations | Conversation longer than the token lifetime | P3 | ⬜ Not run - validate in your tenant |
| [TC-24](#tc-24-clear-error-for-a-wrong-connection-string) | Extended scenarios | Clear error for a wrong connection string | P3 | ⬜ Not run - validate in your tenant |
| [TC-25](#tc-25-page-context-reaches-the-agent-extension) | Extended scenarios | Page context reaches the agent (extension) | P3 | ⬜ Not run - validate in your tenant |
| [TC-26](#tc-26-sharepoint-knowledge-respects-the-users-permissions) | Extended scenarios | SharePoint knowledge respects the user's permissions | P1 | ⬜ Not run - validate in your tenant |
| [TC-27](#tc-27-guest-b2b-users) | Extended scenarios | Guest (B2B) users | P3 | ⬜ Not run - validate in your tenant |

### Sign-in and identity

#### TC-01 Silent single sign-on for MFA-protected users

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | User signed in to SharePoint (MFA satisfied); agent shared with the user. |
| **Steps** | 1. Open the page with the web part.<br>2. Watch the web part while it loads. |
| **Expected result** | The chat connects without a prompt or popup; the footer shows `Microsoft Entra ID SSO · <user>`. |
| **Result in the demo tenant** | 24 of 24 Agents SDK sessions (12 page loads, 5 personas) connected silently with the user's account in the footer. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [21-userA-agentA-sdk-works.png](images/21-userA-agentA-sdk-works.png) · [a.json](test-results/a.json) |

#### TC-02 The agent knows who is chatting

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | Conversation Start greets with `{System.User.DisplayName}`. |
| **Steps** | 1. Open the page.<br>2. Read the greeting. |
| **Expected result** | The greeting contains the user's display name. |
| **Result in the demo tenant** | `Hello Agents SDK User A! ...` - every shared tester and the maker were greeted by name on every site. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [21-userA-agentA-sdk-works.png](images/21-userA-agentA-sdk-works.png) |

#### TC-03 One-click sign-in when silent SSO is blocked

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | Silent SSO cannot complete (third-party cookies blocked, Conditional Access needing interaction, or an embedding host). Simulated by clearing the MSAL cache and blocking the hidden sign-in frame; two Agents SDK web parts on the page. |
| **Steps** | 1. Open the page and wait for *Sign in to chat with this agent*.<br>2. Select **Sign in with Microsoft** right away.<br>3. Complete the popup if prompted, then ask a question. |
| **Expected result** | The popup opens on the first click, the chat connects and answers, and the other agent on the page connects without a second sign-in. |
| **Result in the demo tenant** | Pass on 1.0.3.0: answer 20.5 s after the click; the second web part reconnected by itself. **Defect found and fixed:** on 1.0.1.0 the first click was blocked by the popup blocker while another web part was still trying silent sign-in. |
| **Status** | ✅ Pass |
| **Automation** | run_extra_scenarios.py |
| **Evidence** | [40-tc03-sign-in-prompt.png](images/40-tc03-sign-in-prompt.png) · [41-tc03-after-popup-sign-in.png](images/41-tc03-after-popup-sign-in.png) · [39-tc03-v1.0.1-popup-blocked.png](images/39-tc03-v1.0.1-popup-blocked.png) · [tc03-tc22-userA-siteA-v1.0.3.json](test-results/tc03-tc22-userA-siteA-v1.0.3.json) |

#### TC-04 No anonymous access to the agent endpoint

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | The agent's connection string. |
| **Steps** | 1. POST to the connection string without an `Authorization` header (command in tutorial step 11). |
| **Expected result** | HTTP 401. |
| **Result in the demo tenant** | HTTP 401 for Agent A and Agent B (re-checked on 7 Oct 2026). |
| **Status** | ✅ Pass |
| **Automation** | Manual (one command) |
| **Evidence** | - |

### Access control

#### TC-05 Shared users can chat

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | Agent shared with the user (End user access), or the user owns it. |
| **Steps** | 1. Open the page.<br>2. Send *Give me one tip for running shorter meetings.* |
| **Expected result** | Greeting and an on-topic answer (`Agent A (standard harness): ...`). |
| **Result in the demo tenant** | 8 of 8 runs: User A on Sites A-C, User B on Sites A-B, MarioR on Sites A-C. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [21-userA-agentA-sdk-works.png](images/21-userA-agentA-sdk-works.png) · [summary.md](test-results/summary.md) |

#### TC-06 Unshared users are blocked

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | Agent not shared with the user; the user can open the page. |
| **Steps** | 1. Open the page.<br>2. Send a question. |
| **Expected result** | *You don't have access to talk to this bot, contact the owner.* and no answer. |
| **Result in the demo tenant** | 4 of 4 runs blocked: User C on Sites B-C, User D on Sites A and C. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [22-userD-agentA-sdk-not-shared.png](images/22-userD-agentA-sdk-not-shared.png) |

#### TC-07 Removing a user from sharing blocks them

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | User B is shared and can chat. |
| **Steps** | 1. Remove User B from the agent's sharing.<br>2. Reload the page about 2 minutes later. |
| **Expected result** | The blocked message instead of an answer. |
| **Result in the demo tenant** | *You don't have access to talk to this bot, contact the owner.* |
| **Status** | ✅ Pass |
| **Automation** | Manual change + run_matrix.py |
| **Evidence** | [28-userB-after-revoke.png](images/28-userB-after-revoke.png) · [b-after-revocation.json](test-results/b-after-revocation.json) |

#### TC-08 Re-sharing restores access

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | User B was removed (TC-07). |
| **Steps** | 1. Share the agent with User B again.<br>2. Reload the page. |
| **Expected result** | Greeting by name and an answer. |
| **Result in the demo tenant** | `Hello Agents SDK User B! ...` and an answer. |
| **Status** | ✅ Pass |
| **Automation** | Manual change + run_matrix.py |
| **Evidence** | [b-after-regrant.json](test-results/b-after-regrant.json) |

#### TC-09 Access through an Entra security group

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | Agent shared with a security group; the test user is a member and not shared individually. |
| **Steps** | 1. Open the page as the member.<br>2. Remove the user from the group, wait for the change to apply, reload. |
| **Expected result** | The member can chat; after removal the user is blocked. |
| **Result in the demo tenant** | Not run - individual sharing was tested. Recommended before go-live because production access should be group-based. |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | run_matrix.py (after the group change) |
| **Evidence** | - |

#### TC-10 Agent sharing and site permissions stay independent

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | One shared user without site access; one unshared user with site access. |
| **Steps** | 1. Open the page as the shared user without site access.<br>2. Open it as the unshared site member. |
| **Expected result** | 1: SharePoint *You need access* - the web part never loads. 2: The page opens and the agent refuses. |
| **Result in the demo tenant** | 3 access-denied runs (User B Site C, User C Site A, User D Site B) and 4 agent-blocked runs. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [25-site-access-denied.png](images/25-site-access-denied.png) · [summary.md](test-results/summary.md) |

### Sites and pages

#### TC-11 One package and one Entra app serve many sites

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | Web part deployed tenant-wide. |
| **Steps** | 1. Add the web part with the same connection string on Sites A, B and C.<br>2. Open each page as a shared user. |
| **Expected result** | The agent works on every site without per-site app registrations or packages. |
| **Result in the demo tenant** | Agent A worked on all three sites with one package and one SPA app registration. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [20-site-page-overview.png](images/20-site-page-overview.png) |

#### TC-12 Visitor (Read) permission is enough

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | Shared users with Visitor, Member or Owner permission. |
| **Steps** | 1. Open the page with each permission level. |
| **Expected result** | The chat works for every level. |
| **Result in the demo tenant** | User A: Visitor (Site A), Member (B), Owner (C). User B: Member (A), Visitor (B). All worked. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [summary.md](test-results/summary.md) |

#### TC-13 No host-site dependency

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | Agent shared with User C, who has no access to Site A. |
| **Steps** | 1. Open the Site B page as User C.<br>2. Ask a question. |
| **Expected result** | The chat works; no access to any other site is needed. |
| **Result in the demo tenant** | Greeting and answer. The Agent link (SharePoint channel) blocked the same user with *You need a Microsoft 365 Copilot license...*. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [37-sdk-userC-no-site-a-works.png](images/37-sdk-userC-no-site-a-works.png) · [c-sdk-shared-no-site-a.json](test-results/c-sdk-shared-no-site-a.json) |

#### TC-14 Several agents on one page

| | |
|---|---|
| **Priority** | P3 |
| **Preconditions** | Two or more web parts with different connection strings on one page. |
| **Steps** | 1. Open the page. |
| **Expected result** | Each web part connects to its own agent and behaves independently. |
| **Result in the demo tenant** | Every test page hosted three agent web parts; Agent A answered while Agent B returned its harness message. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | [20-site-page-overview.png](images/20-site-page-overview.png) |

#### TC-15 Page opened in Teams, Viva Connections or the mobile app

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | Applies when these entry points are used. |
| **Steps** | 1. Open the page inside the host.<br>2. Chat. |
| **Expected result** | The chat connects silently, or the one-click sign-in appears and its popup is allowed by the host. |
| **Result in the demo tenant** | Not run. Embedded hosts can block the hidden sign-in frame or popups - validate before announcing these entry points. |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | Manual |
| **Evidence** | - |

### Licensing and capacity

#### TC-16 Users without Microsoft 365 Copilot licenses

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | Test users without Microsoft 365 Copilot or Copilot Studio licenses. |
| **Steps** | 1. Chat as these users. |
| **Expected result** | Chat works. |
| **Result in the demo tenant** | Every test user had Microsoft 365 E5 (no Teams) only; all shared users chatted. |
| **Status** | ✅ Pass |
| **Automation** | run_matrix.py |
| **Evidence** | - |

#### TC-17 Usage is billed to Copilot Studio capacity

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | A few completed conversations. |
| **Steps** | 1. Open the environment's Copilot Studio consumption (Copilot Credits) in the Power Platform admin center. |
| **Expected result** | Agents SDK conversations appear as Copilot Studio consumption. |
| **Result in the demo tenant** | Not run. |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | Manual |
| **Evidence** | - |

### Agent types

#### TC-18 GitHub Copilot harness agent through the Agents SDK

| | |
|---|---|
| **Priority** | P3 |
| **Preconditions** | Agent built on the GitHub Copilot harness, shared with the user. |
| **Steps** | 1. Open a page whose web part points to that agent. |
| **Expected result** | Not supported (Microsoft Learn: the client library supports standard-harness agents). |
| **Result in the demo tenant** | *This action doesn't support agents built with the GitHub Copilot harness.* - 12 of 12 page loads, including the owner. |
| **Status** | ⛔ Confirmed limitation |
| **Automation** | run_matrix.py |
| **Evidence** | [23-agentB-sdk-not-supported.png](images/23-agentB-sdk-not-supported.png) |

#### TC-19 Hosted web chat embed is not an access boundary (comparison)

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | GitHub Copilot harness agent set to Authenticate with Microsoft, App-channel embed on a page. |
| **Steps** | 1. Open the page as a user the agent is not shared with.<br>2. Open the embed URL in a signed-out browser. |
| **Expected result** | Users without access are refused. |
| **Result in the demo tenant** | The embed answered unshared users (4 page loads) and a signed-out browser. Use the Agents SDK path for access-controlled agents. |
| **Status** | ⚠️ Security finding |
| **Automation** | run_matrix.py |
| **Evidence** | [24-userB-agentB-embed-answers.png](images/24-userB-agentB-embed-answers.png) · [26-embed-anonymous-browser.png](images/26-embed-anonymous-browser.png) |

### Lifecycle and operations

#### TC-20 A web part update reaches every site

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | Web part deployed tenant-wide. |
| **Steps** | 1. Increase the version, rebuild and upload with **Replace** > **Enable app**.<br>2. Reload a page on any site. |
| **Expected result** | The new version runs on every site without editing pages. |
| **Result in the demo tenant** | 1.0.1.0 -> 1.0.2.0 -> 1.0.3.0 on 7 Oct 2026: pages loaded the new bundle right away and the regression run passed. |
| **Status** | ✅ Pass |
| **Automation** | Manual + run_matrix.py |
| **Evidence** | [43-app-catalog-v1.0.3-enabled.png](images/43-app-catalog-v1.0.3-enabled.png) · [regression-v1.0.3-userA-siteA.json](test-results/regression-v1.0.3-userA-siteA.json) |

#### TC-21 Published agent changes reach every site

| | |
|---|---|
| **Priority** | P2 |
| **Preconditions** | Agent in use on several pages. |
| **Steps** | 1. Change the agent (for example the greeting) and publish.<br>2. Select **New chat** on any page. |
| **Expected result** | The next conversation uses the change; no page or package change is needed. |
| **Result in the demo tenant** | Not run. |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | Manual |
| **Evidence** | - |

#### TC-22 New chat starts a fresh conversation

| | |
|---|---|
| **Priority** | P3 |
| **Preconditions** | An ongoing conversation. |
| **Steps** | 1. Select **New chat**.<br>2. Ask a question. |
| **Expected result** | Transcript cleared, greeting shown again, the agent answers. |
| **Result in the demo tenant** | Pass on 1.0.3.0. |
| **Status** | ✅ Pass |
| **Automation** | run_extra_scenarios.py |
| **Evidence** | [42-tc22-new-chat.png](images/42-tc22-new-chat.png) · [tc03-tc22-userA-siteA-v1.0.3.json](test-results/tc03-tc22-userA-siteA-v1.0.3.json) |

#### TC-23 Conversation longer than the token lifetime

| | |
|---|---|
| **Priority** | P3 |
| **Preconditions** | A conversation left open for more than about an hour. |
| **Steps** | 1. Send a message after the access token has expired. |
| **Expected result** | If the message fails, **New chat** signs in again silently and the chat continues. |
| **Result in the demo tenant** | Not run. Known design limit: the client receives a token, not a token provider. |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | Manual |
| **Evidence** | - |

### Extended scenarios

#### TC-24 Clear error for a wrong connection string

| | |
|---|---|
| **Priority** | P3 |
| **Preconditions** | Web part with a mistyped agent schema name or an unpublished agent. |
| **Steps** | 1. Open the page. |
| **Expected result** | A readable error, for example *Agent not found (404). Verify the connection string / schema name and that the agent is published.* |
| **Result in the demo tenant** | Not run. |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | Manual |
| **Evidence** | - |

#### TC-25 Page context reaches the agent (extension)

| | |
|---|---|
| **Priority** | P3 |
| **Preconditions** | Web part extended to post a context event after the first message; `OnEventActivity` topic stores it. |
| **Steps** | 1. Open the page on a project site.<br>2. Ask a project question. |
| **Expected result** | The agent uses the site or page context. |
| **Result in the demo tenant** | Not run - not implemented in this web part. |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | Manual |
| **Evidence** | - |

#### TC-26 SharePoint knowledge respects the user's permissions

| | |
|---|---|
| **Priority** | P1 |
| **Preconditions** | Applies when the agent uses SharePoint knowledge. User X can read a document, User Y cannot. |
| **Steps** | 1. Ask X and Y the same question about that document. |
| **Expected result** | X gets a grounded answer; Y gets no content from the document. |
| **Result in the demo tenant** | Not run through the Agents SDK (documented by Microsoft for Teams, Power Apps and Microsoft 365 Copilot). |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | Manual |
| **Evidence** | - |

#### TC-27 Guest (B2B) users

| | |
|---|---|
| **Priority** | P3 |
| **Preconditions** | Applies when guests use the sites. A guest with page access and agent sharing. |
| **Steps** | 1. Open the page as the guest.<br>2. Chat. |
| **Expected result** | The guest signs in (silently or with one click) and chats if shared; otherwise the agent refuses. |
| **Result in the demo tenant** | Not run. |
| **Status** | ⬜ Not run - validate in your tenant |
| **Automation** | Manual |
| **Evidence** | - |
