# Copilot Studio agents on SharePoint with the Microsoft 365 Agents SDK

Publish **authenticated** Copilot Studio agents to **any number of SharePoint sites** through the
**Microsoft 365 Agents SDK** (Copilot Studio > Channels > App / Web app > *Microsoft 365 Agents SDK* connection string),
with silent Microsoft Entra ID single sign-on - validated with a 4-user x 3-site x 2-agent permission matrix.

**Web pages: https://patmeh1.github.io/copilot-studio-agents-sdk-sharepoint/** - GitHub Pages serves the `docs/` folder
of this repository (tutorial, use cases and test scenarios, customer use case, Excel test plan and deck). The App-channel
embed URL of the GitHub Copilot harness agent is not published anywhere in the repository because it opens the agent
without sign-in (screenshot 04 is covered, raw results and test settings omit it).

| Start here | |
|---|---|
| [Tutorial, use cases and test scenarios (web page)](https://patmeh1.github.io/copilot-studio-agents-sdk-sharepoint/Agents-SDK-SharePoint-Tutorial-and-Test-Scenarios.html) · [file](docs/Agents-SDK-SharePoint-Tutorial-and-Test-Scenarios.html) | Self-contained page: **Agents SDK tutorial, 12 use cases and 27 test scenarios** with expected and actual results |
| [docs/Agents-SDK-SharePoint-Tutorial.md](docs/Agents-SDK-SharePoint-Tutorial.md) | Tutorial: 13 steps with checkpoints and screenshots |
| [docs/Agents-SDK-SharePoint-Use-Cases-and-Test-Scenarios.md](docs/Agents-SDK-SharePoint-Use-Cases-and-Test-Scenarios.md) | Use cases and test scenarios (preconditions, steps, expected result, result in the demo tenant, evidence) |
| [docs/Agents-SDK-SharePoint-Test-Plan.xlsx](docs/Agents-SDK-SharePoint-Test-Plan.xlsx) | Test plan to run in your tenant - record your results, the summary sheet counts them |
| [Customer use case (web page)](https://patmeh1.github.io/copilot-studio-agents-sdk-sharepoint/Agents-SDK-SharePoint-Use-Case.html) · [file](docs/Agents-SDK-SharePoint-Use-Case.html) | Self-contained page: the customer's SharePoint channel + Agent link setup compared with the Agents SDK web part |
| [docs/Agents-SDK-SharePoint-Guide.md](docs/Agents-SDK-SharePoint-Guide.md) | Step-by-step guide driven by variables, findings, test matrix, troubleshooting |
| [docs/Copilot-Studio-Agents-SDK-SharePoint.pptx](docs/Copilot-Studio-Agents-SDK-SharePoint.pptx) | Presentation of the approach and results |
| [docs/test-results](docs/test-results) | Raw results of every run (`summary.md`, `tc03-tc22-*.json`, `agentlink-*.json`) and the scenario catalog (`agents-sdk-test-catalog.json`) |

## Results in one table

| | Agent A - **standard harness** | Agent B - **GitHub Copilot harness** |
|---|---|---|
| Via Microsoft 365 Agents SDK web part (this repo) | **Works** - SSO, user identity reaches the agent, Copilot Studio sharing enforced on every site, **no host-site dependency** | **Not supported** - `This action doesn't support agents built with the GitHub Copilot harness.` |
| Via SharePoint channel + Agent link web part | **Works** - but every user needs Read on the `.agent` file in the deployment site (Read on just that file is enough) | SharePoint channel not available for this harness |
| Via the App-channel embed code | not offered for authenticated standard agents | Renders, but **no authentication / sharing enforced** (answered unshared users and a signed-out browser) |

Test scenarios: **27** - 16 passed, 1 confirmed limitation (GitHub Copilot harness), 1 security finding (hosted embed),
9 to validate in your tenant. One defect was found and fixed during testing: on pages with several agents the
*Sign in with Microsoft* popup could be blocked (fixed in web part 1.0.2.0, current package **1.0.3.0**).

## Repository layout

```
docs/                         Tutorial, use cases + test scenarios (md, html, xlsx), guide, use-case HTML, deck, screenshots, raw results
scripts/
  variables.sample.ps1        All tenant-specific values in one place
  New-AgentsSdkSpaApp.ps1     Entra SPA app: Power Platform API / CopilotStudio.Copilots.Invoke + consent
  Build-SpfxPackage.ps1       Stamps tenant defaults and builds the .sppkg (Node 22)
spfx-copilot-agent/           SPFx 1.21.1 solution - "Copilot Studio Agent (Agents SDK)" web part
tests/agents-sdk-matrix/      Playwright + Edge harness: permission matrix + extra scenarios (sign-in fallback, New chat)
tools/                        Generators: sdk_content.py (single source for tutorial/use cases/tests), build_sdk_docs.py,
                              build_sdk_html.py, build_usecase_html.py and build_deck.js
```

## Quick start

```powershell
Copy-Item scripts\variables.sample.ps1 scripts\variables.ps1   # fill in
. .\scripts\variables.ps1
az login --tenant $TenantDomain --allow-no-subscriptions
.\scripts\New-AgentsSdkSpaApp.ps1 -TenantId $TenantDomain -SharePointRootUrl $SharePointRootUrl
.\scripts\Build-SpfxPackage.ps1 -AppClientId <appClientId from previous step> -Version 1.0.0.0 -NodePath $Node22Path
# Upload spfx-copilot-agent\sharepoint\solution\copilot-studio-agent-webpart.sppkg to the tenant app catalog,
# "Enable this app and add it to all sites", then add the web part to pages and paste the agent connection string.
```

Full walkthrough: [docs/Agents-SDK-SharePoint-Tutorial.md](docs/Agents-SDK-SharePoint-Tutorial.md).

## Regenerating the documents

```powershell
python tools\build_sdk_docs.py      # tutorial + use cases/test scenarios (md), test plan (xlsx), scenario catalog (json)
python tools\build_sdk_html.py      # docs\Agents-SDK-SharePoint-Tutorial-and-Test-Scenarios.html
python tools\build_usecase_html.py  # docs\Agents-SDK-SharePoint-Use-Case.html
$env:NODE_PATH = (npm root -g); node tools\build_deck.js   # docs\Copilot-Studio-Agents-SDK-SharePoint.pptx
python tools\build_site_index.py    # docs\index.html (landing page of the GitHub Pages site) + docs\.nojekyll
```

Open the test plan in Excel once after regenerating it so the summary formulas are calculated. Pushing to `main`
republishes the web pages (GitHub Pages, source `main` / `docs`).
