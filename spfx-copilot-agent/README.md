# Copilot Studio Agent (Agents SDK) - SPFx web part

SharePoint Framework 1.21.1 web part that embeds a Copilot Studio agent through the **Microsoft 365 Agents SDK**
(`@microsoft/agents-copilotstudio-client`) and renders it with **Bot Framework Web Chat** (Fluent theme).

* **Authentication**: MSAL.js (`@azure/msal-browser` 4.30, LTS) with a dedicated Entra **SPA** app registration
  (redirect URI = SharePoint origin). Silent SSO via `acquireTokenSilent` -> `ssoSilent(login_hint = SharePoint user)`,
  popup fallback (*Sign in with Microsoft*) when silent SSO is blocked - one interactive sign-in serves every agent web part
  on the page. SPFx `AadTokenProvider` is **not** used because it
  cannot obtain `CopilotStudio.Copilots.Invoke` ([sp-dev-docs#10775](https://github.com/SharePoint/sp-dev-docs/issues/10775)).
* **Token audience** is derived from the connection string (`CopilotStudioClient.scopeFromSettings`, e.g.
  `https://api.powerplatform.com/.default`).
* **Modes**: `sdk` (default) - Agents SDK; `embed` - Copilot Studio hosted web chat iframe (comparison/fallback only, it does
  not enforce authentication for GitHub Copilot harness agents - see the guide).
* Everything is bundled into the package (no CDN scripts, no SharePoint *API access* approvals).

## Properties

| Property | Description |
|---|---|
| `title` | Header text |
| `mode` | `sdk` or `embed` |
| `connectionString` | Agents SDK connection string (Copilot Studio > Channels > Web app / App) |
| `environmentId` + `agentIdentifier` | Alternative to the connection string (schema name) |
| `embedUrl` | Embed code or iframe URL (embed mode) |
| `appClientId` / `tenantId` | Override the tenant defaults from `src/config/defaults.ts` / the current site's tenant |
| `height`, `showHeader`, `useFluentTheme` | Appearance |

## Code map

| Path | |
|---|---|
| `src/services/AgentAuthService.ts` | MSAL instance per (client, tenant), serialized silent token requests, popup flow and page-wide signed-in event |
| `src/services/AgentConnection.ts` | `ConnectionSettings`, `CopilotStudioClient`, Web Chat bridge with error capture, friendly error mapping |
| `src/webparts/copilotStudioAgent/components/CopilotStudioAgent.tsx` | UI states (signing in, sign-in needed, error, chat) and embed mode |
| `src/config/defaults.ts` | Tenant defaults written by `scripts/Build-SpfxPackage.ps1` |

## Build

Requires Node **22** (`>=22.14 <23`).

```powershell
npm install
npx gulp bundle --ship
npx gulp package-solution --ship   # -> sharepoint/solution/copilot-studio-agent-webpart.sppkg
```

or use `..\scripts\Build-SpfxPackage.ps1`, which also stamps the tenant defaults and version.

## Version history

| Version | Change |
|---|---|
| 1.0.3.0 | A web part still trying silent sign-on picks up an account that another web part on the page signed in interactively. |
| 1.0.2.0 | *Sign in with Microsoft* opens the popup directly from the click instead of queueing behind other web parts' silent attempts (the popup blocker rejected it on pages with several agents - test TC-03). One interactive sign-in now reconnects every agent web part on the page. |
| 1.0.1.0 | Embed mode (hosted web chat) for comparison. |
| 1.0.0.0 | First release: Agents SDK mode with MSAL single sign-on. |
