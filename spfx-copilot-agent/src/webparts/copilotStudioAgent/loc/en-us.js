define([], function() {
  return {
    "PropertyPaneDescription": "Chat with a Copilot Studio agent through the Microsoft 365 Agents SDK. Users are signed in silently with Microsoft Entra ID (SSO).",
    "AgentGroupName": "Agent",
    "AuthGroupName": "Authentication (Microsoft Entra ID)",
    "AppearanceGroupName": "Appearance",
    "TitleFieldLabel": "Title",
    "ModeFieldLabel": "Integration mode",
    "ModeSdk": "Microsoft 365 Agents SDK (direct, Entra ID SSO)",
    "ModeEmbed": "Copilot Studio hosted web chat (embed code)",
    "EmbedUrlFieldLabel": "Embed code or URL",
    "EmbedUrlFieldDescription": "Copilot Studio > your agent > Channels > App > Embed code. Use this for agents the Agents SDK cannot reach (for example GitHub Copilot harness agents).",
    "EmbedUrlInvalid": "Expected a Copilot Studio web chat URL (https://copilotstudio.microsoft.com/... or https://copilotstudio.preview.microsoft.com/...).",
    "ConnectionStringFieldLabel": "Agents SDK connection string",
    "ConnectionStringFieldDescription": "Copilot Studio > your agent > Channels > App > Microsoft 365 Agents SDK > Connection string.",
    "ConnectionStringInvalid": "Expected https://<env>.environment.api.powerplatform.com/copilotstudio/... (copy it from the App channel).",
    "OrLabel": "…or use the environment ID and the agent schema name:",
    "EnvironmentIdFieldLabel": "Environment ID",
    "AgentIdentifierFieldLabel": "Agent schema name",
    "AppClientIdFieldLabel": "Entra app (client) ID",
    "AppClientIdFieldDescription": "SPA app registration with Power Platform API / CopilotStudio.Copilots.Invoke. Leave empty to use the tenant default.",
    "TenantIdFieldLabel": "Tenant ID",
    "TenantIdFieldDescription": "Leave empty to use the tenant of this SharePoint site.",
    "GuidInvalid": "Enter a GUID.",
    "ShowHeaderFieldLabel": "Show header",
    "FluentThemeFieldLabel": "Modern (Fluent) chat theme",
    "HeightFieldLabel": "Height (px)"
  }
});
