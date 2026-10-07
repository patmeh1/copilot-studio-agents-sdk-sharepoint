# -----------------------------------------------------------------------------------------------
# Variables for recreating "authenticated Copilot Studio agents on many SharePoint sites via the
# Microsoft 365 Agents SDK". Copy to variables.ps1, fill in, then dot-source:  . .\variables.ps1
# Values shown in comments are the ones used in the validated test tenant.
# -----------------------------------------------------------------------------------------------

# Tenant
$TenantDomain      = '<tenant>.onmicrosoft.com'                 # M365CPI03719864.onmicrosoft.com
$TenantId          = '<tenant-guid>'                            # c13b5dfa-60e0-4b6d-bfad-5a8f9b8e5c0d
$SharePointRootUrl = 'https://<tenant>.sharepoint.com'          # https://m365cpi03719864.sharepoint.com
$AppCatalogUrl     = "$SharePointRootUrl/sites/appcatalog"      # from /_api/SP_TenantSettings_Current

# Copilot Studio (Settings > Advanced > Metadata, or the App/Web app channel)
$EnvironmentId     = '<environment-guid>'                       # afdeda35-52fa-e145-9bc0-7b43e756300a
$AgentSchemaName   = '<publisherprefix_AgentSchemaName>'        # new_AgentAStandardharness
# Channels > Web app (standard harness) or Channels > App (new experience) > Microsoft 365 Agents SDK.
# Always prefer copying it from Copilot Studio; for the public cloud it is derived like this:
$envHost = ($EnvironmentId -replace '-', '')
$envHost = $envHost.Substring(0, $envHost.Length - 2) + '.' + $envHost.Substring($envHost.Length - 2)
$AgentConnectionString = "https://$envHost.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/$AgentSchemaName/conversations?api-version=2022-03-01-preview"

# Entra ID SPA app registration used by the web part (MSAL)
$AppDisplayName    = 'Copilot Studio Agents SDK - SharePoint'
$AppClientId       = '<output of New-AgentsSdkSpaApp.ps1>'      # 1f06ba87-8706-4216-bb08-f8d10c3118cf

# SPFx build
$Node22Path        = 'C:\tools\node-v22.x-win-x64'              # SPFx 1.21.1 needs Node >= 22.14 < 23
$PackageVersion    = '1.0.3.0'                                  # increase for every re-deployment (1.0.3.0 is deployed in the demo tenant)

# Target sites (any number - the same package and app registration serve them all)
$SiteUrls = @(
    "$SharePointRootUrl/sites/AgentsSDK-SiteA",
    "$SharePointRootUrl/sites/AgentsSDK-SiteB",
    "$SharePointRootUrl/sites/AgentsSDK-SiteC"
)

# Who may chat with the agent (Copilot Studio > Share). Prefer Entra security groups in production.
$AgentUsers = @('usera@<tenant>.onmicrosoft.com', 'userb@<tenant>.onmicrosoft.com')
