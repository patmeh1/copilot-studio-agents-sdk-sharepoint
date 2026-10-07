<#
.SYNOPSIS
  Creates (or updates) the Microsoft Entra ID single-page-application (SPA) registration used by the
  "Copilot Studio Agent (Agents SDK)" SharePoint web part to call Copilot Studio agents through the
  Microsoft 365 Agents SDK (Power Platform API, delegated scope CopilotStudio.Copilots.Invoke).

.DESCRIPTION
  Idempotent. Uses the Azure CLI (az) so it runs the same way on any workstation.
    1. Verifies az is signed in to the expected tenant.
    2. Makes sure the "Power Platform API" service principal (appId 8578e004-a5c6-46e7-913e-12f58912df43) exists.
    3. Creates the app registration (or re-uses one with the same display name).
    4. Sets the SPA redirect URI(s) (the SharePoint tenant origin works for every site in the tenant).
    5. Adds delegated permissions: Power Platform API / CopilotStudio.Copilots.Invoke and
       Microsoft Graph / openid, profile, offline_access, User.Read.
    6. Creates the enterprise app (service principal) and tries to grant tenant-wide admin consent.

  Sign in first (use an isolated config dir if you work with several tenants):
    $env:AZURE_CONFIG_DIR = "$HOME\.azure-<tenant>"      # optional
    az login --tenant <tenant>.onmicrosoft.com --allow-no-subscriptions

.EXAMPLE
  .\New-AgentsSdkSpaApp.ps1 -TenantId contoso.onmicrosoft.com -SharePointRootUrl https://contoso.sharepoint.com
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $TenantId,
    [Parameter(Mandatory)] [string] $SharePointRootUrl,
    [string] $AppDisplayName = 'Copilot Studio Agents SDK - SharePoint',
    [string[]] $AdditionalRedirectUris = @(),
    [switch] $SkipAdminConsent
)

$ErrorActionPreference = 'Stop'
$PowerPlatformApiAppId = '8578e004-a5c6-46e7-913e-12f58912df43'
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$GraphScopes = @{
    'openid'         = '37f7f235-527c-4136-accd-4a02d197296e'
    'profile'        = '14dad69e-099b-42c9-810b-d002981feec1'
    'offline_access' = '7427e0e9-2fba-42fe-b0c0-848c9e6a8182'
    'User.Read'      = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'
}

function Invoke-Az {
    # Simple function (no param block) so az switches such as -o are passed through untouched.
    $out = & az @args --only-show-errors 2>&1
    if ($LASTEXITCODE -ne 0) { throw "az $($args -join ' ') failed: $out" }
    return $out
}

# 1. Signed-in tenant check
$account = Invoke-Az account show -o json | ConvertFrom-Json
$tenantDetails = Invoke-Az rest --method GET --uri 'https://graph.microsoft.com/v1.0/organization?$select=id,verifiedDomains' -o json | ConvertFrom-Json
$org = $tenantDetails.value[0]
if ($TenantId -ne $org.id -and -not ($org.verifiedDomains.name -contains $TenantId)) {
    throw "az is signed in to tenant $($account.tenantId), not '$TenantId'. Run: az login --tenant $TenantId --allow-no-subscriptions"
}
$tenantGuid = $org.id
Write-Host "Tenant: $tenantGuid  (signed in as $($account.user.name))" -ForegroundColor Cyan

# 2. Power Platform API service principal
$ppSp = & az ad sp show --id $PowerPlatformApiAppId -o json --only-show-errors 2>$null | ConvertFrom-Json
if (-not $ppSp) {
    Write-Host 'Power Platform API service principal not found - creating it...' -ForegroundColor Yellow
    $ppSp = Invoke-Az ad sp create --id $PowerPlatformApiAppId -o json | ConvertFrom-Json
}
$invokeScope = $ppSp.oauth2PermissionScopes | Where-Object { $_.value -eq 'CopilotStudio.Copilots.Invoke' }
if (-not $invokeScope) { throw 'Scope CopilotStudio.Copilots.Invoke not found on the Power Platform API service principal.' }

# 3. App registration
$app = (Invoke-Az ad app list --display-name $AppDisplayName -o json | ConvertFrom-Json) | Select-Object -First 1
if (-not $app) {
    Write-Host "Creating app registration '$AppDisplayName'..." -ForegroundColor Cyan
    $app = Invoke-Az ad app create --display-name $AppDisplayName --sign-in-audience AzureADMyOrg -o json | ConvertFrom-Json
} else {
    Write-Host "Re-using existing app registration '$AppDisplayName' ($($app.appId))" -ForegroundColor Cyan
}

# 4 + 5. SPA redirect URIs and delegated permissions
$redirects = @($SharePointRootUrl.TrimEnd('/')) + $AdditionalRedirectUris | Where-Object { $_ } | Select-Object -Unique
$body = @{
    spa                    = @{ redirectUris = @($redirects) }
    requiredResourceAccess = @(
        @{ resourceAppId = $PowerPlatformApiAppId; resourceAccess = @(@{ id = $invokeScope.id; type = 'Scope' }) },
        @{ resourceAppId = $GraphAppId; resourceAccess = @($GraphScopes.Values | ForEach-Object { @{ id = $_; type = 'Scope' } }) }
    )
} | ConvertTo-Json -Depth 6
$bodyFile = New-TemporaryFile
Set-Content -Path $bodyFile -Value $body -Encoding utf8
Invoke-Az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" --headers 'Content-Type=application/json' --body "@$bodyFile" | Out-Null
Remove-Item $bodyFile

# 6. Enterprise app + admin consent
$sp = & az ad sp show --id $app.appId -o json --only-show-errors 2>$null | ConvertFrom-Json
if (-not $sp) { $sp = Invoke-Az ad sp create --id $app.appId -o json | ConvertFrom-Json }

$consent = 'skipped'
if (-not $SkipAdminConsent) {
    $consent = 'failed'
    for ($i = 1; $i -le 6; $i++) {
        & az ad app permission admin-consent --id $app.appId --only-show-errors 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { $consent = 'granted'; break }
        Start-Sleep -Seconds 10
    }
    if ($consent -ne 'granted') {
        Write-Warning ("Tenant-wide admin consent could not be granted with the signed-in account (needs Global Admin, " +
            "Privileged Role Admin, Cloud Application Admin or Application Admin). Ask an admin to open " +
            "https://login.microsoftonline.com/$tenantGuid/adminconsent?client_id=$($app.appId) " +
            "or Entra admin center > App registrations > $AppDisplayName > API permissions > Grant admin consent. " +
            "Until then each user is asked to consent once (if your user-consent policy allows it).")
    }
}

$result = [ordered]@{
    tenantId     = $tenantGuid
    appClientId  = $app.appId
    appObjectId  = $app.id
    displayName  = $AppDisplayName
    redirectUris = $redirects
    scope        = 'https://api.powerplatform.com/CopilotStudio.Copilots.Invoke'
    adminConsent = $consent
}
$result | ConvertTo-Json -Depth 4
