<#
.SYNOPSIS
  Creates the licensed test personas (User A-D) used by the Agents SDK x SharePoint test matrix and writes their
  credentials to <SecretsDir>\test-users.json (keep that folder outside source control).

.DESCRIPTION
  Requires Azure CLI signed in to the tenant with User Administrator (or higher) rights.
  Existing users are left untouched unless -ResetPassword is specified.

.EXAMPLE
  .\New-AgentsSdkTestUsers.ps1 -TenantDomain contoso.onmicrosoft.com -SkuPartNumber 'Microsoft_365_E5_(no_Teams)'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $TenantDomain,
    [string] $SkuPartNumber = 'Microsoft_365_E5_(no_Teams)',
    [string] $UsageLocation = 'US',
    [string] $SecretsDir = (Join-Path $HOME '.agents-sdk-test'),
    [switch] $ResetPassword
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $SecretsDir | Out-Null
$secretsFile = Join-Path $SecretsDir 'test-users.json'
$existing = @{}
if (Test-Path $secretsFile) { (Get-Content $secretsFile -Raw | ConvertFrom-Json) | ForEach-Object { $existing[$_.upn.ToLower()] = $_ } }

$skuId = az rest --method GET --uri 'https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuPartNumber,skuId' `
    --query "value[?skuPartNumber=='$SkuPartNumber'].skuId | [0]" -o tsv
if (-not $skuId) { throw "SKU $SkuPartNumber not found in the tenant." }

$personas = @(
    @{ l = 'a'; access = 'Agent A + Agent B' }, @{ l = 'b'; access = 'Agent A only' },
    @{ l = 'c'; access = 'Agent B only' }, @{ l = 'd'; access = 'No agent access' }
)
$out = @()
foreach ($p in $personas) {
    $upn = "agentsdk.user$($p.l)@$TenantDomain"
    $id = az ad user show --id $upn --query id -o tsv 2>$null
    $password = $null
    if (-not $id -or $ResetPassword) {
        $password = (-join ((48..57) + (65..90) + (97..122) | Get-Random -Count 14 | ForEach-Object { [char]$_ })) + '!a9'
    }
    if (-not $id) {
        $id = az ad user create --display-name "Agents SDK User $($p.l.ToUpper())" --user-principal-name $upn --password $password `
            --force-change-password-next-sign-in false --mail-nickname "agentsdk.user$($p.l)" --query id -o tsv --only-show-errors
    } elseif ($ResetPassword) {
        az ad user update --id $upn --password $password --force-change-password-next-sign-in false --only-show-errors
    } else {
        $password = $existing[$upn.ToLower()].password
    }
    $patch = New-TemporaryFile
    @{ usageLocation = $UsageLocation; jobTitle = "Test persona - $($p.access)" } | ConvertTo-Json | Set-Content $patch
    az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/users/$id" --headers 'Content-Type=application/json' --body "@$patch" --only-show-errors | Out-Null
    $lic = New-TemporaryFile
    @{ addLicenses = @(@{ skuId = $skuId }); removeLicenses = @() } | ConvertTo-Json -Depth 4 | Set-Content $lic
    az rest --method POST --uri "https://graph.microsoft.com/v1.0/users/$id/assignLicense" --headers 'Content-Type=application/json' --body "@$lic" --only-show-errors | Out-Null
    Remove-Item $patch, $lic
    $out += [ordered]@{ persona = "User $($p.l.ToUpper())"; upn = $upn; id = $id; password = $password; agentAccess = $p.access }
    Write-Host "User $($p.l.ToUpper()): $upn" -ForegroundColor Cyan
}
$out | ConvertTo-Json | Set-Content $secretsFile -Encoding utf8
Write-Host "Credentials written to $secretsFile (do not commit)." -ForegroundColor Yellow
