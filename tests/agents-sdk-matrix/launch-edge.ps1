<#
.SYNOPSIS
  Starts one Microsoft Edge window per test persona, each with its own profile and remote-debugging port
  (ports must match PERSONAS in settings.py). Sign-in state is kept in the profiles between runs.
#>
param(
    [string] $ProfileRoot = (Join-Path $env:LOCALAPPDATA 'AgentsSdkTestEdge'),
    [hashtable] $Ports = @{ mario = 9333; a = 9341; b = 9342; c = 9343; d = 9344 }
)
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found.' }
foreach ($persona in $Ports.Keys) {
    $port = $Ports[$persona]
    try { Invoke-RestMethod "http://127.0.0.1:$port/json/version" -TimeoutSec 2 | Out-Null; Write-Host "$persona already running on $port"; continue } catch { }
    $profileDir = Join-Path $ProfileRoot $persona
    New-Item -ItemType Directory -Force $profileDir | Out-Null
    Start-Process -FilePath $edge -ArgumentList @("--remote-debugging-port=$port", "--user-data-dir=$profileDir", '--no-first-run', '--no-default-browser-check', 'about:blank')
    Write-Host "$persona -> Edge on port $port (profile $profileDir)"
}
