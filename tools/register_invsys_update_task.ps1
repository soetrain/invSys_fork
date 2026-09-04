[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ReleaseRoot,
    [string]$CacheRoot = (Join-Path $env:LOCALAPPDATA "invSys\\Addins"),
    [string]$TaskName = "invSys.StationUpdate",
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$updater = Join-Path $scriptRoot "update_invsys_station.ps1"
if (-not (Test-Path -LiteralPath $updater -PathType Leaf)) { throw "Station updater was not found: $updater" }
$agentSources = @(
    "invsys_release_common.ps1",
    "register_current_addins.ps1",
    "update_invsys_station.ps1",
    "rollback_invsys_station_release.ps1"
)
foreach ($name in $agentSources) {
    if (-not (Test-Path -LiteralPath (Join-Path $scriptRoot $name) -PathType Leaf)) { throw "Station agent source was not found: $name" }
}
$agentHashMaterial = ($agentSources | ForEach-Object { (Get-FileHash -LiteralPath (Join-Path $scriptRoot $_) -Algorithm SHA256).Hash.ToLowerInvariant() }) -join ""
$sha256 = [Security.Cryptography.SHA256]::Create()
try { $agentId = "D16-" + (($sha256.ComputeHash([Text.Encoding]::UTF8.GetBytes($agentHashMaterial)) | ForEach-Object { $_.ToString("x2") }) -join "").Substring(0, 12) }
finally { $sha256.Dispose() }
$stationRoot = Split-Path -Parent $CacheRoot
$agentRoot = Join-Path $stationRoot "Deployment\\Agents"
$agentPath = Join-Path $agentRoot $agentId
$agentUpdater = Join-Path $agentPath "update_invsys_station.ps1"
$command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -ReleaseRoot "{1}" -CacheRoot "{2}"' -f $agentUpdater, $ReleaseRoot, $CacheRoot
if (-not $Apply) {
    Write-Output "TaskName=$TaskName"
    Write-Output "Trigger=AtLogOn"
    Write-Output "Trigger=Every 15 minutes"
    Write-Output "AgentId=$agentId"
    Write-Output "AgentPath=$agentPath"
    Write-Output "Command=$command"
    exit 0
}

if (-not (Test-Path -LiteralPath $agentPath -PathType Container)) {
    if (-not (Test-Path -LiteralPath $agentRoot -PathType Container)) { New-Item -ItemType Directory -Path $agentRoot -Force | Out-Null }
    $staging = Join-Path $agentRoot ("." + $agentId + ".staging-" + [guid]::NewGuid().ToString("N"))
    try {
        New-Item -ItemType Directory -Path $staging -Force | Out-Null
        foreach ($name in $agentSources) {
            $source = Join-Path $scriptRoot $name
            $target = Join-Path $staging $name
            Copy-Item -LiteralPath $source -Destination $target -Force
            if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash) { throw "Station agent hash verification failed for $name" }
        }
        Move-Item -LiteralPath $staging -Destination $agentPath
    }
    finally { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force } }
}
foreach ($name in $agentSources) {
    $sourceHash = (Get-FileHash -LiteralPath (Join-Path $scriptRoot $name) -Algorithm SHA256).Hash
    $agentFile = Join-Path $agentPath $name
    if (-not (Test-Path -LiteralPath $agentFile -PathType Leaf) -or $sourceHash -ne (Get-FileHash -LiteralPath $agentFile -Algorithm SHA256).Hash) { throw "Existing station agent does not match the verified source: $name" }
}

$logon = New-ScheduledTaskTrigger -AtLogOn
$periodic = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(1)) -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 365)
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`" -ReleaseRoot `"{1}`" -CacheRoot `"{2}`"" -f $agentUpdater, $ReleaseRoot, $CacheRoot)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger @($logon, $periodic) -Settings $settings -Description "invSys D16 five-package station update; defers while Excel is open." -Force | Out-Null
Write-Host "INVSYS_UPDATE_TASK_REGISTERED"
