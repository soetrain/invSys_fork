[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ReleaseRoot,
    [string]$CacheRoot = (Join-Path $env:LOCALAPPDATA "invSys\\Addins"),
    [string]$TaskName = "invSys.StationUpdate",
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$updater = Join-Path $PSScriptRoot "update_invsys_station.ps1"
if (-not (Test-Path -LiteralPath $updater -PathType Leaf)) { throw "Station updater was not found: $updater" }
$command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -ReleaseRoot "{1}" -CacheRoot "{2}"' -f $updater, $ReleaseRoot, $CacheRoot
if (-not $Apply) {
    Write-Output "TaskName=$TaskName"
    Write-Output "Trigger=AtLogOn"
    Write-Output "Trigger=Every 15 minutes"
    Write-Output "Command=$command"
    exit 0
}

$logon = New-ScheduledTaskTrigger -AtLogOn
$periodic = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(1)) -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 365)
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`" -ReleaseRoot `"{1}`" -CacheRoot `"{2}`"" -f $updater, $ReleaseRoot, $CacheRoot)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger @($logon, $periodic) -Settings $settings -Description "invSys D16 five-package station update; defers while Excel is open." -Force | Out-Null
Write-Host "INVSYS_UPDATE_TASK_REGISTERED"
