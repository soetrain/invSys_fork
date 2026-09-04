[CmdletBinding()]
param(
    [string]$ReleaseRoot = "",
    [string]$CacheRoot = (Join-Path $env:LOCALAPPDATA "invSys\Addins"),
    [string]$ExcelOptionsKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options",
    [string]$AddinManagerKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager",
    [string]$ExcelProcessName = "EXCEL",
    [switch]$RegisterPeriodicUpdater
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ReleaseRoot)) {
    $ReleaseRoot = $PSScriptRoot
}
$rootPath = (Resolve-Path -LiteralPath $ReleaseRoot).Path
$pointerPath = Join-Path $rootPath "current-release.json"
if (-not (Test-Path -LiteralPath $pointerPath -PathType Leaf)) {
    throw "invSys NAS release pointer was not found: $pointerPath"
}

$pointer = Get-Content -LiteralPath $pointerPath -Raw | ConvertFrom-Json
$setupRelative = [string]$pointer.stationSetup
if ([string]::IsNullOrWhiteSpace($setupRelative) -or
    -not $setupRelative.StartsWith("StationSetup/", [StringComparison]::OrdinalIgnoreCase) -or
    -not $setupRelative.EndsWith("/install_invsys_station_from_nas.ps1", [StringComparison]::OrdinalIgnoreCase) -or
    $setupRelative.Contains("..")) {
    throw "The current invSys NAS release does not declare a valid StationSetup entry point."
}

$setupPath = Join-Path $rootPath ($setupRelative.Replace("/", "\"))
if (-not (Test-Path -LiteralPath $setupPath -PathType Leaf)) {
    throw "The current invSys StationSetup installer was not found: $setupPath"
}

& $setupPath -ReleaseRoot $rootPath -CacheRoot $CacheRoot `
    -ExcelOptionsKey $ExcelOptionsKey -AddinManagerKey $AddinManagerKey `
    -ExcelProcessName $ExcelProcessName -RegisterPeriodicUpdater:$RegisterPeriodicUpdater
