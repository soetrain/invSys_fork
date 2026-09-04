[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ReleaseRoot,
    [string]$CacheRoot = (Join-Path $env:LOCALAPPDATA "invSys\Addins"),
    [string]$ExcelOptionsKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options",
    [string]$AddinManagerKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager",
    [string]$ExcelProcessName = "EXCEL",
    [switch]$RegisterPeriodicUpdater
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-StationSetupFiles {
    param([string]$SetupRoot)

    $manifestPath = Join-Path $SetupRoot "station-setup-manifest.json"
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "StationSetup manifest was not found: $manifestPath"
    }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    if (@($manifest.files).Count -eq 0) { throw "StationSetup manifest contains no files." }

    foreach ($file in @($manifest.files)) {
        $name = [string]$file.name
        if ([string]::IsNullOrWhiteSpace($name) -or $name.Contains("..") -or $name.Contains("\") -or $name.Contains("/")) {
            throw "StationSetup manifest contains an invalid file name."
        }
        $path = Join-Path $SetupRoot $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "StationSetup file is missing: $name" }
        $actualHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actualHash -ne ([string]$file.sha256).ToLowerInvariant()) {
            throw "StationSetup file hash mismatch: $name"
        }
    }
}

$rootPath = (Resolve-Path -LiteralPath $ReleaseRoot).Path
$setupRoot = $PSScriptRoot
Assert-StationSetupFiles -SetupRoot $setupRoot

if (Get-Process -Name $ExcelProcessName -ErrorAction SilentlyContinue) {
    Write-Output "INVSYS_STATION_SETUP_DEFERRED_EXCEL_OPEN"
    Write-Output "Close all Excel windows, then run the same NAS StationSetup command again."
    exit 0
}

$updater = Join-Path $setupRoot "update_invsys_station.ps1"
if (-not (Test-Path -LiteralPath $updater -PathType Leaf)) { throw "Station updater was not found in StationSetup." }
& $updater -ReleaseRoot $ReleaseRoot -CacheRoot $CacheRoot `
    -ExcelOptionsKey $ExcelOptionsKey -AddinManagerKey $AddinManagerKey `
    -ExcelProcessName $ExcelProcessName

if ($RegisterPeriodicUpdater) {
    try {
        $taskInstaller = Join-Path $setupRoot "register_invsys_update_task.ps1"
        if (-not (Test-Path -LiteralPath $taskInstaller -PathType Leaf)) { throw "Task installer was not found in StationSetup." }
        & $taskInstaller -ReleaseRoot $ReleaseRoot -CacheRoot $CacheRoot -Apply
    }
    catch {
        Write-Warning "Warning: periodic updater was not registered: $($_.Exception.Message)"
    }
}

$statusPath = Join-Path $CacheRoot "update-status.json"
if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { throw "Station updater did not produce a local update status." }
$status = Get-Content -LiteralPath $statusPath -Raw | ConvertFrom-Json
if ([string]$status.status -ne "APPLIED") { throw "Station update did not apply: $([string]$status.status)" }

Write-Output "INVSYS_STATION_SETUP_APPLIED"
Write-Output ("ReleaseId=" + [string]$status.releaseId)
Write-Output "Next step: open Excel, use Server Sign In, select the warehouse, then use invSys Sign In."
