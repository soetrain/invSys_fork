[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$docs = (Resolve-Path (Join-Path $repo "..\invSys_docs")).Path

function Read-Source([string]$path) {
    Get-Content -Raw -LiteralPath $path
}

$publisher = Read-Source (Join-Path $repo "tools\publish_invsys_release.ps1")
$installerPath = Join-Path $repo "tools\install_invsys_station_from_nas.ps1"
$bootstrapPath = Join-Path $repo "tools\start_invsys_nas_station_setup.ps1"
$register = Read-Source (Join-Path $repo "tools\register_invsys_update_task.ps1")
$spec = Read-Source (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$failures = [System.Collections.Generic.List[string]]::new()
$passes = [System.Collections.Generic.List[string]]::new()

function Check([string]$name, [bool]$passed, [string]$contract) {
    if ($passed) {
        $passes.Add($name)
        Write-Host "PASS $name"
    } else {
        $failures.Add("${name}: ${contract}")
        Write-Host "FAIL $name - $contract"
    }
}

$installer = if (Test-Path -LiteralPath $installerPath) { Read-Source $installerPath } else { "" }
$bootstrap = if (Test-Path -LiteralPath $bootstrapPath) { Read-Source $bootstrapPath } else { "" }

Check "StationSetup.PublisherStagesVersionedEntryPoint" (
    $publisher.Contains('StationSetup') -and
    $publisher.Contains('station-setup-manifest.json') -and
    $publisher.Contains('start_invsys_nas_station_setup.ps1')
) "The NAS publisher must expose a versioned station-setup entry point beside the immutable release feed."

Check "StationSetup.BootstrapResolvesCurrentRelease" (
    $bootstrap.Contains('current-release.json') -and
    $bootstrap.Contains('stationSetup') -and
    $bootstrap.Contains('install_invsys_station_from_nas.ps1')
) "The user-runnable NAS bootstrap must resolve only the current feed-declared setup script."

Check "StationSetup.InstallsFivePackageRelease" (
    $installer.Contains('update_invsys_station.ps1') -and
    $installer.Contains('-ReleaseRoot $ReleaseRoot') -and
    $installer.Contains('INVSYS_STATION_SETUP_APPLIED')
) "The station setup must reuse the hash-validating updater and report a successful local registration handoff."

Check "StationSetup.TaskIsOptional" (
    $installer.Contains('[switch]$RegisterPeriodicUpdater') -and
    $installer.Contains('register_invsys_update_task.ps1') -and
    $installer.Contains('Warning: periodic updater was not registered')
) "Manual NAS setup must succeed even when optional periodic task registration is unavailable."

Check "StationSetup.NoGitOrWarehouseAuthority" (
    -not $installer.Contains('git clone') -and
    -not $installer.Contains('BootstrapWarehouse') -and
    -not $installer.Contains('CreateWarehouse') -and
    -not $installer.Contains('invSys.Data.Inventory')
) "Station setup must not acquire from Git or write warehouse/auth/inventory authority."

Check "StationSetup.CreateWarehouseRemainsCapabilityGated" (
    $spec.Contains('**Create New Warehouse** remains an `ADMIN_MAINT` action') -and
    $register.Contains('Register-ScheduledTask')
) "NAS deployment access is separate from the existing invSys-admin and optional local task boundaries."

if ($failures.Count -eq 0) {
    $scratch = Join-Path ([IO.Path]::GetTempPath()) ("invSys-StationSetup-" + [guid]::NewGuid().ToString("N"))
    $registryPath = "HKCU:\Software\invSysTest\Slice4bj\" + [guid]::NewGuid().ToString("N")
    try {
        $feed = Join-Path $scratch "feed"
        $cache = Join-Path $scratch "cache"
        $source = Join-Path $repo "deploy\current"
        New-Item -ItemType Directory -Path $scratch -Force | Out-Null

        & (Join-Path $repo "tools\publish_invsys_release.ps1") `
            -SourceRoot $source -ReleaseRoot $feed -ReleaseId "test-station-r1" -GitCommit "test-station" | Out-Null
        New-Item -Path $registryPath -Force | Out-Null
        Set-ItemProperty -Path $registryPath -Name "OPEN" -Value '"C:\ThirdParty.xlam"' -Type String

        & (Join-Path $feed "Install-invSys-Station.ps1") `
            -ReleaseRoot $feed -CacheRoot $cache -ExcelOptionsKey $registryPath `
            -AddinManagerKey ($registryPath + "\Manager") -ExcelProcessName "invSysNoExcelProcess"

        $pointer = Get-Content -LiteralPath (Join-Path $feed "current-release.json") -Raw | ConvertFrom-Json
        $local = Get-Content -LiteralPath (Join-Path $cache "current-release.json") -Raw | ConvertFrom-Json
        $setupManifest = Get-Content -LiteralPath (Join-Path $feed "StationSetup/test-station-r1/station-setup-manifest.json") -Raw | ConvertFrom-Json
        $registered = Get-ItemProperty -Path $registryPath
        $openValues = @($registered.PSObject.Properties |
            Where-Object { $_.Name -match '^OPEN\d*$' } |
            ForEach-Object { [string]$_.Value })
        $cachedPackages = @("invSys.Core.xlam", "invSys.Inventory.Domain.xlam", "invSys.Designs.Domain.xlam", "invSys.Operations.xlam", "invSys.Admin.xlam") |
            ForEach-Object { Test-Path -LiteralPath (Join-Path $cache ("Releases/test-station-r1/" + $_)) }

        Check "StationSetup.Integration.NasToCacheAndLeaves" (
            $pointer.stationSetup -eq "StationSetup/test-station-r1/install_invsys_station_from_nas.ps1" -and
            $local.releaseId -eq "test-station-r1" -and
            (@($cachedPackages | Where-Object { -not $_ }).Count -eq 0) -and
            (@($openValues | Where-Object { $_ -match 'invSys\.Operations\.xlam' }).Count -eq 1) -and
            (@($openValues | Where-Object { $_ -match 'invSys\.Admin\.xlam' }).Count -eq 1) -and
            (@($openValues | Where-Object { $_ -match 'ThirdParty\.xlam' }).Count -eq 1)
        ) "The NAS launcher must cache a complete release and register only Operations/Admin without disturbing other add-ins."

        Check "StationSetup.Integration.VersionedHashManifest" (
            $setupManifest.releaseId -eq "test-station-r1" -and
            @($setupManifest.files).Count -ge 5 -and
            @($setupManifest.files | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.sha256) }).Count -eq 0
        ) "The NAS StationSetup sidecar is versioned and hash-described beside the immutable release."
    }
    finally {
        if (Test-Path -LiteralPath $registryPath) { Remove-Item -LiteralPath $registryPath -Recurse -Force }
        if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Recurse -Force }
    }
}

Write-Host "RESULT passed=$($passes.Count) failed=$($failures.Count)"
if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "  $_" }
    exit 1
}
