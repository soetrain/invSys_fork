[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SourceRoot,
    [Parameter(Mandatory = $true)][string]$ReleaseRoot,
    [Parameter(Mandatory = $true)][string]$ReleaseId,
    [string]$GitCommit = "",
    [ValidateRange(3, 20)][int]$RetainCount = 3
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "invsys_release_common.ps1")

Assert-InvSysReleaseId -ReleaseId $ReleaseId
$sourcePath = (Resolve-Path -LiteralPath $SourceRoot).Path
if (-not (Test-Path -LiteralPath $ReleaseRoot)) { New-Item -ItemType Directory -Path $ReleaseRoot -Force | Out-Null }
$rootPath = (Resolve-Path -LiteralPath $ReleaseRoot).Path
$releasesPath = Join-Path $rootPath "Releases"
if (-not (Test-Path -LiteralPath $releasesPath)) { New-Item -ItemType Directory -Path $releasesPath -Force | Out-Null }
$finalPath = Join-Path $releasesPath $ReleaseId
if (Test-Path -LiteralPath $finalPath) { throw "Release already exists and is immutable: $finalPath" }

function Publish-StationSetup {
    param(
        [string]$ToolRoot,
        [string]$FeedRoot,
        [string]$SetupReleaseId
    )

    $setupNames = @(
        "install_invsys_station_from_nas.ps1",
        "invsys_release_common.ps1",
        "register_current_addins.ps1",
        "register_invsys_update_task.ps1",
        "update_invsys_station.ps1"
    )
    foreach ($name in $setupNames) {
        if (-not (Test-Path -LiteralPath (Join-Path $toolRoot $name) -PathType Leaf)) {
            throw "StationSetup source was not found: $name"
        }
    }
    $bootstrapSource = Join-Path $toolRoot "start_invsys_nas_station_setup.ps1"
    if (-not (Test-Path -LiteralPath $bootstrapSource -PathType Leaf)) {
        throw "StationSetup bootstrap source was not found: start_invsys_nas_station_setup.ps1"
    }
    $cmdLauncherSource = Join-Path $toolRoot "Install-invSys-Station.cmd"
    if (-not (Test-Path -LiteralPath $cmdLauncherSource -PathType Leaf)) {
        throw "StationSetup command launcher source was not found: Install-invSys-Station.cmd"
    }

    $setupRoot = Join-Path $FeedRoot "StationSetup"
    if (-not (Test-Path -LiteralPath $setupRoot)) { New-Item -ItemType Directory -Path $setupRoot -Force | Out-Null }
    $finalSetupPath = Join-Path $setupRoot $SetupReleaseId
    if (Test-Path -LiteralPath $finalSetupPath) { throw "StationSetup release already exists and is immutable: $finalSetupPath" }
    $stagingSetupPath = Join-Path $setupRoot ("." + $SetupReleaseId + ".staging-" + [guid]::NewGuid().ToString("N"))
    try {
        New-Item -ItemType Directory -Path $stagingSetupPath -Force | Out-Null
        $files = foreach ($name in $setupNames) {
            $source = Join-Path $toolRoot $name
            $destination = Join-Path $stagingSetupPath $name
            Copy-Item -LiteralPath $source -Destination $destination -Force
            [ordered]@{
                name = $name
                sha256 = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant()
            }
        }
        Write-InvSysJsonAtomic -Path (Join-Path $stagingSetupPath "station-setup-manifest.json") -Value ([ordered]@{
            releaseId = $SetupReleaseId
            files = $files
        })
        Move-Item -LiteralPath $stagingSetupPath -Destination $finalSetupPath
    }
    finally {
        if (Test-Path -LiteralPath $stagingSetupPath) { Remove-Item -LiteralPath $stagingSetupPath -Recurse -Force }
    }

    Copy-Item -LiteralPath $bootstrapSource -Destination (Join-Path $FeedRoot "Install-invSys-Station.ps1") -Force
    Copy-Item -LiteralPath $cmdLauncherSource -Destination (Join-Path $FeedRoot "Install-invSys-Station.cmd") -Force
}

$packageNames = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam",
    "invSys.Admin.xlam"
)
foreach ($name in $packageNames) {
    $file = Join-Path $sourcePath $name
    if (-not (Test-Path -LiteralPath $file -PathType Leaf) -or (Get-Item -LiteralPath $file).Length -le 0) {
        throw "Source package is missing or empty: $file"
    }
}

if ([string]::IsNullOrWhiteSpace($GitCommit)) {
    $GitCommit = (& git -C $sourcePath rev-parse HEAD 2>$null)
    if ($LASTEXITCODE -ne 0) { $GitCommit = "unavailable" }
}

$stagingPath = Join-Path $releasesPath ("." + $ReleaseId + ".staging-" + [guid]::NewGuid().ToString("N"))
try {
    New-Item -ItemType Directory -Path $stagingPath -Force | Out-Null
    $packages = foreach ($name in $packageNames) {
        $sourceFile = Join-Path $sourcePath $name
        $destinationFile = Join-Path $stagingPath $name
        Copy-Item -LiteralPath $sourceFile -Destination $destinationFile -Force
        [ordered]@{
            name = $name
            sha256 = (Get-FileHash -LiteralPath $destinationFile -Algorithm SHA256).Hash.ToLowerInvariant()
            bytes = (Get-Item -LiteralPath $destinationFile).Length
        }
    }
    $manifest = [ordered]@{
        releaseId = $ReleaseId
        packageSetVersion = "R1-5"
        gitCommit = ([string]$GitCommit).Trim()
        publishedAtUtc = [DateTime]::UtcNow.ToString("o")
        compatibility = "invSys Release 1 / Architecture v4.11 D16"
        packages = $packages
    }
    Write-InvSysJsonAtomic -Path (Join-Path $stagingPath "release-manifest.json") -Value $manifest
    [void](Assert-InvSysReleaseManifest -ReleaseDirectory $stagingPath -ExpectedReleaseId $ReleaseId)
    Move-Item -LiteralPath $stagingPath -Destination $finalPath
    Publish-StationSetup -ToolRoot $PSScriptRoot -FeedRoot $rootPath -SetupReleaseId $ReleaseId
    $pointer = [ordered]@{
        releaseId = $ReleaseId
        manifest = ("Releases/{0}/release-manifest.json" -f $ReleaseId)
        stationSetup = ("StationSetup/{0}/install_invsys_station_from_nas.ps1" -f $ReleaseId)
        publishedAtUtc = $manifest.publishedAtUtc
    }
    Write-InvSysJsonAtomic -Path (Join-Path $rootPath "current-release.json") -Value $pointer
}
finally {
    if (Test-Path -LiteralPath $stagingPath) { Remove-Item -LiteralPath $stagingPath -Recurse -Force }
}

$keep = @($ReleaseId)
$existing = @(Get-ChildItem -LiteralPath $releasesPath -Directory | Sort-Object LastWriteTimeUtc -Descending)
foreach ($directory in $existing) {
    if ($keep.Count -ge $RetainCount) { break }
    if ($keep -notcontains $directory.Name) { $keep += $directory.Name }
}
foreach ($directory in $existing) {
    if ($keep -notcontains $directory.Name) { Remove-Item -LiteralPath $directory.FullName -Recurse -Force }
}

Write-Host "INVSYS_RELEASE_PUBLISHED"
Write-Host "ReleaseId=$ReleaseId"
Write-Host "ReleaseRoot=$rootPath"
