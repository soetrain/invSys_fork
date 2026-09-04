[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ReleaseRoot,
    [string]$CacheRoot = (Join-Path $env:LOCALAPPDATA "invSys\\Addins"),
    [string]$RegisterScriptPath = (Join-Path $PSScriptRoot "register_current_addins.ps1"),
    [string]$ExcelProcessName = "EXCEL",
    [string]$ExcelOptionsKey = "HKCU:\\Software\\Microsoft\\Office\\16.0\\Excel\\Options",
    [string]$AddinManagerKey = "HKCU:\\Software\\Microsoft\\Office\\16.0\\Excel\\Add-in Manager"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "invsys_release_common.ps1")

$rootPath = (Resolve-Path -LiteralPath $ReleaseRoot).Path
if (-not (Test-Path -LiteralPath $RegisterScriptPath -PathType Leaf)) { throw "Registration script was not found: $RegisterScriptPath" }
$pointer = Get-InvSysReleasePointer -Root $rootPath
$releaseId = [string]$pointer.releaseId
if (Get-Process -Name $ExcelProcessName -ErrorAction SilentlyContinue) {
    Write-InvSysUpdateStatus -CacheRoot $CacheRoot -Status "DEFERRED_EXCEL_OPEN" -ReleaseId $releaseId -Detail "Excel is running; no add-ins were copied or registered."
    Write-Host "INVSYS_UPDATE_DEFERRED_EXCEL_OPEN"
    exit 0
}

$remoteRelease = Join-Path (Join-Path $rootPath "Releases") $releaseId
$manifest = Assert-InvSysReleaseManifest -ReleaseDirectory $remoteRelease -ExpectedReleaseId $releaseId
$cacheReleases = Join-Path $CacheRoot "Releases"
if (-not (Test-Path -LiteralPath $cacheReleases)) { New-Item -ItemType Directory -Path $cacheReleases -Force | Out-Null }
$cachedRelease = Join-Path $cacheReleases $releaseId
if (Test-Path -LiteralPath $cachedRelease) {
    try { [void](Assert-InvSysReleaseManifest -ReleaseDirectory $cachedRelease -ExpectedReleaseId $releaseId) }
    catch { Remove-Item -LiteralPath $cachedRelease -Recurse -Force }
}
if (-not (Test-Path -LiteralPath $cachedRelease)) {
    $staging = Join-Path $cacheReleases ("." + $releaseId + ".staging-" + [guid]::NewGuid().ToString("N"))
    try {
        New-Item -ItemType Directory -Path $staging -Force | Out-Null
        foreach ($package in @($manifest.packages)) { Copy-Item -LiteralPath (Join-Path $remoteRelease ([string]$package.name)) -Destination (Join-Path $staging ([string]$package.name)) -Force }
        Copy-Item -LiteralPath (Join-Path $remoteRelease "release-manifest.json") -Destination (Join-Path $staging "release-manifest.json") -Force
        [void](Assert-InvSysReleaseManifest -ReleaseDirectory $staging -ExpectedReleaseId $releaseId)
        Move-Item -LiteralPath $staging -Destination $cachedRelease
    }
    finally { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force } }
}

$priorPointerPath = Join-Path $CacheRoot "current-release.json"
$priorPointer = if (Test-Path -LiteralPath $priorPointerPath) { Get-Content -LiteralPath $priorPointerPath -Raw } else { $null }
$priorExcelOptions = Get-InvSysRegistrySnapshot -Path $ExcelOptionsKey
$priorAddinManager = Get-InvSysRegistrySnapshot -Path $AddinManagerKey
try {
    & $RegisterScriptPath -AddinsRoot $cachedRelease -ExcelOptionsKey $ExcelOptionsKey -AddinManagerKey $AddinManagerKey -ExcelProcessName $ExcelProcessName
    Write-InvSysJsonAtomic -Path $priorPointerPath -Value ([ordered]@{ releaseId = $releaseId; manifest = "Releases/$releaseId/release-manifest.json"; appliedAtUtc = [DateTime]::UtcNow.ToString("o") })
    Write-InvSysUpdateStatus -CacheRoot $CacheRoot -Status "APPLIED" -ReleaseId $releaseId -Detail "Hash-verified five-package release registered for next Excel startup."
}
catch {
    Restore-InvSysRegistrySnapshot -Snapshot $priorExcelOptions
    Restore-InvSysRegistrySnapshot -Snapshot $priorAddinManager
    if ($null -ne $priorPointer) { Set-Content -LiteralPath $priorPointerPath -Value $priorPointer -Encoding UTF8 }
    Write-InvSysUpdateStatus -CacheRoot $CacheRoot -Status "FAILED_RESTORED" -ReleaseId $releaseId -Detail "Registration failed; prior known-good pointer was preserved."
    throw
}

$keep = @($releaseId)
foreach ($directory in @(Get-ChildItem -LiteralPath $cacheReleases -Directory | Sort-Object LastWriteTimeUtc -Descending)) {
    if ($keep.Count -ge 3) { break }
    if ($keep -notcontains $directory.Name) { $keep += $directory.Name }
}
foreach ($directory in @(Get-ChildItem -LiteralPath $cacheReleases -Directory)) {
    if ($keep -notcontains $directory.Name) { Remove-Item -LiteralPath $directory.FullName -Recurse -Force }
}
Write-Host "INVSYS_UPDATE_APPLIED"
Write-Host "ReleaseId=$releaseId"
