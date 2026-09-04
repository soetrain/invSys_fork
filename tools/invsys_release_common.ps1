Set-StrictMode -Version Latest

function Get-InvSysReleasePackageNames {
    return @(
        "invSys.Core.xlam",
        "invSys.Inventory.Domain.xlam",
        "invSys.Designs.Domain.xlam",
        "invSys.Operations.xlam",
        "invSys.Admin.xlam"
    )
}

function Assert-InvSysReleaseId {
    param([Parameter(Mandatory = $true)][string]$ReleaseId)
    if ($ReleaseId -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,119}$') {
        throw "ReleaseId is invalid. Use letters, digits, dot, underscore, or hyphen only."
    }
}

function Read-InvSysJsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required JSON file was not found: $Path"
    }
    try { return (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json) }
    catch { throw "Invalid JSON file ${Path}: $($_.Exception.Message)" }
}

function Write-InvSysJsonAtomic {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )
    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $temporary = Join-Path $parent (([IO.Path]::GetFileName($Path)) + ".new-" + [guid]::NewGuid().ToString("N"))
    try {
        $Value | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $temporary -Encoding UTF8
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Assert-InvSysReleaseManifest {
    param(
        [Parameter(Mandatory = $true)][string]$ReleaseDirectory,
        [Parameter(Mandatory = $false)][string]$ExpectedReleaseId = ""
    )
    $manifestPath = Join-Path $ReleaseDirectory "release-manifest.json"
    $manifest = Read-InvSysJsonFile -Path $manifestPath
    $releaseId = [string]$manifest.releaseId
    Assert-InvSysReleaseId -ReleaseId $releaseId
    if (-not [string]::IsNullOrWhiteSpace($ExpectedReleaseId) -and $releaseId -ne $ExpectedReleaseId) {
        throw "Release manifest identity mismatch. Expected $ExpectedReleaseId but found $releaseId."
    }
    if ([string]$manifest.packageSetVersion -ne "R1-5") {
        throw "Release $releaseId is not an R1-5 package set."
    }
    $expectedNames = @(Get-InvSysReleasePackageNames | Sort-Object)
    $packages = @($manifest.packages)
    $actualNames = @($packages | ForEach-Object { [string]$_.name } | Sort-Object)
    if ($actualNames.Count -ne $expectedNames.Count -or @(Compare-Object $expectedNames $actualNames).Count -ne 0) {
        throw "Release $releaseId does not contain exactly the five normative XLAM packages."
    }
    foreach ($package in $packages) {
        $name = [string]$package.name
        $hash = ([string]$package.sha256).ToLowerInvariant()
        if ($hash -notmatch '^[a-f0-9]{64}$') { throw "Release $releaseId has an invalid SHA-256 for $name." }
        $packagePath = Join-Path $ReleaseDirectory $name
        if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf)) {
            throw "Release $releaseId is incomplete; missing $name."
        }
        $actualHash = (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actualHash -ne $hash) { throw "Release $releaseId hash mismatch for $name." }
    }
    return $manifest
}

function Get-InvSysReleasePointer {
    param([Parameter(Mandatory = $true)][string]$Root)
    $pointer = Read-InvSysJsonFile -Path (Join-Path $Root "current-release.json")
    Assert-InvSysReleaseId -ReleaseId ([string]$pointer.releaseId)
    return $pointer
}

function Write-InvSysUpdateStatus {
    param(
        [Parameter(Mandatory = $true)][string]$CacheRoot,
        [Parameter(Mandatory = $true)][string]$Status,
        [string]$ReleaseId = "",
        [string]$Detail = ""
    )
    Write-InvSysJsonAtomic -Path (Join-Path $CacheRoot "update-status.json") -Value ([ordered]@{
        status = $Status
        releaseId = $ReleaseId
        timestampUtc = [DateTime]::UtcNow.ToString("o")
        detail = $Detail
    })
}

function Get-InvSysRegistrySnapshot {
    param([Parameter(Mandatory = $true)][string]$Path)
    $exists = Test-Path -LiteralPath $Path
    $values = @()
    if ($exists) {
        $item = Get-ItemProperty -Path $Path
        $values = @($item.PSObject.Properties |
            Where-Object { $_.Name -notin @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider") } |
            ForEach-Object { [pscustomobject]@{ Name = $_.Name; Value = $_.Value } })
    }
    return [pscustomobject]@{ Path = $Path; Exists = $exists; Values = $values }
}

function Restore-InvSysRegistrySnapshot {
    param([Parameter(Mandatory = $true)][object]$Snapshot)
    if (-not $Snapshot.Exists) {
        if (Test-Path -LiteralPath $Snapshot.Path) { Remove-Item -LiteralPath $Snapshot.Path -Recurse -Force }
        return
    }
    if (-not (Test-Path -LiteralPath $Snapshot.Path)) { New-Item -Path $Snapshot.Path -Force | Out-Null }
    $current = Get-ItemProperty -Path $Snapshot.Path
    foreach ($property in $current.PSObject.Properties) {
        if ($property.Name -in @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider")) { continue }
        Remove-ItemProperty -Path $Snapshot.Path -Name $property.Name -ErrorAction SilentlyContinue
    }
    foreach ($value in @($Snapshot.Values)) {
        Set-ItemProperty -Path $Snapshot.Path -Name ([string]$value.Name) -Value $value.Value -Type String
    }
}
