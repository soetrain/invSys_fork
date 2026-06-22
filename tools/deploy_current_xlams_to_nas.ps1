[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$TargetRoot,

    [Parameter(Mandatory = $false)]
    [string]$SourceRoot = "deploy/current",

    [Parameter(Mandatory = $false)]
    [bool]$Backup = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$requiredXlams = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Receiving.xlam",
    "invSys.Shipping.xlam",
    "invSys.Production.xlam",
    "invSys.Admin.xlam"
)

function Resolve-ExistingOrAbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue
    )

    if (Test-Path -LiteralPath $PathValue) {
        $resolved = Resolve-Path -LiteralPath $PathValue
        if (-not [string]::IsNullOrWhiteSpace($resolved.ProviderPath)) {
            return $resolved.ProviderPath
        }

        return $resolved.Path
    }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }

    return Join-Path (Get-Location).Path $PathValue
}

function Test-IsWarehouseRuntimeRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue
    )

    $trimmed = $PathValue.TrimEnd("\", "/")
    return ($trimmed -match "(?i)[\\/]invSysWH\d+$")
}

function Get-ValidatedSourceFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
        throw "Source root not found: $Root"
    }

    foreach ($name in $requiredXlams) {
        $path = Join-Path $Root $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Required XLAM missing: $path"
        }

        $item = Get-Item -LiteralPath $path
        if ($item.Length -le 0) {
            throw "Required XLAM is empty: $path"
        }

        $item
    }
}

$sourceRootPath = Resolve-ExistingOrAbsolutePath -PathValue $SourceRoot
$targetRootPath = Resolve-ExistingOrAbsolutePath -PathValue $TargetRoot
$sourceFiles = @(Get-ValidatedSourceFiles -Root $sourceRootPath)

if (-not (Test-Path -LiteralPath $targetRootPath -PathType Container)) {
    if (Test-IsWarehouseRuntimeRoot -PathValue $targetRootPath) {
        throw "Target appears to be a warehouse runtime root and does not exist. Refusing to create it without explicit setup: $targetRootPath"
    }

    if ($PSCmdlet.ShouldProcess($targetRootPath, "Create target directory")) {
        New-Item -ItemType Directory -Path $targetRootPath -Force | Out-Null
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupRoot = Join-Path $targetRootPath ("_backup_xlam_" + $timestamp)
$copied = @()
$backedUp = @()

foreach ($source in $sourceFiles) {
    $target = Join-Path $targetRootPath $source.Name

    if ($Backup -and (Test-Path -LiteralPath $target -PathType Leaf)) {
        if (-not (Test-Path -LiteralPath $backupRoot -PathType Container)) {
            if ($PSCmdlet.ShouldProcess($backupRoot, "Create backup directory")) {
                New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
            }
        }

        $backupTarget = Join-Path $backupRoot $source.Name
        if ($PSCmdlet.ShouldProcess($target, "Back up to $backupTarget")) {
            Copy-Item -LiteralPath $target -Destination $backupTarget -Force
            $backupItem = Get-Item -LiteralPath $backupTarget
            $backedUp += [pscustomobject]@{
                Name = $backupItem.Name
                Path = $backupItem.FullName
                Size = $backupItem.Length
            }
        }
    }

    if ($PSCmdlet.ShouldProcess($target, "Copy $($source.FullName)")) {
        Copy-Item -LiteralPath $source.FullName -Destination $target -Force
    }

    if (-not $WhatIfPreference) {
        $targetItem = Get-Item -LiteralPath $target
        if ($targetItem.Length -ne $source.Length) {
            throw "Copied file size mismatch for $($source.Name). Source=$($source.Length) Target=$($targetItem.Length)"
        }

        $copied += [pscustomobject]@{
            Name = $source.Name
            SourcePath = $source.FullName
            TargetPath = $targetItem.FullName
            Size = $source.Length
            SourceLastWriteTime = $source.LastWriteTime.ToString("o")
            TargetLastWriteTime = $targetItem.LastWriteTime.ToString("o")
        }
    }
}

if (-not $WhatIfPreference) {
    $manifest = [pscustomobject]@{
        DeployedAt = (Get-Date).ToString("o")
        SourceRoot = $sourceRootPath
        TargetRoot = $targetRootPath
        BackupEnabled = $Backup
        BackupRoot = $(if ($backedUp.Count -gt 0) { $backupRoot } else { $null })
        Files = $copied
        Backups = $backedUp
    }

    $manifestPath = Join-Path $targetRootPath ("deploy_manifest_" + $timestamp + ".json")
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding ASCII
    Write-Host "DEPLOY_XLAMS_OK"
    Write-Host "TargetRoot=$targetRootPath"
    Write-Host "Manifest=$manifestPath"
}
else {
    Write-Host "DEPLOY_XLAMS_WHATIF_OK"
    Write-Host "TargetRoot=$targetRootPath"
}
