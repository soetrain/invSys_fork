<#
.SYNOPSIS
Scaffold build script for invSys XLAM packaging.

.DESCRIPTION
Provides a stable command entry point for build automation while keeping
current behavior non-destructive. This script reports intended outputs and
does not modify existing add-ins unless explicit implementation is added.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SourceRoot = "src",

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = "deploy/current",

    [Parameter(Mandatory = $false)]
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "invSys build-xlam.ps1"
Write-Host "SourceRoot: $SourceRoot"
Write-Host "OutputRoot: $OutputRoot"

if (-not (Test-Path -LiteralPath $SourceRoot)) {
    throw "Source root not found: $SourceRoot"
}

if (-not (Test-Path -LiteralPath $OutputRoot)) {
    New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
}

Write-Host "Planned action:"
Write-Host "- Build role/domain/core XLAM artifacts from exported VBA sources."
Write-Host "- Place release-ready files in '$OutputRoot'."

if (-not $Apply) {
    Write-Host "Dry run only. Re-run with -Apply after build mapping is finalized."
    exit 0
}

Write-Warning "Build implementation is intentionally deferred to prevent accidental overwrite."
Write-Warning "Add explicit project-to-XLAM mapping before enabling writes."
