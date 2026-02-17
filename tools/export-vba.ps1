<#
.SYNOPSIS
Scaffold export script for invSys VBA sources.

.DESCRIPTION
This script is intentionally conservative. It validates inputs and prints
the expected export plan without modifying existing source by default.
Use this as the standard entry point referenced by the design document.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = "src",

    [Parameter(Mandatory = $false)]
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "invSys export-vba.ps1"
Write-Host "WorkbookPath: $WorkbookPath"
Write-Host "OutputRoot:   $OutputRoot"

if (-not $WorkbookPath) {
    Write-Warning "No workbook provided. Example:"
    Write-Warning "  .\\tools\\export-vba.ps1 -WorkbookPath C:\\path\\invSys.Receiving.xlam"
    Write-Host "No changes made."
    exit 0
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

if (-not (Test-Path -LiteralPath $OutputRoot)) {
    New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
}

Write-Host "Planned action:"
Write-Host "- Export VBA components from workbook into repository structure under '$OutputRoot'."

if (-not $Apply) {
    Write-Host "Dry run only. Re-run with -Apply to enable implementation once finalized."
    exit 0
}

Write-Warning "Export implementation is intentionally deferred to avoid accidental overwrite."
Write-Warning "Add explicit component mapping before enabling write operations."
