[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current",

    [Parameter(Mandatory = $false)]
    [string]$AddinsRoot = "",

    [Parameter(Mandatory = $false)]
    [string]$ExcelOptionsKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options",

    [Parameter(Mandatory = $false)]
    [string]$AddinManagerKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager",

    [Parameter(Mandatory = $false)]
    [string]$ExcelProcessName = "EXCEL"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoPath = (Resolve-Path $RepoRoot).Path
$deployPath = if ([string]::IsNullOrWhiteSpace($AddinsRoot)) {
    Join-Path $repoPath $DeployRoot
}
else {
    (Resolve-Path -LiteralPath $AddinsRoot).Path
}

$installOrder = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam",
    "invSys.Admin.xlam"
)
$uninstallOrder = @($installOrder.Clone())
[array]::Reverse($uninstallOrder)
$startupOrder = @(
    "invSys.Operations.xlam",
    "invSys.Admin.xlam"
)

foreach ($fileName in $installOrder) {
    $path = Join-Path $deployPath $fileName
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Expected add-in not found: $path"
    }
}

function Remove-InvSysAddinManagerEntries {
    param(
        [string]$RegistryPath
    )

    if (-not (Test-Path $RegistryPath)) { return }
    $props = Get-ItemProperty -Path $RegistryPath
    foreach ($prop in $props.PSObject.Properties) {
        if ($prop.Name -in @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider")) { continue }
        if ($prop.Name -like "*invSys*") {
            Write-Output ("- remove Add-in Manager entry " + $prop.Name)
            Remove-ItemProperty -Path $RegistryPath -Name $prop.Name -ErrorAction SilentlyContinue
        }
    }
}

function Set-ExcelOpenOrder {
    param(
        [string]$RegistryPath,
        [string[]]$OrderedPaths
    )

    if (-not (Test-Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    $props = Get-ItemProperty -Path $RegistryPath
    foreach ($prop in $props.PSObject.Properties) {
        if ($prop.Name -notmatch '^OPEN\d*$') { continue }
        if ([string]$prop.Value -notmatch '(?i)invsys\.') { continue }
        Write-Output ("- remove " + $prop.Name + "=" + [string]$prop.Value)
        Remove-ItemProperty -Path $RegistryPath -Name $prop.Name -ErrorAction SilentlyContinue
    }

    $props = Get-ItemProperty -Path $RegistryPath
    for ($i = 0; $i -lt $OrderedPaths.Count; $i++) {
        $slot = 0
        do {
            $name = if ($slot -eq 0) { "OPEN" } else { "OPEN$slot" }
            $slot++
        } while ($null -ne $props.PSObject.Properties[$name])
        $value = '"' + $OrderedPaths[$i] + '"'
        Write-Output ("- set " + $name + "=" + $value)
        Set-ItemProperty -Path $RegistryPath -Name $name -Value $value -Type String
        $props = Get-ItemProperty -Path $RegistryPath
    }
}

$orderedPaths = @()
foreach ($fileName in $startupOrder) {
    $orderedPaths += (Join-Path $deployPath $fileName)
}

if (Get-Process -Name $ExcelProcessName -ErrorAction SilentlyContinue) {
    throw "Close all Excel windows before registering invSys add-ins."
}

Write-Output "Using registry-only leaf XLAM startup registration..."
Write-Output "- Core and Domain XLAMs are not explicitly opened; Operations/Admin use bridge calls instead of domain references"

Write-Output "Pruning Add-in Manager entries..."
Remove-InvSysAddinManagerEntries -RegistryPath $addinManagerKey

Write-Output "Setting Excel OPEN order..."
Set-ExcelOpenOrder -RegistryPath $excelOptionsKey -OrderedPaths $orderedPaths
