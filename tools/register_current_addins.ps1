[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoPath = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repoPath $DeployRoot

$installOrder = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Receiving.xlam",
    "invSys.Shipping.xlam",
    "invSys.Production.xlam",
    "invSys.Admin.xlam"
)
$uninstallOrder = @($installOrder.Clone())
[array]::Reverse($uninstallOrder)

foreach ($fileName in $installOrder) {
    $path = Join-Path $deployPath $fileName
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Expected add-in not found: $path"
    }
}

$excelOptionsKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$addinManagerKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager"

function Release-ComObject {
    param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Find-AddInByName {
    param(
        [object]$Excel,
        [string]$Name
    )

    foreach ($addin in $Excel.AddIns) {
        if ([string]::Equals([string]$addin.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $addin
        }
    }
    return $null
}

function Remove-StaleAddinManagerEntries {
    param(
        [string]$RegistryPath,
        [string]$KeepPrefix
    )

    if (-not (Test-Path $RegistryPath)) { return }
    $props = Get-ItemProperty -Path $RegistryPath
    foreach ($prop in $props.PSObject.Properties) {
        if ($prop.Name -in @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider")) { continue }
        if ($prop.Name -like "*invSys*" -and -not $prop.Name.StartsWith($KeepPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Output ("- remove stale Add-in Manager entry " + $prop.Name)
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

    for ($i = 0; $i -lt 10; $i++) {
        $name = if ($i -eq 0) { "OPEN" } else { "OPEN$i" }
        if ($i -lt $OrderedPaths.Count) {
            $value = '"' + $OrderedPaths[$i] + '"'
            Write-Output ("- set " + $name + "=" + $value)
            Set-ItemProperty -Path $RegistryPath -Name $name -Value $value -Type String
        }
        else {
            Remove-ItemProperty -Path $RegistryPath -Name $name -ErrorAction SilentlyContinue
        }
    }
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    Write-Output "Disabling registered invSys add-ins..."
    foreach ($fileName in $uninstallOrder) {
        $addin = Find-AddInByName -Excel $excel -Name $fileName
        if ($null -ne $addin -and [bool]$addin.Installed) {
            Write-Output ("- disable " + $fileName + " (" + $addin.FullName + ")")
            $addin.Installed = $false
        }
    }

    Write-Output "Registering deploy/current add-ins..."
    foreach ($fileName in $installOrder) {
        $targetPath = Join-Path $deployPath $fileName
        $addin = Find-AddInByName -Excel $excel -Name $fileName
        if ($null -eq $addin -or -not [string]::Equals([string]$addin.FullName, $targetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Output ("- register " + $targetPath)
            $addin = $excel.AddIns.Add($targetPath, $false)
        }
        if (-not [bool]$addin.Installed) {
            Write-Output ("- enable " + $fileName)
            $addin.Installed = $true
        }
    }

    Write-Output "Pruning stale registry entries..."
    Remove-StaleAddinManagerEntries -RegistryPath $addinManagerKey -KeepPrefix $deployPath

    Write-Output "Setting Excel OPEN order..."
    $orderedPaths = @()
    foreach ($fileName in $installOrder) {
        $orderedPaths += (Join-Path $deployPath $fileName)
    }
    Set-ExcelOpenOrder -RegistryPath $excelOptionsKey -OrderedPaths $orderedPaths

    Write-Output ""
    Write-Output "Active invSys add-ins:"
    foreach ($fileName in $installOrder) {
        $addin = Find-AddInByName -Excel $excel -Name $fileName
        if ($null -ne $addin) {
            Write-Output ("- " + $addin.Name + " | Installed=" + [string][bool]$addin.Installed + " | " + $addin.FullName)
        }
    }
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}
