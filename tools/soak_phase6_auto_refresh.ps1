[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [int]$IntervalSeconds = 2,

    [Parameter(Mandatory = $false)]
    [int]$WaitSeconds = 5
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Run-WorkbookMacro {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName,
        [object[]]$Arguments = @()
    )

    $macro = "'$WorkbookName'!$MacroName"
    switch ($Arguments.Count) {
        0 { return $Excel.Run($macro) }
        1 { return $Excel.Run($macro, $Arguments[0]) }
        2 { return $Excel.Run($macro, $Arguments[0], $Arguments[1]) }
        3 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2]) }
        4 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3]) }
        5 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4]) }
        default { throw "Too many macro arguments for $MacroName" }
    }
}

function Resolve-WorkbookByName {
    param(
        [object]$Excel,
        [string]$WorkbookName
    )

    foreach ($wb in $Excel.Workbooks) {
        if ([string]::Equals([string]$wb.Name, $WorkbookName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $wb
        }
    }
    return $null
}

function Get-Worksheet {
    param(
        [object]$Workbook,
        [string]$WorksheetName
    )

    if ($null -eq $Workbook) { return $null }
    try { return $Workbook.Worksheets.Item($WorksheetName) } catch { return $null }
}

function Get-ListObject {
    param(
        [object]$Worksheet,
        [string]$TableName
    )

    if ($null -eq $Worksheet) { return $null }
    try { return $Worksheet.ListObjects.Item($TableName) } catch { return $null }
}

function Get-Shape {
    param(
        [object]$Worksheet,
        [string]$ShapeName
    )

    if ($null -eq $Worksheet) { return $null }
    try { return $Worksheet.Shapes.Item($ShapeName) } catch { return $null }
}

function Get-TableValue {
    param(
        [object]$ListObject,
        [int]$RowIndex,
        [string]$ColumnName
    )

    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) { return $null }
    $colIndex = $ListObject.ListColumns.Item($ColumnName).Index
    return $ListObject.DataBodyRange.Cells($RowIndex, $colIndex).Value2
}

function Set-TableValue {
    param(
        [object]$ListObject,
        [int]$RowIndex,
        [string]$ColumnName,
        [object]$Value
    )

    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) { return }
    $colIndex = $ListObject.ListColumns.Item($ColumnName).Index
    $valueToWrite = $Value
    if ($Value -is [bool]) {
        $valueToWrite = if ($Value) { "TRUE" } else { "FALSE" }
    }
    elseif ($Value -is [datetime]) {
        $valueToWrite = [datetime]$Value
    }
    elseif ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        $valueToWrite = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", $Value)
    }
    $ListObject.DataBodyRange.Cells($RowIndex, $colIndex).Value = $valueToWrite
}

function Write-ResultFile {
    param(
        [string]$Path,
        [hashtable]$Data
    )

    $lines = @()
    $lines += "# Phase 6 Auto Refresh Soak Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    foreach ($key in ($Data.Keys | Sort-Object)) {
        $lines += "- ${key}: $($Data[$key])"
    }
    [System.IO.File]::WriteAllLines($Path, $lines)
}

function Convert-ExcelDateValue {
    param([object]$Value)

    if ($null -eq $Value) { return [datetime]::MinValue }
    if ($Value -is [datetime]) { return [datetime]$Value }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return [datetime]::MinValue }

    try {
        return [datetime]::FromOADate([double]$Value)
    }
    catch {
        return [datetime]::Parse($text)
    }
}

function Convert-ExcelDoubleValue {
    param([object]$Value)

    if ($null -eq $Value) { return 0.0 }
    if ($Value -is [double]) { return [double]$Value }
    if ($Value -is [float]) { return [double]$Value }
    if ($Value -is [decimal]) { return [double]$Value }
    if ($Value -is [int]) { return [double]$Value }
    if ($Value -is [long]) { return [double]$Value }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return 0.0 }
    return [double]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture)
}

$repo = (Resolve-Path $RepoRoot).Path
$boundaryScript = Join-Path $repo "tools/run_phase6_lan_boundary_validation.ps1"
$setupScript = Join-Path $repo "tools/setup_lan_station.ps1"
$deployPath = Join-Path $repo "deploy/current"
$resultPath = Join-Path $repo "tests/unit/phase6_auto_refresh_soak_results.md"

$boundaryOutput = & powershell -NoProfile -ExecutionPolicy Bypass -File $boundaryScript -RepoRoot $repo 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "LAN boundary seed failed.`n$($boundaryOutput -join [Environment]::NewLine)"
}

$sessionRootLine = ($boundaryOutput | Where-Object { $_ -like "SESSION_ROOT=*" } | Select-Object -Last 1)
if ([string]::IsNullOrWhiteSpace($sessionRootLine)) {
    throw "SESSION_ROOT was not returned by run_phase6_lan_boundary_validation.ps1"
}

$sessionRoot = $sessionRootLine.Substring("SESSION_ROOT=".Length)
$warehouseId = "WH89"
$stationId = "S5"
$sharedRoot = Join-Path $sessionRoot "runtime"
$sharedConfigPath = Join-Path $sharedRoot ($warehouseId + ".invSys.Config.xlsb")
$sharedSnapshotPath = Join-Path $sharedRoot ($warehouseId + ".invSys.Snapshot.Inventory.xlsb")
$localConfigRoot = Join-Path $sessionRoot "soak_station_config"
$stationInboxRoot = Join-Path $sessionRoot "soak_station_inbox"
$operatorPath = Join-Path $sessionRoot "soak_station_operator.xlsb"

New-Item -ItemType Directory -Path $localConfigRoot -Force | Out-Null
New-Item -ItemType Directory -Path $stationInboxRoot -Force | Out-Null

$excel = $null
$wbConfig = $null
$wbCore = $null
$wbInventoryDomain = $null
$wbReceiving = $null
$wbLocalConfig = $null
$wbOperator = $null
$wbSnapshot = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    $wbConfig = $excel.Workbooks.Open($sharedConfigPath)
    $loWh = Get-ListObject -Worksheet (Get-Worksheet -Workbook $wbConfig -WorksheetName "WarehouseConfig") -TableName "tblWarehouseConfig"
    if ($null -eq $loWh) { throw "tblWarehouseConfig was not found in shared config workbook." }
    Set-TableValue -ListObject $loWh -RowIndex 1 -ColumnName "FF_AutoSnapshot" -Value $true
    Set-TableValue -ListObject $loWh -RowIndex 1 -ColumnName "AutoRefreshIntervalSeconds" -Value $IntervalSeconds
    $wbConfig.Save()
    $wbConfig.Close($false)
    $wbConfig = $null

    $setupOutput = & powershell -NoProfile -ExecutionPolicy Bypass -File $setupScript `
        -RepoRoot $repo `
        -WarehouseId $warehouseId `
        -StationId $stationId `
        -SharedRuntimeRoot $sharedRoot `
        -StationInboxRoot $stationInboxRoot `
        -LocalConfigRoot $localConfigRoot `
        -RoleDefault "RECEIVE" `
        -StationName "SOAK-STATION-05" `
        -StationUserId $env:USERNAME `
        -CreateOperatorWorkbook `
        -OperatorWorkbookPath $operatorPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "setup_lan_station.ps1 failed.`n$($setupOutput -join [Environment]::NewLine)"
    }

    $roleReadyLine = ($setupOutput | Where-Object { $_ -like "RoleReady=*" } | Select-Object -Last 1)
    if ([string]::IsNullOrWhiteSpace($roleReadyLine) -or -not $roleReadyLine.EndsWith("True", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "setup_lan_station.ps1 did not reach RoleReady=True.`n$($setupOutput -join [Environment]::NewLine)"
    }

    $wbCore = $excel.Workbooks.Open((Join-Path $deployPath "invSys.Core.xlam"))
    $wbInventoryDomain = $excel.Workbooks.Open((Join-Path $deployPath "invSys.Inventory.Domain.xlam"))
    $wbReceiving = $excel.Workbooks.Open((Join-Path $deployPath "invSys.Receiving.xlam"))
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $wbCore.Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($sharedRoot))

    $wbLocalConfig = $excel.Workbooks.Open((Join-Path $localConfigRoot ($warehouseId + ".invSys.Config.xlsb")))
    $wbOperator = $excel.Workbooks.Open($operatorPath)
    [void]$wbOperator.Activate()
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $wbReceiving.Name -MacroName "modReceivingInit.InitReceivingAddin")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $wbReceiving.Name -MacroName "modReceivingInit.EnsureReceivingSurfaceForWorkbook" -Arguments @($wbOperator))

    $wsInv = Get-Worksheet -Workbook $wbOperator -WorksheetName "InventoryManagement"
    $loInv = Get-ListObject -Worksheet $wsInv -TableName "invSys"
    if ($null -eq $loInv -or $null -eq $loInv.DataBodyRange) {
        throw "Operator workbook invSys table was not available for soak validation."
    }

    $initialQty = Convert-ExcelDoubleValue -Value (Get-TableValue -ListObject $loInv -RowIndex 1 -ColumnName "TOTAL INV")
    $initialRefreshRaw = Get-TableValue -ListObject $loInv -RowIndex 1 -ColumnName "LastRefreshUTC"
    $initialRefresh = Convert-ExcelDateValue -Value $initialRefreshRaw

    $wbSnapshot = $excel.Workbooks.Open($sharedSnapshotPath)
    $loSnap = Get-ListObject -Worksheet (Get-Worksheet -Workbook $wbSnapshot -WorksheetName "InventorySnapshot") -TableName "tblInventorySnapshot"
    if ($null -eq $loSnap -or $null -eq $loSnap.DataBodyRange) {
        throw "Shared snapshot table was not available for soak mutation."
    }

    $soakQty = $initialQty + 9
    Set-TableValue -ListObject $loSnap -RowIndex 1 -ColumnName "QtyOnHand" -Value $soakQty
    Set-TableValue -ListObject $loSnap -RowIndex 1 -ColumnName "QtyAvailable" -Value $soakQty
    $wbSnapshot.Save()
    $wbSnapshot.Close($false)
    $wbSnapshot = $null

    Start-Sleep -Seconds $WaitSeconds

    $wsInv = Get-Worksheet -Workbook $wbOperator -WorksheetName "InventoryManagement"
    $loInv = Get-ListObject -Worksheet $wsInv -TableName "invSys"
    $postQty = Convert-ExcelDoubleValue -Value (Get-TableValue -ListObject $loInv -RowIndex 1 -ColumnName "TOTAL INV")
    $postRefreshRaw = Get-TableValue -ListObject $loInv -RowIndex 1 -ColumnName "LastRefreshUTC"
    $postRefresh = Convert-ExcelDateValue -Value $postRefreshRaw
    $postIsStale = [string](Get-TableValue -ListObject $loInv -RowIndex 1 -ColumnName "IsStale")
    $statusShape = Get-Shape -Worksheet $wsInv -ShapeName "invSysReadModelStatus"
    $statusText = if ($null -eq $statusShape) { "" } else { [string]$statusShape.TextFrame.Characters().Text }

    $passed = ($postQty -eq $soakQty) `
        -and ($postRefresh -gt $initialRefresh) `
        -and ($postIsStale.Trim().ToUpperInvariant() -in @("FALSE", "0")) `
        -and ($statusText -like "*INVENTORY SNAPSHOT CURRENT*")

    $result = @{
        SessionRoot = $sessionRoot
        WarehouseId = $warehouseId
        StationId = $stationId
        IntervalSeconds = $IntervalSeconds
        WaitSeconds = $WaitSeconds
        InitialQty = $initialQty
        SoakQty = $soakQty
        PostQty = $postQty
        InitialRefreshUTC = $initialRefresh.ToString("yyyy-MM-dd HH:mm:ss")
        PostRefreshUTC = $postRefresh.ToString("yyyy-MM-dd HH:mm:ss")
        IsStale = $postIsStale
        StatusText = $statusText
        Passed = $passed
    }
    Write-ResultFile -Path $resultPath -Data $result

    if (-not $passed) {
        throw "Auto refresh soak failed. Results written to $resultPath"
    }

    Write-Output "PHASE6_AUTO_REFRESH_SOAK_OK"
    Write-Output ("RESULTS=" + $resultPath)
    Write-Output ("SESSION_ROOT=" + $sessionRoot)
    Write-Output ("INITIAL_QTY=" + $initialQty)
    Write-Output ("POST_QTY=" + $postQty)
    Write-Output ("INITIAL_REFRESH_UTC=" + $result.InitialRefreshUTC)
    Write-Output ("POST_REFRESH_UTC=" + $result.PostRefreshUTC)
}
catch {
    Write-Error ("SOAK_FAIL line " + $_.InvocationInfo.ScriptLineNumber + " :: " + $_.InvocationInfo.PositionMessage + " :: " + $_.Exception.Message)
    throw
}
finally {
    foreach ($wb in @($wbSnapshot, $wbOperator, $wbLocalConfig, $wbReceiving, $wbInventoryDomain, $wbCore, $wbConfig)) {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
            Release-ComObject $wb
        }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}
