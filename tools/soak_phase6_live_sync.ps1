[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [int]$Iterations = 5,

    [Parameter(Mandatory = $false)]
    [int]$PollTimeoutSeconds = 20,

    [Parameter(Mandatory = $false)]
    [int]$PollIntervalMilliseconds = 250,

    [Parameter(Mandatory = $false)]
    [int]$AutoRefreshIntervalSeconds = 2
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
        6 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5]) }
        default { throw "Too many macro arguments for $MacroName" }
    }
}

function Import-BasModule {
    param(
        [object]$VbProject,
        [string]$BasPath
    )

    if (-not (Test-Path $BasPath)) {
        throw "Missing BAS module: $BasPath"
    }
    [void]$VbProject.VBComponents.Import($BasPath)
}

function Add-BootstrapModule {
    param([object]$Workbook)

    $comp = $Workbook.VBProject.VBComponents.Add(1)
    $comp.Name = "modHarnessBootstrap"
    $comp.CodeModule.AddFromString("Public Function HarnessPing() As Long: HarnessPing = 1: End Function")
    return $comp
}

function New-HarnessWorkbook {
    param(
        [object]$Excel,
        [string]$HarnessPath,
        [string[]]$ModulePaths
    )

    $wb = $Excel.Workbooks.Add()
    [void](Add-BootstrapModule -Workbook $wb)
    $vbProject = $wb.VBProject
    [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $wb.Name -MacroName "HarnessPing")

    foreach ($modulePath in $ModulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $modulePath
        [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $wb.Name -MacroName "HarnessPing")
    }

    $wb.SaveAs($HarnessPath, 52)
    return $wb
}

function New-ExcelApp {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1
    return $excel
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

function Ensure-ListObjectRow {
    param([object]$ListObject)

    if ($null -eq $ListObject) { throw "ListObject missing." }
    if ($ListObject.ListRows.Count -gt 0) {
        return $ListObject.ListRows.Item(1).Range
    }

    try {
        $row = $ListObject.ListRows.Add($null, $false)
    }
    catch {
        $row = $ListObject.ListRows.Add()
    }
    return $row.Range
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
    if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        $valueToWrite = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", $Value)
    }
    $ListObject.DataBodyRange.Cells($RowIndex, $colIndex).Value = $valueToWrite
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

function Update-SharedConfig {
    param(
        [string]$ConfigPath,
        [int]$AutoRefreshSeconds
    )

    $excel = $null
    $wb = $null
    try {
        $excel = New-ExcelApp
        $wb = $excel.Workbooks.Open($ConfigPath)
        $ws = Get-Worksheet -Workbook $wb -WorksheetName "WarehouseConfig"
        $lo = Get-ListObject -Worksheet $ws -TableName "tblWarehouseConfig"
        if ($null -eq $lo) { throw "tblWarehouseConfig was not found in shared config workbook." }
        Set-TableValue -ListObject $lo -RowIndex 1 -ColumnName "FF_AutoSnapshot" -Value "TRUE"
        Set-TableValue -ListObject $lo -RowIndex 1 -ColumnName "AutoRefreshIntervalSeconds" -Value $AutoRefreshSeconds
        $wb.Save()
    }
    finally {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
            Release-ComObject $wb
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            Release-ComObject $excel
        }
    }
}

function Get-SyncLogTail {
    param(
        [string]$Path,
        [int]$TailCount = 6
    )

    if (-not (Test-Path $Path)) { return "" }
    try {
        return ((Get-Content -Path $Path -Tail $TailCount) -join " || ")
    }
    catch {
        return ""
    }
}

function Write-ResultFile {
    param(
        [string]$Path,
        [hashtable]$Summary,
        [System.Collections.Generic.List[object]]$Rows
    )

    $lines = @()
    $lines += "# Phase 6 Bidirectional Live Sync Soak Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    foreach ($key in ($Summary.Keys | Sort-Object)) {
        $lines += "- ${key}: $($Summary[$key])"
    }
    $lines += ""
    $lines += "| Iteration | Direction | Qty | ExpectedTotal | Processed | BatchMs | CatchupMs | RuntimeReadOnly | ObservedTotal | ObservedReceived | Result | Detail |"
    $lines += "|---|---|---:|---:|---:|---:|---:|---|---:|---:|---|---|"
    foreach ($row in $Rows) {
        $lines += "| $($row.Iteration) | $($row.Direction) | $($row.Qty) | $($row.ExpectedTotal) | $($row.Processed) | $($row.BatchMs) | $($row.CatchupMs) | $($row.RuntimeReadOnly) | $($row.ObservedTotal) | $($row.ObservedReceived) | $($row.Result) | $($row.Detail) |"
    }
    [System.IO.File]::WriteAllLines($Path, $lines)
}

function Wait-ForProjection {
    param(
        [object]$ListObject,
        [double]$ExpectedTotal,
        [double]$ExpectedReceived,
        [int]$TimeoutSeconds,
        [int]$PollIntervalMilliseconds,
        [string]$SyncLogPath = ""
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $observedTotal = 0.0
    $observedReceived = 0.0
    $runtimeReadOnly = ""
    $detail = ""

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        Start-Sleep -Milliseconds $PollIntervalMilliseconds
        $observedTotal = Convert-ExcelDoubleValue (Get-TableValue -ListObject $ListObject -RowIndex 1 -ColumnName "TOTAL INV")
        $observedReceived = Convert-ExcelDoubleValue (Get-TableValue -ListObject $ListObject -RowIndex 1 -ColumnName "RECEIVED")
        if ($SyncLogPath -ne "") {
            $detail = Get-SyncLogTail -Path $SyncLogPath -TailCount 10
            if ($detail -match 'RuntimeReadOnly=(True|False)') {
                $runtimeReadOnly = $Matches[1]
            }
        }
        if (($observedTotal -eq $ExpectedTotal) -and ($observedReceived -eq $ExpectedReceived)) {
            $sw.Stop()
            return [pscustomobject]@{
                Passed          = $true
                CatchupMs       = [math]::Round($sw.Elapsed.TotalMilliseconds, 0)
                ObservedTotal   = $observedTotal
                ObservedReceived= $observedReceived
                RuntimeReadOnly = $runtimeReadOnly
                Detail          = $detail
            }
        }
    }

    $sw.Stop()
    if ($detail -eq "" -and $SyncLogPath -ne "") {
        $detail = Get-SyncLogTail -Path $SyncLogPath -TailCount 12
        if ($detail -match 'RuntimeReadOnly=(True|False)') {
            $runtimeReadOnly = $Matches[1]
        }
    }
    return [pscustomobject]@{
        Passed           = $false
        CatchupMs        = [math]::Round($sw.Elapsed.TotalMilliseconds, 0)
        ObservedTotal    = $observedTotal
        ObservedReceived = $observedReceived
        RuntimeReadOnly  = $runtimeReadOnly
        Detail           = $detail
    }
}

$repo = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repo "deploy/current"
$boundaryScript = Join-Path $repo "tools/run_phase6_lan_boundary_validation.ps1"
$setupScript = Join-Path $repo "tools/setup_lan_station.ps1"
$resultPath = Join-Path $repo "tests/unit/phase6_live_sync_soak_results.md"
$fixturesRoot = Join-Path $repo "tests/fixtures"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$sessionRoot = Join-Path $fixturesRoot "phase6_live_sync_$stamp"
$canonicalRoot = Join-Path $sessionRoot "runtime"
$stationConfigRoot = Join-Path $sessionRoot "station_s2_config"
$stationInboxRoot = Join-Path $sessionRoot "station_s2_inbox"
$operatorPath = Join-Path $sessionRoot "WH89_S2_Receiving_Operator.xlsb"
$sourceWorkbookPath = Join-Path $sessionRoot "WH89_FRODECO.inventory_management.xlsb"
$warehouseId = "WH89"
$stationA = "S1"
$stationB = "S2"
$sku = "SKU-LAN-BOUNDARY-001"
$sharedConfigPath = Join-Path $canonicalRoot ($warehouseId + ".invSys.Config.xlsb")

$modulePaths = @(
    (Join-Path $repo "src/Core/Modules/modConfigDefaults.bas"),
    (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
    (Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas"),
    (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas"),
    (Join-Path $repo "src/Core/Modules/modRoleEventJson.bas"),
    (Join-Path $repo "src/Core/Modules/modSystemIdentity.bas"),
    (Join-Path $repo "src/Core/Modules/modOperatorReadModel.bas"),
    (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
    (Join-Path $repo "src/Core/Modules/modWarehouseSync.bas"),
    (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
    (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
    (Join-Path $repo "src/Core/Modules/modConfig.bas"),
    (Join-Path $repo "src/Core/Modules/modStationIdentity.bas"),
    (Join-Path $repo "src/Core/Modules/modAuth.bas"),
    (Join-Path $repo "src/Core/Modules/modAuthSession.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventoryBridgeApi.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
    (Join-Path $repo "tests/unit/TestPhase2Helpers.bas"),
    (Join-Path $repo "tests/unit/TestPhase6LanBoundary.bas")
)

$excelSource = $null
$excelOperator = $null
$excelHarness = $null
$wbCoreSource = $null
$wbInventoryDomainSource = $null
$wbSource = $null
$wbCoreOperator = $null
$wbInventoryDomainOperator = $null
$wbReceivingOperator = $null
$wbLocalConfig = $null
$wbOperator = $null
$wbHarness = $null
$sourceLogPath = ""
$rows = New-Object 'System.Collections.Generic.List[object]'

try {
    if (Test-Path $sessionRoot) { Remove-Item $sessionRoot -Recurse -Force }
    $boundaryOutput = & powershell -NoProfile -ExecutionPolicy Bypass -File $boundaryScript -RepoRoot $repo 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "LAN boundary seed failed.`n$($boundaryOutput -join [Environment]::NewLine)"
    }
    $sessionRootLine = ($boundaryOutput | Where-Object { $_ -like "SESSION_ROOT=*" } | Select-Object -Last 1)
    if ([string]::IsNullOrWhiteSpace($sessionRootLine)) {
        throw "SESSION_ROOT was not returned by run_phase6_lan_boundary_validation.ps1"
    }
    $sessionRoot = $sessionRootLine.Substring("SESSION_ROOT=".Length)
    $canonicalRoot = Join-Path $sessionRoot "runtime"
    $stationConfigRoot = Join-Path $sessionRoot "station_s2_config"
    $stationInboxRoot = Join-Path $sessionRoot "station_s2_inbox"
    $operatorPath = Join-Path $sessionRoot "WH89_S2_Receiving_Operator.xlsb"
    $sourceWorkbookPath = Join-Path $sessionRoot "WH89_FRODECO.inventory_management.xlsb"
    $sharedConfigPath = Join-Path $canonicalRoot ($warehouseId + ".invSys.Config.xlsb")

    New-Item -ItemType Directory -Path $stationConfigRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $stationInboxRoot -Force | Out-Null

    Update-SharedConfig -ConfigPath $sharedConfigPath -AutoRefreshSeconds $AutoRefreshIntervalSeconds

    $setupOutput = & powershell -NoProfile -ExecutionPolicy Bypass -File $setupScript `
        -RepoRoot $repo `
        -WarehouseId $warehouseId `
        -StationId $stationB `
        -SharedRuntimeRoot $canonicalRoot `
        -StationInboxRoot $stationInboxRoot `
        -LocalConfigRoot $stationConfigRoot `
        -RoleDefault "RECEIVE" `
        -StationName "SOAK-STATION-S2" `
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

    $excelSource = New-ExcelApp
    $wbCoreSource = $excelSource.Workbooks.Open((Join-Path $deployPath "invSys.Core.xlam"))
    $wbInventoryDomainSource = $excelSource.Workbooks.Open((Join-Path $deployPath "invSys.Inventory.Domain.xlam"))
    [void](Run-WorkbookMacro -Excel $excelSource -WorkbookName $wbCoreSource.Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($canonicalRoot))
    [void](Run-WorkbookMacro -Excel $excelSource -WorkbookName $wbInventoryDomainSource.Name -MacroName "modInventoryInit.ResetSyncLog")
    [void](Run-WorkbookMacro -Excel $excelSource -WorkbookName $wbInventoryDomainSource.Name -MacroName "modInventoryInit.InitInventoryDomainAddin")
    $sourceLogPath = [string](Run-WorkbookMacro -Excel $excelSource -WorkbookName $wbInventoryDomainSource.Name -MacroName "modInventoryInit.GetSyncLogPath")

    $wbSource = $excelSource.Workbooks.Add()
    [void](Run-WorkbookMacro -Excel $excelSource -WorkbookName $wbCoreSource.Name -MacroName "modRoleWorkbookSurfaces.EnsureInventoryManagementSurface" -Arguments @($wbSource))
    $wbSource.SaveAs($sourceWorkbookPath, 50)

    $wsSource = Get-Worksheet -Workbook $wbSource -WorksheetName "InventoryManagement"
    $loSource = Get-ListObject -Worksheet $wsSource -TableName "invSys"
    if ($null -eq $loSource) { throw "Source workbook invSys table was not created." }
    [void](Ensure-ListObjectRow -ListObject $loSource)
    Set-TableValue -ListObject $loSource -RowIndex 1 -ColumnName "ROW" -Value 1001
    Set-TableValue -ListObject $loSource -RowIndex 1 -ColumnName "ITEM_CODE" -Value $sku
    Set-TableValue -ListObject $loSource -RowIndex 1 -ColumnName "ITEM" -Value "LAN Boundary Item"
    Set-TableValue -ListObject $loSource -RowIndex 1 -ColumnName "UOM" -Value "EA"
    Set-TableValue -ListObject $loSource -RowIndex 1 -ColumnName "LOCATION" -Value "A1"
    Set-TableValue -ListObject $loSource -RowIndex 1 -ColumnName "SnapshotId" -Value ($warehouseId + ".invSys.Snapshot.Inventory.xlsb|seed")
    $wbSource.Save()

    $excelOperator = New-ExcelApp
    $wbCoreOperator = $excelOperator.Workbooks.Open((Join-Path $deployPath "invSys.Core.xlam"))
    $wbInventoryDomainOperator = $excelOperator.Workbooks.Open((Join-Path $deployPath "invSys.Inventory.Domain.xlam"))
    $wbReceivingOperator = $excelOperator.Workbooks.Open((Join-Path $deployPath "invSys.Operations.xlam"))
    [void](Run-WorkbookMacro -Excel $excelOperator -WorkbookName $wbCoreOperator.Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($canonicalRoot))
    $wbLocalConfig = $excelOperator.Workbooks.Open((Join-Path $stationConfigRoot ($warehouseId + ".invSys.Config.xlsb")))
    $wbOperator = $excelOperator.Workbooks.Open($operatorPath)
    [void]$wbOperator.Activate()
    [void](Run-WorkbookMacro -Excel $excelOperator -WorkbookName $wbReceivingOperator.Name -MacroName "modReceivingInit.InitReceivingAddin")
    [void](Run-WorkbookMacro -Excel $excelOperator -WorkbookName $wbReceivingOperator.Name -MacroName "modReceivingInit.EnsureReceivingSurfaceForWorkbook" -Arguments @($wbOperator))

    $wsOperator = Get-Worksheet -Workbook $wbOperator -WorksheetName "InventoryManagement"
    $loOperator = Get-ListObject -Worksheet $wsOperator -TableName "invSys"
    if ($null -eq $loOperator) { throw "Operator workbook invSys table was not available." }

    $excelHarness = New-ExcelApp
    $harnessPath = Join-Path $sessionRoot "LiveSyncHarness.xlsm"
    $wbHarness = New-HarnessWorkbook -Excel $excelHarness -HarnessPath $harnessPath -ModulePaths $modulePaths
    $attach = [string](Run-WorkbookMacro -Excel $excelHarness -WorkbookName $wbHarness.Name -MacroName "TestPhase6LanBoundary.LanBoundaryAttachToCanonicalRoot" -Arguments @($canonicalRoot, $warehouseId, $stationA))
    if ($attach -notlike "OK*") {
        throw "Harness attach failed: $attach"
    }

    Start-Sleep -Seconds ([Math]::Max(6, $AutoRefreshIntervalSeconds + 3))

    $currentTotal = Convert-ExcelDoubleValue (Get-TableValue -ListObject $loSource -RowIndex 1 -ColumnName "TOTAL INV")
    $batchAll = New-Object 'System.Collections.Generic.List[double]'
    $catchupAll = New-Object 'System.Collections.Generic.List[double]'

    for ($iteration = 1; $iteration -le $Iterations; $iteration++) {
        foreach ($directionInfo in @(
            @{ Direction = "S2->S1"; StationId = $stationB; WaitListObject = $loSource; WaitLogPath = $sourceLogPath },
            @{ Direction = "S1->S2"; StationId = $stationA; WaitListObject = $loOperator; WaitLogPath = "" }
        )) {
            $qty = [double]$iteration
            $currentTotal += $qty

            $queueResult = [string](Run-WorkbookMacro -Excel $excelHarness -WorkbookName $wbHarness.Name -MacroName "TestPhase6LanBoundary.LanBoundaryQueueReceiveOnly" -Arguments @($warehouseId, $directionInfo.StationId, $sku, $qty, "A1", "bidirectional-soak-$iteration-$($directionInfo.Direction)"))
            if ($queueResult -notlike "OK*") {
                $rows.Add([pscustomobject]@{
                    Iteration        = $iteration
                    Direction        = $directionInfo.Direction
                    Qty              = $qty
                    ExpectedTotal    = $currentTotal
                    Processed        = 0
                    BatchMs          = 0
                    CatchupMs        = 0
                    RuntimeReadOnly  = ""
                    ObservedTotal    = Convert-ExcelDoubleValue (Get-TableValue -ListObject $directionInfo.WaitListObject -RowIndex 1 -ColumnName "TOTAL INV")
                    ObservedReceived = Convert-ExcelDoubleValue (Get-TableValue -ListObject $directionInfo.WaitListObject -RowIndex 1 -ColumnName "RECEIVED")
                    Result           = "FAIL"
                    Detail           = ($queueResult -replace '\|', '/')
                }) | Out-Null
                throw "Queue failed for $($directionInfo.Direction): $queueResult"
            }

            $batchSw = [System.Diagnostics.Stopwatch]::StartNew()
            $batchReport = [string](Run-WorkbookMacro -Excel $excelHarness -WorkbookName $wbHarness.Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @($warehouseId, 500))
            $batchSw.Stop()
            $batchMs = [math]::Round($batchSw.Elapsed.TotalMilliseconds, 0)
            $batchAll.Add($batchMs) | Out-Null

            $processedCount = 0
            if ($batchReport -match 'Processed=(\d+)') {
                $processedCount = [int]$Matches[1]
            }

            $waitResult = Wait-ForProjection -ListObject $directionInfo.WaitListObject `
                -ExpectedTotal $currentTotal `
                -ExpectedReceived $qty `
                -TimeoutSeconds $PollTimeoutSeconds `
                -PollIntervalMilliseconds $PollIntervalMilliseconds `
                -SyncLogPath $directionInfo.WaitLogPath

            if ($waitResult.Passed) {
                $catchupAll.Add([double]$waitResult.CatchupMs) | Out-Null
            }

            $detail = if ([string]::IsNullOrWhiteSpace($waitResult.Detail)) { $batchReport } else { $waitResult.Detail }
            $detail = ($detail -replace '\|', '/')
            if ($detail.Length -gt 220) { $detail = $detail.Substring(0, 220) }

            $rows.Add([pscustomobject]@{
                Iteration        = $iteration
                Direction        = $directionInfo.Direction
                Qty              = $qty
                ExpectedTotal    = $currentTotal
                Processed        = $processedCount
                BatchMs          = $batchMs
                CatchupMs        = $waitResult.CatchupMs
                RuntimeReadOnly  = $waitResult.RuntimeReadOnly
                ObservedTotal    = $waitResult.ObservedTotal
                ObservedReceived = $waitResult.ObservedReceived
                Result           = $(if ($processedCount -ge 1 -and $waitResult.Passed) { "PASS" } else { "FAIL" })
                Detail           = $detail
            }) | Out-Null

            if ($processedCount -lt 1 -or -not $waitResult.Passed) {
                throw "Catch-up failed for $($directionInfo.Direction). Results written to $resultPath"
            }
        }
    }

    $passedRows = @($rows | Where-Object { $_.Result -eq "PASS" })
    $failedRows = @($rows | Where-Object { $_.Result -ne "PASS" })
    $s2toS1Rows = @($rows | Where-Object { $_.Direction -eq "S2->S1" -and $_.Result -eq "PASS" })
    $s1toS2Rows = @($rows | Where-Object { $_.Direction -eq "S1->S2" -and $_.Result -eq "PASS" })
    $summary = @{
        IterationsRequested         = $Iterations
        LegsExecuted                = $rows.Count
        LegsPassed                  = $passedRows.Count
        LegsFailed                  = $failedRows.Count
        PollTimeoutSeconds          = $PollTimeoutSeconds
        PollIntervalMs              = $PollIntervalMilliseconds
        AutoRefreshIntervalSeconds  = $AutoRefreshIntervalSeconds
        AverageBatchMs              = $(if ($batchAll.Count -gt 0) { [math]::Round((($batchAll | Measure-Object -Average).Average), 0) } else { 0 })
        AverageCatchupMs            = $(if ($catchupAll.Count -gt 0) { [math]::Round((($catchupAll | Measure-Object -Average).Average), 0) } else { 0 })
        MaxCatchupMs                = $(if ($catchupAll.Count -gt 0) { [math]::Round((($catchupAll | Measure-Object -Maximum).Maximum), 0) } else { 0 })
        AverageCatchupMs_S2toS1     = $(if ($s2toS1Rows.Count -gt 0) { [math]::Round((($s2toS1Rows.CatchupMs | Measure-Object -Average).Average), 0) } else { 0 })
        AverageCatchupMs_S1toS2     = $(if ($s1toS2Rows.Count -gt 0) { [math]::Round((($s1toS2Rows.CatchupMs | Measure-Object -Average).Average), 0) } else { 0 })
        SessionRoot                 = $sessionRoot
        CanonicalRoot               = $canonicalRoot
        SourceWorkbook              = $sourceWorkbookPath
        OperatorWorkbook            = $operatorPath
        SyncLogPath                 = $sourceLogPath
    }

    Write-ResultFile -Path $resultPath -Summary $summary -Rows $rows
    if ($failedRows.Count -gt 0) {
        throw "Bidirectional live sync soak failed. Results written to $resultPath"
    }

    Write-Output "PHASE6_BIDIRECTIONAL_LIVE_SYNC_SOAK_OK"
    Write-Output "RESULTS=$resultPath"
}
finally {
    if ($null -ne $wbHarness) {
        try { $wbHarness.Close($false) } catch {}
        Release-ComObject $wbHarness
    }
    if ($null -ne $excelHarness) {
        try { $excelHarness.Quit() } catch {}
        Release-ComObject $excelHarness
    }
    if ($null -ne $wbOperator) {
        try { $wbOperator.Close($false) } catch {}
        Release-ComObject $wbOperator
    }
    if ($null -ne $wbLocalConfig) {
        try { $wbLocalConfig.Close($false) } catch {}
        Release-ComObject $wbLocalConfig
    }
    if ($null -ne $wbReceivingOperator) {
        try { $wbReceivingOperator.Close($false) } catch {}
        Release-ComObject $wbReceivingOperator
    }
    if ($null -ne $wbInventoryDomainOperator) {
        try { $wbInventoryDomainOperator.Close($false) } catch {}
        Release-ComObject $wbInventoryDomainOperator
    }
    if ($null -ne $wbCoreOperator) {
        try { $wbCoreOperator.Close($false) } catch {}
        Release-ComObject $wbCoreOperator
    }
    if ($null -ne $excelOperator) {
        try { $excelOperator.Quit() } catch {}
        Release-ComObject $excelOperator
    }
    if ($null -ne $wbSource) {
        try { $wbSource.Close($false) } catch {}
        Release-ComObject $wbSource
    }
    if ($null -ne $wbInventoryDomainSource) {
        try { $wbInventoryDomainSource.Close($false) } catch {}
        Release-ComObject $wbInventoryDomainSource
    }
    if ($null -ne $wbCoreSource) {
        try { $wbCoreSource.Close($false) } catch {}
        Release-ComObject $wbCoreSource
    }
    if ($null -ne $excelSource) {
        try { $excelSource.Quit() } catch {}
        Release-ComObject $excelSource
    }
}
