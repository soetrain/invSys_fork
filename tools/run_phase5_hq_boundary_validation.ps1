Param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject {
    Param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Import-BasModule {
    Param(
        [object]$VbProject,
        [string]$BasPath
    )

    if (-not (Test-Path $BasPath)) {
        throw "Missing BAS module: $BasPath"
    }
    [void]$VbProject.VBComponents.Import($BasPath)
}

function Run-WorkbookMacro {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName,
        [object[]]$Arguments = @()
    )

    $fullMacro = "'$WorkbookName'!$MacroName"
    switch ($Arguments.Count) {
        0 { return $Excel.Run($fullMacro) }
        1 { return $Excel.Run($fullMacro, $Arguments[0]) }
        2 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1]) }
        3 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2]) }
        4 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3]) }
        5 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4]) }
        6 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5]) }
        7 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6]) }
        8 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7]) }
        default { throw "Run-WorkbookMacro supports at most 8 arguments." }
    }
}

function Add-BootstrapModule {
    Param([object]$Workbook)
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
    $bootstrap = Add-BootstrapModule -Workbook $wb
    $vbProject = $wb.VBProject
    [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $wb.Name -MacroName "HarnessPing")

    foreach ($m in $ModulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
        [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $wb.Name -MacroName "HarnessPing")
    }

    $wb.SaveAs($HarnessPath, 52)
    return $wb
}

function New-ExcelApp {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    return $excel
}

function Add-ResultRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Check,
        [bool]$Passed,
        [string]$Detail
    )

    $Rows.Add([pscustomobject]@{
        Check  = $Check
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

function Write-Results {
    param(
        [string]$ResultPath,
        [System.Collections.Generic.List[object]]$Rows
    )

    $passedCount = @($Rows | Where-Object { $_.Passed }).Count
    $failedCount = $Rows.Count - $passedCount

    $lines = @()
    $lines += "# Phase 5 HQ Boundary Validation Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += ""
    $lines += "| Check | Result |"
    $lines += "|---|---|"
    foreach ($row in $Rows) {
        $detail = if ($row.Passed) { "PASS" } else { "FAIL" }
        if (-not [string]::IsNullOrWhiteSpace($row.Detail)) {
            $detail = "$detail - $($row.Detail)"
        }
        $lines += "| $($row.Check) | $detail |"
    }
    [System.IO.File]::WriteAllLines($ResultPath, $lines)

    return [pscustomobject]@{
        Passed = $passedCount
        Failed = $failedCount
        Total  = $Rows.Count
    }
}

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$sessionRoot = Join-Path $fixtures "phase5_hq_boundary_$stamp"
$shareRoot = Join-Path $sessionRoot "share"
$rootA = Join-Path $sessionRoot "wh97"
$rootB = Join-Path $sessionRoot "wh98"
$resultPath = Join-Path $repo "tests/unit/phase5_hq_boundary_results.md"

$warehouseA = "WH97"
$warehouseB = "WH98"
$stationA = "S1"
$stationB = "S2"
$sku = "SKU-HQ-BOUNDARY-001"

$modulePaths = @(
    (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
    (Join-Path $repo "src/Core/Modules/modConfigDefaults.bas"),
    (Join-Path $repo "src/Core/Modules/modConfig.bas"),
    (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
    (Join-Path $repo "src/Core/Modules/modAuth.bas"),
    (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
    (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas"),
    (Join-Path $repo "src/Core/Modules/modWarehouseSync.bas"),
    (Join-Path $repo "src/Core/Modules/modHqAggregator.bas"),
    (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventoryBridgeApi.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
    (Join-Path $repo "tests/unit/TestPhase2Helpers.bas"),
    (Join-Path $repo "tests/unit/TestPhase5HqBoundary.bas")
)

$excelSetup = $null
$excelA = $null
$excelB = $null
$excelHq = $null
$harnessSetup = $null
$harnessA = $null
$harnessB = $null
$harnessHq = $null
$rows = New-Object 'System.Collections.Generic.List[object]'

try {
    if (Test-Path $sessionRoot) { Remove-Item $sessionRoot -Recurse -Force }
    New-Item -ItemType Directory -Path $sessionRoot | Out-Null

    $excelSetup = New-ExcelApp
    $setupPath = Join-Path $sessionRoot "Phase5HqBoundary_Inventory.Domain_Setup.xlsm"
    $harnessSetup = New-HarnessWorkbook -Excel $excelSetup -HarnessPath $setupPath -ModulePaths $modulePaths

    $seedA = [string](Run-WorkbookMacro -Excel $excelSetup -WorkbookName $harnessSetup.Name -MacroName "TestPhase5HqBoundary.HqBoundarySeedWarehouseRoot" -Arguments @($rootA, $shareRoot, $warehouseA, $stationA, $sku))
    $seedB = [string](Run-WorkbookMacro -Excel $excelSetup -WorkbookName $harnessSetup.Name -MacroName "TestPhase5HqBoundary.HqBoundarySeedWarehouseRoot" -Arguments @($rootB, $shareRoot, $warehouseB, $stationB, $sku))
    Add-ResultRow -Rows $rows -Check "Setup.WH97" -Passed ($seedA -like "OK*") -Detail $seedA
    Add-ResultRow -Rows $rows -Check "Setup.WH98" -Passed ($seedB -like "OK*") -Detail $seedB

    $harnessSetup.Close($false)
    Release-ComObject $harnessSetup
    $harnessSetup = $null
    $excelSetup.Quit()
    Release-ComObject $excelSetup
    $excelSetup = $null

    $excelA = New-ExcelApp
    $excelB = New-ExcelApp
    $excelHq = New-ExcelApp
    $pathA = Join-Path $sessionRoot "Phase5HqBoundary_Inventory.Domain_A.xlsm"
    $pathB = Join-Path $sessionRoot "Phase5HqBoundary_Inventory.Domain_B.xlsm"
    $pathHq = Join-Path $sessionRoot "Phase5HqBoundary_Inventory.Domain_HQ.xlsm"
    $harnessA = New-HarnessWorkbook -Excel $excelA -HarnessPath $pathA -ModulePaths $modulePaths
    $harnessB = New-HarnessWorkbook -Excel $excelB -HarnessPath $pathB -ModulePaths $modulePaths
    $harnessHq = New-HarnessWorkbook -Excel $excelHq -HarnessPath $pathHq -ModulePaths $modulePaths

    $publishA1 = [string](Run-WorkbookMacro -Excel $excelA -WorkbookName $harnessA.Name -MacroName "TestPhase5HqBoundary.HqBoundaryWarehouseRunAndPublish" -Arguments @($rootA, $shareRoot, $warehouseA, $stationA, $sku, [double]5, "A1", "wh97-initial"))
    $publishB1 = [string](Run-WorkbookMacro -Excel $excelB -WorkbookName $harnessB.Name -MacroName "TestPhase5HqBoundary.HqBoundaryWarehouseRunAndPublish" -Arguments @($rootB, $shareRoot, $warehouseB, $stationB, $sku, [double]8, "B1", "wh98-initial"))
    Add-ResultRow -Rows $rows -Check "Publish.WH97.Initial" -Passed (($publishA1 -like "OK*") -and ($publishA1 -match "Processed=1")) -Detail $publishA1
    Add-ResultRow -Rows $rows -Check "Publish.WH98.Initial" -Passed (($publishB1 -like "OK*") -and ($publishB1 -match "Processed=1")) -Detail $publishB1

    $aggregate1 = [string](Run-WorkbookMacro -Excel $excelHq -WorkbookName $harnessHq.Name -MacroName "TestPhase5HqBoundary.HqBoundaryRunAggregatorAndRead" -Arguments @($shareRoot, $warehouseA, $warehouseB, $sku))
    $aggregate1Pass = ($aggregate1 -like "OK*") -and ($aggregate1 -match "QtyA=5") -and ($aggregate1 -match "QtyB=8") -and ($aggregate1 -match "Skipped=0") -and ($aggregate1 -match "Warehouses=2")
    Add-ResultRow -Rows $rows -Check "Aggregate.Initial" -Passed $aggregate1Pass -Detail $aggregate1

    $publishB2 = [string](Run-WorkbookMacro -Excel $excelB -WorkbookName $harnessB.Name -MacroName "TestPhase5HqBoundary.HqBoundaryWarehouseRunAndPublish" -Arguments @($rootB, $shareRoot, $warehouseB, $stationB, $sku, [double]3, "B1", "wh98-catchup"))
    Add-ResultRow -Rows $rows -Check "Publish.WH98.Catchup" -Passed (($publishB2 -like "OK*") -and ($publishB2 -match "Processed=1")) -Detail $publishB2

    $aggregate2 = [string](Run-WorkbookMacro -Excel $excelHq -WorkbookName $harnessHq.Name -MacroName "TestPhase5HqBoundary.HqBoundaryRunAggregatorAndRead" -Arguments @($shareRoot, $warehouseA, $warehouseB, $sku))
    $aggregate2Pass = ($aggregate2 -like "OK*") -and ($aggregate2 -match "QtyA=5") -and ($aggregate2 -match "QtyB=11") -and ($aggregate2 -match "Skipped=0") -and ($aggregate2 -match "Warehouses=2")
    Add-ResultRow -Rows $rows -Check "Aggregate.Catchup" -Passed $aggregate2Pass -Detail $aggregate2

    $summary = Write-Results -ResultPath $resultPath -Rows $rows

    Write-Output "PHASE5_HQ_BOUNDARY_OK"
    Write-Output "RESULTS=$resultPath"
    Write-Output "SESSION_ROOT=$sessionRoot"
    Write-Output "PASSED=$($summary.Passed) FAILED=$($summary.Failed) TOTAL=$($summary.Total)"
}
finally {
    if ($null -ne $harnessA) {
        try { $harnessA.Close($false) } catch {}
        Release-ComObject $harnessA
    }
    if ($null -ne $harnessB) {
        try { $harnessB.Close($false) } catch {}
        Release-ComObject $harnessB
    }
    if ($null -ne $harnessHq) {
        try { $harnessHq.Close($false) } catch {}
        Release-ComObject $harnessHq
    }
    if ($null -ne $excelA) {
        try { $excelA.Quit() } catch {}
        Release-ComObject $excelA
    }
    if ($null -ne $excelB) {
        try { $excelB.Quit() } catch {}
        Release-ComObject $excelB
    }
    if ($null -ne $excelHq) {
        try { $excelHq.Quit() } catch {}
        Release-ComObject $excelHq
    }
    if ($null -ne $harnessSetup) {
        try { $harnessSetup.Close($false) } catch {}
        Release-ComObject $harnessSetup
    }
    if ($null -ne $excelSetup) {
        try { $excelSetup.Quit() } catch {}
        Release-ComObject $excelSetup
    }
}
