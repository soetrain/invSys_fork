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
        default { throw "Run-WorkbookMacro supports at most 6 arguments." }
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
    $lines += "# Phase 6 LAN Boundary Validation Results"
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

function Write-ProgressLog {
    param(
        [string]$LogPath,
        [string]$Message
    )

    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $LogPath -Value "[$stamp] $Message"
}

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$sessionRoot = Join-Path $fixtures "phase6_lan_boundary_$stamp"
$canonicalRoot = Join-Path $sessionRoot "runtime"
$publishedRoot = Join-Path $sessionRoot "published"
$operatorPathB = Join-Path $sessionRoot "stationB_operator.xlsb"
$resultPath = Join-Path $repo "tests/unit/phase6_lan_boundary_results.md"
$progressPath = Join-Path $sessionRoot "lan_boundary_progress.log"

$warehouseId = "WH89"
$stationA = "S1"
$stationB = "S2"
$sku = "SKU-LAN-BOUNDARY-001"

$modulePaths = @(
    (Join-Path $repo "src/Core/Modules/modConfigDefaults.bas"),
    (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
    (Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas"),
    (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas"),
    (Join-Path $repo "src/Core/Modules/modOperatorReadModel.bas"),
    (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
    (Join-Path $repo "src/Core/Modules/modWarehouseSync.bas"),
    (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
    (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
    (Join-Path $repo "src/Core/Modules/modConfig.bas"),
    (Join-Path $repo "src/Core/Modules/modAuth.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventoryBridgeApi.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
    (Join-Path $repo "tests/unit/TestPhase2Helpers.bas"),
    (Join-Path $repo "tests/unit/TestPhase6LanBoundary.bas")
)

$excelSetup = $null
$excelA = $null
$excelB = $null
$harnessSetup = $null
$harnessA = $null
$harnessB = $null
$rows = New-Object 'System.Collections.Generic.List[object]'

try {
    if (Test-Path $sessionRoot) { Remove-Item $sessionRoot -Recurse -Force }
    New-Item -ItemType Directory -Path $sessionRoot | Out-Null
    New-Item -ItemType File -Path $progressPath -Force | Out-Null

    Write-ProgressLog -LogPath $progressPath -Message "Creating setup harness."
    $excelSetup = New-ExcelApp
    $setupPath = Join-Path $sessionRoot "LanBoundary_Inventory.Domain_Setup.xlsm"
    $harnessSetup = New-HarnessWorkbook -Excel $excelSetup -HarnessPath $setupPath -ModulePaths $modulePaths

    Write-ProgressLog -LogPath $progressPath -Message "Running seed macro."
    $seed = [string](Run-WorkbookMacro -Excel $excelSetup -WorkbookName $harnessSetup.Name -MacroName "TestPhase6LanBoundary.LanBoundarySeedCanonicalRoot" -Arguments @($canonicalRoot, $publishedRoot, $warehouseId, $stationA, $stationB, $sku))
    Add-ResultRow -Rows $rows -Check "Setup.SharedRoot" -Passed ($seed -like "OK*") -Detail $seed
    Write-ProgressLog -LogPath $progressPath -Message "Seed result: $seed"

    $harnessSetup.Close($false)
    Release-ComObject $harnessSetup
    $harnessSetup = $null
    $excelSetup.Quit()
    Release-ComObject $excelSetup
    $excelSetup = $null

    Write-ProgressLog -LogPath $progressPath -Message "Creating session A/B harnesses."
    $excelA = New-ExcelApp
    $excelB = New-ExcelApp
    $pathA = Join-Path $sessionRoot "LanBoundary_Inventory.Domain_A.xlsm"
    $pathB = Join-Path $sessionRoot "LanBoundary_Inventory.Domain_B.xlsm"
    $harnessA = New-HarnessWorkbook -Excel $excelA -HarnessPath $pathA -ModulePaths $modulePaths
    $harnessB = New-HarnessWorkbook -Excel $excelB -HarnessPath $pathB -ModulePaths $modulePaths

    Write-ProgressLog -LogPath $progressPath -Message "Attaching session A."
    $attachA = [string](Run-WorkbookMacro -Excel $excelA -WorkbookName $harnessA.Name -MacroName "TestPhase6LanBoundary.LanBoundaryAttachToCanonicalRoot" -Arguments @($canonicalRoot, $warehouseId, $stationA))
    Write-ProgressLog -LogPath $progressPath -Message "Attach A result: $attachA"
    Write-ProgressLog -LogPath $progressPath -Message "Attaching session B."
    $attachB = [string](Run-WorkbookMacro -Excel $excelB -WorkbookName $harnessB.Name -MacroName "TestPhase6LanBoundary.LanBoundaryAttachToCanonicalRoot" -Arguments @($canonicalRoot, $warehouseId, $stationB))
    Write-ProgressLog -LogPath $progressPath -Message "Attach B result: $attachB"
    Add-ResultRow -Rows $rows -Check "Attach.SessionA" -Passed ($attachA -like "OK*") -Detail $attachA
    Add-ResultRow -Rows $rows -Check "Attach.SessionB" -Passed ($attachB -like "OK*") -Detail $attachB

    Write-ProgressLog -LogPath $progressPath -Message "Holding canonical inventory in session A."
    $hold = [string](Run-WorkbookMacro -Excel $excelA -WorkbookName $harnessA.Name -MacroName "TestPhase6LanBoundary.LanBoundaryHoldCanonicalInventory" -Arguments @($warehouseId))
    Write-ProgressLog -LogPath $progressPath -Message "Hold result: $hold"
    $holdPass = ($hold -like "OK*") -and ($hold -match "ReadOnly=False")
    Add-ResultRow -Rows $rows -Check "Lock.SessionAHold" -Passed $holdPass -Detail $hold

    Write-ProgressLog -LogPath $progressPath -Message "Running lock-denied receive in session B."
    $lockFail = [string](Run-WorkbookMacro -Excel $excelB -WorkbookName $harnessB.Name -MacroName "TestPhase6LanBoundary.LanBoundaryQueueAndRunReceive" -Arguments @($warehouseId, $stationB, $sku, [double]5, "A1", "station-b-lock-check"))
    Write-ProgressLog -LogPath $progressPath -Message "Lock-denied result: $lockFail"
    $lockFailPass = ($lockFail -like "OK*") -and ($lockFail -match "Processed=0") -and ($lockFail -match "read-only or locked by another Excel session")
    Add-ResultRow -Rows $rows -Check "Lock.SessionBDeniedByFileBoundary" -Passed $lockFailPass -Detail $lockFail

    Write-ProgressLog -LogPath $progressPath -Message "Closing canonical inventory in session A."
    $closeHold = [string](Run-WorkbookMacro -Excel $excelA -WorkbookName $harnessA.Name -MacroName "TestPhase6LanBoundary.LanBoundaryCloseCanonicalInventory" -Arguments @($warehouseId))
    Write-ProgressLog -LogPath $progressPath -Message "Close hold result: $closeHold"
    Add-ResultRow -Rows $rows -Check "Lock.SessionARelease" -Passed ($closeHold -like "OK*") -Detail $closeHold

    Write-ProgressLog -LogPath $progressPath -Message "Queueing receive in session A."
    $queueA = [string](Run-WorkbookMacro -Excel $excelA -WorkbookName $harnessA.Name -MacroName "TestPhase6LanBoundary.LanBoundaryQueueReceiveOnly" -Arguments @($warehouseId, $stationA, $sku, [double]7, "A1", "station-a-publish"))
    Write-ProgressLog -LogPath $progressPath -Message "Queue A result: $queueA"
    $eventIdA = if ($queueA -match 'EventID=([^|]+)') { $matches[1] } else { "" }
    Write-ProgressLog -LogPath $progressPath -Message "Running batch for queued session A event."
    $processA = [string](Run-WorkbookMacro -Excel $excelA -WorkbookName $harnessA.Name -MacroName "TestPhase6LanBoundary.LanBoundaryRunBatchForEvent" -Arguments @($warehouseId, $stationA, $eventIdA))
    Write-ProgressLog -LogPath $progressPath -Message "Process A result: $processA"
    $processAPass = ($processA -like "OK*") -and ($processA -match "Processed=1") -and ($processA -match "Status=PROCESSED")
    Add-ResultRow -Rows $rows -Check "Lock.SessionARetryAfterRelease" -Passed $processAPass -Detail $processA

    Write-ProgressLog -LogPath $progressPath -Message "Publishing snapshot from session A."
    $publish = [string](Run-WorkbookMacro -Excel $excelA -WorkbookName $harnessA.Name -MacroName "TestPhase6LanBoundary.LanBoundaryPublishCurrentSnapshot" -Arguments @($warehouseId, $publishedRoot))
    Write-ProgressLog -LogPath $progressPath -Message "Publish result: $publish"
    $publishedSnapshotPath = Join-Path $publishedRoot "$warehouseId.invSys.Snapshot.Inventory.xlsb"
    $publishPass = ($publish -like "OK*") -and (Test-Path $publishedSnapshotPath)
    Add-ResultRow -Rows $rows -Check "Publish.SessionAToSharedSnapshot" -Passed $publishPass -Detail $publish

    Write-ProgressLog -LogPath $progressPath -Message "Building saved operator in session B."
    $buildOpB = [string](Run-WorkbookMacro -Excel $excelB -WorkbookName $harnessB.Name -MacroName "TestPhase6LanBoundary.LanBoundaryBuildSavedReceivingOperator" -Arguments @($operatorPathB, $sku, "REF-LAN-B-001", "SNAP-LAN-B-OLD", [double]0, "Z1"))
    Write-ProgressLog -LogPath $progressPath -Message "Build operator result: $buildOpB"
    Add-ResultRow -Rows $rows -Check "Operator.BuildStationB" -Passed ($buildOpB -like "OK*") -Detail $buildOpB

    Write-ProgressLog -LogPath $progressPath -Message "Refreshing station B from published snapshot."
    $refreshB = [string](Run-WorkbookMacro -Excel $excelB -WorkbookName $harnessB.Name -MacroName "TestPhase6LanBoundary.LanBoundaryRefreshSavedOperatorFromRoot" -Arguments @($operatorPathB, $warehouseId, $publishedRoot, "SHAREPOINT"))
    Write-ProgressLog -LogPath $progressPath -Message "Refresh B result: $refreshB"
    $refreshBPass = ($refreshB -like "OK*") -and ($refreshB -match "TotalInv=7") -and ($refreshB -match "SourceType=SHAREPOINT") -and ($refreshB -match "IsStale=False")
    Add-ResultRow -Rows $rows -Check "Refresh.SessionBReadsPublishedSnapshot" -Passed $refreshBPass -Detail $refreshB

    $summary = Write-Results -ResultPath $resultPath -Rows $rows
    Write-ProgressLog -LogPath $progressPath -Message "Summary: Passed=$($summary.Passed) Failed=$($summary.Failed) Total=$($summary.Total)"

    Write-Output "PHASE6_LAN_BOUNDARY_OK"
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
    if ($null -ne $excelA) {
        try { $excelA.Quit() } catch {}
        Release-ComObject $excelA
    }
    if ($null -ne $excelB) {
        try { $excelB.Quit() } catch {}
        Release-ComObject $excelB
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
