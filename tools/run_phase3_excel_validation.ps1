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

function Run-TestFunction {
    Param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$FunctionName
    )

    $fullMacro = "'$WorkbookName'!$FunctionName"
    $result = $Excel.Run($fullMacro)
    if ($null -eq $result) { return 0 }
    return [int]$result
}

function Add-BootstrapModule {
    Param([object]$Workbook)
    $comp = $Workbook.VBProject.VBComponents.Add(1)
    $comp.Name = "modHarnessBootstrap"
    $comp.CodeModule.AddFromString("Public Function HarnessPing() As Long: HarnessPing = 1: End Function")
    return $comp
}

function Add-TestWrappers {
    Param(
        [object]$BootstrapComponent,
        [string[]]$TargetFunctions
    )

    $cm = $BootstrapComponent.CodeModule
    $wrappers = @()
    for ($i = 0; $i -lt $TargetFunctions.Count; $i++) {
        $fn = $TargetFunctions[$i]
        $wrapper = "RunT" + ($i + 1)
        $line = @"
Public Function $wrapper() As Long
$wrapper = Application.Run("$fn")
End Function
"@
        $cm.AddFromString($line)
        $wrappers += $wrapper
    }
    return ,$wrappers
}

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$harnessPath = Join-Path $fixtures "Phase3_TestHarness.xlsm"
$resultPath = Join-Path $repo "tests/unit/phase3_test_results.md"

$excel = $null
$harness = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $modulePaths = @(
        (Join-Path $repo "src/Core/Modules/modConfigDefaults.bas"),
        (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas"),
        (Join-Path $repo "src/Core/Modules/modConfig.bas"),
        (Join-Path $repo "src/Core/Modules/modAuth.bas"),
        (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
        (Join-Path $repo "src/Core/Modules/modItemSearch.bas"),
        (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
        (Join-Path $repo "src/Core/Modules/modWarehouseSync.bas"),
        (Join-Path $repo "src/Core/Modules/modOperatorReadModel.bas"),
        (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleUiAccess.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
        (Join-Path $repo "src/Receiving/Modules/modReceivingEventCreator.bas"),
        (Join-Path $repo "src/Shipping/Modules/modShippingEventCreator.bas"),
        (Join-Path $repo "src/Production/Modules/modProductionEventCreator.bas"),
        (Join-Path $repo "tests/unit/TestPhase2Helpers.bas"),
        (Join-Path $repo "tests/unit/TestCoreItemSearch.bas"),
        (Join-Path $repo "tests/unit/TestCoreRoleEventWriter.bas"),
        (Join-Path $repo "tests/unit/TestCoreRoleUiAccess.bas"),
        (Join-Path $repo "tests/unit/TestPhase3RoleFlows.bas")
    )

    $allTests = @(
        "TestCoreRoleEventWriter.TestQueueReceiveEvent_WritesInboxRow",
        "TestCoreRoleEventWriter.TestOpenInboxWorkbook_UsesStationPathInboxRoot",
        "TestCoreRoleEventWriter.TestQueueShipEvent_WritesInboxRow",
        "TestCoreRoleEventWriter.TestQueuePayloadEvent_DeniedWithoutCapability",
        "TestCoreRoleEventWriter.TestBuildPayloadJson_WithObjectItems",
        "TestCoreRoleUiAccess.TestCanCurrentUserPerformCapability_Allow",
        "TestCoreRoleUiAccess.TestCanCurrentUserPerformCapability_Deny",
        "TestCoreRoleUiAccess.TestApplyShapeCapability_TogglesVisibility",
        "TestCoreItemSearch.TestNormalizeSearchText_CollapsesWhitespace",
        "TestCoreItemSearch.TestAnyTextMatchesSearch_MatchesAcrossFields",
        "TestCoreItemSearch.TestIdentifiersMatch_UsesTokenOverlap",
        "TestCoreItemSearch.TestResolveSearchCaption_ReturnsRoleSpecificText",
        "TestCoreItemSearch.TestShouldDefaultShippableForRole_UsesRoleDefaults",
        "TestPhase3RoleFlows.TestReceivingRoleFlow_QueuesAndProcessesEvent",
        "TestPhase3RoleFlows.TestShippingRoleFlow_QueuesAndProcessesEvent",
        "TestPhase3RoleFlows.TestProductionRoleFlow_QueuesAndProcessesEvent"
    )

    if (Test-Path $harnessPath) { Remove-Item $harnessPath -Force }
    $harness = $excel.Workbooks.Add()
    $bootstrap = Add-BootstrapModule -Workbook $harness
    $vbProject = $harness.VBProject
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")

    foreach ($m in $modulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
        [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    }

    $wrapperNames = Add-TestWrappers -BootstrapComponent $bootstrap -TargetFunctions $allTests
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    $harness.SaveAs($harnessPath, 52)

    $testRows = @()
    for ($i = 0; $i -lt $allTests.Count; $i++) {
        $name = $allTests[$i]
        $wrapperName = $wrapperNames[$i]
        $passed = Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName $wrapperName
        $testRows += [pscustomobject]@{
            TestName = $name
            Passed   = ($passed -eq 1)
        }
    }

    $passedCount = @($testRows | Where-Object { $_.Passed }).Count
    $failedCount = $testRows.Count - $passedCount

    $lines = @()
    $lines += "# Phase 3 VBA Test Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += ""
    $lines += "| Test | Result |"
    $lines += "|---|---|"
    foreach ($r in $testRows) {
        $lines += "| $($r.TestName) | $([string]::Join('', $(if ($r.Passed) {'PASS'} else {'FAIL'}))) |"
    }
    [System.IO.File]::WriteAllLines($resultPath, $lines)

    Write-Output "PHASE3_VALIDATION_OK"
    Write-Output "HARNESS=$harnessPath"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$passedCount FAILED=$failedCount TOTAL=$($testRows.Count)"
}
finally {
    if ($null -ne $harness) {
        try { $harness.Close($true) } catch {}
        Release-ComObject $harness
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}
