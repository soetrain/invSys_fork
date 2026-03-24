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

function Convert-SampleToXlsb {
    Param(
        [object]$Excel,
        [string]$InputPath,
        [string]$OutputPath
    )

    if (-not (Test-Path $InputPath)) {
        throw "Missing input workbook: $InputPath"
    }

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    $wb = $Excel.Workbooks.Open($InputPath)
    try {
        $wb.SaveAs($OutputPath, 50)
    }
    finally {
        $wb.Close($false)
        Release-ComObject $wb
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
    foreach ($fn in $TargetFunctions) {
        $wrapper = "Run_" + ($fn -replace "[^A-Za-z0-9_]", "_")
        $line = "Public Function $wrapper() As Long: $wrapper = $fn(): End Function"
        $cm.AddFromString($line)
    }
}

function Ensure-WorksheetEditable {
    Param(
        [object]$Worksheet,
        [string]$Context
    )

    if ($null -eq $Worksheet) {
        throw "Worksheet missing while preparing $Context"
    }

    if (-not $Worksheet.ProtectContents) {
        return
    }

    $Worksheet.Unprotect()
    if ($Worksheet.ProtectContents) {
        throw "Worksheet '$($Worksheet.Name)' is protected and could not be unprotected while preparing $Context"
    }
}

function Ensure-Phase2AuthData {
    Param(
        [object]$Excel,
        [string]$AuthPath
    )

    $wb = $Excel.Workbooks.Open($AuthPath)
    try {
        $wsUsers = $wb.Worksheets("Users")
        $wsCaps = $wb.Worksheets("Capabilities")
        try { $loUsers = $wsUsers.ListObjects("tblUsers") } catch { $loUsers = $null }
        try { $loCaps = $wsCaps.ListObjects("tblCapabilities") } catch { $loCaps = $null }

        if ($null -eq $loUsers) {
            $userRange = $wsUsers.Range("A1").CurrentRegion
            $loUsers = $wsUsers.ListObjects.Add(1, $userRange, $null, 1)
            $loUsers.Name = "tblUsers"
        }
        if ($null -eq $loCaps) {
            $capRange = $wsCaps.Range("A1").CurrentRegion
            $loCaps = $wsCaps.ListObjects.Add(1, $capRange, $null, 1)
            $loCaps.Name = "tblCapabilities"
        }

        $hasSvc = $false
        foreach ($cell in $loUsers.ListColumns("UserId").DataBodyRange.Cells) {
            if ($cell.Value2 -eq "svc_processor") { $hasSvc = $true; break }
        }
        if (-not $hasSvc) {
            Ensure-WorksheetEditable -Worksheet $wsUsers -Context "tblUsers"
            $row = $loUsers.ListRows.Add()
            $row.Range.Cells(1, $loUsers.ListColumns("UserId").Index).Value2 = "svc_processor"
            $row.Range.Cells(1, $loUsers.ListColumns("DisplayName").Index).Value2 = "Processor Service"
            $row.Range.Cells(1, $loUsers.ListColumns("Status").Index).Value2 = "Active"
        }

        $hasCap = $false
        foreach ($row in $loCaps.ListRows) {
            $uid = $row.Range.Cells(1, $loCaps.ListColumns("UserId").Index).Value2
            $cap = $row.Range.Cells(1, $loCaps.ListColumns("Capability").Index).Value2
            if ($uid -eq "svc_processor" -and $cap -eq "INBOX_PROCESS") {
                $hasCap = $true
                break
            }
        }
        if (-not $hasCap) {
            Ensure-WorksheetEditable -Worksheet $wsCaps -Context "tblCapabilities"
            $row = $loCaps.ListRows.Add()
            $row.Range.Cells(1, $loCaps.ListColumns("UserId").Index).Value2 = "svc_processor"
            $row.Range.Cells(1, $loCaps.ListColumns("Capability").Index).Value2 = "INBOX_PROCESS"
            $row.Range.Cells(1, $loCaps.ListColumns("WarehouseId").Index).Value2 = "WH1"
            $row.Range.Cells(1, $loCaps.ListColumns("StationId").Index).Value2 = "*"
            $row.Range.Cells(1, $loCaps.ListColumns("Status").Index).Value2 = "Active"
        }

        $wb.Save()
    }
    finally {
        $wb.Close($true)
        Release-ComObject $wb
    }
}

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$cfgXlsx = Join-Path $fixtures "WH1.invSys.Config.sample.xlsx"
$authXlsx = Join-Path $fixtures "WH1.invSys.Auth.sample.xlsx"
$cfgXlsb = Join-Path $fixtures "WH1.invSys.Config.xlsb"
$authXlsb = Join-Path $fixtures "WH1.invSys.Auth.xlsb"
$harnessRunStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$harnessPath = Join-Path $fixtures ("Phase2_TestHarness_" + $harnessRunStamp + ".xlsm")
$resultPath = Join-Path $repo "tests/unit/phase2_test_results.md"

& (Join-Path $repo "tools/create_phase1_fixture_xlsx.ps1") -OutputDir $fixtures | Out-Null

$excel = $null
$harness = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    Convert-SampleToXlsb -Excel $excel -InputPath $cfgXlsx -OutputPath $cfgXlsb
    Convert-SampleToXlsb -Excel $excel -InputPath $authXlsx -OutputPath $authXlsb
    Ensure-Phase2AuthData -Excel $excel -AuthPath $authXlsb

    & (Join-Path $repo "tools/create_phase2_fixture_xlsx.ps1") -OutputDir $fixtures | Out-Null

    $modulePaths = @(
        (Join-Path $repo "src/Core/Modules/modConfigDefaults.bas"),
        (Join-Path $repo "src/Core/Modules/modConfig.bas"),
        (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
        (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
        (Join-Path $repo "src/Core/Modules/modAuth.bas"),
        (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
        (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
        (Join-Path $repo "tests/unit/TestCoreConfig.bas"),
        (Join-Path $repo "tests/unit/TestCoreAuth.bas"),
        (Join-Path $repo "tests/unit/TestInventorySchema.bas"),
        (Join-Path $repo "tests/unit/TestPhase2Helpers.bas"),
        (Join-Path $repo "tests/unit/TestCoreLockManager.bas"),
        (Join-Path $repo "tests/unit/TestInventoryApply.bas"),
        (Join-Path $repo "tests/unit/TestCoreProcessor.bas")
    )

    $allTests = @(
        "TestCoreConfig.TestLoad_ValidConfig",
        "TestCoreConfig.TestLoad_MissingRequiredKey",
        "TestCoreConfig.TestPrecedence_StationOverridesWarehouse",
        "TestCoreConfig.TestGetRequired_MissingKey",
        "TestCoreConfig.TestGetBool_TypeConversion",
        "TestCoreConfig.TestReload_UpdatedValue",
        "TestCoreAuth.TestCanPerform_Allow",
        "TestCoreAuth.TestCanPerform_Deny_MissingCapability",
        "TestCoreAuth.TestCanPerform_WildcardStation",
        "TestCoreAuth.TestCanPerform_DisabledUser",
        "TestCoreAuth.TestCanPerform_ExpiredCapability",
        "TestCoreAuth.TestRequire_RaisesOnDeny",
        "TestInventorySchema.TestEnsureInventorySchema_RecreatesTables",
        "TestInventorySchema.TestEnsureInventorySchema_AddsMissingColumns",
        "TestInventorySchema.TestEnsureInventorySchema_RemovesBlankSeedRow",
        "TestCoreLockManager.TestAcquireReleaseLock_Lifecycle",
        "TestCoreLockManager.TestHeartbeat_ExtendsExpiry",
        "TestInventoryApply.TestApplyReceive_ValidEvent",
        "TestInventoryApply.TestApplyReceive_InvalidSKU",
        "TestInventoryApply.TestApplyReceive_Duplicate",
        "TestInventoryApply.TestApplyReceive_ProtectedSheetReturnsClearError",
        "TestInventoryApply.TestApplyShip_MultiLineEvent",
        "TestInventoryApply.TestApplyProdConsume_MultiLineEvent",
        "TestInventoryApply.TestApplyProdComplete_MultiLineEvent",
        "TestCoreProcessor.TestRunBatch_ProcessesInboxRow",
        "TestCoreProcessor.TestRunBatch_DuplicateMarkedSkipDup",
        "TestCoreProcessor.TestRunBatch_ProcessesShipRow",
        "TestCoreProcessor.TestRunBatch_ProcessesProdConsumeRow",
        "TestCoreProcessor.TestRunBatch_ProcessesProdCompleteRow"
    )

    $harness = $excel.Workbooks.Add()
    $bootstrap = Add-BootstrapModule -Workbook $harness
    $vbProject = $harness.VBProject
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")

    foreach ($m in $modulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
        [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    }

    Add-TestWrappers -BootstrapComponent $bootstrap -TargetFunctions $allTests
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    $harness.SaveAs($harnessPath, 52)

    $testRows = @()
    foreach ($name in $allTests) {
        $wrapperName = "Run_" + ($name -replace "[^A-Za-z0-9_]", "_")
        $passed = Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName $wrapperName
        $testRows += [pscustomobject]@{
            TestName = $name
            Passed   = ($passed -eq 1)
        }
    }

    $passedCount = ($testRows | Where-Object { $_.Passed }).Count
    $failedCount = $testRows.Count - $passedCount

    $lines = @()
    $lines += "# Phase 2 VBA Test Results"
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

    Write-Output "PHASE2_VALIDATION_OK"
    Write-Output "CONFIG_XLSB=$cfgXlsb"
    Write-Output "AUTH_XLSB=$authXlsb"
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
