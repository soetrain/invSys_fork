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
        $wb.SaveAs($OutputPath, 50)  # xlExcel12 (.xlsb)
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
    $comp = $Workbook.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
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

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$cfgXlsx = Join-Path $fixtures "WH1.invSys.Config.sample.xlsx"
$authXlsx = Join-Path $fixtures "WH1.invSys.Auth.sample.xlsx"
$cfgXlsb = Join-Path $fixtures "WH1.invSys.Config.xlsb"
$authXlsb = Join-Path $fixtures "WH1.invSys.Auth.xlsb"
$harnessPath = Join-Path $fixtures "Phase1_TestHarness.xlsm"
$resultPath = Join-Path $repo "tests/unit/phase1_test_results.md"

$modulePaths = @(
    (Join-Path $repo "src/Core/Modules/modConfigDefaults.bas"),
    (Join-Path $repo "src/Core/Modules/modConfig.bas"),
    (Join-Path $repo "src/Core/Modules/modAuth.bas"),
    (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
    (Join-Path $repo "tests/unit/TestCoreConfig.bas"),
    (Join-Path $repo "tests/unit/TestCoreAuth.bas"),
    (Join-Path $repo "tests/unit/TestInventorySchema.bas")
)

$configTests = @(
    "TestCoreConfig.TestLoad_ValidConfig",
    "TestCoreConfig.TestLoad_MissingRequiredKey",
    "TestCoreConfig.TestPrecedence_StationOverridesWarehouse",
    "TestCoreConfig.TestGetRequired_MissingKey",
    "TestCoreConfig.TestGetBool_TypeConversion",
    "TestCoreConfig.TestReload_UpdatedValue"
)

$authTests = @(
    "TestCoreAuth.TestCanPerform_Allow",
    "TestCoreAuth.TestCanPerform_Deny_MissingCapability",
    "TestCoreAuth.TestCanPerform_WildcardStation",
    "TestCoreAuth.TestCanPerform_DisabledUser",
    "TestCoreAuth.TestCanPerform_ExpiredCapability",
    "TestCoreAuth.TestRequire_RaisesOnDeny"
)

$schemaTests = @(
    "TestInventorySchema.TestEnsureInventorySchema_RecreatesTables",
    "TestInventorySchema.TestEnsureInventorySchema_AddsMissingColumns"
)

$excel = $null
$harness = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    # 1) Convert sample fixtures to production-style xlsb names.
    Convert-SampleToXlsb -Excel $excel -InputPath $cfgXlsx -OutputPath $cfgXlsb
    Convert-SampleToXlsb -Excel $excel -InputPath $authXlsx -OutputPath $authXlsb

    # 2) Build test harness workbook and import modules.
    if (Test-Path $harnessPath) { Remove-Item $harnessPath -Force }
    $harness = $excel.Workbooks.Add()
    $bootstrap = Add-BootstrapModule -Workbook $harness
    $vbProject = $harness.VBProject

    # Warm-up run before imports.
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")

    foreach ($m in $modulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
        # Warm-up after each import; this avoids occasional COM macro resolution failures.
        [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    }

    $allTargetFunctions = @($configTests + $authTests + $schemaTests)
    Add-TestWrappers -BootstrapComponent $bootstrap -TargetFunctions $allTargetFunctions
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    $harness.SaveAs($harnessPath, 52)  # xlOpenXMLWorkbookMacroEnabled

    # 3) Execute tests and capture pass/fail.
    $testRows = @()
    foreach ($name in $configTests + $authTests + $schemaTests) {
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
    $lines += "# Phase 1 VBA Test Results"
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

    Write-Output "VALIDATION_OK"
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
