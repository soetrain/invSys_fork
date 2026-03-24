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
    try {
        $result = $Excel.Run($fullMacro)
    }
    catch {
        throw "Excel.Run failed for $fullMacro :: $($_.Exception.Message)"
    }
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
        $errCell = "A" + ($i + 1)
        $line = @"
Public Function $wrapper() As Long
On Error GoTo ErrHandler
ThisWorkbook.Worksheets(1).Range("$errCell").Value = ""
$wrapper = Application.Run("$fn")
Exit Function
ErrHandler:
ThisWorkbook.Worksheets(1).Range("$errCell").Value = Err.Description
$wrapper = 0
End Function
"@
        $cm.AddFromString($line)
        $wrappers += $wrapper
    }
    return ,$wrappers
}

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$harnessPath = Join-Path $fixtures "Phase6_TestHarness.xlsm"
$resultPath = Join-Path $repo "tests/unit/phase6_test_results.md"

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
        (Join-Path $repo "src/Admin/Modules/modAdminConsole.bas"),
        (Join-Path $repo "tests/unit/TestPhase6CoreSurfaces.bas"),
        (Join-Path $repo "tests/unit/TestPhase6RoleSurfaces.bas")
    )

    $allTests = @(
        "TestPhase6CoreSurfaces.TestOpenOrCreateConfigWorkbookRuntime_CreatesCanonicalWorkbook",
        "TestPhase6CoreSurfaces.TestLoadConfig_AutoBootstrapsCanonicalWorkbook",
        "TestPhase6CoreSurfaces.TestLoadConfig_BlankContextAutoBootstrapsDefaultRuntimeWorkbook",
        "TestPhase6CoreSurfaces.TestLoadConfig_QuarantinesContaminatedConfigSheet",
        "TestPhase6CoreSurfaces.TestLoadAuth_AutoBootstrapsCanonicalWorkbook",
        "TestPhase6CoreSurfaces.TestLoadAuth_BootstrapGrantsCurrentOperatorCapabilities",
        "TestPhase6RoleSurfaces.TestEnsureReceivingWorkbookSurface_CreatesExpectedTables",
        "TestPhase6RoleSurfaces.TestEnsureReceivingWorkbookSurface_RecreatesDeletedArtifacts",
        "TestPhase6RoleSurfaces.TestEnsureShippingWorkbookSurface_CreatesExpectedTables",
        "TestPhase6RoleSurfaces.TestEnsureShippingWorkbookSurface_RecreatesDeletedArtifacts",
        "TestPhase6RoleSurfaces.TestEnsureProductionWorkbookSurface_CreatesExpectedTables",
        "TestPhase6RoleSurfaces.TestEnsureProductionWorkbookSurface_RecreatesDeletedArtifacts",
        "TestPhase6RoleSurfaces.TestEnsureAdminWorkbookSurface_CreatesExpectedTables"
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
        $errorText = [string]$harness.Worksheets.Item(1).Range("A$($i + 1)").Value2
        $testRows += [pscustomobject]@{
            TestName = $name
            Passed   = ($passed -eq 1)
            Error    = $errorText
        }
    }

    $passedCount = @($testRows | Where-Object { $_.Passed }).Count
    $failedCount = $testRows.Count - $passedCount

    $lines = @()
    $lines += "# Phase 6 VBA Test Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += ""
    $lines += "| Test | Result |"
    $lines += "|---|---|"
    foreach ($r in $testRows) {
        $detail = if ($r.Passed) { "PASS" } elseif ([string]::IsNullOrWhiteSpace($r.Error)) { "FAIL" } else { "FAIL - $($r.Error)" }
        $lines += "| $($r.TestName) | $detail |"
    }
    [System.IO.File]::WriteAllLines($resultPath, $lines)

    Write-Output "PHASE6_VALIDATION_OK"
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
