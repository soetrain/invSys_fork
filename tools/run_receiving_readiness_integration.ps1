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

function New-NormalizedImportFile {
    Param([string]$SourcePath)

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-harness-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $tempPath = Join-Path $tempDir ([System.IO.Path]::GetFileName($SourcePath))
    $raw = Get-Content -LiteralPath $SourcePath -Raw
    $normalized = $raw -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($tempPath, $normalized, [System.Text.Encoding]::ASCII)
    return $tempPath
}

function Remove-ExistingVBComponent {
    Param(
        [object]$VbProject,
        [string]$ComponentName
    )

    try {
        $existing = $VbProject.VBComponents.Item($ComponentName)
    }
    catch {
        $existing = $null
    }

    if ($null -eq $existing) { return }
    if ($existing.Type -eq 100) {
        throw "Refusing to remove document component '$ComponentName'."
    }
    [void]$VbProject.VBComponents.Remove($existing)
}

function Import-ClassModule {
    Param(
        [object]$VbProject,
        [string]$ClassPath
    )

    if (-not (Test-Path $ClassPath)) {
        throw "Missing class module: $ClassPath"
    }

    $componentName = [System.IO.Path]::GetFileNameWithoutExtension($ClassPath)
    Remove-ExistingVBComponent -VbProject $VbProject -ComponentName $componentName
    $normalizedPath = New-NormalizedImportFile -SourcePath $ClassPath
    try {
        [void]$VbProject.VBComponents.Import($normalizedPath)
    }
    finally {
        Remove-Item -LiteralPath (Split-Path $normalizedPath -Parent) -Recurse -Force -ErrorAction SilentlyContinue
    }

    $component = $VbProject.VBComponents.Item($componentName)
    if ($component.Type -ne 2) {
        throw "Class import failed for $ClassPath; component '$componentName' resolved as type $($component.Type)."
    }
}

function Run-Macro {
    Param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$FunctionName
    )

    $fullMacro = "'$WorkbookName'!$FunctionName"
    return $Excel.Run($fullMacro)
}

function Parse-PackedMap {
    Param([string]$Packed)

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Packed)) { return $map }

    foreach ($part in ($Packed -split '\|')) {
        if ($part -match '^(.*?)=(.*)$') {
            $map[$matches[1]] = $matches[2]
        }
    }
    return $map
}

function Sanitize-MarkdownCell {
    Param([string]$Text)

    if ($null -eq $Text) { return "" }
    return (($Text -replace '\|', ' ; ') -replace "`r|`n", ' ').Trim()
}

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$harnessStamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$harnessPath = Join-Path $fixtures "ReceivingReadiness_Integration_Harness_$harnessStamp.xlsm"
$resultPath = Join-Path $repo "tests/integration/receiving-readiness-results.md"

$excel = $null
$harness = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $modulePaths = @(
        (Join-Path $repo "src/Core/Modules/modConfigDefaults.bas"),
        (Join-Path $repo "src/Core/Modules/modDeploymentPaths.bas"),
        (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas"),
        (Join-Path $repo "src/Core/Modules/modOperatorReadModel.bas"),
        (Join-Path $repo "src/Core/Modules/modPerfLog.bas"),
        (Join-Path $repo "src/Core/Modules/modDiagnostics.bas"),
        (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
        (Join-Path $repo "src/Core/Modules/modWarehouseSync.bas"),
        (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
        (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
        (Join-Path $repo "src/Core/Modules/modConfig.bas"),
        (Join-Path $repo "src/Core/Modules/modAuth.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryPublisher.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryBridgeApi.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
        (Join-Path $repo "src/Receiving/Modules/modReceivingInit.bas"),
        (Join-Path $repo "tests/unit/TestPhase2Helpers.bas"),
        (Join-Path $repo "tests/unit/TestStub_modTS_Received.bas"),
        (Join-Path $repo "tests/integration/test_ReceivingReadiness.bas"),
        (Join-Path $repo "tests/integration/TestReceivingReadinessEntry.bas")
    )
    $classPaths = @(
        (Join-Path $repo "src/Receiving/ClassModules/cAppEvents.cls")
    )

    $harness = $excel.Workbooks.Add()
    $vbProject = $harness.VBProject
    foreach ($m in $modulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
    }
    foreach ($c in $classPaths) {
        Import-ClassModule -VbProject $vbProject -ClassPath $c
    }

    $harness.SaveAs($harnessPath, 52)

    $passed = [int](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "TestReceivingReadinessEntry.RunReceivingReadinessIntegration")
    $context = [string](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "TestReceivingReadinessEntry.GetReceivingReadinessIntegrationContext")
    $rowsPacked = [string](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "TestReceivingReadinessEntry.GetReceivingReadinessIntegrationRows")

    $contextMap = Parse-PackedMap -Packed $context
    $rows = @()
    if (-not [string]::IsNullOrWhiteSpace($rowsPacked)) {
        foreach ($line in ($rowsPacked -split "`r?`n")) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $parts = $line -split "`t", 3
            $rows += [pscustomobject]@{
                Check  = if ($parts.Count -ge 1) { $parts[0] } else { "" }
                Result = if ($parts.Count -ge 2) { $parts[1] } else { "FAIL" }
                Detail = if ($parts.Count -ge 3) { $parts[2] } else { "" }
            }
        }
    }

    $passCount = @($rows | Where-Object { $_.Result -eq "PASS" }).Count
    $failCount = @($rows | Where-Object { $_.Result -ne "PASS" }).Count
    $overall = if ($passed -eq 1) { "PASS" } else { "FAIL" }

    $lines = @()
    $lines += "# Receiving Readiness Integration Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Overall: $overall"
    $lines += "- Harness: $harnessPath"
    if ($contextMap.ContainsKey("Summary")) { $lines += "- Summary: $($contextMap['Summary'])" }
    $lines += "- Passed checks: $passCount"
    $lines += "- Failed checks: $failCount"
    $lines += ""
    $lines += "| Check | Result | Detail |"
    $lines += "|---|---|---|"
    foreach ($row in $rows) {
        $lines += "| $(Sanitize-MarkdownCell $row.Check) | $(Sanitize-MarkdownCell $row.Result) | $(Sanitize-MarkdownCell $row.Detail) |"
    }
    [System.IO.File]::WriteAllLines($resultPath, $lines)

    Write-Output "RECEIVING_READINESS_INTEGRATION_OK"
    Write-Output "HARNESS=$harnessPath"
    Write-Output "RESULTS=$resultPath"
    Write-Output "OVERALL=$overall PASSED=$passCount FAILED=$failCount TOTAL=$($rows.Count)"
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
