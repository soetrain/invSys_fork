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

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-confirm-writes-" + [guid]::NewGuid().ToString("N"))
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

function Start-ExcelEnterDismissal {
    Param([int]$Seconds = 20)

    Start-Job -ScriptBlock {
        param($durationSeconds)
        $shell = $null
        try {
            $shell = New-Object -ComObject WScript.Shell
            $stopAt = (Get-Date).AddSeconds($durationSeconds)
            while ((Get-Date) -lt $stopAt) {
                Start-Sleep -Milliseconds 400
                try { [void]$shell.AppActivate("Microsoft Excel") } catch {}
                try { $shell.SendKeys("~") } catch {}
            }
        }
        finally {
            if ($null -ne $shell) {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch {}
            }
        }
    } -ArgumentList $Seconds
}

function Invoke-MacroWithDismiss {
    Param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$FunctionName,
        [int]$DismissSeconds = 20
    )

    $job = Start-ExcelEnterDismissal -Seconds $DismissSeconds
    try {
        return Run-Macro -Excel $Excel -WorkbookName $WorkbookName -FunctionName $FunctionName
    }
    finally {
        if ($null -ne $job) {
            try { Wait-Job -Job $job -Timeout ($DismissSeconds + 2) | Out-Null } catch {}
            try { Receive-Job -Job $job -ErrorAction SilentlyContinue | Out-Null } catch {}
            try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
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
$harnessPath = Join-Path $fixtures "ConfirmWritesTester_Integration_Harness_$harnessStamp.xlsm"
$resultPath = Join-Path $repo "tests/integration/confirm-writes-tester-results.md"

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
        (Join-Path $repo "src/Core/Modules/modWarehouseBootstrap.bas"),
        (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleUiAccess.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas"),
        (Join-Path $repo "src/Core/Modules/modOperatorReadModel.bas"),
        (Join-Path $repo "src/Core/Modules/modPerfLog.bas"),
        (Join-Path $repo "src/Core/Modules/modDiagnostics.bas"),
        (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
        (Join-Path $repo "src/Core/Modules/modWarehouseSync.bas"),
        (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
        (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
        (Join-Path $repo "src/Core/Modules/modUiQuiet.bas"),
        (Join-Path $repo "src/Core/Modules/modConfig.bas"),
        (Join-Path $repo "src/Core/Modules/modAuth.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryInit.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryPublisher.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryBridgeApi.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
        (Join-Path $repo "src/Admin/Modules/modTesterSetup.bas"),
        (Join-Path $repo "src/Receiving/Modules/modReceivingInit.bas"),
        (Join-Path $repo "src/Receiving/Modules/modTS_Received.bas"),
        (Join-Path $repo "tests/unit/TestPhase2Helpers.bas"),
        (Join-Path $repo "tests/integration/test_ConfirmWrites_Tester.bas"),
        (Join-Path $repo "tests/integration/TestConfirmWritesTesterEntry.bas")
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

    $passed = [int](Invoke-MacroWithDismiss -Excel $excel -WorkbookName $harness.Name -FunctionName "TestConfirmWritesTesterEntry.RunConfirmWritesTesterIntegration")
    $context = [string](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "TestConfirmWritesTesterEntry.GetConfirmWritesTesterIntegrationContext")
    $rowsPacked = [string](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "TestConfirmWritesTesterEntry.GetConfirmWritesTesterIntegrationRows")

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
    $lines += "# Confirm Writes Tester Integration Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Machine: $env:COMPUTERNAME"
    $lines += "- Overall: $overall"
    $lines += "- Harness: $harnessPath"
    if ($contextMap.ContainsKey("Summary")) { $lines += "- Summary: $($contextMap['Summary'])" }
    if ($contextMap.ContainsKey("RuntimeUser")) { $lines += "- Runtime user: $($contextMap['RuntimeUser'])" }
    if ($contextMap.ContainsKey("TesterUser")) { $lines += "- Tester user: $($contextMap['TesterUser'])" }
    $lines += "- Passed checks: $passCount"
    $lines += "- Failed checks: $failCount"
    $lines += ""
    $lines += "| Check | Result | Detail |"
    $lines += "|---|---|---|"
    foreach ($row in $rows) {
        $lines += "| $(Sanitize-MarkdownCell $row.Check) | $(Sanitize-MarkdownCell $row.Result) | $(Sanitize-MarkdownCell $row.Detail) |"
    }
    [System.IO.File]::WriteAllLines($resultPath, $lines)

    Write-Output "CONFIRM_WRITES_TESTER_INTEGRATION_OK"
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
