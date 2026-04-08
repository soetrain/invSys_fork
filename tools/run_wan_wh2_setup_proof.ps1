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

    if (-not (Test-Path -LiteralPath $BasPath)) {
        throw "Missing BAS module: $BasPath"
    }
    [void]$VbProject.VBComponents.Import($BasPath)
}

function New-NormalizedImportFile {
    Param([string]$SourcePath)

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-wan-wh2-" + [guid]::NewGuid().ToString("N"))
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

    if (-not (Test-Path -LiteralPath $ClassPath)) {
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

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$harnessStamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$harnessPath = Join-Path $fixtures "WanWh2Setup_Proof_Harness_$harnessStamp.xlsm"

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
        (Join-Path $repo "src/Core/Modules/modUiQuiet.bas"),
        (Join-Path $repo "src/Core/Modules/modConfig.bas"),
        (Join-Path $repo "src/Core/Modules/modAuth.bas"),
        (Join-Path $repo "src/Core/Modules/modDiagnostics.bas"),
        (Join-Path $repo "src/Core/Modules/modPerfLog.bas"),
        (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
        (Join-Path $repo "src/Core/Modules/modWarehouseSync.bas"),
        (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
        (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryInit.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryPublisher.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryBridgeApi.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
        (Join-Path $repo "tests/integration/prove_wan_wh2_setup.bas"),
        (Join-Path $repo "tests/integration/TestWanWh2SetupEntry.bas")
    )
    $classPaths = @()

    $harness = $excel.Workbooks.Add()
    $vbProject = $harness.VBProject
    foreach ($m in $modulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
    }
    foreach ($c in $classPaths) {
        Import-ClassModule -VbProject $vbProject -ClassPath $c
    }

    $harness.SaveAs($harnessPath, 52)

    $passed = [int](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "TestWanWh2SetupEntry.RunWanWh2SetupProof")
    $context = [string](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "TestWanWh2SetupEntry.GetWanWh2SetupContext")
    $contextMap = Parse-PackedMap -Packed $context
    $resultPath = if ($contextMap.ContainsKey("ResultPath")) { $contextMap["ResultPath"] } else { "" }

    Write-Output "WAN_WH2_SETUP_PROOF_OK"
    Write-Output "HARNESS=$harnessPath"
    if ($resultPath) { Write-Output "RESULTS=$resultPath" }
    if ($contextMap.ContainsKey("Summary")) { Write-Output "SUMMARY=$($contextMap['Summary'])" }
    if ($contextMap.ContainsKey("Passed") -and $contextMap.ContainsKey("Failed")) {
        Write-Output "OVERALL=$(if ($passed -eq 1) { 'PASS' } else { 'FAIL' }) PASSED=$($contextMap['Passed']) FAILED=$($contextMap['Failed'])"
    }
}
finally {
    if ($null -ne $harness) {
        try { $harness.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    Release-ComObject $harness
    Release-ComObject $excel
}
