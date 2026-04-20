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
$harnessPath = Join-Path $fixtures "CreateWarehouse_Integration_Harness_$harnessStamp.xlsm"
$resultPath = Join-Path $repo "tests/integration/create-warehouse-results.md"

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
        (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas"),
        (Join-Path $repo "src/Core/Modules/modOperatorReadModel.bas"),
        (Join-Path $repo "src/Core/Modules/modPerfLog.bas"),
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
        (Join-Path $repo "tests/integration/test_CreateWarehouse.bas")
    )

    $harness = $excel.Workbooks.Add()
    $vbProject = $harness.VBProject
    foreach ($m in $modulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
    }

    $harness.SaveAs($harnessPath, 52)

    $passed = [int](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "test_CreateWarehouse.TestCreateWarehouse_EndToEndLifecycle")
    $context = [string](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "test_CreateWarehouse.GetCreateWarehouseContextPacked")
    $rowsPacked = [string](Run-Macro -Excel $excel -WorkbookName $harness.Name -FunctionName "test_CreateWarehouse.GetCreateWarehouseEvidenceRows")

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
    $lines += "# Create Warehouse Integration Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Overall: $overall"
    $lines += "- Harness: $harnessPath"
    if ($contextMap.ContainsKey("WarehouseId")) { $lines += "- Warehouse: $($contextMap['WarehouseId'])" }
    if ($contextMap.ContainsKey("StationId")) { $lines += "- Station: $($contextMap['StationId'])" }
    if ($contextMap.ContainsKey("LocalRoot")) { $lines += "- Local root: $($contextMap['LocalRoot'])" }
    if ($contextMap.ContainsKey("SharePointRoot")) { $lines += "- SharePoint root: $($contextMap['SharePointRoot'])" }
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

    Write-Output "CREATE_WAREHOUSE_INTEGRATION_OK"
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
