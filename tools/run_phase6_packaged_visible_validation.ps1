[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current",

    [Parameter(Mandatory = $false)]
    [int]$PauseSeconds = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Add-ResultRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Check,
        [bool]$Passed,
        [string]$Detail = ""
    )

    $Rows.Add([pscustomobject]@{
        Check  = $Check
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
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
        default { throw "Run-WorkbookMacro supports at most 5 arguments." }
    }
}

function Get-WorksheetSafe {
    param(
        [object]$Workbook,
        [string]$WorksheetName
    )

    try {
        return $Workbook.Worksheets.Item($WorksheetName)
    }
    catch {
        return $null
    }
}

$repo = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repo $DeployRoot
$resultPath = Join-Path $repo "tests/unit/phase6_visible_packaged_results.md"
$runtimeRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-phase6-visible-" + [guid]::NewGuid().ToString("N"))

$openOrder = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Receiving.xlam",
    "invSys.Shipping.xlam",
    "invSys.Production.xlam",
    "invSys.Admin.xlam"
)

$visibleSpecs = @(
    @{
        Name = "Receiving"
        File = "invSys.Receiving.xlam"
        SafeMacro = "modTS_Received.EnsureGeneratedButtons"
        RevealSheets = @("ReceivedTally", "ReceivedLog", "InventoryManagement")
        ActivateSheet = "ReceivedTally"
    }
    @{
        Name = "Shipping"
        File = "invSys.Shipping.xlam"
        SafeMacro = "modTS_Shipments.InitializeShipmentsUI"
        RevealSheets = @("ShipmentsTally", "AggregateBoxBOM_Log", "AggregatePackages_Log")
        ActivateSheet = "ShipmentsTally"
    }
    @{
        Name = "Production"
        File = "invSys.Production.xlam"
        SafeMacro = "mProduction.InitializeProductionUI"
        RevealSheets = @("Production", "Recipes", "TemplatesTable", "ProductionLog", "BatchCodesLog")
        ActivateSheet = "Production"
    }
    @{
        Name = "Admin"
        File = "invSys.Admin.xlam"
        SafeMacro = "modAdmin.Admin_Click"
        RevealSheets = @("AdminConsole", "UserCredentials", "Emails", "AdminAudit", "PoisonQueue")
        ActivateSheet = "AdminConsole"
    }
)

$resultRows = New-Object 'System.Collections.Generic.List[object]'
$excel = $null
$openedWorkbooks = New-Object 'System.Collections.Generic.List[object]'
$workbookMap = @{}

try {
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $true
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    foreach ($fileName in $openOrder) {
        $path = Join-Path $deployPath $fileName
        if (-not (Test-Path -LiteralPath $path)) {
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $false -Detail "Missing packaged XLAM: $path"
            continue
        }

        try {
            $wb = $excel.Workbooks.Open($path)
            $openedWorkbooks.Add($wb) | Out-Null
            $workbookMap[$fileName] = $wb
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $true -Detail "Opened from $path"
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $false -Detail $_.Exception.Message
        }
    }

    if ($workbookMap.ContainsKey("invSys.Core.xlam")) {
        try {
            [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($runtimeRoot))
            Add-ResultRow -Rows $resultRows -Check "Core.RuntimeRootOverride" -Passed $true -Detail $runtimeRoot
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Core.RuntimeRootOverride" -Passed $false -Detail $_.Exception.Message
        }
    }

    foreach ($spec in $visibleSpecs) {
        $fileName = $spec.File
        if (-not $workbookMap.ContainsKey($fileName)) {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).VisibleSession" -Passed $false -Detail "Workbook not open."
            continue
        }

        $wb = $workbookMap[$fileName]
        try {
            $wb.IsAddin = $false
            $wb.Activate()
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RevealWorkbook" -Passed $true -Detail "IsAddin=False for visible inspection."
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RevealWorkbook" -Passed $false -Detail $_.Exception.Message
        }

        try {
            [void](Run-WorkbookMacro -Excel $excel -WorkbookName $wb.Name -MacroName $spec.SafeMacro)
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).SafeMacro" -Passed $true -Detail $spec.SafeMacro
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).SafeMacro" -Passed $false -Detail $_.Exception.Message
        }

        foreach ($sheetName in $spec.RevealSheets) {
            $ws = Get-WorksheetSafe -Workbook $wb -WorksheetName $sheetName
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Sheet.$sheetName" -Passed ($null -ne $ws) -Detail $sheetName
        }

        $targetSheet = Get-WorksheetSafe -Workbook $wb -WorksheetName $spec.ActivateSheet
        if ($null -ne $targetSheet) {
            try {
                $targetSheet.Activate()
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Activate.$($spec.ActivateSheet)" -Passed $true -Detail "Activated for visible inspection."
            }
            catch {
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Activate.$($spec.ActivateSheet)" -Passed $false -Detail $_.Exception.Message
            }
        }
        else {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Activate.$($spec.ActivateSheet)" -Passed $false -Detail "Worksheet missing."
        }

        Start-Sleep -Seconds $PauseSeconds
    }
}
finally {
    $failedCount = @($resultRows | Where-Object { -not $_.Passed }).Count
    $passedCount = $resultRows.Count - $failedCount

    $lines = @()
    $lines += "# Phase 6 Visible Packaged Validation Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Deploy root: $deployPath"
    $lines += "- Runtime root override: $runtimeRoot"
    $lines += "- Excel visible: true"
    $lines += "- Pause seconds per workbook: $PauseSeconds"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += ""
    $lines += "| Check | Result | Detail |"
    $lines += "|---|---|---|"
    foreach ($row in $resultRows) {
        $result = if ($row.Passed) { "PASS" } else { "FAIL" }
        $detail = [string]$row.Detail
        $detail = $detail.Replace("|", "/")
        $lines += "| $($row.Check) | $result | $detail |"
    }
    [System.IO.File]::WriteAllLines($resultPath, $lines)

    Start-Sleep -Seconds 2

    foreach ($wb in $openedWorkbooks) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }

    Remove-Item -LiteralPath $runtimeRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$failed = @($resultRows | Where-Object { -not $_.Passed }).Count
if ($failed -gt 0) {
    Write-Output "PHASE6_VISIBLE_PACKAGED_VALIDATION_FAILED"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$($resultRows.Count - $failed) FAILED=$failed TOTAL=$($resultRows.Count)"
    exit 1
}

Write-Output "PHASE6_VISIBLE_PACKAGED_VALIDATION_OK"
Write-Output "RESULTS=$resultPath"
Write-Output "PASSED=$($resultRows.Count) FAILED=0 TOTAL=$($resultRows.Count)"
