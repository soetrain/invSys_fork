[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current"
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
        [string]$MacroName
    )

    $fullMacro = "'$WorkbookName'!$MacroName"
    [void]$Excel.Run($fullMacro)
}

function Begin-QuietUi {
    param([object]$Excel)

    try {
        [void]$Excel.Run("'invSys.Core.xlam'!modUiQuiet.BeginQuietUi")
    }
    catch {}
}

function End-QuietUi {
    param([object]$Excel)

    try {
        [void]$Excel.Run("'invSys.Core.xlam'!modUiQuiet.EndQuietUi")
    }
    catch {}
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

function Get-ListObjectSafe {
    param(
        [object]$Worksheet,
        [string]$TableName
    )

    if ($null -eq $Worksheet) { return $null }
    try {
        return $Worksheet.ListObjects.Item($TableName)
    }
    catch {
        return $null
    }
}

function Test-WorkbookSurface {
    param(
        [object]$Workbook,
        [hashtable[]]$TableSpecs
    )

    foreach ($spec in $TableSpecs) {
        $ws = Get-WorksheetSafe -Workbook $Workbook -WorksheetName $spec.Sheet
        if ($null -eq $ws) {
            return "Missing worksheet: $($spec.Sheet)"
        }

        $lo = Get-ListObjectSafe -Worksheet $ws -TableName $spec.Table
        if ($null -eq $lo) {
            return "Missing table: $($spec.Table)"
        }

        foreach ($columnName in $spec.Columns) {
            $hasColumn = $false
            foreach ($lc in $lo.ListColumns) {
                if ([string]::Equals($lc.Name, $columnName, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $hasColumn = $true
                    break
                }
            }
            if (-not $hasColumn) {
                return "Missing column '$columnName' in table $($spec.Table)"
            }
        }
    }

    return "OK"
}

function Test-VbComponentCode {
    param(
        [object]$Workbook,
        [string]$ComponentName,
        [string[]]$MustContain = @(),
        [string[]]$MustNotContain = @()
    )

    try {
        $component = $Workbook.VBProject.VBComponents.Item($ComponentName)
    }
    catch {
        return "Missing VB component: $ComponentName"
    }

    $lineCount = $component.CodeModule.CountOfLines
    if ($lineCount -le 0) {
        return "VB component has no code: $ComponentName"
    }

    $code = $component.CodeModule.Lines(1, $lineCount)
    foreach ($needle in $MustContain) {
        if ($code.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
            return "Missing expected code text '$needle' in $ComponentName"
        }
    }
    foreach ($needle in $MustNotContain) {
        if ($code.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return "Found retired code text '$needle' in $ComponentName"
        }
    }

    return "OK"
}

$repo = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repo $DeployRoot
$resultPath = Join-Path $repo "tests/unit/phase6_packaged_xlam_results.md"

$openOrder = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Receiving.xlam",
    "invSys.Shipping.xlam",
    "invSys.Production.xlam",
    "invSys.Admin.xlam"
)

$validationSpecs = @(
    @{
        Name = "Receiving"
        File = "invSys.Receiving.xlam"
        TargetFile = "WH1.Receiving.Operator.xlsx"
        InitMacro = "modReceivingInit.InitReceivingAddin"
        SafeMacro = "modTS_Received.EnsureGeneratedButtons"
        Tables = @(
            @{ Sheet = "ReceivedTally"; Table = "ReceivedTally"; Columns = @("REF_NUMBER", "ITEMS", "QUANTITY", "ROW") },
            @{ Sheet = "ReceivedTally"; Table = "AggregateReceived"; Columns = @("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW") },
            @{ Sheet = "ReceivedLog"; Table = "ReceivedLog"; Columns = @("SNAPSHOT_ID", "ENTRY_DATE", "REF_NUMBER", "ITEMS", "QUANTITY", "UOM", "VENDOR", "LOCATION", "ITEM_CODE", "ROW") },
            @{ Sheet = "InventoryManagement"; Table = "invSys"; Columns = @("ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION") }
        )
    },
    @{
        Name = "Shipping"
        File = "invSys.Shipping.xlam"
        TargetFile = "WH1.Shipping.Operator.xlsx"
        InitMacro = "modShippingInit.InitShippingAddin"
        SafeMacro = "modTS_Shipments.InitializeShipmentsUI"
        Tables = @(
            @{ Sheet = "ShippingBackend"; Table = "ShipmentsTally"; Columns = @("LINE_ID", "SERVER_RESERVE_EVENT_ID", "REF_NUMBER", "ITEMS", "QUANTITY", "ROW", "UOM", "LOCATION", "DESCRIPTION") },
            @{ Sheet = "ShippingBackend"; Table = "AggregatePackages"; Columns = @("ROW", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION") },
            @{ Sheet = "ShippingBackend"; Table = "AggregateBoxBOM_Log"; Columns = @("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP") },
            @{ Sheet = "ShippingBackend"; Table = "AggregatePackages_Log"; Columns = @("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP") }
        )
        FormCode = @(
            @{
                Component = "frmShipmentsTally"
                MustContain = @("NAS Inv", "Projected Inv", "Locked")
                MustNotContain = @("Current Inv", "Posted")
            }
        )
    },
    @{
        Name = "Production"
        File = "invSys.Production.xlam"
        TargetFile = "WH1.Production.Operator.xlsx"
        InitMacro = "modProductionInit.InitProductionAddin"
        SafeMacro = "mProduction.InitializeProductionUI"
        Tables = @(
            @{ Sheet = "TemplatesTable"; Table = "TemplatesTable"; Columns = @("TEMPLATE_SCOPE", "RECIPE_ID", "INGREDIENT_ID", "PROCESS", "TARGET_TABLE", "TARGET_COLUMN", "FORMULA", "GUID", "NOTES", "ACTIVE", "CREATED_AT", "UPDATED_AT") },
            @{ Sheet = "ProductionLog"; Table = "ProductionLog"; Columns = @("TIMESTAMP", "RECIPE", "RECIPE_ID", "DEPARTMENT", "DESCRIPTION", "PROCESS", "OUTPUT", "PREDICTED OUTPUT", "REAL OUTPUT", "BATCH", "BATCH_ID", "RECALL CODE", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW", "INPUT/OUTPUT", "INGREDIENT_ID", "GUID") },
            @{ Sheet = "BatchCodesLog"; Table = "BatchCodesLog"; Columns = @("RECIPE", "RECIPE_ID", "PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "TIMESTAMP", "LOCATION", "USER", "GUID") }
        )
    },
    @{
        Name = "Admin"
        File = "invSys.Admin.xlam"
        TargetFile = "WH1.Admin.Console.xlsx"
        InitMacro = "modAdminInit.InitAdminAddin"
        SafeMacro = ""
        Tables = @(
            @{ Sheet = "UserCredentials"; Table = "UserCredentials"; Columns = @("USER_ID", "USERNAME", "PIN", "ROLE", "STATUS", "LAST LOGIN") },
            @{ Sheet = "Emails"; Table = "Emails"; Columns = @("EMAIL_ID", "EMAIL_ADDRESS", "DISPLAY_NAME", "STATUS") },
            @{ Sheet = "AdminAudit"; Table = "tblAdminAudit"; Columns = @("LoggedAtUTC", "Action", "UserId", "WarehouseId", "StationId", "TargetType", "TargetId", "Reason", "Detail", "Result") },
            @{ Sheet = "PoisonQueue"; Table = "tblAdminPoisonQueue"; Columns = @("SourceWorkbook", "SourceTable", "RowIndex", "EventID", "ParentEventId", "UndoOfEventId", "EventType", "CreatedAtUTC", "WarehouseId", "StationId", "UserId", "SKU", "Qty", "Location", "Note", "PayloadJson", "Status", "RetryCount", "ErrorCode", "ErrorMessage", "FailedAtUTC") }
        )
    }
)

$resultRows = New-Object 'System.Collections.Generic.List[object]'
$excel = $null
$openedWorkbooks = New-Object 'System.Collections.Generic.List[object]'
$workbookMap = @{}
$targetWorkbooks = New-Object 'System.Collections.Generic.List[object]'
$targetRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-packaged-surfaces-" + [guid]::NewGuid().ToString("N"))

try {
    New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
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
            Add-ResultRow -Rows $resultRows -Check "$fileName.IsAddin" -Passed ([bool]$wb.IsAddin) -Detail ("IsAddin=" + [string]$wb.IsAddin)
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $false -Detail $_.Exception.Message
        }
    }

    foreach ($spec in $validationSpecs) {
        $fileName = $spec.File
        if (-not $workbookMap.ContainsKey($fileName)) {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Init" -Passed $false -Detail "Workbook not open."
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Surface" -Passed $false -Detail "Workbook not open."
            if ($spec.SafeMacro -ne "") {
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).SafeMacro" -Passed $false -Detail "Workbook not open."
            }
            if ($spec.ContainsKey("FormCode")) {
                foreach ($formSpec in $spec.FormCode) {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).$($formSpec.Component).Code" -Passed $false -Detail "Workbook not open."
                }
            }
            continue
        }

        $wb = $workbookMap[$fileName]
        $targetWb = $null

        if ($spec.ContainsKey("FormCode")) {
            foreach ($formSpec in $spec.FormCode) {
                $codeResult = Test-VbComponentCode -Workbook $wb -ComponentName $formSpec.Component -MustContain $formSpec.MustContain -MustNotContain $formSpec.MustNotContain
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).$($formSpec.Component).Code" -Passed ($codeResult -eq "OK") -Detail $codeResult
            }
        }

        try {
            $targetPath = Join-Path $targetRoot $spec.TargetFile
            $targetWb = $excel.Workbooks.Add()
            $targetWorkbooks.Add($targetWb) | Out-Null
            $targetWb.SaveAs($targetPath, 51)
            $targetWb.Activate()
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).TargetWorkbook" -Passed $false -Detail $_.Exception.Message
        }

        try {
            if ($null -ne $targetWb) { $targetWb.Activate() }
            Run-WorkbookMacro -Excel $excel -WorkbookName $wb.Name -MacroName $spec.InitMacro
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Init" -Passed $true -Detail $spec.InitMacro
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Init" -Passed $false -Detail $_.Exception.Message
        }

        if ($spec.SafeMacro -ne "") {
            try {
                if ($null -ne $targetWb) { $targetWb.Activate() }
                Begin-QuietUi -Excel $excel
                Run-WorkbookMacro -Excel $excel -WorkbookName $wb.Name -MacroName $spec.SafeMacro
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).SafeMacro" -Passed $true -Detail $spec.SafeMacro
            }
            catch {
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).SafeMacro" -Passed $false -Detail $_.Exception.Message
            }
            finally {
                End-QuietUi -Excel $excel
            }
        }

        $surfaceWorkbook = if ($null -ne $targetWb) { $targetWb } else { $wb }
        $surfaceResult = Test-WorkbookSurface -Workbook $surfaceWorkbook -TableSpecs $spec.Tables
        Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Surface" -Passed ($surfaceResult -eq "OK") -Detail $surfaceResult
    }
}
finally {
    $failedCount = @($resultRows | Where-Object { -not $_.Passed }).Count
    $passedCount = $resultRows.Count - $failedCount

    $lines = @()
    $lines += "# Phase 6 Packaged XLAM Validation Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Deploy root: $deployPath"
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

    foreach ($wb in $openedWorkbooks) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    foreach ($wb in $targetWorkbooks) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    Remove-Item -LiteralPath $targetRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$failed = @($resultRows | Where-Object { -not $_.Passed }).Count
if ($failed -gt 0) {
    Write-Output "PHASE6_PACKAGED_XLAM_VALIDATION_FAILED"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$($resultRows.Count - $failed) FAILED=$failed TOTAL=$($resultRows.Count)"
    exit 1
}

Write-Output "PHASE6_PACKAGED_XLAM_VALIDATION_OK"
Write-Output "RESULTS=$resultPath"
Write-Output "PASSED=$($resultRows.Count) FAILED=0 TOTAL=$($resultRows.Count)"
