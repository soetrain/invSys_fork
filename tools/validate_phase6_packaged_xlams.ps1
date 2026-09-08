[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw "Close Excel before packaged validation." }
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PackagedValidationProcess {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@

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

function Run-WorkbookMacro1 {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName,
        [object]$Argument1
    )

    $fullMacro = "'$WorkbookName'!$MacroName"
    return $Excel.Run($fullMacro, $Argument1)
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

function Test-VbComponentPresence {
    param(
        [object]$Workbook,
        [string]$ComponentName,
        [bool]$ShouldExist
    )

    $exists = $false
    try {
        $null = $Workbook.VBProject.VBComponents.Item($ComponentName)
        $exists = $true
    }
    catch {}

    if ($exists -eq $ShouldExist) { return "OK" }
    if ($ShouldExist) { return "Missing VB component: $ComponentName" }
    return "Retired VB component is still packaged: $ComponentName"
}

$repo = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repo $DeployRoot
$resultPath = Join-Path $repo "tests/unit/phase6_packaged_xlam_results.md"

function ConvertTo-SafePackagedEvidenceText {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) { return "" }
    $safe = $Text
    foreach ($sensitiveRoot in @($targetRoot, $repo, $env:USERPROFILE)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$sensitiveRoot)) {
            $safe = $safe.Replace([string]$sensitiveRoot, "<redacted-path>")
        }
    }
    $safe = [regex]::Replace($safe, '<redacted-path>(?:\\[^ ;|]+)+', '<redacted-path>')
    return $safe
}

function Set-PackagedValidationRoot {
    param([object]$Excel, [string]$Phase)
    [void]$Excel.Run("'invSys.Core.xlam'!modRuntimeWorkbooks.SetCoreDataRootOverride", $targetRoot)
    $actual = [string]$Excel.Run("'invSys.Core.xlam'!modRuntimeWorkbooks.GetCoreDataRootOverride")
    $bound = [string]::Equals($actual, $targetRoot, [StringComparison]::OrdinalIgnoreCase)
    if (-not $bound) { throw "Packaged fixture runtime root was not bound." }
    if ($Phase -eq "Initial") {
        $setup = [string]$Excel.Run("'invSys.Core.xlam'!modConfig.EnsureStationConfigEntryForAutomation",
            "WH1", "S1", "Packaged validation", (Join-Path $targetRoot "inbox"), "ADMIN",
            (Join-Path $targetRoot "WH1.invSys.Config.xlsb"), $targetRoot)
        Add-ResultRow -Rows $resultRows -Check "Fixture.ConfigProvisioned" -Passed ($setup -eq "OK") -Detail "Explicit Core setup provisioned Config inside the disposable fixture."
        if ($setup -ne "OK") { throw "Packaged Config fixture provisioning failed." }
    }
    $loaded = [bool]$Excel.Run("'invSys.Core.xlam'!modConfig.LoadConfig", "WH1", "S1")
    $bound = $bound -and $loaded
    Add-ResultRow -Rows $resultRows -Check "$Phase.IsolatedRuntimeRoot" -Passed $bound -Detail "Core runtime root is the generated fixture directory."
    if (-not $bound) { throw "Packaged fixture runtime root was not bound." }
}

function Close-PackagedValidationSession {
    param([object]$Excel, [string]$Phase)
    [uint32]$ownedProcessId = 0
    [void][PackagedValidationProcess]::GetWindowThreadProcessId([IntPtr]$Excel.Hwnd, [ref]$ownedProcessId)
    if ($ownedProcessId -eq 0 -or $null -eq $Excel.Workbooks) { throw "Owned Excel session is unavailable for cleanup." }
    $books = @($Excel.Workbooks)
    foreach ($book in $books) {
        if ([string]::IsNullOrWhiteSpace([string]$book.Path)) { continue }
        $path = [IO.Path]::GetFullPath([string]$book.FullName)
        $fixture = $path.StartsWith($targetRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
        $package = [bool]$book.IsAddin -and $path.StartsWith($deployPath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
        if (-not ($fixture -or $package)) { throw "A workbook outside the packaged fixture/package roots remains open; cleanup refused." }
    }
    $Excel.EnableEvents = $false
    $Excel.DisplayAlerts = $false
    foreach ($book in @($books | Sort-Object { [bool]$_.IsAddin })) { $book.Close($false) }
    $empty = $null -ne $Excel.Workbooks -and [int]$Excel.Workbooks.Count -eq 0
    Add-ResultRow -Rows $resultRows -Check "$Phase.OwnedWorkbooksClosed" -Passed $empty -Detail "Every workbook in the isolated Excel session closed without saving further changes."
    if (-not $empty) { throw "Owned Excel session still contains workbooks." }
    $Excel.Quit()
    Release-ComObject $Excel
    $owned = Get-Process -Id $ownedProcessId -ErrorAction SilentlyContinue
    # Only this HWND-identified process, after a real zero-workbook check.
    if ($null -ne $owned -and -not $owned.WaitForExit(1000)) { Stop-Process -Id $ownedProcessId }
}

$openOrder = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam",
    "invSys.Admin.xlam"
)

$validationSpecs = @(
    @{
        Name = "Receiving"
        File = "invSys.Operations.xlam"
        TargetFile = "WH1.Receiving.Operator.xlsx"
        InitMacro = "modReceivingInit.InitReceivingAddin"
        SafeMacro = "modTS_Received.EnsureGeneratedButtons"
        Tables = @(
            @{ Sheet = "ReceivedTally"; Table = "ReceivedTally"; Columns = @("REF_NUMBER", "ITEMS", "QUANTITY", "System_Key") },
            @{ Sheet = "ReceivedTally"; Table = "AggregateReceived"; Columns = @("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "System_Key") },
            @{ Sheet = "ReceivedLog"; Table = "ReceivedLog"; Columns = @("SNAPSHOT_ID", "ENTRY_DATE", "REF_NUMBER", "ITEMS", "QUANTITY", "UOM", "VENDOR", "LOCATION", "ITEM_CODE", "System_Key") },
            @{ Sheet = "InventoryManagement"; Table = "invSys"; Columns = @("System_Key", "SKU", "ITEM_CODE", "ITEM", "QtyOnHand", "LOCATION", "Condition", "LastRefreshUTC", "SnapshotId", "SourceType", "IsStale") }
        )
    },
    @{
        Name = "Shipping"
        File = "invSys.Operations.xlam"
        TargetFile = "WH1.Shipping.Operator.xlsx"
        InitMacro = "modShippingInit.InitShippingAddin"
        SafeMacro = "modTS_Shipments.InitializeShipmentsUI"
        FormSmokeMacro = "modTS_Shipments.ShippingTabbedNavigationSmokeForWorkbook"
        Tables = @(
            @{ Sheet = "ShippingBackend"; Table = "ShipmentsTally"; Columns = @("LINE_ID", "SERVER_RESERVE_EVENT_ID", "REF_NUMBER", "ITEMS", "QUANTITY", "System_Key", "UOM", "LOCATION", "DESCRIPTION") },
            @{ Sheet = "ShippingBackend"; Table = "AggregatePackages"; Columns = @("System_Key", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION") },
            @{ Sheet = "ShippingBackend"; Table = "AggregateBoxBOM_Log"; Columns = @("GUID", "USER", "ACTION", "System_Key", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP") },
            @{ Sheet = "ShippingBackend"; Table = "AggregatePackages_Log"; Columns = @("GUID", "USER", "ACTION", "System_Key", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP") }
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
        File = "invSys.Operations.xlam"
        TargetFile = "WH1.Production.Operator.xlsx"
        InitMacro = "modProductionInit.InitProductionAddin"
        SafeMacro = "mProduction.InitializeProductionUI"
        FormSmokeMacro = "mProduction.ProductionFormInitializeSmokeForWorkbook"
        Tables = @(
            @{ Sheet = "TemplatesTable"; Table = "TemplatesTable"; Columns = @("TEMPLATE_SCOPE", "RECIPE_ID", "INGREDIENT_ID", "PROCESS", "TARGET_TABLE", "TARGET_COLUMN", "FORMULA", "GUID", "NOTES", "ACTIVE", "CREATED_AT", "UPDATED_AT") },
            @{ Sheet = "ProductionLog"; Table = "ProductionLog"; Columns = @("TIMESTAMP", "RECIPE", "RECIPE_ID", "DEPARTMENT", "DESCRIPTION", "PROCESS", "OUTPUT", "PREDICTED OUTPUT", "REAL OUTPUT", "BATCH", "BATCH_ID", "RECALL CODE", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "ITEM", "UOM", "QUANTITY", "LOCATION", "System_Key", "INPUT/OUTPUT", "INGREDIENT_ID", "GUID") },
            @{ Sheet = "BatchCodesLog"; Table = "BatchCodesLog"; Columns = @("RECIPE", "RECIPE_ID", "PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "TIMESTAMP", "LOCATION", "USER", "GUID") }
        )
    },
    @{
        Name = "Admin"
        File = "invSys.Admin.xlam"
        TargetFile = "WH1.Admin.Console.xlsx"
        InitMacro = "modAdminInit.InitAdminAddin"
        SafeMacro = ""
        FormSmokeMacro = "modAdmin.AdminSettingsFormInitializeSmokeForWorkbook"
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

    $sentinelWb = $excel.Workbooks.Add()
    $targetWorkbooks.Add($sentinelWb) | Out-Null
    $sentinelWs = $sentinelWb.Worksheets.Item(1)
    $sentinelWs.Name = "OperatorSentinel"
    $sentinelWs.Range("A1").Value2 = "UNCHANGED"
    $sentinelWs.Range("A3").Value2 = "ROW"
    $sentinelWs.Range("B3").Value2 = "LOCATION"
    $sentinelWs.Range("A4").Value2 = 99
    $sentinelWs.Range("B4").Value2 = "CLEARVIEW"
    $sentinelTable = $sentinelWs.ListObjects.Add(1, $sentinelWs.Range("A3:B4"), $null, 1)
    $sentinelTable.Name = "invSys"
    $sentinelWb.Names.Add("RunLocation", $sentinelWs.Range("B4")) | Out-Null
    $sentinelWb.Activate()

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
            if ($fileName -eq "invSys.Core.xlam") { Set-PackagedValidationRoot -Excel $excel -Phase "Initial" }
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $true -Detail "Opened from $path"
            Add-ResultRow -Rows $resultRows -Check "$fileName.IsAddin" -Passed ([bool]$wb.IsAddin) -Detail ("IsAddin=" + [string]$wb.IsAddin)
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $false -Detail $_.Exception.Message
        }

        if ($fileName -eq "invSys.Designs.Domain.xlam") {
            $sentinelUnchanged = (
                [string]$sentinelWs.Range("A1").Value2 -eq "UNCHANGED" -and
                [int]$sentinelWs.Range("A4").Value2 -eq 99 -and
                [string]$sentinelWs.Range("B4").Value2 -eq "CLEARVIEW" -and
                $sentinelWs.ListObjects.Count -eq 1 -and
                $sentinelWb.Worksheets.Count -eq 1
            )
            Add-ResultRow -Rows $resultRows -Check "DomainStartup.OperatorIsolation" -Passed $sentinelUnchanged -Detail "Core, Inventory Domain, and Designs Domain left the active operator sentinel unchanged."
        }
    }

    $autoOpenMacros = @(
        @{ File = "invSys.Inventory.Domain.xlam"; Macro = "modInventoryInit.Auto_Open" },
        @{ File = "invSys.Designs.Domain.xlam"; Macro = "modDesignsInit.Auto_Open" },
        @{ File = "invSys.Operations.xlam"; Macro = "modOperationsInit.Auto_Open" },
        @{ File = "invSys.Admin.xlam"; Macro = "modAdminInit.Auto_Open" }
    )
    $autoOpenError = ""
    foreach ($autoOpenSpec in $autoOpenMacros) {
        if (-not $workbookMap.ContainsKey($autoOpenSpec.File)) {
            $autoOpenError = "Workbook not open: $($autoOpenSpec.File)"
            break
        }
        try {
            $sentinelWb.Activate()
            Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap[$autoOpenSpec.File].Name -MacroName $autoOpenSpec.Macro
        }
        catch {
            $autoOpenError = "$($autoOpenSpec.Macro): $($_.Exception.Message)"
            break
        }
    }
    $sentinelUnchanged = (
        $autoOpenError -eq "" -and
        [string]$sentinelWs.Range("A1").Value2 -eq "UNCHANGED" -and
        [int]$sentinelWs.Range("A4").Value2 -eq 99 -and
        [string]$sentinelWs.Range("B4").Value2 -eq "CLEARVIEW" -and
        $sentinelWs.ListObjects.Count -eq 1 -and
        $sentinelWb.Worksheets.Count -eq 1
    )
    $autoOpenDetail = if ($autoOpenError -eq "") {
        "All Domain/Role/Admin Auto_Open entry points left the active operator workbook unchanged."
    }
    else {
        $autoOpenError
    }
    Add-ResultRow -Rows $resultRows -Check "XLAMStartup.ExplicitAutoOpenOperatorIsolation" -Passed $sentinelUnchanged -Detail $autoOpenDetail

    try {
        $eventSentinelPath = Join-Path $targetRoot "Unrelated.Operator.xlsx"
        $eventSentinelWb = $excel.Workbooks.Add()
        $eventSentinelWs = $eventSentinelWb.Worksheets.Item(1)
        $eventSentinelWs.Range("A1").Value2 = "EVENT-UNCHANGED"
        $eventSentinelWb.SaveAs($eventSentinelPath, 51)
        $eventSentinelWb.Close($false)
        Release-ComObject $eventSentinelWb
        $eventSentinelWb = $excel.Workbooks.Open($eventSentinelPath)
        $targetWorkbooks.Add($eventSentinelWb) | Out-Null
        $eventSentinelWs = $eventSentinelWb.Worksheets.Item(1)
        $eventIsolationPassed = (
            [string]$eventSentinelWs.Range("A1").Value2 -eq "EVENT-UNCHANGED" -and
            $eventSentinelWb.Worksheets.Count -eq 1 -and
            $eventSentinelWs.ListObjects.Count -eq 0
        )
        Add-ResultRow -Rows $resultRows -Check "RoleEvents.WorkbookOpenOperatorIsolation" -Passed $eventIsolationPassed -Detail "WorkbookOpen/NewWorkbook handlers left an unrelated operator workbook unchanged."
    }
    catch {
        Add-ResultRow -Rows $resultRows -Check "RoleEvents.WorkbookOpenOperatorIsolation" -Passed $false -Detail $_.Exception.Message
    }

    try {
        $roleNameSentinelWb = $excel.Workbooks.Add()
        $targetWorkbooks.Add($roleNameSentinelWb) | Out-Null
        $roleSheetNames = @("Production", "ShipmentsTally", "ReceivedTally")
        foreach ($roleSheetName in $roleSheetNames) {
            $roleNameWs = $roleNameSentinelWb.Worksheets.Add()
            $roleNameWs.Name = $roleSheetName
            [void]$roleNameWs.Activate()
            [void]$roleNameWs.Range("A1").Select()
            $roleNameWs.Range("A1").Value2 = "$roleSheetName-UNCHANGED"
        }
        $roleNameIsolationPassed = ($roleNameSentinelWb.Worksheets.Count -eq 4)
        foreach ($roleSheetName in $roleSheetNames) {
            $roleNameWs = $roleNameSentinelWb.Worksheets.Item($roleSheetName)
            $roleNameIsolationPassed = (
                $roleNameIsolationPassed -and
                [string]$roleNameWs.Range("A1").Value2 -eq "$roleSheetName-UNCHANGED" -and
                $roleNameWs.ListObjects.Count -eq 0
            )
        }
        Add-ResultRow -Rows $resultRows -Check "RoleEvents.NamedSheetOperatorIsolation" -Passed $roleNameIsolationPassed -Detail "Selection/change handlers ignored unrelated workbooks whose sheet names resembled role sheets but lacked role-owned tables."
    }
    catch {
        Add-ResultRow -Rows $resultRows -Check "RoleEvents.NamedSheetOperatorIsolation" -Passed $false -Detail $_.Exception.Message
    }

    $componentSpecs = @(
        @{ File = "invSys.Core.xlam"; Component = "modInventoryDomainBridge"; Exists = $true },
        @{ File = "invSys.Core.xlam"; Component = "modDesignsDomainBridge"; Exists = $true },
        @{ File = "invSys.Core.xlam"; Component = "modInventoryViewerData"; Exists = $true },
        @{ File = "invSys.Core.xlam"; Component = "frmItemSearch"; Exists = $true },
        @{ File = "invSys.Inventory.Domain.xlam"; Component = "modInventoryApply"; Exists = $true },
        @{ File = "invSys.Inventory.Domain.xlam"; Component = "modInventoryQueries"; Exists = $true },
        @{ File = "invSys.Inventory.Domain.xlam"; Component = "modInvMan"; Exists = $false },
        @{ File = "invSys.Inventory.Domain.xlam"; Component = "cInventoryAppEvents"; Exists = $false },
        @{ File = "invSys.Designs.Domain.xlam"; Component = "modDesignsApply"; Exists = $true },
        @{ File = "invSys.Designs.Domain.xlam"; Component = "modDesignsQueries"; Exists = $true },
        @{ File = "invSys.Designs.Domain.xlam"; Component = "modDesignsSchema"; Exists = $true },
        @{ File = "invSys.Admin.xlam"; Component = "modAdminConsole"; Exists = $true },
        @{ File = "invSys.Admin.xlam"; Component = "modAdminDesignLifecycle"; Exists = $true },
        @{ File = "invSys.Operations.xlam"; Component = "modInventoryViewer"; Exists = $true },
        @{ File = "invSys.Operations.xlam"; Component = "frmInventoryViewer"; Exists = $true },
        @{ File = "invSys.Admin.xlam"; Component = "frmAdminControls"; Exists = $false },
        @{ File = "invSys.Admin.xlam"; Component = "frmAdminEmail"; Exists = $false },
        @{ File = "invSys.Admin.xlam"; Component = "frmEditUser"; Exists = $false },
        @{ File = "invSys.Admin.xlam"; Component = "ufAdminItemSearch"; Exists = $false },
        @{ File = "invSys.Admin.xlam"; Component = "ufDynItemSearchTemplate"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "frmCreateRecipeTable"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "frmCreateSubstitutionList"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "frmIngredientPalette"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "frmSubstitution"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "ufProductionItemSearch"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "frmReceivingSavedList"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "ufReceivingItemSearch"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "frmShippingCreateList"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "frmShippingSavedList"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "ufShippingItemSearch"; Exists = $false },
        @{ File = "invSys.Operations.xlam"; Component = "ufDynItemSearchTemplate"; Exists = $false }
    )
    foreach ($componentSpec in $componentSpecs) {
        if (-not $workbookMap.ContainsKey($componentSpec.File)) {
            Add-ResultRow -Rows $resultRows -Check "$($componentSpec.File).$($componentSpec.Component)" -Passed $false -Detail "Workbook not open."
            continue
        }
        $presenceResult = Test-VbComponentPresence -Workbook $workbookMap[$componentSpec.File] -ComponentName $componentSpec.Component -ShouldExist $componentSpec.Exists
        Add-ResultRow -Rows $resultRows -Check "$($componentSpec.File).$($componentSpec.Component)" -Passed ($presenceResult -eq "OK") -Detail $presenceResult
    }

    foreach ($spec in $validationSpecs) {
        $fileName = $spec.File
        if (-not $workbookMap.ContainsKey($fileName)) {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Init" -Passed $false -Detail "Workbook not open."
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Surface" -Passed $false -Detail "Workbook not open."
            if ($spec.SafeMacro -ne "") {
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).SafeMacro" -Passed $false -Detail "Workbook not open."
            }
            if ($spec.ContainsKey("FormSmokeMacro")) {
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).FormInitialize" -Passed $false -Detail "Workbook not open."
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

        if ($spec.ContainsKey("FormSmokeMacro")) {
            try {
                if ($null -eq $targetWb) {
                    throw "Production target workbook is not open."
                }
                $targetWb.Activate()
                Begin-QuietUi -Excel $excel
                $formSmoke = [string](Run-WorkbookMacro1 -Excel $excel -WorkbookName $wb.Name -MacroName $spec.FormSmokeMacro -Argument1 $targetWb)
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).FormInitialize" -Passed $formSmoke.StartsWith("OK|", [System.StringComparison]::Ordinal) -Detail $formSmoke
            }
            catch {
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).FormInitialize" -Passed $false -Detail $_.Exception.Message
            }
            finally {
                End-QuietUi -Excel $excel
            }
        }

        $surfaceWorkbook = if ($null -ne $targetWb) { $targetWb } else { $wb }
        $surfaceResult = Test-WorkbookSurface -Workbook $surfaceWorkbook -TableSpecs $spec.Tables
        Add-ResultRow -Rows $resultRows -Check "$($spec.Name).Surface" -Passed ($surfaceResult -eq "OK") -Detail $surfaceResult
    }

    if ($workbookMap.ContainsKey("invSys.Admin.xlam")) {
        try {
            $adminMacro = "'$($workbookMap["invSys.Admin.xlam"].Name)'!modAdminConsole.ReissuePoisonReceiveEventReportForAutomation"
            $adminSmoke = [string]$excel.Run($adminMacro, "__PACKAGED_SMOKE_MISSING__.xlsb", "EVT-MISSING", "SKU-SMOKE", 1, "A1")
            $adminSmokePassed = $adminSmoke.StartsWith("FAIL|Report=Source workbook not open:", [System.StringComparison]::Ordinal)
            Add-ResultRow -Rows $resultRows -Check "Admin.PoisonReissue.PackagedSurface" -Passed $adminSmokePassed -Detail $adminSmoke
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Admin.PoisonReissue.PackagedSurface" -Passed $false -Detail $_.Exception.Message
        }
        try {
            $designLifecycleLayout = [int]$excel.Run("'$($workbookMap["invSys.Admin.xlam"].Name)'!modAdminDesignLifecycle.DesignLifecycleFormLayoutSmokeForAutomation")
            Add-ResultRow -Rows $resultRows -Check "Admin.DesignLifecycle.LegacyMigrationControl" -Passed ($designLifecycleLayout -eq 1) -Detail "LayoutReady=$designLifecycleLayout"
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Admin.DesignLifecycle.LegacyMigrationControl" -Passed $false -Detail $_.Exception.Message
        }
        try {
            $editSelection = [string]$excel.Run("'$($workbookMap["invSys.Admin.xlam"].Name)'!modAdmin.InventoryEditSelectionContractForAutomation")
            $editSelectionPassed = $editSelection -match '^OK\|' -and
                $editSelection -match '(?:^|\|)ComboSelected=True(?:\||$)' -and
                $editSelection -match '(?:^|\|)FieldsLoaded=True(?:\||$)' -and
                $editSelection -match '(?:^|\|)UtilityReady=True(?:\||$)'
            Add-ResultRow -Rows $resultRows -Check "Admin.EditItemComboSelection" -Passed $editSelectionPassed -Detail $editSelection
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Admin.EditItemComboSelection" -Passed $false -Detail $_.Exception.Message
        }
        try {
            $adminAdd = [string]$excel.Run("'$($workbookMap["invSys.Admin.xlam"].Name)'!modAdmin.InventoryAddVisibilityDropdownContractForAutomation")
            $adminAddPassed = $adminAdd -match '^OK\|' -and
                $adminAdd -match '(?:^|\|)SubmitHandler=True(?:\||$)' -and
                $adminAdd -match '(?:^|\|)ExactEntityCreate=True(?:\||$)' -and
                $adminAdd -match '(?:^|\|)ZeroQtyAccepted=True(?:\||$)' -and
                $adminAdd -match '(?:^|\|)NegativeRejected=True(?:\||$)' -and
                $adminAdd -match '(?:^|\|)LocationDropdown=True(?:\||$)' -and
                $adminAdd -match '(?:^|\|)CategoryDropdown=True(?:\||$)'
            Add-ResultRow -Rows $resultRows -Check "Admin.InventoryAddVisibilityDropdowns" -Passed $adminAddPassed -Detail $adminAdd
            Add-ResultRow -Rows $resultRows -Check "Admin.InventoryZeroStartingQuantity" -Passed $adminAddPassed -Detail $adminAdd
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Admin.InventoryAddVisibilityDropdowns" -Passed $false -Detail $_.Exception.Message
        }
        try {
            $inventoryDelete = [string]$excel.Run("'$($workbookMap["invSys.Admin.xlam"].Name)'!modAdmin.InventoryDeleteContractForAutomation")
            $inventoryDeletePassed = $inventoryDelete -match '^OK\|' -and
                $inventoryDelete -match '(?:^|\|)DeleteHandler=True(?:\||$)' -and
                $inventoryDelete -match '(?:^|\|)ExactKey=True(?:\||$)' -and
                $inventoryDelete -match '(?:^|\|)UtilityZeroDelta=True(?:\||$)'
            Add-ResultRow -Rows $resultRows -Check "Admin.InventoryDeleteItem" -Passed $inventoryDeletePassed -Detail $inventoryDelete
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Admin.InventoryDeleteItem" -Passed $false -Detail $_.Exception.Message
        }
        try {
            $inventoryWorksheetWb = $excel.Workbooks.Add()
            $targetWorkbooks.Add($inventoryWorksheetWb) | Out-Null
            $inventoryWorksheetPath = Join-Path $targetRoot "Admin.Inventory.Worksheet.Contract.xlsx"
            $inventoryWorksheetWb.SaveAs($inventoryWorksheetPath, 51)
            $inventoryWorksheetWb.Activate()
            $inventoryWorksheet = [string]$excel.Run(
                "'$($workbookMap["invSys.Admin.xlam"].Name)'!modAdmin.InventoryWorksheetContractForAutomation",
                $inventoryWorksheetWb)
            $inventoryWorksheetPassed = $inventoryWorksheet -match '^OK\|' -and
                $inventoryWorksheet -match '(?:^|\|)TableCreated=True(?:\||$)' -and
                $inventoryWorksheet -match '(?:^|\|)Preflight=True(?:\||$)' -and
                $inventoryWorksheet -match '(?:^|\|)Utility=True(?:\||$)' -and
                $inventoryWorksheet -match '(?:^|\|)ZeroCounted=True(?:\||$)'
            Add-ResultRow -Rows $resultRows -Check "Admin.InventoryWorksheetActions" -Passed $inventoryWorksheetPassed -Detail $inventoryWorksheet
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Admin.InventoryWorksheetActions" -Passed $false -Detail $_.Exception.Message
        }
    }

    if ($workbookMap.ContainsKey("invSys.Core.xlam") -and $workbookMap.ContainsKey("invSys.Inventory.Domain.xlam")) {
        try {
            $inventoryZeroCreate = [string]$excel.Run("'$($workbookMap["invSys.Inventory.Domain.xlam"].Name)'!modInventoryApply.InventoryZeroCreateContractForAutomation")
            $inventoryZeroCreatePassed = $inventoryZeroCreate -match '^OK\|' -and
                $inventoryZeroCreate -match '(?:^|\|)ZeroEntityActive=True(?:\||$)' -and
                $inventoryZeroCreate -match '(?:^|\|)ZeroVisible=True(?:\||$)' -and
                $inventoryZeroCreate -match '(?:^|\|)NegativeRejected=True(?:\||$)'
            Add-ResultRow -Rows $resultRows -Check "InventoryDomain.InventoryZeroCreate" -Passed $inventoryZeroCreatePassed -Detail $inventoryZeroCreate
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "InventoryDomain.InventoryZeroCreate" -Passed $false -Detail $_.Exception.Message
        }
        try {
            $inventoryRetirement = [string]$excel.Run("'$($workbookMap["invSys.Inventory.Domain.xlam"].Name)'!modInventoryApply.InventoryRetirementContractForAutomation")
            $inventoryRetirementPassed = $inventoryRetirement -match '^OK\|' -and
                $inventoryRetirement -match '(?:^|\|)CountedRetired=True(?:\||$)' -and
                $inventoryRetirement -match '(?:^|\|)UtilityRetired=True(?:\||$)' -and
                $inventoryRetirement -match '(?:^|\|)ActiveRows=0(?:\||$)'
            Add-ResultRow -Rows $resultRows -Check "InventoryDomain.InventoryRetirement" -Passed $inventoryRetirementPassed -Detail $inventoryRetirement
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "InventoryDomain.InventoryRetirement" -Passed $false -Detail $_.Exception.Message
        }
        try {
            $workbookMap["invSys.Inventory.Domain.xlam"].Close($false)
            $diagnosticMacro = "'$($workbookMap["invSys.Core.xlam"].Name)'!modInventoryDomainBridge.DiagnoseInventoryDomainBridge"
            $diagnostic = [string]$excel.Run($diagnosticMacro)
            $passed = -not [string]::IsNullOrWhiteSpace($diagnostic) -and `
                      -not $diagnostic.StartsWith("Inventory Domain unavailable", [System.StringComparison]::OrdinalIgnoreCase)
            Add-ResultRow -Rows $resultRows -Check "InventoryDomain.PeerAutoLoad" -Passed $passed -Detail $diagnostic
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "InventoryDomain.PeerAutoLoad" -Passed $false -Detail $_.Exception.Message
        }
    }

    if ($workbookMap.ContainsKey("invSys.Core.xlam") -and $workbookMap.ContainsKey("invSys.Designs.Domain.xlam")) {
        try {
            $workbookMap["invSys.Designs.Domain.xlam"].Close($false)
            $diagnosticMacro = "'$($workbookMap["invSys.Core.xlam"].Name)'!modDesignsDomainBridge.DiagnoseDesignsDomainBridge"
            $diagnostic = [string]$excel.Run($diagnosticMacro)
            $domainReopened = $false
            foreach ($candidate in $excel.Workbooks) {
                if ([string]::Equals([string]$candidate.Name, "invSys.Designs.Domain.xlam", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $domainReopened = $true
                    break
                }
            }
            $passed = -not [string]::IsNullOrWhiteSpace($diagnostic) -and `
                      -not $diagnostic.StartsWith("Designs Domain unavailable", [System.StringComparison]::OrdinalIgnoreCase)
            Add-ResultRow -Rows $resultRows -Check "DesignsDomain.PeerAutoLoad" -Passed $passed -Detail "$diagnostic; WorkbookOpen=$domainReopened"
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "DesignsDomain.PeerAutoLoad" -Passed $false -Detail $_.Exception.Message
        }
    }

    # Cross a real Excel process boundary, then reopen the full packaged set and
    # every saved role/Admin workbook. This catches stale XLAM references,
    # workbook identity drift, and startup mutation that an in-process reopen
    # cannot expose.
    foreach ($targetWb in $targetWorkbooks) {
        try {
            if (-not [string]::IsNullOrWhiteSpace([string]$targetWb.Path)) { $targetWb.Save() }
        }
        catch {}
    }
    Close-PackagedValidationSession -Excel $excel -Phase "Initial"
    foreach ($targetWb in $targetWorkbooks) {
        Release-ComObject $targetWb
    }
    foreach ($addinWb in $openedWorkbooks) {
        Release-ComObject $addinWb
    }
    $excel = $null

    $openedWorkbooks = New-Object 'System.Collections.Generic.List[object]'
    $targetWorkbooks = New-Object 'System.Collections.Generic.List[object]'
    $workbookMap = @{}
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    foreach ($fileName in $openOrder) {
        try {
            $reopenPath = Join-Path $deployPath $fileName
            $reopenedAddin = $excel.Workbooks.Open($reopenPath)
            $openedWorkbooks.Add($reopenedAddin) | Out-Null
            $workbookMap[$fileName] = $reopenedAddin
            if ($fileName -eq "invSys.Core.xlam") { Set-PackagedValidationRoot -Excel $excel -Phase "Restart" }
            $sameIdentity = [string]::Equals([string]$reopenedAddin.FullName, [string](Resolve-Path $reopenPath).Path, [System.StringComparison]::OrdinalIgnoreCase)
            Add-ResultRow -Rows $resultRows -Check "Restart.$fileName" -Passed ([bool]$reopenedAddin.IsAddin -and $sameIdentity) -Detail "IsAddin=$($reopenedAddin.IsAddin); FullName=$($reopenedAddin.FullName)"
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Restart.$fileName" -Passed $false -Detail $_.Exception.Message
        }
    }

    foreach ($spec in $validationSpecs) {
        try {
            $operatorPath = Join-Path $targetRoot $spec.TargetFile
            $reopenedOperator = $excel.Workbooks.Open($operatorPath)
            $targetWorkbooks.Add($reopenedOperator) | Out-Null
            $surfaceResult = Test-WorkbookSurface -Workbook $reopenedOperator -TableSpecs $spec.Tables
            $identityOk = [string]::Equals([string]$reopenedOperator.FullName, [string](Resolve-Path $operatorPath).Path, [System.StringComparison]::OrdinalIgnoreCase)
            Add-ResultRow -Rows $resultRows -Check "Restart.$($spec.Name).SavedWorkbook" -Passed ($identityOk -and $surfaceResult -eq "OK") -Detail "FullName=$($reopenedOperator.FullName); Surface=$surfaceResult"
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Restart.$($spec.Name).SavedWorkbook" -Passed $false -Detail $_.Exception.Message
        }
    }

    if ($workbookMap.ContainsKey("invSys.Core.xlam")) {
        try {
            $inventoryDiagnostic = [string]$excel.Run("'$($workbookMap["invSys.Core.xlam"].Name)'!modInventoryDomainBridge.DiagnoseInventoryDomainBridge")
            $designsDiagnostic = [string]$excel.Run("'$($workbookMap["invSys.Core.xlam"].Name)'!modDesignsDomainBridge.DiagnoseDesignsDomainBridge")
            $domainRestartOk = -not [string]::IsNullOrWhiteSpace($inventoryDiagnostic) `
                -and -not [string]::IsNullOrWhiteSpace($designsDiagnostic) `
                -and -not $inventoryDiagnostic.StartsWith("Inventory Domain unavailable", [System.StringComparison]::OrdinalIgnoreCase) `
                -and -not $designsDiagnostic.StartsWith("Designs Domain unavailable", [System.StringComparison]::OrdinalIgnoreCase)
            Add-ResultRow -Rows $resultRows -Check "Restart.DomainBridges" -Passed $domainRestartOk -Detail "Inventory=$inventoryDiagnostic; Designs=$designsDiagnostic"
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Restart.DomainBridges" -Passed $false -Detail $_.Exception.Message
        }
    }
}
finally {
    if ($null -ne $excel) {
        try { Close-PackagedValidationSession -Excel $excel -Phase "Final" }
        catch { Add-ResultRow -Rows $resultRows -Check "Final.Cleanup" -Passed $false -Detail $_.Exception.Message }
    }
    $failedCount = @($resultRows | Where-Object { -not $_.Passed }).Count
    $passedCount = $resultRows.Count - $failedCount

    $lines = @()
    $lines += "# Phase 6 Packaged XLAM Validation Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Deploy root: $(ConvertTo-SafePackagedEvidenceText $deployPath)"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += ""
    $lines += "| Check | Result | Detail |"
    $lines += "|---|---|---|"
    foreach ($row in $resultRows) {
        $result = if ($row.Passed) { "PASS" } else { "FAIL" }
        $detail = ConvertTo-SafePackagedEvidenceText ([string]$row.Detail)
        $detail = $detail.Replace("|", "/")
        $lines += "| $($row.Check) | $result | $detail |"
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($resultPath, (($lines -join "`n") + "`n"), $utf8NoBom)

    foreach ($wb in $openedWorkbooks) {
        Release-ComObject $wb
    }
    foreach ($wb in $targetWorkbooks) {
        Release-ComObject $wb
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
