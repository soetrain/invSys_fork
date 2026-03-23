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

function Start-ExcelEnterDismissal {
    param([int]$Seconds = 8)

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

function Invoke-WorkbookMacroWithDismiss {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName,
        [int]$DismissSeconds = 8
    )

    $job = Start-ExcelEnterDismissal -Seconds $DismissSeconds
    try {
        return Run-WorkbookMacro -Excel $Excel -WorkbookName $WorkbookName -MacroName $MacroName
    }
    finally {
        if ($null -ne $job) {
            try { Wait-Job -Job $job -Timeout ($DismissSeconds + 2) | Out-Null } catch {}
            try { Receive-Job -Job $job -ErrorAction SilentlyContinue | Out-Null } catch {}
            try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue } catch {}
        }
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

function Get-ColumnIndexSafe {
    param(
        [object]$ListObject,
        [string]$ColumnName
    )

    if ($null -eq $ListObject) { return 0 }
    foreach ($lc in $ListObject.ListColumns) {
        if ([string]::Equals([string]$lc.Name, $ColumnName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [int]$lc.Index
        }
    }
    return 0
}

function Clear-ListObjectRows {
    param([object]$ListObject)

    if ($null -eq $ListObject) { return }
    while ($ListObject.ListRows.Count -gt 0) {
        $ListObject.ListRows.Item($ListObject.ListRows.Count).Delete()
    }
}

function Add-ListObjectRow {
    param(
        [object]$ListObject,
        [hashtable]$Values
    )

    if ($null -eq $ListObject) { throw "ListObject missing." }
    $targetRange = $null
    if ($ListObject.ListRows.Count -eq 0) {
        try {
            $targetRange = $ListObject.InsertRowRange
        }
        catch {
            $targetRange = $null
        }
    }

    if ($null -eq $targetRange) {
        try {
            $row = $ListObject.ListRows.Add($null, $false)
        }
        catch {
            $row = $ListObject.ListRows.Add()
        }
        $targetRange = $row.Range
    }

    foreach ($key in $Values.Keys) {
        $idx = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName ([string]$key)
        if ($idx -gt 0) {
            $value = $Values[$key]
            try {
                if ($null -eq $value) {
                    $targetRange.Cells.Item(1, [int]$idx).Value2 = $null
                }
                elseif ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal] -or $value -is [single]) {
                    $targetRange.Cells.Item(1, [int]$idx).Value2 = [double]$value
                }
                elseif ($value -is [datetime]) {
                    $targetRange.Cells.Item(1, [int]$idx).Value2 = $value.ToOADate()
                }
                else {
                    $targetRange.Cells.Item(1, [int]$idx).Value2 = [string]$value
                }
            }
            catch {
                throw "Failed to seed table '$($ListObject.Name)' column '$key' with value '$value': $($_.Exception.Message)"
            }
        }
    }
}

function Get-RowCountSafe {
    param([object]$ListObject)
    if ($null -eq $ListObject) { return 0 }
    return [int]$ListObject.ListRows.Count
}

function Find-RowIndexByValue {
    param(
        [object]$ListObject,
        [string]$ColumnName,
        [object]$ExpectedValue
    )

    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) { return 0 }
    $idx = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName $ColumnName
    if ($idx -le 0) { return 0 }

    for ($i = 1; $i -le $ListObject.ListRows.Count; $i++) {
        $actual = [string]$ListObject.DataBodyRange.Cells.Item([int]$i, [int]$idx).Value2
        if ($actual -eq [string]$ExpectedValue) {
            return $i
        }
    }
    return 0
}

function Get-RowValueSafe {
    param(
        [object]$ListObject,
        [int]$RowIndex,
        [string]$ColumnName
    )

    if ($null -eq $ListObject -or $RowIndex -le 0 -or $null -eq $ListObject.DataBodyRange) { return $null }
    $idx = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName $ColumnName
    if ($idx -le 0) { return $null }
    return $ListObject.DataBodyRange.Cells.Item([int]$RowIndex, [int]$idx).Value2
}

function Add-Table {
    param(
        [object]$Worksheet,
        [string]$TableName,
        [object[]]$Headers,
        [object[][]]$Rows
    )

    $colCount = $Headers.Count
    $rowCount = [Math]::Max($Rows.Count, 1)

    $Worksheet.Range("A1").Resize(1, $colCount).Value = ,$Headers
    if ($Rows.Count -gt 0) {
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            for ($c = 0; $c -lt $colCount; $c++) {
                $value = $null
                if ($Rows[$r] -is [System.Array]) {
                    if ($c -le $Rows[$r].GetUpperBound(0)) {
                        $value = $Rows[$r][$c]
                    }
                }
                else {
                    if ($c -eq 0) { $value = $Rows[$r] }
                }

                if ($null -eq $value) {
                    $Worksheet.Cells($r + 2, $c + 1).Value2 = $null
                }
                elseif ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal] -or $value -is [single]) {
                    $Worksheet.Cells($r + 2, $c + 1).Value2 = [double]$value
                }
                elseif ($value -is [datetime]) {
                    $Worksheet.Cells($r + 2, $c + 1).Value2 = $value.ToOADate()
                }
                else {
                    $Worksheet.Cells($r + 2, $c + 1).Value2 = [string]$value
                }
            }
        }
    }
    else {
        $Worksheet.Range("A2").Resize(1, $colCount).Value = ,([object[]]::new($colCount))
    }

    $endCell = $Worksheet.Cells($rowCount + 1, $colCount)
    $range = $Worksheet.Range("A1", $endCell)
    $listObject = $Worksheet.ListObjects.Add(1, $range, $null, 1)
    $listObject.Name = $TableName
    for ($i = 0; $i -lt $colCount; $i++) {
        $listObject.HeaderRowRange.Cells.Item(1, $i + 1).Value2 = [string]$Headers[$i]
        $listObject.ListColumns.Item($i + 1).Name = [string]$Headers[$i]
    }
    return $listObject
}

function Save-NewWorkbook {
    param(
        [object]$Workbook,
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
    $Workbook.SaveAs($Path, 50)
}

function New-ConfigWorkbook {
    param(
        [object]$Excel,
        [string]$Path,
        [string]$WarehouseId,
        [string]$StationId,
        [string]$RuntimeRoot
    )

    $wb = $Excel.Workbooks.Add()
    $wsWh = $wb.Worksheets.Item(1)
    $wsWh.Name = "WarehouseConfig"
    $wsSt = $wb.Worksheets.Add()
    $wsSt.Name = "StationConfig"

    $loWh = Add-Table -Worksheet $wsWh -TableName "tblWarehouseConfig" -Headers @(
        "WarehouseId", "WarehouseName", "Timezone", "DefaultLocation",
        "BatchSize", "LockTimeoutMinutes", "HeartbeatIntervalSeconds", "MaxLockHoldMinutes",
        "SnapshotCadence", "BackupCadence", "PathDataRoot", "PathBackupRoot", "PathSharePointRoot",
        "DesignsEnabled", "PoisonRetryMax", "AuthCacheTTLSeconds", "ProcessorServiceUserId",
        "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot"
    ) -Rows @()
    Clear-ListObjectRows $loWh
    Add-ListObjectRow -ListObject $loWh -Values @{
        "WarehouseId" = $WarehouseId
        "WarehouseName" = "Main Warehouse"
        "Timezone" = "UTC"
        "DefaultLocation" = "A1"
        "BatchSize" = 500
        "LockTimeoutMinutes" = 3
        "HeartbeatIntervalSeconds" = 30
        "MaxLockHoldMinutes" = 2
        "SnapshotCadence" = "PER_BATCH"
        "BackupCadence" = "DAILY"
        "PathDataRoot" = $RuntimeRoot
        "PathBackupRoot" = (Join-Path $RuntimeRoot "Backups")
        "PathSharePointRoot" = ""
        "DesignsEnabled" = $false
        "PoisonRetryMax" = 3
        "AuthCacheTTLSeconds" = 300
        "ProcessorServiceUserId" = "svc_processor"
        "FF_DesignsEnabled" = $false
        "FF_OutlookAlerts" = $false
        "FF_AutoSnapshot" = $true
    }

    $loSt = Add-Table -Worksheet $wsSt -TableName "tblStationConfig" -Headers @(
        "StationId", "WarehouseId", "StationName", "RoleDefault"
    ) -Rows @()
    Clear-ListObjectRows $loSt
    Add-ListObjectRow -ListObject $loSt -Values @{
        "StationId" = $StationId
        "WarehouseId" = $WarehouseId
        "StationName" = $env:COMPUTERNAME
        "RoleDefault" = "RECEIVE"
    }

    Save-NewWorkbook -Workbook $wb -Path $Path
    return $wb
}

function New-AuthWorkbook {
    param(
        [object]$Excel,
        [string]$Path,
        [string]$WarehouseId,
        [string]$StationId,
        [string[]]$CurrentUserIds
    )

    $resolvedUserIds = @($CurrentUserIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($resolvedUserIds.Count -eq 0) {
        $resolvedUserIds = @("user1")
    }

    $wb = $Excel.Workbooks.Add()
    $wsUsers = $wb.Worksheets.Item(1)
    $wsUsers.Name = "Users"
    $wsCaps = $wb.Worksheets.Add()
    $wsCaps.Name = "Capabilities"

    $loUsers = Add-Table -Worksheet $wsUsers -TableName "tblUsers" -Headers @(
        "UserId", "DisplayName", "PinHash", "Status", "ValidFrom", "ValidTo"
    ) -Rows @()
    Clear-ListObjectRows $loUsers
    foreach ($userId in $resolvedUserIds) {
        Add-ListObjectRow -ListObject $loUsers -Values @{
            "UserId" = $userId
            "DisplayName" = $userId
            "PinHash" = ""
            "Status" = "Active"
            "ValidFrom" = ""
            "ValidTo" = ""
        }
    }
    Add-ListObjectRow -ListObject $loUsers -Values @{
        "UserId" = "svc_processor"
        "DisplayName" = "Processor Service"
        "PinHash" = ""
        "Status" = "Active"
        "ValidFrom" = ""
        "ValidTo" = ""
    }

    $loCaps = Add-Table -Worksheet $wsCaps -TableName "tblCapabilities" -Headers @(
        "UserId", "Capability", "WarehouseId", "StationId", "Status", "ValidFrom", "ValidTo"
    ) -Rows @()
    Clear-ListObjectRows $loCaps
    foreach ($userId in $resolvedUserIds) {
        Add-ListObjectRow -ListObject $loCaps -Values @{
            "UserId" = $userId
            "Capability" = "RECEIVE_POST"
            "WarehouseId" = $WarehouseId
            "StationId" = $StationId
            "Status" = "ACTIVE"
            "ValidFrom" = ""
            "ValidTo" = ""
        }
        Add-ListObjectRow -ListObject $loCaps -Values @{
            "UserId" = $userId
            "Capability" = "SHIP_POST"
            "WarehouseId" = $WarehouseId
            "StationId" = $StationId
            "Status" = "ACTIVE"
            "ValidFrom" = ""
            "ValidTo" = ""
        }
        Add-ListObjectRow -ListObject $loCaps -Values @{
            "UserId" = $userId
            "Capability" = "PROD_POST"
            "WarehouseId" = $WarehouseId
            "StationId" = $StationId
            "Status" = "ACTIVE"
            "ValidFrom" = ""
            "ValidTo" = ""
        }
    }
    Add-ListObjectRow -ListObject $loCaps -Values @{
        "UserId" = "svc_processor"
        "Capability" = "INBOX_PROCESS"
        "WarehouseId" = $WarehouseId
        "StationId" = "*"
        "Status" = "ACTIVE"
        "ValidFrom" = ""
        "ValidTo" = ""
    }

    Save-NewWorkbook -Workbook $wb -Path $Path
    return $wb
}

function New-InventoryWorkbook {
    param(
        [object]$Excel,
        [string]$Path,
        [object[]]$SkuRows
    )

    $wb = $Excel.Workbooks.Add()
    $wsLog = $wb.Worksheets.Item(1)
    $wsLog.Name = "InventoryLog"
    Add-Table -Worksheet $wsLog -TableName "tblInventoryLog" -Headers @(
        "EventID", "UndoOfEventId", "AppliedSeq", "EventType", "OccurredAtUTC", "AppliedAtUTC",
        "WarehouseId", "StationId", "UserId", "SKU", "QtyDelta", "Location", "Note"
    ) -Rows @() | Out-Null
    Clear-ListObjectRows (Get-ListObjectSafe -Worksheet $wsLog -TableName "tblInventoryLog")

    $wsApplied = $wb.Worksheets.Add()
    $wsApplied.Name = "AppliedEvents"
    Add-Table -Worksheet $wsApplied -TableName "tblAppliedEvents" -Headers @(
        "EventID", "UndoOfEventId", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status"
    ) -Rows @() | Out-Null
    Clear-ListObjectRows (Get-ListObjectSafe -Worksheet $wsApplied -TableName "tblAppliedEvents")

    $wsLocks = $wb.Worksheets.Add()
    $wsLocks.Name = "Locks"
    Add-Table -Worksheet $wsLocks -TableName "tblLocks" -Headers @(
        "LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", "ExpiresAtUTC", "HeartbeatAtUTC", "Status"
    ) -Rows @(
        @("INVENTORY", "", "", "", "", "", "", "EXPIRED")
    ) | Out-Null

    $wsSku = $wb.Worksheets.Add()
    $wsSku.Name = "SkuCatalog"
    $rows = @()
    foreach ($sku in $SkuRows) {
        $rows += ,@([string]$sku)
    }
    Add-Table -Worksheet $wsSku -TableName "tblSkuCatalog" -Headers @("SKU") -Rows $rows | Out-Null

    Save-NewWorkbook -Workbook $wb -Path $Path
    return $wb
}

function New-InboxWorkbook {
    param(
        [object]$Excel,
        [string]$Path,
        [string]$SheetName,
        [string]$TableName
    )

    $wb = $Excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = $SheetName
    Add-Table -Worksheet $ws -TableName $TableName -Headers @(
        "EventID", "ParentEventId", "UndoOfEventId", "EventType", "CreatedAtUTC", "WarehouseId", "StationId",
        "UserId", "SKU", "Qty", "Location", "Note", "PayloadJson", "Status", "RetryCount", "ErrorCode", "ErrorMessage", "FailedAtUTC"
    ) -Rows @() | Out-Null
    Clear-ListObjectRows (Get-ListObjectSafe -Worksheet $ws -TableName $TableName)

    Save-NewWorkbook -Workbook $wb -Path $Path
    return $wb
}

function Build-PayloadContains {
    param(
        [object]$ListObject,
        [int]$RowIndex,
        [string]$ExpectedText
    )

    $payload = [string](Get-RowValueSafe -ListObject $ListObject -RowIndex $RowIndex -ColumnName "PayloadJson")
    return ($payload -like ("*" + $ExpectedText + "*"))
}

$repo = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repo $DeployRoot
$resultPath = Join-Path $repo "tests/unit/phase6_live_role_workflow_results.md"
$runtimeRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-phase6-live-" + [guid]::NewGuid().ToString("N"))
$warehouseId = "WH1"
$stationId = "S1"
$currentUserId = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "user1" } else { $env:USERNAME }

$configPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Config.xlsb")
$authPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Auth.xlsb")
$inventoryPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Data.Inventory.xlsb")
$receiveInboxPath = Join-Path $runtimeRoot ("invSys.Inbox.Receiving." + $stationId + ".xlsb")
$shipInboxPath = Join-Path $runtimeRoot ("invSys.Inbox.Shipping." + $stationId + ".xlsb")
$prodInboxPath = Join-Path $runtimeRoot ("invSys.Inbox.Production." + $stationId + ".xlsb")

$openOrder = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Receiving.xlam",
    "invSys.Shipping.xlam",
    "invSys.Production.xlam"
)

$resultRows = New-Object 'System.Collections.Generic.List[object]'
$excel = $null
$openedWorkbooks = New-Object 'System.Collections.Generic.List[object]'
$workbookMap = @{}
$currentStep = "Startup"

try {
    $currentStep = "Create runtime root"
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null

    $currentStep = "Start Excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1
    $authUserIds = @($currentUserId, [string]$excel.UserName, "user1") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    $currentStep = "Seed runtime workbooks"
    $runtimeBooks = @(
        (New-ConfigWorkbook -Excel $excel -Path $configPath -WarehouseId $warehouseId -StationId $stationId -RuntimeRoot $runtimeRoot),
        (New-AuthWorkbook -Excel $excel -Path $authPath -WarehouseId $warehouseId -StationId $stationId -CurrentUserIds $authUserIds),
        (New-InventoryWorkbook -Excel $excel -Path $inventoryPath -SkuRows @("SKU-REC", "SKU-SHIP", "SKU-FG")),
        (New-InboxWorkbook -Excel $excel -Path $receiveInboxPath -SheetName "InboxReceive" -TableName "tblInboxReceive"),
        (New-InboxWorkbook -Excel $excel -Path $shipInboxPath -SheetName "InboxShip" -TableName "tblInboxShip"),
        (New-InboxWorkbook -Excel $excel -Path $prodInboxPath -SheetName "InboxProd" -TableName "tblInboxProd")
    )
    foreach ($wb in $runtimeBooks) {
        $openedWorkbooks.Add($wb) | Out-Null
    }

    $currentStep = "Open packaged add-ins"
    foreach ($fileName in $openOrder) {
        $path = Join-Path $deployPath $fileName
        $wb = $excel.Workbooks.Open($path)
        $openedWorkbooks.Add($wb) | Out-Null
        $workbookMap[$fileName] = $wb
    }

    $currentStep = "Set core runtime override"
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($runtimeRoot))
    Add-ResultRow -Rows $resultRows -Check "Core.RuntimeRootOverride" -Passed $true -Detail $runtimeRoot

    $currentStep = "Capture core auth diagnostics"
    $resolvedUserId = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRoleEventWriter.ResolveCurrentUserId")
    $configLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modConfig.LoadConfig" -Arguments @($warehouseId, $stationId))
    $resolvedWarehouseId = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modConfig.GetWarehouseId")
    $resolvedStationId = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modConfig.GetStationId")
    $resolvedDataRoot = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modConfig.GetString" -Arguments @("PathDataRoot", ""))
    $authLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.LoadAuth" -Arguments @($warehouseId))
    $authReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.ValidateAuth")
    $receiveAllowed = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.CanPerform" -Arguments @("RECEIVE_POST", $resolvedUserId, $warehouseId, $stationId))
    $shipAllowed = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.CanPerform" -Arguments @("SHIP_POST", $resolvedUserId, $warehouseId, $stationId))
    $prodAllowed = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.CanPerform" -Arguments @("PROD_POST", $resolvedUserId, $warehouseId, $stationId))
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.User" -Passed (-not [string]::IsNullOrWhiteSpace($resolvedUserId)) -Detail ("ResolvedUser=" + $resolvedUserId + "; SeededUsers=" + (($authUserIds -join ",") + ",svc_processor"))
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.Config" -Passed $configLoaded -Detail ("WarehouseId=" + $resolvedWarehouseId + "; StationId=" + $resolvedStationId + "; PathDataRoot=" + $resolvedDataRoot)
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.AuthLoad" -Passed $authLoaded -Detail $authReport
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.ReceiveCapability" -Passed $receiveAllowed -Detail ("User=" + $resolvedUserId + "; WarehouseId=" + $warehouseId + "; StationId=" + $stationId)
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.ShipCapability" -Passed $shipAllowed -Detail ("User=" + $resolvedUserId + "; WarehouseId=" + $warehouseId + "; StationId=" + $stationId)
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.ProdCapability" -Passed $prodAllowed -Detail ("User=" + $resolvedUserId + "; WarehouseId=" + $warehouseId + "; StationId=" + $stationId)

    $currentStep = "Init role add-ins"
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Receiving.xlam"].Name -MacroName "modReceivingInit.InitReceivingAddin")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modShippingInit.InitShippingAddin")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Production.xlam"].Name -MacroName "modProductionInit.InitProductionAddin")

    $currentStep = "Stage Receiving workflow"
    $wbReceive = $workbookMap["invSys.Receiving.xlam"]
    $wsReceive = Get-WorksheetSafe -Workbook $wbReceive -WorksheetName "ReceivedTally"
    $wsReceiveLog = Get-WorksheetSafe -Workbook $wbReceive -WorksheetName "ReceivedLog"
    $wsReceiveInv = Get-WorksheetSafe -Workbook $wbReceive -WorksheetName "InventoryManagement"
    $loReceivedTally = Get-ListObjectSafe -Worksheet $wsReceive -TableName "ReceivedTally"
    $loAggReceived = Get-ListObjectSafe -Worksheet $wsReceive -TableName "AggregateReceived"
    $loReceiveLog = Get-ListObjectSafe -Worksheet $wsReceiveLog -TableName "ReceivedLog"
    $loReceiveInv = Get-ListObjectSafe -Worksheet $wsReceiveInv -TableName "invSys"
    $loInboxReceive = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $runtimeBooks[3] -WorksheetName "InboxReceive") -TableName "tblInboxReceive"
    $loInventoryLog = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $runtimeBooks[2] -WorksheetName "InventoryLog") -TableName "tblInventoryLog"

    Clear-ListObjectRows $loReceivedTally
    Clear-ListObjectRows $loAggReceived
    Clear-ListObjectRows $loReceiveLog
    Clear-ListObjectRows $loReceiveInv
    Add-ListObjectRow -ListObject $loReceiveInv -Values @{
        "ROW" = 101; "ITEM_CODE" = "SKU-REC"; "ITEM" = "Receive Widget"; "UOM" = "EA"; "LOCATION" = "A1";
        "DESCRIPTION" = "Receive Widget"; "RECEIVED" = 0; "TOTAL INV" = 10; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }
    Add-ListObjectRow -ListObject $loReceivedTally -Values @{
        "REF_NUMBER" = "REF-LIVE-001"; "ITEMS" = "Receive Widget"; "QUANTITY" = 7; "ROW" = 101
    }
    Add-ListObjectRow -ListObject $loAggReceived -Values @{
        "REF_NUMBER" = "REF-LIVE-001"; "ITEM_CODE" = "SKU-REC"; "VENDORS" = "Vendor A"; "VENDOR_CODE" = "V001";
        "DESCRIPTION" = "Receive Widget"; "ITEM" = "Receive Widget"; "UOM" = "EA"; "QUANTITY" = 7; "LOCATION" = "A1"; "ROW" = 101
    }
    $receiveQueueDiag = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $wbReceive.Name -MacroName "modTS_Received.ValidateQueueReceiveEventsFromCurrentWorkbook")
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.QueueDiagnostic" -Passed ([string]::Equals($receiveQueueDiag, "OK", [System.StringComparison]::OrdinalIgnoreCase)) -Detail $receiveQueueDiag
    $receiveInboxBefore = Get-RowCountSafe $loInboxReceive
    $inventoryLogBefore = Get-RowCountSafe $loInventoryLog
    $currentStep = "Run Receiving ConfirmWrites"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $wbReceive.Name -MacroName "modTS_Received.ConfirmWrites")

    $receiveLocalOk = ([double](Get-RowValueSafe -ListObject $loReceiveInv -RowIndex 1 -ColumnName "RECEIVED")) -eq 7 -and (Get-RowCountSafe $loReceiveLog) -eq 1
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.Local" -Passed $receiveLocalOk -Detail "RECEIVED=$((Get-RowValueSafe -ListObject $loReceiveInv -RowIndex 1 -ColumnName 'RECEIVED')); LogRows=$(Get-RowCountSafe $loReceiveLog)"

    $receiveInboxAfter = Get-RowCountSafe $loInboxReceive
    $receiveQueuedRow = Find-RowIndexByValue -ListObject $loInboxReceive -ColumnName "SKU" -ExpectedValue "SKU-REC"
    $receiveQueuedOk = ($receiveInboxAfter -eq ($receiveInboxBefore + 1)) -and ($receiveQueuedRow -gt 0) -and ([double](Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName "Qty") -eq 7)
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.Queue" -Passed $receiveQueuedOk -Detail "InboxRows=$receiveInboxAfter; Row=$receiveQueuedRow"

    $receiveRunBatchReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @($warehouseId, 500))
    $receiveRunBatch = 0
    if ($receiveRunBatchReport -match 'Processed=(\d+)') { $receiveRunBatch = [int]$Matches[1] }
    $receiveProcessedOk = ($receiveRunBatch -ge 1) -and ([string](Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName "Status") -eq "PROCESSED")
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.Process" -Passed $receiveProcessedOk -Detail "RunBatch=$receiveRunBatch; Status=$((Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName 'Status')); ErrorCode=$((Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName 'ErrorCode')); ErrorMessage=$((Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName 'ErrorMessage')); $receiveRunBatchReport"

    $receiveLogRow = Find-RowIndexByValue -ListObject $loInventoryLog -ColumnName "EventType" -ExpectedValue "RECEIVE"
    $receiveInventoryOk = ($receiveLogRow -gt 0) -and ([double](Get-RowValueSafe -ListObject $loInventoryLog -RowIndex $receiveLogRow -ColumnName "QtyDelta") -eq 7)
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.InventoryLog" -Passed $receiveInventoryOk -Detail "InventoryLogRowsBefore=$inventoryLogBefore; Row=$receiveLogRow"

    $currentStep = "Stage Shipping workflow"
    $wbShip = $workbookMap["invSys.Shipping.xlam"]
    $wsShip = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $wsShipInv = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "InventoryManagement"
    $loAggPackages = Get-ListObjectSafe -Worksheet $wsShip -TableName "AggregatePackages"
    $loShipInv = Get-ListObjectSafe -Worksheet $wsShipInv -TableName "invSys"
    $loInboxShip = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $runtimeBooks[4] -WorksheetName "InboxShip") -TableName "tblInboxShip"

    Clear-ListObjectRows $loAggPackages
    Clear-ListObjectRows $loShipInv
    Add-ListObjectRow -ListObject $loShipInv -Values @{
        "ROW" = 201; "ITEM_CODE" = "SKU-SHIP"; "ITEM" = "Ship Widget"; "UOM" = "EA"; "LOCATION" = "DOCK";
        "DESCRIPTION" = "Ship Widget"; "SHIPMENTS" = 5; "TOTAL INV" = 20; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }
    Add-ListObjectRow -ListObject $loAggPackages -Values @{
        "ROW" = 201; "ITEM_CODE" = "SKU-SHIP"; "ITEM" = "Ship Widget"; "QUANTITY" = 5; "UOM" = "EA"; "LOCATION" = "DOCK"
    }

    $shipQueueDiag = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $wbShip.Name -MacroName "modTS_Shipments.ValidateQueueShipmentsSentEventFromCurrentWorkbook")
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnShipmentsSent.QueueDiagnostic" -Passed ([string]::Equals($shipQueueDiag, "OK", [System.StringComparison]::OrdinalIgnoreCase)) -Detail $shipQueueDiag
    $shipInboxBefore = Get-RowCountSafe $loInboxShip
    $currentStep = "Run Shipping BtnShipmentsSent"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $wbShip.Name -MacroName "modTS_Shipments.BtnShipmentsSent")

    $shipLocalOk = ([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName "SHIPMENTS")) -eq 0
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnShipmentsSent.Local" -Passed $shipLocalOk -Detail "SHIPMENTS=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'SHIPMENTS'))"

    $shipInboxAfter = Get-RowCountSafe $loInboxShip
    $shipQueuedRow = 0
    for ($i = 1; $i -le $shipInboxAfter; $i++) {
        if (Build-PayloadContains -ListObject $loInboxShip -RowIndex $i -ExpectedText '"SKU":"SKU-SHIP"') {
            $shipQueuedRow = $i
            break
        }
    }
    $shipQueuedOk = ($shipInboxAfter -eq ($shipInboxBefore + 1)) -and ($shipQueuedRow -gt 0)
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnShipmentsSent.Queue" -Passed $shipQueuedOk -Detail "InboxRows=$shipInboxAfter; Row=$shipQueuedRow"

    $shipRunBatchReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @($warehouseId, 500))
    $shipRunBatch = 0
    if ($shipRunBatchReport -match 'Processed=(\d+)') { $shipRunBatch = [int]$Matches[1] }
    $shipProcessedOk = ($shipRunBatch -ge 1) -and ([string](Get-RowValueSafe -ListObject $loInboxShip -RowIndex $shipQueuedRow -ColumnName "Status") -eq "PROCESSED")
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnShipmentsSent.Process" -Passed $shipProcessedOk -Detail "RunBatch=$shipRunBatch; Status=$((Get-RowValueSafe -ListObject $loInboxShip -RowIndex $shipQueuedRow -ColumnName 'Status')); ErrorCode=$((Get-RowValueSafe -ListObject $loInboxShip -RowIndex $shipQueuedRow -ColumnName 'ErrorCode')); ErrorMessage=$((Get-RowValueSafe -ListObject $loInboxShip -RowIndex $shipQueuedRow -ColumnName 'ErrorMessage')); $shipRunBatchReport"

    $shipLogRow = Find-RowIndexByValue -ListObject $loInventoryLog -ColumnName "EventType" -ExpectedValue "SHIP"
    $shipInventoryOk = ($shipLogRow -gt 0) -and ([double](Get-RowValueSafe -ListObject $loInventoryLog -RowIndex $shipLogRow -ColumnName "QtyDelta") -eq -5)
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnShipmentsSent.InventoryLog" -Passed $shipInventoryOk -Detail "InventoryLogRow=$shipLogRow"

    $currentStep = "Stage Production workflow"
    $wbProd = $workbookMap["invSys.Production.xlam"]
    $wsProd = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "Production"
    $wsProdRecipes = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "Recipes"
    $wsProdInv = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "InventoryManagement"
    $wsPalette = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "IngredientPalette"
    $loChooseRecipe = Get-ListObjectSafe -Worksheet $wsProd -TableName "IP_ChooseRecipe"
    $loChooseIngredient = Get-ListObjectSafe -Worksheet $wsProd -TableName "IP_ChooseIngredient"
    $loChooseItem = Get-ListObjectSafe -Worksheet $wsProd -TableName "IP_ChooseItem"
    $loRecipes = Get-ListObjectSafe -Worksheet $wsProdRecipes -TableName "Recipes"
    $loPalette = Get-ListObjectSafe -Worksheet $wsPalette -TableName "IngredientPalette"
    $loProdInv = Get-ListObjectSafe -Worksheet $wsProdInv -TableName "invSys"
    $loProductionOutput = Get-ListObjectSafe -Worksheet $wsProd -TableName "ProductionOutput"
    $loInboxProd = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $runtimeBooks[5] -WorksheetName "InboxProd") -TableName "tblInboxProd"

    Clear-ListObjectRows $loChooseRecipe
    Clear-ListObjectRows $loChooseIngredient
    Clear-ListObjectRows $loChooseItem
    Clear-ListObjectRows $loRecipes
    Clear-ListObjectRows $loPalette
    Clear-ListObjectRows $loProdInv
    Clear-ListObjectRows $loProductionOutput

    Add-ListObjectRow -ListObject $loRecipes -Values @{
        "RECIPE" = "Blend A"; "RECIPE_ID" = "R-001"; "DESCRIPTION" = "Blend A"; "DEPARTMENT" = "Kitchen"; "PROCESS" = "Mix";
        "DIAGRAM_ID" = "D-1"; "INPUT/OUTPUT" = "INPUT"; "INGREDIENT" = "Sugar"; "PERCENT" = 50; "UOM" = "LB";
        "AMOUNT" = 2; "RECIPE_LIST_ROW" = 1; "INGREDIENT_ID" = "ING-001"; "GUID" = "REC-ING-1"
    }
    Add-ListObjectRow -ListObject $loChooseRecipe -Values @{
        "RECIPE_NAME" = "Blend A"; "DESCRIPTION" = "Blend A"; "GUID" = "PAL-REC-1"; "RECIPE_ID" = "R-001"
    }
    Add-ListObjectRow -ListObject $loChooseIngredient -Values @{
        "INGREDIENT" = "Sugar"; "UOM" = "LB"; "QUANTITY" = 2; "DESCRIPTION" = "INPUT"; "GUID" = "PAL-ING-1";
        "RECIPE_ID" = "R-001"; "INGREDIENT_ID" = "ING-001"; "PROCESS" = "Mix"
    }
    Add-ListObjectRow -ListObject $loChooseItem -Values @{
        "ITEMS" = "Sugar Bin"; "UOM" = "LB"; "DESCRIPTION" = "Granulated"; "ROW" = 301; "RECIPE_ID" = "R-001"; "INGREDIENT_ID" = "ING-001"
    }
    Add-ListObjectRow -ListObject $loProdInv -Values @{
        "ROW" = 301; "ITEM_CODE" = "SKU-SUGAR"; "ITEM" = "Sugar Bin"; "UOM" = "LB"; "LOCATION" = "BIN-A";
        "DESCRIPTION" = "Granulated"; "TOTAL INV" = 100; "MADE" = 0; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }

    $currentStep = "Run Production BtnSavePalette"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $wbProd.Name -MacroName "mProduction.BtnSavePalette")
    $paletteRow = Find-RowIndexByValue -ListObject $loPalette -ColumnName "RECIPE_ID" -ExpectedValue "R-001"
    $paletteOk = ($paletteRow -gt 0) -and ([string](Get-RowValueSafe -ListObject $loPalette -RowIndex $paletteRow -ColumnName "INGREDIENT_ID") -eq "ING-001") -and ([string](Get-RowValueSafe -ListObject $loPalette -RowIndex $paletteRow -ColumnName "ITEM") -eq "Sugar Bin")
    Add-ResultRow -Rows $resultRows -Check "Production.BtnSavePalette" -Passed $paletteOk -Detail "PaletteRow=$paletteRow"

    Add-ListObjectRow -ListObject $loProdInv -Values @{
        "ROW" = 401; "ITEM_CODE" = "SKU-FG"; "ITEM" = "Finished Good"; "UOM" = "EA"; "LOCATION" = "FG";
        "DESCRIPTION" = "Finished Good"; "MADE" = 8; "TOTAL INV" = 0; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }
    Add-ListObjectRow -ListObject $loProductionOutput -Values @{
        "PROCESS" = "Mix"; "OUTPUT" = "Finished Good"; "UOM" = "EA"; "REAL OUTPUT" = 8; "BATCH" = "B-001"; "RECALL CODE" = "RC-001"; "ROW" = 401
    }

    $prodQueueDiag = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $wbProd.Name -MacroName "mProduction.ValidateQueueProductionCompleteEventFromCurrentWorkbook")
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToTotalInv.QueueDiagnostic" -Passed ([string]::Equals($prodQueueDiag, "OK", [System.StringComparison]::OrdinalIgnoreCase)) -Detail $prodQueueDiag
    $prodInboxBefore = Get-RowCountSafe $loInboxProd
    $currentStep = "Run Production BtnToTotalInv"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $wbProd.Name -MacroName "mProduction.BtnToTotalInv")

    $prodLocalOk = ([double](Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName "MADE")) -eq 0 -and ([double](Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName "TOTAL INV")) -eq 8
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToTotalInv.Local" -Passed $prodLocalOk -Detail "MADE=$((Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName 'MADE')); TOTAL_INV=$((Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName 'TOTAL INV'))"

    $prodInboxAfter = Get-RowCountSafe $loInboxProd
    $prodQueuedRow = 0
    for ($i = 1; $i -le $prodInboxAfter; $i++) {
        if (([string](Get-RowValueSafe -ListObject $loInboxProd -RowIndex $i -ColumnName "EventType") -eq "PROD_COMPLETE") -and (Build-PayloadContains -ListObject $loInboxProd -RowIndex $i -ExpectedText '"SKU":"SKU-FG"')) {
            $prodQueuedRow = $i
            break
        }
    }
    $prodQueuedOk = ($prodInboxAfter -eq ($prodInboxBefore + 1)) -and ($prodQueuedRow -gt 0)
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToTotalInv.Queue" -Passed $prodQueuedOk -Detail "InboxRows=$prodInboxAfter; Row=$prodQueuedRow"

    $prodRunBatchReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @($warehouseId, 500))
    $prodRunBatch = 0
    if ($prodRunBatchReport -match 'Processed=(\d+)') { $prodRunBatch = [int]$Matches[1] }
    $prodProcessedOk = ($prodRunBatch -ge 1) -and ([string](Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodQueuedRow -ColumnName "Status") -eq "PROCESSED")
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToTotalInv.Process" -Passed $prodProcessedOk -Detail "RunBatch=$prodRunBatch; Status=$((Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodQueuedRow -ColumnName 'Status')); ErrorCode=$((Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodQueuedRow -ColumnName 'ErrorCode')); ErrorMessage=$((Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodQueuedRow -ColumnName 'ErrorMessage')); $prodRunBatchReport"

    $prodLogRow = Find-RowIndexByValue -ListObject $loInventoryLog -ColumnName "EventType" -ExpectedValue "PROD_COMPLETE"
    $prodInventoryOk = ($prodLogRow -gt 0) -and ([double](Get-RowValueSafe -ListObject $loInventoryLog -RowIndex $prodLogRow -ColumnName "QtyDelta") -eq 8)
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToTotalInv.InventoryLog" -Passed $prodInventoryOk -Detail "InventoryLogRow=$prodLogRow"
}
catch {
    Add-ResultRow -Rows $resultRows -Check "Harness.Exception" -Passed $false -Detail ("Step=" + $currentStep + "; " + $_.Exception.Message)
    throw
}
finally {
    $failedCount = @($resultRows | Where-Object { -not $_.Passed }).Count
    $passedCount = $resultRows.Count - $failedCount

    $lines = @()
    $lines += "# Phase 6 Live Role Workflow Validation Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Deploy root: $deployPath"
    $lines += "- Runtime root override: $runtimeRoot"
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
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}

$failed = @($resultRows | Where-Object { -not $_.Passed }).Count
if ($failed -gt 0) {
    Write-Output "PHASE6_LIVE_ROLE_VALIDATION_FAILED"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$($resultRows.Count - $failed) FAILED=$failed TOTAL=$($resultRows.Count)"
    exit 1
}

Write-Output "PHASE6_LIVE_ROLE_VALIDATION_OK"
Write-Output "RESULTS=$resultPath"
Write-Output "PASSED=$($resultRows.Count) FAILED=0 TOTAL=$($resultRows.Count)"
