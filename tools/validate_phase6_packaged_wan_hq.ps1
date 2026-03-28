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
        6 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5]) }
        7 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6]) }
        8 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7]) }
        9 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7], $Arguments[8]) }
        10 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7], $Arguments[8], $Arguments[9]) }
        11 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7], $Arguments[8], $Arguments[9], $Arguments[10]) }
        12 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7], $Arguments[8], $Arguments[9], $Arguments[10], $Arguments[11]) }
        13 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7], $Arguments[8], $Arguments[9], $Arguments[10], $Arguments[11], $Arguments[12]) }
        default { throw "Run-WorkbookMacro supports at most 13 arguments." }
    }
}

function Resolve-WorkbookSafe {
    param(
        [object]$Excel,
        [string]$WorkbookName
    )

    if ([string]::IsNullOrWhiteSpace($WorkbookName)) { return $null }
    try { return $Excel.Workbooks.Item($WorkbookName) } catch {}
    foreach ($wb in $Excel.Workbooks) {
        try {
            if ([string]::Equals([string]$wb.Name, $WorkbookName, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $wb
            }
        }
        catch {}
    }
    return $null
}

function Get-WorksheetSafe {
    param(
        [object]$Workbook,
        [string]$WorksheetName
    )

    try { return $Workbook.Worksheets.Item($WorksheetName) } catch { return $null }
}

function Get-ListObjectSafe {
    param(
        [object]$Worksheet,
        [string]$TableName
    )

    if ($null -eq $Worksheet) { return $null }
    try { return $Worksheet.ListObjects.Item($TableName) } catch { return $null }
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
        try { $targetRange = $ListObject.InsertRowRange } catch { $targetRange = $null }
    }

    if ($null -eq $targetRange) {
        try { $row = $ListObject.ListRows.Add($null, $false) } catch { $row = $ListObject.ListRows.Add() }
        $targetRange = $row.Range
    }

    foreach ($key in $Values.Keys) {
        $idx = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName ([string]$key)
        if ($idx -le 0) { continue }
        $value = $Values[$key]
        if ($null -eq $value) {
            $targetRange.Cells.Item(1, [int]$idx).Value2 = $null
        }
        elseif ($value -is [datetime]) {
            $targetRange.Cells.Item(1, [int]$idx).Value2 = $value.ToOADate()
        }
        elseif ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal] -or $value -is [single]) {
            $targetRange.Cells.Item(1, [int]$idx).Value2 = [double]$value
        }
        else {
            $targetRange.Cells.Item(1, [int]$idx).Value2 = [string]$value
        }
    }
}

function Get-RowValueSafe {
    param(
        [object]$ListObject,
        [int]$RowIndex,
        [string]$ColumnName
    )

    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange -or $RowIndex -le 0) { return $null }
    $idx = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName $ColumnName
    if ($idx -le 0) { return $null }
    return $ListObject.DataBodyRange.Cells.Item([int]$RowIndex, [int]$idx).Value2
}

function Find-RowIndexByWarehouseSku {
    param(
        [object]$ListObject,
        [string]$WarehouseId,
        [string]$Sku
    )

    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) { return 0 }
    $whIdx = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName "WarehouseId"
    $skuIdx = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName "SKU"
    if ($whIdx -le 0 -or $skuIdx -le 0) { return 0 }

    for ($i = 1; $i -le $ListObject.ListRows.Count; $i++) {
        $wh = [string]$ListObject.DataBodyRange.Cells.Item([int]$i, [int]$whIdx).Value2
        $skuVal = [string]$ListObject.DataBodyRange.Cells.Item([int]$i, [int]$skuIdx).Value2
        if ($wh -eq $WarehouseId -and $skuVal -eq $Sku) {
            return $i
        }
    }
    return 0
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
                elseif ($value -is [datetime]) {
                    $Worksheet.Cells($r + 2, $c + 1).Value2 = $value.ToOADate()
                }
                elseif ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal] -or $value -is [single]) {
                    $Worksheet.Cells($r + 2, $c + 1).Value2 = [double]$value
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

    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $Workbook.SaveAs($Path, 50)
}

function Ensure-Folder {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Add-BootstrapModule {
    param([object]$Workbook)
    $comp = $Workbook.VBProject.VBComponents.Add(1)
    $comp.Name = "modHarnessBootstrap"
    $comp.CodeModule.AddFromString(@"
Public Function QueueReceiveHarness(ByVal warehouseId As String, _
                                    ByVal stationId As String, _
                                    ByVal sku As String, _
                                    ByVal qty As Double, _
                                    ByVal locationVal As String, _
                                    ByVal noteVal As String) As String
    Dim eventId As String
    Dim errorMessage As String
    Dim ok As Boolean

    ok = Application.Run("'invSys.Core.xlam'!modRoleEventWriter.QueueReceiveEvent", _
                         warehouseId, stationId, "user1", sku, qty, locationVal, noteVal, "", "", 0, Nothing, eventId, errorMessage)
    QueueReceiveHarness = CStr(Abs(CLng(ok))) & "|" & eventId & "|" & errorMessage
End Function
"@)
    return $comp
}

function New-HelperWorkbook {
    param(
        [object]$Excel,
        [string]$Path
    )

    $wb = $Excel.Workbooks.Add()
    [void](Add-BootstrapModule -Workbook $wb)
    $wb.SaveAs($Path, 52)
    return $wb
}

function New-ConfigWorkbook {
    param(
        [object]$Excel,
        [string]$Path,
        [string]$WarehouseId,
        [string]$StationId,
        [string]$RuntimeRoot,
        [string]$ShareRoot
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
        "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot", "AutoRefreshIntervalSeconds"
    ) -Rows @()
    Clear-ListObjectRows $loWh
    Add-ListObjectRow -ListObject $loWh -Values @{
        "WarehouseId" = $WarehouseId
        "WarehouseName" = $WarehouseId
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
        "PathSharePointRoot" = $ShareRoot
        "DesignsEnabled" = $false
        "PoisonRetryMax" = 3
        "AuthCacheTTLSeconds" = 300
        "ProcessorServiceUserId" = "svc_processor"
        "FF_DesignsEnabled" = $false
        "FF_OutlookAlerts" = $false
        "FF_AutoSnapshot" = $true
        "AutoRefreshIntervalSeconds" = 0
    }

    $loSt = Add-Table -Worksheet $wsSt -TableName "tblStationConfig" -Headers @(
        "StationId", "WarehouseId", "StationName", "RoleDefault"
    ) -Rows @()
    Clear-ListObjectRows $loSt
    Add-ListObjectRow -ListObject $loSt -Values @{
        "StationId" = $StationId
        "WarehouseId" = $WarehouseId
        "StationName" = $StationId
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
        [string]$StationId
    )

    $wb = $Excel.Workbooks.Add()
    $wsUsers = $wb.Worksheets.Item(1)
    $wsUsers.Name = "Users"
    $wsCaps = $wb.Worksheets.Add()
    $wsCaps.Name = "Capabilities"

    $loUsers = Add-Table -Worksheet $wsUsers -TableName "tblUsers" -Headers @(
        "UserId", "DisplayName", "PinHash", "Status", "ValidFrom", "ValidTo"
    ) -Rows @()
    Clear-ListObjectRows $loUsers
    Add-ListObjectRow -ListObject $loUsers -Values @{ "UserId" = "user1"; "DisplayName" = "user1"; "PinHash" = ""; "Status" = "Active"; "ValidFrom" = ""; "ValidTo" = "" }
    Add-ListObjectRow -ListObject $loUsers -Values @{ "UserId" = "svc_processor"; "DisplayName" = "Processor Service"; "PinHash" = ""; "Status" = "Active"; "ValidFrom" = ""; "ValidTo" = "" }

    $loCaps = Add-Table -Worksheet $wsCaps -TableName "tblCapabilities" -Headers @(
        "UserId", "Capability", "WarehouseId", "StationId", "Status", "ValidFrom", "ValidTo"
    ) -Rows @()
    Clear-ListObjectRows $loCaps
    Add-ListObjectRow -ListObject $loCaps -Values @{ "UserId" = "user1"; "Capability" = "RECEIVE_POST"; "WarehouseId" = $WarehouseId; "StationId" = $StationId; "Status" = "ACTIVE"; "ValidFrom" = ""; "ValidTo" = "" }
    Add-ListObjectRow -ListObject $loCaps -Values @{ "UserId" = "svc_processor"; "Capability" = "INBOX_PROCESS"; "WarehouseId" = $WarehouseId; "StationId" = "*"; "Status" = "ACTIVE"; "ValidFrom" = ""; "ValidTo" = "" }

    Save-NewWorkbook -Workbook $wb -Path $Path
    return $wb
}

function New-InventoryWorkbook {
    param(
        [object]$Excel,
        [string]$Path,
        [string]$Sku
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
    Add-Table -Worksheet $wsSku -TableName "tblSkuCatalog" -Headers @("SKU") -Rows @(@($Sku)) | Out-Null

    Save-NewWorkbook -Workbook $wb -Path $Path
    return $wb
}

function New-InboxWorkbook {
    param(
        [object]$Excel,
        [string]$Path
    )

    $wb = $Excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "InboxReceive"
    Add-Table -Worksheet $ws -TableName "tblInboxReceive" -Headers @(
        "EventID", "ParentEventId", "UndoOfEventId", "EventType", "CreatedAtUTC", "WarehouseId", "StationId",
        "UserId", "SKU", "Qty", "Location", "Note", "PayloadJson", "Status", "RetryCount", "ErrorCode", "ErrorMessage", "FailedAtUTC"
    ) -Rows @() | Out-Null
    Clear-ListObjectRows (Get-ListObjectSafe -Worksheet $ws -TableName "tblInboxReceive")

    Save-NewWorkbook -Workbook $wb -Path $Path
    return $wb
}

function Open-PackagedAddins {
    param(
        [object]$Excel,
        [string]$DeployPath
    )

    $map = @{}
    foreach ($fileName in @("invSys.Core.xlam", "invSys.Inventory.Domain.xlam")) {
        $path = Join-Path $DeployPath $fileName
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Missing packaged XLAM: $path"
        }
        $map[$fileName] = $Excel.Workbooks.Open($path)
    }
    return $map
}

function Close-WorkbookByName {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [bool]$SaveChanges = $false
    )

    $wb = Resolve-WorkbookSafe -Excel $Excel -WorkbookName $WorkbookName
    if ($null -ne $wb) {
        try { $wb.Close($SaveChanges) } catch {}
    }
}

function Copy-LocalSnapshotToShare {
    param(
        [object]$Excel,
        [string]$RuntimeRoot,
        [string]$WarehouseId,
        [string]$ShareRoot
    )

    $snapshotName = "$WarehouseId.invSys.Snapshot.Inventory.xlsb"
    Close-WorkbookByName -Excel $Excel -WorkbookName $snapshotName -SaveChanges $true
    $sourcePath = Join-Path $RuntimeRoot $snapshotName
    $targetPath = Join-Path (Join-Path $ShareRoot "Snapshots") $snapshotName
    if (Test-Path -LiteralPath $targetPath) { Remove-Item -LiteralPath $targetPath -Force }
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    return $targetPath
}

function Seed-InboxReceiveRowOpen {
    param(
        [object]$Excel,
        [string]$InboxPath,
        [string]$WarehouseId,
        [string]$StationId,
        [string]$Sku,
        [double]$Qty,
        [string]$Location,
        [string]$Note
    )

    $wb = $Excel.Workbooks.Open($InboxPath)
    $ws = Get-WorksheetSafe -Workbook $wb -WorksheetName "InboxReceive"
    if ($null -ne $ws) {
        try { $ws.Unprotect() } catch {}
    }
    $lo = Get-ListObjectSafe -Worksheet $ws -TableName "tblInboxReceive"
    $eventId = "EVT-$WarehouseId-" + (Get-Date -Format "yyyyMMddHHmmssfff")
    Add-ListObjectRow -ListObject $lo -Values @{
        "EventID" = $eventId
        "ParentEventId" = ""
        "UndoOfEventId" = ""
        "EventType" = "RECEIVE"
        "CreatedAtUTC" = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        "WarehouseId" = $WarehouseId
        "StationId" = $StationId
        "UserId" = "user1"
        "SKU" = $Sku
        "Qty" = $Qty
        "Location" = $Location
        "Note" = $Note
        "PayloadJson" = ""
        "Status" = "NEW"
        "RetryCount" = 0
        "ErrorCode" = ""
        "ErrorMessage" = ""
        "FailedAtUTC" = ""
    }
    $wb.Save()
    if ($null -ne $ws) {
        try { $ws.Protect($null, $true, $true) } catch {}
    }
    return $eventId
}

function Write-Results {
    param(
        [string]$ResultPath,
        [System.Collections.Generic.List[object]]$Rows,
        [string]$DeployPath,
        [string]$SessionRoot
    )

    $passedCount = @($Rows | Where-Object { $_.Passed }).Count
    $failedCount = $Rows.Count - $passedCount
    $lines = @()
    $lines += "# Phase 6 Packaged WAN HQ Validation Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Deploy root: $DeployPath"
    $lines += "- Session root: $SessionRoot"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += ""
    $lines += "| Check | Result | Detail |"
    $lines += "|---|---|---|"
    foreach ($row in $Rows) {
        $result = if ($row.Passed) { "PASS" } else { "FAIL" }
        $detail = ([string]$row.Detail).Replace("|", "/")
        $lines += "| $($row.Check) | $result | $detail |"
    }
    [System.IO.File]::WriteAllLines($ResultPath, $lines)
    return [pscustomobject]@{ Passed = $passedCount; Failed = $failedCount; Total = $Rows.Count }
}

$repo = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repo $DeployRoot
$resultPath = Join-Path $repo "tests/unit/phase6_packaged_wan_hq_results.md"
$sessionRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-phase6-wanhq-" + [guid]::NewGuid().ToString("N"))
$rootA = Join-Path $sessionRoot "WH97"
$rootB = Join-Path $sessionRoot "WH98"
$shareRoot = Join-Path $sessionRoot "Share"
$sku = "SKU-WAN-HQ-001"

$excelSetup = $null
$excelA = $null
$excelB = $null
$excelHq = $null
$helperA = $null
$helperB = $null
$workbookSets = @()
$resultRows = New-Object 'System.Collections.Generic.List[object]'

try {
    Ensure-Folder $sessionRoot
    Ensure-Folder $rootA
    Ensure-Folder $rootB
    Ensure-Folder $shareRoot
    Ensure-Folder (Join-Path $shareRoot "Snapshots")
    Ensure-Folder (Join-Path $shareRoot "Global")

    $excelSetup = New-Object -ComObject Excel.Application
    $excelSetup.Visible = $false
    $excelSetup.DisplayAlerts = $false
    $excelSetup.EnableEvents = $false

    $cfgA = New-ConfigWorkbook -Excel $excelSetup -Path (Join-Path $rootA "WH97.invSys.Config.xlsb") -WarehouseId "WH97" -StationId "S1" -RuntimeRoot $rootA -ShareRoot $shareRoot
    $authA = New-AuthWorkbook -Excel $excelSetup -Path (Join-Path $rootA "WH97.invSys.Auth.xlsb") -WarehouseId "WH97" -StationId "S1"
    $invA = New-InventoryWorkbook -Excel $excelSetup -Path (Join-Path $rootA "WH97.invSys.Data.Inventory.xlsb") -Sku $sku
    $inboxA = New-InboxWorkbook -Excel $excelSetup -Path (Join-Path $rootA "invSys.Inbox.Receiving.S1.xlsb")
    $cfgB = New-ConfigWorkbook -Excel $excelSetup -Path (Join-Path $rootB "WH98.invSys.Config.xlsb") -WarehouseId "WH98" -StationId "S2" -RuntimeRoot $rootB -ShareRoot $shareRoot
    $authB = New-AuthWorkbook -Excel $excelSetup -Path (Join-Path $rootB "WH98.invSys.Auth.xlsb") -WarehouseId "WH98" -StationId "S2"
    $invB = New-InventoryWorkbook -Excel $excelSetup -Path (Join-Path $rootB "WH98.invSys.Data.Inventory.xlsb") -Sku $sku
    $inboxB = New-InboxWorkbook -Excel $excelSetup -Path (Join-Path $rootB "invSys.Inbox.Receiving.S2.xlsb")
    foreach ($wb in @($cfgA,$authA,$invA,$inboxA,$cfgB,$authB,$invB,$inboxB)) { try { $wb.Close($false) } catch {} }
    Add-ResultRow -Rows $resultRows -Check "Setup.RuntimeRoots" -Passed $true -Detail ("SessionRoot=" + $sessionRoot)

    $excelSetup.Quit()
    Release-ComObject $excelSetup
    $excelSetup = $null

    foreach ($excelVarName in @("excelA","excelB","excelHq")) {
        $excelObj = New-Object -ComObject Excel.Application
        $excelObj.Visible = $false
        $excelObj.DisplayAlerts = $false
        $excelObj.EnableEvents = $false
        Set-Variable -Name $excelVarName -Value $excelObj
    }

    $wbSetA = Open-PackagedAddins -Excel $excelA -DeployPath $deployPath
    $wbSetB = Open-PackagedAddins -Excel $excelB -DeployPath $deployPath
    $wbSetHq = Open-PackagedAddins -Excel $excelHq -DeployPath $deployPath
    $helperA = New-HelperWorkbook -Excel $excelA -Path (Join-Path $sessionRoot "PackagedWanHelperA.xlsm")
    $helperB = New-HelperWorkbook -Excel $excelB -Path (Join-Path $sessionRoot "PackagedWanHelperB.xlsm")
    $workbookSets += $wbSetA, $wbSetB, $wbSetHq
    Add-ResultRow -Rows $resultRows -Check "Packaged.OpenA" -Passed $true -Detail "Core+Inventory.Domain"
    Add-ResultRow -Rows $resultRows -Check "Packaged.OpenB" -Passed $true -Detail "Core+Inventory.Domain"
    Add-ResultRow -Rows $resultRows -Check "Packaged.OpenHQ" -Passed $true -Detail "Core+Inventory.Domain"

    [void](Run-WorkbookMacro -Excel $excelA -WorkbookName $wbSetA["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($rootA))
    [void](Run-WorkbookMacro -Excel $excelB -WorkbookName $wbSetB["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($rootB))
    Add-ResultRow -Rows $resultRows -Check "Packaged.RuntimeOverrides" -Passed $true -Detail ("WH97=" + $rootA + "; WH98=" + $rootB)

    $queueA1 = Seed-InboxReceiveRowOpen -Excel $excelA -InboxPath (Join-Path $rootA "invSys.Inbox.Receiving.S1.xlsb") -WarehouseId "WH97" -StationId "S1" -Sku $sku -Qty 5 -Location "A1" -Note "wan-hq-wh97-initial"
    $queueB1 = Seed-InboxReceiveRowOpen -Excel $excelB -InboxPath (Join-Path $rootB "invSys.Inbox.Receiving.S2.xlsb") -WarehouseId "WH98" -StationId "S2" -Sku $sku -Qty 8 -Location "B1" -Note "wan-hq-wh98-initial"
    $reportA1 = [string](Run-WorkbookMacro -Excel $excelA -WorkbookName $wbSetA["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @("WH97", 500))
    $reportB1 = [string](Run-WorkbookMacro -Excel $excelB -WorkbookName $wbSetB["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @("WH98", 500))
    $publishA1 = Copy-LocalSnapshotToShare -Excel $excelA -RuntimeRoot $rootA -WarehouseId "WH97" -ShareRoot $shareRoot
    $publishB1 = Copy-LocalSnapshotToShare -Excel $excelB -RuntimeRoot $rootB -WarehouseId "WH98" -ShareRoot $shareRoot
    Add-ResultRow -Rows $resultRows -Check "Publish.WH97.Initial" -Passed (($queueA1 -ne "") -and ($reportA1 -match "Processed=1")) -Detail ("EventID=" + $queueA1 + "; " + $reportA1 + "; " + $publishA1)
    Add-ResultRow -Rows $resultRows -Check "Publish.WH98.Initial" -Passed (($queueB1 -ne "") -and ($reportB1 -match "Processed=1")) -Detail ("EventID=" + $queueB1 + "; " + $reportB1 + "; " + $publishB1)

    $agg1Ok = [bool](Run-WorkbookMacro -Excel $excelHq -WorkbookName $wbSetHq["invSys.Core.xlam"].Name -MacroName "modHqAggregator.RunHQAggregation" -Arguments @($shareRoot, "", ""))
    $wbGlobal1 = $excelHq.Workbooks.Open((Join-Path $shareRoot "Global\\invSys.Global.InventorySnapshot.xlsb"))
    $loGlobal1 = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbGlobal1 -WorksheetName "GlobalInventorySnapshot") -TableName "tblGlobalInventorySnapshot"
    $loStatus1 = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbGlobal1 -WorksheetName "GlobalSnapshotStatus") -TableName "tblGlobalSnapshotStatus"
    $rowA1 = Find-RowIndexByWarehouseSku -ListObject $loGlobal1 -WarehouseId "WH97" -Sku $sku
    $rowB1 = Find-RowIndexByWarehouseSku -ListObject $loGlobal1 -WarehouseId "WH98" -Sku $sku
    $agg1Pass = $agg1Ok -and $rowA1 -gt 0 -and $rowB1 -gt 0 -and ([double](Get-RowValueSafe -ListObject $loGlobal1 -RowIndex $rowA1 -ColumnName "QtyOnHand")) -eq 5 -and ([double](Get-RowValueSafe -ListObject $loGlobal1 -RowIndex $rowB1 -ColumnName "QtyOnHand")) -eq 8 -and ([int](Get-RowValueSafe -ListObject $loStatus1 -RowIndex 1 -ColumnName "WarehouseCount")) -eq 2
    Add-ResultRow -Rows $resultRows -Check "Aggregate.Initial" -Passed $agg1Pass -Detail ("QtyA=" + (Get-RowValueSafe -ListObject $loGlobal1 -RowIndex $rowA1 -ColumnName "QtyOnHand") + "; QtyB=" + (Get-RowValueSafe -ListObject $loGlobal1 -RowIndex $rowB1 -ColumnName "QtyOnHand"))
    try { $wbGlobal1.Close($false) } catch {}

    Start-Sleep -Milliseconds 1100
    $queueB2 = Seed-InboxReceiveRowOpen -Excel $excelB -InboxPath (Join-Path $rootB "invSys.Inbox.Receiving.S2.xlsb") -WarehouseId "WH98" -StationId "S2" -Sku $sku -Qty 3 -Location "B1" -Note "wan-hq-wh98-catchup"
    $reportB2 = [string](Run-WorkbookMacro -Excel $excelB -WorkbookName $wbSetB["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @("WH98", 500))
    $publishB2 = Copy-LocalSnapshotToShare -Excel $excelB -RuntimeRoot $rootB -WarehouseId "WH98" -ShareRoot $shareRoot
    Add-ResultRow -Rows $resultRows -Check "Publish.WH98.Catchup" -Passed (($queueB2 -ne "") -and ($reportB2 -match "Processed=1")) -Detail ("EventID=" + $queueB2 + "; " + $reportB2 + "; " + $publishB2)

    $agg2Ok = [bool](Run-WorkbookMacro -Excel $excelHq -WorkbookName $wbSetHq["invSys.Core.xlam"].Name -MacroName "modHqAggregator.RunHQAggregation" -Arguments @($shareRoot, "", ""))
    $wbGlobal2 = $excelHq.Workbooks.Open((Join-Path $shareRoot "Global\\invSys.Global.InventorySnapshot.xlsb"))
    $loGlobal2 = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbGlobal2 -WorksheetName "GlobalInventorySnapshot") -TableName "tblGlobalInventorySnapshot"
    $loStatus2 = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbGlobal2 -WorksheetName "GlobalSnapshotStatus") -TableName "tblGlobalSnapshotStatus"
    $rowA2 = Find-RowIndexByWarehouseSku -ListObject $loGlobal2 -WarehouseId "WH97" -Sku $sku
    $rowB2 = Find-RowIndexByWarehouseSku -ListObject $loGlobal2 -WarehouseId "WH98" -Sku $sku
    $agg2Pass = $agg2Ok -and $rowA2 -gt 0 -and $rowB2 -gt 0 -and ([double](Get-RowValueSafe -ListObject $loGlobal2 -RowIndex $rowA2 -ColumnName "QtyOnHand")) -eq 5 -and ([double](Get-RowValueSafe -ListObject $loGlobal2 -RowIndex $rowB2 -ColumnName "QtyOnHand")) -eq 11 -and ([int](Get-RowValueSafe -ListObject $loStatus2 -RowIndex 1 -ColumnName "SkippedSnapshotFileCount")) -eq 0
    Add-ResultRow -Rows $resultRows -Check "Aggregate.Catchup" -Passed $agg2Pass -Detail ("QtyA=" + (Get-RowValueSafe -ListObject $loGlobal2 -RowIndex $rowA2 -ColumnName "QtyOnHand") + "; QtyB=" + (Get-RowValueSafe -ListObject $loGlobal2 -RowIndex $rowB2 -ColumnName "QtyOnHand"))
    try { $wbGlobal2.Close($false) } catch {}
}
finally {
    $summary = Write-Results -ResultPath $resultPath -Rows $resultRows -DeployPath $deployPath -SessionRoot $sessionRoot

    foreach ($set in $workbookSets) {
        foreach ($name in $set.Keys) {
            try { $set[$name].Close($false) } catch {}
            Release-ComObject $set[$name]
        }
    }
    if ($null -ne $helperA) { try { $helperA.Close($false) } catch {}; Release-ComObject $helperA }
    if ($null -ne $helperB) { try { $helperB.Close($false) } catch {}; Release-ComObject $helperB }
    if ($null -ne $excelA) { try { $excelA.Quit() } catch {}; Release-ComObject $excelA }
    if ($null -ne $excelB) { try { $excelB.Quit() } catch {}; Release-ComObject $excelB }
    if ($null -ne $excelHq) { try { $excelHq.Quit() } catch {}; Release-ComObject $excelHq }
    if ($null -ne $excelSetup) { try { $excelSetup.Quit() } catch {}; Release-ComObject $excelSetup }
}

$failed = @($resultRows | Where-Object { -not $_.Passed }).Count
if ($failed -gt 0) {
    Write-Output "PHASE6_PACKAGED_WAN_HQ_FAILED"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$($resultRows.Count - $failed) FAILED=$failed TOTAL=$($resultRows.Count)"
    exit 1
}

Write-Output "PHASE6_PACKAGED_WAN_HQ_OK"
Write-Output "RESULTS=$resultPath"
Write-Output "PASSED=$($resultRows.Count) FAILED=0 TOTAL=$($resultRows.Count)"
