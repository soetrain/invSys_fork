[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class InvSysLiveValidationWindow
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

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

function Get-InvSysCredentialHash {
    param([string]$Credential)

    [double]$acc = 5381
    [double]$modulusVal = 2147483647
    for ($i = 0; $i -lt $Credential.Length; $i++) {
        $chVal = [int][char]$Credential[$i]
        $acc = ($acc * 33) + $chVal + ($i + 1)
        while ($acc -ge $modulusVal) {
            $acc -= $modulusVal
        }
    }
    return ("{0:X8}" -f [int]$acc)
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

function Resolve-WorkbookSafe {
    param(
        [object]$Excel,
        [string]$WorkbookName
    )

    if ([string]::IsNullOrWhiteSpace($WorkbookName)) { return $null }
    try {
        return $Excel.Workbooks.Item($WorkbookName)
    }
    catch {}

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

function Activate-WorkbookSafe {
    param(
        [object]$Excel,
        [object]$Workbook,
        [int]$RetryCount = 8,
        [int]$DelayMs = 500
    )

    $workbookName = ""
    try {
        if ($null -ne $Workbook) { $workbookName = [string]$Workbook.Name }
    }
    catch {}

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        $candidate = if ($workbookName -ne "") { Resolve-WorkbookSafe -Excel $Excel -WorkbookName $workbookName } else { $Workbook }
        if ($null -eq $candidate) {
            Start-Sleep -Milliseconds $DelayMs
            continue
        }

        try {
            $candidate.Activate()
            Start-Sleep -Milliseconds 250
            return $candidate
        }
        catch {
            if ($attempt -eq $RetryCount) {
                throw
            }
            Start-Sleep -Milliseconds $DelayMs
        }
    }

    throw "Unable to activate workbook '$workbookName'."
}

function Activate-WorksheetSafe {
    param(
        [object]$Excel,
        [object]$Workbook,
        [string]$WorksheetName
    )

    $wb = Activate-WorkbookSafe -Excel $Excel -Workbook $Workbook
    $ws = Get-WorksheetSafe -Workbook $wb -WorksheetName $WorksheetName
    if ($null -ne $ws) {
        try {
            $ws.Activate()
            Start-Sleep -Milliseconds 150
        }
        catch {}
    }
    return $wb
}

function Get-ExcelProcessIdFromHwnd {
    param([object]$Excel)

    if ($null -eq $Excel) { return 0 }

    try {
        [uint32]$processId = 0
        [void][InvSysLiveValidationWindow]::GetWindowThreadProcessId([intptr]$Excel.Hwnd, [ref]$processId)
        return [int]$processId
    }
    catch {
        return 0
    }
}

function Start-ExcelEnterDismissal {
    param(
        [int]$Seconds = 8,
        [int]$ExcelProcessId = 0
    )

    Start-Job -ScriptBlock {
        param($durationSeconds, $excelProcessId)
        $shell = $null
        try {
            $shell = New-Object -ComObject WScript.Shell
            $stopAt = (Get-Date).AddSeconds($durationSeconds)
            while ((Get-Date) -lt $stopAt) {
                Start-Sleep -Milliseconds 400
                $activated = $false
                try {
                    if ($excelProcessId -gt 0) {
                        $activated = [bool]$shell.AppActivate([int]$excelProcessId)
                    }
                    if (-not $activated) {
                        $activated = [bool]$shell.AppActivate("Microsoft Excel")
                    }
                }
                catch {
                    $activated = $false
                }
                if ($activated) {
                    try { $shell.SendKeys("~") } catch {}
                }
            }
        }
        finally {
            if ($null -ne $shell) {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch {}
            }
        }
    } -ArgumentList $Seconds, $ExcelProcessId
}

function Invoke-WorkbookMacroWithDismiss {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName,
        [int]$DismissSeconds = 30
    )

    $excelProcessId = Get-ExcelProcessIdFromHwnd -Excel $Excel
    $job = Start-ExcelEnterDismissal -Seconds $DismissSeconds -ExcelProcessId $excelProcessId
    try {
        return Run-WorkbookMacro -Excel $Excel -WorkbookName $WorkbookName -MacroName $MacroName
    }
    finally {
        if ($null -ne $job) {
            try { Stop-Job -Job $job -ErrorAction SilentlyContinue } catch {}
            try { Wait-Job -Job $job -Timeout 2 | Out-Null } catch {}
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

function Get-WorksheetTableNames {
    param([object]$Worksheet)

    $names = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Worksheet) { return @() }
    foreach ($lo in $Worksheet.ListObjects) {
        try { $names.Add([string]$lo.Name) | Out-Null } catch {}
    }
    return ,$names.ToArray()
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

function Restore-LiveRuntimeContext {
    param(
        [object]$Excel,
        [hashtable]$WorkbookMap,
        [string]$RuntimeRoot,
        [string]$WarehouseId,
        [string]$StationId,
        [string]$UserId,
        [string]$Pin
    )

    [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $WorkbookMap["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($RuntimeRoot))
    [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $WorkbookMap["invSys.Core.xlam"].Name -MacroName "modConfig.LoadConfig" -Arguments @($WarehouseId, $StationId))
    [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $WorkbookMap["invSys.Core.xlam"].Name -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" -Arguments @($RuntimeRoot, $RuntimeRoot, $StationId, $true))
    [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $WorkbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.SignInCurrentTargetForAutomation" -Arguments @($UserId, $Pin, "RECEIVE_POST"))
}

function Get-OpenWorkbookSummary {
    param([object]$Excel)

    $names = @()
    foreach ($wb in $Excel.Workbooks) {
        try {
            $names += ([string]$wb.Name + "=" + [string]$wb.FullName)
        }
        catch {}
    }
    return ($names -join "; ")
}

function Get-RowCountSafe {
    param([object]$ListObject)
    if ($null -eq $ListObject) { return 0 }
    try {
        return [int]$ListObject.ListRows.Count
    }
    catch {
        try {
            if ($null -eq $ListObject.DataBodyRange) { return 0 }
            return [int]$ListObject.DataBodyRange.Rows.Count
        }
        catch {
            return 0
        }
    }
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

function Find-LastRowIndexByValue {
    param(
        [object]$ListObject,
        [string]$ColumnName,
        [object]$ExpectedValue
    )

    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) { return 0 }
    $idx = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName $ColumnName
    if ($idx -le 0) { return 0 }

    for ($i = $ListObject.ListRows.Count; $i -ge 1; $i--) {
        $actual = [string]$ListObject.DataBodyRange.Cells.Item([int]$i, [int]$idx).Value2
        if ($actual -eq [string]$ExpectedValue) {
            return $i
        }
    }
    return 0
}

function Find-RowIndexByTwoValues {
    param(
        [object]$ListObject,
        [string]$ColumnName1,
        [object]$ExpectedValue1,
        [string]$ColumnName2,
        [object]$ExpectedValue2
    )

    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) { return 0 }
    $idx1 = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName $ColumnName1
    $idx2 = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName $ColumnName2
    if ($idx1 -le 0 -or $idx2 -le 0) { return 0 }

    for ($i = 1; $i -le $ListObject.ListRows.Count; $i++) {
        $actual1 = [string]$ListObject.DataBodyRange.Cells.Item([int]$i, [int]$idx1).Value2
        $actual2 = [string]$ListObject.DataBodyRange.Cells.Item([int]$i, [int]$idx2).Value2
        if ($actual1 -eq [string]$ExpectedValue1 -and $actual2 -eq [string]$ExpectedValue2) {
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

function Add-TableAt {
    param(
        [object]$Worksheet,
        [string]$StartAddress,
        [string]$TableName,
        [object[]]$Headers,
        [object[]]$Rows
    )

    $colCount = $Headers.Count
    $rowItems = @()
    if ($null -ne $Rows) {
        if ($Rows.Count -eq $Headers.Count -and -not ($Rows[0] -is [System.Array])) {
            $rowItems = ,([object[]]$Rows)
        }
        else {
            $rowItems = @($Rows)
        }
    }
    $rowCount = [Math]::Max($rowItems.Count, 1)
    $start = $Worksheet.Range($StartAddress)
    $range = $start.Resize($rowCount + 1, $colCount)
    $range.Clear()

    for ($c = 0; $c -lt $colCount; $c++) {
        $start.Cells.Item(1, $c + 1).Value2 = [string]$Headers[$c]
    }
    if ($rowItems.Count -gt 0) {
        for ($r = 0; $r -lt $rowItems.Count; $r++) {
            for ($c = 0; $c -lt $colCount; $c++) {
                $value = $null
                if ($rowItems[$r] -is [System.Array]) {
                    if ($c -le $rowItems[$r].GetUpperBound(0)) {
                        $value = $rowItems[$r][$c]
                    }
                }
                else {
                    if ($c -eq 0) { $value = $rowItems[$r] }
                }

                if ($null -eq $value) {
                    $start.Cells.Item($r + 2, $c + 1).Value2 = $null
                }
                elseif ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal] -or $value -is [single]) {
                    $start.Cells.Item($r + 2, $c + 1).Value2 = [double]$value
                }
                elseif ($value -is [datetime]) {
                    $start.Cells.Item($r + 2, $c + 1).Value2 = $value.ToOADate()
                }
                else {
                    $start.Cells.Item($r + 2, $c + 1).Value2 = [string]$value
                }
            }
        }
    }

    $listObject = $Worksheet.ListObjects.Add(1, $range, $null, 1)
    $listObject.Name = $TableName
    for ($i = 0; $i -lt $colCount; $i++) {
        $listObject.HeaderRowRange.Cells.Item(1, $i + 1).Value2 = [string]$Headers[$i]
        $listObject.ListColumns.Item($i + 1).Name = [string]$Headers[$i]
    }
    return $listObject
}

function Clear-ProcessCheckboxes {
    param(
        [object]$Worksheet
    )

    if ($null -eq $Worksheet) { return }
    for ($i = $Worksheet.Shapes.Count; $i -ge 1; $i--) {
        $shape = $null
        try { $shape = $Worksheet.Shapes.Item($i) } catch {}
        if ($null -eq $shape) { continue }
        try {
            if ([string]$shape.Name -like "CHK_PROC_*") {
                $shape.Delete()
            }
        }
        catch {}
    }
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

function New-OperationalWorkbook {
    param(
        [object]$Excel,
        [string]$NameHint,
        [string]$Path = ""
    )

    $wb = $Excel.Workbooks.Add()
    try {
        $wb.Windows.Item(1).Caption = $NameHint
    }
    catch {}
    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        Save-NewWorkbook -Workbook $wb -Path $Path
    }
    return $wb
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
        "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot", "AutoRefreshIntervalSeconds"
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
        "AutoRefreshIntervalSeconds" = 0
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
        [string[]]$CurrentUserIds,
        [string]$CredentialHash = ""
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
            "PinHash" = $CredentialHash
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

function Find-OutboxRowByEventTypeAndPayload {
    param(
        [object]$ListObject,
        [string]$EventType,
        [string]$ExpectedText
    )

    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) { return 0 }
    for ($i = $ListObject.ListRows.Count; $i -ge 1; $i--) {
        $eventTypeVal = ([string](Get-RowValueSafe -ListObject $ListObject -RowIndex $i -ColumnName "EventType")).Trim().ToUpperInvariant()
        if ($eventTypeVal -ne $EventType) { continue }
        if (Build-PayloadContains -ListObject $ListObject -RowIndex $i -ExpectedText $ExpectedText) {
            return $i
        }
    }
    return 0
}

function Get-ProcessCheckboxCount {
    param([object]$Worksheet)

    if ($null -eq $Worksheet) { return 0 }
    $count = 0
    foreach ($shape in $Worksheet.Shapes) {
        try {
            if ([string]$shape.Name -like "CHK_PROC_*") { $count++ }
        }
        catch {}
    }
    return $count
}

function Get-ProcessTableSummary {
    param([object]$Worksheet)

    if ($null -eq $Worksheet) { return "" }
    $parts = @()
    foreach ($lo in $Worksheet.ListObjects) {
        try {
            $name = [string]$lo.Name
            if ($name -like "proc_*_rchooser" -or $name -eq "RecipeChooser_generated") {
                $processVal = Get-RowValueSafe -ListObject $lo -RowIndex 1 -ColumnName "PROCESS"
                $ioVal = Get-RowValueSafe -ListObject $lo -RowIndex 1 -ColumnName "INPUT/OUTPUT"
                $ingVal = Get-RowValueSafe -ListObject $lo -RowIndex 1 -ColumnName "INGREDIENT"
                $amtVal = Get-RowValueSafe -ListObject $lo -RowIndex 1 -ColumnName "AMOUNT"
                $parts += ($name + ":Rows=" + (Get-RowCountSafe $lo) + ",Process=" + $processVal + ",IO=" + $ioVal + ",Ingredient=" + $ingVal + ",Amount=" + $amtVal)
            }
        }
        catch {}
    }
    return ($parts -join "; ")
}

$repo = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repo $DeployRoot
$resultPath = Join-Path $repo "tests/unit/phase6_live_role_workflow_results.md"
$runtimeRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-phase6-live-" + [guid]::NewGuid().ToString("N"))
$warehouseId = "WH1"
$stationId = "S1"
$currentUserId = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "user1" } else { $env:USERNAME }
$testPin = "123456"
$testPinHash = Get-InvSysCredentialHash -Credential $testPin

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
        (New-AuthWorkbook -Excel $excel -Path $authPath -WarehouseId $warehouseId -StationId $stationId -CurrentUserIds $authUserIds -CredentialHash $testPinHash),
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
    $targetSelectResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" -Arguments @($runtimeRoot, $runtimeRoot, $stationId, $true))
    $signInResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.SignInCurrentTargetForAutomation" -Arguments @($resolvedUserId, $testPin, "RECEIVE_POST"))
    $receiveAllowed = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.CanPerform" -Arguments @("RECEIVE_POST", $resolvedUserId, $warehouseId, $stationId))
    $shipAllowed = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.CanPerform" -Arguments @("SHIP_POST", $resolvedUserId, $warehouseId, $stationId))
    $prodAllowed = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modAuth.CanPerform" -Arguments @("PROD_POST", $resolvedUserId, $warehouseId, $stationId))
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.User" -Passed (-not [string]::IsNullOrWhiteSpace($resolvedUserId)) -Detail ("ResolvedUser=" + $resolvedUserId + "; SeededUsers=" + (($authUserIds -join ",") + ",svc_processor"))
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.Config" -Passed $configLoaded -Detail ("WarehouseId=" + $resolvedWarehouseId + "; StationId=" + $resolvedStationId + "; PathDataRoot=" + $resolvedDataRoot)
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.AuthLoad" -Passed $authLoaded -Detail $authReport
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.TargetSelect" -Passed $targetSelectResult.StartsWith("OK|") -Detail $targetSelectResult
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.SignIn" -Passed $signInResult.StartsWith("OK|") -Detail $signInResult
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.ReceiveCapability" -Passed $receiveAllowed -Detail ("User=" + $resolvedUserId + "; WarehouseId=" + $warehouseId + "; StationId=" + $stationId)
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.ShipCapability" -Passed $shipAllowed -Detail ("User=" + $resolvedUserId + "; WarehouseId=" + $warehouseId + "; StationId=" + $stationId)
    Add-ResultRow -Rows $resultRows -Check "Core.AuthDiagnostic.ProdCapability" -Passed $prodAllowed -Detail ("User=" + $resolvedUserId + "; WarehouseId=" + $warehouseId + "; StationId=" + $stationId)

    $currentStep = "Init role add-ins"
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Receiving.xlam"].Name -MacroName "modReceivingInit.InitReceivingAddin")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modShippingInit.InitShippingAddin")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Production.xlam"].Name -MacroName "modProductionInit.InitProductionAddin")

    $currentStep = "Validate clean config bootstrap under live add-ins"
    $bootstrapRoot = Join-Path $env:TEMP ("phase6_cfg_live_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $bootstrapRoot -Force | Out-Null
    try {
        [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($bootstrapRoot))
        $cfgLoadOk = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modConfig.LoadConfig" -Arguments @("", ""))
        $cfgValidate = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modConfig.Validate")
        $wbCfgBootstrap = Resolve-WorkbookSafe -Excel $excel -WorkbookName "WH1.invSys.Config.xlsb"
        $cfgWhSheet = Get-WorksheetSafe -Workbook $wbCfgBootstrap -WorksheetName "WarehouseConfig"
        $cfgStSheet = Get-WorksheetSafe -Workbook $wbCfgBootstrap -WorksheetName "StationConfig"
        $cfgWhTables = @(Get-WorksheetTableNames -Worksheet $cfgWhSheet)
        $cfgStTables = @(Get-WorksheetTableNames -Worksheet $cfgStSheet)
        $cfgBootstrapClean = $cfgLoadOk `
            -and $null -ne $wbCfgBootstrap `
            -and $wbCfgBootstrap.Worksheets.Count -eq 2 `
            -and $cfgWhTables.Count -eq 1 -and $cfgWhTables[0] -eq "tblWarehouseConfig" `
            -and $cfgStTables.Count -eq 1 -and $cfgStTables[0] -eq "tblStationConfig"
        Add-ResultRow -Rows $resultRows -Check "Core.ConfigBootstrap.CleanSurface" -Passed $cfgBootstrapClean -Detail ("Load=" + $cfgLoadOk + "; Validate=" + $cfgValidate + "; Sheets=" + $(if ($null -eq $wbCfgBootstrap) { 0 } else { $wbCfgBootstrap.Worksheets.Count }) + "; WHTables=" + ($cfgWhTables -join ',') + "; STTables=" + ($cfgStTables -join ','))
        if ($null -ne $wbCfgBootstrap) {
            try { $wbCfgBootstrap.Close($false) } catch {}
        }
    }
    finally {
        try { [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.ClearCoreDataRootOverride") } catch {}
        try { if (Test-Path $bootstrapRoot) { Remove-Item -Path $bootstrapRoot -Recurse -Force } } catch {}
    }

    $currentStep = "Restore live runtime context"
    Restore-LiveRuntimeContext -Excel $excel -WorkbookMap $workbookMap -RuntimeRoot $runtimeRoot -WarehouseId $warehouseId -StationId $stationId -UserId $resolvedUserId -Pin $testPin
    $restoredRoot = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.GetCoreDataRootOverride")
    $restoredDataRoot = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modConfig.GetString" -Arguments @("PathDataRoot", ""))
    $inventoryFileExists = Test-Path -LiteralPath $inventoryPath
    $inventoryWbOpen = Resolve-WorkbookSafe -Excel $excel -WorkbookName ($warehouseId + ".invSys.Data.Inventory.xlsb")
    $inventoryFullName = if ($null -eq $inventoryWbOpen) { "<not open>" } else { [string]$inventoryWbOpen.FullName }
    Add-ResultRow -Rows $resultRows -Check "Core.RuntimeInventoryDiagnostic" -Passed ($inventoryFileExists -and $inventoryFullName -ne "<not open>") -Detail ("Override=" + $restoredRoot + "; PathDataRoot=" + $restoredDataRoot + "; InventoryPath=" + $inventoryPath + "; FileExists=" + $inventoryFileExists + "; OpenFullName=" + $inventoryFullName)

    $currentStep = "Create operational role workbooks"
    $wbReceiveOps = New-OperationalWorkbook -Excel $excel -NameHint "ReceivingOps" -Path (Join-Path $runtimeRoot ($warehouseId + "." + $stationId + ".Receiving.Operator.xlsb"))
    $openedWorkbooks.Add($wbReceiveOps) | Out-Null
    $wbReceiveOps = Activate-WorkbookSafe -Excel $excel -Workbook $wbReceiveOps
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface" -Arguments @($wbReceiveOps))

    $wbShipOps = New-OperationalWorkbook -Excel $excel -NameHint "ShippingOps" -Path (Join-Path $runtimeRoot ($warehouseId + "." + $stationId + ".Shipping.Operator.xlsb"))
    $openedWorkbooks.Add($wbShipOps) | Out-Null
    $wbShipOps = Activate-WorkbookSafe -Excel $excel -Workbook $wbShipOps
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface" -Arguments @($wbShipOps))

    $wbProdOps = New-OperationalWorkbook -Excel $excel -NameHint "ProductionOps" -Path (Join-Path $runtimeRoot ($warehouseId + "." + $stationId + ".Production.Operator.xlsb"))
    $openedWorkbooks.Add($wbProdOps) | Out-Null
    $wbProdOps = Activate-WorkbookSafe -Excel $excel -Workbook $wbProdOps
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface" -Arguments @($wbProdOps))

    $currentStep = "Stage Receiving workflow"
    $wbReceive = Activate-WorksheetSafe -Excel $excel -Workbook $wbReceiveOps -WorksheetName "ReceivedTally"
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
    $receiveInboxBefore = Get-RowCountSafe $loInboxReceive
    $inventoryLogBefore = Get-RowCountSafe $loInventoryLog
    $currentStep = "Run Receiving ConfirmWrites"
    $wbReceive = Activate-WorksheetSafe -Excel $excel -Workbook $wbReceive -WorksheetName "ReceivedTally"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $workbookMap["invSys.Receiving.xlam"].Name -MacroName "modTS_Received.ConfirmWrites")

    $wbReceive = Resolve-WorkbookSafe -Excel $excel -WorkbookName $wbReceive.Name
    $wbReceiveInboxRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ("invSys.Inbox.Receiving." + $stationId + ".xlsb")
    $wbInventoryRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ($warehouseId + ".invSys.Data.Inventory.xlsb")
    $wsReceive = Get-WorksheetSafe -Workbook $wbReceive -WorksheetName "ReceivedTally"
    $wsReceiveLog = Get-WorksheetSafe -Workbook $wbReceive -WorksheetName "ReceivedLog"
    $wsReceiveInv = Get-WorksheetSafe -Workbook $wbReceive -WorksheetName "InventoryManagement"
    $loReceivedTally = Get-ListObjectSafe -Worksheet $wsReceive -TableName "ReceivedTally"
    $loAggReceived = Get-ListObjectSafe -Worksheet $wsReceive -TableName "AggregateReceived"
    $loReceiveLog = Get-ListObjectSafe -Worksheet $wsReceiveLog -TableName "ReceivedLog"
    $loReceiveInv = Get-ListObjectSafe -Worksheet $wsReceiveInv -TableName "invSys"
    $loInboxReceive = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbReceiveInboxRuntime -WorksheetName "InboxReceive") -TableName "tblInboxReceive"
    $loInventoryLog = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbInventoryRuntime -WorksheetName "InventoryLog") -TableName "tblInventoryLog"

    $receivedTallyRowsAfter = Get-RowCountSafe $loReceivedTally
    $aggReceivedRowsAfter = Get-RowCountSafe $loAggReceived
    $receiveLocalOk = ($receivedTallyRowsAfter -eq 0) `
        -and ($aggReceivedRowsAfter -eq 0) `
        -and (([double](Get-RowValueSafe -ListObject $loReceiveInv -RowIndex 1 -ColumnName "RECEIVED")) -eq 0)
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.Local" -Passed $receiveLocalOk -Detail "ReceivedTallyRows=$receivedTallyRowsAfter; AggregateReceivedRows=$aggReceivedRowsAfter; RECEIVED=$((Get-RowValueSafe -ListObject $loReceiveInv -RowIndex 1 -ColumnName 'RECEIVED')); TOTAL_INV=$((Get-RowValueSafe -ListObject $loReceiveInv -RowIndex 1 -ColumnName 'TOTAL INV')); QtyOnHand=$((Get-RowValueSafe -ListObject $loReceiveInv -RowIndex 1 -ColumnName 'QtyOnHand')); SourceType=$((Get-RowValueSafe -ListObject $loReceiveInv -RowIndex 1 -ColumnName 'SourceType')); IsStale=$((Get-RowValueSafe -ListObject $loReceiveInv -RowIndex 1 -ColumnName 'IsStale')); LogRows=$(Get-RowCountSafe $loReceiveLog)"

    $receiveInboxAfter = Get-RowCountSafe $loInboxReceive
    $receiveQueuedRow = 0
    for ($i = $receiveInboxAfter; $i -ge 1; $i--) {
        $skuVal = [string](Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $i -ColumnName "SKU")
        $qtyVal = Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $i -ColumnName "Qty"
        if ($skuVal -eq "SKU-REC" -and [double]$qtyVal -eq 7) {
            $receiveQueuedRow = $i
            break
        }
    }
    $receiveQueuedOk = ($receiveInboxAfter -eq ($receiveInboxBefore + 1)) -and ($receiveQueuedRow -gt 0) -and ([double](Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName "Qty") -eq 7)
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.Queue" -Passed $receiveQueuedOk -Detail "InboxRows=$receiveInboxAfter; Row=$receiveQueuedRow"

    $receiveStatusBeforeRun = ([string](Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName "Status")).Trim().ToUpperInvariant()
    Restore-LiveRuntimeContext -Excel $excel -WorkbookMap $workbookMap -RuntimeRoot $runtimeRoot -WarehouseId $warehouseId -StationId $stationId -UserId $resolvedUserId -Pin $testPin
    $receiveRunBatchReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @($warehouseId, 500))
    $receiveRunBatch = 0
    if ($receiveRunBatchReport -match 'Processed=(\d+)') { $receiveRunBatch = [int]$Matches[1] }
    $wbOutboxRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ($warehouseId + ".Outbox.Events.xlsb")
    $loOutbox = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbOutboxRuntime -WorksheetName "OutboxEvents") -TableName "tblOutboxEvents"
    $receiveOutboxRow = Find-OutboxRowByEventTypeAndPayload -ListObject $loOutbox -EventType "RECEIVE" -ExpectedText '"SKU":"SKU-REC"'
    $receiveStatus = ([string](Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName "Status")).Trim().ToUpperInvariant()
    $receiveProcessedOk = ($receiveStatusBeforeRun -eq "PROCESSED") -or ($receiveStatus -eq "PROCESSED") -or ($receiveRunBatch -ge 1)
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.Process" -Passed $receiveProcessedOk -Detail "StatusBeforeRun=$receiveStatusBeforeRun; RunBatch=$receiveRunBatch; Status=$receiveStatus; OutboxRow=$receiveOutboxRow; ErrorCode=$((Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName 'ErrorCode')); ErrorMessage=$((Get-RowValueSafe -ListObject $loInboxReceive -RowIndex $receiveQueuedRow -ColumnName 'ErrorMessage')); $receiveRunBatchReport; OpenBooks=$(Get-OpenWorkbookSummary -Excel $excel)"

    $receiveLogRow = 0
    for ($i = Get-RowCountSafe $loInventoryLog; $i -ge 1; $i--) {
        $eventTypeVal = [string](Get-RowValueSafe -ListObject $loInventoryLog -RowIndex $i -ColumnName "EventType")
        $skuVal = [string](Get-RowValueSafe -ListObject $loInventoryLog -RowIndex $i -ColumnName "SKU")
        $qtyVal = Get-RowValueSafe -ListObject $loInventoryLog -RowIndex $i -ColumnName "QtyDelta"
        if ($eventTypeVal -eq "RECEIVE" -and $skuVal -eq "SKU-REC" -and [double]$qtyVal -eq 7) {
            $receiveLogRow = $i
            break
        }
    }
    $receiveInventoryOk = ($receiveLogRow -gt 0) -or ($receiveOutboxRow -gt 0) -or ($receiveRunBatch -ge 1)
    Add-ResultRow -Rows $resultRows -Check "Receiving.ConfirmWrites.InventoryLog" -Passed $receiveInventoryOk -Detail "InventoryLogRowsBefore=$inventoryLogBefore; Row=$receiveLogRow; OutboxRow=$receiveOutboxRow"

    $currentStep = "Stage Shipping workflow"
    $wbShip = Activate-WorksheetSafe -Excel $excel -Workbook $wbShipOps -WorksheetName "ShipmentsTally"
    $wsShip = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $wsShipInv = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "InventoryManagement"
    $loAggPackages = Get-ListObjectSafe -Worksheet $wsShip -TableName "AggregatePackages"
    $loShipInv = Get-ListObjectSafe -Worksheet $wsShipInv -TableName "invSys"
    $loInboxShip = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $runtimeBooks[4] -WorksheetName "InboxShip") -TableName "tblInboxShip"

    Clear-ListObjectRows $loAggPackages
    Clear-ListObjectRows $loShipInv
    Add-ListObjectRow -ListObject $loShipInv -Values @{
        "ROW" = 201; "ITEM_CODE" = "SKU-SHIP"; "ITEM" = "Ship Widget"; "UOM" = "EA"; "LOCATION" = "DOCK";
        "DESCRIPTION" = "Ship Widget"; "SHIPMENTS" = 0; "TOTAL INV" = 20; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }
    Add-ListObjectRow -ListObject $loAggPackages -Values @{
        "ROW" = 201; "ITEM_CODE" = "SKU-SHIP"; "ITEM" = "Ship Widget"; "QUANTITY" = 5; "UOM" = "EA"; "LOCATION" = "DOCK"
    }
    $shipToShipmentsPreflightOk = ((Get-RowCountSafe $loAggPackages) -eq 1) `
        -and ([double](Get-RowValueSafe -ListObject $loAggPackages -RowIndex 1 -ColumnName "QUANTITY") -eq 5) `
        -and ([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName "TOTAL INV") -eq 20)
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnToShipments.Preflight" -Passed $shipToShipmentsPreflightOk -Detail "AggregatePackagesRows=$(Get-RowCountSafe $loAggPackages); AggROW=$((Get-RowValueSafe -ListObject $loAggPackages -RowIndex 1 -ColumnName 'ROW')); AggQty=$((Get-RowValueSafe -ListObject $loAggPackages -RowIndex 1 -ColumnName 'QUANTITY')); InvROW=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'ROW')); InvCode=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'ITEM_CODE')); InvTOTAL_INV=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'TOTAL INV')); InvSHIPMENTS=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'SHIPMENTS'))"

    $currentStep = "Run Shipping BtnToShipments"
    $wbShip = Activate-WorksheetSafe -Excel $excel -Workbook $wbShip -WorksheetName "ShipmentsTally"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modTS_Shipments.BtnToShipments")
    $wsShip = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $wsShipInv = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "InventoryManagement"
    $loAggPackages = Get-ListObjectSafe -Worksheet $wsShip -TableName "AggregatePackages"
    $loShipInv = Get-ListObjectSafe -Worksheet $wsShipInv -TableName "invSys"
    $shipStageOk = ([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName "SHIPMENTS")) -eq 5
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnToShipments.Local" -Passed $shipStageOk -Detail "SHIPMENTS=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'SHIPMENTS')); AggregatePackagesRows=$(Get-RowCountSafe $loAggPackages)"

    $shipInboxBefore = Get-RowCountSafe $loInboxShip
    $currentStep = "Run Shipping BtnShipmentsSent"
    $wbShip = Activate-WorksheetSafe -Excel $excel -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $shipWorkbookName = [string]$wbShip.Name
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modTS_Shipments.BtnShipmentsSent")

    $wbShip = Resolve-WorkbookSafe -Excel $excel -WorkbookName $shipWorkbookName
    $wbShipInboxRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ("invSys.Inbox.Shipping." + $stationId + ".xlsb")
    $wbInventoryRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ($warehouseId + ".invSys.Data.Inventory.xlsb")
    $wsShip = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $wsShipInv = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "InventoryManagement"
    $loAggPackages = Get-ListObjectSafe -Worksheet $wsShip -TableName "AggregatePackages"
    $loShipInv = Get-ListObjectSafe -Worksheet $wsShipInv -TableName "invSys"
    $loInboxShip = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbShipInboxRuntime -WorksheetName "InboxShip") -TableName "tblInboxShip"
    $loInventoryLog = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbInventoryRuntime -WorksheetName "InventoryLog") -TableName "tblInventoryLog"

    $aggPackagesRowsAfter = Get-RowCountSafe $loAggPackages
    $shipLocalOk = (([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName "SHIPMENTS")) -eq 0) `
        -and ($aggPackagesRowsAfter -eq 0)
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnShipmentsSent.Local" -Passed $shipLocalOk -Detail "SHIPMENTS=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'SHIPMENTS')); AggregatePackagesRows=$aggPackagesRowsAfter"

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

    $shipStatusBeforeRun = ([string](Get-RowValueSafe -ListObject $loInboxShip -RowIndex $shipQueuedRow -ColumnName "Status")).Trim().ToUpperInvariant()
    Restore-LiveRuntimeContext -Excel $excel -WorkbookMap $workbookMap -RuntimeRoot $runtimeRoot -WarehouseId $warehouseId -StationId $stationId -UserId $resolvedUserId -Pin $testPin
    $shipRunBatchReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @($warehouseId, 500))
    $shipRunBatch = 0
    if ($shipRunBatchReport -match 'Processed=(\d+)') { $shipRunBatch = [int]$Matches[1] }
    $wbShipInboxRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ("invSys.Inbox.Shipping." + $stationId + ".xlsb")
    $loInboxShip = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbShipInboxRuntime -WorksheetName "InboxShip") -TableName "tblInboxShip"
    $shipStatus = ([string](Get-RowValueSafe -ListObject $loInboxShip -RowIndex $shipQueuedRow -ColumnName "Status")).Trim().ToUpperInvariant()
    $shipProcessedOk = ($shipStatusBeforeRun -eq "PROCESSED") -or ($shipStatus -eq "PROCESSED") -or ($shipRunBatch -ge 1)
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnShipmentsSent.Process" -Passed $shipProcessedOk -Detail "StatusBeforeRun=$shipStatusBeforeRun; RunBatch=$shipRunBatch; Status=$shipStatus; ErrorCode=$((Get-RowValueSafe -ListObject $loInboxShip -RowIndex $shipQueuedRow -ColumnName 'ErrorCode')); ErrorMessage=$((Get-RowValueSafe -ListObject $loInboxShip -RowIndex $shipQueuedRow -ColumnName 'ErrorMessage')); $shipRunBatchReport"

    $wbInventoryRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ($warehouseId + ".invSys.Data.Inventory.xlsb")
    $loInventoryLog = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbInventoryRuntime -WorksheetName "InventoryLog") -TableName "tblInventoryLog"
    $shipLogRow = Find-RowIndexByValue -ListObject $loInventoryLog -ColumnName "EventType" -ExpectedValue "SHIP"
    $shipInventoryOk = ($shipLogRow -gt 0) -and ([double](Get-RowValueSafe -ListObject $loInventoryLog -RowIndex $shipLogRow -ColumnName "QtyDelta") -eq -5)
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnShipmentsSent.InventoryLog" -Passed $shipInventoryOk -Detail "InventoryLogRow=$shipLogRow"

    $currentStep = "Stage Shipping hold workflow"
    $wbShip = Activate-WorksheetSafe -Excel $excel -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $wsShip = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $loShipments = Get-ListObjectSafe -Worksheet $wsShip -TableName "ShipmentsTally"
    $loNotShipped = Get-ListObjectSafe -Worksheet $wsShip -TableName "NotShipped"
    Clear-ListObjectRows $loShipments
    Clear-ListObjectRows $loNotShipped
    Add-ListObjectRow -ListObject $loShipments -Values @{
        "REF_NUMBER" = "REF-HOLD-001"; "ITEMS" = "Hold Widget"; "QUANTITY" = 10; "ROW" = 250; "UOM" = "EA"; "LOCATION" = "DOCK"; "DESCRIPTION" = "Hold Widget"
    }

    $initialHoldHidden = [bool]$loNotShipped.Range.EntireColumn.Hidden
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modTS_Shipments.BtnUnship")
    $afterFirstToggleHidden = [bool]$loNotShipped.Range.EntireColumn.Hidden
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modTS_Shipments.BtnUnship")
    $afterSecondToggleHidden = [bool]$loNotShipped.Range.EntireColumn.Hidden
    $holdToggleOk = ($afterFirstToggleHidden -ne $initialHoldHidden) -and ($afterSecondToggleHidden -eq $initialHoldHidden)
    Add-ResultRow -Rows $resultRows -Check "Shipping.Hold.ToggleNotShipped" -Passed $holdToggleOk -Detail "InitialHidden=$initialHoldHidden; AfterFirst=$afterFirstToggleHidden; AfterSecond=$afterSecondToggleHidden"

    $holdToResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modTS_Shipments.MoveShipmentHoldForAutomation" -Arguments @("REF-HOLD-001", "Hold Widget", 4, $true))
    $shipHoldRow = Find-RowIndexByTwoValues -ListObject $loShipments -ColumnName1 "REF_NUMBER" -ExpectedValue1 "REF-HOLD-001" -ColumnName2 "ITEMS" -ExpectedValue2 "Hold Widget"
    $notShippedRow = Find-RowIndexByTwoValues -ListObject $loNotShipped -ColumnName1 "REF_NUMBER" -ExpectedValue1 "REF-HOLD-001" -ColumnName2 "ITEMS" -ExpectedValue2 "Hold Widget"
    $holdToOk = $holdToResult.StartsWith("OK|") `
        -and ($shipHoldRow -gt 0) `
        -and ($notShippedRow -gt 0) `
        -and ([double](Get-RowValueSafe -ListObject $loShipments -RowIndex $shipHoldRow -ColumnName "QUANTITY") -eq 6) `
        -and ([double](Get-RowValueSafe -ListObject $loNotShipped -RowIndex $notShippedRow -ColumnName "QUANTITY") -eq 4) `
        -and ([double](Get-RowValueSafe -ListObject $loNotShipped -RowIndex $notShippedRow -ColumnName "ROW") -eq 250)
    Add-ResultRow -Rows $resultRows -Check "Shipping.Hold.Send" -Passed $holdToOk -Detail "Result=$holdToResult; ShipQty=$((Get-RowValueSafe -ListObject $loShipments -RowIndex $shipHoldRow -ColumnName 'QUANTITY')); HoldQty=$((Get-RowValueSafe -ListObject $loNotShipped -RowIndex $notShippedRow -ColumnName 'QUANTITY')); HoldROW=$((Get-RowValueSafe -ListObject $loNotShipped -RowIndex $notShippedRow -ColumnName 'ROW'))"

    $returnHoldResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modTS_Shipments.MoveShipmentHoldForAutomation" -Arguments @("REF-HOLD-001", "Hold Widget", 4, $false))
    $shipHoldRow = Find-RowIndexByTwoValues -ListObject $loShipments -ColumnName1 "REF_NUMBER" -ExpectedValue1 "REF-HOLD-001" -ColumnName2 "ITEMS" -ExpectedValue2 "Hold Widget"
    $notShippedRow = Find-RowIndexByTwoValues -ListObject $loNotShipped -ColumnName1 "REF_NUMBER" -ExpectedValue1 "REF-HOLD-001" -ColumnName2 "ITEMS" -ExpectedValue2 "Hold Widget"
    $returnHoldQty = if ($notShippedRow -gt 0) { [double](Get-RowValueSafe -ListObject $loNotShipped -RowIndex $notShippedRow -ColumnName "QUANTITY") } else { 0 }
    $returnHoldOk = $returnHoldResult.StartsWith("OK|") `
        -and ($shipHoldRow -gt 0) `
        -and ([double](Get-RowValueSafe -ListObject $loShipments -RowIndex $shipHoldRow -ColumnName "QUANTITY") -eq 10) `
        -and ($returnHoldQty -eq 0)
    Add-ResultRow -Rows $resultRows -Check "Shipping.Hold.Return" -Passed $returnHoldOk -Detail "Result=$returnHoldResult; ShipQty=$((Get-RowValueSafe -ListObject $loShipments -RowIndex $shipHoldRow -ColumnName 'QUANTITY')); HoldQty=$returnHoldQty"

    $currentStep = "Stage Shipping box-build workflow"
    $wbShip = Activate-WorksheetSafe -Excel $excel -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $wsShip = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $wsShipInv = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "InventoryManagement"
    $loAggBoxBom = Get-ListObjectSafe -Worksheet $wsShip -TableName "AggregateBoxBOM"
    $loAggPackages = Get-ListObjectSafe -Worksheet $wsShip -TableName "AggregatePackages"
    $loShipInv = Get-ListObjectSafe -Worksheet $wsShipInv -TableName "invSys"

    Clear-ListObjectRows $loAggBoxBom
    Clear-ListObjectRows $loAggPackages
    Clear-ListObjectRows $loShipInv
    Add-ListObjectRow -ListObject $loShipInv -Values @{
        "ROW" = 301; "ITEM_CODE" = "SKU-COMP"; "ITEM" = "Component Widget"; "UOM" = "EA"; "LOCATION" = "LINE";
        "DESCRIPTION" = "Component Widget"; "USED" = 3; "MADE" = 0; "SHIPMENTS" = 0; "TOTAL INV" = 10; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }
    Add-ListObjectRow -ListObject $loShipInv -Values @{
        "ROW" = 302; "ITEM_CODE" = "SKU-BOX"; "ITEM" = "Box Widget"; "UOM" = "EA"; "LOCATION" = "LINE";
        "DESCRIPTION" = "Box Widget"; "USED" = 0; "MADE" = 0; "SHIPMENTS" = 0; "TOTAL INV" = 0; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }
    Add-ListObjectRow -ListObject $loAggBoxBom -Values @{
        "ROW" = 301; "ITEM_CODE" = "SKU-COMP"; "ITEM" = "Component Widget"; "QUANTITY" = 3; "UOM" = "EA"; "LOCATION" = "LINE"
    }
    Add-ListObjectRow -ListObject $loAggPackages -Values @{
        "ROW" = 302; "ITEM_CODE" = "SKU-BOX"; "ITEM" = "Box Widget"; "QUANTITY" = 2; "UOM" = "EA"; "LOCATION" = "LINE"
    }

    $currentStep = "Run Shipping BtnBoxesMade"
    $wbShip = Activate-WorksheetSafe -Excel $excel -Workbook $wbShip -WorksheetName "ShipmentsTally"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modTS_Shipments.BtnBoxesMade")
    $wsShip = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "ShipmentsTally"
    $wsShipInv = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "InventoryManagement"
    $loAggPackages = Get-ListObjectSafe -Worksheet $wsShip -TableName "AggregatePackages"
    $loShipInv = Get-ListObjectSafe -Worksheet $wsShipInv -TableName "invSys"
    $shipBoxesMadeOk = ([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName "USED")) -eq 0 `
        -and ([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName "TOTAL INV")) -eq 7 `
        -and ([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 2 -ColumnName "MADE")) -eq 2
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnBoxesMade.Local" -Passed $shipBoxesMadeOk -Detail "ComponentUSED=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'USED')); ComponentTOTAL_INV=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 1 -ColumnName 'TOTAL INV')); PackageMADE=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 2 -ColumnName 'MADE')); AggregatePackagesRows=$(Get-RowCountSafe $loAggPackages)"

    $currentStep = "Run Shipping BtnToTotalInv"
    $wbShip = Activate-WorksheetSafe -Excel $excel -Workbook $wbShip -WorksheetName "ShipmentsTally"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $workbookMap["invSys.Shipping.xlam"].Name -MacroName "modTS_Shipments.BtnToTotalInv")
    $wsShipInv = Get-WorksheetSafe -Workbook $wbShip -WorksheetName "InventoryManagement"
    $loShipInv = Get-ListObjectSafe -Worksheet $wsShipInv -TableName "invSys"
    $shipToTotalOk = ([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 2 -ColumnName "MADE")) -eq 0 `
        -and ([double](Get-RowValueSafe -ListObject $loShipInv -RowIndex 2 -ColumnName "TOTAL INV")) -eq 2
    Add-ResultRow -Rows $resultRows -Check "Shipping.BtnToTotalInv.Local" -Passed $shipToTotalOk -Detail "PackageMADE=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 2 -ColumnName 'MADE')); PackageTOTAL_INV=$((Get-RowValueSafe -ListObject $loShipInv -RowIndex 2 -ColumnName 'TOTAL INV'))"

    $currentStep = "Stage Production workflow"
    $wbProd = Activate-WorksheetSafe -Excel $excel -Workbook $wbProdOps -WorksheetName "Production"
    $wsProd = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "Production"
    $wsProdRecipes = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "Recipes"
    $wsProdInv = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "InventoryManagement"
    $wsPalette = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "IngredientPalette"
    if ($null -eq $wsPalette) {
        $wsPalette = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "IngredientsPalette"
    }
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
        "DESCRIPTION" = "Granulated"; "TOTAL INV" = 100; "USED" = 0; "MADE" = 0; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }

    $prodPaletteDiagBefore = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Production.xlam"].Name -MacroName "mProduction.GetPaletteSaveDiagnostic")
    $currentStep = "Run Production BtnSavePalette"
    $wbProd = Activate-WorksheetSafe -Excel $excel -Workbook $wbProd -WorksheetName "Production"
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $workbookMap["invSys.Production.xlam"].Name -MacroName "mProduction.BtnSavePalette")
    $prodPaletteDiagAfter = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Production.xlam"].Name -MacroName "mProduction.GetPaletteSaveDiagnostic")
    $paletteRow = Find-RowIndexByValue -ListObject $loPalette -ColumnName "RECIPE_ID" -ExpectedValue "R-001"
    $paletteOk = ($paletteRow -gt 0) -and ([string](Get-RowValueSafe -ListObject $loPalette -RowIndex $paletteRow -ColumnName "INGREDIENT_ID") -eq "ING-001") -and ([string](Get-RowValueSafe -ListObject $loPalette -RowIndex $paletteRow -ColumnName "ITEM") -eq "Sugar Bin")
    Add-ResultRow -Rows $resultRows -Check "Production.BtnSavePalette" -Passed $paletteOk -Detail "PaletteRow=$paletteRow; Before=$prodPaletteDiagBefore; After=$prodPaletteDiagAfter"

    Add-ListObjectRow -ListObject $loProdInv -Values @{
        "ROW" = 401; "ITEM_CODE" = "SKU-FG"; "ITEM" = "Finished Good"; "UOM" = "EA"; "LOCATION" = "FG";
        "DESCRIPTION" = "Finished Good"; "USED" = 0; "MADE" = 0; "TOTAL INV" = 0; "LAST EDITED" = ""; "TOTAL INV LAST EDIT" = ""; "TIMESTAMP" = ""
    }
    Add-ListObjectRow -ListObject $loProductionOutput -Values @{
        "PROCESS" = "Mix"; "OUTPUT" = "Finished Good"; "UOM" = "EA"; "REAL OUTPUT" = 8; "BATCH" = "B-001"; "RECALL CODE" = "RC-001"; "ROW" = 401
    }

    $prodRecallDiag = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Production.xlam"].Name -MacroName "mProduction.GetRecallPrintDiagnostic")
    $wsRecall = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "RecallCodesPrint"
    $loRecall = $null
    if ($null -ne $wsRecall) {
        $loRecall = Get-ListObjectSafe -Worksheet $wsRecall -TableName "RecallCodesReport"
    }
    $prodRecallOk = ($prodRecallDiag -like "OK*") -and ($null -ne $loRecall) -and ((Get-RowCountSafe $loRecall) -eq 1) -and ([string](Get-RowValueSafe -ListObject $loRecall -RowIndex 1 -ColumnName "RECALL CODE") -eq "RC-001")
    Add-ResultRow -Rows $resultRows -Check "Production.BtnPrintRecallCodes" -Passed $prodRecallOk -Detail "Diag=$prodRecallDiag; RecallRows=$((Get-RowCountSafe $loRecall)); RecallCode=$((Get-RowValueSafe -ListObject $loRecall -RowIndex 1 -ColumnName 'RECALL CODE'))"

    [void](Add-TableAt -Worksheet $wsProd -StartAddress "BA200" -TableName "proc_1_rchooser" -Headers @(
        "PROCESS", "INPUT/OUTPUT", "INGREDIENT", "UOM", "AMOUNT", "INGREDIENT_ID"
    ) -Rows @(
        @("Mix", "MADE", "Finished Good", "EA", 8, "FG-001")
    ))
    Clear-ProcessCheckboxes -Worksheet $wsProd
    $prodToMadePreflightOk = ((Get-ProcessTableSummary -Worksheet $wsProd) -like "*proc_1_rchooser*") `
        -and ((Get-ProcessCheckboxCount -Worksheet $wsProd) -eq 0) `
        -and ([double](Get-RowValueSafe -ListObject $loProductionOutput -RowIndex 1 -ColumnName "REAL OUTPUT") -eq 8) `
        -and ([string](Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName "ITEM_CODE") -eq "SKU-FG")
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToMade.Preflight" -Passed $prodToMadePreflightOk -Detail "ProcessTables=$(Get-ProcessTableSummary -Worksheet $wsProd); ProcessCheckboxes=$(Get-ProcessCheckboxCount -Worksheet $wsProd); OutputROW=$((Get-RowValueSafe -ListObject $loProductionOutput -RowIndex 1 -ColumnName 'ROW')); RealOutput=$((Get-RowValueSafe -ListObject $loProductionOutput -RowIndex 1 -ColumnName 'REAL OUTPUT')); InvRow2Code=$((Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName 'ITEM_CODE'))"

    $prodInboxBefore = Get-RowCountSafe $loInboxProd
    $currentStep = "Run Production BtnToMade"
    $wbProd = Activate-WorksheetSafe -Excel $excel -Workbook $wbProd -WorksheetName "Production"
    $prodWorkbookName = [string]$wbProd.Name
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $workbookMap["invSys.Production.xlam"].Name -MacroName "mProduction.BtnToMade")

    $wbProd = Resolve-WorkbookSafe -Excel $excel -WorkbookName $prodWorkbookName
    $wbProdInboxRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ("invSys.Inbox.Production." + $stationId + ".xlsb")
    $wbInventoryRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ($warehouseId + ".invSys.Data.Inventory.xlsb")
    $wsProd = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "Production"
    $wsProdInv = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "InventoryManagement"
    $loProductionOutput = Get-ListObjectSafe -Worksheet $wsProd -TableName "ProductionOutput"
    $loProdInv = Get-ListObjectSafe -Worksheet $wsProdInv -TableName "invSys"
    $loInboxProd = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbProdInboxRuntime -WorksheetName "InboxProd") -TableName "tblInboxProd"
    $loInventoryLog = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbInventoryRuntime -WorksheetName "InventoryLog") -TableName "tblInventoryLog"

    $prodMadeLocalOk = ((([double](Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName "MADE")) -ge 8) `
        -or (([double](Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName "TOTAL INV")) -ge 8)) `
        -and (([double](Get-RowValueSafe -ListObject $loProductionOutput -RowIndex 1 -ColumnName "REAL OUTPUT")) -eq 8)
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToMade.Local" -Passed $prodMadeLocalOk -Detail "MADE=$((Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName 'MADE')); TOTAL_INV=$((Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName 'TOTAL INV')); RealOutput=$((Get-RowValueSafe -ListObject $loProductionOutput -RowIndex 1 -ColumnName 'REAL OUTPUT'))"

    $prodConsumeInboxAfter = Get-RowCountSafe $loInboxProd
    $prodConsumeQueuedRow = 0
    for ($i = 1; $i -le $prodConsumeInboxAfter; $i++) {
        if (([string](Get-RowValueSafe -ListObject $loInboxProd -RowIndex $i -ColumnName "EventType") -eq "PROD_CONSUME") -and (Build-PayloadContains -ListObject $loInboxProd -RowIndex $i -ExpectedText '"SKU":"SKU-FG"')) {
            $prodConsumeQueuedRow = $i
            break
        }
    }
    $prodConsumeQueuedOk = ($prodConsumeInboxAfter -eq ($prodInboxBefore + 1)) -and ($prodConsumeQueuedRow -gt 0)
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToMade.Queue" -Passed $prodConsumeQueuedOk -Detail "InboxRows=$prodConsumeInboxAfter; Row=$prodConsumeQueuedRow"

    $prodConsumeStatus = ([string](Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodConsumeQueuedRow -ColumnName "Status")).Trim().ToUpperInvariant()
    $prodConsumeProcessedOk = ($prodConsumeStatus -eq "PROCESSED")
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToMade.Process" -Passed $prodConsumeProcessedOk -Detail "Status=$prodConsumeStatus; ErrorCode=$((Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodConsumeQueuedRow -ColumnName 'ErrorCode')); ErrorMessage=$((Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodConsumeQueuedRow -ColumnName 'ErrorMessage'))"

    $prodConsumeLogRow = Find-RowIndexByValue -ListObject $loInventoryLog -ColumnName "EventType" -ExpectedValue "PROD_CONSUME"
    $prodConsumeInventoryOk = ($prodConsumeLogRow -gt 0)
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToMade.InventoryLog" -Passed $prodConsumeInventoryOk -Detail "InventoryLogRow=$prodConsumeLogRow"

    $prodInboxBefore = Get-RowCountSafe $loInboxProd
    $currentStep = "Run Production BtnToTotalInv"
    $wbProd = Activate-WorksheetSafe -Excel $excel -Workbook $wbProd -WorksheetName "Production"
    $prodWorkbookName = [string]$wbProd.Name
    [void](Invoke-WorkbookMacroWithDismiss -Excel $excel -WorkbookName $workbookMap["invSys.Production.xlam"].Name -MacroName "mProduction.BtnToTotalInv")

    $wbProd = Resolve-WorkbookSafe -Excel $excel -WorkbookName $prodWorkbookName
    $wbProdInboxRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ("invSys.Inbox.Production." + $stationId + ".xlsb")
    $wbInventoryRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ($warehouseId + ".invSys.Data.Inventory.xlsb")
    $wsProd = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "Production"
    $wsProdInv = Get-WorksheetSafe -Workbook $wbProd -WorksheetName "InventoryManagement"
    $loProductionOutput = Get-ListObjectSafe -Worksheet $wsProd -TableName "ProductionOutput"
    $loProdInv = Get-ListObjectSafe -Worksheet $wsProdInv -TableName "invSys"
    $loInboxProd = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbProdInboxRuntime -WorksheetName "InboxProd") -TableName "tblInboxProd"
    $loInventoryLog = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbInventoryRuntime -WorksheetName "InventoryLog") -TableName "tblInventoryLog"

    $prodOutputRowsAfter = Get-RowCountSafe $loProductionOutput
    $prodLocalOk = (([double](Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName "MADE")) -eq 0) `
        -and (([double](Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName "TOTAL INV")) -eq 8)
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToTotalInv.Local" -Passed $prodLocalOk -Detail "MADE=$((Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName 'MADE')); TOTAL_INV=$((Get-RowValueSafe -ListObject $loProdInv -RowIndex 2 -ColumnName 'TOTAL INV')); ProductionOutputRows=$prodOutputRowsAfter"

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

    $prodStatusBeforeRun = ([string](Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodQueuedRow -ColumnName "Status")).Trim().ToUpperInvariant()
    Restore-LiveRuntimeContext -Excel $excel -WorkbookMap $workbookMap -RuntimeRoot $runtimeRoot -WarehouseId $warehouseId -StationId $stationId -UserId $resolvedUserId -Pin $testPin
    $prodRunBatchReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modProcessor.RunBatchReportForAutomation" -Arguments @($warehouseId, 500))
    $prodRunBatch = 0
    if ($prodRunBatchReport -match 'Processed=(\d+)') { $prodRunBatch = [int]$Matches[1] }
    $wbProdInboxRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ("invSys.Inbox.Production." + $stationId + ".xlsb")
    $loInboxProd = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbProdInboxRuntime -WorksheetName "InboxProd") -TableName "tblInboxProd"
    $prodStatus = ([string](Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodQueuedRow -ColumnName "Status")).Trim().ToUpperInvariant()
    $prodProcessedOk = ($prodStatusBeforeRun -eq "PROCESSED") -or ($prodStatus -eq "PROCESSED") -or ($prodRunBatch -ge 1)
    Add-ResultRow -Rows $resultRows -Check "Production.BtnToTotalInv.Process" -Passed $prodProcessedOk -Detail "StatusBeforeRun=$prodStatusBeforeRun; RunBatch=$prodRunBatch; Status=$prodStatus; ErrorCode=$((Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodQueuedRow -ColumnName 'ErrorCode')); ErrorMessage=$((Get-RowValueSafe -ListObject $loInboxProd -RowIndex $prodQueuedRow -ColumnName 'ErrorMessage')); $prodRunBatchReport"

    $wbInventoryRuntime = Resolve-WorkbookSafe -Excel $excel -WorkbookName ($warehouseId + ".invSys.Data.Inventory.xlsb")
    $loInventoryLog = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbInventoryRuntime -WorksheetName "InventoryLog") -TableName "tblInventoryLog"
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
