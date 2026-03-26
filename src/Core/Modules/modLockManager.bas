Attribute VB_Name = "modLockManager"
Option Explicit

Public Const LOCK_STATUS_HELD As String = "HELD"
Public Const LOCK_STATUS_EXPIRED As String = "EXPIRED"
Public Const LOCK_STATUS_BROKEN As String = "BROKEN"

Public Function AcquireLock(ByVal lockName As String, _
                            Optional ByVal warehouseId As String = "", _
                            Optional ByVal ownerUserId As String = "", _
                            Optional ByVal ownerStationId As String = "", _
                            Optional ByVal inventoryWb As Workbook = Nothing, _
                            Optional ByRef runId As String = "", _
                            Optional ByRef message As String = "") As Boolean
    On Error GoTo FailAcquire

    Dim wb As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim expiresAt As Variant
    Dim existingRunId As String
    Dim nowTs As Date

    lockName = UCase$(SafeTrimLock(lockName))
    If lockName = "" Then
        message = "Lock name is required."
        Exit Function
    End If

    Set wb = ResolveInventoryWorkbookLock(warehouseId, inventoryWb)
    If wb Is Nothing Then
        message = "Inventory workbook not found."
        Exit Function
    End If
    If wb.ReadOnly Then
        message = "Inventory workbook is read-only or locked by another Excel session."
        Exit Function
    End If

    If Not EnsureInventorySchemaBridge(wb) Then
        message = "Unable to validate inventory schema."
        Exit Function
    End If

    Set lo = FindListObjectByNameLock(wb, "tblLocks")
    If lo Is Nothing Then
        message = "tblLocks not found."
        Exit Function
    End If

    SetSheetProtectionLock lo.Parent, False

    rowIndex = FindLockRow(lo, lockName)
    If rowIndex = 0 Then rowIndex = FindReusableBlankRow(lo)
    If rowIndex = 0 Then rowIndex = EnsureLockRow(lo, lockName)

    existingRunId = SafeTrimLock(GetCellByColumnLock(lo, rowIndex, "RunId"))
    expiresAt = GetCellByColumnLock(lo, rowIndex, "ExpiresAtUTC")

    If IsHeldAndActive(lo, rowIndex) Then
        If runId <> "" And StrComp(existingRunId, runId, vbTextCompare) = 0 Then
            AcquireLock = UpdateHeartbeat(lockName, runId, wb)
            If AcquireLock Then message = "Lock refreshed."
            Exit Function
        End If
        message = "Lock already held by another processor."
        Exit Function
    End If

    If warehouseId = "" Then warehouseId = modConfig.GetString("WarehouseId", "")
    If ownerUserId = "" Then ownerUserId = modConfig.GetString("ProcessorServiceUserId", "svc_processor")
    If ownerStationId = "" Then ownerStationId = modConfig.GetString("StationId", "")
    If runId = "" Then runId = CreateRunIdLock(lockName, warehouseId)

    nowTs = Now
    SetCellByColumnLock lo, rowIndex, "LockName", lockName
    SetCellByColumnLock lo, rowIndex, "OwnerStationId", ownerStationId
    SetCellByColumnLock lo, rowIndex, "OwnerUserId", ownerUserId
    SetCellByColumnLock lo, rowIndex, "RunId", runId
    SetCellByColumnLock lo, rowIndex, "AcquiredAtUTC", nowTs
    SetCellByColumnLock lo, rowIndex, "ExpiresAtUTC", DateAdd("n", GetLockTimeoutMinutes(), nowTs)
    SetCellByColumnLock lo, rowIndex, "HeartbeatAtUTC", nowTs
    SetCellByColumnLock lo, rowIndex, "Status", LOCK_STATUS_HELD

    SaveLockWorkbookIfWritable wb
    AcquireLock = True
    message = "Lock acquired."
    SetSheetProtectionLock lo.Parent, True
    Exit Function

FailAcquire:
    On Error Resume Next
    If Not lo Is Nothing Then SetSheetProtectionLock lo.Parent, True
    On Error GoTo 0
    message = "AcquireLock failed: " & Err.Description
End Function

Public Function UpdateHeartbeat(ByVal lockName As String, _
                                Optional ByVal runId As String = "", _
                                Optional ByVal inventoryWb As Workbook = Nothing) As Boolean
    On Error GoTo FailHeartbeat

    Dim wb As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim nowTs As Date
    Dim existingRunId As String

    lockName = UCase$(SafeTrimLock(lockName))
    If lockName = "" Then Exit Function

    Set wb = ResolveInventoryWorkbookLock(modConfig.GetString("WarehouseId", ""), inventoryWb)
    If wb Is Nothing Then Exit Function
    If Not EnsureInventorySchemaBridge(wb) Then Exit Function

    Set lo = FindListObjectByNameLock(wb, "tblLocks")
    If lo Is Nothing Then Exit Function

    SetSheetProtectionLock lo.Parent, False

    rowIndex = FindLockRow(lo, lockName)
    If rowIndex = 0 Then Exit Function

    existingRunId = SafeTrimLock(GetCellByColumnLock(lo, rowIndex, "RunId"))
    If runId <> "" Then
        If StrComp(existingRunId, runId, vbTextCompare) <> 0 Then Exit Function
    End If

    nowTs = Now
    SetCellByColumnLock lo, rowIndex, "HeartbeatAtUTC", nowTs
    SetCellByColumnLock lo, rowIndex, "ExpiresAtUTC", DateAdd("n", GetLockTimeoutMinutes(), nowTs)
    SetCellByColumnLock lo, rowIndex, "Status", LOCK_STATUS_HELD
    SaveLockWorkbookIfWritable wb
    UpdateHeartbeat = True
    SetSheetProtectionLock lo.Parent, True
    Exit Function

FailHeartbeat:
    On Error Resume Next
    If Not lo Is Nothing Then SetSheetProtectionLock lo.Parent, True
    On Error GoTo 0
    UpdateHeartbeat = False
End Function

Public Function ReleaseLock(ByVal lockName As String, _
                            Optional ByVal runId As String = "", _
                            Optional ByVal inventoryWb As Workbook = Nothing) As Boolean
    On Error GoTo FailRelease

    Dim wb As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim existingRunId As String
    Dim nowTs As Date

    lockName = UCase$(SafeTrimLock(lockName))
    If lockName = "" Then Exit Function

    Set wb = ResolveInventoryWorkbookLock(modConfig.GetString("WarehouseId", ""), inventoryWb)
    If wb Is Nothing Then Exit Function
    If Not EnsureInventorySchemaBridge(wb) Then Exit Function

    Set lo = FindListObjectByNameLock(wb, "tblLocks")
    If lo Is Nothing Then Exit Function

    SetSheetProtectionLock lo.Parent, False

    rowIndex = FindLockRow(lo, lockName)
    If rowIndex = 0 Then Exit Function

    existingRunId = SafeTrimLock(GetCellByColumnLock(lo, rowIndex, "RunId"))
    If runId <> "" Then
        If existingRunId <> "" And StrComp(existingRunId, runId, vbTextCompare) <> 0 Then Exit Function
    End If

    nowTs = Now
    SetCellByColumnLock lo, rowIndex, "ExpiresAtUTC", nowTs
    SetCellByColumnLock lo, rowIndex, "HeartbeatAtUTC", nowTs
    SetCellByColumnLock lo, rowIndex, "Status", LOCK_STATUS_EXPIRED
    SaveLockWorkbookIfWritable wb
    ReleaseLock = True
    SetSheetProtectionLock lo.Parent, True
    Exit Function

FailRelease:
    On Error Resume Next
    If Not lo Is Nothing Then SetSheetProtectionLock lo.Parent, True
    On Error GoTo 0
    ReleaseLock = False
End Function

Public Function BreakLock(ByVal lockName As String, _
                          Optional ByVal warehouseId As String = "", _
                          Optional ByVal breakerUserId As String = "", _
                          Optional ByVal reason As String = "", _
                          Optional ByVal inventoryWb As Workbook = Nothing, _
                          Optional ByRef message As String = "") As Boolean
    On Error GoTo FailBreak

    Dim wb As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim nowTs As Date

    lockName = UCase$(SafeTrimLock(lockName))
    If lockName = "" Then
        message = "Lock name is required."
        Exit Function
    End If

    Set wb = ResolveInventoryWorkbookLock(warehouseId, inventoryWb)
    If wb Is Nothing Then
        message = "Inventory workbook not found."
        Exit Function
    End If
    If Not EnsureInventorySchemaBridge(wb) Then
        message = "Unable to validate inventory schema."
        Exit Function
    End If

    Set lo = FindListObjectByNameLock(wb, "tblLocks")
    If lo Is Nothing Then
        message = "tblLocks not found."
        Exit Function
    End If

    SetSheetProtectionLock lo.Parent, False
    rowIndex = FindLockRow(lo, lockName)
    If rowIndex = 0 Then
        message = "Lock not found."
        GoTo CleanExit
    End If

    nowTs = Now
    SetCellByColumnLock lo, rowIndex, "HeartbeatAtUTC", nowTs
    SetCellByColumnLock lo, rowIndex, "ExpiresAtUTC", nowTs
    SetCellByColumnLock lo, rowIndex, "Status", LOCK_STATUS_BROKEN
    If breakerUserId <> "" Then SetCellByColumnLock lo, rowIndex, "OwnerUserId", breakerUserId

    SaveLockWorkbookIfWritable wb
    BreakLock = True
    If reason <> "" Then
        message = "Lock broken: " & reason
    Else
        message = "Lock broken."
    End If

CleanExit:
    On Error Resume Next
    If Not lo Is Nothing Then SetSheetProtectionLock lo.Parent, True
    On Error GoTo 0
    Exit Function

FailBreak:
    On Error Resume Next
    If Not lo Is Nothing Then SetSheetProtectionLock lo.Parent, True
    On Error GoTo 0
    message = "BreakLock failed: " & Err.Description
End Function

Private Function ResolveInventoryWorkbookLock(ByVal warehouseId As String, ByVal inventoryWb As Workbook) As Workbook
    If Not inventoryWb Is Nothing Then
        Set ResolveInventoryWorkbookLock = inventoryWb
    Else
        Set ResolveInventoryWorkbookLock = ResolveInventoryWorkbookBridge(warehouseId)
    End If
End Function

Private Function GetLockTimeoutMinutes() As Long
    GetLockTimeoutMinutes = modConfig.GetLong("LockTimeoutMinutes", 3)
    If GetLockTimeoutMinutes <= 0 Then GetLockTimeoutMinutes = 3
End Function

Private Function IsHeldAndActive(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim statusVal As String
    Dim expiresAt As Variant

    statusVal = UCase$(SafeTrimLock(GetCellByColumnLock(lo, rowIndex, "Status")))
    expiresAt = GetCellByColumnLock(lo, rowIndex, "ExpiresAtUTC")

    If statusVal <> LOCK_STATUS_HELD Then Exit Function
    If Not IsDate(expiresAt) Then Exit Function

    IsHeldAndActive = (CDate(expiresAt) > Now)
End Function

Private Function EnsureLockRow(ByVal lo As ListObject, ByVal lockName As String) As Long
    Dim r As ListRow
    Set r = lo.ListRows.Add
    EnsureLockRow = r.Index
    SetCellByColumnLock lo, EnsureLockRow, "LockName", lockName
    SetCellByColumnLock lo, EnsureLockRow, "Status", LOCK_STATUS_EXPIRED
End Function

Private Function FindLockRow(ByVal lo As ListObject, ByVal lockName As String) As Long
    Dim i As Long
    Dim currentName As String

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        currentName = UCase$(SafeTrimLock(GetCellByColumnLock(lo, i, "LockName")))
        If currentName = UCase$(lockName) Then
            FindLockRow = i
            Exit Function
        End If
    Next i
End Function

Private Function FindReusableBlankRow(ByVal lo As ListObject) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If SafeTrimLock(GetCellByColumnLock(lo, i, "LockName")) = "" And _
           SafeTrimLock(GetCellByColumnLock(lo, i, "RunId")) = "" Then
            FindReusableBlankRow = i
            Exit Function
        End If
    Next i
End Function

Private Function CreateRunIdLock(ByVal lockName As String, ByVal warehouseId As String) As String
    Randomize
    CreateRunIdLock = "RUN-" & IIf(warehouseId = "", "WHX", warehouseId) & "-" & _
                      UCase$(lockName) & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
                      Format$(CLng(Rnd() * 1000000), "000000")
End Function

Private Function GetCellByColumnLock(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexLock(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnLock = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Sub SetCellByColumnLock(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexLock(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetColumnIndexLock(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexLock = i
            Exit Function
        End If
    Next i
End Function

Private Function FindListObjectByNameLock(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameLock = ws.ListObjects(tableName)
        If Not FindListObjectByNameLock Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function SafeTrimLock(ByVal v As Variant) As String
    On Error Resume Next
    SafeTrimLock = Trim$(CStr(v))
End Function

Private Sub SetSheetProtectionLock(ByVal ws As Worksheet, ByVal protectAfter As Boolean)
    If ws Is Nothing Then Exit Sub
    If protectAfter Then
        On Error Resume Next
        ws.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    Else
        If Not ws.ProtectContents Then Exit Sub
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
        If ws.ProtectContents Then
            Err.Raise vbObjectError + 2301, "modLockManager.SetSheetProtectionLock", _
                      "Worksheet '" & ws.Name & "' is protected and could not be unprotected. " & _
                      "Excel automation cannot update tblLocks while the sheet remains protected."
        End If
    End If
End Sub

Private Sub SaveLockWorkbookIfWritable(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If Trim$(wb.Path) = "" Then Exit Sub
    wb.Save
End Sub
