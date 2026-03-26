Attribute VB_Name = "TestPhase6LanBoundary"
Option Explicit

Public Function LanBoundarySeedCanonicalRoot(ByVal rootPath As String, _
                                             ByVal publishedRoot As String, _
                                             ByVal warehouseId As String, _
                                             ByVal stationA As String, _
                                             ByVal stationB As String, _
                                             ByVal sku As String) As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInboxA As Workbook
    Dim wbInboxB As Workbook
    Dim currentUser As String

    On Error GoTo FailSeed
    EnsureFolderRecursiveLan rootPath
    If Trim$(publishedRoot) <> "" Then EnsureFolderRecursiveLan publishedRoot

    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook(warehouseId, stationA, rootPath, "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", rootPath
    If Trim$(publishedRoot) <> "" Then TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathSharePointRoot", publishedRoot
    EnsureStationConfigRowLan wbCfg, warehouseId, stationB
    wbCfg.Save

    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook(warehouseId, rootPath)
    currentUser = modRoleEventWriter.ResolveCurrentUserId()
    If currentUser = "" Then currentUser = "user1"
    TestPhase2Helpers.AddCapability wbAuth, currentUser, "RECEIVE_POST", warehouseId, stationA, "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, currentUser, "RECEIVE_POST", warehouseId, stationB, "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", warehouseId, "*", "ACTIVE"
    wbAuth.Save

    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook(warehouseId, rootPath, Array(sku))
    wbInv.Save

    Set wbInboxA = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook(stationA, rootPath)
    Set wbInboxB = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook(stationB, rootPath)
    wbInboxA.Save
    wbInboxB.Save

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Call modConfig.LoadConfig(warehouseId, stationA)
    Call modAuth.LoadAuth(warehouseId)
    CloseWorkbookByNameLan warehouseId & ".invSys.Config.xlsb", True
    CloseWorkbookByNameLan warehouseId & ".invSys.Auth.xlsb", True

    LanBoundarySeedCanonicalRoot = "OK|Root=" & rootPath & "|Published=" & publishedRoot
    Exit Function

FailSeed:
    LanBoundarySeedCanonicalRoot = "ERR|" & Err.Description
End Function

Public Function LanBoundaryAttachToCanonicalRoot(ByVal rootPath As String, _
                                                 ByVal warehouseId As String, _
                                                 ByVal stationId As String) As String
    On Error GoTo FailAttach

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig(warehouseId, stationId) Then
        LanBoundaryAttachToCanonicalRoot = "ERR|Config|" & modConfig.Validate()
        Exit Function
    End If
    If Not modAuth.LoadAuth(warehouseId) Then
        LanBoundaryAttachToCanonicalRoot = "ERR|Auth|" & modAuth.ValidateAuth()
        Exit Function
    End If
    CloseWorkbookByNameLan warehouseId & ".invSys.Config.xlsb", True
    CloseWorkbookByNameLan warehouseId & ".invSys.Auth.xlsb", True

    LanBoundaryAttachToCanonicalRoot = "OK|Warehouse=" & modConfig.GetWarehouseId() & "|Station=" & modConfig.GetStationId()
    Exit Function

FailAttach:
    LanBoundaryAttachToCanonicalRoot = "ERR|" & Err.Description
End Function

Public Function LanBoundaryHoldCanonicalInventory(ByVal warehouseId As String) As String
    Dim wbInv As Workbook

    On Error GoTo FailHold
    Set wbInv = ResolveInventoryWorkbookBridge(warehouseId)
    If wbInv Is Nothing Then
        LanBoundaryHoldCanonicalInventory = "ERR|Inventory workbook not found."
        Exit Function
    End If

    LanBoundaryHoldCanonicalInventory = "OK|Path=" & wbInv.FullName & "|ReadOnly=" & CStr(wbInv.ReadOnly)
    Exit Function

FailHold:
    LanBoundaryHoldCanonicalInventory = "ERR|" & Err.Description
End Function

Public Function LanBoundaryCloseCanonicalInventory(ByVal warehouseId As String) As String
    Dim wbInv As Workbook

    On Error GoTo FailClose
    Set wbInv = FindWorkbookByNameLan(warehouseId & ".invSys.Data.Inventory.xlsb")
    If wbInv Is Nothing Then
        LanBoundaryCloseCanonicalInventory = "OK|AlreadyClosed"
        Exit Function
    End If
    wbInv.Close SaveChanges:=True
    LanBoundaryCloseCanonicalInventory = "OK|Closed"
    Exit Function

FailClose:
    LanBoundaryCloseCanonicalInventory = "ERR|" & Err.Description
End Function

Public Function LanBoundaryQueueAndRunReceive(ByVal warehouseId As String, _
                                              ByVal stationId As String, _
                                              ByVal sku As String, _
                                              ByVal qty As Double, _
                                              ByVal locationVal As String, _
                                              ByVal noteVal As String) As String
    Dim queued As String
    Dim eventIdOut As String

    queued = LanBoundaryQueueReceiveOnly(warehouseId, stationId, sku, qty, locationVal, noteVal)
    If Left$(queued, 3) <> "OK|" Then
        LanBoundaryQueueAndRunReceive = queued
        Exit Function
    End If

    eventIdOut = GetTaggedValueLan(queued, "EventID")
    LanBoundaryQueueAndRunReceive = LanBoundaryRunBatchForEvent(warehouseId, stationId, eventIdOut)
End Function

Public Function LanBoundaryQueueReceiveOnly(ByVal warehouseId As String, _
                                            ByVal stationId As String, _
                                            ByVal sku As String, _
                                            ByVal qty As Double, _
                                            ByVal locationVal As String, _
                                            ByVal noteVal As String) As String
    Dim wbInbox As Workbook
    Dim eventIdOut As String
    Dim errorMessage As String
    Dim currentUser As String

    On Error GoTo FailQueue
    currentUser = modRoleEventWriter.ResolveCurrentUserId()
    If currentUser = "" Then currentUser = "user1"

    Set wbInbox = modRoleEventWriter.OpenInboxWorkbook(CORE_EVENT_TYPE_RECEIVE, warehouseId, stationId, errorMessage)
    If wbInbox Is Nothing Then
        LanBoundaryQueueReceiveOnly = "ERR|OpenInbox|" & errorMessage
        Exit Function
    End If

    If Not modRoleEventWriter.QueueReceiveEvent(warehouseId, stationId, currentUser, sku, qty, locationVal, noteVal, "", "", Now, wbInbox, eventIdOut, errorMessage) Then
        LanBoundaryQueueReceiveOnly = "ERR|Queue|" & errorMessage
        CloseConfigAndAuthLan warehouseId
        Exit Function
    End If

    LanBoundaryQueueReceiveOnly = "OK|EventID=" & eventIdOut
    CloseConfigAndAuthLan warehouseId
    Exit Function

FailQueue:
    CloseConfigAndAuthLan warehouseId
    LanBoundaryQueueReceiveOnly = "ERR|" & Err.Description
End Function

Public Function LanBoundaryRunBatchForEvent(ByVal warehouseId As String, _
                                            ByVal stationId As String, _
                                            ByVal eventIdOut As String) As String
    Dim wbInbox As Workbook
    Dim errorMessage As String
    Dim report As String
    Dim processedCount As Long
    Dim loInbox As ListObject
    Dim rowIndex As Long

    On Error GoTo FailRunBatch
    Set wbInbox = modRoleEventWriter.OpenInboxWorkbook(CORE_EVENT_TYPE_RECEIVE, warehouseId, stationId, errorMessage)
    If wbInbox Is Nothing Then
        LanBoundaryRunBatchForEvent = "ERR|OpenInbox|" & errorMessage
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 500, report)
    Set loInbox = FindTableByNameLan(wbInbox, "tblInboxReceive")
    If loInbox Is Nothing Then
        LanBoundaryRunBatchForEvent = "ERR|InboxTableMissing"
        Exit Function
    End If
    rowIndex = FindRowByColumnValueLan(loInbox, "EventID", eventIdOut)
    If rowIndex = 0 Then
        LanBoundaryRunBatchForEvent = "ERR|InboxRowMissing|EventID=" & eventIdOut
        Exit Function
    End If

    LanBoundaryRunBatchForEvent = "OK|EventID=" & eventIdOut & _
        "|Processed=" & CStr(processedCount) & _
        "|Report=" & EscapePipeLan(report) & _
        "|Status=" & EscapePipeLan(CStr(GetTableValueLan(loInbox, rowIndex, "Status"))) & _
        "|ErrorCode=" & EscapePipeLan(CStr(GetTableValueLan(loInbox, rowIndex, "ErrorCode"))) & _
        "|ErrorMessage=" & EscapePipeLan(CStr(GetTableValueLan(loInbox, rowIndex, "ErrorMessage")))
    CloseConfigAndAuthLan warehouseId
    Exit Function

FailRunBatch:
    CloseConfigAndAuthLan warehouseId
    LanBoundaryRunBatchForEvent = "ERR|" & Err.Description
End Function

Public Function LanBoundaryPublishCurrentSnapshot(ByVal warehouseId As String, ByVal publishedRoot As String) As String
    Dim report As String
    Dim sourcePath As String
    Dim targetPath As String

    On Error GoTo FailPublish
    If Not modWarehouseSync.GenerateWarehouseSnapshot(warehouseId, Nothing, "", Nothing, report) Then
        LanBoundaryPublishCurrentSnapshot = "ERR|Snapshot|" & report
        Exit Function
    End If

    sourcePath = report
    If Trim$(sourcePath) = "" Then
        LanBoundaryPublishCurrentSnapshot = "ERR|SnapshotPathMissing"
        Exit Function
    End If

    EnsureFolderRecursiveLan publishedRoot
    targetPath = NormalizeFolderPathLan(publishedRoot) & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    CloseWorkbookByPathLan sourcePath, True
    CloseWorkbookByPathLan targetPath, False
    On Error Resume Next
    Kill targetPath
    On Error GoTo FailPublish
    FileCopy sourcePath, targetPath

    LanBoundaryPublishCurrentSnapshot = "OK|PublishedPath=" & targetPath
    Exit Function

FailPublish:
    LanBoundaryPublishCurrentSnapshot = "ERR|" & Err.Description
End Function

Public Function LanBoundaryBuildSavedReceivingOperator(ByVal operatorPath As String, _
                                                       ByVal sku As String, _
                                                       ByVal refNumber As String, _
                                                       ByVal snapshotLogId As String, _
                                                       ByVal totalInv As Double, _
                                                       ByVal locationVal As String) As String
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject

    On Error GoTo FailBuild
    EnsureFolderRecursiveLan GetParentFolderLan(operatorPath)

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then
        LanBoundaryBuildSavedReceivingOperator = "ERR|" & report
        Exit Function
    End If

    Set loInv = FindTableByNameLan(wb, "invSys")
    Set loRecv = FindTableByNameLan(wb, "ReceivedTally")
    Set loLog = FindTableByNameLan(wb, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then
        LanBoundaryBuildSavedReceivingOperator = "ERR|Saved operator tables missing."
        Exit Function
    End If

    AddInvSysSeedRowLan loInv, 999, sku, "LAN Boundary Item", "EA", locationVal, totalInv
    AddReceivedTallyRowLan loRecv, refNumber, "LAN Boundary Item", 1, 999
    AddReceivedLogRowLan loLog, snapshotLogId, refNumber, "LAN Boundary Item", 1, "EA", "Vendor", locationVal, sku, 999

    wb.SaveAs Filename:=operatorPath, FileFormat:=50
    wb.Close SaveChanges:=False
    LanBoundaryBuildSavedReceivingOperator = "OK|OperatorPath=" & operatorPath
    Exit Function

FailBuild:
    LanBoundaryBuildSavedReceivingOperator = "ERR|" & Err.Description
End Function

Public Function LanBoundaryRefreshSavedOperatorFromRoot(ByVal operatorPath As String, _
                                                        ByVal warehouseId As String, _
                                                        ByVal rootPath As String, _
                                                        ByVal sourceType As String) As String
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim rowIndex As Long

    On Error GoTo FailRefresh
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wb = OpenWorkbookByPathLan(operatorPath)
    If wb Is Nothing Then
        LanBoundaryRefreshSavedOperatorFromRoot = "ERR|OperatorOpenFailed"
        Exit Function
    End If

    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then
        LanBoundaryRefreshSavedOperatorFromRoot = "ERR|Surface|" & report
        Exit Function
    End If
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wb, warehouseId, sourceType, report) Then
        LanBoundaryRefreshSavedOperatorFromRoot = "ERR|Refresh|" & report
        Exit Function
    End If

    Set loInv = FindTableByNameLan(wb, "invSys")
    If loInv Is Nothing Then
        LanBoundaryRefreshSavedOperatorFromRoot = "ERR|invSysMissing"
        Exit Function
    End If
    rowIndex = 1
    LanBoundaryRefreshSavedOperatorFromRoot = "OK|TotalInv=" & CStr(GetTableValueLan(loInv, rowIndex, "TOTAL INV")) & _
        "|QtyAvailable=" & CStr(GetTableValueLan(loInv, rowIndex, "QtyAvailable")) & _
        "|SnapshotId=" & EscapePipeLan(CStr(GetTableValueLan(loInv, rowIndex, "SnapshotId"))) & _
        "|SourceType=" & EscapePipeLan(CStr(GetTableValueLan(loInv, rowIndex, "SourceType"))) & _
        "|IsStale=" & EscapePipeLan(CStr(GetTableValueLan(loInv, rowIndex, "IsStale"))) & _
        "|Path=" & wb.FullName
    Exit Function

FailRefresh:
    LanBoundaryRefreshSavedOperatorFromRoot = "ERR|" & Err.Description
End Function

Private Sub EnsureStationConfigRowLan(ByVal wbCfg As Workbook, ByVal warehouseId As String, ByVal stationId As String)
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim lr As ListRow

    If wbCfg Is Nothing Then Exit Sub
    Set lo = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    rowIndex = FindRowByColumnValueLan(lo, "StationId", stationId)
    If rowIndex = 0 Then
        Set lr = lo.ListRows.Add
        rowIndex = lr.Index
    End If
    SetTableCellLan lo, rowIndex, "StationId", stationId
    SetTableCellLan lo, rowIndex, "WarehouseId", warehouseId
    SetTableCellLan lo, rowIndex, "StationName", stationId
    SetTableCellLan lo, rowIndex, "RoleDefault", "RECEIVE"
End Sub

Private Function OpenWorkbookByPathLan(ByVal targetPath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            Set OpenWorkbookByPathLan = wb
            Exit Function
        End If
    Next wb
    If Len(Dir$(targetPath)) > 0 Then Set OpenWorkbookByPathLan = Application.Workbooks.Open(targetPath)
End Function

Private Sub CloseWorkbookByPathLan(ByVal targetPath As String, ByVal saveChanges As Boolean)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=saveChanges
            Exit For
        End If
    Next wb
End Sub

Private Sub CloseWorkbookByNameLan(ByVal workbookName As String, ByVal saveChanges As Boolean)
    Dim wb As Workbook

    Set wb = FindWorkbookByNameLan(workbookName)
    If wb Is Nothing Then Exit Sub
    wb.Close SaveChanges:=saveChanges
End Sub

Private Sub CloseConfigAndAuthLan(ByVal warehouseId As String)
    CloseWorkbookByNameLan warehouseId & ".invSys.Config.xlsb", False
    CloseWorkbookByNameLan warehouseId & ".invSys.Auth.xlsb", False
End Sub

Private Function FindWorkbookByNameLan(ByVal workbookName As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            Set FindWorkbookByNameLan = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindTableByNameLan(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTableByNameLan = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTableByNameLan Is Nothing Then Exit Function
    Next ws
End Function

Private Function FindRowByColumnValueLan(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(GetTableValueLan(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByColumnValueLan = i
            Exit Function
        End If
    Next i
End Function

Private Function GetTableValueLan(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    GetTableValueLan = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value
End Function

Private Sub SetTableCellLan(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueIn
End Sub

Private Sub AddInvSysSeedRowLan(ByVal lo As ListObject, ByVal rowValue As Long, ByVal sku As String, ByVal itemName As String, ByVal uom As String, ByVal locationVal As String, ByVal totalInv As Double)
    Dim lr As ListRow

    Set lr = lo.ListRows.Add
    SetTableCellLan lo, lr.Index, "ROW", rowValue
    SetTableCellLan lo, lr.Index, "ITEM_CODE", sku
    SetTableCellLan lo, lr.Index, "ITEM", itemName
    SetTableCellLan lo, lr.Index, "UOM", uom
    SetTableCellLan lo, lr.Index, "LOCATION", locationVal
    SetTableCellLan lo, lr.Index, "TOTAL INV", totalInv
End Sub

Private Sub AddReceivedTallyRowLan(ByVal lo As ListObject, ByVal refNumber As String, ByVal itemName As String, ByVal qty As Double, ByVal rowValue As Long)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf Trim$(CStr(GetTableValueLan(lo, 1, "REF_NUMBER"))) = "" _
        And Trim$(CStr(GetTableValueLan(lo, 1, "ITEMS"))) = "" _
        And NzDblLan(GetTableValueLan(lo, 1, "QUANTITY")) = 0 Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCellLan lo, lr.Index, "REF_NUMBER", refNumber
    SetTableCellLan lo, lr.Index, "ITEMS", itemName
    SetTableCellLan lo, lr.Index, "QUANTITY", qty
    SetTableCellLan lo, lr.Index, "ROW", rowValue
End Sub

Private Sub AddReceivedLogRowLan(ByVal lo As ListObject, _
                                 ByVal snapshotId As String, _
                                 ByVal refNumber As String, _
                                 ByVal itemName As String, _
                                 ByVal qty As Double, _
                                 ByVal uom As String, _
                                 ByVal vendorName As String, _
                                 ByVal locationVal As String, _
                                 ByVal sku As String, _
    ByVal rowValue As Long)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValueLan(lo, 1, "SNAPSHOT_ID"))) = "" _
        And Trim$(CStr(GetTableValueLan(lo, 1, "REF_NUMBER"))) = "" _
        And NzDblLan(GetTableValueLan(lo, 1, "QUANTITY")) = 0 Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCellLan lo, lr.Index, "SNAPSHOT_ID", snapshotId
    SetTableCellLan lo, lr.Index, "ENTRY_DATE", CDate("2026-03-25 08:00:00")
    SetTableCellLan lo, lr.Index, "REF_NUMBER", refNumber
    SetTableCellLan lo, lr.Index, "ITEMS", itemName
    SetTableCellLan lo, lr.Index, "QUANTITY", qty
    SetTableCellLan lo, lr.Index, "UOM", uom
    SetTableCellLan lo, lr.Index, "VENDOR", vendorName
    SetTableCellLan lo, lr.Index, "LOCATION", locationVal
    SetTableCellLan lo, lr.Index, "ITEM_CODE", sku
    SetTableCellLan lo, lr.Index, "ROW", rowValue
End Sub

Private Function GetParentFolderLan(ByVal pathIn As String) As String
    Dim sepPos As Long

    sepPos = InStrRev(pathIn, "\")
    If sepPos > 1 Then GetParentFolderLan = Left$(pathIn, sepPos - 1)
End Function

Private Function NormalizeFolderPathLan(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathLan = folderPath
End Function

Private Sub EnsureFolderRecursiveLan(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 1 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveLan parentPath
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function EscapePipeLan(ByVal textIn As String) As String
    EscapePipeLan = Replace$(Replace$(textIn, "|", "/"), vbCrLf, " ")
End Function

Private Function NzDblLan(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If Trim$(CStr(valueIn)) = "" Then Exit Function
    If IsNumeric(valueIn) Then NzDblLan = CDbl(valueIn)
End Function

Private Function GetTaggedValueLan(ByVal encoded As String, ByVal tagName As String) As String
    Dim parts() As String
    Dim i As Long
    Dim prefix As String

    prefix = tagName & "="
    parts = Split(encoded, "|")
    For i = LBound(parts) To UBound(parts)
        If StrComp(Left$(parts(i), Len(prefix)), prefix, vbTextCompare) = 0 Then
            GetTaggedValueLan = Mid$(parts(i), Len(prefix) + 1)
            Exit Function
        End If
    Next i
End Function
