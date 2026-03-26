Attribute VB_Name = "TestPhase5HqBoundary"
Option Explicit

Public Function HqBoundarySeedWarehouseRoot(ByVal rootPath As String, _
                                            ByVal shareRoot As String, _
                                            ByVal warehouseId As String, _
                                            ByVal stationId As String, _
                                            ByVal sku As String) As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook

    On Error GoTo FailSeed
    EnsureFolderRecursiveHqBoundary rootPath
    EnsureFolderRecursiveHqBoundary shareRoot
    EnsureFolderRecursiveHqBoundary NormalizeFolderPathHqBoundary(shareRoot) & "Snapshots"
    EnsureFolderRecursiveHqBoundary NormalizeFolderPathHqBoundary(shareRoot) & "Global"

    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook(warehouseId, stationId, rootPath, "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", rootPath
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathSharePointRoot", shareRoot
    wbCfg.Save

    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook(warehouseId, rootPath)
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", warehouseId, stationId, "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", warehouseId, "*", "ACTIVE"
    wbAuth.Save

    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook(warehouseId, rootPath, Array(sku))
    wbInv.Save

    Set wbInbox = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook(stationId, rootPath)
    wbInbox.Save

    HqBoundarySeedWarehouseRoot = "OK|Warehouse=" & warehouseId & "|Station=" & stationId
    Exit Function

FailSeed:
    HqBoundarySeedWarehouseRoot = "ERR|" & Err.Description
End Function

Public Function HqBoundaryWarehouseRunAndPublish(ByVal rootPath As String, _
                                                 ByVal shareRoot As String, _
                                                 ByVal warehouseId As String, _
                                                 ByVal stationId As String, _
                                                 ByVal sku As String, _
                                                 ByVal qty As Double, _
                                                 ByVal locationVal As String, _
                                                 ByVal noteVal As String) As String
    Dim wbInbox As Workbook
    Dim report As String
    Dim processedCount As Long
    Dim eventId As String
    Dim localSnapshotPath As String
    Dim publishedSnapshotPath As String

    On Error GoTo FailRun
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig(warehouseId, stationId) Then
        HqBoundaryWarehouseRunAndPublish = "ERR|Config|" & modConfig.Validate()
        Exit Function
    End If
    If Not modAuth.LoadAuth(warehouseId) Then
        HqBoundaryWarehouseRunAndPublish = "ERR|Auth|" & modAuth.ValidateAuth()
        Exit Function
    End If

    Set wbInbox = OpenWorkbookByPathHqBoundary(NormalizeFolderPathHqBoundary(rootPath) & "invSys.Inbox.Receiving." & stationId & ".xlsb")
    If wbInbox Is Nothing Then
        HqBoundaryWarehouseRunAndPublish = "ERR|InboxOpen"
        CloseConfigAndAuthHqBoundary warehouseId
        Exit Function
    End If

    eventId = BuildBoundaryEventId(warehouseId)
    TestPhase2Helpers.AddInboxReceiveRow wbInbox, eventId, Now, warehouseId, stationId, "user1", sku, qty, locationVal, noteVal
    processedCount = modProcessor.RunBatch(warehouseId, 500, report)

    localSnapshotPath = NormalizeFolderPathHqBoundary(rootPath) & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    publishedSnapshotPath = NormalizeFolderPathHqBoundary(shareRoot) & "Snapshots\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    CloseWorkbookByPathHqBoundary localSnapshotPath, True
    CopyFileReplacingHqBoundary localSnapshotPath, publishedSnapshotPath
    CloseConfigAndAuthHqBoundary warehouseId

    HqBoundaryWarehouseRunAndPublish = "OK|EventID=" & eventId & _
        "|Processed=" & CStr(processedCount) & _
        "|Report=" & EscapePipeHqBoundary(report) & _
        "|PublishedPath=" & publishedSnapshotPath
    Exit Function

FailRun:
    CloseConfigAndAuthHqBoundary warehouseId
    HqBoundaryWarehouseRunAndPublish = "ERR|" & Err.Description
End Function

Public Function HqBoundaryRunAggregatorAndRead(ByVal shareRoot As String, _
                                               ByVal warehouseA As String, _
                                               ByVal warehouseB As String, _
                                               ByVal sku As String) As String
    Dim report As String
    Dim wbGlobal As Workbook
    Dim loGlobal As ListObject
    Dim loStatus As ListObject
    Dim rowA As Long
    Dim rowB As Long
    Dim globalPath As String

    On Error GoTo FailAggregate
    If Not modHqAggregator.RunHQAggregation(shareRoot, "", report) Then
        HqBoundaryRunAggregatorAndRead = "ERR|Aggregate|" & report
        Exit Function
    End If

    globalPath = NormalizeFolderPathHqBoundary(shareRoot) & "Global\invSys.Global.InventorySnapshot.xlsb"
    Set wbGlobal = OpenWorkbookByPathHqBoundary(globalPath)
    If wbGlobal Is Nothing Then
        HqBoundaryRunAggregatorAndRead = "ERR|GlobalOpen"
        Exit Function
    End If

    Set loGlobal = wbGlobal.Worksheets("GlobalInventorySnapshot").ListObjects("tblGlobalInventorySnapshot")
    Set loStatus = wbGlobal.Worksheets("GlobalSnapshotStatus").ListObjects("tblGlobalSnapshotStatus")
    rowA = FindWarehouseSkuRowHqBoundary(loGlobal, warehouseA, sku)
    rowB = FindWarehouseSkuRowHqBoundary(loGlobal, warehouseB, sku)
    If rowA = 0 Or rowB = 0 Then
        HqBoundaryRunAggregatorAndRead = "ERR|RowsMissing"
        Exit Function
    End If

    HqBoundaryRunAggregatorAndRead = "OK|Report=" & EscapePipeHqBoundary(report) & _
        "|QtyA=" & CStr(TestPhase2Helpers.GetRowValue(loGlobal, rowA, "QtyOnHand")) & _
        "|QtyB=" & CStr(TestPhase2Helpers.GetRowValue(loGlobal, rowB, "QtyOnHand")) & _
        "|SourceA=" & EscapePipeHqBoundary(CStr(TestPhase2Helpers.GetRowValue(loGlobal, rowA, "SourceSnapshot"))) & _
        "|SourceB=" & EscapePipeHqBoundary(CStr(TestPhase2Helpers.GetRowValue(loGlobal, rowB, "SourceSnapshot"))) & _
        "|Skipped=" & CStr(TestPhase2Helpers.GetRowValue(loStatus, 1, "SkippedSnapshotFileCount")) & _
        "|Warehouses=" & CStr(TestPhase2Helpers.GetRowValue(loStatus, 1, "WarehouseCount"))
    Exit Function

FailAggregate:
    HqBoundaryRunAggregatorAndRead = "ERR|" & Err.Description
End Function

Private Function BuildBoundaryEventId(ByVal warehouseId As String) As String
    Randomize
    BuildBoundaryEventId = "EVT-" & warehouseId & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(CLng(Rnd() * 1000000), "000000")
End Function

Private Function OpenWorkbookByPathHqBoundary(ByVal targetPath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            Set OpenWorkbookByPathHqBoundary = wb
            Exit Function
        End If
    Next wb
    If Len(Dir$(targetPath)) > 0 Then Set OpenWorkbookByPathHqBoundary = Application.Workbooks.Open(targetPath)
End Function

Private Sub CloseWorkbookByPathHqBoundary(ByVal targetPath As String, ByVal saveChanges As Boolean)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=saveChanges
            Exit For
        End If
    Next wb
End Sub

Private Sub CloseConfigAndAuthHqBoundary(ByVal warehouseId As String)
    CloseWorkbookByPathHqBoundary NormalizeFolderPathHqBoundary(modRuntimeWorkbooks.GetCoreDataRootOverride()) & warehouseId & ".invSys.Config.xlsb", False
    CloseWorkbookByPathHqBoundary NormalizeFolderPathHqBoundary(modRuntimeWorkbooks.GetCoreDataRootOverride()) & warehouseId & ".invSys.Auth.xlsb", False
End Sub

Private Function FindWarehouseSkuRowHqBoundary(ByVal lo As ListObject, ByVal warehouseId As String, ByVal sku As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "WarehouseId")), warehouseId, vbTextCompare) = 0 And _
           StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "SKU")), sku, vbTextCompare) = 0 Then
            FindWarehouseSkuRowHqBoundary = i
            Exit Function
        End If
    Next i
End Function

Private Sub CopyFileReplacingHqBoundary(ByVal sourcePath As String, ByVal targetPath As String)
    On Error Resume Next
    Kill targetPath
    On Error GoTo 0
    FileCopy sourcePath, targetPath
End Sub

Private Function NormalizeFolderPathHqBoundary(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathHqBoundary = folderPath
End Function

Private Sub EnsureFolderRecursiveHqBoundary(ByVal folderPath As String)
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
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveHqBoundary parentPath
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function EscapePipeHqBoundary(ByVal textIn As String) As String
    EscapePipeHqBoundary = Replace$(Replace$(textIn, "|", "/"), vbCrLf, " ")
End Function
