Attribute VB_Name = "TestPhase5Sync"
Option Explicit

Public Sub RunPhase5SyncTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestRunBatch_WritesOutboxAndSnapshot(), passed, failed
    Tally TestManualCopy_PublishesWarehouseArtifacts(), passed, failed
    Tally TestHqAggregation_TwoWarehousesPreservesPerWarehouseQty(), passed, failed

    Debug.Print "Phase 5 sync tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestRunBatch_WritesOutboxAndSnapshot() As Long
    Dim tempRoot As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim wbOutbox As Workbook
    Dim wbSnap As Workbook
    Dim loOutbox As ListObject
    Dim loSnap As ListObject
    Dim report As String

    On Error GoTo CleanFail
    tempRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase5Local")
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHS5", "S1", "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", tempRoot
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHS5")
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHS5", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WHS5", "*", "ACTIVE"
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHS5", tempRoot, Array("SKU-001"))
    Set wbInbox = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook("S1", tempRoot)
    TestPhase2Helpers.AddInboxReceiveRow wbInbox, "EVT-P5-001", Now, "WHS5", "S1", "user1", "SKU-001", 9, "A1", "phase5"

    If modProcessor.RunBatch("WHS5", 500, report) <> 1 Then GoTo CleanExit

    Set wbOutbox = OpenWorkbookIfNeeded(tempRoot & "\WHS5.Outbox.Events.xlsb")
    Set wbSnap = OpenWorkbookIfNeeded(tempRoot & "\WHS5.invSys.Snapshot.Inventory.xlsb")
    If wbOutbox Is Nothing Or wbSnap Is Nothing Then GoTo CleanExit

    Set loOutbox = wbOutbox.Worksheets("OutboxEvents").ListObjects("tblOutboxEvents")
    Set loSnap = wbSnap.Worksheets("InventorySnapshot").ListObjects("tblInventorySnapshot")
    If FindRowByColumnValue(loOutbox, "EventID", "EVT-P5-001") = 0 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loOutbox, FindRowByColumnValue(loOutbox, "EventID", "EVT-P5-001"), "RunId")) = "" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loOutbox, FindRowByColumnValue(loOutbox, "EventID", "EVT-P5-001"), "WarehouseId")) <> "WHS5" Then GoTo CleanExit
    If FindRowByColumnValue(loSnap, "SKU", "SKU-001") = 0 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loSnap, FindRowByColumnValue(loSnap, "SKU", "SKU-001"), "QtyOnHand")) <> 9 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loSnap, FindRowByColumnValue(loSnap, "SKU", "SKU-001"), "QtyAvailable")) <> 9 Then GoTo CleanExit
    If InStr(1, CStr(TestPhase2Helpers.GetRowValue(loSnap, FindRowByColumnValue(loSnap, "SKU", "SKU-001"), "LocationSummary")), "A1", vbTextCompare) = 0 Then GoTo CleanExit

    TestRunBatch_WritesOutboxAndSnapshot = 1

CleanExit:
    TestPhase2Helpers.CloseAndDeleteWorkbook wbSnap
    TestPhase2Helpers.CloseAndDeleteWorkbook wbOutbox
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInbox
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestManualCopy_PublishesWarehouseArtifacts() As Long
    Dim localRoot As String
    Dim shareRoot As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim wbOutbox As Workbook
    Dim wbSnap As Workbook
    Dim report As String

    On Error GoTo CleanFail
    localRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase5PublishLocal")
    shareRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase5PublishShare")
    CreateFolderIfMissing shareRoot & "\Events"
    CreateFolderIfMissing shareRoot & "\Snapshots"

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHP5", "S1", "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", localRoot
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathSharePointRoot", shareRoot
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHP5")
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHP5", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WHP5", "*", "ACTIVE"
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHP5", localRoot, Array("SKU-001"))
    Set wbInbox = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook("S1", localRoot)
    TestPhase2Helpers.AddInboxReceiveRow wbInbox, "EVT-P5-002", Now, "WHP5", "S1", "user1", "SKU-001", 3, "A1", "publish"

    If modProcessor.RunBatch("WHP5", 500, report) <> 1 Then GoTo CleanExit

    Set wbOutbox = OpenWorkbookIfNeeded(localRoot & "\WHP5.Outbox.Events.xlsb")
    Set wbSnap = OpenWorkbookIfNeeded(localRoot & "\WHP5.invSys.Snapshot.Inventory.xlsb")
    If wbOutbox Is Nothing Or wbSnap Is Nothing Then GoTo CleanExit
    wbOutbox.Save
    wbSnap.Save
    wbOutbox.Close SaveChanges:=False
    Set wbOutbox = Nothing
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    FileCopy localRoot & "\WHP5.Outbox.Events.xlsb", shareRoot & "\Events\WHP5.Outbox.Events.xlsb"
    FileCopy localRoot & "\WHP5.invSys.Snapshot.Inventory.xlsb", shareRoot & "\Snapshots\WHP5.invSys.Snapshot.Inventory.xlsb"

    If Len(Dir$(shareRoot & "\Events\WHP5.Outbox.Events.xlsb")) = 0 Then GoTo CleanExit
    If Len(Dir$(shareRoot & "\Snapshots\WHP5.invSys.Snapshot.Inventory.xlsb")) = 0 Then GoTo CleanExit

    TestManualCopy_PublishesWarehouseArtifacts = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbSnap
    TestPhase2Helpers.CloseNoSave wbOutbox
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInbox
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestHqAggregation_TwoWarehousesPreservesPerWarehouseQty() As Long
    Dim shareRoot As String
    Dim localRoot1 As String
    Dim localRoot2 As String
    Dim wbCfg1 As Workbook
    Dim wbAuth1 As Workbook
    Dim wbInv1 As Workbook
    Dim wbInbox1 As Workbook
    Dim wbCfg2 As Workbook
    Dim wbAuth2 As Workbook
    Dim wbInv2 As Workbook
    Dim wbInbox2 As Workbook
    Dim wbGlobal As Workbook
    Dim loGlobal As ListObject
    Dim report As String

    On Error GoTo CleanFail
    shareRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase5Share")
    localRoot1 = TestPhase2Helpers.BuildUniqueTestFolder("Phase5WH1")
    localRoot2 = TestPhase2Helpers.BuildUniqueTestFolder("Phase5WH2")
    CreateFolderIfMissing shareRoot & "\Snapshots"
    CreateFolderIfMissing shareRoot & "\Global"

    Set wbCfg1 = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH51", "S1", "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg1, "PathDataRoot", localRoot1
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg1, "PathSharePointRoot", shareRoot
    Set wbAuth1 = TestPhase2Helpers.BuildPhase2AuthWorkbook("WH51")
    TestPhase2Helpers.AddCapability wbAuth1, "user1", "RECEIVE_POST", "WH51", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth1, "svc_processor", "INBOX_PROCESS", "WH51", "*", "ACTIVE"
    Set wbInv1 = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WH51", localRoot1, Array("SKU-001"))
    Set wbInbox1 = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook("S1", localRoot1)
    TestPhase2Helpers.AddInboxReceiveRow wbInbox1, "EVT-WH51-001", Now, "WH51", "S1", "user1", "SKU-001", 100, "A1", "wh1"

    If modProcessor.RunBatch("WH51", 500, report) <> 1 Then GoTo CleanExit
    CloseIfOpen localRoot1 & "\WH51.invSys.Snapshot.Inventory.xlsb"
    FileCopy localRoot1 & "\WH51.invSys.Snapshot.Inventory.xlsb", shareRoot & "\Snapshots\WH51.invSys.Snapshot.Inventory.xlsb"

    Set wbCfg2 = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH52", "S2", "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg2, "PathDataRoot", localRoot2
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg2, "PathSharePointRoot", shareRoot
    Set wbAuth2 = TestPhase2Helpers.BuildPhase2AuthWorkbook("WH52")
    TestPhase2Helpers.AddCapability wbAuth2, "user1", "RECEIVE_POST", "WH52", "S2", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth2, "svc_processor", "INBOX_PROCESS", "WH52", "*", "ACTIVE"
    Set wbInv2 = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WH52", localRoot2, Array("SKU-001"))
    Set wbInbox2 = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook("S2", localRoot2)
    TestPhase2Helpers.AddInboxReceiveRow wbInbox2, "EVT-WH52-001", Now, "WH52", "S2", "user1", "SKU-001", 50, "A1", "wh2"

    If modProcessor.RunBatch("WH52", 500, report) <> 1 Then GoTo CleanExit
    CloseIfOpen localRoot2 & "\WH52.invSys.Snapshot.Inventory.xlsb"
    FileCopy localRoot2 & "\WH52.invSys.Snapshot.Inventory.xlsb", shareRoot & "\Snapshots\WH52.invSys.Snapshot.Inventory.xlsb"

    If Not modHqAggregator.RunHQAggregation(shareRoot, "", report) Then GoTo CleanExit
    Set wbGlobal = OpenWorkbookIfNeeded(shareRoot & "\Global\invSys.Global.InventorySnapshot.xlsb")
    If wbGlobal Is Nothing Then GoTo CleanExit
    Set loGlobal = wbGlobal.Worksheets("GlobalInventorySnapshot").ListObjects("tblGlobalInventorySnapshot")
    If FindWarehouseSkuQty(loGlobal, "WH51", "SKU-001") <> 100 Then GoTo CleanExit
    If FindWarehouseSkuQty(loGlobal, "WH52", "SKU-001") <> 50 Then GoTo CleanExit

    TestHqAggregation_TwoWarehousesPreservesPerWarehouseQty = 1

CleanExit:
    TestPhase2Helpers.CloseAndDeleteWorkbook wbGlobal
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInbox2
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInv2
    TestPhase2Helpers.CloseNoSave wbAuth2
    TestPhase2Helpers.CloseNoSave wbCfg2
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInbox1
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInv1
    TestPhase2Helpers.CloseNoSave wbAuth1
    TestPhase2Helpers.CloseNoSave wbCfg1
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function OpenWorkbookIfNeeded(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenWorkbookIfNeeded = wb
            Exit Function
        End If
    Next wb
    If Len(Dir$(fullPath)) > 0 Then Set OpenWorkbookIfNeeded = Application.Workbooks.Open(fullPath)
End Function

Private Sub CloseIfOpen(ByVal fullPath As String)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=True
            Exit For
        End If
    Next wb
End Sub

Private Sub CreateFolderIfMissing(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function FindRowByColumnValue(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByColumnValue = i
            Exit Function
        End If
    Next i
End Function

Private Function FindWarehouseSkuQty(ByVal lo As ListObject, ByVal warehouseId As String, ByVal sku As String) As Double
    Dim i As Long
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "WarehouseId")), warehouseId, vbTextCompare) = 0 And _
           StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "SKU")), sku, vbTextCompare) = 0 Then
            FindWarehouseSkuQty = CDbl(TestPhase2Helpers.GetRowValue(lo, i, "QtyOnHand"))
            Exit Function
        End If
    Next i
End Function

Private Sub Tally(ByVal testResult As Long, ByRef passed As Long, ByRef failed As Long)
    If testResult = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
