Attribute VB_Name = "TestAdminConsole"
Option Explicit

Public Sub RunAdminConsoleTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestBreakInventoryLock_WritesBrokenStatusAndAudit(), passed, failed
    Tally TestRunProcessorFromConsole_ProcessesInboxAndWritesAudit(), passed, failed
    Tally TestReissuePoisonEvent_CreatesChildAndReruns(), passed, failed
    Tally TestGenerateInventorySnapshot_WritesWorkbookAndAudit(), passed, failed
    Tally TestPublishWarehouseArtifacts_WritesAuditAndPublishesSnapshot(), passed, failed

    Debug.Print "Admin.Console tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestBreakInventoryLock_WritesBrokenStatusAndAudit() As Long
    Dim tempFolder As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbAdmin As Workbook
    Dim loLocks As ListObject
    Dim loAudit As ListObject
    Dim runId As String
    Dim report As String

    On Error GoTo CleanFail
    tempFolder = TestPhase2Helpers.BuildUniqueTestFolder("Phase4BreakLock")
    modRuntimeWorkbooks.SetCoreDataRootOverride tempFolder
    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook("WHA4", "ADM1", tempFolder, "ADMIN")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", tempFolder
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook("WHA4", tempFolder)
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHA4", tempFolder, Array("SKU-001"))
    Set wbAdmin = Application.Workbooks.Add
    Call modAdminConsole.EnsureAdminSchema(wbAdmin, report)

    TestPhase2Helpers.AddCapability wbAuth, "admin1", "ADMIN_MAINT", "WHA4", "ADM1", "ACTIVE"
    If Not modLockManager.AcquireLock("INVENTORY", "WHA4", "svc_processor", "ADM1", wbInv, runId, report) Then GoTo CleanExit
    If Not modAdminConsole.BreakInventoryLock("manual break", "admin1", "WHA4", wbInv, wbAdmin, report) Then GoTo CleanExit

    Set loLocks = wbInv.Worksheets("Locks").ListObjects("tblLocks")
    Set loAudit = wbAdmin.Worksheets("AdminAudit").ListObjects("tblAdminAudit")
    If UCase$(CStr(TestPhase2Helpers.GetRowValue(loLocks, 1, "Status"))) <> "BROKEN" Then GoTo CleanExit
    If FindAuditRowByAction(loAudit, "BREAK_LOCK") = 0 Then GoTo CleanExit
    If InStr(1, CStr(TestPhase2Helpers.GetRowValue(loAudit, FindAuditRowByAction(loAudit, "BREAK_LOCK"), "Reason")), "manual break", vbTextCompare) = 0 Then GoTo CleanExit

    TestBreakInventoryLock_WritesBrokenStatusAndAudit = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    TestPhase2Helpers.CloseNoSave wbAdmin
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRunProcessorFromConsole_ProcessesInboxAndWritesAudit() As Long
    Dim tempFolder As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim wbAdmin As Workbook
    Dim loInbox As ListObject
    Dim loLog As ListObject
    Dim loAudit As ListObject
    Dim report As String

    On Error GoTo CleanFail
    tempFolder = TestPhase2Helpers.BuildUniqueTestFolder("Phase4RunProcessor")
    modRuntimeWorkbooks.SetCoreDataRootOverride tempFolder
    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook("WHP4", "ADM1", tempFolder, "ADMIN")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", tempFolder
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook("WHP4", tempFolder)
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHP4", tempFolder, Array("SKU-001"))
    Set wbInbox = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook("ADM1", tempFolder)
    Set wbAdmin = Application.Workbooks.Add
    Call modAdminConsole.EnsureAdminSchema(wbAdmin, report)

    TestPhase2Helpers.AddCapability wbAuth, "admin1", "ADMIN_MAINT", "WHP4", "ADM1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHP4", "ADM1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WHP4", "*", "ACTIVE"
    TestPhase2Helpers.AddInboxReceiveRow wbInbox, "EVT-ADMIN-PROC-001", Now, "WHP4", "ADM1", "user1", "SKU-001", 4, "A1", "admin run"

    If modAdminConsole.RunProcessorFromConsole("admin1", "WHP4", wbAdmin, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    Set loAudit = wbAdmin.Worksheets("AdminAudit").ListObjects("tblAdminAudit")
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, 1, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If FindRowByColumnValue(loLog, "EventID", "EVT-ADMIN-PROC-001") = 0 Then GoTo CleanExit
    If FindAuditRowByAction(loAudit, "RUN_PROCESSOR") = 0 Then GoTo CleanExit

    TestRunProcessorFromConsole_ProcessesInboxAndWritesAudit = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    TestPhase2Helpers.CloseNoSave wbAdmin
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReissuePoisonEvent_CreatesChildAndReruns() As Long
    Dim tempFolder As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim wbAdmin As Workbook
    Dim loInbox As ListObject
    Dim loAudit As ListObject
    Dim corrections As Object
    Dim newEventId As String
    Dim report As String
    Dim poisonCount As Long
    Dim newRow As Long

    On Error GoTo CleanFail
    tempFolder = TestPhase2Helpers.BuildUniqueTestFolder("Phase4ReissuePoison")
    modRuntimeWorkbooks.SetCoreDataRootOverride tempFolder
    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook("WHQ4", "ADM1", tempFolder, "ADMIN")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", tempFolder
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook("WHQ4", tempFolder)
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHQ4", tempFolder, Array("SKU-001"))
    Set wbInbox = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook("ADM1", tempFolder)
    Set wbAdmin = Application.Workbooks.Add
    Call modAdminConsole.EnsureAdminSchema(wbAdmin, report)

    TestPhase2Helpers.AddCapability wbAuth, "admin1", "ADMIN_MAINT", "WHQ4", "ADM1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHQ4", "ADM1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WHQ4", "*", "ACTIVE"
    TestPhase2Helpers.AddInboxReceiveRow wbInbox, "EVT-POISON-001", Now, "WHQ4", "ADM1", "user1", "BAD-SKU", 6, "A1", "bad sku"

    If modProcessor.RunBatch("WHQ4", 500, report) <> 0 Then GoTo CleanExit
    poisonCount = modAdminConsole.RefreshPoisonQueue(wbAdmin, report)
    If poisonCount <> 1 Then GoTo CleanExit

    Set corrections = CreateObject("Scripting.Dictionary")
    corrections.CompareMode = vbTextCompare
    corrections.Add "SKU", "SKU-001"
    corrections.Add "Note", "fixed sku"

    If Not modAdminConsole.ReissuePoisonEvent(wbInbox.Name, "tblInboxReceive", "EVT-POISON-001", "admin1", corrections, "fix sku", wbAdmin, newEventId, report) Then GoTo CleanExit
    If modAdminConsole.RunProcessorFromConsole("admin1", "WHQ4", wbAdmin, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    Set loAudit = wbAdmin.Worksheets("AdminAudit").ListObjects("tblAdminAudit")
    newRow = FindRowByColumnValue(loInbox, "EventID", newEventId)
    If newRow = 0 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, 1, "Status")) <> "POISON" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, newRow, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, newRow, "ParentEventId")) <> "EVT-POISON-001" Then GoTo CleanExit
    If FindAuditRowByAction(loAudit, "REISSUE_POISON") = 0 Then GoTo CleanExit

    TestReissuePoisonEvent_CreatesChildAndReruns = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    TestPhase2Helpers.CloseNoSave wbAdmin
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestGenerateInventorySnapshot_WritesWorkbookAndAudit() As Long
    Dim runtimeRoot As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbAdmin As Workbook
    Dim wbSnap As Workbook
    Dim loAudit As ListObject
    Dim loSnap As ListObject
    Dim evt As Object
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim outPath As String
    Dim tempFolder As String

    On Error GoTo CleanFail
    tempFolder = TestPhase2Helpers.BuildUniqueTestFolder("Phase4Snapshot")
    runtimeRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase4SnapshotRuntime")
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    outPath = tempFolder & "\snapshot.test.xlsb"

    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook("WHS4", "ADM1", runtimeRoot, "ADMIN")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", runtimeRoot
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook("WHS4", runtimeRoot)
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHS4", runtimeRoot, Array("SKU-001", "SKU-002"))
    Set wbAdmin = Application.Workbooks.Add
    Call modAdminConsole.EnsureAdminSchema(wbAdmin, report)

    TestPhase2Helpers.AddCapability wbAuth, "admin1", "ADMIN_MAINT", "WHS4", "ADM1", "ACTIVE"
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-SNAP-001", "WHS4", "ADM1", "user1", "SKU-001", 5, "A1", "snap1")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-SNAP", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-SNAP-002", "WHS4", "ADM1", "user1", "SKU-002", 2, "A1", "snap2")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-SNAP", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    If Not modAdminConsole.GenerateInventorySnapshot("admin1", "WHS4", wbInv, outPath, wbAdmin, report) Then GoTo CleanExit
    If Len(Dir$(outPath)) = 0 Then GoTo CleanExit

    Set wbSnap = Application.Workbooks.Open(outPath)
    Set loSnap = wbSnap.Worksheets("InventorySnapshot").ListObjects("tblInventorySnapshot")
    Set loAudit = wbAdmin.Worksheets("AdminAudit").ListObjects("tblAdminAudit")
    If FindRowByColumnValue(loSnap, "SKU", "SKU-001") = 0 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loSnap, FindRowByColumnValue(loSnap, "SKU", "SKU-001"), "QtyAvailable")) <> 5 Then GoTo CleanExit
    If InStr(1, CStr(TestPhase2Helpers.GetRowValue(loSnap, FindRowByColumnValue(loSnap, "SKU", "SKU-001"), "LocationSummary")), "A1", vbTextCompare) = 0 Then GoTo CleanExit
    If FindAuditRowByAction(loAudit, "GENERATE_SNAPSHOT") = 0 Then GoTo CleanExit

    TestGenerateInventorySnapshot_WritesWorkbookAndAudit = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    TestPhase2Helpers.CloseAndDeleteWorkbook wbSnap
    TestPhase2Helpers.CloseNoSave wbAdmin
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestPublishWarehouseArtifacts_WritesAuditAndPublishesSnapshot() As Long
    Dim tempFolder As String
    Dim shareRoot As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbAdmin As Workbook
    Dim wbSnap As Workbook
    Dim loAudit As ListObject
    Dim loSnap As ListObject
    Dim evt As Object
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String

    On Error GoTo CleanFail
    tempFolder = TestPhase2Helpers.BuildUniqueTestFolder("AdminWanPublishLocal")
    shareRoot = TestPhase2Helpers.BuildUniqueTestFolder("AdminWanPublishShare")

    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook("WHA5", "ADM1", tempFolder, "ADMIN")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", tempFolder
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathSharePointRoot", shareRoot
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook("WHA5", tempFolder)
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHA5", tempFolder, Array("SKU-001"))
    Set wbAdmin = Application.Workbooks.Add
    Call modAdminConsole.EnsureAdminSchema(wbAdmin, report)

    TestPhase2Helpers.AddCapability wbAuth, "admin1", "ADMIN_MAINT", "WHA5", "ADM1", "ACTIVE"
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-ADMIN-WAN-001", "WHA5", "ADM1", "user1", "SKU-001", 9, "A1", "admin-wan")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-ADMIN-WAN", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    If Not modAdminConsole.PublishWarehouseArtifacts("admin1", "WHA5", wbInv, wbAdmin, report) Then GoTo CleanExit

    Set wbSnap = Application.Workbooks.Open(shareRoot & "\Snapshots\WHA5.invSys.Snapshot.Inventory.xlsb")
    Set loSnap = wbSnap.Worksheets("InventorySnapshot").ListObjects("tblInventorySnapshot")
    Set loAudit = wbAdmin.Worksheets("AdminAudit").ListObjects("tblAdminAudit")
    If FindRowByColumnValue(loSnap, "SKU", "SKU-001") = 0 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loSnap, FindRowByColumnValue(loSnap, "SKU", "SKU-001"), "QtyOnHand")) <> 9 Then GoTo CleanExit
    If FindAuditRowByAction(loAudit, "PUBLISH_WAN") = 0 Then GoTo CleanExit

    TestPublishWarehouseArtifacts_WritesAuditAndPublishesSnapshot = 1

CleanExit:
    TestPhase2Helpers.CloseAndDeleteWorkbook wbSnap
    TestPhase2Helpers.CloseNoSave wbAdmin
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function FindAuditRowByAction(ByVal lo As ListObject, ByVal actionName As String) As Long
    Dim i As Long
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = lo.ListRows.Count To 1 Step -1
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "Action")), actionName, vbTextCompare) = 0 Then
            FindAuditRowByAction = i
            Exit Function
        End If
    Next i
End Function

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

Private Sub Tally(ByVal testResult As Long, ByRef passed As Long, ByRef failed As Long)
    If testResult = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
