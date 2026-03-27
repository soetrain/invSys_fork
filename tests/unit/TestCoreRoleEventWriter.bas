Attribute VB_Name = "TestCoreRoleEventWriter"
Option Explicit

Public Sub RunCoreRoleEventWriterTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestQueueReceiveEvent_WritesInboxRow(), passed, failed
    Tally TestOpenInboxWorkbook_UsesStationPathInboxRoot(), passed, failed
    Tally TestQueueShipEvent_WritesInboxRow(), passed, failed
    Tally TestQueuePayloadEvent_DeniedWithoutCapability(), passed, failed
    Tally TestBuildPayloadJson_WithObjectItems(), passed, failed

    Debug.Print "Core.RoleEventWriter tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestQueueReceiveEvent_WritesInboxRow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim eventIdOut As String
    Dim errorMessage As String

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHR1", "R1", "RECEIVE")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHR1")
    Set wbInbox = TestPhase2Helpers.BuildReceiveInboxWorkbook("R1")
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHR1", "R1", "ACTIVE"

    On Error GoTo CleanFail
    If Not modRoleEventWriter.QueueReceiveEvent("WHR1", "R1", "user1", "SKU-001", 4, "A1", "receive test", "", "", Now, wbInbox, eventIdOut, errorMessage) Then GoTo CleanExit

    Set lo = wbInbox.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    If lo.ListRows.Count <> 2 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "EventID")) <> eventIdOut Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "EventType")) <> EVENT_TYPE_RECEIVE Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "UserId")) <> "user1" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "SKU")) <> "SKU-001" Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(lo, 2, "Qty")) <> 4 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "Status")) <> "NEW" Then GoTo CleanExit

    TestQueueReceiveEvent_WritesInboxRow = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestOpenInboxWorkbook_UsesStationPathInboxRoot() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim inboxRoot As String
    Dim expectedPath As String
    Dim errorMessage As String

    inboxRoot = Environ$("TEMP") & "\invsys_role_writer_" & Format$(Now, "yyyymmdd_hhnnss")
    If Len(Dir$(inboxRoot, vbDirectory)) = 0 Then MkDir inboxRoot

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHR2", "R2", "RECEIVE")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHR2")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", Environ$("TEMP") & "\invsys_wrong_data_root"
    TestPhase2Helpers.SetStationConfigValue wbCfg, "PathInboxRoot", inboxRoot
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHR2", "R2", "ACTIVE"

    On Error GoTo CleanFail
    Set wbInbox = modRoleEventWriter.OpenInboxWorkbook(EVENT_TYPE_RECEIVE, "WHR2", "R2", errorMessage)
    If wbInbox Is Nothing Then GoTo CleanExit

    expectedPath = inboxRoot & "\invSys.Inbox.Receiving.R2.xlsb"
    If StrComp(wbInbox.FullName, expectedPath, vbTextCompare) <> 0 Then GoTo CleanExit
    If Len(Dir$(expectedPath, vbNormal)) = 0 Then GoTo CleanExit

    TestOpenInboxWorkbook_UsesStationPathInboxRoot = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    On Error Resume Next
    If expectedPath <> "" Then
        If Len(Dir$(expectedPath, vbNormal)) > 0 Then Kill expectedPath
    End If
    If Len(Dir$(inboxRoot, vbDirectory)) > 0 Then RmDir inboxRoot
    On Error GoTo 0
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestQueueShipEvent_WritesInboxRow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim errorMessage As String

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHS1", "H1", "SHIP")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHS1")
    Set wbInbox = TestPhase2Helpers.BuildShipInboxWorkbook("H1")
    TestPhase2Helpers.AddCapability wbAuth, "user1", "SHIP_POST", "WHS1", "H1", "ACTIVE"

    Dim payloadItems As Collection
    Set payloadItems = New Collection
    payloadItems.Add modRoleEventWriter.CreatePayloadItem(101, "SKU-001", 2, "DOCK", "line 1")
    payloadItems.Add modRoleEventWriter.CreatePayloadItem(102, "SKU-002", 3, "DOCK", "line 2")
    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)

    On Error GoTo CleanFail
    If Not modRoleEventWriter.QueuePayloadEvent(EVENT_TYPE_SHIP, "WHS1", "H1", "user1", payloadJson, "ship test", "", "", Now, wbInbox, eventIdOut, errorMessage) Then GoTo CleanExit

    Set lo = wbInbox.Worksheets("InboxShip").ListObjects("tblInboxShip")
    If lo.ListRows.Count <> 2 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "EventID")) <> eventIdOut Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "EventType")) <> EVENT_TYPE_SHIP Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "PayloadJson")) <> payloadJson Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "Status")) <> "NEW" Then GoTo CleanExit

    TestQueueShipEvent_WritesInboxRow = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestQueuePayloadEvent_DeniedWithoutCapability() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim errorMessage As String

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHP1", "P1", "PROD")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHP1")
    Set wbInbox = TestPhase2Helpers.BuildProductionInboxWorkbook("P1")

    Dim payloadItems As Collection
    Set payloadItems = New Collection
    payloadItems.Add modRoleEventWriter.CreatePayloadItem(201, "SKU-001", 1, "LINE1", "made line", "MADE")
    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)

    On Error GoTo CleanFail
    If modRoleEventWriter.QueuePayloadEvent(EVENT_TYPE_PROD_COMPLETE, "WHP1", "P1", "user1", payloadJson, "prod test", "", "", Now, wbInbox, eventIdOut, errorMessage) Then GoTo CleanExit
    If InStr(1, errorMessage, "PROD_POST", vbTextCompare) = 0 Then GoTo CleanExit

    Set lo = wbInbox.Worksheets("InboxProd").ListObjects("tblInboxProd")
    If lo.ListRows.Count <> 0 Then GoTo CleanExit

    TestQueuePayloadEvent_DeniedWithoutCapability = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestBuildPayloadJson_WithObjectItems() As Long
    Dim item1 As Object
    Dim item2 As Object
    Dim payloadJson As String

    On Error GoTo CleanFail
    Set item1 = modRoleEventWriter.CreatePayloadItem(101, "SKU-001", 2, "DOCK", "line 1")
    Set item2 = modRoleEventWriter.CreatePayloadItem(102, "SKU-002", 3, "DOCK", "line 2")
    payloadJson = modRoleEventWriter.BuildPayloadJson(item1, item2)

    If InStr(1, payloadJson, """SKU"":""SKU-001""", vbTextCompare) = 0 Then GoTo CleanExit
    If InStr(1, payloadJson, """SKU"":""SKU-002""", vbTextCompare) = 0 Then GoTo CleanExit

    TestBuildPayloadJson_WithObjectItems = 1

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Sub Tally(ByVal resultIn As Long, ByRef passed As Long, ByRef failed As Long)
    If resultIn = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
