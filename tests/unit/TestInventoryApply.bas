Attribute VB_Name = "TestInventoryApply"
Option Explicit

Public Sub RunInventoryApplyTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestApplyReceive_ValidEvent(), passed, failed
    Tally TestApplyReceive_InvalidSKU(), passed, failed
    Tally TestApplyReceive_Duplicate(), passed, failed
    Tally TestApplyReceive_ProtectedSheetReturnsClearError(), passed, failed
    Tally TestApplyShip_MultiLineEvent(), passed, failed
    Tally TestApplyProdConsume_MultiLineEvent(), passed, failed
    Tally TestApplyProdComplete_MultiLineEvent(), passed, failed

    Debug.Print "InventoryDomain.Apply tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestApplyReceive_ValidEvent() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject
    Dim loApplied As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-001", "WH1", "S1", "user1", "SKU-001", 5, "A1", "first receipt")

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(statusOut) <> "APPLIED" Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    Set loApplied = wbInv.Worksheets("AppliedEvents").ListObjects("tblAppliedEvents")
    If loLog.ListRows.Count <> 1 Then GoTo CleanExit
    If loApplied.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loLog, 1, "EventID")) <> "EVT-001" Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 1, "QtyDelta")) <> 5 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loApplied, 1, "Status")) <> "APPLIED" Then GoTo CleanExit

    TestApplyReceive_ValidEvent = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyReceive_ProtectedSheetReturnsClearError() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-004", "WH1", "S1", "user1", "SKU-001", 3)

    wbInv.Worksheets("InventoryLog").Unprotect
    wbInv.Worksheets("InventoryLog").Protect Password:="pw"

    On Error GoTo CleanFail
    If modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(errorCode) <> "INVENTORY_SCHEMA_INVALID" Then GoTo CleanExit
    If InStr(1, errorMessage, "could not be unprotected", vbTextCompare) = 0 Then GoTo CleanExit

    TestApplyReceive_ProtectedSheetReturnsClearError = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyReceive_InvalidSKU() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-002", "WH1", "S1", "user1", "BAD-SKU", 5)

    On Error GoTo CleanFail
    If modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(errorCode) <> "INVALID_SKU" Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If loLog.ListRows.Count <> 0 Then GoTo CleanExit

    TestApplyReceive_InvalidSKU = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyReceive_Duplicate() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-003", "WH1", "S1", "user1", "SKU-001", 1)

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(statusOut) <> "SKIP_DUP" Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If loLog.ListRows.Count <> 1 Then GoTo CleanExit

    TestApplyReceive_Duplicate = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyShip_MultiLineEvent() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim payloadJson As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject
    Dim loApplied As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001", "SKU-002"))
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(101, "SKU-001", 4, "DOCK", "shipment line 1"), _
        TestPhase2Helpers.CreatePayloadItem(102, "SKU-002", 2, "DOCK", "shipment line 2"))
    Set evt = TestPhase2Helpers.CreatePayloadEvent("EVT-SHIP-001", EVENT_TYPE_SHIP, "WH1", "S1", "user1", payloadJson)

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(statusOut) <> "APPLIED" Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    Set loApplied = wbInv.Worksheets("AppliedEvents").ListObjects("tblAppliedEvents")
    If loLog.ListRows.Count <> 2 Then GoTo CleanExit
    If loApplied.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loLog, 1, "EventType")) <> EVENT_TYPE_SHIP Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 1, "QtyDelta")) <> -4 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> -2 Then GoTo CleanExit

    TestApplyShip_MultiLineEvent = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyProdConsume_MultiLineEvent() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim payloadJson As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-COMP", "SKU-FG"))
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(201, "SKU-COMP", 6, "LINE1", "component use", "USED"), _
        TestPhase2Helpers.CreatePayloadItem(202, "SKU-FG", 2, "LINE1", "finished staged", "MADE"))
    Set evt = TestPhase2Helpers.CreatePayloadEvent("EVT-PROD-001", EVENT_TYPE_PROD_CONSUME, "WH1", "S1", "user1", payloadJson)

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If loLog.ListRows.Count <> 2 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 1, "QtyDelta")) <> -6 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> 2 Then GoTo CleanExit

    TestApplyProdConsume_MultiLineEvent = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyProdComplete_MultiLineEvent() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim payloadJson As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-FG1", "SKU-FG2"))
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(301, "SKU-FG1", 5, "FG", "completed lot 1", "COMPLETE"), _
        TestPhase2Helpers.CreatePayloadItem(302, "SKU-FG2", 1, "FG", "completed lot 2", "MADE"))
    Set evt = TestPhase2Helpers.CreatePayloadEvent("EVT-PROD-002", EVENT_TYPE_PROD_COMPLETE, "WH1", "S1", "user1", payloadJson)

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If loLog.ListRows.Count <> 2 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 1, "QtyDelta")) <> 5 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> 1 Then GoTo CleanExit

    TestApplyProdComplete_MultiLineEvent = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Sub Tally(ByVal testResult As Long, ByRef passed As Long, ByRef failed As Long)
    If testResult = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
