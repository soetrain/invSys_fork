Attribute VB_Name = "modTS_Received"
Option Explicit

Private Const SHEET_RECEIVING As String = "ReceivedTally"
Private Const TABLE_STAGING As String = "ReceivedTally"
Private Const TABLE_AGGREGATE As String = "AggregateReceived"
Private Const TABLE_INVENTORY As String = "invSys"
Private Const SHEET_INVENTORY As String = "InventoryManagement"
Private Const SHEET_LOG As String = "ReceivedLog"

Private mLastConfirmSucceeded As Boolean
Private mLastConfirmStatus As String
Private mReceivingLauncherForm As frmReceiving
Private mReceivingLauncherWorkbookName As String
Private mReceivingLauncherFormTerminated As Boolean

Public Sub EnsureGeneratedButtons()
    Dim wb As Workbook
    Dim report As String

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then
        ShowReceivingMessage "Activate a Receiving operator workbook before refreshing.", vbExclamation
        Exit Sub
    End If
    If Not RefreshReceivingUiForWorkbook(wb, "LOCAL", report) Then
        ShowReceivingMessage report, vbExclamation
    End If
End Sub

Public Sub InitializeReceivingUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    Dim wb As Workbook
    Dim report As String
    Dim ws As Worksheet

    Set wb = ResolveReceivingWorkbook(targetWb)
    If wb Is Nothing Then Exit Sub
    If Not modOperationsPrimitiveBridge.EnsureReceivingWorkbookSurface(wb.Name, report) Then
        Err.Raise vbObjectError + 7680, "modTS_Received.InitializeReceivingUiForWorkbook", report
    End If
    Set ws = WorkbookSheet(wb, SHEET_RECEIVING)
    If ws Is Nothing Then Exit Sub
    EnsureReceivingButtons ws
    modOperationsPrimitiveBridge.ApplyShapeCapability _
        wb.Name, SHEET_RECEIVING, "btnConfirmWrites", "RECEIVE_POST"
    modOperationsPrimitiveBridge.InitializeReceivingAutoSnapshot wb.Name
    EnforceReceivingSupportSheetsHidden wb
End Sub

Public Function RefreshReceivingUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing, _
                                              Optional ByVal sourceType As String = "LOCAL", _
                                              Optional ByRef report As String = "", Optional ByRef refreshState As String = "") As Boolean
    Dim wb As Workbook
    refreshState = "FAILED"
    Set wb = ResolveReceivingWorkbook(targetWb)
    If wb Is Nothing Then
        report = "Activate a Receiving operator workbook before refreshing."
        Exit Function
    End If
    InitializeReceivingUiForWorkbook wb
    RefreshReceivingUiForWorkbook = _
        modOperationsPrimitiveBridge.RefreshInventoryReadModel( _
            wb.Name, "", sourceType, report, refreshState)
    EnforceReceivingSupportSheetsHidden wb
End Function

Public Sub ShowReceivingForm(Optional ByVal userControlAction As Boolean = False)
    On Error GoTo ErrHandler

    Dim wb As Workbook, preferredWorkbookName As String, workbookName As String
    Dim report As String, launcherStage As String, activityId As String, notice As String, outcome As String
    outcome = "FAILED"
    If userControlAction Then activityId = modActivity.BeginAction("RECEIVING_OPEN", modActivity.CaptureContext(), notice)

    launcherStage = "capture active workbook"
    If Not Application.ActiveWorkbook Is Nothing Then
        preferredWorkbookName = Application.ActiveWorkbook.Name
    End If

    launcherStage = "resolve or provision Receiving workbook"
    If Not modOperationsPrimitiveBridge.OpenOrCreateCurrentReceivingOperatorWorkbook( _
            preferredWorkbookName, workbookName, report) Then
        If Trim$(report) = "" Then
            report = "The station-local Receiving operator workbook could not be opened."
        End If
        ShowReceivingMessage report, vbExclamation
        GoTo Done
    End If

    launcherStage = "capture resolved Receiving workbook"
    Set wb = modOperationsInit.ResolveOpenWorkbookByName(workbookName)
    If wb Is Nothing Then
        ShowReceivingMessage "The resolved Receiving operator workbook is no longer open.", vbExclamation
        GoTo Done
    End If

    launcherStage = "activate Receiving workbook"
    wb.Activate

    If mReceivingLauncherFormTerminated Then
        Set mReceivingLauncherForm = Nothing
        mReceivingLauncherWorkbookName = vbNullString
        mReceivingLauncherFormTerminated = False
    End If

    outcome = "REUSED"
    If Not IsReceivingLauncherFormReusable(wb) Then
        outcome = "OPENED"
        launcherStage = "replace Receiving form binding"
        On Error Resume Next
        If Not mReceivingLauncherForm Is Nothing Then
            If mReceivingLauncherForm.Visible Then Unload mReceivingLauncherForm
        End If
        Set mReceivingLauncherForm = Nothing
        On Error GoTo ErrHandler
        Set mReceivingLauncherForm = New frmReceiving
        mReceivingLauncherFormTerminated = False
        mReceivingLauncherWorkbookName = wb.Name
        launcherStage = "bind Receiving form"
        mReceivingLauncherForm.SetOperatorWorkbook wb
        launcherStage = "initialize Receiving form"
        mReceivingLauncherForm.InitializeFromReceiving
    End If

    EnforceReceivingSupportSheetsHidden wb
    launcherStage = "show Receiving form"
    If Not mReceivingLauncherForm.Visible Then
        mReceivingLauncherForm.Show vbModeless
    End If
Done:
    modReceivingActivityAction.FinishLifecycle activityId, outcome, notice
    Exit Sub

ErrHandler:
    ShowReceivingMessage _
        "Receiving form failed [Stage=" & Trim$(launcherStage) & _
        "; Err.Number=" & CStr(Err.Number) & _
        "; Err.Source=" & modOperationsInit.SanitizeLauncherErrorSource(Err.Source) & _
        "]: " & Err.Description, vbCritical
    outcome = "FAILED"
    Resume Done
End Sub

Public Sub NotifyReceivingLauncherFormTerminating(ByVal terminatingForm As frmReceiving)
    If terminatingForm Is Nothing Then Exit Sub
    If mReceivingLauncherForm Is Nothing Then Exit Sub
    If Not (terminatingForm Is mReceivingLauncherForm) Then Exit Sub
    mReceivingLauncherFormTerminated = True
End Sub

Private Function IsReceivingLauncherFormReusable(ByVal operatorWb As Workbook) As Boolean
    Dim visibleState As Boolean

    If operatorWb Is Nothing Then Exit Function
    If mReceivingLauncherFormTerminated Then Exit Function
    If mReceivingLauncherForm Is Nothing Then Exit Function
    If StrComp(mReceivingLauncherWorkbookName, operatorWb.Name, vbTextCompare) <> 0 Then Exit Function

    On Error GoTo Disappeared
    visibleState = mReceivingLauncherForm.Visible
    If Not mReceivingLauncherForm.CanReuseFor(operatorWb) Then Exit Function
    IsReceivingLauncherFormReusable = visibleState
Disappeared:
End Function

Public Sub HandleReceivingOperatorWorkbookClosing(ByVal operatorWb As Workbook)
    If operatorWb Is Nothing Then Exit Sub
    If StrComp(mReceivingLauncherWorkbookName, operatorWb.Name, vbTextCompare) <> 0 Then Exit Sub

    On Error Resume Next
    If Not mReceivingLauncherForm Is Nothing Then Unload mReceivingLauncherForm
    Set mReceivingLauncherForm = Nothing
    mReceivingLauncherWorkbookName = vbNullString
    mReceivingLauncherFormTerminated = False
    On Error GoTo 0
End Sub

Public Function RunReceivingConfirmWritesFormActionForTest(ByVal operatorWb As Workbook, _
                                                           Optional ByVal activatedWb As Workbook = Nothing) As String
    Dim frm As frmReceiving

    Set frm = New frmReceiving
    If Not activatedWb Is Nothing Then activatedWb.Activate
    RunReceivingConfirmWritesFormActionForTest = _
        frm.TestRunConfirmWritesActionForWorkbook(operatorWb, activatedWb)
    Unload frm
End Function

Public Function RunReceivingPurchasingTabContractForTest(ByVal operatorWb As Workbook) As String
    Dim frm As frmReceiving

    Set frm = New frmReceiving
    RunReceivingPurchasingTabContractForTest = frm.TestPurchasingTabContract(operatorWb)
    Unload frm
End Function

Public Function RunReceivingReturnsTabContractForTest(ByVal operatorWb As Workbook) As String
    Dim frm As frmReceiving

    Set frm = New frmReceiving
    RunReceivingReturnsTabContractForTest = frm.TestReturnsTabContract(operatorWb)
    Unload frm
End Function

Public Function RunReceivingInboundReturnFormActionForTest(ByVal operatorWb As Workbook) As String
    Dim frm As frmReceiving

    Set frm = New frmReceiving
    RunReceivingInboundReturnFormActionForTest = _
        frm.TestStageInboundReturnActionForWorkbook(operatorWb)
    Unload frm
End Function

Public Function RunReceivingProtectedDispositionFormActionForTest(ByVal operatorWb As Workbook) As String
    Dim frm As frmReceiving

    Set frm = New frmReceiving
    RunReceivingProtectedDispositionFormActionForTest = _
        frm.TestStageProtectedDispositionActionForWorkbook(operatorWb)
    Unload frm
End Function

Public Function RunReceivingSearchAndHeaderContractTest() As String
    On Error GoTo Failed

    ShowReceivingForm
    If mReceivingLauncherForm Is Nothing Then
        RunReceivingSearchAndHeaderContractTest = "FAIL|FormNotOpen"
    Else
        RunReceivingSearchAndHeaderContractTest = _
            mReceivingLauncherForm.TestReceivingSearchAndHeaderContract()
    End If
    Exit Function

Failed:
    RunReceivingSearchAndHeaderContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunReceivingRefreshFormActionForTest(ByVal operatorWorkbookName As String, _
                                                     Optional ByVal filterText As String = "") As String
    Dim frm As frmReceiving
    Dim operatorWb As Workbook
    Dim itemRows As Variant
    Dim historyRows As Variant
    Dim loaderItemRowsBefore As Long
    Dim loaderHistoryRowsBefore As Long
    Dim actionReport As String
    Dim directHistoryRows As Long

    On Error GoTo Failed
    Set operatorWb = modOperationsInit.ResolveOpenWorkbookByName(operatorWorkbookName)
    If operatorWb Is Nothing Then
        RunReceivingRefreshFormActionForTest = "FAIL|Operator workbook is not open."
        Exit Function
    End If

    Set frm = New frmReceiving
    itemRows = LoadReceivingItemChoicesForWorkbook(operatorWb)
    historyRows = LoadReceivingEntriesHistoryForWorkbook(operatorWb)
    loaderItemRowsBefore = VariantArrayRowCount(itemRows)
    loaderHistoryRowsBefore = VariantArrayRowCount(historyRows)
    actionReport = frm.TestRefreshInventoryActionForWorkbook(operatorWb, filterText)
    itemRows = LoadReceivingItemChoicesForWorkbook(operatorWb)
    historyRows = LoadReceivingEntriesHistoryForWorkbook(operatorWb)
    directHistoryRows = frm.TestSearchInventoryCount(historyRows, filterText)
    RunReceivingRefreshFormActionForTest = _
        actionReport & _
        "|LoaderItemRowsBefore=" & CStr(loaderItemRowsBefore) & _
        "|LoaderHistoryRowsBefore=" & CStr(loaderHistoryRowsBefore) & _
        "|LoaderItemRowsAfter=" & CStr(VariantArrayRowCount(itemRows)) & _
        "|LoaderHistoryRowsAfter=" & CStr(VariantArrayRowCount(historyRows)) & _
        "|DirectHistoryRows=" & CStr(directHistoryRows)
CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    Set operatorWb = Nothing
    On Error GoTo 0
    Exit Function

Failed:
    RunReceivingRefreshFormActionForTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    Resume CleanExit
End Function

Private Function VariantArrayRowCount(ByVal values As Variant) As Long
    On Error GoTo NoRows
    If IsEmpty(values) Or Not IsArray(values) Then Exit Function
    VariantArrayRowCount = UBound(values, 1) - LBound(values, 1) + 1
NoRows:
End Function

Public Function ReceivingFormInitializeSmokeForWorkbook(ByVal operatorWb As Workbook) As String
    Dim frm As frmReceiving

    On Error GoTo Failed
    Set frm = New frmReceiving
    ReceivingFormInitializeSmokeForWorkbook = frm.TestInitializeForWorkbook(operatorWb)
CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    On Error GoTo 0
    Exit Function
Failed:
    ReceivingFormInitializeSmokeForWorkbook = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    Resume CleanExit
End Function

Public Sub EnforceReceivingSupportSheetsHidden(ByVal wb As Workbook)
    Dim names As Variant
    Dim nameValue As Variant
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    names = Array(SHEET_RECEIVING, SHEET_INVENTORY, SHEET_LOG)
    For Each nameValue In names
        Set ws = WorkbookSheet(wb, CStr(nameValue))
        If Not ws Is Nothing Then
            If CanHideWorksheet(wb, ws) Then ws.Visible = xlSheetVeryHidden
        End If
    Next nameValue
End Sub

Public Function LoadReceivingFormInventoryForWorkbook(ByVal operatorWb As Workbook, _
                                                      Optional ByVal filterText As String = "") As Variant
    Dim inventoryTable As ListObject
    Dim sourceValues As Variant
    Dim outputValues() As Variant
    Dim trimmedValues() As Variant
    Dim recordIndex As Long
    Dim fieldIndex As Long
    Dim outputIndex As Long
    Dim matchedIndex As Long
    Dim searchText As String
    Dim searchableText As String
    Dim systemKey As String
    Dim itemCode As String
    Dim itemName As String
    Dim uomValue As String
    Dim qtyValue As Double
    Dim qtyText As String
    Dim locationValue As String
    Dim conditionValue As String
    Dim lotNumber As String
    Dim descriptionValue As String
    Dim vendorValue As String
    Dim groupKey As String
    Dim groupIndex As Object

    Set inventoryTable = FindTable(operatorWb, TABLE_INVENTORY)
    If inventoryTable Is Nothing Or inventoryTable.DataBodyRange Is Nothing Then Exit Function
    If ColumnIndex(inventoryTable, "System_Key") = 0 _
       Or ColumnIndex(inventoryTable, "ITEM_CODE") = 0 Then Exit Function

    searchText = LCase$(Trim$(filterText))
    sourceValues = inventoryTable.DataBodyRange.Value2
    ReDim outputValues(1 To UBound(sourceValues, 1), 1 To 10)
    Set groupIndex = CreateObject("Scripting.Dictionary")
    groupIndex.CompareMode = vbTextCompare
    For recordIndex = 1 To UBound(sourceValues, 1)
        systemKey = CellText(inventoryTable, recordIndex, "System_Key")
        itemCode = CellText(inventoryTable, recordIndex, "ITEM_CODE")
        itemName = CellText(inventoryTable, recordIndex, "ITEM")
        If itemName = "" Then itemName = CellText(inventoryTable, recordIndex, "ItemName")
        uomValue = CellText(inventoryTable, recordIndex, "UOM")
        qtyText = CellText(inventoryTable, recordIndex, "QtyAvailable")
        If qtyText = "" Then qtyText = CellText(inventoryTable, recordIndex, "TOTAL INV")
        If IsNumeric(qtyText) Then qtyValue = CDbl(qtyText) Else qtyValue = 0
        locationValue = CellText(inventoryTable, recordIndex, "LOCATION")
        lotNumber = ReceivingInventoryLotNumber(inventoryTable, recordIndex)
        conditionValue = UCase$(CellText(inventoryTable, recordIndex, "Condition"))
        If conditionValue = "" Then conditionValue = "GOOD"
        descriptionValue = CellText(inventoryTable, recordIndex, "DESCRIPTION")
        vendorValue = CellText(inventoryTable, recordIndex, "VENDOR(s)")
        If systemKey = "" Or itemCode = "" Then GoTo NextInventoryRecord

        groupKey = UCase$(itemCode) & Chr$(30) & UCase$(uomValue) & Chr$(30) & _
                   UCase$(locationValue) & Chr$(30) & UCase$(lotNumber) & Chr$(30) & conditionValue
        If groupIndex.Exists(groupKey) Then
            outputValues(CLng(groupIndex(groupKey)), 5) = _
                CDbl(outputValues(CLng(groupIndex(groupKey)), 5)) + qtyValue
        Else
            outputIndex = outputIndex + 1
            groupIndex.Add groupKey, outputIndex
            outputValues(outputIndex, 1) = systemKey
            outputValues(outputIndex, 2) = itemCode
            outputValues(outputIndex, 3) = itemName
            outputValues(outputIndex, 4) = uomValue
            outputValues(outputIndex, 5) = qtyValue
            outputValues(outputIndex, 6) = locationValue
            outputValues(outputIndex, 7) = lotNumber
            outputValues(outputIndex, 8) = conditionValue
            outputValues(outputIndex, 9) = descriptionValue
            outputValues(outputIndex, 10) = vendorValue
        End If
NextInventoryRecord:
    Next recordIndex

    If outputIndex = 0 Then Exit Function

    For recordIndex = 1 To outputIndex
        If CDbl(outputValues(recordIndex, 5)) <= 0 Then GoTo NextGroupedCount
        searchableText = LCase$(CStr(outputValues(recordIndex, 2)) & " " & _
                                    CStr(outputValues(recordIndex, 3)) & " " & _
                                     CStr(outputValues(recordIndex, 6)) & " " & _
                                     CStr(outputValues(recordIndex, 7)) & " " & _
                                     CStr(outputValues(recordIndex, 8)) & " " & _
                                     CStr(outputValues(recordIndex, 9)) & " " & _
                                     CStr(outputValues(recordIndex, 10)))
        If searchText = "" Or InStr(1, searchableText, searchText, vbTextCompare) > 0 Then _
            matchedIndex = matchedIndex + 1
NextGroupedCount:
    Next recordIndex
    If matchedIndex = 0 Then Exit Function

    ReDim trimmedValues(1 To matchedIndex, 1 To 10)
    matchedIndex = 0
    For recordIndex = 1 To outputIndex
        If CDbl(outputValues(recordIndex, 5)) <= 0 Then GoTo NextGroupedCopy
        searchableText = LCase$(CStr(outputValues(recordIndex, 2)) & " " & _
                                    CStr(outputValues(recordIndex, 3)) & " " & _
                                     CStr(outputValues(recordIndex, 6)) & " " & _
                                     CStr(outputValues(recordIndex, 7)) & " " & _
                                     CStr(outputValues(recordIndex, 8)) & " " & _
                                     CStr(outputValues(recordIndex, 9)) & " " & _
                                     CStr(outputValues(recordIndex, 10)))
        If searchText = "" Or InStr(1, searchableText, searchText, vbTextCompare) > 0 Then
            matchedIndex = matchedIndex + 1
            For fieldIndex = 1 To 10
                trimmedValues(matchedIndex, fieldIndex) = outputValues(recordIndex, fieldIndex)
            Next fieldIndex
        End If
NextGroupedCopy:
    Next recordIndex
    LoadReceivingFormInventoryForWorkbook = trimmedValues
End Function

Public Function LoadReceivingFormInventory(Optional ByVal filterText As String = "") As Variant
    Dim wb As Workbook

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Function
    LoadReceivingFormInventory = _
        LoadReceivingFormInventoryForWorkbook(wb, filterText)
End Function

Public Function LoadReceivingItemChoicesForWorkbook(ByVal operatorWb As Workbook) As Variant
    Dim sourceRows As Variant
    Dim seen As Object
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim itemCode As String
    Dim r As Long
    Dim outRow As Long

    sourceRows = LoadReceivingFormInventoryForWorkbook(operatorWb, "")
    If IsEmpty(sourceRows) Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    ReDim result(1 To UBound(sourceRows, 1), 1 To 3)
    For r = 1 To UBound(sourceRows, 1)
        itemCode = Trim$(CStr(sourceRows(r, 2)))
        If itemCode = "" Or seen.Exists(itemCode) Then GoTo NextRow
        seen.Add itemCode, True
        outRow = outRow + 1
        result(outRow, 1) = sourceRows(r, 1)
        result(outRow, 2) = itemCode
        result(outRow, 3) = sourceRows(r, 3)
NextRow:
    Next r
    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 3)
    For r = 1 To outRow
        trimmed(r, 1) = result(r, 1)
        trimmed(r, 2) = result(r, 2)
        trimmed(r, 3) = result(r, 3)
    Next r
    LoadReceivingItemChoicesForWorkbook = trimmed
End Function

Public Function LoadReceivingEntriesHistoryForWorkbook(ByVal operatorWb As Workbook) As Variant
    Dim historyTable As ListObject
    Dim result() As Variant
    Dim sourceRow As Long
    Dim outputRow As Long
    Dim rowCount As Long

    Set historyTable = FindTable(operatorWb, "ReceivedLog")
    If historyTable Is Nothing Or historyTable.DataBodyRange Is Nothing Then Exit Function
    rowCount = historyTable.ListRows.Count
    ReDim result(1 To rowCount, 1 To 10)
    For sourceRow = rowCount To 1 Step -1
        outputRow = outputRow + 1
        result(outputRow, 1) = CellText(historyTable, sourceRow, "ENTRY_DATE")
        result(outputRow, 2) = CellText(historyTable, sourceRow, "RECEIPT_TYPE")
        result(outputRow, 3) = CellText(historyTable, sourceRow, "REF_NUMBER")
        result(outputRow, 4) = CellText(historyTable, sourceRow, "ITEMS")
        result(outputRow, 5) = CellText(historyTable, sourceRow, "QUANTITY")
        result(outputRow, 6) = CellText(historyTable, sourceRow, "UOM")
        result(outputRow, 7) = CellText(historyTable, sourceRow, "LOCATION")
        result(outputRow, 8) = CellText(historyTable, sourceRow, "LOT_NUMBER")
        result(outputRow, 9) = CellText(historyTable, sourceRow, "Condition")
        result(outputRow, 10) = CellText(historyTable, sourceRow, "RETURN_REASON")
    Next sourceRow
    LoadReceivingEntriesHistoryForWorkbook = result
End Function

Public Function LoadReceivingStagingViewForWorkbook(ByVal operatorWb As Workbook) As Variant
    Dim targetTable As ListObject
    Dim result() As Variant
    Dim rowIndex As Long

    Set targetTable = FindTable(operatorWb, TABLE_STAGING)
    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    ReDim result(1 To targetTable.ListRows.Count, 1 To 10)
    For rowIndex = 1 To targetTable.ListRows.Count
        result(rowIndex, 1) = CellText(targetTable, rowIndex, "REF_NUMBER")
        result(rowIndex, 2) = CellText(targetTable, rowIndex, "RECEIPT_TYPE")
        result(rowIndex, 3) = CellText(targetTable, rowIndex, "ITEMS")
        result(rowIndex, 4) = CellText(targetTable, rowIndex, "QUANTITY")
        result(rowIndex, 5) = CellText(targetTable, rowIndex, "UOM")
        result(rowIndex, 6) = CellText(targetTable, rowIndex, "LOCATION")
        result(rowIndex, 7) = CellText(targetTable, rowIndex, "LOT_NUMBER")
        result(rowIndex, 8) = CellText(targetTable, rowIndex, "VENDOR")
        result(rowIndex, 9) = CellText(targetTable, rowIndex, "Condition")
        result(rowIndex, 10) = CellText(targetTable, rowIndex, "RETURN_REASON")
    Next rowIndex
    LoadReceivingStagingViewForWorkbook = result
End Function

Public Function LoadReceivingAggregateViewForWorkbook(ByVal operatorWb As Workbook) As Variant
    Dim targetTable As ListObject
    Dim result() As Variant
    Dim rowIndex As Long

    Set targetTable = FindTable(operatorWb, TABLE_AGGREGATE)
    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    ReDim result(1 To targetTable.ListRows.Count, 1 To 10)
    For rowIndex = 1 To targetTable.ListRows.Count
        result(rowIndex, 1) = CellText(targetTable, rowIndex, "REF_NUMBER")
        result(rowIndex, 2) = CellText(targetTable, rowIndex, "RECEIPT_TYPE")
        result(rowIndex, 3) = CellText(targetTable, rowIndex, "ITEM_CODE")
        result(rowIndex, 4) = CellText(targetTable, rowIndex, "ITEM")
        result(rowIndex, 5) = CellText(targetTable, rowIndex, "UOM")
        result(rowIndex, 6) = CellText(targetTable, rowIndex, "QUANTITY")
        result(rowIndex, 7) = CellText(targetTable, rowIndex, "LOCATION")
        result(rowIndex, 8) = CellText(targetTable, rowIndex, "LOT_NUMBER")
        result(rowIndex, 9) = CellText(targetTable, rowIndex, "Condition")
        result(rowIndex, 10) = CellText(targetTable, rowIndex, "RETURN_REASON")
    Next rowIndex
    LoadReceivingAggregateViewForWorkbook = result
End Function

Public Function LoadReceivingFormTableForWorkbook(ByVal operatorWb As Workbook, _
                                                  ByVal tableName As String) As Variant
    Dim targetTable As ListObject

    Set targetTable = FindTable(operatorWb, tableName)
    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    LoadReceivingFormTableForWorkbook = targetTable.DataBodyRange.Value2
End Function

Public Function LoadReceivingFormTable(ByVal tableName As String) As Variant
    Dim wb As Workbook

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Function
    LoadReceivingFormTable = LoadReceivingFormTableForWorkbook(wb, tableName)
End Function

Public Function StageReceivingFormItemForWorkbook(ByVal targetWb As Workbook, _
                                                  ByVal refNumber As String, _
                                                  ByVal sourceSystemKey As String, _
                                                  ByVal itemCodeValue As String, _
                                                  ByVal qty As Double, _
                                                  ByRef report As String, _
                                                  Optional ByVal locationOverride As String = "", _
                                                  Optional ByVal lotNumber As String = "", _
                                                  Optional ByVal conditionValue As String = "GOOD", _
                                                  Optional ByVal receiptType As String = "RECEIPT", _
                                                  Optional ByVal returnReason As String = "") As Boolean
    On Error GoTo Failed

    Dim inventoryTable As ListObject
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject
    Dim inventoryIndex As Long
    Dim stagingIndex As Long
    Dim aggregateIndex As Long
    Dim receivingSystemKey As String
    Dim eventId As String
    Dim itemCode As String
    Dim itemName As String
    Dim uomValue As String
    Dim locationValue As String
    Dim vendorValue As String
    Dim stagingRecord As ListRow
    Dim failureStage As String
    Dim previousEvents As Boolean
    Dim eventStateCaptured As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    refNumber = Trim$(refNumber)
    sourceSystemKey = Trim$(sourceSystemKey)
    itemCodeValue = Trim$(itemCodeValue)
    If targetWb Is Nothing Then report = "Receiving workbook was not provided.": Exit Function
    If refNumber = "" Then report = "Ref number is required.": Exit Function
    If sourceSystemKey = "" And itemCodeValue = "" Then report = "Select an inventory item first.": Exit Function
    If qty <= 0 Then report = "Quantity must be greater than zero.": Exit Function
    receiptType = UCase$(Trim$(receiptType))
    If receiptType = "" Then receiptType = "RECEIPT"
    If receiptType <> "RECEIPT" And receiptType <> "RETURN" And receiptType <> "DUMP" Then
        report = "Receiving type must be RECEIPT, RETURN, or DUMP."
        Exit Function
    End If
    returnReason = Trim$(returnReason)
    If receiptType <> "RECEIPT" And returnReason = "" Then
        report = "Disposition reason is required."
        Exit Function
    End If
    If receiptType = "RETURN" Or receiptType = "DUMP" Then
        StageReceivingFormItemForWorkbook = StageInventoryDispositionForWorkbook( _
            targetWb, refNumber, sourceSystemKey, itemCodeValue, qty, report, _
            receiptType, returnReason)
        Exit Function
    End If
    conditionValue = NormalizeReceivingCondition(conditionValue)
    If conditionValue = "" Then report = "Choose a valid receiving condition.": Exit Function

    failureStage = "resolve Receiving inventory and staging tables"
    Set inventoryTable = FindTable(targetWb, TABLE_INVENTORY)
    Set stagingTable = FindTable(targetWb, TABLE_STAGING)
    Set aggregateTable = FindTable(targetWb, TABLE_AGGREGATE)
    If inventoryTable Is Nothing Or stagingTable Is Nothing Or aggregateTable Is Nothing Then
        report = "Receiving inventory or staging tables are missing."
        Exit Function
    End If

    failureStage = "resolve selected inventory entity"
    If sourceSystemKey <> "" Then
        inventoryIndex = FindTableRecord(inventoryTable, "System_Key", sourceSystemKey)
    End If
    If inventoryIndex = 0 And itemCodeValue <> "" Then
        inventoryIndex = FindTableRecord(inventoryTable, "ITEM_CODE", itemCodeValue)
    End If
    If inventoryIndex = 0 Then
        report = "The selected inventory item was not found by System_Key or ITEM_CODE."
        Exit Function
    End If

    failureStage = "read selected inventory entity"
    sourceSystemKey = CellText(inventoryTable, inventoryIndex, "System_Key")
    itemCode = CellText(inventoryTable, inventoryIndex, "ITEM_CODE")
    itemName = CellText(inventoryTable, inventoryIndex, "ITEM")
    If itemName = "" Then itemName = CellText(inventoryTable, inventoryIndex, "ItemName")
    uomValue = CellText(inventoryTable, inventoryIndex, "UOM")
    vendorValue = CellText(inventoryTable, inventoryIndex, "VENDOR(s)")
    locationValue = CellText(inventoryTable, inventoryIndex, "LOCATION")
    If Trim$(locationOverride) <> "" Then locationValue = Trim$(locationOverride)
    lotNumber = Trim$(lotNumber)
    If locationValue = "" Then report = "Receive location is required.": Exit Function

    previousEvents = Application.EnableEvents
    eventStateCaptured = True
    Application.EnableEvents = False

    failureStage = "find existing receipt staging row"
    stagingIndex = FindExistingStagingRecord( _
        stagingTable, refNumber, sourceSystemKey, locationValue, lotNumber, _
        conditionValue, receiptType, returnReason)
    If stagingIndex > 0 Then
        failureStage = "update receipt staging quantity"
        receivingSystemKey = CellText(stagingTable, stagingIndex, "System_Key")
        eventId = CellText(stagingTable, stagingIndex, "EventId")
        SetCellValue stagingTable, stagingIndex, "QUANTITY", _
                     CellNumber(stagingTable, stagingIndex, "QUANTITY") + qty
    Else
        failureStage = "create receipt event identity"
        receivingSystemKey = modRoleEventWriter.CreateSystemKey()
        eventId = modRoleEventWriter.CreateSystemKey()
        failureStage = "add receipt staging row"
        Set stagingRecord = FirstBlankOrNewRecord(stagingTable)
        stagingIndex = stagingRecord.Index
        failureStage = "populate receipt staging row"
        SetCellText stagingTable, stagingIndex, "REF_NUMBER", refNumber
        SetCellText stagingTable, stagingIndex, "RECEIPT_TYPE", receiptType
        SetCellText stagingTable, stagingIndex, "ITEMS", itemName
        SetCellValue stagingTable, stagingIndex, "QUANTITY", qty
        SetCellText stagingTable, stagingIndex, "UOM", uomValue
        SetCellText stagingTable, stagingIndex, "VENDOR", vendorValue
        SetCellText stagingTable, stagingIndex, "LOCATION", locationValue
        SetCellText stagingTable, stagingIndex, "LOT_NUMBER", lotNumber
        SetCellText stagingTable, stagingIndex, "Condition", conditionValue
        SetCellText stagingTable, stagingIndex, "RETURN_REASON", returnReason
        SetCellText stagingTable, stagingIndex, "System_Key", receivingSystemKey
        SetCellText stagingTable, stagingIndex, "ITEM_CODE", itemCode
        SetCellText stagingTable, stagingIndex, "Source_System_Key", sourceSystemKey
        SetCellText stagingTable, stagingIndex, "EventId", eventId
        SetCellText stagingTable, stagingIndex, "WorkflowState", "STAGED"
    End If

    failureStage = "populate receipt aggregate"
    aggregateIndex = FindTableRecord(aggregateTable, "System_Key", receivingSystemKey)
    If aggregateIndex = 0 Then aggregateIndex = FirstBlankOrNewRecord(aggregateTable).Index
    PopulateAggregateRecord aggregateTable, aggregateIndex, inventoryTable, inventoryIndex, _
                            refNumber, receivingSystemKey, eventId, _
                            CellNumber(stagingTable, stagingIndex, "QUANTITY"), _
                            locationValue, lotNumber, conditionValue, receiptType, returnReason

    report = "Staged " & CStr(qty) & " " & uomValue & " of " & itemName & _
             "; System_Key=" & receivingSystemKey
    StageReceivingFormItemForWorkbook = True
CleanExit:
    On Error Resume Next
    If eventStateCaptured Then Application.EnableEvents = previousEvents
    On Error GoTo 0
    Exit Function
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    report = "Receiving staging failed: Stage=" & failureStage & _
             "; Error=" & CStr(errorNumber) & _
             "; Source=" & ReceivingErrorSource(errorSource) & _
             "; Description=" & errorDescription
    Resume CleanExit
End Function

Public Function StageInventoryDispositionForWorkbook(ByVal targetWb As Workbook, _
                                                      ByVal refNumber As String, _
                                                      ByVal selectedSystemKey As String, _
                                                      ByVal itemCodeValue As String, _
                                                      ByVal qty As Double, _
                                                      ByRef report As String, _
                                                      ByVal dispositionType As String, _
                                                      ByVal dispositionReason As String) As Boolean
    On Error GoTo Failed

    Dim inventoryTable As ListObject
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject
    Dim selectedIndex As Long
    Dim rowIndex As Long
    Dim nextIndex As Long
    Dim allocationQty As Double
    Dim availableQty As Double
    Dim totalAvailable As Double
    Dim remainingQty As Double
    Dim itemCode As String
    Dim uomValue As String
    Dim locationValue As String
    Dim lotNumber As String
    Dim conditionValue As String
    Dim lastSystemKey As String
    Dim aggregateReport As String
    Dim allocationCount As Long
    Dim failureStage As String
    Dim previousEvents As Boolean
    Dim eventStateCaptured As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    refNumber = Trim$(refNumber)
    selectedSystemKey = Trim$(selectedSystemKey)
    itemCodeValue = Trim$(itemCodeValue)
    dispositionType = UCase$(Trim$(dispositionType))
    dispositionReason = Trim$(dispositionReason)
    If targetWb Is Nothing Then report = "Receiving workbook was not provided.": Exit Function
    If refNumber = "" Then report = "Disposition reference is required.": Exit Function
    If qty <= 0 Then report = "Quantity must be greater than zero.": Exit Function
    If dispositionType <> "RETURN" And dispositionType <> "DUMP" Then
        report = "Disposition must be RETURN or DUMP."
        Exit Function
    End If
    If dispositionReason = "" Then report = "Disposition reason is required.": Exit Function

    failureStage = "resolve Receiving inventory and staging tables"
    Set inventoryTable = FindTable(targetWb, TABLE_INVENTORY)
    Set stagingTable = FindTable(targetWb, TABLE_STAGING)
    Set aggregateTable = FindTable(targetWb, TABLE_AGGREGATE)
    If inventoryTable Is Nothing Or stagingTable Is Nothing Or aggregateTable Is Nothing Then
        report = "Receiving inventory or staging tables are missing."
        Exit Function
    End If
    failureStage = "resolve selected inventory entity"
    If selectedSystemKey <> "" Then selectedIndex = FindTableRecord(inventoryTable, "System_Key", selectedSystemKey)
    If selectedIndex = 0 And itemCodeValue <> "" Then selectedIndex = FindTableRecord(inventoryTable, "ITEM_CODE", itemCodeValue)
    If selectedIndex = 0 Then report = "The selected inventory item is no longer available.": Exit Function

    failureStage = "read selected inventory group"
    itemCode = CellText(inventoryTable, selectedIndex, "ITEM_CODE")
    uomValue = CellText(inventoryTable, selectedIndex, "UOM")
    locationValue = CellText(inventoryTable, selectedIndex, "LOCATION")
    lotNumber = ReceivingInventoryLotNumber(inventoryTable, selectedIndex)
    conditionValue = NormalizeReceivingCondition(CellText(inventoryTable, selectedIndex, "Condition"))
    If conditionValue = "" Then conditionValue = "GOOD"

    failureStage = "calculate exact available inventory"
    For rowIndex = 1 To inventoryTable.ListRows.Count
        If InventoryRowMatchesDispositionGroup(inventoryTable, rowIndex, itemCode, uomValue, _
                                               locationValue, lotNumber, conditionValue) Then
            totalAvailable = totalAvailable + DispositionAvailableForInventoryRow( _
                inventoryTable, rowIndex, stagingTable)
        End If
    Next rowIndex
    If totalAvailable + 0.0000001 < qty Then
        report = "Disposition quantity exceeds available inventory for the selected item, location, lot, and Condition. Available=" & _
                 Format$(totalAvailable, "0.###") & "; Requested=" & Format$(qty, "0.###") & "."
        Exit Function
    End If

    previousEvents = Application.EnableEvents
    eventStateCaptured = True
    Application.EnableEvents = False
    remainingQty = qty
    Do While remainingQty > 0.0000001
        failureStage = "select exact inventory allocation"
        nextIndex = NextDispositionInventoryRecord(inventoryTable, stagingTable, itemCode, uomValue, _
                                                   locationValue, lotNumber, conditionValue, lastSystemKey)
        If nextIndex = 0 Then
            report = "Unable to allocate disposition quantity across exact inventory entities."
            GoTo CleanExit
        End If
        availableQty = DispositionAvailableForInventoryRow(inventoryTable, nextIndex, stagingTable)
        allocationQty = availableQty
        If allocationQty > remainingQty Then allocationQty = remainingQty
        failureStage = "stage exact inventory allocation"
        If Not StageDispositionAllocation(stagingTable, inventoryTable, nextIndex, refNumber, _
                                          allocationQty, dispositionType, dispositionReason, report) Then GoTo CleanExit
        allocationCount = allocationCount + 1
        remainingQty = remainingQty - allocationQty
        lastSystemKey = CellText(inventoryTable, nextIndex, "System_Key")
    Loop

    failureStage = "rebuild disposition aggregate"
    If Not RebuildAggregationForWorkbook(targetWb, aggregateReport) Then
        report = aggregateReport
        GoTo CleanExit
    End If
    report = "Staged " & dispositionType & " of " & Format$(qty, "0.###") & " " & uomValue & _
             " across " & CStr(allocationCount) & " exact inventory allocation(s)."
    StageInventoryDispositionForWorkbook = True
CleanExit:
    On Error Resume Next
    If eventStateCaptured Then Application.EnableEvents = previousEvents
    On Error GoTo 0
    Exit Function
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    report = "Inventory disposition staging failed: Stage=" & failureStage & _
             "; Error=" & CStr(errorNumber) & _
             "; Source=" & ReceivingErrorSource(errorSource) & _
             "; Description=" & errorDescription
    Resume CleanExit
End Function

Private Function StageDispositionAllocation(ByVal stagingTable As ListObject, _
                                            ByVal inventoryTable As ListObject, _
                                            ByVal inventoryIndex As Long, _
                                            ByVal refNumber As String, _
                                            ByVal qty As Double, _
                                            ByVal dispositionType As String, _
                                            ByVal dispositionReason As String, _
                                            ByRef report As String) As Boolean
    On Error GoTo Failed

    Dim stagingIndex As Long
    Dim stagingRecord As ListRow
    Dim systemKey As String
    Dim eventId As String
    Dim locationValue As String
    Dim lotNumber As String
    Dim conditionValue As String
    Dim itemName As String
    Dim failureStage As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    failureStage = "read exact inventory allocation"
    systemKey = CellText(inventoryTable, inventoryIndex, "System_Key")
    locationValue = CellText(inventoryTable, inventoryIndex, "LOCATION")
    lotNumber = ReceivingInventoryLotNumber(inventoryTable, inventoryIndex)
    conditionValue = NormalizeReceivingCondition(CellText(inventoryTable, inventoryIndex, "Condition"))
    If conditionValue = "" Then conditionValue = "GOOD"
    itemName = CellText(inventoryTable, inventoryIndex, "ITEM")
    If itemName = "" Then itemName = CellText(inventoryTable, inventoryIndex, "ItemName")

    failureStage = "find existing disposition staging row"
    stagingIndex = FindExistingStagingRecord(stagingTable, refNumber, systemKey, locationValue, _
                                             lotNumber, conditionValue, dispositionType, dispositionReason)
    If stagingIndex > 0 Then
        failureStage = "update disposition staging quantity"
        SetCellValue stagingTable, stagingIndex, "QUANTITY", _
                     CellNumber(stagingTable, stagingIndex, "QUANTITY") + qty
    Else
        failureStage = "create disposition event identity"
        eventId = modRoleEventWriter.CreateSystemKey()
        failureStage = "add disposition staging row"
        Set stagingRecord = FirstBlankOrNewRecord(stagingTable)
        stagingIndex = stagingRecord.Index
        failureStage = "populate disposition staging row"
        SetCellText stagingTable, stagingIndex, "REF_NUMBER", refNumber
        SetCellText stagingTable, stagingIndex, "RECEIPT_TYPE", dispositionType
        SetCellText stagingTable, stagingIndex, "ITEMS", itemName
        SetCellValue stagingTable, stagingIndex, "QUANTITY", qty
        SetCellText stagingTable, stagingIndex, "UOM", CellText(inventoryTable, inventoryIndex, "UOM")
        SetCellText stagingTable, stagingIndex, "VENDOR", CellText(inventoryTable, inventoryIndex, "VENDOR(s)")
        SetCellText stagingTable, stagingIndex, "LOCATION", locationValue
        SetCellText stagingTable, stagingIndex, "LOT_NUMBER", lotNumber
        SetCellText stagingTable, stagingIndex, "Condition", conditionValue
        SetCellText stagingTable, stagingIndex, "RETURN_REASON", dispositionReason
        SetCellText stagingTable, stagingIndex, "System_Key", systemKey
        SetCellText stagingTable, stagingIndex, "ITEM_CODE", CellText(inventoryTable, inventoryIndex, "ITEM_CODE")
        SetCellText stagingTable, stagingIndex, "Source_System_Key", systemKey
        SetCellText stagingTable, stagingIndex, "EventId", eventId
        SetCellText stagingTable, stagingIndex, "WorkflowState", "STAGED"
    End If
    StageDispositionAllocation = True
    Exit Function
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    report = "Inventory disposition staging failed: Stage=" & failureStage & _
             "; Error=" & CStr(errorNumber) & _
             "; Source=" & ReceivingErrorSource(errorSource) & _
             "; Description=" & errorDescription
End Function

Private Function InventoryRowMatchesDispositionGroup(ByVal inventoryTable As ListObject, _
                                                     ByVal rowIndex As Long, _
                                                     ByVal itemCode As String, _
                                                     ByVal uomValue As String, _
                                                     ByVal locationValue As String, _
                                                     ByVal lotNumber As String, _
                                                     ByVal conditionValue As String) As Boolean
    If StrComp(CellText(inventoryTable, rowIndex, "ITEM_CODE"), itemCode, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CellText(inventoryTable, rowIndex, "UOM"), uomValue, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CellText(inventoryTable, rowIndex, "LOCATION"), locationValue, vbTextCompare) <> 0 Then Exit Function
    If StrComp(ReceivingInventoryLotNumber(inventoryTable, rowIndex), lotNumber, vbTextCompare) <> 0 Then Exit Function
    If StrComp(NormalizeReceivingCondition(CellText(inventoryTable, rowIndex, "Condition")), _
               conditionValue, vbTextCompare) <> 0 Then Exit Function
    InventoryRowMatchesDispositionGroup = True
End Function

Private Function DispositionAvailableForInventoryRow(ByVal inventoryTable As ListObject, _
                                                     ByVal rowIndex As Long, _
                                                     ByVal stagingTable As ListObject) As Double
    Dim qtyText As String
    Dim availableQty As Double
    Dim systemKey As String

    qtyText = CellText(inventoryTable, rowIndex, "QtyAvailable")
    If qtyText = "" Then qtyText = CellText(inventoryTable, rowIndex, "TOTAL INV")
    If IsNumeric(qtyText) Then availableQty = CDbl(qtyText)
    systemKey = CellText(inventoryTable, rowIndex, "System_Key")
    availableQty = availableQty - StagedDispositionQtyForSystemKey(stagingTable, systemKey)
    If availableQty > 0 Then DispositionAvailableForInventoryRow = availableQty
End Function

Private Function StagedDispositionQtyForSystemKey(ByVal stagingTable As ListObject, _
                                                  ByVal systemKey As String) As Double
    Dim rowIndex As Long
    Dim stagedType As String

    If stagingTable Is Nothing Or stagingTable.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To stagingTable.ListRows.Count
        stagedType = UCase$(CellText(stagingTable, rowIndex, "RECEIPT_TYPE"))
        If (stagedType = "RETURN" Or stagedType = "DUMP") _
           And StrComp(CellText(stagingTable, rowIndex, "Source_System_Key"), _
                       systemKey, vbBinaryCompare) = 0 Then
            StagedDispositionQtyForSystemKey = StagedDispositionQtyForSystemKey + _
                CellNumber(stagingTable, rowIndex, "QUANTITY")
        End If
    Next rowIndex
End Function

Private Function NextDispositionInventoryRecord(ByVal inventoryTable As ListObject, _
                                                ByVal stagingTable As ListObject, _
                                                ByVal itemCode As String, _
                                                ByVal uomValue As String, _
                                                ByVal locationValue As String, _
                                                ByVal lotNumber As String, _
                                                ByVal conditionValue As String, _
                                                ByVal afterSystemKey As String) As Long
    Dim rowIndex As Long
    Dim candidateKey As String
    Dim selectedKey As String

    For rowIndex = 1 To inventoryTable.ListRows.Count
        If InventoryRowMatchesDispositionGroup(inventoryTable, rowIndex, itemCode, uomValue, _
                                               locationValue, lotNumber, conditionValue) Then
            candidateKey = CellText(inventoryTable, rowIndex, "System_Key")
            If candidateKey <> "" _
               And StrComp(candidateKey, afterSystemKey, vbBinaryCompare) > 0 _
               And DispositionAvailableForInventoryRow(inventoryTable, rowIndex, stagingTable) > 0 Then
                If selectedKey = "" Or StrComp(candidateKey, selectedKey, vbBinaryCompare) < 0 Then
                    selectedKey = candidateKey
                    NextDispositionInventoryRecord = rowIndex
                End If
            End If
        End If
    Next rowIndex
End Function

Public Function RebuildAggregationForWorkbook(ByVal targetWb As Workbook, _
                                              Optional ByRef report As String = "") As Boolean
    On Error GoTo Failed

    Dim inventoryTable As ListObject
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject
    Dim stagingIndex As Long
    Dim inventoryIndex As Long
    Dim aggregateIndex As Long
    Dim sourceSystemKey As String
    Dim receivingSystemKey As String
    Dim aggregateGroupKey As String
    Dim aggregateIndexes As Object
    Dim receiptType As String
    Dim conditionValue As String
    Dim returnReason As String

    Set inventoryTable = FindTable(targetWb, TABLE_INVENTORY)
    Set stagingTable = FindTable(targetWb, TABLE_STAGING)
    Set aggregateTable = FindTable(targetWb, TABLE_AGGREGATE)
    If inventoryTable Is Nothing Or stagingTable Is Nothing Or aggregateTable Is Nothing Then
        report = "Receiving inventory or staging tables are missing."
        Exit Function
    End If
    If Not aggregateTable.DataBodyRange Is Nothing Then aggregateTable.DataBodyRange.Delete
    If Not stagingTable.DataBodyRange Is Nothing Then
        For stagingIndex = stagingTable.ListRows.Count To 1 Step -1
            If CellText(stagingTable, stagingIndex, "System_Key") = "" _
               And CellText(stagingTable, stagingIndex, "ITEM_CODE") = "" _
               And CellText(stagingTable, stagingIndex, "REF_NUMBER") = "" _
               And CellText(stagingTable, stagingIndex, "EventId") = "" _
               And Abs(CellNumber(stagingTable, stagingIndex, "QUANTITY")) < 0.0000001 Then
                stagingTable.ListRows(stagingIndex).Delete
            End If
        Next stagingIndex
    End If
    If stagingTable.DataBodyRange Is Nothing Then
        report = "OK|Rows=0"
        RebuildAggregationForWorkbook = True
        Exit Function
    End If

    Set aggregateIndexes = CreateObject("Scripting.Dictionary")
    aggregateIndexes.CompareMode = vbTextCompare

    For stagingIndex = 1 To stagingTable.ListRows.Count
        receivingSystemKey = CellText(stagingTable, stagingIndex, "System_Key")
        sourceSystemKey = CellText(stagingTable, stagingIndex, "Source_System_Key")
        If receivingSystemKey = "" Then
            report = "ReceivedTally contains a blank System_Key."
            Exit Function
        End If
        inventoryIndex = FindTableRecord(inventoryTable, "System_Key", sourceSystemKey)
        If inventoryIndex = 0 Then
            inventoryIndex = FindTableRecord( _
                inventoryTable, "ITEM_CODE", CellText(stagingTable, stagingIndex, "ITEM_CODE"))
        End If
        If inventoryIndex = 0 Then
            report = "A staged source inventory item is no longer available."
            Exit Function
        End If
        receiptType = CellText(stagingTable, stagingIndex, "RECEIPT_TYPE")
        If receiptType = "" Then receiptType = "RECEIPT"
        conditionValue = NormalizeReceivingCondition( _
            CellText(stagingTable, stagingIndex, "Condition"))
        If conditionValue = "" Then conditionValue = "GOOD"
        returnReason = CellText(stagingTable, stagingIndex, "RETURN_REASON")
        aggregateGroupKey = BuildReceivingAggregateGroupKey( _
            receiptType, CellText(stagingTable, stagingIndex, "ITEM_CODE"), _
            CellText(stagingTable, stagingIndex, "UOM"), _
            CellText(stagingTable, stagingIndex, "LOCATION"), _
            CellText(stagingTable, stagingIndex, "LOT_NUMBER"), conditionValue)
        If aggregateIndexes.Exists(aggregateGroupKey) Then
            aggregateIndex = CLng(aggregateIndexes(aggregateGroupKey))
            SetCellValue aggregateTable, aggregateIndex, "QUANTITY", _
                CellNumber(aggregateTable, aggregateIndex, "QUANTITY") + _
                CellNumber(stagingTable, stagingIndex, "QUANTITY")
            SetCellText aggregateTable, aggregateIndex, "REF_NUMBER", _
                AppendDistinctReceivingValue( _
                    CellText(aggregateTable, aggregateIndex, "REF_NUMBER"), _
                    CellText(stagingTable, stagingIndex, "REF_NUMBER"))
            SetCellText aggregateTable, aggregateIndex, "RETURN_REASON", _
                AppendDistinctReceivingValue( _
                    CellText(aggregateTable, aggregateIndex, "RETURN_REASON"), returnReason)
        Else
            aggregateIndex = FirstBlankOrNewRecord(aggregateTable).Index
            aggregateIndexes.Add aggregateGroupKey, aggregateIndex
            PopulateAggregateRecord aggregateTable, aggregateIndex, inventoryTable, inventoryIndex, _
                                    CellText(stagingTable, stagingIndex, "REF_NUMBER"), _
                                    receivingSystemKey, _
                                    CellText(stagingTable, stagingIndex, "EventId"), _
                                    CellNumber(stagingTable, stagingIndex, "QUANTITY"), _
                                    CellText(stagingTable, stagingIndex, "LOCATION"), _
                                    CellText(stagingTable, stagingIndex, "LOT_NUMBER"), _
                                    conditionValue, receiptType, returnReason
            SetCellText aggregateTable, aggregateIndex, "WorkflowState", _
                        CellText(stagingTable, stagingIndex, "WorkflowState")
        End If
    Next stagingIndex
    report = "OK|Rows=" & CStr(aggregateIndexes.Count) & _
             "|SourceRows=" & CStr(stagingTable.ListRows.Count)
    RebuildAggregationForWorkbook = True
    Exit Function
Failed:
    report = "Receiving aggregation rebuild failed: " & Err.Description
End Function

Public Sub RebuildAggregation()
    Dim wb As Workbook
    Dim report As String

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Sub
    If Not RebuildAggregationForWorkbook(wb, report) Then
        ShowReceivingMessage report, vbExclamation
    End If
End Sub

Public Sub ClearReceivingFormStagingForWorkbook(ByVal operatorWb As Workbook, Optional ByRef changed As Boolean = False)
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject

    Set stagingTable = FindTable(operatorWb, TABLE_STAGING)
    Set aggregateTable = FindTable(operatorWb, TABLE_AGGREGATE)
    modReceivingPostingService.ClearReceivingStaging stagingTable, aggregateTable, changed
End Sub

Public Sub ClearReceivingFormStaging()
    Dim wb As Workbook

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Sub
    ClearReceivingFormStagingForWorkbook wb
End Sub

Public Sub ConfirmWrites()
    Dim wb As Workbook
    Dim report As String

    mLastConfirmSucceeded = False
    mLastConfirmStatus = "Confirm Writes did not complete."
    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then
        mLastConfirmStatus = "Activate a Receiving operator workbook before confirming writes."
        ShowReceivingMessage mLastConfirmStatus, vbExclamation
        Exit Sub
    End If
    mLastConfirmSucceeded = modReceivingPostingService.ExecuteConfirmWrites(wb, report)
    mLastConfirmStatus = report
    If Not mLastConfirmSucceeded Then ShowReceivingMessage report, vbExclamation
End Sub

Public Sub RecordConfirmWritesResult(ByVal succeeded As Boolean, ByVal statusText As String)
    mLastConfirmSucceeded = succeeded
    mLastConfirmStatus = statusText
End Sub

Public Function LastConfirmWritesSucceeded() As Boolean
    LastConfirmWritesSucceeded = mLastConfirmSucceeded
End Function

Public Function LastConfirmWritesStatus() As String
    LastConfirmWritesStatus = mLastConfirmStatus
End Function

Public Sub ShowReceivingDynamicItemSearch(ByVal targetCell As Range)
    ' The supported Receiving search is hosted by frmReceiving.
End Sub

Public Sub HandleReceivingSelectionChange(ByVal target As Range)
    ' Receiving support sheets are projections; operator selection has no write authority.
End Sub

Public Sub HandleReceivingSheetChange(ByVal target As Range)
    On Error GoTo CleanExit

    Dim stagingTable As ListObject
    Dim qtyColumn As ListColumn
    Dim changedCells As Range
    Dim cell As Range
    Dim recordIndex As Long
    Dim targetWb As Workbook
    Dim previousEvents As Boolean
    Dim eventStateCaptured As Boolean

    If target Is Nothing Then Exit Sub
    Set stagingTable = target.ListObject
    If stagingTable Is Nothing Then Exit Sub
    If StrComp(stagingTable.Name, TABLE_STAGING, vbTextCompare) <> 0 Then Exit Sub
    Set qtyColumn = stagingTable.ListColumns("QUANTITY")
    If qtyColumn.DataBodyRange Is Nothing Then Exit Sub
    Set changedCells = Application.Intersect(target, qtyColumn.DataBodyRange)
    If changedCells Is Nothing Then Exit Sub

    Set targetWb = stagingTable.Parent.Parent
    previousEvents = Application.EnableEvents
    eventStateCaptured = True
    Application.EnableEvents = False
    For Each cell In changedCells.Cells
        recordIndex = cell.Row - stagingTable.DataBodyRange.Row + 1
        SyncQuantityFromStaging recordIndex, NzDbl(cell.Value), targetWb
    Next cell
CleanExit:
    On Error Resume Next
    If eventStateCaptured Then Application.EnableEvents = previousEvents
    On Error GoTo 0
End Sub

Public Sub SyncQuantityFromStaging(ByVal stagingRecordIndex As Long, _
                                   ByVal newQty As Double, _
                                   Optional ByVal targetWb As Workbook = Nothing)
    Dim stagingTable As ListObject
    Dim rebuildReport As String

    If stagingRecordIndex <= 0 Then Exit Sub
    If targetWb Is Nothing Then Set targetWb = Application.ActiveWorkbook
    Set stagingTable = FindTable(targetWb, TABLE_STAGING)
    If stagingTable Is Nothing Then Exit Sub
    If stagingRecordIndex > stagingTable.ListRows.Count Then Exit Sub
    SetCellValue stagingTable, stagingRecordIndex, "QUANTITY", newQty
    Call RebuildAggregationForWorkbook(targetWb, rebuildReport)
End Sub

Public Function NzDbl(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If IsNumeric(valueIn) Then NzDbl = CDbl(valueIn)
End Function

Private Sub PopulateAggregateRecord(ByVal aggregateTable As ListObject, _
                                    ByVal aggregateIndex As Long, _
                                    ByVal inventoryTable As ListObject, _
                                    ByVal inventoryIndex As Long, _
                                    ByVal refNumber As String, _
                                    ByVal receivingSystemKey As String, _
                                    ByVal eventId As String, _
                                    ByVal qty As Double, _
                                    ByVal locationOverride As String, _
                                    ByVal lotNumber As String, _
                                    ByVal conditionValue As String, _
                                    ByVal receiptType As String, _
                                    ByVal returnReason As String)
    SetCellText aggregateTable, aggregateIndex, "REF_NUMBER", refNumber
    SetCellText aggregateTable, aggregateIndex, "RECEIPT_TYPE", receiptType
    SetCellText aggregateTable, aggregateIndex, "ITEM_CODE", _
                CellText(inventoryTable, inventoryIndex, "ITEM_CODE")
    SetCellText aggregateTable, aggregateIndex, "VENDORS", _
                CellText(inventoryTable, inventoryIndex, "VENDOR(s)")
    SetCellText aggregateTable, aggregateIndex, "VENDOR_CODE", _
                CellText(inventoryTable, inventoryIndex, "VENDOR_CODE")
    SetCellText aggregateTable, aggregateIndex, "DESCRIPTION", _
                CellText(inventoryTable, inventoryIndex, "DESCRIPTION")
    SetCellText aggregateTable, aggregateIndex, "ITEM", _
                CellText(inventoryTable, inventoryIndex, "ITEM")
    If CellText(aggregateTable, aggregateIndex, "ITEM") = "" Then
        SetCellText aggregateTable, aggregateIndex, "ITEM", _
                    CellText(inventoryTable, inventoryIndex, "ItemName")
    End If
    SetCellText aggregateTable, aggregateIndex, "UOM", _
                CellText(inventoryTable, inventoryIndex, "UOM")
    SetCellValue aggregateTable, aggregateIndex, "QUANTITY", qty
    If Trim$(locationOverride) = "" Then _
        locationOverride = CellText(inventoryTable, inventoryIndex, "LOCATION")
    SetCellText aggregateTable, aggregateIndex, "LOCATION", locationOverride
    SetCellText aggregateTable, aggregateIndex, "LOT_NUMBER", lotNumber
    SetCellText aggregateTable, aggregateIndex, "Condition", conditionValue
    SetCellText aggregateTable, aggregateIndex, "RETURN_REASON", returnReason
    SetCellText aggregateTable, aggregateIndex, "System_Key", receivingSystemKey
    SetCellText aggregateTable, aggregateIndex, "EventId", eventId
    SetCellText aggregateTable, aggregateIndex, "WorkflowState", "STAGED"
End Sub

Private Function FindExistingStagingRecord(ByVal stagingTable As ListObject, _
                                           ByVal refNumber As String, _
                                           ByVal sourceSystemKey As String, _
                                           ByVal locationValue As String, _
                                           ByVal lotNumber As String, _
                                           ByVal conditionValue As String, _
                                           ByVal receiptType As String, _
                                           ByVal returnReason As String) As Long
    Dim recordIndex As Long

    If stagingTable Is Nothing Or stagingTable.DataBodyRange Is Nothing Then Exit Function
    For recordIndex = 1 To stagingTable.ListRows.Count
        If StrComp(CellText(stagingTable, recordIndex, "REF_NUMBER"), _
                   refNumber, vbTextCompare) = 0 _
           And StrComp(CellText(stagingTable, recordIndex, "Source_System_Key"), _
                       sourceSystemKey, vbBinaryCompare) = 0 _
           And StrComp(CellText(stagingTable, recordIndex, "LOCATION"), _
                       locationValue, vbTextCompare) = 0 _
           And StrComp(CellText(stagingTable, recordIndex, "LOT_NUMBER"), _
                       lotNumber, vbTextCompare) = 0 _
           And StrComp(CellText(stagingTable, recordIndex, "Condition"), _
                       conditionValue, vbTextCompare) = 0 _
           And StrComp(CellText(stagingTable, recordIndex, "RECEIPT_TYPE"), _
                       receiptType, vbTextCompare) = 0 _
           And StrComp(CellText(stagingTable, recordIndex, "RETURN_REASON"), _
                       returnReason, vbTextCompare) = 0 Then
            FindExistingStagingRecord = recordIndex
            Exit Function
        End If
    Next recordIndex
End Function

Private Function NormalizeReceivingCondition(ByVal conditionValue As String) As String
    conditionValue = UCase$(Trim$(conditionValue))
    If conditionValue = "" Then conditionValue = "GOOD"
    Select Case conditionValue
        Case "GOOD", "BAD", "DAMAGED", "EXPIRED", "REJECTED"
            NormalizeReceivingCondition = conditionValue
    End Select
End Function

Private Function ReceivingErrorSource(ByVal errorSource As String) As String
    errorSource = Replace$(errorSource, vbCr, " ")
    errorSource = Replace$(errorSource, vbLf, " ")
    errorSource = Replace$(errorSource, ";", ",")
    errorSource = Trim$(errorSource)
    If errorSource = "" Then errorSource = "(none)"
    ReceivingErrorSource = errorSource
End Function

Private Function ReceivingInventoryLotNumber(ByVal inventoryTable As ListObject, _
                                             ByVal rowIndex As Long) As String
    ReceivingInventoryLotNumber = CellText(inventoryTable, rowIndex, "LOT_NUMBER")
    If ReceivingInventoryLotNumber = "" Then
        ReceivingInventoryLotNumber = ExtractReceivingJsonString( _
            CellText(inventoryTable, rowIndex, "AttributesJson"), "LOT_NUMBER")
    End If
End Function

Private Function ExtractReceivingJsonString(ByVal jsonText As String, _
                                            ByVal propertyName As String) As String
    Dim propertyToken As String
    Dim propertyPos As Long
    Dim colonPos As Long
    Dim valueStart As Long
    Dim valueEnd As Long

    jsonText = Trim$(jsonText)
    propertyToken = Chr$(34) & propertyName & Chr$(34)
    propertyPos = InStr(1, jsonText, propertyToken, vbTextCompare)
    If propertyPos = 0 Then Exit Function
    colonPos = InStr(propertyPos + Len(propertyToken), jsonText, ":", vbBinaryCompare)
    If colonPos = 0 Then Exit Function
    valueStart = InStr(colonPos + 1, jsonText, Chr$(34), vbBinaryCompare)
    If valueStart = 0 Then Exit Function
    valueEnd = InStr(valueStart + 1, jsonText, Chr$(34), vbBinaryCompare)
    If valueEnd = 0 Then Exit Function
    ExtractReceivingJsonString = Mid$(jsonText, valueStart + 1, valueEnd - valueStart - 1)
End Function

Private Function BuildReceivingAggregateGroupKey(ByVal receiptType As String, _
                                                 ByVal itemCode As String, _
                                                 ByVal uomValue As String, _
                                                 ByVal locationValue As String, _
                                                 ByVal lotNumber As String, _
                                                 ByVal conditionValue As String) As String
    BuildReceivingAggregateGroupKey = _
        UCase$(Trim$(receiptType)) & Chr$(30) & _
        UCase$(Trim$(itemCode)) & Chr$(30) & _
        UCase$(Trim$(uomValue)) & Chr$(30) & _
        UCase$(Trim$(locationValue)) & Chr$(30) & _
        UCase$(Trim$(lotNumber)) & Chr$(30) & _
        UCase$(Trim$(conditionValue))
End Function

Private Function AppendDistinctReceivingValue(ByVal existingValues As String, _
                                               ByVal valueToAdd As String) As String
    Dim candidate As Variant

    existingValues = Trim$(existingValues)
    valueToAdd = Trim$(valueToAdd)
    If valueToAdd = "" Then
        AppendDistinctReceivingValue = existingValues
        Exit Function
    End If
    For Each candidate In Split(existingValues, ",")
        If StrComp(Trim$(CStr(candidate)), valueToAdd, vbTextCompare) = 0 Then
            AppendDistinctReceivingValue = existingValues
            Exit Function
        End If
    Next candidate
    If existingValues = "" Then
        AppendDistinctReceivingValue = valueToAdd
    Else
        AppendDistinctReceivingValue = existingValues & ", " & valueToAdd
    End If
End Function

Private Function ResolveReceivingWorkbook(Optional ByVal preferredWb As Workbook = Nothing) As Workbook
    If Not preferredWb Is Nothing Then
        If Not preferredWb.IsAddin Then
            Set ResolveReceivingWorkbook = preferredWb
            Exit Function
        End If
    End If
End Function

Private Function WorkbookSheet(ByVal wb As Workbook, _
                               ByVal sheetName As String) As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set WorkbookSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindTable(ByVal wb As Workbook, _
                           ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function FindTableRecord(ByVal targetTable As ListObject, _
                                 ByVal columnName As String, _
                                 ByVal matchValue As String) As Long
    Dim recordIndex As Long
    Dim columnNumber As Long

    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    columnNumber = ColumnIndex(targetTable, columnName)
    If columnNumber = 0 Then Exit Function
    For recordIndex = 1 To targetTable.ListRows.Count
        If StrComp(Trim$(CStr(targetTable.DataBodyRange.Cells(recordIndex, columnNumber).Value)), _
                   Trim$(matchValue), vbTextCompare) = 0 Then
            FindTableRecord = recordIndex
            Exit Function
        End If
    Next recordIndex
End Function

Private Function FirstBlankOrNewRecord(ByVal targetTable As ListObject) As ListRow
    Dim candidate As ListRow
    Dim column As ListColumn
    Dim hasValue As Boolean

    For Each candidate In targetTable.ListRows
        hasValue = False
        For Each column In targetTable.ListColumns
            If Trim$(CStr(candidate.Range.Cells(1, column.Index).Value)) <> "" Then
                hasValue = True
                Exit For
            End If
        Next column
        If Not hasValue Then
            Set FirstBlankOrNewRecord = candidate
            Exit Function
        End If
    Next candidate
    Set FirstBlankOrNewRecord = targetTable.ListRows.Add
End Function

Private Function ColumnIndex(ByVal targetTable As ListObject, _
                             ByVal columnName As String) As Long
    Dim column As ListColumn

    If targetTable Is Nothing Then Exit Function
    For Each column In targetTable.ListColumns
        If StrComp(Trim$(column.Name), Trim$(columnName), vbTextCompare) = 0 Then
            ColumnIndex = column.Index
            Exit Function
        End If
    Next column
End Function

Private Function CellText(ByVal targetTable As ListObject, _
                          ByVal recordIndex As Long, _
                          ByVal columnName As String) As String
    Dim columnNumber As Long
    Dim valueIn As Variant

    columnNumber = ColumnIndex(targetTable, columnName)
    If columnNumber = 0 Or targetTable.DataBodyRange Is Nothing Then Exit Function
    valueIn = targetTable.DataBodyRange.Cells(recordIndex, columnNumber).Value
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    CellText = Trim$(CStr(valueIn))
End Function

Private Function CellNumber(ByVal targetTable As ListObject, _
                            ByVal recordIndex As Long, _
                            ByVal columnName As String) As Double
    Dim valueIn As String

    valueIn = CellText(targetTable, recordIndex, columnName)
    If IsNumeric(valueIn) Then CellNumber = CDbl(valueIn)
End Function

Private Sub SetCellText(ByVal targetTable As ListObject, _
                        ByVal recordIndex As Long, _
                        ByVal columnName As String, _
                        ByVal valueText As String)
    SetCellValue targetTable, recordIndex, columnName, valueText
End Sub

Private Sub SetCellValue(ByVal targetTable As ListObject, _
                         ByVal recordIndex As Long, _
                         ByVal columnName As String, _
                         ByVal valueIn As Variant)
    Dim columnNumber As Long

    columnNumber = ColumnIndex(targetTable, columnName)
    If columnNumber = 0 Then
        Err.Raise vbObjectError + 7681, "modTS_Received.SetCellValue", _
                  "Receiving column is missing: " & columnName
    End If
    targetTable.DataBodyRange.Cells(recordIndex, columnNumber).Value = valueIn
End Sub

Private Function CanHideWorksheet(ByVal wb As Workbook, _
                                  ByVal sheetToHide As Worksheet) As Boolean
    Dim ws As Worksheet
    Dim visibleCount As Long

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then visibleCount = visibleCount + 1
    Next ws
    CanHideWorksheet = (sheetToHide.Visible <> xlSheetVisible Or visibleCount > 1)
End Function

Private Sub EnsureReceivingButtons(ByVal ws As Worksheet)
    Dim button As Shape
    Dim anchor As Range

    If ws Is Nothing Then Exit Sub
    DeleteReceivingActionButton ws, "btnUndoMacro"
    DeleteReceivingActionButton ws, "btnRedoMacro"
    On Error Resume Next
    Set button = ws.Shapes("btnConfirmWrites")
    On Error GoTo 0
    Set anchor = ws.Range("C1")
    If button Is Nothing Then
        Set button = ws.Shapes.AddFormControl( _
            xlButtonControl, anchor.Left, 6, 118, 20)
        button.Name = "btnConfirmWrites"
    End If
    button.Left = anchor.Left
    button.Top = 6
    button.Width = 118
    button.Height = 20
    button.TextFrame.Characters.Text = "Confirm Writes"
    button.OnAction = "'" & ThisWorkbook.Name & "'!modTS_Received.ConfirmWrites"
End Sub

Private Sub DeleteReceivingActionButton(ByVal ws As Worksheet, _
                                        ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Public Sub ShowReceivingMessage(ByVal messageText As String, ByVal style As VbMsgBoxStyle)
    modReceivingActivityAction.ShowMessage messageText, style
End Sub
