Attribute VB_Name = "modReceivingPostingService"
Option Explicit

Private Const SHEET_RECEIVING As String = "ReceivedTally"
Private Const SHEET_RECEIVED_LOG As String = "ReceivedLog"
Private Const TABLE_STAGING As String = "ReceivedTally"
Private Const TABLE_RECEIVED_TALLY As String = "ReceivedTally"
Private Const TABLE_AGGREGATE As String = "AggregateReceived"
Private Const TABLE_LOG As String = "ReceivedLog"

Public Function ExecuteConfirmWrites(ByVal operatorWb As Workbook, _
                                     Optional ByRef report As String = "", _
                                     Optional ByRef outcome As String = "", _
                                     Optional ByRef sourceEventIds As String = "", _
                                     Optional ByRef submissionState As String = "Unknown") As Boolean
    On Error GoTo Failed

    Dim aggregateTable As ListObject
    Dim stagingTable As ListObject
    Dim logTable As ListObject
    Dim states As Collection
    Dim state As cReceivingWorkflowState
    Dim rowIndex As Long
    Dim userId As String
    Dim queueError As String
    Dim queuePayload As String
    Dim runtimeReport As String
    Dim queuedWorkHandled As Boolean
    Dim queuedCount As Long
    Dim persistenceSummary As String

    outcome = "REJECTED": sourceEventIds = "": submissionState = "Unknown"
    If operatorWb Is Nothing Then
        report = "Receiving operator workbook was not provided."
        Exit Function
    End If
    If Not WorkbookIsOpenReceivingService(operatorWb) Then
        report = "The captured Receiving operator workbook is no longer open."
        Exit Function
    End If
    If Not modRoleUiAccess.CanCurrentUserPerformCapability( _
        "RECEIVE_POST", "", "", "", report) Then
        outcome = "DENIED"
        Exit Function
    End If

    Set stagingTable = FindTableReceivingService(operatorWb, TABLE_STAGING)
    Set aggregateTable = FindTableReceivingService(operatorWb, TABLE_AGGREGATE)
    Set logTable = FindTableReceivingService(operatorWb, TABLE_LOG)
    If stagingTable Is Nothing Or aggregateTable Is Nothing Or logTable Is Nothing Then
        report = "Receiving staging or log tables are missing."
        Exit Function
    End If
    For rowIndex = stagingTable.ListRows.Count To 1 Step -1
        If CellText(stagingTable, rowIndex, "System_Key") = "" _
           And CellText(stagingTable, rowIndex, "ITEM_CODE") = "" _
           And CellText(stagingTable, rowIndex, "REF_NUMBER") = "" _
           And CellText(stagingTable, rowIndex, "EventId") = "" _
           And Abs(CellNumber(stagingTable, rowIndex, "QUANTITY")) < 0.0000001 Then
            stagingTable.ListRows(rowIndex).Delete
        End If
    Next rowIndex
    If stagingTable.DataBodyRange Is Nothing Then
        report = "ReceivedTally has no rows to confirm."
        Exit Function
    End If

    userId = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If userId = "" Then
        report = "Unable to resolve current user identity."
        Exit Function
    End If

    Set states = BuildValidatedStates(stagingTable, report)
    If states Is Nothing Then Exit Function

    For rowIndex = 1 To states.Count
        Set state = states(rowIndex)
        If StrComp(state.CurrentState, state.StateValidated, vbBinaryCompare) = 0 Then
            If queuePayload <> "" Then queuePayload = queuePayload & vbLf
            queuePayload = queuePayload & BuildReceivingQueueJson(state)
        End If
    Next rowIndex
    queueError = ""
    sourceEventIds = ReceivingSourceEventIds(states)
    outcome = "FAILED"
    If queuePayload <> "" Then
        If Not modRoleEventWriter.QueueReceiveEventBatchServer( _
            "", "", userId, queuePayload, queueError, queuedCount) Then
            report = "Inbox queue failed for staged receipt batch: " & queueError
            Exit Function
        End If
        submissionState = "Submitted"
        For rowIndex = 1 To states.Count
            Set state = states(rowIndex)
            If StrComp(state.CurrentState, state.StateValidated, vbBinaryCompare) = 0 Then
                state.MarkSubmitted
                WriteWorkflowState stagingTable, rowIndex, state
            End If
        Next rowIndex
    End If

    submissionState = "Submitted"
    outcome = "PENDING"
    If Not modOperationsPrimitiveBridge.RunBatchAndRefreshOperatorWorkbook( _
        operatorWb.Name, "", "LOCAL", runtimeReport, False, queuedWorkHandled) Then
        report = "Receiving rows remain submitted; processor application or snapshot refresh " & _
                 "did not complete. " & runtimeReport
        Exit Function
    End If

    For rowIndex = 1 To states.Count
        Set state = states(rowIndex)
        If StrComp(state.CurrentState, state.StateSubmitted, vbBinaryCompare) = 0 Then
            state.MarkProcessorApplied
            WriteWorkflowState stagingTable, rowIndex, state
        End If
        If StrComp(state.CurrentState, state.StateProcessorApplied, vbBinaryCompare) = 0 Then
            state.MarkSnapshotRefreshed
            WriteWorkflowState stagingTable, rowIndex, state
        End If
        AppendReceivedLog logTable, stagingTable, rowIndex, state, userId
    Next rowIndex

    ClearReceivingStaging stagingTable, aggregateTable
    For Each state In states
        If StrComp(state.CurrentState, state.StateSnapshotRefreshed, vbBinaryCompare) = 0 Then
            state.MarkReady
        End If
    Next state

    report = "OK|State=READY|Lines=" & CStr(states.Count) & _
             "|Queued=" & CStr(queuedCount) & _
             "|ProcessorApplied=True|SnapshotRefreshed=True|StagingCleared=True"
    If queuedCount > 0 Then
        persistenceSummary = "receiving inbox batch saved"
    Else
        persistenceSummary = "previously submitted receiving inbox retained"
    End If
    If queuedWorkHandled Then
        persistenceSummary = persistenceSummary & "; processor durability saves retained"
    Else
        persistenceSummary = persistenceSummary & "; no new processor write required"
    End If
    report = report & vbCrLf & "Persistence summary: " & persistenceSummary & "."
    outcome = "CONFIRMED"
    ExecuteConfirmWrites = True
    Exit Function

Failed:
    outcome = "FAILED"
    report = "Receiving Confirm Writes failed: " & Err.Description
End Function

' Optional observation preparation must not prevent the owning business command.
Private Function ReceivingSourceEventIds(ByVal states As Collection) As String
    Dim state As cReceivingWorkflowState, ids As String
    On Error GoTo Unavailable
    For Each state In states
        If ids <> "" Then ids = ids & vbLf
        ids = ids & state.EventId
    Next state
    ReceivingSourceEventIds = ids
Unavailable:
End Function

Public Sub ClearReceivingStaging(ByVal stagingTable As ListObject, _
                                 ByVal aggregateTable As ListObject, Optional ByRef changed As Boolean = False)
    changed = False
    If Not stagingTable Is Nothing Then
        If Not stagingTable.DataBodyRange Is Nothing Then
            stagingTable.DataBodyRange.Delete
            changed = True
        End If
    End If
    If Not aggregateTable Is Nothing Then
        If Not aggregateTable.DataBodyRange Is Nothing Then
            aggregateTable.DataBodyRange.Delete
            changed = True
        End If
    End If
End Sub

Public Function ReceivingWorkflowContractProbe(ByVal contractName As String) As String
    On Error GoTo Failed

    Dim state As cReceivingWorkflowState
    Dim firstEventId As String

    Set state = New cReceivingWorkflowState
    Select Case UCase$(Trim$(contractName))
        Case "SEQUENCE"
            state.Initialize "SYS-RECEIVING-PROBE", "", "SKU-PROBE", 2, "DOCK", "probe"
            state.MarkValidated
            state.MarkSubmitted
            state.MarkProcessorApplied
            state.MarkSnapshotRefreshed
            state.MarkReady
            ReceivingWorkflowContractProbe = "OK|State=" & state.CurrentState
        Case "IDENTITY"
            state.Initialize "SYS-RECEIVING-PROBE", "", "SKU-PROBE", 2, "DOCK", "probe"
            state.EnsureEventIdentities
            firstEventId = state.EventId
            state.EnsureEventIdentities
            If firstEventId = "" Or StrComp(firstEventId, state.EventId, vbBinaryCompare) <> 0 Then
                ReceivingWorkflowContractProbe = "FAIL|Event identity changed."
            Else
                ReceivingWorkflowContractProbe = _
                    "OK|System_Key=" & state.SystemKey & "|EventIdStable=True"
            End If
        Case "MISSING_SYSTEM_KEY"
            state.Initialize "", "", "SKU-PROBE", 2, "DOCK", "probe"
            state.EnsureEventIdentities
            ReceivingWorkflowContractProbe = "FAIL|Blank System_Key was accepted."
        Case Else
            ReceivingWorkflowContractProbe = "FAIL|UnknownContract=" & contractName
    End Select
    Exit Function

Failed:
    If UCase$(Trim$(contractName)) = "MISSING_SYSTEM_KEY" Then
        ReceivingWorkflowContractProbe = "OK|Rejected=True|Message=" & Err.Description
    Else
        ReceivingWorkflowContractProbe = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    End If
End Function

Private Function BuildValidatedStates(ByVal stagingTable As ListObject, _
                                      ByRef report As String) As Collection
    On Error GoTo Failed

    Dim states As Collection
    Dim state As cReceivingWorkflowState
    Dim rowIndex As Long
    Dim stateValue As String

    If Not RequiredColumnsPresent(stagingTable, _
        Array("REF_NUMBER", "RECEIPT_TYPE", "ITEM_CODE", "ITEMS", "UOM", "QUANTITY", "LOCATION", "LOT_NUMBER", "Condition", "RETURN_REASON", _
              "System_Key", "EventId", "WorkflowState")) Then
        report = "ReceivedTally is missing required Receiving workflow columns."
        Exit Function
    End If

    Set states = New Collection
    For rowIndex = 1 To stagingTable.ListRows.Count
        Set state = New cReceivingWorkflowState
        stateValue = CellText(stagingTable, rowIndex, "WorkflowState")
        state.Initialize _
            CellText(stagingTable, rowIndex, "System_Key"), _
            CellText(stagingTable, rowIndex, "EventId"), _
            CellText(stagingTable, rowIndex, "ITEM_CODE"), _
            CellNumber(stagingTable, rowIndex, "QUANTITY"), _
            CellText(stagingTable, rowIndex, "LOCATION"), _
            BuildEventNote(stagingTable, rowIndex), _
            CellText(stagingTable, rowIndex, "Condition"), _
            BuildReceivingAttributesJson(stagingTable, rowIndex), stateValue, _
            EventTypeForReceiptType(CellText(stagingTable, rowIndex, "RECEIPT_TYPE"))

        Select Case state.CurrentState
            Case state.StateStaged
                state.MarkValidated
            Case state.StateValidated, state.StateSubmitted, _
                 state.StateProcessorApplied, state.StateSnapshotRefreshed
                state.EnsureEventIdentities
            Case Else
                report = "ReceivedTally contains an invalid completed workflow row."
                Exit Function
        End Select
        WriteWorkflowState stagingTable, rowIndex, state
        states.Add state
    Next rowIndex
    Set BuildValidatedStates = states
    Exit Function

Failed:
    report = "Receiving validation failed: " & Err.Description
End Function

Private Function BuildReceivingAttributesJson(ByVal aggregateTable As ListObject, _
                                              ByVal rowIndex As Long) As String
    Dim lotNumber As String
    Dim receiptType As String
    Dim returnReason As String
    Dim result As String

    lotNumber = CellText(aggregateTable, rowIndex, "LOT_NUMBER")
    receiptType = CellText(aggregateTable, rowIndex, "RECEIPT_TYPE")
    If receiptType = "" Then receiptType = "RECEIPT"
    returnReason = CellText(aggregateTable, rowIndex, "RETURN_REASON")
    result = "{""RECEIPT_TYPE"":""" & EscapeJsonReceiving(receiptType) & """"
    If lotNumber <> "" Then
        result = result & ",""LOT_NUMBER"":""" & EscapeJsonReceiving(lotNumber) & """"
    End If
    If returnReason <> "" Then
        result = result & ",""RETURN_REASON"":""" & EscapeJsonReceiving(returnReason) & """"
    End If
    BuildReceivingAttributesJson = result & "}"
End Function

Private Function EscapeJsonReceiving(ByVal textIn As String) As String
    EscapeJsonReceiving = Replace$(textIn, "\", "\\")
    EscapeJsonReceiving = Replace$(EscapeJsonReceiving, Chr$(34), "\" & Chr$(34))
    EscapeJsonReceiving = Replace$(EscapeJsonReceiving, vbCrLf, "\n")
    EscapeJsonReceiving = Replace$(EscapeJsonReceiving, vbCr, "\n")
    EscapeJsonReceiving = Replace$(EscapeJsonReceiving, vbLf, "\n")
    EscapeJsonReceiving = Replace$(EscapeJsonReceiving, vbTab, "\t")
End Function

Private Function BuildReceivingQueueJson(ByVal state As cReceivingWorkflowState) As String
    BuildReceivingQueueJson = _
        "{""EventType"":""" & EscapeJsonReceiving(state.EventType) & _
        """,""EventID"":""" & EscapeJsonReceiving(state.EventId) & _
        """,""System_Key"":""" & EscapeJsonReceiving(state.SystemKey) & _
        """,""SKU"":""" & EscapeJsonReceiving(state.Sku) & _
        """,""Qty"":" & Replace$(CStr(state.Qty), Application.DecimalSeparator, ".") & _
        ",""Location"":""" & EscapeJsonReceiving(state.Location) & _
        """,""Note"":""" & EscapeJsonReceiving(state.Note) & _
        """,""Condition"":""" & EscapeJsonReceiving(state.ConditionValue) & _
        """,""AttributesJson"":""" & EscapeJsonReceiving(state.AttributesJson) & """}"
End Function

Private Function EventTypeForReceiptType(ByVal receiptType As String) As String
    receiptType = UCase$(Trim$(receiptType))
    Select Case receiptType
        Case "RETURN", "DUMP"
            EventTypeForReceiptType = receiptType
        Case Else
            EventTypeForReceiptType = "RECEIVE"
    End Select
End Function

Private Sub WriteWorkflowState(ByVal targetTable As ListObject, _
                               ByVal rowIndex As Long, _
                               ByVal state As cReceivingWorkflowState)
    SetCellText targetTable, rowIndex, "System_Key", state.SystemKey
    SetCellText targetTable, rowIndex, "EventId", state.EventId
    SetCellText targetTable, rowIndex, "WorkflowState", state.CurrentState
End Sub

Private Sub WriteWorkflowStateBySystemKey(ByVal targetTable As ListObject, _
                                          ByVal state As cReceivingWorkflowState)
    Dim rowIndex As Long

    rowIndex = FindTableRow(targetTable, "System_Key", state.SystemKey)
    If rowIndex = 0 Then Exit Sub
    WriteWorkflowState targetTable, rowIndex, state
End Sub

Private Sub AppendReceivedLog(ByVal logTable As ListObject, _
                              ByVal stagingTable As ListObject, _
                              ByVal stagingIndex As Long, _
                              ByVal state As cReceivingWorkflowState, _
                              ByVal userId As String)
    Dim targetRow As ListRow
    Dim logIndex As Long

    If FindTableRow(logTable, "EventId", state.EventId) > 0 Then Exit Sub
    Set targetRow = FirstBlankOrNewRow(logTable)
    logIndex = targetRow.Index
    SetCellText logTable, logIndex, "SNAPSHOT_ID", state.EventId
    SetCellValue logTable, logIndex, "ENTRY_DATE", Now
    SetCellText logTable, logIndex, "USER", userId
    SetCellText logTable, logIndex, "RECEIPT_TYPE", _
                CellText(stagingTable, stagingIndex, "RECEIPT_TYPE")
    SetCellText logTable, logIndex, "REF_NUMBER", _
                CellText(stagingTable, stagingIndex, "REF_NUMBER")
    SetCellText logTable, logIndex, "ITEMS", _
                CellText(stagingTable, stagingIndex, "ITEMS")
    SetCellValue logTable, logIndex, "QUANTITY", state.Qty
    SetCellText logTable, logIndex, "UOM", _
                CellText(stagingTable, stagingIndex, "UOM")
    SetCellText logTable, logIndex, "VENDOR", _
                CellText(stagingTable, stagingIndex, "VENDOR")
    SetCellText logTable, logIndex, "LOCATION", state.Location
    SetCellText logTable, logIndex, "LOT_NUMBER", _
                CellText(stagingTable, stagingIndex, "LOT_NUMBER")
    SetCellText logTable, logIndex, "Condition", state.ConditionValue
    SetCellText logTable, logIndex, "RETURN_REASON", _
                CellText(stagingTable, stagingIndex, "RETURN_REASON")
    SetCellText logTable, logIndex, "ITEM_CODE", state.Sku
    SetCellText logTable, logIndex, "System_Key", state.SystemKey
    SetCellText logTable, logIndex, "EventId", state.EventId
End Sub

Private Function BuildEventNote(ByVal stagingTable As ListObject, _
                                ByVal rowIndex As Long) As String
    BuildEventNote = "RECEIPT_TYPE=" & CellText(stagingTable, rowIndex, "RECEIPT_TYPE") & _
                     "; REF_NUMBER=" & CellText(stagingTable, rowIndex, "REF_NUMBER") & _
                     "; CONDITION=" & CellText(stagingTable, rowIndex, "Condition")
    If CellText(stagingTable, rowIndex, "ITEMS") <> "" Then
        BuildEventNote = BuildEventNote & _
                         "; ITEM=" & CellText(stagingTable, rowIndex, "ITEMS")
    End If
    If CellText(stagingTable, rowIndex, "VENDOR") <> "" Then
        BuildEventNote = BuildEventNote & _
                         "; VENDORS=" & CellText(stagingTable, rowIndex, "VENDOR")
    End If
    If CellText(stagingTable, rowIndex, "LOT_NUMBER") <> "" Then
        BuildEventNote = BuildEventNote & _
                         "; LOT_NUMBER=" & CellText(stagingTable, rowIndex, "LOT_NUMBER")
    End If
    If CellText(stagingTable, rowIndex, "RETURN_REASON") <> "" Then
        BuildEventNote = BuildEventNote & _
                         "; RETURN_REASON=" & CellText(stagingTable, rowIndex, "RETURN_REASON")
    End If
End Function

Private Function FindTableReceivingService(ByVal wb As Workbook, _
                                           ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    If StrComp(tableName, TABLE_LOG, vbTextCompare) = 0 Then
        Set ws = wb.Worksheets(SHEET_RECEIVED_LOG)
    Else
        Set ws = wb.Worksheets(SHEET_RECEIVING)
    End If
    If Not ws Is Nothing Then
        Set FindTableReceivingService = ws.ListObjects(tableName)
    End If
    On Error GoTo 0
End Function

Private Function RequiredColumnsPresent(ByVal targetTable As ListObject, _
                                        ByVal names As Variant) As Boolean
    Dim nameValue As Variant
    Dim requiredColumn As ListColumn

    For Each nameValue In names
        Set requiredColumn = FindColumnReceivingService(targetTable, CStr(nameValue))
        If requiredColumn Is Nothing Then Exit Function
    Next nameValue
    RequiredColumnsPresent = True
End Function

Private Function FindColumnReceivingService(ByVal targetTable As ListObject, _
                                            ByVal columnName As String) As ListColumn
    Dim column As ListColumn

    If targetTable Is Nothing Then Exit Function
    For Each column In targetTable.ListColumns
        If StrComp(Trim$(column.Name), Trim$(columnName), vbTextCompare) = 0 Then
            Set FindColumnReceivingService = column
            Exit Function
        End If
    Next column
End Function

Private Function CellText(ByVal targetTable As ListObject, _
                          ByVal rowIndex As Long, _
                          ByVal columnName As String) As String
    Dim valueIn As Variant
    Dim targetColumn As ListColumn

    Set targetColumn = FindColumnReceivingService(targetTable, columnName)
    If targetColumn Is Nothing Then Exit Function
    valueIn = targetTable.DataBodyRange.Cells(rowIndex, targetColumn.Index).Value
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    CellText = Trim$(CStr(valueIn))
End Function

Private Function CellNumber(ByVal targetTable As ListObject, _
                            ByVal rowIndex As Long, _
                            ByVal columnName As String) As Double
    Dim valueIn As Variant
    Dim targetColumn As ListColumn

    Set targetColumn = FindColumnReceivingService(targetTable, columnName)
    If targetColumn Is Nothing Then Exit Function
    valueIn = targetTable.DataBodyRange.Cells(rowIndex, targetColumn.Index).Value
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If IsNumeric(valueIn) Then CellNumber = CDbl(valueIn)
End Function

Private Sub SetCellText(ByVal targetTable As ListObject, _
                        ByVal rowIndex As Long, _
                        ByVal columnName As String, _
                        ByVal valueText As String)
    SetCellValue targetTable, rowIndex, columnName, valueText
End Sub

Private Sub SetCellValue(ByVal targetTable As ListObject, _
                         ByVal rowIndex As Long, _
                         ByVal columnName As String, _
                         ByVal valueIn As Variant)
    Dim targetColumn As ListColumn

    Set targetColumn = FindColumnReceivingService(targetTable, columnName)
    If targetColumn Is Nothing Then
        Err.Raise vbObjectError + 7671, "modReceivingPostingService.SetCellValue", _
                  "Receiving column is missing: " & columnName
    End If
    targetTable.DataBodyRange.Cells(rowIndex, targetColumn.Index).Value = valueIn
End Sub

Private Function FindTableRow(ByVal targetTable As ListObject, _
                              ByVal columnName As String, _
                              ByVal matchValue As String) As Long
    Dim rowIndex As Long
    Dim targetColumn As ListColumn
    Dim valueIn As Variant

    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    Set targetColumn = FindColumnReceivingService(targetTable, columnName)
    If targetColumn Is Nothing Then Exit Function
    For rowIndex = 1 To targetTable.ListRows.Count
        valueIn = targetTable.DataBodyRange.Cells(rowIndex, targetColumn.Index).Value
        If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then GoTo ContinueRow
        If StrComp(Trim$(CStr(valueIn)), _
                   Trim$(matchValue), vbBinaryCompare) = 0 Then
            FindTableRow = rowIndex
            Exit Function
        End If
ContinueRow:
    Next rowIndex
End Function

Private Function FirstBlankOrNewRow(ByVal targetTable As ListObject) As ListRow
    Dim candidate As ListRow

    If targetTable.DataBodyRange Is Nothing Then
        Set FirstBlankOrNewRow = targetTable.ListRows.Add
        Exit Function
    End If
    For Each candidate In targetTable.ListRows
        If Application.WorksheetFunction.CountA(candidate.Range) = 0 Then
            Set FirstBlankOrNewRow = candidate
            Exit Function
        End If
    Next candidate
    Set FirstBlankOrNewRow = targetTable.ListRows.Add
End Function

Private Function WorkbookIsOpenReceivingService(ByVal wb As Workbook) As Boolean
    Dim candidate As Workbook

    If wb Is Nothing Then Exit Function
    For Each candidate In Application.Workbooks
        If candidate Is wb Then
            WorkbookIsOpenReceivingService = True
            Exit Function
        End If
    Next candidate
End Function
