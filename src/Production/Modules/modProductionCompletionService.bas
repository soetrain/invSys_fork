Attribute VB_Name = "modProductionCompletionService"
Option Explicit

Private Const SESSION_FORMAT As String = "PSS1"
Private Const SESSION_WORKBOOK_NAME As String = "__invSys_ProductionSession"
Private Const SESSION_WORKBOOK_CHUNK_SIZE As Long = 200
Private Const EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Private Const EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"

Public Function CreateProductionSession(Optional ByVal sessionId As String = "", _
                                        Optional ByVal designId As String = "", _
                                        Optional ByVal designVersion As String = "", _
                                        Optional ByVal qtyPlanned As Double = 0, _
                                        Optional ByVal location As String = "") As cProductionRunSession
    Dim session As cProductionRunSession

    If Trim$(sessionId) = "" Then sessionId = modRoleEventWriter.CreateSystemKey()
    Set session = New cProductionRunSession
    session.Initialize sessionId, designId, designVersion, qtyPlanned, location
    Set CreateProductionSession = session
End Function

Public Function CreateProductionCompletionResult(ByVal session As cProductionRunSession, _
                                                 Optional ByVal message As String = "") As cProductionCompletionResult
    Dim result As cProductionCompletionResult

    Set result = New cProductionCompletionResult
    result.InitializeFromSession session, message
    Set CreateProductionCompletionResult = result
End Function

Public Function CreateProductionSessionFromWorkbook(ByVal wb As Workbook, _
                                                    ByVal outputRowNumber As Long, _
                                                    Optional ByRef report As String = "") As cProductionRunSession
    Dim loCheck As ListObject
    Dim loOutput As ListObject
    Dim session As cProductionRunSession
    Dim allocations As Object
    Dim allocation As Object
    Dim key As Variant
    Dim rowIndex As Long
    Dim systemKey As String
    Dim sku As String
    Dim locationValue As String
    Dim conditionValue As String
    Dim attributesJson As String
    Dim qty As Double

    On Error GoTo FailCreate
    If wb Is Nothing Then
        report = "Operator workbook is required."
        Exit Function
    End If
    Set loCheck = FindWorkbookTable(wb, "Prod_invSys_Check")
    Set loOutput = FindWorkbookTable(wb, "ProductionOutput")
    If loCheck Is Nothing Or loCheck.DataBodyRange Is Nothing Then
        report = "No checked-in production input rows were found."
        Exit Function
    End If
    If loOutput Is Nothing Or loOutput.DataBodyRange Is Nothing Then
        report = "ProductionOutput staging was not found."
        Exit Function
    End If
    If outputRowNumber < 1 Or outputRowNumber > loOutput.ListRows.Count Then
        report = "Select a valid Production Output row before completing the run."
        Exit Function
    End If

    Set allocations = CreateObject("Scripting.Dictionary")
    allocations.CompareMode = vbTextCompare
    For rowIndex = 1 To loCheck.ListRows.Count
        systemKey = Trim$(TableText(loCheck, rowIndex, "System_Key"))
        qty = TableNumber(loCheck, rowIndex, "USED")
        If systemKey = "" And qty <= 0 Then GoTo NextInput
        If systemKey = "" Then
            report = "Checked-in Production input row " & CStr(rowIndex) & _
                     " is missing immutable System_Key identity."
            Exit Function
        End If
        If qty <= 0 Then GoTo NextInput
        sku = TableText(loCheck, rowIndex, "ITEM_CODE")
        If sku = "" Then sku = TableText(loCheck, rowIndex, "SKU")
        locationValue = TableText(loCheck, rowIndex, "LOCATION")
        If sku = "" Then sku = InventoryValueBySystemKey(wb, systemKey, "ITEM_CODE")
        If sku = "" Then sku = InventoryValueBySystemKey(wb, systemKey, "SKU")
        If locationValue = "" Then locationValue = InventoryValueBySystemKey(wb, systemKey, "LOCATION")
        If allocations.Exists(systemKey) Then
            Set allocation = allocations(systemKey)
            allocation("Qty") = CDbl(allocation("Qty")) + qty
        Else
            Set allocation = CreateObject("Scripting.Dictionary")
            allocation.CompareMode = vbTextCompare
            allocation("Qty") = qty
            allocation("SKU") = sku
            allocation("Location") = locationValue
            allocations.Add systemKey, allocation
        End If
NextInput:
    Next rowIndex
    If allocations.Count = 0 Then
        report = "No checked-in production input quantities were found."
        Exit Function
    End If

    sku = TableText(loOutput, outputRowNumber, "ITEM_CODE")
    If sku = "" Then sku = TableText(loOutput, outputRowNumber, "SKU")
    qty = TableNumber(loOutput, outputRowNumber, "REAL OUTPUT")
    locationValue = TableText(loOutput, outputRowNumber, "LOCATION")
    conditionValue = TableText(loOutput, outputRowNumber, "Condition")
    attributesJson = TableText(loOutput, outputRowNumber, "AttributesJson")
    If sku = "" Then
        report = "Selected Production output is missing ITEM_CODE/SKU."
        Exit Function
    End If
    If qty <= 0 Then
        report = "Selected Production output quantity must be greater than zero."
        Exit Function
    End If

    Set session = CreateProductionSession("", "", "", qty, locationValue)
    For Each key In allocations.Keys
        Set allocation = allocations(key)
        If locationValue = "" Then locationValue = CStr(allocation("Location"))
        session.AddInputAllocation CStr(key), CDbl(allocation("Qty")), _
                                   CStr(allocation("SKU")), CStr(allocation("Location"))
    Next key
    If session.Location = "" And locationValue <> "" Then
        Set session = RecreateSessionWithLocation(session, locationValue)
    End If
    systemKey = session.EnsureOutputIdentity(sku, qty, locationValue, conditionValue, attributesJson)
    session.EnsureEventIdentities
    SetTableText loOutput, outputRowNumber, "System_Key", systemKey
    If Not SaveProductionSessionToWorkbook(wb, session, report) Then Exit Function

    Set CreateProductionSessionFromWorkbook = session
    report = "Production session prepared; SessionId=" & session.SessionId & _
             "; OutputSystemKey=" & session.OutputSystemKey
    Exit Function

FailCreate:
    report = "CreateProductionSessionFromWorkbook failed: " & Err.Description
End Function

Public Function ExecuteProductionSession(ByVal wb As Workbook, _
                                         ByVal session As cProductionRunSession, _
                                         Optional ByRef report As String = "", _
                                         Optional ByVal action As cProductionWorksheetAction = Nothing) As cProductionCompletionResult
    Dim queueReport As String
    Dim runtimeReport As String
    Dim persistenceReport As String
    Dim runtimeSucceeded As Boolean
    Dim queuedWorkHandled As Boolean

    On Error GoTo FailExecute
    If wb Is Nothing Or session Is Nothing Then
        report = "Operator workbook and Production session are required."
        Exit Function
    End If
    If Not QueueProductionSessionEvents(session, queueReport, action) Then
        report = queueReport
        SaveProductionSessionToWorkbook wb, session, persistenceReport
        Set ExecuteProductionSession = CreateProductionCompletionResult(session, report)
        Exit Function
    End If
    SaveProductionSessionToWorkbook wb, session, persistenceReport

    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    runtimeSucceeded = modOperationsPrimitiveBridge.RunBatchAndRefreshOperatorWorkbook( _
        wb.Name, "", "LOCAL", runtimeReport, True, queuedWorkHandled)
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    If runtimeSucceeded And queuedWorkHandled Then
        session.RecordProcessorResult True, True, True
        session.RecordRefreshResult True
        report = queueReport
        If runtimeReport <> "" Then report = report & "; " & runtimeReport
    ElseIf queuedWorkHandled Then
        session.RecordProcessorResult True, True, True
        session.RecordRefreshResult False, "REFRESH_FAILED", runtimeReport
        report = runtimeReport
    Else
        session.RecordProcessorResult False, False, False, "PROCESSOR_NOT_VERIFIED", runtimeReport
        report = runtimeReport
    End If
    SaveProductionSessionToWorkbook wb, session, persistenceReport
    If persistenceReport <> "" And InStr(1, persistenceReport, "failed", vbTextCompare) > 0 Then
        If report <> "" Then report = report & "; "
        report = report & persistenceReport
    End If
    Set ExecuteProductionSession = CreateProductionCompletionResult(session, report)
    Exit Function

FailExecute:
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    report = "ExecuteProductionSession failed: " & Err.Description
    If Not session Is Nothing Then
        session.RecordFailure "EXECUTION_EXCEPTION", report
        Set ExecuteProductionSession = CreateProductionCompletionResult(session, report)
    End If
End Function

Public Function BuildProductionConsumePayload(ByVal session As cProductionRunSession) As String
    Dim items As Collection
    Dim item As Object
    Dim allocation As Object
    Dim i As Long

    RequireSessionPrepared session, "BuildProductionConsumePayload"
    Set items = New Collection
    For i = 1 To session.InputCount
        Set allocation = session.InputAllocation(i)
        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = vbTextCompare
        item("System_Key") = CStr(allocation("System_Key"))
        item("SKU") = CStr(allocation("SKU"))
        item("Qty") = CDbl(allocation("Qty"))
        item("Location") = CStr(allocation("Location"))
        item("IoType") = "USED"
        items.Add item
    Next i
    BuildProductionConsumePayload = modProductionJson.BuildJsonArray(items)
End Function

Public Function BuildProductionCompletePayload(ByVal session As cProductionRunSession) As String
    Dim items As Collection
    Dim item As Object

    RequireSessionPrepared session, "BuildProductionCompletePayload"
    Set items = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("System_Key") = session.OutputSystemKey
    item("SKU") = session.OutputSku
    item("Qty") = session.OutputQty
    item("Location") = session.Location
    item("Condition") = session.OutputCondition
    item("AttributesJson") = session.OutputAttributesJson
    item("IoType") = "MADE"
    items.Add item
    BuildProductionCompletePayload = modProductionJson.BuildJsonArray(items)
End Function

Public Function QueueProductionSessionEvents(ByVal session As cProductionRunSession, _
                                             Optional ByRef report As String = "", _
                                             Optional ByVal action As cProductionWorksheetAction = Nothing) As Boolean
    Dim consumeEventId As String
    Dim completeEventId As String
    Dim errNotes As String

    On Error GoTo FailQueue
    RequireSessionPrepared session, "QueueProductionSessionEvents"
    session.EnsureEventIdentities

    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    consumeEventId = session.ConsumeEventId
    If Not modProductionCompleteActions.QueueInventory( _
            EVENT_TYPE_PROD_CONSUME, BuildProductionConsumePayload(session), _
            "PRODUCTION_SESSION_CONSUME:" & session.SessionId, consumeEventId, errNotes, action) Then
        If errNotes = "" Then errNotes = "Unable to queue the Production consume event."
        session.RecordFailure "CONSUME_QUEUE_FAILED", errNotes
        report = errNotes
        Exit Function
    End If
    If StrComp(consumeEventId, session.ConsumeEventId, vbBinaryCompare) <> 0 Then _
        Err.Raise vbObjectError + 7640, "modProductionCompletionService.QueueProductionSessionEvents", _
                  "The queued consume event did not preserve its allocated identity."
    session.MarkConsumeQueued

    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    completeEventId = session.CompleteEventId
    errNotes = ""
    If Not modProductionCompleteActions.QueueInventory( _
            EVENT_TYPE_PROD_COMPLETE, BuildProductionCompletePayload(session), _
            "PRODUCTION_SESSION_COMPLETE:" & session.SessionId, completeEventId, errNotes, action) Then
        If errNotes = "" Then errNotes = "Unable to queue the Production completion event."
        session.RecordFailure "COMPLETE_QUEUE_FAILED", errNotes
        report = errNotes
        Exit Function
    End If
    If StrComp(completeEventId, session.CompleteEventId, vbBinaryCompare) <> 0 Then _
        Err.Raise vbObjectError + 7641, "modProductionCompletionService.QueueProductionSessionEvents", _
                  "The queued completion event did not preserve its allocated identity."
    session.MarkCompleteQueued

    QueueProductionSessionEvents = True
    report = "ConsumeEvent=" & session.ConsumeEventId & _
             "; CompleteEvent=" & session.CompleteEventId & _
             "; OutputSystemKey=" & session.OutputSystemKey
    Exit Function

FailQueue:
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    report = "QueueProductionSessionEvents failed: " & Err.Description
    If Not session Is Nothing Then session.RecordFailure "QUEUE_EXCEPTION", report
End Function

Public Function SerializeProductionSession(ByVal session As cProductionRunSession) As String
    Dim parts As Collection
    Dim inputs As String
    Dim allocation As Object
    Dim i As Long

    If session Is Nothing Then _
        Err.Raise vbObjectError + 7642, "modProductionCompletionService.SerializeProductionSession", _
                  "Production session is required."

    Set parts = New Collection
    parts.Add SESSION_FORMAT
    parts.Add HexEncode(session.SessionId)
    parts.Add HexEncode(session.DesignId)
    parts.Add HexEncode(session.DesignVersion)
    parts.Add InvariantNumber(session.QtyPlanned)
    parts.Add HexEncode(session.Location)
    parts.Add HexEncode(session.OutputSystemKey)
    parts.Add HexEncode(session.OutputSku)
    parts.Add InvariantNumber(session.OutputQty)
    parts.Add HexEncode(session.OutputCondition)
    parts.Add HexEncode(session.OutputAttributesJson)
    parts.Add HexEncode(session.ConsumeEventId)
    parts.Add HexEncode(session.CompleteEventId)
    parts.Add BoolToken(session.ConsumeQueued)
    parts.Add BoolToken(session.CompleteQueued)
    parts.Add BoolToken(session.ConsumeApplied)
    parts.Add BoolToken(session.CompleteApplied)
    parts.Add BoolToken(session.ProcessorVerified)
    parts.Add BoolToken(session.RefreshVerified)
    parts.Add HexEncode(session.Status)
    parts.Add HexEncode(session.FailureCode)
    parts.Add HexEncode(session.FailureMessage)
    parts.Add BoolToken(session.CompensationRequired)

    For i = 1 To session.InputCount
        Set allocation = session.InputAllocation(i)
        If inputs <> "" Then inputs = inputs & ","
        inputs = inputs & HexEncode(CStr(allocation("System_Key"))) & ":" & _
                 InvariantNumber(CDbl(allocation("Qty"))) & ":" & _
                 HexEncode(CStr(allocation("SKU"))) & ":" & _
                 HexEncode(CStr(allocation("Location")))
    Next i
    parts.Add inputs
    SerializeProductionSession = JoinCollection(parts, "|")
End Function

Public Function DeserializeProductionSession(ByVal serializedState As String, _
                                             Optional ByRef report As String = "") As cProductionRunSession
    Dim fields As Variant
    Dim inputRows As Variant
    Dim inputFields As Variant
    Dim session As cProductionRunSession
    Dim i As Long

    On Error GoTo FailDeserialize
    fields = Split(serializedState, "|")
    If UBound(fields) <> 23 Or CStr(fields(0)) <> SESSION_FORMAT Then
        report = "Unsupported or malformed Production session state."
        Exit Function
    End If

    Set session = New cProductionRunSession
    session.Initialize HexDecode(CStr(fields(1))), HexDecode(CStr(fields(2))), _
                       HexDecode(CStr(fields(3))), ParseInvariantNumber(CStr(fields(4))), _
                       HexDecode(CStr(fields(5)))

    If CStr(fields(23)) <> "" Then
        inputRows = Split(CStr(fields(23)), ",")
        For i = LBound(inputRows) To UBound(inputRows)
            inputFields = Split(CStr(inputRows(i)), ":")
            If UBound(inputFields) <> 3 Then _
                Err.Raise vbObjectError + 7643, "modProductionCompletionService.DeserializeProductionSession", _
                          "Malformed input allocation state."
            session.AddInputAllocation HexDecode(CStr(inputFields(0))), _
                                       ParseInvariantNumber(CStr(inputFields(1))), _
                                       HexDecode(CStr(inputFields(2))), _
                                       HexDecode(CStr(inputFields(3)))
        Next i
    End If

    If CStr(fields(6)) <> "" Then
        session.RestoreOutputIdentity HexDecode(CStr(fields(6))), HexDecode(CStr(fields(7))), _
                                      ParseInvariantNumber(CStr(fields(8))), HexDecode(CStr(fields(5))), _
                                      HexDecode(CStr(fields(9))), HexDecode(CStr(fields(10)))
    End If
    If CStr(fields(11)) <> "" Or CStr(fields(12)) <> "" Then
        session.RestoreEventIdentities HexDecode(CStr(fields(11))), HexDecode(CStr(fields(12)))
    End If
    session.RestoreProgress TokenBool(CStr(fields(13))), TokenBool(CStr(fields(14))), _
                            TokenBool(CStr(fields(15))), TokenBool(CStr(fields(16))), _
                            TokenBool(CStr(fields(17))), TokenBool(CStr(fields(18))), _
                            HexDecode(CStr(fields(19))), HexDecode(CStr(fields(20))), _
                            HexDecode(CStr(fields(21))), TokenBool(CStr(fields(22)))

    Set DeserializeProductionSession = session
    report = ""
    Exit Function

FailDeserialize:
    report = "DeserializeProductionSession failed: " & Err.Description
End Function

Public Function SaveProductionSessionToWorkbook(ByVal wb As Workbook, _
                                                ByVal session As cProductionRunSession, _
                                                Optional ByRef report As String = "") As Boolean
    Dim serializedState As String
    Dim chunkValue As String
    Dim refersToValue As String
    Dim chunkCount As Long
    Dim chunkIndex As Long

    On Error GoTo FailSave
    If wb Is Nothing Then
        report = "Operator workbook is required."
        Exit Function
    End If
    serializedState = SerializeProductionSession(session)
    DeleteProductionSessionNames wb
    chunkCount = (Len(serializedState) + SESSION_WORKBOOK_CHUNK_SIZE - 1) \ SESSION_WORKBOOK_CHUNK_SIZE
    If chunkCount < 1 Then chunkCount = 1

    For chunkIndex = 1 To chunkCount
        chunkValue = Mid$(serializedState, _
                          ((chunkIndex - 1) * SESSION_WORKBOOK_CHUNK_SIZE) + 1, _
                          SESSION_WORKBOOK_CHUNK_SIZE)
        refersToValue = "=""" & Replace$(chunkValue, """", """""") & """"
        wb.Names.Add Name:=SessionChunkName(chunkIndex), RefersTo:=refersToValue, Visible:=False
    Next chunkIndex
    wb.Names.Add Name:=SESSION_WORKBOOK_NAME, RefersTo:="=" & CStr(chunkCount), Visible:=False
    SaveProductionSessionToWorkbook = True
    report = "Production session state saved in " & CStr(chunkCount) & " metadata chunks."
    Exit Function

FailSave:
    On Error Resume Next
    If Not wb Is Nothing Then DeleteProductionSessionNames wb
    On Error GoTo 0
    report = "SaveProductionSessionToWorkbook failed: " & Err.Description
End Function

Public Function LoadProductionSessionFromWorkbook(ByVal wb As Workbook, _
                                                  Optional ByRef report As String = "") As cProductionRunSession
    Dim stateName As Name
    Dim chunkName As Name
    Dim serializedState As String
    Dim chunkCount As Long
    Dim chunkIndex As Long
    Dim chunkValue As String

    On Error GoTo FailLoad
    If wb Is Nothing Then
        report = "Operator workbook is required."
        Exit Function
    End If
    On Error Resume Next
    Set stateName = wb.Names(SESSION_WORKBOOK_NAME)
    On Error GoTo FailLoad
    If stateName Is Nothing Then
        report = "No persisted Production session exists in the operator workbook."
        Exit Function
    End If
    chunkCount = CLng(Replace$(CStr(stateName.RefersTo), "=", ""))
    If chunkCount < 1 Then _
        Err.Raise vbObjectError + 7652, "modProductionCompletionService.LoadProductionSessionFromWorkbook", _
                  "Persisted Production session chunk manifest is invalid."
    For chunkIndex = 1 To chunkCount
        Set chunkName = Nothing
        On Error Resume Next
        Set chunkName = wb.Names(SessionChunkName(chunkIndex))
        On Error GoTo FailLoad
        If chunkName Is Nothing Then _
            Err.Raise vbObjectError + 7653, "modProductionCompletionService.LoadProductionSessionFromWorkbook", _
                      "Persisted Production session chunk " & CStr(chunkIndex) & " is missing."
        chunkValue = CStr(chunkName.RefersTo)
        If Left$(chunkValue, 2) = "=""" And Right$(chunkValue, 1) = """" Then
            chunkValue = Mid$(chunkValue, 3, Len(chunkValue) - 3)
            chunkValue = Replace$(chunkValue, """""", """")
        End If
        serializedState = serializedState & chunkValue
    Next chunkIndex
    Set LoadProductionSessionFromWorkbook = DeserializeProductionSession(serializedState, report)
    Exit Function

FailLoad:
    report = "LoadProductionSessionFromWorkbook failed: " & Err.Description
End Function

Public Function ClearProductionSessionFromWorkbook(ByVal wb As Workbook, _
                                                   Optional ByRef report As String = "") As Boolean
    On Error GoTo FailClear
    If wb Is Nothing Then
        report = "Operator workbook is required."
        Exit Function
    End If
    DeleteProductionSessionNames wb
    ClearProductionSessionFromWorkbook = True
    report = "Production session state cleared."
    Exit Function

FailClear:
    report = "ClearProductionSessionFromWorkbook failed: " & Err.Description
End Function

Private Sub DeleteProductionSessionNames(ByVal wb As Workbook)
    Dim i As Long
    Dim localName As String

    If wb Is Nothing Then Exit Sub
    For i = wb.Names.Count To 1 Step -1
        localName = CStr(wb.Names(i).Name)
        If InStrRev(localName, "!") > 0 Then localName = Mid$(localName, InStrRev(localName, "!") + 1)
        localName = Replace$(localName, "'", "")
        If StrComp(localName, SESSION_WORKBOOK_NAME, vbTextCompare) = 0 _
           Or LCase$(Left$(localName, Len(SESSION_WORKBOOK_NAME) + 1)) = _
              LCase$(SESSION_WORKBOOK_NAME & "_") Then
            wb.Names(i).Delete
        End If
    Next i
End Sub

Private Function SessionChunkName(ByVal chunkIndex As Long) As String
    SessionChunkName = SESSION_WORKBOOK_NAME & "_" & Format$(chunkIndex, "0000")
End Function

Public Function ProductionSessionContractProbe(ByVal contractName As String) As String
    On Error GoTo ProbeFailed
    Select Case UCase$(Trim$(contractName))
        Case "IDENTITY"
            ProductionSessionContractProbe = ProbeIdentity()
        Case "INVALID_INPUT_KEYS"
            ProductionSessionContractProbe = ProbeInvalidInputKeys()
        Case "EVENT_IDENTITIES"
            ProductionSessionContractProbe = ProbeEventIdentities()
        Case "READY_SEQUENCE"
            ProductionSessionContractProbe = ProbeReadySequence()
        Case "COMPENSATION"
            ProductionSessionContractProbe = ProbeCompensation()
        Case "RESTART"
            ProductionSessionContractProbe = ProbeRestart()
        Case "RESULT_ENVELOPE"
            ProductionSessionContractProbe = ProbeResultEnvelope()
        Case Else
            ProductionSessionContractProbe = "FAIL|UnknownContract=" & contractName
    End Select
    Exit Function

ProbeFailed:
    ProductionSessionContractProbe = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function ProbeIdentity() As String
    Dim session As cProductionRunSession
    Dim firstOutputKey As String
    Dim secondOutputKey As String

    Set session = CreateProductionSession("SESSION-IDENTITY")
    session.AddInputAllocation "SYS-INPUT-A", 2, "SKU-A", "A1"
    session.AddInputAllocation "SYS-INPUT-B", 3, "SKU-B", "A1"
    firstOutputKey = session.EnsureOutputIdentity("SKU-OUTPUT", 5, "A1")
    secondOutputKey = session.EnsureOutputIdentity("SKU-OUTPUT", 5, "A1")
    ProbeIdentity = "OK|InputCount=" & CStr(session.InputCount) & _
                    "|InputKeysPreserved=" & BoolText( _
                        CStr(session.InputAllocation(1)("System_Key")) = "SYS-INPUT-A" And _
                        CStr(session.InputAllocation(2)("System_Key")) = "SYS-INPUT-B") & _
                    "|OutputKeyNonblank=" & BoolText(firstOutputKey <> "") & _
                    "|OutputKeyStable=" & BoolText(firstOutputKey = secondOutputKey)
End Function

Private Function ProbeInvalidInputKeys() As String
    Dim blankRejected As Boolean
    Dim duplicateRejected As Boolean
    Dim session As cProductionRunSession

    Set session = CreateProductionSession("SESSION-INVALID")
    On Error Resume Next
    session.AddInputAllocation "", 1
    blankRejected = (Err.Number <> 0)
    Err.Clear
    session.AddInputAllocation "SYS-DUP", 1
    session.AddInputAllocation "SYS-DUP", 1
    duplicateRejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    ProbeInvalidInputKeys = "OK|BlankRejected=" & BoolText(blankRejected) & _
                            "|DuplicateRejected=" & BoolText(duplicateRejected)
End Function

Private Function ProbeEventIdentities() As String
    Dim session As cProductionRunSession
    Dim consumeId As String
    Dim completeId As String

    Set session = CreatePreparedProbeSession("SESSION-EVENTS")
    session.EnsureEventIdentities
    consumeId = session.ConsumeEventId
    completeId = session.CompleteEventId
    session.EnsureEventIdentities
    ProbeEventIdentities = "OK|ConsumeNonblank=" & BoolText(consumeId <> "") & _
                           "|CompleteNonblank=" & BoolText(completeId <> "") & _
                           "|Distinct=" & BoolText(StrComp(consumeId, completeId, vbTextCompare) <> 0) & _
                           "|Stable=" & BoolText(consumeId = session.ConsumeEventId And completeId = session.CompleteEventId)
End Function

Private Function ProbeReadySequence() As String
    Dim session As cProductionRunSession
    Dim beforeProcessor As Boolean
    Dim beforeRefresh As Boolean

    Set session = CreatePreparedProbeSession("SESSION-READY")
    session.EnsureEventIdentities
    session.MarkConsumeQueued
    session.MarkCompleteQueued
    beforeProcessor = session.ReadyForNextBatch
    session.RecordProcessorResult True, True, True
    beforeRefresh = session.ReadyForNextBatch
    session.RecordRefreshResult True
    ProbeReadySequence = "OK|BeforeProcessor=" & BoolText(beforeProcessor) & _
                         "|BeforeRefresh=" & BoolText(beforeRefresh) & _
                         "|AfterRefresh=" & BoolText(session.ReadyForNextBatch)
End Function

Private Function ProbeCompensation() As String
    Dim beforeConsume As cProductionRunSession
    Dim afterConsume As cProductionRunSession
    Dim afterComplete As cProductionRunSession

    Set beforeConsume = CreatePreparedProbeSession("SESSION-COMP-1")
    beforeConsume.RecordFailure "PRE_QUEUE", "pre-queue failure"

    Set afterConsume = CreatePreparedProbeSession("SESSION-COMP-2")
    afterConsume.RecordProcessorResult True, False, False, "COMPLETE_FAILED", "output failed"

    Set afterComplete = CreatePreparedProbeSession("SESSION-COMP-3")
    afterComplete.RecordProcessorResult True, True, True
    afterComplete.RecordFailure "REFRESH_FAILED", "refresh failed"

    ProbeCompensation = "OK|BeforeConsume=" & BoolText(beforeConsume.CompensationRequired) & _
                        "|AfterConsume=" & BoolText(afterConsume.CompensationRequired) & _
                        "|AfterComplete=" & BoolText(afterComplete.CompensationRequired)
End Function

Private Function ProbeRestart() As String
    Dim beforeRestart As cProductionRunSession
    Dim afterRestart As cProductionRunSession
    Dim serializedState As String
    Dim report As String

    Set beforeRestart = CreatePreparedProbeSession("SESSION-RESTART")
    beforeRestart.EnsureEventIdentities
    beforeRestart.MarkConsumeQueued
    beforeRestart.MarkCompleteQueued
    beforeRestart.RecordProcessorResult True, True, True
    serializedState = SerializeProductionSession(beforeRestart)
    Set afterRestart = DeserializeProductionSession(serializedState, report)
    If afterRestart Is Nothing Then
        ProbeRestart = "FAIL|" & report
        Exit Function
    End If

    ProbeRestart = "OK|StatePreserved=" & BoolText( _
                        beforeRestart.Status = afterRestart.Status And _
                        beforeRestart.InputCount = afterRestart.InputCount And _
                        beforeRestart.ProcessorVerified = afterRestart.ProcessorVerified) & _
                   "|OutputKeyPreserved=" & BoolText(beforeRestart.OutputSystemKey = afterRestart.OutputSystemKey) & _
                   "|EventIdsPreserved=" & BoolText( _
                        beforeRestart.ConsumeEventId = afterRestart.ConsumeEventId And _
                        beforeRestart.CompleteEventId = afterRestart.CompleteEventId)
End Function

Private Function ProbeResultEnvelope() As String
    Dim session As cProductionRunSession
    Dim result As cProductionCompletionResult

    Set session = CreatePreparedProbeSession("SESSION-RESULT")
    session.EnsureEventIdentities
    session.MarkConsumeQueued
    session.MarkCompleteQueued
    session.RecordProcessorResult True, True, True
    session.RecordRefreshResult True
    Set result = CreateProductionCompletionResult(session)

    ProbeResultEnvelope = "OK|Status=" & result.Status & _
                          "|ProcessorVerified=" & BoolText(result.ProcessorVerified) & _
                          "|RefreshVerified=" & BoolText(result.RefreshVerified) & _
                          "|CompensationRequired=" & BoolText(result.CompensationRequired)
End Function

Private Function CreatePreparedProbeSession(ByVal sessionId As String) As cProductionRunSession
    Dim session As cProductionRunSession

    Set session = CreateProductionSession(sessionId, "DESIGN-1", "1", 5, "A1")
    session.AddInputAllocation "SYS-INPUT-" & sessionId, 5, "SKU-INPUT", "A1"
    session.EnsureOutputIdentity "SKU-OUTPUT", 5, "A1", "GOOD", "{}", "SYS-OUTPUT-" & sessionId
    Set CreatePreparedProbeSession = session
End Function

Private Sub RequireSessionPrepared(ByVal session As cProductionRunSession, ByVal memberName As String)
    If session Is Nothing Then _
        Err.Raise vbObjectError + 7644, "modProductionCompletionService." & memberName, _
                  "Production session is required."
    If session.InputCount = 0 Then _
        Err.Raise vbObjectError + 7645, "modProductionCompletionService." & memberName, _
                  "Production session has no input allocations."
    If session.OutputSystemKey = "" Or session.OutputSku = "" Or session.OutputQty <= 0 Then _
        Err.Raise vbObjectError + 7646, "modProductionCompletionService." & memberName, _
                  "Production session output identity is incomplete."
End Sub

Private Function JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim i As Long
    For i = 1 To values.Count
        If i > 1 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(values(i))
    Next i
End Function

Private Function HexEncode(ByVal valueIn As String) As String
    Dim i As Long
    Dim codeUnit As Long
    For i = 1 To Len(valueIn)
        codeUnit = AscW(Mid$(valueIn, i, 1))
        If codeUnit < 0 Then codeUnit = codeUnit + 65536
        HexEncode = HexEncode & Right$("0000" & Hex$(codeUnit), 4)
    Next i
End Function

Private Function HexDecode(ByVal valueIn As String) As String
    Dim i As Long
    Dim codeUnit As Long
    If Len(valueIn) Mod 4 <> 0 Then _
        Err.Raise vbObjectError + 7647, "modProductionCompletionService.HexDecode", _
                  "Malformed encoded Production session text."
    For i = 1 To Len(valueIn) Step 4
        codeUnit = CLng("&H" & Mid$(valueIn, i, 4))
        If codeUnit > 32767 Then codeUnit = codeUnit - 65536
        HexDecode = HexDecode & ChrW$(codeUnit)
    Next i
End Function

Private Function InvariantNumber(ByVal valueIn As Double) As String
    InvariantNumber = CStr(valueIn)
    InvariantNumber = Replace$(InvariantNumber, Application.International(xlDecimalSeparator), ".")
End Function

Private Function ParseInvariantNumber(ByVal valueIn As String) As Double
    valueIn = Replace$(valueIn, ".", Application.International(xlDecimalSeparator))
    If Not IsNumeric(valueIn) Then _
        Err.Raise vbObjectError + 7648, "modProductionCompletionService.ParseInvariantNumber", _
                  "Malformed numeric Production session value."
    ParseInvariantNumber = CDbl(valueIn)
End Function

Private Function BoolToken(ByVal valueIn As Boolean) As String
    If valueIn Then
        BoolToken = "1"
    Else
        BoolToken = "0"
    End If
End Function

Private Function TokenBool(ByVal valueIn As String) As Boolean
    If valueIn <> "0" And valueIn <> "1" Then _
        Err.Raise vbObjectError + 7649, "modProductionCompletionService.TokenBool", _
                  "Malformed Boolean Production session value."
    TokenBool = (valueIn = "1")
End Function

Private Function BoolText(ByVal valueIn As Boolean) As String
    BoolText = IIf(valueIn, "TRUE", "FALSE")
End Function

Private Function FindWorkbookTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set FindWorkbookTable = lo
            Exit Function
        End If
    Next ws
End Function

Private Function TableColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Function
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(headerName), vbTextCompare) = 0 Then
            TableColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function TableText(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal headerName As String) As String
    Dim columnIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    columnIndex = TableColumnIndex(lo, headerName)
    If columnIndex = 0 Then Exit Function
    On Error Resume Next
    TableText = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, columnIndex).Value))
    On Error GoTo 0
End Function

Private Function TableNumber(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal headerName As String) As Double
    Dim rawValue As Variant
    Dim columnIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    columnIndex = TableColumnIndex(lo, headerName)
    If columnIndex = 0 Then Exit Function
    rawValue = lo.DataBodyRange.Cells(rowIndex, columnIndex).Value
    If IsNumeric(rawValue) Then TableNumber = CDbl(rawValue)
End Function

Private Sub SetTableText(ByVal lo As ListObject, ByVal rowIndex As Long, _
                         ByVal headerName As String, ByVal valueOut As String)
    Dim columnIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Sub
    columnIndex = TableColumnIndex(lo, headerName)
    If columnIndex = 0 Then _
        Err.Raise vbObjectError + 7654, "modProductionCompletionService.SetTableText", _
                  "Production staging table '" & lo.Name & "' is missing '" & headerName & "'."
    lo.DataBodyRange.Cells(rowIndex, columnIndex).Value = valueOut
End Sub

Private Function InventoryValueBySystemKey(ByVal wb As Workbook, _
                                           ByVal systemKey As String, _
                                           ByVal headerName As String) As String
    Dim lo As ListObject
    Dim cSystemKey As Long
    Dim cValue As Long
    Dim rowIndex As Long

    Set lo = FindWorkbookTable(wb, "invSys")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cSystemKey = TableColumnIndex(lo, "System_Key")
    cValue = TableColumnIndex(lo, headerName)
    If cSystemKey = 0 Or cValue = 0 Then Exit Function
    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, cSystemKey).Value)), _
                   Trim$(systemKey), vbTextCompare) = 0 Then
            InventoryValueBySystemKey = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, cValue).Value))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function RecreateSessionWithLocation(ByVal sourceSession As cProductionRunSession, _
                                             ByVal locationValue As String) As cProductionRunSession
    Dim replacement As cProductionRunSession
    Dim allocation As Object
    Dim i As Long

    Set replacement = CreateProductionSession(sourceSession.SessionId, sourceSession.DesignId, _
                                              sourceSession.DesignVersion, sourceSession.QtyPlanned, _
                                              locationValue)
    For i = 1 To sourceSession.InputCount
        Set allocation = sourceSession.InputAllocation(i)
        replacement.AddInputAllocation CStr(allocation("System_Key")), CDbl(allocation("Qty")), _
                                       CStr(allocation("SKU")), CStr(allocation("Location"))
    Next i
    Set RecreateSessionWithLocation = replacement
End Function
