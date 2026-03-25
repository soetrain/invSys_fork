Attribute VB_Name = "modInventoryApply"
Option Explicit

Public Const APPLY_STATUS_APPLIED As String = "APPLIED"
Public Const APPLY_STATUS_SKIP_DUP As String = "SKIP_DUP"

Public Const EVENT_TYPE_RECEIVE As String = "RECEIVE"
Public Const EVENT_TYPE_SHIP As String = "SHIP"
Public Const EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Public Const EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"

Public Function ApplyEvent(ByVal evt As Object, _
                           Optional ByVal inventoryWb As Workbook = Nothing, _
                           Optional ByVal runId As String = "", _
                           Optional ByRef statusOut As String = "", _
                           Optional ByRef errorCode As String = "", _
                           Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo FailApply

    Dim wb As Workbook
    Dim loLog As ListObject
    Dim loApplied As ListObject
    Dim schemaReport As String
    Dim eventId As String
    Dim eventType As String
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim sourceInbox As String
    Dim undoOfEventId As String
    Dim occurredAt As Date
    Dim appliedAt As Date
    Dim appliedSeq As Long
    Dim linesToApply As Collection
    Dim lineItem As Variant
    Dim r As ListRow

    Set wb = ResolveInventoryWorkbook(GetEventString(evt, "WarehouseId"), inventoryWb)
    If wb Is Nothing Then
        errorCode = "INVENTORY_WORKBOOK_NOT_FOUND"
        errorMessage = "Inventory workbook not found."
        Exit Function
    End If

    If Not modInventorySchema.EnsureInventorySchema(wb, schemaReport) Then
        errorCode = "INVENTORY_SCHEMA_INVALID"
        If schemaReport <> "" Then
            errorMessage = schemaReport
        Else
            errorMessage = "Unable to validate inventory schema."
        End If
        Exit Function
    End If

    Set loLog = FindListObjectByNameApply(wb, "tblInventoryLog")
    Set loApplied = FindListObjectByNameApply(wb, "tblAppliedEvents")
    If loLog Is Nothing Or loApplied Is Nothing Then
        errorCode = "INVENTORY_TABLE_MISSING"
        errorMessage = "Required inventory tables not found."
        Exit Function
    End If

    SetSheetProtectionApply loLog.Parent, False
    SetSheetProtectionApply loApplied.Parent, False

    eventId = GetEventString(evt, "EventID")
    eventType = NormalizeEventType(GetEventString(evt, "EventType"))
    warehouseId = GetEventString(evt, "WarehouseId")
    stationId = GetEventString(evt, "StationId")
    userId = GetEventString(evt, "UserId")
    sourceInbox = GetEventString(evt, "SourceInbox")
    undoOfEventId = GetEventString(evt, "UndoOfEventId")

    If eventId = "" Then
        errorCode = "INVALID_EVENT"
        errorMessage = "EventID is required."
        GoTo CleanExit
    End If
    If Not TryGetEventDate(evt, "CreatedAtUTC", occurredAt) Then
        errorCode = "INVALID_EVENT"
        errorMessage = "CreatedAtUTC is required and must be a valid date."
        GoTo CleanExit
    End If
    If warehouseId = "" Or stationId = "" Or userId = "" Then
        errorCode = "INVALID_EVENT"
        errorMessage = "WarehouseId, StationId, and UserId are required."
        GoTo CleanExit
    End If
    If eventType = "" Then
        errorCode = "INVALID_EVENT_TYPE"
        errorMessage = "EventType is required."
        GoTo CleanExit
    End If

    If AppliedEventExists(loApplied, eventId) Then
        statusOut = APPLY_STATUS_SKIP_DUP
        ApplyEvent = True
        GoTo CleanExit
    End If

    Set linesToApply = BuildApplyLines(evt, wb, eventType, errorCode, errorMessage)
    If linesToApply Is Nothing Then GoTo CleanExit
    If linesToApply.Count = 0 Then
        errorCode = "INVALID_PAYLOAD"
        errorMessage = "Event did not produce any inventory lines."
        GoTo CleanExit
    End If

    appliedAt = Now
    appliedSeq = GetNextAppliedSeq(wb)
    If runId = "" Then runId = "RUN-" & Format$(appliedAt, "yyyymmddhhnnss")

    For Each lineItem In linesToApply
        Set r = loLog.ListRows.Add
        SetTableRowValue loLog, r.Index, "EventID", eventId
        SetTableRowValue loLog, r.Index, "UndoOfEventId", undoOfEventId
        SetTableRowValue loLog, r.Index, "AppliedSeq", appliedSeq
        SetTableRowValue loLog, r.Index, "EventType", eventType
        SetTableRowValue loLog, r.Index, "OccurredAtUTC", occurredAt
        SetTableRowValue loLog, r.Index, "AppliedAtUTC", appliedAt
        SetTableRowValue loLog, r.Index, "WarehouseId", warehouseId
        SetTableRowValue loLog, r.Index, "StationId", stationId
        SetTableRowValue loLog, r.Index, "UserId", userId
        SetTableRowValue loLog, r.Index, "SKU", CStr(lineItem("SKU"))
        SetTableRowValue loLog, r.Index, "QtyDelta", CDbl(lineItem("QtyDelta"))
        SetTableRowValue loLog, r.Index, "Location", CStr(lineItem("Location"))
        SetTableRowValue loLog, r.Index, "Note", CStr(lineItem("Note"))
    Next lineItem

    Set r = loApplied.ListRows.Add
    SetTableRowValue loApplied, r.Index, "EventID", eventId
    SetTableRowValue loApplied, r.Index, "UndoOfEventId", undoOfEventId
    SetTableRowValue loApplied, r.Index, "AppliedSeq", appliedSeq
    SetTableRowValue loApplied, r.Index, "AppliedAtUTC", appliedAt
    SetTableRowValue loApplied, r.Index, "RunId", runId
    SetTableRowValue loApplied, r.Index, "SourceInbox", sourceInbox
    SetTableRowValue loApplied, r.Index, "Status", APPLY_STATUS_APPLIED

    RebuildInventoryProjections wb
    RefreshLedgerStatus wb, warehouseId, appliedSeq, eventId, appliedAt
    SaveInventoryWorkbookIfWritable wb
    statusOut = APPLY_STATUS_APPLIED
    ApplyEvent = True

CleanExit:
    On Error Resume Next
    If Not loLog Is Nothing Then SetSheetProtectionApply loLog.Parent, True
    If Not loApplied Is Nothing Then SetSheetProtectionApply loApplied.Parent, True
    On Error GoTo 0
    Exit Function

FailApply:
    Dim failNumber As Long
    Dim failDescription As String

    failNumber = Err.Number
    failDescription = Err.Description
    On Error Resume Next
    If Not loLog Is Nothing Then SetSheetProtectionApply loLog.Parent, True
    If Not loApplied Is Nothing Then SetSheetProtectionApply loApplied.Parent, True
    On Error GoTo 0
    If errorCode = "" Then errorCode = "APPLY_EXCEPTION"
    If errorMessage = "" Then errorMessage = CStr(failNumber) & ": " & failDescription
End Function

Private Sub SaveInventoryWorkbookIfWritable(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If Trim$(wb.Path) = "" Then Exit Sub
    wb.Save
End Sub

Public Function ApplyReceiveEvent(ByVal evt As Object, _
                                  Optional ByVal inventoryWb As Workbook = Nothing, _
                                  Optional ByVal runId As String = "", _
                                  Optional ByRef statusOut As String = "", _
                                  Optional ByRef errorCode As String = "", _
                                  Optional ByRef errorMessage As String = "") As Boolean
    ApplyReceiveEvent = ApplyEvent(evt, inventoryWb, runId, statusOut, errorCode, errorMessage)
End Function

Public Function ResolveInventoryWorkbook(Optional ByVal warehouseId As String = "", _
                                         Optional ByVal inventoryWb As Workbook = Nothing) As Workbook
    Dim wb As Workbook
    Dim targetPath As String

    If Not inventoryWb Is Nothing Then
        Set ResolveInventoryWorkbook = inventoryWb
        Exit Function
    End If

    targetPath = BuildCanonicalInventoryPath(warehouseId)
    For Each wb In Application.Workbooks
        If targetPath <> "" Then
            If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
                Set ResolveInventoryWorkbook = wb
                Exit Function
            End If
        End If

        If IsInventoryWorkbookName(wb.Name) Then
            If warehouseId = "" Or InStr(1, wb.Name, warehouseId, vbTextCompare) > 0 Then
                Set ResolveInventoryWorkbook = wb
                Exit Function
            End If
        End If
    Next wb

    Set ResolveInventoryWorkbook = OpenOrCreateCanonicalInventoryWorkbook(warehouseId)
End Function

Private Function BuildApplyLines(ByVal evt As Object, _
                                 ByVal wb As Workbook, _
                                 ByVal eventType As String, _
                                 ByRef errorCode As String, _
                                 ByRef errorMessage As String) As Collection
    Select Case eventType
        Case EVENT_TYPE_RECEIVE
            Set BuildApplyLines = BuildReceiveLines(evt, wb, errorCode, errorMessage)
        Case EVENT_TYPE_SHIP, EVENT_TYPE_PROD_CONSUME, EVENT_TYPE_PROD_COMPLETE
            Set BuildApplyLines = BuildPayloadLines(evt, wb, eventType, errorCode, errorMessage)
        Case Else
            errorCode = "INVALID_EVENT_TYPE"
            errorMessage = "Unsupported EventType '" & eventType & "'."
    End Select
End Function

Private Function BuildReceiveLines(ByVal evt As Object, _
                                   ByVal wb As Workbook, _
                                   ByRef errorCode As String, _
                                   ByRef errorMessage As String) As Collection
    Dim sku As String
    Dim qty As Double
    Dim lineItem As Object

    sku = GetEventString(evt, "SKU")
    If sku = "" Then
        errorCode = "INVALID_SKU"
        errorMessage = "SKU is required."
        Exit Function
    End If
    If Not TryGetEventDouble(evt, "Qty", qty) Then
        errorCode = "INVALID_QTY"
        errorMessage = "Qty is required and must be numeric."
        Exit Function
    End If
    If qty <= 0 Then
        errorCode = "INVALID_QTY"
        errorMessage = "Qty must be greater than zero."
        Exit Function
    End If
    If Not ValidateSkuExists(wb, sku) Then
        errorCode = "INVALID_SKU"
        errorMessage = "SKU not found in inventory catalog."
        Exit Function
    End If

    Set BuildReceiveLines = New Collection
    Set lineItem = CreateObject("Scripting.Dictionary")
    lineItem.CompareMode = vbTextCompare
    lineItem("SKU") = sku
    lineItem("QtyDelta") = qty
    lineItem("Location") = GetEventString(evt, "Location")
    lineItem("Note") = GetEventString(evt, "Note")
    BuildReceiveLines.Add lineItem
End Function

Private Function BuildPayloadLines(ByVal evt As Object, _
                                   ByVal wb As Workbook, _
                                   ByVal eventType As String, _
                                   ByRef errorCode As String, _
                                   ByRef errorMessage As String) As Collection
    Dim payloadJson As String
    Dim parsedItems As Collection
    Dim rawItem As Variant
    Dim lineItem As Object
    Dim sku As String
    Dim qty As Double
    Dim qtyDelta As Double
    Dim locationVal As String
    Dim noteVal As String
    Dim ioType As String

    payloadJson = GetEventString(evt, "PayloadJson")
    If payloadJson = "" Then
        errorCode = "INVALID_PAYLOAD"
        errorMessage = "PayloadJson is required for event type '" & eventType & "'."
        Exit Function
    End If

    Set parsedItems = ParsePayloadJsonArray(payloadJson, errorMessage)
    If parsedItems Is Nothing Then
        errorCode = "INVALID_PAYLOAD"
        If errorMessage = "" Then errorMessage = "PayloadJson could not be parsed."
        Exit Function
    End If
    If parsedItems.Count = 0 Then
        errorCode = "INVALID_PAYLOAD"
        errorMessage = "PayloadJson did not contain any line items."
        Exit Function
    End If

    Set BuildPayloadLines = New Collection
    For Each rawItem In parsedItems
        sku = SafeTrimApply(GetDictionaryValue(rawItem, "SKU"))
        If sku = "" Then
            errorCode = "INVALID_SKU"
            errorMessage = "Every payload line item requires SKU."
            Set BuildPayloadLines = Nothing
            Exit Function
        End If
        If Not TryGetDictionaryDouble(rawItem, "Qty", qty) Then
            errorCode = "INVALID_QTY"
            errorMessage = "Every payload line item requires numeric Qty."
            Set BuildPayloadLines = Nothing
            Exit Function
        End If
        If qty <= 0 Then
            errorCode = "INVALID_QTY"
            errorMessage = "Payload Qty must be greater than zero."
            Set BuildPayloadLines = Nothing
            Exit Function
        End If
        If Not ValidateSkuExists(wb, sku) Then
            errorCode = "INVALID_SKU"
            errorMessage = "SKU '" & sku & "' not found in inventory catalog."
            Set BuildPayloadLines = Nothing
            Exit Function
        End If

        ioType = UCase$(SafeTrimApply(GetDictionaryValue(rawItem, "IoType")))
        qtyDelta = ResolvePayloadQtyDelta(eventType, ioType, qty, errorCode, errorMessage)
        If errorCode <> "" Then
            Set BuildPayloadLines = Nothing
            Exit Function
        End If

        locationVal = SafeTrimApply(GetDictionaryValue(rawItem, "Location"))
        If locationVal = "" Then locationVal = GetEventString(evt, "Location")

        noteVal = ComposeLineNote(eventType, rawItem, GetEventString(evt, "Note"))

        Set lineItem = CreateObject("Scripting.Dictionary")
        lineItem.CompareMode = vbTextCompare
        lineItem("SKU") = sku
        lineItem("QtyDelta") = qtyDelta
        lineItem("Location") = locationVal
        lineItem("Note") = noteVal
        BuildPayloadLines.Add lineItem
    Next rawItem
End Function

Private Function ResolvePayloadQtyDelta(ByVal eventType As String, _
                                        ByVal ioType As String, _
                                        ByVal qty As Double, _
                                        ByRef errorCode As String, _
                                        ByRef errorMessage As String) As Double
    Select Case eventType
        Case EVENT_TYPE_SHIP
            ResolvePayloadQtyDelta = -qty
        Case EVENT_TYPE_PROD_CONSUME
            Select Case ioType
                Case "USED"
                    ResolvePayloadQtyDelta = -qty
                Case "MADE"
                    ResolvePayloadQtyDelta = qty
                Case Else
                    errorCode = "INVALID_PAYLOAD"
                    errorMessage = "PROD_CONSUME payload line items require IoType USED or MADE."
            End Select
        Case EVENT_TYPE_PROD_COMPLETE
            If ioType <> "" And ioType <> "MADE" And ioType <> "COMPLETE" Then
                errorCode = "INVALID_PAYLOAD"
                errorMessage = "PROD_COMPLETE payload line items may only use IoType MADE or COMPLETE."
            Else
                ResolvePayloadQtyDelta = qty
            End If
        Case Else
            errorCode = "INVALID_EVENT_TYPE"
            errorMessage = "Unsupported EventType '" & eventType & "'."
    End Select
End Function

Private Function ComposeLineNote(ByVal eventType As String, ByVal rawItem As Object, ByVal eventNote As String) As String
    Dim itemNote As String
    Dim ioType As String
    Dim rowVal As String
    Dim detail As String

    itemNote = SafeTrimApply(GetDictionaryValue(rawItem, "Note"))
    ioType = SafeTrimApply(GetDictionaryValue(rawItem, "IoType"))
    rowVal = SafeTrimApply(GetDictionaryValue(rawItem, "Row"))

    If rowVal <> "" Then detail = "ROW=" & rowVal
    If ioType <> "" Then
        If detail <> "" Then detail = detail & "; "
        detail = detail & "IO=" & UCase$(ioType)
    End If

    ComposeLineNote = itemNote
    If ComposeLineNote = "" Then ComposeLineNote = eventNote
    If detail <> "" Then
        If ComposeLineNote <> "" Then
            ComposeLineNote = ComposeLineNote & " | " & detail
        Else
            ComposeLineNote = detail
        End If
    End If
    If ComposeLineNote = "" Then ComposeLineNote = eventType
End Function

Private Function ParsePayloadJsonArray(ByVal payloadJson As String, ByRef errorMessage As String) As Collection
    Dim idx As Long
    idx = 1
    SkipJsonWhitespace payloadJson, idx
    If Not ConsumeJsonChar(payloadJson, idx, "[") Then
        errorMessage = "PayloadJson must start with '['."
        Exit Function
    End If

    Set ParsePayloadJsonArray = New Collection
    SkipJsonWhitespace payloadJson, idx
    If ConsumeJsonChar(payloadJson, idx, "]") Then Exit Function

    Do
        Dim item As Object
        Set item = ParsePayloadJsonObject(payloadJson, idx, errorMessage)
        If item Is Nothing Then
            Set ParsePayloadJsonArray = Nothing
            Exit Function
        End If
        ParsePayloadJsonArray.Add item
        SkipJsonWhitespace payloadJson, idx
        If ConsumeJsonChar(payloadJson, idx, "]") Then Exit Do
        If Not ConsumeJsonChar(payloadJson, idx, ",") Then
            errorMessage = "PayloadJson array is missing a comma separator."
            Set ParsePayloadJsonArray = Nothing
            Exit Function
        End If
    Loop

    SkipJsonWhitespace payloadJson, idx
    If idx <= Len(payloadJson) Then
        errorMessage = "PayloadJson contains unexpected trailing characters."
        Set ParsePayloadJsonArray = Nothing
    End If
End Function

Private Function ParsePayloadJsonObject(ByVal payloadJson As String, ByRef idx As Long, ByRef errorMessage As String) As Object
    Dim item As Object

    SkipJsonWhitespace payloadJson, idx
    If Not ConsumeJsonChar(payloadJson, idx, "{") Then
        errorMessage = "PayloadJson object must start with '{'."
        Exit Function
    End If

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare

    SkipJsonWhitespace payloadJson, idx
    If ConsumeJsonChar(payloadJson, idx, "}") Then
        Set ParsePayloadJsonObject = item
        Exit Function
    End If

    Do
        Dim key As String
        key = ParseJsonString(payloadJson, idx, errorMessage)
        If errorMessage <> "" Then
            Exit Function
        End If

        SkipJsonWhitespace payloadJson, idx
        If Not ConsumeJsonChar(payloadJson, idx, ":") Then
            errorMessage = "PayloadJson object is missing ':' after key '" & key & "'."
            Exit Function
        End If

        item(key) = ParseJsonValue(payloadJson, idx, errorMessage)
        If errorMessage <> "" Then
            Exit Function
        End If

        SkipJsonWhitespace payloadJson, idx
        If ConsumeJsonChar(payloadJson, idx, "}") Then Exit Do
        If Not ConsumeJsonChar(payloadJson, idx, ",") Then
            errorMessage = "PayloadJson object is missing a comma separator."
            Exit Function
        End If
    Loop

    Set ParsePayloadJsonObject = item
End Function

Private Function ParseJsonValue(ByVal payloadJson As String, ByRef idx As Long, ByRef errorMessage As String) As Variant
    Dim ch As String

    SkipJsonWhitespace payloadJson, idx
    If idx > Len(payloadJson) Then
        errorMessage = "Unexpected end of PayloadJson."
        Exit Function
    End If

    ch = Mid$(payloadJson, idx, 1)
    Select Case ch
        Case """"
            ParseJsonValue = ParseJsonString(payloadJson, idx, errorMessage)
        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            ParseJsonValue = ParseJsonNumber(payloadJson, idx, errorMessage)
        Case "t"
            If MatchJsonLiteral(payloadJson, idx, "true") Then
                ParseJsonValue = True
            Else
                errorMessage = "Invalid literal in PayloadJson."
            End If
        Case "f"
            If MatchJsonLiteral(payloadJson, idx, "false") Then
                ParseJsonValue = False
            Else
                errorMessage = "Invalid literal in PayloadJson."
            End If
        Case "n"
            If MatchJsonLiteral(payloadJson, idx, "null") Then
                ParseJsonValue = vbNullString
            Else
                errorMessage = "Invalid literal in PayloadJson."
            End If
        Case Else
            errorMessage = "Unsupported value in PayloadJson at position " & CStr(idx) & "."
    End Select
End Function

Private Function ParseJsonString(ByVal payloadJson As String, ByRef idx As Long, ByRef errorMessage As String) As String
    Dim result As String
    Dim ch As String
    Dim esc As String

    SkipJsonWhitespace payloadJson, idx
    If Not ConsumeJsonChar(payloadJson, idx, """") Then
        errorMessage = "Expected string value in PayloadJson."
        Exit Function
    End If

    Do While idx <= Len(payloadJson)
        ch = Mid$(payloadJson, idx, 1)
        idx = idx + 1
        If ch = """" Then
            ParseJsonString = result
            Exit Function
        End If
        If ch = "\" Then
            If idx > Len(payloadJson) Then
                errorMessage = "Incomplete escape sequence in PayloadJson."
                Exit Function
            End If
            esc = Mid$(payloadJson, idx, 1)
            idx = idx + 1
            Select Case esc
                Case """", "\", "/"
                    result = result & esc
                Case "b"
                    result = result & Chr$(8)
                Case "f"
                    result = result & Chr$(12)
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case Else
                    errorMessage = "Unsupported escape sequence '\\" & esc & "' in PayloadJson."
                    Exit Function
            End Select
        Else
            result = result & ch
        End If
    Loop

    errorMessage = "Unterminated string in PayloadJson."
End Function

Private Function ParseJsonNumber(ByVal payloadJson As String, ByRef idx As Long, ByRef errorMessage As String) As Double
    Dim startPos As Long
    Dim ch As String
    Dim token As String

    startPos = idx
    Do While idx <= Len(payloadJson)
        ch = Mid$(payloadJson, idx, 1)
        If InStr(1, "-+0123456789.eE", ch, vbBinaryCompare) = 0 Then Exit Do
        idx = idx + 1
    Loop

    token = Mid$(payloadJson, startPos, idx - startPos)
    If token = "" Or Not IsNumeric(token) Then
        errorMessage = "Invalid numeric value in PayloadJson."
        Exit Function
    End If
    ParseJsonNumber = CDbl(token)
End Function

Private Function MatchJsonLiteral(ByVal payloadJson As String, ByRef idx As Long, ByVal literalValue As String) As Boolean
    If StrComp(Mid$(payloadJson, idx, Len(literalValue)), literalValue, vbBinaryCompare) <> 0 Then Exit Function
    idx = idx + Len(literalValue)
    MatchJsonLiteral = True
End Function

Private Sub SkipJsonWhitespace(ByVal payloadJson As String, ByRef idx As Long)
    Do While idx <= Len(payloadJson)
        Select Case Mid$(payloadJson, idx, 1)
            Case " ", vbTab, vbCr, vbLf
                idx = idx + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Function ConsumeJsonChar(ByVal payloadJson As String, ByRef idx As Long, ByVal expectedChar As String) As Boolean
    SkipJsonWhitespace payloadJson, idx
    If idx > Len(payloadJson) Then Exit Function
    If Mid$(payloadJson, idx, 1) <> expectedChar Then Exit Function
    idx = idx + 1
    ConsumeJsonChar = True
End Function

Private Function AppliedEventExists(ByVal lo As ListObject, ByVal eventId As String) As Boolean
    Dim i As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimApply(GetCellByColumnApply(lo, i, "EventID")), eventId, vbTextCompare) = 0 Then
            AppliedEventExists = True
            Exit Function
        End If
    Next i
End Function

Private Function ValidateSkuExists(ByVal wb As Workbook, ByVal sku As String) As Boolean
    Dim hasCatalog As Boolean

    ValidateSkuExists = SearchSkuInTable(FindListObjectByNameApply(wb, "tblSkuCatalog"), sku, hasCatalog)
    If ValidateSkuExists Then Exit Function
    If SearchSkuInTable(FindListObjectByNameApply(wb, "invSys"), sku, hasCatalog) Then
        ValidateSkuExists = True
        Exit Function
    End If
    If SearchSkuInTable(FindListObjectByNameApply(wb, "tblItemSearchIndex"), sku, hasCatalog) Then
        ValidateSkuExists = True
        Exit Function
    End If

    If Not hasCatalog Then ValidateSkuExists = True
End Function

Private Function SearchSkuInTable(ByVal lo As ListObject, ByVal sku As String, ByRef hasCatalog As Boolean) As Boolean
    Dim idx As Long
    Dim i As Long
    Dim valueInRow As String

    If lo Is Nothing Then Exit Function

    idx = GetColumnIndexApply(lo, "SKU")
    If idx = 0 Then idx = GetColumnIndexApply(lo, "ITEM_CODE")
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    hasCatalog = True
    For i = 1 To lo.ListRows.Count
        valueInRow = SafeTrimApply(lo.DataBodyRange.Cells(i, idx).Value)
        If valueInRow <> "" Then
            If StrComp(valueInRow, sku, vbTextCompare) = 0 Then
                SearchSkuInTable = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function GetNextAppliedSeq(ByVal wb As Workbook) As Long
    Dim lo As ListObject
    Dim idx As Long
    Dim i As Long
    Dim currentVal As Variant

    Set lo = FindListObjectByNameApply(wb, "tblAppliedEvents")
    If lo Is Nothing Then
        GetNextAppliedSeq = 1
        Exit Function
    End If

    idx = GetColumnIndexApply(lo, "AppliedSeq")
    If idx = 0 Or lo.DataBodyRange Is Nothing Then
        GetNextAppliedSeq = 1
        Exit Function
    End If

    For i = 1 To lo.ListRows.Count
        currentVal = lo.DataBodyRange.Cells(i, idx).Value
        If IsNumeric(currentVal) Then
            If CLng(currentVal) > GetNextAppliedSeq Then GetNextAppliedSeq = CLng(currentVal)
        End If
    Next i

    GetNextAppliedSeq = GetNextAppliedSeq + 1
End Function

Private Function NormalizeEventType(ByVal eventType As String) As String
    eventType = UCase$(SafeTrimApply(eventType))
    If eventType = "" Then
        NormalizeEventType = EVENT_TYPE_RECEIVE
    Else
        NormalizeEventType = eventType
    End If
End Function

Private Function TryGetEventDate(ByVal evt As Object, ByVal key As String, ByRef valueOut As Date) As Boolean
    Dim rawValue As Variant
    If Not TryGetEventValue(evt, key, rawValue) Then Exit Function
    If Not IsDate(rawValue) Then Exit Function
    valueOut = CDate(rawValue)
    TryGetEventDate = True
End Function

Private Function TryGetEventDouble(ByVal evt As Object, ByVal key As String, ByRef valueOut As Double) As Boolean
    Dim rawValue As Variant
    If Not TryGetEventValue(evt, key, rawValue) Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function
    valueOut = CDbl(rawValue)
    TryGetEventDouble = True
End Function

Private Function GetEventString(ByVal evt As Object, ByVal key As String) As String
    Dim rawValue As Variant
    If TryGetEventValue(evt, key, rawValue) Then
        GetEventString = SafeTrimApply(rawValue)
    End If
End Function

Private Function TryGetEventValue(ByVal evt As Object, ByVal key As String, ByRef valueOut As Variant) As Boolean
    On Error Resume Next
    If evt Is Nothing Then Exit Function
    valueOut = evt(key)
    TryGetEventValue = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetDictionaryValue(ByVal d As Object, ByVal key As String) As Variant
    On Error Resume Next
    If d Is Nothing Then Exit Function
    GetDictionaryValue = d(key)
    On Error GoTo 0
End Function

Private Function TryGetDictionaryDouble(ByVal d As Object, ByVal key As String, ByRef valueOut As Double) As Boolean
    Dim rawValue As Variant
    rawValue = GetDictionaryValue(d, key)
    If IsNumeric(rawValue) Then
        valueOut = CDbl(rawValue)
        TryGetDictionaryDouble = True
    End If
End Function

Private Sub SetTableRowValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexApply(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetCellByColumnApply(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexApply(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnApply = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Function GetColumnIndexApply(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexApply = i
            Exit Function
        End If
    Next i
End Function

Private Function FindListObjectByNameApply(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameApply = ws.ListObjects(tableName)
        If Not FindListObjectByNameApply Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function WorkbookHasListObjectApply(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasListObjectApply = Not (FindListObjectByNameApply(wb, tableName) Is Nothing)
End Function

Private Function IsInventoryWorkbookName(ByVal wbName As String) As Boolean
    Dim n As String
    n = LCase$(wbName)
    IsInventoryWorkbookName = (n Like "wh*.invsys.data.inventory.xlsb") Or _
                              (n Like "wh*.invsys.data.inventory.xlsx") Or _
                              (n Like "wh*.invsys.data.inventory.xlsm")
End Function

Private Function OpenOrCreateCanonicalInventoryWorkbook(ByVal warehouseId As String) As Workbook
    On Error GoTo FailOpen

    Dim targetPath As String
    Dim wb As Workbook
    Dim prevEvents As Boolean
    Dim eventsSuppressed As Boolean

    targetPath = BuildCanonicalInventoryPath(warehouseId)
    If targetPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            Set OpenOrCreateCanonicalInventoryWorkbook = wb
            Exit Function
        End If
    Next wb

    EnsureFolderRecursiveApply GetParentFolderApply(targetPath)
    If Len(Dir$(targetPath)) > 0 Then
        Set wb = Application.Workbooks.Open(targetPath)
    Else
        prevEvents = Application.EnableEvents
        Application.EnableEvents = False
        eventsSuppressed = True
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        wb.SaveAs Filename:=targetPath, FileFormat:=50
        Application.EnableEvents = prevEvents
        eventsSuppressed = False
    End If

    If modInventorySchema.EnsureInventorySchema(wb) Then
        Set OpenOrCreateCanonicalInventoryWorkbook = wb
    End If
    Exit Function

FailOpen:
    On Error Resume Next
    If eventsSuppressed Then Application.EnableEvents = prevEvents
    On Error GoTo 0
End Function

Private Function BuildCanonicalInventoryPath(ByVal warehouseId As String) As String
    Dim resolvedWh As String
    Dim rootPath As String

    resolvedWh = Trim$(warehouseId)
    If resolvedWh = "" Then resolvedWh = SafeTrimApply(modConfig.GetString("WarehouseId", "WH1"))
    If resolvedWh = "" Then resolvedWh = "WH1"

    rootPath = SafeTrimApply(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = SafeTrimApply(modConfig.GetString("PathDataRoot", ""))
    If rootPath = "" Then rootPath = "C:\invSys\" & resolvedWh & "\"

    BuildCanonicalInventoryPath = NormalizeFolderPathApply(rootPath) & resolvedWh & ".invSys.Data.Inventory.xlsb"
End Function

Private Function NormalizeFolderPathApply(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathApply = folderPath
End Function

Private Function GetParentFolderApply(ByVal fullPath As String) As String
    Dim lastSlash As Long

    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then GetParentFolderApply = Left$(fullPath, lastSlash - 1)
End Function

Private Sub EnsureFolderRecursiveApply(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    parentPath = GetParentFolderApply(folderPath)
    If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveApply parentPath

    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
End Sub

Private Function SafeTrimApply(ByVal v As Variant) As String
    On Error Resume Next
    SafeTrimApply = Trim$(CStr(v))
End Function

Private Sub SetSheetProtectionApply(ByVal ws As Worksheet, ByVal protectAfter As Boolean)
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
            Err.Raise vbObjectError + 2201, "modInventoryApply.SetSheetProtectionApply", _
                      "Worksheet '" & ws.Name & "' is protected and could not be unprotected. " & _
                      "Excel automation cannot add table rows while the sheet remains protected."
        End If
    End If
End Sub

Private Sub RebuildInventoryProjections(ByVal wb As Workbook)
    Dim loLog As ListObject
    Dim loSku As ListObject
    Dim loLoc As ListObject
    Dim skuQty As Object
    Dim skuLast As Object
    Dim locQty As Object
    Dim locLast As Object
    Dim rowIndex As Long
    Dim sku As String
    Dim locationVal As String
    Dim qtyDelta As Double
    Dim appliedAt As Variant

    If wb Is Nothing Then Exit Sub

    Set loLog = FindListObjectByNameApply(wb, "tblInventoryLog")
    Set loSku = FindListObjectByNameApply(wb, "tblSkuBalance")
    Set loLoc = FindListObjectByNameApply(wb, "tblLocationBalance")
    If loLog Is Nothing Or loSku Is Nothing Or loLoc Is Nothing Then Exit Sub

    SetSheetProtectionApply loSku.Parent, False
    SetSheetProtectionApply loLoc.Parent, False

    Set skuQty = CreateObject("Scripting.Dictionary")
    skuQty.CompareMode = vbTextCompare
    Set skuLast = CreateObject("Scripting.Dictionary")
    skuLast.CompareMode = vbTextCompare
    Set locQty = CreateObject("Scripting.Dictionary")
    locQty.CompareMode = vbTextCompare
    Set locLast = CreateObject("Scripting.Dictionary")
    locLast.CompareMode = vbTextCompare

    If Not loLog.DataBodyRange Is Nothing Then
        For rowIndex = 1 To loLog.ListRows.Count
            sku = SafeTrimApply(GetCellByColumnApply(loLog, rowIndex, "SKU"))
            If sku = "" Then GoTo ContinueLoop

            qtyDelta = 0#
            If IsNumeric(GetCellByColumnApply(loLog, rowIndex, "QtyDelta")) Then qtyDelta = CDbl(GetCellByColumnApply(loLog, rowIndex, "QtyDelta"))
            locationVal = SafeTrimApply(GetCellByColumnApply(loLog, rowIndex, "Location"))
            appliedAt = GetCellByColumnApply(loLog, rowIndex, "AppliedAtUTC")

            AccumulateProjectionScalars skuQty, skuLast, sku, qtyDelta, appliedAt
            AccumulateProjectionScalars locQty, locLast, sku & "|" & locationVal, qtyDelta, appliedAt
ContinueLoop:
        Next rowIndex
    End If

    RewriteSkuProjectionTable loSku, skuQty, skuLast
    RewriteLocationProjectionTable loLoc, locQty, locLast

    SetSheetProtectionApply loSku.Parent, True
    SetSheetProtectionApply loLoc.Parent, True
End Sub

Private Sub RefreshLedgerStatus(ByVal wb As Workbook, _
                                ByVal warehouseId As String, _
                                ByVal appliedSeq As Long, _
                                ByVal eventId As String, _
                                ByVal appliedAt As Date)
    Dim loStatus As ListObject
    Dim loLog As ListObject
    Dim loApplied As ListObject
    Dim loSku As ListObject
    Dim loLoc As ListObject
    Dim rowIndex As Long

    Set loStatus = FindListObjectByNameApply(wb, "tblInventoryLedgerStatus")
    Set loLog = FindListObjectByNameApply(wb, "tblInventoryLog")
    Set loApplied = FindListObjectByNameApply(wb, "tblAppliedEvents")
    Set loSku = FindListObjectByNameApply(wb, "tblSkuBalance")
    Set loLoc = FindListObjectByNameApply(wb, "tblLocationBalance")
    If loStatus Is Nothing Then Exit Sub

    SetSheetProtectionApply loStatus.Parent, False
    rowIndex = EnsureWritableLedgerStatusRow(loStatus)

    SetTableRowValue loStatus, rowIndex, "WarehouseId", warehouseId
    SetTableRowValue loStatus, rowIndex, "LastAppliedSeq", appliedSeq
    SetTableRowValue loStatus, rowIndex, "LastEventId", eventId
    SetTableRowValue loStatus, rowIndex, "LastAppliedAtUTC", appliedAt
    SetTableRowValue loStatus, rowIndex, "TotalEventRows", CountTableRowsApply(loLog)
    SetTableRowValue loStatus, rowIndex, "TotalAppliedEvents", CountTableRowsApply(loApplied)
    SetTableRowValue loStatus, rowIndex, "DistinctSkuCount", CountTableRowsApply(loSku)
    SetTableRowValue loStatus, rowIndex, "DistinctLocationCount", CountTableRowsApply(loLoc)
    SetTableRowValue loStatus, rowIndex, "ProjectionRebuiltAtUTC", Now
    SetTableRowValue loStatus, rowIndex, "Notes", "Authoritative store: tblInventoryLog + tblAppliedEvents; projections are derived."

    SetSheetProtectionApply loStatus.Parent, True
End Sub

Private Function EnsureWritableLedgerStatusRow(ByVal lo As ListObject) As Long
    If lo Is Nothing Then Exit Function

    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        EnsureWritableLedgerStatusRow = 1
        Exit Function
    End If

    If lo.ListRows.Count = 1 And TableRowIsBlankApply(lo, 1) Then
        EnsureWritableLedgerStatusRow = 1
    Else
        Do While lo.ListRows.Count > 1
            lo.ListRows(lo.ListRows.Count).Delete
        Loop
        EnsureWritableLedgerStatusRow = 1
    End If
End Function

Private Function CountTableRowsApply(ByVal lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    CountTableRowsApply = lo.ListRows.Count
End Function

Private Function TableRowIsBlankApply(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim colIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then
        TableRowIsBlankApply = True
        Exit Function
    End If
    If rowIndex <= 0 Or rowIndex > lo.ListRows.Count Then Exit Function

    TableRowIsBlankApply = True
    For colIndex = 1 To lo.ListColumns.Count
        If SafeTrimApply(lo.DataBodyRange.Cells(rowIndex, colIndex).Value) <> "" Then
            TableRowIsBlankApply = False
            Exit Function
        End If
    Next colIndex
End Function

Private Sub AccumulateProjectionScalars(ByVal qtyDict As Object, _
                                        ByVal lastDict As Object, _
                                        ByVal dictKey As String, _
                                        ByVal qtyDelta As Double, _
                                        ByVal appliedAt As Variant)
    If qtyDict Is Nothing Or lastDict Is Nothing Then Exit Sub

    If qtyDict.Exists(dictKey) Then
        qtyDict(dictKey) = CDbl(qtyDict(dictKey)) + qtyDelta
    Else
        qtyDict.Add dictKey, qtyDelta
    End If

    If IsDate(appliedAt) Then
        If Not lastDict.Exists(dictKey) Then
            lastDict.Add dictKey, CDate(appliedAt)
        ElseIf CDate(appliedAt) > CDate(lastDict(dictKey)) Then
            lastDict(dictKey) = CDate(appliedAt)
        End If
    End If
End Sub

Private Sub RewriteSkuProjectionTable(ByVal lo As ListObject, ByVal qtyDict As Object, ByVal lastDict As Object)
    Dim key As Variant
    Dim r As ListRow

    If lo Is Nothing Then Exit Sub

    ClearProjectionRows lo
    If qtyDict Is Nothing Then Exit Sub

    For Each key In qtyDict.Keys
        Set r = lo.ListRows.Add
        SetTableRowValue lo, r.Index, "SKU", CStr(key)
        SetTableRowValue lo, r.Index, "QtyOnHand", CDbl(qtyDict(key))
        If lastDict.Exists(CStr(key)) Then SetTableRowValue lo, r.Index, "LastAppliedUTC", CDate(lastDict(key))
    Next key
End Sub

Private Sub RewriteLocationProjectionTable(ByVal lo As ListObject, ByVal qtyDict As Object, ByVal lastDict As Object)
    Dim key As Variant
    Dim parts() As String
    Dim r As ListRow

    If lo Is Nothing Then Exit Sub

    ClearProjectionRows lo
    If qtyDict Is Nothing Then Exit Sub

    For Each key In qtyDict.Keys
        parts = Split(CStr(key), "|", 2)
        Set r = lo.ListRows.Add
        SetTableRowValue lo, r.Index, "SKU", parts(0)
        If UBound(parts) >= 1 Then SetTableRowValue lo, r.Index, "Location", parts(1)
        SetTableRowValue lo, r.Index, "QtyOnHand", CDbl(qtyDict(key))
        If lastDict.Exists(CStr(key)) Then SetTableRowValue lo, r.Index, "LastAppliedUTC", CDate(lastDict(key))
    Next key
End Sub

Private Sub ClearProjectionRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
End Sub
