Attribute VB_Name = "modInventoryApply"
Option Explicit

Public Const APPLY_STATUS_APPLIED As String = "APPLIED"
Public Const APPLY_STATUS_SKIP_DUP As String = "SKIP_DUP"

Public Const EVENT_TYPE_RECEIVE As String = "RECEIVE"
Public Const EVENT_TYPE_SHIP As String = "SHIP"
Public Const EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Public Const EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"
Public Const EVENT_TYPE_MIGRATION_SEED As String = "MIGRATION_SEED"

Private mSourceSyncStampCache As Object

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
    Dim migrationSourceId As String
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
    migrationSourceId = GetEventString(evt, "MigrationSourceId")

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
        SetTableRowValue loLog, r.Index, "MigrationSourceId", migrationSourceId
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

Public Function RefreshInvSysFromCanonicalRuntime(ByVal sourceWb As Workbook, _
                                                  Optional ByVal warehouseId As String = "", _
                                                  Optional ByRef report As String = "") As Boolean
    On Error GoTo FailRefresh

    Dim runtimeWb As Workbook
    Dim runtimePath As String
    Dim runtimeWasOpen As Boolean
    Dim loSource As ListObject
    Dim loSku As ListObject
    Dim loLoc As ListObject
    Dim loLog As ListObject
    Dim skuQty As Object
    Dim skuLast As Object
    Dim locSummary As Object
    Dim latestEventType As Object
    Dim latestEventQty As Object
    Dim rowIndex As Long
    Dim sku As String
    Dim sourceSheetWasProtected As Boolean
    Dim sourceSheet As Worksheet
    Dim runtimeRows As Long
    Dim sourceRows As Long
    Dim matchedCount As Long
    Dim changedCount As Long
    Dim configLoadResult As Boolean
    Dim resolvedRootPath As String
    Dim runtimeOpenedReadOnly As Boolean
    Dim runtimeFileStamp As String
    Dim cachedRuntimeFileStamp As String

    If sourceWb Is Nothing Then
        report = "Source workbook not resolved."
        modInventoryInit.AppendSyncLogEntry "TRACE", report
        Exit Function
    End If

    If Trim$(warehouseId) = "" Then warehouseId = ResolveWarehouseIdFromSourceWorkbookApply(sourceWb)
    If Trim$(warehouseId) = "" Then
        report = "SrcWb=" & sourceWb.Name & "|WH=<blank>|Result=WarehouseId not resolved for source workbook."
        modInventoryInit.AppendSyncLogEntry "TRACE", report
        Exit Function
    End If

    If (Not modConfig.IsLoaded()) _
       Or StrComp(SafeTrimApply(modConfig.GetWarehouseId()), warehouseId, vbTextCompare) <> 0 Then
        configLoadResult = modConfig.LoadConfig(warehouseId, "")
    Else
        configLoadResult = True
    End If
    resolvedRootPath = SafeTrimApply(modConfig.GetString("PathDataRoot", ""))

    Set loSource = FindListObjectByNameApply(sourceWb, "invSys")
    If loSource Is Nothing Then
        report = "SrcWb=" & sourceWb.Name & "|WH=" & warehouseId & "|ConfigLoad=" & CStr(configLoadResult) & "|PathDataRoot=" & resolvedRootPath & "|Result=Source invSys table not found."
        modInventoryInit.AppendSyncLogEntry "TRACE", report
        Exit Function
    End If

    runtimePath = BuildCanonicalInventoryPath(warehouseId)
    runtimeWasOpen = WorkbookIsAlreadyOpenApply(runtimePath)
    runtimeFileStamp = ResolveFileStampApply(runtimePath)
    cachedRuntimeFileStamp = GetSourceSyncStampApply(sourceWb)
    If runtimeFileStamp <> "" Then
        If StrComp(runtimeFileStamp, cachedRuntimeFileStamp, vbTextCompare) = 0 Then
            report = "SrcWb=" & sourceWb.Name & "|WH=" & warehouseId & "|ConfigLoad=" & CStr(configLoadResult) & "|PathDataRoot=" & resolvedRootPath & "|RuntimePath=" & runtimePath & "|RuntimeStamp=" & runtimeFileStamp & "|Result=UNCHANGED"
            modInventoryInit.AppendSyncLogEntry "TRACE", report
            RefreshInvSysFromCanonicalRuntime = True
            Exit Function
        End If
    End If
    Set runtimeWb = ResolveInventoryWorkbook(warehouseId)
    If runtimeWb Is Nothing And FileExistsApply(runtimePath) Then
        Set runtimeWb = OpenWorkbookReadOnlyApply(runtimePath)
        If Not runtimeWb Is Nothing Then runtimeOpenedReadOnly = True
    End If
    If runtimeWb Is Nothing Then
        report = "SrcWb=" & sourceWb.Name & "|WH=" & warehouseId & "|ConfigLoad=" & CStr(configLoadResult) & "|PathDataRoot=" & resolvedRootPath & "|RuntimePath=" & runtimePath & "|RuntimeWasOpen=" & CStr(runtimeWasOpen) & "|Result=Canonical runtime inventory workbook not found."
        modInventoryInit.AppendSyncLogEntry "TRACE", report
        Exit Function
    End If
    If Not runtimeWasOpen Then HideWorkbookWindowsApply runtimeWb

    Set loSku = FindListObjectByNameApply(runtimeWb, "tblSkuBalance")
    Set loLoc = FindListObjectByNameApply(runtimeWb, "tblLocationBalance")
    Set loLog = FindListObjectByNameApply(runtimeWb, "tblInventoryLog")
    If loSku Is Nothing Then
        report = "SrcWb=" & sourceWb.Name & "|WH=" & warehouseId & "|ConfigLoad=" & CStr(configLoadResult) & "|PathDataRoot=" & resolvedRootPath & "|RuntimePath=" & runtimePath & "|RuntimeWasOpen=" & CStr(runtimeWasOpen) & "|Result=Canonical runtime projection table tblSkuBalance not found."
        modInventoryInit.AppendSyncLogEntry "TRACE", report
        GoTo CleanExit
    End If
    runtimeRows = loSku.ListRows.Count
    sourceRows = loSource.ListRows.Count

    Set skuQty = CreateObject("Scripting.Dictionary")
    skuQty.CompareMode = vbTextCompare
    Set skuLast = CreateObject("Scripting.Dictionary")
    skuLast.CompareMode = vbTextCompare
    Set locSummary = CreateObject("Scripting.Dictionary")
    locSummary.CompareMode = vbTextCompare
    Set latestEventType = CreateObject("Scripting.Dictionary")
    latestEventType.CompareMode = vbTextCompare
    Set latestEventQty = CreateObject("Scripting.Dictionary")
    latestEventQty.CompareMode = vbTextCompare

    BuildSkuProjectionDictionariesApply loSku, skuQty, skuLast
    BuildLocationSummaryDictionaryApply loLoc, locSummary
    BuildLatestMovementDictionariesApply loLog, latestEventType, latestEventQty

    Set sourceSheet = loSource.Parent
    sourceSheetWasProtected = sourceSheet.ProtectContents
    SetSheetProtectionApply sourceSheet, False

    If Not loSource.DataBodyRange Is Nothing Then
        For rowIndex = 1 To loSource.ListRows.Count
            sku = ResolveInvSysSkuApply(loSource, rowIndex)
            If sku <> "" Then
                If skuQty.Exists(sku) Or locSummary.Exists(sku) Then
                    matchedCount = matchedCount + 1
                    If CanonicalRuntimeRowWouldChangeApply(loSource, rowIndex, skuQty, locSummary, latestEventType, latestEventQty, sku) Then changedCount = changedCount + 1
                End If
                ApplyCanonicalRuntimeRowApply loSource, rowIndex, skuQty, skuLast, locSummary, latestEventType, latestEventQty, sku
            End If
        Next rowIndex
    End If

    If runtimeFileStamp <> "" Then SetSourceSyncStampApply sourceWb, runtimeFileStamp
    report = "SrcWb=" & sourceWb.Name & "|WH=" & warehouseId & "|ConfigLoad=" & CStr(configLoadResult) & "|PathDataRoot=" & resolvedRootPath & "|RuntimePath=" & runtimePath & "|RuntimeStamp=" & runtimeFileStamp & "|RuntimeWasOpen=" & CStr(runtimeWasOpen) & "|RuntimeReadOnly=" & CStr(runtimeOpenedReadOnly) & "|RuntimeRows=" & CStr(runtimeRows) & "|SrcInvSysRows=" & CStr(sourceRows) & "|MatchedSKUs=" & CStr(matchedCount) & "|ChangedRows=" & CStr(changedCount) & "|Result=OK"
    modInventoryInit.AppendSyncLogEntry "TRACE", report
    RefreshInvSysFromCanonicalRuntime = True

CleanExit:
    On Error Resume Next
    If sourceSheetWasProtected Then SetSheetProtectionApply sourceSheet, True
    If Not runtimeWasOpen Then CloseWorkbookQuietlyApply runtimeWb
    On Error GoTo 0
    Exit Function

FailRefresh:
    report = "SrcWb=" & sourceWb.Name & "|WH=" & warehouseId & "|ConfigLoad=" & CStr(configLoadResult) & "|PathDataRoot=" & resolvedRootPath & "|RuntimePath=" & runtimePath & "|RuntimeStamp=" & runtimeFileStamp & "|RuntimeWasOpen=" & CStr(runtimeWasOpen) & "|RuntimeReadOnly=" & CStr(runtimeOpenedReadOnly) & "|RuntimeRows=" & CStr(runtimeRows) & "|SrcInvSysRows=" & CStr(sourceRows) & "|MatchedSKUs=" & CStr(matchedCount) & "|ChangedRows=" & CStr(changedCount) & "|Result=RefreshInvSysFromCanonicalRuntime failed: " & Err.Description
    modInventoryInit.AppendSyncLogEntry "TRACE", report
    On Error Resume Next
    If sourceSheetWasProtected Then SetSheetProtectionApply sourceSheet, True
    If Not runtimeWasOpen Then CloseWorkbookQuietlyApply runtimeWb
    On Error GoTo 0
End Function

Private Sub EnsureSourceSyncStampCacheApply()
    If mSourceSyncStampCache Is Nothing Then
        Set mSourceSyncStampCache = CreateObject("Scripting.Dictionary")
        mSourceSyncStampCache.CompareMode = vbTextCompare
    End If
End Sub

Private Function BuildSourceSyncCacheKeyApply(ByVal wb As Workbook) As String
    If wb Is Nothing Then Exit Function
    If Trim$(wb.FullName) <> "" Then
        BuildSourceSyncCacheKeyApply = LCase$(Trim$(wb.FullName))
    Else
        BuildSourceSyncCacheKeyApply = LCase$(Trim$(wb.Name))
    End If
End Function

Private Function GetSourceSyncStampApply(ByVal wb As Workbook) As String
    Dim cacheKey As String

    EnsureSourceSyncStampCacheApply
    cacheKey = BuildSourceSyncCacheKeyApply(wb)
    If cacheKey = "" Then Exit Function
    If mSourceSyncStampCache.Exists(cacheKey) Then GetSourceSyncStampApply = CStr(mSourceSyncStampCache(cacheKey))
End Function

Private Sub SetSourceSyncStampApply(ByVal wb As Workbook, ByVal fileStamp As String)
    Dim cacheKey As String

    EnsureSourceSyncStampCacheApply
    cacheKey = BuildSourceSyncCacheKeyApply(wb)
    If cacheKey = "" Then Exit Sub
    If mSourceSyncStampCache.Exists(cacheKey) Then mSourceSyncStampCache.Remove cacheKey
    mSourceSyncStampCache.Add cacheKey, fileStamp
End Sub

Private Function BuildApplyLines(ByVal evt As Object, _
                                 ByVal wb As Workbook, _
                                 ByVal eventType As String, _
                                 ByRef errorCode As String, _
                                 ByRef errorMessage As String) As Collection
    Select Case eventType
        Case EVENT_TYPE_RECEIVE
            Set BuildApplyLines = BuildReceiveLines(evt, wb, errorCode, errorMessage)
        Case EVENT_TYPE_SHIP, EVENT_TYPE_PROD_CONSUME, EVENT_TYPE_PROD_COMPLETE, EVENT_TYPE_MIGRATION_SEED
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
        If eventType = EVENT_TYPE_MIGRATION_SEED Then
            EnsureSkuCatalogFromPayloadLineApply wb, rawItem
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
        Case EVENT_TYPE_MIGRATION_SEED
            If ioType <> "" And ioType <> "MADE" And ioType <> "SEED" And ioType <> "IMPORT" Then
                errorCode = "INVALID_PAYLOAD"
                errorMessage = "MIGRATION_SEED payload line items may only use IoType MADE, SEED, or IMPORT."
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

Private Sub EnsureSkuCatalogFromPayloadLineApply(ByVal wb As Workbook, ByVal rawItem As Object)
    Dim lo As ListObject
    Dim sku As String
    Dim rowIndex As Long
    Dim r As ListRow

    On Error GoTo CleanExit
    If wb Is Nothing Then Exit Sub
    If rawItem Is Nothing Then Exit Sub

    Set lo = FindListObjectByNameApply(wb, "tblSkuCatalog")
    If lo Is Nothing Then Exit Sub

    sku = SafeTrimApply(GetDictionaryValue(rawItem, "SKU"))
    If sku = "" Then Exit Sub

    rowIndex = FindRowByColumnValueApply(lo, "SKU", sku)
    If rowIndex = 0 Then
        SetSheetProtectionApply lo.Parent, False
        Set r = lo.ListRows.Add
        rowIndex = r.Index
        SetTableRowValue lo, rowIndex, "SKU", sku
        SetTableRowValue lo, rowIndex, "ITEM_CODE", ResolvePayloadTextApply(rawItem, "ITEM_CODE", sku)
        SetTableRowValue lo, rowIndex, "ITEM", ResolvePayloadTextApply(rawItem, "ITEM", sku)
        SetTableRowValue lo, rowIndex, "UOM", ResolvePayloadTextApply(rawItem, "UOM", "")
        SetTableRowValue lo, rowIndex, "LOCATION", ResolvePayloadTextApply(rawItem, "LOCATION", ResolvePayloadTextApply(rawItem, "Location", ""))
        SetTableRowValue lo, rowIndex, "DESCRIPTION", ResolvePayloadTextApply(rawItem, "DESCRIPTION", "")
        SetTableRowValue lo, rowIndex, "VENDOR(s)", ResolvePayloadTextApply(rawItem, "VENDOR(s)", "")
        SetTableRowValue lo, rowIndex, "VENDOR_CODE", ResolvePayloadTextApply(rawItem, "VENDOR_CODE", "")
        SetTableRowValue lo, rowIndex, "CATEGORY", ResolvePayloadTextApply(rawItem, "CATEGORY", "")
        SetSheetProtectionApply lo.Parent, True
    End If

CleanExit:
    On Error Resume Next
    If Not lo Is Nothing Then SetSheetProtectionApply lo.Parent, True
    On Error GoTo 0
End Sub

Private Function ResolvePayloadTextApply(ByVal rawItem As Object, ByVal keyName As String, ByVal defaultValue As String) As String
    ResolvePayloadTextApply = SafeTrimApply(GetDictionaryValue(rawItem, keyName))
    If ResolvePayloadTextApply = "" Then ResolvePayloadTextApply = defaultValue
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

Private Function FindRowByColumnValueApply(ByVal lo As ListObject, ByVal columnName As String, ByVal matchValue As String) As Long
    Dim idx As Long
    Dim rowIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idx = GetColumnIndexApply(lo, columnName)
    If idx = 0 Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(SafeTrimApply(lo.DataBodyRange.Cells(rowIndex, idx).Value), SafeTrimApply(matchValue), vbTextCompare) = 0 Then
            FindRowByColumnValueApply = rowIndex
            Exit Function
        End If
    Next rowIndex
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
        If IsWorkbookFileLockedApply(targetPath) Then Exit Function
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

Private Function IsWorkbookFileLockedApply(ByVal targetPath As String) As Boolean
    Dim fileNum As Integer

    If Len(Dir$(targetPath)) = 0 Then Exit Function

    On Error GoTo Locked
    fileNum = FreeFile
    Open targetPath For Binary Access Read Write Lock Read Write As #fileNum
    Close #fileNum
    Exit Function

Locked:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    IsWorkbookFileLockedApply = True
End Function

Private Function BuildCanonicalInventoryPath(ByVal warehouseId As String) As String
    Dim resolvedWh As String
    Dim rootPath As String

    resolvedWh = Trim$(warehouseId)
    If resolvedWh = "" Then resolvedWh = SafeTrimApply(modConfig.GetString("WarehouseId", "WH1"))
    If resolvedWh = "" Then resolvedWh = "WH1"

    rootPath = SafeTrimApply(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = SafeTrimApply(modConfig.GetString("PathDataRoot", ""))
    If rootPath = "" Then rootPath = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(resolvedWh, True)

    BuildCanonicalInventoryPath = NormalizeFolderPathApply(rootPath) & resolvedWh & ".invSys.Data.Inventory.xlsb"
End Function

Private Function ResolveWarehouseIdFromSourceWorkbookApply(ByVal wb As Workbook) As String
    Dim lo As ListObject
    Dim idx As Long
    Dim rowIndex As Long
    Dim snapshotId As String

    If wb Is Nothing Then Exit Function

    Set lo = FindListObjectByNameApply(wb, "invSys")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idx = GetColumnIndexApply(lo, "SnapshotId")
    If idx > 0 Then
        For rowIndex = 1 To lo.ListRows.Count
            snapshotId = SafeTrimApply(lo.DataBodyRange.Cells(rowIndex, idx).Value)
            ResolveWarehouseIdFromSourceWorkbookApply = ResolveWarehouseIdFromSnapshotIdApply(snapshotId)
            If ResolveWarehouseIdFromSourceWorkbookApply <> "" Then Exit Function
        Next rowIndex
    End If

    ResolveWarehouseIdFromSourceWorkbookApply = InferWarehouseIdFromWorkbookNameApply(wb.Name)
    If ResolveWarehouseIdFromSourceWorkbookApply = "" Then
        ResolveWarehouseIdFromSourceWorkbookApply = SafeTrimApply(modConfig.GetWarehouseId())
    End If
End Function

Private Function ResolveWarehouseIdFromSnapshotIdApply(ByVal snapshotId As String) As String
    Dim markerPos As Long

    snapshotId = Trim$(snapshotId)
    If snapshotId = "" Then Exit Function
    markerPos = InStr(1, snapshotId, ".invSys.Snapshot.Inventory.xls", vbTextCompare)
    If markerPos > 1 Then ResolveWarehouseIdFromSnapshotIdApply = Left$(snapshotId, markerPos - 1)
End Function

Private Function InferWarehouseIdFromWorkbookNameApply(ByVal wbName As String) As String
    Dim markerPos As Long

    markerPos = InStr(1, wbName, ".invSys.", vbTextCompare)
    If markerPos > 1 Then
        InferWarehouseIdFromWorkbookNameApply = Left$(wbName, markerPos - 1)
        Exit Function
    End If

    markerPos = InStr(1, wbName, "_", vbTextCompare)
    If markerPos > 1 Then InferWarehouseIdFromWorkbookNameApply = Left$(wbName, markerPos - 1)
End Function

Private Sub BuildSkuProjectionDictionariesApply(ByVal loSku As ListObject, _
                                                ByVal skuQty As Object, _
                                                ByVal skuLast As Object)
    Dim rowIndex As Long
    Dim sku As String
    Dim appliedAt As Variant

    If loSku Is Nothing Then Exit Sub
    If loSku.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To loSku.ListRows.Count
        sku = SafeTrimApply(GetCellByColumnApply(loSku, rowIndex, "SKU"))
        If sku = "" Then GoTo ContinueLoop

        skuQty(sku) = NzDblApply(GetCellByColumnApply(loSku, rowIndex, "QtyOnHand"))
        appliedAt = GetCellByColumnApply(loSku, rowIndex, "LastAppliedUTC")
        If IsDate(appliedAt) Then skuLast(sku) = CDate(appliedAt)
ContinueLoop:
    Next rowIndex
End Sub

Private Sub BuildLocationSummaryDictionaryApply(ByVal loLoc As ListObject, ByVal locSummary As Object)
    Dim rowIndex As Long
    Dim sku As String
    Dim locationVal As String
    Dim qtyOnHand As Double
    Dim fragment As String

    If loLoc Is Nothing Then Exit Sub
    If loLoc.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To loLoc.ListRows.Count
        sku = SafeTrimApply(GetCellByColumnApply(loLoc, rowIndex, "SKU"))
        If sku = "" Then GoTo ContinueLoop
        locationVal = SafeTrimApply(GetCellByColumnApply(loLoc, rowIndex, "Location"))
        qtyOnHand = NzDblApply(GetCellByColumnApply(loLoc, rowIndex, "QtyOnHand"))
        If locationVal = "" Then GoTo ContinueLoop

        fragment = locationVal & "=" & FormatQuantityApply(qtyOnHand)
        If locSummary.Exists(sku) Then
            locSummary(sku) = CStr(locSummary(sku)) & "; " & fragment
        Else
            locSummary(sku) = fragment
        End If
ContinueLoop:
    Next rowIndex
End Sub

Private Sub BuildLatestMovementDictionariesApply(ByVal loLog As ListObject, _
                                                 ByVal latestEventType As Object, _
                                                 ByVal latestEventQty As Object)
    Dim rowIndex As Long
    Dim sku As String
    Dim eventType As String
    Dim qtyDelta As Double
    Dim appliedAt As Variant
    Dim stamp As Double
    Dim bestStamp As Double
    Dim stampMap As Object

    If loLog Is Nothing Then Exit Sub
    If loLog.DataBodyRange Is Nothing Then Exit Sub

    Set stampMap = CreateObject("Scripting.Dictionary")
    stampMap.CompareMode = vbTextCompare

    For rowIndex = 1 To loLog.ListRows.Count
        sku = SafeTrimApply(GetCellByColumnApply(loLog, rowIndex, "SKU"))
        If sku = "" Then GoTo ContinueLoop

        eventType = NormalizeEventType(SafeTrimApply(GetCellByColumnApply(loLog, rowIndex, "EventType")))
        qtyDelta = NzDblApply(GetCellByColumnApply(loLog, rowIndex, "QtyDelta"))
        appliedAt = GetCellByColumnApply(loLog, rowIndex, "AppliedAtUTC")
        If IsDate(appliedAt) Then
            stamp = CDbl(CDate(appliedAt))
        Else
            stamp = CDbl(rowIndex)
        End If

        If stampMap.Exists(sku) Then bestStamp = CDbl(stampMap(sku))
        If (Not stampMap.Exists(sku)) Or stamp >= bestStamp Then
            stampMap(sku) = stamp
            latestEventType(sku) = eventType
            latestEventQty(sku) = Abs(qtyDelta)
        End If
ContinueLoop:
    Next rowIndex
End Sub

Private Sub ApplyCanonicalRuntimeRowApply(ByVal loSource As ListObject, _
                                          ByVal rowIndex As Long, _
                                          ByVal skuQty As Object, _
                                          ByVal skuLast As Object, _
                                          ByVal locSummary As Object, _
                                          ByVal latestEventType As Object, _
                                          ByVal latestEventQty As Object, _
                                          ByVal sku As String)
    Dim qtyOnHand As Double
    Dim summaryText As String
    Dim appliedAt As Variant
    Dim primaryLocation As String

    If skuQty.Exists(sku) Then qtyOnHand = CDbl(skuQty(sku))
    If locSummary.Exists(sku) Then summaryText = CStr(locSummary(sku))
    If skuLast.Exists(sku) Then appliedAt = skuLast(sku)

    SetInvSysValueApply loSource, rowIndex, "TOTAL INV", qtyOnHand
    SetInvSysValueApply loSource, rowIndex, "QtyAvailable", qtyOnHand
    If summaryText <> "" Then
        SetInvSysValueApply loSource, rowIndex, "LocationSummary", summaryText
        primaryLocation = ResolvePrimaryLocationFromSummaryApply(summaryText)
        If primaryLocation <> "" Then SetInvSysValueApply loSource, rowIndex, "LOCATION", primaryLocation
    Else
        SetInvSysValueApply loSource, rowIndex, "LocationSummary", vbNullString
    End If
    If IsDate(appliedAt) Then
        SetInvSysValueApply loSource, rowIndex, "LAST EDITED", CDate(appliedAt)
        SetInvSysValueApply loSource, rowIndex, "TOTAL INV LAST EDIT", CDate(appliedAt)
    Else
        SetInvSysValueApply loSource, rowIndex, "LAST EDITED", vbNullString
        SetInvSysValueApply loSource, rowIndex, "TOTAL INV LAST EDIT", vbNullString
    End If
    ApplyLatestMovementToInvSysApply loSource, rowIndex, latestEventType, latestEventQty, sku
    SetInvSysValueApply loSource, rowIndex, "LastRefreshUTC", Now
    SetInvSysValueApply loSource, rowIndex, "SourceType", "CANONICAL_RUNTIME"
    SetInvSysValueApply loSource, rowIndex, "IsStale", False
End Sub

Private Function CanonicalRuntimeRowWouldChangeApply(ByVal loSource As ListObject, _
                                                     ByVal rowIndex As Long, _
                                                     ByVal skuQty As Object, _
                                                     ByVal locSummary As Object, _
                                                     ByVal latestEventType As Object, _
                                                     ByVal latestEventQty As Object, _
                                                     ByVal sku As String) As Boolean
    Dim qtyOnHand As Double
    Dim summaryText As String
    Dim primaryLocation As String
    Dim expectedReceived As Double
    Dim expectedUsed As Double
    Dim expectedMade As Double
    Dim expectedShipments As Double

    If skuQty.Exists(sku) Then qtyOnHand = CDbl(skuQty(sku))
    If locSummary.Exists(sku) Then summaryText = CStr(locSummary(sku))
    If summaryText <> "" Then primaryLocation = ResolvePrimaryLocationFromSummaryApply(summaryText)
    ResolveLatestMovementValuesApply latestEventType, latestEventQty, sku, expectedReceived, expectedUsed, expectedMade, expectedShipments

    If ValuesDifferNumericApply(GetCellByColumnApply(loSource, rowIndex, "TOTAL INV"), qtyOnHand) Then
        CanonicalRuntimeRowWouldChangeApply = True
        Exit Function
    End If
    If ValuesDifferNumericApply(GetCellByColumnApply(loSource, rowIndex, "QtyAvailable"), qtyOnHand) Then
        CanonicalRuntimeRowWouldChangeApply = True
        Exit Function
    End If
    If ValuesDifferTextApply(GetCellByColumnApply(loSource, rowIndex, "LocationSummary"), summaryText) Then
        CanonicalRuntimeRowWouldChangeApply = True
        Exit Function
    End If
    If primaryLocation <> "" Then
        If ValuesDifferTextApply(GetCellByColumnApply(loSource, rowIndex, "LOCATION"), primaryLocation) Then
            CanonicalRuntimeRowWouldChangeApply = True
            Exit Function
        End If
    End If
    If ValuesDifferNumericApply(GetCellByColumnApply(loSource, rowIndex, "RECEIVED"), expectedReceived) Then
        CanonicalRuntimeRowWouldChangeApply = True
        Exit Function
    End If
    If ValuesDifferNumericApply(GetCellByColumnApply(loSource, rowIndex, "USED"), expectedUsed) Then
        CanonicalRuntimeRowWouldChangeApply = True
        Exit Function
    End If
    If ValuesDifferNumericApply(GetCellByColumnApply(loSource, rowIndex, "MADE"), expectedMade) Then
        CanonicalRuntimeRowWouldChangeApply = True
        Exit Function
    End If
    If ValuesDifferNumericApply(GetCellByColumnApply(loSource, rowIndex, "SHIPMENTS"), expectedShipments) Then
        CanonicalRuntimeRowWouldChangeApply = True
    End If
End Function

Private Sub ApplyLatestMovementToInvSysApply(ByVal loSource As ListObject, _
                                             ByVal rowIndex As Long, _
                                             ByVal latestEventType As Object, _
                                             ByVal latestEventQty As Object, _
                                             ByVal sku As String)
    Dim expectedReceived As Double
    Dim expectedUsed As Double
    Dim expectedMade As Double
    Dim expectedShipments As Double

    ResolveLatestMovementValuesApply latestEventType, latestEventQty, sku, expectedReceived, expectedUsed, expectedMade, expectedShipments
    SetInvSysValueApply loSource, rowIndex, "RECEIVED", expectedReceived
    SetInvSysValueApply loSource, rowIndex, "USED", expectedUsed
    SetInvSysValueApply loSource, rowIndex, "MADE", expectedMade
    SetInvSysValueApply loSource, rowIndex, "SHIPMENTS", expectedShipments
End Sub

Private Sub ResolveLatestMovementValuesApply(ByVal latestEventType As Object, _
                                             ByVal latestEventQty As Object, _
                                             ByVal sku As String, _
                                             ByRef receivedOut As Double, _
                                             ByRef usedOut As Double, _
                                             ByRef madeOut As Double, _
                                             ByRef shipmentsOut As Double)
    Dim eventType As String
    Dim qty As Double

    If latestEventType Is Nothing Then Exit Sub
    If latestEventQty Is Nothing Then Exit Sub
    If Not latestEventType.Exists(sku) Then Exit Sub

    eventType = UCase$(SafeTrimApply(latestEventType(sku)))
    If latestEventQty.Exists(sku) Then qty = NzDblApply(latestEventQty(sku))

    Select Case eventType
        Case EVENT_TYPE_RECEIVE
            receivedOut = qty
        Case EVENT_TYPE_SHIP
            shipmentsOut = qty
        Case EVENT_TYPE_PROD_CONSUME
            usedOut = qty
        Case EVENT_TYPE_PROD_COMPLETE
            madeOut = qty
    End Select
End Sub

Private Function ResolveInvSysSkuApply(ByVal lo As ListObject, ByVal rowIndex As Long) As String
    ResolveInvSysSkuApply = SafeTrimApply(GetCellByColumnApply(lo, rowIndex, "ITEM_CODE"))
    If ResolveInvSysSkuApply = "" Then ResolveInvSysSkuApply = SafeTrimApply(GetCellByColumnApply(lo, rowIndex, "SKU"))
End Function

Private Sub SetInvSysValueApply(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    idx = GetColumnIndexApply(lo, columnName)
    If idx = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function ValuesDifferNumericApply(ByVal currentValue As Variant, ByVal expectedValue As Double) As Boolean
    ValuesDifferNumericApply = (Abs(NzDblApply(currentValue) - expectedValue) > 0.0001#)
End Function

Private Function ValuesDifferTextApply(ByVal currentValue As Variant, ByVal expectedValue As String) As Boolean
    ValuesDifferTextApply = (StrComp(SafeTrimApply(currentValue), SafeTrimApply(expectedValue), vbTextCompare) <> 0)
End Function

Private Function ResolvePrimaryLocationFromSummaryApply(ByVal summaryText As String) As String
    Dim firstFragment As String
    Dim eqPos As Long

    summaryText = Trim$(summaryText)
    If summaryText = "" Then Exit Function
    firstFragment = Trim$(Split(summaryText, ";")(0))
    eqPos = InStr(1, firstFragment, "=", vbTextCompare)
    If eqPos > 1 Then
        ResolvePrimaryLocationFromSummaryApply = Trim$(Left$(firstFragment, eqPos - 1))
    Else
        ResolvePrimaryLocationFromSummaryApply = firstFragment
    End If
End Function

Private Function WorkbookIsAlreadyOpenApply(ByVal fullPath As String) As Boolean
    Dim wb As Workbook

    fullPath = Trim$(fullPath)
    If fullPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            WorkbookIsAlreadyOpenApply = True
            Exit Function
        End If
    Next wb
End Function

Private Function FileExistsApply(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExistsApply = (Len(Dir$(fullPath)) > 0)
    On Error GoTo 0
End Function

Private Function ResolveFileStampApply(ByVal fullPath As String) As String
    On Error GoTo FailStamp

    If Trim$(fullPath) = "" Then Exit Function
    If Not FileExistsApply(fullPath) Then Exit Function
    ResolveFileStampApply = Format$(FileDateTime(fullPath), "yyyymmddhhnnss")
    Exit Function

FailStamp:
    ResolveFileStampApply = vbNullString
End Function

Private Function OpenWorkbookReadOnlyApply(ByVal fullPath As String) As Workbook
    On Error GoTo FailOpen

    If Trim$(fullPath) = "" Then Exit Function
    If Not FileExistsApply(fullPath) Then Exit Function

    Set OpenWorkbookReadOnlyApply = Application.Workbooks.Open(Filename:=fullPath, ReadOnly:=True, Notify:=False)
    If Not OpenWorkbookReadOnlyApply Is Nothing Then HideWorkbookWindowsApply OpenWorkbookReadOnlyApply
    Exit Function

FailOpen:
    Set OpenWorkbookReadOnlyApply = Nothing
End Function

Private Sub HideWorkbookWindowsApply(ByVal wb As Workbook)
    Dim i As Long

    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To wb.Windows.Count
        wb.Windows(i).Visible = False
    Next i
    modUiQuiet.ReactivateQuietOwner
    On Error GoTo 0
End Sub

Private Sub CloseWorkbookQuietlyApply(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    HideWorkbookWindowsApply wb
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function FormatQuantityApply(ByVal qty As Double) As String
    FormatQuantityApply = Replace$(Format$(qty, "0.########"), ",", "")
End Function

Private Function NzDblApply(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzDblApply = CDbl(valueIn)
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
