Attribute VB_Name = "TestPhase2Helpers"
Option Explicit

Public Function BuildPhase2ConfigWorkbook(ByVal whId As String, ByVal stId As String, _
                                          Optional ByVal roleDefault As String = "RECEIVE", _
                                          Optional ByVal processorServiceUserId As String = "svc_processor") As Workbook
    Dim wb As Workbook
    Dim wsWh As Worksheet
    Dim wsSt As Worksheet
    Dim p As String

    Set wb = Application.Workbooks.Add
    Set wsWh = wb.Worksheets(1)
    wsWh.Name = "WarehouseConfig"
    Set wsSt = wb.Worksheets.Add(After:=wsWh)
    wsSt.Name = "StationConfig"

    wsWh.Range("A1").Resize(1, 21).Value = Array( _
        "WarehouseId", "WarehouseName", "Timezone", "DefaultLocation", _
        "BatchSize", "LockTimeoutMinutes", "HeartbeatIntervalSeconds", "MaxLockHoldMinutes", _
        "SnapshotCadence", "BackupCadence", "PathDataRoot", "PathBackupRoot", "PathSharePointRoot", _
        "DesignsEnabled", "PoisonRetryMax", "AuthCacheTTLSeconds", "ProcessorServiceUserId", _
        "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot", "AutoRefreshIntervalSeconds")
    wsWh.Range("A2").Resize(1, 21).Value = Array( _
        whId, "Main Warehouse", "UTC", "A1", _
        500, 3, 30, 2, _
        "PER_BATCH", "DAILY", "C:\invSys\" & whId & "\", "C:\invSys\Backups\" & whId & "\", "", _
        False, 3, 300, processorServiceUserId, _
        False, False, True, 0)
    wsWh.ListObjects.Add(xlSrcRange, wsWh.Range("A1:U2"), , xlYes).Name = "tblWarehouseConfig"

    wsSt.Range("A1").Resize(1, 5).Value = Array("StationId", "WarehouseId", "StationName", "PathInboxRoot", "RoleDefault")
    wsSt.Range("A2").Resize(1, 5).Value = Array(stId, whId, Environ$("COMPUTERNAME"), "", roleDefault)
    wsSt.ListObjects.Add(xlSrcRange, wsSt.Range("A1:E2"), , xlYes).Name = "tblStationConfig"

    p = BuildUniqueTestWorkbookPath(whId & ".invSys.Config")
    SaveWorkbookAsTestFile wb, p, 51
    Set BuildPhase2ConfigWorkbook = wb
End Function

Public Sub SetWarehouseConfigValue(ByVal wb As Workbook, ByVal keyName As String, ByVal valueIn As Variant)
    Dim lo As ListObject
    Dim idx As Long

    If wb Is Nothing Then Exit Sub
    Set lo = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    idx = lo.ListColumns(keyName).Index
    lo.DataBodyRange.Cells(1, idx).Value = valueIn
End Sub

Public Sub SetStationConfigValue(ByVal wb As Workbook, ByVal keyName As String, ByVal valueIn As Variant)
    Dim lo As ListObject
    Dim idx As Long

    If wb Is Nothing Then Exit Sub
    Set lo = wb.Worksheets("StationConfig").ListObjects("tblStationConfig")
    idx = lo.ListColumns(keyName).Index
    lo.DataBodyRange.Cells(1, idx).Value = valueIn
End Sub

Public Function BuildPhase2AuthWorkbook(ByVal whId As String, _
                                        Optional ByVal processorServiceUserId As String = "svc_processor") As Workbook
    Dim wb As Workbook
    Dim wsUsers As Worksheet
    Dim wsCaps As Worksheet
    Dim p As String

    Set wb = Application.Workbooks.Add
    Set wsUsers = wb.Worksheets(1)
    wsUsers.Name = "Users"
    Set wsCaps = wb.Worksheets.Add(After:=wsUsers)
    wsCaps.Name = "Capabilities"

    wsUsers.Range("A1").Resize(1, 6).Value = Array("UserId", "DisplayName", "PinHash", "Status", "ValidFrom", "ValidTo")
    wsUsers.Range("A2").Resize(1, 6).Value = Array("user1", "User One", "", "Active", "", "")
    wsUsers.Range("A3").Resize(1, 6).Value = Array("user2", "User Two", "", "Active", "", "")
    wsUsers.Range("A4").Resize(1, 6).Value = Array(processorServiceUserId, "Processor Service", "", "Active", "", "")
    wsUsers.ListObjects.Add(xlSrcRange, wsUsers.Range("A1:F4"), , xlYes).Name = "tblUsers"

    wsCaps.Range("A1").Resize(1, 7).Value = Array("UserId", "Capability", "WarehouseId", "StationId", "Status", "ValidFrom", "ValidTo")
    wsCaps.Range("A2").Resize(1, 7).Value = Array("", "", "", "", "", "", "")
    wsCaps.ListObjects.Add(xlSrcRange, wsCaps.Range("A1:G2"), , xlYes).Name = "tblCapabilities"

    p = BuildUniqueTestWorkbookPath(whId & ".invSys.Auth")
    SaveWorkbookAsTestFile wb, p, 51
    Set BuildPhase2AuthWorkbook = wb
End Function

Public Function BuildPhase2InventoryWorkbook(ByVal whId As String, _
                                             Optional ByVal skuList As Variant, _
                                             Optional ByVal persistWorkbook As Boolean = True) As Workbook
    Dim wb As Workbook
    Dim wsSku As Worksheet
    Dim loSku As ListObject
    Dim lastRow As Long
    Dim i As Long
    Dim p As String

    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Name = "InventoryLog"
    Call modInventorySchema.EnsureInventorySchema(wb)
    DeleteAllTableRows wb.Worksheets("InventoryLog").ListObjects("tblInventoryLog"), True
    DeleteAllTableRows wb.Worksheets("AppliedEvents").ListObjects("tblAppliedEvents"), True
    DeleteAllTableRows wb.Worksheets("Locks").ListObjects("tblLocks"), True
    DeleteAllTableRows wb.Worksheets("SkuBalance").ListObjects("tblSkuBalance"), True
    DeleteAllTableRows wb.Worksheets("LocationBalance").ListObjects("tblLocationBalance"), True

    If Not IsMissing(skuList) Then
        Set wsSku = wb.Worksheets("SkuCatalog")
        Set loSku = wsSku.ListObjects("tblSkuCatalog")
        DeleteAllTableRows loSku, True

        If IsArray(skuList) Then
            For i = LBound(skuList) To UBound(skuList)
                AppendSkuCatalogRow loSku, CStr(skuList(i))
            Next i
        ElseIf CStr(skuList) <> "" Then
            AppendSkuCatalogRow loSku, CStr(skuList)
        End If
    End If

    If persistWorkbook Then
        p = BuildUniqueTestWorkbookPath(whId & ".invSys.Data.Inventory")
        SaveWorkbookAsTestFile wb, p, 51
    End If
    Set BuildPhase2InventoryWorkbook = wb
End Function

Public Function BuildCanonicalInventoryWorkbook(ByVal whId As String, ByVal rootPath As String, Optional ByVal skuList As Variant) As Workbook
    Dim wb As Workbook
    Dim targetPath As String

    Set wb = BuildPhase2InventoryWorkbook(whId, skuList, False)
    targetPath = EnsureCanonicalFolder(rootPath) & "\" & whId & ".invSys.Data.Inventory.xlsb"
    SaveWorkbookAsCanonicalFile wb, targetPath, 50
    Set BuildCanonicalInventoryWorkbook = wb
End Function

Private Sub AppendSkuCatalogRow(ByVal loSku As ListObject, ByVal skuValue As String)
    Dim r As ListRow

    If loSku Is Nothing Then Exit Sub
    skuValue = Trim$(skuValue)
    If skuValue = "" Then Exit Sub

    EnsureTableSheetEditable loSku, "tblSkuCatalog"
    Set r = loSku.ListRows.Add
    SetTableRowValue loSku, r.Index, "SKU", skuValue
    SetTableRowValue loSku, r.Index, "ITEM_CODE", skuValue
    SetTableRowValue loSku, r.Index, "ITEM", skuValue
End Sub

Public Function BuildCanonicalConfigWorkbook(ByVal whId As String, _
                                             ByVal stId As String, _
                                             ByVal rootPath As String, _
                                             Optional ByVal roleDefault As String = "RECEIVE", _
                                             Optional ByVal processorServiceUserId As String = "svc_processor") As Workbook
    Dim wb As Workbook
    Dim targetPath As String

    Set wb = BuildPhase2ConfigWorkbook(whId, stId, roleDefault, processorServiceUserId)
    targetPath = EnsureCanonicalFolder(rootPath) & "\" & whId & ".invSys.Config.xlsb"
    SaveWorkbookAsCanonicalFile wb, targetPath, 50
    Set BuildCanonicalConfigWorkbook = wb
End Function

Public Function BuildCanonicalAuthWorkbook(ByVal whId As String, _
                                           ByVal rootPath As String, _
                                           Optional ByVal processorServiceUserId As String = "svc_processor") As Workbook
    Dim wb As Workbook
    Dim targetPath As String

    Set wb = BuildPhase2AuthWorkbook(whId, processorServiceUserId)
    targetPath = EnsureCanonicalFolder(rootPath) & "\" & whId & ".invSys.Auth.xlsb"
    SaveWorkbookAsCanonicalFile wb, targetPath, 50
    Set BuildCanonicalAuthWorkbook = wb
End Function

Public Function BuildPhase2InboxWorkbook(Optional ByVal stationId As String = "S1") As Workbook
    Set BuildPhase2InboxWorkbook = BuildReceiveInboxWorkbook(stationId)
End Function

Public Function BuildReceiveInboxWorkbook(Optional ByVal stationId As String = "S1", _
                                          Optional ByVal persistWorkbook As Boolean = True) As Workbook
    Dim wb As Workbook
    Dim report As String
    Dim p As String

    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Name = "InboxReceive"
    Call modProcessor.EnsureReceiveInboxSchema(wb, report)
    DeleteAllTableRows wb.Worksheets("InboxReceive").ListObjects("tblInboxReceive"), False

    If persistWorkbook Then
        p = BuildUniqueTestWorkbookPath("invSys.Inbox.Receiving." & stationId)
        SaveWorkbookAsTestFile wb, p, 51
    End If
    Set BuildReceiveInboxWorkbook = wb
End Function

Public Function BuildCanonicalReceiveInboxWorkbook(ByVal stationId As String, ByVal rootPath As String) As Workbook
    Dim wb As Workbook
    Dim targetPath As String

    Set wb = BuildReceiveInboxWorkbook(stationId, False)
    targetPath = EnsureCanonicalFolder(rootPath) & "\invSys.Inbox.Receiving." & stationId & ".xlsb"
    SaveWorkbookAsCanonicalFile wb, targetPath, 50
    Set BuildCanonicalReceiveInboxWorkbook = wb
End Function

Public Function BuildShipInboxWorkbook(Optional ByVal stationId As String = "S1", _
                                       Optional ByVal persistWorkbook As Boolean = True) As Workbook
    Dim wb As Workbook
    Dim report As String
    Dim p As String

    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Name = "InboxShip"
    Call modProcessor.EnsureShipInboxSchema(wb, report)
    DeleteAllTableRows wb.Worksheets("InboxShip").ListObjects("tblInboxShip"), False

    If persistWorkbook Then
        p = BuildUniqueTestWorkbookPath("invSys.Inbox.Shipping." & stationId)
        SaveWorkbookAsTestFile wb, p, 51
    End If
    Set BuildShipInboxWorkbook = wb
End Function

Public Function BuildCanonicalShipInboxWorkbook(ByVal stationId As String, ByVal rootPath As String) As Workbook
    Dim wb As Workbook
    Dim targetPath As String

    Set wb = BuildShipInboxWorkbook(stationId, False)
    targetPath = EnsureCanonicalFolder(rootPath) & "\invSys.Inbox.Shipping." & stationId & ".xlsb"
    SaveWorkbookAsCanonicalFile wb, targetPath, 50
    Set BuildCanonicalShipInboxWorkbook = wb
End Function

Public Function BuildProductionInboxWorkbook(Optional ByVal stationId As String = "S1", _
                                             Optional ByVal persistWorkbook As Boolean = True) As Workbook
    Dim wb As Workbook
    Dim report As String
    Dim p As String

    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Name = "InboxProd"
    Call modProcessor.EnsureProductionInboxSchema(wb, report)
    DeleteAllTableRows wb.Worksheets("InboxProd").ListObjects("tblInboxProd"), False

    If persistWorkbook Then
        p = BuildUniqueTestWorkbookPath("invSys.Inbox.Production." & stationId)
        SaveWorkbookAsTestFile wb, p, 51
    End If
    Set BuildProductionInboxWorkbook = wb
End Function

Public Function BuildCanonicalProductionInboxWorkbook(ByVal stationId As String, ByVal rootPath As String) As Workbook
    Dim wb As Workbook
    Dim targetPath As String

    Set wb = BuildProductionInboxWorkbook(stationId, False)
    targetPath = EnsureCanonicalFolder(rootPath) & "\invSys.Inbox.Production." & stationId & ".xlsb"
    SaveWorkbookAsCanonicalFile wb, targetPath, 50
    Set BuildCanonicalProductionInboxWorkbook = wb
End Function

Public Sub AddCapability(ByVal wb As Workbook, _
                         ByVal userId As String, _
                         ByVal capability As String, _
                         ByVal whId As String, _
                         ByVal stId As String, _
                         ByVal status As String, _
                         Optional ByVal validFrom As String = "", _
                         Optional ByVal validTo As String = "")
    Dim lo As ListObject
    Dim r As ListRow

    EnsureUserExists wb, userId
    Set lo = wb.Worksheets("Capabilities").ListObjects("tblCapabilities")
    EnsureTableSheetEditable lo, "tblCapabilities"
    Set r = lo.ListRows.Add
    r.Range.Cells(1, lo.ListColumns("UserId").Index).Value = userId
    r.Range.Cells(1, lo.ListColumns("Capability").Index).Value = capability
    r.Range.Cells(1, lo.ListColumns("WarehouseId").Index).Value = whId
    r.Range.Cells(1, lo.ListColumns("StationId").Index).Value = stId
    r.Range.Cells(1, lo.ListColumns("Status").Index).Value = status
    r.Range.Cells(1, lo.ListColumns("ValidFrom").Index).Value = validFrom
    r.Range.Cells(1, lo.ListColumns("ValidTo").Index).Value = validTo
End Sub

Public Sub SetUserPinHash(ByVal wb As Workbook, _
                          ByVal userId As String, _
                          ByVal pinHash As String)
    Dim lo As ListObject
    Dim i As Long

    EnsureUserExists wb, userId
    Set lo = wb.Worksheets("Users").ListObjects("tblUsers")
    EnsureTableSheetEditable lo, "tblUsers"
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 Then
            lo.DataBodyRange.Cells(i, lo.ListColumns("PinHash").Index).Value = pinHash
            Exit Sub
        End If
    Next i
End Sub

Public Sub AddInboxReceiveRow(ByVal wb As Workbook, _
                              ByVal eventId As String, _
                              ByVal createdAtUtc As Variant, _
                              ByVal whId As String, _
                              ByVal stId As String, _
                              ByVal userId As String, _
                              ByVal sku As String, _
                              ByVal qty As Double, _
                              Optional ByVal locationVal As String = "", _
                              Optional ByVal noteVal As String = "")
    Dim lo As ListObject
    Dim r As ListRow

    Set lo = wb.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    EnsureTableSheetEditable lo, "tblInboxReceive"
    Set r = lo.ListRows.Add
    SetInboxRowCommon lo, r.Index, eventId, EVENT_TYPE_RECEIVE, createdAtUtc, whId, stId, userId, noteVal
    SetTableRowValue lo, r.Index, "SKU", sku
    SetTableRowValue lo, r.Index, "Qty", qty
    SetTableRowValue lo, r.Index, "Location", locationVal
End Sub

Public Sub AddInboxShipRow(ByVal wb As Workbook, _
                           ByVal eventId As String, _
                           ByVal createdAtUtc As Variant, _
                           ByVal whId As String, _
                           ByVal stId As String, _
                           ByVal userId As String, _
                           ByVal payloadJson As String, _
                           Optional ByVal noteVal As String = "")
    Dim lo As ListObject
    Dim r As ListRow

    Set lo = wb.Worksheets("InboxShip").ListObjects("tblInboxShip")
    EnsureTableSheetEditable lo, "tblInboxShip"
    Set r = lo.ListRows.Add
    SetInboxRowCommon lo, r.Index, eventId, EVENT_TYPE_SHIP, createdAtUtc, whId, stId, userId, noteVal
    SetTableRowValue lo, r.Index, "PayloadJson", payloadJson
End Sub

Public Sub AddInboxProductionRow(ByVal wb As Workbook, _
                                 ByVal eventId As String, _
                                 ByVal eventType As String, _
                                 ByVal createdAtUtc As Variant, _
                                 ByVal whId As String, _
                                 ByVal stId As String, _
                                 ByVal userId As String, _
                                 ByVal payloadJson As String, _
                                 Optional ByVal noteVal As String = "")
    Dim lo As ListObject
    Dim r As ListRow

    Set lo = wb.Worksheets("InboxProd").ListObjects("tblInboxProd")
    EnsureTableSheetEditable lo, "tblInboxProd"
    Set r = lo.ListRows.Add
    SetInboxRowCommon lo, r.Index, eventId, eventType, createdAtUtc, whId, stId, userId, noteVal
    SetTableRowValue lo, r.Index, "PayloadJson", payloadJson
End Sub

Public Function CreateReceiveEvent(ByVal eventId As String, _
                                   ByVal whId As String, _
                                   ByVal stId As String, _
                                   ByVal userId As String, _
                                   ByVal sku As String, _
                                   ByVal qty As Double, _
                                   Optional ByVal locationVal As String = "", _
                                   Optional ByVal noteVal As String = "", _
                                   Optional ByVal createdAtUtc As Variant = Empty, _
                                   Optional ByVal sourceInbox As String = "test-inbox") As Object
    Dim evt As Object

    Set evt = CreateBaseEvent(eventId, EVENT_TYPE_RECEIVE, whId, stId, userId, createdAtUtc, sourceInbox)
    evt("SKU") = sku
    evt("Qty") = qty
    evt("Location") = locationVal
    evt("Note") = noteVal
    Set CreateReceiveEvent = evt
End Function

Public Function CreatePayloadEvent(ByVal eventId As String, _
                                   ByVal eventType As String, _
                                   ByVal whId As String, _
                                   ByVal stId As String, _
                                   ByVal userId As String, _
                                   ByVal payloadJson As String, _
                                   Optional ByVal noteVal As String = "", _
                                   Optional ByVal createdAtUtc As Variant = Empty, _
                                   Optional ByVal sourceInbox As String = "test-inbox") As Object
    Dim evt As Object

    Set evt = CreateBaseEvent(eventId, eventType, whId, stId, userId, createdAtUtc, sourceInbox)
    evt("PayloadJson") = payloadJson
    evt("Note") = noteVal
    Set CreatePayloadEvent = evt
End Function

Public Function CreatePayloadItem(ByVal rowVal As Long, _
                                  ByVal sku As String, _
                                  ByVal qty As Double, _
                                  Optional ByVal locationVal As String = "", _
                                  Optional ByVal noteVal As String = "", _
                                  Optional ByVal ioType As String = "") As Object
    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("Row") = rowVal
    item("SKU") = sku
    item("Qty") = qty
    If locationVal <> "" Then item("Location") = locationVal
    If noteVal <> "" Then item("Note") = noteVal
    If ioType <> "" Then item("IoType") = ioType
    Set CreatePayloadItem = item
End Function

Public Function BuildPayloadJson(ParamArray items() As Variant) As String
    Dim i As Long
    Dim item As Object

    BuildPayloadJson = "["
    For i = LBound(items) To UBound(items)
        Set item = items(i)
        If i > LBound(items) Then BuildPayloadJson = BuildPayloadJson & ","
        BuildPayloadJson = BuildPayloadJson & DictionaryToJson(item)
    Next i
    BuildPayloadJson = BuildPayloadJson & "]"
End Function

Public Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        If Not ws.ListObjects(tableName) Is Nothing Then
            TableExists = True
            Exit Function
        End If
    Next ws
    On Error GoTo 0
End Function

Public Function GetRowValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = lo.ListColumns(columnName).Index
    GetRowValue = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Public Sub CloseNoSave(ByVal wb As Workbook)
    Dim p As String
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    p = wb.FullName
    wb.Close SaveChanges:=False
    If InStr(1, p, ".test.", vbTextCompare) > 0 Then
        If Len(Dir$(p)) > 0 Then Kill p
    End If
    On Error GoTo 0
End Sub

Public Sub CloseAndDeleteWorkbook(ByVal wb As Workbook)
    Dim p As String

    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    p = wb.FullName
    wb.Close SaveChanges:=False
    If Len(p) > 0 Then
        If Len(Dir$(p)) > 0 Then Kill p
    End If
    On Error GoTo 0
End Sub

Private Function CreateBaseEvent(ByVal eventId As String, _
                                 ByVal eventType As String, _
                                 ByVal whId As String, _
                                 ByVal stId As String, _
                                 ByVal userId As String, _
                                 ByVal createdAtUtc As Variant, _
                                 ByVal sourceInbox As String) As Object
    Dim evt As Object

    Set evt = CreateObject("Scripting.Dictionary")
    evt.CompareMode = vbTextCompare
    evt("EventID") = eventId
    evt("EventType") = eventType
    evt("CreatedAtUTC") = IIf(IsEmpty(createdAtUtc), Now, createdAtUtc)
    evt("WarehouseId") = whId
    evt("StationId") = stId
    evt("UserId") = userId
    evt("SourceInbox") = sourceInbox
    Set CreateBaseEvent = evt
End Function

Private Sub SetInboxRowCommon(ByVal lo As ListObject, _
                              ByVal rowIndex As Long, _
                              ByVal eventId As String, _
                              ByVal eventType As String, _
                              ByVal createdAtUtc As Variant, _
                              ByVal whId As String, _
                              ByVal stId As String, _
                              ByVal userId As String, _
                              ByVal noteVal As String)
    SetTableRowValue lo, rowIndex, "EventID", eventId
    SetTableRowValue lo, rowIndex, "EventType", eventType
    SetTableRowValue lo, rowIndex, "CreatedAtUTC", createdAtUtc
    SetTableRowValue lo, rowIndex, "WarehouseId", whId
    SetTableRowValue lo, rowIndex, "StationId", stId
    SetTableRowValue lo, rowIndex, "UserId", userId
    SetTableRowValue lo, rowIndex, "Note", noteVal
    SetTableRowValue lo, rowIndex, "Status", "NEW"
    SetTableRowValue lo, rowIndex, "RetryCount", 0
End Sub

Private Function DictionaryToJson(ByVal d As Object) As String
    Dim key As Variant
    Dim parts As Collection
    Dim part As Variant

    Set parts = New Collection
    For Each key In d.Keys
        parts.Add """" & EscapeJson(CStr(key)) & """:" & JsonValue(d(key))
    Next key

    DictionaryToJson = "{"
    For Each part In parts
        If Right$(DictionaryToJson, 1) <> "{" Then DictionaryToJson = DictionaryToJson & ","
        DictionaryToJson = DictionaryToJson & CStr(part)
    Next part
    DictionaryToJson = DictionaryToJson & "}"
End Function

Private Function JsonValue(ByVal valueIn As Variant) As String
    Select Case VarType(valueIn)
        Case vbBoolean
            JsonValue = LCase$(CStr(valueIn))
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            JsonValue = Replace$(CStr(valueIn), ",", "")
        Case Else
            JsonValue = """" & EscapeJson(CStr(valueIn)) & """"
    End Select
End Function

Private Function EscapeJson(ByVal textIn As String) As String
    textIn = Replace$(textIn, "\", "\\")
    textIn = Replace$(textIn, Chr$(34), "\" & Chr$(34))
    textIn = Replace$(textIn, vbCrLf, "\n")
    textIn = Replace$(textIn, vbCr, "\n")
    textIn = Replace$(textIn, vbLf, "\n")
    EscapeJson = textIn
End Function

Private Sub SaveWorkbookAsTestFile(ByVal wb As Workbook, ByVal pathOut As String, ByVal fileFormat As Long)
    EnsureCanonicalFolder Left$(pathOut, InStrRev(pathOut, "\") - 1)
    CloseWorkbookByFullName pathOut
    On Error Resume Next
    Kill pathOut
    On Error GoTo 0
    wb.SaveAs Filename:=pathOut, FileFormat:=fileFormat
End Sub

Private Sub SaveWorkbookAsCanonicalFile(ByVal wb As Workbook, ByVal pathOut As String, ByVal fileFormat As Long)
    EnsureCanonicalFolder Left$(pathOut, InStrRev(pathOut, "\") - 1)
    CloseWorkbookByFullName pathOut
    On Error Resume Next
    Kill pathOut
    On Error GoTo 0
    wb.SaveAs Filename:=pathOut, FileFormat:=fileFormat
End Sub

Private Function BuildUniqueTestWorkbookPath(ByVal baseName As String) As String
    BuildUniqueTestWorkbookPath = Environ$("TEMP") & "\" & baseName & "." & BuildSafeTestSuffix() & ".test.xlsx"
End Function

Public Function BuildUniqueTestFolder(ByVal baseName As String) As String
    Dim tempRoot As String

    tempRoot = Environ$("TEMP")
    If Right$(tempRoot, 1) = "\" Then
        BuildUniqueTestFolder = tempRoot & baseName & "_" & BuildSafeTestSuffix()
    Else
        BuildUniqueTestFolder = tempRoot & "\" & baseName & "_" & BuildSafeTestSuffix()
    End If
    If Len(Dir$(BuildUniqueTestFolder, vbDirectory)) = 0 Then MkDir BuildUniqueTestFolder
End Function

Private Function EnsureCanonicalFolder(ByVal folderPath As String) As String
    EnsureCanonicalFolder = folderPath
    If Trim$(folderPath) = "" Then Exit Function
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Function
    CreateFolderRecursive folderPath
End Function

Private Sub CloseWorkbookByFullName(ByVal fullNameIn As String)
    Dim wb As Workbook

    If Trim$(fullNameIn) = "" Then Exit Sub
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullNameIn, vbTextCompare) = 0 Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
            Exit For
        End If
    Next wb
End Sub

Private Sub SetTableRowValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueOut
End Sub

Private Sub DeleteAllTableRows(ByVal lo As ListObject, ByVal reprotectAfter As Boolean)
    On Error Resume Next
    lo.Parent.Unprotect
    On Error GoTo 0
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
    If reprotectAfter Then
        On Error Resume Next
        lo.Parent.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

Private Sub EnsureTableSheetEditable(ByVal lo As ListObject, ByVal tableName As String)
    If lo Is Nothing Then Exit Sub
    If Not lo.Parent.ProtectContents Then Exit Sub

    On Error Resume Next
    lo.Parent.Unprotect
    On Error GoTo 0

    If lo.Parent.ProtectContents Then
        Err.Raise vbObjectError + 2601, "TestPhase2Helpers.EnsureTableSheetEditable", _
                  "Worksheet '" & lo.Parent.Name & "' is protected and could not be unprotected before writing to " & tableName & "."
    End If
End Sub

Private Sub EnsureUserExists(ByVal wb As Workbook, ByVal userId As String)
    Dim lo As ListObject
    Dim r As ListRow
    Dim i As Long

    If wb Is Nothing Then Exit Sub
    userId = Trim$(userId)
    If userId = "" Then Exit Sub

    Set lo = wb.Worksheets("Users").ListObjects("tblUsers")
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            If StrComp(CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 Then Exit Sub
        Next i
    End If

    EnsureTableSheetEditable lo, "tblUsers"
    Set r = lo.ListRows.Add
    r.Range.Cells(1, lo.ListColumns("UserId").Index).Value = userId
    r.Range.Cells(1, lo.ListColumns("DisplayName").Index).Value = userId
    r.Range.Cells(1, lo.ListColumns("PinHash").Index).Value = ""
    r.Range.Cells(1, lo.ListColumns("Status").Index).Value = "Active"
    r.Range.Cells(1, lo.ListColumns("ValidFrom").Index).Value = ""
    r.Range.Cells(1, lo.ListColumns("ValidTo").Index).Value = ""
End Sub

Private Function BuildSafeTestSuffix() As String
    Dim token As String

    token = Format$(Now, "yyyymmdd_hhnnss") & "_" & Right$(SanitizePathToken(CreateGuidToken()), 8)
    BuildSafeTestSuffix = SanitizePathToken(token)
End Function

Private Function CreateGuidToken() As String
    On Error Resume Next
    CreateGuidToken = CreateObject("Scriptlet.TypeLib").GUID
    On Error GoTo 0
    If CreateGuidToken = "" Then CreateGuidToken = Format$(Timer * 1000, "00000000")
End Function

Private Function SanitizePathToken(ByVal valueIn As String) As String
    Dim i As Long
    Dim ch As String

    valueIn = Replace$(valueIn, Chr$(0), "")
    For i = 1 To Len(valueIn)
        ch = Mid$(valueIn, i, 1)
        If ch Like "[A-Za-z0-9_]" Then SanitizePathToken = SanitizePathToken & ch
    Next i
End Function

Private Sub CreateFolderRecursive(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 0 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then CreateFolderRecursive parentPath
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub
