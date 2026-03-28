Attribute VB_Name = "modWarehouseSync"
Option Explicit

Private Const SHEET_OUTBOX As String = "OutboxEvents"
Private Const TABLE_OUTBOX As String = "tblOutboxEvents"

Private Const SHEET_SNAPSHOT As String = "InventorySnapshot"
Private Const TABLE_SNAPSHOT As String = "tblInventorySnapshot"
Private Const SNAPSHOT_SOURCE_LOG As String = "LOG_FALLBACK"
Private Const SNAPSHOT_SOURCE_PROJECTION As String = "PROJECTION"
Private Const SNAPSHOT_SOURCE_MANAGED_SURFACE As String = "MANAGED_SURFACE"

Public Function AppendEventToOutbox(ByVal evt As Object, _
                                    Optional ByVal inventoryWb As Workbook = Nothing, _
                                    Optional ByVal outboxWb As Workbook = Nothing, _
                                    Optional ByVal runId As String = "", _
                                    Optional ByRef report As String = "") As Boolean
    On Error GoTo FailAppend

    Dim warehouseId As String
    Dim wbOutbox As Workbook
    Dim loOutbox As ListObject
    Dim appliedMeta As Object
    Dim eventId As String
    Dim rowIndex As Long
    Dim r As ListRow

    warehouseId = GetEventStringSync(evt, "WarehouseId")
    eventId = GetEventStringSync(evt, "EventID")
    If eventId = "" Then
        report = "Outbox write requires EventID."
        Exit Function
    End If

    Set wbOutbox = ResolveOutboxWorkbook(warehouseId, outboxWb, True)
    If wbOutbox Is Nothing Then
        report = "Outbox workbook not found."
        Exit Function
    End If
    If Not EnsureOutboxSchema(wbOutbox, report) Then Exit Function

    Set loOutbox = wbOutbox.Worksheets(SHEET_OUTBOX).ListObjects(TABLE_OUTBOX)
    Set appliedMeta = ResolveAppliedMeta(eventId, inventoryWb)
    If appliedMeta Is Nothing Then
        report = "Applied metadata not found for EventID " & eventId
        Exit Function
    End If

    rowIndex = FindRowByValueSync(loOutbox, "EventID", eventId)
    If rowIndex = 0 Then
        EnsureTableSheetEditableSync loOutbox, TABLE_OUTBOX
        Set r = loOutbox.ListRows.Add
        rowIndex = r.Index
    End If

    SetTableRowValueSync loOutbox, rowIndex, "EventID", eventId
    SetTableRowValueSync loOutbox, rowIndex, "UndoOfEventId", GetEventStringSync(evt, "UndoOfEventId")
    SetTableRowValueSync loOutbox, rowIndex, "EventType", GetEventStringSync(evt, "EventType")
    SetTableRowValueSync loOutbox, rowIndex, "WarehouseId", warehouseId
    SetTableRowValueSync loOutbox, rowIndex, "StationId", GetEventStringSync(evt, "StationId")
    SetTableRowValueSync loOutbox, rowIndex, "OccurredAtUTC", GetEventValueSync(evt, "CreatedAtUTC")
    SetTableRowValueSync loOutbox, rowIndex, "AppliedAtUTC", appliedMeta("AppliedAtUTC")
    SetTableRowValueSync loOutbox, rowIndex, "AppliedByUserId", GetEventStringSync(evt, "UserId")
    SetTableRowValueSync loOutbox, rowIndex, "RunId", ResolveStringSync(appliedMeta, "RunId", runId)
    SetTableRowValueSync loOutbox, rowIndex, "DeltaJson", BuildDeltaJsonForOutbox(evt)
    SaveWorkbookSync wbOutbox

    report = "OK"
    AppendEventToOutbox = True
    Exit Function

FailAppend:
    report = "AppendEventToOutbox failed: " & Err.Description
End Function

Public Function EnsureOutboxSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                   Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim i As Long

    If targetWb Is Nothing Then
        Set wb = ResolveOutboxWorkbook(modConfig.GetWarehouseId(), Nothing, True)
    Else
        Set wb = targetWb
    End If
    If wb Is Nothing Then
        report = "Outbox workbook not resolved."
        Exit Function
    End If

    headers = Array("EventID", "UndoOfEventId", "EventType", "WarehouseId", "StationId", "OccurredAtUTC", _
                    "AppliedAtUTC", "AppliedByUserId", "RunId", "DeltaJson")

    NormalizeWorkbookSheetsSync wb, Array(SHEET_OUTBOX)
    Set ws = EnsureWorksheetSync(wb, SHEET_OUTBOX)
    EnsureWorksheetEditableSync ws
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_OUTBOX)
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCellSync(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers))), , xlYes)
        lo.Name = TABLE_OUTBOX
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnSync lo, CStr(headers(i))
    Next i
    RemoveBlankSeedRowSync lo

    report = "OK"
    EnsureOutboxSchema = True
    Exit Function

FailEnsure:
    report = "EnsureOutboxSchema failed: " & Err.Description
End Function

Public Function GenerateWarehouseSnapshot(Optional ByVal warehouseId As String = "", _
                                          Optional ByVal inventoryWb As Workbook = Nothing, _
                                          Optional ByVal outputPath As String = "", _
                                          Optional ByVal snapshotWb As Workbook = Nothing, _
                                          Optional ByRef report As String = "") As Boolean
    On Error GoTo FailSnapshot

    Dim wbInv As Workbook
    Dim wbSnap As Workbook
    Dim snapshotRows As Object
    Dim savePath As String

    If warehouseId = "" Then warehouseId = modConfig.GetWarehouseId()
    Set wbInv = ResolveInventoryWorkbookBridge(warehouseId, inventoryWb)
    If wbInv Is Nothing Then
        report = "Inventory workbook not found."
        Exit Function
    End If

    Set snapshotRows = BuildSnapshotRowsSync(wbInv, warehouseId, report)
    If snapshotRows Is Nothing Then
        If report = "" Then report = "Snapshot rows could not be built."
        Exit Function
    End If

    Set wbSnap = ResolveSnapshotWorkbook(warehouseId, outputPath, snapshotWb, True)
    If wbSnap Is Nothing Then
        report = "Snapshot workbook not resolved."
        Exit Function
    End If
    savePath = wbSnap.FullName
    If Not EnsureSnapshotSchema(wbSnap, report) Then Exit Function
    WriteSnapshotRows wbSnap, warehouseId, snapshotRows
    wbSnap.Save

    report = savePath
    GenerateWarehouseSnapshot = True
    Exit Function

FailSnapshot:
    report = "GenerateWarehouseSnapshot failed: " & Err.Description
End Function

Public Function ResolveOutboxWorkbook(Optional ByVal warehouseId As String = "", _
                                      Optional ByVal targetWb As Workbook = Nothing, _
                                      Optional ByVal createIfMissing As Boolean = False) As Workbook
    Dim targetPath As String

    If Not targetWb Is Nothing Then
        Set ResolveOutboxWorkbook = targetWb
        Exit Function
    End If

    targetPath = ResolveOutboxPath(warehouseId)
    Set ResolveOutboxWorkbook = ResolveWorkbookByPathSync(targetPath, createIfMissing, False)
End Function

Public Function ResolveSnapshotWorkbook(Optional ByVal warehouseId As String = "", _
                                        Optional ByVal outputPath As String = "", _
                                        Optional ByVal targetWb As Workbook = Nothing, _
                                        Optional ByVal createIfMissing As Boolean = False) As Workbook
    Dim targetPath As String

    If Not targetWb Is Nothing Then
        Set ResolveSnapshotWorkbook = targetWb
        Exit Function
    End If

    targetPath = ResolveSnapshotPath(warehouseId, outputPath)
    Set ResolveSnapshotWorkbook = ResolveWorkbookByPathSync(targetPath, createIfMissing, Not createIfMissing)
End Function

Private Function EnsureSnapshotSchema(ByVal wb As Workbook, ByRef report As String) As Boolean
    On Error GoTo FailEnsure

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim i As Long

    headers = Array("WarehouseId", "SKU", "ITEM", "UOM", "LOCATION", "DESCRIPTION", "VENDOR(s)", "VENDOR_CODE", "CATEGORY", _
                    "QtyOnHand", "QtyAvailable", "LocationSummary", "LastAppliedAtUTC")
    NormalizeWorkbookSheetsSync wb, Array(SHEET_SNAPSHOT)
    Set ws = EnsureWorksheetSync(wb, SHEET_SNAPSHOT)
    EnsureWorksheetEditableSync ws

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_SNAPSHOT)
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCellSync(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers))), , xlYes)
        lo.Name = TABLE_SNAPSHOT
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnSync lo, CStr(headers(i))
    Next i
    RemoveBlankSeedRowSync lo
    ApplySnapshotColumnFormatsSync lo

    report = "OK"
    EnsureSnapshotSchema = True
    Exit Function

FailEnsure:
    report = "EnsureSnapshotSchema failed: " & Err.Description
End Function

Private Sub WriteSnapshotRows(ByVal wb As Workbook, _
                              ByVal warehouseId As String, _
                              ByVal snapshotRows As Object)
    Dim lo As ListObject
    Dim key As Variant
    Dim rowIndex As Long
    Dim entry As Object

    Set lo = wb.Worksheets(SHEET_SNAPSHOT).ListObjects(TABLE_SNAPSHOT)
    DeleteAllRowsSync lo

    If snapshotRows Is Nothing Or snapshotRows.Count = 0 Then
        EnsureTableSheetEditableSync lo, TABLE_SNAPSHOT
        lo.ListRows.Add
        SetTableRowValueSync lo, 1, "WarehouseId", warehouseId
        SetTableRowValueSync lo, 1, "SKU", ""
        SetTableRowValueSync lo, 1, "ITEM", ""
        SetTableRowValueSync lo, 1, "UOM", ""
        SetTableRowValueSync lo, 1, "LOCATION", ""
        SetTableRowValueSync lo, 1, "DESCRIPTION", ""
        SetTableRowValueSync lo, 1, "VENDOR(s)", ""
        SetTableRowValueSync lo, 1, "VENDOR_CODE", ""
        SetTableRowValueSync lo, 1, "CATEGORY", ""
        SetTableRowValueSync lo, 1, "QtyOnHand", 0
        SetTableRowValueSync lo, 1, "QtyAvailable", 0
        SetTableRowValueSync lo, 1, "LocationSummary", vbNullString
        SetTableRowValueSync lo, 1, "LastAppliedAtUTC", vbNullString
        Exit Sub
    End If

    For Each key In snapshotRows.Keys
        EnsureTableSheetEditableSync lo, TABLE_SNAPSHOT
        lo.ListRows.Add
        rowIndex = lo.ListRows.Count
        Set entry = snapshotRows(key)
        SetTableRowValueSync lo, rowIndex, "WarehouseId", ResolveStringSync(entry, "WarehouseId", warehouseId)
        SetTableRowValueSync lo, rowIndex, "SKU", ResolveStringSync(entry, "SKU", CStr(key))
        SetTableRowValueSync lo, rowIndex, "ITEM", ResolveStringSync(entry, "ITEM", ResolveStringSync(entry, "ItemName", ResolveStringSync(entry, "SKU", CStr(key))))
        SetTableRowValueSync lo, rowIndex, "UOM", ResolveStringSync(entry, "UOM", "")
        SetTableRowValueSync lo, rowIndex, "LOCATION", ResolveStringSync(entry, "LOCATION", "")
        SetTableRowValueSync lo, rowIndex, "DESCRIPTION", ResolveStringSync(entry, "DESCRIPTION", "")
        SetTableRowValueSync lo, rowIndex, "VENDOR(s)", ResolveStringSync(entry, "VENDOR(s)", ResolveStringSync(entry, "VENDORS", ""))
        SetTableRowValueSync lo, rowIndex, "VENDOR_CODE", ResolveStringSync(entry, "VENDOR_CODE", "")
        SetTableRowValueSync lo, rowIndex, "CATEGORY", ResolveStringSync(entry, "CATEGORY", "")
        SetTableRowValueSync lo, rowIndex, "QtyOnHand", ResolveNumberSync(entry, "QtyOnHand")
        SetTableRowValueSync lo, rowIndex, "QtyAvailable", ResolveNumberSync(entry, "QtyAvailable")
        SetTableRowValueSync lo, rowIndex, "LocationSummary", ResolveStringSync(entry, "LocationSummary", "")
        If entry.Exists("LastAppliedAtUTC") Then SetTableRowValueSync lo, rowIndex, "LastAppliedAtUTC", entry("LastAppliedAtUTC")
    Next key
End Sub

Private Function BuildSnapshotRowsSync(ByVal wbInv As Workbook, _
                                       ByVal warehouseId As String, _
                                       ByRef report As String) As Object
    Dim snapshotRows As Object

    Set snapshotRows = BuildSnapshotRowsFromProjectionsSync(wbInv, warehouseId)
    If Not snapshotRows Is Nothing Then
        AppendCatalogRowsSync snapshotRows, wbInv, warehouseId
        report = SNAPSHOT_SOURCE_PROJECTION
        Set BuildSnapshotRowsSync = snapshotRows
        Exit Function
    End If

    Set snapshotRows = BuildSnapshotRowsFromLogSync(wbInv, warehouseId)
    If Not snapshotRows Is Nothing Then
        AppendCatalogRowsSync snapshotRows, wbInv, warehouseId
        report = SNAPSHOT_SOURCE_LOG
        Set BuildSnapshotRowsSync = snapshotRows
        Exit Function
    End If

    Set snapshotRows = BuildSnapshotRowsFromManagedSurfaceSync(wbInv, warehouseId)
    If Not snapshotRows Is Nothing Then
        AppendCatalogRowsSync snapshotRows, wbInv, warehouseId
        report = SNAPSHOT_SOURCE_MANAGED_SURFACE
        Set BuildSnapshotRowsSync = snapshotRows
        Exit Function
    End If

    report = "Inventory snapshot source tables not found."
End Function

Private Function BuildSnapshotRowsFromProjectionsSync(ByVal wbInv As Workbook, ByVal warehouseId As String) As Object
    Dim loSku As ListObject
    Dim loLoc As ListObject
    Dim rows As Object
    Dim rowIndex As Long
    Dim sku As String
    Dim entry As Object
    Dim qtyOnHand As Double

    Set loSku = FindListObjectByNameSync(wbInv, "tblSkuBalance")
    Set loLoc = FindListObjectByNameSync(wbInv, "tblLocationBalance")
    If loSku Is Nothing Then Exit Function

    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbTextCompare

    If Not loSku.DataBodyRange Is Nothing Then
        For rowIndex = 1 To loSku.ListRows.Count
            sku = SafeTrimSync(GetCellByColumnSync(loSku, rowIndex, "SKU"))
            If sku = "" Then GoTo ContinueSkuLoop

            Set entry = EnsureSnapshotEntrySync(rows, sku, warehouseId)
            qtyOnHand = NzDblSync(GetCellByColumnSync(loSku, rowIndex, "QtyOnHand"))
            entry("QtyOnHand") = qtyOnHand
            entry("QtyAvailable") = qtyOnHand
            If IsDate(GetCellByColumnSync(loSku, rowIndex, "LastAppliedUTC")) Then
                entry("LastAppliedAtUTC") = CDate(GetCellByColumnSync(loSku, rowIndex, "LastAppliedUTC"))
            End If
ContinueSkuLoop:
        Next rowIndex
    End If

    If Not loLoc Is Nothing Then AppendLocationSummariesSync rows, loLoc, warehouseId
    Set BuildSnapshotRowsFromProjectionsSync = rows
End Function

Private Function BuildSnapshotRowsFromLogSync(ByVal wbInv As Workbook, ByVal warehouseId As String) As Object
    Dim loLog As ListObject
    Dim rows As Object
    Dim rowIndex As Long
    Dim sku As String
    Dim locationVal As String
    Dim qty As Double
    Dim entry As Object
    Dim rowDate As Variant

    Set loLog = FindListObjectByNameSync(wbInv, "tblInventoryLog")
    If loLog Is Nothing Then Exit Function

    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbTextCompare

    If Not loLog.DataBodyRange Is Nothing Then
        For rowIndex = 1 To loLog.ListRows.Count
            sku = SafeTrimSync(GetCellByColumnSync(loLog, rowIndex, "SKU"))
            If sku = "" Then GoTo ContinueLogLoop

            Set entry = EnsureSnapshotEntrySync(rows, sku, warehouseId)
            qty = 0
            If IsNumeric(GetCellByColumnSync(loLog, rowIndex, "QtyDelta")) Then qty = CDbl(GetCellByColumnSync(loLog, rowIndex, "QtyDelta"))
            entry("QtyOnHand") = ResolveNumberSync(entry, "QtyOnHand") + qty
            entry("QtyAvailable") = ResolveNumberSync(entry, "QtyOnHand")

            rowDate = GetCellByColumnSync(loLog, rowIndex, "AppliedAtUTC")
            If IsDate(rowDate) Then
                If (Not entry.Exists("LastAppliedAtUTC")) Or CDate(rowDate) > CDate(entry("LastAppliedAtUTC")) Then
                    entry("LastAppliedAtUTC") = CDate(rowDate)
                End If
            End If

            locationVal = SafeTrimSync(GetCellByColumnSync(loLog, rowIndex, "Location"))
            If locationVal <> "" Then AppendLocationFragmentSync entry, locationVal, qty
ContinueLogLoop:
        Next rowIndex
    End If

    Set BuildSnapshotRowsFromLogSync = rows
End Function

Private Function BuildSnapshotRowsFromManagedSurfaceSync(ByVal wbInv As Workbook, ByVal warehouseId As String) As Object
    Dim loInv As ListObject
    Dim rows As Object
    Dim rowIndex As Long
    Dim sku As String
    Dim entry As Object
    Dim qtyOnHand As Double
    Dim qtyAvailable As Double
    Dim locationSummary As String
    Dim lastApplied As Variant

    Set loInv = FindListObjectByNameSync(wbInv, "invSys")
    If loInv Is Nothing Then Exit Function
    If loInv.DataBodyRange Is Nothing Then Exit Function

    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbTextCompare

    For rowIndex = 1 To loInv.ListRows.Count
        sku = ResolveCatalogCellTextSync(loInv, rowIndex, "ITEM_CODE")
        If sku = "" Then sku = ResolveCatalogCellTextSync(loInv, rowIndex, "SKU")
        If sku = "" Then GoTo ContinueManagedLoop

        Set entry = EnsureSnapshotEntrySync(rows, sku, warehouseId)
        qtyOnHand = NzDblSync(GetCellByColumnSync(loInv, rowIndex, "TOTAL INV"))
        qtyAvailable = qtyOnHand
        If GetColumnIndexSync(loInv, "QtyAvailable") > 0 Then
            qtyAvailable = NzDblSync(GetCellByColumnSync(loInv, rowIndex, "QtyAvailable"))
        End If
        locationSummary = ResolveCatalogCellTextSync(loInv, rowIndex, "LocationSummary")
        If locationSummary = "" Then locationSummary = ResolveCatalogCellTextSync(loInv, rowIndex, "LOCATION")

        entry("QtyOnHand") = qtyOnHand
        entry("QtyAvailable") = qtyAvailable
        entry("LocationSummary") = NormalizeManagedLocationSummarySync(locationSummary, qtyOnHand)
        ApplyManagedSurfaceMetadataSync entry, loInv, rowIndex

        lastApplied = GetCellByColumnSync(loInv, rowIndex, "LAST EDITED")
        If Not IsDate(lastApplied) Then lastApplied = GetCellByColumnSync(loInv, rowIndex, "TOTAL INV LAST EDIT")
        If IsDate(lastApplied) Then entry("LastAppliedAtUTC") = CDate(lastApplied)
ContinueManagedLoop:
    Next rowIndex

    Set BuildSnapshotRowsFromManagedSurfaceSync = rows
End Function

Private Sub ApplyManagedSurfaceMetadataSync(ByVal entry As Object, ByVal loInv As ListObject, ByVal rowIndex As Long)
    ApplyCatalogValueIfPresentSync entry, "ITEM", ResolveCatalogCellTextSync(loInv, rowIndex, "ITEM")
    ApplyCatalogValueIfPresentSync entry, "UOM", ResolveCatalogCellTextSync(loInv, rowIndex, "UOM")
    ApplyCatalogValueIfPresentSync entry, "LOCATION", ResolveCatalogCellTextSync(loInv, rowIndex, "LOCATION")
    ApplyCatalogValueIfPresentSync entry, "DESCRIPTION", ResolveCatalogCellTextSync(loInv, rowIndex, "DESCRIPTION")
    ApplyCatalogValueIfPresentSync entry, "VENDOR(s)", ResolveCatalogCellTextSync(loInv, rowIndex, "VENDOR(s)")
    ApplyCatalogValueIfPresentSync entry, "VENDOR_CODE", ResolveCatalogCellTextSync(loInv, rowIndex, "VENDOR_CODE")
    ApplyCatalogValueIfPresentSync entry, "CATEGORY", ResolveCatalogCellTextSync(loInv, rowIndex, "CATEGORY")
End Sub

Private Function NormalizeManagedLocationSummarySync(ByVal locationSummary As String, ByVal qtyOnHand As Double) As String
    locationSummary = SafeTrimSync(locationSummary)
    If locationSummary <> "" Then
        NormalizeManagedLocationSummarySync = locationSummary
        Exit Function
    End If
    If qtyOnHand = 0 Then Exit Function
    NormalizeManagedLocationSummarySync = "(blank)=" & FormatQuantitySync(qtyOnHand)
End Function

Private Sub AppendLocationSummariesSync(ByVal snapshotRows As Object, _
                                        ByVal loLoc As ListObject, _
                                        ByVal warehouseId As String)
    Dim rowIndex As Long
    Dim sku As String
    Dim locationVal As String
    Dim qtyOnHand As Double
    Dim entry As Object
    Dim rowDate As Variant

    If snapshotRows Is Nothing Or loLoc Is Nothing Then Exit Sub
    If loLoc.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To loLoc.ListRows.Count
        sku = SafeTrimSync(GetCellByColumnSync(loLoc, rowIndex, "SKU"))
        If sku = "" Then GoTo ContinueLocLoop

        locationVal = SafeTrimSync(GetCellByColumnSync(loLoc, rowIndex, "Location"))
        qtyOnHand = NzDblSync(GetCellByColumnSync(loLoc, rowIndex, "QtyOnHand"))
        Set entry = EnsureSnapshotEntrySync(snapshotRows, sku, warehouseId)
        AppendLocationFragmentSync entry, locationVal, qtyOnHand

        rowDate = GetCellByColumnSync(loLoc, rowIndex, "LastAppliedUTC")
        If IsDate(rowDate) Then
            If (Not entry.Exists("LastAppliedAtUTC")) Or CDate(rowDate) > CDate(entry("LastAppliedAtUTC")) Then
                entry("LastAppliedAtUTC") = CDate(rowDate)
            End If
        End If
ContinueLocLoop:
    Next rowIndex
End Sub

Private Function EnsureSnapshotEntrySync(ByVal rows As Object, _
                                         ByVal sku As String, _
                                         ByVal warehouseId As String) As Object
    Dim entry As Object
    Dim locationTotals As Object

    If rows.Exists(sku) Then
        Set EnsureSnapshotEntrySync = rows(sku)
        Exit Function
    End If

    Set entry = CreateObject("Scripting.Dictionary")
    entry.CompareMode = vbTextCompare
    entry("WarehouseId") = warehouseId
    entry("SKU") = sku
    entry("QtyOnHand") = 0#
    entry("QtyAvailable") = 0#
    entry("LocationSummary") = vbNullString
    Set locationTotals = CreateObject("Scripting.Dictionary")
    locationTotals.CompareMode = vbTextCompare
    entry.Add "LocationTotals", locationTotals
    rows.Add sku, entry
    Set EnsureSnapshotEntrySync = entry
End Function

Private Sub AppendCatalogRowsSync(ByVal snapshotRows As Object, ByVal wbInv As Workbook, ByVal warehouseId As String)
    Dim loCatalog As ListObject

    If snapshotRows Is Nothing Then Exit Sub
    If wbInv Is Nothing Then Exit Sub

    Set loCatalog = FindListObjectByNameSync(wbInv, "invSys")
    ApplyCatalogTableToSnapshotRowsSync snapshotRows, loCatalog, warehouseId

    Set loCatalog = FindListObjectByNameSync(wbInv, "tblItemSearchIndex")
    ApplyCatalogTableToSnapshotRowsSync snapshotRows, loCatalog, warehouseId

    Set loCatalog = FindListObjectByNameSync(wbInv, "tblSkuCatalog")
    ApplyCatalogTableToSnapshotRowsSync snapshotRows, loCatalog, warehouseId
End Sub

Private Sub ApplyCatalogTableToSnapshotRowsSync(ByVal snapshotRows As Object, _
                                                ByVal loCatalog As ListObject, _
                                                ByVal warehouseId As String)
    Dim rowIndex As Long
    Dim sku As String
    Dim entry As Object
    Dim itemValue As String
    Dim uomValue As String
    Dim locationValue As String
    Dim descriptionValue As String
    Dim vendorValue As String
    Dim vendorCodeValue As String
    Dim categoryValue As String

    If snapshotRows Is Nothing Then Exit Sub
    If loCatalog Is Nothing Then Exit Sub
    If loCatalog.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To loCatalog.ListRows.Count
        sku = ResolveCatalogCellTextSync(loCatalog, rowIndex, "SKU")
        If sku = "" Then sku = ResolveCatalogCellTextSync(loCatalog, rowIndex, "ITEM_CODE")
        If sku = "" Then GoTo ContinueLoop

        Set entry = EnsureSnapshotEntrySync(snapshotRows, sku, warehouseId)
        itemValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "ITEM")
        If itemValue = "" Then itemValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "ItemName")
        If itemValue = "" Then itemValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "NAME")
        If itemValue = "" Then itemValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "SKU")
        If itemValue = "" Then itemValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "ITEM_CODE")

        uomValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "UOM")
        If uomValue = "" Then uomValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "UNITOFMEASURE")
        If uomValue = "" Then uomValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "UNITOFMEASUREMENT")
        If uomValue = "" Then uomValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "UNIT")

        locationValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "LOCATION")
        If locationValue = "" Then locationValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "DEFAULTLOCATION")
        If locationValue = "" Then locationValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "PRIMARYLOCATION")

        descriptionValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "DESCRIPTION")
        If descriptionValue = "" Then descriptionValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "DESC")

        vendorValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "VENDOR(s)")
        If vendorValue = "" Then vendorValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "VENDORS")
        If vendorValue = "" Then vendorValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "VENDOR")

        vendorCodeValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "VENDOR_CODE")
        If vendorCodeValue = "" Then vendorCodeValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "VENDORCODE")

        categoryValue = ResolveCatalogCellTextSync(loCatalog, rowIndex, "CATEGORY")

        ApplyCatalogValueIfPresentSync entry, "ITEM", itemValue
        ApplyCatalogValueIfPresentSync entry, "UOM", uomValue
        ApplyCatalogValueIfPresentSync entry, "LOCATION", locationValue
        ApplyCatalogValueIfPresentSync entry, "DESCRIPTION", descriptionValue
        ApplyCatalogValueIfPresentSync entry, "VENDOR(s)", vendorValue
        ApplyCatalogValueIfPresentSync entry, "VENDOR_CODE", vendorCodeValue
        ApplyCatalogValueIfPresentSync entry, "CATEGORY", categoryValue
ContinueLoop:
    Next rowIndex
End Sub

Private Sub ApplyCatalogValueIfPresentSync(ByVal entry As Object, ByVal key As String, ByVal valueIn As String)
    If entry Is Nothing Then Exit Sub
    valueIn = SafeTrimSync(valueIn)
    If valueIn = "" Then Exit Sub
    entry(key) = valueIn
End Sub

Private Function ResolveCatalogCellTextSync(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim idx As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idx = GetColumnIndexSync(lo, columnName)
    If idx = 0 Then Exit Function
    ResolveCatalogCellTextSync = SafeTrimSync(lo.DataBodyRange.Cells(rowIndex, idx).Value)
End Function

Private Sub AppendLocationFragmentSync(ByVal entry As Object, ByVal locationVal As String, ByVal qtyOnHand As Double)
    Dim totals As Object
    Dim label As String

    If entry Is Nothing Then Exit Sub

    label = NormalizeLocationLabelForSummarySync(locationVal)
    Set totals = EnsureLocationTotalsSync(entry)
    If totals.Exists(label) Then
        totals(label) = CDbl(totals(label)) + qtyOnHand
    Else
        totals.Add label, qtyOnHand
    End If

    entry("LocationSummary") = BuildLocationSummarySync(totals)
End Sub

Private Function EnsureLocationTotalsSync(ByVal entry As Object) As Object
    If entry Is Nothing Then Exit Function

    On Error Resume Next
    If entry.Exists("LocationTotals") Then
        If IsObject(entry("LocationTotals")) Then
            Set EnsureLocationTotalsSync = entry("LocationTotals")
            Exit Function
        End If
    End If
    On Error GoTo 0

    Set EnsureLocationTotalsSync = CreateObject("Scripting.Dictionary")
    EnsureLocationTotalsSync.CompareMode = vbTextCompare
    entry.Add "LocationTotals", EnsureLocationTotalsSync
End Function

Private Function BuildLocationSummarySync(ByVal totals As Object) As String
    Dim key As Variant
    Dim fragment As String

    If totals Is Nothing Then Exit Function
    For Each key In totals.Keys
        fragment = CStr(key) & "=" & FormatQuantitySync(CDbl(totals(key)))
        If BuildLocationSummarySync = "" Then
            BuildLocationSummarySync = fragment
        Else
            BuildLocationSummarySync = BuildLocationSummarySync & "; " & fragment
        End If
    Next key
End Function

Private Function NormalizeLocationLabelForSummarySync(ByVal locationVal As String) As String
    Dim label As String
    Dim eqPos As Long
    Dim suffixText As String

    label = Trim$(locationVal)
    If label = "" Then
        NormalizeLocationLabelForSummarySync = "(blank)"
        Exit Function
    End If

    eqPos = InStrRev(label, "=")
    If eqPos > 1 Then
        suffixText = Trim$(Mid$(label, eqPos + 1))
        suffixText = Replace$(suffixText, ",", "")
        If suffixText <> "" Then
            If IsNumeric(suffixText) Then label = Trim$(Left$(label, eqPos - 1))
        End If
    End If

    If label = "" Then label = "(blank)"
    NormalizeLocationLabelForSummarySync = label
End Function

Private Function FormatQuantitySync(ByVal qtyIn As Double) As String
    If Abs(qtyIn - CLng(qtyIn)) < 0.0000001 Then
        FormatQuantitySync = CStr(CLng(qtyIn))
    Else
        FormatQuantitySync = Trim$(Format$(qtyIn, "0.########"))
    End If
End Function

Private Function NzDblSync(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzDblSync = CDbl(valueIn)
End Function

Private Function ResolveNumberSync(ByVal dict As Object, ByVal keyName As String) As Double
    If dict Is Nothing Then Exit Function
    If Not dict.Exists(keyName) Then Exit Function
    ResolveNumberSync = NzDblSync(dict(keyName))
End Function

Private Function ResolveAppliedMeta(ByVal eventId As String, ByVal inventoryWb As Workbook) As Object
    Dim wb As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim meta As Object

    Set wb = ResolveInventoryWorkbookBridge("", inventoryWb)
    If wb Is Nothing Then Exit Function

    Set lo = FindListObjectByNameSync(wb, "tblAppliedEvents")
    If lo Is Nothing Then Exit Function
    rowIndex = FindRowByValueSync(lo, "EventID", eventId)
    If rowIndex = 0 Then Exit Function

    Set meta = CreateObject("Scripting.Dictionary")
    meta.CompareMode = vbTextCompare
    meta("AppliedAtUTC") = GetCellByColumnSync(lo, rowIndex, "AppliedAtUTC")
    meta("RunId") = GetCellByColumnSync(lo, rowIndex, "RunId")
    meta("Status") = GetCellByColumnSync(lo, rowIndex, "Status")
    meta("SourceInbox") = GetCellByColumnSync(lo, rowIndex, "SourceInbox")
    Set ResolveAppliedMeta = meta
End Function

Private Function BuildDeltaJsonForOutbox(ByVal evt As Object) As String
    Dim payloadJson As String
    Dim items As Collection
    Dim item As Object

    payloadJson = GetEventStringSync(evt, "PayloadJson")
    If payloadJson <> "" Then
        BuildDeltaJsonForOutbox = payloadJson
        Exit Function
    End If

    Set items = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("SKU") = GetEventStringSync(evt, "SKU")
    item("QtyDelta") = GetEventValueSync(evt, "Qty")
    item("Location") = GetEventStringSync(evt, "Location")
    item("Note") = GetEventStringSync(evt, "Note")
    items.Add item
    BuildDeltaJsonForOutbox = modRoleEventWriter.BuildPayloadJsonFromCollection(items)
End Function

Private Function ResolveOutboxPath(ByVal warehouseId As String) As String
    Dim rootPath As String
    If warehouseId = "" Then warehouseId = modConfig.GetWarehouseId()
    rootPath = Trim$(GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = modConfig.GetString("PathDataRoot", Environ$("TEMP"))
    ResolveOutboxPath = NormalizeFolderPathSync(rootPath) & warehouseId & ".Outbox.Events.xlsb"
End Function

Private Function ResolveSnapshotPath(ByVal warehouseId As String, ByVal outputPath As String) As String
    Dim rootPath As String
    If Trim$(outputPath) <> "" Then
        ResolveSnapshotPath = outputPath
        Exit Function
    End If
    If warehouseId = "" Then warehouseId = modConfig.GetWarehouseId()
    rootPath = Trim$(GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = modConfig.GetString("PathDataRoot", Environ$("TEMP"))
    ResolveSnapshotPath = NormalizeFolderPathSync(rootPath) & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
End Function

Private Function ResolveWorkbookByPathSync(ByVal targetPath As String, _
                                           ByVal createIfMissing As Boolean, _
                                           Optional ByVal openReadOnly As Boolean = False) As Workbook
    On Error GoTo FailOpen

    Dim wb As Workbook
    Dim fileExists As Boolean
    Dim prevEvents As Boolean
    Dim eventsSuppressed As Boolean
    Dim prevAlerts As Boolean
    Dim alertsSuppressed As Boolean

    If Trim$(targetPath) = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            If (Not openReadOnly) And wb.ReadOnly Then Exit Function
            Set ResolveWorkbookByPathSync = wb
            Exit Function
        End If
    Next wb

    fileExists = FileExistsSync(targetPath)
    If fileExists Or (IsUncPathSync(targetPath) And Not createIfMissing) Then
        prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        alertsSuppressed = True
        On Error Resume Next
        Set ResolveWorkbookByPathSync = Application.Workbooks.Open( _
            Filename:=targetPath, _
            UpdateLinks:=0, _
            ReadOnly:=openReadOnly, _
            IgnoreReadOnlyRecommended:=True, _
            Notify:=False, _
            AddToMru:=False)
        If Err.Number <> 0 Then
            Err.Clear
            Set ResolveWorkbookByPathSync = Nothing
        End If
        On Error GoTo FailOpen
        Application.DisplayAlerts = prevAlerts
        alertsSuppressed = False
        Exit Function
    End If

    If Not createIfMissing Then Exit Function

    EnsureFolderForFileSync targetPath
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    eventsSuppressed = True
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    Application.EnableEvents = prevEvents
    eventsSuppressed = False
    Set ResolveWorkbookByPathSync = wb
    Exit Function

FailOpen:
    On Error Resume Next
    If eventsSuppressed Then Application.EnableEvents = prevEvents
    If alertsSuppressed Then Application.DisplayAlerts = prevAlerts
    On Error GoTo 0
End Function

Private Function EnsureWorksheetSync(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheetSync = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheetSync Is Nothing Then
        Set EnsureWorksheetSync = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheetSync.Name = sheetName
    End If
End Function

Private Sub NormalizeWorkbookSheetsSync(ByVal wb As Workbook, ByVal wantedSheets As Variant)
    Dim i As Long
    Dim ws As Worksheet
    Dim prevAlerts As Boolean

    If wb Is Nothing Then Exit Sub

    For i = LBound(wantedSheets) To UBound(wantedSheets)
        EnsureWorksheetSync wb, CStr(wantedSheets(i))
    Next i

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If Not WorksheetNameInSetSync(ws.Name, wantedSheets) Then ws.Delete
    Next i
    Application.DisplayAlerts = prevAlerts
End Sub

Private Sub EnsureWorksheetEditableSync(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 3001, "modWarehouseSync.EnsureWorksheetEditableSync", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected."
    End If
End Sub

Private Sub EnsureTableSheetEditableSync(ByVal lo As ListObject, ByVal tableName As String)
    If lo Is Nothing Then Exit Sub
    EnsureWorksheetEditableSync lo.Parent
    If lo.Parent.ProtectContents Then
        Err.Raise vbObjectError + 3002, "modWarehouseSync.EnsureTableSheetEditableSync", _
                  "Worksheet '" & lo.Parent.Name & "' is protected and could not be unprotected before updating " & tableName & "."
    End If
End Sub

Private Function GetNextTableStartCellSync(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCellSync = ws.Range("A1")
    Else
        Set GetNextTableStartCellSync = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Sub EnsureListColumnSync(ByVal lo As ListObject, ByVal columnName As String)
    If GetColumnIndexSync(lo, columnName) > 0 Then Exit Sub
    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
End Sub

Private Sub RemoveBlankSeedRowSync(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If lo.ListRows.Count <> 1 Then Exit Sub
    If Not TableRowIsBlankSync(lo, 1) Then Exit Sub
    EnsureTableSheetEditableSync lo, lo.Name
    lo.ListRows(1).Delete
End Sub

Private Sub DeleteAllRowsSync(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableSheetEditableSync lo, lo.Name
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
End Sub

Private Function WorksheetNameInSetSync(ByVal sheetName As String, ByVal sheetNames As Variant) As Boolean
    Dim i As Long

    For i = LBound(sheetNames) To UBound(sheetNames)
        If StrComp(CStr(sheetNames(i)), sheetName, vbTextCompare) = 0 Then
            WorksheetNameInSetSync = True
            Exit Function
        End If
    Next i
End Function

Private Function TableRowIsBlankSync(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim colIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > lo.ListRows.Count Then Exit Function

    TableRowIsBlankSync = True
    For colIndex = 1 To lo.ListColumns.Count
        If SafeTrimSync(lo.DataBodyRange.Cells(rowIndex, colIndex).Value) <> "" Then
            TableRowIsBlankSync = False
            Exit Function
        End If
    Next colIndex
End Function

Private Function FindListObjectByNameSync(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameSync = ws.ListObjects(tableName)
        If Not FindListObjectByNameSync Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function FindRowByValueSync(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimSync(GetCellByColumnSync(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByValueSync = i
            Exit Function
        End If
    Next i
End Function

Private Sub SaveWorkbookSync(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If wb.Path = "" Then Exit Sub
    wb.Save
End Sub

Private Function GetCellByColumnSync(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexSync(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnSync = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Sub SetTableRowValueSync(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexSync(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetColumnIndexSync(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexSync = i
            Exit Function
        End If
    Next i
End Function

Private Sub ApplySnapshotColumnFormatsSync(ByVal lo As ListObject)
    Dim qtyCols As Variant
    Dim dateCols As Variant
    Dim key As Variant
    Dim idx As Long

    If lo Is Nothing Then Exit Sub

    qtyCols = Array("QtyOnHand", "QtyAvailable")
    For Each key In qtyCols
        idx = GetColumnIndexSync(lo, CStr(key))
        If idx > 0 Then lo.ListColumns(idx).Range.NumberFormat = "0.########"
    Next key

    dateCols = Array("LastAppliedAtUTC")
    For Each key In dateCols
        idx = GetColumnIndexSync(lo, CStr(key))
        If idx > 0 Then lo.ListColumns(idx).Range.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    Next key
End Sub

Private Function GetEventStringSync(ByVal evt As Object, ByVal key As String) As String
    Dim v As Variant
    v = GetEventValueSync(evt, key)
    GetEventStringSync = SafeTrimSync(v)
End Function

Private Function GetEventValueSync(ByVal evt As Object, ByVal key As String) As Variant
    On Error Resume Next
    If evt Is Nothing Then Exit Function
    GetEventValueSync = evt(key)
    On Error GoTo 0
End Function

Private Function ResolveStringSync(ByVal d As Object, ByVal key As String, ByVal fallbackValue As String) As String
    On Error Resume Next
    ResolveStringSync = SafeTrimSync(d(key))
    On Error GoTo 0
    If ResolveStringSync = "" Then ResolveStringSync = fallbackValue
End Function

Private Function SafeTrimSync(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimSync = Trim$(CStr(valueIn))
End Function

Private Function NormalizeFolderPathSync(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then
        NormalizeFolderPathSync = Environ$("TEMP") & "\"
        Exit Function
    End If
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathSync = folderPath
End Function

Private Sub EnsureFolderForFileSync(ByVal filePath As String)
    Dim folderPath As String
    Dim sepPos As Long

    sepPos = InStrRev(filePath, "\")
    If sepPos <= 0 Then Exit Sub
    folderPath = Left$(filePath, sepPos - 1)
    CreateFolderRecursiveSync folderPath
End Sub

Private Sub CreateFolderRecursiveSync(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long
    Dim fso As Object

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If FolderExistsSync(folderPath) Then Exit Sub

    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    If IsUncShareRootSync(folderPath) Then Exit Sub

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 0 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Not FolderExistsSync(parentPath) Then CreateFolderRecursiveSync parentPath
    End If

    If FolderExistsSync(folderPath) Then Exit Sub

    If IsUncPathSync(folderPath) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder folderPath
    Else
        MkDir folderPath
    End If
End Sub

Private Function FileExistsSync(ByVal fullPath As String) As Boolean
    Dim fso As Object

    fullPath = Trim$(Replace$(fullPath, "/", "\"))
    If fullPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsSync = fso.FileExists(fullPath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsSync = (Len(Dir$(fullPath, vbNormal)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function FolderExistsSync(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) = "\" And Len(folderPath) > 3 Then folderPath = Left$(folderPath, Len(folderPath) - 1)

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsSync = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExistsSync = (Len(Dir$(folderPath, vbDirectory)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function IsUncPathSync(ByVal folderPath As String) As Boolean
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    IsUncPathSync = (Left$(folderPath, 2) = "\\")
End Function

Private Function IsUncShareRootSync(ByVal folderPath As String) As Boolean
    Dim trimmedPath As String
    Dim parts() As String

    trimmedPath = Trim$(Replace$(folderPath, "/", "\"))
    If Right$(trimmedPath, 1) = "\" And Len(trimmedPath) > 3 Then trimmedPath = Left$(trimmedPath, Len(trimmedPath) - 1)
    If Left$(trimmedPath, 2) <> "\\" Then Exit Function

    trimmedPath = Mid$(trimmedPath, 3)
    If trimmedPath = "" Then Exit Function

    parts = Split(trimmedPath, "\")
    IsUncShareRootSync = (UBound(parts) = 1)
End Function
