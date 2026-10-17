Attribute VB_Name = "modAdminInventoryWorksheet"
Option Explicit

Private Const EDITOR_SHEET As String = "invSys Inventory Editor"
Private Const TABLE_PREFIX As String = "invSys_Inventory_"
Private Const TEMPLATE_ROWS As Long = 20
Private Const TABLE_GAP_ROWS As Long = 5

Private Const HDR_ACTION As String = "Action"
Private Const HDR_ITEM_CODE As String = "Item Code"
Private Const HDR_ITEM_NAME As String = "Item Name"
Private Const HDR_UOM As String = "UOM"
Private Const HDR_QTY_MODE As String = "Qty Mode"
Private Const HDR_QUANTITY As String = "Quantity"
Private Const HDR_LOCATION As String = "Default Location"
Private Const HDR_CATEGORY As String = "Category"
Private Const HDR_DESCRIPTION As String = "Description"
Private Const HDR_VENDORS As String = "Vendor(s)"
Private Const HDR_VENDOR_CODE As String = "Vendor Code"
Private Const HDR_EXTERNAL_CODE As String = "External Code"
Private Const HDR_PICTURE As String = "Picture Path/URL"
Private Const HDR_EDIT_REASON As String = "Edit Reason"
Private Const HDR_STATUS As String = "Upload Status"
Private Const HDR_RESULT As String = "Upload Result"

Private mAutomationEnabled As Boolean
Private mAutomationCatalog As Object
Private mLastAutomationReport As String

Public Function CreateInventoryWorksheetTable(ByVal wb As Workbook, _
                                              ByRef tableName As String, _
                                              ByRef report As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim headerRow As Long
    Dim tableRange As Range
    Dim idx As Long

    On Error GoTo Failed
    If Not IsUsableInventoryWorkbook(wb) Then
        report = "The captured Admin operator workbook is unavailable."
        Exit Function
    End If
    If wb.ReadOnly Then
        report = "Open the Admin operator workbook for editing before creating an inventory table."
        Exit Function
    End If
    If Trim$(wb.Path) = "" Then
        report = "Save the Admin operator workbook before creating an inventory table."
        Exit Function
    End If

    Set ws = EnsureInventoryEditorSheet(wb)
    headerRow = NextInventoryTableHeaderRow(ws)
    ws.Cells(headerRow - 2, 1).Value2 = "invSys Inventory worksheet"
    ws.Cells(headerRow - 1, 1).Value2 = _
        "Use ADD for new items and EDIT for an exact existing Item Code. Select a cell in this table before upload."

    headers = Array("Action", "Item Code", "Item Name", "UOM", "Qty Mode", "Quantity", _
                    "Default Location", "Category", "Description", "Vendor(s)", _
                    "Vendor Code", "External Code", "Picture Path/URL", "Edit Reason", _
                    "Upload Status", "Upload Result")
    For idx = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, idx + 1).Value2 = CStr(headers(idx))
    Next idx
    Set tableRange = ws.Range(ws.Cells(headerRow, 1), _
                              ws.Cells(headerRow + TEMPLATE_ROWS, UBound(headers) + 1))
    Set lo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tableName = BuildUniqueInventoryTableName(wb)
    lo.Name = tableName
    lo.TableStyle = "TableStyleMedium2"
    lo.ListColumns(HDR_ITEM_CODE).DataBodyRange.NumberFormat = "@"
    ApplyInventoryWorksheetValidation lo
    FormatInventoryWorksheet ws, lo
    wb.Save
    wb.Activate
    ws.Visible = xlSheetVisible
    ws.Activate
    lo.DataBodyRange.Cells(1, 1).Select
    report = "Created Inventory table " & tableName & ". Fill or paste rows, select a cell in the table, then choose Upload Selected Inventory Table."
    CreateInventoryWorksheetTable = True
    Exit Function

Failed:
    report = "Inventory worksheet creation failed: " & Err.Description
End Function

Public Function UploadSelectedInventoryWorksheetTable(ByVal wb As Workbook, _
                                                       ByVal warehouseId As String, _
                                                       ByVal stationId As String, _
                                                       ByVal userId As String, _
                                                       ByRef report As String, _
                                                       Optional ByVal interactiveContext As String = "") As Boolean
    Dim lo As ListObject
    Dim records As Collection
    Dim record As Object
    Dim applyReport As String
    Dim succeeded As Long
    Dim failed As Long
    Dim remaining As Long
    Dim idx As Long, values As Variant, tableName As String

    On Error GoTo Failed
    mLastAutomationReport = ""
    If wb Is Nothing Then report = "The captured Admin operator workbook is unavailable.": Exit Function
    If wb.ReadOnly Then report = "Open the Admin operator workbook for editing before uploading inventory.": Exit Function
    If interactiveContext <> "" Then
        If Not modInventoryCommands.CheckOperatorWorkbook(interactiveContext, wb.Name, report) Then Exit Function
        values = modInventoryCommands.SourceValues(interactiveContext)
        If Not IsArray(values) Then Exit Function
        If values(0) <> warehouseId Or values(1) <> stationId Or values(2) <> userId Then
            report = "The captured inventory command source differs. Reopen Add/Edit Inventory.": Exit Function
        End If
        tableName = modInventoryCommands.StagingTableName(interactiveContext)
        Set lo = FindInventoryTableByName(wb, tableName)
        If lo Is Nothing Then report = "The captured staging table is unavailable. Reopen Add/Edit Inventory.": Exit Function
    Else
        If Not mAutomationEnabled Then
            If Not modAuth.LoadAuth(warehouseId) Then report = "Provisioned authority is unavailable.": Exit Function
            If Not modAuth.HasProvisionedCapabilityForSystem("ADMIN_MAINT", userId, warehouseId, stationId) Then
                report = "The provisioned user lacks ADMIN_MAINT.": Exit Function
            End If
        End If
        If Not FindSelectedInventoryWorksheetTable(wb, lo, report) Then Exit Function
    End If
    If Not PreflightInventoryWorksheetTable(lo, warehouseId, records, report, interactiveContext) Then
        If Not SaveInventoryWorksheetWorkbook(wb, interactiveContext, report) Then GoTo Partial
        mLastAutomationReport = report
        Exit Function
    End If
    If records Is Nothing Or records.Count = 0 Then
        report = "The selected Inventory table has no pending rows to upload."
        mLastAutomationReport = report
        Exit Function
    End If

    ' Every pending row has passed preflight before the first catalog/event write.
    For Each record In records
        If CStr(record("Action")) = "ADD" And interactiveContext <> "" Then
            If CStr(record("CreateEventId")) = "" Then
                record("CreateEventId") = modInventoryCreateReceipts.NewEventId(interactiveContext, report)
                If CStr(record("CreateEventId")) = "" Then GoTo Partial
            End If
        End If
        SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_ITEM_CODE, CStr(record("ItemCode")), interactiveContext
        SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_STATUS, "READY", interactiveContext
        If CStr(record("CreateEventId")) <> "" Then
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_RESULT, _
                "Inventory upload prepared; application is unconfirmed." & vbCrLf & "Event: " & CStr(record("CreateEventId")), interactiveContext
        Else
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_RESULT, "Validated; awaiting upload.", interactiveContext
        End If
    Next record
    ' The exact event identifier and generated Item Code must be independently
    ' durable in the existing result boundary before any create can be queued.
    If interactiveContext <> "" Then
        If Not SaveInventoryWorksheetWorkbook(wb, interactiveContext, report) Then GoTo Partial
    End If

    For idx = 1 To records.Count
        Set record = records(idx)
        applyReport = ""
        If ApplyInventoryWorksheetRecord(record, warehouseId, stationId, userId, applyReport, interactiveContext) Then
            succeeded = succeeded + 1
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_STATUS, "IMPORTED", interactiveContext
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_RESULT, applyReport, interactiveContext
        Else
            failed = failed + 1
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_STATUS, "FAILED", interactiveContext
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_RESULT, applyReport, interactiveContext
            Exit For
        End If
    Next idx
    If failed > 0 Then
        For idx = succeeded + failed + 1 To records.Count
            Set record = records(idx)
            remaining = remaining + 1
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_STATUS, "NOT IMPORTED", interactiveContext
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_RESULT, _
                "A prior row failed; correct it and upload again.", interactiveContext
        Next idx
    End If

    If Not SaveInventoryWorksheetWorkbook(wb, interactiveContext, report) Then GoTo Partial
    report = "Uploaded " & CStr(succeeded) & " Inventory row(s) from " & lo.Name & "."
    If failed > 0 Then report = report & " One row failed and " & CStr(remaining) & " later row(s) were not imported."
    report = report & " Refresh Inventory Viewer and role pickers to see applied changes."
    mLastAutomationReport = report & "|" & InventoryWorksheetEvidenceForTest(lo)
    UploadSelectedInventoryWorksheetTable = (failed = 0 And succeeded > 0)
    Exit Function

Partial:
    If succeeded > 0 Then report = "Partial result: inventory application was saved; worksheet result is unconfirmed." & vbCrLf & applyReport & vbCrLf & report
    mLastAutomationReport = report
    Exit Function
Failed:
    report = "Inventory worksheet upload failed: " & Err.Description
    If succeeded > 0 Then report = "Partial result: inventory application was saved; worksheet result is unconfirmed." & vbCrLf & applyReport & vbCrLf & report
    If succeeded = 0 And applyReport <> "" Then report = applyReport & vbCrLf & report
    mLastAutomationReport = report
End Function

Public Function PreflightInventoryWorksheetTable(ByVal lo As ListObject, _
                                                 ByVal warehouseId As String, _
                                                 ByRef records As Collection, _
                                                 ByRef report As String, _
                                                 Optional ByVal interactiveContext As String = "") As Boolean
    Dim pending As New Collection
    Dim record As Object
    Dim rowError As String
    Dim errors As String
    Dim rowIndex As Long
    Dim baseOrdinal As Long
    Dim addOrdinal As Long
    Dim generatedCodes As Object

    On Error GoTo Failed
    ' Validate every pending row before any authoritative write is attempted.
    If lo Is Nothing Then report = "The selected Inventory table is unavailable.": Exit Function
    If Not ValidateInventoryWorksheetHeaders(lo, report) Then Exit Function
    Set generatedCodes = CreateObject("Scripting.Dictionary")
    generatedCodes.CompareMode = vbTextCompare
    If mAutomationEnabled Then
        baseOrdinal = 1
    Else
        baseOrdinal = modAdmin.NextInventoryCodeOrdinalForWorksheet(warehouseId)
    End If
    If baseOrdinal <= 0 Then baseOrdinal = 1

    For rowIndex = 1 To lo.ListRows.Count
        If InventoryWorksheetRowHasBusinessData(lo, rowIndex) Then
            If UCase$(InventoryWorksheetText(lo, rowIndex, HDR_STATUS)) <> "IMPORTED" Then
                rowError = ""
                Set record = Nothing
                If ValidateInventoryWorksheetRow(lo, rowIndex, warehouseId, baseOrdinal, _
                        addOrdinal, generatedCodes, record, rowError, interactiveContext) Then
                    pending.Add record
                Else
                    SetInventoryWorksheetValue lo, rowIndex, HDR_STATUS, "VALIDATION ERROR", interactiveContext
                    SetInventoryWorksheetValue lo, rowIndex, HDR_RESULT, rowError, interactiveContext
                    If errors <> "" Then errors = errors & vbCrLf
                    errors = errors & "Row " & CStr(rowIndex) & ": " & rowError
                End If
            End If
        End If
    Next rowIndex
    If errors <> "" Then
        For Each record In pending
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_STATUS, "NOT UPLOADED", interactiveContext
            SetInventoryWorksheetValue lo, CLng(record("TableRow")), HDR_RESULT, _
                "Resolve every validation error before upload.", interactiveContext
        Next record
        report = "Inventory table validation failed. No inventory or catalog changes were made." & vbCrLf & errors
        Exit Function
    End If
    If pending.Count = 0 Then
        report = "The selected Inventory table has no pending business rows."
        Exit Function
    End If
    Set records = pending
    report = "Inventory table preflight passed for " & CStr(pending.Count) & " row(s)."
    PreflightInventoryWorksheetTable = True
    Exit Function

Failed:
    report = "Inventory worksheet preflight failed: " & Err.Description
End Function

Public Function FindSelectedInventoryWorksheetTable(ByVal wb As Workbook, _
                                                    ByRef lo As ListObject, _
                                                    ByRef report As String) As Boolean
    Dim selectedRange As Range
    Dim candidate As ListObject
    Dim ws As Worksheet

    If Not IsUsableInventoryWorkbook(wb) Then
        report = "The captured Admin operator workbook is unavailable."
        Exit Function
    End If
    On Error Resume Next
    Set selectedRange = Application.Selection
    On Error GoTo 0
    If selectedRange Is Nothing Then
        report = "Select a cell in an invSys Inventory table before upload."
        Exit Function
    End If
    If Not (selectedRange.Parent.Parent Is wb) Then
        report = "The selection is not in the captured Admin operator workbook."
        Exit Function
    End If
    For Each ws In wb.Worksheets
        For Each candidate In ws.ListObjects
            If IsInventoryWorksheetTable(candidate) Then
                If Not Intersect(selectedRange, candidate.Range) Is Nothing Then
                    Set lo = candidate
                    FindSelectedInventoryWorksheetTable = True
                    Exit Function
                End If
            End If
        Next candidate
    Next ws
    report = "Select a cell in an invSys Inventory table before upload."
End Function

Public Function ValidateInventoryWorksheetHeaders(ByVal lo As ListObject, _
                                                   ByRef report As String) As Boolean
    Dim required As Variant
    Dim headerValue As Variant
    Dim lc As ListColumn
    Dim normalized As String
    Dim seen As Object

    required = Array(HDR_ACTION, HDR_ITEM_CODE, HDR_ITEM_NAME, HDR_UOM, HDR_QTY_MODE, _
                     HDR_QUANTITY, HDR_LOCATION, HDR_CATEGORY, HDR_DESCRIPTION, HDR_VENDORS, _
                     HDR_VENDOR_CODE, HDR_EXTERNAL_CODE, HDR_PICTURE, HDR_EDIT_REASON, _
                     HDR_STATUS, HDR_RESULT)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For Each lc In lo.ListColumns
        normalized = NormalizeInventoryWorksheetHeader(CStr(lc.Name))
        If normalized = Chr$(82) & Chr$(79) & Chr$(87) Or _
                normalized = "SYSTEMKEY" Or normalized = "CONDITION" Then
            report = "Inventory staging tables may not contain the prohibited header " & lc.Name & "."
            Exit Function
        End If
        If seen.Exists(normalized) Then
            report = "Duplicate normalized Inventory header: " & lc.Name & "."
            Exit Function
        End If
        seen(normalized) = True
    Next lc
    For Each headerValue In required
        If InventoryWorksheetColumnIndex(lo, CStr(headerValue)) = 0 Then
            report = "Inventory table is missing required header: " & CStr(headerValue) & "."
            Exit Function
        End If
    Next headerValue
    ValidateInventoryWorksheetHeaders = True
End Function

Public Sub BeginInventoryWorksheetAutomation(ByVal existingSku As String)
    Set mAutomationCatalog = CreateObject("Scripting.Dictionary")
    mAutomationCatalog.CompareMode = vbTextCompare
    existingSku = Trim$(existingSku)
    If existingSku <> "" Then mAutomationCatalog(existingSku) = 77
    mAutomationEnabled = True
    mLastAutomationReport = ""
End Sub

Public Sub EndInventoryWorksheetAutomation()
    mAutomationEnabled = False
    Set mAutomationCatalog = Nothing
End Sub

Public Function InventoryWorksheetAutomationEnabled() As Boolean
    InventoryWorksheetAutomationEnabled = mAutomationEnabled
End Function

Public Function LastInventoryWorksheetAutomationReport() As String
    LastInventoryWorksheetAutomationReport = mLastAutomationReport
End Function

Public Function PopulateInventoryWorksheetContractRowsForTest(ByVal wb As Workbook, _
                                                              ByVal tableName As String) As Boolean
    Dim lo As ListObject

    Set lo = FindInventoryTableByName(wb, tableName)
    If lo Is Nothing Then Exit Function
    WriteContractTestRow lo, 1, "ADD", "", "Bulk Flour", "LB", "COUNTED", 0, "CLEARVIEW", "", ""
    WriteContractTestRow lo, 2, "ADD", "", "Piped Water", "LB", "UTILITY", "", "CLEARVIEW", "", ""
    WriteContractTestRow lo, 3, "EDIT", "EXISTING-SKU", "Existing Item", "EA", "COUNTED", 10, "CLEARVIEW", "", "bulk edit"
    PopulateInventoryWorksheetContractRowsForTest = True
End Function

Public Function SelectInventoryWorksheetTableForTest(ByVal wb As Workbook, _
                                                     ByVal tableName As String) As Boolean
    Dim lo As ListObject

    Set lo = FindInventoryTableByName(wb, tableName)
    If lo Is Nothing Then Exit Function
    wb.Activate
    lo.Parent.Activate
    lo.DataBodyRange.Cells(1, 1).Select
    SelectInventoryWorksheetTableForTest = True
End Function

Public Function CountInventoryWorksheetTables(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If IsInventoryWorksheetTable(lo) Then CountInventoryWorksheetTables = CountInventoryWorksheetTables + 1
        Next lo
    Next ws
End Function

Public Function InventoryWorksheetEvidenceForTest(ByVal lo As ListObject) As String
    Dim tableCreated As Boolean
    Dim preflightPassed As Boolean
    Dim utilityReady As Boolean
    Dim exactEdit As Boolean
    Dim generatedCode As Boolean
    Dim statusesReady As Boolean
    Dim zeroCounted As Boolean

    If lo Is Nothing Then
        InventoryWorksheetEvidenceForTest = "TableCreated=False"
        Exit Function
    End If
    tableCreated = IsInventoryWorksheetTable(lo)
    generatedCode = InventoryWorksheetText(lo, 1, HDR_ITEM_CODE) <> "" And _
                    InventoryWorksheetText(lo, 2, HDR_ITEM_CODE) <> ""
    utilityReady = InStr(1, InventoryWorksheetText(lo, 2, HDR_RESULT), _
                         "ITEM_KIND=UTILITY", vbTextCompare) > 0
    exactEdit = InventoryWorksheetText(lo, 3, HDR_ITEM_CODE) = "EXISTING-SKU" And _
                UCase$(InventoryWorksheetText(lo, 3, HDR_STATUS)) = "IMPORTED"
    zeroCounted = (UCase$(InventoryWorksheetText(lo, 1, HDR_STATUS)) = "IMPORTED" And _
                   Val(InventoryWorksheetText(lo, 1, HDR_QUANTITY)) = 0)
    statusesReady = UCase$(InventoryWorksheetText(lo, 1, HDR_STATUS)) = "IMPORTED" And _
                    UCase$(InventoryWorksheetText(lo, 2, HDR_STATUS)) = "IMPORTED" And exactEdit
    preflightPassed = generatedCode And statusesReady
    InventoryWorksheetEvidenceForTest = _
        "TableCreated=" & CStr(tableCreated) & _
        "|Preflight=" & CStr(preflightPassed) & _
        "|ZeroCounted=" & CStr(zeroCounted) & _
        "|Utility=" & CStr(utilityReady) & _
        "|ExactEdit=" & CStr(exactEdit) & _
        "|GeneratedCode=" & CStr(generatedCode) & _
        "|Statuses=" & CStr(statusesReady)
End Function

Private Function ValidateInventoryWorksheetRow(ByVal lo As ListObject, _
                                               ByVal rowIndex As Long, _
                                               ByVal warehouseId As String, _
                                               ByVal baseOrdinal As Long, _
                                               ByRef addOrdinal As Long, _
                                               ByVal generatedCodes As Object, _
                                               ByRef record As Object, _
                                               ByRef rowError As String, _
                                                Optional ByVal interactiveContext As String = "") As Boolean
    Dim actionName As String
    Dim sku As String
    Dim itemName As String
    Dim uom As String
    Dim qtyMode As String
    Dim qtyText As String
    Dim qty As Double
    Dim hasQty As Boolean
    Dim editReason As String
    Dim statusText As String
    Dim codeOrdinal As Long
    Dim customFields As Object, receiptEvent As String

    actionName = UCase$(InventoryWorksheetText(lo, rowIndex, HDR_ACTION))
    sku = InventoryWorksheetText(lo, rowIndex, HDR_ITEM_CODE)
    itemName = InventoryWorksheetText(lo, rowIndex, HDR_ITEM_NAME)
    uom = UCase$(InventoryWorksheetText(lo, rowIndex, HDR_UOM))
    qtyMode = UCase$(InventoryWorksheetText(lo, rowIndex, HDR_QTY_MODE))
    qtyText = InventoryWorksheetText(lo, rowIndex, HDR_QUANTITY)
    editReason = InventoryWorksheetText(lo, rowIndex, HDR_EDIT_REASON)
    statusText = UCase$(InventoryWorksheetText(lo, rowIndex, HDR_STATUS))
    If qtyMode = "" Then qtyMode = "COUNTED"

    If actionName <> "ADD" And actionName <> "EDIT" Then
        rowError = "Action must be ADD or EDIT."
        Exit Function
    End If
    If itemName = "" Then rowError = "Item Name is required.": Exit Function
    If uom = "" Then rowError = "UOM is required.": Exit Function
    If Not IsConfiguredInventoryWorksheetUom(uom) Then
        rowError = uom & " is not in the warehouse UOM catalog."
        Exit Function
    End If
    Select Case qtyMode
        Case "COUNTED", "UTILITY", "SERVICE", "NOT COUNTED"
        Case Else
            rowError = "Qty Mode must be COUNTED, UTILITY, SERVICE, or NOT COUNTED."
            Exit Function
    End Select
    If qtyText <> "" Then
        If Not IsNumeric(qtyText) Then rowError = "Quantity must be numeric.": Exit Function
        qty = CDbl(qtyText)
        hasQty = True
    End If
    If qtyMode = "COUNTED" Then
        If actionName = "ADD" And (Not hasQty Or qty < 0) Then
            rowError = "COUNTED ADD requires an explicit nonnegative Quantity."
            Exit Function
        End If
        If actionName = "EDIT" And hasQty And qty < 0 Then
            rowError = "COUNTED EDIT Quantity cannot be negative."
            Exit Function
        End If
    Else
        qty = 0
        hasQty = False
    End If

    If actionName = "ADD" Then
        receiptEvent = modInventoryCreateReceipts.EventIdFromResult(InventoryWorksheetText(lo, rowIndex, HDR_RESULT))
        If receiptEvent <> "" And (sku = "" Or interactiveContext = "") Then
            rowError = "The prepared create requires its captured source, operator table and Item Code.": Exit Function
        End If
        If sku <> "" And receiptEvent = "" And statusText <> "FAILED" And statusText <> "NOT IMPORTED" Then
            rowError = "ADD Item Code is generated by invSys; leave it blank."
            Exit Function
        End If
        addOrdinal = addOrdinal + 1
        codeOrdinal = baseOrdinal + addOrdinal - 1
        If sku = "" Then sku = modAdmin.GenerateInventoryItemCodeForWorksheet(warehouseId, codeOrdinal)
        If receiptEvent = "" And InventoryCatalogContains(warehouseId, sku) Then
            rowError = "Generated ADD Item Code already exists: " & sku & "."
            Exit Function
        End If
    Else
        If sku = "" Then rowError = "EDIT requires the exact existing Item Code.": Exit Function
        If editReason = "" Then rowError = "EDIT requires Edit Reason.": Exit Function
        If Not InventoryCatalogContains(warehouseId, sku) Then
            rowError = "EDIT Item Code was not found in the warehouse catalog: " & sku & "."
            Exit Function
        End If
    End If
    If generatedCodes.Exists(sku) Then
        rowError = "Duplicate Item Code in pending rows: " & sku & "."
        Exit Function
    End If
    generatedCodes(sku) = True

    Set customFields = CreateObject("Scripting.Dictionary")
    customFields.CompareMode = vbTextCompare
    AppendInventoryWorksheetCustomFields lo, rowIndex, customFields
    If qtyMode <> "COUNTED" Then
        customFields("TRACK_QTY") = "FALSE"
        If qtyMode = "NOT COUNTED" Then
            customFields("ITEM_KIND") = "NON_COUNTED"
        Else
            customFields("ITEM_KIND") = qtyMode
        End If
    Else
        customFields("TRACK_QTY") = "TRUE"
        customFields("ITEM_KIND") = "INVENTORY"
    End If

    Set record = CreateObject("Scripting.Dictionary")
    record.CompareMode = vbTextCompare
    record("TableRow") = rowIndex
    record("Action") = actionName
    record("ItemCode") = sku
    record("ItemName") = itemName
    record("UOM") = uom
    record("QtyMode") = qtyMode
    record("Quantity") = qty
    record("HasQuantity") = hasQty
    record("Location") = InventoryWorksheetText(lo, rowIndex, HDR_LOCATION)
    record("Category") = InventoryWorksheetText(lo, rowIndex, HDR_CATEGORY)
    record("Description") = InventoryWorksheetText(lo, rowIndex, HDR_DESCRIPTION)
    record("Vendors") = InventoryWorksheetText(lo, rowIndex, HDR_VENDORS)
    record("VendorCode") = InventoryWorksheetText(lo, rowIndex, HDR_VENDOR_CODE)
    record("ExternalCode") = InventoryWorksheetText(lo, rowIndex, HDR_EXTERNAL_CODE)
    record("Picture") = InventoryWorksheetText(lo, rowIndex, HDR_PICTURE)
    record("EditReason") = editReason
    Set record("CustomFields") = customFields
    record("CreateEventId") = receiptEvent
    If receiptEvent <> "" Then
        If Not modAdmin.ValidateWorksheetCreateReceipt(record, interactiveContext, receiptEvent, rowError) Then Exit Function
    End If
    ValidateInventoryWorksheetRow = True
End Function

Public Function StoreQueuedCreateReceipt(ByVal context As String, ByVal row As Long, ByVal sku As String, _
        ByVal eventId As String, ByVal systemKey As String, ByRef report As String) As Boolean
    Dim book As Workbook, table As ListObject, name As String
    On Error GoTo Failed
    name = modInventoryCommands.OperatorWorkbookName(context)
    If name = "" Then Exit Function
    Set book = Application.Workbooks(name)
    Set table = FindInventoryTableByName(book, modInventoryCommands.StagingTableName(context))
    If table Is Nothing Then Exit Function
    If InventoryWorksheetText(table, row, HDR_ITEM_CODE) <> sku Then Exit Function
    SetInventoryWorksheetValue table, row, HDR_STATUS, "QUEUED", context
    SetInventoryWorksheetValue table, row, HDR_RESULT, _
        "Inventory event queued; application is unconfirmed." & vbCrLf & "Event: " & eventId & _
        vbCrLf & "System_Key: " & systemKey, context
    StoreQueuedCreateReceipt = SaveInventoryWorksheetWorkbook(book, context, report)
    Exit Function
Failed:
    report = "The event was queued, but its worksheet receipt save is unconfirmed. Event: " & eventId
End Function

Private Function ApplyInventoryWorksheetRecord(ByVal record As Object, _
                                               ByVal warehouseId As String, _
                                               ByVal stationId As String, _
                                               ByVal userId As String, _
                                               ByRef report As String, _
                                               Optional ByVal interactiveContext As String = "") As Boolean
    Dim utilitySuffix As String

    If mAutomationEnabled Then
        If UCase$(CStr(record("QtyMode"))) = "UTILITY" Then utilitySuffix = " ITEM_KIND=UTILITY"
        report = "Dry-run " & CStr(record("Action")) & " applied for " & _
                 CStr(record("ItemCode")) & "." & utilitySuffix
        ApplyInventoryWorksheetRecord = True
        Exit Function
    End If
    ApplyInventoryWorksheetRecord = modAdmin.ApplyInventoryWorksheetRecordForWarehouse( _
        warehouseId, stationId, userId, record, report, interactiveContext)
End Function

Private Function InventoryCatalogContains(ByVal warehouseId As String, _
                                          ByVal sku As String) As Boolean
    If mAutomationEnabled Then
        If Not mAutomationCatalog Is Nothing Then
            If mAutomationCatalog.Exists(sku) Then
                InventoryCatalogContains = True
            End If
        End If
    Else
        InventoryCatalogContains = modAdmin.InventoryItemCatalogContainsForWorksheet( _
            warehouseId, sku)
    End If
End Function

Private Sub AppendInventoryWorksheetCustomFields(ByVal lo As ListObject, _
                                                 ByVal rowIndex As Long, _
                                                 ByVal customFields As Object)
    Dim lc As ListColumn
    Dim headerName As String
    Dim valueText As String

    For Each lc In lo.ListColumns
        headerName = Trim$(CStr(lc.Name))
        If Not IsManagedInventoryWorksheetHeader(headerName) Then
            valueText = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, lc.Index).Value2))
            If valueText <> "" Then customFields(headerName) = valueText
        End If
    Next lc
End Sub

Private Function IsManagedInventoryWorksheetHeader(ByVal headerName As String) As Boolean
    Select Case NormalizeInventoryWorksheetHeader(headerName)
        Case "ACTION", "ITEMCODE", "ITEMNAME", "UOM", "QTYMODE", "QUANTITY", _
             "DEFAULTLOCATION", "CATEGORY", "DESCRIPTION", "VENDORS", "VENDORCODE", _
             "EXTERNALCODE", "PICTUREPATHURL", "EDITREASON", "UPLOADSTATUS", "UPLOADRESULT"
            IsManagedInventoryWorksheetHeader = True
    End Select
End Function

Private Function IsConfiguredInventoryWorksheetUom(ByVal uom As String) As Boolean
    Dim packed As String
    Dim parts As Variant
    Dim part As Variant

    packed = modUomSettings.GetConfiguredUomsPackedText()
    parts = Split(packed, "|")
    For Each part In parts
        If StrComp(Trim$(CStr(part)), uom, vbTextCompare) = 0 Then
            IsConfiguredInventoryWorksheetUom = True
            Exit Function
        End If
    Next part
End Function

Private Sub ApplyInventoryWorksheetValidation(ByVal lo As ListObject)
    Dim uomList As String

    uomList = Replace$(modUomSettings.GetConfiguredUomsPackedText(), "|", ",")
    ApplyListValidation lo.ListColumns(HDR_ACTION).DataBodyRange, "ADD,EDIT", _
        "Choose ADD or EDIT", "Select whether this row creates or edits a catalog item."
    ApplyListValidation lo.ListColumns(HDR_QTY_MODE).DataBodyRange, _
        "COUNTED,UTILITY,SERVICE,NOT COUNTED", "Choose Qty Mode", _
        "Select COUNTED, UTILITY, SERVICE, or NOT COUNTED."
    If uomList <> "" And Len(uomList) <= 255 Then
        ApplyListValidation lo.ListColumns(HDR_UOM).DataBodyRange, uomList, _
            "Choose a warehouse UOM", "Select a UOM maintained in Settings."
    End If
End Sub

Private Sub ApplyListValidation(ByVal target As Range, _
                                ByVal listText As String, _
                                ByVal errorTitle As String, _
                                ByVal errorMessage As String)
    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    With target.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=listText
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = errorTitle
        .ErrorMessage = errorMessage
    End With
End Sub

Private Function EnsureInventoryEditorSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set EnsureInventoryEditorSheet = wb.Worksheets(EDITOR_SHEET)
    On Error GoTo 0
    If EnsureInventoryEditorSheet Is Nothing Then
        Set EnsureInventoryEditorSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureInventoryEditorSheet.Name = EDITOR_SHEET
    End If
    EnsureInventoryEditorSheet.Visible = xlSheetVisible
End Function

Private Function NextInventoryTableHeaderRow(ByVal ws As Worksheet) As Long
    Dim lo As ListObject
    Dim bottomRow As Long

    NextInventoryTableHeaderRow = 4
    For Each lo In ws.ListObjects
        If IsInventoryWorksheetTable(lo) Then
            bottomRow = lo.Range.Row + lo.Range.Rows.Count - 1
            If bottomRow + TABLE_GAP_ROWS > NextInventoryTableHeaderRow Then _
                NextInventoryTableHeaderRow = bottomRow + TABLE_GAP_ROWS
        End If
    Next lo
End Function

Private Function BuildUniqueInventoryTableName(ByVal wb As Workbook) As String
    Dim ordinal As Long
    Dim candidate As String

    ordinal = 1
    Do
        candidate = TABLE_PREFIX & Format$(ordinal, "000")
        If FindInventoryTableByName(wb, candidate) Is Nothing Then Exit Do
        ordinal = ordinal + 1
    Loop
    BuildUniqueInventoryTableName = candidate
End Function

Private Function FindInventoryTableByName(ByVal wb As Workbook, _
                                          ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindInventoryTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function IsInventoryWorksheetTable(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    IsInventoryWorksheetTable = (StrComp(CStr(lo.Parent.Name), EDITOR_SHEET, vbTextCompare) = 0 And _
        StrComp(Left$(lo.Name, Len(TABLE_PREFIX)), TABLE_PREFIX, vbTextCompare) = 0)
End Function

Private Function InventoryWorksheetColumnIndex(ByVal lo As ListObject, _
                                               ByVal headerName As String) As Long
    Dim lc As ListColumn
    Dim normalized As String

    normalized = NormalizeInventoryWorksheetHeader(headerName)
    For Each lc In lo.ListColumns
        If NormalizeInventoryWorksheetHeader(CStr(lc.Name)) = normalized Then
            InventoryWorksheetColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function NormalizeInventoryWorksheetHeader(ByVal headerName As String) As String
    headerName = UCase$(Trim$(headerName))
    headerName = Replace$(headerName, " ", "")
    headerName = Replace$(headerName, "_", "")
    headerName = Replace$(headerName, "-", "")
    headerName = Replace$(headerName, "/", "")
    headerName = Replace$(headerName, "(", "")
    headerName = Replace$(headerName, ")", "")
    NormalizeInventoryWorksheetHeader = headerName
End Function

Private Function InventoryWorksheetText(ByVal lo As ListObject, _
                                        ByVal rowIndex As Long, _
                                        ByVal headerName As String) As String
    Dim colIndex As Long

    colIndex = InventoryWorksheetColumnIndex(lo, headerName)
    If colIndex = 0 Or lo.DataBodyRange Is Nothing Then Exit Function
    InventoryWorksheetText = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, colIndex).Value2))
End Function

Private Sub SetInventoryWorksheetValue(ByVal lo As ListObject, ByVal rowIndex As Long, _
        ByVal headerName As String, ByVal value As Variant, Optional ByVal interactiveContext As String = "")
    Dim colIndex As Long, report As String
    If interactiveContext <> "" Then
        If Not modInventoryCommands.PlanStageCell(interactiveContext, rowIndex, headerName, CStr(value), report) Then
            Err.Raise vbObjectError + 2141, , report
        End If
    End If
    colIndex = InventoryWorksheetColumnIndex(lo, headerName)
    If colIndex = 0 Or lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 2141, , "The captured staging field is unavailable."
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value2 = value
    If interactiveContext <> "" Then
        If Not modInventoryCommands.ConfirmStageCell(interactiveContext, report) Then Err.Raise vbObjectError + 2141, , report
    End If
End Sub

Private Function InventoryWorksheetRowHasBusinessData(ByVal lo As ListObject, _
                                                      ByVal rowIndex As Long) As Boolean
    InventoryWorksheetRowHasBusinessData = _
        InventoryWorksheetText(lo, rowIndex, HDR_ACTION) <> "" Or _
        InventoryWorksheetText(lo, rowIndex, HDR_ITEM_CODE) <> "" Or _
        InventoryWorksheetText(lo, rowIndex, HDR_ITEM_NAME) <> "" Or _
        InventoryWorksheetText(lo, rowIndex, HDR_UOM) <> "" Or _
        InventoryWorksheetText(lo, rowIndex, HDR_QUANTITY) <> ""
End Function

Private Sub FormatInventoryWorksheet(ByVal ws As Worksheet, ByVal lo As ListObject)
    lo.Range.Columns.AutoFit
    lo.ListColumns(HDR_DESCRIPTION).Range.ColumnWidth = 28
    lo.ListColumns(HDR_PICTURE).Range.ColumnWidth = 28
    lo.ListColumns(HDR_RESULT).Range.ColumnWidth = 42
    lo.Range.VerticalAlignment = xlTop
End Sub

Private Function SaveInventoryWorksheetWorkbook(ByVal wb As Workbook, _
        Optional ByVal interactiveContext As String = "", Optional ByRef report As String = "") As Boolean
    If wb Is Nothing Then Exit Function
    If interactiveContext <> "" Then
        If Not modInventoryCommands.CheckOperatorWorkbook(interactiveContext, wb.Name, report) Then Exit Function
    End If
    If Trim$(wb.Path) <> "" Then wb.Save
    If interactiveContext <> "" Then
        If Not modInventoryCommands.ConfirmStageSaved(interactiveContext, report) Then Exit Function
    End If
    SaveInventoryWorksheetWorkbook = wb.Saved
End Function

Private Function IsUsableInventoryWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo Unavailable
    If wb Is Nothing Then Exit Function
    If wb.Worksheets.Count < 1 Then Exit Function
    IsUsableInventoryWorkbook = True
Unavailable:
End Function

Private Sub WriteContractTestRow(ByVal lo As ListObject, _
                                 ByVal rowIndex As Long, _
                                 ByVal actionName As String, _
                                 ByVal sku As String, _
                                 ByVal itemName As String, _
                                 ByVal uom As String, _
                                 ByVal qtyMode As String, _
                                 ByVal qty As Variant, _
                                 ByVal locationValue As String, _
                                 ByVal categoryValue As String, _
                                 ByVal editReason As String)
    SetInventoryWorksheetValue lo, rowIndex, HDR_ACTION, actionName
    SetInventoryWorksheetValue lo, rowIndex, HDR_ITEM_CODE, sku
    SetInventoryWorksheetValue lo, rowIndex, HDR_ITEM_NAME, itemName
    SetInventoryWorksheetValue lo, rowIndex, HDR_UOM, uom
    SetInventoryWorksheetValue lo, rowIndex, HDR_QTY_MODE, qtyMode
    SetInventoryWorksheetValue lo, rowIndex, HDR_QUANTITY, qty
    SetInventoryWorksheetValue lo, rowIndex, HDR_LOCATION, locationValue
    SetInventoryWorksheetValue lo, rowIndex, HDR_CATEGORY, categoryValue
    SetInventoryWorksheetValue lo, rowIndex, HDR_EDIT_REASON, editReason
End Sub
