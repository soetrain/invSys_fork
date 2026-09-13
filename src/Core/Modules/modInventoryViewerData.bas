Attribute VB_Name = "modInventoryViewerData"
Option Explicit

Private Const SNAPSHOT_TABLE As String = "tblInventorySnapshot"
Private Const SNAPSHOT_EVENT_TABLE As String = "tblInventoryEvents"

Public Function LoadCurrentInventoryViewerData() As String
    Dim warehouseId As String
    Dim snapshotName As String
    Dim snapshotWasOpen As Boolean
    Dim snapshotWb As Workbook
    Dim snapshotTable As ListObject
    Dim quantities As Object
    Dim displayRows As Object
    Dim rowIndex As Long
    Dim itemCode As String
    Dim itemName As String
    Dim uom As String
    Dim locationValue As String
    Dim conditionValue As String
    Dim inventoryState As String
    Dim groupKey As String
    Dim qty As Double
    Dim key As Variant
    Dim fields As Variant
    Dim resultText As String
    Dim visibleCount As Long

    On Error GoTo FailLoad

    If Not modAuth.IsSignedIn() Then
        LoadCurrentInventoryViewerData = "FAIL" & vbTab & "Sign in to invSys before opening Inventory Viewer."
        Exit Function
    End If
    warehouseId = Trim$(modNasConnection.GetCurrentTargetWarehouseId())
    If warehouseId = "" Then
        LoadCurrentInventoryViewerData = "FAIL" & vbTab & "Select a warehouse before opening Inventory Viewer."
        Exit Function
    End If

    snapshotName = warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    snapshotWasOpen = ViewerWorkbookNameIsOpen(snapshotName)
    Set snapshotWb = modWarehouseSync.ResolveSnapshotWorkbook(warehouseId, "", Nothing, False)
    If snapshotWb Is Nothing Then
        LoadCurrentInventoryViewerData = "FAIL" & vbTab & _
            "No published inventory snapshot is available. Run an inventory refresh, then try again."
        Exit Function
    End If
    Set snapshotTable = ViewerFindTable(snapshotWb, SNAPSHOT_TABLE)
    If snapshotTable Is Nothing Then
        LoadCurrentInventoryViewerData = "FAIL" & vbTab & _
            "The published inventory snapshot does not contain tblInventorySnapshot."
        GoTo CleanExit
    End If

    Set quantities = CreateObject("Scripting.Dictionary")
    quantities.CompareMode = vbTextCompare
    Set displayRows = CreateObject("Scripting.Dictionary")
    displayRows.CompareMode = vbTextCompare

    If Not snapshotTable.DataBodyRange Is Nothing Then
        For rowIndex = 1 To snapshotTable.ListRows.Count
            itemCode = ViewerCellText(snapshotTable, rowIndex, "SKU")
            If itemCode = "" Then itemCode = ViewerCellText(snapshotTable, rowIndex, "ITEM_CODE")
            If itemCode = "" Then GoTo ContinueRow
            itemName = ViewerCellText(snapshotTable, rowIndex, "ITEM")
            uom = ViewerCellText(snapshotTable, rowIndex, "UOM")
            locationValue = ViewerCellText(snapshotTable, rowIndex, "LOCATION")
            conditionValue = UCase$(ViewerCellText(snapshotTable, rowIndex, "Condition"))
            inventoryState = UCase$(ViewerCellText(snapshotTable, rowIndex, "InventoryState"))
            If inventoryState = "RETIRED" Then GoTo ContinueRow
            qty = ViewerCellNumber(snapshotTable, rowIndex, "QtyAvailable")
            groupKey = itemCode & Chr$(30) & itemName & Chr$(30) & uom & _
                Chr$(30) & locationValue & Chr$(30) & conditionValue
            If quantities.Exists(groupKey) Then
                quantities(groupKey) = CDbl(quantities(groupKey)) + qty
            Else
                quantities.Add groupKey, qty
                displayRows.Add groupKey, Array(itemCode, itemName, uom, locationValue, conditionValue)
            End If
ContinueRow:
        Next rowIndex
    End If

    For Each key In quantities.Keys
        If CDbl(quantities(key)) >= 0 Then visibleCount = visibleCount + 1
    Next key
    resultText = "OK" & vbTab & ViewerEscape(warehouseId) & vbTab & _
        Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & CStr(visibleCount)
    For Each key In quantities.Keys
        If CDbl(quantities(key)) < 0 Then GoTo ContinueDisplayGroup
        fields = displayRows(key)
        resultText = resultText & vbCrLf & _
            ViewerEscape(CStr(fields(0))) & vbTab & _
            ViewerEscape(CStr(fields(1))) & vbTab & _
            ViewerEscape(CStr(fields(2))) & vbTab & _
            Format$(CDbl(quantities(key)), "0.########") & vbTab & _
            ViewerEscape(CStr(fields(3))) & vbTab & _
            ViewerEscape(CStr(fields(4)))
ContinueDisplayGroup:
    Next key
    LoadCurrentInventoryViewerData = resultText

CleanExit:
    On Error Resume Next
    If Not snapshotWb Is Nothing Then
        If Not snapshotWasOpen Then snapshotWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Exit Function

FailLoad:
    LoadCurrentInventoryViewerData = "FAIL" & vbTab & _
        "Inventory Viewer could not read the published snapshot: " & Err.Description
    Resume CleanExit
End Function

Public Function LoadCurrentInventoryEventViewerData() As String
    Dim warehouseId As String
    Dim snapshotName As String
    Dim snapshotWasOpen As Boolean
    Dim snapshotWb As Workbook
    Dim eventTable As ListObject
    Dim rowIndex As Long
    Dim eventType As String
    Dim friendlyType As String
    Dim noteText As String
    Dim referenceText As String
    Dim itemText As String
    Dim eventDateText As String
    Dim resultText As String
    Dim visibleCount As Long

    On Error GoTo FailLoad
    If Not modAuth.IsSignedIn() Then
        LoadCurrentInventoryEventViewerData = "FAIL" & vbTab & "Sign in to invSys before opening Inventory Viewer."
        Exit Function
    End If
    warehouseId = Trim$(modNasConnection.GetCurrentTargetWarehouseId())
    If warehouseId = "" Then
        LoadCurrentInventoryEventViewerData = "FAIL" & vbTab & "Select a warehouse before opening Inventory Viewer."
        Exit Function
    End If

    snapshotName = warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    snapshotWasOpen = ViewerWorkbookNameIsOpen(snapshotName)
    Set snapshotWb = modWarehouseSync.ResolveSnapshotWorkbook(warehouseId, "", Nothing, False)
    If snapshotWb Is Nothing Then
        LoadCurrentInventoryEventViewerData = "FAIL" & vbTab & "No published inventory snapshot is available."
        Exit Function
    End If
    Set eventTable = ViewerFindTable(snapshotWb, SNAPSHOT_EVENT_TABLE)
    If eventTable Is Nothing Then
        LoadCurrentInventoryEventViewerData = "FAIL" & vbTab & _
            "The published snapshot does not yet contain the Events projection. Refresh warehouse inventory, then try again."
        GoTo CleanExit
    End If

    resultText = "OK" & vbTab & ViewerEscape(warehouseId) & vbTab & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "0" & _
        vbTab & "DETAIL1" & vbTab & modTrainingWire.UtcTimestamp()
    If Not eventTable.DataBodyRange Is Nothing Then
        For rowIndex = eventTable.ListRows.Count To 1 Step -1
            eventType = UCase$(ViewerCellText(eventTable, rowIndex, "EventType"))
            friendlyType = ViewerFriendlyEventType(eventType)
            If friendlyType <> "" Then
                noteText = ViewerCellText(eventTable, rowIndex, "Note")
                referenceText = ViewerFirstNoteToken(noteText, Array("Reference", "Ref", "PO", "BOL", "ReceiptId", "DispositionRef"))
                itemText = ViewerFirstNoteToken(noteText, Array("Item", "Box", "Package"))
                If itemText = "" Then itemText = ViewerCellText(eventTable, rowIndex, "SKU")
                eventDateText = ViewerFirstNonBlank(ViewerCellText(eventTable, rowIndex, "AppliedAtUTC"), ViewerCellText(eventTable, rowIndex, "OccurredAtUTC"))
                If IsNumeric(eventDateText) Then
                    eventDateText = Format$(CDate(CDbl(eventDateText)), "yyyy-mm-dd hh:nn:ss")
                ElseIf IsDate(eventDateText) Then
                    eventDateText = Format$(CDate(eventDateText), "yyyy-mm-dd hh:nn:ss")
                End If
                visibleCount = visibleCount + 1
                resultText = resultText & vbCrLf & _
                    ViewerEscape(eventDateText) & vbTab & _
                    ViewerEscape(friendlyType) & vbTab & ViewerEscape(referenceText) & vbTab & ViewerEscape(itemText) & vbTab & _
                    ViewerEscape(ViewerCellText(eventTable, rowIndex, "QtyDelta")) & vbTab & _
                    ViewerEscape(ViewerFirstNonBlank(ViewerCellText(eventTable, rowIndex, "UOM"), ViewerFirstNoteToken(noteText, Array("UOM")))) & vbTab & _
                    ViewerEscape(ViewerCellText(eventTable, rowIndex, "Location")) & vbTab & _
                    ViewerEscape(ViewerCellText(eventTable, rowIndex, "Condition")) & vbTab & _
                    ViewerEscape(ViewerCellText(eventTable, rowIndex, "UserId")) & vbTab & "Published inventory event. Select for detail." & vbTab & _
                    ViewerEscape(ViewerRawCell(eventTable, rowIndex, "EventID")) & vbTab & ViewerEscape(ViewerRawCell(eventTable, rowIndex, "System_Key")) & vbTab & _
                    ViewerEscape(ViewerRawCell(eventTable, rowIndex, "EventType")) & vbTab & ViewerEscape(ViewerRawCell(eventTable, rowIndex, "SKU")) & vbTab & _
                    ViewerEscape(ViewerRawCell(eventTable, rowIndex, "StationId")) & vbTab & ViewerEscape(ViewerRecordedTime(eventTable, rowIndex, "OccurredAtUTC")) & vbTab & _
                    ViewerEscape(ViewerRecordedTime(eventTable, rowIndex, "AppliedAtUTC")) & vbTab & "Inventory"
            End If
        Next rowIndex
    End If
    resultText = Replace(resultText, vbTab & "0" & vbTab & "DETAIL1", vbTab & CStr(visibleCount) & vbTab & "DETAIL1", 1, 1)
    LoadCurrentInventoryEventViewerData = resultText

CleanExit:
    On Error Resume Next
    If Not snapshotWb Is Nothing Then
        If Not snapshotWasOpen Then snapshotWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Exit Function

FailLoad:
    LoadCurrentInventoryEventViewerData = "FAIL" & vbTab & _
        "Inventory Viewer could not read the published Events projection: " & Err.Description
    Resume CleanExit
End Function

Private Function ViewerRawCell(ByVal table As ListObject, ByVal rowIndex As Long, ByVal headerName As String) As String
    Dim columnIndex As Long, value As Variant
    columnIndex = ViewerColumnIndex(table, headerName)
    If columnIndex = 0 Or table.DataBodyRange Is Nothing Then Exit Function
    value = table.DataBodyRange.Cells(rowIndex, columnIndex).Value2
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    ViewerRawCell = CStr(value)
End Function

Private Function ViewerRecordedTime(ByVal table As ListObject, ByVal rowIndex As Long, ByVal headerName As String) As String
    Dim value As String
    value = ViewerCellText(table, rowIndex, headerName)
    If value = "" Then Exit Function
    If IsNumeric(value) Then
        ViewerRecordedTime = Format$(CDate(CDbl(value)), "yyyy-mm-dd hh:nn:ss")
    ElseIf IsDate(value) Then
        ViewerRecordedTime = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss")
    End If
End Function

Private Function ViewerFriendlyEventType(ByVal eventType As String) As String
    Select Case UCase$(Trim$(eventType))
        Case "RECEIVE": ViewerFriendlyEventType = "Receipt"
        Case "RETURN": ViewerFriendlyEventType = "Return"
        Case "DUMP": ViewerFriendlyEventType = "Dump"
        Case "BOX_BUILD": ViewerFriendlyEventType = "Box Made"
        Case "BOX_UNBOX": ViewerFriendlyEventType = "Box Unboxed"
        Case "SHIP": ViewerFriendlyEventType = "Shipped"
        Case "SHIP_RELEASE": ViewerFriendlyEventType = "Remove"
        Case "PROD_CONSUME": ViewerFriendlyEventType = "Production Input Consumed"
        Case "PROD_COMPLETE": ViewerFriendlyEventType = "Production Output Created"
        Case "ADMIN_INVENTORY_ADJUST": ViewerFriendlyEventType = "Inventory Adjustment"
    End Select
End Function

Private Function ViewerFirstNoteToken(ByVal noteText As String, ByVal tokenNames As Variant) As String
    Dim tokenName As Variant
    For Each tokenName In tokenNames
        ViewerFirstNoteToken = ViewerNoteToken(noteText, CStr(tokenName))
        If ViewerFirstNoteToken <> "" Then Exit Function
    Next tokenName
End Function

Private Function ViewerNoteToken(ByVal noteText As String, ByVal tokenName As String) As String
    Dim parts As Variant
    Dim part As Variant
    Dim prefix As String
    prefix = LCase$(Trim$(tokenName)) & "="
    parts = Split(noteText, ";")
    For Each part In parts
        If LCase$(Left$(Trim$(CStr(part)), Len(prefix))) = prefix Then
            ViewerNoteToken = Trim$(Mid$(Trim$(CStr(part)), Len(prefix) + 1))
            Exit Function
        End If
    Next part
End Function

Private Function ViewerFirstNonBlank(ByVal firstValue As String, ByVal secondValue As String) As String
    ViewerFirstNonBlank = Trim$(firstValue)
    If ViewerFirstNonBlank = "" Then ViewerFirstNonBlank = Trim$(secondValue)
End Function

Private Function ViewerWorkbookNameIsOpen(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            ViewerWorkbookNameIsOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function ViewerFindTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set ViewerFindTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not ViewerFindTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function ViewerColumnIndex(ByVal table As ListObject, ByVal headerName As String) As Long
    Dim columnIndex As Long

    If table Is Nothing Then Exit Function
    For columnIndex = 1 To table.ListColumns.Count
        If StrComp(Trim$(table.ListColumns(columnIndex).Name), Trim$(headerName), vbTextCompare) = 0 Then
            ViewerColumnIndex = columnIndex
            Exit Function
        End If
    Next columnIndex
End Function

Private Function ViewerCellText(ByVal table As ListObject, _
                                ByVal rowIndex As Long, _
                                ByVal headerName As String) As String
    Dim columnIndex As Long
    Dim valueIn As Variant

    columnIndex = ViewerColumnIndex(table, headerName)
    If columnIndex = 0 Or table.DataBodyRange Is Nothing Then Exit Function
    valueIn = table.DataBodyRange.Cells(rowIndex, columnIndex).Value2
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    ViewerCellText = Trim$(CStr(valueIn))
End Function

Private Function ViewerCellNumber(ByVal table As ListObject, _
                                  ByVal rowIndex As Long, _
                                  ByVal headerName As String) As Double
    Dim valueText As String

    valueText = ViewerCellText(table, rowIndex, headerName)
    If IsNumeric(valueText) Then ViewerCellNumber = CDbl(valueText)
End Function

Private Function ViewerEscape(ByVal valueIn As String) As String
    valueIn = Replace(valueIn, "\", "\\")
    valueIn = Replace(valueIn, vbTab, "\t")
    valueIn = Replace(valueIn, vbCr, "\r")
    valueIn = Replace(valueIn, vbLf, "\n")
    ViewerEscape = valueIn
End Function
