Attribute VB_Name = "modInventoryViewerData"
Option Explicit

Private Const SNAPSHOT_TABLE As String = "tblInventorySnapshot"

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
    LoadCurrentInventoryEventViewerData = modPublishedEventsReader.ReadCurrent()
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
