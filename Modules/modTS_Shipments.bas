Attribute VB_Name = "modTS_Shipments"
Option Explicit

' =============================================================
' Module: modTS_Shipments
' Purpose: All logic for the ShippingTally system (box builder,
'          holding subsystem, confirm/build/ship macros, logging).
' Notes:
'   - Buttons are generated dynamically (similar to modTS_Received).
'   - ShippingBOM sheet stores one ListObject per BOM (Box Name).
'   - BOM entries store ROW/QUANTITY/UOM only; item metadata is
'     resolved from invSys (InventoryManagement!invSys).
'   - Hold subsystem keeps packages on NotShipped until released.
'   - Additional confirm/build/ship routines will be implemented in
'     subsequent iterations (placeholders provided below).
' =============================================================

' ===== constants =====
Private Const SHEET_SHIPMENTS As String = "ShipmentsTally"
Private Const SHEET_INV As String = "InventoryManagement"
Private Const SHEET_BOM As String = "ShippingBOM"

Private Const TABLE_SHIPMENTS As String = "ShipmentsTally"
Private Const TABLE_NOTSHIPPED As String = "NotShipped"
Private Const TABLE_AGG_BOM As String = "AggregateBoxBOM"
Private Const TABLE_AGG_PACK As String = "AggregatePackages"
Private Const TABLE_BOX_BUILDER As String = "BoxBuilder"
Private Const TABLE_BOX_BOM As String = "BoxBOM"
Private Const TABLE_CHECK_INV As String = "Check_invSys"

Private Const BTN_SHOW_BUILDER As String = "BTN_SHOW_BUILDER"
Private Const BTN_HIDE_BUILDER As String = "BTN_HIDE_BUILDER"
Private Const BTN_SAVE_BOX As String = "BTN_SAVE_BOX"
Private Const BTN_UNSHIP As String = "BTN_UNSHIP"
Private Const BTN_SEND_HOLD As String = "BTN_SEND_HOLD"
Private Const BTN_RETURN_HOLD As String = "BTN_RETURN_HOLD"
Private Const BTN_CONFIRM_INV As String = "BTN_CONFIRM_INV"
Private Const BTN_BOXES_MADE As String = "BTN_BOXES_MADE"
Private Const BTN_TO_TOTALINV As String = "BTN_TO_TOTALINV"
Private Const BTN_TO_SHIPMENTS As String = "BTN_TO_SHIPMENTS"
Private Const BTN_SHIPMENTS_SENT As String = "BTN_SHIPMENTS_SENT"

Private Const SHIPPING_BOM_BLOCK_ROWS As Long = 52
Private Const SHIPPING_BOM_DATA_ROWS As Long = 50
Private Const SHIPPING_BOM_COLS As Long = 3 ' ROW, QUANTITY, UOM

Private mDynSearch As cDynItemSearch

' ===== public entry points =====
Public Sub InitializeShipmentsUI()
    EnsureShipmentsButtons
End Sub

Public Sub BtnShowBuilder()
    ToggleBuilderTables True
End Sub

Public Sub BtnHideBuilder()
    ToggleBuilderTables False
End Sub

Public Sub BtnSaveBox()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim loMeta As ListObject: Set loMeta = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loMeta Is Nothing Or loBom Is Nothing Then
        MsgBox "Box Builder tables not found on ShipmentsTally sheet.", vbExclamation
        Exit Sub
    End If

    EnsureTableHasRow loMeta
    EnsureColumnExists loMeta, "ROW"
    EnsureBoxBomEntryColumns loBom

    Dim boxName As String
    boxName = Trim$(NzStr(ValueFromTable(loMeta, "Box Name")))
    If boxName = "" Then
        MsgBox "Enter a Box Name before saving.", vbExclamation
        Exit Sub
    End If
    Dim boxUOM As String: boxUOM = Trim$(NzStr(ValueFromTable(loMeta, "UOM")))
    Dim boxLoc As String: boxLoc = Trim$(NzStr(ValueFromTable(loMeta, "LOCATION")))
    Dim boxDesc As String: boxDesc = Trim$(NzStr(ValueFromTable(loMeta, "DESCRIPTION")))
    If boxUOM = "" Then
        MsgBox "Box Builder UOM is required.", vbExclamation
        Exit Sub
    End If

    EnsureTableHasRow loBom

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim components As Collection
    Dim syncNotes As String
    Set components = CollectBomComponents(loBom, invLo, syncNotes)
    If components.count = 0 Then
        MsgBox "Add at least one valid component row (ROW/QUANTITY) to the BoxBOM table.", vbExclamation
        Exit Sub
    End If
    If components.count > SHIPPING_BOM_DATA_ROWS Then
        MsgBox "BOM exceeds the 50-row limit. Remove extra rows and try again.", vbExclamation
        Exit Sub
    End If

    Dim boxRowValue As Long
    boxRowValue = EnsureInvSysItem(boxName, boxUOM, boxLoc, boxDesc, invLo)
    If boxRowValue = 0 Then Exit Sub
    Dim cBoxRowField As Long: cBoxRowField = ColumnIndex(loMeta, "ROW")
    If cBoxRowField > 0 Then
        loMeta.DataBodyRange.Cells(1, cBoxRowField).Value = boxRowValue
    End If

    Dim wsBOM As Worksheet: Set wsBOM = SheetExists(SHEET_BOM)
    If wsBOM Is Nothing Then
        MsgBox "ShippingBOM sheet not found.", vbCritical
        Exit Sub
    End If
    Dim bomTable As ListObject, blockRange As Range
    Set bomTable = EnsureBomTable(wsBOM, boxName, blockRange)
    If bomTable Is Nothing Then Exit Sub

    WriteBomData bomTable, blockRange, components
    PropagateBomMetadata wsBOM, components

    Dim finalMsg As String
    finalMsg = "Saved BOM '" & boxName & "' (invSys ROW " & boxRowValue & ", " & components.count & " components)."
    If Len(syncNotes) > 0 Then
        finalMsg = finalMsg & vbCrLf & syncNotes
    End If
    MsgBox finalMsg, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BTN_SAVE_BOX failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnUnship()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = GetListObject(ws, TABLE_NOTSHIPPED)
    If lo Is Nothing Then
        MsgBox "NotShipped table not found.", vbExclamation
        Exit Sub
    End If
    Dim isHidden As Boolean
    isHidden = lo.Range.EntireColumn.Hidden
    lo.Range.EntireColumn.Hidden = Not isHidden
End Sub

Public Sub BtnSendHold()
    MoveSelectionToHold True
End Sub

Public Sub BtnReturnHold()
    MoveSelectionToHold False
End Sub

Public Sub BtnConfirmInventory()
    ' Placeholder for full confirm workflow
    MsgBox "BTN_CONFIRM_INV logic pending implementation.", vbInformation
End Sub

Public Sub BtnBoxesMade()
    ' Placeholder for BOM build workflow
    MsgBox "BTN_BOXES_MADE logic pending implementation.", vbInformation
End Sub

Public Sub BtnToTotalInv()
    MsgBox "BTN_TO_TOTALINV logic pending implementation.", vbInformation
End Sub

Public Sub BtnToShipments()
    MsgBox "BTN_TO_SHIPMENTS logic pending implementation.", vbInformation
End Sub

Public Sub BtnShipmentsSent()
    MsgBox "BTN_SHIPMENTS_SENT logic pending implementation.", vbInformation
End Sub

Public Sub ShowDynamicItemSearch(ByVal targetCell As Range)
    On Error GoTo ErrHandler
    If targetCell Is Nothing Then Exit Sub
    If mDynSearch Is Nothing Then Set mDynSearch = New cDynItemSearch
    mDynSearch.ShowForCell targetCell
    Exit Sub
ErrHandler:
    On Error Resume Next
    frmItemSearch.Show vbModeless
End Sub

' ===== button scaffolding =====
Private Sub EnsureShipmentsButtons()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim leftA As Double: leftA = ws.Columns("A").Left + 4
    Dim nextTop As Double: nextTop = ws.Rows(2).Top

    EnsureButtonCustom ws, BTN_SHOW_BUILDER, "Show builder", "modTS_Shipments.BtnShowBuilder", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_HIDE_BUILDER, "Hide builder", "modTS_Shipments.BtnHideBuilder", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_SAVE_BOX, "Save box", "modTS_Shipments.BtnSaveBox", leftA, nextTop
    nextTop = nextTop + 28
    EnsureButtonCustom ws, BTN_CONFIRM_INV, "Confirm inventory", "modTS_Shipments.BtnConfirmInventory", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_BOXES_MADE, "Boxes made", "modTS_Shipments.BtnBoxesMade", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_TO_TOTALINV, "To TotalInv", "modTS_Shipments.BtnToTotalInv", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_TO_SHIPMENTS, "To Shipments", "modTS_Shipments.BtnToShipments", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_SHIPMENTS_SENT, "Shipments sent", "modTS_Shipments.BtnShipmentsSent", leftA, nextTop

    Dim loHold As ListObject: Set loHold = GetListObject(ws, TABLE_NOTSHIPPED)
    If Not loHold Is Nothing Then
        Dim topBand As Double
        topBand = loHold.HeaderRowRange.Top - 24
        Dim leftBand As Double
        leftBand = loHold.HeaderRowRange.Left
        EnsureButtonCustom ws, BTN_UNSHIP, "Toggle NotShipped", "modTS_Shipments.BtnUnship", leftBand, topBand
        EnsureButtonCustom ws, BTN_SEND_HOLD, "Send to hold", "modTS_Shipments.BtnSendHold", leftBand + 120, topBand
        EnsureButtonCustom ws, BTN_RETURN_HOLD, "Return from hold", "modTS_Shipments.BtnReturnHold", leftBand + 240, topBand
    End If
End Sub

Private Sub EnsureButtonCustom(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String, leftPos As Double, topPos As Double)
    Const BTN_WIDTH As Double = 118
    Const BTN_HEIGHT As Double = 20
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, BTN_WIDTH, BTN_HEIGHT)
        shp.Name = shapeName
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = onActionMacro
    Else
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = BTN_WIDTH
        shp.Height = BTN_HEIGHT
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = onActionMacro
    End If
End Sub

' ===== builder helpers =====
Private Sub ToggleBuilderTables(ByVal makeVisible As Boolean)
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo1 As ListObject: Set lo1 = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim lo2 As ListObject: Set lo2 = GetListObject(ws, TABLE_BOX_BOM)
    If lo1 Is Nothing And lo2 Is Nothing Then Exit Sub

    Dim firstCol As Long: firstCol = 0
    Dim lastCol As Long: lastCol = 0

    Dim arrTables As Variant
    arrTables = Array(lo1, lo2)
    Dim idx As Long
    Dim lo As ListObject
    For idx = LBound(arrTables) To UBound(arrTables)
        Set lo = arrTables(idx)
        If Not lo Is Nothing Then
            Dim startCol As Long
            Dim endCol As Long
            startCol = lo.HeaderRowRange.Column
            endCol = startCol + lo.HeaderRowRange.Columns.Count - 1
            If firstCol = 0 Or startCol < firstCol Then firstCol = startCol
            If endCol > lastCol Then lastCol = endCol
        End If
    Next idx

    If firstCol = 0 Or lastCol = 0 Then Exit Sub

    ws.Range(ws.Columns(firstCol), ws.Columns(lastCol)).EntireColumn.Hidden = Not makeVisible
End Sub

Public Sub ApplyItemSelection(targetCell As Range, lo As ListObject, rowIndex As Long, _
    ByVal itemName As String, ByVal itemCode As String, ByVal itemRow As Long, _
    ByVal uom As String, ByVal location As String, ByVal vendor As String)

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
    If rowIndex <= 0 Or rowIndex > lo.ListRows.Count Then rowIndex = lo.ListRows.Count
    Dim cItems As Long: cItems = ColumnIndex(lo, "ITEMS")
    If cItems > 0 Then lo.DataBodyRange.Cells(rowIndex, cItems).Value = itemName
    ' Future enhancement: capture ROW/UOM metadata once staging columns are defined.
End Sub

Private Function CollectBomComponents(loBom As ListObject, invLo As ListObject, ByRef syncNotes As String) As Collection
    Dim result As New Collection
    If loBom Is Nothing Or invLo Is Nothing Then
        Set CollectBomComponents = result
        Exit Function
    End If

    Dim cName As Long: cName = ColumnIndex(loBom, "BoxBOM")
    Dim cRow As Long: cRow = ColumnIndex(loBom, "ROW")
    Dim cQty As Long: cQty = ColumnIndex(loBom, "QUANTITY")
    Dim cUom As Long: cUom = ColumnIndex(loBom, "UOM")
    Dim cLoc As Long: cLoc = ColumnIndex(loBom, "LOCATION")
    Dim cDesc As Long: cDesc = ColumnIndex(loBom, "DESCRIPTION")
    If cName = 0 Or cRow = 0 Or cQty = 0 Or cUom = 0 Then
        MsgBox "BoxBOM table must include BoxBOM, ROW, QUANTITY, and UOM columns.", vbExclamation
        Exit Function
    End If

    If loBom.DataBodyRange Is Nothing Then
        Set CollectBomComponents = result
        Exit Function
    End If

    Dim invRowCol As Long: invRowCol = ColumnIndex(invLo, "ROW")
    Dim invItemCol As Long: invItemCol = ColumnIndex(invLo, "ITEM")
    Dim invUomCol As Long: invUomCol = ColumnIndex(invLo, "UOM")
    Dim invLocCol As Long: invLocCol = ColumnIndex(invLo, "LOCATION")
    Dim invDescCol As Long: invDescCol = ColumnIndex(invLo, "DESCRIPTION")
    If invRowCol = 0 Then
        MsgBox "invSys table must contain a ROW column.", vbCritical
        Exit Function
    End If

    Dim arr As Variant: arr = loBom.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim partName As String: partName = Trim$(NzStr(arr(r, cName)))
        Dim partRow As Long: partRow = NzLng(arr(r, cRow))
        Dim qty As Double: qty = NzDbl(arr(r, cQty))
        Dim uomVal As String: uomVal = Trim$(NzStr(arr(r, cUom)))

        If partName = "" And partRow = 0 And qty = 0 Then GoTo NextComponent
        If qty <= 0 Then
            Err.Raise vbObjectError + 1, , "Component row " & r & " has no quantity."
        End If

        Dim invIdx As Long
        Dim partResolvedName As String
        If partRow > 0 Then
            invIdx = FindInvRowIndexByRow(invLo, partRow)
            If invIdx = 0 And partName <> "" Then
                invIdx = FindInvRowIndexByItem(invLo, partName)
                If invIdx > 0 Then
                    Dim resolvedRow As Long
                    resolvedRow = NzLng(invLo.DataBodyRange.Cells(invIdx, invRowCol).Value)
                    If resolvedRow <> partRow Then
                        partRow = resolvedRow
                        AppendSyncMessage syncNotes, "Updated ROW for '" & partName & "' to " & resolvedRow & "."
                    End If
                End If
            End If
            If invIdx = 0 Then
                Err.Raise vbObjectError + 2, , "Component row " & partRow & " not found in invSys. Update BOM before saving."
            End If
        ElseIf partName <> "" Then
            invIdx = FindInvRowIndexByItem(invLo, partName)
            If invIdx = 0 Then
                Err.Raise vbObjectError + 3, , "Component '" & partName & "' not found in invSys."
            End If
            partRow = NzLng(invLo.DataBodyRange.Cells(invIdx, invRowCol).Value)
        Else
            Err.Raise vbObjectError + 4, , "Component row " & r & " is missing both item name and ROW."
        End If

        Dim actualUom As String, actualLoc As String, actualDesc As String
        Dim actualItem As String
        If invItemCol > 0 Then actualItem = NzStr(invLo.DataBodyRange.Cells(invIdx, invItemCol).Value)
        If invUomCol > 0 Then actualUom = NzStr(invLo.DataBodyRange.Cells(invIdx, invUomCol).Value)
        If invLocCol > 0 Then actualLoc = NzStr(invLo.DataBodyRange.Cells(invIdx, invLocCol).Value)
        If invDescCol > 0 Then actualDesc = NzStr(invLo.DataBodyRange.Cells(invIdx, invDescCol).Value)
        If actualItem <> "" Then partResolvedName = actualItem Else partResolvedName = partName
        If actualUom = "" Then actualUom = uomVal
        If StrComp(uomVal, actualUom, vbTextCompare) <> 0 Then
            AppendSyncMessage syncNotes, "UOM for '" & partResolvedName & "' reset to " & actualUom & "."
        End If
        uomVal = actualUom

        If cName > 0 And partResolvedName <> "" Then
            loBom.DataBodyRange.Cells(r, cName).Value = partResolvedName
        End If
        loBom.DataBodyRange.Cells(r, cRow).Value = partRow
        loBom.DataBodyRange.Cells(r, cUom).Value = uomVal
        If cLoc > 0 Then loBom.DataBodyRange.Cells(r, cLoc).Value = actualLoc
        If cDesc > 0 Then loBom.DataBodyRange.Cells(r, cDesc).Value = actualDesc

        Dim entry(1 To 3) As Variant
        entry(1) = partRow
        entry(2) = qty
        entry(3) = uomVal
        result.Add entry
NextComponent:
    Next

    Set CollectBomComponents = result
End Function

Private Sub EnsureBoxBomEntryColumns(loBom As ListObject)
    If loBom Is Nothing Then Exit Sub
    EnsureColumnExists loBom, "BoxBOM"
    EnsureColumnExists loBom, "QUANTITY", "BoxBOM"
    EnsureColumnExists loBom, "ROW"
    EnsureColumnExists loBom, "UOM"
    EnsureColumnExists loBom, "LOCATION"
    EnsureColumnExists loBom, "DESCRIPTION"
End Sub

Private Sub EnsureColumnExists(lo As ListObject, colName As String, Optional afterColumn As String = "")
    If lo Is Nothing Then Exit Sub
    If ColumnIndex(lo, colName) > 0 Then Exit Sub
    Dim insertPos As Long
    If afterColumn <> "" Then insertPos = ColumnIndex(lo, afterColumn)
    Dim newCol As ListColumn
    If insertPos > 0 Then
        Set newCol = lo.ListColumns.Add(insertPos + 1)
    Else
        Set newCol = lo.ListColumns.Add
    End If
    newCol.Name = colName
End Sub

Private Sub PropagateBomMetadata(ws As Worksheet, comps As Collection)
    If ws Is Nothing Then Exit Sub
    If comps Is Nothing Then Exit Sub
    If comps.count = 0 Then Exit Sub
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To comps.count
        Dim info As Variant
        info = comps(i)
        Dim rowVal As Long: rowVal = NzLng(info(1))
        Dim uomVal As String: uomVal = NzStr(info(3))
        If rowVal > 0 Then
            If Not seen.Exists(rowVal) Then
                seen(rowVal) = True
                SyncSavedBomRows ws, rowVal, uomVal
            End If
        End If
    Next
End Sub

Private Sub SyncSavedBomRows(ws As Worksheet, ByVal rowValue As Long, ByVal uomValue As String)
    If ws Is Nothing Or rowValue = 0 Then Exit Sub
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        Dim cRow As Long: cRow = ColumnIndex(lo, "ROW")
        Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
        If cRow = 0 Or cUom = 0 Then GoTo NextTable
        If lo.DataBodyRange Is Nothing Then GoTo NextTable
        Dim lr As ListRow
        For Each lr In lo.ListRows
            If NzLng(lr.Range.Cells(1, cRow).Value) = rowValue Then
                lr.Range.Cells(1, cUom).Value = uomValue
            End If
        Next lr
NextTable:
    Next lo
End Sub

Private Sub AppendSyncMessage(ByRef target As String, ByVal text As String)
    If Len(text) = 0 Then Exit Sub
    If Len(target) = 0 Then
        target = text
    Else
        target = target & vbCrLf & text
    End If
End Sub

Private Function EnsureBomTable(ws As Worksheet, ByVal boxName As String, ByRef blockRange As Range) As ListObject
    Dim cleanName As String: cleanName = SafeTableName(boxName)

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(cleanName)
    On Error GoTo 0
    If Not lo Is Nothing Then
        Set blockRange = BlockRangeFromHeader(ws, lo.HeaderRowRange.Row)
        If blockRange Is Nothing Then
            Set blockRange = lo.Range
        End If
        lo.Resize blockRange
        lo.HeaderRowRange.Cells(1, 1).Value = "ROW"
        lo.HeaderRowRange.Cells(1, 2).Value = "QUANTITY"
        lo.HeaderRowRange.Cells(1, 3).Value = "UOM"
        Set EnsureBomTable = lo
        Exit Function
    End If

    Dim startRow As Long: startRow = NextAvailableBomRow(ws)
    If startRow = 0 Then
        MsgBox "ShippingBOM sheet has no space for additional BOMs.", vbCritical
        Exit Function
    End If
    Set blockRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + SHIPPING_BOM_DATA_ROWS, SHIPPING_BOM_COLS))
    blockRange.Clear
    blockRange.Rows(1).Cells(1, 1).Value = "ROW"
    blockRange.Rows(1).Cells(1, 2).Value = "QUANTITY"
    blockRange.Rows(1).Cells(1, 3).Value = "UOM"
    Set lo = ws.ListObjects.Add(xlSrcRange, blockRange, , xlYes)
    lo.Name = cleanName
    Set EnsureBomTable = lo
End Function

Private Sub WriteBomData(lo As ListObject, blockRange As Range, comps As Collection)
    If lo Is Nothing Then Exit Sub
    If blockRange Is Nothing Then Set blockRange = lo.Range
    lo.Resize blockRange
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    If comps.count = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To comps.count
        Dim info As Variant
        info = comps(i)
        lo.DataBodyRange.Cells(i, 1).Value = info(1)
        lo.DataBodyRange.Cells(i, 2).Value = info(2)
        lo.DataBodyRange.Cells(i, 3).Value = info(3)
    Next
End Sub

Private Function NextAvailableBomRow(ws As Worksheet) As Long
    Dim totalRows As Long: totalRows = ws.Rows.Count
    Dim startRow As Long
    startRow = 1
    Do
        If startRow + SHIPPING_BOM_BLOCK_ROWS - 1 > totalRows Then
            NextAvailableBomRow = 0
            Exit Function
        End If
        If IsBlockFree(ws, startRow) Then
            NextAvailableBomRow = startRow
            Exit Function
        End If
        startRow = startRow + SHIPPING_BOM_BLOCK_ROWS
    Loop
End Function

Private Function IsBlockFree(ws As Worksheet, startRow As Long) As Boolean
    Dim rg As Range
    Set rg = BlockRangeFromHeader(ws, startRow)
    If rg Is Nothing Then
        IsBlockFree = False
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If Not Intersect(lo.Range, rg) Is Nothing Then
            IsBlockFree = False
            Exit Function
        End If
    Next
    IsBlockFree = True
End Function

Private Function BlockRangeFromHeader(ws As Worksheet, startRow As Long) As Range
    On Error Resume Next
    Set BlockRangeFromHeader = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + SHIPPING_BOM_DATA_ROWS, SHIPPING_BOM_COLS))
    On Error GoTo 0
End Function

Private Function SafeTableName(ByVal sourceName As String) As String
    Dim cleaned As String
    cleaned = Trim$(sourceName)
    If cleaned = "" Then cleaned = "BOM_" & Format(Now, "yyyymmdd_hhnnss")
    Dim i As Long, ch As String, kept As String
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            kept = kept & ch
        Else
            kept = kept & "_"
        End If
    Next
    If kept = "" Then kept = "BOM_" & Format(Now, "yyyymmdd_hhnnss")
    If Not kept Like "[A-Za-z_]*" Then kept = "BOM_" & kept
    SafeTableName = kept
End Function

Private Function ValueFromTable(lo As ListObject, headerName As String) As Variant
    Dim colIdx As Long: colIdx = ColumnIndex(lo, headerName)
    If colIdx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ValueFromTable = lo.DataBodyRange.Cells(1, colIdx).Value
End Function

' ===== hold helpers =====
Private Sub MoveSelectionToHold(ByVal moveToHold As Boolean)
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim loShip As ListObject: Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    Dim loHold As ListObject: Set loHold = GetListObject(ws, TABLE_NOTSHIPPED)
    If loShip Is Nothing Or loHold Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub

    Dim targetTable As ListObject
    Dim sourceTable As ListObject
    If moveToHold Then
        Set sourceTable = loShip
        Set targetTable = loHold
    Else
        Set sourceTable = loHold
        Set targetTable = loShip
    End If

    Dim rngSel As Range
    On Error Resume Next
    Set rngSel = Application.Intersect(Application.Selection, sourceTable.DataBodyRange)
    On Error GoTo 0
    If rngSel Is Nothing Then
        MsgBox "Select rows inside the " & sourceTable.Name & " table first.", vbInformation
        Exit Sub
    End If

    Dim processed As Object: Set processed = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    For Each cell In rngSel.Areas
        Dim r As Range
        For Each r In cell.Rows
            Dim rowIndex As Long
            rowIndex = r.Row - sourceTable.DataBodyRange.Row + 1
            If rowIndex >= 1 And rowIndex <= sourceTable.ListRows.Count Then
                If Not processed.Exists(rowIndex) Then
                    processed(rowIndex) = True
                    HandleHoldRow sourceTable, targetTable, rowIndex, moveToHold
                End If
            End If
        Next r
    Next cell
End Sub

Private Sub HandleHoldRow(sourceTable As ListObject, targetTable As ListObject, rowIndex As Long, moveToHold As Boolean)
    Dim cRef As Long: cRef = ColumnIndex(sourceTable, "REF_NUMBER")
    Dim cItems As Long: cItems = ColumnIndex(sourceTable, "ITEMS")
    Dim cQty As Long: cQty = ColumnIndex(sourceTable, "QUANTITY")
    If cQty = 0 Then
        MsgBox sourceTable.Name & " table needs a QUANTITY column.", vbCritical
        Exit Sub
    End If

    Dim refVal As String: refVal = NzStr(sourceTable.DataBodyRange.Cells(rowIndex, cRef).Value)
    Dim itemVal As String: itemVal = NzStr(sourceTable.DataBodyRange.Cells(rowIndex, cItems).Value)
    Dim qtyVal As Double: qtyVal = NzDbl(sourceTable.DataBodyRange.Cells(rowIndex, cQty).Value)
    If qtyVal <= 0 Then Exit Sub

    Dim prompt As String
    If moveToHold Then
        prompt = "Enter quantity to hold for '" & itemVal & "' (available " & qtyVal & "):"
    Else
        prompt = "Enter quantity to return to shipments for '" & itemVal & "' (available " & qtyVal & "):"
    End If
    Dim qtyInput As Variant
    qtyInput = Application.InputBox(prompt, "Hold quantity", qtyVal, Type:=1)
    If qtyInput = False Then Exit Sub
    Dim qtyMove As Double: qtyMove = CDbl(qtyInput)
    If qtyMove <= 0 Then Exit Sub
    If qtyMove > qtyVal Then qtyMove = qtyVal

    AppendHoldRow targetTable, refVal, itemVal, qtyMove

    Dim newQty As Double
    If moveToHold Then
        newQty = qtyVal - qtyMove
    Else
        newQty = qtyVal - qtyMove
    End If
    If newQty <= 0 Then
        sourceTable.ListRows(rowIndex).Range.ClearContents
    Else
        sourceTable.DataBodyRange.Cells(rowIndex, cQty).Value = newQty
    End If
End Sub

Private Sub AppendHoldRow(targetTable As ListObject, refVal As String, itemVal As String, qtyMove As Double)
    Dim cRef As Long: cRef = ColumnIndex(targetTable, "REF_NUMBER")
    Dim cItems As Long: cItems = ColumnIndex(targetTable, "ITEMS")
    Dim cQty As Long: cQty = ColumnIndex(targetTable, "QUANTITY")
    Dim lr As ListRow: Set lr = targetTable.ListRows.Add
    If cRef > 0 Then lr.Range.Cells(1, cRef).Value = refVal
    If cItems > 0 Then lr.Range.Cells(1, cItems).Value = itemVal
    If cQty > 0 Then lr.Range.Cells(1, cQty).Value = qtyMove
End Sub

' ===== helpers reused from modTS_Received =====
Private Function SheetExists(nameOrCode As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set SheetExists = ws
            Exit Function
        End If
    Next ws
End Function

Private Function GetListObject(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function GetInvSysTable() As ListObject
    Dim wsInv As Worksheet: Set wsInv = SheetExists(SHEET_INV)
    If wsInv Is Nothing Then Exit Function
    On Error Resume Next
    Set GetInvSysTable = wsInv.ListObjects("invSys")
    On Error GoTo 0
End Function

Private Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    ColumnIndex = 0
End Function

Private Function FindInvRowIndexByRow(invLo As ListObject, ByVal rowValue As Long) As Long
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then Exit Function
    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If NzLng(invLo.DataBodyRange.Cells(r, cRow).Value) = rowValue Then
            FindInvRowIndexByRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindInvRowIndexByItem(invLo As ListObject, ByVal itemName As String) As Long
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    Dim cItem As Long: cItem = ColumnIndex(invLo, "ITEM")
    If cItem = 0 Then Exit Function
    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If StrComp(Trim$(NzStr(invLo.DataBodyRange.Cells(r, cItem).Value)), Trim$(itemName), vbTextCompare) = 0 Then
            FindInvRowIndexByItem = r
            Exit Function
        End If
    Next r
End Function

Private Function NextInvSysRowValue(invLo As ListObject) As Long
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then
        NextInvSysRowValue = invLo.ListRows.Count + 1
        Exit Function
    End If
    Dim maxVal As Long: maxVal = 0
    If Not invLo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To invLo.DataBodyRange.Rows.Count
            Dim v As Variant: v = invLo.DataBodyRange.Cells(r, cRow).Value
            If IsNumeric(v) Then
                If CLng(v) > maxVal Then maxVal = CLng(v)
            End If
        Next r
    End If
    NextInvSysRowValue = maxVal + 1
End Function

Private Function EnsureInvSysItem(boxName As String, uom As String, location As String, descr As String, invLo As ListObject) As Long
    If invLo Is Nothing Then Exit Function
    Dim existingIdx As Long
    existingIdx = FindInvRowIndexByItem(invLo, boxName)
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If existingIdx > 0 Then
        EnsureInvSysItem = NzLng(invLo.DataBodyRange.Cells(existingIdx, cRow).Value)
        UpdateInvSysRow invLo.ListRows(existingIdx), boxName, uom, location, descr
        Exit Function
    End If

    Dim lr As ListRow: Set lr = invLo.ListRows.Add
    Dim newRowVal As Long: newRowVal = NextInvSysRowValue(invLo)
    EnsureInvSysItem = newRowVal
    UpdateInvSysRow lr, boxName, uom, location, descr, newRowVal
End Function

Private Sub UpdateInvSysRow(lr As ListRow, boxName As String, uom As String, location As String, descr As String, Optional forceRowValue As Variant)
    If lr Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = lr.Parent
    Dim idx As Long
    If Not IsMissing(forceRowValue) Then
        idx = ColumnIndex(lo, "ROW")
        If idx > 0 Then lr.Range.Cells(1, idx).Value = forceRowValue
    End If
    idx = ColumnIndex(lo, "ITEM")
    If idx > 0 Then lr.Range.Cells(1, idx).Value = boxName
    idx = ColumnIndex(lo, "ITEM_CODE")
    If idx > 0 And Trim$(NzStr(lr.Range.Cells(1, idx).Value)) = "" Then
        lr.Range.Cells(1, idx).Value = boxName
    End If
    idx = ColumnIndex(lo, "UOM")
    If idx > 0 Then lr.Range.Cells(1, idx).Value = uom
    idx = ColumnIndex(lo, "LOCATION")
    If idx > 0 Then lr.Range.Cells(1, idx).Value = location
    idx = ColumnIndex(lo, "DESCRIPTION")
    If idx > 0 Then lr.Range.Cells(1, idx).Value = descr
End Sub

Private Sub EnsureTableHasRow(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Public Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Public Function NzDbl(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzDbl = 0#
    Else
        NzDbl = CDbl(v)
    End If
End Function

Public Function NzLng(v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzLng = 0
    Else
        NzLng = CLng(v)
    End If
End Function

' ===== Workbook/setup helpers (migrated from modTS_Data) =====
Public Sub SetupAllHandlers()
    On Error Resume Next
    ClearTableFilters
    modGlobals.InitializeGlobalVariables
    Application.OnKey "{F3}", "modGlobals.OpenItemSearchForCurrentCell"
    modTS_Received.EnsureGeneratedButtons
    InitializeShipmentsUI
    On Error GoTo 0
End Sub

Public Sub GenerateRowNumbers()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowNumCol As Long
    Dim i As Long
    Dim maxRowNum As Long
    Dim newCol As ListColumn

    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")

    On Error Resume Next
    rowNumCol = tbl.ListColumns("ROW").Index
    On Error GoTo ErrorHandler
    If rowNumCol = 0 Then
        Set newCol = tbl.ListColumns.Add
        newCol.Name = "ROW"
        rowNumCol = newCol.Index
    End If

    maxRowNum = 0
    For i = 1 To tbl.ListRows.Count
        If IsNumeric(tbl.DataBodyRange(i, rowNumCol).Value) Then
            maxRowNum = Application.WorksheetFunction.Max(maxRowNum, tbl.DataBodyRange(i, rowNumCol).Value)
        End If
    Next i

    For i = 1 To tbl.ListRows.Count
        If Trim$(tbl.DataBodyRange(i, rowNumCol).Value & "") = "" Then
            maxRowNum = maxRowNum + 1
            tbl.DataBodyRange(i, rowNumCol).Value = maxRowNum
        End If
    Next i
    MsgBox "Row numbers have been generated successfully.", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error generating row numbers: " & Err.Description, vbExclamation
End Sub

Public Function IsInItemsColumn(Target As Range) As Boolean
    IsInItemsColumn = False
    On Error Resume Next

    Dim lo As ListObject
    Set lo = Target.ListObject
    If lo Is Nothing Then Exit Function

    If lo.Name <> "ShipmentsTally" And lo.Name <> "ReceivedTally" Then Exit Function

    Dim itemsCol As ListColumn
    Set itemsCol = lo.ListColumns("ITEMS")
    On Error GoTo 0
    If itemsCol Is Nothing Then Exit Function

    If Target.Column = itemsCol.Range.Column Then
        If Target.Row > lo.HeaderRowRange.Row Then
            IsInItemsColumn = True
        End If
    End If
End Function

Public Sub ClearTableFilters()
    On Error Resume Next
    If Not ThisWorkbook.Worksheets("ShipmentsTally") Is Nothing Then
        If Not ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("ShipmentsTally") Is Nothing Then
            ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("ShipmentsTally").AutoFilter.ShowAllData
        End If
        If Not ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("invSysData_Shipping") Is Nothing Then
            ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("invSysData_Shipping").AutoFilter.ShowAllData
        End If
    End If
    If Not ThisWorkbook.Worksheets("ReceivedTally") Is Nothing Then
        If Not ThisWorkbook.Worksheets("ReceivedTally").ListObjects("ReceivedTally") Is Nothing Then
            ThisWorkbook.Worksheets("ReceivedTally").ListObjects("ReceivedTally").AutoFilter.ShowAllData
        End If
        If Not ThisWorkbook.Worksheets("ReceivedTally").ListObjects("invSysData_Receiving") Is Nothing Then
            ThisWorkbook.Worksheets("ReceivedTally").ListObjects("invSysData_Receiving").AutoFilter.ShowAllData
        End If
    End If
    On Error GoTo 0
End Sub

Public Function LoadItemList() As Variant
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowCount As Long
    Dim result As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Then GoTo ErrorHandler

    rowCount = tbl.ListRows.Count
    If rowCount = 0 Then GoTo ErrorHandler

    ReDim result(1 To rowCount, 0 To 4)

    Dim itemCodeCol As Integer, rowCol As Integer, itemCol As Integer
    Dim locCol As Integer
    On Error Resume Next
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    rowCol = tbl.ListColumns("ROW").Index
    itemCol = tbl.ListColumns("ITEM").Index
    locCol = tbl.ListColumns("LOCATION").Index
    On Error GoTo ErrorHandler
    If itemCodeCol = 0 Or rowCol = 0 Or itemCol = 0 Then GoTo ErrorHandler

    For i = 1 To rowCount
        result(i, 0) = tbl.DataBodyRange.Cells(i, rowCol).Value
        result(i, 1) = tbl.DataBodyRange.Cells(i, itemCodeCol).Value
        result(i, 2) = tbl.DataBodyRange.Cells(i, itemCol).Value
        If locCol > 0 Then
            result(i, 3) = tbl.DataBodyRange.Cells(i, locCol).Value
        End If
    Next i
    LoadItemList = result
    Exit Function
ErrorHandler:
    LoadItemList = Empty
End Function
