Attribute VB_Name = "modTS_Received"
Option Explicit

' =============================================================
' Module: modTS_Received
' Purpose: Consolidated receiving workflow
'   Item Search -> ReceivedTally -> AggregateReceived -> Confirm -> invSys/ReceivedLog
' Notes:
'   - Replaces legacy Tally form/button and invSysData_Receiving table.
'   - Buttons (Confirm, Undo, Redo) are generated once; recreate if missing, never duplicate.
'   - ReceivedTally is minimal (REF_NUMBER, ITEMS, QUANTITY) for fast entry.
'   - AggregateReceived holds detail: REF_NUMBER, ITEM_CODE, VENDORS, VENDOR_CODE, DESCRIPTION, ITEM, UOM, QUANTITY, LOCATION, ROW.
'   - Confirm only adds QUANTITY to existing invSys rows (items must pre-exist).
'   - ReceivedLog: REF_NUMBER, ITEMS, QUANTITY, UOM, VENDOR, LOCATION, ITEM_CODE, ROW, SNAPSHOT_ID, ENTRY_DATE.
'   - invSys table (InventoryManagement sheet): columns include ROW, ITEM_CODE, ITEM, UOM, LOCATION, RECEIVED, TOTAL INV, TIMESTAMP, etc.
'   - Single-level undo/redo for last confirm.
' =============================================================

' ==== module-level undo/redo state ====
Private mUndoInv As Collection
Private mUndoLogRows As Collection
Private mUndoRT As Variant
Private mUndoAGG As Variant
Private mRedoReady As Boolean

' ==== public entry points =====
Public Sub EnsureGeneratedButtons()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = SheetExists("ReceivedTally")
    If ws Is Nothing Then Exit Sub
    ' Simple guard: check for shapes by name; if missing, create.
    EnsureButton ws, "btnConfirmWrites", "Confirm writes", "modTS_Received.ConfirmWrites"
    EnsureButton ws, "btnUndoMacro", "Undo macro", "modTS_Received.MacroUndo"
    EnsureButton ws, "btnRedoMacro", "Redo macro", "modTS_Received.MacroRedo"
End Sub

' Called by frmItemSearch after user picks an item
Public Sub AddOrMergeFromSearch( _
    ByVal refNumber As String, _
    ByVal itemName As String, _
    ByVal itemCode As String, _
    ByVal qty As Double, _
    ByVal vendors As String, _
    ByVal vendorCode As String, _
    ByVal descr As String, _
    ByVal uom As String, _
    ByVal location As String, _
    ByVal invRow As Long)

    Dim ws As Worksheet: Set ws = SheetExists("ReceivedTally")
    If ws Is Nothing Then Exit Sub

    Dim rt As ListObject, agg As ListObject
    Set rt = ws.ListObjects("ReceivedTally")
    Set agg = ws.ListObjects("AggregateReceived")

    ' Insert/merge into ReceivedTally (fast entry: ITEMS + QUANTITY + REF_NUMBER)
    MergeIntoReceivedTally rt, refNumber, itemName, qty

    ' Insert/merge into AggregateReceived (detailed)
    MergeIntoAggregate agg, refNumber, itemCode, vendors, vendorCode, descr, itemName, uom, qty, location, invRow
End Sub

Public Sub RebuildAggregation()
    Dim ws As Worksheet: Set ws = SheetExists("ReceivedTally")
    If ws Is Nothing Then Exit Sub
    Dim rt As ListObject: Set rt = ws.ListObjects("ReceivedTally")
    Dim agg As ListObject: Set agg = ws.ListObjects("AggregateReceived")
    Dim catalog As ListObject: Set catalog = SheetExists("InventoryManagement").ListObjects("invSys")

    If rt Is Nothing Or agg Is Nothing Or catalog Is Nothing Then Exit Sub
    ClearTable agg

    If rt.DataBodyRange Is Nothing Then Exit Sub
    Dim arr, r As Long
    arr = rt.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        Dim itemName As String, qty As Double
        itemName = NzStr(arr(r, ColumnIndex(rt, "ITEMS")))
        qty = NzDbl(arr(r, ColumnIndex(rt, "QUANTITY")))
        Dim refNumber As String
        refNumber = NzStr(arr(r, ColumnIndex(rt, "REF_NUMBER")))

        Dim itemCode As String, vendors As String, vendorCode As String
        Dim descr As String, uom As String, location As String, invRow As Long
        itemCode = "": vendors = "": vendorCode = "": descr = "": uom = "": location = "": invRow = 0
        LookupInvSys catalog, itemName, itemCode, vendors, vendorCode, descr, uom, location, invRow

        MergeIntoAggregate agg, refNumber, itemCode, vendors, vendorCode, descr, itemName, uom, qty, location, invRow
    Next r
    ' Ensure quantity shows as number, not date
    On Error Resume Next
    agg.ListColumns("QUANTITY").DataBodyRange.NumberFormat = "0.00"
    rt.ListColumns("QUANTITY").DataBodyRange.NumberFormat = "0.00"
    On Error GoTo 0
End Sub

Public Sub ConfirmWrites()
    On Error GoTo ErrHandler
    mRedoReady = False
    Dim wsRT As Worksheet: Set wsRT = SheetExists("ReceivedTally")
    Dim wsAgg As Worksheet: Set wsAgg = SheetExists("ReceivedTally")
    Dim wsInv As Worksheet: Set wsInv = SheetExists("InventoryManagement")
    Dim wsLog As Worksheet: Set wsLog = SheetExists("ReceivedLog")
    If wsRT Is Nothing Or wsAgg Is Nothing Or wsInv Is Nothing Or wsLog Is Nothing Then Exit Sub

    Dim agg As ListObject: Set agg = wsAgg.ListObjects("AggregateReceived")
    Dim inv As ListObject: Set inv = wsInv.ListObjects("invSys")
    Dim logTbl As ListObject: Set logTbl = wsLog.ListObjects("ReceivedLog")
    If agg Is Nothing Or inv Is Nothing Or logTbl Is Nothing Then Exit Sub
    If agg.DataBodyRange Is Nothing Then Exit Sub

    ' Validate and collect rows
    Dim arr, r As Long, errs As String
    arr = agg.DataBodyRange.Value
    Dim cols As Object: Set cols = AggColMap(agg)

    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cols("ITEM"))) = "" And NzStr(arr(r, cols("ITEM_CODE"))) = "" Then errs = errs & "Row " & r & ": ITEM/ITEM_CODE missing" & vbCrLf
        If NzStr(arr(r, cols("UOM"))) = "" Then errs = errs & "Row " & r & ": UOM missing" & vbCrLf
        If NzDbl(arr(r, cols("QUANTITY"))) <= 0 Then errs = errs & "Row " & r & ": QUANTITY <= 0" & vbCrLf
        If NzLng(arr(r, cols("ROW"))) <= 0 Then errs = errs & "Row " & r & ": ROW missing" & vbCrLf
    Next
    If errs <> "" Then
        MsgBox "Cannot confirm:" & vbCrLf & errs, vbExclamation
        Exit Sub
    End If

    ' Capture undo snapshot
    CaptureUndoState agg, inv, logTbl

    Dim snapshotId As String: snapshotId = NewGuid()
    Dim entryDate As Date: entryDate = Now

    ' Apply writes
    For r = 1 To UBound(arr, 1)
        Dim tgtRow As Long: tgtRow = NzLng(arr(r, cols("ROW")))
        Dim qty As Double: qty = NzDbl(arr(r, cols("QUANTITY")))
        Dim invRow As ListRow: Set invRow = FindInvRowByROW(inv, tgtRow)
        If invRow Is Nothing Then
            errs = errs & "Row " & r & ": invSys ROW " & tgtRow & " not found" & vbCrLf
            GoTo Bail
        End If
        Dim invRecvCol As Long: invRecvCol = ColumnIndex(inv, "RECEIVED")
        Dim oldVal As Double: oldVal = NzDbl(invRow.Range.Cells(1, invRecvCol).Value)
        RecordInvDelta invRow.Index, oldVal ' for undo
        invRow.Range.Cells(1, invRecvCol).Value = oldVal + qty

        AppendLogRow logTbl, cols, arr, r, snapshotId, entryDate
    Next

    ' Clear staging on success
    ClearTable wsRT.ListObjects("ReceivedTally")
    ClearTable agg
    mRedoReady = True
    Exit Sub

Bail:
    ' On failure, roll back any partial invSys updates and log rows
    UndoInvDeltas wsInv.ListObjects("invSys")
    DeleteAddedLogRows logTbl
    MsgBox "Confirm failed:" & vbCrLf & errs, vbCritical
    Exit Sub

ErrHandler:
    MsgBox "Error in ConfirmWrites: " & Err.Description, vbCritical
    UndoInvDeltas wsInv.ListObjects("invSys")
    DeleteAddedLogRows wsLog.ListObjects("ReceivedLog")
End Sub

Public Sub MacroUndo()
    Dim wsRT As Worksheet: Set wsRT = SheetExists("ReceivedTally")
    Dim wsAgg As Worksheet: Set wsAgg = SheetExists("ReceivedTally")
    Dim wsInv As Worksheet: Set wsInv = SheetExists("InventoryManagement")
    Dim wsLog As Worksheet: Set wsLog = SheetExists("ReceivedLog")
    If wsRT Is Nothing Or wsAgg Is Nothing Or wsInv Is Nothing Or wsLog Is Nothing Then Exit Sub

    RestoreTable wsRT.ListObjects("ReceivedTally"), mUndoRT
    RestoreTable wsAgg.ListObjects("AggregateReceived"), mUndoAGG
    UndoInvDeltas wsInv.ListObjects("invSys")
    DeleteAddedLogRows wsLog.ListObjects("ReceivedLog")
    mRedoReady = True
End Sub

Public Sub MacroRedo()
    If mRedoReady Then
        ConfirmWrites
    End If
End Sub

' ==== helpers ====
Private Function SheetExists(nameOrCode As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set SheetExists = ws
            Exit Function
        End If
    Next
End Function

Private Sub EnsureButton(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Dim topPos As Double: topPos = 10 + ws.Shapes.Count * 20
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, 10, topPos, 100, 18)
        shp.Name = shapeName
        shp.TextFrame.Characters.Text = caption
        If onActionMacro <> "" Then shp.OnAction = onActionMacro
    End If
End Sub

Private Sub MergeIntoReceivedTally(rt As ListObject, refNumber As String, itemName As String, qty As Double)
    If qty <= 0 Then Exit Sub
    Dim colItem As Long: colItem = ColumnIndex(rt, "ITEMS")
    Dim colQty As Long: colQty = ColumnIndex(rt, "QUANTITY")
    Dim colRef As Long: colRef = ColumnIndex(rt, "REF_NUMBER")
    If colItem = 0 Or colQty = 0 Or colRef = 0 Then Exit Sub

    Dim found As Range
    If Not rt.DataBodyRange Is Nothing Then
        Set found = FindInColumn(rt.ListColumns(colItem).DataBodyRange, itemName)
    End If
    If found Is Nothing Then
        Dim lr As ListRow: Set lr = rt.ListRows.Add
        lr.Range.Cells(1, colRef).Value = refNumber
        lr.Range.Cells(1, colItem).Value = itemName
        lr.Range.Cells(1, colQty).Value = qty
    Else
        Dim rIdx As Long: rIdx = found.Row - rt.DataBodyRange.Rows(1).Row + 1
        rt.DataBodyRange.Cells(rIdx, colQty).Value = NzDbl(rt.DataBodyRange.Cells(rIdx, colQty).Value) + qty
        ' concatenate ref numbers
        Dim existingRef As String: existingRef = NzStr(rt.DataBodyRange.Cells(rIdx, colRef).Value)
        If existingRef = "" Then
            rt.DataBodyRange.Cells(rIdx, colRef).Value = refNumber
        ElseIf InStr(1, existingRef, refNumber, vbTextCompare) = 0 Then
            rt.DataBodyRange.Cells(rIdx, colRef).Value = existingRef & "," & refNumber
        End If
    End If
End Sub

Private Sub MergeIntoAggregate(agg As ListObject, refNumber As String, itemCode As String, vendors As String, vendorCode As String, descr As String, itemName As String, uom As String, qty As Double, location As String, invRow As Long)
    Dim c As Object: Set c = AggColMap(agg)
    If c Is Nothing Then Exit Sub

    Dim matchRow As Range
    ' Only try to merge when we have a resolved item code; otherwise always add a new row
    If Not agg.DataBodyRange Is Nothing Then
        If itemCode <> "" Then
            Set matchRow = FindAggregateMatch(agg, itemCode, itemName, uom, vendors, location, invRow, vendorCode, descr)
        Else
            Set matchRow = Nothing
        End If
    End If

    Dim lr As ListRow
    If matchRow Is Nothing Then
        Set lr = agg.ListRows.Add
    Else
        Set lr = agg.ListRows(matchRow.Row - agg.DataBodyRange.Row + 1)
    End If

    With lr.Range
        .Cells(1, c("REF_NUMBER")).Value = AppendRef(NzStr(.Cells(1, c("REF_NUMBER")).Value), refNumber)
        .Cells(1, c("ITEM_CODE")).Value = itemCode
        .Cells(1, c("VENDORS")).Value = vendors
        .Cells(1, c("VENDOR_CODE")).Value = vendorCode
        .Cells(1, c("DESCRIPTION")).Value = descr
        .Cells(1, c("ITEM")).Value = itemName
        .Cells(1, c("UOM")).Value = uom
        .Cells(1, c("LOCATION")).Value = location
        .Cells(1, c("ROW")).Value = invRow
        If qty > 0 Then
            .Cells(1, c("QUANTITY")).Value = NzDbl(.Cells(1, c("QUANTITY")).Value) + qty
        End If
    End With
End Sub

Private Function FindAggregateMatch(agg As ListObject, itemCode As String, itemName As String, uom As String, vendors As String, location As String, invRow As Long, vendorCode As String, descr As String) As Range
    Dim c As Object: Set c = AggColMap(agg)
    If agg.DataBodyRange Is Nothing Then Exit Function
    If invRow <= 0 Then Exit Function ' no resolved invSys row => no merge
    Dim r As Range
    For Each r In agg.DataBodyRange.Rows
        Dim sameKey As Boolean
        sameKey = False
        If itemCode <> "" Then
            sameKey = (NzStr(r.Cells(1, c("ITEM_CODE")).Value) = itemCode And _
                       NzStr(r.Cells(1, c("ITEM")).Value) = itemName)
        Else
            sameKey = (NzStr(r.Cells(1, c("ITEM")).Value) = itemName And NzStr(r.Cells(1, c("UOM")).Value) = uom)
        End If
        If sameKey Then
            ' Require same ROW to merge
            If NzLng(r.Cells(1, c("ROW")).Value) = invRow And _
               NzStr(r.Cells(1, c("LOCATION")).Value) = location And _
               NzStr(r.Cells(1, c("VENDORS")).Value) = vendors Then
                Set FindAggregateMatch = r
                Exit Function
            End If
        End If
    Next
End Function

Private Function AggColMap(lo As ListObject) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim names: names = Array("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW")
    Dim i As Long
    For i = LBound(names) To UBound(names)
        d(names(i)) = ColumnIndex(lo, CStr(names(i)))
        If d(names(i)) = 0 Then Exit Function
    Next
    Set AggColMap = d
End Function

Private Sub AppendLogRow(logTbl As ListObject, cols As Object, arr As Variant, r As Long, snapshotId As String, entryDate As Date)
    Dim newRow As ListRow: Set newRow = logTbl.ListRows.Add
    With logTbl.ListColumns
        newRow.Range(1, .Item("REF_NUMBER").Index).Value = NzStr(arr(r, cols("REF_NUMBER")))
        newRow.Range(1, .Item("ITEMS").Index).Value = NzStr(arr(r, cols("ITEM")))
        newRow.Range(1, .Item("QUANTITY").Index).Value = NzDbl(arr(r, cols("QUANTITY")))
        newRow.Range(1, .Item("UOM").Index).Value = NzStr(arr(r, cols("UOM")))
        newRow.Range(1, .Item("VENDOR").Index).Value = NzStr(arr(r, cols("VENDORS")))
        newRow.Range(1, .Item("LOCATION").Index).Value = NzStr(arr(r, cols("LOCATION")))
        newRow.Range(1, .Item("ITEM_CODE").Index).Value = NzStr(arr(r, cols("ITEM_CODE")))
        newRow.Range(1, .Item("ROW").Index).Value = NzLng(arr(r, cols("ROW")))
        newRow.Range(1, .Item("SNAPSHOT_ID").Index).Value = snapshotId
        newRow.Range(1, .Item("ENTRY_DATE").Index).Value = entryDate
    End With
    If mUndoLogRows Is Nothing Then Set mUndoLogRows = New Collection
    mUndoLogRows.Add newRow.Index
End Sub

Private Function FindInvRowByROW(inv As ListObject, rowValue As Long) As ListRow
    Dim cRow As Long: cRow = ColumnIndex(inv, "ROW")
    If cRow = 0 Or inv.DataBodyRange Is Nothing Then Exit Function
    Dim cel As Range
    For Each cel In inv.ListColumns(cRow).DataBodyRange.Cells
        If NzLng(cel.Value) = rowValue Then
            Set FindInvRowByROW = inv.ListRows(cel.Row - inv.DataBodyRange.Row + 1)
            Exit Function
        End If
    Next
End Function

Private Sub ClearTable(lo As ListObject)
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.Delete
        End If
    End If
End Sub

' ===== undo helpers =====
Private Sub CaptureUndoState(agg As ListObject, inv As ListObject, logTbl As ListObject)
    Set mUndoInv = New Collection
    Set mUndoLogRows = New Collection
    mUndoRT = SnapshotTable(SheetExists("ReceivedTally").ListObjects("ReceivedTally"))
    mUndoAGG = SnapshotTable(agg)
    ' log rows added will be captured as we append
End Sub

Private Function SnapshotTable(lo As ListObject) As Variant
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        SnapshotTable = Empty
    Else
        SnapshotTable = lo.DataBodyRange.Value
    End If
End Function

Private Sub RestoreTable(lo As ListObject, snap As Variant)
    ClearTable lo
    If IsEmpty(snap) Then Exit Sub
    Dim r As Long, c As Long
    Dim rows As Long: rows = UBound(snap, 1)
    Dim cols As Long: cols = UBound(snap, 2)
    lo.Resize lo.Range.Resize(rows + 1, cols)
    lo.DataBodyRange.Value = snap
End Sub

Private Sub RecordInvDelta(rowIndex As Long, oldVal As Double)
    ' Store simple variant array to avoid UDT/collection coercion issues
    Dim arr(1 To 2) As Variant
    arr(1) = rowIndex
    arr(2) = oldVal
    If mUndoInv Is Nothing Then Set mUndoInv = New Collection
    mUndoInv.Add arr
End Sub

Private Sub UndoInvDeltas(inv As ListObject)
    If mUndoInv Is Nothing Then Exit Sub
    Dim v As Variant
    Dim recvCol As Long: recvCol = ColumnIndex(inv, "RECEIVED")
    For Each v In mUndoInv
        inv.ListRows(CLng(v(1))).Range.Cells(1, recvCol).Value = CDbl(v(2))
    Next
End Sub

Private Sub DeleteAddedLogRows(logTbl As ListObject)
    If mUndoLogRows Is Nothing Then Exit Sub
    Dim idx As Variant
    ' delete from bottom to top
    Dim arr() As Long
    ReDim arr(1 To mUndoLogRows.Count)
    Dim i As Long
    For i = 1 To mUndoLogRows.Count
        arr(i) = CLng(mUndoLogRows(i))
    Next
    QuickSort arr, LBound(arr), UBound(arr)
    For i = UBound(arr) To LBound(arr) Step -1
        If arr(i) <= logTbl.ListRows.Count Then logTbl.ListRows(arr(i)).Delete
    Next
End Sub

Private Sub QuickSort(a() As Long, lo As Long, hi As Long)
    Dim i As Long, j As Long, p As Long, tmp As Long
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSort a, lo, j
    If i < hi Then QuickSort a, i, hi
End Sub

' ===== misc helpers =====
Private Function FindInColumn(rng As Range, value As String) As Range
    Dim cel As Range
    For Each cel In rng.Cells
        If StrComp(NzStr(cel.Value), value, vbTextCompare) = 0 Then
            Set FindInColumn = cel
            Exit Function
        End If
    Next
End Function

Private Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Private Function NzDbl(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzDbl = 0#
    Else
        NzDbl = CDbl(v)
    End If
End Function

Private Function NzLng(v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzLng = 0
    Else
        NzLng = CLng(v)
    End If
End Function

Private Function AppendRef(existingRef As String, newRef As String) As String
    If existingRef = "" Then AppendRef = newRef: Exit Function
    If InStr(1, existingRef, newRef, vbTextCompare) > 0 Then
        AppendRef = existingRef
    Else
        AppendRef = existingRef & "," & newRef
    End If
End Function

Private Function LookupInvSys(catalog As ListObject, itemName As String, ByRef itemCode As String, ByRef vendors As String, ByRef vendorCode As String, ByRef descr As String, ByRef uom As String, ByRef location As String, ByRef invRow As Long)
    Dim cItem As Long: cItem = ColumnIndex(catalog, "ITEM")
    Dim cCode As Long: cCode = ColumnIndex(catalog, "ITEM_CODE")
    Dim cVend As Long: cVend = ColumnIndex(catalog, "VENDOR(s)")
    Dim cLoc As Long: cLoc = ColumnIndex(catalog, "LOCATION")
    Dim cDesc As Long: cDesc = ColumnIndex(catalog, "DESCRIPTION")
    Dim cUOM As Long: cUOM = ColumnIndex(catalog, "UOM")
    Dim cRow As Long: cRow = ColumnIndex(catalog, "ROW")
    If cItem = 0 Or cCode = 0 Then Exit Function
    Dim rng As Range: Set rng = catalog.ListColumns(IIf(itemCode <> "", cCode, cItem)).DataBodyRange
    Dim cel As Range
    For Each cel In rng.Cells
        If StrComp(NzStr(cel.Value), IIf(itemCode <> "", itemCode, itemName), vbTextCompare) = 0 Then
            invRow = NzLng(cel.Offset(0, cRow - cel.Column).Value)
            itemCode = NzStr(cel.Offset(0, cCode - cel.Column).Value)
            itemName = NzStr(cel.Offset(0, cItem - cel.Column).Value)
            vendors = NzStr(cel.Offset(0, cVend - cel.Column).Value)
            vendorCode = "" ' not in catalog headers; left blank
            descr = NzStr(cel.Offset(0, cDesc - cel.Column).Value)
            uom = NzStr(cel.Offset(0, cUOM - cel.Column).Value)
            location = NzStr(cel.Offset(0, cLoc - cel.Column).Value)
            Exit Function
        End If
    Next
End Function

Private Function NewGuid() As String
    NewGuid = CreateObject("Scriptlet.TypeLib").Guid
End Function

' Column index helper (case-insensitive) on a ListObject
Private Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next
    ColumnIndex = 0
End Function

' Load item list for frmItemSearch from invSys (InventoryManagement!invSys)
' Returns a 2D array with columns: ITEM_CODE, ITEM, UOM, LOCATION
Public Function LoadItemList() As Variant
    Dim ws As Worksheet: Set ws = SheetExists("InventoryManagement")
    If ws Is Nothing Then Exit Function
    Dim lo As ListObject: Set lo = ws.ListObjects("invSys")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    Dim cCode As Long, cItem As Long, cUOM As Long, cLoc As Long
    cCode = ColumnIndex(lo, "ITEM_CODE")
    cItem = ColumnIndex(lo, "ITEM")
    cUOM = ColumnIndex(lo, "UOM")
    cLoc = ColumnIndex(lo, "LOCATION")
    If cCode * cItem = 0 Then Exit Function

    Dim src As Variant: src = lo.DataBodyRange.Value
    Dim r As Long, n As Long: n = UBound(src, 1)
    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 4)
    Dim outRow As Long: outRow = 0

    For r = 1 To n
        Dim itm As String: itm = NzStr(src(r, cItem))
        If itm <> "" Then
            outRow = outRow + 1
            outArr(outRow, 1) = NzStr(src(r, cCode)) ' ITEM_CODE
            outArr(outRow, 2) = itm                  ' ITEM
            outArr(outRow, 3) = NzStr(src(r, cUOM))  ' UOM
            outArr(outRow, 4) = NzStr(src(r, cLoc))  ' LOCATION
        End If
    Next

    If outRow = 0 Then Exit Function
    ' Trim to actual count
    ReDim Preserve outArr(1 To outRow, 1 To 4)
    LoadItemList = outArr
End Function
