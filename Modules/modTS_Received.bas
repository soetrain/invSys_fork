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
Private mDynSearch As cDynItemSearch
Private mRowMap As Object ' maps staging row number -> Array(invRow, refNumber)

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

' ==== dynamic search form (ReceivedTally) =====
Public Sub ShowDynamicItemSearch(ByVal targetCell As Range)
    Debug.Print "ShowDynamicItemSearch called, target:", targetCell.Address
    
    If targetCell Is Nothing Then
        Debug.Print "  targetCell is Nothing, exiting"
        Exit Sub
    End If

    If mDynSearch Is Nothing Then
        Debug.Print "  mDynSearch is Nothing, creating new cDynItemSearch"
        On Error GoTo ErrHandler
        Set mDynSearch = New cDynItemSearch
        Debug.Print "  New cDynItemSearch succeeded"
    Else
        Debug.Print "  Reusing existing cDynItemSearch instance"
    End If

    Debug.Print "  Calling mDynSearch.ShowForCell"
    mDynSearch.ShowForCell targetCell
    Debug.Print "  Returned from ShowForCell"
    Exit Sub

ErrHandler:
    Debug.Print "  ERROR creating cDynItemSearch:", Err.Number, Err.Description
    Debug.Print "  Falling back to frmItemSearch"
    On Error Resume Next
    frmItemSearch.Show vbModeless
End Sub

' =========================
' Confirm Writes sub-system
' -------------------------
' - AggregateReceived already holds the summed QUANTITY per invSys ROW and concatenated REF_NUMBER for display.
' - ConfirmWrites uses AGG.QUANTITY directly to add into invSys.RECEIVED (no recompute, no overwrite of AGG).
' - ReceivedLog is per REF: REF/ITEM/QUANTITY from staging; ROW/UOM/LOCATION from AggregateReceived; SNAPSHOT_ID/ENTRY_DATE generated.
' - AggregateReceived is treated as read-only for the user; code clears it only after a successful Confirm.
' =========================

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
    ByVal invRow As Long, _
    Optional ByVal stagingRow As Long = 0)

    Dim ws As Worksheet: Set ws = SheetExists("ReceivedTally")
    If ws Is Nothing Then Exit Sub

    Dim rt As ListObject, agg As ListObject
    Set rt = ws.ListObjects("ReceivedTally")
    Set agg = ws.ListObjects("AggregateReceived")

    ' Insert/merge into ReceivedTally (fast entry: ITEMS + QUANTITY + REF_NUMBER)
    MergeIntoReceivedTally rt, refNumber, itemName, qty

    ' Insert/merge into AggregateReceived (detailed)
    MergeIntoAggregate agg, refNumber, itemCode, vendors, vendorCode, descr, itemName, uom, qty, location, invRow

    ' Track the invRow for this staging row so quantity edits can sync correctly
    If stagingRow > 0 Then
        EnsureRowMap
        mRowMap(CStr(stagingRow)) = Array(invRow, refNumber)
    End If
End Sub

Public Sub RebuildAggregation()
    Dim ws As Worksheet: Set ws = SheetExists("ReceivedTally")
    If ws Is Nothing Then Exit Sub
    Dim rt As ListObject: Set rt = ws.ListObjects("ReceivedTally")
    Dim agg As ListObject: Set agg = ws.ListObjects("AggregateReceived")
    If rt Is Nothing Or agg Is Nothing Then Exit Sub
    ClearTable agg

    If rt.DataBodyRange Is Nothing Then Exit Sub
    ' If staging has no ROW column, we cannot rebuild by ROW; skip quietly
    Dim cRowRT As Long
    cRowRT = ColumnIndex(rt, "ROW")
    If cRowRT = 0 Then
        Debug.Print "RebuildAggregation: staging has no ROW column; skipped."
        Exit Sub
    End If

    Dim arr, r As Long
    arr = rt.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        Dim itemName As String, qty As Double
        itemName = NzStr(arr(r, ColumnIndex(rt, "ITEMS")))
        qty = NzDbl(arr(r, ColumnIndex(rt, "QUANTITY")))
        Dim refNumber As String
        refNumber = NzStr(arr(r, ColumnIndex(rt, "REF_NUMBER")))

        ' If the staging row has a ROW column, use that exact invSys row; otherwise skip
        Dim invRow As Long
        invRow = NzLng(arr(r, cRowRT))
        If invRow > 0 Then
            ' We still need catalog details to display; fetch strictly by ROW
            Dim itemCode As String, vendors As String, vendorCode As String
            Dim descr As String, uom As String, location As String
            Dim catWs As Worksheet: Set catWs = SheetExists("InventoryManagement")
            Dim catLo As ListObject
            If Not catWs Is Nothing Then Set catLo = catWs.ListObjects("invSys")
            LookupInvSysByROW catLo, invRow, itemCode, vendors, vendorCode, descr, itemName, uom, location
            MergeIntoAggregate agg, refNumber, itemCode, vendors, vendorCode, descr, itemName, uom, qty, location, invRow
        End If
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
    arr = agg.DataBodyRange.value
    Dim cols As Object: Set cols = AggColMap(agg)
    Dim refNumRT As String
    Dim itemRT As String
    Dim qtyRT As Double

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

    ' Build ref -> (ROW, UOM, LOCATION, ITEM) map from AggregateReceived for accurate logging
    Dim refMap As Object: Set refMap = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(arr, 1)
        Dim mapRow As Long: mapRow = NzLng(arr(r, cols("ROW")))
        If mapRow > 0 Then
            Dim mapRefs As Variant
            mapRefs = Split(NzStr(arr(r, cols("REF_NUMBER"))), ",")
            Dim mapUOM As String: mapUOM = NzStr(arr(r, cols("UOM")))
            Dim mapLoc As String: mapLoc = NzStr(arr(r, cols("LOCATION")))
            Dim mapItem As String: mapItem = NzStr(arr(r, cols("ITEM")))
            Dim rf As Variant
            For Each rf In mapRefs
                rf = Trim(CStr(rf))
                If rf <> "" Then
                    refMap(rf) = Array(mapRow, mapUOM, mapLoc, mapItem)
                End If
            Next rf
        End If
    Next

    ' Apply writes to invSys (per aggregated row)
    For r = 1 To UBound(arr, 1)
        Dim tgtRow As Long: tgtRow = NzLng(arr(r, cols("ROW")))
        Dim qty As Double: qty = NzDbl(arr(r, cols("QUANTITY")))
        Dim invRow As ListRow: Set invRow = FindInvRowByROW(inv, tgtRow)
        If invRow Is Nothing Then
            errs = errs & "Row " & r & ": invSys ROW " & tgtRow & " not found" & vbCrLf
            GoTo Bail
        End If
        Dim invRecvCol As Long: invRecvCol = ColumnIndex(inv, "RECEIVED")
        Dim oldVal As Double: oldVal = NzDbl(invRow.Range.Cells(1, invRecvCol).value)
        RecordInvDelta invRow.Index, oldVal ' for undo
        invRow.Range.Cells(1, invRecvCol).value = oldVal + qty
    Next

    ' Log per REF_NUMBER using staging (ReceivedTally) quantities + ROW/UOM/LOCATION from refMap
    ' staging table
    Dim rt As ListObject: Set rt = wsRT.ListObjects("ReceivedTally")
    If Not rt Is Nothing And Not rt.DataBodyRange Is Nothing Then
        Dim rtArr As Variant: rtArr = rt.DataBodyRange.value
        Dim rtCols As Object: Set rtCols = CreateObject("Scripting.Dictionary")
        rtCols("REF_NUMBER") = ColumnIndex(rt, "REF_NUMBER")
        rtCols("ITEMS") = ColumnIndex(rt, "ITEMS")
        rtCols("QUANTITY") = ColumnIndex(rt, "QUANTITY")
        Dim rrt As Long
        For rrt = 1 To UBound(rtArr, 1)
            refNumRT = Trim$(NzStr(rtArr(rrt, rtCols("REF_NUMBER"))))
            itemRT = NzStr(rtArr(rrt, rtCols("ITEMS")))
            qtyRT = NzDbl(rtArr(rrt, rtCols("QUANTITY")))
            If refNumRT = "" And itemRT = "" And qtyRT = 0 Then GoTo NextRt

            Dim logRow As Long, logUOM As String, logLoc As String, logItem As String
            If refMap.Exists(refNumRT) Then
                Dim mArr As Variant: mArr = refMap(refNumRT)
                logRow = NzLng(mArr(0))
                logUOM = NzStr(mArr(1))
                logLoc = NzStr(mArr(2))
                logItem = NzStr(mArr(3))
                If logItem = "" Then logItem = itemRT
            Else
                ' fallback: try lookup by item name
                Dim tmpCode As String, tmpVend As String, tmpVCode As String, tmpDesc As String
                LookupInvSys wsInv.ListObjects("invSys"), itemRT, tmpCode, tmpVend, tmpVCode, tmpDesc, logUOM, logLoc, logRow
                logItem = itemRT
            End If

            AppendLogRowFromRT logTbl, refNumRT, logItem, qtyRT, logUOM, logLoc, logRow, snapshotId, entryDate
NextRt:
        Next rrt
    End If

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
    ' Undo last successful ConfirmWrites
    Dim wsRT As Worksheet: Set wsRT = SheetExists("ReceivedTally")
    Dim wsAgg As Worksheet: Set wsAgg = SheetExists("ReceivedTally")
    Dim wsInv As Worksheet: Set wsInv = SheetExists("InventoryManagement")
    Dim wsLog As Worksheet: Set wsLog = SheetExists("ReceivedLog")
    If wsRT Is Nothing Or wsAgg Is Nothing Or wsInv Is Nothing Or wsLog Is Nothing Then Exit Sub
    ' Guard: do we have anything to undo?
    Dim hasUndo As Boolean
    hasUndo = Not IsEmpty(mUndoRT) Or Not IsEmpty(mUndoAGG)
    If mUndoInv Is Nothing Then
        hasUndo = hasUndo Or False
    Else
        hasUndo = hasUndo Or (mUndoInv.count > 0)
    End If
    If mUndoLogRows Is Nothing Then
        hasUndo = hasUndo Or False
    Else
        hasUndo = hasUndo Or (mUndoLogRows.count > 0)
    End If
    If Not hasUndo Then
        MsgBox "Nothing to undo (no confirm snapshot).", vbInformation
        Exit Sub
    End If

    Application.EnableEvents = False
    RestoreTable wsRT.ListObjects("ReceivedTally"), mUndoRT
    RestoreTable wsAgg.ListObjects("AggregateReceived"), mUndoAGG
    UndoInvDeltas wsInv.ListObjects("invSys")
    DeleteAddedLogRows wsLog.ListObjects("ReceivedLog")
    Application.EnableEvents = True
    mRedoReady = True
End Sub

Public Sub MacroRedo()
    If Not mRedoReady Then
        MsgBox "Nothing to redo. Perform an undo first.", vbInformation
        Exit Sub
    End If
    ConfirmWrites
End Sub

' ==== helpers ====
Private Function SheetExists(nameOrCode As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set SheetExists = ws
            Exit Function
        End If
    Next
End Function

Private Sub EnsureButton(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Dim topPos As Double: topPos = 10 + ws.shapes.count * 20
        Set shp = ws.shapes.AddFormControl(xlButtonControl, 10, topPos, 100, 18)
        shp.name = shapeName
        shp.TextFrame.Characters.text = caption
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
        lr.Range.Cells(1, colRef).value = refNumber
        lr.Range.Cells(1, colItem).value = itemName
        lr.Range.Cells(1, colQty).value = qty
    Else
        Dim rIdx As Long: rIdx = found.row - rt.DataBodyRange.rows(1).row + 1
        rt.DataBodyRange.Cells(rIdx, colQty).value = NzDbl(rt.DataBodyRange.Cells(rIdx, colQty).value) + qty
        ' concatenate ref numbers
        Dim existingRef As String: existingRef = NzStr(rt.DataBodyRange.Cells(rIdx, colRef).value)
        If existingRef = "" Then
            rt.DataBodyRange.Cells(rIdx, colRef).value = refNumber
        ElseIf InStr(1, existingRef, refNumber, vbTextCompare) = 0 Then
            rt.DataBodyRange.Cells(rIdx, colRef).value = existingRef & "," & refNumber
        End If
    End If
End Sub

Private Sub MergeIntoAggregate(agg As ListObject, refNumber As String, itemCode As String, vendors As String, vendorCode As String, descr As String, itemName As String, uom As String, qty As Double, location As String, invRow As Long)
    Dim c As Object: Set c = AggColMap(agg)
    If c Is Nothing Then Exit Sub
    If invRow <= 0 Then Exit Sub ' must have resolved invSys row to merge

    Dim matchLR As ListRow
    Set matchLR = FindAggregateMatchByRow(agg, invRow)

    Dim lr As ListRow
    If matchLR Is Nothing Then
        Set lr = agg.ListRows.Add
    Else
        Set lr = matchLR
    End If

    With lr.Range
        .Cells(1, c("REF_NUMBER")).value = AppendRef(NzStr(.Cells(1, c("REF_NUMBER")).value), refNumber)
        .Cells(1, c("ITEM_CODE")).value = itemCode
        .Cells(1, c("VENDORS")).value = vendors
        .Cells(1, c("VENDOR_CODE")).value = vendorCode
        .Cells(1, c("DESCRIPTION")).value = descr
        .Cells(1, c("ITEM")).value = itemName
        .Cells(1, c("UOM")).value = uom
        .Cells(1, c("LOCATION")).value = location
        .Cells(1, c("ROW")).value = invRow
        If qty > 0 Then
            .Cells(1, c("QUANTITY")).value = NzDbl(.Cells(1, c("QUANTITY")).value) + qty
        End If
    End With
End Sub

Private Function FindAggregateMatchByRow(agg As ListObject, invRow As Long) As ListRow
    If agg Is Nothing Or agg.DataBodyRange Is Nothing Then Exit Function
    Dim cRow As Long
    cRow = ColumnIndex(agg, "ROW")
    If cRow = 0 Then Exit Function

    Dim lr As ListRow
    For Each lr In agg.ListRows
        If NzLng(lr.Range.Cells(1, cRow).value) = invRow Then
            Set FindAggregateMatchByRow = lr
            Exit Function
        End If
    Next lr
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
    ' Write only columns that exist in ReceivedLog (current headers: SNAPSHOT_ID, ENTRY_DATE, REF_NUMBER, ITEMS, QUANTITY, UOM, ROW, LOCATION)
    Dim cRef As Long, cItems As Long, cQty As Long, cUOM As Long
    Dim cRow As Long, cLoc As Long, cSnap As Long, cEntry As Long
    cRef = LogColIndex(logTbl, "REF_NUMBER")
    cItems = LogColIndex(logTbl, "ITEMS")
    cQty = LogColIndex(logTbl, "QUANTITY")
    cUOM = LogColIndex(logTbl, "UOM")
    cRow = LogColIndex(logTbl, "ROW")
    cLoc = LogColIndex(logTbl, "LOCATION")
    cSnap = LogColIndex(logTbl, "SNAPSHOT_ID")
    cEntry = LogColIndex(logTbl, "ENTRY_DATE")

    With newRow.Range
        If cRef > 0 Then .Cells(1, cRef).value = NzStr(arr(r, cols("REF_NUMBER")))
        If cItems > 0 Then .Cells(1, cItems).value = NzStr(arr(r, cols("ITEM")))
        If cQty > 0 Then .Cells(1, cQty).value = NzDbl(arr(r, cols("QUANTITY")))
        If cUOM > 0 Then .Cells(1, cUOM).value = NzStr(arr(r, cols("UOM")))
        If cRow > 0 Then .Cells(1, cRow).value = NzLng(arr(r, cols("ROW")))
        If cLoc > 0 Then .Cells(1, cLoc).value = NzStr(arr(r, cols("LOCATION")))
        If cSnap > 0 Then .Cells(1, cSnap).value = snapshotId
        If cEntry > 0 Then .Cells(1, cEntry).value = entryDate
    End With
    If mUndoLogRows Is Nothing Then Set mUndoLogRows = New Collection
    mUndoLogRows.Add newRow.Index
End Sub

Private Sub AppendLogRowFromRT(logTbl As ListObject, ByVal refNum As String, ByVal itemName As String, ByVal qty As Double, ByVal uom As String, ByVal location As String, ByVal invRow As Long, ByVal snapshotId As String, ByVal entryDate As Date)
    If logTbl Is Nothing Then Exit Sub
    Dim newRow As ListRow: Set newRow = logTbl.ListRows.Add
    Dim cRef As Long, cItems As Long, cQty As Long, cUOM As Long
    Dim cRow As Long, cLoc As Long, cSnap As Long, cEntry As Long
    cRef = LogColIndex(logTbl, "REF_NUMBER")
    cItems = LogColIndex(logTbl, "ITEMS")
    cQty = LogColIndex(logTbl, "QUANTITY")
    cUOM = LogColIndex(logTbl, "UOM")
    cRow = LogColIndex(logTbl, "ROW")
    cLoc = LogColIndex(logTbl, "LOCATION")
    cSnap = LogColIndex(logTbl, "SNAPSHOT_ID")
    cEntry = LogColIndex(logTbl, "ENTRY_DATE")

    With newRow.Range
        If cRef > 0 Then .Cells(1, cRef).value = refNum
        If cItems > 0 Then .Cells(1, cItems).value = itemName
        If cQty > 0 Then .Cells(1, cQty).value = qty
        If cUOM > 0 Then .Cells(1, cUOM).value = uom
        If cRow > 0 Then .Cells(1, cRow).value = invRow
        If cLoc > 0 Then .Cells(1, cLoc).value = location
        If cSnap > 0 Then .Cells(1, cSnap).value = snapshotId
        If cEntry > 0 Then .Cells(1, cEntry).value = entryDate
    End With
    If mUndoLogRows Is Nothing Then Set mUndoLogRows = New Collection
    mUndoLogRows.Add newRow.Index
End Sub

Private Function FindInvRowByROW(inv As ListObject, rowValue As Long) As ListRow
    Dim cRow As Long: cRow = ColumnIndex(inv, "ROW")
    If cRow = 0 Or inv.DataBodyRange Is Nothing Then Exit Function
    Dim cel As Range
    For Each cel In inv.ListColumns(cRow).DataBodyRange.Cells
        If NzLng(cel.value) = rowValue Then
            Set FindInvRowByROW = inv.ListRows(cel.row - inv.DataBodyRange.row + 1)
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
    ' Clear row map if we clear staging
    If lo Is Nothing Then Exit Sub
    If StrComp(lo.name, "ReceivedTally", vbTextCompare) = 0 Then
        If Not mRowMap Is Nothing Then mRowMap.RemoveAll
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
        SnapshotTable = lo.DataBodyRange.value
    End If
End Function

Private Sub RestoreTable(lo As ListObject, snap As Variant)
    ClearTable lo
    If IsEmpty(snap) Then Exit Sub
    Dim r As Long, c As Long
    Dim rows As Long: rows = UBound(snap, 1)
    Dim cols As Long: cols = UBound(snap, 2)
    lo.Resize lo.Range.Resize(rows + 1, cols)
    lo.DataBodyRange.value = snap
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
        inv.ListRows(CLng(v(1))).Range.Cells(1, recvCol).value = CDbl(v(2))
    Next
End Sub

Private Sub DeleteAddedLogRows(logTbl As ListObject)
    If mUndoLogRows Is Nothing Then Exit Sub
    If mUndoLogRows.count = 0 Then Exit Sub
    Dim idx As Variant
    ' delete from bottom to top
    Dim arr() As Long
    ReDim arr(1 To mUndoLogRows.count)
    Dim i As Long
    For i = 1 To mUndoLogRows.count
        arr(i) = CLng(mUndoLogRows(i))
    Next
    QuickSort arr, LBound(arr), UBound(arr)
    For i = UBound(arr) To LBound(arr) Step -1
        If arr(i) <= logTbl.ListRows.count Then logTbl.ListRows(arr(i)).Delete
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

' ===== log column tools (optional columns in ReceivedLog) =====
Private Function CriticalLogCols() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim names
    names = Array("REF_NUMBER", "ITEMS", "QUANTITY", "UOM", "VENDOR", "LOCATION", "ITEM_CODE", "ROW", "SNAPSHOT_ID", "ENTRY_DATE")
    Dim i As Long
    For i = LBound(names) To UBound(names)
        d.Add names(i), True
    Next
    Set CriticalLogCols = d
End Function

Public Sub ToggleLogColumn(ByVal colName As String, ByVal enable As Boolean)
    colName = Trim$(colName)
    If colName = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = SheetExists("ReceivedLog")
    If ws Is Nothing Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects("ReceivedLog")
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    Dim crit As Object
    Set crit = CriticalLogCols()

    Dim idx As Long
    idx = ColumnIndex(lo, colName)

    If enable Then
        If idx = 0 Then
            Dim newCol As ListColumn
            Set newCol = lo.ListColumns.Add
            newCol.name = colName
        End If
    Else
        If crit.Exists(colName) Then
            MsgBox colName & " is critical and cannot be removed.", vbInformation
            Exit Sub
        End If
        If idx > 0 Then lo.ListColumns(idx).Delete
    End If
End Sub

' ===== misc helpers =====
Private Function FindInColumn(rng As Range, value As String) As Range
    Dim cel As Range
    For Each cel In rng.Cells
        If StrComp(NzStr(cel.value), value, vbTextCompare) = 0 Then
            Set FindInColumn = cel
            Exit Function
        End If
    Next
End Function

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

Private Function AppendRef(existingRef As String, newRef As String) As String
    If existingRef = "" Then AppendRef = newRef: Exit Function
    If InStr(1, existingRef, newRef, vbTextCompare) > 0 Then
        AppendRef = existingRef
    Else
        AppendRef = existingRef & "," & newRef
    End If
End Function

Private Function LookupInvSys(catalog As ListObject, itemName As String, ByRef itemCode As String, ByRef vendors As String, ByRef vendorCode As String, ByRef descr As String, ByRef uom As String, ByRef location As String, ByRef invRow As Long)
    ' NOTE: this lookup is used when we do not already have ROW.
    ' defaults
    itemCode = "": vendors = "": vendorCode = "": descr = "": uom = "": location = "": invRow = 0

    Dim cItem As Long: cItem = ColumnIndex(catalog, "ITEM")
    Dim cCode As Long: cCode = ColumnIndex(catalog, "ITEM_CODE")
    Dim cVend As Long: cVend = ColumnIndex(catalog, "VENDOR(s)")
    Dim cLoc As Long: cLoc = ColumnIndex(catalog, "LOCATION")
    Dim cDesc As Long: cDesc = ColumnIndex(catalog, "DESCRIPTION")
    Dim cUOM As Long: cUOM = ColumnIndex(catalog, "UOM")
    Dim cRow As Long: cRow = ColumnIndex(catalog, "ROW")
    If cItem = 0 Or cCode = 0 Then Exit Function

    ' Try exact match by code first
    If itemCode <> "" Then
        Dim found As Range
        Set found = FindInColumn(catalog.ListColumns(cCode).DataBodyRange, itemCode)
        If Not found Is Nothing Then
            invRow = NzLng(found.Offset(0, cRow - found.Column).value)
            itemCode = NzStr(found.Offset(0, cCode - found.Column).value)
            itemName = NzStr(found.Offset(0, cItem - found.Column).value)
            vendors = NzStr(found.Offset(0, cVend - found.Column).value)
            descr = NzStr(found.Offset(0, cDesc - found.Column).value)
            uom = NzStr(found.Offset(0, cUOM - found.Column).value)
            location = NzStr(found.Offset(0, cLoc - found.Column).value)
            Exit Function
        End If
    End If

    ' Then try exact match by name
    If itemName <> "" Then
        Dim found2 As Range
        Set found2 = FindInColumn(catalog.ListColumns(cItem).DataBodyRange, itemName)
        If Not found2 Is Nothing Then
            invRow = NzLng(found2.Offset(0, cRow - found2.Column).value)
            itemCode = NzStr(found2.Offset(0, cCode - found2.Column).value)
            itemName = NzStr(found2.Offset(0, cItem - found2.Column).value)
            vendors = NzStr(found2.Offset(0, cVend - found2.Column).value)
            descr = NzStr(found2.Offset(0, cDesc - found2.Column).value)
            uom = NzStr(found2.Offset(0, cUOM - found2.Column).value)
            location = NzStr(found2.Offset(0, cLoc - found2.Column).value)
            Exit Function
        End If
    End If
End Function

' Strict lookup by ROW when we already know the invSys ROW
Private Sub LookupInvSysByROW(catalog As ListObject, ByVal invRow As Long, _
    ByRef itemCode As String, ByRef vendors As String, ByRef vendorCode As String, _
    ByRef descr As String, ByRef itemName As String, ByRef uom As String, ByRef location As String)

    itemCode = "": vendors = "": vendorCode = "": descr = "": itemName = "": uom = "": location = ""
    If catalog Is Nothing Or invRow <= 0 Then Exit Sub
    If catalog.DataBodyRange Is Nothing Then Exit Sub

    Dim cRow As Long: cRow = ColumnIndex(catalog, "ROW")
    Dim cCode As Long: cCode = ColumnIndex(catalog, "ITEM_CODE")
    Dim cItem As Long: cItem = ColumnIndex(catalog, "ITEM")
    Dim cVend As Long: cVend = ColumnIndex(catalog, "VENDOR(s)")
    Dim cDesc As Long: cDesc = ColumnIndex(catalog, "DESCRIPTION")
    Dim cUOM As Long: cUOM = ColumnIndex(catalog, "UOM")
    Dim cLoc As Long: cLoc = ColumnIndex(catalog, "LOCATION")
    If cRow = 0 Then Exit Sub

    Dim cel As Range
    For Each cel In catalog.ListColumns(cRow).DataBodyRange.Cells
        If NzLng(cel.value) = invRow Then
            If cCode > 0 Then itemCode = NzStr(cel.Offset(0, cCode - cel.Column).value)
            If cItem > 0 Then itemName = NzStr(cel.Offset(0, cItem - cel.Column).value)
            If cVend > 0 Then vendors = NzStr(cel.Offset(0, cVend - cel.Column).value)
            If cDesc > 0 Then descr = NzStr(cel.Offset(0, cDesc - cel.Column).value)
            If cUOM > 0 Then uom = NzStr(cel.Offset(0, cUOM - cel.Column).value)
            If cLoc > 0 Then location = NzStr(cel.Offset(0, cLoc - cel.Column).value)
            Exit Sub
        End If
    Next
End Sub

Private Function NewGuid() As String
    NewGuid = CreateObject("Scriptlet.TypeLib").GUID
End Function

' Column index helper for log tables (case-insensitive)
Private Function LogColIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.name, colName, vbTextCompare) = 0 Then
            LogColIndex = lc.Index
            Exit Function
        End If
    Next
    LogColIndex = 0
End Function

' Column index helper (case-insensitive) on a ListObject
Private Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.name, colName, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next
    ColumnIndex = 0
End Function

' Maintain quantity sync from staging (ReceivedTally) to AggregateReceived
Public Sub SyncQuantityFromStaging(ByVal stagingRowIdx As Long, ByVal newQty As Double)
    If stagingRowIdx <= 0 Then Exit Sub
    If mRowMap Is Nothing Then Exit Sub
    If Not mRowMap.Exists(CStr(stagingRowIdx)) Then Exit Sub

    ' Identify invSys ROW for this staging row
    Dim info As Variant
    info = mRowMap(CStr(stagingRowIdx)) ' Array(invRow, refNumber)
    Dim invRow As Long: invRow = CLng(info(0))

    ' Sum all staging quantities that map to the same invSys ROW
    Dim wsRT As Worksheet: Set wsRT = SheetExists("ReceivedTally")
    If wsRT Is Nothing Then Exit Sub
    Dim rt As ListObject: Set rt = wsRT.ListObjects("ReceivedTally")
    If rt Is Nothing Or rt.DataBodyRange Is Nothing Then Exit Sub
    Dim colQtyRT As Long: colQtyRT = ColumnIndex(rt, "QUANTITY")
    If colQtyRT = 0 Then Exit Sub

    Dim totalQty As Double: totalQty = 0
    Dim k As Variant
    For Each k In mRowMap.Keys
        Dim arr As Variant
        arr = mRowMap(k) ' invRow, refNumber
        If CLng(arr(0)) = invRow Then
            Dim sr As Long
            sr = CLng(k)
            If sr >= 1 And sr <= rt.DataBodyRange.rows.count Then
                totalQty = totalQty + NzDbl(rt.DataBodyRange.Cells(sr, colQtyRT).value)
            End If
        End If
    Next k

    ' Update the single aggregate row for this invRow
    Dim wsAgg As Worksheet: Set wsAgg = SheetExists("ReceivedTally")
    If wsAgg Is Nothing Then Exit Sub
    Dim agg As ListObject: Set agg = wsAgg.ListObjects("AggregateReceived")
    If agg Is Nothing Or agg.DataBodyRange Is Nothing Then Exit Sub

    Dim cRowAgg As Long: cRowAgg = ColumnIndex(agg, "ROW")
    Dim cQtyAgg As Long: cQtyAgg = ColumnIndex(agg, "QUANTITY")
    If cRowAgg = 0 Or cQtyAgg = 0 Then Exit Sub

    Dim lr As ListRow
    For Each lr In agg.ListRows
        If NzLng(lr.Range.Cells(1, cRowAgg).value) = invRow Then
            lr.Range.Cells(1, cQtyAgg).value = totalQty
            Exit For
        End If
    Next lr
End Sub

Private Sub EnsureRowMap()
    If mRowMap Is Nothing Then Set mRowMap = CreateObject("Scripting.Dictionary")
End Sub

' Load item list for frmItemSearch from invSys (InventoryManagement!invSys)
' Returns a 2D array with columns: ROW, ITEM_CODE, ITEM, UOM, LOCATION, DESCRIPTION, VENDORS
Public Function LoadItemList() As Variant
    Dim ws As Worksheet: Set ws = SheetExists("InventoryManagement")
    If ws Is Nothing Then Exit Function
    Dim lo As ListObject: Set lo = ws.ListObjects("invSys")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    Dim cRow As Long, cCode As Long, cItem As Long, cUOM As Long, cLoc As Long, cDesc As Long, cVend As Long
    cRow = ColumnIndex(lo, "ROW")
    cCode = ColumnIndex(lo, "ITEM_CODE")
    cItem = ColumnIndex(lo, "ITEM")
    cUOM = ColumnIndex(lo, "UOM")
    cLoc = ColumnIndex(lo, "LOCATION")
    cDesc = ColumnIndex(lo, "DESCRIPTION")
    cVend = ColumnIndex(lo, "VENDOR(s)")
    If cCode * cItem = 0 Or cRow = 0 Then Exit Function

    Dim src As Variant: src = lo.DataBodyRange.value
    Dim r As Long, n As Long: n = UBound(src, 1)
    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 7)
    Dim outRow As Long: outRow = 0

    For r = 1 To n
        Dim itm As String: itm = NzStr(src(r, cItem))
        If itm <> "" Then
            outRow = outRow + 1
            outArr(outRow, 1) = NzStr(src(r, cRow))   ' ROW
            outArr(outRow, 2) = NzStr(src(r, cCode))  ' ITEM_CODE
            outArr(outRow, 3) = itm                   ' ITEM
            outArr(outRow, 4) = NzStr(src(r, cUOM))   ' UOM
            outArr(outRow, 5) = NzStr(src(r, cLoc))   ' LOCATION
            outArr(outRow, 6) = NzStr(src(r, cDesc))  ' DESCRIPTION
            outArr(outRow, 7) = NzStr(src(r, cVend))  ' VENDORS
        End If
    Next

    If outRow = 0 Then Exit Function
    ' Trim to actual count
    ReDim Preserve outArr(1 To outRow, 1 To 7)
    LoadItemList = outArr
End Function

