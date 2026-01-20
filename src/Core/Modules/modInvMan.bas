Attribute VB_Name = "modInvMan"

Public Sub AddGoodsReceived_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim receivedCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, rowCol As Long
    Dim lastEditedCol As Long, totalInvLastEditCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long
    On Error GoTo ErrorHandler
    Call modUR_Transaction.BeginTransaction
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If
    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    receivedCol = tbl.ListColumns("RECEIVED").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    rowCol = tbl.ListColumns("ROW").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    totalInvLastEditCol = tbl.ListColumns("TOTAL INV LAST EDIT").Index
    rowCount = tbl.ListRows.count
    Set rng = tbl.DataBodyRange
    Set LogEntries = New Collection
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    For i = 1 To rowCount
        Dim receivedVal As Variant
        receivedVal = rng.Cells(i, receivedCol).value
        If IsNumeric(receivedVal) And receivedVal > 0 Then
            Dim oldTotalInv As Variant
            oldTotalInv = rng.Cells(i, totalInvCol).value
            ' Update TOTAL INV
            rng.Cells(i, totalInvCol).value = oldTotalInv + receivedVal
            ' Update LAST EDITED (general)
            rng.Cells(i, lastEditedCol).value = Now
            ' Update TOTAL INV LAST EDIT (specific to inventory)
            rng.Cells(i, totalInvLastEditCol).value = Now
            ' Update TOTAL INV LAST EDIT (specific to inventory)
            rng.Cells(i, totalInvLastEditCol).value = Now
            ' Track the change
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                rng.Cells(i, itemCodeCol).value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).value)
            ' Log the change
            LogEntries.Add Array(Environ("USERNAME"), "Added Goods Received", _
                rng.Cells(i, rowCol).Value, rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).value, receivedVal, rng.Cells(i, totalInvCol).value)
            ' Reset RECEIVED
            rng.Cells(i, receivedCol).value = 0
        End If
    Next i
    If LogEntries.count > 0 Then
        insertedCount = LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Goods received successfully.")
    Exit Sub
ErrorHandler:
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("AddGoodsReceived_Click")
    Resume Cleanup
End Sub
Public Sub DeductUsed_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim usedCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, rowCol As Long
    Dim lastEditedCol As Long, totalInvLastEditCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long
    On Error GoTo ErrorHandler
    Call modUR_Transaction.BeginTransaction
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If
    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    usedCol = tbl.ListColumns("USED").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    rowCol = tbl.ListColumns("ROW").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    totalInvLastEditCol = tbl.ListColumns("TOTAL INV LAST EDIT").Index
    rowCount = tbl.ListRows.count
    Set rng = tbl.DataBodyRange
    Set LogEntries = New Collection
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    For i = 1 To rowCount
        Dim usedVal As Variant
        usedVal = rng.Cells(i, usedCol).value
        If IsNumeric(usedVal) And usedVal > 0 Then
            Dim oldTotalInv As Variant
            oldTotalInv = rng.Cells(i, totalInvCol).value
            ' Update TOTAL INV
            rng.Cells(i, totalInvCol).value = oldTotalInv - usedVal
            ' Update LAST EDITED (general)
            rng.Cells(i, lastEditedCol).value = Now
            ' Update TOTAL INV LAST EDIT (specific to inventory)
            rng.Cells(i, totalInvLastEditCol).value = Now
            ' Track the change
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                rng.Cells(i, itemCodeCol).value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).value)
            ' Log the change
            LogEntries.Add Array(Environ("USERNAME"), "Deducted Used Items", _
                rng.Cells(i, rowCol).Value, rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).value, -usedVal, rng.Cells(i, totalInvCol).value)
            ' Reset USED
            rng.Cells(i, usedCol).value = 0
        End If
    Next i
    If LogEntries.count > 0 Then
        insertedCount = LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Used items deducted successfully.")
    Exit Sub
ErrorHandler:
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("DeductUsed_Click")
    Resume Cleanup
End Sub
Public Sub DeductShipments_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim shipmentsCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, rowCol As Long
    Dim lastEditedCol As Long, totalInvLastEditCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long
    On Error GoTo ErrorHandler
    Call modUR_Transaction.BeginTransaction
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If
    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    shipmentsCol = tbl.ListColumns("SHIPMENTS").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    rowCol = tbl.ListColumns("ROW").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    totalInvLastEditCol = tbl.ListColumns("TOTAL INV LAST EDIT").Index
    rowCount = tbl.ListRows.count
    Set rng = tbl.DataBodyRange
    Set LogEntries = New Collection
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    For i = 1 To rowCount
        Dim shipmentsVal As Variant
        shipmentsVal = rng.Cells(i, shipmentsCol).value
        If IsNumeric(shipmentsVal) And shipmentsVal > 0 Then
            Dim oldTotalInv As Variant
            oldTotalInv = rng.Cells(i, totalInvCol).value
            ' Update TOTAL INV
            rng.Cells(i, totalInvCol).value = oldTotalInv - shipmentsVal
            ' Update LAST EDITED (general)
            rng.Cells(i, lastEditedCol).value = Now
            ' Update TOTAL INV LAST EDIT (specific to inventory)
            rng.Cells(i, totalInvLastEditCol).value = Now
            ' Track the change
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                rng.Cells(i, itemCodeCol).value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).value)
            ' Log the change
            LogEntries.Add Array(Environ("USERNAME"), "Shipments Deducted", _
                rng.Cells(i, rowCol).Value, rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).value, -shipmentsVal, rng.Cells(i, totalInvCol).value)
            ' Reset SHIPMENTS
            rng.Cells(i, shipmentsCol).value = 0
        End If
    Next i
    If LogEntries.count > 0 Then
        insertedCount = LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Shipments deducted successfully.")
    Exit Sub
ErrorHandler:
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("DeductShipments_Click")
    Resume Cleanup
End Sub
Public Sub Adjustments_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim adjustmentsCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, rowCol As Long
    Dim lastEditedCol As Long, totalInvLastEditCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long
    On Error GoTo ErrorHandler
    Call modUR_Transaction.BeginTransaction
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If
    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    adjustmentsCol = tbl.ListColumns("ADJUSTMENTS").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    rowCol = tbl.ListColumns("ROW").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    totalInvLastEditCol = tbl.ListColumns("TOTAL INV LAST EDIT").Index
    rowCount = tbl.ListRows.count
    Set rng = tbl.DataBodyRange
    Set LogEntries = New Collection
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    For i = 1 To rowCount
        Dim adjustmentVal As Variant
        adjustmentVal = rng.Cells(i, adjustmentsCol).value
        If IsNumeric(adjustmentVal) And adjustmentVal <> 0 Then
            Dim oldTotalInv As Variant
            oldTotalInv = rng.Cells(i, totalInvCol).value
            ' Update TOTAL INV (positive adds, negative subtracts)
            rng.Cells(i, totalInvCol).value = oldTotalInv + adjustmentVal
            ' Update LAST EDITED (general)
            rng.Cells(i, lastEditedCol).value = Now
            ' Update TOTAL INV LAST EDIT (specific to inventory)
            rng.Cells(i, totalInvLastEditCol).value = Now
            ' Track the change
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                rng.Cells(i, itemCodeCol).value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).value)
            ' Log the change
            LogEntries.Add Array(Environ("USERNAME"), "Inventory Adjustment", _
                rng.Cells(i, rowCol).Value, rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).value, adjustmentVal, rng.Cells(i, totalInvCol).value)
            ' Reset ADJUSTMENTS
            rng.Cells(i, adjustmentsCol).value = 0
        End If
    Next i
    If LogEntries.count > 0 Then
        insertedCount = LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Adjustments applied successfully.")
    Exit Sub
ErrorHandler:
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("Adjustments_Click")
    Resume Cleanup
End Sub
Public Sub AddMadeItems_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim madeCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, rowCol As Long
    Dim lastEditedCol As Long, totalInvLastEditCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long
    On Error GoTo ErrorHandler
    Call modUR_Transaction.BeginTransaction
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If
    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    madeCol = tbl.ListColumns("MADE").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    rowCol = tbl.ListColumns("ROW").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    totalInvLastEditCol = tbl.ListColumns("TOTAL INV LAST EDIT").Index
    rowCount = tbl.ListRows.count
    Set rng = tbl.DataBodyRange
    Set LogEntries = New Collection
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    For i = 1 To rowCount
        Dim madeVal As Variant
        madeVal = rng.Cells(i, madeCol).value
        If IsNumeric(madeVal) And madeVal > 0 Then
            Dim oldTotalInv As Variant
            oldTotalInv = rng.Cells(i, totalInvCol).value
            ' Update TOTAL INV
            rng.Cells(i, totalInvCol).value = oldTotalInv + madeVal
            ' Update LAST EDITED (general)
            rng.Cells(i, lastEditedCol).value = Now
            ' Update TOTAL INV LAST EDIT (specific to inventory)
            rng.Cells(i, totalInvLastEditCol).value = Now
            ' Track the change
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                rng.Cells(i, itemCodeCol).value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).value)
            ' Log the change
            LogEntries.Add Array(Environ("USERNAME"), "Made Items Added", _
                rng.Cells(i, rowCol).Value, rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).value, madeVal, rng.Cells(i, totalInvCol).value)
            ' Reset MADE
            rng.Cells(i, madeCol).value = 0
        End If
    Next i
    If LogEntries.count > 0 Then
        insertedCount = LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Made items added successfully.")
    Exit Sub
ErrorHandler:
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("AddMadeItems_Click")
    Resume Cleanup
End Sub

Public Function ApplyUsedDeltas(deltas As Collection, ByRef errNotes As String, Optional actionLabel As String = "Deducted Used Items") As Double
    ApplyUsedDeltas = 0
    errNotes = ""
    If deltas Is Nothing Or deltas.Count = 0 Then Exit Function
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Then
        errNotes = "invSys table not found."
        ApplyUsedDeltas = -1
        Exit Function
    End If

    Dim usedCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, rowCol As Long
    Dim lastEditedCol As Long, totalInvLastEditCol As Long
    usedCol = tbl.ListColumns("USED").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    rowCol = tbl.ListColumns("ROW").Index
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    totalInvLastEditCol = tbl.ListColumns("TOTAL INV LAST EDIT").Index

    Dim logEntries As New Collection
    Dim delta As Variant
    Dim insertedCount As Long

    On Error GoTo ErrHandler
    Call modUR_Transaction.BeginTransaction
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDblInv(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvRowByRowValue(tbl, rowVal, rowCol)
        If invRow Is Nothing Then
            errNotes = "invSys row " & rowVal & " not found."
            GoTo FailHandler
        End If
        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, totalInvCol)
        Dim usedCell As Range: Set usedCell = invRow.Range.Cells(1, usedCol)
        Dim available As Double: available = NzDblInv(totalCell.Value)
        If qtyVal > available + 0.0000001 Then
            errNotes = "ROW " & rowVal & " requires " & Format$(qtyVal, "0.###") & " but only " & Format$(available, "0.###") & " available."
            GoTo FailHandler
        End If
        totalCell.Value = available - qtyVal
        Dim newUsed As Double: newUsed = Application.WorksheetFunction.Max(0, NzDblInv(usedCell.Value) - qtyVal)
        usedCell.Value = newUsed
        invRow.Range.Cells(1, lastEditedCol).Value = Now
        invRow.Range.Cells(1, totalInvLastEditCol).Value = Now
        Call modUR_Transaction.TrackTransactionChange("CellUpdate", invRow.Range.Cells(1, rowCol).Value, "TOTAL INV", available, totalCell.Value)
        logEntries.Add Array(Environ$("USERNAME"), actionLabel, rowVal, NzStrInv(delta("ITEM_CODE")), NzStrInv(delta("ITEM_NAME")), -qtyVal, NzDblInv(totalCell.Value))
        ApplyUsedDeltas = ApplyUsedDeltas + qtyVal
    Next delta

    If logEntries.Count > 0 Then
        insertedCount = LogMultipleInventoryChanges(logEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction
CleanupUsed:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Function
FailHandler:
    ApplyUsedDeltas = -1
    If modUR_Transaction.IsInTransaction Then Call modUR_Transaction.RollbackTransaction
    GoTo CleanupUsed
ErrHandler:
    errNotes = Err.Description
    If modUR_Transaction.IsInTransaction Then Call modUR_Transaction.RollbackTransaction
    ApplyUsedDeltas = -1
    Resume CleanupUsed
End Function

Public Function ApplyMadeDeltas(deltas As Collection, ByRef errNotes As String, Optional actionLabel As String = "Made Items Added") As Double
    ApplyMadeDeltas = 0
    errNotes = ""
    If deltas Is Nothing Or deltas.Count = 0 Then Exit Function
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Then
        errNotes = "invSys table not found."
        ApplyMadeDeltas = -1
        Exit Function
    End If

    Dim madeCol As Long, itemCodeCol As Long, itemNameCol As Long, rowCol As Long
    Dim lastEditedCol As Long
    madeCol = tbl.ListColumns("MADE").Index
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    rowCol = tbl.ListColumns("ROW").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index

    Dim logEntries As New Collection
    Dim delta As Variant
    Dim insertedCount As Long

    On Error GoTo ErrHandlerMade
    Call modUR_Transaction.BeginTransaction
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDblInv(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvRowByRowValue(tbl, rowVal, rowCol)
        If invRow Is Nothing Then
            errNotes = "Package row " & rowVal & " not found."
            GoTo FailHandlerMade
        End If
        Dim madeCell As Range: Set madeCell = invRow.Range.Cells(1, madeCol)
        Dim newMade As Double: newMade = NzDblInv(madeCell.Value) + qtyVal
        madeCell.Value = newMade
        invRow.Range.Cells(1, lastEditedCol).Value = Now
        logEntries.Add Array(Environ$("USERNAME"), actionLabel, rowVal, NzStrInv(delta("ITEM_CODE")), NzStrInv(delta("ITEM_NAME")), qtyVal, newMade)
        ApplyMadeDeltas = ApplyMadeDeltas + qtyVal
    Next delta

    If logEntries.Count > 0 Then
        insertedCount = LogMultipleInventoryChanges(logEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction
CleanupMade:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Function
FailHandlerMade:
    ApplyMadeDeltas = -1
    If modUR_Transaction.IsInTransaction Then Call modUR_Transaction.RollbackTransaction
    GoTo CleanupMade
ErrHandlerMade:
    errNotes = Err.Description
    If modUR_Transaction.IsInTransaction Then Call modUR_Transaction.RollbackTransaction
    ApplyMadeDeltas = -1
    Resume CleanupMade
End Function

Public Function ApplyMadeToInventoryDeltas(deltas As Collection, ByRef errNotes As String, Optional actionLabel As String = "BTN_TO_TOTALINV - Added To Total Inv") As Double
    ApplyMadeToInventoryDeltas = 0
    errNotes = ""
    If deltas Is Nothing Or deltas.Count = 0 Then Exit Function

    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Then
        errNotes = "invSys table not found."
        ApplyMadeToInventoryDeltas = -1
        Exit Function
    End If

    Dim madeCol As Long, totalInvCol As Long, rowCol As Long
    Dim itemCodeCol As Long, itemNameCol As Long
    Dim lastEditedCol As Long, totalInvLastEditCol As Long
    madeCol = tbl.ListColumns("MADE").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    rowCol = tbl.ListColumns("ROW").Index
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    totalInvLastEditCol = tbl.ListColumns("TOTAL INV LAST EDIT").Index

    Dim logEntries As New Collection
    Dim delta As Variant
    Dim insertedCount As Long

    On Error GoTo ErrHandlerMadeToInv
    Call modUR_Transaction.BeginTransaction
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDblInv(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvRowByRowValue(tbl, rowVal, rowCol)
        If invRow Is Nothing Then
            errNotes = "invSys row " & rowVal & " not found."
            GoTo FailHandlerMadeToInv
        End If

        Dim madeCell As Range: Set madeCell = invRow.Range.Cells(1, madeCol)
        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, totalInvCol)
        Dim stagedQty As Double: stagedQty = NzDblInv(madeCell.Value)
        Dim oldTotal As Double: oldTotal = NzDblInv(totalCell.Value)
        If qtyVal > stagedQty + 0.0000001 Then
            errNotes = "ROW " & rowVal & " only has " & Format$(stagedQty, "0.###") & " staged in MADE but requires " & Format$(qtyVal, "0.###") & "."
            GoTo FailHandlerMadeToInv
        End If

        madeCell.Value = stagedQty - qtyVal
        Dim newTotal As Double: newTotal = oldTotal + qtyVal
        totalCell.Value = newTotal

        invRow.Range.Cells(1, lastEditedCol).Value = Now
        invRow.Range.Cells(1, totalInvLastEditCol).Value = Now
        Call modUR_Transaction.TrackTransactionChange("CellUpdate", rowVal, "TOTAL INV", oldTotal, newTotal)

        logEntries.Add Array(Environ$("USERNAME"), actionLabel, rowVal, NzStrInv(delta("ITEM_CODE")), NzStrInv(delta("ITEM_NAME")), qtyVal, newTotal)
        ApplyMadeToInventoryDeltas = ApplyMadeToInventoryDeltas + qtyVal
    Next delta

    If logEntries.Count > 0 Then
        insertedCount = LogMultipleInventoryChanges(logEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction

CleanupMadeToInv:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Function

FailHandlerMadeToInv:
    ApplyMadeToInventoryDeltas = -1
    If modUR_Transaction.IsInTransaction Then Call modUR_Transaction.RollbackTransaction
    GoTo CleanupMadeToInv

ErrHandlerMadeToInv:
    errNotes = Err.Description
    If modUR_Transaction.IsInTransaction Then Call modUR_Transaction.RollbackTransaction
    ApplyMadeToInventoryDeltas = -1
    Resume CleanupMadeToInv
End Function

Public Function ApplyShipmentDeltas(deltas As Collection, ByRef errNotes As String, Optional actionLabel As String = "BTN_TO_SHIPMENTS - Inventory Staged") As Double
    ApplyShipmentDeltas = 0
    errNotes = ""
    If deltas Is Nothing Or deltas.Count = 0 Then Exit Function

    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Then
        errNotes = "invSys table not found."
        ApplyShipmentDeltas = -1
        Exit Function
    End If

    Dim totalInvCol As Long, shipmentsCol As Long, rowCol As Long
    Dim itemCodeCol As Long, itemNameCol As Long
    Dim lastEditedCol As Long, totalInvLastEditCol As Long
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    shipmentsCol = tbl.ListColumns("SHIPMENTS").Index
    rowCol = tbl.ListColumns("ROW").Index
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    totalInvLastEditCol = tbl.ListColumns("TOTAL INV LAST EDIT").Index

    Dim logEntries As New Collection
    Dim delta As Variant
    Dim insertedCount As Long

    On Error GoTo ErrHandlerShipments
    Call modUR_Transaction.BeginTransaction
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDblInv(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvRowByRowValue(tbl, rowVal, rowCol)
        If invRow Is Nothing Then
            errNotes = "invSys row " & rowVal & " not found."
            GoTo FailHandlerShipments
        End If

        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, totalInvCol)
        Dim shipmentsCell As Range: Set shipmentsCell = invRow.Range.Cells(1, shipmentsCol)
        Dim available As Double: available = NzDblInv(totalCell.Value)
        If qtyVal > available + 0.0000001 Then
            errNotes = "ROW " & rowVal & " needs " & Format$(qtyVal, "0.###") & " but only " & Format$(available, "0.###") & " available."
            GoTo FailHandlerShipments
        End If

        totalCell.Value = available - qtyVal
        Dim newShipments As Double: newShipments = NzDblInv(shipmentsCell.Value) + qtyVal
        shipmentsCell.Value = newShipments
        invRow.Range.Cells(1, lastEditedCol).Value = Now
        invRow.Range.Cells(1, totalInvLastEditCol).Value = Now

        Call modUR_Transaction.TrackTransactionChange("CellUpdate", rowVal, "TOTAL INV", available, totalCell.Value)
        logEntries.Add Array(Environ$("USERNAME"), actionLabel, rowVal, NzStrInv(delta("ITEM_CODE")), NzStrInv(delta("ITEM_NAME")), -qtyVal, NzDblInv(totalCell.Value))
        ApplyShipmentDeltas = ApplyShipmentDeltas + qtyVal
    Next delta

    If logEntries.Count > 0 Then
        insertedCount = LogMultipleInventoryChanges(logEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If
    Call modUR_Transaction.CommitTransaction

CleanupShipments:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Function

FailHandlerShipments:
    ApplyShipmentDeltas = -1
    If modUR_Transaction.IsInTransaction Then Call modUR_Transaction.RollbackTransaction
    GoTo CleanupShipments

ErrHandlerShipments:
    errNotes = Err.Description
    If modUR_Transaction.IsInTransaction Then Call modUR_Transaction.RollbackTransaction
    ApplyShipmentDeltas = -1
    Resume CleanupShipments
End Function
' Log helpers migrated from modInvLogs to keep InventoryLog routines centralized
Public Function LogMultipleInventoryChanges(LogEntries As Collection) As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim newRow As ListRow
    Dim logData As Variant
    Dim rowsInserted As Long
    Dim newLogID As String
    Dim userVal As Variant
    Dim actionVal As Variant
    Dim rowVal As Variant
    Dim itemCodeVal As Variant
    Dim itemNameVal As Variant
    Dim qtyVal As Variant
    Dim newQtyVal As Variant

    Set ws = ThisWorkbook.Sheets("InventoryLog")
    Set tbl = ws.ListObjects("InventoryLog")

    For i = 1 To LogEntries.Count
        logData = LogEntries(i)
        userVal = logData(0)
        actionVal = logData(1)
        If UBound(logData) >= 6 Then
            rowVal = logData(2)
            itemCodeVal = logData(3)
            itemNameVal = logData(4)
            qtyVal = logData(5)
            newQtyVal = logData(6)
        Else
            rowVal = ""
            itemCodeVal = logData(2)
            itemNameVal = logData(3)
            qtyVal = logData(4)
            newQtyVal = logData(5)
        End If

        Set newRow = tbl.ListRows.Add
        newLogID = modUR_Snapshot.GenerateGUID()
        With newRow.Range
            .Cells(1, 1).Value = newLogID
            .Cells(1, 2).Value = userVal
            .Cells(1, 3).Value = actionVal
            .Cells(1, 4).Value = rowVal
            .Cells(1, 5).Value = itemCodeVal
            .Cells(1, 6).Value = itemNameVal
            .Cells(1, 7).Value = qtyVal
            .Cells(1, 8).Value = newQtyVal
            .Cells(1, 9).Value = Now
        End With
        rowsInserted = rowsInserted + 1
    Next i

    LogMultipleInventoryChanges = rowsInserted
End Function

Public Function RemoveLastBulkLogEntries(ByVal CountToRemove As Long) As Collection
    Dim capturedEntries As New Collection
    Dim tbl As ListObject
    Dim i As Long
    Dim lastRow As Long
    Dim rowValues As Variant

    Set tbl = ThisWorkbook.Sheets("InventoryLog").ListObjects("InventoryLog")
    lastRow = tbl.ListRows.Count

    If lastRow = 0 Or CountToRemove <= 0 Then
        Set RemoveLastBulkLogEntries = capturedEntries
        Exit Function
    End If

    For i = 1 To Application.WorksheetFunction.Min(CountToRemove, lastRow)
        rowValues = tbl.ListRows(lastRow).Range.Value
        capturedEntries.Add rowValues
        tbl.ListRows(lastRow).Delete
        lastRow = lastRow - 1
    Next i

    Set RemoveLastBulkLogEntries = capturedEntries
End Function

Public Sub ReAddBulkLogEntries(ByVal LogDataCollection As Collection)
    Dim tbl As ListObject
    Dim i As Long
    Dim logRowData As Variant
    Dim newRow As ListRow

    Set tbl = ThisWorkbook.Sheets("InventoryLog").ListObjects("InventoryLog")
    If LogDataCollection Is Nothing Then Exit Sub

    For i = 1 To LogDataCollection.Count
        logRowData = LogDataCollection(i)
        If IsArray(logRowData) Then
            Set newRow = tbl.ListRows.Add
            With newRow.Range
                .Cells(1, 1).Value = logRowData(1, 1)
                .Cells(1, 2).Value = logRowData(1, 2)
                .Cells(1, 3).Value = logRowData(1, 3)
                .Cells(1, 4).Value = logRowData(1, 4)
                .Cells(1, 5).Value = logRowData(1, 5)
                .Cells(1, 6).Value = logRowData(1, 6)
                .Cells(1, 7).Value = logRowData(1, 7)
                .Cells(1, 8).Value = logRowData(1, 8)
                .Cells(1, 9).Value = logRowData(1, 9)
            End With
        End If
    Next i
End Sub
Public Sub DisplayMessage(msg As String)
    Dim ws As Worksheet
    Dim shp As Shape
    ' Set reference to the sheet
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    ' Check if the shape exists before updating
    On Error Resume Next
    Set shp = ws.shapes("lblMessage")
    On Error GoTo 0
    ' If shape exists, update the text
    If Not shp Is Nothing Then
        shp.TextFrame2.TextRange.text = msg
    Else
        MsgBox "Error: lblMessage text box not found!", vbCritical, "DisplayMessage Error"
    End If
End Sub

Private Function FindInvRowByRowValue(tbl As ListObject, ByVal rowValue As Long, ByVal rowColIndex As Long) As ListRow
    If tbl Is Nothing Or rowValue <= 0 Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    Dim r As Range
    For Each r In tbl.ListColumns(rowColIndex).DataBodyRange.Cells
        If NzDblInv(r.Value) = rowValue Then
            Set FindInvRowByRowValue = tbl.ListRows(r.Row - tbl.DataBodyRange.Row + 1)
            Exit Function
        End If
    Next r
End Function

Private Function NzDblInv(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzDblInv = 0#
    Else
        NzDblInv = CDbl(v)
    End If
End Function

Private Function NzStrInv(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStrInv = ""
    Else
        NzStrInv = CStr(v)
    End If
End Function

















