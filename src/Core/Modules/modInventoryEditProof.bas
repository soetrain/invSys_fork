Attribute VB_Name = "modInventoryEditProof"
Option Explicit
Option Private Module

' Read-only persisted proof. Missing receipt storage is valid only for a new
' owning EDIT; this reader never creates or repairs a table.
Public Function ReadReceipt(ByVal context As String, ByVal command As String, _
        ByRef found As Boolean, ByRef receipt As Variant, ByRef report As String) As Boolean
    Dim path As String, hash As String, copy As String, book As Workbook
    Dim table As ListObject, fields As Object, state As cInventoryTableInput
    Dim headers As Variant, values() As String, row As Long, i As Long, matches As Long
    On Error GoTo Failed
    found = False: receipt = Empty
    path = modInventoryCommands.SourceInventoryPath(context)
    If path = "" Or modAggregationFiles.Dirty(path) Then GoTo Failed
    hash = modTrainingWire.FileSha256(path)
    Set book = modAggregationFiles.OpenCopy(path, copy, True)
    If book Is Nothing Then GoTo Failed
    If hash = "" Or modTrainingWire.FileSha256(copy) <> hash Then GoTo Failed
    Set table = modExecutionTarget.Table(book, "tblInventoryEditReceipts")
    If Not table Is Nothing Then
        headers = modInventoryEditWire.ReceiptHeaders()
        Set fields = modAggregationSources.RequiredColumns(table, headers)
        If fields Is Nothing Then GoTo Failed
        Set state = New cInventoryTableInput
        If Not state.Capture(table, "CommandId", False) Then GoTo Failed
        For row = 1 To table.ListRows.Count
            If CStr(table.DataBodyRange.Cells(row, fields("CommandId")).Value2) <> "" Then
                ReDim values(0 To UBound(headers))
                For i = 0 To UBound(headers)
                    If table.DataBodyRange.Cells(row, fields(headers(i))).HasFormula Then GoTo Failed
                    values(i) = CStr(table.DataBodyRange.Cells(row, fields(headers(i))).Value2)
                Next i
                If Not modInventoryEditPlan.ValidReceipt(values) Then GoTo Failed
                If values(0) = command And command <> "" Then
                    matches = matches + 1: receipt = values
                End If
            End If
        Next row
        If matches > 1 Then GoTo Failed
        If matches = 1 Then found = True
    End If
    If hash <> modTrainingWire.FileSha256(path) Then GoTo Failed
    modAggregationFiles.CloseCopy book, copy
    ReadReceipt = modInventoryCommands.CheckForm(context, report)
    Exit Function
Failed:
    modAggregationFiles.CloseCopy book, copy
    report = "The worksheet EDIT receipt is unconfirmed. No new event was queued."
End Function

Public Function QuantityApplied(ByVal context As String, ByVal receipt As Variant, _
        ByRef applied As Boolean, ByRef report As String) As Boolean
    ' Quantity completion remains in the existing saved Domain event/log proof.
    ' Do not add a second mutable success flag or write a receipt after sign-out.
    Dim path As String, hash As String, copy As String, book As Workbook, table As ListObject
    Dim fields As Object, row As Long, events As Long, lines As Long, delta As Variant
    On Error GoTo Failed
    applied = False
    If CStr(receipt(7)) = "" Then
        If CStr(receipt(12)) <> "NONE" And CStr(receipt(12)) <> "NOOP" Then GoTo Failed
        QuantityApplied = modInventoryCommands.CheckForm(context, report): applied = QuantityApplied
        Exit Function
    End If
    If CStr(receipt(12)) <> "ADJUST" Or CStr(receipt(8)) = "" Then GoTo Failed
    If Not IsNumeric(receipt(9)) Or Not IsNumeric(receipt(10)) Then GoTo Failed
    path = modInventoryCommands.SourceInventoryPath(context)
    If path = "" Or modAggregationFiles.Dirty(path) Then GoTo Failed
    hash = modTrainingWire.FileSha256(path)
    Set book = modAggregationFiles.OpenCopy(path, copy, True)
    If book Is Nothing Then GoTo Failed
    If hash = "" Or modTrainingWire.FileSha256(copy) <> hash Then GoTo Failed
    Set table = modExecutionTarget.Table(book, "tblAppliedEvents")
    Set fields = modAggregationSources.RequiredColumns(table, Array("EventID", "Status"))
    If fields Is Nothing Then GoTo Failed
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = receipt(7) Then
            events = events + 1
            If CStr(table.DataBodyRange.Cells(row, fields("Status")).Value2) <> "APPLIED" Then GoTo Failed
        End If
    Next row
    If events > 1 Then GoTo Failed
    Set table = modExecutionTarget.Table(book, "tblInventoryLog")
    Set fields = modAggregationSources.RequiredColumns(table, _
        Array("EventID", "System_Key", "SKU", "QtyDelta", "WarehouseId", "StationId", "UserId"))
    If fields Is Nothing Then GoTo Failed
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = receipt(7) Then
            lines = lines + 1
            If CStr(table.DataBodyRange.Cells(row, fields("System_Key")).Value2) <> receipt(8) Then GoTo Failed
            If CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2) <> receipt(4) Then GoTo Failed
            If CStr(table.DataBodyRange.Cells(row, fields("WarehouseId")).Value2) <> receipt(1) Then GoTo Failed
            If CStr(table.DataBodyRange.Cells(row, fields("StationId")).Value2) <> receipt(2) Then GoTo Failed
            If CStr(table.DataBodyRange.Cells(row, fields("UserId")).Value2) <> receipt(3) Then GoTo Failed
            delta = table.DataBodyRange.Cells(row, fields("QtyDelta")).Value2
            If Not IsNumeric(delta) Then GoTo Failed
            If Abs(CDbl(delta) - (CDbl(receipt(10)) - CDbl(receipt(9)))) > 0.0000001 Then GoTo Failed
        End If
    Next row
    If lines <> events Then GoTo Failed
    If hash <> modTrainingWire.FileSha256(path) Then GoTo Failed
    modAggregationFiles.CloseCopy book, copy
    applied = (events = 1)
    QuantityApplied = modInventoryCommands.CheckForm(context, report)
    Exit Function
Failed:
    modAggregationFiles.CloseCopy book, copy
    report = "Exact saved EDIT quantity application is unconfirmed. No new event was queued."
End Function
