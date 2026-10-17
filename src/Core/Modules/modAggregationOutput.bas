Attribute VB_Name = "modAggregationOutput"
Option Explicit
Option Private Module

Private mOwned As Workbook

' Only the active owner can supply this staging path. Existing output is copied,
' never opened for writing; workbook-local columns/sheets remain in the candidate.
Public Function OpenCandidate(ByVal context As String, ByVal path As String) As Workbook
    Dim saved As Object, fso As Object, book As Workbook, security As Long
    On Error GoTo Failed
    security = Application.AutomationSecurity
    If Not modAggregationCommands.CheckPhase(context, path) Then Exit Function
    Set saved = modAggregationAuthority.Entry(context)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Application.AutomationSecurity = 3
    If fso.FileExists(CStr(saved("Destination"))) Then
        fso.CopyFile CStr(saved("Destination")), path, False
        If Not modAggregationCommands.CheckPhase(context, path) Then GoTo Finished
        Set book = Application.Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=False, Notify:=False, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    Else
        Set book = Application.Workbooks.Add
    End If
    If Not modAggregationCommands.CheckPhase(context, path) Then GoTo Finished
    Set OpenCandidate = book
    Set mOwned = book
Finished:
    Application.AutomationSecurity = security
    If OpenCandidate Is Nothing And Not book Is Nothing Then book.Close False
    Exit Function
Failed:
    Resume Finished
End Function

Public Sub CloseOwnedStaging()
    On Error Resume Next
    If Not mOwned Is Nothing Then mOwned.Close False
    Set mOwned = Nothing
End Sub

Public Function CaptureExtras(ByVal table As ListObject) As Object
    Dim fields As Object, values As Object, saved As Object, row As Long, column As ListColumn, key As String
    Set saved = CreateObject("Scripting.Dictionary"): saved.CompareMode = vbBinaryCompare
    Set fields = modAggregationSources.RequiredColumns(table, Array("WarehouseId", "System_Key", "SKU", "QtyOnHand", "LastAppliedAtUTC", "SourceSnapshot"))
    If fields Is Nothing Then Err.Raise vbObjectError + 7500
    For row = 1 To table.ListRows.Count
        key = CStr(table.DataBodyRange.Cells(row, fields("WarehouseId")).Value2) & "|" & CStr(table.DataBodyRange.Cells(row, fields("System_Key")).Value2)
        If saved.Exists(key) Then Err.Raise vbObjectError + 7501
        Set values = CreateObject("Scripting.Dictionary")
        For Each column In table.ListColumns
            If Not fields.Exists(Trim$(column.Name)) Then values.Add column.Name, table.DataBodyRange.Cells(row, column.Index).FormulaR1C1
        Next column
        saved.Add key, values
    Next row
    Set CaptureExtras = saved
End Function

Public Sub RestoreExtras(ByVal table As ListObject, ByVal saved As Object)
    Dim fields As Object, values As Object, row As Long, column As ListColumn, key As String
    Set fields = modAggregationSources.RequiredColumns(table, Array("WarehouseId", "System_Key"))
    If fields Is Nothing Then Err.Raise vbObjectError + 7502
    For row = 1 To table.ListRows.Count
        key = CStr(table.DataBodyRange.Cells(row, fields("WarehouseId")).Value2) & "|" & CStr(table.DataBodyRange.Cells(row, fields("System_Key")).Value2)
        If saved.Exists(key) Then
            Set values = saved(key)
            For Each column In table.ListColumns
                If values.Exists(column.Name) Then table.DataBodyRange.Cells(row, column.Index).FormulaR1C1 = values(column.Name)
            Next column
        End If
    Next row
End Sub
