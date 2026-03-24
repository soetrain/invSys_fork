Attribute VB_Name = "TestInventorySchema"
Option Explicit

Public Sub RunInventorySchemaTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestEnsureInventorySchema_RecreatesTables(), passed, failed
    Tally TestEnsureInventorySchema_AddsMissingColumns(), passed, failed
    Tally TestEnsureInventorySchema_RemovesBlankSeedRow(), passed, failed

    Debug.Print "Inventory schema tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestEnsureInventorySchema_RecreatesTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If modInventorySchema.EnsureInventorySchema(wb, report) Then
        If TableExists(wb, "tblInventoryLog") And TableExists(wb, "tblAppliedEvents") And TableExists(wb, "tblLocks") Then
            TestEnsureInventorySchema_RecreatesTables = 1
        End If
    End If

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureInventorySchema_AddsMissingColumns() As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim report As String

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "InventoryLog"
    ws.Range("A1").Resize(1, 2).Value = Array("EventID", "SKU")
    ws.Range("A2").Resize(1, 2).Value = Array("", "")
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)
    lo.Name = "tblInventoryLog"

    On Error GoTo CleanFail
    If modInventorySchema.EnsureInventorySchema(wb, report) Then
        If ColumnExists(lo, "AppliedAtUTC") And ColumnExists(lo, "QtyDelta") Then
            TestEnsureInventorySchema_AddsMissingColumns = 1
        End If
    End If

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureInventorySchema_RemovesBlankSeedRow() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loLog As ListObject
    Dim loApplied As ListObject

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modInventorySchema.EnsureInventorySchema(wb, report) Then GoTo CleanExit

    Set loLog = wb.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    Set loApplied = wb.Worksheets("AppliedEvents").ListObjects("tblAppliedEvents")
    If loLog.ListRows.Count <> 0 Then GoTo CleanExit
    If loApplied.ListRows.Count <> 0 Then GoTo CleanExit

    TestEnsureInventorySchema_RemovesBlankSeedRow = 1

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        If Not ws.ListObjects(tableName) Is Nothing Then
            TableExists = True
            Exit Function
        End If
    Next ws
    On Error GoTo 0
End Function

Private Function ColumnExists(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            ColumnExists = True
            Exit Function
        End If
    Next i
End Function

Private Sub Tally(ByVal testResult As Long, ByRef passed As Long, ByRef failed As Long)
    If testResult = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub

Private Sub CloseNoSave(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub
