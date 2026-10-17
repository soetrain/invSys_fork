Attribute VB_Name = "modProcessorPersistence"
Option Explicit

' Existing processor persistence/ownership rules, shared by its typed callers.
Public Sub SaveWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If Trim$(wb.Path) = "" Then Exit Sub
    wb.Save
End Sub

Public Sub CloseTransient(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    If Not wb.ReadOnly Then
        If wb.Saved = False Then wb.Save
    End If
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Public Function ResolveWorkbookNameProcessor(ByVal wb As Workbook) As String
    If wb Is Nothing Then
        ResolveWorkbookNameProcessor = "<none>"
    ElseIf Trim$(wb.Name) <> "" Then
        ResolveWorkbookNameProcessor = wb.Name
    Else
        ResolveWorkbookNameProcessor = "<unnamed>"
    End If
End Function
