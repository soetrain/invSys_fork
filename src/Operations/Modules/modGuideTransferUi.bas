Attribute VB_Name = "modGuideTransferUi"
Option Explicit
Option Private Module

' Dialog ownership stays in Operations; Core rechecks the captured request.
Public Function Export(ByVal context As String, ByVal key As String, ByRef notice As String, ByRef outcome As String) As Boolean
    Dim path As String
    On Error GoTo Failed
    If Not modGuideTransfer.CanExport(context, key, notice, outcome) Then Exit Function
    path = SelectFileForTransfer(True)
    If path = "" Then notice = "Export cancelled.": outcome = "CANCELLED": Exit Function
    Export = modGuideTransfer.ExportGuide(context, key, path, notice, outcome)
    Exit Function
Failed:
    outcome = "FAILED"
    notice = "Unavailable: the guide file could not be selected for export."
End Function

Public Function Import(ByVal context As String, ByRef key As String, ByRef notice As String, ByRef outcome As String) As Boolean
    Dim path As String
    On Error GoTo Failed
    key = ""
    If Not modGuideTransfer.CanImport(context, notice, outcome) Then Exit Function
    path = SelectFileForTransfer(False)
    If path = "" Then notice = "Import cancelled.": outcome = "CANCELLED": Exit Function
    Import = modGuideTransfer.ImportGuide(context, path, key, notice, outcome)
    Exit Function
Failed:
    outcome = "FAILED"
    notice = "Unavailable: the guide file could not be selected for import."
End Function

Private Function SelectFileForTransfer(ByVal exporting As Boolean) As String
    Dim selected As Variant
    If exporting Then
        selected = Application.GetSaveAsFilename(InitialFileName:="How-To.json", FileFilter:="JSON guide package (*.json),*.json", Title:="Export guide to a new file")
    Else
        selected = Application.GetOpenFilename(FileFilter:="JSON guide package (*.json),*.json", Title:="Import guide", MultiSelect:=False)
    End If
    If VarType(selected) = vbString Then SelectFileForTransfer = CStr(selected)
End Function
