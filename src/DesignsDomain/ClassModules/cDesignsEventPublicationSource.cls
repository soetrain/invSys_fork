VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cDesignsEventPublicationSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FIELDS As String = "EventID|UndoOfEventId|AppliedSeq|EventType|OccurredAtUTC|AppliedAtUTC|WarehouseId|StationId|UserId|DefinitionType|DefinitionId|DefinitionVersion|Note"

Public Function Read(ByVal warehouseId As String, ByVal runtimeRoot As String) As String
    Dim wb As Workbook, owned As Boolean, mode As String, reason As String, body As String
    On Error GoTo Failed
    reason = "ContextMismatch"
    If warehouseId = "" Or runtimeRoot = "" Then GoTo Done
    If Not modNasConnection.IsCurrentTargetAllowed(True) Then GoTo Done
    If StrComp(warehouseId, modNasConnection.GetCurrentTargetWarehouseId(), vbBinaryCompare) <> 0 Then GoTo Done
    If StrComp(Folder(runtimeRoot), Folder(modNasConnection.GetCurrentTargetRuntimeRoot()), vbTextCompare) <> 0 Then GoTo Done
    Set wb = Acquire(Folder(runtimeRoot) & "\" & warehouseId & ".invSys.Data.Designs.xlsb", owned, mode, reason)
    If wb Is Nothing Then GoTo Done
    body = ReadRows(wb, warehouseId, mode, reason)
Done:
    If owned Then CloseSource wb
    If body = "" Then body = Envelope("Unavailable", warehouseId, 0, "None", reason) & vbCrLf & Replace$(FIELDS, "|", vbTab)
    Read = body
    Exit Function
Failed:
    reason = "ReadFailed"
    Resume Done
End Function

Private Function Acquire(ByVal path As String, ByRef owned As Boolean, ByRef mode As String, ByRef reason As String) As Workbook
    Dim wb As Workbook, events As Boolean, security As Long, guarded As Boolean, window As Window
    On Error GoTo Failed
    reason = "MissingSource"
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            reason = "UnsavedSource"
            If Not wb.Saved Then Exit Function
            mode = "Borrowed": Set Acquire = wb
            Exit Function
        End If
    Next wb
    If Not CreateObject("Scripting.FileSystemObject").FileExists(path) Then Exit Function
    events = Application.EnableEvents: security = Application.AutomationSecurity: guarded = True
    Application.EnableEvents = False: Application.AutomationSecurity = 3
    Set wb = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    owned = True
    For Each window In wb.Windows
        window.Visible = False
    Next window
    mode = "ReadOnly": Set Acquire = wb
Done:
    If guarded Then
        Application.AutomationSecurity = security
        Application.EnableEvents = events
    End If
    Exit Function
Failed:
    reason = "ReadFailed"
    If owned Then CloseSource wb
    owned = False
    Resume Done
End Function

Private Function ReadRows(ByVal wb As Workbook, ByVal warehouseId As String, ByVal mode As String, ByRef reason As String) As String
    Dim ws As Worksheet, lo As ListObject, source As ListObject, columns As Object
    Dim names As Variant, values As Variant, rows() As String, fieldsOut() As String, r As Long, c As Long
    reason = "InvalidSchema"
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, "tblDesignEvents", vbTextCompare) = 0 Then Set source = lo
        Next lo
    Next ws
    If source Is Nothing Then Exit Function
    names = Split(FIELDS, "|")
    Set columns = RequiredColumns(source, names)
    If columns Is Nothing Then Exit Function
    ReDim rows(0 To source.ListRows.Count + 1)
    rows(0) = Envelope("Available", warehouseId, source.ListRows.Count, mode, "OK")
    rows(1) = Join(names, vbTab)
    ReDim fieldsOut(0 To UBound(names))
    If Not source.DataBodyRange Is Nothing Then
        values = source.DataBodyRange.Value2
        For r = 1 To UBound(values, 1)
            reason = "WarehouseMismatch"
            If StrComp(CStr(values(r, columns("WarehouseId"))), warehouseId, vbBinaryCompare) <> 0 Then Exit Function
            reason = "InvalidData"
            For c = 0 To UBound(names)
                fieldsOut(c) = EscapeField(FieldText(values(r, columns(names(c))), CStr(names(c))))
            Next c
            rows(r + 1) = Join(fieldsOut, vbTab)
        Next r
    End If
    ReadRows = Join(rows, vbCrLf)
End Function

Private Function RequiredColumns(ByVal source As ListObject, ByVal names As Variant) As Object
    Dim found As Object, required As Object, column As ListColumn, name As Variant, normalized As String
    Set found = CreateObject("Scripting.Dictionary"): found.CompareMode = vbTextCompare
    Set required = CreateObject("Scripting.Dictionary"): required.CompareMode = vbTextCompare
    For Each name In names
        required.Add CStr(name), True
    Next name
    For Each column In source.ListColumns
        normalized = Trim$(column.Name)
        If required.Exists(normalized) Then
            If found.Exists(normalized) Then Exit Function
            found.Add normalized, column.Index
        End If
    Next column
    If found.Count = required.Count Then Set RequiredColumns = found
End Function

Private Function FieldText(ByVal value As Variant, ByVal field As String) As String
    If IsError(value) Or IsNull(value) Then Err.Raise vbObjectError + 2840
    If (field = "OccurredAtUTC" Or field = "AppliedAtUTC") And Not IsEmpty(value) Then
        If IsNumeric(value) Or IsDate(value) Then
            FieldText = Format$(CDate(value), "yyyy-mm-dd\Thh:nn:ss")
            Exit Function
        End If
    End If
    FieldText = CStr(value)
End Function

Private Function EscapeField(ByVal value As String) As String
    value = Replace$(value, "\", "\\")
    value = Replace$(value, vbTab, "\t")
    value = Replace$(value, vbCr, "\r")
    EscapeField = Replace$(value, vbLf, "\n")
End Function

Private Function Envelope(ByVal availability As String, ByVal warehouseId As String, ByVal count As Long, ByVal mode As String, ByVal reason As String) As String
    Envelope = "EVTSRC1" & vbTab & "Designs" & vbTab & availability & vbTab & EscapeField(warehouseId) & vbTab & CStr(count) & vbTab & mode & vbTab & reason
End Function

Private Function Folder(ByVal value As String) As String
    value = Replace$(value, "/", "\")
    Do While Len(value) > 0 And Right$(value, 1) = "\"
        value = Left$(value, Len(value) - 1)
    Loop
    Folder = value
End Function

Private Sub CloseSource(ByVal wb As Workbook)
    Dim events As Boolean
    On Error Resume Next
    events = Application.EnableEvents
    Application.EnableEvents = False
    wb.Close SaveChanges:=False
    Application.EnableEvents = events
End Sub
