Attribute VB_Name = "modShippingPublicationSource"
Option Explicit

Private Const BOM_FIELDS As String = "PackageSystemKey|PackageItem|PackageUOM|PackageLocation|PackageDescription|BomVersion|BomVersionLabel|IsActive|EffectiveFromUTC|EffectiveToUTC|RetiredAtUTC|ComponentSystemKey|ComponentItemCode|ComponentItem|ComponentQty|ComponentUOM|ComponentLocation|ComponentDescription|UpdatedAtUTC|UpdatedBy"
Private Const HOLD_FIELDS As String = "Ref|Item|Qty|UOM|Location|Description|Area|Carrier|ShipmentLineId|ReserveEventId|System_Key"

' D18: fixed publication owner read. Viewer never calls this entry.
Public Function ReadForPublication(ByVal source As String, ByVal warehouseId As String, ByVal runtimeRoot As String) As String
    Dim headers As String, body As String, reason As String, mode As String, count As Long
    On Error GoTo Failed
    reason = "ContextMismatch": mode = "None"
    Select Case source
        Case "ShippingBOM": headers = BOM_FIELDS
        Case "ShippingHolds": headers = HOLD_FIELDS
        Case Else: Exit Function
    End Select
    If Not modTrainingWire.ValidSegment(warehouseId) Then GoTo Done
    If Not modNasConnection.IsCurrentTargetAllowed(True) Then GoTo Done
    If StrComp(warehouseId, modNasConnection.GetCurrentTargetWarehouseId(), vbBinaryCompare) <> 0 Then GoTo Done
    If StrComp(modConfig.NormalizeFolderPathForRuntime(runtimeRoot), modConfig.NormalizeFolderPathForRuntime(modNasConnection.GetCurrentTargetRuntimeRoot()), vbTextCompare) <> 0 Then GoTo Done
    If source = "ShippingBOM" Then
        body = ReadBom(modConfig.NormalizeFolderPathForRuntime(runtimeRoot) & "\" & warehouseId & ".invSys.Data.ShippingBOM.xlsb", count, mode, reason)
    Else
        body = ReadHolds(warehouseId, count, mode, reason)
    End If
Done:
    If reason <> "OK" Then count = 0: body = "": mode = "None"
    ReadForPublication = "EVTSRC1" & vbTab & source & vbTab & IIf(reason = "OK", "Available", "Unavailable") & _
        vbTab & modTS_Shipments.EscapeHoldField(warehouseId) & vbTab & CStr(count) & vbTab & mode & vbTab & reason & vbCrLf & Replace$(headers, "|", vbTab) & body
    Exit Function
Failed:
    reason = "ReadFailed"
    Resume Done
End Function

Private Function ReadBom(ByVal path As String, ByRef count As Long, ByRef mode As String, ByRef reason As String) As String
    Dim wb As Workbook, candidate As Workbook, owned As Boolean, events As Boolean, security As Long
    Dim ws As Worksheet, lo As ListObject, source As ListObject, column As ListColumn, window As Window
    Dim required As Object, columns As Object, names As Variant, name As Variant, values As Variant
    Dim rows() As String, fields() As String, r As Long, c As Long, value As Variant, text As String
    On Error GoTo Failed
    reason = "MissingSource"
    events = Application.EnableEvents: security = Application.AutomationSecurity
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, path, vbTextCompare) = 0 Then
            reason = "UnsavedSource"
            If Not candidate.Saved Then Exit Function
            Set wb = candidate: mode = "Borrowed": Exit For
        End If
    Next candidate
    events = Application.EnableEvents: security = Application.AutomationSecurity
    If wb Is Nothing Then
        If Not CreateObject("Scripting.FileSystemObject").FileExists(path) Then Exit Function
        Application.EnableEvents = False: Application.AutomationSecurity = 3
        Set wb = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        owned = True: mode = "ReadOnly"
        For Each window In wb.Windows
            window.Visible = False
        Next window
        Application.AutomationSecurity = security: Application.EnableEvents = events
    End If
    reason = "InvalidSchema"
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, "tblShippingBOM", vbTextCompare) = 0 Then
                If Not source Is Nothing Then GoTo Done
                Set source = lo
            End If
        Next lo
    Next ws
    If source Is Nothing Then GoTo Done
    names = Split(BOM_FIELDS, "|")
    Set required = CreateObject("Scripting.Dictionary"): required.CompareMode = vbTextCompare
    Set columns = CreateObject("Scripting.Dictionary"): columns.CompareMode = vbTextCompare
    For Each name In names: required.Add CStr(name), True: Next name
    For Each column In source.ListColumns
        text = Trim$(column.Name)
        If required.Exists(text) Then
            If columns.Exists(text) Then GoTo Done
            columns.Add text, column.Index
        End If
    Next column
    If columns.Count <> required.Count Then GoTo Done
    count = source.ListRows.Count
    If count > 0 Then
        values = source.DataBodyRange.Value2
        ReDim rows(1 To count): ReDim fields(0 To UBound(names))
        For r = 1 To count
            For c = 0 To UBound(names)
                value = values(r, CLng(columns(names(c))))
                If IsError(value) Or IsNull(value) Then GoTo Done
                text = CStr(value)
                If Right$(CStr(names(c)), 3) = "UTC" And Len(text) > 0 Then
                    If IsNumeric(value) Or IsDate(value) Then text = Format$(CDate(value), "yyyy-mm-dd\Thh:nn:ss")
                End If
                fields(c) = modTS_Shipments.EscapeHoldField(text)
            Next c
            If fields(0) = "" Or fields(11) = "" Then GoTo Done
            rows(r) = Join(fields, vbTab)
        Next r
        ReadBom = vbCrLf & Join(rows, vbCrLf)
    End If
    reason = "OK"
Done:
    On Error Resume Next
    Application.AutomationSecurity = security: Application.EnableEvents = events
    If owned Then
        Application.EnableEvents = False
        wb.Close SaveChanges:=False
        Application.EnableEvents = events
    End If
    Exit Function
Failed:
    reason = "ReadFailed"
    Resume Done
End Function

Private Function ReadHolds(ByVal warehouseId As String, ByRef count As Long, ByRef mode As String, ByRef reason As String) As String
    Dim root As String, path As String, fso As Object, stream As Object, content As String, line As Variant
    Dim fields As Variant, output(0 To 10) As String, rows As Collection, lines() As String, i As Long, c As Long
    On Error GoTo Failed
    reason = "MissingSource"
    root = Environ$("LOCALAPPDATA")
    If Trim$(root) = "" Then root = Environ$("TEMP")
    If Trim$(root) = "" Then Exit Function
    For i = 1 To Len(warehouseId)
        content = Mid$(warehouseId, i, 1)
        path = path & IIf(content Like "[A-Za-z0-9_-]", content, "_")
    Next i
    ' An aliased filename cannot prove which warehouse owns this local state.
    If StrComp(path, warehouseId, vbBinaryCompare) <> 0 Then reason = "ContextMismatch": Exit Function
    path = modConfig.NormalizeFolderPathForRuntime(root) & "\invSys\shipping_hold_" & path & ".tsv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Set stream = fso.OpenTextFile(path, 1, False, 0)
    content = stream.ReadAll: stream.Close: Set stream = Nothing
    reason = "InvalidSchema": Set rows = New Collection
    For Each line In Split(Replace$(content, vbCrLf, vbLf), vbLf)
        If CStr(line) <> "" Then
            fields = Split(CStr(line), vbTab)
            If UBound(fields) <> 11 Then Exit Function
            If CStr(fields(11)) = "" Or CStr(fields(9)) = "" Then Exit Function
            For c = 0 To 10
                ' Stored fields already use EVTSRC1's escape convention.
                output(c) = CStr(fields(IIf(c < 3, c, c + 1)))
            Next c
            rows.Add Join(output, vbTab)
        End If
    Next line
    count = rows.Count
    If count > 0 Then
        ReDim lines(1 To count)
        For i = 1 To count: lines(i) = rows(i): Next i
        ReadHolds = vbCrLf & Join(lines, vbCrLf)
    End If
    mode = "ReadOnly": reason = "OK"
    Exit Function
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    reason = "ReadFailed"
End Function
