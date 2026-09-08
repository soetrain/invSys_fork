Attribute VB_Name = "modActivityStore"
Option Explicit
Option Private Module

Public Function Append(ByVal target As WarehouseTarget, ByVal recordId As String, _
                       ByVal body As String, ByRef notice As String) As Boolean
    Dim root As String, path As String, pending As String, hash As String, content As String
    Dim bytes() As Byte, file As Integer, opened As Boolean, fso As Object
    On Error GoTo Failed
    notice = "Tracking unavailable: the training record could not be saved."
    If Not modTrainingWire.ValidId(recordId) Or Len(body) = 0 Then Exit Function
    If Not ValidBody(target, recordId, body) Then Exit Function
    hash = modTrainingWire.Sha256(body)
    If Len(hash) <> 64 Then Exit Function
    content = Left$(body, Len(body) - 1) & ",""ContentSha256"":""" & hash & """}"
    If Len(content) > 1048576 Then Exit Function
    root = StoreRoot(target, True)
    If root = "" Then Exit Function
    path = root & "\" & recordId & ".json"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        Append = (ReadAscii(path) = content)
        If Append Then notice = ""
        Exit Function
    End If
    pending = root & "\" & modTrainingWire.NewId() & ".pending"
    If fso.FileExists(pending) Then Exit Function
    bytes = StrConv(content, vbFromUnicode)
    file = FreeFile
    Open pending For Binary Access Write Lock Read Write As #file
    opened = True
    Put #file, , bytes
    Close #file
    opened = False
    If ReadAscii(pending) <> content Then GoTo Failed
    ' Rename on the same filesystem publishes a complete file without overwrite.
    Name pending As path
    pending = ""
    Append = True: notice = ""
    Exit Function
Failed:
    On Error Resume Next
    If opened Then Close #file
    If pending <> "" Then
        If Not fso Is Nothing Then
            If fso.FileExists(pending) Then fso.DeleteFile pending, True
        End If
    End If
    ' A competing identical append is idempotent; a conflicting ID is not.
    If path <> "" And content <> "" Then
        If ReadAscii(path) = content Then Append = True: notice = ""
    End If
End Function

Public Function ReadRecord(ByVal target As WarehouseTarget, ByVal recordId As String) As String
    Dim root As String, text As String, marker As Long, body As String, hash As String
    On Error GoTo Invalid
    If Not modTrainingWire.ValidId(recordId) Then Exit Function
    root = StoreRoot(target, False)
    If root = "" Then Exit Function
    text = ReadAscii(root & "\" & recordId & ".json")
    marker = InStrRev(text, ",""ContentSha256"":""", -1, vbBinaryCompare)
    If marker = 0 Then Exit Function
    body = Left$(text, marker - 1) & "}"
    hash = Mid$(text, marker + Len(",""ContentSha256"":"""), 64)
    If Mid$(text, marker) <> ",""ContentSha256"":""" & hash & """}" Then Exit Function
    If Len(hash) <> 64 Or modTrainingWire.Sha256(body) <> hash Then Exit Function
    If Not ValidBody(target, recordId, body) Then Exit Function
    ReadRecord = text
Invalid:
End Function

Private Function ValidBody(ByVal target As WarehouseTarget, ByVal recordId As String, ByVal body As String) As Boolean
    Dim record As Object, definition As Object, outcome As Object, key As Variant
    Dim allowed As String, field As Variant
    On Error GoTo Invalid
    Set record = modTrainingJson.DecodeObject(body)
    If record Is Nothing Then Exit Function
    allowed = "|SchemaVersion|CatalogVersion|PackageSetVersion|BuildIdentity|RecordId|ActivityId|SequenceId|WarehouseId|StationId|UserId|ControlId|OwnerId|SourceKind|SourceRole|Caption|Surface|Ordinal|OccurredAtUTC|PolicyVersion|EventCode|OutcomeCode|Severity|DataEffect|UserMessage|NextStep|SourceEventRefs|"
    For Each key In record.Keys
        If InStr(1, allowed, "|" & CStr(key) & "|", vbBinaryCompare) = 0 Then Exit Function
        Select Case CStr(key)
            Case "SchemaVersion", "CatalogVersion", "Ordinal", "PolicyVersion"
                If VarType(record(key)) <> vbLong Then Exit Function
            Case "SourceEventRefs"
                If TypeName(record(key)) <> "Collection" Then Exit Function
            Case Else
                If VarType(record(key)) <> vbString Then Exit Function
        End Select
    Next key
    If record.Count <> 26 Then Exit Function
    If record("SchemaVersion") <> 1 Or record("CatalogVersion") <> modActivityCatalog.CATALOG_VERSION Then Exit Function
    If record("RecordId") <> recordId Or Not modTrainingWire.ValidId(record("ActivityId")) Then Exit Function
    If record("WarehouseId") <> target.WarehouseId Or record("SourceKind") <> "User activity" Then Exit Function
    If Not modTrainingWire.ValidSegment(record("StationId")) Or record("UserId") = "" Then Exit Function
    If record("PackageSetVersion") = "" Or record("BuildIdentity") = "" Then Exit Function
    If Not modTrainingWire.ValidUtcTimestamp(record("OccurredAtUTC")) Then Exit Function
    If record("PolicyVersion") < 0 Then Exit Function
    If record("SequenceId") <> "" Then
        If Not modTrainingWire.ValidId(record("SequenceId")) Then Exit Function
        If record("Ordinal") < 1 Or record("Ordinal") > 256 Then Exit Function
    Else
        If record("Ordinal") <> 0 Then Exit Function
    End If
    Set definition = modActivityCatalog.Control(record("ControlId"))
    Set outcome = modActivityCatalog.Outcome(record("ControlId"), record("OutcomeCode"))
    If definition Is Nothing Or outcome Is Nothing Then Exit Function
    For Each field In Array("OwnerId", "Caption", "Surface")
        If record(field) <> definition(field) Then Exit Function
    Next field
    If record("SourceRole") <> definition("Role") Then Exit Function
    For Each field In Array("EventCode", "Severity", "DataEffect", "UserMessage", "NextStep")
        If record(field) <> outcome(field) Then Exit Function
    Next field
    If TypeName(record("SourceEventRefs")) <> "Collection" Then Exit Function
    ' These initial configuration controls have no canonical business event.
    If record("SourceEventRefs").Count <> 0 Then Exit Function
    ValidBody = True
Invalid:
End Function

Private Function StoreRoot(ByVal target As WarehouseTarget, ByVal create As Boolean) As String
    Dim fso As Object, root As String, name As Variant
    If Not modNasConnection.IsWarehouseTargetAllowed(target, True) Then Exit Function
    If Not modTrainingWire.ValidSegment(target.WarehouseId) Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = target.RuntimeRoot
    If Not fso.FolderExists(root) Then Exit Function
    For Each name In Array("Training", "Activity", target.WarehouseId)
        root = fso.BuildPath(root, CStr(name))
        If Not fso.FolderExists(root) Then
            If Not create Or fso.FileExists(root) Then Exit Function
            fso.CreateFolder root
        End If
        If (fso.GetFolder(root).Attributes And &H400) <> 0 Then Exit Function
    Next name
    StoreRoot = root
End Function

Private Function ReadAscii(ByVal path As String) As String
    Dim stream As Object, fso As Object, text As String, i As Long
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    If fso.GetFile(path).Size < 1 Or fso.GetFile(path).Size > 1048576 Then Exit Function
    Set stream = fso.OpenTextFile(path, 1, False, 0)
    text = stream.Read(1048577)
    If Len(text) > 1048576 Then GoTo Failed
    For i = 1 To Len(text)
        If AscW(Mid$(text, i, 1)) < 0 Or AscW(Mid$(text, i, 1)) > 127 Then GoTo Failed
    Next i
    ReadAscii = text
Failed:
    If Not stream Is Nothing Then stream.Close
End Function
