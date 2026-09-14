Attribute VB_Name = "modEventsPublicationStore"
Option Explicit
Option Private Module

Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal existing As LongPtr, ByVal replacement As LongPtr, ByVal flags As Long) As Long

Public Function WriteModel(ByVal model As Object, ByVal runtimeRoot As String, ByRef report As String) As Boolean
    Dim body As String, hash As String, content As String, path As String, pending As String, fso As Object
    Dim bytes() As Byte, file As Integer, opened As Boolean, decoded As Object
    On Error GoTo Failed
    report = "Events publication failed; the prior complete Events artifact was retained."
    If Not ValidModel(model, CStr(model("WarehouseId"))) Then Exit Function
    body = modTrainingJson.EncodeObject(model)
    Set decoded = modTrainingJson.DecodePublicationObject(body)
    If decoded Is Nothing Then Exit Function
    If Not ValidModel(decoded, CStr(model("WarehouseId"))) Then Exit Function
    hash = modTrainingWire.Sha256(body)
    If Len(hash) <> 64 Then Exit Function
    content = Left$(body, Len(body) - 1) & ",""ContentSha256"":""" & hash & """}"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(runtimeRoot) Then Exit Function
    path = fso.BuildPath(runtimeRoot, CStr(model("WarehouseId")) & ".invSys.Snapshot.Events.json")
    pending = fso.BuildPath(runtimeRoot, ".events-" & modTrainingWire.NewId() & ".pending")
    If fso.FileExists(pending) Then Exit Function
    If fso.FileExists(path) Then
        If (fso.GetFile(path).Attributes And &H400) <> 0 Then Exit Function
    End If
    bytes = StrConv(content, vbFromUnicode)
    file = FreeFile
    Open pending For Binary Access Write Lock Read Write As #file
    opened = True: Put #file, , bytes: Close #file: opened = False
    If ReadAscii(pending) <> content Then GoTo Failed
    ' Same-directory rename/replace: no copy fallback or delete of the prior file.
    If MoveFileExW(StrPtr(pending), StrPtr(path), &H1 Or &H8) = 0 Then GoTo Failed
    pending = "": WriteModel = True
    report = "Events publication complete."
    Exit Function
Failed:
    On Error Resume Next
    If opened Then Close #file
    If pending <> "" And Not fso Is Nothing Then
        If fso.FileExists(pending) Then fso.DeleteFile pending, True
    End If
End Function

Public Function Read(ByVal path As String, ByVal warehouseId As String, Optional ByRef contentSha256 As String = "") As Object
    Dim content As String, body As String, hash As String, marker As Long, model As Object
    On Error GoTo Invalid
    contentSha256 = ""
    content = ReadAscii(path)
    marker = InStrRev(content, ",""ContentSha256"":""", -1, vbBinaryCompare)
    If marker = 0 Then Exit Function
    body = Left$(content, marker - 1) & "}"
    hash = Mid$(content, marker + Len(",""ContentSha256"":"""), 64)
    If Mid$(content, marker) <> ",""ContentSha256"":""" & hash & """}" Then Exit Function
    If Len(hash) <> 64 Or modTrainingWire.Sha256(body) <> hash Then Exit Function
    Set model = modTrainingJson.DecodePublicationObject(body)
    If model Is Nothing Then Exit Function
    If ValidModel(model, warehouseId) Then Set Read = model: contentSha256 = hash
Invalid:
End Function

Private Function ValidModel(ByVal model As Object, ByVal warehouse As String) As Boolean
    Dim field As Variant, entry As Object, group As Object, source As String, key As String, seen As Object, sources As Object
    Dim groups As Object, lines As Object, count As Long, current As Variant, names As Variant
    On Error GoTo Invalid
    If model.Count <> 10 Or model("SchemaVersion") <> 1 Then Exit Function
    If Not modTrainingWire.ValidSegment(warehouse) Or model("WarehouseId") <> warehouse Then Exit Function
    If Not modTrainingWire.ValidId(CStr(model("PublicationId"))) Then Exit Function
    If Not modTrainingWire.ValidUtcTimestamp(CStr(model("PublishedAtUTC"))) Then Exit Function
    If VarType(model("PolicyVersion")) <> vbLong And VarType(model("PolicyVersion")) <> vbInteger Then Exit Function
    If model("PolicyVersion") < 0 Then Exit Function
    For Each field In Array("PackageSetVersion", "BuildIdentity")
        If VarType(model(field)) <> vbString Or model(field) = "" Then Exit Function
    Next field
    If TypeName(model("Groups")) <> "Collection" Or model("Groups").Count > 5000 Then Exit Function
    If TypeName(model("CurrentState")) <> "Collection" Then Exit Function
    If model("Coverage").Count <> 1 Or model("Coverage")("Sources").Count <> 5 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbBinaryCompare
    Set sources = CreateObject("Scripting.Dictionary"): sources.CompareMode = vbBinaryCompare
    Set groups = CreateObject("Scripting.Dictionary"): Set lines = CreateObject("Scripting.Dictionary")
    names = Array("Inventory", "Designs", "Activity", "ShippingBOM", "ShippingHolds")
    For Each field In names
        groups.Add CStr(field), 0&: lines.Add CStr(field), 0&
    Next field
    For Each group In model("Groups")
        source = CStr(group("Source"))
        If source <> "Inventory" And source <> "Designs" And source <> "Activity" Then Exit Function
        If group.Count <> 7 Or CStr(group("SourceId")) = "" Then Exit Function
        key = source & Chr$(30) & CStr(group("SourceId"))
        If seen.Exists(key) Then Exit Function
        seen.Add key, True
        If TypeName(group("Lines")) <> "Collection" Or group("Lines").Count < 1 Then Exit Function
        If TypeName(group("Outcomes")) <> "Collection" Then Exit Function
        groups(source) = groups(source) + 1&: lines(source) = lines(source) + group("Lines").Count
    Next group
    For Each entry In model("CurrentState")
        source = CStr(entry("Source"))
        If source <> "ShippingBOM" And source <> "ShippingHolds" Then Exit Function
        If entry.Count <> 4 Or entry("SourceKind") <> "Current state" Then Exit Function
        If TypeName(entry("Lines")) <> "Collection" Or entry("Lines").Count < 1 Then Exit Function
        groups(source) = groups(source) + 1&: lines(source) = lines(source) + entry("Lines").Count
    Next entry
    For Each entry In model("Coverage")("Sources")
        source = CStr(entry("Source"))
        If Not groups.Exists(source) Or sources.Exists(source) Then Exit Function
        sources.Add source, True
        If entry("Scope") <> IIf(source = "ShippingHolds", "Station profile", "Warehouse") Then Exit Function
        If entry("Availability") = "Available" Then
            For Each field In Array("Groups", "Lines")
                For Each current In Array("Available", "Included", "Omitted")
                    If VarType(entry(current & field)) <> vbLong And VarType(entry(current & field)) <> vbInteger Then Exit Function
                    If entry(current & field) < 0 Then Exit Function
                Next current
                If entry("Available" & field) <> entry("Included" & field) + entry("Omitted" & field) Then Exit Function
            Next field
            If entry("IncludedGroups") <> groups(source) Or entry("IncludedLines") <> lines(source) Then Exit Function
        ElseIf entry("Availability") = "Unavailable" Then
            If groups(source) <> 0 Or lines(source) <> 0 Then Exit Function
            For Each field In Array("AvailableGroups", "IncludedGroups", "OmittedGroups", "AvailableLines", "IncludedLines", "OmittedLines")
                If entry(field) <> "Unavailable" Then Exit Function
            Next field
        Else
            Exit Function
        End If
    Next entry
    ValidModel = True
Invalid:
End Function

Private Function ReadAscii(ByVal path As String) As String
    Dim stream As Object
    On Error GoTo Failed
    Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(path, 1, False, 0)
    ReadAscii = stream.ReadAll
Failed:
    If Not stream Is Nothing Then stream.Close
End Function
