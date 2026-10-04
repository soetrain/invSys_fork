Attribute VB_Name = "modExecutionProfileStore"
Option Explicit
Option Private Module

Private Function ReadVersion(ByVal target As WarehouseTarget, ByVal id As String, ByVal version As Long) As Object
    Dim root As String, path As String, text As String, marker As Long, hash As String, body As String, model As Object, fso As Object
    On Error GoTo Invalid
    If Not modTrainingWire.ValidId(id) Or version < 1 Then Exit Function
    root = modRecordingJournal.ChildRoot(target, "ExecutionProfiles", False)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, id & "." & CStr(version) & ".json")
    If Not fso.FileExists(path) Then Exit Function
    If (fso.GetFile(path).Attributes And &H400) <> 0 Then Exit Function
    text = modRecordingJournal.ReadText(path)
    marker = InStrRev(text, ",""ContentSha256"":""", -1, vbBinaryCompare)
    If marker = 0 Then Exit Function
    body = Left$(text, marker - 1) & "}": hash = Mid$(text, marker + Len(",""ContentSha256"":"""), 64)
    If Mid$(text, marker) <> ",""ContentSha256"":""" & hash & """}" Then Exit Function
    If Not modEvaluationModel.IsHash(hash) Or modTrainingWire.Sha256(body) <> hash Then Exit Function
    Set model = modTrainingJson.DecodeObject(body)
    If model Is Nothing Then Exit Function
    If Not modExecutionProfileModel.Validate(target, model) Then Exit Function
    If model("ProfileId") <> id Or model("Version") <> version Then Exit Function
    model.Add "ContentSha256", hash: Set ReadVersion = model
Invalid:
End Function

Private Function ReadChain(ByVal target As WarehouseTarget, ByVal id As String, ByVal version As Long) As Object
    Dim index As Long, current As Object, previous As Object, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    For index = 1 To version
        Set current = ReadVersion(target, id, index)
        If current Is Nothing Then Exit Function
        If seen.Exists(current("RecordId")) Then Exit Function
        seen.Add current("RecordId"), True
        If Not previous Is Nothing Then
            If current("PreviousRecordId") <> previous("RecordId") Or current("PreviousSha256") <> previous("ContentSha256") Then Exit Function
            If modTrainingJson.EncodeObject(current("Guide")) <> modTrainingJson.EncodeObject(previous("Guide")) Then Exit Function
        End If
        Set previous = current
    Next index
    Set ReadChain = current
End Function

Public Function Latest(ByVal target As WarehouseTarget, ByVal guide As Object, ByRef profile As Object, ByRef notice As String) As Boolean
    Dim root As String, fso As Object, file As Object, parts As Variant, version As Long, current As Object
    On Error GoTo Invalid
    Set profile = Nothing
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = modRecordingJournal.JournalRoot(target, False)
    If root = "" Then GoTo Invalid
    root = fso.BuildPath(root, "ExecutionProfiles")
    If fso.FileExists(root) Then GoTo Invalid
    If Not fso.FolderExists(root) Then Latest = True: Exit Function
    root = modRecordingJournal.ChildRoot(target, "ExecutionProfiles", False)
    If root = "" Then GoTo Invalid
    For Each file In fso.GetFolder(root).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "json" Then
            parts = Split(file.Name, ".")
            If UBound(parts) <> 2 Then GoTo Invalid
            version = CLng(parts(1))
            If version < 1 Or CStr(version) <> CStr(parts(1)) Then GoTo Invalid
            Set current = ReadChain(target, CStr(parts(0)), version)
            If current Is Nothing Then GoTo Invalid
            If current("Guide")("ActionPathId") = guide("ActionPathId") And current("Guide")("ContentSha256") = guide("ContentSha256") Then
                If profile Is Nothing Then
                    Set profile = current
                Else
                    If profile("ProfileId") <> current("ProfileId") Then GoTo Invalid
                    If current("Version") > profile("Version") Then Set profile = current
                End If
            End If
        End If
    Next file
    Latest = True: Exit Function
Invalid:
    Set profile = Nothing
    notice = "Unavailable: execution profiles are invalid or ambiguous. Existing files were preserved."
End Function

Public Function Append(ByVal target As WarehouseTarget, ByVal model As Object, ByRef hash As String, ByRef notice As String) As Boolean
    Dim body As String, content As String, root As String, path As String, pending As String
    Dim fso As Object, stream As Object, decoded As Object, previous As Object
    On Error GoTo Failed
    hash = "": notice = "Execution profile could not be saved. Check the required training inputs."
    If Not modExecutionProfileModel.Validate(target, model) Then Exit Function
    body = modTrainingJson.EncodeObject(model)
    If Len(body) + 83 > 1048576 Then notice = "Execution profile exceeds 1 MiB. Nothing was truncated.": Exit Function
    Set decoded = modTrainingJson.DecodeObject(body)
    If decoded Is Nothing Then Exit Function
    If Not modExecutionProfileModel.Validate(target, decoded) Then Exit Function
    hash = modTrainingWire.Sha256(body)
    If Not modEvaluationModel.IsHash(hash) Then Exit Function
    content = Left$(body, Len(body) - 1) & ",""ContentSha256"":""" & hash & """}"
    root = modRecordingJournal.ChildRoot(target, "ExecutionProfiles", True)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, CStr(model("ProfileId")) & "." & CStr(model("Version")) & ".json")
    If fso.FileExists(path) Or fso.FolderExists(path) Then notice = "Execution profile version conflict. Existing versions were preserved.": Exit Function
    If model("Version") > 1 Then
        Set previous = ReadChain(target, CStr(model("ProfileId")), CLng(model("Version")) - 1)
        If previous Is Nothing Then Exit Function
        If previous("RecordId") <> model("PreviousRecordId") Or previous("ContentSha256") <> model("PreviousSha256") Then Exit Function
        If modTrainingJson.EncodeObject(previous("Guide")) <> modTrainingJson.EncodeObject(model("Guide")) Then Exit Function
    End If
    pending = fso.BuildPath(root, modTrainingWire.NewId() & ".pending")
    Set stream = fso.CreateTextFile(pending, False, False)
    stream.Write content: stream.Close: Set stream = Nothing
    If modRecordingJournal.ReadText(pending) <> content Then GoTo Failed
    Name pending As path
    pending = "": Append = True: notice = "": Exit Function
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If pending <> "" And Not fso Is Nothing Then
        If fso.FileExists(pending) Then fso.DeleteFile pending, True
    End If
End Function
