Attribute VB_Name = "modRecordingJournal"
Option Explicit
Option Private Module

Public Function Append(ByVal target As WarehouseTarget, ByVal model As Object, _
                       ByRef hash As String, ByRef notice As String) As Boolean
    Dim body As String, content As String, root As String, path As String, pending As String
    Dim fso As Object, stream As Object, previous As Object
    On Error GoTo Failed
    notice = "Recording unavailable: the journal could not be saved."
    If Not modRecordingModel.Validate(target, model) Then Exit Function
    body = modTrainingJson.EncodeObject(model)
    If Len(body) + 83 > 1048576 Then
        notice = "Incomplete evidence: recording save exceeds 1 MiB. No observations were truncated."
        Exit Function
    End If
    hash = modTrainingWire.Sha256(body)
    If Len(hash) <> 64 Then Exit Function
    content = Left$(body, Len(body) - 1) & ",""ContentSha256"":""" & hash & """}"
    If Len(content) > 1048576 Then Exit Function
    root = JournalRoot(target, True)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, CStr(model("ActionPathId")) & "." & CStr(model("Version")) & ".json")
    If model("Version") > 1 Then
        Set previous = ReadEntry(target, CStr(model("ActionPathId")), CLng(model("Version")) - 1)
        If previous Is Nothing Then Exit Function
        If previous("RecordId") <> model("PreviousRecordId") Or previous("ContentSha256") <> model("PreviousSha256") Then Exit Function
        If previous("Lifecycle") <> "Recording" Or previous("SequenceId") <> model("SequenceId") Then Exit Function
        If previous("CreatedByUserId") <> model("CreatedByUserId") Or previous("PolicyVersion") <> model("PolicyVersion") Then Exit Function
        If previous("ActionCount") > model("ActionCount") Then Exit Function
    End If
    If fso.FileExists(path) Then
        Append = (ReadText(path) = content)
        If Append Then notice = ""
        Exit Function
    End If
    pending = fso.BuildPath(root, modTrainingWire.NewId() & ".pending")
    Set stream = fso.CreateTextFile(pending, False, False)
    stream.Write content: stream.Close: Set stream = Nothing
    If ReadText(pending) <> content Then GoTo Failed
    Name pending As path
    pending = "": Append = True: notice = ""
    Exit Function
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If pending <> "" Then
        If Not fso Is Nothing Then
            If fso.FileExists(pending) Then fso.DeleteFile pending, True
        End If
    End If
End Function

Public Function ReadEntry(ByVal target As WarehouseTarget, ByVal pathId As String, ByVal version As Long) As Object
    Dim root As String, text As String, body As String, hash As String, marker As Long, model As Object
    On Error GoTo Invalid
    If Not modTrainingWire.ValidId(pathId) Or version < 1 Or version > 514 Then Exit Function
    root = JournalRoot(target, False)
    If root = "" Then Exit Function
    text = ReadText(root & "\" & pathId & "." & CStr(version) & ".json")
    marker = InStrRev(text, ",""ContentSha256"":""", -1, vbBinaryCompare)
    If marker = 0 Then Exit Function
    body = Left$(text, marker - 1) & "}"
    hash = Mid$(text, marker + Len(",""ContentSha256"":"""), 64)
    If Mid$(text, marker) <> ",""ContentSha256"":""" & hash & """}" Then Exit Function
    If modTrainingWire.Sha256(body) <> hash Or Len(hash) <> 64 Then Exit Function
    Set model = modTrainingJson.DecodeObject(body)
    If model Is Nothing Then Exit Function
    If Not modRecordingModel.Validate(target, model) Then Exit Function
    If model("ActionPathId") <> pathId Or model("Version") <> version Then Exit Function
    model.Add "ContentSha256", hash
    Set ReadEntry = model
Invalid:
End Function

Public Function JournalRoot(ByVal target As WarehouseTarget, ByVal create As Boolean) As String
    Dim root As String, part As Variant, fso As Object
    If Not modNasConnection.IsWarehouseTargetAllowed(target, True) Then Exit Function
    If Not modTrainingWire.ValidSegment(target.WarehouseId) Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = target.RuntimeRoot
    If Not fso.FolderExists(root) Then Exit Function
    For Each part In Array("Training", "ActionPaths", target.WarehouseId)
        root = fso.BuildPath(root, CStr(part))
        If Not fso.FolderExists(root) Then
            If Not create Or fso.FileExists(root) Then Exit Function
            fso.CreateFolder root
        End If
        If (fso.GetFolder(root).Attributes And &H400) <> 0 Then Exit Function
    Next part
    JournalRoot = root
End Function

Public Function ReadText(ByVal path As String) As String
    Dim fso As Object, stream As Object, text As String
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    If fso.GetFile(path).Size < 1 Or fso.GetFile(path).Size > 1048576 Then Exit Function
    Set stream = fso.OpenTextFile(path, 1, False, 0)
    text = stream.Read(1048577)
    If Len(text) <= 1048576 And Len(modTrainingWire.Sha256(text)) = 64 Then ReadText = text
Done:
    If Not stream Is Nothing Then stream.Close
End Function

' Fixed children share the journal's allowed-target and ancestor checks.
Public Function ChildRoot(ByVal target As WarehouseTarget, ByVal child As String, ByVal create As Boolean) As String
    Dim root As String, fso As Object
    If child <> "Guides" And child <> "Evaluations" Then Exit Function
    root = JournalRoot(target, create)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = fso.BuildPath(root, child)
    If Not fso.FolderExists(root) Then
        If Not create Or fso.FileExists(root) Then Exit Function
        fso.CreateFolder root
    End If
    If (fso.GetFolder(root).Attributes And &H400) = 0 Then ChildRoot = root
End Function
