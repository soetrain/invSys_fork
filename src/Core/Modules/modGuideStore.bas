Attribute VB_Name = "modGuideStore"
Option Explicit
Option Private Module

' Shared internal chain validation; no object crosses an XLAM boundary.
Public Function ReadChain(ByVal target As WarehouseTarget, ByVal id As String, ByVal version As Long) As Object
    Dim current As Object, previous As Object, identities As Object, index As Long
    If Not modTrainingWire.ValidId(id) Or version < 1 Then Exit Function
    Set identities = CreateObject("Scripting.Dictionary")
    index = 1
    Do
        Set current = ReadVersion(target, id, index)
        If current Is Nothing Then Exit Function
        If identities.Exists(current("RecordId")) Then Exit Function
        identities.Add current("RecordId"), True
        If Not previous Is Nothing Then
            If current("PreviousRecordId") <> previous("RecordId") Or current("PreviousSha256") <> previous("ContentSha256") Then Exit Function
        End If
        If index = version Then Exit Do
        Set previous = current: index = index + 1
    Loop
    Set ReadChain = current
End Function

' Immutable guide revisions have their own namespace and no journal-entry limit.
Public Function Append(ByVal target As WarehouseTarget, ByVal model As Object, ByRef hash As String, ByRef notice As String) As Boolean
    Dim body As String, content As String, root As String, path As String, pending As String
    Dim fso As Object, stream As Object, decoded As Object, previous As Object
    On Error GoTo Failed
    hash = "": notice = "Unavailable: the guide version could not be saved. Your edits remain in this draft."
    If Not modGuideModel.Validate(target, model) Then Exit Function
    body = modTrainingJson.EncodeObject(model)
    If Len(body) + 83 > 1048576 Then
        notice = "Guide save exceeds 1 MiB. No authored text or observations were truncated.": Exit Function
    End If
    Set decoded = modTrainingJson.DecodeObject(body)
    If decoded Is Nothing Then Exit Function
    If Not modGuideModel.Validate(target, decoded) Then Exit Function
    hash = modTrainingWire.Sha256(body)
    If Not modEvaluationModel.IsHash(hash) Then Exit Function
    content = Left$(body, Len(body) - 1) & ",""ContentSha256"":""" & hash & """}"
    root = modRecordingJournal.ChildRoot(target, "Guides", True)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, CStr(model("ActionPathId")) & "." & CStr(model("Version")) & ".json")
    If fso.FileExists(path) Or fso.FolderExists(path) Then
        notice = "Guide version conflict. Existing versions and your unsaved edits were preserved.": Exit Function
    End If
    If model("Version") > 1 Then
        Set previous = ReadVersion(target, CStr(model("ActionPathId")), CLng(model("Version")) - 1)
        If previous Is Nothing Then Exit Function
        If previous("RecordId") <> model("PreviousRecordId") Or previous("ContentSha256") <> model("PreviousSha256") Then Exit Function
    End If
    pending = fso.BuildPath(root, modTrainingWire.NewId() & ".pending")
    Set stream = fso.CreateTextFile(pending, False, False)
    stream.Write content: stream.Close: Set stream = Nothing
    If modRecordingJournal.ReadText(pending) <> content Then GoTo Failed
    Name pending As path
    pending = "": Append = True: notice = ""
    Exit Function
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If pending <> "" And Not fso Is Nothing Then
        If fso.FileExists(pending) Then fso.DeleteFile pending, True
    End If
End Function

Public Function ReadVersion(ByVal target As WarehouseTarget, ByVal pathId As String, ByVal version As Long) As Object
    Dim root As String, path As String, text As String, body As String, hash As String, marker As Long
    Dim fso As Object, model As Object
    On Error GoTo Invalid
    If Not modTrainingWire.ValidId(pathId) Or version < 1 Then Exit Function
    root = modRecordingJournal.ChildRoot(target, "Guides", False)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, pathId & "." & CStr(version) & ".json")
    If Not fso.FileExists(path) Then Exit Function
    If (fso.GetFile(path).Attributes And &H400) <> 0 Then Exit Function
    text = modRecordingJournal.ReadText(path)
    marker = InStrRev(text, ",""ContentSha256"":""", -1, vbBinaryCompare)
    If marker = 0 Then Exit Function
    body = Left$(text, marker - 1) & "}"
    hash = Mid$(text, marker + Len(",""ContentSha256"":"""), 64)
    If Mid$(text, marker) <> ",""ContentSha256"":""" & hash & """}" Then Exit Function
    If Not modEvaluationModel.IsHash(hash) Or modTrainingWire.Sha256(body) <> hash Then Exit Function
    Set model = modTrainingJson.DecodeObject(body)
    If model Is Nothing Then Exit Function
    If Not modGuideModel.Validate(target, model) Then Exit Function
    If model("ActionPathId") <> pathId Or model("Version") <> version Then Exit Function
    model.Add "ContentSha256", hash
    Set ReadVersion = model
Invalid:
End Function
