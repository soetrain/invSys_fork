Attribute VB_Name = "modExecutionRunStore"
Option Explicit
Option Private Module

Public Function ReadVersion(ByVal target As WarehouseTarget, ByVal id As String, ByVal revision As Long) As Object
    Dim root As String, path As String, text As String, body As String, hash As String, marker As Long, model As Object, fso As Object
    On Error GoTo Invalid
    If Not modTrainingWire.ValidId(id) Or revision < 1 Then Exit Function
    root = modRecordingJournal.ChildRoot(target, "Runs", False)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, id & "." & CStr(revision) & ".json")
    If Not fso.FileExists(path) Then Exit Function
    If (fso.GetFile(path).Attributes And &H400) <> 0 Then Exit Function
    text = modRecordingJournal.ReadText(path)
    marker = InStrRev(text, ",""ContentSha256"":""", -1, vbBinaryCompare)
    If marker = 0 Then Exit Function
    body = Left$(text, marker - 1) & "}": hash = Mid$(text, marker + Len(",""ContentSha256"":"""), 64)
    If Mid$(text, marker) <> ",""ContentSha256"":""" & hash & """}" Then Exit Function
    If modTrainingWire.Sha256(body) <> hash Then Exit Function
    Set model = modTrainingJson.DecodeObject(body)
    If model Is Nothing Then Exit Function
    If Not modExecutionRunModel.Validate(target, model) Then Exit Function
    If model("RunId") <> id Or model("Revision") <> revision Then Exit Function
    model.Add "ContentSha256", hash: Set ReadVersion = model
Invalid:
End Function

Public Function Append(ByVal target As WarehouseTarget, ByVal model As Object, ByRef hash As String, ByRef notice As String) As Boolean
    Dim body As String, content As String, root As String, path As String, pending As String
    Dim fso As Object, stream As Object, previous As Object, decoded As Object, field As Variant, index As Long
    On Error GoTo Failed
    notice = "Run evidence could not be saved. Further dispatch is stopped; existing work was preserved."
    hash = ""
    If Not modExecutionRunModel.Validate(target, model) Then Exit Function
    body = modTrainingJson.EncodeObject(model)
    If Len(body) + 83 > 1048576 Then Exit Function
    Set decoded = modTrainingJson.DecodeObject(body)
    If decoded Is Nothing Then Exit Function
    If Not modExecutionRunModel.Validate(target, decoded) Then Exit Function
    hash = modTrainingWire.Sha256(body)
    If Not modEvaluationModel.IsHash(hash) Then Exit Function
    root = modRecordingJournal.ChildRoot(target, "Runs", True)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, CStr(model("RunId")) & "." & CStr(model("Revision")) & ".json")
    If fso.FileExists(path) Or fso.FolderExists(path) Then Exit Function
    If model("Revision") > 1 Then
        Set previous = ReadVersion(target, CStr(model("RunId")), CLng(model("Revision")) - 1)
        If previous Is Nothing Then Exit Function
        If model("PreviousRecordId") <> previous("RecordId") Or model("PreviousSha256") <> previous("ContentSha256") Then Exit Function
        For Each field In Array("ContextSha256", "CreatedByUserId", "WarehouseId", "StationId", "Mode", "PackageSetVersion")
            If model(field) <> previous(field) Then Exit Function
        Next field
        For Each field In Array("Guide", "Profile")
            If modTrainingJson.EncodeObject(model(field)) <> modTrainingJson.EncodeObject(previous(field)) Then Exit Function
        Next field
        If previous("Steps").Count > model("Steps").Count Then Exit Function
        For index = 1 To previous("Steps").Count
            If modTrainingJson.EncodeObject(model("Steps")(index)) <> modTrainingJson.EncodeObject(previous("Steps")(index)) Then Exit Function
        Next index
        If modExecutionBinding.BuildsHash(model("PackageBuilds")) <> modExecutionBinding.BuildsHash(previous("PackageBuilds")) Then Exit Function
        If previous("Recording").Count > 0 Then
            If modTrainingJson.EncodeObject(model("Recording")) <> modTrainingJson.EncodeObject(previous("Recording")) Then Exit Function
        End If
    End If
    content = Left$(body, Len(body) - 1) & ",""ContentSha256"":""" & hash & """}"
    pending = fso.BuildPath(root, modTrainingWire.NewId() & ".pending")
    Set stream = fso.CreateTextFile(pending, False, False)
    stream.Write content: stream.Close: Set stream = Nothing
    If modRecordingJournal.ReadText(pending) <> content Then GoTo Failed
    Name pending As path
    pending = "": Append = True: notice = ""
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If pending <> "" And Not fso Is Nothing Then
        If fso.FileExists(pending) Then fso.DeleteFile pending, True
    End If
End Function
