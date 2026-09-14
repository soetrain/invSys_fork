Attribute VB_Name = "modEvaluationStore"
Option Explicit
Option Private Module

Private Function StoreRoot(ByVal target As WarehouseTarget, ByVal create As Boolean) As String
    Dim root As String, fso As Object
    root = modRecordingJournal.JournalRoot(target, create)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = fso.BuildPath(root, "Evaluations")
    If Not fso.FolderExists(root) Then
        If Not create Or fso.FileExists(root) Then Exit Function
        fso.CreateFolder root
    End If
    If (fso.GetFolder(root).Attributes And &H400) <> 0 Then Exit Function
    StoreRoot = root
End Function

Public Function Append(ByVal target As WarehouseTarget, ByVal model As Object, ByRef notice As String) As Boolean
    Dim root As String, body As String, content As String, hash As String, path As String, pending As String
    Dim fso As Object, stream As Object, decoded As Object
    On Error GoTo Failed
    notice = "Incomplete evidence: the diagnostic result could not be saved."
    If Not modEvaluationModel.Validate(model, target.WarehouseId) Then Exit Function
    body = modTrainingJson.EncodeObject(model)
    If Len(body) + 83 > 1048576 Then
        notice = "Incomplete evidence: the diagnostic result exceeds 1 MiB. No evidence was truncated.": Exit Function
    End If
    Set decoded = modTrainingJson.DecodeObject(body)
    If decoded Is Nothing Then Exit Function
    If Not modEvaluationModel.Validate(decoded, target.WarehouseId) Then Exit Function
    hash = modTrainingWire.Sha256(body)
    If Not modEvaluationModel.IsHash(hash) Then Exit Function
    content = Left$(body, Len(body) - 1) & ",""ContentSha256"":""" & hash & """}"
    root = StoreRoot(target, True)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, CStr(model("EvaluationId")) & ".1.json")
    If fso.FileExists(path) Or fso.FolderExists(path) Then Exit Function
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

Public Function Read(ByVal target As WarehouseTarget, ByVal evaluationId As String) As Object
    Dim root As String, path As String, text As String, body As String, hash As String, marker As Long
    Dim fso As Object, model As Object
    On Error GoTo Invalid
    If Not modTrainingWire.ValidId(evaluationId) Then Exit Function
    root = StoreRoot(target, False)
    If root = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(root, evaluationId & ".1.json")
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
    If Not modEvaluationModel.Validate(model, target.WarehouseId) Then Exit Function
    If model("EvaluationId") <> evaluationId Then Exit Function
    Set Read = model
Invalid:
End Function
