Attribute VB_Name = "modGuideTransferWire"
Option Explicit
Option Private Module

Public Function Seal(ByVal body As String) As String
    Dim hash As String
    If Len(body) < 2 Or Len(body) > 1048576 Then Exit Function
    hash = modTrainingWire.Sha256(body)
    If Not modEvaluationModel.IsHash(hash) Then Exit Function
    Seal = Left$(body, Len(body) - 1) & ",""ContentSha256"":""" & hash & """}"
    If Len(Seal) > 1048576 Then Seal = ""
End Function

Private Function OpenRecord(ByVal text As String, ByRef body As String, ByRef hash As String) As Object
    Dim marker As Long
    body = "": hash = ""
    If Len(text) = 0 Or Len(text) > 1048576 Then Exit Function
    marker = InStrRev(text, ",""ContentSha256"":""", -1, vbBinaryCompare)
    If marker = 0 Then Exit Function
    hash = Mid$(text, marker + Len(",""ContentSha256"":"""), 64)
    If Mid$(text, marker) <> ",""ContentSha256"":""" & hash & """}" Then Exit Function
    body = Left$(text, marker - 1) & "}"
    If Not modEvaluationModel.IsHash(hash) Or modTrainingWire.Sha256(body) <> hash Then Exit Function
    Set OpenRecord = modTrainingJson.DecodeObject(body)
End Function

Public Function Read(ByVal path As String, ByRef notice As String) As Object
    Dim model As Object, guide As Object, target As WarehouseTarget, field As Variant
    Dim text As String, body As String, hash As String, guideBody As String, guideHash As String
    On Error GoTo Invalid
    notice = "Unavailable: the transfer file has invalid integrity, schema, bounds or provenance."
    text = modRecordingJournal.ReadText(path)
    Set model = OpenRecord(text, body, hash)
    If model Is Nothing Then Exit Function
    If Not model.Exists("SchemaVersion") Then Exit Function
    If Not modEvaluationModel.IsInteger(model("SchemaVersion")) Then Exit Function
    If model("SchemaVersion") <> 1 Then
        notice = "Unsupported transfer version. Executable content was not imported or discarded.": Exit Function
    End If
    If Not modEvaluationModel.HasFields(model, "SchemaVersion|RecordKind|TransferId|ExportedAtUTC|ExportedByUserId|SourceWarehouseId|Guide") Then Exit Function
    For Each field In Array("RecordKind", "TransferId", "ExportedAtUTC", "ExportedByUserId", "SourceWarehouseId")
        If VarType(model(field)) <> vbString Then Exit Function
    Next field
    If model("RecordKind") <> "GuideTransfer" Or TypeName(model("Guide")) <> "Dictionary" Then Exit Function
    If Not modTrainingWire.ValidId(CStr(model("TransferId"))) Or Not modTrainingWire.ValidUtcTimestamp(CStr(model("ExportedAtUTC"))) Then Exit Function
    If Trim$(CStr(model("ExportedByUserId"))) = "" Or model("ExportedByUserId") <> Trim$(CStr(model("ExportedByUserId"))) Then Exit Function
    If Not modTrainingWire.ValidSegment(CStr(model("SourceWarehouseId"))) Then Exit Function
    Set guide = OpenRecord(modTrainingJson.MemberText(body, "Guide"), guideBody, guideHash)
    If guide Is Nothing Then Exit Function
    Set target = New WarehouseTarget
    target.WarehouseId = CStr(model("SourceWarehouseId"))
    If Not modGuideModel.Validate(target, guide) Then Exit Function
    guide.Add "ContentSha256", guideHash
    Set model("Guide") = guide
    model.Add "ContentSha256", hash
    Set Read = model: notice = ""
Invalid:
End Function

Public Function WriteNew(ByVal path As String, ByVal content As String, ByRef notice As String, Optional ByRef outcome As String = "") As Boolean
    Dim fso As Object, stream As Object, parent As String, pending As String, created As Boolean
    On Error GoTo Failed
    outcome = "REJECTED"
    notice = "Unavailable: the guide package could not be exported."
    If path = "" Or content = "" Or Len(content) > 1048576 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.GetAbsolutePathName(path): parent = fso.GetParentFolderName(path)
    If fso.GetFileName(path) = "" Or Not fso.FolderExists(parent) Then Exit Function
    If fso.FileExists(path) Or fso.FolderExists(path) Then
        notice = "Export destination already exists. Choose a new file.": Exit Function
    End If
    pending = fso.BuildPath(parent, modTrainingWire.NewId() & ".pending")
    Set stream = fso.CreateTextFile(pending, False, False)
    created = True
    stream.Write content: stream.Close: Set stream = Nothing
    If modRecordingJournal.ReadText(pending) <> content Then GoTo Failed
    Name pending As path
    pending = "": WriteNew = True: notice = "Guide exported.": outcome = "COMPLETED"
    Exit Function
Failed:
    outcome = "FAILED"
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If created And pending <> "" And Not fso Is Nothing Then
        If fso.FileExists(pending) Then fso.DeleteFile pending, True
    End If
End Function
