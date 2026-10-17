Attribute VB_Name = "modAggregationRetry"
Option Explicit
Option Private Module

Public Function Pending(ByVal destination As String) As Boolean
    Dim fso As Object, folder As Object, file As Object
    On Error GoTo Unknown
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fso.GetParentFolderName(destination)) Then Exit Function
    Set folder = fso.GetFolder(fso.GetParentFolderName(destination))
    For Each file In folder.Files
        If LCase$(file.Name) Like "invsys.aggregation.*.retry.json" Then Pending = True: Exit Function
    Next file
    Exit Function
Unknown:
    Pending = True
End Function

Public Function Reserve(ByVal destination As String, ByVal sources As Collection) As Object
    Dim record As Object, inputs As New Collection, source As Variant, item As Object, folder As String, id As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(destination)
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    id = modSystemIdentity.NewId()
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "Version", CLng(1): record.Add "Attempt", id
    record.Add "Destination", destination
    record.Add "Candidate", fso.BuildPath(folder, "invSys.Aggregation." & id & ".candidate.xlsb")
    record.Add "Promotion", fso.BuildPath(folder, "invSys.Aggregation." & id & ".promotion.xlsb")
    record.Add "Backup", fso.BuildPath(folder, "invSys.Aggregation." & id & ".previous.xlsb")
    record.Add "Receipt", fso.BuildPath(folder, "invSys.Aggregation." & id & ".retry.json")
    record.Add "PriorExists", fso.FileExists(destination)
    record.Add "PriorHash", modTrainingWire.FileSha256(destination)
    If record("PriorExists") And record("PriorHash") = "" Then Err.Raise vbObjectError + 7490
    record.Add "CandidateHash", "": record.Add "Phase", "STAGING_PENDING"
    record.Add "Persistence", "PRESERVED"
    For Each source In sources
        Set item = CreateObject("Scripting.Dictionary")
        item.Add "Path", source("Path"): item.Add "Hash", source("Hash")
        inputs.Add item
    Next source
    record.Add "Sources", inputs
    Save record, True
    Set Reserve = record
End Function

Public Sub Save(ByVal record As Object, Optional ByVal first As Boolean = False)
    Dim stream As Object, fso As Object, path As String, text As String, actual As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = CStr(record("Receipt")): text = modTrainingJson.EncodeObject(record)
    If first Then
        Set stream = fso.CreateTextFile(path, False, False)
    Else
        Set stream = fso.CreateTextFile(path & ".pending", False, False)
    End If
    stream.Write text: stream.Close
    If Not first Then
        If Not modAggregationFiles.PublishRecord(path & ".pending", path) Then Err.Raise vbObjectError + 7491
    End If
    Set actual = Read(path)
    If actual Is Nothing Then Err.Raise vbObjectError + 7492
    If modTrainingJson.EncodeObject(actual) <> text Then Err.Raise vbObjectError + 7493
End Sub

Public Function Read(ByVal path As String) As Object
    Dim stream As Object, fso As Object
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetFile(path).Size > 1048576 Then Exit Function
    Set stream = fso.OpenTextFile(path, 1, False)
    Set Read = modTrainingJson.DecodeObject(stream.ReadAll)
    stream.Close
Failed:
End Function

Public Function Persistence(ByVal record As Object) As String
    Dim hash As String, exists As Boolean
    On Error GoTo Unknown
    exists = CreateObject("Scripting.FileSystemObject").FileExists(CStr(record("Destination")))
    hash = modTrainingWire.FileSha256(CStr(record("Destination")))
    If Not exists And Not record("PriorExists") Then Persistence = "PRESERVED": Exit Function
    If hash <> "" Then
        If record("CandidateHash") <> "" And hash = record("CandidateHash") Then Persistence = "COMPLETE": Exit Function
        If record("PriorExists") And hash = record("PriorHash") Then Persistence = "PRESERVED": Exit Function
    End If
Unknown:
    If Persistence = "" Then Persistence = "UNKNOWN"
End Function

Public Sub ResolveComplete(ByVal record As Object)
    Dim fso As Object, field As Variant, folder As String, expected As String, path As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Persistence(record) <> "COMPLETE" Then Err.Raise vbObjectError + 7494
    folder = fso.GetParentFolderName(CStr(record("Destination")))
    expected = "invSys.Aggregation." & CStr(record("Attempt")) & "."
    For Each field In Array("Candidate", "Promotion", "Backup")
        path = CStr(record(CStr(field)))
        If StrComp(fso.GetParentFolderName(path), folder, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7495
        If Left$(fso.GetFileName(path), Len(expected)) <> expected Then Err.Raise vbObjectError + 7496
        If fso.FileExists(path) Then fso.DeleteFile path, False
    Next field
    fso.DeleteFile CStr(record("Receipt")), False
End Sub
