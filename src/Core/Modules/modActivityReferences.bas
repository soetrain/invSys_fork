Attribute VB_Name = "modActivityReferences"
Option Explicit
Option Private Module

Public Function Decode(ByVal warehouseId As String, ByVal controlId As String, _
                       ByVal outcomeCode As String, ByVal json As String) As Collection
    Dim wrapper As Object, references As Collection
    Set wrapper = modTrainingJson.DecodeObject("{""SourceEventRefs"":" & json & "}")
    If wrapper Is Nothing Then Exit Function
    If wrapper.Count <> 1 Then Exit Function
    If TypeName(wrapper("SourceEventRefs")) <> "Collection" Then Exit Function
    Set references = wrapper("SourceEventRefs")
    If Valid(warehouseId, controlId, outcomeCode, references) Then Set Decode = references
End Function

Public Function Valid(ByVal warehouseId As String, ByVal controlId As String, _
                      ByVal outcomeCode As String, ByVal references As Collection) As Boolean
    Dim reference As Variant, field As Variant, seen As Object, key As String
    On Error GoTo Invalid
    If (controlId <> "RECEIVING_CONFIRM_WRITES" And controlId <> "DISPOSITION_CONFIRM") Or outcomeCode = "REQUESTED" Or _
       outcomeCode = "DENIED" Or outcomeCode = "REJECTED" Then
        Valid = (references.Count = 0)
        Exit Function
    End If
    If outcomeCode = "CONFIRMED" Or outcomeCode = "PENDING" Then
        If references.Count = 0 Then Exit Function
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    For Each reference In references
        If Not IsObject(reference) Then Exit Function
        If TypeName(reference) <> "Dictionary" Then Exit Function
        If reference.Count <> 4 Then Exit Function
        For Each field In Array("WarehouseId", "SourceKind", "EventId", "SubmissionState")
            If Not reference.Exists(field) Then Exit Function
            If VarType(reference(field)) <> vbString Then Exit Function
        Next field
        If reference("WarehouseId") <> warehouseId Or reference("SourceKind") <> "Inventory" Then Exit Function
        key = reference("EventId")
        If Not ValidIdentity(key) Or seen.Exists(key) Then Exit Function
        seen.Add key, True
        key = reference("SubmissionState")
        If key <> "Submitted" And key <> "Unknown" Then Exit Function
        If outcomeCode = "CONFIRMED" Or outcomeCode = "PENDING" Then
            If key <> "Submitted" Then Exit Function
        End If
    Next reference
    Valid = True
Invalid:
End Function

Public Function Encode(ByVal references As Collection) As String
    Const PREFIX As String = "{""SourceEventRefs"":"
    Dim wrapper As Object, json As String
    Set wrapper = CreateObject("Scripting.Dictionary")
    wrapper.Add "SourceEventRefs", references
    json = modTrainingJson.EncodeObject(wrapper)
    Encode = Mid$(json, Len(PREFIX) + 1, Len(json) - Len(PREFIX) - 1)
End Function

Public Function Inventory(ByVal warehouseId As String, ByVal eventIds As String, _
                          ByVal submissionState As String) As String
    Dim references As New Collection, reference As Object, id As Variant
    On Error GoTo Invalid
    If eventIds <> "" Then
        For Each id In Split(eventIds, vbLf)
            If Not ValidIdentity(CStr(id)) Then Exit Function
            Set reference = CreateObject("Scripting.Dictionary")
            reference.Add "WarehouseId", warehouseId
            reference.Add "SourceKind", "Inventory"
            reference.Add "EventId", CStr(id)
            reference.Add "SubmissionState", submissionState
            references.Add reference
        Next id
    End If
    Inventory = Encode(references)
Invalid:
End Function

Private Function ValidIdentity(ByVal value As String) As Boolean
    Dim i As Long, ch As String
    If Len(value) = 0 Or Len(value) > 128 Then Exit Function
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If Not ch Like "[A-Za-z0-9_-]" Then Exit Function
    Next i
    ValidIdentity = True
End Function
