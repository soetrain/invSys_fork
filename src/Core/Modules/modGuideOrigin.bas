Attribute VB_Name = "modGuideOrigin"
Option Explicit
Option Private Module

Public Function Valid(ByVal model As Object) As Boolean
    Dim origin As Object, field As Variant
    On Error GoTo Invalid
    Set origin = model("TransferOrigin")
    If Not modEvaluationModel.HasFields(origin, "TransferId|TransferSha256|SourceWarehouseId|ActionPathId|Version|RecordId|ContentSha256") Then Exit Function
    For Each field In origin.Keys
        If field = "Version" Then
            If Not modEvaluationModel.IsInteger(origin(field)) Then Exit Function
            If origin(field) < 1 Then Exit Function
        Else
            If VarType(origin(field)) <> vbString Then Exit Function
        End If
    Next field
    For Each field In Array("TransferId", "ActionPathId", "RecordId")
        If Not modTrainingWire.ValidId(CStr(origin(field))) Then Exit Function
    Next field
    If origin("ActionPathId") = model("ActionPathId") Or origin("RecordId") = model("RecordId") Then Exit Function
    If Not modTrainingWire.ValidSegment(CStr(origin("SourceWarehouseId"))) Then Exit Function
    If Not modTrainingWire.ValidSegment(CStr(model("OriginWarehouseId"))) Then Exit Function
    If Not modEvaluationModel.IsHash(CStr(origin("TransferSha256"))) Or Not modEvaluationModel.IsHash(CStr(origin("ContentSha256"))) Then Exit Function
    Valid = True
Invalid:
End Function

Public Function Caption(ByVal model As Object) As String
    Dim origin As Object
    If model("SchemaVersion") <> 2 Then Exit Function
    Set origin = model("TransferOrigin")
    Caption = "Imported origin evidence; not locally observed" & vbCrLf & _
        "Source warehouse: " & CStr(origin("SourceWarehouseId")) & "; guide: " & CStr(origin("ActionPathId")) & "; version: " & CStr(origin("Version")) & vbCrLf & _
        "Source SHA-256: " & CStr(origin("ContentSha256"))
End Function
