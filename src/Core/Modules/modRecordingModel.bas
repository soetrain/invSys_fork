Attribute VB_Name = "modRecordingModel"
Option Explicit
Option Private Module

Public Function Validate(ByVal target As WarehouseTarget, ByVal model As Object) As Boolean
    Dim field As Variant, record As Variant, kind As String, life As String, seen As Object
    Dim fields As String, numbers As String, value As String
    On Error GoTo Invalid
    fields = "|SchemaVersion|RecordKind|RecordType|RecordId|ActionPathId|SequenceId|Version|PreviousRecordId|PreviousSha256|Lifecycle|ReasonCode|ActionCount|CreatedByUserId|CreatedAtUTC|WarehouseId|OriginWarehouseId|PolicyVersion|CatalogVersion|PackageSetVersion|BuildIdentity|Name|Tags|Instructions|Method|Observations|"
    numbers = "|SchemaVersion|Version|ActionCount|PolicyVersion|CatalogVersion|"
    If model.Count <> 25 Then Exit Function
    For Each field In model.Keys
        If InStr(1, fields, "|" & CStr(field) & "|", vbBinaryCompare) = 0 Then Exit Function
        If field = "Tags" Or field = "Observations" Then
            If TypeName(model(field)) <> "Collection" Then Exit Function
        ElseIf InStr(1, numbers, "|" & CStr(field) & "|", vbBinaryCompare) > 0 Then
            If VarType(model(field)) <> vbLong And VarType(model(field)) <> vbInteger Then Exit Function
        Else
            If VarType(model(field)) <> vbString Then Exit Function
        End If
    Next field
    If model("SchemaVersion") <> 1 Or model("RecordKind") <> "Recording" Then Exit Function
    If model("WarehouseId") <> target.WarehouseId Or model("OriginWarehouseId") <> target.WarehouseId Then Exit Function
    For Each field In Array("RecordId", "ActionPathId", "SequenceId")
        If Not modTrainingWire.ValidId(CStr(model(field))) Then Exit Function
    Next field
    If model("RecordId") = model("ActionPathId") Or model("RecordId") = model("SequenceId") Or model("ActionPathId") = model("SequenceId") Then Exit Function
    If model("Version") < 1 Or model("Version") > 514 Then Exit Function
    If model("ActionCount") < 0 Or model("ActionCount") > 256 Then Exit Function
    If model("PolicyVersion") < 0 Or model("CatalogVersion") < 1 Or model("CatalogVersion") > modActivityCatalog.CATALOG_VERSION Then Exit Function
    If model("CreatedByUserId") = "" Or Not modTrainingWire.ValidUtcTimestamp(CStr(model("CreatedAtUTC"))) Then Exit Function
    If model("PackageSetVersion") = "" Or model("BuildIdentity") = "" Then Exit Function
    If model("Name") <> "Recorded sequence" Or model("Method") <> "Diagnostic" Or model("Instructions") <> "" Or model("Tags").Count <> 0 Then Exit Function
    kind = model("RecordType"): life = model("Lifecycle")
    If model("Version") = 1 Then
        If kind <> "Start" Or model("PreviousRecordId") <> "" Or model("PreviousSha256") <> "" Then Exit Function
    Else
        If kind = "Start" Or Not modTrainingWire.ValidId(CStr(model("PreviousRecordId"))) Then Exit Function
        value = model("PreviousSha256")
        If Len(value) <> 64 Or value Like "*[!0-9a-f]*" Then Exit Function
    End If
    Select Case kind
        Case "Start"
            If model("ActionCount") <> 0 Or model("Observations").Count <> 0 Then Exit Function
        Case "Observation"
            If model("Observations").Count <> 1 Then Exit Function
        Case "Close"
            If life <> "Stopped" And life <> "Cancelled" And life <> "Incomplete" Then Exit Function
        Case Else: Exit Function
    End Select
    If kind <> "Close" And life <> "Recording" Then Exit Function
    If life = "Incomplete" Then
        If InStr(1, "|SESSION_CHANGED|VIEWER_CLOSED|POLICY_CHANGED|TRACKING_UNAVAILABLE|ACTION_LIMIT|UNFINISHED_ACTIONS|", "|" & CStr(model("ReasonCode")) & "|", vbBinaryCompare) = 0 Or model("ReasonCode") = "" Then Exit Function
    ElseIf model("ReasonCode") <> "" Then
        Exit Function
    End If
    If model("Observations").Count > 512 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each record In model("Observations")
        If Not modActivityStore.ValidBody(target, CStr(record("RecordId")), modTrainingJson.EncodeObject(record)) Then Exit Function
        If record("SequenceId") <> model("SequenceId") Or record("UserId") <> model("CreatedByUserId") Then Exit Function
        If record("PolicyVersion") <> model("PolicyVersion") Or record("Ordinal") > model("ActionCount") Then Exit Function
        If seen.Exists(record("RecordId")) Then Exit Function
        seen.Add record("RecordId"), True
    Next record
    Validate = True
Invalid:
End Function
