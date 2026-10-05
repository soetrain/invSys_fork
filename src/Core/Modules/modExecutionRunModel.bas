Attribute VB_Name = "modExecutionRunModel"
Option Explicit
Option Private Module

Private Const FIELDS As String = "SchemaVersion|RecordKind|RunId|Revision|RecordId|PreviousRecordId|PreviousSha256|WarehouseId|StationId|CreatedByUserId|CreatedAtUTC|ContextSha256|Guide|Profile|PackageSetVersion|PackageBuilds|Mode|State|ReasonCode|Recording|Steps"

Public Function Create(ByVal target As WarehouseTarget, ByVal context As String, ByVal guide As Object, ByVal profile As Object, _
                       ByVal builds As Collection, ByVal mode As String) As Object
    Dim model As Object, recording As Object, steps As New Collection
    Set model = CreateObject("Scripting.Dictionary"): Set recording = CreateObject("Scripting.Dictionary")
    model.Add "SchemaVersion", 1&: model.Add "RecordKind", "ExecutionRun"
    model.Add "RunId", modTrainingWire.NewId(): model.Add "Revision", 1&: model.Add "RecordId", modTrainingWire.NewId()
    model.Add "PreviousRecordId", "": model.Add "PreviousSha256", ""
    model.Add "WarehouseId", target.WarehouseId: model.Add "StationId", target.StationId
    model.Add "CreatedByUserId", modAuth.GetCurrentUserId(): model.Add "CreatedAtUTC", modTrainingWire.UtcTimestamp()
    model.Add "ContextSha256", modTrainingWire.Sha256(context)
    model.Add "Guide", Reference(guide, "ActionPathId"): model.Add "Profile", Reference(profile, "ProfileId")
    model.Add "PackageSetVersion", CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value)
    model.Add "PackageBuilds", builds: model.Add "Mode", mode: model.Add "State", "Ready": model.Add "ReasonCode", ""
    model.Add "Recording", recording: model.Add "Steps", steps
    Set Create = model
End Function

Private Function Reference(ByVal model As Object, ByVal idField As String) As Object
    Dim value As Object, field As Variant
    Set value = CreateObject("Scripting.Dictionary")
    For Each field In Array(idField, "Version", "RecordId", "ContentSha256"): value.Add CStr(field), model(field): Next field
    Set Reference = value
End Function

Public Function Validate(ByVal target As WarehouseTarget, ByVal model As Object) As Boolean
    Dim field As Variant, entry As Object, seen As Object, index As Long
    On Error GoTo Invalid
    If Not modEvaluationModel.HasFields(model, FIELDS) Then Exit Function
    For Each field In model.Keys
        Select Case CStr(field)
            Case "SchemaVersion", "Revision"
                If Not modEvaluationModel.IsInteger(model(field)) Or model(field) < 1 Then Exit Function
            Case "Guide", "Profile", "Recording"
                If TypeName(model(field)) <> "Dictionary" Then Exit Function
            Case "Steps", "PackageBuilds"
                If TypeName(model(field)) <> "Collection" Then Exit Function
            Case Else
                If VarType(model(field)) <> vbString Then Exit Function
        End Select
    Next field
    If model("SchemaVersion") <> 1 Or model("RecordKind") <> "ExecutionRun" Then Exit Function
    If model("WarehouseId") <> target.WarehouseId Or model("StationId") <> target.StationId Then Exit Function
    For Each field In Array("RunId", "RecordId")
        If Not modTrainingWire.ValidId(CStr(model(field))) Then Exit Function
    Next field
    If model("RunId") = model("RecordId") Or model("CreatedByUserId") = "" Or model("PackageSetVersion") = "" Then Exit Function
    If Not modTrainingWire.ValidUtcTimestamp(CStr(model("CreatedAtUTC"))) Or Not modEvaluationModel.IsHash(CStr(model("ContextSha256"))) Then Exit Function
    If model("Revision") = 1 Then
        If model("PreviousRecordId") <> "" Or model("PreviousSha256") <> "" Then Exit Function
    Else
        If Not modTrainingWire.ValidId(CStr(model("PreviousRecordId"))) Or Not modEvaluationModel.IsHash(CStr(model("PreviousSha256"))) Then Exit Function
        If model("PreviousRecordId") = model("RecordId") Then Exit Function
    End If
    If Not ValidReference(model("Guide"), "ActionPathId") Or Not ValidReference(model("Profile"), "ProfileId") Then Exit Function
    If model("Mode") <> "RunAll" And model("Mode") <> "StepThrough" Then Exit Function
    Select Case model("State")
        Case "Ready", "Running", "Stopped", "Completed", "Blocked", "Failed", "Unknown"
        Case Else: Exit Function
    End Select
    If model("Recording").Count > 0 Then
        If Not modEvaluationModel.HasFields(model("Recording"), "ActionPathId|SequenceId") Then Exit Function
        For Each field In model("Recording").Keys
            If VarType(model("Recording")(field)) <> vbString Then Exit Function
            If Not modTrainingWire.ValidId(CStr(model("Recording")(field))) Then Exit Function
        Next field
    End If
    If model("Steps").Count > 256 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each entry In model("Steps")
        If Not ValidStep(entry, target.WarehouseId) Then Exit Function
        If seen.Exists(entry("StepId")) Then Exit Function
        seen.Add entry("StepId"), True
    Next entry
    If model("PackageBuilds").Count < 3 Or model("PackageBuilds").Count > 5 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each entry In model("PackageBuilds")
        If Not modEvaluationModel.HasFields(entry, "PackageId|BuildIdentity") Then Exit Function
        If VarType(entry("PackageId")) <> vbString Or VarType(entry("BuildIdentity")) <> vbString Then Exit Function
        Select Case entry("PackageId")
            Case "invSys.Core.xlam", "invSys.Inventory.Domain.xlam", "invSys.Designs.Domain.xlam", "invSys.Operations.xlam", "invSys.Admin.xlam"
            Case Else: Exit Function
        End Select
        If seen.Exists(entry("PackageId")) Or entry("BuildIdentity") = "" Then Exit Function
        seen.Add entry("PackageId"), True
    Next entry
    Validate = seen.Exists("invSys.Core.xlam") And seen.Exists("invSys.Inventory.Domain.xlam") And seen.Exists("invSys.Operations.xlam")
Invalid:
End Function

Private Function ValidReference(ByVal value As Object, ByVal idField As String) As Boolean
    Dim field As Variant
    If Not modEvaluationModel.HasFields(value, idField & "|Version|RecordId|ContentSha256") Then Exit Function
    If Not modEvaluationModel.IsInteger(value("Version")) Or value("Version") < 1 Then Exit Function
    For Each field In Array(idField, "RecordId", "ContentSha256")
        If VarType(value(field)) <> vbString Then Exit Function
    Next field
    ValidReference = modTrainingWire.ValidId(CStr(value(idField))) And modTrainingWire.ValidId(CStr(value("RecordId"))) And modEvaluationModel.IsHash(CStr(value("ContentSha256")))
End Function

Private Function ValidStep(ByVal value As Object, ByVal warehouseId As String) As Boolean
    Dim field As Variant
    If Not modEvaluationModel.HasFields(value, "StepId|ControlId|State|ReasonCode|ActivityId|SourceEventRefs|Outputs") Then Exit Function
    For Each field In Array("StepId", "ControlId", "State", "ReasonCode", "ActivityId")
        If VarType(value(field)) <> vbString Then Exit Function
    Next field
    If Not modTrainingWire.ValidId(CStr(value("StepId"))) Or Not modExecutionInputs.Supported(CStr(value("ControlId"))) Then Exit Function
    Select Case value("State")
        Case "Completed", "OperatorCompleted", "Blocked", "Failed", "Unknown"
        Case Else: Exit Function
    End Select
    If value("ActivityId") <> "" Then
        If Not modTrainingWire.ValidId(CStr(value("ActivityId"))) Then Exit Function
    ElseIf value("State") = "Completed" Then
        Exit Function
    End If
    If TypeName(value("SourceEventRefs")) <> "Collection" Or TypeName(value("Outputs")) <> "Collection" Then Exit Function
    ' B0 has no declared output binding; source references retain the ordinary wire.
    If value("Outputs").Count <> 0 Then Exit Function
    For Each field In value("SourceEventRefs")
        If Not modEvaluationModel.HasFields(field, "WarehouseId|SourceKind|EventId|SubmissionState") Then Exit Function
        If VarType(field("WarehouseId")) <> vbString Or VarType(field("SourceKind")) <> vbString Or VarType(field("EventId")) <> vbString Or VarType(field("SubmissionState")) <> vbString Then Exit Function
        If field("WarehouseId") <> warehouseId Or field("SourceKind") <> "Inventory" Or field("EventId") = "" Then Exit Function
        If field("SubmissionState") <> "Submitted" And field("SubmissionState") <> "Unknown" Then Exit Function
    Next field
    ValidStep = True
End Function
