Attribute VB_Name = "modGuideModel"
Option Explicit
Option Private Module

Private Const FIELDS As String = "SchemaVersion|RecordKind|ActionPathId|Version|RecordId|PreviousRecordId|PreviousSha256|WarehouseId|OriginWarehouseId|CreatedByUserId|CreatedAtUTC|Lifecycle|Name|Tags|Instructions|CatalogVersion|PackageSetVersion|BuildIdentity|PolicyVersion|Steps|Observations|SourceRun|ExpectedConclusion"
Private Const SOURCE_FIELDS As String = "ActionPathId|SequenceId|JournalVersion|RecordId|ContentSha256|RecordedByUserId|EntryCreatedAtUTC|Lifecycle|ReasonCode|ActionCount|CapturePolicyVersion|CatalogVersion|PackageSetVersion|BuildIdentity|RestrictedObservationCount"

' Build authored content separately from unchanged original activity bodies.
Public Function Create(ByVal header As Object, ByVal text As Object, ByVal draftSteps As Collection, _
                       ByVal records As Collection, ByVal visible As Object, ByVal policyVersion As Long, _
                       ByVal previous As Object, ByVal expectation As Object) As Object
    Dim model As Object, source As Object, field As Variant, tag As Variant, step As Object, authored As Object, record As Object
    Dim tags As New Collection, steps As New Collection, observations As New Collection, hidden As Long
    Set model = CreateObject("Scripting.Dictionary")
    model.Add "SchemaVersion", 1&: model.Add "RecordKind", "Guide"
    If previous Is Nothing Then
        model.Add "ActionPathId", modTrainingWire.NewId(): model.Add "Version", 1&
        model.Add "RecordId", modTrainingWire.NewId(): model.Add "PreviousRecordId", "": model.Add "PreviousSha256", ""
    Else
        If previous("Version") = 2147483647 Then Err.Raise 6
        model.Add "ActionPathId", previous("ActionPathId"): model.Add "Version", CLng(previous("Version")) + 1&
        model.Add "RecordId", modTrainingWire.NewId(): model.Add "PreviousRecordId", previous("RecordId")
        model.Add "PreviousSha256", previous("ContentSha256")
    End If
    model.Add "WarehouseId", header("WarehouseId"): model.Add "OriginWarehouseId", header("WarehouseId")
    model.Add "CreatedByUserId", modAuth.GetCurrentUserId(): model.Add "CreatedAtUTC", modTrainingWire.UtcTimestamp()
    model.Add "Lifecycle", "Published": model.Add "Name", Trim$(CStr(text("Name")))
    For Each tag In Split(CStr(text("Tags")), ",")
        If Trim$(CStr(tag)) <> "" Then tags.Add Trim$(CStr(tag))
    Next tag
    model.Add "Tags", tags: model.Add "Instructions", CStr(text("Instructions"))
    model.Add "CatalogVersion", modActivityCatalog.CATALOG_VERSION
    model.Add "PackageSetVersion", CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value)
    model.Add "BuildIdentity", CStr(ThisWorkbook.CustomDocumentProperties("invSysBuildIdentity").Value)
    model.Add "PolicyVersion", policyVersion
    For Each step In draftSteps
        Set authored = CreateObject("Scripting.Dictionary")
        authored.Add "StepId", step("StepId"): authored.Add "Method", "How-To"
        For Each field In Array("ControlId", "Caption", "Instruction")
            authored.Add CStr(field), step(field)
        Next field
        authored.Add "SourceActivityId", step("ActivityId"): steps.Add authored
    Next step
    For Each record In records
        If modEvaluationMatches.Permitted(visible, CStr(record("ControlId"))) Then observations.Add record Else hidden = hidden + 1
    Next record
    model.Add "Steps", steps: model.Add "Observations", observations
    Set source = CreateObject("Scripting.Dictionary")
    For Each field In Array("ActionPathId", "SequenceId", "RecordId", "ContentSha256", "Lifecycle", "ReasonCode", "ActionCount", "CatalogVersion", "PackageSetVersion", "BuildIdentity")
        source.Add CStr(field), header(field)
    Next field
    source.Add "JournalVersion", header("Version"): source.Add "RecordedByUserId", header("CreatedByUserId")
    source.Add "EntryCreatedAtUTC", header("CreatedAtUTC"): source.Add "CapturePolicyVersion", header("PolicyVersion")
    source.Add "RestrictedObservationCount", hidden: model.Add "SourceRun", source
    model.Add "ExpectedConclusion", modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(expectation))
    Set Create = model
End Function

Public Function Validate(ByVal target As WarehouseTarget, ByVal model As Object) As Boolean
    Dim field As Variant, tag As Variant, step As Object, record As Object, definition As Object
    Dim identities As Object, actions As Object, stepIds As Object, source As Object, id As String
    On Error GoTo Invalid
    If Not modEvaluationModel.HasFields(model, FIELDS) Then Exit Function
    For Each field In model.Keys
        Select Case CStr(field)
            Case "SchemaVersion", "Version", "CatalogVersion", "PolicyVersion"
                If Not modEvaluationModel.IsInteger(model(field)) Then Exit Function
            Case "Tags", "Steps", "Observations"
                If TypeName(model(field)) <> "Collection" Then Exit Function
            Case "SourceRun", "ExpectedConclusion"
                If TypeName(model(field)) <> "Dictionary" Then Exit Function
            Case Else
                If VarType(model(field)) <> vbString Then Exit Function
        End Select
    Next field
    If model("SchemaVersion") <> 1 Or model("RecordKind") <> "Guide" Or model("Lifecycle") <> "Published" Then Exit Function
    If model("WarehouseId") <> target.WarehouseId Or model("OriginWarehouseId") <> target.WarehouseId Then Exit Function
    If model("Version") < 1 Or model("PolicyVersion") < 0 Then Exit Function
    If model("CatalogVersion") < 1 Or model("CatalogVersion") > modActivityCatalog.CATALOG_VERSION Then Exit Function
    For Each field In Array("CreatedByUserId", "PackageSetVersion", "BuildIdentity", "Name")
        If Trim$(CStr(model(field))) = "" Then Exit Function
    Next field
    If model("Name") <> Trim$(CStr(model("Name"))) Or Not modTrainingWire.ValidUtcTimestamp(CStr(model("CreatedAtUTC"))) Then Exit Function
    Set identities = CreateObject("Scripting.Dictionary"): Set actions = CreateObject("Scripting.Dictionary")
    For Each field In Array("ActionPathId", "RecordId")
        id = CStr(model(field))
        If Not modTrainingWire.ValidId(id) Or identities.Exists(id) Then Exit Function
        identities.Add id, True
    Next field
    If model("Version") = 1 Then
        If model("PreviousRecordId") <> "" Or model("PreviousSha256") <> "" Then Exit Function
    Else
        id = CStr(model("PreviousRecordId"))
        If Not modTrainingWire.ValidId(id) Or identities.Exists(id) Or Not modEvaluationModel.IsHash(CStr(model("PreviousSha256"))) Then Exit Function
        identities.Add id, True
    End If
    For Each tag In model("Tags")
        If VarType(tag) <> vbString Then Exit Function
        If tag = "" Or tag <> Trim$(CStr(tag)) Then Exit Function
    Next tag
    If Not modExpectationModel.Validate(model("ExpectedConclusion"), CLng(model("CatalogVersion"))) Then Exit Function
    Set source = model("SourceRun")
    If Not ValidSource(source, identities) Then Exit Function
    For Each record In model("Observations")
        id = CStr(record("RecordId"))
        If Not modTrainingWire.ValidId(id) Or identities.Exists(id) Then Exit Function
        If Not modActivityStore.ValidBody(target, id, modTrainingJson.EncodeObject(record)) Then Exit Function
        identities.Add id, True
        If source.Count > 0 Then
            If record("SequenceId") <> source("SequenceId") Or record("UserId") <> source("RecordedByUserId") Then Exit Function
            If record("PolicyVersion") <> source("CapturePolicyVersion") Or record("Ordinal") > source("ActionCount") Then Exit Function
        End If
        If record("OutcomeCode") = "REQUESTED" Then
            id = CStr(record("ActivityId"))
            If actions.Exists(id) Then Exit Function
            ' REQUESTED uses its ActivityId as its own RecordId in the source schema.
            If identities.Exists(id) And id <> CStr(record("RecordId")) Then Exit Function
            actions.Add id, record
        End If
    Next record
    If model("Steps").Count < 1 Then Exit Function
    Set stepIds = CreateObject("Scripting.Dictionary")
    For Each step In model("Steps")
        If Not modEvaluationModel.HasFields(step, "StepId|Method|ControlId|Caption|Instruction|SourceActivityId") Then Exit Function
        For Each field In step.Keys
            If VarType(step(field)) <> vbString Then Exit Function
        Next field
        id = CStr(step("StepId"))
        If Not modTrainingWire.ValidId(id) Or stepIds.Exists(id) Or identities.Exists(id) Or actions.Exists(id) Then Exit Function
        stepIds.Add id, True
        If step("Method") <> "How-To" Or Not actions.Exists(step("SourceActivityId")) Then Exit Function
        Set record = actions(step("SourceActivityId"))
        If step("ControlId") <> record("ControlId") Or step("Caption") <> record("Caption") Then Exit Function
        Set definition = modActivityCatalog.Control(CStr(step("ControlId")), CLng(model("CatalogVersion")))
        If definition Is Nothing Then Exit Function
    Next step
    Validate = True
Invalid:
End Function

Private Function ValidSource(ByVal source As Object, ByVal identities As Object) As Boolean
    Dim field As Variant, id As String
    If source.Count = 0 Then ValidSource = True: Exit Function
    If Not modEvaluationModel.HasFields(source, SOURCE_FIELDS) Then Exit Function
    For Each field In source.Keys
        Select Case CStr(field)
            Case "JournalVersion", "ActionCount", "CapturePolicyVersion", "CatalogVersion", "RestrictedObservationCount"
                If Not modEvaluationModel.IsInteger(source(field)) Then Exit Function
                If source(field) < 0 Then Exit Function
            Case Else
                If VarType(source(field)) <> vbString Then Exit Function
        End Select
    Next field
    For Each field In Array("ActionPathId", "SequenceId", "RecordId")
        id = CStr(source(field))
        If Not modTrainingWire.ValidId(id) Or identities.Exists(id) Then Exit Function
        identities.Add id, True
    Next field
    If source("JournalVersion") < 1 Or source("JournalVersion") > 514 Or source("ActionCount") > 256 Or source("RestrictedObservationCount") > 512 Then Exit Function
    If source("CatalogVersion") < 1 Or source("CatalogVersion") > modActivityCatalog.CATALOG_VERSION Then Exit Function
    For Each field In Array("RecordedByUserId", "PackageSetVersion", "BuildIdentity")
        If source(field) = "" Then Exit Function
    Next field
    If Not modTrainingWire.ValidUtcTimestamp(CStr(source("EntryCreatedAtUTC"))) Or Not modEvaluationModel.IsHash(CStr(source("ContentSha256"))) Then Exit Function
    Select Case CStr(source("Lifecycle"))
        Case "Recording", "Stopped", "Cancelled"
            If source("ReasonCode") <> "" Then Exit Function
        Case "Incomplete"
            If source("ReasonCode") = "" Or InStr("|SESSION_CHANGED|VIEWER_CLOSED|POLICY_CHANGED|TRACKING_UNAVAILABLE|ACTION_LIMIT|UNFINISHED_ACTIONS|", "|" & CStr(source("ReasonCode")) & "|") = 0 Then Exit Function
        Case Else: Exit Function
    End Select
    ValidSource = True
End Function
