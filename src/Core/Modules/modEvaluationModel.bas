Attribute VB_Name = "modEvaluationModel"
Option Explicit
Option Private Module

Private Const FIELDS As String = "SchemaVersion|RecordKind|EvaluationId|Version|ActionPathId|SequenceId|JournalVersion|JournalRecordId|JournalSha256|RecordedByUserId|EvaluatedByUserId|EvaluatedAtUTC|WarehouseId|OriginWarehouseId|CatalogVersion|PackageSetVersion|BuildIdentity|CapturePolicyVersion|EvaluationPolicyVersion|ExpectationSource|ExpectedConclusion|ExpectationSha256|Guide|PreviousEvaluationId|ResultState|ReasonCodes|Matches|MissingSteps|FailedSteps|UnavailableSteps|ExtraActivityIds|TerminalSources|Publication"
Private Const NUMBERS As String = "|SchemaVersion|Version|JournalVersion|CatalogVersion|CapturePolicyVersion|EvaluationPolicyVersion|"
Private Const COLLECTIONS As String = "|ReasonCodes|Matches|MissingSteps|FailedSteps|UnavailableSteps|ExtraActivityIds|TerminalSources|"
Private Const REASONS As String = "|NO_EXPECTATION|JOURNAL_UNAVAILABLE|CAPTURE_CANCELLED|CAPTURE_INCOMPLETE|REQUIRED_CAPTURE_UNAVAILABLE|REQUIRED_OBSERVATION_RESTRICTED|REQUIRED_STEP_MISSING|REQUIRED_OUTCOME_MISMATCH|PUBLICATION_UNAVAILABLE|PUBLICATION_STALE|SOURCE_UNAVAILABLE|SOURCE_PENDING|SOURCE_APPLIED|COMMAND_COMPLETED|TERMINAL_NOT_COMPLETED|"

Public Function Create(ByVal header As Object, ByVal definition As Object, ByVal source As String, ByVal policyVersion As Long) As Object
    Dim model As Object, field As Variant, values As Collection, guide As Object
    Set model = CreateObject("Scripting.Dictionary")
    model.Add "SchemaVersion", 1&: model.Add "RecordKind", "Evaluation"
    model.Add "EvaluationId", modTrainingWire.NewId(): model.Add "Version", 1&
    For Each field In Array("ActionPathId", "SequenceId", "WarehouseId", "OriginWarehouseId")
        model.Add CStr(field), header(field)
    Next field
    model.Add "JournalVersion", header("Version"): model.Add "JournalRecordId", header("RecordId")
    model.Add "JournalSha256", header("ContentSha256"): model.Add "RecordedByUserId", header("CreatedByUserId")
    model.Add "EvaluatedByUserId", modAuth.GetCurrentUserId(): model.Add "EvaluatedAtUTC", modTrainingWire.UtcTimestamp()
    model.Add "CatalogVersion", modActivityCatalog.CATALOG_VERSION
    model.Add "PackageSetVersion", CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value)
    model.Add "BuildIdentity", CStr(ThisWorkbook.CustomDocumentProperties("invSysBuildIdentity").Value)
    model.Add "CapturePolicyVersion", header("PolicyVersion"): model.Add "EvaluationPolicyVersion", policyVersion
    model.Add "ExpectationSource", source: model.Add "ExpectedConclusion", definition
    model.Add "ExpectationSha256", modTrainingWire.Sha256(modTrainingJson.EncodeObject(definition))
    Set guide = CreateObject("Scripting.Dictionary"): model.Add "Guide", guide
    model.Add "PreviousEvaluationId", "": model.Add "ResultState", "Incomplete"
    For Each field In Split("ReasonCodes|Matches|MissingSteps|FailedSteps|UnavailableSteps|ExtraActivityIds|TerminalSources", "|")
        Set values = New Collection: model.Add CStr(field), values
    Next field
    Set Create = model
End Function

Public Sub Reason(ByVal model As Object, ByVal code As String)
    Dim value As Variant
    For Each value In model("ReasonCodes")
        If CStr(value) = code Then Exit Sub
    Next value
    model("ReasonCodes").Add code
End Sub

Public Function HasFields(ByVal model As Object, ByVal names As String) As Boolean
    Dim fields As Variant, field As Variant
    On Error GoTo Invalid
    If TypeName(model) <> "Dictionary" Then Exit Function
    fields = Split(names, "|")
    If model.Count <> UBound(fields) + 1 Then Exit Function
    For Each field In fields
        If Not model.Exists(CStr(field)) Then Exit Function
    Next field
    HasFields = True
Invalid:
End Function

Public Function IsHash(ByVal value As String) As Boolean
    IsHash = (Len(value) = 64 And Not value Like "*[!0-9a-f]*")
End Function

Public Function IsInteger(ByVal value As Variant) As Boolean
    IsInteger = (VarType(value) = vbInteger Or VarType(value) = vbLong)
End Function

Public Function Validate(ByVal model As Object, ByVal warehouseId As String) As Boolean
    Dim field As Variant, value As Variant, steps As Object, used As Object, actors As Object, entry As Object, id As String
    Dim order As Object, index As Long, lastStep As Long, lastOrdinal As Long, reasonSet As Object
    On Error GoTo Invalid
    If Not HasFields(model, FIELDS) Then Exit Function
    For Each field In model.Keys
        If InStr(NUMBERS, "|" & CStr(field) & "|") > 0 Then
            If Not IsInteger(model(field)) Or model(field) < 0 Then Exit Function
        ElseIf InStr(COLLECTIONS, "|" & CStr(field) & "|") > 0 Then
            If TypeName(model(field)) <> "Collection" Then Exit Function
            If field <> "TerminalSources" And model(field).Count > 256 Then Exit Function
        ElseIf field <> "ExpectedConclusion" And field <> "Guide" And field <> "Publication" Then
            If VarType(model(field)) <> vbString Then Exit Function
        End If
    Next field
    If model("SchemaVersion") <> 1 Or model("RecordKind") <> "Evaluation" Or model("Version") <> 1 Then Exit Function
    If model("WarehouseId") <> warehouseId Or model("OriginWarehouseId") <> warehouseId Then Exit Function
    If model("JournalVersion") < 1 Or model("JournalVersion") > 514 Then Exit Function
    If model("CatalogVersion") < 1 Or model("CatalogVersion") > modActivityCatalog.CATALOG_VERSION Then Exit Function
    For Each field In Array("EvaluationId", "ActionPathId", "SequenceId", "JournalRecordId")
        If Not modTrainingWire.ValidId(CStr(model(field))) Then Exit Function
    Next field
    If model("PreviousEvaluationId") <> "" Then
        If Not modTrainingWire.ValidId(CStr(model("PreviousEvaluationId"))) Or model("PreviousEvaluationId") = model("EvaluationId") Then Exit Function
    End If
    For Each field In Array("RecordedByUserId", "EvaluatedByUserId", "PackageSetVersion", "BuildIdentity")
        If model(field) = "" Then Exit Function
    Next field
    If Not modTrainingWire.ValidUtcTimestamp(CStr(model("EvaluatedAtUTC"))) Or Not IsHash(CStr(model("JournalSha256"))) Then Exit Function
    If Not modExpectationModel.Validate(model("ExpectedConclusion"), CLng(model("CatalogVersion"))) Then Exit Function
    If model("ExpectationSha256") <> modTrainingWire.Sha256(modTrainingJson.EncodeObject(model("ExpectedConclusion"))) Then Exit Function
    If InStr("|No expectation|Captured expectation|This evaluation|Guide expectation|", "|" & CStr(model("ExpectationSource")) & "|") = 0 Then Exit Function
    If TypeName(model("Guide")) <> "Dictionary" Then Exit Function
    If model("Guide").Count <> 0 Then
        If Not HasFields(model("Guide"), "ActionPathId|Version|ContentSha256") Then Exit Function
        If Not modTrainingWire.ValidId(CStr(model("Guide")("ActionPathId"))) Or Not IsHash(CStr(model("Guide")("ContentSha256"))) Then Exit Function
        If Not IsInteger(model("Guide")("Version")) Or model("Guide")("Version") < 1 Or model("ExpectationSource") <> "Guide expectation" Then Exit Function
    ElseIf model("ExpectationSource") = "Guide expectation" Then
        Exit Function
    End If
    If InStr("|Concluded|Awaiting|Failed|Cancelled|Incomplete|", "|" & CStr(model("ResultState")) & "|") = 0 Then Exit Function
    If model("ReasonCodes").Count = 0 Then Exit Function
    Set reasonSet = CreateObject("Scripting.Dictionary")
    For Each value In model("ReasonCodes")
        If VarType(value) <> vbString Then Exit Function
        If value = "" Or InStr(REASONS, "|" & CStr(value) & "|") = 0 Then Exit Function
        If reasonSet.Exists(value) Then Exit Function
        reasonSet.Add value, True
    Next value
    Set steps = CreateObject("Scripting.Dictionary"): Set used = CreateObject("Scripting.Dictionary"): Set actors = CreateObject("Scripting.Dictionary")
    Set order = CreateObject("Scripting.Dictionary")
    For Each entry In model("ExpectedConclusion")("Steps")
        index = index + 1: steps.Add entry("StepId"), entry: order.Add entry("StepId"), index
    Next entry
    For Each entry In model("Matches")
        If Not HasFields(entry, "StepId|ActivityId|Ordinal|ControlId|OutcomeCode") Then Exit Function
        For Each field In Array("StepId", "ActivityId", "ControlId", "OutcomeCode")
            If VarType(entry(field)) <> vbString Then Exit Function
        Next field
        id = entry("StepId")
        If Not steps.Exists(id) Or used.Exists(id) Then Exit Function
        If Not modTrainingWire.ValidId(CStr(entry("ActivityId"))) Or actors.Exists(entry("ActivityId")) Then Exit Function
        If Not IsInteger(entry("Ordinal")) Or entry("Ordinal") < 1 Or entry("Ordinal") > 256 Then Exit Function
        If entry("Ordinal") <= lastOrdinal Or order(id) <= lastStep Then Exit Function
        lastOrdinal = CLng(entry("Ordinal")): lastStep = CLng(order(id))
        If entry("ControlId") <> steps(id)("ControlId") Or entry("OutcomeCode") <> steps(id)("RequiredOutcome") Then Exit Function
        used.Add id, True: actors.Add entry("ActivityId"), True
    Next entry
    For Each field In Array("MissingSteps", "FailedSteps", "UnavailableSteps")
        For Each value In model(field)
            If VarType(value) <> vbString Then Exit Function
            If Not steps.Exists(value) Or used.Exists(value) Then Exit Function
            used.Add value, True
        Next value
    Next field
    If used.Count <> steps.Count Then Exit Function
    For Each value In model("ExtraActivityIds")
        If VarType(value) <> vbString Then Exit Function
        If Not modTrainingWire.ValidId(CStr(value)) Or actors.Exists(value) Then Exit Function
        actors.Add value, True
    Next value
    If Not modEvaluationSources.ValidSavedSources(model("TerminalSources"), warehouseId) Then Exit Function
    If Not modEvaluationSources.ValidPublication(model("Publication"), warehouseId) Then Exit Function
    If model("ResultState") = "Concluded" Or model("ResultState") = "Awaiting" Then
        If steps.Count = 0 Or model("Matches").Count <> steps.Count Or model("Publication")("Availability") <> "Loaded" Then Exit Function
        If model("ExpectedConclusion")("TerminalKind") = "SourceEventsApplied" Then
            If model("TerminalSources").Count = 0 Then Exit Function
            For Each entry In model("TerminalSources")
                If entry("OwnerStatus") = "Unavailable" Then Exit Function
                If model("ResultState") = "Concluded" And entry("OwnerStatus") <> "Applied" Then Exit Function
            Next entry
        ElseIf model("ResultState") <> "Concluded" Then
            Exit Function
        End If
    End If
    Validate = True
Invalid:
End Function
