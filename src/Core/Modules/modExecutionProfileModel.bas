Attribute VB_Name = "modExecutionProfileModel"
Option Explicit
Option Private Module

Private Const FIELDS As String = "SchemaVersion|RecordKind|ProfileId|Version|RecordId|PreviousRecordId|PreviousSha256|WarehouseId|CreatedByUserId|CreatedAtUTC|Guide|CatalogVersion|PackageSetVersion|AdapterSetVersion|Steps|ExpectedConclusion"

Public Function Create(ByVal guide As Object, ByVal steps As Collection, ByVal previous As Object) As Object
    Dim model As Object, reference As Object, field As Variant
    Set model = CreateObject("Scripting.Dictionary"): Set reference = CreateObject("Scripting.Dictionary")
    model.Add "SchemaVersion", 1&: model.Add "RecordKind", "ExecutionProfile"
    If previous Is Nothing Then
        model.Add "ProfileId", modTrainingWire.NewId(): model.Add "Version", 1&
        model.Add "PreviousRecordId", "": model.Add "PreviousSha256", ""
    Else
        model.Add "ProfileId", previous("ProfileId"): model.Add "Version", CLng(previous("Version")) + 1&
        model.Add "PreviousRecordId", previous("RecordId"): model.Add "PreviousSha256", previous("ContentSha256")
    End If
    model.Add "RecordId", modTrainingWire.NewId(): model.Add "WarehouseId", guide("WarehouseId")
    model.Add "CreatedByUserId", modAuth.GetCurrentUserId(): model.Add "CreatedAtUTC", modTrainingWire.UtcTimestamp()
    For Each field In Array("ActionPathId", "Version", "RecordId", "ContentSha256")
        reference.Add CStr(field), guide(field)
    Next field
    model.Add "Guide", reference: model.Add "CatalogVersion", modActivityCatalog.CATALOG_VERSION
    model.Add "PackageSetVersion", CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value)
    model.Add "AdapterSetVersion", 1&: model.Add "Steps", steps
    model.Add "ExpectedConclusion", guide("ExpectedConclusion")
    Set Create = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(model))
End Function

Public Function NewSteps(ByVal guide As Object) As Collection
    Dim steps As New Collection, original As Object, step As Object
    For Each original In guide("Steps")
        If Not modExecutionInputs.Supported(CStr(original("ControlId"))) Then Exit Function
        Set step = CreateObject("Scripting.Dictionary")
        step.Add "StepId", original("StepId"): step.Add "ControlId", original("ControlId")
        step.Add "AdapterId", original("ControlId"): step.Add "AdapterVersion", 1&
        step.Add "Inputs", modExecutionInputs.Create(CStr(original("ControlId"))): steps.Add step
    Next original
    Set NewSteps = steps
End Function

Public Function Validate(ByVal target As WarehouseTarget, ByVal model As Object) As Boolean
    Dim field As Variant, guide As Object, reference As Object, step As Object, index As Long, expected As Object, position As Long
    On Error GoTo Invalid
    If Not modEvaluationModel.HasFields(model, FIELDS) Then Exit Function
    For Each field In model.Keys
        Select Case CStr(field)
            Case "SchemaVersion", "Version", "CatalogVersion", "AdapterSetVersion"
                If Not modEvaluationModel.IsInteger(model(field)) Or model(field) < 1 Then Exit Function
            Case "Guide", "ExpectedConclusion"
                If TypeName(model(field)) <> "Dictionary" Then Exit Function
            Case "Steps"
                If TypeName(model(field)) <> "Collection" Then Exit Function
            Case Else
                If VarType(model(field)) <> vbString Then Exit Function
        End Select
    Next field
    If model("SchemaVersion") <> 1 Or model("RecordKind") <> "ExecutionProfile" Or model("AdapterSetVersion") <> 1 Then Exit Function
    If model("WarehouseId") <> target.WarehouseId Or model("CatalogVersion") > modActivityCatalog.CATALOG_VERSION Then Exit Function
    If model("CreatedByUserId") = "" Or model("PackageSetVersion") = "" Then Exit Function
    If Not modTrainingWire.ValidUtcTimestamp(CStr(model("CreatedAtUTC"))) Then Exit Function
    If Not modTrainingWire.ValidId(CStr(model("ProfileId"))) Or Not modTrainingWire.ValidId(CStr(model("RecordId"))) Then Exit Function
    If model("ProfileId") = model("RecordId") Then Exit Function
    If model("Version") = 1 Then
        If model("PreviousRecordId") <> "" Or model("PreviousSha256") <> "" Then Exit Function
    Else
        If Not modTrainingWire.ValidId(CStr(model("PreviousRecordId"))) Or Not modEvaluationModel.IsHash(CStr(model("PreviousSha256"))) Then Exit Function
        If model("PreviousRecordId") = model("RecordId") Or model("PreviousRecordId") = model("ProfileId") Then Exit Function
    End If
    Set reference = model("Guide")
    If Not modEvaluationModel.HasFields(reference, "ActionPathId|Version|RecordId|ContentSha256") Then Exit Function
    If Not modEvaluationModel.IsInteger(reference("Version")) Then Exit Function
    For Each field In Array("ActionPathId", "RecordId", "ContentSha256")
        If VarType(reference(field)) <> vbString Then Exit Function
    Next field
    Set guide = modGuideStore.ReadChain(target, CStr(reference("ActionPathId")), CLng(reference("Version")))
    If guide Is Nothing Then Exit Function
    If reference("RecordId") <> guide("RecordId") Or reference("ContentSha256") <> guide("ContentSha256") Then Exit Function
    If model("ExpectedConclusion")("TerminalKind") = "None" Then Exit Function
    If modTrainingJson.EncodeObject(model("ExpectedConclusion")) <> modTrainingJson.EncodeObject(guide("ExpectedConclusion")) Then Exit Function
    If model("Steps").Count = 0 Or model("Steps").Count > 256 Or model("Steps").Count <> guide("Steps").Count Then Exit Function
    For Each step In model("Steps")
        index = index + 1
        If Not modEvaluationModel.HasFields(step, "StepId|ControlId|AdapterId|AdapterVersion|Inputs") Then Exit Function
        For Each field In Array("StepId", "ControlId", "AdapterId")
            If VarType(step(field)) <> vbString Then Exit Function
        Next field
        If Not modEvaluationModel.IsInteger(step("AdapterVersion")) Or step("AdapterVersion") <> 1 Then Exit Function
        If step("StepId") <> guide("Steps")(index)("StepId") Or step("ControlId") <> guide("Steps")(index)("ControlId") Then Exit Function
        If step("AdapterId") <> step("ControlId") Or TypeName(step("Inputs")) <> "Collection" Then Exit Function
        If Not modExecutionInputs.Validate(CStr(step("ControlId")), step("Inputs")) Then Exit Function
    Next step
    For Each expected In model("ExpectedConclusion")("Steps")
        Do
            position = position + 1
            If position > model("Steps").Count Then Exit Function
            If model("Steps")(position)("ControlId") = expected("ControlId") Then Exit Do
        Loop
    Next expected
    Validate = True
Invalid:
End Function
