Attribute VB_Name = "modGuideRevision"
Option Explicit
Option Private Module

' Preserve the published source bodies; only explicit authored edits become a revision.
Public Function Create(ByVal previous As Object, ByVal text As Object, ByVal draftSteps As Collection, _
                       ByVal expectation As Object, ByVal policyVersion As Long) As Object
    Dim model As Object, step As Object, authored As Object, field As Variant, tag As Variant
    Dim steps As New Collection, tags As New Collection
    Set model = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(previous))
    model.Remove "ContentSha256"
    model("Version") = CLng(previous("Version")) + 1&
    model("RecordId") = modTrainingWire.NewId()
    model("PreviousRecordId") = previous("RecordId"): model("PreviousSha256") = previous("ContentSha256")
    model("CreatedByUserId") = modAuth.GetCurrentUserId(): model("CreatedAtUTC") = modTrainingWire.UtcTimestamp()
    model("CatalogVersion") = modActivityCatalog.CATALOG_VERSION: model("PolicyVersion") = policyVersion
    model("PackageSetVersion") = CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value)
    model("BuildIdentity") = CStr(ThisWorkbook.CustomDocumentProperties("invSysBuildIdentity").Value)
    model("Name") = Trim$(CStr(text("Name"))): model("Instructions") = CStr(text("Instructions"))
    ' Unchanged tag text retains the exact ordered collection, including older values.
    If CStr(text("Tags")) <> modGuideModel.TagsText(previous) Then
        For Each tag In Split(CStr(text("Tags")), ",")
            If Trim$(CStr(tag)) <> "" Then tags.Add Trim$(CStr(tag))
        Next tag
        Set model("Tags") = tags
    End If
    For Each step In draftSteps
        Set authored = CreateObject("Scripting.Dictionary")
        authored.Add "StepId", step("StepId"): authored.Add "Method", "How-To"
        For Each field In Array("ControlId", "Caption", "Instruction")
            authored.Add CStr(field), step(field)
        Next field
        authored.Add "SourceActivityId", step("ActivityId"): steps.Add authored
    Next step
    Set model("Steps") = steps
    Set model("ExpectedConclusion") = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(expectation))
    Set Create = model
End Function
