Attribute VB_Name = "modPublishedGuideDraftSource"
Option Explicit
Option Private Module

' Exact published source and current authorization; no observed run is consulted.
Public Function Read(ByVal context As String, ByVal key As String, ByRef model As Object, _
                     ByRef records As Collection, ByRef visible As Object, ByRef policyHash As String, _
                     ByRef notice As String, ByRef policyVersion As Long, Optional ByVal saved As Object = Nothing) As Boolean
    Dim target As WarehouseTarget, policy As Object, parts As Variant, version As Long
    Dim item As Object, current As Object
    On Error GoTo Invalid
    Set model = Nothing: Set records = Nothing: Set visible = Nothing: policyHash = "": policyVersion = 0
    If Not modTrainingReadContext.Read(context, target, policy, notice, policyVersion) Then Exit Function
    notice = "Unavailable: editing a published guide requires ACTION_PATH_MAINT."
    If Not modAuth.CanPerform("ACTION_PATH_MAINT", modAuth.GetCurrentUserId(), target.WarehouseId, target.StationId) Then Exit Function
    notice = "Unavailable: the exact published guide or its predecessor changed. Reopen Edit guide."
    parts = Split(key, "|")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    version = CLng(parts(1))
    If version < 1 Or CStr(version) <> CStr(parts(1)) Then Exit Function
    Set model = modGuideStore.ReadChain(target, CStr(parts(0)), version)
    If model Is Nothing Then Exit Function
    If CStr(model("ContentSha256")) <> CStr(parts(2)) Then GoTo Invalid
    If Not saved Is Nothing Then
        If saved("ActionPathId") <> model("ActionPathId") Then GoTo Invalid
        Set current = modGuideStore.ReadChain(target, CStr(saved("ActionPathId")), CLng(saved("Version")))
        If current Is Nothing Then GoTo Invalid
        If current("RecordId") <> saved("RecordId") Or current("ContentSha256") <> saved("ContentSha256") Then GoTo Invalid
    End If
    Set visible = modEvaluationMatches.PolicyControls(policy, False)
    notice = "Hidden by policy: editing this guide is unavailable. No retained content was removed."
    For Each item In model("Steps")
        If Not modEvaluationMatches.Permitted(visible, CStr(item("ControlId"))) Then GoTo Invalid
    Next item
    For Each item In model("Observations")
        If Not modEvaluationMatches.Permitted(visible, CStr(item("ControlId"))) Then GoTo Invalid
    Next item
    For Each item In model("ExpectedConclusion")("Steps")
        If Not modEvaluationMatches.Permitted(visible, CStr(item("ControlId"))) Then GoTo Invalid
    Next item
    policy("SavedCatalogVersion") = CLng(policy("SavedCatalogVersion"))
    policyHash = modTrainingWire.Sha256(modTrainingJson.EncodeObject(policy))
    If Not modEvaluationModel.IsHash(policyHash) Or context <> modActivity.CaptureContext() Then GoTo Invalid
    Set records = model("Observations")
    notice = "": Read = True
    Exit Function
Invalid:
    Set model = Nothing: Set records = Nothing: Set visible = Nothing: policyHash = ""
    If notice = "" Then notice = "Unavailable: the published guide could not be validated. Reopen Edit guide."
End Function

Public Function SourceCaption(ByVal model As Object) As String
    Dim source As Object
    SourceCaption = "Editing guide: " & CStr(model("ActionPathId")) & "; version: " & CStr(model("Version")) & vbCrLf & _
                    "SHA-256: " & CStr(model("ContentSha256"))
    Set source = model("SourceRun")
    If source.Count > 0 Then
        SourceCaption = SourceCaption & vbCrLf & "Original source sequence: " & CStr(source("SequenceId")) & _
                        "; journal version: " & CStr(source("JournalVersion"))
    End If
    SourceCaption = SourceCaption & vbCrLf & "Authored instructions do not prove that a task ran."
End Function
