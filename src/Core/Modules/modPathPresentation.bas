Attribute VB_Name = "modPathPresentation"
Option Explicit

' Read-only primitive Operations/Core boundary for an explicitly paired guide/run.
Public Function Capture(ByVal context As String, ByVal pathId As String, ByRef key As String, ByRef notice As String) As String
    Dim target As WarehouseTarget, header As Object, records As Collection, guide As Object, binding As String
    key = ""
    If Not ReadPair(context, pathId, target, header, records, guide, binding, notice) Then Exit Function
    key = GuideKey(guide): Capture = binding
End Function

Public Function Read(ByVal context As String, ByVal pathId As String, ByVal binding As String, ByVal key As String, _
                     ByVal evaluationId As String, ByRef instructions As String, ByRef diagnostic As String, _
                     ByRef provenance As String, ByRef method As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, header As Object, records As Collection, guide As Object, currentBinding As String
    Dim choice As String, effective As String, availability As String, request As String, policy As Object
    Dim version As Long, catalog As Long, ignored As String, guideNotice As String, recordNotice As String
    Dim currentRequest As String, currentVersion As Long, currentCatalog As Long, visible As Object
    On Error GoTo Invalid
    instructions = "": diagnostic = "": provenance = "": method = ""
    If Not modActionPathPreference.ReadPreference(context, choice, effective, availability, notice, version, request, catalog) Then GoTo Invalid
    If request = "" Then GoTo Invalid
    Set policy = modTrainingJson.DecodeObject(request)
    If policy Is Nothing Then GoTo Invalid
    policy.Add "SavedCatalogVersion", catalog
    method = choice
    If choice = "Use warehouse default" Then method = CStr(policy("DefaultView"))
    If method <> "How-To" And method <> "Diagnostic" And method <> "Compare both" Then GoTo Invalid
    If Not ReadPair(context, pathId, target, header, records, guide, currentBinding, recordNotice) Then GoTo Invalid
    If binding = "" Or currentBinding <> binding Or GuideKey(guide) <> key Then GoTo Invalid
    If Not modGuideLibraryRead.ReadGuide(context, key, instructions, ignored, provenance, guideNotice) Then GoTo Invalid
    Set visible = modEvaluationMatches.PolicyControls(policy, False)
    diagnostic = "Selected observed recording: " & pathId & "; journal version: " & CStr(header("Version")) & vbCrLf & _
        "Journal SHA-256: " & CStr(header("ContentSha256")) & vbCrLf & recordNotice & vbCrLf & vbCrLf & _
        modGuideDraftSource.Evidence(records, visible) & vbCrLf & SavedResult(context, pathId, target, header, guide, evaluationId)
    provenance = provenance & vbCrLf & "Selected observed recording: " & pathId & "; journal version: " & CStr(header("Version"))
    ' Fail closed if the policy changed while the existing readers validated evidence.
    If Not modActionPathPreference.ReadPreference(context, choice, effective, ignored, notice, currentVersion, currentRequest, currentCatalog) Then GoTo Invalid
    If request <> currentRequest Or version <> currentVersion Or catalog <> currentCatalog Then GoTo Invalid
    If context <> modActivity.CaptureContext() Or binding <> modPathExpectation.EvaluationBinding(context, pathId) Then GoTo Invalid
    notice = availability & " " & guideNotice
    Read = True
    Exit Function
Invalid:
    instructions = "": diagnostic = "": provenance = "": method = ""
    notice = "Unavailable: the paired guide, recording, session or current policy changed. Reopen View guide and run."
End Function

Private Function ReadPair(ByVal context As String, ByVal pathId As String, ByRef target As WarehouseTarget, _
                          ByRef header As Object, ByRef records As Collection, ByRef guide As Object, _
                          ByRef binding As String, ByRef notice As String) As Boolean
    Dim policy As Object, versions As Object
    On Error GoTo Invalid
    binding = "": Set header = Nothing: Set records = Nothing: Set guide = Nothing
    If Not modTrainingReadContext.Read(context, target, policy, notice) Then Exit Function
    binding = modPathExpectation.EvaluationBinding(context, pathId)
    If binding = "" Then GoTo Invalid
    If Not modPathExpectation.GuideReference(context, pathId, guide, notice) Then GoTo Invalid
    If guide Is Nothing Then GoTo Invalid
    If guide.Count = 0 Then GoTo Invalid
    Set versions = modRecordingReader.Versions(target)
    If versions Is Nothing Then GoTo Invalid
    If Not versions.Exists(pathId) Then GoTo Invalid
    Set records = modRecordingReader.ReadRun(target, pathId, CLng(versions(pathId)), header, notice)
    If records Is Nothing Then GoTo Invalid
    If Not modPathExpectation.MatchesHeader(context, pathId, header) Then GoTo Invalid
    If context <> modActivity.CaptureContext() Or binding <> modPathExpectation.EvaluationBinding(context, pathId) Then GoTo Invalid
    ReadPair = True
    Exit Function
Invalid:
    binding = ""
    notice = "Unavailable: explicitly use a permitted published guide for the selected recording, then reopen the view."
End Function

Private Function GuideKey(ByVal guide As Object) As String
    GuideKey = CStr(guide("ActionPathId")) & "|" & CStr(guide("Version")) & "|" & CStr(guide("ContentSha256"))
End Function

Private Function SavedResult(ByVal context As String, ByVal pathId As String, ByVal target As WarehouseTarget, _
                             ByVal header As Object, ByVal guide As Object, ByVal evaluationId As String) As String
    Dim result As Object, text As String, notice As String, entry As Object, id As Variant, group As Variant
    On Error GoTo Unavailable
    SavedResult = "Not evaluated for this guide and recording. Choose Evaluate in Action Paths to create a diagnostic result."
    If evaluationId = "" Then Exit Function
    Set result = modEvaluationStore.Read(target, evaluationId)
    If result Is Nothing Then GoTo Unavailable
    If result("ActionPathId") <> pathId Or result("JournalRecordId") <> header("RecordId") Or _
        result("JournalVersion") <> header("Version") Or result("JournalSha256") <> header("ContentSha256") Then Exit Function
    If result("Guide").Count = 0 Then Exit Function
    If GuideKey(result("Guide")) <> GuideKey(guide) Then Exit Function
    If Not modPathEvaluation.ReadSaved(context, pathId, evaluationId, text, notice) Then GoTo Unavailable
    For Each entry In result("Matches")
        text = text & "Matched expected step " & CStr(entry("StepId")) & ": observed action " & CStr(entry("ActivityId")) & vbCrLf
    Next entry
    For Each id In result("ExtraActivityIds"): text = text & "Additional observed action: " & CStr(id) & vbCrLf: Next id
    For Each group In Array("MissingSteps", "FailedSteps", "UnavailableSteps")
        For Each id In result(CStr(group)): text = text & CStr(group) & ": " & CStr(id) & vbCrLf: Next id
    Next group
    SavedResult = text
    Exit Function
Unavailable:
    SavedResult = "Incomplete evidence: the selected saved diagnostic result could not be verified for this pair."
End Function
