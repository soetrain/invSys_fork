Attribute VB_Name = "modGuideExpectation"
Option Explicit
Option Private Module

' Exact authored intent only. Embedded observations never select an observed run.
Public Function Read(ByVal context As String, ByVal key As String, ByRef reference As Object, _
                     ByRef definition As Object, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, policy As Object, model As Object, visible As Object, step As Object
    Dim parts As Variant, version As Long
    On Error GoTo Invalid
    Set reference = Nothing: Set definition = Nothing
    If Not modTrainingReadContext.Read(context, target, policy, notice) Then Exit Function
    notice = "Incomplete evidence: the exact guide version or its predecessor is unavailable."
    parts = Split(key, "|")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    version = CLng(parts(1))
    If version < 1 Or CStr(version) <> CStr(parts(1)) Then Exit Function
    Set model = modGuideStore.ReadChain(target, CStr(parts(0)), version)
    If model Is Nothing Then Exit Function
    If CStr(model("ContentSha256")) <> CStr(parts(2)) Then Exit Function
    Set visible = modEvaluationMatches.PolicyControls(policy, False)
    notice = "Hidden by policy: the guide expectation is unavailable."
    For Each step In model("ExpectedConclusion")("Steps")
        If Not modEvaluationMatches.Permitted(visible, CStr(step("ControlId"))) Then Exit Function
    Next step
    If context <> modActivity.CaptureContext() Then GoTo Invalid
    Set reference = CreateObject("Scripting.Dictionary")
    reference.Add "ActionPathId", model("ActionPathId")
    reference.Add "Version", model("Version")
    reference.Add "ContentSha256", model("ContentSha256")
    Set definition = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(model("ExpectedConclusion")))
    Read = Not definition Is Nothing
    If Read Then notice = ""
    Exit Function
Invalid:
    Set reference = Nothing: Set definition = Nothing
    notice = "Unavailable: the exact guide expectation could not be validated."
End Function

Public Function Matches(ByVal context As String, ByVal reference As Object, ByVal definition As Object, _
                        ByRef notice As String) As Boolean
    Dim current As Object, expected As Object, key As String
    On Error GoTo Invalid
    key = CStr(reference("ActionPathId")) & "|" & CStr(reference("Version")) & "|" & CStr(reference("ContentSha256"))
    If Not Read(context, key, current, expected, notice) Then Exit Function
    notice = "Incomplete evidence: the saved definition does not match the exact guide version."
    Matches = (modTrainingJson.EncodeObject(definition) = modTrainingJson.EncodeObject(expected))
    If Matches Then notice = ""
    Exit Function
Invalid:
    notice = "Unavailable: the exact guide expectation could not be verified."
End Function
