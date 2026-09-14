Attribute VB_Name = "modEvaluationPresentation"
Option Explicit
Option Private Module

Public Function Caption(ByVal result As Object) As String
    Select Case CStr(result("ResultState"))
        Case "Concluded"
            Caption = "Conclusion observed"
            If result("ExpectedConclusion")("TerminalKind") = "CommandCompleted" Then Caption = Caption & ". Command completed; Domain application not asserted"
        Case "Awaiting": Caption = "Awaiting published result"
        Case "Failed": Caption = "Failed"
        Case "Cancelled": Caption = "Cancelled"
        Case Else: Caption = "Incomplete evidence"
    End Select
End Function

Public Function Render(ByVal result As Object) As String
    Dim text As String, states As Object, entry As Object, step As Object, control As Object
    Dim id As Variant, index As Long
    Set states = CreateObject("Scripting.Dictionary")
    For Each entry In result("Matches")
        states.Add CStr(entry("StepId")), "Matched at action " & CStr(entry("Ordinal"))
    Next entry
    For Each id In result("MissingSteps"): states.Add CStr(id), "Missing expected step": Next id
    For Each id In result("FailedSteps"): states.Add CStr(id), "Outcome mismatch": Next id
    For Each id In result("UnavailableSteps"): states.Add CStr(id), "Evidence unavailable": Next id
    text = Caption(result) & vbCrLf & "Saved diagnostic result: " & CStr(result("EvaluationId")) & vbCrLf & _
        CStr(result("ExpectationSource")) & vbCrLf & "Evaluated at: " & CStr(result("EvaluatedAtUTC")) & vbCrLf & _
        "Publication: " & CStr(result("Publication")("Availability")) & vbCrLf
    If result("Publication")("LoadedAtUTC") <> "" Then text = text & "Publication loaded at: " & CStr(result("Publication")("LoadedAtUTC")) & vbCrLf
    For Each step In result("ExpectedConclusion")("Steps")
        index = index + 1
        Set control = modActivityCatalog.Control(CStr(step("ControlId")), CLng(result("CatalogVersion")))
        text = text & vbCrLf & CStr(index) & ". " & CStr(control("Caption")) & ": " & CStr(states(step("StepId"))) & _
            ". Expected " & CStr(step("RequiredOutcome")) & "." & vbCrLf
    Next step
    text = text & vbCrLf & "Additional observed actions: " & CStr(result("ExtraActivityIds").Count) & vbCrLf
    For Each entry In result("TerminalSources")
        text = text & "Source event: " & CStr(entry("EventId")) & " - " & CStr(entry("OwnerStatus")) & vbCrLf
    Next entry
    Render = text
End Function
