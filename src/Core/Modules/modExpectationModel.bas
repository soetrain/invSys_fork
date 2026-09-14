Attribute VB_Name = "modExpectationModel"
Option Explicit
Option Private Module

Public Function NoneDefinition() As Object
    Dim model As Object, steps As New Collection
    Set model = CreateObject("Scripting.Dictionary")
    model.Add "SchemaVersion", 1&: model.Add "Steps", steps
    model.Add "TerminalStepId", "": model.Add "TerminalKind", "None"
    Set NoneDefinition = model
End Function

' D18 authored intent only. Validation never executes the registered control.
Public Function Validate(ByVal model As Object, ByVal catalogVersion As Long) As Boolean
    Dim field As Variant, step As Variant, seen As Object, definition As Object
    On Error GoTo Invalid
    If model.Count <> 4 Then Exit Function
    For Each field In model.Keys
        Select Case CStr(field)
            Case "SchemaVersion"
                If VarType(model(field)) <> vbInteger And VarType(model(field)) <> vbLong Then Exit Function
                If model(field) <> 1 Then Exit Function
            Case "Steps"
                If TypeName(model(field)) <> "Collection" Then Exit Function
            Case "TerminalStepId", "TerminalKind"
                If VarType(model(field)) <> vbString Then Exit Function
            Case Else: Exit Function
        End Select
    Next field
    If model("Steps").Count > 256 Then Exit Function
    If model("TerminalKind") = "None" Then
        Validate = (model("Steps").Count = 0 And model("TerminalStepId") = "")
        Exit Function
    End If
    If model("TerminalKind") <> "CommandCompleted" And model("TerminalKind") <> "SourceEventsApplied" Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbBinaryCompare
    For Each step In model("Steps")
        If step.Count <> 4 Then Exit Function
        For Each field In step.Keys
            Select Case CStr(field)
                Case "StepId", "ControlId", "RequiredOutcome"
                    If VarType(step(field)) <> vbString Then Exit Function
                    If step(field) = "" Then Exit Function
                Case "RetryAllowed"
                    If VarType(step(field)) <> vbBoolean Then Exit Function
                Case Else: Exit Function
            End Select
        Next field
        If seen.Exists(step("StepId")) Then Exit Function
        seen.Add step("StepId"), True
        Set definition = modActivityCatalog.Control(CStr(step("ControlId")), catalogVersion)
        If definition Is Nothing Then Exit Function
        Set definition = modActivityCatalog.Outcome(CStr(step("ControlId")), CStr(step("RequiredOutcome")))
        If definition Is Nothing Then Exit Function
    Next step
    Validate = (seen.Count > 0 And seen.Exists(model("TerminalStepId")))
Invalid:
End Function
