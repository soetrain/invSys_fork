Attribute VB_Name = "modExpectationDraft"
Option Explicit
Option Private Module

Private mContext As String
Private mSequence As String
Private mId As String
Private mModel As Object

Public Function OpenRecording(ByVal context As String, ByRef projection As String, ByRef notice As String) As Boolean
    Dim definition As Object, sequenceId As String
    On Error GoTo Unavailable
    projection = ""
    If Not modRecordingSession.ReadExpectation(context, sequenceId, definition, notice) Then Exit Function
    Discard ""
    mContext = context: mSequence = sequenceId: mId = modTrainingWire.NewId()
    Set mModel = definition
    projection = Project()
    notice = "Choose expected steps and the conclusion to check. This does not perform the work."
    OpenRecording = True
    Exit Function
Unavailable:
    Discard "": projection = "": notice = "Expected conclusion is unavailable. Reopen the editor."
End Function

Private Function Guard(ByVal context As String, ByVal draftId As String, ByRef notice As String) As Boolean
    Dim sequenceId As String, definition As Object
    notice = "This expectation draft is no longer available. Reopen Expected conclusion."
    If mModel Is Nothing Or draftId = "" Or draftId <> mId Or context <> mContext Then Exit Function
    If Not modRecordingSession.ReadExpectation(context, sequenceId, definition, notice) Then Discard draftId: Exit Function
    If sequenceId <> mSequence Then Discard draftId: Exit Function
    Guard = True
End Function

Public Function Choices(ByVal context As String, ByVal draftId As String, ByVal controlId As String) As String
    Dim id As Variant, definition As Object, notice As String, caption As String
    On Error GoTo Unavailable
    If Not Guard(context, draftId, notice) Then Exit Function
    If controlId = "" Then
        For Each id In modActivityCatalog.ControlIds()
            Set definition = modActivityCatalog.Control(CStr(id))
            Choices = Choices & CStr(id) & vbTab & SafeCaption(CStr(definition("Surface")) & " > " & CStr(definition("Caption"))) & vbCrLf
        Next id
    Else
        For Each id In Array("REQUESTED", "COMPLETED", "UNCHANGED", "DENIED", "REJECTED", "FAILED", _
                             "CONFIRMED", "PENDING", "STAGED", "REFRESHED", "STALE", "CLEARED", "EMPTY", _
                             "OPENED", "REUSED", "CLOSED", "SELECTED")
            Set definition = modActivityCatalog.Outcome(controlId, CStr(id))
            If Not definition Is Nothing Then
                caption = StrConv(LCase$(CStr(id)), vbProperCase)
                Choices = Choices & CStr(id) & vbTab & caption & vbCrLf
            End If
        Next id
    End If
    Exit Function
Unavailable:
    Choices = ""
End Function

Public Function Edit(ByVal context As String, ByVal draftId As String, ByVal command As String, _
                     ByVal stepId As String, ByVal controlId As String, ByVal outcome As String, _
                     ByVal retry As Boolean, ByRef projection As String, ByRef notice As String) As Boolean
    Dim steps As Collection, step As Object, definition As Object, index As Long, destination As Long
    On Error GoTo Unavailable
    projection = ""
    If Not Guard(context, draftId, notice) Then Exit Function
    Set steps = mModel("Steps")
    Select Case command
        Case "Read"
        Case "Add"
            notice = "Choose a registered control and outcome. No more than 256 steps are allowed."
            If steps.Count >= 256 Then Exit Function
            Set definition = modActivityCatalog.Control(controlId)
            If definition Is Nothing Then Exit Function
            Set definition = modActivityCatalog.Outcome(controlId, outcome)
            If definition Is Nothing Then Exit Function
            Set step = CreateObject("Scripting.Dictionary")
            step.Add "StepId", modTrainingWire.NewId(): step.Add "ControlId", controlId
            step.Add "RequiredOutcome", outcome: step.Add "RetryAllowed", retry
            steps.Add step
        Case "Remove", "Up", "Down"
            For index = 1 To steps.Count
                If CStr(steps(index)("StepId")) = stepId Then Exit For
            Next index
            notice = "Select a step to change."
            If index > steps.Count Then Exit Function
            If command = "Remove" Then
                steps.Remove index
                If mModel("TerminalStepId") = stepId Then mModel("TerminalStepId") = "": mModel("TerminalKind") = "None"
            Else
                destination = index + IIf(command = "Up", -1, 1)
                If destination < 1 Or destination > steps.Count Then Exit Function
                Set step = steps(index): steps.Remove index
                If destination > steps.Count Then steps.Add step Else steps.Add step, Before:=destination
            End If
        Case Else: notice = "Expectation edit is unavailable.": Exit Function
    End Select
    projection = Project(): notice = "Draft only. Use for this recording stages the selected conclusion."
    Edit = True
    Exit Function
Unavailable:
    projection = "": notice = "The expectation draft could not be edited. Reopen the editor."
End Function

Public Function UseRecording(ByVal context As String, ByVal draftId As String, ByVal terminalStepId As String, _
                              ByVal terminalKind As String, ByRef notice As String) As Boolean
    Dim definition As Object
    On Error GoTo Unavailable
    If Not Guard(context, draftId, notice) Then Exit Function
    Set definition = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(mModel))
    definition("TerminalStepId") = terminalStepId: definition("TerminalKind") = terminalKind
    UseRecording = modRecordingSession.StageExpectation(context, mSequence, definition, notice)
    If UseRecording Then Discard draftId
    Exit Function
Unavailable:
    notice = "Expected conclusion could not be staged. The recording was not changed."
End Function

Public Sub CloseContext(ByVal context As String, ByVal draftId As String)
    If context = mContext Then Discard draftId
End Sub

Public Sub Discard(ByVal draftId As String)
    If draftId <> "" And draftId <> mId Then Exit Sub
    Set mModel = Nothing
    mContext = "": mSequence = "": mId = ""
End Sub

Private Function Project() As String
    Dim step As Variant, definition As Object
    Project = "EXPECTATION1" & vbTab & mId & vbTab & mSequence & vbTab & _
        CStr(mModel("TerminalStepId")) & vbTab & CStr(mModel("TerminalKind"))
    For Each step In mModel("Steps")
        Set definition = modActivityCatalog.Control(CStr(step("ControlId")))
        Project = Project & vbCrLf & CStr(step("StepId")) & vbTab & CStr(step("ControlId")) & vbTab & _
            CStr(step("RequiredOutcome")) & vbTab & CStr(step("RetryAllowed")) & vbTab & _
            SafeCaption(CStr(definition("Surface")) & " > " & CStr(definition("Caption")))
    Next step
End Function

Private Function SafeCaption(ByVal value As String) As String
    SafeCaption = Replace$(Replace$(Replace$(value, vbTab, " "), vbCr, " "), vbLf, " ")
End Function
