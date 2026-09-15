Attribute VB_Name = "modActionGuideDraft"
Option Explicit

' D18 declared primitive Core boundary. A draft is memory-only authored intent.
Private mContext As String
Private mPathId As String
Private mBinding As String
Private mDraftId As String
Private mPolicyHash As String
Private mHeader As Object
Private mSteps As Collection
Private mText As Object

Public Function CanCreate(ByVal context As String, ByVal pathId As String) As Boolean
    Dim header As Object, records As Collection, visible As Object, policyHash As String, notice As String
    CanCreate = modGuideDraftSource.ReadSelected(context, pathId, header, records, visible, policyHash, notice)
End Function

Public Function OpenDraft(ByVal context As String, ByVal pathId As String, ByRef draftId As String, ByRef notice As String) As Boolean
    Dim header As Object, records As Collection, visible As Object, policyHash As String, record As Object, step As Object
    On Error GoTo Failed
    draftId = ""
    If Not modGuideDraftSource.ReadSelected(context, pathId, header, records, visible, policyHash, notice) Then Exit Function
    Discard
    mContext = context: mPathId = pathId: mBinding = modPathExpectation.SelectionBinding(context, pathId)
    mDraftId = modTrainingWire.NewId(): mPolicyHash = policyHash: Set mHeader = header
    Set mSteps = New Collection: Set mText = CreateObject("Scripting.Dictionary")
    mText.Add "Name", "": mText.Add "Tags", "": mText.Add "Instructions", ""
    For Each record In records
        If record("OutcomeCode") = "REQUESTED" And modEvaluationMatches.Permitted(visible, CStr(record("ControlId"))) Then
            Set step = CreateObject("Scripting.Dictionary")
            step.Add "StepId", modTrainingWire.NewId(): step.Add "ActivityId", CStr(record("ActivityId"))
            step.Add "ControlId", CStr(record("ControlId")): step.Add "Caption", CStr(record("Caption"))
            step.Add "Instruction", "": mSteps.Add step
        End If
    Next record
    draftId = mDraftId: notice = "Draft only. Authored instructions do not prove that a task ran."
    OpenDraft = True
    Exit Function
Failed:
    Discard: draftId = "": notice = "Unavailable: the guide draft could not be created."
End Function

Private Function Guard(ByVal context As String, ByVal draftId As String, ByRef records As Collection, _
                       ByRef visible As Object, ByRef notice As String) As Boolean
    Dim header As Object, policyHash As String
    notice = "Unavailable: the guide draft or captured context changed. Reopen Create guide."
    If mHeader Is Nothing Or draftId = "" Or draftId <> mDraftId Or context <> mContext Then Exit Function
    If mBinding <> modPathExpectation.SelectionBinding(context, mPathId) Then GoTo Invalid
    If Not modGuideDraftSource.ReadSelected(context, mPathId, header, records, visible, policyHash, notice) Then GoTo Invalid
    If header("RecordId") <> mHeader("RecordId") Or header("ContentSha256") <> mHeader("ContentSha256") Then GoTo Invalid
    If policyHash <> mPolicyHash Then GoTo Invalid
    Guard = True
    Exit Function
Invalid:
    Discard
    notice = "Unavailable: the source selection, session, permission or policy changed. Reopen Create guide."
End Function

Public Function ReadDraft(ByVal context As String, ByVal draftId As String, ByRef rows As String, _
                          ByRef source As String, ByRef evidence As String, ByRef notice As String) As Boolean
    Dim records As Collection, visible As Object, step As Object
    On Error GoTo Failed
    rows = "": source = "": evidence = ""
    If Not Guard(context, draftId, records, visible, notice) Then Exit Function
    For Each step In mSteps
        rows = rows & CStr(step("StepId")) & vbTab & CStr(step("Caption")) & vbCrLf
    Next step
    source = modGuideDraftSource.SourceCaption(mHeader)
    evidence = modGuideDraftSource.Evidence(records, visible)
    notice = "Draft only. Changes affect authored instructions; source observations remain unchanged."
    ReadDraft = True
    Exit Function
Failed:
    rows = "": source = "": evidence = "": notice = "Unavailable: the guide draft could not be read."
    Discard
End Function

Public Function ReadText(ByVal context As String, ByVal draftId As String, ByVal field As String, _
                         ByVal stepId As String, ByRef value As String, ByRef notice As String) As Boolean
    Dim records As Collection, visible As Object, step As Object
    On Error GoTo Failed
    value = ""
    If Not Guard(context, draftId, records, visible, notice) Then Exit Function
    If field = "Instruction" Then
        Set step = FindStep(stepId)
        If step Is Nothing Then Exit Function
        value = CStr(step("Instruction"))
    Else
        If Not mText.Exists(field) Then Exit Function
        value = CStr(mText(field))
    End If
    ReadText = True
    Exit Function
Failed:
    value = "": notice = "Unavailable: authored text could not be read."
End Function

Public Function WriteText(ByVal context As String, ByVal draftId As String, ByVal field As String, _
                          ByVal stepId As String, ByVal value As String, ByRef notice As String) As Boolean
    Dim records As Collection, visible As Object, step As Object
    On Error GoTo Failed
    If Not Guard(context, draftId, records, visible, notice) Then Exit Function
    If field = "Instruction" Then
        Set step = FindStep(stepId)
        If step Is Nothing Then Exit Function
        step("Instruction") = value
    Else
        If Not mText.Exists(field) Then Exit Function
        mText(field) = value
    End If
    WriteText = True: notice = "Authored instruction staged in this draft only."
    Exit Function
Failed:
    notice = "Unavailable: authored text could not be staged."
End Function

Public Function EditStep(ByVal context As String, ByVal draftId As String, ByVal stepId As String, _
                         ByVal command As String, ByRef notice As String) As Boolean
    Dim records As Collection, visible As Object, step As Object, index As Long, destination As Long
    On Error GoTo Failed
    If Not Guard(context, draftId, records, visible, notice) Then Exit Function
    For index = 1 To mSteps.Count
        If mSteps(index)("StepId") = stepId Then Exit For
    Next index
    notice = "Select an authored step to change."
    If index > mSteps.Count Then Exit Function
    If command = "Remove" Then
        mSteps.Remove index
    ElseIf command = "Up" Or command = "Down" Then
        destination = index + IIf(command = "Up", -1, 1)
        notice = "The selected step is already at that end of the guide."
        If destination < 1 Or destination > mSteps.Count Then Exit Function
        Set step = mSteps(index): mSteps.Remove index
        If destination > mSteps.Count Then mSteps.Add step Else mSteps.Add step, Before:=destination
    Else
        Exit Function
    End If
    EditStep = True: notice = "Authored step order changed. Original observations are unchanged."
    Exit Function
Failed:
    notice = "Unavailable: the authored step could not be changed."
End Function

Private Function FindStep(ByVal stepId As String) As Object
    Dim step As Object
    For Each step In mSteps
        If step("StepId") = stepId Then Set FindStep = step: Exit Function
    Next step
End Function

Public Sub CloseDraft(ByVal context As String, ByVal draftId As String)
    If context = mContext And draftId = mDraftId Then Discard
End Sub

Private Sub Discard()
    mContext = "": mPathId = "": mBinding = "": mDraftId = "": mPolicyHash = ""
    Set mHeader = Nothing: Set mSteps = Nothing: Set mText = Nothing
End Sub
