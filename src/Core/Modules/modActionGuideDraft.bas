Attribute VB_Name = "modActionGuideDraft"
Option Explicit

' D18 primitive Core boundary. Edits are staged until explicit immutable save.
Private mContext As String
Private mPathId As String
Private mBinding As String
Private mDraftId As String
Private mPolicyHash As String
Private mHeader As Object
Private mSteps As Collection
Private mText As Object
Private mSaved As Object
Private mExpectation As Object
Private mPolicyVersion As Long
Private mPublishedKey As String
Private mCurationId As String

Public Function CanEditPublished(ByVal context As String, ByVal key As String, Optional ByRef notice As String = "") As Boolean
    Dim model As Object, records As Collection, visible As Object, hash As String, version As Long
    CanEditPublished = modPublishedGuideDraftSource.Read(context, key, model, records, visible, hash, notice, version)
End Function

Public Function OpenPublishedDraft(ByVal context As String, ByVal key As String, ByRef draftId As String, ByRef notice As String) As Boolean
    Dim model As Object, records As Collection, visible As Object, hash As String, version As Long
    Dim original As Object, step As Object, field As Variant
    On Error GoTo Failed
    draftId = ""
    If Not modPublishedGuideDraftSource.Read(context, key, model, records, visible, hash, notice, version) Then Exit Function
    Discard
    mContext = context: mPublishedKey = key: mDraftId = modTrainingWire.NewId()
    mPolicyHash = hash: mPolicyVersion = version: Set mHeader = model: Set mSaved = model
    Set mText = CreateObject("Scripting.Dictionary"): Set mSteps = New Collection
    mText.Add "Name", CStr(model("Name")): mText.Add "Tags", modGuideModel.TagsText(model)
    mText.Add "Instructions", CStr(model("Instructions"))
    For Each original In model("Steps")
        Set step = CreateObject("Scripting.Dictionary")
        For Each field In Array("StepId", "ControlId", "Caption", "Instruction")
            step.Add CStr(field), original(field)
        Next field
        step.Add "ActivityId", original("SourceActivityId"): mSteps.Add step
    Next original
    Set mExpectation = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(model("ExpectedConclusion")))
    draftId = mDraftId: notice = "Draft only. Save guide publishes the next immutable version."
    OpenPublishedDraft = True
    Exit Function
Failed:
    Discard: draftId = "": notice = "Unavailable: the published guide could not be opened for editing."
End Function

Public Function CanCreate(ByVal context As String, ByVal pathId As String) As Boolean
    Dim header As Object, records As Collection, visible As Object, policyHash As String, notice As String
    CanCreate = modGuideDraftSource.ReadSelected(context, pathId, header, records, visible, policyHash, notice)
End Function

Public Function OpenDraft(ByVal context As String, ByVal pathId As String, ByRef draftId As String, ByRef notice As String, _
                          Optional ByVal curationId As String = "") As Boolean
    Dim header As Object, records As Collection, visible As Object, policyHash As String, record As Object, step As Object, policyVersion As Long
    Dim binding As String
    On Error GoTo Failed
    draftId = ""
    If curationId <> "" Then
        If pathId <> "" Then Exit Function
        If Not modCuratedGuideSource.ReadSelected(context, curationId, header, records, visible, policyHash, policyVersion, binding, notice) Then Exit Function
    Else
        If Not modGuideDraftSource.ReadSelected(context, pathId, header, records, visible, policyHash, notice, policyVersion) Then Exit Function
        binding = modPathExpectation.SelectionBinding(context, pathId)
    End If
    Discard
    mContext = context: mPathId = pathId: mBinding = binding: mCurationId = curationId
    mDraftId = modTrainingWire.NewId(): mPolicyHash = policyHash: Set mHeader = header
    mPolicyVersion = policyVersion
    Set mSteps = New Collection: Set mText = CreateObject("Scripting.Dictionary")
    Set mExpectation = modExpectationModel.NoneDefinition()
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
    Dim header As Object, policyHash As String, policyVersion As Long, entry As String, binding As String
    entry = "Create guide": If mPublishedKey <> "" Then entry = "Edit guide"
    If mCurationId <> "" Then entry = "Choose tracked actions"
    notice = "Unavailable: the guide draft or captured context changed. Reopen " & entry & "."
    If mHeader Is Nothing Or draftId = "" Or draftId <> mDraftId Or context <> mContext Then Exit Function
    If mCurationId <> "" Then
        If Not modCuratedGuideSource.ReadSelected(context, mCurationId, header, records, visible, policyHash, policyVersion, binding, notice) Then GoTo Invalid
        If binding <> mBinding Or policyHash <> mPolicyHash Or policyVersion <> mPolicyVersion Then GoTo Invalid
        Guard = True: Exit Function
    End If
    If mPublishedKey <> "" Then
        If Not modPublishedGuideDraftSource.Read(context, mPublishedKey, header, records, visible, policyHash, notice, policyVersion, mSaved) Then GoTo Invalid
        If policyHash <> mPolicyHash Or policyVersion <> mPolicyVersion Then GoTo Invalid
        Guard = True: Exit Function
    End If
    If mBinding <> modPathExpectation.SelectionBinding(context, mPathId) Then GoTo Invalid
    If Not modGuideDraftSource.ReadSelected(context, mPathId, header, records, visible, policyHash, notice) Then GoTo Invalid
    If header("RecordId") <> mHeader("RecordId") Or header("ContentSha256") <> mHeader("ContentSha256") Then GoTo Invalid
    If policyHash <> mPolicyHash Then GoTo Invalid
    Guard = True
    Exit Function
Invalid:
    Discard
    notice = "Unavailable: the source selection, session, permission or policy changed. Reopen " & entry & "."
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
    If mCurationId <> "" Then
        source = modCuratedGuideSource.SourceCaption(mHeader)
    ElseIf mPublishedKey = "" Then
        source = modGuideDraftSource.SourceCaption(mHeader)
    Else
        source = modPublishedGuideDraftSource.SourceCaption(mHeader)
    End If
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

' Primitive/serialized projection; mutable expectation objects stay in Core.
Public Function ReadExpectation(ByVal context As String, ByVal draftId As String, ByRef sequenceId As String, _
                                ByRef definition As String, ByRef summary As String, ByRef notice As String) As Boolean
    Dim records As Collection, visible As Object, count As Long
    On Error GoTo Failed
    sequenceId = "": definition = "": summary = ""
    If Not Guard(context, draftId, records, visible, notice) Then Exit Function
    If mCurationId <> "" Then
        sequenceId = ""
    ElseIf mPublishedKey = "" Then
        sequenceId = CStr(mHeader("SequenceId"))
    ElseIf mHeader("SourceRun").Count > 0 Then
        sequenceId = CStr(mHeader("SourceRun")("SequenceId"))
    End If
    definition = modTrainingJson.EncodeObject(mExpectation)
    count = mExpectation("Steps").Count
    summary = "Guide expectation: None"
    If count > 0 Then summary = "Guide expectation: " & CStr(count) & " expected step(s)."
    ReadExpectation = True
    Exit Function
Failed:
    sequenceId = "": definition = "": summary = ""
    notice = "Unavailable: the guide expectation could not be read."
End Function

Public Function StageExpectation(ByVal context As String, ByVal draftId As String, ByVal serialized As String, ByRef notice As String) As Boolean
    Dim records As Collection, visible As Object, definition As Object
    On Error GoTo Failed
    If Not Guard(context, draftId, records, visible, notice) Then Exit Function
    notice = "Choose valid expected steps and a conclusion, or remove all steps and choose None."
    Set definition = modTrainingJson.DecodeObject(serialized)
    If Not modExpectationModel.Validate(definition, modActivityCatalog.CATALOG_VERSION) Then Exit Function
    Set mExpectation = definition
    notice = "Expected conclusion staged for this guide. Save guide publishes it with the version."
    StageExpectation = True
    Exit Function
Failed:
    notice = "Unavailable: the guide expectation could not be staged."
End Function

Public Function SaveDraft(ByVal context As String, ByVal draftId As String, ByRef notice As String) As Boolean
    Dim records As Collection, visible As Object, target As WarehouseTarget, model As Object, hash As String, step As Object
    On Error GoTo Failed
    If Not Guard(context, draftId, records, visible, notice) Then Exit Function
    notice = "A guide name is required."
    If Trim$(CStr(mText("Name"))) = "" Then Exit Function
    notice = "At least one authored step is required."
    If mSteps.Count = 0 Then Exit Function
    notice = "Guide save exceeds 1 MiB. No authored text or observations were truncated."
    ' The complete encoded record cannot be smaller than any one authored field.
    If Len(mText("Name")) >= 1048576 Or Len(mText("Tags")) >= 1048576 Or Len(mText("Instructions")) >= 1048576 Then Exit Function
    For Each step In mSteps
        If Len(step("Instruction")) >= 1048576 Then Exit Function
    Next step
    notice = "Unavailable: the guide version could not be saved. Your edits remain in this draft."
    If Not mSaved Is Nothing Then
        If mSaved("Version") = 2147483647 Then notice = "Guide version limit reached. Your edits remain in this draft.": Exit Function
    End If
    If mPublishedKey = "" Then
        Set model = modGuideModel.Create(mHeader, mText, mSteps, records, visible, mPolicyVersion, mSaved, mExpectation, mCurationId = "")
    Else
        Set model = modGuideRevision.Create(mSaved, mText, mSteps, mExpectation, mPolicyVersion)
    End If
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    If Not modGuideStore.Append(target, model, hash, notice) Then Exit Function
    model.Add "ContentSha256", hash: Set mSaved = model
    notice = "Published guide " & CStr(model("ActionPathId")) & ", version " & CStr(model("Version")) & "."
    SaveDraft = True
    Exit Function
Failed:
    notice = "Unavailable: the guide version could not be saved. Your edits remain in this draft."
End Function

Private Sub Discard()
    modExpectationDraft.CloseGuide mContext, mDraftId
    mContext = "": mPathId = "": mBinding = "": mDraftId = "": mPolicyHash = "": mPublishedKey = "": mCurationId = ""
    Set mHeader = Nothing: Set mSteps = Nothing: Set mText = Nothing: Set mSaved = Nothing: mPolicyVersion = 0
    Set mExpectation = Nothing
End Sub
