Attribute VB_Name = "modExecutionProfile"
Option Explicit

' Declared primitive Core boundary; forms receive projections, never authority objects.
Private mContext As String
Private mKey As String
Private mToken As String
Private mPolicyHash As String
Private mPolicyVersion As Long
Private mGuide As Object
Private mSaved As Object
Private mSteps As Collection

Public Function OpenDraft(ByVal context As String, ByVal key As String, ByRef token As String, ByRef notice As String) As Boolean
    Dim guide As Object, records As Collection, visible As Object, policyHash As String, version As Long, target As WarehouseTarget, saved As Object, copy As Object
    On Error GoTo Failed
    token = ""
    If Not modPublishedGuideDraftSource.Read(context, key, guide, records, visible, policyHash, notice, version) Then Exit Function
    notice = "Save an expected conclusion in the guide before configuring execution."
    If guide("ExpectedConclusion")("TerminalKind") = "None" Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If Not modExecutionProfileStore.Latest(target, guide, saved, notice) Then Exit Function
    Discard
    Set mGuide = guide: Set mSaved = saved
    If saved Is Nothing Then
        Set mSteps = modExecutionProfileModel.NewSteps(guide)
    Else
        Set copy = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(saved))
        Set mSteps = copy("Steps")
    End If
    notice = "Execution is not configured for one or more guide controls. No actions were skipped or run."
    If mSteps Is Nothing Then GoTo Failed
    mContext = context: mKey = key: mPolicyHash = policyHash: mPolicyVersion = version
    mToken = modTrainingWire.NewId(): token = mToken
    notice = "Review training inputs. Saving a profile does not run the guide."
    OpenDraft = True: Exit Function
Failed:
    Discard: token = ""
    If notice = "" Then notice = "Unavailable: execution inputs could not be opened."
End Function

Private Function Guard(ByVal context As String, ByVal token As String, ByRef notice As String) As Boolean
    Dim guide As Object, records As Collection, visible As Object, hash As String, version As Long
    notice = "Unavailable: the execution draft, session, permission or policy changed. Reopen Configure execution."
    If Not MatchesDraft(context, token) Then Exit Function
    If Not modPublishedGuideDraftSource.Read(context, mKey, guide, records, visible, hash, notice, version) Then GoTo Invalid
    If hash <> mPolicyHash Or version <> mPolicyVersion Then GoTo Invalid
    Guard = True: Exit Function
Invalid:
    Discard
    If notice = "" Then notice = "Unavailable: the execution profile policy changed. Reopen Configure execution."
End Function

Public Function ReadDraft(ByVal context As String, ByVal token As String, ByRef rows As String, ByRef guideCaption As String, _
                          ByRef profileCaption As String, ByRef notice As String) As Boolean
    Dim step As Object, index As Long
    On Error GoTo Failed
    rows = "": guideCaption = "": profileCaption = ""
    If Not Guard(context, token, notice) Then Exit Function
    guideCaption = "Guide " & CStr(mGuide("ActionPathId")) & "; version " & CStr(mGuide("Version")) & vbCrLf & "SHA-256: " & CStr(mGuide("ContentSha256"))
    For Each step In mSteps
        index = index + 1
        rows = rows & CStr(step("StepId")) & vbTab & CStr(index) & ". " & CStr(step("ControlId")) & vbCrLf
    Next step
    profileCaption = "Execution profile not saved."
    If Not mSaved Is Nothing Then profileCaption = "Profile " & CStr(mSaved("ProfileId")) & "; version " & CStr(mSaved("Version")) & vbCrLf & "SHA-256: " & CStr(mSaved("ContentSha256"))
    notice = "Training inputs are separate from observations. Save does not execute."
    ReadDraft = True: Exit Function
Failed:
    rows = "": guideCaption = "": profileCaption = "": notice = "Unavailable: execution inputs could not be read."
End Function

Public Function InputNames(ByVal context As String, ByVal token As String, ByVal stepId As String) As String
    Dim step As Object, notice As String
    On Error GoTo Failed
    If Not Guard(context, token, notice) Then Exit Function
    For Each step In mSteps
        If step("StepId") = stepId Then InputNames = modExecutionInputs.Names(CStr(step("ControlId"))): Exit Function
    Next step
Failed:
End Function

Private Function FindInput(ByVal stepId As String, ByVal name As String) As Object
    Dim step As Object, item As Object
    For Each step In mSteps
        If step("StepId") = stepId Then
            For Each item In step("Inputs")
                If item("Name") = name Then Set FindInput = item: Exit Function
            Next item
            Exit Function
        End If
    Next step
End Function

Public Function ReadInput(ByVal context As String, ByVal token As String, ByVal stepId As String, ByVal name As String, _
                          ByRef kind As String, ByRef value As String, ByRef notice As String) As Boolean
    Dim item As Object, binding As Object
    On Error GoTo Failed
    kind = "": value = ""
    If Not Guard(context, token, notice) Then Exit Function
    Set item = FindInput(stepId, name)
    If item Is Nothing Then Exit Function
    Set binding = item("Binding"): kind = CStr(binding("Kind"))
    If kind = "Prompt" Then value = CStr(binding("PromptId")) Else value = CStr(binding("Value"))
    ReadInput = True: notice = "": Exit Function
Failed:
    notice = "Unavailable: the selected execution input could not be read."
End Function

Public Function WriteInput(ByVal context As String, ByVal token As String, ByVal stepId As String, ByVal name As String, _
                           ByVal kind As String, ByVal value As String, ByRef notice As String) As Boolean
    Dim item As Object, binding As Object
    On Error GoTo Failed
    If Not Guard(context, token, notice) Then Exit Function
    Set item = FindInput(stepId, name)
    If item Is Nothing Then Exit Function
    Set binding = CreateObject("Scripting.Dictionary"): binding.Add "Kind", kind
    If kind = "Prompt" Then binding.Add "PromptId", value Else binding.Add "Value", value
    notice = "Input is invalid for this registered action. Previous draft values were preserved."
    If Not modExecutionInputs.ValidBinding(name, binding) Then Exit Function
    Set item("Binding") = binding
    WriteInput = True: notice = "Input staged. Save profile preserves the reviewed inputs.": Exit Function
Failed:
    notice = "Unavailable: the execution input could not be staged."
End Function

Public Function SaveDraft(ByVal context As String, ByVal token As String, ByRef notice As String) As Boolean
    Dim model As Object, target As WarehouseTarget, hash As String
    On Error GoTo Failed
    If Not Guard(context, token, notice) Then Exit Function
    Set model = modExecutionProfileModel.Create(mGuide, mSteps, mSaved)
    Set target = modNasConnection.GetCurrentTarget()
    If Not modExecutionProfileStore.Append(target, model, hash, notice) Then Exit Function
    model.Add "ContentSha256", hash: Set mSaved = model
    SaveDraft = True: notice = "Execution profile saved. No workflow was run.": Exit Function
Failed:
    notice = "Unavailable: the execution profile could not be saved. Existing versions were preserved."
End Function

Public Sub CloseDraft(ByVal context As String, ByVal token As String)
    If MatchesDraft(context, token) Then Discard
End Sub

Private Function MatchesDraft(ByVal context As String, ByVal token As String) As Boolean
    MatchesDraft = (token <> "" And token = mToken And context = mContext And Not mGuide Is Nothing)
End Function

Private Sub Discard()
    mContext = "": mKey = "": mToken = "": mPolicyHash = "": mPolicyVersion = 0
    Set mGuide = Nothing: Set mSaved = Nothing: Set mSteps = Nothing
End Sub
