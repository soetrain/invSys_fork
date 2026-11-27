Attribute VB_Name = "modExecutionRun"
Option Explicit

' Declared primitive boundary. Operations dispatches; Core owns binding and evidence.
Private mContext As String
Private mKey As String
Private mToken As String
Private mPolicyHash As String
Private mBuildHash As String
Private mHash As String
Private mEntity As String
Private mTarget As WarehouseTarget
Private mGuide As Object
Private mProfile As Object
Private mBuilds As Collection
Private mRun As Object
Private mPending As Long
Private mSeedClaimed As Boolean
Private mNeedsEntity As Boolean
Private mBefore As Long
Private mStop As Boolean
Private mWriteFailed As Boolean
Private mStarting As Boolean
Private mOperatorReturning As Boolean

Public Function OpenSetup(ByVal context As String, ByVal key As String, ByRef token As String, ByRef snapshot As String, _
                          ByRef guideCaption As String, ByRef profileCaption As String, ByRef targetCaption As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, guide As Object, profile As Object, builds As Collection, policyHash As String, step As Object
    On Error GoTo Failed
    token = "": snapshot = "": guideCaption = "": profileCaption = "": targetCaption = ""
    If mStarting Or mPending > 0 Then notice = "A run action is still returning. Stop it before opening another setup.": Exit Function
    If Not mRun Is Nothing Then
        If mRun("State") = "Running" Then notice = "Stop the current run before opening another setup.": Exit Function
    End If
    If Not modExecutionBinding.Read(context, key, target, guide, profile, policyHash, builds, snapshot, notice) Then Exit Function
    Set mTarget = target: Set mGuide = guide: Set mProfile = profile: Set mBuilds = builds: Set mRun = Nothing
    mContext = context: mKey = key: mPolicyHash = policyHash: mBuildHash = modExecutionBinding.BuildsHash(builds)
    mToken = modTrainingWire.NewId(): token = mToken: mPending = 0: mBefore = 0: mStop = False: mWriteFailed = False: mHash = "": mEntity = ""
    guideCaption = "Guide " & CStr(guide("ActionPathId")) & "; version " & CStr(guide("Version")) & vbCrLf & CStr(guide("ContentSha256"))
    profileCaption = "Profile " & CStr(profile("ProfileId")) & "; version " & CStr(profile("Version")) & vbCrLf & CStr(profile("ContentSha256"))
    targetCaption = "Training warehouse: " & target.WarehouseId & "; station: " & target.StationId & "; user: " & modAuth.GetCurrentUserId()
    notice = "Review the captured Training target and choose an exact local source entity. Start Run executes ordinary Receiving actions."
    mNeedsEntity = False
    For Each step In profile("Steps")
        If Left$(step("ControlId"), 10) = "RECEIVING_" Then mNeedsEntity = True
    Next step
    If Not mNeedsEntity Then
        snapshot = "": notice = "Review the captured Training target. Start Run applies the reviewed built-in Seed kit."
    End If
    If modOperatorStepEvidence.IncludesCompletion(profile("Steps")) Then notice = "Review the Training target. Production Complete Run requires an operator to bind the prepared selection and click Complete Run."
    OpenSetup = True
Failed:
End Function

Private Function Matches(ByVal context As String, ByVal token As String) As Boolean
    Matches = (token <> "" And token = mToken And context = mContext And Not mProfile Is Nothing)
End Function

Private Function Guard(ByVal context As String, ByVal token As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, guide As Object, profile As Object, builds As Collection, hash As String, snapshot As String
    notice = "Run blocked: the captured context, guide, inputs, permission, recording policy or package changed. Reopen setup."
    If Not Matches(context, token) Or mWriteFailed Then Exit Function
    If Not modExecutionBinding.Read(context, mKey, target, guide, profile, hash, builds, snapshot, notice) Then Exit Function
    If profile("ContentSha256") <> mProfile("ContentSha256") Or hash <> mPolicyHash Then Exit Function
    If modExecutionBinding.BuildsHash(builds) <> mBuildHash Then Exit Function
    Guard = True
End Function

Public Function Start(ByVal context As String, ByVal token As String, ByVal entity As String, ByVal mode As String, ByRef notice As String) As Boolean
    On Error GoTo Done
    If mStarting Then notice = "Start Run is already in progress.": Exit Function
    mStarting = True
    Start = StartAttempt(context, token, entity, mode, notice)
Done:
    mStarting = False
End Function

Private Function StartAttempt(ByVal context As String, ByVal token As String, ByVal entity As String, ByVal mode As String, ByRef notice As String) As Boolean
    Dim active As Boolean, canStart As Boolean, header As Object, observations As Collection, recording As Object
    On Error GoTo Failed
    If mPending > 0 Or Not mRun Is Nothing Then notice = "This setup already has an attempt. Close it and explicitly open a new setup.": Exit Function
    If mode <> "RunAll" And mode <> "StepThrough" Then notice = "Choose Run all or Step through.": Exit Function
    If Not Guard(context, token, notice) Then Exit Function
    If mNeedsEntity Then
        If Not modExecutionTarget.ContainsEntity(context, entity, notice) Then Exit Function
    ElseIf entity <> "" Then
        notice = "Seed has no source entity input.": Exit Function
    End If
    notice = modActionRecording.Status(context, active, canStart)
    If active Or Not canStart Then notice = "Stop the current recording and enable required capture before starting a run.": Exit Function
    If mStop Then notice = "Run was stopped before its first action.": Exit Function
    Set mRun = modExecutionRunModel.Create(mTarget, context, mGuide, mProfile, mBuilds, mode)
    If Not modExecutionRunStore.Append(mTarget, mRun, mHash, notice) Then mWriteFailed = True: Exit Function
    mEntity = entity
    If Not modActionRecording.Control(context, "Start", notice) Then EndRun "Blocked", "CAPTURE_UNAVAILABLE", notice: Exit Function
    If Not modRecordingSession.ExecutionEvidence(context, header, observations) Then GoTo Failed
    Set recording = mRun("Recording")
    recording.Add "ActionPathId", header("ActionPathId"): recording.Add "SequenceId", header("SequenceId")
    mRun("State") = "Running"
    StartAttempt = SaveRevision(notice)
    If StartAttempt Then notice = "Run started. Select Verify run after execution to check business results."
    Exit Function
Failed:
    If Not mRun Is Nothing Then EndRun "Unknown", "CAPTURE_UNAVAILABLE", notice
End Function

Public Function BeginStep(ByVal context As String, ByVal token As String, ByRef controlId As String, ByRef inputs As String, ByRef notice As String, Optional ByRef stepId As String = "") As Boolean
    Dim header As Object, observations As Collection, step As Object, item As Object, binding As Object
    On Error GoTo Failed
    controlId = "": inputs = "": stepId = ""
    If mStarting Or Not Matches(context, token) Or mPending > 0 Or mRun Is Nothing Then Exit Function
    If mRun("State") <> "Running" Or mWriteFailed Then Exit Function
    If mStop Then EndRun "Stopped", "STOP_REQUESTED", notice: Exit Function
    If Not Guard(context, token, notice) Then EndRun "Blocked", "CONTEXT_CHANGED", notice: Exit Function
    If Not modRecordingSession.ActiveFor(context) Then EndRun "Blocked", "CAPTURE_UNAVAILABLE", notice: Exit Function
    If mRun("Steps").Count >= mProfile("Steps").Count Then EndRun "Completed", "", notice: Exit Function
    If Not modRecordingSession.ExecutionEvidence(context, header, observations) Then GoTo Failed
    If header("SequenceId") <> mRun("Recording")("SequenceId") Then GoTo Failed
    Set step = mProfile("Steps")(mRun("Steps").Count + 1)
    controlId = CStr(step("ControlId"))
    For Each item In step("Inputs")
        Set binding = item("Binding")
        If inputs <> "" Then inputs = inputs & vbTab
        If binding("Kind") = "Prompt" Then inputs = inputs & mEntity Else inputs = inputs & CStr(binding("Value"))
    Next item
    mBefore = observations.Count: mPending = mRun("Steps").Count + 1
    stepId = CStr(step("StepId")): mSeedClaimed = False
    notice = "Running: " & CStr(mGuide("Steps")(mPending)("Caption"))
    If controlId = modOperatorStepEvidence.CONTROL_ID Then
        mRun("OperatorClaims").Add modOperatorStepEvidence.NewClaim(stepId)
        mOperatorReturning = False
        If Not SaveRevision(notice) Then mPending = 0: Exit Function
        notice = "Operator action required. Select the prepared Process or worksheet output in Production, then bind that selection here."
    End If
    BeginStep = True: Exit Function
Failed:
    If Not mRun Is Nothing Then EndRun "Unknown", "CAPTURE_UNAVAILABLE", notice
End Function

Private Function OperatorPending() As Boolean
    If mPending <= 0 Or mRun Is Nothing Then Exit Function
    OperatorPending = (mProfile("Steps")(mPending)("ControlId") = modOperatorStepEvidence.CONTROL_ID)
End Function

Public Function HasOperatorStep(ByVal context As String, ByVal token As String) As Boolean
    If Matches(context, token) Then HasOperatorStep = modOperatorStepEvidence.IncludesCompletion(mProfile("Steps"))
End Function

Private Function CurrentClaim(ByVal context As String, ByVal token As String, ByVal stepId As String, ByRef notice As String) As Object
    Dim saved As Object, claim As Object, header As Object, observations As Collection
    notice = "This operator step is no longer eligible in the captured Training session."
    If Not Matches(context, token) Or Not OperatorPending() Or mStop Then Exit Function
    If mRun("State") <> "Running" Or mProfile("Steps")(mPending)("StepId") <> stepId Then Exit Function
    If Not Guard(context, token, notice) Or Not modRecordingSession.ActiveFor(context) Then Exit Function
    Set saved = modExecutionRunStore.ReadVersion(mTarget, CStr(mRun("RunId")), CLng(mRun("Revision")))
    If saved Is Nothing Then Exit Function
    If saved("ContentSha256") <> mHash Then Exit Function
    If Not modRecordingSession.ExecutionEvidence(context, header, observations) Then Exit Function
    If header("SequenceId") <> saved("Recording")("SequenceId") Then Exit Function
    Set claim = mRun("OperatorClaims")(mRun("OperatorClaims").Count)
    If claim("StepId") = stepId And Not mStop And context = modActivity.CaptureContext() Then Set CurrentClaim = claim
End Function

Public Function SelectionTicket(ByVal context As String, ByVal token As String, ByVal stepId As String, ByRef notice As String) As String
    Dim claim As Object
    Set claim = CurrentClaim(context, token, stepId, notice)
    If claim Is Nothing Then Exit Function
    If claim("Kind") = "" Then SelectionTicket = claim("BindingId")
End Function

Public Function BindSelection(ByVal context As String, ByVal token As String, ByVal stepId As String, ByVal ticket As String, _
                              ByVal kind As String, ByVal digest As String, ByRef notice As String) As Boolean
    Dim claim As Object, header As Object, observations As Collection
    Set claim = CurrentClaim(context, token, stepId, notice)
    If claim Is Nothing Then Exit Function
    If claim("BindingId") <> ticket Or claim("Kind") <> "" Then Exit Function
    If kind <> "WorksheetOutput" And kind <> "ReusableProcess" Then Exit Function
    If Not modEvaluationModel.IsHash(digest) Then Exit Function
    If Not modRecordingSession.ExecutionEvidence(context, header, observations) Then Exit Function
    If header("SequenceId") <> mRun("Recording")("SequenceId") Then Exit Function
    claim("AfterObservation") = CLng(observations.Count)
    claim("Kind") = kind: claim("SelectionSha256") = digest
    BindSelection = SaveRevision(notice)
    If BindSelection Then notice = "Selection bound. Click its ordinary Complete Run button in Production, then check the saved evidence here."
End Function

Public Function ClaimCompletion(ByVal context As String, ByVal token As String, ByVal stepId As String, ByVal ticket As String, _
                                ByVal digest As String, ByVal activityId As String, ByVal atWrite As Boolean) As Boolean
    Dim claim As Object, terminal As Object, notice As String
    Set claim = CurrentClaim(context, token, stepId, notice)
    If claim Is Nothing Then Exit Function
    If digest = "" Or claim("SelectionSha256") <> digest Or claim("BindingId") <> ticket Then Exit Function
    If Not modOperatorStepEvidence.ReadPair(context, CStr(mRun("Recording")("SequenceId")), activityId, terminal, CLng(claim("AfterObservation"))) Then Exit Function
    If Not terminal Is Nothing Then Exit Function
    If atWrite Then
        If claim("ActivityId") <> activityId Or claim("OwnerValidated") Then Exit Function
        claim("OwnerValidated") = True
    Else
        If claim("ActivityId") <> "" Or mOperatorReturning Then Exit Function
        claim("ActivityId") = activityId
    End If
    ClaimCompletion = SaveRevision(notice)
    If ClaimCompletion Then mOperatorReturning = True
End Function

Public Sub OperatorReturned(ByVal token As String, ByVal activityId As String)
    Dim claim As Object, notice As String
    If token <> mToken Or Not OperatorPending() Or Not mOperatorReturning Then Exit Sub
    Set claim = mRun("OperatorClaims")(mRun("OperatorClaims").Count)
    If claim("ActivityId") <> activityId Then Exit Sub
    mOperatorReturning = False
    If mStop Then mPending = 0: EndRun "Stopped", "STOP_REQUESTED", notice
End Sub

Public Function CheckCompletion(ByVal context As String, ByVal token As String, ByVal stepId As String, ByRef notice As String, _
                                Optional ByVal selectionStillBound As Boolean = True) As Boolean
    Dim claim As Object, terminal As Object, row As Object, outputs As New Collection
    On Error GoTo Failed
    If Not Matches(context, token) Or Not OperatorPending() Then Exit Function
    Set claim = CurrentClaim(context, token, stepId, notice)
    If claim Is Nothing Then
        If Not mOperatorReturning Then mPending = 0: EndRun "Blocked", "CONTEXT_CHANGED", notice
        Exit Function
    End If
    notice = "Waiting for the bound selection's ordinary Complete Run action."
    If claim("ActivityId") = "" And claim("Kind") <> "" And Not selectionStillBound Then
        mPending = 0: EndRun "Blocked", "CONTEXT_CHANGED", notice
        notice = "The bound Production workbook or prepared selection changed. This run cannot credit another completion."
        Exit Function
    End If
    If Not modOperatorStepEvidence.ReadPair(context, CStr(mRun("Recording")("SequenceId")), CStr(claim("ActivityId")), terminal, CLng(claim("AfterObservation"))) Then Exit Function
    If terminal Is Nothing Then notice = "The claimed completion has no saved owner result yet. Check again after it returns.": Exit Function
    mOperatorReturning = False
    Set row = CreateObject("Scripting.Dictionary")
    row.Add "StepId", stepId: row.Add "ControlId", modOperatorStepEvidence.CONTROL_ID
    row.Add "State", "Failed": row.Add "ReasonCode", terminal("OutcomeCode"): row.Add "ActivityId", claim("ActivityId")
    row.Add "SourceEventRefs", terminal("SourceEventRefs"): row.Add "Outputs", outputs
    If terminal("OutcomeCode") = "CONFIRMED" And claim("OwnerValidated") Then
        If Not modOperatorStepEvidence.Applied(context, mTarget.WarehouseId, terminal) Then
            notice = "Completion recorded. Refresh Events and check again to verify every Inventory event. Later steps are waiting."
            Exit Function
        End If
        row("State") = "OperatorCompleted"
    ElseIf terminal("OutcomeCode") = "CONFIRMED" Then
        row("State") = "Blocked"
    ElseIf terminal("OutcomeCode") = "PENDING" Or (terminal("OutcomeCode") = "FAILED" And terminal("SourceEventRefs").Count > 0) Then
        row("State") = "Unknown"
    End If
    mRun("Steps").Add row: mPending = 0
    If Not SaveRevision(notice) Then Exit Function
    If row("State") <> "OperatorCompleted" Then
        EndRun CStr(row("State")), "OWNER_NOT_COMPLETED", notice
        notice = "Production reported " & CStr(terminal("OutcomeCode")) & ". The bound step is unfulfilled. Inspect the saved owner evidence before starting another run."
        Exit Function
    End If
    CheckCompletion = True
    If mRun("Steps").Count = mProfile("Steps").Count Then
        EndRun "Completed", "", notice
    Else
        notice = "Operator completion verified. Select Next Step to continue."
    End If
    Exit Function
Failed:
    mPending = 0: EndRun "Unknown", "CAPTURE_UNAVAILABLE", notice
End Function

Public Function PendingSeed(ByVal token As String, ByVal stepId As String, ByVal kit As String, ByVal claim As Boolean) As Boolean
    Dim step As Object, context As String, notice As String, saved As Object, header As Object, observations As Collection
    On Error GoTo Failed
    context = modActivity.CaptureContext()
    If mPending = 0 Or mSeedClaimed Or mStop Or mRun Is Nothing Then Exit Function
    If mRun("State") <> "Running" Or Not Matches(context, token) Then Exit Function
    Set step = mProfile("Steps")(mPending)
    If step("ControlId") <> "ADMIN_DEMO_SEED" Or step("StepId") <> stepId Then Exit Function
    If step("Inputs")(1)("Binding")("Value") <> kit Or kit <> "R1_WORKFLOW_KIT" Then Exit Function
    If Not Guard(context, token, notice) Then Exit Function
    Set saved = modExecutionRunStore.ReadVersion(mTarget, CStr(mRun("RunId")), CLng(mRun("Revision")))
    If saved Is Nothing Then Exit Function
    If saved("ContentSha256") <> mHash Or saved("Steps").Count + 1 <> mPending Then Exit Function
    If Not modRecordingSession.ActiveFor(context) Then Exit Function
    If Not modRecordingSession.ExecutionEvidence(context, header, observations) Then Exit Function
    If header("SequenceId") <> saved("Recording")("SequenceId") Or observations.Count <> mBefore Then Exit Function
    If mStop Or mSeedClaimed Or context <> modActivity.CaptureContext() Then Exit Function
    If claim Then mSeedClaimed = True
    PendingSeed = True
Failed:
End Function

Public Function FinishStep(ByVal context As String, ByVal token As String, ByVal delivered As Boolean, ByRef notice As String) As Boolean
    Dim header As Object, observations As Collection, entry As Object, outcome As Object, attempt As Object, definition As Object
    Dim index As Long, controlId As String, valid As Boolean, refs As New Collection, outputs As New Collection, current As Boolean
    On Error GoTo Failed
    If Not Matches(context, token) Or mPending = 0 Or mRun Is Nothing Then Exit Function
    controlId = CStr(mProfile("Steps")(mPending)("ControlId"))
    If controlId = modOperatorStepEvidence.CONTROL_ID Then Exit Function
    If modRecordingSession.ExecutionEvidence(context, header, observations) Then
        If header("SequenceId") = mRun("Recording")("SequenceId") And observations.Count = mBefore + 2 Then
            Set attempt = observations(mBefore + 1): Set outcome = observations(mBefore + 2)
            valid = (attempt("OutcomeCode") = "REQUESTED" And attempt("ControlId") = controlId And _
                outcome("ControlId") = controlId And attempt("ActivityId") = outcome("ActivityId"))
        End If
    End If
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "StepId", mProfile("Steps")(mPending)("StepId"): entry.Add "ControlId", controlId
    entry.Add "State", "Unknown": entry.Add "ReasonCode", "CAPTURE_UNAVAILABLE": entry.Add "ActivityId", ""
    entry.Add "SourceEventRefs", refs: entry.Add "Outputs", outputs
    If valid Then
        entry("ActivityId") = outcome("ActivityId"): entry("ReasonCode") = outcome("OutcomeCode")
        Set entry("SourceEventRefs") = outcome("SourceEventRefs")
        entry("State") = "Failed"
        If delivered And PositiveOwnerOutcome(controlId, CStr(outcome("OutcomeCode"))) Then entry("State") = "Completed"
    End If
    mRun("Steps").Add entry: mPending = 0
    If Not SaveRevision(notice) Then Exit Function
    If entry("State") <> "Completed" Then EndRun CStr(entry("State")), "OWNER_NOT_COMPLETED", notice: Exit Function
    If Not Guard(context, token, notice) Then EndRun "Blocked", "CONTEXT_CHANGED", notice: Exit Function
    If Not modRecordingSession.ActiveFor(context) Then EndRun "Blocked", "CAPTURE_UNAVAILABLE", notice: Exit Function
    FinishStep = True
    If mStop Then
        EndRun "Stopped", "STOP_REQUESTED", notice
    ElseIf mRun("Steps").Count = mProfile("Steps").Count Then
        EndRun "Completed", "", notice
    Else
        notice = "Step completed through its ordinary owner."
    End If
    Exit Function
Failed:
    mPending = 0
    If Not mRun Is Nothing Then EndRun "Unknown", "DISPATCH_FAILED", notice
End Function

Private Function PositiveOwnerOutcome(ByVal controlId As String, ByVal outcome As String) As Boolean
    Select Case controlId
        Case "ADMIN_DEMO_SEED": PositiveOwnerOutcome = (outcome = "APPLIED" Or outcome = "UNCHANGED")
        Case "RECEIVING_OPEN": PositiveOwnerOutcome = (outcome = "OPENED" Or outcome = "REUSED")
        Case "RECEIVING_REFRESH": PositiveOwnerOutcome = (outcome = "REFRESHED")
        Case "RECEIVING_CLEAR": PositiveOwnerOutcome = (outcome = "CLEARED" Or outcome = "EMPTY")
        Case "RECEIVING_SELECT_ITEM": PositiveOwnerOutcome = (outcome = "SELECTED")
        Case "RECEIVING_ADD_SELECTED": PositiveOwnerOutcome = (outcome = "STAGED")
        Case "RECEIVING_CONFIRM_WRITES": PositiveOwnerOutcome = (outcome = "CONFIRMED")
    End Select
End Function

Private Function SaveRevision(ByRef notice As String) As Boolean
    Dim hash As String
    If mWriteFailed Or mRun Is Nothing Then Exit Function
    mRun("PreviousRecordId") = mRun("RecordId"): mRun("PreviousSha256") = mHash
    mRun("RecordId") = modTrainingWire.NewId(): mRun("Revision") = CLng(mRun("Revision")) + 1
    mRun("CreatedAtUTC") = modTrainingWire.UtcTimestamp()
    SaveRevision = modExecutionRunStore.Append(mTarget, mRun, hash, notice)
    If SaveRevision Then
        mHash = hash
    Else
        mWriteFailed = True: mStop = True
        modRecordingSession.InterruptContext mContext, "TRACKING_UNAVAILABLE"
    End If
End Function

Private Sub EndRun(ByVal state As String, ByVal reason As String, ByRef notice As String)
    Dim ignored As String, closed As Boolean
    If mRun Is Nothing Or mPending > 0 Then Exit Sub
    If modRecordingSession.ActiveFor(mContext) Then
        closed = modActionRecording.Control(mContext, "Stop", ignored)
        If Not closed Then modRecordingSession.InterruptContext mContext, "SESSION_CHANGED"
        If Not closed And state = "Completed" Then state = "Unknown": reason = "CAPTURE_UNAVAILABLE"
    End If
    mRun("State") = state: mRun("ReasonCode") = reason
    If SaveRevision(notice) Then notice = state & IIf(reason = "", ". Dispatch is complete; select Verify run for business proof.", ": " & Replace$(LCase$(reason), "_", " ") & ". Partial work was preserved.")
End Sub

Public Sub RequestStop(ByVal context As String, ByVal token As String, ByRef notice As String)
    If Not Matches(context, token) Then Exit Sub
    mStop = True
    If OperatorPending() And Not mOperatorReturning Then mPending = 0
    notice = "Stop requested. The current owner will return before later dispatch is prevented."
    If mPending = 0 And Not mRun Is Nothing Then
        If mRun("State") = "Running" Then EndRun "Stopped", "STOP_REQUESTED", notice
    End If
End Sub

Public Function State(ByVal context As String, ByVal token As String, ByRef completed As Long, ByRef total As Long) As String
    completed = 0: total = 0
    If Not Matches(context, token) Then State = "Unavailable": Exit Function
    total = mProfile("Steps").Count: State = "Ready"
    If Not mRun Is Nothing Then completed = mRun("Steps").Count: State = CStr(mRun("State"))
    If mWriteFailed Then State = "Unknown"
End Function

Public Function Verify(ByVal context As String, ByVal token As String, ByRef text As String, ByRef notice As String) As Boolean
    Dim header As Object, observations As Collection, pathId As String, binding As String, evaluationId As String, saved As Object
    On Error GoTo Failed
    text = "": notice = "Verify requires this completed or stopped attempt and its fresh saved recording."
    If Not Matches(context, token) Or context <> modActivity.CaptureContext() Or mRun Is Nothing Or mPending > 0 Then Exit Function
    If mRun("State") = "Running" Or mWriteFailed Then Exit Function
    Set saved = modExecutionRunStore.ReadVersion(mTarget, CStr(mRun("RunId")), CLng(mRun("Revision")))
    If saved Is Nothing Then Exit Function
    If saved("ContentSha256") <> mHash Then Exit Function
    If Not modRecordingSession.ExecutionEvidence(context, header, observations) Then Exit Function
    If header("RecordType") <> "Close" Or header("SequenceId") <> mRun("Recording")("SequenceId") Then Exit Function
    pathId = CStr(header("ActionPathId"))
    modPathExpectation.BindRun context, header
    binding = modPathExpectation.SelectionBinding(context, pathId)
    If Not modPathExpectation.StageGuide(context, pathId, binding, mKey, notice) Then Exit Function
    If saved("SchemaVersion") = 2 Then
        Verify = modPathEvaluation.Evaluate(context, pathId, "", evaluationId, text, notice, CStr(saved("RunId")), CLng(saved("Revision")), mHash)
    Else
        Verify = modPathEvaluation.Evaluate(context, pathId, "", evaluationId, text, notice)
    End If
Failed:
End Function

Public Function CloseSetup(ByVal context As String, ByVal token As String) As Boolean
    Dim notice As String
    If Not Matches(context, token) Then CloseSetup = True: Exit Function
    RequestStop context, token, notice
    If mStarting Or mPending > 0 Then Exit Function
    mToken = "": mContext = "": mKey = "": Set mGuide = Nothing: Set mProfile = Nothing: Set mRun = Nothing: Set mTarget = Nothing
    CloseSetup = True
End Function
