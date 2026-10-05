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
Private mBefore As Long
Private mStop As Boolean
Private mWriteFailed As Boolean
Private mStarting As Boolean

Public Function OpenSetup(ByVal context As String, ByVal key As String, ByRef token As String, ByRef snapshot As String, _
                          ByRef guideCaption As String, ByRef profileCaption As String, ByRef targetCaption As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, guide As Object, profile As Object, builds As Collection, policyHash As String
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
    If Not modExecutionTarget.ContainsEntity(context, entity, notice) Then Exit Function
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

Public Function BeginStep(ByVal context As String, ByVal token As String, ByRef controlId As String, ByRef inputs As String, ByRef notice As String) As Boolean
    Dim header As Object, observations As Collection, step As Object, item As Object, binding As Object
    On Error GoTo Failed
    controlId = "": inputs = ""
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
    notice = "Running: " & CStr(mGuide("Steps")(mPending)("Caption"))
    BeginStep = True: Exit Function
Failed:
    If Not mRun Is Nothing Then EndRun "Unknown", "CAPTURE_UNAVAILABLE", notice
End Function

Public Function FinishStep(ByVal context As String, ByVal token As String, ByVal delivered As Boolean, ByRef notice As String) As Boolean
    Dim header As Object, observations As Collection, entry As Object, outcome As Object, attempt As Object, definition As Object
    Dim index As Long, controlId As String, valid As Boolean, refs As New Collection, outputs As New Collection, current As Boolean
    On Error GoTo Failed
    If Not Matches(context, token) Or mPending = 0 Or mRun Is Nothing Then Exit Function
    controlId = CStr(mProfile("Steps")(mPending)("ControlId"))
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
    Verify = modPathEvaluation.Evaluate(context, pathId, "", evaluationId, text, notice)
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
