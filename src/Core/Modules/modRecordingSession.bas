Attribute VB_Name = "modRecordingSession"
Option Explicit
Option Private Module

Private mTarget As WarehouseTarget
Private mContext As String
Private mHeader As Object
Private mObservations As Collection
Private mSeen As Object
Private mPending As Object
Private mActive As Boolean
Private mBusy As Boolean
Private mCount As Long
Private mVersion As Long
Private mPreviousId As String
Private mPreviousHash As String
Private mReport As String
Private mExpectation As Object

Public Function ExecuteCommand(ByVal context As String, ByVal command As String, ByRef report As String) As Boolean
    On Error GoTo Failed
    report = "Recording unavailable: the session or warehouse changed. Reopen Viewer."
    If context = "" Or context <> modActivity.CaptureContext() Or mBusy Then Exit Function
    Select Case command
        Case "Start": ExecuteCommand = Start(context, report)
        Case "Stop", "Cancel"
            If Not mActive Or mContext <> context Then report = "No recording is active.": Exit Function
            If command = "Cancel" Then
                ExecuteCommand = CloseRun("Cancelled", "")
            ElseIf mPending.Count > 0 Then
                ExecuteCommand = CloseRun("Incomplete", "UNFINISHED_ACTIONS")
            Else
                ExecuteCommand = CloseRun("Stopped", "")
            End If
            report = mReport
        Case Else: report = "Recording command is unavailable."
    End Select
    Exit Function
Failed:
    mBusy = False
    Interrupt "TRACKING_UNAVAILABLE"
    report = "Recording unavailable: the command could not be completed."
End Function

Private Function Start(ByVal context As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget, captureEnabled As Boolean, version As Long, tags As New Collection
    If mActive And mContext = context Then report = "A recording is already active.": Exit Function
    If mActive Then Interrupt "SESSION_CHANGED"
    Set target = modNasConnection.GetCurrentTarget()
    If Not ReadPolicySnapshot(target, captureEnabled, version, report) Then Exit Function
    If Not captureEnabled Then report = "Recording capture is off.": Exit Function
    If context <> modActivity.CaptureContext() Then report = "Recording unavailable: session changed.": Exit Function
    Set mTarget = New WarehouseTarget
    mTarget.WarehouseId = target.WarehouseId: mTarget.StationId = target.StationId
    mTarget.RuntimeRoot = target.RuntimeRoot: mTarget.HubRoot = target.HubRoot
    mTarget.ConfigPath = target.ConfigPath: mTarget.SourceType = target.SourceType
    Set mHeader = CreateObject("Scripting.Dictionary")
    mHeader.Add "SchemaVersion", 2&: mHeader.Add "RecordKind", "Recording"
    mHeader.Add "ActionPathId", modTrainingWire.NewId(): mHeader.Add "SequenceId", modTrainingWire.NewId()
    mHeader.Add "WarehouseId", mTarget.WarehouseId: mHeader.Add "OriginWarehouseId", mTarget.WarehouseId
    mHeader.Add "CreatedByUserId", modAuth.GetCurrentUserId()
    mHeader.Add "PolicyVersion", version: mHeader.Add "CatalogVersion", modActivityCatalog.CATALOG_VERSION
    mHeader.Add "PackageSetVersion", CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value)
    mHeader.Add "BuildIdentity", CStr(ThisWorkbook.CustomDocumentProperties("invSysBuildIdentity").Value)
    mHeader.Add "Name", "Recorded sequence": mHeader.Add "Tags", tags
    mHeader.Add "Instructions", "": mHeader.Add "Method", "Diagnostic"
    mContext = context: mCount = 0: mVersion = 0: mPreviousId = "": mPreviousHash = ""
    Set mObservations = New Collection
    Set mExpectation = modExpectationModel.NoneDefinition()
    modExpectationDraft.Discard ""
    Set mSeen = CreateObject("Scripting.Dictionary"): Set mPending = CreateObject("Scripting.Dictionary")
    mBusy = True
    Start = AppendEntry("Start", "Recording", "", mObservations)
    mBusy = False: mActive = Start
    If Start Then mReport = "Recording: 0 / 256 actions."
    report = mReport
End Function

Public Function ReadStatus(ByVal context As String, ByRef active As Boolean, ByRef canStart As Boolean) As String
    Dim captureEnabled As Boolean, version As Long, notice As String, target As WarehouseTarget
    On Error GoTo Unavailable
    active = False: canStart = False
    ReadStatus = "Recording unavailable: session changed. Reopen Viewer."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    If mBusy Then ReadStatus = "Recording status is being updated.": Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If Not ReadPolicySnapshot(target, captureEnabled, version, notice) Then ReadStatus = notice: Exit Function
    If context <> modActivity.CaptureContext() Then Exit Function
    If mContext = context And mActive Then
        If version <> CLng(mHeader("PolicyVersion")) Or Not captureEnabled Then
            ReadStatus = "Incomplete evidence: tracking policy changed. Reopen Viewer.": Exit Function
        End If
        active = True
        ReadStatus = "Recording: " & CStr(mCount) & " / 256 actions.": Exit Function
    End If
    canStart = captureEnabled
    ReadStatus = IIf(canStart, "Ready to record: 0 / 256 actions.", "Recording capture is off.")
    If mContext = context And mReport <> "" And canStart Then ReadStatus = mReport
    Exit Function
Unavailable:
    active = False: canStart = False
    ReadStatus = "Recording unavailable: tracking policy could not be read."
End Function

Private Function ReadPolicySnapshot(ByVal target As WarehouseTarget, ByRef captureEnabled As Boolean, ByRef version As Long, ByRef report As String) As Boolean
    Dim collect As Boolean, visible As Boolean
    ReadPolicySnapshot = modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, report, Nothing, captureEnabled)
End Function

Public Sub Prepare(ByVal action As Object, ByVal captureEnabled As Boolean, ByVal eligible As Boolean, ByRef collect As Boolean)
    Dim definition As Object
    action.Add "SequenceId", "": action.Add "Ordinal", 0&: action.Add "CaptureCollected", False
    If Not mActive Then Exit Sub
    If mContext <> action("Context") Then Interrupt "SESSION_CHANGED": Exit Sub
    If CLng(action("PolicyVersion")) <> CLng(mHeader("PolicyVersion")) Then Interrupt "POLICY_CHANGED": Exit Sub
    If Not captureEnabled Then Interrupt "POLICY_CHANGED": Exit Sub
    Set definition = modActivityCatalog.Control(CStr(action("ControlId")))
    If Not eligible Then
        If definition("Class") = "Command" Then Interrupt "TRACKING_UNAVAILABLE"
        Exit Sub
    End If
    If Not collect And definition("Class") = "Navigation" Then collect = True: action("CaptureCollected") = True
    If Not collect Then Interrupt "TRACKING_UNAVAILABLE": Exit Sub
    If mCount >= 256 Then Interrupt "ACTION_LIMIT": Exit Sub
    mCount = mCount + 1
    action("SequenceId") = mHeader("SequenceId"): action("Ordinal") = mCount
End Sub

Public Function ActiveFor(ByVal context As String) As Boolean
    ActiveFor = (mActive And mContext = context)
End Function

Public Function ReadExpectation(ByVal context As String, ByRef sequenceId As String, _
                                ByRef definition As Object, ByRef notice As String) As Boolean
    Dim active As Boolean, canStart As Boolean
    sequenceId = "": Set definition = Nothing
    notice = ReadStatus(context, active, canStart)
    If Not active Or mBusy Or mContext <> context Then Exit Function
    sequenceId = CStr(mHeader("SequenceId"))
    Set definition = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(mExpectation))
    ReadExpectation = Not definition Is Nothing
End Function

Public Function StageExpectation(ByVal context As String, ByVal sequenceId As String, _
                                 ByVal definition As Object, ByRef notice As String) As Boolean
    Dim currentSequence As String, current As Object
    If Not ReadExpectation(context, currentSequence, current, notice) Then Exit Function
    If sequenceId = "" Or sequenceId <> currentSequence Then notice = "This recording has ended. Reopen Expected conclusion.": Exit Function
    notice = "Choose valid expected steps and a conclusion, or remove all steps and choose None."
    If Not modExpectationModel.Validate(definition, CLng(mHeader("CatalogVersion"))) Then Exit Function
    Set mExpectation = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(definition))
    StageExpectation = Not mExpectation Is Nothing
    If StageExpectation Then notice = "Expected conclusion staged for this recording. Stop freezes it with the observations."
End Function

Public Sub Observe(ByVal body As String, ByRef notice As String)
    Dim record As Object, one As New Collection, id As String, activityId As String
    On Error GoTo Failed
    If Not mActive Or mBusy Then Exit Sub
    Set record = modTrainingJson.DecodeObject(body)
    If record Is Nothing Then GoTo Failed
    If record("SequenceId") <> mHeader("SequenceId") Then Exit Sub
    id = record("RecordId"): activityId = record("ActivityId")
    If mSeen.Exists(id) Then
        If mSeen(id) <> body Then GoTo Failed
        Exit Sub
    End If
    one.Add record
    mBusy = True
    If Not AppendEntry("Observation", "Recording", "", one) Then
        notice = mReport: mActive = False: mBusy = False: Exit Sub
    End If
    mSeen.Add id, body: mObservations.Add record
    If record("OutcomeCode") = "REQUESTED" Then
        mPending.Add activityId, True
    ElseIf mPending.Exists(activityId) Then
        mPending.Remove activityId
    End If
    mBusy = False
    If mCount = 256 And mPending.Count = 0 Then Interrupt "ACTION_LIMIT"
    Exit Sub
Failed:
    mBusy = False
    Interrupt "TRACKING_UNAVAILABLE"
    notice = "Incomplete evidence: recording could not retain an observation."
End Sub

Public Sub InterruptContext(ByVal context As String, ByVal reason As String)
    If context <> "" And context = mContext Then Interrupt reason
End Sub

Public Sub InterruptAction(ByVal action As Object, ByVal reason As String)
    If Not mActive Or mBusy Then Exit Sub
    If action Is Nothing Then Exit Sub
    If Not action.Exists("Context") Or Not action.Exists("SequenceId") Then Exit Sub
    If action("Context") = mContext And action("SequenceId") = mHeader("SequenceId") Then Interrupt reason
End Sub

Public Sub Interrupt(ByVal reason As String)
    If Not mActive Or mBusy Then Exit Sub
    CloseRun "Incomplete", reason
End Sub

Private Function CloseRun(ByVal life As String, ByVal reason As String) As Boolean
    On Error GoTo Failed
    mBusy = True
    CloseRun = AppendEntry("Close", life, reason, mObservations)
    If CloseRun Then
        Select Case life
            Case "Stopped": mReport = "Stopped. Capture frozen; conclusion not evaluated."
            Case "Cancelled": mReport = "Cancelled. Observed work was not undone."
            Case Else: mReport = "Incomplete evidence: " & Replace$(LCase$(reason), "_", " ") & "."
        End Select
        If reason = "ACTION_LIMIT" Then mReport = "Partial: action limit reached. 256 / 256 actions."
    End If
Failed:
    If Err.Number <> 0 Then mReport = "Incomplete evidence: recording closure could not be saved."
    mActive = False: mBusy = False
    modExpectationDraft.Discard ""
    Set mExpectation = Nothing
End Function

Private Function AppendEntry(ByVal kind As String, ByVal life As String, ByVal reason As String, ByVal observations As Collection) As Boolean
    Dim model As Object, field As Variant, hash As String
    Set model = CreateObject("Scripting.Dictionary")
    For Each field In mHeader.Keys: model.Add field, mHeader(field): Next field
    model.Add "RecordType", kind: model.Add "RecordId", modTrainingWire.NewId()
    model.Add "Version", mVersion + 1: model.Add "PreviousRecordId", mPreviousId
    model.Add "PreviousSha256", mPreviousHash: model.Add "Lifecycle", life
    model.Add "ReasonCode", reason: model.Add "ActionCount", mCount
    model.Add "CreatedAtUTC", modTrainingWire.UtcTimestamp(): model.Add "Observations", observations
    If kind = "Close" Then
        model.Add "ExpectedConclusion", mExpectation
    Else
        model.Add "ExpectedConclusion", modExpectationModel.NoneDefinition()
    End If
    AppendEntry = modRecordingJournal.Append(mTarget, model, hash, mReport)
    If AppendEntry Then
        mVersion = mVersion + 1: mPreviousId = model("RecordId"): mPreviousHash = hash
    End If
End Function
