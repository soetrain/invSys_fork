Attribute VB_Name = "modActivity"
Option Explicit

Private mBusy As Boolean
Private mContext As String
Private mActions As Object

' Primitive context is opaque local UI state; it never enters a persisted record.
Public Function CaptureContext() As String
    Dim target As WarehouseTarget
    On Error GoTo Unavailable
    If Not modAuth.IsSignedIn() Or Not modNasConnection.IsCurrentTargetAllowed(True) Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    CaptureContext = modTrainingJson.Quote(target.WarehouseId) & modTrainingJson.Quote(target.StationId) & _
        modTrainingJson.Quote(target.RuntimeRoot) & modTrainingJson.Quote(target.ConfigPath) & _
        modTrainingJson.Quote(modAuth.GetCurrentUserId()) & CStr(modAuthSession.Version())
Unavailable:
End Function

Public Function BeginAction(ByVal controlId As String, ByVal capturedContext As String, _
                            Optional ByRef notice As String = "") As String
    Dim target As WarehouseTarget, definition As Object, action As Object, id As String
    Dim version As Long, collect As Boolean, visible As Boolean, current As String, key As Variant
    On Error GoTo Failed
    notice = ""
    If mBusy Then Exit Function
    current = CaptureContext()
    If current = "" Then Exit Function
    notice = "Tracking unavailable: the form context changed. Reopen the form."
    If capturedContext = "" Or current <> capturedContext Then Exit Function
    Set definition = modActivityCatalog.Control(controlId)
    If definition Is Nothing Then notice = "Tracking unavailable: the control is unregistered.": Exit Function
    mBusy = True
    Set target = modNasConnection.GetCurrentTarget()
    If Not modActivityPolicy.ReadPolicy(target, controlId, version, collect, visible, notice) Then GoTo CleanExit
    If Not collect Then GoTo CleanExit
    If mActions Is Nothing Or mContext <> current Then
        Set mActions = CreateObject("Scripting.Dictionary")
        mContext = current
    End If
    If mActions.Count >= 512 Then
        For Each key In mActions.Keys
            If mActions(key).Exists("OutcomeBody") Then mActions.Remove key: Exit For
        Next key
        If mActions.Count >= 512 Then notice = "Tracking unavailable: too many unfinished actions.": GoTo CleanExit
    End If
    id = modTrainingWire.NewId()
    Set action = CreateObject("Scripting.Dictionary")
    action.Add "ActivityId", id
    action.Add "Context", current
    action.Add "ControlId", controlId
    action.Add "PolicyVersion", version
    action.Add "Target", target
    action.Add "UserId", modAuth.GetCurrentUserId()
    action.Add "AttemptBody", MakeBody(action, id, "REQUESTED")
    If Not modActivityStore.Append(target, id, action("AttemptBody"), notice) Then GoTo CleanExit
    mActions.Add id, action
    BeginAction = id
CleanExit:
    mBusy = False
    Exit Function
Failed:
    notice = "Tracking unavailable: the action could not be recorded."
    Resume CleanExit
End Function

Public Function FinishAction(ByVal activityId As String, ByVal outcomeCode As String, _
                             Optional ByRef notice As String = "", _
                             Optional ByVal sourceReferencesJson As String = "[]") As Boolean
    Dim action As Object, target As WarehouseTarget, definition As Object, outcome As Object
    Dim version As Long, collect As Boolean, visible As Boolean, ignored As String, permitted As Boolean
    Dim references As Collection, canonicalReferences As String
    On Error GoTo Failed
    If activityId = "" Then Exit Function
    notice = "Tracking unavailable: the action context is no longer current."
    If mBusy Or mActions Is Nothing Then Exit Function
    If Not mActions.Exists(activityId) Then Exit Function
    Set action = mActions(activityId)
    If CaptureContext() = "" Or action("Context") <> CaptureContext() Then Exit Function
    Set target = action("Target")
    Set definition = modActivityCatalog.Control(action("ControlId"))
    Set outcome = modActivityCatalog.Outcome(action("ControlId"), outcomeCode)
    If outcome Is Nothing Or outcomeCode = "REQUESTED" Then Exit Function
    Set references = modActivityReferences.Decode(target.WarehouseId, action("ControlId"), outcomeCode, sourceReferencesJson)
    If references Is Nothing Then notice = "Tracking unavailable: source references are invalid.": Exit Function
    canonicalReferences = modActivityReferences.Encode(references)
    mBusy = True
    If Not modActivityPolicy.ReadPolicy(target, action("ControlId"), version, collect, visible, notice) Then GoTo CleanExit
    If Not collect Or version <> CLng(action("PolicyVersion")) Then
        notice = "Tracking unavailable: the tracking policy changed during this action."
        GoTo CleanExit
    End If
    If outcomeCode = "COMPLETED" Or outcomeCode = "UNCHANGED" Or outcomeCode = "CONFIRMED" Or outcomeCode = "PENDING" Then
        permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached(definition("Capability"), ignored)
        If Not permitted And definition("Role") = "Production" Then _
            permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", ignored)
        If Not permitted Then notice = "Tracking unavailable: completion is not authorized.": GoTo CleanExit
    End If
    If action.Exists("OutcomeBody") Then
        If action("OutcomeCode") <> outcomeCode Or action("ReferencesJson") <> canonicalReferences Then
            notice = "Tracking unavailable: conflicting completion.": GoTo CleanExit
        End If
    Else
        action.Add "OutcomeCode", outcomeCode
        action.Add "ReferencesJson", canonicalReferences
        action.Add "SourceEventRefs", references
        action.Add "OutcomeId", modTrainingWire.NewId()
        action.Add "OutcomeBody", MakeBody(action, action("OutcomeId"), outcomeCode)
    End If
    FinishAction = modActivityStore.Append(target, action("OutcomeId"), action("OutcomeBody"), notice)
CleanExit:
    mBusy = False
    Exit Function
Failed:
    notice = "Tracking unavailable: the result could not be recorded."
    Resume CleanExit
End Function

' Serialization only: callers supply owner-observed identities and submission facts.
Public Function InventorySourceReferences(ByVal activityId As String, ByVal eventIds As String, _
                                          ByVal submissionState As String) As String
    Dim target As WarehouseTarget
    On Error GoTo Unavailable
    If mActions Is Nothing Then Exit Function
    If Not mActions.Exists(activityId) Then Exit Function
    Set target = mActions(activityId)("Target")
    InventorySourceReferences = modActivityReferences.Inventory(target.WarehouseId, eventIds, submissionState)
Unavailable:
End Function

Public Function ReadActivityRecord(ByVal recordId As String) As String
    Dim target As WarehouseTarget, text As String, record As Object
    Dim version As Long, collect As Boolean, visible As Boolean, notice As String
    On Error GoTo Unavailable
    ReadActivityRecord = "UNAVAILABLE|Recorded activity is unavailable."
    If CaptureContext() = "" Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    text = modActivityStore.ReadRecord(target, recordId)
    If text = "" Then Exit Function
    Set record = modTrainingJson.DecodeObject(text)
    If record Is Nothing Then Exit Function
    If Not modActivityPolicy.ReadPolicy(target, record("ControlId"), version, collect, visible, notice) Then Exit Function
    If Not visible Then ReadActivityRecord = "UNAVAILABLE|Hidden by policy.": Exit Function
    ReadActivityRecord = "OK|" & text
Unavailable:
End Function

Private Function MakeBody(ByVal action As Object, ByVal recordId As String, ByVal outcomeCode As String) As String
    Dim record As Object, definition As Object, outcome As Object, key As Variant
    Dim target As WarehouseTarget, source As Workbook, references As Collection
    Set record = CreateObject("Scripting.Dictionary")
    Set definition = modActivityCatalog.Control(action("ControlId"))
    Set outcome = modActivityCatalog.Outcome(action("ControlId"), outcomeCode)
    Set target = action("Target")
    If definition("Role") = "Admin" Then
        Set source = Application.Workbooks("invSys.Admin.xlam")
    Else
        Set source = Application.Workbooks("invSys.Operations.xlam")
    End If
    record.Add "SchemaVersion", 1&
    record.Add "CatalogVersion", modActivityCatalog.CATALOG_VERSION
    record.Add "PackageSetVersion", CStr(source.CustomDocumentProperties("invSysPackageSetVersion").Value)
    record.Add "BuildIdentity", CStr(source.CustomDocumentProperties("invSysBuildIdentity").Value)
    record.Add "RecordId", recordId
    record.Add "ActivityId", action("ActivityId")
    record.Add "SequenceId", ""
    record.Add "WarehouseId", target.WarehouseId
    record.Add "StationId", target.StationId
    record.Add "UserId", action("UserId")
    record.Add "ControlId", action("ControlId")
    record.Add "OwnerId", definition("OwnerId")
    record.Add "SourceKind", "User activity"
    record.Add "SourceRole", definition("Role")
    record.Add "Caption", definition("Caption")
    record.Add "Surface", definition("Surface")
    record.Add "Ordinal", 0&
    record.Add "OccurredAtUTC", modTrainingWire.UtcTimestamp()
    record.Add "PolicyVersion", CLng(action("PolicyVersion"))
    For Each key In outcome.Keys
        record.Add key, outcome(key)
    Next key
    If action.Exists("SourceEventRefs") Then
        Set references = action("SourceEventRefs")
    Else
        Set references = New Collection
    End If
    record.Add "SourceEventRefs", references
    MakeBody = modTrainingJson.EncodeObject(record)
End Function
