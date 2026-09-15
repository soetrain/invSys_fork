Attribute VB_Name = "modGuideDraftSource"
Option Explicit
Option Private Module

' Original training evidence stays in Core; no authority workbook is opened here.
Public Function ReadSelected(ByVal context As String, ByVal pathId As String, ByRef header As Object, _
                             ByRef records As Collection, ByRef visible As Object, ByRef policyHash As String, _
                             ByRef notice As String, Optional ByRef policyVersion As Long = 0) As Boolean
    Dim target As WarehouseTarget, policy As Object, current As Object, definition As Object
    Dim binding As String, source As String, version As Long, collect As Boolean, show As Boolean
    On Error GoTo Invalid
    Set header = Nothing: Set records = Nothing: Set visible = Nothing: policyHash = "": policyVersion = 0
    notice = "Unavailable: the selected recording, session or guide permission changed. Reopen Action Paths."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    binding = modPathExpectation.SelectionBinding(context, pathId)
    If binding = "" Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    If Not modAuth.CanPerform("ACTION_PATH_MAINT", modAuth.GetCurrentUserId(), target.WarehouseId, target.StationId) Then Exit Function
    If Not modPathExpectation.ReadSelected(context, pathId, header, definition, source, notice) Then Exit Function
    Set policy = CreateObject("Scripting.Dictionary")
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, show, notice, policy) Then GoTo Invalid
    Set records = modRecordingReader.ReadRun(target, pathId, CLng(header("Version")), current, notice)
    If records Is Nothing Then GoTo Invalid
    If Not modPathExpectation.MatchesHeader(context, pathId, current) Then GoTo Invalid
    If binding <> modPathExpectation.SelectionBinding(context, pathId) Or context <> modActivity.CaptureContext() Then GoTo Invalid
    Set visible = modEvaluationMatches.PolicyControls(policy, False)
    ' ReadPolicy validated this integer; Excel's Value2 projects it as Double.
    policy("SavedCatalogVersion") = CLng(policy("SavedCatalogVersion"))
    policyHash = modTrainingWire.Sha256(modTrainingJson.EncodeObject(policy))
    If Not modEvaluationModel.IsHash(policyHash) Then GoTo Invalid
    policyVersion = version: ReadSelected = True
    Exit Function
Invalid:
    Set header = Nothing: Set records = Nothing: Set visible = Nothing: policyHash = ""
    notice = "Unavailable: the source recording or current policy could not be validated. Reopen Action Paths."
End Function

Public Function SourceCaption(ByVal header As Object) As String
    Dim lifecycle As String
    lifecycle = CStr(header("Lifecycle"))
    If header("RecordType") <> "Close" Then lifecycle = "Interrupted"
    SourceCaption = "Source Action Path: " & CStr(header("ActionPathId")) & vbCrLf & _
        "Sequence: " & CStr(header("SequenceId")) & "; journal version: " & CStr(header("Version")) & _
        "; " & lifecycle & ". Capture lifecycle is not a conclusion."
End Function

Public Function Evidence(ByVal records As Collection, ByVal visible As Object) As String
    Dim record As Object, reference As Object, hidden As Long
    For Each record In records
        If modEvaluationMatches.Permitted(visible, CStr(record("ControlId"))) Then
            Evidence = Evidence & "Observed control " & CStr(record("Ordinal")) & ": " & CStr(record("Caption")) & _
                " - " & CStr(record("OutcomeCode")) & vbCrLf & CStr(record("ActivityId")) & vbCrLf & _
                CStr(record("OccurredAtUTC")) & vbCrLf
            For Each reference In record("SourceEventRefs")
                Evidence = Evidence & "Source event: " & CStr(reference("EventId")) & " (" & _
                    CStr(reference("SubmissionState")) & "; application not asserted)" & vbCrLf
            Next reference
            Evidence = Evidence & vbCrLf
        Else
            hidden = hidden + 1
        End If
    Next record
    If hidden > 0 Then Evidence = "Incomplete evidence: current policy restricts " & CStr(hidden) & " observation(s)." & vbCrLf & Evidence
End Function
