Attribute VB_Name = "modPathEvaluation"
Option Explicit

' Declared primitive Operations/Core bridge. All evidence and file objects stay in Core.
Private mBusy As Boolean

Public Function AcceptLoaded(ByVal context As String, ByVal publicationId As String, ByVal loadedAtUTC As String) As Boolean
    AcceptLoaded = modLoadedEvents.Accept(context, publicationId, loadedAtUTC)
End Function

Public Sub ClearLoaded(ByVal context As String)
    modLoadedEvents.ClearContext context
End Sub

Public Sub StaleLoaded(ByVal context As String)
    modLoadedEvents.MarkStale context
End Sub

Public Function SelectedBinding(ByVal context As String, ByVal pathId As String) As String
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    SelectedBinding = modPathExpectation.SelectionBinding(context, pathId)
End Function

Public Function SelectedRunCaption(ByVal context As String, ByVal pathId As String, ByVal binding As String) As String
    SelectedRunCaption = modPathExpectation.RunCaption(context, pathId, binding)
End Function

Public Function UseGuide(ByVal context As String, ByVal pathId As String, ByVal binding As String, _
                         ByVal key As String, ByRef notice As String) As Boolean
    UseGuide = modPathExpectation.StageGuide(context, pathId, binding, key, notice)
End Function

Public Function Evaluate(ByVal context As String, ByVal pathId As String, ByVal previousId As String, _
                         ByRef evaluationId As String, ByRef text As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, policy As Object, header As Object, definition As Object, source As String
    Dim currentHeader As Object, records As Collection, result As Object, historical As Object, visible As Object
    Dim publication As Object, provenance As Object, terminal As Object, previous As Object, step As Object
    Dim version As Long, collect As Boolean, show As Boolean, available As Boolean, ignored As String
    Dim binding As String, guide As Object
    On Error GoTo Failed
    evaluationId = "": text = "": notice = "Incomplete evidence: evaluation is unavailable."
    If mBusy Then Exit Function
    mBusy = True
    binding = modPathExpectation.EvaluationBinding(context, pathId)
    If binding = "" Then GoTo Done
    If Not ReadContext(context, target, policy, notice, version) Then GoTo Done
    If Not modPathExpectation.ReadSelected(context, pathId, header, definition, source, notice) Then GoTo Done
    If binding <> modPathExpectation.EvaluationBinding(context, pathId) Then GoTo Failed
    Set result = modEvaluationModel.Create(header, definition, source, version)
    If Not modPathExpectation.GuideReference(context, pathId, guide, notice) Then GoTo Done
    Set result("Guide") = guide
    available = modLoadedEvents.Read(context, publication, provenance): result.Add "Publication", provenance
    Set records = modRecordingReader.ReadRun(target, pathId, CLng(header("Version")), currentHeader, ignored)
    If Not records Is Nothing Then
        If currentHeader("RecordId") <> header("RecordId") Or currentHeader("ContentSha256") <> header("ContentSha256") Then Set records = Nothing
    End If
    If records Is Nothing Then
        For Each step In definition("Steps"): result("UnavailableSteps").Add CStr(step("StepId")): Next step
        modEvaluationModel.Reason result, "JOURNAL_UNAVAILABLE"
    Else
        Set historical = CreateObject("Scripting.Dictionary")
        If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, show, ignored, historical, requestedVersion:=CLng(header("PolicyVersion"))) Then Set historical = Nothing
        Set historical = modEvaluationMatches.PolicyControls(historical, True)
        Set visible = modEvaluationMatches.PolicyControls(policy, False)
        Set terminal = modEvaluationMatches.Match(result, records, historical, visible)
        Classify result, header, terminal, publication
    End If
    If previousId <> "" Then
        Set previous = modEvaluationStore.Read(target, previousId)
        If Not previous Is Nothing Then
            If previous("ActionPathId") = pathId And previous("JournalSha256") = header("ContentSha256") Then result("PreviousEvaluationId") = previousId
        End If
    End If
    If context <> modActivity.CaptureContext() Then GoTo Failed
    If binding <> modPathExpectation.EvaluationBinding(context, pathId) Then GoTo Failed
    If Not modPathExpectation.GuideReference(context, pathId, guide, notice) Then GoTo Done
    If binding <> modPathExpectation.EvaluationBinding(context, pathId) Then GoTo Failed
    If Not modEvaluationStore.Append(target, result, notice) Then GoTo Done
    evaluationId = CStr(result("EvaluationId"))
    Evaluate = ReadSaved(context, pathId, evaluationId, text, notice)
Done:
    mBusy = False
    Exit Function
Failed:
    text = "": notice = "Incomplete evidence: evaluation could not be completed."
    Resume Done
End Function

Private Sub Classify(ByVal result As Object, ByVal header As Object, ByVal terminal As Object, ByVal publication As Object)
    If Not terminal Is Nothing Then
        If result("ExpectedConclusion")("TerminalKind") = "SourceEventsApplied" Then modEvaluationSources.Retain result, terminal, publication
    End If
    If header("Lifecycle") = "Cancelled" Then
        result("ResultState") = "Cancelled": modEvaluationModel.Reason result, "CAPTURE_CANCELLED"
    ElseIf header("Lifecycle") <> "Stopped" Or header("RecordType") <> "Close" Then
        modEvaluationModel.Reason result, "CAPTURE_INCOMPLETE"
    ElseIf result("UnavailableSteps").Count > 0 Then
        result("ResultState") = "Incomplete"
    ElseIf result("ExpectedConclusion")("TerminalKind") = "None" Then
        modEvaluationModel.Reason result, "NO_EXPECTATION"
    ElseIf result("Publication")("Availability") <> "Loaded" Then
        modEvaluationModel.Reason result, IIf(result("Publication")("Availability") = "Stale", "PUBLICATION_STALE", "PUBLICATION_UNAVAILABLE")
    ElseIf result("MissingSteps").Count > 0 Or result("FailedSteps").Count > 0 Then
        result("ResultState") = "Failed"
    ElseIf terminal Is Nothing Then
        modEvaluationModel.Reason result, "TERMINAL_NOT_COMPLETED"
    ElseIf result("ExpectedConclusion")("TerminalKind") = "SourceEventsApplied" Then
        modEvaluationSources.Assess result
    ElseIf modEvaluationMatches.CommandCompleted(terminal) Then
        result("ResultState") = "Concluded": modEvaluationModel.Reason result, "COMMAND_COMPLETED"
    Else
        result("ResultState") = "Failed": modEvaluationModel.Reason result, "TERMINAL_NOT_COMPLETED"
    End If
End Sub

Public Function ReadSaved(ByVal context As String, ByVal pathId As String, ByVal evaluationId As String, _
                          ByRef text As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, policy As Object, result As Object, header As Object, records As Collection
    Dim visible As Object, record As Object, step As Object, displayed As Object, id As Variant
    Dim binding As String
    On Error GoTo Invalid
    text = "": notice = "Incomplete evidence: the saved diagnostic result is unavailable."
    binding = modPathExpectation.SelectionBinding(context, pathId)
    If binding = "" Then Exit Function
    If Not ReadContext(context, target, policy, notice) Then Exit Function
    notice = "Incomplete evidence: the saved diagnostic result is unavailable."
    Set result = modEvaluationStore.Read(target, evaluationId)
    If result Is Nothing Then Exit Function
    If result("ActionPathId") <> pathId Then Exit Function
    Set records = modRecordingReader.ReadRun(target, pathId, CLng(result("JournalVersion")), header, notice)
    notice = "Incomplete evidence: the saved diagnostic result cannot be verified."
    If records Is Nothing Then Exit Function
    If Not modPathExpectation.MatchesHeader(context, pathId, header) Then Exit Function
    If Not modEvaluationJournal.Matches(result, header, records) Then Exit Function
    If result("ExpectationSource") = "Guide expectation" Then
        If Not modGuideExpectation.Matches(context, result("Guide"), result("ExpectedConclusion"), notice) Then Exit Function
    End If
    Set visible = modEvaluationMatches.PolicyControls(policy, False)
    Set displayed = CreateObject("Scripting.Dictionary")
    For Each record In result("Matches"): displayed.Add CStr(record("ActivityId")), True: Next record
    For Each id In result("ExtraActivityIds"): displayed.Add CStr(id), True: Next id
    For Each record In records
        If displayed.Exists(record("ActivityId")) Then
            If Not modEvaluationMatches.Permitted(visible, CStr(record("ControlId"))) Then Exit Function
        End If
    Next record
    For Each step In result("ExpectedConclusion")("Steps")
        If Not modEvaluationMatches.Permitted(visible, CStr(step("ControlId"))) Then Exit Function
    Next step
    If context <> modActivity.CaptureContext() Then Exit Function
    If binding <> modPathExpectation.SelectionBinding(context, pathId) Then Exit Function
    text = modEvaluationPresentation.Render(result)
    notice = modEvaluationPresentation.Caption(result)
    ReadSaved = True
    Exit Function
Invalid:
    text = "": notice = "Incomplete evidence: the saved diagnostic result could not be read."
End Function

Private Function ReadContext(ByVal context As String, ByRef target As WarehouseTarget, ByRef policy As Object, ByRef notice As String, Optional ByRef version As Long = 0) As Boolean
    Dim collect As Boolean, visible As Boolean
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    Set target = modNasConnection.GetCurrentTarget(): Set policy = CreateObject("Scripting.Dictionary")
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice, policy) Then Exit Function
    ReadContext = (context = modActivity.CaptureContext())
End Function
