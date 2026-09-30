Attribute VB_Name = "modProductionWorksheetActions"
Option Explicit
Option Private Module

Public Enum ProductionWorksheetCommand
    WorksheetSend = 1
    WorksheetAddItem
    WorksheetRetrieve
End Enum

Public Function Execute(ByVal form As frmProduction, ByVal operatorBook As Workbook, _
                        ByVal context As String, ByVal command As ProductionWorksheetCommand, _
                        ByVal loading As Boolean, ByRef busy As Boolean) As String
    Dim action As cProductionWorksheetAction, report As String, suffix As String, rejected As Boolean
    Dim errorNumber As Long, errorSource As String, errorText As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    Select Case command
        Case WorksheetSend: suffix = "SEND"
        Case WorksheetAddItem: suffix = "ADD_ITEM"
        Case WorksheetRetrieve: suffix = "RETRIEVE"
        Case Else: Exit Function
    End Select
    busy = True
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_PROCESS_WORKSHEET_" & suffix, context, operatorBook, report) Then GoTo Done
    Select Case command
        Case WorksheetSend
            If form.SendProcessWorksheetDraft(action, report, rejected) Then
                action.OutcomeCode = "STAGED"
            ElseIf rejected Then
                action.OutcomeCode = "REJECTED"
            End If
        Case WorksheetAddItem
            If modProductionProcessWorksheet.AddAcceptableItemPairToSelectedTable(operatorBook, report) Then
                If action.CanContinue(report) Then action.OutcomeCode = "STAGED"
            Else
                action.OutcomeCode = "REJECTED"
            End If
        Case WorksheetRetrieve
            Retrieve form, action, report
    End Select
Done:
    If Not action Is Nothing Then action.Finish report
    busy = False
    Execute = report
    ' Add originally propagates unexpected service errors; retain that behavior.
    If errorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise errorNumber, errorSource, errorText
    End If
    Exit Function
Failed:
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    If command = WorksheetAddItem Then
        errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    ElseIf command = WorksheetSend Then
        report = "Process worksheet creation failed: " & Err.Description
    Else
        report = "Process worksheet retrieval failed: " & Err.Description
    End If
    Resume Done
End Function

Private Sub Retrieve(ByVal form As frmProduction, ByVal action As cProductionWorksheetAction, ByRef report As String)
    Dim tables As Collection, imports As New Collection, tableName As Variant
    Dim prior As cProductionWorksheetDraft, draft As cProductionWorksheetDraft
    Dim facts As cProductionLifecycleFacts, rejected As Boolean, completed As Boolean
    Dim succeeded As Long, failed As Long, failureDetails As String, deleteReport As String, validationReport As String
    Set tables = modProductionProcessWorksheet.FindSelectedProcessWorksheetTables(action.OperatorWorkbook, report)
    If tables Is Nothing Then action.OutcomeCode = "REJECTED": Exit Sub
    If tables.Count = 0 Then action.OutcomeCode = "REJECTED": Exit Sub
    Set prior = form.CaptureProcessWorksheetDraft()
    ' Preserve D15's complete validation pass before the first owning save.
    For Each tableName In tables
        If Not action.CanContinue(report) Then Exit Sub
        Set draft = New cProductionWorksheetDraft: draft.TableName = CStr(tableName)
        If Not ReadWorksheetDraft(action.OperatorWorkbook, draft, report, rejected) Then GoTo ValidationFailed
        If Not action.CanContinue(report) Then Exit Sub
        rejected = True
        If Not form.LoadProcessWorksheetDraft(draft, report) Then
            report = "Process worksheet retrieval failed: " & report
            GoTo ValidationFailed
        End If
        If Not form.ValidateProcessWorksheetDraft(validationReport) Then
            report = "Process worksheet retrieval failed: " & validationReport
            GoTo ValidationFailed
        End If
        imports.Add draft
    Next tableName

    For Each draft In imports
        If Not action.CanContinue(report) Then Exit Sub
        If Not form.LoadProcessWorksheetDraft(draft, report) Then
            failed = failed + 1: failureDetails = report
            GoTo NextImport
        End If
        If Not action.CanContinue(report) Then Exit Sub
        Set facts = New cProductionLifecycleFacts
        completed = form.SubmitProcessWorksheetDraft(draft, facts)
        action.ObserveSubmission facts
        If Not action.CanContinue(report) Then Exit Sub
        If completed Then
            If modProductionProcessWorksheet.DeleteProcessWorksheetTable(action.OperatorWorkbook, draft.TableName, deleteReport) Then
                succeeded = succeeded + 1
            Else
                failed = failed + 1: failureDetails = deleteReport
            End If
        Else
            failed = failed + 1: failureDetails = form.TestStatusText()
        End If
NextImport:
    Next draft
    If Not action.CanContinue(report) Then Exit Sub
    report = "Retrieved " & CStr(succeeded) & " selected Process table(s) as DRAFT"
    If failed > 0 Then report = report & "; " & CStr(failed) & " table(s) remain"
    report = report & "."
    If failureDetails <> "" Then report = report & " " & failureDetails
    If failed = 0 And succeeded = tables.Count Then action.OutcomeCode = "CONFIRMED"
    Exit Sub
ValidationFailed:
    If Not action.CanContinue(validationReport) Then report = validationReport: Exit Sub
    Call form.LoadProcessWorksheetDraft(prior, validationReport)
    If rejected Then action.OutcomeCode = "REJECTED"
    report = CStr(tableName) & ": " & report & " No selected table was saved or removed."
End Sub

Private Function ReadWorksheetDraft(ByVal operatorBook As Workbook, ByVal draft As cProductionWorksheetDraft, _
                                    ByRef report As String, ByRef rejected As Boolean) As Boolean
    Dim identity As String, version As String, name As String, description As String, payload As String
    ' Retain the service's output variables locally and assign the snapshot explicitly.
    If Not modProductionProcessWorksheet.ReadProcessDraftFromWorksheet(operatorBook, draft.TableName, _
            identity, version, name, description, payload, report, rejected) Then Exit Function
    draft.ProcessId = identity: draft.ProcessVersion = version
    draft.ProcessName = name: draft.Description = description: draft.Payload = payload
    ReadWorksheetDraft = True
End Function
