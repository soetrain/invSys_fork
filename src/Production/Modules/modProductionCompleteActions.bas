Attribute VB_Name = "modProductionCompleteActions"
Option Explicit
Option Private Module

' Keep one operator completion active across the owner's yielding work.
Public Sub Execute(ByVal owner As frmProduction, ByVal context As String, _
                   ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean)
    Dim priorLoading As Boolean, priorBusy As Boolean
    Dim action As cProductionWorksheetAction, report As String
    Dim number As Long, source As String, description As String, helpFile As String, helpContext As Long
    If loading Or busy Then Exit Sub
    priorLoading = loading: priorBusy = busy
    On Error GoTo Failed
    busy = True
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_RUN_COMPLETE", context, operatorBook, report) Then
        If action.OutcomeCode = "DENIED" Then report = "Production permission changed. Reopen Production before continuing."
        GoTo Done
    End If
    Set owner.RunActionContinuation = action
    owner.CompleteProductionRun action
    report = owner.RunAllocationStatus()
Done:
    Set owner.RunActionContinuation = Nothing
    loading = priorLoading
    If Not action Is Nothing Then action.Finish report
    busy = priorBusy
    If report <> "" Then owner.ShowStatus report
    If number <> 0 Then
        On Error GoTo 0
        Err.Raise number, source, description, helpFile, helpContext
    End If
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    helpFile = Err.HelpFile: helpContext = Err.HelpContext
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    Resume Done
End Sub

' Same-project adapter: writer facts determine both outcome and source references.
Public Function QueueInventory(ByVal eventType As String, ByVal payload As String, ByVal note As String, _
                               ByRef eventId As String, ByRef report As String, _
                               ByVal action As cProductionWorksheetAction) As Boolean
    Dim attempted As Boolean, submitted As Boolean
    submitted = modRoleEventWriter.QueuePayloadEventCurrent(eventType, "", payload, note, eventId, report, "", attempted)
    If Not action Is Nothing Then action.ObserveInventorySubmission eventId, submitted, attempted
    QueueInventory = submitted
End Function

' Completion owners retain only event identities acknowledged by their writer.
Public Sub AppendEventId(ByRef eventIds As String, ByVal eventId As String)
    If eventIds <> "" Then eventIds = eventIds & ","
    eventIds = eventIds & eventId
End Sub

Public Sub AppendProcessorReport(ByRef reports As String, ByVal report As String)
    If Trim$(report) = "" Then Exit Sub
    If reports <> "" Then reports = reports & " | "
    reports = reports & report
End Sub
