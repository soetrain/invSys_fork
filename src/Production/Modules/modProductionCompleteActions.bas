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
    modProductionOperatorBinding.ObserveOwner owner, action.RecordedActivityId
    owner.CompleteProductionRun action
    ' Native workbook closure may already have unloaded this exact form.
    If modOperationsFormLifetime.IsLoaded(owner) Then report = owner.RunAllocationStatus()
Done:
    If modOperationsFormLifetime.IsLoaded(owner) Then Set owner.RunActionContinuation = Nothing
    loading = priorLoading
    If Not action Is Nothing Then
        action.Finish report
        ' Release the claim's call-lifetime guard even when the form misses its notification.
        modProductionOperatorBinding.OwnerExited action.RecordedActivityId
        modProductionOperatorBinding.Returned action.RecordedActivityId
    End If
    busy = priorBusy
    If report <> "" And modOperationsFormLifetime.IsLoaded(owner) Then owner.ShowStatus report
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

Public Function FindOutputRow(ByVal owner As frmProduction, ByVal lo As ListObject, _
                                              ByVal rowVal As String, _
                                              ByVal procVal As String, _
                                              ByVal outputVal As String, _
                                              Optional ByVal fallbackRow As Long = 0) As Long
    Dim r As Long
    Dim cRow As Long
    Dim cProc As Long
    Dim cOutput As Long
    Dim wantedRow As String

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cRow = owner.ProductionColumnIndex(lo, "System_Key")
    cProc = owner.ProductionColumnIndex(lo, "PROCESS")
    cOutput = owner.ProductionColumnIndex(lo, "OUTPUT")
    wantedRow = rowVal

    If wantedRow <> "" Then
        If cRow = 0 Then Exit Function
        For r = 1 To lo.ListRows.Count
            If StrComp(owner.NzStr(lo.DataBodyRange.Cells(r, cRow).Value2), wantedRow, vbBinaryCompare) = 0 Then
                If FindOutputRow <> 0 Then FindOutputRow = 0: Exit Function
                FindOutputRow = r
            End If
        Next r
        Exit Function
    End If

    If cProc > 0 And cOutput > 0 Then
        For r = 1 To lo.ListRows.Count
            If StrComp(Trim$(owner.NzStr(lo.DataBodyRange.Cells(r, cProc).Value)), Trim$(procVal), vbTextCompare) = 0 _
               And StrComp(Trim$(owner.NzStr(lo.DataBodyRange.Cells(r, cOutput).Value)), Trim$(outputVal), vbTextCompare) = 0 Then
                FindOutputRow = r
                Exit Function
            End If
        Next r
    End If

    If fallbackRow >= 1 And fallbackRow <= lo.ListRows.Count Then
        FindOutputRow = fallbackRow
    End If
End Function
