Attribute VB_Name = "modEvaluationJournal"
Option Explicit
Option Private Module

' Validate saved references against the selected immutable journal, without
' consulting a newer publication or recalculating a previous conclusion.
Public Function Matches(ByVal result As Object, ByVal header As Object, ByVal records As Collection) As Boolean
    Dim observations As Object, actions As Object, record As Object, entry As Object, terminal As Object
    Dim key As String, id As Variant, reference As Object, index As Long, field As Variant
    On Error GoTo Invalid
    If header("RecordId") <> result("JournalRecordId") Or header("ContentSha256") <> result("JournalSha256") Then Exit Function
    If header("SequenceId") <> result("SequenceId") Or header("CreatedByUserId") <> result("RecordedByUserId") Then Exit Function
    If header("PolicyVersion") <> result("CapturePolicyVersion") Or header("OriginWarehouseId") <> result("OriginWarehouseId") Then Exit Function
    Set observations = CreateObject("Scripting.Dictionary"): Set actions = CreateObject("Scripting.Dictionary")
    For Each record In records
        observations.Add CStr(record("ActivityId")) & vbTab & CStr(record("OutcomeCode")), record
        If record("OutcomeCode") = "REQUESTED" Then actions.Add CStr(record("ActivityId")), True
    Next record
    For Each entry In result("Matches")
        key = CStr(entry("ActivityId")) & vbTab & CStr(entry("OutcomeCode"))
        If Not observations.Exists(key) Then Exit Function
        Set record = observations(key)
        If record("Ordinal") <> entry("Ordinal") Or record("ControlId") <> entry("ControlId") Then Exit Function
        If entry("StepId") = result("ExpectedConclusion")("TerminalStepId") Then Set terminal = record
    Next entry
    For Each id In result("ExtraActivityIds")
        If Not actions.Exists(id) Then Exit Function
    Next id
    If terminal Is Nothing Or result("ExpectedConclusion")("TerminalKind") <> "SourceEventsApplied" Then
        If result("TerminalSources").Count <> 0 Then Exit Function
    Else
        If result("TerminalSources").Count <> terminal("SourceEventRefs").Count Then Exit Function
        For Each entry In result("TerminalSources")
            index = index + 1: Set reference = terminal("SourceEventRefs")(index)
            For Each field In Array("WarehouseId", "SourceKind", "EventId", "SubmissionState")
                If entry(field) <> reference(field) Then Exit Function
            Next field
        Next entry
    End If
    Matches = True
Invalid:
End Function
