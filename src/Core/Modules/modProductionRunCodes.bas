Attribute VB_Name = "modProductionRunCodes"
Option Explicit
Option Private Module

' D18 catalogs24-26: local Run preparation, allocation and presentation.
Public Function ControlIds(Optional ByVal version As Long = 26) As Variant
    Dim ids As Variant
    ids = Array("PRODUCTION_RUN_SCALE", "PRODUCTION_RUN_CLEAR", "PRODUCTION_RUN_LOAD", _
        "PRODUCTION_RUN_LOADER_REFRESH", "PRODUCTION_RUN_MANAGER_REFRESH", "PRODUCTION_RUN_ALLOCATE", _
        "PRODUCTION_RUN_TREE_ALLOCATE", "PRODUCTION_RUN_TREE_EXPAND", "PRODUCTION_RUN_TREE_COLLAPSE")
    If version >= 25 Then
        ReDim Preserve ids(0 To 9)
        ids(9) = "PRODUCTION_RUN_CHECK_IN"
    End If
    If version >= 26 Then
        ReDim Preserve ids(0 To 10)
        ids(10) = "PRODUCTION_RUN_NEXT_BATCH"
    End If
    ControlIds = ids
End Function

Public Function Control(ByVal id As String, Optional ByVal version As Long = 26) As Object
    Dim ids As Variant, captions As Variant, index As Long, record As Object, surface As String
    ids = ControlIds(version)
    captions = Array("Apply Scale", "Clear Run", "Load Recipe", "Refresh", "Refresh", "Apply", _
                     "Apply", "Expand", "Collapse", "Check In", "Next Batch")
    For index = LBound(ids) To UBound(ids)
        If id = ids(index) Then
            surface = "Operations > Production > Production Run - List"
            If index >= 6 And index <= 8 Then surface = "Operations > Production > Production Run - Tree"
            Set record = modProductionControlCatalog.Command(id, "PRODUCTION_RUN_LOCAL", CStr(captions(index)), surface)
            If id = "PRODUCTION_RUN_TREE_EXPAND" Or id = "PRODUCTION_RUN_TREE_COLLAPSE" Then record("Class") = "Navigation"
            Set Control = record
            Exit Function
        End If
    Next index
End Function

Public Function PositiveOutcome(ByVal id As String) As String
    Select Case id
        Case "PRODUCTION_RUN_SCALE", "PRODUCTION_RUN_CLEAR", "PRODUCTION_RUN_LOAD", _
             "PRODUCTION_RUN_ALLOCATE", "PRODUCTION_RUN_TREE_ALLOCATE", _
             "PRODUCTION_RUN_CHECK_IN", "PRODUCTION_RUN_NEXT_BATCH": PositiveOutcome = "STAGED"
        Case "PRODUCTION_RUN_LOADER_REFRESH", "PRODUCTION_RUN_MANAGER_REFRESH": PositiveOutcome = "REFRESHED"
        Case "PRODUCTION_RUN_TREE_EXPAND", "PRODUCTION_RUN_TREE_COLLAPSE": PositiveOutcome = "PRESENTED"
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim positive As String, severity As String, effect As String, message As String, nextStep As String
    Dim record As Object
    positive = PositiveOutcome(id)
    If positive = "" Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED"
            effect = "Unknown": message = "Production Run action requested."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; no Run owner work was performed."
        Case "REJECTED"
            severity = "Warning": message = "Selection or validation prevented the Run action; saved authority was not changed."
            nextStep = "Review the form's validation result; mirrored inputs or earlier local allocation changes may remain."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The Run action failed or could not finish; earlier local preparation may remain."
            nextStep = "Inspect the local run and required sources before retrying; no rollback or inventory application is established."
        Case "STAGED", "REFRESHED", "PRESENTED"
            If code <> positive Then Exit Function
            Select Case code
                Case "STAGED": message = "Run preparation or allocation finished locally; saved inventory and definitions were not changed."
                Case "REFRESHED": message = "The local Run view was refreshed; source availability, freshness and completeness were not established."
                Case "PRESENTED": message = "Run choices were presented locally; inventory and allocation were not certified."
            End Select
        Case Else: Exit Function
    End Select
    If id = "PRODUCTION_RUN_CHECK_IN" And code = "STAGED" Then _
        message = "Check In finished local validation and staging; inventory was not reserved, consumed or submitted."
    If id = "PRODUCTION_RUN_NEXT_BATCH" And code = "STAGED" Then _
        message = "Next Batch finished local preparation; saved inventory and definitions were not changed."
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
