Attribute VB_Name = "modProductionUomCodes"
Option Explicit
Option Private Module

' Catalog 15: local workbench use is distinct from publishing catalog changes.
Public Function Control(ByVal id As String) As Object
    Dim record As Object
    If id <> "PRODUCTION_UOM_EDIT" Then Exit Function
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_UOM_STAGING"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", "Edit UOM Catalog on Sheet"
    record.Add "Surface", "Operations > Production > Production Settings > UOM Catalog"
    record.Add "Capability", "PROD_POST": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal code As String) As Object
    Dim record As Object, severity As String, effect As String, message As String
    Dim nextStep As String
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED": effect = "Unknown": message = "UOM workbench opening requested."
        Case "OPENED"
            message = "A UOM workbench was created locally; the saved catalog was not changed."
            nextStep = "Review the staging sheet; use the separate Retrieve command to request publication."
        Case "REUSED"
            message = "The existing UOM draft was reopened with its edits retained; the saved catalog was not reloaded."
            nextStep = "Review existing edits before requesting publication through Retrieve UOM Catalog."
        Case "REJECTED"
            severity = "Warning": message = "UOM staging validation prevented opening; no cells were changed."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; staging was not changed."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The UOM workbench could not be opened; verify the staging worksheet before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", "PRODUCTION_UOM_EDIT_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
