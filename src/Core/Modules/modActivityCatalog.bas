Attribute VB_Name = "modActivityCatalog"
Option Explicit
Option Private Module

Public Const CATALOG_VERSION As Long = 1

Public Function ControlIds() As Variant
    ControlIds = Array("ADMIN_SETTINGS_SAVE_VALUE", "PRODUCTION_UOM_RETRIEVE")
End Function

Public Function Control(ByVal controlId As String) As Object
    Dim record As Object
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", controlId
    record.Add "OwnerId", "CORE_CONFIGURATION"
    record.Add "Class", "Command"
    Select Case controlId
        Case "ADMIN_SETTINGS_SAVE_VALUE"
            record.Add "Role", "Admin"
            record.Add "Caption", "Save Value"
            record.Add "Surface", "Admin > Settings"
            record.Add "Capability", "ADMIN_MAINT"
            record.Add "CodePrefix", "CONFIG_SAVE_"
        Case "PRODUCTION_UOM_RETRIEVE"
            record.Add "Role", "Production"
            record.Add "Caption", "Retrieve UOM Catalog"
            record.Add "Surface", "Operations > Production > UOM Catalog"
            record.Add "Capability", "PROD_POST"
            record.Add "CodePrefix", "UOM_RETRIEVE_"
        Case Else: Exit Function
    End Select
    Set Control = record
End Function

Public Function Outcome(ByVal controlId As String, ByVal outcomeCode As String) As Object
    Dim record As Object, definition As Object, message As String
    Set definition = Control(controlId)
    If definition Is Nothing Then Exit Function
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", definition("CodePrefix") & outcomeCode
    record.Add "OutcomeCode", outcomeCode
    Select Case outcomeCode
        Case "REQUESTED"
            record.Add "Severity", "Info": record.Add "DataEffect", "Unknown"
            message = "Action requested."
        Case "COMPLETED"
            record.Add "Severity", "Info": record.Add "DataEffect", "Changed"
            message = "Configuration change saved."
        Case "UNCHANGED"
            record.Add "Severity", "Info": record.Add "DataEffect", "Unchanged"
            message = "Configuration already matches."
        Case "DENIED"
            record.Add "Severity", "Blocked": record.Add "DataEffect", "Unchanged"
            message = "The action was not authorized."
        Case "REJECTED"
            record.Add "Severity", "Warning": record.Add "DataEffect", "Unchanged"
            message = "Validation prevented the change."
        Case "FAILED"
            record.Add "Severity", "Error": record.Add "DataEffect", "Unknown"
            message = "The action failed. Its final data state requires verification."
        Case Else: Exit Function
    End Select
    record.Add "UserMessage", message
    If outcomeCode = "FAILED" Then
        record.Add "NextStep", "Inspect the owning workflow before retrying."
    ElseIf outcomeCode = "DENIED" Or outcomeCode = "REJECTED" Then
        record.Add "NextStep", "Review the owning form's prerequisites and permissions."
    Else
        record.Add "NextStep", ""
    End If
    Set Outcome = record
End Function
