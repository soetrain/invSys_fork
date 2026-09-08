Attribute VB_Name = "modActivityCatalog"
Option Explicit
Option Private Module

Public Const CATALOG_VERSION As Long = 4

Public Function ControlIds(Optional ByVal version As Long = CATALOG_VERSION) As Variant
    If version = 1 Then
        ControlIds = Array("ADMIN_SETTINGS_SAVE_VALUE", "PRODUCTION_UOM_RETRIEVE")
    ElseIf version = 2 Then
        ControlIds = Array("ADMIN_SETTINGS_SAVE_VALUE", "PRODUCTION_UOM_RETRIEVE", "RECEIVING_CONFIRM_WRITES")
    ElseIf version = 3 Then
        ControlIds = Array("ADMIN_SETTINGS_SAVE_VALUE", "PRODUCTION_UOM_RETRIEVE", "RECEIVING_CONFIRM_WRITES", _
                           "RECEIVING_ADD_SELECTED", "DISPOSITION_ADD_SELECTED", "DISPOSITION_CONFIRM")
    ElseIf version = 4 Then
        ControlIds = Array("ADMIN_SETTINGS_SAVE_VALUE", "PRODUCTION_UOM_RETRIEVE", "RECEIVING_CONFIRM_WRITES", _
                           "RECEIVING_ADD_SELECTED", "DISPOSITION_ADD_SELECTED", "DISPOSITION_CONFIRM", _
                           "RECEIVING_REFRESH", "RECEIVING_CLEAR")
    End If
End Function

Public Function Control(ByVal controlId As String, Optional ByVal version As Long = CATALOG_VERSION) As Object
    Dim record As Object
    If version < 1 Or version > CATALOG_VERSION Then Exit Function
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
        Case "RECEIVING_CONFIRM_WRITES"
            If version < 2 Then Exit Function
            record("OwnerId") = "RECEIVING_WORKFLOW"
            record.Add "Role", "Receiving"
            record.Add "Caption", "Confirm Writes"
            record.Add "Surface", "Operations > Receiving"
            record.Add "Capability", "RECEIVE_POST"
            record.Add "CodePrefix", "RECEIVE_CONFIRM_"
        Case "RECEIVING_ADD_SELECTED", "DISPOSITION_ADD_SELECTED", "DISPOSITION_CONFIRM"
            If version < 3 Then Exit Function
            record("OwnerId") = "RECEIVING_DISPOSITION"
            record.Add "Role", "Receiving"
            record.Add "Surface", "Operations > Receiving > Returns"
            record.Add "Capability", "RECEIVE_POST"
            If controlId = "RECEIVING_ADD_SELECTED" Then
                record("OwnerId") = "RECEIVING_STAGING"
                record("Surface") = "Operations > Receiving"
                record.Add "Caption", "Add Selected"
                record.Add "CodePrefix", "RECEIVE_ADD_"
            ElseIf controlId = "DISPOSITION_ADD_SELECTED" Then
                record.Add "Caption", "Add Disposition"
                record.Add "CodePrefix", "DISPOSITION_ADD_"
            Else
                record.Add "Caption", "Confirm Dispositions"
                record.Add "CodePrefix", "DISPOSITION_CONFIRM_"
            End If
        Case "RECEIVING_REFRESH", "RECEIVING_CLEAR"
            If version < 4 Then Exit Function
            record("OwnerId") = "RECEIVING_WORKFLOW"
            record.Add "Role", "Receiving"
            record.Add "Surface", "Operations > Receiving"
            record.Add "Capability", "RECEIVE_POST"
            If controlId = "RECEIVING_CLEAR" Then
                record("OwnerId") = "RECEIVING_STAGING"
                record.Add "Caption", "Clear"
                record.Add "CodePrefix", "RECEIVE_CLEAR_"
            Else
                record.Add "Caption", "Refresh"
                record.Add "CodePrefix", "RECEIVE_REFRESH_"
            End If
        Case Else: Exit Function
    End Select
    Set Control = record
End Function

Public Function Outcome(ByVal controlId As String, ByVal outcomeCode As String) As Object
    Dim record As Object, definition As Object, message As String
    Set definition = Control(controlId)
    If definition Is Nothing Then Exit Function
    If definition("Role") = "Receiving" Then
        If controlId = "RECEIVING_ADD_SELECTED" Or controlId = "DISPOSITION_ADD_SELECTED" Then
            Set Outcome = modReceivingActivityCodes.StagingOutcome(outcomeCode, definition("CodePrefix"))
        ElseIf controlId = "RECEIVING_REFRESH" Or controlId = "RECEIVING_CLEAR" Then
            Set Outcome = modReceivingActivityCodes.LocalOutcome(outcomeCode, controlId = "RECEIVING_CLEAR")
        Else
            Set Outcome = modReceivingActivityCodes.Outcome(outcomeCode, controlId = "DISPOSITION_CONFIRM")
        End If
        Exit Function
    End If
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
