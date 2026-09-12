Attribute VB_Name = "modActivityCatalog"
Option Explicit
Option Private Module

Public Const CATALOG_VERSION As Long = 7

Public Function ControlIds(Optional ByVal version As Long = CATALOG_VERSION) As Variant
    Dim ids As Variant
    If version = 7 Then
        ids = ControlIds(6)
        ReDim Preserve ids(LBound(ids) To UBound(ids) + 1)
        ids(UBound(ids)) = "RECEIVING_WORKSHEET_CONFIRM"
        ControlIds = ids
        Exit Function
    End If
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
    ElseIf version = 5 Then
        ControlIds = Array("ADMIN_SETTINGS_SAVE_VALUE", "PRODUCTION_UOM_RETRIEVE", "RECEIVING_CONFIRM_WRITES", _
                           "RECEIVING_ADD_SELECTED", "DISPOSITION_ADD_SELECTED", "DISPOSITION_CONFIRM", _
                           "RECEIVING_REFRESH", "RECEIVING_CLEAR", "RECEIVING_OPEN", "RECEIVING_CLOSE")
    ElseIf version = 6 Then
        ControlIds = Array("ADMIN_SETTINGS_SAVE_VALUE", "PRODUCTION_UOM_RETRIEVE", "RECEIVING_CONFIRM_WRITES", _
                           "RECEIVING_ADD_SELECTED", "DISPOSITION_ADD_SELECTED", "DISPOSITION_CONFIRM", _
                           "RECEIVING_REFRESH", "RECEIVING_CLEAR", "RECEIVING_OPEN", "RECEIVING_CLOSE", _
                           "RECEIVING_PAGE_RECEIPTS", "RECEIVING_PAGE_RETURNS", "RECEIVING_PAGE_PURCHASING", _
                           "RECEIVING_SELECT_ITEM", "DISPOSITION_SELECT_ITEM", _
                           "RECEIVING_SELECT_AGGREGATE", "DISPOSITION_SELECT_AGGREGATE", _
                           "RECEIVING_SELECT_HISTORY", "DISPOSITION_SELECT_HISTORY", _
                           "RECEIVING_SELECT_STAGED", "DISPOSITION_SELECT_STAGED", _
                           "RECEIVING_SELECT_CONDITION", "DISPOSITION_SELECT_KIND")
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
        Case "RECEIVING_WORKSHEET_CONFIRM"
            If version < 7 Then Exit Function
            record("OwnerId") = "RECEIVING_WORKFLOW"
            record.Add "Role", "Receiving"
            record.Add "Caption", "Confirm Writes"
            record.Add "Surface", "Operations > Receiving > Received Tally"
            record.Add "Capability", "RECEIVE_POST"
            record.Add "CodePrefix", "RECEIVE_WORKSHEET_CONFIRM_"
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
        Case "RECEIVING_OPEN", "RECEIVING_CLOSE"
            If version < 5 Then Exit Function
            record("OwnerId") = "RECEIVING_WORKFLOW"
            record.Add "Role", "Receiving"
            record.Add "Surface", "Operations > Receiving"
            record.Add "Capability", "RECEIVE_POST"
            If controlId = "RECEIVING_OPEN" Then
                record.Add "Caption", "Receiving"
                record.Add "CodePrefix", "RECEIVE_OPEN_"
            Else
                record.Add "Caption", "Close"
                record.Add "CodePrefix", "RECEIVE_CLOSE_"
            End If
        Case Else
            If version >= 6 Then Set Control = modReceivingNavigationCodes.Control(controlId)
            Exit Function
    End Select
    Set Control = record
End Function

Public Function Outcome(ByVal controlId As String, ByVal outcomeCode As String) As Object
    Dim record As Object, definition As Object, message As String
    Set definition = Control(controlId)
    If definition Is Nothing Then Exit Function
    If definition("Class") = "Navigation" Then
        Set Outcome = modReceivingNavigationCodes.Outcome(controlId, outcomeCode)
        Exit Function
    End If
    If definition("Role") = "Receiving" Then
        If controlId = "RECEIVING_ADD_SELECTED" Or controlId = "DISPOSITION_ADD_SELECTED" Then
            Set Outcome = modReceivingActivityCodes.StagingOutcome(outcomeCode, definition("CodePrefix"))
        ElseIf controlId = "RECEIVING_REFRESH" Or controlId = "RECEIVING_CLEAR" Then
            Set Outcome = modReceivingActivityCodes.LocalOutcome(outcomeCode, controlId = "RECEIVING_CLEAR")
        ElseIf controlId = "RECEIVING_OPEN" Or controlId = "RECEIVING_CLOSE" Then
            Set Outcome = modReceivingActivityCodes.LifecycleOutcome(outcomeCode, controlId = "RECEIVING_CLOSE")
        Else
            Set record = modReceivingActivityCodes.Outcome(outcomeCode, controlId = "DISPOSITION_CONFIRM")
            If controlId = "RECEIVING_WORKSHEET_CONFIRM" And Not record Is Nothing Then
                record("EventCode") = definition("CodePrefix") & outcomeCode
                record("UserMessage") = "Worksheet Confirm Writes: " & record("UserMessage")
                record("NextStep") = Replace(record("NextStep"), "owning form", "Receiving workflow")
            End If
            Set Outcome = record
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
