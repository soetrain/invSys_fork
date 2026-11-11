Attribute VB_Name = "modInventoryActivityCodes"
Option Explicit
Option Private Module

Public Function SelectionIds() As Variant
    SelectionIds = Array("ADMIN_INVENTORY_SELECT_ITEM", "ADMIN_INVENTORY_SELECT_UOM", _
        "ADMIN_INVENTORY_SELECT_QUANTITY_MODE", "ADMIN_INVENTORY_SELECT_LOCATION", _
        "ADMIN_INVENTORY_SELECT_CATEGORY", "ADMIN_INVENTORY_SELECT_CUSTOM_FIELD", _
        "ADMIN_INVENTORY_CUSTOM_FIELD_ADD", "ADMIN_INVENTORY_CUSTOM_FIELD_REMOVE")
End Function

Public Function Control(ByVal id As String, Optional ByVal version As Long = 47) As Object
    Dim caption As String, record As Object, kind As String
    kind = "Navigation"
    Select Case id
        Case "ADMIN_INVENTORY_OPEN": caption = "Add/Edit Inventory"
        Case "ADMIN_INVENTORY_ADD_MODE": caption = "Add Item Mode"
        Case "ADMIN_INVENTORY_EDIT_MODE": caption = "Edit Item"
        Case "ADMIN_INVENTORY_CLOSE": caption = "Close"
        Case Else
            If version < 45 Then Exit Function
            Select Case id
                Case "ADMIN_INVENTORY_SELECT_ITEM": caption = "Select Inventory Item"
                Case "ADMIN_INVENTORY_SELECT_UOM": caption = "Select UOM"
                Case "ADMIN_INVENTORY_SELECT_QUANTITY_MODE": caption = "Select Quantity Mode"
                Case "ADMIN_INVENTORY_SELECT_LOCATION": caption = "Select Default Location"
                Case "ADMIN_INVENTORY_SELECT_CATEGORY": caption = "Select Category"
                Case "ADMIN_INVENTORY_SELECT_CUSTOM_FIELD": caption = "Select Additional Field"
                Case "ADMIN_INVENTORY_CUSTOM_FIELD_ADD": caption = "Add Additional Field": kind = "Command"
                Case "ADMIN_INVENTORY_CUSTOM_FIELD_REMOVE": caption = "Remove Additional Field": kind = "Command"
                Case "ADMIN_INVENTORY_UOM_ADD"
                    If version < 46 Then Exit Function
                    caption = "Add UOM": kind = "Command"
                Case "ADMIN_INVENTORY_CREATE_TABLE", "ADMIN_INVENTORY_UPLOAD_TABLE"
                    If version < 47 Then Exit Function
                    kind = "Command"
                    If id = "ADMIN_INVENTORY_CREATE_TABLE" Then caption = "Create Inventory Table" Else caption = "Upload Selected Inventory Table"
                Case Else: Exit Function
            End Select
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "ADMIN_INVENTORY"
    record.Add "Class", kind: record.Add "Role", "Admin": record.Add "Caption", caption
    record.Add "Surface", "Admin > Inventory": record.Add "Capability", "ADMIN_MAINT"
    record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function PositiveOutcome(ByVal id As String, ByVal code As String) As Boolean
    Select Case id
        Case "ADMIN_INVENTORY_CREATE_TABLE": PositiveOutcome = (code = "SAVED")
        Case "ADMIN_INVENTORY_UPLOAD_TABLE": PositiveOutcome = (code = "CONFIRMED")
        Case "ADMIN_INVENTORY_OPEN": PositiveOutcome = (code = "OPENED")
        Case "ADMIN_INVENTORY_ADD_MODE", "ADMIN_INVENTORY_EDIT_MODE": PositiveOutcome = (code = "SELECTED")
        Case "ADMIN_INVENTORY_CLOSE": PositiveOutcome = (code = "CLOSED")
        Case "ADMIN_INVENTORY_SELECT_ITEM", "ADMIN_INVENTORY_SELECT_UOM", _
             "ADMIN_INVENTORY_SELECT_QUANTITY_MODE", "ADMIN_INVENTORY_SELECT_LOCATION", _
             "ADMIN_INVENTORY_SELECT_CATEGORY", "ADMIN_INVENTORY_SELECT_CUSTOM_FIELD"
            PositiveOutcome = (code = "SELECTED")
        Case "ADMIN_INVENTORY_CUSTOM_FIELD_ADD": PositiveOutcome = (code = "STAGED")
        Case "ADMIN_INVENTORY_UOM_ADD": PositiveOutcome = (code = "STAGED")
        Case "ADMIN_INVENTORY_CUSTOM_FIELD_REMOVE": PositiveOutcome = (code = "STAGED" Or code = "UNCHANGED")
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, message As String, severity As String, effect As String
    If id = "ADMIN_INVENTORY_CREATE_TABLE" Or id = "ADMIN_INVENTORY_UPLOAD_TABLE" Then
        Set Outcome = modInventoryWorksheetActivity.Outcome(id, code): Exit Function
    End If
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = definition("Caption") & " requested."
        Case "FAILED": severity = "Error": message = "Inventory navigation did not finish. Inspect the owning form; no rollback is asserted."
        Case "CANCELLED", "REFUSED"
            If id <> "ADMIN_INVENTORY_UOM_ADD" Then Exit Function
            effect = "Unchanged"
            If code = "CANCELLED" Then
                message = "UOM entry was cancelled. No catalog or inventory save is asserted."
            Else
                severity = "Blocked"
                message = "UOM entry could not continue under the captured context. No catalog or inventory save is asserted."
            End If
        Case "REJECTED"
            If id <> "ADMIN_INVENTORY_CUSTOM_FIELD_ADD" Then Exit Function
            severity = "Warning": effect = "Unchanged"
            message = "The additional field was not staged. Read the form result. No inventory change is asserted."
        Case Else
            If Not PositiveOutcome(id, code) Then Exit Function
            effect = "Unchanged"
            Select Case code
                Case "OPENED": message = "Add/Edit Inventory opened. No inventory change is asserted."
                Case "SELECTED": message = definition("Caption") & " selected. No inventory change is asserted."
                Case "CLOSED": message = "Add/Edit Inventory closed. No inventory change or rollback is asserted."
                Case "STAGED": message = "The additional-field draft was updated. No inventory change is asserted."
                Case "UNCHANGED": message = "No additional field was selected to remove. The draft is unchanged."
            End Select
            If id = "ADMIN_INVENTORY_UOM_ADD" Then message = "The UOM was selected in the local draft. No catalog or inventory save is asserted."
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
