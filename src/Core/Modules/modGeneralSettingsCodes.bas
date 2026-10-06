Attribute VB_Name = "modGeneralSettingsCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Split("ADMIN_SETTINGS_RELOAD|ADMIN_SETTINGS_SELECT_CONFIG|ADMIN_CARRIER_ADD|" & _
        "ADMIN_CARRIER_REMOVE|ADMIN_CARRIER_RESET|ADMIN_CARRIER_SELECT|ADMIN_UOM_SELECT|" & _
        "ADMIN_CONNECTION_SELECT|ADMIN_CONNECTION_SAVE", "|")
End Function

Public Function Control(ByVal id As String) As Object
    Dim record As Object, caption As String, owner As String, kind As String, section As String
    owner = "ADMIN_SETTINGS_UI": kind = "Navigation"
    Select Case id
        Case "ADMIN_SETTINGS_RELOAD": caption = "Reload": owner = "CORE_CONFIGURATION": kind = "Command"
        Case "ADMIN_SETTINGS_SELECT_CONFIG": caption = "Selected config key"
        Case "ADMIN_CARRIER_ADD": caption = "Add": owner = "CORE_LOCAL_SETTINGS": kind = "Command"
        Case "ADMIN_CARRIER_REMOVE": caption = "Remove": owner = "CORE_LOCAL_SETTINGS": kind = "Command"
        Case "ADMIN_CARRIER_RESET": caption = "Reset": owner = "CORE_LOCAL_SETTINGS": kind = "Command"
        Case "ADMIN_CARRIER_SELECT": caption = "Carrier"
        Case "ADMIN_UOM_SELECT": caption = "UOM"
        Case "ADMIN_CONNECTION_SELECT": caption = "Connection option"
        Case "ADMIN_CONNECTION_SAVE": caption = "Save Connection Option": owner = "CORE_LOCAL_SETTINGS": kind = "Command"
        Case Else: Exit Function
    End Select
    If Left$(id, 14) = "ADMIN_CARRIER_" Then section = " > Shipping Carriers"
    If id = "ADMIN_UOM_SELECT" Then section = " > Recipe UOM Catalog"
    If Left$(id, 17) = "ADMIN_CONNECTION_" Then section = " > Server Connection"
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", owner: record.Add "Class", kind
    record.Add "Role", "Admin": record.Add "Caption", caption
    record.Add "Surface", "Admin > Settings > General" & section
    record.Add "Capability", IIf(kind = "Command", "ADMIN_MAINT", "")
    record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function PositiveOutcome(ByVal id As String, ByVal code As String) As Boolean
    Select Case id
        Case "ADMIN_CARRIER_ADD", "ADMIN_CARRIER_REMOVE", "ADMIN_CARRIER_RESET", "ADMIN_CONNECTION_SAVE"
            PositiveOutcome = (code = "COMPLETED" Or code = "UNCHANGED")
        Case "ADMIN_SETTINGS_RELOAD": PositiveOutcome = (code = "REFRESHED")
        Case "ADMIN_SETTINGS_SELECT_CONFIG", "ADMIN_CARRIER_SELECT", "ADMIN_UOM_SELECT"
            PositiveOutcome = (code = "SELECTED")
        Case "ADMIN_CONNECTION_SELECT": PositiveOutcome = (code = "STAGED")
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, severity As String, effect As String, message As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED": effect = "Unknown": message = "General Settings action requested."
        Case "DENIED": severity = "Blocked": message = "The Settings command was not authorized."
        Case "REJECTED": severity = "Warning": message = "Settings validation prevented the action."
        Case "CANCELLED"
            If id <> "ADMIN_CARRIER_RESET" Then Exit Function
            severity = "Notice": message = "Carrier reset was cancelled."
        Case "FAILED"
            severity = "Error": message = "Settings read or staging failed. Saved settings were not changed."
            If definition("OwnerId") = "CORE_LOCAL_SETTINGS" Then
                effect = "Unknown": message = "Local Settings save could not be verified. Reopen Settings to verify."
            End If
        Case "COMPLETED", "UNCHANGED", "REFRESHED", "SELECTED", "STAGED"
            If Not PositiveOutcome(id, code) Then Exit Function
            Select Case code
                Case "COMPLETED": effect = "Changed": message = "Settings saved for this Windows user; warehouse inventory was not changed."
                Case "UNCHANGED": message = "Settings already match for this Windows user."
                Case "REFRESHED": message = "Canonical config reloaded. Saved settings were not changed."
                Case "SELECTED": message = "Settings display selection changed. Saved settings were not changed."
                Case "STAGED": message = "Connection option staged; not saved."
            End Select
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
