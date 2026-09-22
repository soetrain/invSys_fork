Attribute VB_Name = "modSettingsActivityCodes"
Option Explicit
Option Private Module

' D18 catalog 11: fixed vocabulary only; no staged or selected values.
Public Function ControlIds() As Variant
    ControlIds = Split("ADMIN_TRACKING_SELECT_CONTROL|ADMIN_TRACKING_CAPTURE|ADMIN_TRACKING_ADMIN_VISIBLE|" & _
        "ADMIN_TRACKING_DEFAULT_VIEW|ADMIN_TRACKING_COLLECT|ADMIN_TRACKING_VISIBLE|ADMIN_TRACKING_SEQUENCE|" & _
        "ADMIN_TRACKING_SAVE|ADMIN_TRACKING_RESET|ADMIN_TRACKING_RELOAD|" & _
        "ADMIN_DETAIL_SELECT_FAMILY|ADMIN_DETAIL_SELECT_FIELD|ADMIN_DETAIL_SHOW_FIELD|ADMIN_DETAIL_MOVE_UP|" & _
        "ADMIN_DETAIL_MOVE_DOWN|ADMIN_DETAIL_SAVE|ADMIN_DETAIL_RESET|ADMIN_DETAIL_RELOAD|" & _
        "ADMIN_PATH_PREFERENCE_SELECT|ADMIN_PATH_PREFERENCE_SAVE|ADMIN_PATH_PREFERENCE_RESET|ADMIN_PATH_PREFERENCE_RELOAD|" & _
        "VIEWER_PATH_PREFERENCE_SELECT|VIEWER_PATH_PREFERENCE_SAVE|VIEWER_PATH_PREFERENCE_RESET|VIEWER_PATH_PREFERENCE_RELOAD", "|")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, section As String, kind As String, owner As String, role As String, surface As String
    Dim record As Object
    kind = "Command": owner = "ADMIN_SETTINGS_UI": role = "Admin"
    Select Case id
        Case "ADMIN_TRACKING_SELECT_CONTROL": caption = "Family / control": kind = "Navigation"
        Case "ADMIN_TRACKING_CAPTURE": caption = "Capture recorded controls": kind = "Navigation"
        Case "ADMIN_TRACKING_ADMIN_VISIBLE": caption = "Show optional Admin events": kind = "Navigation"
        Case "ADMIN_TRACKING_DEFAULT_VIEW": caption = "Warehouse default view": kind = "Navigation"
        Case "ADMIN_TRACKING_COLLECT": caption = "Collect selected control": kind = "Navigation"
        Case "ADMIN_TRACKING_VISIBLE": caption = "Visible in Viewer": kind = "Navigation"
        Case "ADMIN_TRACKING_SEQUENCE": caption = "Eligible for recorded sequence": kind = "Navigation"
        Case "ADMIN_TRACKING_SAVE": caption = "Save Tracking Policy": owner = "CORE_CONFIGURATION"
        Case "ADMIN_DETAIL_SELECT_FAMILY": caption = "Event family": kind = "Navigation"
        Case "ADMIN_DETAIL_SELECT_FIELD": caption = "Field": kind = "Navigation"
        Case "ADMIN_DETAIL_SHOW_FIELD": caption = "Show selected field": kind = "Navigation"
        Case "ADMIN_DETAIL_MOVE_UP": caption = "Move Up"
        Case "ADMIN_DETAIL_MOVE_DOWN": caption = "Move Down"
        Case "ADMIN_DETAIL_SAVE": caption = "Save Detail Profile": owner = "CORE_CONFIGURATION"
        Case "ADMIN_PATH_PREFERENCE_SELECT", "VIEWER_PATH_PREFERENCE_SELECT": caption = "Preferred Action Path view": kind = "Navigation"
        Case "ADMIN_PATH_PREFERENCE_SAVE", "VIEWER_PATH_PREFERENCE_SAVE": caption = "Save My Preference"
        Case "ADMIN_TRACKING_RESET", "ADMIN_DETAIL_RESET", "ADMIN_PATH_PREFERENCE_RESET", "VIEWER_PATH_PREFERENCE_RESET": caption = "Reset to Default"
        Case "ADMIN_TRACKING_RELOAD", "ADMIN_DETAIL_RELOAD", "ADMIN_PATH_PREFERENCE_RELOAD", "VIEWER_PATH_PREFERENCE_RELOAD": caption = "Reload"
        Case Else: Exit Function
    End Select
    If PersonalControl(id) Then
        owner = "CORE_PERSONAL_PREFERENCE": section = "Action Paths"
    ElseIf Left$(id, 15) = "ADMIN_TRACKING_" Then
        section = "Tracking"
    Else
        section = "Event Detail"
    End If
    surface = "Admin > Settings > Event Tracking > " & section
    If Left$(id, 7) = "VIEWER_" Then role = "Viewer": surface = "Operations > Viewer > Settings > Event Tracking"
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", owner: record.Add "Class", kind
    record.Add "Role", role: record.Add "Caption", caption: record.Add "Surface", surface
    record.Add "Capability", IIf(owner = "CORE_CONFIGURATION", "ADMIN_MAINT", "")
    record.Add "CodePrefix", id & "_"
    If PersonalControl(id) Then record.Add "AuthorityMode", "SIGNED_IN_CONTEXT"
    Set Control = record
End Function

Public Function PersonalControl(ByVal id As String) As Boolean
    Select Case id
        Case "ADMIN_PATH_PREFERENCE_SELECT", "ADMIN_PATH_PREFERENCE_SAVE", "ADMIN_PATH_PREFERENCE_RESET", "ADMIN_PATH_PREFERENCE_RELOAD", _
             "VIEWER_PATH_PREFERENCE_SELECT", "VIEWER_PATH_PREFERENCE_SAVE", "VIEWER_PATH_PREFERENCE_RESET", "VIEWER_PATH_PREFERENCE_RELOAD"
            PersonalControl = True
    End Select
End Function

Private Function SuccessCodes(ByVal id As String) As String
    Select Case id
        Case "ADMIN_TRACKING_SAVE": SuccessCodes = "|"
        Case "ADMIN_DETAIL_SAVE": SuccessCodes = "|COMPLETED|"
        Case "ADMIN_PATH_PREFERENCE_SAVE", "VIEWER_PATH_PREFERENCE_SAVE": SuccessCodes = "|COMPLETED|UNCHANGED|"
        Case "ADMIN_TRACKING_RELOAD", "ADMIN_DETAIL_RELOAD", "ADMIN_PATH_PREFERENCE_RELOAD", "VIEWER_PATH_PREFERENCE_RELOAD": SuccessCodes = "|REFRESHED|"
        Case "ADMIN_TRACKING_SELECT_CONTROL", "ADMIN_DETAIL_SELECT_FAMILY", "ADMIN_DETAIL_SELECT_FIELD": SuccessCodes = "|SELECTED|"
        Case "ADMIN_TRACKING_CAPTURE", "ADMIN_TRACKING_ADMIN_VISIBLE", "ADMIN_TRACKING_DEFAULT_VIEW", _
             "ADMIN_TRACKING_COLLECT", "ADMIN_TRACKING_VISIBLE", "ADMIN_TRACKING_SEQUENCE", "ADMIN_TRACKING_RESET", _
             "ADMIN_DETAIL_SHOW_FIELD", "ADMIN_DETAIL_MOVE_UP", "ADMIN_DETAIL_MOVE_DOWN", "ADMIN_DETAIL_RESET", _
             "ADMIN_PATH_PREFERENCE_SELECT", "ADMIN_PATH_PREFERENCE_RESET", "VIEWER_PATH_PREFERENCE_SELECT", "VIEWER_PATH_PREFERENCE_RESET"
            SuccessCodes = "|STAGED|"
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, severity As String, effect As String, message As String, nextStep As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED": effect = "Unknown": message = "Settings action requested."
        Case "DENIED": severity = "Blocked": message = "The Settings action was not authorized."
        Case "REJECTED": severity = "Warning": message = "Settings validation prevented the action."
        Case "FAILED"
            severity = "Error"
            Select Case id
                Case "ADMIN_TRACKING_SAVE", "ADMIN_DETAIL_SAVE", "ADMIN_PATH_PREFERENCE_SAVE", "VIEWER_PATH_PREFERENCE_SAVE"
                    effect = "Unknown": message = "Settings save could not be verified."
                Case Else: message = "The Settings display or staging action failed. Saved settings were not changed."
            End Select
            nextStep = "Reload the owning Settings editor before retrying."
        Case "COMPLETED", "UNCHANGED", "REFRESHED", "SELECTED", "STAGED"
            If InStr(1, SuccessCodes(id), "|" & code & "|", vbBinaryCompare) = 0 Or code = "" Then Exit Function
            Select Case code
                Case "COMPLETED"
                    effect = "Changed": message = "Detail profile saved."
                    If PersonalControl(id) Then message = "Personal Action Path preference saved."
                Case "UNCHANGED": message = "Personal Action Path preference already matches."
                Case "REFRESHED": message = "Settings reloaded. Saved settings were not changed."
                Case "SELECTED": message = "Settings display selection changed. Saved settings were not changed."
                Case "STAGED"
                    message = "Settings changes staged; not saved."
                    nextStep = "Use the owning Save command to apply staged settings."
            End Select
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
