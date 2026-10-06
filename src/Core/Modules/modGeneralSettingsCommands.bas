Attribute VB_Name = "modGeneralSettingsCommands"
Option Explicit

' Owner outcomes for the General UI; these direct services create no activity.
Public Function ReloadConfig(ByRef outcome As String) As Boolean
    outcome = "FAILED"
    ReloadConfig = modConfig.Reload()
    If ReloadConfig Then outcome = "REFRESHED"
End Function

Public Sub SaveConnectionOption(ByVal requireManualEntry As Boolean, ByRef outcome As String)
    Dim rawValue As String, current As Boolean, desired As String
    outcome = "FAILED"
    rawValue = UCase$(Trim$(GetSetting("invSys", "NAS", "RequireManualServerCredentials", "FALSE")))
    current = (rawValue = "TRUE" Or rawValue = "YES" Or rawValue = "1" Or rawValue = "ON")
    If current = requireManualEntry Then outcome = "UNCHANGED": Exit Sub
    desired = IIf(requireManualEntry, "TRUE", "FALSE")
    modNasConnection.SetRequireManualServerCredentials requireManualEntry
    If GetSetting("invSys", "NAS", "RequireManualServerCredentials", "") <> desired Then _
        Err.Raise 5, , "Connection option save could not be verified."
    outcome = "COMPLETED"
End Sub
