Attribute VB_Name = "modActionPathPreference"
Option Explicit

Private Const SETTINGS_APP As String = "invSys"
Private Const SETTINGS_SECTION As String = "ActionPathPreferencesV1"
Private Const DEFAULT_CHOICE As String = "Use warehouse default"

' D18 primitive personal boundary; warehouse policy remains read-only here.
Public Function ReadPreference(ByVal context As String, ByRef choice As String, _
                               ByRef effective As String, ByRef evidence As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget, key As String, saved As String, model As Object
    Dim version As Long, collect As Boolean, visible As Boolean, notice As String
    On Error GoTo Failed
    choice = "": effective = "Unavailable": evidence = "Diagnostic evidence unavailable."
    If Not PreferenceKey(context, key, target, report) Then Exit Function
    saved = GetSetting(SETTINGS_APP, SETTINGS_SECTION, key, DEFAULT_CHOICE)
    choice = saved
    report = "Personal preference loaded."
    If Not ValidChoice(saved) Then
        choice = DEFAULT_CHOICE
        report = "Saved preference is invalid; using warehouse default."
    End If
    Set model = CreateObject("Scripting.Dictionary")
    If modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice, model) Then
        effective = choice
        If choice = DEFAULT_CHOICE Then effective = CStr(model("DefaultView"))
        effective = effective & IIf(choice = DEFAULT_CHOICE, " (warehouse default)", " (personal choice)") & _
            "; tracking policy version " & CStr(version)
        If CBool(model("ViewerActionPathCaptureEnabled")) Then
            evidence = "Diagnostic evidence depends on recorded controls; a view choice creates no evidence."
        Else
            evidence = "Diagnostic evidence unavailable: recorded-control capture is off."
        End If
    Else
        report = report & " Tracking policy unavailable; no effective view or tracking permission is inferred."
        evidence = "Diagnostic evidence unavailable: tracking policy could not be read."
    End If
    If Not PreferenceKey(context, key, target, notice) Then report = notice: choice = "": Exit Function
    ReadPreference = True
    Exit Function
Failed:
    choice = "": effective = "Unavailable": evidence = "Diagnostic evidence unavailable."
    report = "Personal preference could not be read. No settings were changed."
End Function

Public Function SavePreference(ByVal context As String, ByVal choice As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget, key As String
    On Error GoTo Failed
    If Not PreferenceKey(context, key, target, report) Then Exit Function
    report = "Invalid Action Path view. No settings were changed."
    If Not ValidChoice(choice) Then Exit Function
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, key, choice
    If GetSetting(SETTINGS_APP, SETTINGS_SECTION, key, "") <> choice Then GoTo Failed
    SavePreference = True
    report = "Your Action Path preference was saved."
    Exit Function
Failed:
    report = "Personal preference save could not be verified. Reload before retrying."
End Function

Public Function Choices() As String
    Choices = DEFAULT_CHOICE & vbLf & "How-To" & vbLf & "Diagnostic" & vbLf & "Compare both"
End Function

Private Function ValidChoice(ByVal choice As String) As Boolean
    ValidChoice = (choice = DEFAULT_CHOICE Or choice = "How-To" Or choice = "Diagnostic" Or choice = "Compare both")
End Function

Private Function PreferenceKey(ByVal context As String, ByRef key As String, ByRef target As WarehouseTarget, ByRef report As String) As Boolean
    report = "Session or warehouse changed. Reopen Settings."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    key = "U" & IdentityPart(modAuth.GetCurrentUserId()) & "W" & IdentityPart(target.WarehouseId)
    PreferenceKey = True
End Function

Private Function IdentityPart(ByVal value As String) As String
    Dim index As Long
    If value = "" Then Err.Raise 5
    For index = 1 To Len(value)
        IdentityPart = IdentityPart & Right$("0000" & Hex$(AscW(Mid$(value, index, 1)) And &HFFFF&), 4)
    Next index
End Function
