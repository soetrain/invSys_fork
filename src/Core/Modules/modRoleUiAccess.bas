Attribute VB_Name = "modRoleUiAccess"
Option Explicit

Private Const SHAPE_VISIBLE_FALSE As Long = 0
Private Const SHAPE_VISIBLE_TRUE As Long = -1

Public Function CanCurrentUserPerformCapability(ByVal capability As String, _
                                                Optional ByVal userId As String = "", _
                                                Optional ByVal warehouseId As String = "", _
                                                Optional ByVal stationId As String = "", _
                                                Optional ByRef errorMessage As String = "") As Boolean
    Dim resolvedWh As String
    Dim resolvedSt As String
    Dim resolvedUser As String
    Dim currentTarget As WarehouseTarget

    resolvedWh = Trim$(warehouseId)
    resolvedSt = Trim$(stationId)

    If CapabilityRequiresNasTargetAccess(capability) Then
        If Not modNasConnection.IsCurrentTargetAllowed(True) Then
            errorMessage = "A connected NAS warehouse target is required before using role controls."
            Exit Function
        End If
        If Not modAuth.IsSignedIn() Then
            errorMessage = "Current invSys user is not signed in."
            Exit Function
        End If
        resolvedUser = Trim$(userId)
        If resolvedUser = "" Then resolvedUser = Trim$(modAuth.GetCurrentUserId())
        If resolvedUser = "" Then
            errorMessage = "Current invSys user is not signed in."
            Exit Function
        End If
        If StrComp(resolvedUser, Trim$(modAuth.GetCurrentUserId()), vbTextCompare) <> 0 Then
            errorMessage = "Requested user does not match the signed-in invSys user."
            Exit Function
        End If
        Set currentTarget = modNasConnection.GetCurrentTarget()
        If currentTarget Is Nothing Then
            errorMessage = "A connected NAS warehouse target is required before using role controls."
            Exit Function
        End If
        If resolvedWh = "" Then resolvedWh = currentTarget.WarehouseId
        If resolvedSt = "" Then resolvedSt = currentTarget.StationId
    End If

    If Not modConfig.LoadConfig(resolvedWh, resolvedSt) Then
        errorMessage = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If

    If resolvedWh = "" Then resolvedWh = modConfig.GetWarehouseId()
    If resolvedSt = "" Then resolvedSt = modConfig.GetStationId()
    If resolvedWh = "" Or resolvedSt = "" Then
        errorMessage = "WarehouseId and StationId are required."
        Exit Function
    End If

    resolvedUser = Trim$(userId)
    If resolvedUser = "" Then resolvedUser = modRoleEventWriter.ResolveCurrentUserId()
    If resolvedUser = "" Then
        errorMessage = "Unable to resolve current user identity."
        Exit Function
    End If

    If Not modAuth.LoadAuth(resolvedWh) Then
        errorMessage = "Auth load failed: " & modAuth.ValidateAuth()
        Exit Function
    End If

    If Not modAuth.CanPerform(capability, resolvedUser, resolvedWh, resolvedSt, "ROLE_UI", capability & ":" & resolvedUser) Then
        errorMessage = "Current user lacks " & capability & " capability." & vbCrLf & _
                       "User=" & ValueOrBlankRoleUi(resolvedUser) & _
                       "; Warehouse=" & ValueOrBlankRoleUi(resolvedWh) & _
                       "; Station=" & ValueOrBlankRoleUi(resolvedSt) & _
                       "; Auth=" & ValueOrBlankRoleUi(modAuth.GetResolvedAuthWorkbookName())
        Exit Function
    End If

    CanCurrentUserPerformCapability = True
End Function

Public Function CanCurrentUserPerformCapabilityCached(ByVal capability As String, _
                                                      Optional ByRef errorMessage As String = "") As Boolean
    Dim resolvedUser As String
    Dim currentTarget As WarehouseTarget

    If CapabilityRequiresNasTargetAccess(capability) Then
        If Not modNasConnection.IsCurrentTargetAllowed(True) Then
            errorMessage = "A connected NAS warehouse target is required before using role controls."
            Exit Function
        End If
    End If

    If Not modAuth.IsSignedIn() Then
        errorMessage = "Current invSys user is not signed in."
        Exit Function
    End If

    resolvedUser = Trim$(modAuth.GetCurrentUserId())
    If resolvedUser = "" Then
        errorMessage = "Current invSys user is not signed in."
        Exit Function
    End If

    If CapabilityRequiresNasTargetAccess(capability) Then
        Set currentTarget = modNasConnection.GetCurrentTarget()
        If currentTarget Is Nothing Then
            errorMessage = "A connected NAS warehouse target is required before using role controls."
            Exit Function
        End If
    End If

    If Not currentTarget Is Nothing Then
        If Not modAuth.CanPerform(capability, resolvedUser, currentTarget.WarehouseId, currentTarget.StationId, "RIBBON", capability & ":" & resolvedUser) Then
            errorMessage = "Current user lacks " & capability & " capability."
            Exit Function
        End If
    ElseIf Not modAuth.CanPerform(capability, resolvedUser, "", "", "RIBBON", capability & ":" & resolvedUser) Then
        errorMessage = "Current user lacks " & capability & " capability."
        Exit Function
    End If

    CanCurrentUserPerformCapabilityCached = True
End Function

Private Function CapabilityRequiresNasTargetAccess(ByVal capability As String) As Boolean
    Select Case UCase$(Trim$(capability))
        Case "RECEIVE_POST", "SHIP_POST", "PROD_POST", "ADMIN_MAINT"
            CapabilityRequiresNasTargetAccess = True
    End Select
End Function

Private Function ValueOrBlankRoleUi(ByVal valueIn As String) As String
    valueIn = Trim$(valueIn)
    If valueIn = "" Then
        ValueOrBlankRoleUi = "<blank>"
    Else
        ValueOrBlankRoleUi = valueIn
    End If
End Function

Public Function RequireCurrentUserCapability(ByVal capability As String, _
                                             Optional ByVal deniedMessage As String = "", _
                                             Optional ByVal userId As String = "", _
                                             Optional ByVal warehouseId As String = "", _
                                             Optional ByVal stationId As String = "", _
                                             Optional ByRef errorMessage As String = "") As Boolean
    RequireCurrentUserCapability = CanCurrentUserPerformCapability(capability, userId, warehouseId, stationId, errorMessage)
    If RequireCurrentUserCapability Then Exit Function

    If deniedMessage = "" Then deniedMessage = errorMessage
    If deniedMessage <> "" Then MsgBox deniedMessage, vbExclamation
End Function

Public Function RequireCurrentUserCapabilityCached(ByVal capability As String, _
                                                   Optional ByVal deniedMessage As String = "", _
                                                   Optional ByRef errorMessage As String = "") As Boolean
    RequireCurrentUserCapabilityCached = CanCurrentUserPerformCapabilityCached(capability, errorMessage)
    If RequireCurrentUserCapabilityCached Then Exit Function

    If deniedMessage = "" Then deniedMessage = errorMessage
    If deniedMessage <> "" Then MsgBox deniedMessage, vbExclamation
End Function

Public Sub ApplyShapeCapability(ByVal ws As Worksheet, _
                                ByVal shapeName As String, _
                                ByVal capability As String, _
                                Optional ByVal userId As String = "", _
                                Optional ByVal warehouseId As String = "", _
                                Optional ByVal stationId As String = "")
    Dim shp As Shape
    Dim errorMessage As String

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    shp.Visible = IIf(CanCurrentUserPerformCapability(capability, userId, warehouseId, stationId, errorMessage), SHAPE_VISIBLE_TRUE, SHAPE_VISIBLE_FALSE)
End Sub
