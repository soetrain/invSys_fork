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

    resolvedWh = Trim$(warehouseId)
    resolvedSt = Trim$(stationId)

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
        errorMessage = "Current user lacks " & capability & " capability."
        Exit Function
    End If

    CanCurrentUserPerformCapability = True
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
