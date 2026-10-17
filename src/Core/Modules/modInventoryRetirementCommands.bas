Attribute VB_Name = "modInventoryRetirementCommands"
Option Explicit

' Primitive coordinator boundary; headless callers supply a provisioned user.
Public Function RetireActiveEntities(ByVal warehouse As String, ByVal station As String, _
        ByVal user As String, ByVal sku As String, ByVal reason As String, ByRef retiredCount As Long, _
        ByRef eventId As String, ByRef report As String, Optional ByVal context As String = "") As Boolean
    Dim inputs As cInventoryCommandInputs, values As Variant, payload As String, role As Variant
    Dim path As String, setupReport As String, queueError As String, batchReport As String
    Dim applied As Boolean, processed As Long, confirmed As Boolean
    On Error GoTo Failed
    retiredCount = 0: eventId = ""
    If Trim$(reason) = "" Or Trim$(sku) = "" Then report = "An exact Item Code and deletion reason are required.": Exit Function
    If context <> "" Then
        values = modInventoryCommands.SourceValues(context)
        If Not IsArray(values) Then GoTo Refused
        If values(0) <> warehouse Or values(1) <> station Or values(2) <> user Then GoTo Refused
        If Not modInventoryCommands.BeginRetirement(context, sku, reason, payload, retiredCount, report) Then Exit Function
    Else
        If Not HeadlessAuthority(warehouse, station, user, report) Then Exit Function
        Set inputs = New cInventoryCommandInputs
        If Not inputs.CaptureSource(warehouse) Then GoTo Refused
        If Not inputs.BeginRetirement(sku, reason, payload, retiredCount) Then GoTo Refused
    End If
    If retiredCount = 0 Then RetireActiveEntities = True: Exit Function
    For Each role In Array("RECEIVE", "SHIP", "PRODUCTION")
        If Not CommandCurrent(context, inputs, warehouse, station, user, report) Then Exit Function
        If Not modConfig.EnsureStationInbox(warehouse, station, CStr(role), "", path, setupReport) Then
            report = "Station inbox setup failed: " & setupReport: Exit Function
        End If
        If Not CommandCurrent(context, inputs, warehouse, station, user, report) Then Exit Function
    Next role
    If Not modInventoryCommandReceipts.FindExact(warehouse, station, user, payload, eventId, report) Then Exit Function
    If Not CommandCurrent(context, inputs, warehouse, station, user, report) Then Exit Function
    If eventId = "" Then
        If Not modRoleEventWriter.QueueAdminInventoryAdjustEvent(warehouse, station, user, payload, _
                "Admin delete inventory item " & sku & ". Reason: " & reason, 0, eventId, queueError, "") Then
            report = "Inventory retirement could not be queued: " & queueError: GoTo Unconfirmed
        End If
    End If
    If context <> "" Then
        If Not modInventoryCommands.BeginRetirementSave(context, report) Then GoTo Unconfirmed
    Else
        If Not CommandCurrent(context, inputs, warehouse, station, user, report) Then GoTo Unconfirmed
        If Not inputs.BeginSave("Retirement") Then GoTo Unconfirmed
    End If
    processed = modProcessor.RunBatch(warehouse, 0, batchReport, eventId, applied, True)
    If Not applied Then report = batchReport: GoTo Unconfirmed
    If context <> "" Then
        confirmed = modInventoryCommands.ConfirmRetirement(context, eventId, report)
    ElseIf HeadlessAuthority(warehouse, station, user, report) Then
        confirmed = inputs.ConfirmEventResult(eventId, "Retirement")
    End If
    If Not confirmed Then
        report = "Partial result: exact retirement event application was saved; confirmation failed." & _
                 vbCrLf & "Event: " & eventId & vbCrLf & report
        Exit Function
    End If
    report = "Exact retirement event application was saved." & vbCrLf & "Event: " & eventId
    RetireActiveEntities = True
    Exit Function
Refused:
    report = "The inventory command source or authority changed. Reopen Add/Edit Inventory."
    Exit Function
Unconfirmed:
    If context <> "" Then modInventoryCommands.Release context
    report = "Inventory retirement application is unconfirmed." & vbCrLf & report
    If eventId <> "" Then report = report & vbCrLf & "Event: " & eventId
    Exit Function
Failed:
    report = "Inventory retirement command failed."
    GoTo Unconfirmed
End Function

Private Function CommandCurrent(ByVal context As String, ByVal inputs As cInventoryCommandInputs, _
        ByVal warehouse As String, ByVal station As String, ByVal user As String, ByRef report As String) As Boolean
    If context <> "" Then
        CommandCurrent = modInventoryCommands.CheckForm(context, report)
    ElseIf HeadlessAuthority(warehouse, station, user, report) Then
        CommandCurrent = inputs.Unchanged()
        If Not CommandCurrent Then report = "The selected inventory source changed."
    End If
End Function

Private Function HeadlessAuthority(ByVal warehouse As String, ByVal station As String, _
        ByVal user As String, ByRef report As String) As Boolean
    If warehouse = "" Or station = "" Or user = "" Then GoTo Refused
    If Not modAuth.LoadAuth(warehouse) Then GoTo Refused
    If Not modAuth.HasProvisionedCapabilityForSystem("ADMIN_MAINT", user, warehouse, station) Then GoTo Refused
    HeadlessAuthority = True
    Exit Function
Refused:
    report = "The provisioned user lacks ADMIN_MAINT for this warehouse and station."
End Function
