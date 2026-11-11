Attribute VB_Name = "modAdminInventoryCreate"
Option Explicit
Option Private Module

' The owning Admin boundary generates a fresh entity identity only when Core
' has independently verified a new catalog-only worksheet completion.
Public Function PrepareWorksheetEdit(ByVal context As String, ByVal row As Long, _
        ByRef reference As String, ByRef report As String) As Boolean
    Dim needsKey As Boolean, systemKey As String
    PrepareWorksheetEdit = modInventoryEditCommands.Prepare(context, row, reference, report, "", needsKey)
    If PrepareWorksheetEdit Or Not needsKey Then Exit Function
    If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
    systemKey = modRoleEventWriter.CreateSystemKey()
    PrepareWorksheetEdit = modInventoryEditCommands.Prepare(context, row, reference, report, systemKey, needsKey)
End Function

' Explicit completion of the retained catalog-only case. This never revives a
' retired key or allocates quantity among existing entities.
Public Function CompleteCatalogItem(ByVal warehouseId As String, ByVal stationId As String, _
        ByVal userId As String, ByVal sku As String, ByVal itemName As String, ByVal uom As String, _
        ByVal locationVal As String, ByVal targetQty As Double, ByVal editReason As String, _
        ByRef report As String, Optional ByVal interactiveContext As String = "") As Boolean
    Dim item As Object, role As Variant, inboxPath As String
    Dim payloadJson As String, eventIdOut As String, queueError As String, batchReport As String
    Dim processedCount As Long, applied As Boolean, systemKey As String, noteText As String
    On Error GoTo Failed
    If targetQty <= 0 Or Trim$(editReason) = "" Or Trim$(sku) = "" Then
        report = "An exact catalog item, positive target and edit reason are required.": Exit Function
    End If
    If interactiveContext <> "" Then
        If Not modInventoryCommands.PrepareCreationInboxes(interactiveContext, report) Then Exit Function
    Else
    For Each role In Array("RECEIVE", "SHIP", "PRODUCTION")
        If Not CommandCurrent(interactiveContext, warehouseId, stationId, userId, report) Then Exit Function
        If Not modConfig.EnsureStationInbox(warehouseId, stationId, CStr(role), "", inboxPath, report) Then Exit Function
        If Not CommandCurrent(interactiveContext, warehouseId, stationId, userId, report) Then Exit Function
    Next role
    End If
    systemKey = modRoleEventWriter.CreateSystemKey()
    noteText = "Admin completed catalog-only inventory item. Reason: " & editReason & _
               "; SubmittedQty=" & CStr(targetQty)
    Set item = CreateObject("Scripting.Dictionary"): item.CompareMode = vbTextCompare
    item("ITEM_CODE") = sku: item("ITEM") = itemName: item("UOM") = uom
    item("LOCATION") = locationVal: item("TOTAL INV") = targetQty
    item("QtyAvailable") = targetQty: item("Reason") = editReason
    payloadJson = modInventoryCreateWire.BuildPayload(systemKey, sku, targetQty, _
        locationVal, "GOOD", "", noteText, item.Keys, PrimitiveFieldValues(item))
    If payloadJson = "" Or payloadJson = "[]" Then report = "Inventory creation payload was empty.": Exit Function
    If interactiveContext <> "" Then
        If Not modInventoryCommands.BeginCreate(interactiveContext, payloadJson, report) Then Exit Function
    End If
    If Not CommandCurrent(interactiveContext, warehouseId, stationId, userId, report) Then Exit Function
    If Not modInventoryCreateWire.QueueCreate(warehouseId, stationId, userId, _
            payloadJson, noteText, eventIdOut, queueError) Then
        report = "Inventory creation event could not be queued: " & queueError: GoTo Unconfirmed
    End If
    If Not CommandCurrent(interactiveContext, warehouseId, stationId, userId, report) Then GoTo Unconfirmed
    If interactiveContext <> "" Then
        If Not modInventoryCommands.BeginCreateSave(interactiveContext, report) Then GoTo Unconfirmed
    End If
    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport, eventIdOut, applied, True)
    If Not applied Then report = batchReport: GoTo Unconfirmed
    If Not CommandCurrentAfterSave(interactiveContext, warehouseId, stationId, userId, eventIdOut, report) Then
        report = "Partial result: exact event application was saved, but command confirmation failed." & _
                 vbCrLf & "Event: " & eventIdOut & vbCrLf & report
        Exit Function
    End If
    report = "Catalog-only inventory item completed as a managed entity. Exact event application was saved." & _
             vbCrLf & "Warehouse: " & warehouseId & vbCrLf & "SKU: " & sku & _
             vbCrLf & "System_Key: " & systemKey & vbCrLf & "Event: " & eventIdOut & _
             vbCrLf & "Starting quantity: " & CStr(targetQty) & vbCrLf & "Reason: " & editReason & _
             vbCrLf & "Processor: " & batchReport & vbCrLf & "Refresh Inventory Viewer and Production to see the item."
    CompleteCatalogItem = True
    Exit Function
Unconfirmed:
    If interactiveContext <> "" Then modInventoryCommands.Release interactiveContext
    report = "Inventory creation application is unconfirmed." & vbCrLf & report
    If eventIdOut <> "" Then report = report & vbCrLf & "Event: " & eventIdOut
    Exit Function
Failed:
    report = "Complete catalog-only inventory item failed: " & Err.Description
    GoTo Unconfirmed
End Function

Public Function PrimitiveFieldValues(ByVal fields As Object) As Variant
    Dim values As Variant, key As Variant, i As Long
    If fields.Count = 0 Then PrimitiveFieldValues = Array(): Exit Function
    ReDim values(0 To fields.Count - 1)
    For Each key In fields.Keys
        If IsObject(fields(key)) Then
            ' Preserve the existing serializer's null representation locally;
            ' an object must never cross the XLAM boundary inside a Variant.
            values(i) = Null
        ElseIf IsArray(fields(key)) Then
            Err.Raise vbObjectError + 2192, , "Inventory custom fields must be scalar values."
        Else
            values(i) = fields(key)
        End If
        i = i + 1
    Next key
    PrimitiveFieldValues = values
End Function

Private Function CommandCurrent(ByVal context As String, ByVal warehouse As String, _
        ByVal station As String, ByVal user As String, ByRef report As String) As Boolean
    Dim values As Variant
    If context <> "" Then
        values = modInventoryCommands.SourceValues(context)
        If Not IsArray(values) Then GoTo Refused
        If values(0) <> warehouse Or values(1) <> station Or values(2) <> user Then GoTo Refused
        CommandCurrent = True
    Else
        If Trim$(user) = "" Or Trim$(station) = "" Then GoTo Refused
        If Not modAuth.LoadAuth(warehouse) Then GoTo Refused
        CommandCurrent = modAuth.HasProvisionedCapabilityForSystem("ADMIN_MAINT", user, warehouse, station)
    End If
    If CommandCurrent Then Exit Function
Refused:
    report = "The inventory source or authority changed. Reopen Add/Edit Inventory."
End Function

Private Function CommandCurrentAfterSave(ByVal context As String, ByVal warehouse As String, _
        ByVal station As String, ByVal user As String, ByVal eventId As String, ByRef report As String) As Boolean
    If context <> "" Then
        CommandCurrentAfterSave = modInventoryCommands.ConfirmCreate(context, eventId, report)
    Else
        CommandCurrentAfterSave = CommandCurrent(context, warehouse, station, user, report)
    End If
End Function
