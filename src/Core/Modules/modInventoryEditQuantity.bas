Attribute VB_Name = "modInventoryEditQuantity"
Option Explicit
Option Private Module

' Called only after the catalog and immutable EDIT receipt were confirmed saved.
Public Function ApplyPending(ByVal context As String, ByVal receipt As Variant, _
        ByVal reason As String, ByRef observedEvent As String, ByRef report As String) As Boolean
    Dim creating As Boolean, queued As Boolean, applied As Boolean, payload As String
    Dim eventId As String, attempted As Boolean, submitted As Boolean, batchReport As String, processed As Long
    On Error GoTo Failed
    creating = (CStr(receipt(12)) = "CREATE")
    If creating Then
        If Not modInventoryEditCreation.PendingUnchanged(context, receipt) Then GoTo Changed
        If Not modInventoryCommands.PrepareCreationInboxes(context, report) Then Exit Function
        If Not modInventoryEditCreation.PendingUnchanged(context, receipt) Then GoTo Changed
    Else
        If Not modInventoryEditPlan.PendingEntityUnchanged(context, receipt) Then GoTo Changed
    End If
    eventId = CStr(receipt(7))
    If creating Then
        If Not modInventoryCreateReceipts.ReadPendingCreation(context, eventId, queued, payload, report) Then Exit Function
    Else
        If Not modInventoryCreateReceipts.ReadPendingAdjustment(context, eventId, queued, payload, report) Then Exit Function
    End If
    If queued And payload <> CStr(receipt(11)) Then GoTo Failed
    If queued Then observedEvent = eventId & vbTab & "Submitted"
    If Not modInventoryCommands.SelectSku(context, CStr(receipt(4)), report) Then Exit Function
    If creating Then
        If Not modInventoryCommands.BeginCreate(context, CStr(receipt(11)), report) Then Exit Function
    Else
        If Not modInventoryCommands.BeginQuantity(context, CStr(receipt(8)), CDbl(receipt(10)), report) Then Exit Function
    End If
    If Not queued Then
        If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
        If Not WriteAvailable(context, report) Then Exit Function
        If creating Then
            If Not modInventoryEditCreation.PendingUnchanged(context, receipt) Then GoTo Changed
            submitted = modInventoryCreateWire.QueueCreate(CStr(receipt(1)), CStr(receipt(2)), CStr(receipt(3)), _
                CStr(receipt(11)), "Admin completed catalog-only inventory item. Reason: " & reason, eventId, report, observedEvent)
        Else
            submitted = modRoleEventWriter.QueueAdminInventoryAdjustEvent(CStr(receipt(1)), CStr(receipt(2)), _
                CStr(receipt(3)), CStr(receipt(11)), "Admin set inventory quantity. Reason: " & reason, _
                0, eventId, report, "", attempted)
            If submitted Then
                observedEvent = eventId & vbTab & "Submitted"
            ElseIf attempted Then
                observedEvent = eventId & vbTab & "Unknown"
            End If
        End If
        If Not submitted Then Exit Function
        If eventId <> CStr(receipt(7)) Then GoTo Failed
    End If
    If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
    If Not WriteAvailable(context, report) Then Exit Function
    If creating Then
        If Not modInventoryEditCreation.PendingUnchanged(context, receipt) Then GoTo Changed
        If Not modInventoryCommands.BeginCreateSave(context, report) Then Exit Function
    Else
        If Not modInventoryCommands.BeginQuantitySave(context, report) Then Exit Function
    End If
    processed = modProcessor.RunBatch(CStr(receipt(1)), 0, batchReport, eventId, applied, True)
    If Not applied Then GoTo Failed
    If creating Then
        If Not modInventoryCommands.ConfirmCreate(context, eventId, report) Then Exit Function
    Else
        If Not modInventoryCommands.ConfirmQuantity(context, eventId, report) Then Exit Function
    End If
    If Not modInventoryEditProof.QuantityApplied(context, receipt, applied, report) Then Exit Function
    ApplyPending = applied
    Exit Function
Changed:
    report = "The original pending quantity plan changed. No quantity event was resubmitted."
    Exit Function
Failed:
    report = "Exact EDIT quantity completion is unconfirmed. " & report
End Function

Public Function WriteAvailable(ByVal context As String, ByRef report As String) As Boolean
    Dim path As String, book As Workbook
    On Error GoTo Refused
    path = modInventoryCommands.SourceInventoryPath(context)
    If path = "" Then GoTo Refused
    For Each book In Application.Workbooks
        If modAggregationFiles.SameFile(book.FullName, path) Then
            report = "Close the Inventory workbook before uploading pending EDIT changes."
            Exit Function
        End If
    Next book
    WriteAvailable = True
    Exit Function
Refused:
    report = "The pending EDIT source could not be confirmed."
End Function
