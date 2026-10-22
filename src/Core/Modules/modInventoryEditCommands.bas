Attribute VB_Name = "modInventoryEditCommands"
Option Explicit

Private mPrepared As Object

' Whole-selection identity preflight runs before any owning inventory write.
Public Function CheckSelection(ByVal context As String, ByRef report As String) As Boolean
    Dim book As Workbook, table As ListObject, column As ListColumn, fields As Object, seen As Object
    Dim row As Long, text As String, command As String, action As String, hasEdit As Boolean
    Dim found As Boolean, receipt As Variant
    On Error GoTo Refused
    text = modInventoryCommands.OperatorWorkbookName(context)
    If text = "" Then GoTo Refused
    Set book = Application.Workbooks(text)
    Set table = modExecutionTarget.Table(book, modInventoryCommands.StagingTableName(context))
    If table Is Nothing Then GoTo Refused
    Set fields = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For Each column In table.ListColumns
        fields.Add modInventoryEditWire.WorksheetHeader(column.Name), column.Index
    Next column
    If Not fields.Exists("ACTION") Or Not fields.Exists("UPLOADRESULT") Then GoTo Refused
    For row = 1 To table.ListRows.Count
        action = UCase$(Trim$(CStr(table.DataBodyRange.Cells(row, fields("ACTION")).Value2)))
        text = CStr(table.DataBodyRange.Cells(row, fields("UPLOADRESULT")).Value2)
        command = modInventoryEditWire.ReferenceField(text, "Command:")
        If action = "EDIT" Then hasEdit = True
        If command <> "" Then
            If action <> "EDIT" Or seen.Exists(command) Then GoTo Refused
            seen.Add command, True
        ElseIf InStr(1, text, "Command:", vbTextCompare) > 0 Or _
            InStr(1, text, "Edit binding:", vbTextCompare) > 0 Then
            GoTo Refused
        End If
    Next row
    If hasEdit Then
        If Not modInventoryEditProof.ReadReceipt(context, "", found, receipt, report) Then Exit Function
    End If
    CheckSelection = modInventoryCommands.CheckForm(context, report)
    Exit Function
Refused:
    report = "An EDIT command receipt is missing, duplicated or belongs to another action. No inventory changes were made."
End Function

' A prepared reference is saved by the existing operator result boundary.
' Its presence alone does not prove application.
Public Function Prepare(ByVal context As String, ByVal row As Long, _
        ByRef reference As String, ByRef report As String) As Boolean
    Dim submitted As cInventoryEditRow, command As String, eventId As String
    Dim receipt As Variant, catalog As Variant, found As Boolean, applied As Boolean, entry As Object
    Dim existing As String, preparedSaved As Boolean, prior As Object
    On Error GoTo Refused
    reference = ""
    Set submitted = New cInventoryEditRow
    If Not submitted.Read(context, row, report) Then Exit Function
    existing = submitted.ResultText
    command = modInventoryEditWire.ReferenceField(existing, "Command:")
    eventId = modInventoryEditWire.ReferenceField(existing, "Event:")
    If command = "" Then
        If submitted.Imported Or eventId <> "" Or _
            InStr(1, existing, "Edit binding:", vbTextCompare) > 0 Then GoTo Refused
        Select Case submitted.Status
            Case "READY", "FAILED", "NOT IMPORTED", "QUEUED": GoTo Refused
        End Select
        command = UCase$(modSystemIdentity.NewId())
    Else
        If Not modTrainingWire.ValidId(LCase$(command)) Then GoTo Refused
        If modInventoryEditWire.ReferenceField(existing, "Edit binding:") = "" Then GoTo Refused
    End If
    If Not modInventoryEditProof.ReadReceipt(context, command, found, receipt, report) Then Exit Function
    If found Then
        If Not modInventoryEditPlan.Matches(submitted, existing, receipt) Then GoTo Refused
        If Not modInventoryEditProof.QuantityApplied(context, receipt, applied, report) Then Exit Function
        If submitted.Imported And Not applied Then GoTo Refused
    Else
        If submitted.Imported Then GoTo Refused
        If Not modInventoryEditPlan.Build(context, submitted, command, eventId, receipt, catalog, report) Then Exit Function
        If modInventoryEditWire.ReferenceField(existing, "Command:") <> "" Then
            If Not modInventoryEditPlan.Matches(submitted, existing, receipt) Then GoTo Refused
        End If
    End If
    ' Preparation covers the whole selection before its first canonical write.
    ' Completed receipts need only read-only proof, even when the source is open.
    If Not found Or Not applied Then
        If Not PendingWriteAvailable(context, report) Then Exit Function
    End If
    reference = modInventoryEditWire.Reference(CStr(receipt(0)), CStr(receipt(14)), CStr(receipt(7)))
    If Not mPrepared Is Nothing Then
        If mPrepared.Exists(context & "|" & command) Then
            Set prior = mPrepared(context & "|" & command)
            If CLng(prior("Row")) = row And CStr(prior("InputHash")) = submitted.InputHash Then
                If Join(prior("Receipt"), vbNullChar) = Join(receipt, vbNullChar) Then preparedSaved = prior("PreparedSaved")
            End If
        End If
    End If
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "Row", row: entry.Add "InputHash", submitted.InputHash
    entry.Add "Reason", submitted.SubmittedText("EDITREASON")
    entry.Add "Receipt", receipt: entry.Add "Catalog", catalog
    entry.Add "Saved", found: entry.Add "Applied", applied: entry.Add "Writing", False
    entry.Add "PreparedSaved", preparedSaved
    If mPrepared Is Nothing Then Set mPrepared = CreateObject("Scripting.Dictionary")
    Set mPrepared(context & "|" & command) = entry
    Prepare = modInventoryCommands.CheckForm(context, report)
    Exit Function
Refused:
    report = "The EDIT command receipt does not confirm this row. No new event was queued."
End Function

' Confirm all prepared row references in the one saved operator boundary.
' Later per-row result text can change without invalidating that earlier proof.
Public Function ConfirmPreparedSelection(ByVal context As String, ByRef report As String) As Boolean
    Dim key As Variant, entry As Object, submitted As cInventoryEditRow, found As Boolean
    If Not mPrepared Is Nothing Then
        For Each key In mPrepared.Keys
            If Left$(CStr(key), Len(context) + 1) = context & "|" Then found = True
        Next key
    End If
    If Not found Then ConfirmPreparedSelection = True: Exit Function
    If Not modInventoryCommands.ConfirmStageSaved(context, report) Then Exit Function
    If Not mPrepared Is Nothing Then
        For Each key In mPrepared.Keys
            If Left$(CStr(key), Len(context) + 1) = context & "|" Then
                Set entry = mPrepared(key)
                Set submitted = New cInventoryEditRow
                If Not submitted.Read(context, CLng(entry("Row")), report) Then Exit Function
                If Not modInventoryEditPlan.Matches(submitted, submitted.ResultText, entry("Receipt")) Then Exit Function
                If Not entry("Saved") Or Not entry("Applied") Then
                    If Not PendingWriteAvailable(context, report) Then Exit Function
                End If
                entry("PreparedSaved") = True
            End If
        Next key
    End If
    ConfirmPreparedSelection = modInventoryCommands.CheckForm(context, report)
End Function

Public Function Execute(ByVal context As String, ByVal row As Long, ByRef report As String, _
                        Optional ByRef observedEvent As String = "") As Boolean
    Dim reference As String, command As String, entry As Object, receipt As Variant, catalog As Variant
    Dim names As Variant, values As Variant, headers As Variant, i As Long, result As Variant
    Dim found As Boolean, savedReceipt As Variant, applied As Boolean, queued As Boolean, payload As String
    Dim eventId As String, batchReport As String, processed As Long
    Dim attempted As Boolean, submitted As Boolean
    On Error GoTo Failed
    observedEvent = ""
    If Not Prepare(context, row, reference, report) Then Exit Function
    command = modInventoryEditWire.ReferenceField(reference, "Command:")
    Set entry = mPrepared(context & "|" & command): receipt = entry("Receipt")
    ' Independent operator-save proof was captured for the whole prepared batch.
    If Not entry("PreparedSaved") Then GoTo Failed
    If Not entry("Saved") Then
        If Not modInventoryCommands.BeginCatalog(context, CStr(receipt(4)), report) Then GoTo Failed
        catalog = entry("Catalog"): names = catalog(0): values = catalog(1)
        For i = LBound(names) To UBound(names)
            If Not modInventoryCommands.PlanCatalogField(context, CStr(receipt(4)), CStr(names(i)), _
                CStr(values(i)), report) Then GoTo Failed
        Next i
        headers = modInventoryEditWire.ReceiptHeaders()
        If Not modInventoryCommands.PlanEditReceipt(context, headers, receipt, report) Then GoTo Failed
        entry("Writing") = True
        result = modInventoryDomainBridge.SaveWorksheetEditCatalogBridge(context, command)
        entry("Writing") = False
        If Not IsArray(result) Then GoTo Failed
        report = CStr(result(1))
        If Not CBool(result(0)) Then GoTo Failed
        If Not modInventoryEditProof.ReadReceipt(context, command, found, savedReceipt, report) Then GoTo Failed
        If Not found Then GoTo Failed
        If Join(savedReceipt, vbNullChar) <> Join(receipt, vbNullChar) Then GoTo Failed
        entry("Saved") = True
    End If
    If Not modInventoryEditProof.QuantityApplied(context, receipt, applied, report) Then GoTo Failed
    If applied And CStr(receipt(12)) = "ADJUST" Then observedEvent = CStr(receipt(7)) & vbTab & "Submitted"
    If Not applied Then
        If Not modInventoryEditPlan.PendingEntityUnchanged(context, receipt) Then
            report = "The original pending quantity plan changed. No adjustment was resubmitted."
            GoTo Failed
        End If
        eventId = CStr(receipt(7))
        If Not modInventoryCreateReceipts.ReadPendingAdjustment(context, eventId, queued, payload, report) Then GoTo Failed
        If queued And payload <> CStr(receipt(11)) Then GoTo Failed
        If queued Then observedEvent = eventId & vbTab & "Submitted"
        If Not modInventoryCommands.SelectSku(context, CStr(receipt(4)), report) Then GoTo Failed
        If Not modInventoryCommands.BeginQuantity(context, CStr(receipt(8)), CDbl(receipt(10)), report) Then GoTo Failed
        If Not queued Then
            If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
            If Not PendingWriteAvailable(context, report) Then GoTo Failed
            submitted = modRoleEventWriter.QueueAdminInventoryAdjustEvent(CStr(receipt(1)), CStr(receipt(2)), _
                CStr(receipt(3)), CStr(receipt(11)), "Admin set inventory quantity. Reason: " & CStr(entry("Reason")), _
                0, eventId, report, "", attempted)
            If submitted Then
                observedEvent = eventId & vbTab & "Submitted"
            ElseIf attempted Then
                observedEvent = eventId & vbTab & "Unknown"
            End If
            If Not submitted Then GoTo Failed
            If eventId <> CStr(receipt(7)) Then GoTo Failed
        End If
        If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
        If Not PendingWriteAvailable(context, report) Then GoTo Failed
        If Not modInventoryCommands.BeginQuantitySave(context, report) Then GoTo Failed
        processed = modProcessor.RunBatch(CStr(receipt(1)), 0, batchReport, eventId, applied, True)
        If Not applied Then GoTo Failed
        If Not modInventoryCommands.ConfirmQuantity(context, eventId, report) Then GoTo Failed
        If Not modInventoryEditProof.QuantityApplied(context, receipt, applied, report) Then GoTo Failed
        If Not applied Then GoTo Failed
    End If
    report = "Catalog changes and the EDIT receipt are confirmed saved." & vbCrLf
    If CStr(receipt(12)) = "ADJUST" Then
        report = report & "The original quantity event is confirmed applied."
    ElseIf CStr(receipt(12)) = "NOOP" Then
        report = report & "Quantity matched the target; no quantity event was needed."
    Else
        report = report & "No quantity change was requested."
    End If
    report = report & vbCrLf & reference
    Execute = True
    Exit Function
Failed:
    report = "Partial result: EDIT completion is unconfirmed. " & report
    If Not entry Is Nothing Then
        If entry("Saved") Then report = report & vbCrLf & "Catalog changes and the EDIT receipt were saved; quantity completion is unconfirmed."
    End If
    If reference <> "" Then report = report & vbCrLf & reference
    modInventoryCommands.Release context
End Function

' Only the in-flight Core command can expose the plan to the owning Domain.
' The envelope is path, catalog names/values, receipt names/values: scalars only.
Public Function ClaimWritePlan(ByVal context As String, ByVal command As String) As Variant
    Dim entry As Object, submitted As cInventoryEditRow, report As String, catalog As Variant
    On Error GoTo Refused
    If mPrepared Is Nothing Then Exit Function
    Set entry = mPrepared(context & "|" & command)
    If Not entry("Writing") Or entry("Saved") Then Exit Function
    Set submitted = New cInventoryEditRow
    If Not submitted.Read(context, CLng(entry("Row")), report) Then Exit Function
    If submitted.InputHash <> entry("InputHash") Then Exit Function
    If Not modInventoryEditPlan.Matches(submitted, submitted.ResultText, entry("Receipt")) Then Exit Function
    If Not PendingWriteAvailable(context, report) Then Exit Function
    catalog = entry("Catalog")
    ClaimWritePlan = Array(modInventoryCommands.SourceInventoryPath(context), _
        catalog(0), catalog(1), modInventoryEditWire.ReceiptHeaders(), entry("Receipt"))
Refused:
End Function

Private Function PendingWriteAvailable(ByVal context As String, ByRef report As String) As Boolean
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
    PendingWriteAvailable = True
    Exit Function
Refused:
    report = "The pending EDIT source could not be confirmed."
End Function

Public Sub ReleaseContext(ByVal context As String)
    Dim key As Variant
    If mPrepared Is Nothing Then Exit Sub
    For Each key In mPrepared.Keys
        If Left$(CStr(key), Len(context) + 1) = context & "|" Then mPrepared.Remove key
    Next key
End Sub
