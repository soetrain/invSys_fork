Attribute VB_Name = "modAuthUserCommands"
Option Explicit

Private mWriting As Boolean

' D8-USERS-01: transient interactive context, never an account record or log value.
Public Function CaptureContext() As String
    Dim target As WarehouseTarget
    On Error GoTo Unavailable
    If Not modAuth.IsSignedIn() Or Not modNasConnection.IsCurrentTargetAllowed(False) Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    If target.AuthPath = "" Then Exit Function
    CaptureContext = modTrainingJson.Quote(target.WarehouseId) & modTrainingJson.Quote(target.StationId) & _
        modTrainingJson.Quote(target.RuntimeRoot) & modTrainingJson.Quote(target.AuthPath) & _
        modTrainingJson.Quote(target.ConfigPath) & modTrainingJson.Quote(target.HubRoot) & _
        CStr(target.SourceType) & modTrainingJson.Quote(modAuth.GetCurrentUserId()) & CStr(modAuthSession.Version())
Unavailable:
End Function

' Capability bits retain the six existing form choices, in their displayed order.
Public Function SaveUser(ByVal context As String, ByVal authPath As String, ByVal userId As String, _
                         ByVal displayName As String, ByVal credential As String, ByVal warehouseScope As String, _
                         ByVal stationScope As String, ByVal capabilityMask As Long, ByRef report As String, _
                         Optional ByRef outcome As String = "", Optional ByVal guideChoice As Long = 0) As Boolean
    SaveUser = Execute("SAVE", context, authPath, userId, displayName, credential, warehouseScope, stationScope, capabilityMask, report, outcome, guideChoice)
End Function

Public Function DeactivateUser(ByVal context As String, ByVal authPath As String, _
                               ByVal userId As String, ByRef report As String, Optional ByRef outcome As String = "") As Boolean
    DeactivateUser = Execute("DEACTIVATE", context, authPath, userId, "", "", "", "", 0, report, outcome)
End Function

Public Function DeleteUser(ByVal context As String, ByVal authPath As String, _
                           ByVal userId As String, ByRef report As String, Optional ByRef outcome As String = "") As Boolean
    DeleteUser = Execute("DELETE", context, authPath, userId, "", "", "", "", 0, report, outcome)
End Function

Private Function Execute(ByVal command As String, ByVal context As String, ByVal authPath As String, _
                         ByVal userId As String, ByVal displayName As String, ByVal credential As String, _
                         ByVal warehouseScope As String, ByVal stationScope As String, _
                         ByVal capabilityMask As Long, ByRef report As String, ByRef outcome As String, _
                         Optional ByVal guideChoice As Long = 0) As Boolean
    Dim target As WarehouseTarget, wb As Workbook, candidate As Workbook, opened As Boolean
    Dim users As ListObject, capabilities As ListObject, userFields As Object, capabilityFields As Object
    Dim userRow As Long, row As Long, index As Long, capability As Variant, capabilityRow As Long
    Dim names As Variant, selected As Boolean, actor As String
    On Error GoTo Failed
    outcome = "REJECTED"
    report = "Another Users & Roles command is still running."
    If mWriting Then Exit Function
    mWriting = True
    If Not SameContext(context, report) Then outcome = "DENIED": GoTo CleanExit
    Set target = modNasConnection.GetCurrentTarget()
    If StrComp(Replace$(Trim$(authPath), "/", "\"), Replace$(target.AuthPath, "/", "\"), vbTextCompare) <> 0 Then
        report = "Select the signed-in warehouse's Auth workbook. Reopen Users & Roles after changing warehouses."
        outcome = "DENIED"
        GoTo CleanExit
    End If
    userId = Trim$(userId): displayName = Trim$(displayName)
    warehouseScope = Trim$(warehouseScope): stationScope = Trim$(stationScope)
    If userId = "" Then report = "User ID is required.": GoTo CleanExit
    If command = "SAVE" Then
        If warehouseScope = "" Then report = "Warehouse scope is required.": GoTo CleanExit
        If displayName = "" Then displayName = userId
        If stationScope = "" Then stationScope = "*"
        If capabilityMask < 0 Or capabilityMask > 63 Then report = "Invalid role choices.": GoTo CleanExit
        If guideChoice < 0 Or guideChoice > 2 Then report = "Invalid guide permission choice.": GoTo CleanExit
        If guideChoice <> 0 Then
            If warehouseScope = "*" Or stationScope = "*" Or _
                StrComp(warehouseScope, target.WarehouseId, vbTextCompare) <> 0 Or _
                StrComp(stationScope, target.StationId, vbTextCompare) <> 0 Then
                report = "Maintain Action Paths requires the signed-in warehouse and exact station."
                outcome = "DENIED": GoTo CleanExit
            End If
        End If
    End If
    If Len(Dir$(target.AuthPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) = 0 Then
        report = "Auth workbook is missing. Use explicit warehouse provisioning."
        GoTo CleanExit
    End If
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, target.AuthPath, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(Filename:=target.AuthPath, UpdateLinks:=0, ReadOnly:=False, _
            IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        opened = True
    End If
    If wb.ReadOnly Or Not wb.Saved Then
        report = "Auth workbook is read-only or has unsaved changes. Close it before changing users and roles."
        GoTo CleanExit
    End If
    Set userFields = RequiredFields(wb, "tblUsers", Array("UserId", "DisplayName", "PinHash", "Status", "ValidFrom", "ValidTo"), users)
    Set capabilityFields = RequiredFields(wb, "tblCapabilities", Array("UserId", "Capability", "WarehouseId", "StationId", "Status", "ValidFrom", "ValidTo"), capabilities)
    If userFields Is Nothing Or capabilityFields Is Nothing Then
        report = "Auth workbook has missing or ambiguous required tables or columns. Use explicit provisioning."
        GoTo CleanExit
    End If
    ' Refresh from the exact clean authority; a cached grant cannot authorize a write.
    If Not modAuth.LoadAuth(target.WarehouseId) Then
        report = "Auth workbook is invalid. Users and roles were not changed."
        GoTo CleanExit
    End If
    actor = modAuth.GetCurrentUserId()
    If Not modAuth.CanPerform("ADMIN_MAINT", actor, target.WarehouseId, target.StationId, "USERS_ROLES") Then
        report = "Current user lacks ADMIN_MAINT capability."
        outcome = "DENIED"
        GoTo CleanExit
    End If
    userRow = MatchUser(users, userFields, userId)
    If command = "SAVE" Then
        If credential = "" Then
            If userRow = 0 Then
                report = "PIN/password is required for a new auth user.": GoTo CleanExit
            ElseIf CellText(users, userFields, userRow, "PinHash") = "" Then
                report = "PIN/password is required for a new auth user.": GoTo CleanExit
            End If
        End If
    ElseIf command = "DEACTIVATE" And userRow = 0 Then
        report = "User was not found.": GoTo CleanExit
    End If
    If Not SameContext(context, report) Then outcome = "DENIED": GoTo CleanExit
    If wb.ReadOnly Or Not wb.Saved Then report = "Auth workbook changed before the command could write.": GoTo CleanExit
    Select Case command
        Case "SAVE"
            If userRow = 0 Then userRow = users.ListRows.Add.Index
            PutText users, userFields, userRow, "UserId", userId
            PutText users, userFields, userRow, "DisplayName", displayName
            If credential <> "" Then modAuth.WriteUserCredential users.DataBodyRange.Cells(userRow, userFields("PinHash")), credential
            PutText users, userFields, userRow, "Status", "ACTIVE"
            names = Array("ADMIN_MAINT", "RECEIVE_POST", "RECEIVE_VIEW", "SHIP_POST", "PROD_POST", "INBOX_PROCESS")
            If guideChoice <> 0 Then
                ReDim Preserve names(6)
                names(6) = "ACTION_PATH_MAINT"
                If guideChoice = 1 Then capabilityMask = capabilityMask Or 64
            End If
            For index = 0 To UBound(names)
                capability = names(index): selected = ((capabilityMask And CLng(2 ^ index)) <> 0)
                capabilityRow = MatchCapability(capabilities, capabilityFields, userId, CStr(capability), warehouseScope, stationScope)
                If selected Then
                    If capabilityRow = 0 Then capabilityRow = capabilities.ListRows.Add.Index
                    PutText capabilities, capabilityFields, capabilityRow, "UserId", userId
                    PutText capabilities, capabilityFields, capabilityRow, "Capability", CStr(capability)
                    PutText capabilities, capabilityFields, capabilityRow, "WarehouseId", warehouseScope
                    PutText capabilities, capabilityFields, capabilityRow, "StationId", stationScope
                    PutText capabilities, capabilityFields, capabilityRow, "Status", "ACTIVE"
                ElseIf capabilityRow > 0 Then
                    PutText capabilities, capabilityFields, capabilityRow, "Status", "INACTIVE"
                End If
            Next index
        Case "DEACTIVATE": PutText users, userFields, userRow, "Status", "INACTIVE"
        Case "DELETE"
            If userRow > 0 Then users.ListRows(userRow).Delete
            For row = capabilities.ListRows.Count To 1 Step -1
                If StrComp(CellText(capabilities, capabilityFields, row, "UserId"), userId, vbTextCompare) = 0 Then capabilities.ListRows(row).Delete
            Next row
    End Select
    If Not SameContext(context, report) Then outcome = "DENIED": GoTo CleanExit
    wb.Save
    If Not wb.Saved Then Err.Raise vbObjectError + 7210, , "Auth persistence was not confirmed."
    Execute = True
    outcome = "SAVED"
    report = ""
    ' Invalidate cached role presentation only after actual persistence.
    On Error Resume Next
    modAuth.ReloadAuth target.WarehouseId
    modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
    On Error GoTo Failed
CleanExit:
    On Error Resume Next
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
    mWriting = False
    Exit Function
Failed:
    outcome = "FAILED"
    report = "Users & Roles could not persist the command. Inspect the current Auth state before retrying."
    Execute = False
    Resume CleanExit
End Function

Private Function SameContext(ByVal context As String, ByRef report As String) As Boolean
    report = "The invSys session or warehouse changed. Reopen Users & Roles."
    If context = "" Then Exit Function
    SameContext = (context = CaptureContext())
    If SameContext Then report = ""
End Function

Private Function RequiredFields(ByVal wb As Workbook, ByVal tableName As String, ByVal required As Variant, _
                                ByRef table As ListObject) As Object
    Dim sheet As Worksheet, candidate As ListObject, column As ListColumn, fields As Object, name As Variant, key As String
    For Each sheet In wb.Worksheets
        For Each candidate In sheet.ListObjects
            If StrComp(candidate.Name, tableName, vbTextCompare) = 0 Then
                If Not table Is Nothing Then Exit Function
                Set table = candidate
            End If
        Next candidate
    Next sheet
    If table Is Nothing Then Exit Function
    Set fields = CreateObject("Scripting.Dictionary"): fields.CompareMode = vbTextCompare
    For Each name In required
        For Each column In table.ListColumns
            key = Trim$(column.Name)
            If StrComp(key, CStr(name), vbTextCompare) = 0 Then
                If fields.Exists(CStr(name)) Then Exit Function
                fields.Add CStr(name), column.Index
            End If
        Next column
        If Not fields.Exists(CStr(name)) Then Exit Function
    Next name
    Set RequiredFields = fields
End Function

Private Function CellText(ByVal table As ListObject, ByVal fields As Object, ByVal row As Long, ByVal name As String) As String
    Dim value As Variant
    value = table.DataBodyRange.Cells(row, CLng(fields(name))).Value2
    If Not IsError(value) And Not IsNull(value) And Not IsEmpty(value) Then CellText = Trim$(CStr(value))
End Function

Private Sub PutText(ByVal table As ListObject, ByVal fields As Object, ByVal row As Long, ByVal name As String, ByVal value As String)
    table.DataBodyRange.Cells(row, CLng(fields(name))).Value2 = value
End Sub

Private Function MatchUser(ByVal table As ListObject, ByVal fields As Object, ByVal userId As String) As Long
    Dim row As Long
    For row = 1 To table.ListRows.Count
        If StrComp(CellText(table, fields, row, "UserId"), userId, vbTextCompare) = 0 Then MatchUser = row: Exit Function
    Next row
End Function

Private Function MatchCapability(ByVal table As ListObject, ByVal fields As Object, ByVal userId As String, _
                                 ByVal capability As String, ByVal warehouse As String, ByVal station As String) As Long
    Dim row As Long
    For row = 1 To table.ListRows.Count
        If StrComp(CellText(table, fields, row, "UserId"), userId, vbTextCompare) = 0 And _
           StrComp(CellText(table, fields, row, "Capability"), capability, vbTextCompare) = 0 And _
           StrComp(CellText(table, fields, row, "WarehouseId"), warehouse, vbTextCompare) = 0 And _
           StrComp(CellText(table, fields, row, "StationId"), station, vbTextCompare) = 0 Then
            MatchCapability = row: Exit Function
        End If
    Next row
End Function
