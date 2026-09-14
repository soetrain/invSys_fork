Attribute VB_Name = "modTrackingPolicyCommand"
Option Explicit
Option Private Module

Public Function Save(ByVal context As String, ByVal expectedVersion As Long, _
                     ByVal request As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget, model As Object, wb As Workbook, candidate As Workbook
    Dim headers As ListObject, controls As ListObject, headerRow As ListRow, controlRow As ListRow
    Dim created As Collection, sheet As Worksheet, row As Variant, field As Variant
    Dim version As Long, collect As Boolean, visible As Boolean, notice As String
    Dim headerCount As Long, controlCount As Long, opened As Boolean, changed As Boolean, saved As Boolean
    Dim index As Long
    On Error GoTo Failed
    If Not modTrackingPolicyModel.AuthorizeEditor(context, target, report) Then Exit Function
    report = "Invalid tracking policy request. No configuration was changed."
    Set model = modTrackingPolicyModel.Decode(request)
    If model Is Nothing Or expectedVersion < 0 Then Exit Function
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, target.ConfigPath, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        If Len(Dir$(target.ConfigPath)) = 0 Then report = "Configuration workbook is missing.": Exit Function
        Set wb = Application.Workbooks.Open(target.ConfigPath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
        opened = True
    End If
    If wb.ReadOnly Or Not wb.Saved Then
        report = "Configuration is read-only, locked, or has unsaved changes. Close it before saving Tracking Policy."
        GoTo CleanExit
    End If
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice) Then
        report = "Tracking policy or required configuration is invalid. No configuration was changed."
        GoTo CleanExit
    End If
    If expectedVersion <> version Then
        report = "Tracking policy changed. Reload before saving."
        GoTo CleanExit
    End If
    If version = 2147483647 Then report = "Tracking policy version limit reached.": GoTo CleanExit
    Set headers = FindPolicyTable(wb, "tblEventTrackingPolicies")
    Set controls = FindPolicyTable(wb, "tblEventTrackingControls")
    Set created = New Collection
    If headers Is Nothing Then
        If wb.ProtectStructure Then report = "Configuration workbook structure is protected.": GoTo CleanExit
    Else
        If headers.Parent.ProtectContents Or controls.Parent.ProtectContents Then
            report = "Tracking policy worksheet is protected.": GoTo CleanExit
        End If
        headerCount = headers.ListRows.Count
        controlCount = controls.ListRows.Count
    End If
    ' Workbook opening/validation may run Excel callbacks. Recheck before mutation.
    If Not modTrackingPolicyModel.AuthorizeEditor(context, target, report) Then GoTo CleanExit
    modRecordingSession.InterruptContext context, "POLICY_CHANGED"
    changed = True
    If headers Is Nothing Then
        Set headers = CreatePolicyTable(wb, "tblEventTrackingPolicies", Array("PolicyVersion", "SchemaVersion", _
            "CatalogVersion", "CreatedAtUTC", "CreatedByUserId", "DefaultView", _
            "ViewerActionPathCaptureEnabled", "AdminViewerEventLoggingEnabled"), created)
        Set controls = CreatePolicyTable(wb, "tblEventTrackingControls", _
            Array("PolicyVersion", "ControlId", "Collect", "Visible", "SequenceEligible"), created)
    End If
    Set headerRow = headers.ListRows.Add
    SetPolicyCell headers, headerRow, "PolicyVersion", version + 1
    SetPolicyCell headers, headerRow, "SchemaVersion", 1
    SetPolicyCell headers, headerRow, "CatalogVersion", model("CatalogVersion")
    SetPolicyCell headers, headerRow, "CreatedAtUTC", modTrainingWire.UtcTimestamp()
    SetPolicyCell headers, headerRow, "CreatedByUserId", modAuth.GetCurrentUserId()
    For Each field In Array("DefaultView", "ViewerActionPathCaptureEnabled", "AdminViewerEventLoggingEnabled")
        SetPolicyCell headers, headerRow, CStr(field), model(field)
    Next field
    For Each row In model("Controls")
        Set controlRow = controls.ListRows.Add
        SetPolicyCell controls, controlRow, "PolicyVersion", version + 1
        For Each field In Array("ControlId", "Collect", "Visible", "SequenceEligible")
            SetPolicyCell controls, controlRow, CStr(field), row(field)
        Next field
    Next row
    wb.Save
    ' Excel BeforeSave can cancel without raising an error.
    If Not wb.Saved Then Err.Raise 5
    saved = True
    Save = True
    report = "Tracking policy version " & CStr(version + 1) & " saved."
CleanExit:
    On Error Resume Next
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Exit Function
Failed:
    On Error Resume Next
    If changed And Not saved Then
        If Not controls Is Nothing Then
            For index = controls.ListRows.Count To controlCount + 1 Step -1
                controls.ListRows(index).Delete
            Next index
        End If
        If Not headers Is Nothing Then
            For index = headers.ListRows.Count To headerCount + 1 Step -1
                headers.ListRows(index).Delete
            Next index
        End If
        If Not created Is Nothing Then
            Dim alerts As Boolean
            alerts = Application.DisplayAlerts
            Application.DisplayAlerts = False
            For Each sheet In created
                sheet.Delete
            Next sheet
            Application.DisplayAlerts = alerts
        End If
    End If
    report = "Tracking policy save could not be verified. Reload before retrying."
    Save = False
    Resume CleanExit
End Function

Private Function FindPolicyTable(ByVal wb As Workbook, ByVal name As String) As ListObject
    Dim sheet As Worksheet, table As ListObject
    For Each sheet In wb.Worksheets
        For Each table In sheet.ListObjects
            If StrComp(table.Name, name, vbTextCompare) = 0 Then Set FindPolicyTable = table: Exit Function
        Next table
    Next sheet
End Function

Private Function CreatePolicyTable(ByVal wb As Workbook, ByVal name As String, ByVal fields As Variant, _
                                   ByVal created As Collection) As ListObject
    Set CreatePolicyTable = modEventSettingsTables.CreateTable(wb, name, fields, created)
End Function

Private Sub SetPolicyCell(ByVal table As ListObject, ByVal row As ListRow, ByVal header As String, ByVal value As Variant)
    modEventSettingsTables.WriteCell table, row, header, value
End Sub
