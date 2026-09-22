Attribute VB_Name = "modEventDetailCommand"
Option Explicit
Option Private Module

Public Function Save(ByVal context As String, ByVal expectedVersion As Long, ByVal request As String, ByRef report As String, Optional ByRef outcome As String = "") As Boolean
    Dim target As WarehouseTarget, model As Object, wb As Workbook, headers As ListObject, fields As ListObject
    Dim created As Collection, row As ListRow, entry As Variant, key As Variant, sheet As Worksheet
    Dim opened As Boolean, changed As Boolean, persisted As Boolean, alerts As Boolean
    Dim version As Long, headerCount As Long, fieldCount As Long, index As Long, priorRequest As String, notice As String
    On Error GoTo Failed
    outcome = "DENIED"
    If Not modTrackingPolicyModel.AuthorizeEditor(context, target, report) Then Exit Function
    outcome = "REJECTED"
    report = "Invalid detail profile request. No configuration was changed."
    Set model = modEventDetailModel.Decode(request)
    If model Is Nothing Or expectedVersion < 0 Then Exit Function
    If Len(Dir$(target.ConfigPath)) = 0 Then report = "Configuration workbook is missing.": Exit Function
    Set wb = modEventDetailStore.ConfigWorkbook(target, False, opened)
    If wb.ReadOnly Or Not wb.Saved Then
        report = "Configuration is read-only, locked, or has unsaved changes. Close it before saving the detail profile."
        GoTo CleanExit
    End If
    If Not modEventDetailStore.ReadProfile(target, version, priorRequest, notice) Then
        report = "Detail profile or required configuration is invalid. No configuration was changed."
        GoTo CleanExit
    End If
    If expectedVersion <> version Then report = "Detail profile changed. Reload before saving.": GoTo CleanExit
    If version = 2147483647 Then report = "Detail profile version limit reached.": GoTo CleanExit
    modEventDetailStore.FindTables wb, headers, fields
    Set created = New Collection
    If headers Is Nothing Then
        If wb.ProtectStructure Then report = "Configuration workbook structure is protected.": GoTo CleanExit
    Else
        If headers.Parent.ProtectContents Or fields.Parent.ProtectContents Then
            report = "Detail profile worksheet is protected."
            GoTo CleanExit
        End If
        headerCount = headers.ListRows.Count: fieldCount = fields.ListRows.Count
    End If
    outcome = "DENIED"
    If Not modTrackingPolicyModel.AuthorizeEditor(context, target, report) Then GoTo CleanExit
    outcome = "FAILED"
    changed = True
    If headers Is Nothing Then
        Set headers = modEventSettingsTables.CreateTable(wb, "tblEventDetailProfiles", _
            Array("ProfileVersion", "SchemaVersion", "CreatedAtUTC", "CreatedByUserId"), created)
        Set fields = modEventSettingsTables.CreateTable(wb, "tblEventDetailFields", _
            Array("ProfileVersion", "EventFamily", "FieldId", "Enabled", "DisplayOrder"), created)
    End If
    Set row = headers.ListRows.Add
    modEventSettingsTables.WriteCell headers, row, "ProfileVersion", version + 1
    modEventSettingsTables.WriteCell headers, row, "SchemaVersion", 1
    modEventSettingsTables.WriteCell headers, row, "CreatedAtUTC", modTrainingWire.UtcTimestamp()
    modEventSettingsTables.WriteCell headers, row, "CreatedByUserId", modAuth.GetCurrentUserId()
    For Each entry In model("Fields")
        Set row = fields.ListRows.Add
        modEventSettingsTables.WriteCell fields, row, "ProfileVersion", version + 1
        For Each key In Array("EventFamily", "FieldId", "Enabled", "DisplayOrder")
            modEventSettingsTables.WriteCell fields, row, CStr(key), entry(key)
        Next key
    Next entry
    wb.Save
    If Not wb.Saved Then Err.Raise 5
    persisted = True
    Save = True
    outcome = "COMPLETED"
    report = "Detail profile version " & CStr(version + 1) & " saved."
CleanExit:
    On Error Resume Next
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Exit Function
Failed:
    outcome = "FAILED"
    On Error Resume Next
    If changed And Not persisted Then
        If Not fields Is Nothing Then
            For index = fields.ListRows.Count To fieldCount + 1 Step -1
                fields.ListRows(index).Delete
            Next index
        End If
        If Not headers Is Nothing Then
            For index = headers.ListRows.Count To headerCount + 1 Step -1
                headers.ListRows(index).Delete
            Next index
        End If
        If Not created Is Nothing Then
            alerts = Application.DisplayAlerts: Application.DisplayAlerts = False
            For Each sheet In created
                sheet.Delete
            Next sheet
            Application.DisplayAlerts = alerts
        End If
    End If
    report = "Detail profile save could not be verified. Reload before retrying."
    Save = False
    Resume CleanExit
End Function
