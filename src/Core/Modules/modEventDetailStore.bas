Attribute VB_Name = "modEventDetailStore"
Option Explicit
Option Private Module

Public Function ReadProfile(ByVal target As WarehouseTarget, ByRef version As Long, ByRef request As String, ByRef report As String) As Boolean
    Dim wb As Workbook, headers As ListObject, fields As ListObject, opened As Boolean, row As Long
    Dim seen As Object, latest As Long, selected As Long, value As Variant, model As Object, rows As Collection, entry As Object, catalog As Object
    On Error GoTo Failed
    version = 0
    request = modEventDetailSettings.DefaultRequest()
    report = "Detail profile unavailable; built-in display defaults apply. No configuration was changed."
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then Exit Function
    Set wb = ConfigWorkbook(target, True, opened)
    If Not opened And Not wb.Saved Then GoTo CleanExit
    FindTables wb, headers, fields
    If headers Is Nothing And fields Is Nothing Then
        ReadProfile = True
        report = "Built-in detail defaults; no profile saved."
        GoTo CleanExit
    End If
    If headers Is Nothing Or fields Is Nothing Then GoTo CleanExit
    Set seen = CreateObject("Scripting.Dictionary")
    For row = 1 To headers.ListRows.Count
        value = modTrackingPolicyModel.TableValue(headers, row, "ProfileVersion")
        If Not modEventDetailModel.IntegerInRange(value, 1, 2147483647) Then GoTo CleanExit
        If seen.Exists(CStr(value)) Then GoTo CleanExit
        seen.Add CStr(value), True
        If value > latest Then latest = CLng(value): selected = row
    Next row
    If selected = 0 Then GoTo CleanExit
    If Not modEventDetailModel.IntegerInRange(modTrackingPolicyModel.TableValue(headers, selected, "SchemaVersion"), 1, 1) Then GoTo CleanExit
    If Not modTrainingWire.ValidUtcTimestamp(CStr(modTrackingPolicyModel.TableValue(headers, selected, "CreatedAtUTC"))) Then GoTo CleanExit
    If CStr(modTrackingPolicyModel.TableValue(headers, selected, "CreatedByUserId")) = "" Then GoTo CleanExit
    Set model = CreateObject("Scripting.Dictionary")
    model.Add "SchemaVersion", 1
    Set rows = New Collection
    Set catalog = modEventDetailCatalog.Fields()
    For row = 1 To fields.ListRows.Count
        value = modTrackingPolicyModel.TableValue(fields, row, "ProfileVersion")
        If Not modEventDetailModel.IntegerInRange(value, 1, 2147483647) Then GoTo CleanExit
        If Not seen.Exists(CStr(value)) Then GoTo CleanExit
        If value = latest Then
            Set entry = CreateObject("Scripting.Dictionary")
            For Each value In Array("EventFamily", "FieldId", "Enabled", "DisplayOrder")
                entry.Add CStr(value), modTrackingPolicyModel.TableValue(fields, row, CStr(value))
            Next value
            ' Excel Value2 numbers are Double; validate before integer wire conversion.
            If Not modEventDetailModel.IntegerInRange(entry("DisplayOrder"), 1, catalog.Count) Then GoTo CleanExit
            entry("DisplayOrder") = CLng(entry("DisplayOrder"))
            rows.Add entry
        End If
    Next row
    model.Add "Fields", rows
    Set model = modEventDetailModel.Decode(modTrainingJson.EncodeObject(model))
    If model Is Nothing Then GoTo CleanExit
    version = latest
    request = modTrainingJson.EncodeObject(model)
    report = "Detail profile version " & CStr(version) & " loaded."
    ReadProfile = True
CleanExit:
    On Error Resume Next
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Exit Function
Failed:
    ReadProfile = False
    Resume CleanExit
End Function

Public Function ConfigWorkbook(ByVal target As WarehouseTarget, ByVal readOnly As Boolean, ByRef opened As Boolean) As Workbook
    Dim candidate As Workbook
    opened = False
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, target.ConfigPath, vbTextCompare) = 0 Then
            Set ConfigWorkbook = candidate
            Exit Function
        End If
    Next candidate
    Set ConfigWorkbook = Application.Workbooks.Open(target.ConfigPath, UpdateLinks:=0, ReadOnly:=readOnly, AddToMru:=False)
    opened = True
End Function

Public Sub FindTables(ByVal wb As Workbook, ByRef headers As ListObject, ByRef fields As ListObject)
    Dim sheet As Worksheet, table As ListObject
    For Each sheet In wb.Worksheets
        For Each table In sheet.ListObjects
            If StrComp(table.Name, "tblEventDetailProfiles", vbTextCompare) = 0 Then Set headers = table
            If StrComp(table.Name, "tblEventDetailFields", vbTextCompare) = 0 Then Set fields = table
        Next table
    Next sheet
End Sub
