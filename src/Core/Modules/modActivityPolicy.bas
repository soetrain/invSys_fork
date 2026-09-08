Attribute VB_Name = "modActivityPolicy"
Option Explicit
Option Private Module

Public Function ReadPolicy(ByVal target As WarehouseTarget, ByVal controlId As String, _
                           ByRef version As Long, ByRef collect As Boolean, _
                           ByRef visible As Boolean, ByRef notice As String) As Boolean
    Dim wb As Workbook, candidate As Workbook, headers As ListObject, controls As ListObject
    Dim opened As Boolean, row As Long, selected As Long, number As Long, latest As Long
    Dim seen As Object, rowsSeen As Object, key As String, definition As Object, ids As Variant
    Dim catalogVersion As Long
    On Error GoTo Failed
    collect = False: visible = False: version = 0
    notice = "Tracking unavailable: configuration could not be validated."
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then Exit Function
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, target.ConfigPath, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(target.ConfigPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        opened = True
    End If
    ' Excel can mark a freshly opened read-only workbook dirty during calculation.
    ' It came from the saved file; only a pre-existing dirty workbook has edits
    ' this reader cannot distinguish from the persisted policy.
    If Not opened Then
        If Not wb.Saved Then GoTo CleanExit
    End If
    Set headers = FindTable(wb, "tblEventTrackingPolicies")
    Set controls = FindTable(wb, "tblEventTrackingControls")
    If headers Is Nothing And controls Is Nothing Then
        Set definition = modActivityCatalog.Control(controlId)
        If definition Is Nothing Then GoTo CleanExit
        collect = (definition("Class") = "Command"): visible = True
        ReadPolicy = True: notice = ""
        GoTo CleanExit
    End If
    notice = "Tracking unavailable: the saved tracking policy is invalid."
    If headers Is Nothing Or controls Is Nothing Then GoTo CleanExit
    Set seen = CreateObject("Scripting.Dictionary")
    For row = 1 To headers.ListRows.Count
        number = PositiveInteger(CellValue(headers, row, "PolicyVersion"))
        If number = 0 Or seen.Exists(CStr(number)) Then GoTo CleanExit
        seen.Add CStr(number), True
        If number > latest Then latest = number: selected = row
    Next row
    If selected = 0 Then GoTo CleanExit
    If PositiveInteger(CellValue(headers, selected, "SchemaVersion")) <> 1 Then GoTo CleanExit
    catalogVersion = PositiveInteger(CellValue(headers, selected, "CatalogVersion"))
    If catalogVersion < 1 Or catalogVersion > modActivityCatalog.CATALOG_VERSION Then GoTo CleanExit
    key = CStr(CellValue(headers, selected, "DefaultView"))
    If key <> "How-To" And key <> "Diagnostic" And key <> "Compare both" Then GoTo CleanExit
    If CStr(CellValue(headers, selected, "CreatedByUserId")) = "" Then GoTo CleanExit
    key = CStr(CellValue(headers, selected, "CreatedAtUTC"))
    If Not modTrainingWire.ValidUtcTimestamp(key) Then GoTo CleanExit
    If Not IsBooleanValue(CellValue(headers, selected, "ViewerActionPathCaptureEnabled")) Then GoTo CleanExit
    If Not IsBooleanValue(CellValue(headers, selected, "AdminViewerEventLoggingEnabled")) Then GoTo CleanExit
    Set rowsSeen = CreateObject("Scripting.Dictionary")
    For row = 1 To controls.ListRows.Count
        number = PositiveInteger(CellValue(controls, row, "PolicyVersion"))
        If number = 0 Or Not seen.Exists(CStr(number)) Then GoTo CleanExit
        If number = latest Then
            key = CStr(CellValue(controls, row, "ControlId"))
            If rowsSeen.Exists(key) Then GoTo CleanExit
            Set definition = modActivityCatalog.Control(key, catalogVersion)
            If definition Is Nothing Then GoTo CleanExit
            rowsSeen.Add key, True
            If Not IsBooleanValue(CellValue(controls, row, "Collect")) Then GoTo CleanExit
            If Not IsBooleanValue(CellValue(controls, row, "Visible")) Then GoTo CleanExit
            If Not IsBooleanValue(CellValue(controls, row, "SequenceEligible")) Then GoTo CleanExit
            If key = controlId Then
                collect = CBool(CellValue(controls, row, "Collect"))
                visible = CBool(CellValue(controls, row, "Visible"))
                If definition("Role") = "Admin" Then visible = visible And CBool(CellValue(headers, selected, "AdminViewerEventLoggingEnabled"))
            End If
        End If
    Next row
    ' A saved profile must explicitly cover every registered control.
    ids = modActivityCatalog.ControlIds(catalogVersion)
    If rowsSeen.Count <> UBound(ids) - LBound(ids) + 1 Then GoTo CleanExit
    If Not rowsSeen.Exists(controlId) Then
        notice = "Tracking unavailable: the saved policy does not include this control."
        GoTo CleanExit
    End If
    version = latest
    ReadPolicy = True: notice = ""
CleanExit:
    If Not ReadPolicy Then collect = False: visible = False
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Exit Function
Failed:
    ReadPolicy = False
    Resume CleanExit
End Function

Private Function FindTable(ByVal wb As Workbook, ByVal name As String) As ListObject
    Dim ws As Worksheet, table As ListObject
    For Each ws In wb.Worksheets
        For Each table In ws.ListObjects
            If StrComp(table.Name, name, vbTextCompare) = 0 Then Set FindTable = table: Exit Function
        Next table
    Next ws
End Function

Private Function CellValue(ByVal table As ListObject, ByVal row As Long, ByVal name As String) As Variant
    Dim column As ListColumn, index As Long
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), name, vbTextCompare) = 0 Then
            If index > 0 Then Err.Raise 5
            index = column.Index
        End If
    Next column
    If index = 0 Then Err.Raise 5
    CellValue = table.DataBodyRange.Cells(row, index).Value2
End Function

Private Function PositiveInteger(ByVal value As Variant) As Long
    On Error GoTo Invalid
    If VarType(value) = vbBoolean Or Not IsNumeric(value) Then Exit Function
    If CDbl(value) < 1 Or CDbl(value) <> Fix(CDbl(value)) Then Exit Function
    PositiveInteger = CLng(value)
Invalid:
End Function

Private Function IsBooleanValue(ByVal value As Variant) As Boolean
    IsBooleanValue = (VarType(value) = vbBoolean)
End Function
