Attribute VB_Name = "modExecutionTarget"
Option Explicit
Option Private Module

' Replay accepts the generated runtime's own authority paths. Reads never repair it.
Public Function Read(ByVal context As String, ByRef target As WarehouseTarget, ByRef snapshot As String, ByRef notice As String) As Boolean
    Dim fso As Object, wb As Workbook, opened As Boolean, warehouse As ListObject, station As ListObject
    Dim root As String, suffix As Variant, path As String, row As Long, foundWarehouse As Long, foundStation As Long
    On Error GoTo Failed
    snapshot = "": notice = "Run requires the captured generated Training warehouse and its own authority files."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If Not modNasConnection.IsWarehouseTargetAllowed(target, True) Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = fso.GetAbsolutePathName(target.RuntimeRoot)
    If Not fso.FolderExists(root) Then Exit Function
    If (fso.GetFolder(root).Attributes And &H400) <> 0 Then Exit Function
    If StrComp(fso.BuildPath(root, target.WarehouseId & ".invSys.Config.xlsb"), target.ConfigPath, vbTextCompare) <> 0 Then Exit Function
    For Each suffix In Array(".invSys.Config.xlsb", ".invSys.Auth.xlsb", ".invSys.Data.Inventory.xlsb", ".invSys.Snapshot.Inventory.xlsb")
        path = fso.BuildPath(root, target.WarehouseId & CStr(suffix))
        If Not fso.FileExists(path) Then Exit Function
        If (fso.GetFile(path).Attributes And &H400) <> 0 Then Exit Function
    Next suffix
    Set wb = OpenRead(target.ConfigPath, opened)
    If wb Is Nothing Then Exit Function
    Set warehouse = Table(wb, "tblWarehouseConfig"): Set station = Table(wb, "tblStationConfig")
    For row = 1 To warehouse.ListRows.Count
        If CStr(Value(warehouse, row, "WarehouseId")) = target.WarehouseId Then
            foundWarehouse = foundWarehouse + 1
            If CStr(Value(warehouse, row, "WarehousePurpose")) <> "Training" Then GoTo Failed
            If StrComp(fso.GetAbsolutePathName(CStr(Value(warehouse, row, "PathDataRoot"))), root, vbTextCompare) <> 0 Then GoTo Failed
        End If
    Next row
    For row = 1 To station.ListRows.Count
        If CStr(Value(station, row, "WarehouseId")) = target.WarehouseId And CStr(Value(station, row, "StationId")) = target.StationId Then
            foundStation = foundStation + 1
            path = fso.GetAbsolutePathName(CStr(Value(station, row, "PathInboxRoot")))
            If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
            If StrComp(path, fso.BuildPath(root, "inbox"), vbTextCompare) <> 0 Then GoTo Failed
        End If
    Next row
    If foundWarehouse <> 1 Or foundStation <> 1 Or context <> modActivity.CaptureContext() Then GoTo Failed
    snapshot = fso.BuildPath(root, target.WarehouseId & ".invSys.Snapshot.Inventory.xlsb")
    Read = True: notice = ""
Failed:
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

Public Function OpenRead(ByVal path As String, ByRef opened As Boolean) As Workbook
    Dim wb As Workbook, window As window
    opened = False
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            If wb.Saved Then Set OpenRead = wb
            Exit Function
        End If
    Next wb
    Set wb = Application.Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    opened = True
    For Each window In wb.Windows: window.Visible = False: Next window
    Set OpenRead = wb
End Function

Public Function Table(ByVal wb As Workbook, ByVal name As String) As ListObject
    Dim sheet As Worksheet, candidate As ListObject
    For Each sheet In wb.Worksheets
        For Each candidate In sheet.ListObjects
            If StrComp(candidate.Name, name, vbTextCompare) = 0 Then Set Table = candidate: Exit Function
        Next candidate
    Next sheet
End Function

Private Function Value(ByVal table As ListObject, ByVal row As Long, ByVal header As String) As Variant
    Dim column As ListColumn, index As Long
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), header, vbTextCompare) = 0 Then
            If index <> 0 Then Err.Raise 5, , "Ambiguous execution source header."
            index = column.Index
        End If
    Next column
    If index = 0 Then Err.Raise 5, , "Missing execution source header."
    Value = table.DataBodyRange.Cells(row, index).Value2
End Function

Public Function ContainsEntity(ByVal context As String, ByVal key As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, path As String, wb As Workbook, opened As Boolean, table As ListObject, row As Long, matches As Long
    On Error GoTo Failed
    If key = "" Or Not Read(context, target, path, notice) Then Exit Function
    Set wb = OpenRead(path, opened)
    Set table = modExecutionTarget.Table(wb, "tblInventorySnapshot")
    For row = 1 To table.ListRows.Count
        If CStr(Value(table, row, "System_Key")) = key Then
            If CStr(Value(table, row, "WarehouseId")) <> target.WarehouseId Then GoTo Failed
            matches = matches + 1
        End If
    Next row
    ContainsEntity = (matches = 1 And context = modActivity.CaptureContext())
Failed:
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not ContainsEntity Then notice = "Select one exact available entity in this Training warehouse."
End Function
