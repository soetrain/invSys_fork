Attribute VB_Name = "modAggregationSources"
Option Explicit
Option Private Module

Public Function RequiredColumns(ByVal table As ListObject, ByVal required As Variant) As Object
    Dim fields As Object, name As Variant, column As ListColumn, index As Long
    If table Is Nothing Then Exit Function
    Set fields = CreateObject("Scripting.Dictionary"): fields.CompareMode = vbTextCompare
    For Each column In table.ListColumns
        If UCase$(Trim$(column.Name)) = "ROW" Then Exit Function
    Next column
    For Each name In required
        index = 0
        For Each column In table.ListColumns
            If StrComp(Trim$(column.Name), CStr(name), vbTextCompare) = 0 Then
                If index <> 0 Then Exit Function
                index = column.Index
            End If
        Next column
        If index = 0 Then Exit Function
        fields.Add CStr(name), index
    Next name
    Set RequiredColumns = fields
End Function

Public Function Feed(ByVal configPath As String, ByVal warehouse As String, ByRef raw As String) As String
    Dim book As Workbook, copy As String, table As ListObject, columns As Object, row As Long, matches As Long
    On Error GoTo Finished
    raw = ""
    If modAggregationFiles.Dirty(configPath) Then Exit Function
    Set book = modAggregationFiles.OpenCopy(configPath, copy)
    If book Is Nothing Then GoTo Finished
    Set table = modExecutionTarget.Table(book, "tblWarehouseConfig")
    Set columns = RequiredColumns(table, Array("WarehouseId", "PathSharePointRoot"))
    If columns Is Nothing Then GoTo Finished
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, columns("WarehouseId")).Value2) = warehouse Then
            matches = matches + 1
            raw = CStr(table.DataBodyRange.Cells(row, columns("PathSharePointRoot")).Value2)
        End If
    Next row
    If matches = 1 And raw <> "" Then Feed = modAggregationFiles.Canonical(raw)
Finished:
    modAggregationFiles.CloseCopy book, copy
End Function

Public Function Inspect(ByVal warehouse As String, ByVal configPath As String, ByVal snapshot As String) As Object
    Dim record As Object, rows As Object, feedPath As String, raw As String, hash As String, fso As Object
    On Error GoTo Invalid
    snapshot = modAggregationFiles.Canonical(snapshot): configPath = modAggregationFiles.Canonical(configPath)
    If snapshot = "" Or configPath = "" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    feedPath = Feed(configPath, warehouse, raw)
    If feedPath = "" Then Exit Function
    If StrComp(fso.GetParentFolderName(snapshot), modAggregationFiles.Canonical(feedPath & "\Snapshots"), vbTextCompare) <> 0 Then Exit Function
    If Not LCase$(fso.GetFileName(snapshot)) Like LCase$(warehouse & ".invSys.Snapshot.Inventory.xls*") Then Exit Function
    hash = modTrainingWire.FileSha256(snapshot)
    If hash = "" Or modAggregationFiles.Dirty(snapshot) Then Exit Function
    Set rows = ReadRows(snapshot, "tblInventorySnapshot", warehouse)
    If rows Is Nothing Then Exit Function
    If hash <> modTrainingWire.FileSha256(snapshot) Then Exit Function
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "Warehouse", warehouse: record.Add "Config", configPath
    record.Add "Feed", raw: record.Add "Path", snapshot: record.Add "Hash", hash
    record.Add "Stamp", CStr(fso.GetFile(snapshot).DateLastModified)
    record.Add "Rows", rows
    Set Inspect = record
Invalid:
End Function

Public Function Unchanged(ByVal record As Object) As Boolean
    Dim current As Object
    On Error GoTo Invalid
    Set current = Inspect(CStr(record("Warehouse")), CStr(record("Config")), CStr(record("Path")))
    If current Is Nothing Then Exit Function
    Unchanged = (current("Hash") = record("Hash") And current("Stamp") = record("Stamp") And current("Feed") = record("Feed"))
Invalid:
End Function

Public Function ReadRows(ByVal path As String, ByVal tableName As String, Optional ByVal warehouse As String = "") As Object
    Dim book As Workbook, copy As String, table As ListObject, columns As Object, rows As Object, entry As Object
    Dim names As Variant, name As Variant, row As Long, key As String, wh As String, identity As String, populated As Boolean
    On Error GoTo Finished
    Set book = modAggregationFiles.OpenCopy(path, copy)
    If book Is Nothing Then GoTo Finished
    Set table = modExecutionTarget.Table(book, tableName)
    names = Array("WarehouseId", "System_Key", "SKU", "QtyOnHand", "LastAppliedAtUTC")
    If tableName = "tblGlobalInventorySnapshot" Then names = Array("WarehouseId", "System_Key", "SKU", "QtyOnHand", "LastAppliedAtUTC", "SourceSnapshot")
    Set columns = RequiredColumns(table, names)
    If columns Is Nothing Then GoTo Finished
    Set rows = CreateObject("Scripting.Dictionary"): rows.CompareMode = vbBinaryCompare
    For row = 1 To table.ListRows.Count
        key = CStr(table.DataBodyRange.Cells(row, columns("System_Key")).Value2)
        wh = CStr(table.DataBodyRange.Cells(row, columns("WarehouseId")).Value2)
        populated = False
        For Each name In names
            If CStr(table.DataBodyRange.Cells(row, columns(CStr(name))).Value2) <> "" Then populated = True
        Next name
        If populated Then
            If key = "" Or key <> Trim$(key) Or wh = "" Then GoTo Finished
            If warehouse <> "" And wh <> warehouse Then GoTo Finished
            If Not IsNumeric(table.DataBodyRange.Cells(row, columns("QtyOnHand")).Value2) Then GoTo Finished
            identity = wh & "|" & key
            If rows.Exists(identity) Then GoTo Finished
            Set entry = CreateObject("Scripting.Dictionary")
            For Each name In names
                entry.Add CStr(name), table.DataBodyRange.Cells(row, columns(CStr(name))).Value2
            Next name
            If CStr(entry("SKU")) = "" Then GoTo Finished
            rows.Add identity, entry
        End If
    Next row
    Set ReadRows = rows
Finished:
    modAggregationFiles.CloseCopy book, copy
End Function

Public Function Complete(ByVal path As String, ByVal sources As Collection) As Boolean
    Dim actual As Object, expected As Object, record As Variant, key As Variant, name As Variant, entry As Object
    Dim warehouses As Object
    On Error GoTo Invalid
    Set expected = CreateObject("Scripting.Dictionary"): expected.CompareMode = vbBinaryCompare
    Set warehouses = CreateObject("Scripting.Dictionary"): warehouses.CompareMode = vbBinaryCompare
    For Each record In sources
        For Each key In record("Rows").Keys
            If expected.Exists(key) Then Exit Function
            Set entry = CreateObject("Scripting.Dictionary")
            For Each name In record("Rows")(key).Keys
                entry.Add CStr(name), record("Rows")(key)(name)
            Next name
            entry.Add "SourceSnapshot", record("Path")
            expected.Add key, entry
            warehouses(CStr(record("Rows")(key)("WarehouseId"))) = True
        Next key
    Next record
    Set actual = ReadRows(path, "tblGlobalInventorySnapshot")
    If actual Is Nothing Then Exit Function
    If actual.Count <> expected.Count Then Exit Function
    For Each key In expected.Keys
        If Not actual.Exists(key) Then Exit Function
        For Each name In expected(key).Keys
            If name = "QtyOnHand" Then
                If CDbl(actual(key)(name)) <> CDbl(expected(key)(name)) Then Exit Function
            Else
                If StrComp(CStr(actual(key)(name)), CStr(expected(key)(name)), vbBinaryCompare) <> 0 Then Exit Function
            End If
        Next name
    Next key
    Complete = StatusValid(path, sources.Count, warehouses.Count)
Invalid:
End Function

Private Function StatusValid(ByVal path As String, ByVal sourceCount As Long, ByVal warehouseCount As Long) As Boolean
    Dim book As Workbook, copy As String, table As ListObject, columns As Object, name As Variant, expected As Object
    On Error GoTo Finished
    Set book = modAggregationFiles.OpenCopy(path, copy)
    If book Is Nothing Then GoTo Finished
    Set table = modExecutionTarget.Table(book, "tblGlobalSnapshotStatus")
    Set columns = RequiredColumns(table, Array("Scope", "AuthorityLevel", "AuthoritativeStore", "VisibilityRule", "GeneratedAtUTC", _
        "SnapshotsFolder", "SnapshotFileCount", "SkippedSnapshotFileCount", "WarehouseCount"))
    If columns Is Nothing Then GoTo Finished
    If table.ListRows.Count <> 1 Then GoTo Finished
    Set expected = CreateObject("Scripting.Dictionary")
    expected.Add "Scope", "GLOBAL": expected.Add "AuthorityLevel", "ADVISORY_ONLY"
    expected.Add "AuthoritativeStore", "Warehouse-local WHx.invSys.Data.Inventory.xlsb"
    expected.Add "VisibilityRule", "Never overrides warehouse-local authoritative balances"
    expected.Add "SnapshotsFolder", "Explicit selected sources\"
    For Each name In expected.Keys
        If CStr(table.DataBodyRange.Cells(1, columns(CStr(name))).Value2) <> expected(name) Then GoTo Finished
    Next name
    If Not modTrainingWire.ValidUtcTimestamp(CStr(table.DataBodyRange.Cells(1, columns("GeneratedAtUTC")).Value2)) Then GoTo Finished
    For Each name In Array("SnapshotFileCount", "SkippedSnapshotFileCount", "WarehouseCount")
        If Not IsNumeric(table.DataBodyRange.Cells(1, columns(CStr(name))).Value2) Then GoTo Finished
    Next name
    If CDbl(table.DataBodyRange.Cells(1, columns("SnapshotFileCount")).Value2) <> sourceCount Then GoTo Finished
    If CDbl(table.DataBodyRange.Cells(1, columns("SkippedSnapshotFileCount")).Value2) <> 0 Then GoTo Finished
    If CDbl(table.DataBodyRange.Cells(1, columns("WarehouseCount")).Value2) <> warehouseCount Then GoTo Finished
    StatusValid = True
Finished:
    modAggregationFiles.CloseCopy book, copy
End Function
