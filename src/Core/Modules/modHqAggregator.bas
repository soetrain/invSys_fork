Attribute VB_Name = "modHqAggregator"
Option Explicit

Private Const SHEET_GLOBAL_SNAPSHOT As String = "GlobalInventorySnapshot"
Private Const TABLE_GLOBAL_SNAPSHOT As String = "tblGlobalInventorySnapshot"
Private Const SHEET_GLOBAL_STATUS As String = "GlobalSnapshotStatus"
Private Const TABLE_GLOBAL_STATUS As String = "tblGlobalSnapshotStatus"
Private Const TABLE_WAREHOUSE_SNAPSHOT As String = "tblInventorySnapshot"

Public Function RunHQAggregation(Optional ByVal sharePointRoot As String = "", _
                                 Optional ByVal outputPath As String = "", _
                                 Optional ByRef report As String = "") As Boolean
    Dim snapshotsFolder As String

    If sharePointRoot = "" Then
        If Not modConfig.LoadConfig("", "") Then
            report = "Config load failed: " & modConfig.Validate()
            Exit Function
        End If
        sharePointRoot = modConfig.GetString("PathSharePointRoot", "")
    End If
    If Trim$(sharePointRoot) = "" Then
        report = "PathSharePointRoot not configured."
        Exit Function
    End If

    snapshotsFolder = NormalizeFolderPathHq(sharePointRoot) & "Snapshots"
    If Trim$(outputPath) = "" Then outputPath = NormalizeFolderPathHq(sharePointRoot) & "Global\invSys.Global.InventorySnapshot.xlsb"
    RunHQAggregation = GenerateGlobalSnapshotFromFolder(snapshotsFolder, outputPath, report)
End Function

Public Function GenerateGlobalSnapshotFromFolder(ByVal snapshotsFolder As String, _
                                                 ByVal outputPath As String, _
                                                 Optional ByRef report As String = "") As Boolean
    On Error GoTo FailAggregate

    Dim fileName As String
    Dim tempFolder As String
    Dim tempFile As String
    Dim wbSnap As Workbook
    Dim globalRows As Object
    Dim key As String
    Dim lo As ListObject
    Dim i As Long
    Dim snapshotFileCount As Long

    If Trim$(snapshotsFolder) = "" Then
        report = "Snapshots folder is required."
        Exit Function
    End If

    Set globalRows = CreateObject("Scripting.Dictionary")
    globalRows.CompareMode = vbTextCompare
    tempFolder = Environ$("TEMP") & "\invSysHQ_" & Format$(Now, "yyyymmdd_hhnnss")
    CreateFolderRecursiveHq tempFolder

    fileName = Dir$(NormalizeFolderPathHq(snapshotsFolder) & "*.invSys.Snapshot.Inventory.xls*")
    Do While fileName <> ""
        tempFile = NormalizeFolderPathHq(tempFolder) & fileName
        On Error Resume Next
        Kill tempFile
        On Error GoTo FailAggregate
        FileCopy NormalizeFolderPathHq(snapshotsFolder) & fileName, tempFile

        Set wbSnap = Application.Workbooks.Open(tempFile, ReadOnly:=True)
        snapshotFileCount = snapshotFileCount + 1
        Set lo = FindListObjectByNameHq(wbSnap, TABLE_WAREHOUSE_SNAPSHOT)
        If Not lo Is Nothing Then
            For i = 1 To lo.ListRows.Count
                If SafeTrimHq(GetCellByColumnHq(lo, i, "SKU")) <> "" Then
                    key = SafeTrimHq(GetCellByColumnHq(lo, i, "WarehouseId")) & "|" & SafeTrimHq(GetCellByColumnHq(lo, i, "SKU"))
                    MergeSnapshotRow globalRows, key, lo, i, fileName
                End If
            Next i
        End If
        wbSnap.Close SaveChanges:=False
        Set wbSnap = Nothing

        fileName = Dir$
    Loop

    WriteGlobalSnapshotWorkbook outputPath, globalRows, snapshotsFolder, snapshotFileCount
    report = "Rows=" & CStr(globalRows.Count) & "; SnapshotFiles=" & CStr(snapshotFileCount)
    GenerateGlobalSnapshotFromFolder = True
    Exit Function

FailAggregate:
    On Error Resume Next
    If Not wbSnap Is Nothing Then wbSnap.Close SaveChanges:=False
    On Error GoTo 0
    report = "GenerateGlobalSnapshotFromFolder failed: " & Err.Description
End Function

Private Sub MergeSnapshotRow(ByVal globalRows As Object, _
                             ByVal key As String, _
                             ByVal lo As ListObject, _
                             ByVal rowIndex As Long, _
                             ByVal sourceFile As String)
    Dim entry As Object
    Dim currentDate As Variant
    Dim existingDate As Variant

    If globalRows.Exists(key) Then
        Set entry = globalRows(key)
        currentDate = GetCellByColumnHq(lo, rowIndex, "LastAppliedAtUTC")
        existingDate = entry("LastAppliedAtUTC")
        If IsDate(currentDate) And IsDate(existingDate) Then
            If CDate(currentDate) <= CDate(existingDate) Then Exit Sub
        End If
    Else
        Set entry = CreateObject("Scripting.Dictionary")
        entry.CompareMode = vbTextCompare
        globalRows.Add key, entry
    End If

    entry("WarehouseId") = GetCellByColumnHq(lo, rowIndex, "WarehouseId")
    entry("SKU") = GetCellByColumnHq(lo, rowIndex, "SKU")
    entry("QtyOnHand") = GetCellByColumnHq(lo, rowIndex, "QtyOnHand")
    entry("LastAppliedAtUTC") = GetCellByColumnHq(lo, rowIndex, "LastAppliedAtUTC")
    entry("SourceSnapshot") = sourceFile
End Sub

Private Sub WriteGlobalSnapshotWorkbook(ByVal outputPath As String, _
                                        ByVal globalRows As Object, _
                                        ByVal snapshotsFolder As String, _
                                        ByVal snapshotFileCount As Long)
    Dim wb As Workbook
    Dim wsSnap As Worksheet
    Dim wsStatus As Worksheet
    Dim loSnap As ListObject
    Dim loStatus As ListObject
    Dim snapHeaders As Variant
    Dim statusHeaders As Variant
    Dim startCell As Range
    Dim i As Long
    Dim key As Variant
    Dim rowIndex As Long
    Dim generatedAt As Date

    EnsureFolderForFileHq outputPath
    CloseWorkbookByFullNameHq outputPath
    On Error Resume Next
    Kill outputPath
    On Error GoTo 0

    Set wb = Application.Workbooks.Add
    generatedAt = Now
    snapHeaders = Array("WarehouseId", "SKU", "QtyOnHand", "LastAppliedAtUTC", "SourceSnapshot")
    statusHeaders = Array("Scope", "AuthorityLevel", "AuthoritativeStore", "VisibilityRule", "GeneratedAtUTC", _
                          "SnapshotsFolder", "SnapshotFileCount", "WarehouseCount")

    Set wsSnap = wb.Worksheets(1)
    wsSnap.Name = SHEET_GLOBAL_SNAPSHOT
    Set startCell = wsSnap.Range("A1")
    For i = LBound(snapHeaders) To UBound(snapHeaders)
        startCell.Offset(0, i - LBound(snapHeaders)).Value = snapHeaders(i)
    Next i

    Set loSnap = wsSnap.ListObjects.Add(xlSrcRange, wsSnap.Range(startCell, startCell.Offset(1, UBound(snapHeaders) - LBound(snapHeaders))), , xlYes)
    loSnap.Name = TABLE_GLOBAL_SNAPSHOT
    If loSnap.DataBodyRange Is Nothing Then loSnap.ListRows.Add
    DeleteAllRowsHq loSnap

    For Each key In globalRows.Keys
        loSnap.ListRows.Add
        rowIndex = loSnap.ListRows.Count
        SetTableRowValueHq loSnap, rowIndex, "WarehouseId", globalRows(key)("WarehouseId")
        SetTableRowValueHq loSnap, rowIndex, "SKU", globalRows(key)("SKU")
        SetTableRowValueHq loSnap, rowIndex, "QtyOnHand", globalRows(key)("QtyOnHand")
        SetTableRowValueHq loSnap, rowIndex, "LastAppliedAtUTC", globalRows(key)("LastAppliedAtUTC")
        SetTableRowValueHq loSnap, rowIndex, "SourceSnapshot", globalRows(key)("SourceSnapshot")
    Next key

    Set wsStatus = wb.Worksheets.Add(After:=wsSnap)
    wsStatus.Name = SHEET_GLOBAL_STATUS
    Set startCell = wsStatus.Range("A1")
    For i = LBound(statusHeaders) To UBound(statusHeaders)
        startCell.Offset(0, i - LBound(statusHeaders)).Value = statusHeaders(i)
    Next i

    Set loStatus = wsStatus.ListObjects.Add(xlSrcRange, wsStatus.Range(startCell, startCell.Offset(1, UBound(statusHeaders) - LBound(statusHeaders))), , xlYes)
    loStatus.Name = TABLE_GLOBAL_STATUS
    If loStatus.DataBodyRange Is Nothing Then loStatus.ListRows.Add
    DeleteAllRowsHq loStatus
    loStatus.ListRows.Add
    SetTableRowValueHq loStatus, 1, "Scope", "GLOBAL"
    SetTableRowValueHq loStatus, 1, "AuthorityLevel", "ADVISORY_ONLY"
    SetTableRowValueHq loStatus, 1, "AuthoritativeStore", "Warehouse-local WHx.invSys.Data.Inventory.xlsb"
    SetTableRowValueHq loStatus, 1, "VisibilityRule", "Never overrides warehouse-local authoritative balances"
    SetTableRowValueHq loStatus, 1, "GeneratedAtUTC", generatedAt
    SetTableRowValueHq loStatus, 1, "SnapshotsFolder", NormalizeFolderPathHq(snapshotsFolder)
    SetTableRowValueHq loStatus, 1, "SnapshotFileCount", snapshotFileCount
    SetTableRowValueHq loStatus, 1, "WarehouseCount", CountWarehouseIdsHq(globalRows)

    wsSnap.Cells.EntireColumn.AutoFit
    wsStatus.Cells.EntireColumn.AutoFit

    wb.SaveAs Filename:=outputPath, FileFormat:=50
    wb.Close SaveChanges:=True
End Sub

Private Function CountWarehouseIdsHq(ByVal globalRows As Object) As Long
    Dim seen As Object
    Dim key As Variant
    Dim warehouseId As String

    If globalRows Is Nothing Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For Each key In globalRows.Keys
        warehouseId = SafeTrimHq(globalRows(key)("WarehouseId"))
        If warehouseId <> "" Then
            If Not seen.Exists(warehouseId) Then seen.Add warehouseId, True
        End If
    Next key

    CountWarehouseIdsHq = seen.Count
End Function

Private Function FindListObjectByNameHq(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameHq = ws.ListObjects(tableName)
        If Not FindListObjectByNameHq Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function GetCellByColumnHq(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexHq(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnHq = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Sub SetTableRowValueHq(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexHq(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetColumnIndexHq(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexHq = i
            Exit Function
        End If
    Next i
End Function

Private Sub DeleteAllRowsHq(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
End Sub

Private Function NormalizeFolderPathHq(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathHq = folderPath
End Function

Private Function SafeTrimHq(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimHq = Trim$(CStr(valueIn))
End Function

Private Sub EnsureFolderForFileHq(ByVal filePath As String)
    Dim folderPath As String
    Dim sepPos As Long

    sepPos = InStrRev(filePath, "\")
    If sepPos <= 0 Then Exit Sub
    folderPath = Left$(filePath, sepPos - 1)
    CreateFolderRecursiveHq folderPath
End Sub

Private Sub CreateFolderRecursiveHq(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 0 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then CreateFolderRecursiveHq parentPath
    End If
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Sub CloseWorkbookByFullNameHq(ByVal fullNameIn As String)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullNameIn, vbTextCompare) = 0 Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
            Exit For
        End If
    Next wb
End Sub
