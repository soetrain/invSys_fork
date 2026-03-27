Attribute VB_Name = "modOperatorReadModel"
Option Explicit

Private Const SHEET_INVENTORY_MANAGEMENT As String = "InventoryManagement"
Private Const TABLE_INVSYS As String = "invSys"
Private Const SHEET_SNAPSHOT As String = "InventorySnapshot"
Private Const TABLE_SNAPSHOT As String = "tblInventorySnapshot"

Public Function RefreshInventoryReadModelForWorkbook(Optional ByVal targetWb As Workbook = Nothing, _
                                                     Optional ByVal warehouseId As String = "", _
                                                     Optional ByVal sourceType As String = "LOCAL", _
                                                     Optional ByRef report As String = "") As Boolean
    On Error GoTo FailRefresh

    Dim wb As Workbook
    Dim loInv As ListObject
    Dim wbSnap As Workbook
    Dim loSnap As ListObject
    Dim snapshotRows As Object
    Dim snapshotId As String
    Dim refreshUtc As Date
    Dim normalizedSource As String
    Dim resolvedWarehouseId As String
    Dim configValidation As String
    Dim snapshotPath As String
    Dim snapshotAlreadyOpen As Boolean

    Set wb = ResolveOperatorWorkbook(targetWb)
    If wb Is Nothing Then
        report = "Operator workbook not resolved."
        Exit Function
    End If

    Set loInv = FindListObjectReadModel(wb, TABLE_INVSYS)
    If loInv Is Nothing Then
        report = "invSys table not found."
        Exit Function
    End If

    refreshUtc = Now
    normalizedSource = NormalizeSourceType(sourceType)
    resolvedWarehouseId = ResolveWarehouseIdReadModel(warehouseId)
    If Not modConfig.IsLoaded() Then
        Call modConfig.LoadConfig(resolvedWarehouseId, "")
        configValidation = modConfig.Validate()
    End If

    snapshotPath = ResolveSnapshotPathReadModel(resolvedWarehouseId)
    snapshotAlreadyOpen = WorkbookIsOpenByPathReadModel(snapshotPath)
    Set wbSnap = ResolveSnapshotWorkbook(resolvedWarehouseId, "", Nothing, False)
    If wbSnap Is Nothing Then
        MarkReadModelState loInv, refreshUtc, vbNullString, "CACHED", True
        report = "Snapshot workbook not found; operator read model marked stale."
        If configValidation <> "" Then report = report & " " & configValidation
        RefreshInventoryReadModelForWorkbook = True
        GoTo CleanExit
    End If

    Set loSnap = FindListObjectReadModel(wbSnap, TABLE_SNAPSHOT)
    If loSnap Is Nothing Then
        MarkReadModelState loInv, refreshUtc, vbNullString, "CACHED", True
        report = "Snapshot table not found; operator read model marked stale."
        If configValidation <> "" Then report = report & " " & configValidation
        RefreshInventoryReadModelForWorkbook = True
        GoTo CleanExit
    End If

    Set snapshotRows = BuildSnapshotDictionary(loSnap)
    snapshotId = BuildSnapshotId(wbSnap)
    ApplySnapshotToInvSys loInv, snapshotRows, refreshUtc, snapshotId, normalizedSource
    report = "OK"
    RefreshInventoryReadModelForWorkbook = True
    
CleanExit:
    If Not snapshotAlreadyOpen Then CloseWorkbookQuietlyReadModel wbSnap
    Exit Function

FailRefresh:
    report = "RefreshInventoryReadModelForWorkbook failed: " & Err.Description
    If Not snapshotAlreadyOpen Then CloseWorkbookQuietlyReadModel wbSnap
End Function

Public Sub RefreshCurrentWorkbookInventoryReadModel()
    On Error GoTo FailRefreshCurrent

    Dim report As String
    Dim wb As Workbook

    Set wb = ResolveOperatorWorkbook(Nothing)

    If wb Is Nothing Then
        MsgBox "No active operator workbook was available for refresh.", vbExclamation
        Exit Sub
    End If

    If wb.IsAddin Then
        MsgBox "Activate the operator workbook before refreshing invSys.", vbExclamation
        Exit Sub
    End If

    If Not RefreshInventoryReadModelForWorkbook(wb, "", "LOCAL", report) Then
        MsgBox report, vbExclamation
    ElseIf report <> "OK" Then
        MsgBox report, vbInformation
    End If
    Exit Sub

FailRefreshCurrent:
    MsgBox "RefreshCurrentWorkbookInventoryReadModel failed: " & Err.Description, vbExclamation
End Sub

Public Function DiagnoseInventoryReadModelRefresh(Optional ByVal targetWb As Workbook = Nothing, _
                                                  Optional ByVal warehouseId As String = "", _
                                                  Optional ByVal sourceType As String = "LOCAL") As String
    On Error GoTo FailDiagnose

    Dim wb As Workbook
    Dim loInv As ListObject
    Dim wbSnap As Workbook
    Dim loSnap As ListObject
    Dim snapshotRows As Object
    Dim refreshReport As String
    Dim resolvedWarehouseId As String
    Dim snapshotPath As String
    Dim normalizedSource As String
    Dim configLoadedBefore As Boolean
    Dim configLoadResult As Boolean
    Dim beforeRows As Long
    Dim afterRows As Long
    Dim snapshotTableRows As Long
    Dim snapshotDictRows As Long
    Dim refreshResult As Boolean
    Dim snapshotAlreadyOpen As Boolean
    Dim snapshotOpenProbe As String

    Set wb = ResolveOperatorWorkbook(targetWb)
    If wb Is Nothing Then
        DiagnoseInventoryReadModelRefresh = "TargetWorkbook=<none>" & vbCrLf & _
                                            "Result=FAIL" & vbCrLf & _
                                            "Report=Operator workbook not resolved."
        Exit Function
    End If

    Set loInv = FindListObjectReadModel(wb, TABLE_INVSYS)
    beforeRows = GetListRowCountReadModel(loInv)
    normalizedSource = NormalizeSourceType(sourceType)
    resolvedWarehouseId = ResolveWarehouseIdReadModel(warehouseId)

    configLoadedBefore = modConfig.IsLoaded()
    If configLoadedBefore Then
        configLoadResult = True
    Else
        configLoadResult = modConfig.LoadConfig(resolvedWarehouseId, "")
    End If

    snapshotPath = ResolveSnapshotPathReadModel(resolvedWarehouseId)
    snapshotAlreadyOpen = WorkbookIsOpenByPathReadModel(snapshotPath)
    Set wbSnap = ResolveSnapshotWorkbook(resolvedWarehouseId, "", Nothing, False)
    If Not wbSnap Is Nothing Then
        Set loSnap = FindListObjectReadModel(wbSnap, TABLE_SNAPSHOT)
        snapshotTableRows = GetListRowCountReadModel(loSnap)
        Set snapshotRows = BuildSnapshotDictionary(loSnap)
        If Not snapshotRows Is Nothing Then snapshotDictRows = snapshotRows.Count
    Else
        snapshotOpenProbe = ProbeSnapshotOpenReadModel(snapshotPath)
    End If

    refreshResult = RefreshInventoryReadModelForWorkbook(wb, resolvedWarehouseId, normalizedSource, refreshReport)
    afterRows = GetListRowCountReadModel(loInv)

    DiagnoseInventoryReadModelRefresh = Join(Array( _
        "TargetWorkbook=" & wb.FullName, _
        "WarehouseId=" & resolvedWarehouseId, _
        "SourceType=" & normalizedSource, _
        "ConfigLoadedBefore=" & CStr(configLoadedBefore), _
        "ConfigLoadResult=" & CStr(configLoadResult), _
        "ConfigWorkbook=" & modConfig.GetResolvedWorkbookName(), _
        "PathDataRoot=" & modConfig.GetString("PathDataRoot", "<missing>"), _
        "PathInboxRoot=" & modConfig.GetString("PathInboxRoot", "<missing>"), _
        "SnapshotPath=" & snapshotPath, _
        "SnapshotFileExists=" & CStr(FileExistsReadModel(snapshotPath)), _
        "SnapshotWorkbookResolved=" & CStr(Not wbSnap Is Nothing), _
        "SnapshotWorkbook=" & ResolveWorkbookNameReadModel(wbSnap), _
        "SnapshotTableResolved=" & CStr(Not loSnap Is Nothing), _
        "SnapshotTableRows=" & CStr(snapshotTableRows), _
        "SnapshotDictionaryRows=" & CStr(snapshotDictRows), _
        "SnapshotOpenProbe=" & snapshotOpenProbe, _
        "InvSysRowsBefore=" & CStr(beforeRows), _
        "RefreshResult=" & CStr(refreshResult), _
        "RefreshReport=" & refreshReport, _
        "InvSysRowsAfter=" & CStr(afterRows), _
        "ConfigValidation=" & modConfig.Validate()), vbCrLf)
    If Not snapshotAlreadyOpen Then CloseWorkbookQuietlyReadModel wbSnap
    Exit Function

FailDiagnose:
    If Not snapshotAlreadyOpen Then CloseWorkbookQuietlyReadModel wbSnap
    DiagnoseInventoryReadModelRefresh = "Result=FAIL" & vbCrLf & _
                                        "Error=" & Err.Description
End Function

Private Function ResolveOperatorWorkbook(ByVal targetWb As Workbook) As Workbook
    If Not targetWb Is Nothing Then
        Set ResolveOperatorWorkbook = targetWb
        Exit Function
    End If

    If Not Application.ActiveWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook.IsAddin Then
            Set ResolveOperatorWorkbook = Application.ActiveWorkbook
        End If
    End If
End Function

Private Function ResolveWarehouseIdReadModel(ByVal warehouseId As String) As String
    ResolveWarehouseIdReadModel = Trim$(warehouseId)
    If ResolveWarehouseIdReadModel = "" Then ResolveWarehouseIdReadModel = Trim$(modConfig.GetWarehouseId())
    If ResolveWarehouseIdReadModel = "" Then ResolveWarehouseIdReadModel = "WH1"
End Function

Private Function ResolveSnapshotPathReadModel(ByVal warehouseId As String) As String
    Dim rootPath As String

    rootPath = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = Trim$(modConfig.GetString("PathDataRoot", Environ$("TEMP")))
    ResolveSnapshotPathReadModel = NormalizeFolderPathReadModel(rootPath) & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
End Function

Private Function WorkbookIsOpenByPathReadModel(ByVal targetPath As String) As Boolean
    Dim wb As Workbook

    targetPath = Trim$(targetPath)
    If targetPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            WorkbookIsOpenByPathReadModel = True
            Exit Function
        End If
    Next wb
End Function

Private Function NormalizeSourceType(ByVal sourceType As String) As String
    NormalizeSourceType = UCase$(Trim$(sourceType))
    If NormalizeSourceType = "" Then NormalizeSourceType = "LOCAL"
    Select Case NormalizeSourceType
        Case "LOCAL", "SHAREPOINT", "CACHED"
        Case Else
            NormalizeSourceType = "LOCAL"
    End Select
End Function

Private Function FindListObjectReadModel(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindListObjectReadModel = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindListObjectReadModel Is Nothing Then Exit Function
    Next ws
End Function

Private Function BuildSnapshotDictionary(ByVal loSnap As ListObject) As Object
    Dim dict As Object
    Dim skuIdx As Long
    Dim qtyOnHandIdx As Long
    Dim qtyAvailableIdx As Long
    Dim locationSummaryIdx As Long
    Dim appliedIdx As Long
    Dim i As Long
    Dim sku As String
    Dim payload As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    If loSnap Is Nothing Then
        Set BuildSnapshotDictionary = dict
        Exit Function
    End If
    If loSnap.DataBodyRange Is Nothing Then
        Set BuildSnapshotDictionary = dict
        Exit Function
    End If

    skuIdx = GetColumnIndexReadModel(loSnap, "SKU")
    qtyOnHandIdx = GetColumnIndexReadModel(loSnap, "QtyOnHand")
    qtyAvailableIdx = GetColumnIndexReadModel(loSnap, "QtyAvailable")
    locationSummaryIdx = GetColumnIndexReadModel(loSnap, "LocationSummary")
    appliedIdx = GetColumnIndexReadModel(loSnap, "LastAppliedAtUTC")
    If skuIdx = 0 Or qtyOnHandIdx = 0 Then
        Set BuildSnapshotDictionary = dict
        Exit Function
    End If

    For i = 1 To loSnap.ListRows.Count
        sku = Trim$(CStr(loSnap.DataBodyRange.Cells(i, skuIdx).Value))
        If sku = "" Then GoTo ContinueLoop
        payload = Array( _
            NzDblReadModel(loSnap.DataBodyRange.Cells(i, qtyOnHandIdx).Value), _
            ResolveSnapshotQtyAvailable(loSnap, i, qtyAvailableIdx, qtyOnHandIdx), _
            ResolveSnapshotLocationSummary(loSnap, i, locationSummaryIdx), _
            ResolveSnapshotLastApplied(loSnap, i, appliedIdx))
        dict(sku) = payload
ContinueLoop:
    Next i

    Set BuildSnapshotDictionary = dict
End Function

Private Sub ApplySnapshotToInvSys(ByVal loInv As ListObject, _
                                  ByVal snapshotRows As Object, _
                                  ByVal refreshUtc As Date, _
                                  ByVal snapshotId As String, _
                                  ByVal sourceType As String)
    Dim rowIndex As Long
    Dim sku As String
    Dim payload As Variant
    Dim qtyOnHand As Double
    Dim qtyAvailable As Double
    Dim locationSummary As String
    Dim lastApplied As Variant

    If loInv Is Nothing Then Exit Sub

    EnsureInvSysRowsForSnapshot loInv, snapshotRows
    If loInv.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To loInv.ListRows.Count
        sku = ResolveInvSysSku(loInv, rowIndex)
        SyncDisplayAliases loInv, rowIndex

        If sku <> "" And Not snapshotRows Is Nothing And snapshotRows.Exists(sku) Then
            payload = snapshotRows(sku)
            qtyOnHand = NzDblReadModel(payload(0))
            qtyAvailable = NzDblReadModel(payload(1))
            locationSummary = Trim$(CStr(payload(2)))
            lastApplied = payload(3)
            ApplyReadModelValues loInv, rowIndex, qtyOnHand, qtyAvailable, locationSummary, lastApplied, refreshUtc, snapshotId, sourceType, False
        ElseIf sku <> "" Then
            ApplyReadModelValues loInv, rowIndex, 0, 0, vbNullString, Empty, refreshUtc, snapshotId, sourceType, False
        Else
            ApplyReadModelValues loInv, rowIndex, NzDblReadModel(GetReadModelValue(loInv, rowIndex, "TOTAL INV")), _
                                NzDblReadModel(GetReadModelValue(loInv, rowIndex, "QtyAvailable")), _
                                CStr(GetReadModelValue(loInv, rowIndex, "LocationSummary")), _
                                GetReadModelValue(loInv, rowIndex, "LAST EDITED"), refreshUtc, snapshotId, sourceType, False
        End If
    Next rowIndex
End Sub

Private Sub EnsureInvSysRowsForSnapshot(ByVal loInv As ListObject, ByVal snapshotRows As Object)
    Dim key As Variant
    Dim rowIndex As Long

    If loInv Is Nothing Then Exit Sub
    If snapshotRows Is Nothing Then Exit Sub

    For Each key In snapshotRows.Keys
        If Trim$(CStr(key)) <> "" Then
            rowIndex = FindInvSysRowBySku(loInv, CStr(key))
            If rowIndex = 0 Then
                rowIndex = AppendInvSysRow(loInv)
                If rowIndex > 0 Then SeedInvSysRow loInv, rowIndex, CStr(key)
            End If
        End If
    Next key
End Sub

Private Function FindInvSysRowBySku(ByVal loInv As ListObject, ByVal sku As String) As Long
    Dim rowIndex As Long

    If loInv Is Nothing Then Exit Function
    If loInv.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To loInv.ListRows.Count
        If StrComp(ResolveInvSysSku(loInv, rowIndex), sku, vbTextCompare) = 0 Then
            FindInvSysRowBySku = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function AppendInvSysRow(ByVal loInv As ListObject) As Long
    If loInv Is Nothing Then Exit Function
    loInv.ListRows.Add
    AppendInvSysRow = loInv.ListRows.Count
End Function

Private Sub SeedInvSysRow(ByVal loInv As ListObject, ByVal rowIndex As Long, ByVal sku As String)
    If loInv Is Nothing Then Exit Sub
    If rowIndex <= 0 Then Exit Sub

    SetReadModelValue loInv, rowIndex, "ITEM_CODE", sku
    SetReadModelValue loInv, rowIndex, "ITEM", sku
End Sub

Private Sub ApplyReadModelValues(ByVal loInv As ListObject, _
                                 ByVal rowIndex As Long, _
                                 ByVal qtyOnHand As Double, _
                                 ByVal qtyAvailable As Double, _
                                 ByVal locationSummary As String, _
                                 ByVal lastApplied As Variant, _
                                 ByVal refreshUtc As Date, _
                                 ByVal snapshotId As String, _
                                 ByVal sourceType As String, _
                                 ByVal isStale As Boolean)
    locationSummary = NormalizeLocationSummaryReadModel(locationSummary)
    SetReadModelValue loInv, rowIndex, "TOTAL INV", qtyOnHand
    SetReadModelValue loInv, rowIndex, "QtyAvailable", qtyAvailable
    SetReadModelValue loInv, rowIndex, "LocationSummary", locationSummary
    If locationSummary <> "" Then
        SetReadModelValue loInv, rowIndex, "LOCATION", ResolvePrimaryLocationReadModel(locationSummary, GetReadModelValue(loInv, rowIndex, "LOCATION"))
    End If
    If Not IsEmpty(lastApplied) And Not IsNull(lastApplied) And CStr(lastApplied) <> "" Then
        SetReadModelValue loInv, rowIndex, "LAST EDITED", lastApplied
        SetReadModelValue loInv, rowIndex, "TOTAL INV LAST EDIT", lastApplied
    Else
        SetReadModelValue loInv, rowIndex, "LAST EDITED", vbNullString
        SetReadModelValue loInv, rowIndex, "TOTAL INV LAST EDIT", vbNullString
    End If
    SetReadModelValue loInv, rowIndex, "LastRefreshUTC", refreshUtc
    SetReadModelValue loInv, rowIndex, "SnapshotId", snapshotId
    SetReadModelValue loInv, rowIndex, "SourceType", sourceType
    SetReadModelValue loInv, rowIndex, "IsStale", isStale
End Sub

Private Sub MarkReadModelState(ByVal loInv As ListObject, _
                               ByVal refreshUtc As Date, _
                               ByVal snapshotId As String, _
                               ByVal sourceType As String, _
                               ByVal isStale As Boolean)
    Dim rowIndex As Long

    If loInv Is Nothing Then Exit Sub
    If loInv.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To loInv.ListRows.Count
        SyncDisplayAliases loInv, rowIndex
        SetReadModelValue loInv, rowIndex, "LastRefreshUTC", refreshUtc
        SetReadModelValue loInv, rowIndex, "SnapshotId", snapshotId
        SetReadModelValue loInv, rowIndex, "SourceType", sourceType
        SetReadModelValue loInv, rowIndex, "IsStale", isStale
    Next rowIndex
End Sub

Private Sub SyncDisplayAliases(ByVal loInv As ListObject, ByVal rowIndex As Long)
    Dim sku As String
    Dim itemName As String

    sku = ResolveInvSysSku(loInv, rowIndex)
    itemName = Trim$(CStr(GetReadModelValue(loInv, rowIndex, "ITEM")))
    If itemName = "" Then itemName = Trim$(CStr(GetReadModelValue(loInv, rowIndex, "ItemName")))

    If sku <> "" Then SetReadModelValue loInv, rowIndex, "ITEM_CODE", sku
    If itemName <> "" Then
        SetReadModelValue loInv, rowIndex, "ITEM", itemName
    End If
End Sub

Private Function ResolveInvSysSku(ByVal loInv As ListObject, ByVal rowIndex As Long) As String
    ResolveInvSysSku = Trim$(CStr(GetReadModelValue(loInv, rowIndex, "ITEM_CODE")))
    If ResolveInvSysSku = "" Then ResolveInvSysSku = Trim$(CStr(GetReadModelValue(loInv, rowIndex, "SKU")))
End Function

Private Function BuildSnapshotId(ByVal wbSnap As Workbook) As String
    Dim modifiedUtc As String

    If wbSnap Is Nothing Then Exit Function
    On Error Resume Next
    modifiedUtc = Format$(FileDateTime(wbSnap.FullName), "yyyymmddhhnnss")
    On Error GoTo 0
    If modifiedUtc = "" Then modifiedUtc = Format$(Now, "yyyymmddhhnnss")
    BuildSnapshotId = wbSnap.Name & "|" & modifiedUtc
End Function

Private Function GetColumnIndexReadModel(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexReadModel = i
            Exit Function
        End If
    Next i
End Function

Private Function GetReadModelValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim colIndex As Long

    colIndex = GetColumnIndexReadModel(lo, columnName)
    If colIndex = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetReadModelValue = lo.DataBodyRange.Cells(rowIndex, colIndex).Value
End Function

Private Sub SetReadModelValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim colIndex As Long

    colIndex = GetColumnIndexReadModel(lo, columnName)
    If colIndex = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = valueOut
End Sub

Private Function NzDblReadModel(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzDblReadModel = CDbl(valueIn)
End Function

Private Function ResolveSnapshotQtyAvailable(ByVal loSnap As ListObject, _
                                             ByVal rowIndex As Long, _
                                             ByVal qtyAvailableIdx As Long, _
                                             ByVal qtyOnHandIdx As Long) As Double
    If qtyAvailableIdx > 0 Then
        ResolveSnapshotQtyAvailable = NzDblReadModel(loSnap.DataBodyRange.Cells(rowIndex, qtyAvailableIdx).Value)
    ElseIf qtyOnHandIdx > 0 Then
        ResolveSnapshotQtyAvailable = NzDblReadModel(loSnap.DataBodyRange.Cells(rowIndex, qtyOnHandIdx).Value)
    End If
End Function

Private Function ResolveSnapshotLocationSummary(ByVal loSnap As ListObject, _
                                                ByVal rowIndex As Long, _
                                                ByVal locationSummaryIdx As Long) As String
    If locationSummaryIdx = 0 Then Exit Function
    ResolveSnapshotLocationSummary = Trim$(CStr(loSnap.DataBodyRange.Cells(rowIndex, locationSummaryIdx).Value))
End Function

Private Function ResolveSnapshotLastApplied(ByVal loSnap As ListObject, _
                                            ByVal rowIndex As Long, _
                                            ByVal appliedIdx As Long) As Variant
    If appliedIdx = 0 Then Exit Function
    ResolveSnapshotLastApplied = loSnap.DataBodyRange.Cells(rowIndex, appliedIdx).Value
End Function

Private Function NormalizeFolderPathReadModel(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then
        NormalizeFolderPathReadModel = Environ$("TEMP") & "\"
        Exit Function
    End If
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathReadModel = folderPath
End Function

Private Function FileExistsReadModel(ByVal fullPath As String) As Boolean
    Dim fso As Object

    fullPath = Trim$(Replace$(fullPath, "/", "\"))
    If fullPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsReadModel = fso.FileExists(fullPath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsReadModel = (Len(Dir$(fullPath, vbNormal)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function ResolveWorkbookNameReadModel(ByVal wb As Workbook) As String
    If wb Is Nothing Then
        ResolveWorkbookNameReadModel = "<none>"
    Else
        ResolveWorkbookNameReadModel = wb.FullName
    End If
End Function

Private Function GetListRowCountReadModel(ByVal lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetListRowCountReadModel = lo.ListRows.Count
End Function

Private Function ProbeSnapshotOpenReadModel(ByVal snapshotPath As String) As String
    Dim wb As Workbook
    Dim loSnap As ListObject
    Dim prevAlerts As Boolean

    snapshotPath = Trim$(snapshotPath)
    If snapshotPath = "" Then
        ProbeSnapshotOpenReadModel = "NoPath"
        Exit Function
    End If

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    Set wb = Application.Workbooks.Open( _
        Filename:=snapshotPath, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, _
        Notify:=False, _
        AddToMru:=False)
    If Err.Number <> 0 Then
        ProbeSnapshotOpenReadModel = "OpenError " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        Application.DisplayAlerts = prevAlerts
        Exit Function
    End If
    On Error GoTo 0

    If wb Is Nothing Then
        ProbeSnapshotOpenReadModel = "OpenReturnedNothing"
        Application.DisplayAlerts = prevAlerts
        Exit Function
    End If

    Set loSnap = FindListObjectReadModel(wb, TABLE_SNAPSHOT)
    If loSnap Is Nothing Then
        ProbeSnapshotOpenReadModel = "OpenedNoTable " & wb.FullName
    Else
        ProbeSnapshotOpenReadModel = "OpenedRows=" & CStr(GetListRowCountReadModel(loSnap)) & " " & wb.FullName
    End If

    CloseWorkbookQuietlyReadModel wb
    Application.DisplayAlerts = prevAlerts
End Function

Private Sub CloseWorkbookQuietlyReadModel(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function ResolvePrimaryLocationReadModel(ByVal locationSummary As String, ByVal existingLocation As Variant) As String
    Dim summaryText As String
    Dim firstFragment As String
    Dim eqPos As Long
    Dim rawLocation As String

    summaryText = Trim$(locationSummary)
    If summaryText = "" Then
        ResolvePrimaryLocationReadModel = Trim$(CStr(existingLocation))
        Exit Function
    End If

    firstFragment = Split(summaryText, ";")(0)
    firstFragment = Trim$(firstFragment)
    eqPos = InStr(1, firstFragment, "=", vbTextCompare)
    If eqPos > 1 Then
        rawLocation = NormalizeDisplayLocationReadModel(Trim$(Left$(firstFragment, eqPos - 1)))
        If rawLocation <> "" Then
            ResolvePrimaryLocationReadModel = rawLocation
            Exit Function
        End If
    End If

    ResolvePrimaryLocationReadModel = NormalizeDisplayLocationReadModel(Trim$(CStr(existingLocation)))
End Function

Private Function NormalizeDisplayLocationReadModel(ByVal locationText As String) As String
    Dim eqPos As Long
    Dim suffixText As String

    locationText = Trim$(locationText)
    If locationText = "" Then Exit Function

    eqPos = InStrRev(locationText, "=")
    If eqPos > 1 Then
        suffixText = Trim$(Mid$(locationText, eqPos + 1))
        suffixText = Replace$(suffixText, ",", "")
        If suffixText <> "" Then
            If IsNumeric(suffixText) Then locationText = Trim$(Left$(locationText, eqPos - 1))
        End If
    End If

    NormalizeDisplayLocationReadModel = locationText
End Function

Private Function NormalizeLocationSummaryReadModel(ByVal locationSummary As String) As String
    Dim summaryText As String
    Dim fragments As Variant
    Dim fragment As Variant
    Dim fragmentText As String
    Dim eqPos As Long
    Dim label As String
    Dim qtyText As String
    Dim totals As Object

    summaryText = Trim$(locationSummary)
    If summaryText = "" Then Exit Function

    fragments = Split(summaryText, ";")
    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare

    For Each fragment In fragments
        fragmentText = Trim$(CStr(fragment))
        If fragmentText <> "" Then
            eqPos = InStrRev(fragmentText, "=")
            If eqPos <= 1 Then
                NormalizeLocationSummaryReadModel = summaryText
                Exit Function
            End If

            label = NormalizeDisplayLocationReadModel(Trim$(Left$(fragmentText, eqPos - 1)))
            If label = "" Then label = "(blank)"

            qtyText = Trim$(Mid$(fragmentText, eqPos + 1))
            qtyText = Replace$(qtyText, ",", "")
            If qtyText = "" Or Not IsNumeric(qtyText) Then
                NormalizeLocationSummaryReadModel = summaryText
                Exit Function
            End If

            If totals.Exists(label) Then
                totals(label) = CDbl(totals(label)) + CDbl(qtyText)
            Else
                totals.Add label, CDbl(qtyText)
            End If
        End If
    Next fragment

    NormalizeLocationSummaryReadModel = BuildNormalizedLocationSummaryReadModel(totals)
End Function

Private Function BuildNormalizedLocationSummaryReadModel(ByVal totals As Object) As String
    Dim key As Variant
    Dim fragment As String

    If totals Is Nothing Then Exit Function

    For Each key In totals.Keys
        fragment = CStr(key) & "=" & FormatQuantityReadModel(CDbl(totals(key)))
        If BuildNormalizedLocationSummaryReadModel = "" Then
            BuildNormalizedLocationSummaryReadModel = fragment
        Else
            BuildNormalizedLocationSummaryReadModel = BuildNormalizedLocationSummaryReadModel & "; " & fragment
        End If
    Next key
End Function

Private Function FormatQuantityReadModel(ByVal qtyIn As Double) As String
    If Abs(qtyIn - CLng(qtyIn)) < 0.0000001 Then
        FormatQuantityReadModel = CStr(CLng(qtyIn))
    Else
        FormatQuantityReadModel = Replace$(Format$(qtyIn, "0.########"), ",", "")
    End If
End Function
