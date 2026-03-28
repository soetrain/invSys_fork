Attribute VB_Name = "modInventoryPublisher"
Option Explicit

Private Const TABLE_INVENTORY_LOG_PUBLISHER As String = "tblInventoryLog"
Private Const TABLE_APPLIED_EVENTS_PUBLISHER As String = "tblAppliedEvents"
Private Const TABLE_SKU_BALANCE_PUBLISHER As String = "tblSkuBalance"
Private Const TABLE_LOCATION_BALANCE_PUBLISHER As String = "tblLocationBalance"
Private Const TABLE_LEDGER_STATUS_PUBLISHER As String = "tblInventoryLedgerStatus"
Private Const TABLE_SKU_CATALOG_PUBLISHER As String = "tblSkuCatalog"
Private Const TABLE_INVSYS_PUBLISHER As String = "invSys"
Private Const TABLE_RECEIVED_TALLY_PUBLISHER As String = "ReceivedTally"
Private Const TABLE_SHIPMENTS_TALLY_PUBLISHER As String = "ShipmentsTally"
Private Const TABLE_PRODUCTION_OUTPUT_PUBLISHER As String = "ProductionOutput"
Private Const TABLE_ADMIN_AUDIT_PUBLISHER As String = "tblAdminAudit"
Private Const TABLE_RECIPES_PUBLISHER As String = "Recipes"
Private Const TABLE_TEMPLATES_PUBLISHER As String = "TemplatesTable"
Private Const TABLE_INGREDIENT_PALETTE_PUBLISHER As String = "IngredientPalette"
Private Const MIN_RECENT_PUBLISH_SECONDS As Long = 5

Private mRecentPublishes As Object

Public Function PublishOpenInventorySnapshots(Optional ByRef report As String = "") As Long
    On Error GoTo FailPublish

    Dim wb As Workbook
    Dim publishReport As String
    Dim detail As String

    For Each wb In Application.Workbooks
        publishReport = vbNullString
        If IsInventorySourceWorkbookPublisher(wb) Then
            If EnsureSnapshotPublicationForWorkbook(wb, publishReport) Then
                If StrComp(Trim$(publishReport), "SKIPPED_RECENT", vbTextCompare) <> 0 Then
                    PublishOpenInventorySnapshots = PublishOpenInventorySnapshots + 1
                End If
            End If
            If publishReport <> "" Then
                If detail <> "" Then detail = detail & "; "
                detail = detail & wb.Name & "=" & publishReport
            End If
        End If
    Next wb

    report = detail
    Exit Function

FailPublish:
    report = "PublishOpenInventorySnapshots failed: " & Err.Description
End Function

Public Function EnsureSnapshotPublicationForWorkbook(Optional ByVal targetWb As Workbook = Nothing, _
                                                     Optional ByRef report As String = "") As Boolean
    On Error GoTo FailPublish

    Dim wb As Workbook
    Dim publishWb As Workbook
    Dim resolvedWarehouseId As String
    Dim snapshotPath As String
    Dim publishKey As String
    Dim runtimeWasOpen As Boolean
    Dim runtimePath As String

    Set wb = targetWb
    If wb Is Nothing Then Set wb = ResolveCandidateInventoryWorkbookPublisher()
    If wb Is Nothing Then
        report = "Inventory source workbook not resolved."
        Exit Function
    End If
    If Not IsInventorySourceWorkbookPublisher(wb) Then
        report = "Workbook is not an inventory source workbook."
        Exit Function
    End If
    If Not TryResolveInventorySourceWarehouseIdPublisher(wb, resolvedWarehouseId) Then
        report = "Inventory source warehouse could not be resolved."
        Exit Function
    End If
    If Not EnsureWarehouseConfigLoadedPublisher(resolvedWarehouseId, report) Then Exit Function

    Set publishWb = wb
    If RequiresRuntimeCatalogSyncPublisher(wb) Then
        runtimePath = ResolveRuntimeInventoryPathPublisher(resolvedWarehouseId)
        runtimeWasOpen = WorkbookIsOpenByPathPublisher(runtimePath)
        Set publishWb = modInventoryApply.ResolveInventoryWorkbook(resolvedWarehouseId)
        If publishWb Is Nothing Then
            report = "Canonical runtime inventory workbook could not be resolved."
            Exit Function
        End If
        If Not SyncManagedCatalogFromWorkbookPublisher(wb, publishWb, report) Then GoTo CleanExit
    End If

    publishKey = BuildPublishKeyPublisher(publishWb, resolvedWarehouseId)
    If ShouldSkipRecentPublishPublisher(publishKey) Then
        report = "SKIPPED_RECENT"
        EnsureSnapshotPublicationForWorkbook = True
        GoTo CleanExit
    End If

    snapshotPath = vbNullString
    If Not modWarehouseSync.GenerateWarehouseSnapshot(resolvedWarehouseId, publishWb, "", Nothing, snapshotPath) Then
        report = snapshotPath
        GoTo CleanExit
    End If

    RecordRecentPublishPublisher publishKey
    report = snapshotPath
    EnsureSnapshotPublicationForWorkbook = True
    
CleanExit:
    If Not publishWb Is Nothing Then
        If RequiresRuntimeCatalogSyncPublisher(wb) Then
            If Not runtimeWasOpen Then CloseWorkbookQuietlyPublisher publishWb
        End If
    End If
    Exit Function

FailPublish:
    report = "EnsureSnapshotPublicationForWorkbook failed: " & Err.Description
    If Not publishWb Is Nothing Then
        If RequiresRuntimeCatalogSyncPublisher(wb) Then
            If Not runtimeWasOpen Then CloseWorkbookQuietlyPublisher publishWb
        End If
    End If
End Function

Public Sub HandlePotentialInventoryWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    Dim report As String

    Call EnsureSnapshotPublicationForWorkbook(targetWb, report)
End Sub

Private Function ResolveCandidateInventoryWorkbookPublisher() As Workbook
    If Not Application.ActiveWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook.IsAddin Then
            If IsInventorySourceWorkbookPublisher(Application.ActiveWorkbook) Then
                Set ResolveCandidateInventoryWorkbookPublisher = Application.ActiveWorkbook
            End If
        End If
    End If
End Function

Private Function IsInventorySourceWorkbookPublisher(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function

    IsInventorySourceWorkbookPublisher = IsRuntimeInventoryWorkbookPublisher(wb) Or IsManagedCatalogSourceWorkbookPublisher(wb)
End Function

Private Function IsRuntimeInventoryWorkbookPublisher(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function

    IsRuntimeInventoryWorkbookPublisher = WorkbookHasTablePublisher(wb, TABLE_INVENTORY_LOG_PUBLISHER) _
        And WorkbookHasTablePublisher(wb, TABLE_APPLIED_EVENTS_PUBLISHER) _
        And WorkbookHasTablePublisher(wb, TABLE_SKU_BALANCE_PUBLISHER) _
        And WorkbookHasTablePublisher(wb, TABLE_LOCATION_BALANCE_PUBLISHER)
End Function

Private Function IsLegacyManagedInventoryWorkbookPublisher(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function
    If HasRoleOperationalTablesPublisher(wb) Then Exit Function

    IsLegacyManagedInventoryWorkbookPublisher = Not (FindManagedInventoryTablePublisher(wb) Is Nothing)
End Function

Private Function IsManagedCatalogSourceWorkbookPublisher(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function
    If IsRuntimeInventoryWorkbookPublisher(wb) Then Exit Function
    If FindManagedInventoryTablePublisher(wb) Is Nothing Then Exit Function

    If IsLegacyManagedInventoryWorkbookPublisher(wb) Then
        IsManagedCatalogSourceWorkbookPublisher = True
        Exit Function
    End If

    IsManagedCatalogSourceWorkbookPublisher = HasManagedCatalogSourceMarkersPublisher(wb)
End Function

Private Function RequiresRuntimeCatalogSyncPublisher(ByVal wb As Workbook) As Boolean
    RequiresRuntimeCatalogSyncPublisher = IsManagedCatalogSourceWorkbookPublisher(wb) And Not IsRuntimeInventoryWorkbookPublisher(wb)
End Function

Private Function HasRoleOperationalTablesPublisher(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function

    HasRoleOperationalTablesPublisher = WorkbookHasTablePublisher(wb, TABLE_RECEIVED_TALLY_PUBLISHER) _
        Or WorkbookHasTablePublisher(wb, TABLE_SHIPMENTS_TALLY_PUBLISHER) _
        Or WorkbookHasTablePublisher(wb, TABLE_PRODUCTION_OUTPUT_PUBLISHER) _
        Or WorkbookHasTablePublisher(wb, TABLE_ADMIN_AUDIT_PUBLISHER)
End Function

Private Function HasManagedCatalogSourceMarkersPublisher(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then Exit Function

    wbName = LCase$(Trim$(wb.Name))
    If wbName Like "*inventory_management*.xls*" Then
        HasManagedCatalogSourceMarkersPublisher = True
        Exit Function
    End If

    If CountRoleOperationalTablesPublisher(wb) > 1 Then
        HasManagedCatalogSourceMarkersPublisher = True
        Exit Function
    End If

    HasManagedCatalogSourceMarkersPublisher = WorkbookHasTablePublisher(wb, TABLE_RECIPES_PUBLISHER) _
        Or WorkbookHasTablePublisher(wb, TABLE_TEMPLATES_PUBLISHER) _
        Or WorkbookHasTablePublisher(wb, TABLE_INGREDIENT_PALETTE_PUBLISHER)
End Function

Private Function CountRoleOperationalTablesPublisher(ByVal wb As Workbook) As Long
    If wb Is Nothing Then Exit Function

    If WorkbookHasTablePublisher(wb, TABLE_RECEIVED_TALLY_PUBLISHER) Then CountRoleOperationalTablesPublisher = CountRoleOperationalTablesPublisher + 1
    If WorkbookHasTablePublisher(wb, TABLE_SHIPMENTS_TALLY_PUBLISHER) Then CountRoleOperationalTablesPublisher = CountRoleOperationalTablesPublisher + 1
    If WorkbookHasTablePublisher(wb, TABLE_PRODUCTION_OUTPUT_PUBLISHER) Then CountRoleOperationalTablesPublisher = CountRoleOperationalTablesPublisher + 1
    If WorkbookHasTablePublisher(wb, TABLE_ADMIN_AUDIT_PUBLISHER) Then CountRoleOperationalTablesPublisher = CountRoleOperationalTablesPublisher + 1
End Function

Private Function FindManagedInventoryTablePublisher(ByVal wb As Workbook) As ListObject
    Set FindManagedInventoryTablePublisher = FindListObjectByNamePublisher(wb, TABLE_INVSYS_PUBLISHER)
End Function

Private Function WorkbookHasTablePublisher(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasTablePublisher = Not (FindListObjectByNamePublisher(wb, tableName) Is Nothing)
End Function

Private Function FindListObjectByNamePublisher(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindListObjectByNamePublisher = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindListObjectByNamePublisher Is Nothing Then Exit Function
    Next ws
End Function

Private Function TryResolveInventorySourceWarehouseIdPublisher(ByVal wb As Workbook, ByRef warehouseId As String) As Boolean
    warehouseId = ResolveWarehouseIdFromLedgerStatusPublisher(wb)
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    warehouseId = ResolveWarehouseIdFromInventoryWorkbookNamePublisher(wb.Name)
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    warehouseId = ResolveWarehouseIdFromSiblingConfigPublisher(wb)
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    warehouseId = ResolveWarehouseIdFromOpenConfigPublisher()
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    warehouseId = ResolveWarehouseIdFromRuntimeConfigScanPublisher()
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    If modConfig.IsLoaded() Then
        warehouseId = Trim$(modConfig.GetWarehouseId())
        TryResolveInventorySourceWarehouseIdPublisher = (warehouseId <> "")
    End If
End Function

Private Function ResolveWarehouseIdFromLedgerStatusPublisher(ByVal wb As Workbook) As String
    Dim lo As ListObject
    Dim idx As Long

    Set lo = FindListObjectByNamePublisher(wb, TABLE_LEDGER_STATUS_PUBLISHER)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idx = lo.ListColumns("WarehouseId").Index
    ResolveWarehouseIdFromLedgerStatusPublisher = Trim$(CStr(lo.DataBodyRange.Cells(1, idx).Value))
End Function

Private Function ResolveWarehouseIdFromInventoryWorkbookNamePublisher(ByVal wbName As String) As String
    Dim markerPos As Long

    markerPos = InStr(1, wbName, ".invSys.Data.Inventory.", vbTextCompare)
    If markerPos > 1 Then ResolveWarehouseIdFromInventoryWorkbookNamePublisher = Left$(wbName, markerPos - 1)
End Function

Private Function ResolveWarehouseIdFromSiblingConfigPublisher(ByVal wb As Workbook) As String
    Dim folderPath As String
    Dim fileName As String
    Dim candidate As String
    Dim current As String

    If wb Is Nothing Then Exit Function
    folderPath = GetParentFolderPublisher(wb.FullName)
    If folderPath = "" Then Exit Function

    fileName = Dir$(NormalizeFolderPathPublisher(folderPath) & "*.invSys.Config.xls*")
    Do While fileName <> ""
        candidate = ResolveWarehouseIdFromConfigWorkbookNamePublisher(fileName)
        If candidate <> "" Then
            If current = "" Then
                current = candidate
            ElseIf StrComp(current, candidate, vbTextCompare) <> 0 Then
                Exit Function
            End If
        End If
        fileName = Dir$
    Loop

    ResolveWarehouseIdFromSiblingConfigPublisher = current
End Function

Private Function ResolveWarehouseIdFromOpenConfigPublisher() As String
    Dim wb As Workbook
    Dim candidate As String
    Dim current As String

    For Each wb In Application.Workbooks
        candidate = ResolveWarehouseIdFromConfigWorkbookNamePublisher(wb.Name)
        If candidate <> "" Then
            If current = "" Then
                current = candidate
            ElseIf StrComp(current, candidate, vbTextCompare) <> 0 Then
                Exit Function
            End If
        End If
    Next wb

    ResolveWarehouseIdFromOpenConfigPublisher = current
End Function

Private Function ResolveWarehouseIdFromRuntimeConfigScanPublisher() As String
    Dim wbCfg As Workbook
    Dim report As String
    Dim wasOpen As Boolean

    Set wbCfg = FindOpenConfigWorkbookPublisher()
    wasOpen = Not wbCfg Is Nothing
    If wbCfg Is Nothing Then Set wbCfg = modRuntimeWorkbooks.OpenFirstRuntimeConfigWorkbook(report)
    If wbCfg Is Nothing Then Exit Function

    ResolveWarehouseIdFromRuntimeConfigScanPublisher = ResolveWarehouseIdFromConfigWorkbookNamePublisher(wbCfg.Name)
    If ResolveWarehouseIdFromRuntimeConfigScanPublisher = "" Then
        ResolveWarehouseIdFromRuntimeConfigScanPublisher = ResolveWarehouseIdFromConfigTablePublisher(wbCfg)
    End If

    If Not wasOpen Then CloseWorkbookQuietlyPublisher wbCfg
End Function

Private Function ResolveWarehouseIdFromConfigTablePublisher(ByVal wbCfg As Workbook) As String
    Dim lo As ListObject

    Set lo = FindListObjectByNamePublisher(wbCfg, "tblWarehouseConfig")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ResolveWarehouseIdFromConfigTablePublisher = Trim$(CStr(lo.DataBodyRange.Cells(1, lo.ListColumns("WarehouseId").Index).Value))
End Function

Private Function ResolveWarehouseIdFromConfigWorkbookNamePublisher(ByVal wbName As String) As String
    Dim markerPos As Long

    markerPos = InStr(1, wbName, ".invSys.Config.", vbTextCompare)
    If markerPos > 1 Then ResolveWarehouseIdFromConfigWorkbookNamePublisher = Left$(wbName, markerPos - 1)
End Function

Private Function EnsureWarehouseConfigLoadedPublisher(ByVal warehouseId As String, ByRef report As String) As Boolean
    If Trim$(warehouseId) = "" Then
        report = "WarehouseId not resolved."
        Exit Function
    End If

    If modConfig.IsLoaded() Then
        If StrComp(Trim$(modConfig.GetWarehouseId()), Trim$(warehouseId), vbTextCompare) = 0 Then
            EnsureWarehouseConfigLoadedPublisher = True
            Exit Function
        End If
    End If

    EnsureWarehouseConfigLoadedPublisher = modConfig.LoadConfig(warehouseId, "")
    If Not EnsureWarehouseConfigLoadedPublisher Then report = "Config load failed for " & warehouseId & "."
End Function

Private Function BuildPublishKeyPublisher(ByVal wb As Workbook, ByVal warehouseId As String) As String
    If wb Is Nothing Then Exit Function
    BuildPublishKeyPublisher = LCase$(wb.FullName & "|" & warehouseId)
End Function

Private Function ResolveRuntimeInventoryPathPublisher(ByVal warehouseId As String) As String
    Dim rootPath As String

    rootPath = Trim$(modConfig.GetString("PathDataRoot", ""))
    If rootPath = "" Then Exit Function
    If Right$(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
    ResolveRuntimeInventoryPathPublisher = rootPath & warehouseId & ".invSys.Data.Inventory.xlsb"
End Function

Private Function WorkbookIsOpenByPathPublisher(ByVal fullPath As String) As Boolean
    Dim wb As Workbook

    If Trim$(fullPath) = "" Then Exit Function
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            WorkbookIsOpenByPathPublisher = True
            Exit Function
        End If
    Next wb
End Function

Private Function SyncManagedCatalogFromWorkbookPublisher(ByVal sourceWb As Workbook, _
                                                         ByVal runtimeWb As Workbook, _
                                                         ByRef report As String) As Boolean
    On Error GoTo FailSync

    Dim schemaReport As String
    Dim sourceLo As ListObject
    Dim targetLo As ListObject
    Dim targetWs As Worksheet
    Dim sheetWasProtected As Boolean
    Dim rowIndex As Long
    Dim sku As String
    Dim seen As Object

    If sourceWb Is Nothing Or runtimeWb Is Nothing Then
        report = "Source or runtime workbook not resolved."
        Exit Function
    End If
    If Not modInventorySchema.EnsureInventorySchema(runtimeWb, schemaReport) Then
        report = schemaReport
        Exit Function
    End If

    Set sourceLo = FindManagedInventoryTablePublisher(sourceWb)
    If sourceLo Is Nothing Then
        report = "Managed inventory source table not found."
        Exit Function
    End If

    Set targetLo = FindListObjectByNamePublisher(runtimeWb, TABLE_SKU_CATALOG_PUBLISHER)
    If targetLo Is Nothing Then
        report = "Runtime SKU catalog table not found."
        Exit Function
    End If

    Set targetWs = targetLo.Parent
    sheetWasProtected = targetWs.ProtectContents
    EnsureWorksheetEditablePublisher targetWs, TABLE_SKU_CATALOG_PUBLISHER

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    ClearListObjectRowsPublisher targetLo

    If Not sourceLo.DataBodyRange Is Nothing Then
        For rowIndex = 1 To sourceLo.ListRows.Count
            sku = ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "ITEM_CODE")
            If sku = "" Then sku = ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "SKU")
            If sku = "" Then GoTo ContinueLoop
            If seen.Exists(sku) Then GoTo ContinueLoop
            seen.Add sku, True
            AppendCatalogRowPublisher targetLo, sourceLo, rowIndex, sku
ContinueLoop:
        Next rowIndex
    End If

    runtimeWb.Save
    report = "CatalogRows=" & CStr(seen.Count)
    SyncManagedCatalogFromWorkbookPublisher = True
    If sheetWasProtected Then RestoreWorksheetProtectionPublisher targetWs
    Exit Function

FailSync:
    report = "SyncManagedCatalogFromWorkbookPublisher failed: " & Err.Description
    On Error Resume Next
    If sheetWasProtected Then RestoreWorksheetProtectionPublisher targetWs
End Function

Private Sub ClearListObjectRowsPublisher(ByVal lo As ListObject)
    Do While Not lo Is Nothing
        If lo.DataBodyRange Is Nothing Then Exit Do
        lo.ListRows(1).Delete
    Loop
End Sub

Private Sub AppendCatalogRowPublisher(ByVal targetLo As ListObject, _
                                      ByVal sourceLo As ListObject, _
                                      ByVal rowIndex As Long, _
                                      ByVal sku As String)
    Dim lr As ListRow
    Dim itemValue As String

    If targetLo Is Nothing Then Exit Sub
    Set lr = targetLo.ListRows.Add

    itemValue = ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "ITEM")
    If itemValue = "" Then itemValue = sku

    SetCatalogCellPublisher targetLo, lr.Index, "SKU", sku
    SetCatalogCellPublisher targetLo, lr.Index, "ITEM_CODE", sku
    SetCatalogCellPublisher targetLo, lr.Index, "ITEM", itemValue
    SetCatalogCellPublisher targetLo, lr.Index, "UOM", ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "UOM")
    SetCatalogCellPublisher targetLo, lr.Index, "LOCATION", ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "LOCATION")
    SetCatalogCellPublisher targetLo, lr.Index, "DESCRIPTION", ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "DESCRIPTION")
    SetCatalogCellPublisher targetLo, lr.Index, "VENDOR(s)", ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "VENDOR(s)")
    SetCatalogCellPublisher targetLo, lr.Index, "VENDOR_CODE", ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "VENDOR_CODE")
    SetCatalogCellPublisher targetLo, lr.Index, "CATEGORY", ResolveCatalogSourceValuePublisher(sourceLo, rowIndex, "CATEGORY")
End Sub

Private Function ResolveCatalogSourceValuePublisher(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim idx As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    On Error Resume Next
    idx = lo.ListColumns(columnName).Index
    On Error GoTo 0
    If idx = 0 Then Exit Function
    ResolveCatalogSourceValuePublisher = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, idx).Value))
End Function

Private Sub SetCatalogCellPublisher(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    If lo Is Nothing Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueIn
End Sub

Private Sub EnsureWorksheetEditablePublisher(ByVal ws As Worksheet, ByVal context As String)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 4601, "modInventoryPublisher.EnsureWorksheetEditablePublisher", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected before updating " & context & "."
    End If
End Sub

Private Sub RestoreWorksheetProtectionPublisher(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0

    If Not ws.ProtectContents Then
        Err.Raise vbObjectError + 4602, "modInventoryPublisher.RestoreWorksheetProtectionPublisher", _
                  "Worksheet '" & ws.Name & "' could not be reprotected after catalog sync."
    End If
End Sub

Private Sub EnsureRecentPublishRegistryPublisher()
    If mRecentPublishes Is Nothing Then
        Set mRecentPublishes = CreateObject("Scripting.Dictionary")
        mRecentPublishes.CompareMode = vbTextCompare
    End If
End Sub

Private Function ShouldSkipRecentPublishPublisher(ByVal publishKey As String) As Boolean
    EnsureRecentPublishRegistryPublisher
    If publishKey = "" Then Exit Function
    If Not mRecentPublishes.Exists(publishKey) Then Exit Function
    ShouldSkipRecentPublishPublisher = (DateDiff("s", CDate(mRecentPublishes(publishKey)), Now) < MIN_RECENT_PUBLISH_SECONDS)
End Function

Private Sub RecordRecentPublishPublisher(ByVal publishKey As String)
    EnsureRecentPublishRegistryPublisher
    If publishKey = "" Then Exit Sub
    mRecentPublishes(publishKey) = Now
End Sub

Private Function GetParentFolderPublisher(ByVal fullPath As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(fullPath, "\")
    If slashPos > 1 Then GetParentFolderPublisher = Left$(fullPath, slashPos - 1)
End Function

Private Function NormalizeFolderPathPublisher(ByVal folderPath As String) As String
    NormalizeFolderPathPublisher = Trim$(folderPath)
    If NormalizeFolderPathPublisher = "" Then Exit Function
    If Right$(NormalizeFolderPathPublisher, 1) <> "\" Then NormalizeFolderPathPublisher = NormalizeFolderPathPublisher & "\"
End Function

Private Function FindOpenConfigWorkbookPublisher() As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If LCase$(wb.Name) Like "*.invsys.config.xlsb" Then
            Set FindOpenConfigWorkbookPublisher = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub CloseWorkbookQuietlyPublisher(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub
