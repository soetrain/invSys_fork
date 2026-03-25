Attribute VB_Name = "TestPhase6CoreSurfaces"
Option Explicit

Public Function TestOpenOrCreateConfigWorkbookRuntime_CreatesCanonicalWorkbook() As Long
    Dim rootPath As String
    Dim wb As Workbook
    Dim loWh As ListObject
    Dim loSt As ListObject
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_cfg_surface")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wb = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH61", "S1", rootPath, report)
    If wb Is Nothing Then GoTo CleanExit

    Set loWh = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wb.Worksheets("StationConfig").ListObjects("tblStationConfig")

    If StrComp(wb.Name, "WH61.invSys.Config.xlsb", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loWh, 1, "WarehouseId")), "WH61", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loSt, 1, "StationId")), "S1", vbTextCompare) = 0 _
       And Len(Dir$(rootPath & "\WH61.invSys.Config.xlsb")) > 0 Then
        TestOpenOrCreateConfigWorkbookRuntime_CreatesCanonicalWorkbook = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wb
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLoadConfig_AutoBootstrapsCanonicalWorkbook() As Long
    Dim rootPath As String
    Dim wb As Workbook

    rootPath = BuildRuntimeTestRoot("phase6_cfg_load")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH62", "S2") Then GoTo CleanExit

    Set wb = FindWorkbookByName("WH62.invSys.Config.xlsb")
    If Not wb Is Nothing _
       And StrComp(modConfig.GetWarehouseId(), "WH62", vbTextCompare) = 0 _
       And StrComp(modConfig.GetStationId(), "S2", vbTextCompare) = 0 _
       And Len(Dir$(rootPath & "\WH62.invSys.Config.xlsb")) > 0 Then
        TestLoadConfig_AutoBootstrapsCanonicalWorkbook = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wb
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLoadConfig_BlankContextAutoBootstrapsDefaultRuntimeWorkbook() As Long
    Dim rootPath As String
    Dim wb As Workbook

    rootPath = BuildRuntimeTestRoot("phase6_cfg_blank")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("", "") Then GoTo CleanExit

    Set wb = FindWorkbookByName("WH1.invSys.Config.xlsb")
    If Not wb Is Nothing _
       And StrComp(modConfig.GetWarehouseId(), "WH1", vbTextCompare) = 0 _
       And StrComp(modConfig.GetStationId(), "S1", vbTextCompare) = 0 _
       And Len(Dir$(rootPath & "\WH1.invSys.Config.xlsb")) > 0 Then
        TestLoadConfig_BlankContextAutoBootstrapsDefaultRuntimeWorkbook = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wb
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLoadConfig_QuarantinesContaminatedConfigSheet() As Long
    Dim rootPath As String
    Dim wb As Workbook
    Dim loSt As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_cfg_quarantine")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wb = CreateContaminatedConfigWorkbook(rootPath, "WH64")
    If wb Is Nothing Then GoTo CleanExit

    If Not modConfig.LoadConfig("WH64", "S4") Then GoTo CleanExit
    Set wb = FindWorkbookByName("WH64.invSys.Config.xlsb")
    If wb Is Nothing Then GoTo CleanExit

    Set loSt = wb.Worksheets("StationConfig").ListObjects("tblStationConfig")
    If Not loSt Is Nothing _
       And FindWorksheetByPrefix(wb, "StationConfig_Stale") > 0 _
       And StrComp(CStr(GetTableValue(loSt, 1, "StationId")), "S4", vbTextCompare) = 0 Then
        TestLoadConfig_QuarantinesContaminatedConfigSheet = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wb
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLoadAuth_AutoBootstrapsCanonicalWorkbook() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim loUsers As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_auth_load")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH63", "S3") Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH63") Then GoTo CleanExit

    Set wbCfg = FindWorkbookByName("WH63.invSys.Config.xlsb")
    Set wbAuth = FindWorkbookByName("WH63.invSys.Auth.xlsb")
    If wbAuth Is Nothing Then GoTo CleanExit

    Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    If FindUserRow(loUsers, "svc_processor") > 0 _
       And Not wbCfg Is Nothing _
       And Len(Dir$(rootPath & "\WH63.invSys.Auth.xlsb")) > 0 Then
        TestLoadAuth_AutoBootstrapsCanonicalWorkbook = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbAuth
    CloseWorkbookIfOpen wbCfg
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLoadAuth_BootstrapGrantsCurrentOperatorCapabilities() As Long
    Dim rootPath As String
    Dim currentUser As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_caps")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH65", "S5") Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH65") Then GoTo CleanExit

    currentUser = Trim$(Environ$("USERNAME"))
    If currentUser = "" Then currentUser = Trim$(Application.UserName)
    If currentUser = "" Then GoTo CleanExit

    If modAuth.CanPerform("RECEIVE_POST", currentUser, "WH65", "S5", "TEST", "AUTH-RECV") _
       And modAuth.CanPerform("SHIP_POST", currentUser, "WH65", "S5", "TEST", "AUTH-SHIP") _
       And modAuth.CanPerform("PROD_POST", currentUser, "WH65", "S5", "TEST", "AUTH-PROD") _
       And modAuth.CanPerform("INBOX_PROCESS", "svc_processor", "WH65", "S5", "TEST", "AUTH-PROC") Then
        TestLoadAuth_BootstrapGrantsCurrentOperatorCapabilities = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestResolveInventoryWorkbookBridge_PrefersCanonicalWorkbookOverOperatorSurface() As Long
    Dim rootPath As String
    Dim wbOperator As Workbook
    Dim wbInventory As Workbook
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_inv_bridge")

    On Error GoTo CleanFail
    Set wbOperator = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wbOperator, report) Then GoTo CleanExit

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wbInventory = modInventoryDomainBridge.ResolveInventoryWorkbookBridge("WH66")
    If wbInventory Is Nothing Then GoTo CleanExit

    If StrComp(wbInventory.Name, "WH66.invSys.Data.Inventory.xlsb", vbTextCompare) = 0 _
       And StrComp(wbInventory.Name, wbOperator.Name, vbTextCompare) <> 0 _
       And Len(Dir$(rootPath & "\WH66.invSys.Data.Inventory.xlsb")) > 0 Then
        TestResolveInventoryWorkbookBridge_PrefersCanonicalWorkbookOverOperatorSurface = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInventory
    CloseWorkbookIfOpen wbOperator
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureInventoryManagementSurface_RemovesDomainArtifacts() As Long
    Dim wb As Workbook
    Dim report As String

    On Error GoTo CleanFail
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    AddNamedWorksheetWithMarker wb, "InventoryLog", "legacy-log"
    AddNamedWorksheetWithMarker wb, "AppliedEvents", "legacy-applied"
    AddNamedWorksheetWithMarker wb, "Locks", "legacy-locks"

    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wb, report) Then GoTo CleanExit

    If WorksheetExistsByName(wb, "InventoryManagement") _
       And Not WorksheetExistsByName(wb, "InventoryLog") _
       And Not WorksheetExistsByName(wb, "AppliedEvents") _
       And Not WorksheetExistsByName(wb, "Locks") _
       And HasTableByName(wb, "invSys") Then
        TestEnsureInventoryManagementSurface_RemovesDomainArtifacts = 1
    End If

CleanExit:
    CloseWorkbookIfOpen wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestOpenOrCreateConfigWorkbookRuntime_PrunesUnexpectedSheets() As Long
    Dim rootPath As String
    Dim wb As Workbook
    Dim extraWs As Worksheet
    Dim targetPath As String
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_cfg_prune")

    On Error GoTo CleanFail
    targetPath = rootPath & "\WH67.invSys.Config.xlsb"
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "WarehouseConfig"
    Set extraWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    extraWs.Name = "StationConfig"
    Set extraWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    extraWs.Name = "ReceivedTally"
    extraWs.Range("A1").Value = "legacy-surface"
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    wb.Close SaveChanges:=False
    Set wb = Nothing

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wb = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH67", "S7", rootPath, report)
    If wb Is Nothing Then GoTo CleanExit

    If wb.Worksheets.Count = 2 _
       And WorksheetExistsByName(wb, "WarehouseConfig") _
       And WorksheetExistsByName(wb, "StationConfig") _
       And Not WorksheetExistsByName(wb, "ReceivedTally") Then
        TestOpenOrCreateConfigWorkbookRuntime_PrunesUnexpectedSheets = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wb
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRefreshInventoryReadModelFromSnapshot_UpdatesReadModelAndMetadata() As Long
    Dim rootPath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH68", "S8") Then GoTo CleanExit
    SetConfigWarehouseValue "WH68.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    AddInvSysSeedRow loInv, 901, "SKU-RM-001", "Read Model Item", "EA", "A1", 99

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH68", "SKU-RM-001", 7, CDate("2026-03-24 17:30:00"))
    If wbSnap Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH68", "LOCAL", report) Then GoTo CleanExit

    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 7 _
       And CDbl(GetTableValue(loInv, 1, "QtyOnHand")) = 7 _
       And CDbl(GetTableValue(loInv, 1, "QtyAvailable")) = 7 _
       And StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "A1", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "SKU")), "SKU-RM-001", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "ItemName")), "Read Model Item", vbTextCompare) = 0 _
       And InStr(1, CStr(GetTableValue(loInv, 1, "LocationSummary")), "A1", vbTextCompare) > 0 _
       And CBool(GetTableValue(loInv, 1, "IsStale")) = False _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "LOCAL", vbTextCompare) = 0 _
       And Trim$(CStr(GetTableValue(loInv, 1, "SnapshotId"))) <> "" _
       And IsDate(GetTableValue(loInv, 1, "LastRefreshUTC")) _
       And IsDate(GetTableValue(loInv, 1, "LastAppliedUTC")) Then
        TestRefreshInventoryReadModelFromSnapshot_UpdatesReadModelAndMetadata = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRefreshInventoryReadModel_MissingSnapshotMarksStaleWithoutMutatingReceivingTally() As Long
    Dim rootPath As String
    Dim wbOps As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loRecv As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model_missing")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH69", "S9") Then GoTo CleanExit
    SetConfigWarehouseValue "WH69.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    Set loRecv = wbOps.Worksheets("ReceivedTally").ListObjects("ReceivedTally")
    AddInvSysSeedRow loInv, 902, "SKU-RM-002", "Stale Item", "EA", "B1", 12
    AddReceivedTallyRow loRecv, "REF-ST-001", "Stale Item", 3, 902

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH69", "LOCAL", report) Then GoTo CleanExit

    If CBool(GetTableValue(loInv, 1, "IsStale")) = True _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "CACHED", vbTextCompare) = 0 _
       And CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 12 _
       And loRecv.ListRows.Count = 1 _
       And StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), "REF-ST-001", vbTextCompare) = 0 Then
        TestRefreshInventoryReadModel_MissingSnapshotMarksStaleWithoutMutatingReceivingTally = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function GetTableValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    GetTableValue = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value
End Function

Private Function FindUserRow(ByVal lo As ListObject, ByVal userId As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(i, lo.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 Then
            FindUserRow = i
            Exit Function
        End If
    Next i
End Function

Private Function FindWorkbookByName(ByVal workbookName As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            Set FindWorkbookByName = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindWorksheetByPrefix(ByVal wb As Workbook, ByVal prefixText As String) As Long
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        If StrComp(Left$(ws.Name, Len(prefixText)), prefixText, vbTextCompare) = 0 Then
            FindWorksheetByPrefix = ws.Index
            Exit Function
        End If
    Next ws
End Function

Private Function WorksheetExistsByName(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            WorksheetExistsByName = True
            Exit Function
        End If
    Next ws
End Function

Private Function HasTableByName(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    HasTableByName = Not FindTableByName(wb, tableName) Is Nothing
End Function

Private Function FindTableByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTableByName = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTableByName Is Nothing Then Exit Function
    Next ws
End Function

Private Sub AddNamedWorksheetWithMarker(ByVal wb As Workbook, ByVal sheetName As String, ByVal markerText As String)
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
    ws.Range("A1").Value = markerText
End Sub

Private Sub AddInvSysSeedRow(ByVal lo As ListObject, ByVal rowValue As Long, ByVal sku As String, ByVal itemName As String, ByVal uom As String, ByVal locationVal As String, ByVal totalInv As Double)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    Set lr = lo.ListRows.Add
    SetTableCell lo, lr.Index, "ROW", rowValue
    SetTableCell lo, lr.Index, "ITEM_CODE", sku
    SetTableCell lo, lr.Index, "ITEM", itemName
    SetTableCell lo, lr.Index, "UOM", uom
    SetTableCell lo, lr.Index, "LOCATION", locationVal
    SetTableCell lo, lr.Index, "TOTAL INV", totalInv
End Sub

Private Sub AddReceivedTallyRow(ByVal lo As ListObject, ByVal refNumber As String, ByVal itemName As String, ByVal qty As Double, ByVal rowValue As Long)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf Trim$(CStr(GetTableValue(lo, 1, "REF_NUMBER"))) = "" _
        And Trim$(CStr(GetTableValue(lo, 1, "ITEMS"))) = "" _
        And NzDblForTest(GetTableValue(lo, 1, "QUANTITY")) = 0 Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "REF_NUMBER", refNumber
    SetTableCell lo, lr.Index, "ITEMS", itemName
    SetTableCell lo, lr.Index, "QUANTITY", qty
    SetTableCell lo, lr.Index, "ROW", rowValue
End Sub

Private Sub SetTableCell(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    If lo Is Nothing Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueIn
End Sub

Private Sub SetConfigWarehouseValue(ByVal workbookName As String, ByVal columnName As String, ByVal valueIn As Variant)
    Dim wb As Workbook
    Dim lo As ListObject

    Set wb = FindWorkbookByName(workbookName)
    If wb Is Nothing Then Exit Sub
    Set lo = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    If lo Is Nothing Then Exit Sub
    lo.DataBodyRange.Cells(1, lo.ListColumns(columnName).Index).Value = valueIn
    wb.Save
End Sub

Private Function NzDblForTest(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzDblForTest = CDbl(valueIn)
End Function

Private Function CreateSnapshotWorkbook(ByVal rootPath As String, ByVal warehouseId As String, ByVal sku As String, ByVal qtyOnHand As Double, ByVal lastAppliedUtc As Date) As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim targetPath As String

    targetPath = rootPath & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "InventorySnapshot"
    ws.Range("A1").Value = "WarehouseId"
    ws.Range("B1").Value = "SKU"
    ws.Range("C1").Value = "QtyOnHand"
    ws.Range("D1").Value = "QtyAvailable"
    ws.Range("E1").Value = "LocationSummary"
    ws.Range("F1").Value = "LastAppliedAtUTC"
    ws.Range("A2").Value = warehouseId
    ws.Range("B2").Value = sku
    ws.Range("C2").Value = qtyOnHand
    ws.Range("D2").Value = qtyOnHand
    ws.Range("E2").Value = "A1=" & CStr(CLng(qtyOnHand))
    ws.Range("F2").Value = lastAppliedUtc
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F2"), , xlYes)
    lo.Name = "tblInventorySnapshot"
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    Set CreateSnapshotWorkbook = wb
End Function

Private Sub CloseWorkbookIfOpen(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function BuildRuntimeTestRoot(ByVal baseName As String) As String
    BuildRuntimeTestRoot = Environ$("TEMP") & "\" & baseName & "_" & Format$(Now, "yyyymmdd_hhnnss")
    If Len(Dir$(BuildRuntimeTestRoot, vbDirectory)) = 0 Then MkDir BuildRuntimeTestRoot
End Function

Private Sub DeleteRuntimeRoot(ByVal rootPath As String)
    Dim fileName As String

    On Error Resume Next
    fileName = Dir$(rootPath & "\*.*")
    Do While fileName <> ""
        Kill rootPath & "\" & fileName
        fileName = Dir$
    Loop
    If Len(Dir$(rootPath, vbDirectory)) > 0 Then RmDir rootPath
    On Error GoTo 0
End Sub

Private Function CreateContaminatedConfigWorkbook(ByVal rootPath As String, ByVal warehouseId As String) As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim targetPath As String

    targetPath = rootPath & "\" & warehouseId & ".invSys.Config.xlsb"
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "WarehouseConfig"
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "StationConfig"
    ws.Range("A1").Value = "PROCESS"
    ws.Range("B1").Value = "OUTPUT"
    ws.Range("C1").Value = "ROW"
    ws.Range("A2").Value = "Mix"
    ws.Range("B2").Value = "Widget"
    ws.Range("C2").Value = 1
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C2"), , xlYes)
    lo.Name = "ProductionOutput"
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    Set CreateContaminatedConfigWorkbook = wb
End Function
