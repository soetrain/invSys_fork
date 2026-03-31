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
    If StrComp(modConfig.GetResolvedWorkbookName(), "WH62.invSys.Config.xlsb", vbTextCompare) = 0 _
       And StrComp(modConfig.GetWarehouseId(), "WH62", vbTextCompare) = 0 _
       And StrComp(modConfig.GetStationId(), "S2", vbTextCompare) = 0 _
       And Len(Dir$(rootPath & "\WH62.invSys.Config.xlsb")) > 0 _
       And (wb Is Nothing Or StrComp(wb.FullName, rootPath & "\WH62.invSys.Config.xlsb", vbTextCompare) = 0) Then
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
    If StrComp(modConfig.GetResolvedWorkbookName(), "WH1.invSys.Config.xlsb", vbTextCompare) = 0 _
       And StrComp(modConfig.GetWarehouseId(), "WH1", vbTextCompare) = 0 _
       And StrComp(modConfig.GetStationId(), "S1", vbTextCompare) = 0 _
       And Len(Dir$(rootPath & "\WH1.invSys.Config.xlsb")) > 0 _
       And (wb Is Nothing Or StrComp(wb.FullName, rootPath & "\WH1.invSys.Config.xlsb", vbTextCompare) = 0) Then
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

Public Function TestEnsureStationBootstrap_CreatesLocalConfigAndInbox() As Long
    Dim rootPath As String
    Dim sharedRoot As String
    Dim localRoot As String
    Dim inboxRoot As String
    Dim sharedConfigPath As String
    Dim localConfigPath As String
    Dim inboxPath As String
    Dim report As String
    Dim failureReason As String
    Dim wbSharedCfg As Workbook
    Dim wbLocalCfg As Workbook
    Dim wbInbox As Workbook
    Dim loSt As ListObject
    Dim rowIndex As Long

    rootPath = BuildRuntimeTestRoot("phase6_station_bootstrap")
    sharedRoot = rootPath & "\shared"
    localRoot = rootPath & "\local_cfg"
    inboxRoot = rootPath & "\station_inbox"
    sharedConfigPath = sharedRoot & "\WH63.invSys.Config.xlsb"
    localConfigPath = localRoot & "\WH63.invSys.Config.xlsb"

    On Error GoTo CleanFail
    MkDir sharedRoot
    MkDir localRoot
    MkDir inboxRoot

    If Not modConfig.EnsureStationConfigEntry("WH63", "S2", "ARCTIC-RAPTOR", inboxRoot & "\", "RECEIVE", sharedConfigPath, sharedRoot & "\", report) Then
        failureReason = "Shared config bootstrap failed: " & report
        GoTo CleanExit
    End If

    If Not modConfig.EnsureStationConfigEntry("WH63", "S2", "ARCTIC-RAPTOR", inboxRoot & "\", "RECEIVE", localConfigPath, sharedRoot & "\", report) Then
        failureReason = "Local config bootstrap failed: " & report
        GoTo CleanExit
    End If

    If Not modConfig.EnsureStationInbox("WH63", "S2", "RECEIVE", localConfigPath, inboxPath, report) Then
        failureReason = "Station inbox bootstrap failed: " & report
        GoTo CleanExit
    End If

    Set wbSharedCfg = FindWorkbookByFullPathForTest(sharedConfigPath)
    If wbSharedCfg Is Nothing Then
        If Len(Dir$(sharedConfigPath, vbNormal)) = 0 Then
            failureReason = "Shared config workbook was not created on disk."
            GoTo CleanExit
        End If
        Set wbSharedCfg = Application.Workbooks.Open(sharedConfigPath)
    End If
    If wbSharedCfg Is Nothing Then
        failureReason = "Shared config workbook could not be opened for verification."
        GoTo CleanExit
    End If
    Set loSt = wbSharedCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    rowIndex = FindRowByColumnValueInTable(loSt, "StationId", "S2")
    If rowIndex = 0 Then
        failureReason = "Shared config did not contain station row S2."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loSt, rowIndex, "PathInboxRoot")), inboxRoot & "\", vbTextCompare) <> 0 Then
        failureReason = "Shared config PathInboxRoot was not updated."
        GoTo CleanExit
    End If
    CloseWorkbookIfOpen wbSharedCfg
    Set wbSharedCfg = Nothing

    Set wbLocalCfg = FindWorkbookByFullPathForTest(localConfigPath)
    If wbLocalCfg Is Nothing Then
        If Len(Dir$(localConfigPath, vbNormal)) = 0 Then
            failureReason = "Local config workbook was not created on disk."
            GoTo CleanExit
        End If
        Set wbLocalCfg = Application.Workbooks.Open(localConfigPath)
    End If
    If wbLocalCfg Is Nothing Then
        failureReason = "Local config workbook could not be opened for verification."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(wbLocalCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig"), 1, "PathDataRoot")), sharedRoot, vbTextCompare) <> 0 Then
        failureReason = "Local config PathDataRoot did not point at shared runtime root."
        GoTo CleanExit
    End If

    If StrComp(inboxPath, inboxRoot & "\invSys.Inbox.Receiving.S2.xlsb", vbTextCompare) <> 0 Then
        failureReason = "Returned inbox path did not match expected station inbox."
        GoTo CleanExit
    End If
    If Len(Dir$(inboxPath, vbNormal)) = 0 Then
        failureReason = "Station inbox workbook was not created on disk."
        GoTo CleanExit
    End If

    Set wbInbox = FindWorkbookByName("invSys.Inbox.Receiving.S2.xlsb")
    If wbInbox Is Nothing Then
        Set wbInbox = Application.Workbooks.Open(inboxPath)
    End If
    If wbInbox Is Nothing Then
        failureReason = "Station inbox workbook could not be opened for verification."
        GoTo CleanExit
    End If
    If FindTableByName(wbInbox, "tblInboxReceive") Is Nothing Then
        failureReason = "Station inbox workbook did not contain tblInboxReceive."
        GoTo CleanExit
    End If

    TestEnsureStationBootstrap_CreatesLocalConfigAndInbox = 1

CleanExit:
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbLocalCfg
    CloseWorkbookIfOpen wbSharedCfg
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7111, "TestEnsureStationBootstrap_CreatesLocalConfigAndInbox", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
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
    If Not wbAuth Is Nothing Then Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    If StrComp(modConfig.GetResolvedWorkbookName(), "WH63.invSys.Config.xlsb", vbTextCompare) = 0 _
       And StrComp(modAuth.GetResolvedAuthWorkbookName(), "WH63.invSys.Auth.xlsb", vbTextCompare) = 0 _
       And Len(Dir$(rootPath & "\WH63.invSys.Auth.xlsb")) > 0 _
       And Len(Dir$(rootPath & "\WH63.invSys.Config.xlsb")) > 0 _
       And (wbCfg Is Nothing Or StrComp(wbCfg.FullName, rootPath & "\WH63.invSys.Config.xlsb", vbTextCompare) = 0) _
       And (wbAuth Is Nothing Or (Not loUsers Is Nothing And FindUserRow(loUsers, "svc_processor") > 0 And StrComp(wbAuth.FullName, rootPath & "\WH63.invSys.Auth.xlsb", vbTextCompare) = 0)) Then
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
       And CDbl(GetTableValue(loInv, 1, "QtyAvailable")) = 7 _
       And StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "A1", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "ITEM_CODE")), "SKU-RM-001", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "ITEM")), "Read Model Item", vbTextCompare) = 0 _
       And InStr(1, CStr(GetTableValue(loInv, 1, "LocationSummary")), "A1", vbTextCompare) > 0 _
       And CBool(GetTableValue(loInv, 1, "IsStale")) = False _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "LOCAL", vbTextCompare) = 0 _
       And Trim$(CStr(GetTableValue(loInv, 1, "SnapshotId"))) <> "" _
       And IsDate(GetTableValue(loInv, 1, "LastRefreshUTC")) _
       And IsDate(GetTableValue(loInv, 1, "LAST EDITED")) _
       And IsDate(GetTableValue(loInv, 1, "TOTAL INV LAST EDIT")) Then
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

Public Function TestRefreshInventoryReadModelFromSharePoint_UpdatesReadModelAndMetadata() As Long
    Dim rootPath As String
    Dim shareRoot As String
    Dim snapshotRoot As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model_sharepoint")
    shareRoot = rootPath & "\Share"
    snapshotRoot = shareRoot & "\Snapshots"

    On Error GoTo CleanFail
    MkDir shareRoot
    MkDir snapshotRoot
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH68SP", "S8") Then GoTo CleanExit
    SetConfigWarehouseValue "WH68SP.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    SetConfigWarehouseValue "WH68SP.invSys.Config.xlsb", "PathSharePointRoot", shareRoot
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    AddInvSysSeedRow loInv, 950, "SKU-RM-SP-001", "SharePoint Item", "EA", "OLD", 2

    Set wbSnap = CreateSnapshotWorkbook(snapshotRoot, "WH68SP", "SKU-RM-SP-001", 13, CDate("2026-03-30 08:10:00"), _
                                        12, "SP1=12", "SharePoint Item", "EA", "SP1")
    If wbSnap Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelFromSharePointForWorkbook(wbOps, "WH68SP", report) Then GoTo CleanExit

    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 13 _
       And CDbl(GetTableValue(loInv, 1, "QtyAvailable")) = 12 _
       And StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "SP1", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "SHAREPOINT", vbTextCompare) = 0 _
       And CBool(GetTableValue(loInv, 1, "IsStale")) = False _
       And InStr(1, CStr(GetTableValue(loInv, 1, "SnapshotId")), "WH68SP.invSys.Snapshot.Inventory.xlsb|", vbTextCompare) = 1 _
       And IsDate(GetTableValue(loInv, 1, "LastRefreshUTC")) Then
        TestRefreshInventoryReadModelFromSharePoint_UpdatesReadModelAndMetadata = 1
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

Public Function TestRefreshInventoryReadModelFromSharePoint_StaleSnapshotMarksReadModelStale() As Long
    Dim rootPath As String
    Dim shareRoot As String
    Dim snapshotRoot As String
    Dim canonicalPath As String
    Dim stalePath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model_sharepoint_stale")
    shareRoot = rootPath & "\Share"
    snapshotRoot = shareRoot & "\Snapshots"
    canonicalPath = snapshotRoot & "\WH68ST.invSys.Snapshot.Inventory.xlsb"
    stalePath = snapshotRoot & "\WH68ST.stale.invSys.Snapshot.Inventory.xlsb"

    On Error GoTo CleanFail
    MkDir shareRoot
    MkDir snapshotRoot
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH68ST", "S8") Then GoTo CleanExit
    SetConfigWarehouseValue "WH68ST.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    SetConfigWarehouseValue "WH68ST.invSys.Config.xlsb", "PathSharePointRoot", shareRoot
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    AddInvSysSeedRow loInv, 951, "SKU-RM-SP-STALE", "Stale Share Item", "EA", "OLD", 4

    Set wbSnap = CreateSnapshotWorkbook(snapshotRoot, "WH68ST", "SKU-RM-SP-STALE", 21, CDate("2026-03-30 08:20:00"), _
                                        19, "SP2=19", "Stale Share Item", "EA", "SP2")
    If wbSnap Is Nothing Then GoTo CleanExit
    wbSnap.SaveCopyAs stalePath
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing
    Kill canonicalPath

    If Not modOperatorReadModel.RefreshInventoryReadModelFromSharePointForWorkbook(wbOps, "WH68ST", report) Then GoTo CleanExit

    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 21 _
       And CDbl(GetTableValue(loInv, 1, "QtyAvailable")) = 19 _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "SHAREPOINT", vbTextCompare) = 0 _
       And CBool(GetTableValue(loInv, 1, "IsStale")) = True _
       And InStr(1, CStr(GetTableValue(loInv, 1, "SnapshotId")), "WH68ST.stale.invSys.Snapshot.Inventory.xlsb|", vbTextCompare) = 1 Then
        TestRefreshInventoryReadModelFromSharePoint_StaleSnapshotMarksReadModelStale = 1
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

Public Function TestRefreshInventoryReadModelFromCache_PreservesLocalStagingAndLogs() As Long
    Dim rootPath As String
    Dim wbOps As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model_cached")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH69C", "S9") Then GoTo CleanExit
    SetConfigWarehouseValue "WH69C.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    Set loRecv = wbOps.Worksheets("ReceivedTally").ListObjects("ReceivedTally")
    Set loLog = wbOps.Worksheets("ReceivedLog").ListObjects("ReceivedLog")
    AddInvSysSeedRow loInv, 952, "SKU-RM-CACHED", "Cached Item", "EA", "C1", 15
    SetTableCell loInv, 1, "SnapshotId", "WH69C.invSys.Snapshot.Inventory.xlsb|20260330070000"
    AddReceivedTallyRow loRecv, "REF-CACHED-001", "Cached Item", 5, 952
    AddReceivedLogRow loLog, "WH69C.invSys.Snapshot.Inventory.xlsb|20260330070000", "REF-CACHED-001", "Cached Item", 5, "EA", "Vendor", "C1", "SKU-RM-CACHED", 952

    If Not modOperatorReadModel.RefreshInventoryReadModelFromCacheForWorkbook(wbOps, "WH69C", report) Then GoTo CleanExit

    If CBool(GetTableValue(loInv, 1, "IsStale")) = True _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "CACHED", vbTextCompare) = 0 _
       And CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 15 _
       And StrComp(CStr(GetTableValue(loInv, 1, "SnapshotId")), "WH69C.invSys.Snapshot.Inventory.xlsb|20260330070000", vbTextCompare) = 0 _
       And loRecv.ListRows.Count = 1 _
       And loLog.ListRows.Count = 1 _
       And StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), "REF-CACHED-001", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loLog, 1, "REF_NUMBER")), "REF-CACHED-001", vbTextCompare) = 0 Then
        TestRefreshInventoryReadModelFromCache_PreservesLocalStagingAndLogs = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRefreshInventoryReadModelFromSnapshot_AddsRowsWhenInvSysStartsEmpty() As Long
    Dim rootPath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model_empty")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH68C", "S8") Then GoTo CleanExit
    SetConfigWarehouseValue "WH68C.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    If Not loInv.DataBodyRange Is Nothing Then GoTo CleanExit

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH68C", "SKU-RM-EMPTY", 11, CDate("2026-03-24 18:15:00"))
    If wbSnap Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH68C", "LOCAL", report) Then GoTo CleanExit

    If loInv.ListRows.Count = 1 _
       And StrComp(CStr(GetTableValue(loInv, 1, "ITEM_CODE")), "SKU-RM-EMPTY", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "ITEM")), "SKU-RM-EMPTY", vbTextCompare) = 0 _
       And CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 11 _
       And CDbl(GetTableValue(loInv, 1, "QtyAvailable")) = 11 _
       And CBool(GetTableValue(loInv, 1, "IsStale")) = False _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "LOCAL", vbTextCompare) = 0 Then
        TestRefreshInventoryReadModelFromSnapshot_AddsRowsWhenInvSysStartsEmpty = 1
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

Public Function TestRefreshInventoryReadModelFromSnapshot_AppliesCatalogMetadataForZeroQtyRows() As Long
    Dim rootPath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model_catalog")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH68D", "S8") Then GoTo CleanExit
    SetConfigWarehouseValue "WH68D.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    If Not loInv.DataBodyRange Is Nothing Then GoTo CleanExit

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH68D", "SKU-RM-CAT", 0, CDate("2026-03-24 18:45:00"), _
                                        0, "", "Catalog Item", "CS", "R9", "Catalog Desc", "Vendor C", "VC-9", "raw")
    If wbSnap Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH68D", "LOCAL", report) Then GoTo CleanExit

    If loInv.ListRows.Count = 1 _
       And StrComp(CStr(GetTableValue(loInv, 1, "ITEM_CODE")), "SKU-RM-CAT", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "ITEM")), "Catalog Item", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "UOM")), "CS", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "R9", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "DESCRIPTION")), "Catalog Desc", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "VENDOR(s)")), "Vendor C", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "VENDOR_CODE")), "VC-9", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "CATEGORY")), "raw", vbTextCompare) = 0 _
       And CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 0 _
       And CDbl(GetTableValue(loInv, 1, "QtyAvailable")) = 0 Then
        TestRefreshInventoryReadModelFromSnapshot_AppliesCatalogMetadataForZeroQtyRows = 1
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

Public Function TestRefreshInventoryReadModelFromSnapshot_NormalizesLegacyLocationSummary() As Long
    Dim rootPath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model_legacy_summary")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH68B", "S8") Then GoTo CleanExit
    SetConfigWarehouseValue "WH68B.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    AddInvSysSeedRow loInv, 903, "SKU-RM-LEGACY", "Legacy Summary Item", "EA", "CLEARVIEW=50", 0

    Set wbSnap = CreateSnapshotWorkbook( _
        rootPath, _
        "WH68B", _
        "SKU-RM-LEGACY", _
        200, _
        CDate("2026-03-24 22:50:10"), _
        200, _
        "CLEARVIEW=50; CLEARVIEW=50=50; (blank)=100")
    If wbSnap Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH68B", "LOCAL", report) Then GoTo CleanExit

    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 200 _
       And CDbl(GetTableValue(loInv, 1, "QtyAvailable")) = 200 _
       And StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "CLEARVIEW", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loInv, 1, "LocationSummary")), "CLEARVIEW=100; (blank)=100", vbTextCompare) = 0 _
       And CBool(GetTableValue(loInv, 1, "IsStale")) = False _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "LOCAL", vbTextCompare) = 0 Then
        TestRefreshInventoryReadModelFromSnapshot_NormalizesLegacyLocationSummary = 1
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

Public Function TestRefreshInventoryReadModel_MissingSharePointSnapshotMarksCachedWithoutMutatingLocalTables() As Long
    Dim rootPath As String
    Dim shareRoot As String
    Dim wbOps As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_read_model_missing_sharepoint")
    shareRoot = rootPath & "\Share"

    On Error GoTo CleanFail
    MkDir shareRoot
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH69SP", "S9") Then GoTo CleanExit
    SetConfigWarehouseValue "WH69SP.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    SetConfigWarehouseValue "WH69SP.invSys.Config.xlsb", "PathSharePointRoot", shareRoot
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = wbOps.Worksheets("InventoryManagement").ListObjects("invSys")
    Set loRecv = wbOps.Worksheets("ReceivedTally").ListObjects("ReceivedTally")
    Set loLog = wbOps.Worksheets("ReceivedLog").ListObjects("ReceivedLog")
    AddInvSysSeedRow loInv, 953, "SKU-RM-SP-MISS", "Missing Share Item", "EA", "D1", 17
    SetTableCell loInv, 1, "SnapshotId", "WH69SP.invSys.Snapshot.Inventory.xlsb|20260330070500"
    AddReceivedTallyRow loRecv, "REF-SP-MISS-001", "Missing Share Item", 6, 953
    AddReceivedLogRow loLog, "WH69SP.invSys.Snapshot.Inventory.xlsb|20260330070500", "REF-SP-MISS-001", "Missing Share Item", 6, "EA", "Vendor", "D1", "SKU-RM-SP-MISS", 953

    If Not modOperatorReadModel.RefreshInventoryReadModelFromSharePointForWorkbook(wbOps, "WH69SP", report) Then GoTo CleanExit

    If CBool(GetTableValue(loInv, 1, "IsStale")) = True _
       And StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "CACHED", vbTextCompare) = 0 _
       And CDbl(GetTableValue(loInv, 1, "TOTAL INV")) = 17 _
       And StrComp(CStr(GetTableValue(loInv, 1, "SnapshotId")), "WH69SP.invSys.Snapshot.Inventory.xlsb|20260330070500", vbTextCompare) = 0 _
       And loRecv.ListRows.Count = 1 _
       And loLog.ListRows.Count = 1 _
       And StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), "REF-SP-MISS-001", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loLog, 1, "REF_NUMBER")), "REF-SP-MISS-001", vbTextCompare) = 0 Then
        TestRefreshInventoryReadModel_MissingSharePointSnapshotMarksCachedWithoutMutatingLocalTables = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestSavedReceivingWorkbook_StaleSharePointSnapshotShowsVisibleMetadataWithoutMutatingLocalTables() As Long
    Dim rootPath As String
    Dim shareRoot As String
    Dim snapshotRoot As String
    Dim canonicalPath As String
    Dim stalePath As String
    Dim operatorPath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject
    Dim invRow As Long

    rootPath = BuildRuntimeTestRoot("phase6_saved_receiving_sharepoint_stale")
    shareRoot = rootPath & "\Share"
    snapshotRoot = shareRoot & "\Snapshots"
    canonicalPath = snapshotRoot & "\WH70SP.invSys.Snapshot.Inventory.xlsb"
    stalePath = snapshotRoot & "\WH70SP.stale.invSys.Snapshot.Inventory.xlsb"

    On Error GoTo CleanFail
    MkDir shareRoot
    MkDir snapshotRoot
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH70SP", "S10") Then GoTo CleanExit
    SetConfigWarehouseValue "WH70SP.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    SetConfigWarehouseValue "WH70SP.invSys.Config.xlsb", "PathSharePointRoot", shareRoot
    If Not modConfig.Reload() Then GoTo CleanExit

    operatorPath = rootPath & "\WH70SP_S10_Receiving_Operator.xlsb"
    BuildSavedReceivingOperatorWorkbookForTest operatorPath, "SKU-RM-SP-ST-OP", "REF-SP-ST-001", "SNAP-SP-ST-OLD", 4, "OLD"
    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then
        failureReason = "Saved receiving operator workbook did not reopen."
        GoTo CleanExit
    End If
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set wbSnap = CreateSnapshotWorkbook(snapshotRoot, "WH70SP", "SKU-RM-SP-ST-OP", 21, CDate("2026-03-30 08:35:00"), _
                                        19, "SP2=19", "Saved Stale Share Item", "EA", "SP2")
    If wbSnap Is Nothing Then GoTo CleanExit
    wbSnap.SaveCopyAs stalePath
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing
    Kill canonicalPath

    If Not modOperatorReadModel.RefreshInventoryReadModelFromSharePointForWorkbook(wbOps, "WH70SP", report) Then
        failureReason = "RefreshInventoryReadModelFromSharePointForWorkbook failed: " & report
        GoTo CleanExit
    End If

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    Set loLog = FindTableByName(wbOps, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then
        failureReason = "Saved receiving workbook tables were missing after stale SharePoint refresh."
        GoTo CleanExit
    End If

    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-RM-SP-ST-OP")
    If invRow = 0 Then
        failureReason = "Saved receiving workbook did not expose the refreshed SharePoint SKU."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "TOTAL INV")) <> 21 Then
        failureReason = "Saved receiving workbook TOTAL INV did not reflect the stale SharePoint snapshot."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "QtyAvailable")) <> 19 Then
        failureReason = "Saved receiving workbook QtyAvailable did not reflect the stale SharePoint snapshot."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, invRow, "SourceType")), "SHAREPOINT", vbTextCompare) <> 0 Then
        failureReason = "Saved receiving workbook SourceType was not SHAREPOINT for the stale snapshot case."
        GoTo CleanExit
    End If
    If CBool(GetTableValue(loInv, invRow, "IsStale")) <> True Then
        failureReason = "Saved receiving workbook did not remain visibly stale for the stale SharePoint snapshot case."
        GoTo CleanExit
    End If
    If InStr(1, CStr(GetTableValue(loInv, invRow, "SnapshotId")), "WH70SP.stale.invSys.Snapshot.Inventory.xlsb|", vbTextCompare) <> 1 Then
        failureReason = "Saved receiving workbook SnapshotId did not show the stale SharePoint artifact."
        GoTo CleanExit
    End If
    If Not IsDate(GetTableValue(loInv, invRow, "LastRefreshUTC")) Then
        failureReason = "Saved receiving workbook LastRefreshUTC was not populated."
        GoTo CleanExit
    End If
    If loRecv.ListRows.Count <> 1 Then
        failureReason = "ReceivedTally changed during stale SharePoint refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), "REF-SP-ST-001", vbTextCompare) <> 0 Then
        failureReason = "ReceivedTally REF_NUMBER changed during stale SharePoint refresh."
        GoTo CleanExit
    End If
    If loLog.ListRows.Count <> 1 Then
        failureReason = "ReceivedLog changed during stale SharePoint refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loLog, 1, "REF_NUMBER")), "REF-SP-ST-001", vbTextCompare) <> 0 Then
        failureReason = "ReceivedLog REF_NUMBER changed during stale SharePoint refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loLog, 1, "SNAPSHOT_ID")), "SNAP-SP-ST-OLD", vbTextCompare) <> 0 Then
        failureReason = "ReceivedLog SNAPSHOT_ID changed during stale SharePoint refresh."
        GoTo CleanExit
    End If

    TestSavedReceivingWorkbook_StaleSharePointSnapshotShowsVisibleMetadataWithoutMutatingLocalTables = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7113, "TestSavedReceivingWorkbook_StaleSharePointSnapshotShowsVisibleMetadataWithoutMutatingLocalTables", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestSavedReceivingWorkbook_MissingSnapshotDoesNotBlockQueueAndRefresh() As Long
    Dim rootPath As String
    Dim operatorPath As String
    Dim currentUser As String
    Dim wbOps As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim report As String
    Dim failureReason As String
    Dim eventIdOut As String
    Dim processedCount As Long
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loInventoryLog As ListObject
    Dim invRow As Long
    Dim logRow As Long

    rootPath = BuildRuntimeTestRoot("phase6_missing_snapshot_queue")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH70", "S10") Then GoTo CleanExit
    SetConfigWarehouseValue "WH70.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH70") Then GoTo CleanExit

    currentUser = ResolveCurrentTestUserId()
    EnsureAuthCapabilityForTest "WH70", currentUser, "RECEIVE_POST", "WH70", "*"
    EnsureAuthCapabilityForTest "WH70", "svc_processor", "INBOX_PROCESS", "WH70", "*"

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH70", Array("SKU-RM-QUEUE"))
    Set wbInbox = CreateCanonicalReceiveInboxWorkbookForTest(rootPath, "S10")
    If wbInv Is Nothing Or wbInbox Is Nothing Then
        failureReason = "Canonical runtime workbooks for stale-queue test were not created."
        GoTo CleanExit
    End If

    operatorPath = rootPath & "\WH70_S10_Receiving_Operator.xlsb"
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    If loInv Is Nothing Or loRecv Is Nothing Then
        failureReason = "Saved receiving workbook surface was incomplete."
        GoTo CleanExit
    End If

    AddInvSysSeedRow loInv, 910, "SKU-RM-QUEUE", "Stale Queue Item", "EA", "B1", 12
    AddReceivedTallyRow loRecv, "REF-ST-QUEUE-001", "Stale Queue Item", 3, 910
    wbOps.SaveAs Filename:=operatorPath, FileFormat:=50
    wbOps.Close SaveChanges:=False
    Set wbOps = Nothing

    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH70", "LOCAL", report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    If loInv Is Nothing Or loRecv Is Nothing Then
        failureReason = "Saved workbook tables were missing after stale refresh."
        GoTo CleanExit
    End If
    If CBool(GetTableValue(loInv, 1, "IsStale")) <> True Then
        failureReason = "invSys was not marked stale when the snapshot was missing."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "CACHED", vbTextCompare) <> 0 Then
        failureReason = "invSys SourceType was not CACHED for the missing snapshot case."
        GoTo CleanExit
    End If
    If loRecv.ListRows.Count <> 1 Then
        failureReason = "ReceivedTally changed during stale refresh."
        GoTo CleanExit
    End If

    If Not modRoleEventWriter.QueueReceiveEvent("WH70", "S10", currentUser, "SKU-RM-QUEUE", 4, "A1", "stale-queue", "", "", Now, wbInbox, eventIdOut, report) Then
        failureReason = "QueueReceiveEvent failed while invSys was stale: " & report
        GoTo CleanExit
    End If
    If Trim$(eventIdOut) = "" Then
        failureReason = "QueueReceiveEvent did not return an EventID."
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, eventIdOut, "NEW") Then
        failureReason = "Queued inbox row was not NEW after stale workbook posting."
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch("WH70", 500, report)
    If processedCount <> 1 Then
        failureReason = "RunBatch did not process the stale-workbook receive event. " & report & _
                        "; Inbox=" & DescribeInboxRowStateForTest(wbInbox, eventIdOut)
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, eventIdOut, "PROCESSED") Then
        failureReason = "Processed inbox row was not marked PROCESSED."
        GoTo CleanExit
    End If

    Set loInventoryLog = FindTableByName(wbInv, "tblInventoryLog")
    If loInventoryLog Is Nothing Then
        failureReason = "Canonical inventory log was missing after RunBatch."
        GoTo CleanExit
    End If
    logRow = FindRowByColumnValueInTable(loInventoryLog, "EventID", eventIdOut)
    If logRow = 0 Then
        failureReason = "Canonical inventory log did not record the stale-workbook event."
        GoTo CleanExit
    End If

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH70", "LOCAL", report) Then
        failureReason = "RefreshInventoryReadModelForWorkbook failed after processor catch-up: " & report
        GoTo CleanExit
    End If

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    If loInv Is Nothing Or loRecv Is Nothing Then
        failureReason = "Saved workbook tables were missing after processor catch-up refresh."
        GoTo CleanExit
    End If
    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-RM-QUEUE")
    If invRow = 0 Then
        failureReason = "invSys did not refresh the queued SKU after processor catch-up."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "TOTAL INV")) <> 4 Then
        failureReason = "invSys TOTAL INV did not reflect the processed stale-workbook receive event."
        GoTo CleanExit
    End If
    If CBool(GetTableValue(loInv, invRow, "IsStale")) <> False Then
        failureReason = "invSys remained stale after processor catch-up refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, invRow, "SourceType")), "LOCAL", vbTextCompare) <> 0 Then
        failureReason = "invSys SourceType was not LOCAL after processor catch-up refresh."
        GoTo CleanExit
    End If
    If loRecv.ListRows.Count <> 1 Then
        failureReason = "ReceivedTally changed after stale-workbook queue/process/refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), "REF-ST-QUEUE-001", vbTextCompare) <> 0 Then
        failureReason = "ReceivedTally REF_NUMBER was not preserved across stale-workbook processing."
        GoTo CleanExit
    End If

    TestSavedReceivingWorkbook_MissingSnapshotDoesNotBlockQueueAndRefresh = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7106, "TestSavedReceivingWorkbook_MissingSnapshotDoesNotBlockQueueAndRefresh", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestSavedReceivingWorkbook_FullRuntimeCloseReopenReloadsCanonicalWorkbooks() As Long
    Dim rootPath As String
    Dim operatorPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim processedCount As Long
    Dim eventIdOut As String
    Dim wbOps As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject
    Dim loInventoryLog As ListObject
    Dim invRow As Long
    Dim logRow As Long

    rootPath = BuildRuntimeTestRoot("phase6_full_reopen_runtime")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH78", "S18") Then GoTo CleanExit
    SetConfigWarehouseValue "WH78.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH78") Then GoTo CleanExit

    currentUser = ResolveCurrentTestUserId()
    EnsureAuthCapabilityForTest "WH78", currentUser, "RECEIVE_POST", "WH78", "*"
    EnsureAuthCapabilityForTest "WH78", "svc_processor", "INBOX_PROCESS", "WH78", "*"

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH78", Array("SKU-RM-RESTART"))
    Set wbInbox = CreateCanonicalReceiveInboxWorkbookForTest(rootPath, "S18")
    If wbInv Is Nothing Or wbInbox Is Nothing Then
        failureReason = "Canonical inventory/inbox workbooks could not be created for full reopen test."
        GoTo CleanExit
    End If

    AddInboxReceiveEventRowForTest FindTableByName(wbInbox, "tblInboxReceive"), "EVT-RESTART-001", "WH78", "S18", currentUser, "SKU-RM-RESTART", 9, "A1", "restart-seed"
    wbInbox.Save
    processedCount = modProcessor.RunBatch("WH78", 500, report)
    If processedCount <> 1 Then
        failureReason = "Initial RunBatch did not seed the canonical runtime state. " & report & _
                        "; Inbox=" & DescribeInboxRowStateForTest(wbInbox, "EVT-RESTART-001")
        GoTo CleanExit
    End If

    operatorPath = rootPath & "\WH78_S18_Receiving_Operator.xlsb"
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    Set loLog = FindTableByName(wbOps, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then
        failureReason = "Saved receiving workbook surface was incomplete before restart simulation."
        GoTo CleanExit
    End If

    AddInvSysSeedRow loInv, 911, "SKU-RM-RESTART", "Restart Item", "EA", "Z9", 1
    AddReceivedTallyRow loRecv, "REF-RESTART-001", "Restart Item", 2, 911
    AddReceivedLogRow loLog, "SNAP-RESTART-OLD", "REF-RESTART-001", "Restart Item", 2, "EA", "Vendor R", "Z9", "SKU-RM-RESTART", 911
    wbOps.SaveAs Filename:=operatorPath, FileFormat:=50
    wbOps.Close SaveChanges:=False
    Set wbOps = Nothing

    CloseWorkbookByNameIfOpen "WH78.invSys.Config.xlsb"
    CloseWorkbookByNameIfOpen "WH78.invSys.Auth.xlsb"
    CloseWorkbookByNameIfOpen "WH78.invSys.Data.Inventory.xlsb"
    CloseWorkbookByNameIfOpen "WH78.invSys.Snapshot.Inventory.xlsb"
    CloseWorkbookByNameIfOpen "WH78.Outbox.Events.xlsb"
    CloseWorkbookByNameIfOpen "invSys.Inbox.Receiving.S18.xlsb"
    modRuntimeWorkbooks.ClearCoreDataRootOverride

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH78", "S18") Then
        failureReason = "LoadConfig failed after full runtime close/reopen boundary."
        GoTo CleanExit
    End If
    If Not modAuth.LoadAuth("WH78") Then
        failureReason = "LoadAuth failed after full runtime close/reopen boundary."
        GoTo CleanExit
    End If

    Set wbCfg = FindWorkbookByName("WH78.invSys.Config.xlsb")
    Set wbAuth = FindWorkbookByName("WH78.invSys.Auth.xlsb")
    If StrComp(modConfig.GetResolvedWorkbookName(), "WH78.invSys.Config.xlsb", vbTextCompare) <> 0 Then
        failureReason = "LoadConfig did not resolve the canonical config workbook after runtime reload."
        GoTo CleanExit
    End If
    If StrComp(modAuth.GetResolvedAuthWorkbookName(), "WH78.invSys.Auth.xlsb", vbTextCompare) <> 0 Then
        failureReason = "LoadAuth did not resolve the canonical auth workbook after runtime reload."
        GoTo CleanExit
    End If
    If Not wbCfg Is Nothing Then
        If StrComp(wbCfg.FullName, rootPath & "\WH78.invSys.Config.xlsb", vbTextCompare) <> 0 Then
            failureReason = "Config workbook reopened at an unexpected path."
            GoTo CleanExit
        End If
    End If
    If Not wbAuth Is Nothing Then
        If StrComp(wbAuth.FullName, rootPath & "\WH78.invSys.Auth.xlsb", vbTextCompare) <> 0 Then
            failureReason = "Auth workbook reopened at an unexpected path."
            GoTo CleanExit
        End If
    End If
    If StrComp(NormalizeTestPath(rootPath), NormalizeTestPath(modConfig.GetString("PathDataRoot", "")), vbTextCompare) <> 0 Then
        failureReason = "PathDataRoot did not reload to the canonical runtime root. Actual=" & modConfig.GetString("PathDataRoot", "")
        GoTo CleanExit
    End If

    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then
        failureReason = "Saved receiving workbook could not be reopened after runtime reload."
        GoTo CleanExit
    End If
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH78", "LOCAL", report) Then
        failureReason = "RefreshInventoryReadModelForWorkbook failed after runtime reload: " & report
        GoTo CleanExit
    End If

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    Set loLog = FindTableByName(wbOps, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then
        failureReason = "Saved receiving workbook surfaces were missing after runtime reload."
        GoTo CleanExit
    End If
    If StrComp(wbOps.FullName, operatorPath, vbTextCompare) <> 0 Then
        failureReason = "Saved receiving workbook reopened at an unexpected path after runtime reload."
        GoTo CleanExit
    End If
    If loRecv.ListRows.Count <> 1 Or loLog.ListRows.Count <> 1 Then
        failureReason = "Workbook-local receiving tables changed across full runtime close/reopen."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), "REF-RESTART-001", vbTextCompare) <> 0 Then
        failureReason = "ReceivedTally REF_NUMBER was not preserved across full runtime close/reopen."
        GoTo CleanExit
    End If
    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-RM-RESTART")
    If invRow = 0 Then
        failureReason = "invSys did not refresh the canonical SKU after runtime reload."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "TOTAL INV")) <> 9 Then
        failureReason = "invSys TOTAL INV did not reload from the canonical snapshot after runtime reload."
        GoTo CleanExit
    End If
    If CBool(GetTableValue(loInv, invRow, "IsStale")) <> False Then
        failureReason = "invSys was stale after runtime reload despite a canonical snapshot."
        GoTo CleanExit
    End If
    If InStr(1, CStr(GetTableValue(loInv, invRow, "SnapshotId")), "WH78.invSys.Snapshot.Inventory.xlsb|", vbTextCompare) <> 1 Then
        failureReason = "invSys SnapshotId was not refreshed after runtime reload."
        GoTo CleanExit
    End If

    Set wbInbox = Application.Workbooks.Open(rootPath & "\invSys.Inbox.Receiving.S18.xlsb")
    If wbInbox Is Nothing Then
        failureReason = "Receive inbox workbook could not be explicitly reopened after runtime reload."
        GoTo CleanExit
    End If
    If StrComp(wbInbox.FullName, rootPath & "\invSys.Inbox.Receiving.S18.xlsb", vbTextCompare) <> 0 Then
        failureReason = "Receive inbox workbook reopened at an unexpected path after runtime reload."
        GoTo CleanExit
    End If

    If Not modRoleEventWriter.QueueReceiveEvent("WH78", "S18", currentUser, "SKU-RM-RESTART", 4, "A1", "restart-post", "", "", Now, wbInbox, eventIdOut, report) Then
        failureReason = "QueueReceiveEvent failed after runtime reload: " & report
        GoTo CleanExit
    End If
    If Trim$(eventIdOut) = "" Then
        failureReason = "QueueReceiveEvent did not return an EventID after runtime reload."
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch("WH78", 500, report)
    If processedCount <> 1 Then
        failureReason = "RunBatch did not process the post-restart receive event. " & report & _
                        "; Inbox=" & DescribeInboxRowStateForTest(wbInbox, eventIdOut)
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, eventIdOut, "PROCESSED") Then
        failureReason = "Post-restart receive inbox row was not marked PROCESSED."
        GoTo CleanExit
    End If

    Set wbInv = FindWorkbookByName("WH78.invSys.Data.Inventory.xlsb")
    If wbInv Is Nothing Then
        If Len(Dir$(rootPath & "\WH78.invSys.Data.Inventory.xlsb")) = 0 Then
            failureReason = "Canonical inventory workbook file was not present after post-restart RunBatch."
            GoTo CleanExit
        End If
        Set wbInv = Application.Workbooks.Open(rootPath & "\WH78.invSys.Data.Inventory.xlsb")
        If wbInv Is Nothing Then
            failureReason = "Canonical inventory workbook could not be reopened for verification after post-restart RunBatch."
            GoTo CleanExit
        End If
    End If
    Set loInventoryLog = FindTableByName(wbInv, "tblInventoryLog")
    If loInventoryLog Is Nothing Then
        failureReason = "Canonical inventory log was missing after post-restart RunBatch."
        GoTo CleanExit
    End If
    logRow = FindRowByColumnValueInTable(loInventoryLog, "EventID", eventIdOut)
    If logRow = 0 Then
        failureReason = "Canonical inventory log did not record the post-restart receive event."
        GoTo CleanExit
    End If

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH78", "LOCAL", report) Then
        failureReason = "RefreshInventoryReadModelForWorkbook failed after post-restart RunBatch: " & report
        GoTo CleanExit
    End If
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    If loInv Is Nothing Or loRecv Is Nothing Then
        failureReason = "Saved receiving workbook surfaces were missing after post-restart refresh."
        GoTo CleanExit
    End If
    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-RM-RESTART")
    If invRow = 0 Then
        failureReason = "invSys lost the canonical SKU after post-restart refresh."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "TOTAL INV")) <> 13 Then
        failureReason = "invSys TOTAL INV did not include the post-restart receive event."
        GoTo CleanExit
    End If
    If loRecv.ListRows.Count <> 1 Then
        failureReason = "ReceivedTally changed after post-restart refresh."
        GoTo CleanExit
    End If

    TestSavedReceivingWorkbook_FullRuntimeCloseReopenReloadsCanonicalWorkbooks = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    CloseWorkbookIfOpen wbAuth
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookByNameIfOpen "WH78.invSys.Snapshot.Inventory.xlsb"
    CloseWorkbookByNameIfOpen "WH78.Outbox.Events.xlsb"
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7107, "TestSavedReceivingWorkbook_FullRuntimeCloseReopenReloadsCanonicalWorkbooks", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestSavedReceivingWorkbook_ReopenRefreshPreservesLocalTables() As Long
    Dim rootPath As String
    Dim operatorPath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_saved_operator_reopen")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH71", "S11") Then GoTo CleanExit
    SetConfigWarehouseValue "WH71.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    operatorPath = rootPath & "\WH71_S11_Receiving_Operator.xlsb"
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    Set loLog = FindTableByName(wbOps, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then
        failureReason = "Initial saved operator workbook surface did not resolve expected tables."
        GoTo CleanExit
    End If

    AddInvSysSeedRow loInv, 904, "SKU-RM-REOPEN", "Saved Workbook Item", "EA", "B2", 1
    AddReceivedTallyRow loRecv, "REF-REOPEN-001", "Saved Workbook Item", 3, 904
    AddReceivedLogRow loLog, "SNAP-OLD-001", "REF-REOPEN-001", "Saved Workbook Item", 3, "EA", "Vendor A", "B2", "SKU-RM-REOPEN", 904

    wbOps.SaveAs Filename:=operatorPath, FileFormat:=50
    wbOps.Close SaveChanges:=False
    Set wbOps = Nothing

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH71", "SKU-RM-REOPEN", 12, CDate("2026-03-25 09:45:00"))
    If wbSnap Is Nothing Then GoTo CleanExit
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH71", "LOCAL", report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    Set loLog = FindTableByName(wbOps, "ReceivedLog")
    If loInv Is Nothing Then
        failureReason = "invSys table was missing after reopen/refresh."
        GoTo CleanExit
    End If
    If loRecv Is Nothing Then
        failureReason = "ReceivedTally table was missing after reopen/refresh."
        GoTo CleanExit
    End If
    If loLog Is Nothing Then
        failureReason = "ReceivedLog table was missing after reopen/refresh."
        GoTo CleanExit
    End If

    If StrComp(wbOps.FullName, operatorPath, vbTextCompare) <> 0 Then
        failureReason = "Operator workbook reopened at unexpected path."
        GoTo CleanExit
    End If
    If StrComp(wbOps.Name, "WH71_S11_Receiving_Operator.xlsb", vbTextCompare) <> 0 Then
        failureReason = "Operator workbook reopened with unexpected name."
        GoTo CleanExit
    End If
    If StrComp(wbOps.Name, "WH71.invSys.Config.xlsb", vbTextCompare) = 0 Then
        failureReason = "Operator workbook identity drifted to runtime config workbook."
        GoTo CleanExit
    End If

    If loRecv.ListRows.Count <> 1 Then
        failureReason = "ReceivedTally row count changed across reopen/refresh."
        GoTo CleanExit
    End If
    If loLog.ListRows.Count <> 1 Then
        failureReason = "ReceivedLog row count changed across reopen/refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), "REF-REOPEN-001", vbTextCompare) <> 0 Then
        failureReason = "ReceivedTally REF_NUMBER was not preserved."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loRecv, 1, "QUANTITY")) <> 3 Then
        failureReason = "ReceivedTally QUANTITY was not preserved."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loLog, 1, "SNAPSHOT_ID")), "SNAP-OLD-001", vbTextCompare) <> 0 Then
        failureReason = "ReceivedLog SNAPSHOT_ID was not preserved."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loLog, 1, "REF_NUMBER")), "REF-REOPEN-001", vbTextCompare) <> 0 Then
        failureReason = "ReceivedLog REF_NUMBER was not preserved."
        GoTo CleanExit
    End If

    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 12 Then
        failureReason = "invSys TOTAL INV did not refresh from snapshot."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "QtyAvailable")) <> 12 Then
        failureReason = "invSys QtyAvailable did not refresh from snapshot."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "A1", vbTextCompare) <> 0 Then
        failureReason = "invSys LOCATION did not refresh to primary snapshot location."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, 1, "ITEM_CODE")), "SKU-RM-REOPEN", vbTextCompare) <> 0 Then
        failureReason = "invSys ITEM_CODE drifted across reopen/refresh."
        GoTo CleanExit
    End If
    If InStr(1, CStr(GetTableValue(loInv, 1, "SnapshotId")), "WH71.invSys.Snapshot.Inventory.xlsb|", vbTextCompare) <> 1 Then
        failureReason = "invSys SnapshotId was not refreshed."
        GoTo CleanExit
    End If
    If CBool(GetTableValue(loInv, 1, "IsStale")) <> False Then
        failureReason = "invSys was marked stale after successful refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "LOCAL", vbTextCompare) <> 0 Then
        failureReason = "invSys SourceType was not LOCAL after refresh."
        GoTo CleanExit
    End If
    If Not IsDate(GetTableValue(loInv, 1, "LastRefreshUTC")) Then
        failureReason = "invSys LastRefreshUTC was not populated."
        GoTo CleanExit
    End If
    If Not IsDate(GetTableValue(loInv, 1, "LAST EDITED")) Then
        failureReason = "invSys LAST EDITED was not populated."
        GoTo CleanExit
    End If

    TestSavedReceivingWorkbook_ReopenRefreshPreservesLocalTables = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7101, "TestSavedReceivingWorkbook_ReopenRefreshPreservesLocalTables", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestReceivingSetupUi_ForceRefreshesRegisteredWorkbook() As Long
    Dim rootPath As String
    Dim operatorPath As String
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject
    Dim invRow As Long

    rootPath = BuildRuntimeTestRoot("phase6_receiving_setup_refresh")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH82", "S23") Then GoTo CleanExit
    SetConfigWarehouseValue "WH82.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    operatorPath = rootPath & "\WH82_S23_Receiving_Operator.xlsb"
    BuildSavedReceivingOperatorWorkbookForTest operatorPath, "SKU-SETUP-001", "REF-SETUP-001", "SNAP-SETUP-OLD", 0, "Z9"

    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then GoTo CleanExit
    wbOps.Activate
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then
        failureReason = "EnsureReceivingWorkbookSurface failed: " & report
        GoTo CleanExit
    End If
    modOperatorReadModel.InitializeAutoSnapshotForWorkbook wbOps

    Set loInv = FindTableByName(wbOps, "invSys")
    If loInv Is Nothing Then
        failureReason = "invSys table missing after receiving setup initialization."
        GoTo CleanExit
    End If
    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-SETUP-001")
    If invRow = 0 Then
        failureReason = "Seed invSys row missing before forced setup refresh."
        GoTo CleanExit
    End If
    If CBool(GetTableValue(loInv, invRow, "IsStale")) <> True Then
        failureReason = "Receiving setup initialization did not mark the missing snapshot as stale."
        GoTo CleanExit
    End If

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH82", "SKU-SETUP-001", 14, CDate("2026-03-28 11:45:00"), _
                                        14, "B7=14", "Setup Refresh Item", "EA", "B7", "Setup refresh desc", "Vendor Setup", "VS-1", "receiving")
    If wbSnap Is Nothing Then
        failureReason = "Snapshot workbook could not be created for setup refresh test."
        GoTo CleanExit
    End If
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then
        failureReason = "EnsureReceivingWorkbookSurface failed during forced refresh: " & report
        GoTo CleanExit
    End If
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "", "LOCAL", report) Then
        failureReason = "RefreshInventoryReadModelForWorkbook failed: " & report
        GoTo CleanExit
    End If

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    Set loLog = FindTableByName(wbOps, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then
        failureReason = "Receiving tables were missing after forced setup refresh."
        GoTo CleanExit
    End If

    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-SETUP-001")
    If invRow = 0 Then
        failureReason = "Forced setup refresh did not retain the target SKU."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "TOTAL INV")) <> 14 Then
        failureReason = "Forced setup refresh did not update TOTAL INV from the shared snapshot."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, invRow, "LOCATION")), "B7", vbTextCompare) <> 0 Then
        failureReason = "Forced setup refresh did not update LOCATION from the shared snapshot."
        GoTo CleanExit
    End If
    If CBool(GetTableValue(loInv, invRow, "IsStale")) <> False Then
        failureReason = "Forced setup refresh left invSys marked stale."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, invRow, "SourceType")), "LOCAL", vbTextCompare) <> 0 Then
        failureReason = "Forced setup refresh did not preserve LOCAL source type."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), "REF-SETUP-001", vbTextCompare) <> 0 Then
        failureReason = "Receiving staging row was not preserved across forced setup refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loLog, 1, "SNAPSHOT_ID")), "SNAP-SETUP-OLD", vbTextCompare) <> 0 Then
        failureReason = "Receiving log row was not preserved across forced setup refresh."
        GoTo CleanExit
    End If

    TestReceivingSetupUi_ForceRefreshesRegisteredWorkbook = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7111, "TestReceivingSetupUi_ForceRefreshesRegisteredWorkbook", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestInventoryPublisher_PublishesSnapshotForOpenInventoryWorkbook() As Long
    Dim rootPath As String
    Dim report As String
    Dim failureReason As String
    Dim publishCount As Long
    Dim snapshotPath As String
    Dim wbInv As Workbook
    Dim wbRuntime As Workbook
    Dim wbSnap As Workbook
    Dim loRuntimeCatalog As ListObject
    Dim loSnap As ListObject
    Dim rowSku1 As Long
    Dim rowSku2 As Long

    rootPath = BuildRuntimeTestRoot("phase6_inventory_open_publish")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH83", "S24") Then GoTo CleanExit
    SetConfigWarehouseValue "WH83.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbInv = CreateManagedInventoryDonorWorkbookForTest(rootPath, "FRODECO.inventory_management.xlsb")
    If wbInv Is Nothing Then
        failureReason = "Inventory source workbook could not be created."
        GoTo CleanExit
    End If
    AddInvSysSeedRow FindTableByName(wbInv, "invSys"), 1001, "SKU-PUB-001", "Publish Item 1", "EA", "A1", 7
    AddInvSysSeedRow FindTableByName(wbInv, "invSys"), 1002, "SKU-PUB-002", "Publish Item 2", "EA", "B2", 0
    wbInv.Save

    publishCount = modInventoryPublisher.PublishOpenInventorySnapshots(report)
    If publishCount < 1 Then
        failureReason = "PublishOpenInventorySnapshots did not publish the open inventory workbook. " & report
        GoTo CleanExit
    End If

    Set wbRuntime = modInventoryApply.ResolveInventoryWorkbook("WH83")
    If wbRuntime Is Nothing Then
        failureReason = "Canonical runtime inventory workbook was not created."
        GoTo CleanExit
    End If
    Set loRuntimeCatalog = FindTableByName(wbRuntime, "tblSkuCatalog")
    If loRuntimeCatalog Is Nothing Then
        failureReason = "Canonical runtime SKU catalog was not created."
        GoTo CleanExit
    End If
    If FindRowByColumnValueInTable(loRuntimeCatalog, "SKU", "SKU-PUB-001") = 0 Or FindRowByColumnValueInTable(loRuntimeCatalog, "SKU", "SKU-PUB-002") = 0 Then
        failureReason = "Canonical runtime SKU catalog did not receive the donor workbook managed inventory rows."
        GoTo CleanExit
    End If

    snapshotPath = rootPath & "\WH83.invSys.Snapshot.Inventory.xlsb"
    If Len(Dir$(snapshotPath)) = 0 Then
        failureReason = "Snapshot workbook was not published for the open inventory workbook."
        GoTo CleanExit
    End If

    Set wbSnap = Application.Workbooks.Open(snapshotPath)
    Set loSnap = FindTableByName(wbSnap, "tblInventorySnapshot")
    If loSnap Is Nothing Then
        failureReason = "Published snapshot table was missing."
        GoTo CleanExit
    End If
    rowSku1 = FindRowByColumnValueInTable(loSnap, "SKU", "SKU-PUB-001")
    rowSku2 = FindRowByColumnValueInTable(loSnap, "SKU", "SKU-PUB-002")
    If rowSku1 = 0 Or rowSku2 = 0 Then
        failureReason = "Published snapshot did not include the full catalog list from the open inventory workbook."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loSnap, rowSku1, "QtyOnHand")) <> 7 Then
        failureReason = "Published snapshot did not preserve managed inventory quantities from the source workbook."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loSnap, rowSku2, "QtyOnHand")) <> 0 Then
        failureReason = "Published snapshot did not preserve zero quantities for catalog-only rows."
        GoTo CleanExit
    End If

    TestInventoryPublisher_PublishesSnapshotForOpenInventoryWorkbook = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbRuntime
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7112, "TestInventoryPublisher_PublishesSnapshotForOpenInventoryWorkbook", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestLanSharedSnapshot_TwoSavedOperatorWorkbooksRefreshWithoutCrossContamination() As Long
    Dim rootPath As String
    Dim operatorPathA As String
    Dim operatorPathB As String
    Dim wbOpsA As Workbook
    Dim wbOpsB As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_lan_shared_snapshot")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH72", "S11") Then GoTo CleanExit
    SetConfigWarehouseValue "WH72.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    operatorPathA = rootPath & "\WH72_S11_Receiving_Operator.xlsb"
    operatorPathB = rootPath & "\WH72_S12_Receiving_Operator.xlsb"

    Set wbOpsA = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOpsA, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOpsA, "invSys")
    Set loRecv = FindTableByName(wbOpsA, "ReceivedTally")
    Set loLog = FindTableByName(wbOpsA, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then GoTo CleanExit
    AddInvSysSeedRow loInv, 905, "SKU-LAN-001", "LAN Shared Item", "EA", "B2", 2
    AddReceivedTallyRow loRecv, "REF-LAN-A", "LAN Shared Item", 4, 905
    AddReceivedLogRow loLog, "SNAP-LAN-A", "REF-LAN-A", "LAN Shared Item", 4, "EA", "Vendor A", "B2", "SKU-LAN-001", 905
    wbOpsA.SaveAs Filename:=operatorPathA, FileFormat:=50
    wbOpsA.Close SaveChanges:=False
    Set wbOpsA = Nothing

    Set wbOpsB = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOpsB, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOpsB, "invSys")
    Set loRecv = FindTableByName(wbOpsB, "ReceivedTally")
    Set loLog = FindTableByName(wbOpsB, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then GoTo CleanExit
    AddInvSysSeedRow loInv, 906, "SKU-LAN-001", "LAN Shared Item", "EA", "C3", 3
    AddReceivedTallyRow loRecv, "REF-LAN-B", "LAN Shared Item", 5, 906
    AddReceivedLogRow loLog, "SNAP-LAN-B", "REF-LAN-B", "LAN Shared Item", 5, "EA", "Vendor B", "C3", "SKU-LAN-001", 906
    wbOpsB.SaveAs Filename:=operatorPathB, FileFormat:=50
    wbOpsB.Close SaveChanges:=False
    Set wbOpsB = Nothing

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH72", "SKU-LAN-001", 25, CDate("2026-03-25 10:15:00"))
    If wbSnap Is Nothing Then GoTo CleanExit
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    Set wbOpsA = Application.Workbooks.Open(operatorPathA)
    If wbOpsA Is Nothing Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOpsA, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOpsA, "WH72", "LOCAL", report) Then GoTo CleanExit

    Set wbOpsB = Application.Workbooks.Open(operatorPathB)
    If wbOpsB Is Nothing Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOpsB, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOpsB, "WH72", "LOCAL", report) Then GoTo CleanExit

    If Not AssertLanWorkbookState(wbOpsA, operatorPathA, "REF-LAN-A", "SNAP-LAN-A", 25, "SKU-LAN-001", "WH72.invSys.Snapshot.Inventory.xlsb|") Then GoTo CleanExit
    If Not AssertLanWorkbookState(wbOpsB, operatorPathB, "REF-LAN-B", "SNAP-LAN-B", 25, "SKU-LAN-001", "WH72.invSys.Snapshot.Inventory.xlsb|") Then GoTo CleanExit

    TestLanSharedSnapshot_TwoSavedOperatorWorkbooksRefreshWithoutCrossContamination = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOpsA
    CloseWorkbookIfOpen wbOpsB
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLanTwoStationProcessorRun_RespectsLockAndPreservesOperatorWorkbooks() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim wbInv As Workbook
    Dim wbInboxA As Workbook
    Dim wbInboxB As Workbook
    Dim wbOpsA As Workbook
    Dim wbOpsB As Workbook
    Dim loLocks As ListObject
    Dim loSku As ListObject
    Dim loLoc As ListObject
    Dim runIdA As String
    Dim runIdB As String
    Dim message As String
    Dim processedCount As Long
    Dim operatorPathA As String
    Dim operatorPathB As String

    rootPath = BuildRuntimeTestRoot("phase6_lan_processor")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH75", "S11") Then GoTo CleanExit
    SetConfigWarehouseValue "WH75.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH75") Then GoTo CleanExit

    currentUser = ResolveCurrentTestUserId()
    EnsureAuthCapabilityForTest "WH75", currentUser, "RECEIVE_POST", "WH75", "*"
    EnsureAuthCapabilityForTest "WH75", "svc_processor", "INBOX_PROCESS", "WH75", "*"

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH75", Array("SKU-LAN-LOCK"))
    If wbInv Is Nothing Then
        failureReason = "Canonical inventory workbook could not be created."
        GoTo CleanExit
    End If

    Set wbInboxA = CreateCanonicalReceiveInboxWorkbookForTest(rootPath, "S11")
    Set wbInboxB = CreateCanonicalReceiveInboxWorkbookForTest(rootPath, "S12")
    If wbInboxA Is Nothing Or wbInboxB Is Nothing Then
        failureReason = "LAN inbox workbooks could not be created."
        GoTo CleanExit
    End If

    AddInboxReceiveEventRowForTest FindTableByName(wbInboxA, "tblInboxReceive"), "EVT-LAN-001", "WH75", "S11", currentUser, "SKU-LAN-LOCK", 4, "A1", "lan-station-a"
    AddInboxReceiveEventRowForTest FindTableByName(wbInboxB, "tblInboxReceive"), "EVT-LAN-002", "WH75", "S12", currentUser, "SKU-LAN-LOCK", 6, "B1", "lan-station-b"
    wbInboxA.Save
    wbInboxB.Save

    operatorPathA = rootPath & "\WH75_S11_Receiving_Operator.xlsb"
    operatorPathB = rootPath & "\WH75_S12_Receiving_Operator.xlsb"
    BuildSavedReceivingOperatorWorkbookForTest operatorPathA, "SKU-LAN-LOCK", "REF-LAN-OP-A", "SNAP-OLD-LAN-A", 0, "Z1"
    BuildSavedReceivingOperatorWorkbookForTest operatorPathB, "SKU-LAN-LOCK", "REF-LAN-OP-B", "SNAP-OLD-LAN-B", 0, "Z2"

    If Not modLockManager.AcquireLock("INVENTORY", "WH75", "svc_processor", "S11", wbInv, runIdA, message) Then
        failureReason = "Station S11 could not acquire inventory lock."
        GoTo CleanExit
    End If
    If modLockManager.AcquireLock("INVENTORY", "WH75", "svc_processor", "S12", wbInv, runIdB, message) Then
        failureReason = "Station S12 acquired inventory lock while S11 still held it."
        GoTo CleanExit
    End If

    Set loLocks = FindTableByName(wbInv, "tblLocks")
    If loLocks Is Nothing Then
        failureReason = "tblLocks not found in canonical inventory workbook."
        GoTo CleanExit
    End If
    If UCase$(CStr(GetTableValue(loLocks, 1, "Status"))) <> "HELD" Then
        failureReason = "Lock row was not HELD during contention."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loLocks, 1, "OwnerStationId")), "S11", vbTextCompare) <> 0 Then
        failureReason = "Lock row owner station drifted during contention."
        GoTo CleanExit
    End If

    If Not modLockManager.ReleaseLock("INVENTORY", runIdA, wbInv) Then
        failureReason = "Station S11 could not release inventory lock."
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch("WH75", 500, report)
    If processedCount <> 2 Then
        failureReason = "RunBatch did not process both LAN inbox rows. " & report & _
                        "; S11=" & DescribeInboxRowStateForTest(wbInboxA, "EVT-LAN-001") & _
                        "; S12=" & DescribeInboxRowStateForTest(wbInboxB, "EVT-LAN-002")
        GoTo CleanExit
    End If

    If Not AssertInboxRowStatusForTest(wbInboxA, "EVT-LAN-001", "PROCESSED") Then
        failureReason = "Station S11 inbox row was not marked PROCESSED."
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInboxB, "EVT-LAN-002", "PROCESSED") Then
        failureReason = "Station S12 inbox row was not marked PROCESSED."
        GoTo CleanExit
    End If

    Set loSku = wbInv.Worksheets("SkuBalance").ListObjects("tblSkuBalance")
    Set loLoc = wbInv.Worksheets("LocationBalance").ListObjects("tblLocationBalance")
    If loSku Is Nothing Or loLoc Is Nothing Then
        failureReason = "Projection tables missing after LAN processor run."
        GoTo CleanExit
    End If
    If FindRowByColumnValueInTable(loSku, "SKU", "SKU-LAN-LOCK") = 0 Then
        failureReason = "Projected SKU balance row missing after LAN processor run."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loSku, FindRowByColumnValueInTable(loSku, "SKU", "SKU-LAN-LOCK"), "QtyOnHand")) <> 10 Then
        failureReason = "Projected SKU balance did not equal combined LAN quantity."
        GoTo CleanExit
    End If
    If loLoc.ListRows.Count <> 2 Then
        failureReason = "Location projection did not retain both LAN station locations."
        GoTo CleanExit
    End If

    Set wbOpsA = Application.Workbooks.Open(operatorPathA)
    Set wbOpsB = Application.Workbooks.Open(operatorPathB)
    If wbOpsA Is Nothing Or wbOpsB Is Nothing Then
        failureReason = "Saved LAN operator workbook(s) could not be reopened."
        GoTo CleanExit
    End If

    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOpsA, report) Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOpsB, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOpsA, "WH75", "LOCAL", report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOpsB, "WH75", "LOCAL", report) Then GoTo CleanExit

    If Not AssertLanWorkbookState(wbOpsA, operatorPathA, "REF-LAN-OP-A", "SNAP-OLD-LAN-A", 10, "SKU-LAN-LOCK", "WH75.invSys.Snapshot.Inventory.xlsb|") Then
        failureReason = "Station S11 operator workbook was contaminated by LAN refresh."
        GoTo CleanExit
    End If
    If Not AssertLanWorkbookState(wbOpsB, operatorPathB, "REF-LAN-OP-B", "SNAP-OLD-LAN-B", 10, "SKU-LAN-LOCK", "WH75.invSys.Snapshot.Inventory.xlsb|") Then
        failureReason = "Station S12 operator workbook was contaminated by LAN refresh."
        GoTo CleanExit
    End If

    TestLanTwoStationProcessorRun_RespectsLockAndPreservesOperatorWorkbooks = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOpsA
    CloseWorkbookIfOpen wbOpsB
    CloseWorkbookIfOpen wbInboxA
    CloseWorkbookIfOpen wbInboxB
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7102, "TestLanTwoStationProcessorRun_RespectsLockAndPreservesOperatorWorkbooks", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestProcessor_DiscoversClosedConfiguredStationInboxWorkbook() As Long
    Dim rootPath As String
    Dim stationRoot As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim wbInboxCheck As Workbook
    Dim loInbox As ListObject
    Dim loSku As ListObject
    Dim processedCount As Long

    rootPath = BuildRuntimeTestRoot("phase6_lan_closed_inbox")
    stationRoot = rootPath & "\station_S22"

    On Error GoTo CleanFail
    If Len(Dir$(stationRoot, vbDirectory)) = 0 Then MkDir stationRoot

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH81", "S21") Then GoTo CleanExit
    SetConfigWarehouseValue "WH81.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    EnsureConfigStationRowValue "WH81.invSys.Config.xlsb", "S21", "WH81", "RoleDefault", "RECEIVE"
    EnsureConfigStationRowValue "WH81.invSys.Config.xlsb", "S22", "WH81", "PathInboxRoot", stationRoot & "\"
    If Not modConfig.Reload() Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH81") Then GoTo CleanExit

    currentUser = ResolveCurrentTestUserId()
    EnsureAuthCapabilityForTest "WH81", currentUser, "RECEIVE_POST", "WH81", "*"
    EnsureAuthCapabilityForTest "WH81", "svc_processor", "INBOX_PROCESS", "WH81", "*"

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH81", Array("SKU-LAN-DISK"))
    If wbInv Is Nothing Then
        failureReason = "Canonical inventory workbook could not be created."
        GoTo CleanExit
    End If

    Set wbInbox = CreateCanonicalReceiveInboxWorkbookForTest(stationRoot, "S22")
    If wbInbox Is Nothing Then
        failureReason = "Configured station inbox workbook could not be created."
        GoTo CleanExit
    End If

    Set loInbox = FindTableByName(wbInbox, "tblInboxReceive")
    AddInboxReceiveEventRowForTest loInbox, "EVT-LAN-DISK-001", "WH81", "S22", currentUser, "SKU-LAN-DISK", 5, "A1", "closed-configured-inbox"
    wbInbox.Save
    wbInbox.Close SaveChanges:=True
    Set wbInbox = Nothing

    processedCount = modProcessor.RunBatch("WH81", 500, report)
    If processedCount <> 1 Then
        failureReason = "RunBatch did not process the configured closed inbox workbook. " & report
        GoTo CleanExit
    End If

    Set wbInboxCheck = Application.Workbooks.Open(stationRoot & "\invSys.Inbox.Receiving.S22.xlsb")
    If Not AssertInboxRowStatusForTest(wbInboxCheck, "EVT-LAN-DISK-001", "PROCESSED") Then
        failureReason = "Configured station inbox row was not marked PROCESSED after closed-file discovery."
        GoTo CleanExit
    End If

    Set loSku = wbInv.Worksheets("SkuBalance").ListObjects("tblSkuBalance")
    If loSku Is Nothing Then
        failureReason = "Projected SKU balance table missing after closed inbox processing."
        GoTo CleanExit
    End If
    If FindRowByColumnValueInTable(loSku, "SKU", "SKU-LAN-DISK") = 0 Then
        failureReason = "Projected SKU balance row missing after closed inbox processing."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loSku, FindRowByColumnValueInTable(loSku, "SKU", "SKU-LAN-DISK"), "QtyOnHand")) <> 5 Then
        failureReason = "Projected SKU balance did not reflect closed inbox processing."
        GoTo CleanExit
    End If

    TestProcessor_DiscoversClosedConfiguredStationInboxWorkbook = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInboxCheck
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    CloseWorkbookByNameIfOpen "WH81.invSys.Config.xlsb"
    CloseWorkbookByNameIfOpen "WH81.invSys.Auth.xlsb"
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7110, "TestProcessor_DiscoversClosedConfiguredStationInboxWorkbook", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestSavedShippingWorkbook_RefreshPreservesStagingAndLogs() As Long
    Dim rootPath As String
    Dim operatorPath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loShipLog As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_saved_shipping_refresh")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH73", "S13") Then GoTo CleanExit
    SetConfigWarehouseValue "WH73.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    operatorPath = rootPath & "\WH73_S13_Shipping_Operator.xlsb"
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loShipLog = FindTableByName(wbOps, "AggregatePackages_Log")
    If loInv Is Nothing Or loShip Is Nothing Or loShipLog Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 907, "SKU-SHIP-001", "Shipping Refresh Item", "EA", "D4", 5
    AddShippingTallyRow loShip, "REF-SHIP-001", "Shipping Refresh Item", 6, 907, "EA", "D4", "ship note"
    AddAggregatePackagesLogRow loShipLog, "GUID-SHIP-001", "user1", "ADD", 907, "SKU-SHIP-001", "Shipping Refresh Item", 6, "6"

    wbOps.SaveAs Filename:=operatorPath, FileFormat:=50
    wbOps.Close SaveChanges:=False
    Set wbOps = Nothing

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH73", "SKU-SHIP-001", 18, CDate("2026-03-25 11:00:00"))
    If wbSnap Is Nothing Then GoTo CleanExit
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH73", "LOCAL", report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loShipLog = FindTableByName(wbOps, "AggregatePackages_Log")
    If loInv Is Nothing Or loShip Is Nothing Or loShipLog Is Nothing Then GoTo CleanExit

    If loShip.ListRows.Count <> 1 Then GoTo CleanExit
    If loShipLog.ListRows.Count <> 1 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loShip, 1, "REF_NUMBER")), "REF-SHIP-001", vbTextCompare) <> 0 Then GoTo CleanExit
    If CDbl(GetTableValue(loShip, 1, "QUANTITY")) <> 6 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loShipLog, 1, "GUID")), "GUID-SHIP-001", vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loShipLog, 1, "USER")), "user1", vbTextCompare) <> 0 Then GoTo CleanExit

    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 18 Then GoTo CleanExit
    If CDbl(GetTableValue(loInv, 1, "QtyAvailable")) <> 18 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loInv, 1, "ITEM_CODE")), "SKU-SHIP-001", vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "A1", vbTextCompare) <> 0 Then GoTo CleanExit
    If CBool(GetTableValue(loInv, 1, "IsStale")) <> False Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "LOCAL", vbTextCompare) <> 0 Then GoTo CleanExit

    TestSavedShippingWorkbook_RefreshPreservesStagingAndLogs = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestSavedShippingWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs() As Long
    Dim rootPath As String
    Dim operatorPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim eventIdOut As String
    Dim payloadJson As String
    Dim processedCount As Long
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loShipLog As ListObject
    Dim loInventoryLog As ListObject
    Dim invRow As Long
    Dim logRow As Long
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String

    rootPath = BuildRuntimeTestRoot("phase6_saved_shipping_post")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH79", "S19") Then GoTo CleanExit
    SetConfigWarehouseValue "WH79.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH79") Then GoTo CleanExit

    currentUser = ResolveCurrentTestUserId()
    EnsureAuthCapabilityForTest "WH79", currentUser, "SHIP_POST", "WH79", "*"
    EnsureAuthCapabilityForTest "WH79", "svc_processor", "INBOX_PROCESS", "WH79", "*"

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH79", Array("SKU-SHIP-POST"))
    Set wbInbox = CreateCanonicalShipInboxWorkbookForTest(rootPath, "S19")
    If wbInv Is Nothing Or wbInbox Is Nothing Then
        failureReason = "Canonical shipping runtime workbooks could not be created."
        GoTo CleanExit
    End If

    Set evt = CreateReceiveEventForTest("EVT-SHIP-SEED-001", "WH79", "S19", currentUser, "SKU-SHIP-POST", 10, "A1", "shipping seed")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-SHIP-SEED-001", statusOut, errorCode, errorMessage) Then
        failureReason = "Canonical shipping seed event failed: " & errorCode & "; " & errorMessage
        GoTo CleanExit
    End If

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH79", "SKU-SHIP-POST", 10, CDate("2026-03-25 12:15:00"))
    If wbSnap Is Nothing Then GoTo CleanExit
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    operatorPath = rootPath & "\WH79_S19_Shipping_Operator.xlsb"
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loShipLog = FindTableByName(wbOps, "AggregatePackages_Log")
    If loInv Is Nothing Or loShip Is Nothing Or loShipLog Is Nothing Then
        failureReason = "Saved shipping workbook surface was incomplete."
        GoTo CleanExit
    End If

    AddInvSysSeedRow loInv, 912, "SKU-SHIP-POST", "Shipping Post Item", "EA", "D4", 1
    AddShippingTallyRow loShip, "REF-SHIP-POST-001", "Shipping Post Item", 6, 912, "EA", "D4", "ship workflow"
    AddAggregatePackagesLogRow loShipLog, "GUID-SHIP-POST-001", currentUser, "ADD", 912, "SKU-SHIP-POST", "Shipping Post Item", 6, "6"
    wbOps.SaveAs Filename:=operatorPath, FileFormat:=50
    wbOps.Close SaveChanges:=False
    Set wbOps = Nothing

    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH79", "LOCAL", report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loShipLog = FindTableByName(wbOps, "AggregatePackages_Log")
    If loInv Is Nothing Or loShip Is Nothing Or loShipLog Is Nothing Then
        failureReason = "Saved shipping workbook tables were missing after reopen/refresh."
        GoTo CleanExit
    End If

    payloadJson = modRoleEventWriter.BuildPayloadJson( _
        modRoleEventWriter.CreatePayloadItem( _
            CLng(GetTableValue(loShip, 1, "ROW")), _
            CStr(GetTableValue(loInv, 1, "ITEM_CODE")), _
            CDbl(GetTableValue(loShip, 1, "QUANTITY")), _
            CStr(GetTableValue(loShip, 1, "LOCATION")), _
            CStr(GetTableValue(loShip, 1, "DESCRIPTION"))))

    If Not modRoleEventWriter.QueuePayloadEvent(CORE_EVENT_TYPE_SHIP, "WH79", "S19", currentUser, payloadJson, "saved-shipping-post", "", "", Now, wbInbox, eventIdOut, report) Then
        failureReason = "QueuePayloadEvent failed from saved shipping workbook: " & report
        GoTo CleanExit
    End If
    If Trim$(eventIdOut) = "" Then
        failureReason = "QueuePayloadEvent did not return an EventID for saved shipping workbook."
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch("WH79", 500, report)
    If processedCount <> 1 Then
        failureReason = "RunBatch did not process the saved shipping event. " & report
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, eventIdOut, "PROCESSED") Then
        failureReason = "Saved shipping inbox row was not marked PROCESSED."
        GoTo CleanExit
    End If

    Set loInventoryLog = FindTableByName(wbInv, "tblInventoryLog")
    If loInventoryLog Is Nothing Then
        failureReason = "Canonical inventory log was missing after saved shipping process."
        GoTo CleanExit
    End If
    logRow = FindRowByColumnValueInTable(loInventoryLog, "EventID", eventIdOut)
    If logRow = 0 Then
        failureReason = "Canonical inventory log did not record the saved shipping event."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInventoryLog, logRow, "EventType")), CORE_EVENT_TYPE_SHIP, vbTextCompare) <> 0 Then
        failureReason = "Canonical inventory log recorded unexpected event type for saved shipping workflow."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInventoryLog, logRow, "QtyDelta")) <> -6 Then
        failureReason = "Canonical inventory log QtyDelta was not negative for saved shipping workflow."
        GoTo CleanExit
    End If

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH79", "LOCAL", report) Then
        failureReason = "RefreshInventoryReadModelForWorkbook failed after saved shipping process: " & report
        GoTo CleanExit
    End If
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loShipLog = FindTableByName(wbOps, "AggregatePackages_Log")
    If loInv Is Nothing Or loShip Is Nothing Or loShipLog Is Nothing Then
        failureReason = "Saved shipping workbook tables were missing after process/refresh."
        GoTo CleanExit
    End If
    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-SHIP-POST")
    If invRow = 0 Then
        failureReason = "invSys did not retain shipping SKU after process/refresh."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "TOTAL INV")) <> 4 Then
        failureReason = "invSys TOTAL INV did not reflect saved shipping processing."
        GoTo CleanExit
    End If
    If loShip.ListRows.Count <> 1 Or loShipLog.ListRows.Count <> 1 Then
        failureReason = "Shipping staging/log tables changed after saved workflow processing."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loShip, 1, "REF_NUMBER")), "REF-SHIP-POST-001", vbTextCompare) <> 0 Then
        failureReason = "ShipmentsTally REF_NUMBER was not preserved across saved workflow processing."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loShipLog, 1, "GUID")), "GUID-SHIP-POST-001", vbTextCompare) <> 0 Then
        failureReason = "AggregatePackages_Log GUID was not preserved across saved workflow processing."
        GoTo CleanExit
    End If

    TestSavedShippingWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7108, "TestSavedShippingWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs() As Long
    Dim rootPath As String
    Dim operatorPath As String
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loProd As ListObject
    Dim loProdLog As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_saved_production_refresh")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH74", "S14") Then GoTo CleanExit
    SetConfigWarehouseValue "WH74.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    operatorPath = rootPath & "\WH74_S14_Production_Operator.xlsb"
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loProd = FindTableByName(wbOps, "ProductionOutput")
    Set loProdLog = FindTableByName(wbOps, "ProductionLog")
    If loInv Is Nothing Or loProd Is Nothing Or loProdLog Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 908, "SKU-PROD-001", "Production Refresh Item", "EA", "E5", 8
    AddProductionOutputRow loProd, "Blend", "Production Refresh Item", "EA", 7, "BATCH-001", "RECALL-001", 908
    AddProductionLogRow loProdLog, "Blend", "REC-001", "Production Refresh Item", "EA", 7, "E5", 908, "SKU-PROD-001", "GUID-PROD-001"

    wbOps.SaveAs Filename:=operatorPath, FileFormat:=50
    wbOps.Close SaveChanges:=False
    Set wbOps = Nothing

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH74", "SKU-PROD-001", 33, CDate("2026-03-25 11:30:00"))
    If wbSnap Is Nothing Then GoTo CleanExit
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wbOps, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH74", "LOCAL", report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loProd = FindTableByName(wbOps, "ProductionOutput")
    Set loProdLog = FindTableByName(wbOps, "ProductionLog")
    If loInv Is Nothing Or loProd Is Nothing Or loProdLog Is Nothing Then GoTo CleanExit

    If loProd.ListRows.Count <> 1 Then GoTo CleanExit
    If loProdLog.ListRows.Count <> 1 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loProd, 1, "PROCESS")), "Blend", vbTextCompare) <> 0 Then GoTo CleanExit
    If CDbl(GetTableValue(loProd, 1, "REAL OUTPUT")) <> 7 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loProdLog, 1, "GUID")), "GUID-PROD-001", vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loProdLog, 1, "ITEM_CODE")), "SKU-PROD-001", vbTextCompare) <> 0 Then GoTo CleanExit

    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 33 Then GoTo CleanExit
    If CDbl(GetTableValue(loInv, 1, "QtyAvailable")) <> 33 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loInv, 1, "ITEM_CODE")), "SKU-PROD-001", vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "A1", vbTextCompare) <> 0 Then GoTo CleanExit
    If CBool(GetTableValue(loInv, 1, "IsStale")) <> False Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "LOCAL", vbTextCompare) <> 0 Then GoTo CleanExit

    TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs() As Long
    Dim rootPath As String
    Dim operatorPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim eventIdOut As String
    Dim payloadJson As String
    Dim processedCount As Long
    Dim wbOps As Workbook
    Dim wbSnap As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loProd As ListObject
    Dim loProdLog As ListObject
    Dim loInventoryLog As ListObject
    Dim invRow As Long
    Dim logRow As Long

    rootPath = BuildRuntimeTestRoot("phase6_saved_production_post")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH80", "S20") Then GoTo CleanExit
    SetConfigWarehouseValue "WH80.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH80") Then GoTo CleanExit

    currentUser = ResolveCurrentTestUserId()
    EnsureAuthCapabilityForTest "WH80", currentUser, "PROD_POST", "WH80", "*"
    EnsureAuthCapabilityForTest "WH80", "svc_processor", "INBOX_PROCESS", "WH80", "*"

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH80", Array("SKU-PROD-POST"))
    Set wbInbox = CreateCanonicalProductionInboxWorkbookForTest(rootPath, "S20")
    If wbInv Is Nothing Or wbInbox Is Nothing Then
        failureReason = "Canonical production runtime workbooks could not be created."
        GoTo CleanExit
    End If

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH80", "SKU-PROD-POST", 0, CDate("2026-03-25 12:45:00"))
    If wbSnap Is Nothing Then GoTo CleanExit
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    operatorPath = rootPath & "\WH80_S20_Production_Operator.xlsb"
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wbOps, report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loProd = FindTableByName(wbOps, "ProductionOutput")
    Set loProdLog = FindTableByName(wbOps, "ProductionLog")
    If loInv Is Nothing Or loProd Is Nothing Or loProdLog Is Nothing Then
        failureReason = "Saved production workbook surface was incomplete."
        GoTo CleanExit
    End If

    AddInvSysSeedRow loInv, 913, "SKU-PROD-POST", "Production Post Item", "EA", "E5", 0
    AddProductionOutputRow loProd, "Blend", "Production Post Item", "EA", 7, "BATCH-POST-001", "RECALL-POST-001", 913
    AddProductionLogRow loProdLog, "Blend", "REC-POST-001", "Production Post Item", "EA", 7, "E5", 913, "SKU-PROD-POST", "GUID-PROD-POST-001"
    wbOps.SaveAs Filename:=operatorPath, FileFormat:=50
    wbOps.Close SaveChanges:=False
    Set wbOps = Nothing

    Set wbOps = Application.Workbooks.Open(operatorPath)
    If wbOps Is Nothing Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wbOps, report) Then GoTo CleanExit
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH80", "LOCAL", report) Then GoTo CleanExit

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loProd = FindTableByName(wbOps, "ProductionOutput")
    Set loProdLog = FindTableByName(wbOps, "ProductionLog")
    If loInv Is Nothing Or loProd Is Nothing Or loProdLog Is Nothing Then
        failureReason = "Saved production workbook tables were missing after reopen/refresh."
        GoTo CleanExit
    End If

    payloadJson = modRoleEventWriter.BuildPayloadJson( _
        modRoleEventWriter.CreatePayloadItem( _
            CLng(GetTableValue(loProd, 1, "ROW")), _
            "SKU-PROD-POST", _
            CDbl(GetTableValue(loProd, 1, "REAL OUTPUT")), _
            "FG", _
            CStr(GetTableValue(loProd, 1, "PROCESS")), _
            "COMPLETE"))

    If Not modRoleEventWriter.QueuePayloadEvent(CORE_EVENT_TYPE_PROD_COMPLETE, "WH80", "S20", currentUser, payloadJson, "saved-production-post", "", "", Now, wbInbox, eventIdOut, report) Then
        failureReason = "QueuePayloadEvent failed from saved production workbook: " & report
        GoTo CleanExit
    End If
    If Trim$(eventIdOut) = "" Then
        failureReason = "QueuePayloadEvent did not return an EventID for saved production workbook."
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch("WH80", 500, report)
    If processedCount <> 1 Then
        failureReason = "RunBatch did not process the saved production event. " & report
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, eventIdOut, "PROCESSED") Then
        failureReason = "Saved production inbox row was not marked PROCESSED."
        GoTo CleanExit
    End If

    Set loInventoryLog = FindTableByName(wbInv, "tblInventoryLog")
    If loInventoryLog Is Nothing Then
        failureReason = "Canonical inventory log was missing after saved production process."
        GoTo CleanExit
    End If
    logRow = FindRowByColumnValueInTable(loInventoryLog, "EventID", eventIdOut)
    If logRow = 0 Then
        failureReason = "Canonical inventory log did not record the saved production event."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInventoryLog, logRow, "EventType")), CORE_EVENT_TYPE_PROD_COMPLETE, vbTextCompare) <> 0 Then
        failureReason = "Canonical inventory log recorded unexpected event type for saved production workflow."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInventoryLog, logRow, "QtyDelta")) <> 7 Then
        failureReason = "Canonical inventory log QtyDelta was not positive for saved production workflow."
        GoTo CleanExit
    End If

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, "WH80", "LOCAL", report) Then
        failureReason = "RefreshInventoryReadModelForWorkbook failed after saved production process: " & report
        GoTo CleanExit
    End If
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loProd = FindTableByName(wbOps, "ProductionOutput")
    Set loProdLog = FindTableByName(wbOps, "ProductionLog")
    If loInv Is Nothing Or loProd Is Nothing Or loProdLog Is Nothing Then
        failureReason = "Saved production workbook tables were missing after process/refresh."
        GoTo CleanExit
    End If
    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-PROD-POST")
    If invRow = 0 Then
        failureReason = "invSys did not retain production SKU after process/refresh."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "TOTAL INV")) <> 7 Then
        failureReason = "invSys TOTAL INV did not reflect saved production processing."
        GoTo CleanExit
    End If
    If loProd.ListRows.Count <> 1 Or loProdLog.ListRows.Count <> 1 Then
        failureReason = "Production staging/log tables changed after saved workflow processing."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loProd, 1, "PROCESS")), "Blend", vbTextCompare) <> 0 Then
        failureReason = "ProductionOutput PROCESS was not preserved across saved workflow processing."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loProdLog, 1, "GUID")), "GUID-PROD-POST-001", vbTextCompare) <> 0 Then
        failureReason = "ProductionLog GUID was not preserved across saved workflow processing."
        GoTo CleanExit
    End If

    TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7109, "TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestSavedAdminWorkbook_ReopenRefreshReissuePreservesAudit() As Long
    Dim rootPath As String
    Dim adminPath As String
    Dim currentUser As String
    Dim wbAdmin As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim loAudit As ListObject
    Dim loPoison As ListObject
    Dim loInbox As ListObject
    Dim loLog As ListObject
    Dim corrections As Object
    Dim report As String
    Dim newEventId As String
    Dim poisonCount As Long
    Dim failureReason As String

    rootPath = BuildRuntimeTestRoot("phase6_saved_admin_reissue")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH76", "ADM1") Then GoTo CleanExit
    SetConfigWarehouseValue "WH76.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH76") Then GoTo CleanExit

    currentUser = ResolveCurrentTestUserId()
    EnsureAuthCapabilityForTest "WH76", currentUser, "ADMIN_MAINT", "WH76", "*"
    EnsureAuthCapabilityForTest "WH76", currentUser, "RECEIVE_POST", "WH76", "*"
    EnsureAuthCapabilityForTest "WH76", "svc_processor", "INBOX_PROCESS", "WH76", "*"

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH76", Array("SKU-001"))
    If wbInv Is Nothing Then
        failureReason = "Canonical inventory workbook could not be created."
        GoTo CleanExit
    End If
    Set wbInbox = CreateCanonicalReceiveInboxWorkbookForTest(rootPath, "ADM1")
    If wbInbox Is Nothing Then
        failureReason = "Canonical admin inbox workbook could not be created."
        GoTo CleanExit
    End If

    AddInboxReceiveEventRowForTest FindTableByName(wbInbox, "tblInboxReceive"), "EVT-ADMIN-POISON-001", "WH76", "ADM1", currentUser, "BAD-SKU", 6, "A1", "bad sku"
    If modProcessor.RunBatch("WH76", 500, report) <> 0 Then
        failureReason = "Initial processor run did not return poison-only result. " & report
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, "EVT-ADMIN-POISON-001", "POISON") Then
        failureReason = "Poison seed event was not left in POISON status."
        GoTo CleanExit
    End If

    adminPath = rootPath & "\WH76_ADM1_Admin_Operator.xlsb"
    Set wbAdmin = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wbAdmin, report) Then
        failureReason = "Initial admin legacy surface failed: " & report
        GoTo CleanExit
    End If
    If Not modAdminConsole.EnsureAdminSchema(wbAdmin, report) Then
        failureReason = "Initial admin schema failed: " & report
        GoTo CleanExit
    End If

    Set loAudit = FindTableByName(wbAdmin, "tblAdminAudit")
    If loAudit Is Nothing Then
        failureReason = "Initial admin audit table was missing."
        GoTo CleanExit
    End If
    AddAdminAuditRow loAudit, "SEED_ADMIN", currentUser, "WH76", "ADM1", "WORKBOOK", "WH76_ADM1_Admin_Operator", "seed", "seed row", "OK"
    wbAdmin.SaveAs Filename:=adminPath, FileFormat:=50
    wbAdmin.Close SaveChanges:=False
    Set wbAdmin = Nothing

    Set wbAdmin = Application.Workbooks.Open(adminPath)
    If wbAdmin Is Nothing Then
        failureReason = "Saved admin workbook could not be reopened."
        GoTo CleanExit
    End If
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wbAdmin, report) Then
        failureReason = "Reopened admin legacy surface failed: " & report
        GoTo CleanExit
    End If
    If Not modAdminConsole.EnsureAdminSchema(wbAdmin, report) Then
        failureReason = "Reopened admin schema failed: " & report
        GoTo CleanExit
    End If

    If Not modAdminConsole.RefreshAdminConsole(wbAdmin, report) Then
        failureReason = "RefreshAdminConsole failed after reopen: " & report
        GoTo CleanExit
    End If

    Set loAudit = FindTableByName(wbAdmin, "tblAdminAudit")
    Set loPoison = FindTableByName(wbAdmin, "tblAdminPoisonQueue")
    If loAudit Is Nothing Or loPoison Is Nothing Then
        failureReason = "Admin audit or poison queue table was missing after reopen."
        GoTo CleanExit
    End If
    If StrComp(wbAdmin.FullName, adminPath, vbTextCompare) <> 0 Then
        failureReason = "Saved admin workbook identity drifted after reopen."
        GoTo CleanExit
    End If
    If loAudit.ListRows.Count <> 1 Then
        failureReason = "Admin audit row count changed across reopen/refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loAudit, 1, "Action")), "SEED_ADMIN", vbTextCompare) <> 0 Then
        failureReason = "Seed admin audit row did not survive reopen/refresh."
        GoTo CleanExit
    End If
    poisonCount = loPoison.ListRows.Count
    If poisonCount <> 1 Then
        failureReason = "Admin poison queue count was not rebuilt correctly after reopen."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loPoison, 1, "EventID")), "EVT-ADMIN-POISON-001", vbTextCompare) <> 0 Then
        failureReason = "Admin poison queue did not point at the poisoned event."
        GoTo CleanExit
    End If
    If CLng(wbAdmin.Worksheets("AdminConsole").Range("B7").Value) <> 1 Then
        failureReason = "Admin console poison count was not refreshed after reopen."
        GoTo CleanExit
    End If

    Set corrections = CreateObject("Scripting.Dictionary")
    corrections.CompareMode = vbTextCompare
    corrections.Add "SKU", "SKU-001"
    corrections.Add "Note", "fixed sku"

    If Not modAdminConsole.ReissuePoisonEvent(wbInbox.Name, "tblInboxReceive", "EVT-ADMIN-POISON-001", currentUser, corrections, "fix sku", wbAdmin, newEventId, report) Then
        failureReason = "ReissuePoisonEvent failed from saved admin workbook: " & report
        GoTo CleanExit
    End If
    If newEventId = "" Then
        failureReason = "ReissuePoisonEvent did not return a new child EventID."
        GoTo CleanExit
    End If
    If modAdminConsole.RunProcessorFromConsole(currentUser, "WH76", wbAdmin, report) <> 1 Then
        failureReason = "RunProcessorFromConsole did not process the reissued event. " & report
        GoTo CleanExit
    End If

    Set loInbox = FindTableByName(wbInbox, "tblInboxReceive")
    Set loLog = FindTableByName(wbInv, "tblInventoryLog")
    Set loAudit = FindTableByName(wbAdmin, "tblAdminAudit")
    Set loPoison = FindTableByName(wbAdmin, "tblAdminPoisonQueue")
    If loInbox Is Nothing Or loLog Is Nothing Or loAudit Is Nothing Or loPoison Is Nothing Then
        failureReason = "Admin workflow tables were missing after reissue/processor run."
        GoTo CleanExit
    End If

    If FindRowByColumnValueInTable(loInbox, "EventID", newEventId) = 0 Then
        failureReason = "Reissued child event row was not found in the inbox."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInbox, 1, "Status")), "POISON", vbTextCompare) <> 0 Then
        failureReason = "Original poisoned row lost POISON status after reissue."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInbox, FindRowByColumnValueInTable(loInbox, "EventID", newEventId), "Status")), "PROCESSED", vbTextCompare) <> 0 Then
        failureReason = "Reissued child row was not processed."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInbox, FindRowByColumnValueInTable(loInbox, "EventID", newEventId), "ParentEventId")), "EVT-ADMIN-POISON-001", vbTextCompare) <> 0 Then
        failureReason = "Reissued child row did not preserve ParentEventId."
        GoTo CleanExit
    End If
    If FindRowByColumnValueInTable(loLog, "EventID", newEventId) = 0 Then
        failureReason = "Canonical inventory log did not record the reissued event."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loLog, FindRowByColumnValueInTable(loLog, "EventID", newEventId), "SKU")), "SKU-001", vbTextCompare) <> 0 Then
        failureReason = "Canonical inventory log recorded the wrong SKU for the reissued event."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loLog, FindRowByColumnValueInTable(loLog, "EventID", newEventId), "QtyDelta")) <> 6 Then
        failureReason = "Canonical inventory log recorded the wrong quantity for the reissued event."
        GoTo CleanExit
    End If
    If FindRowByColumnValueInTable(loAudit, "Action", "REISSUE_POISON") = 0 Then
        failureReason = "Admin audit did not record REISSUE_POISON."
        GoTo CleanExit
    End If
    If FindRowByColumnValueInTable(loAudit, "Action", "RUN_PROCESSOR") = 0 Then
        failureReason = "Admin audit did not record RUN_PROCESSOR."
        GoTo CleanExit
    End If
    If loAudit.ListRows.Count <> 3 Then
        failureReason = "Admin audit row count drifted after reissue/processor run."
        GoTo CleanExit
    End If
    If loPoison.ListRows.Count <> 1 Then
        failureReason = "Admin poison queue count drifted after reissue/processor run."
        GoTo CleanExit
    End If
    If CLng(wbAdmin.Worksheets("AdminConsole").Range("B8").Value) <> 1 Then
        failureReason = "Admin console processed count was not refreshed after processor run."
        GoTo CleanExit
    End If

    TestSavedAdminWorkbook_ReopenRefreshReissuePreservesAudit = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbAdmin
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7105, "TestSavedAdminWorkbook_ReopenRefreshReissuePreservesAudit", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestApplyReceive_RebuildsDeletedProjectionTablesInCanonicalWorkbook() As Long
    Dim rootPath As String
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim report As String
    Dim loSku As ListObject
    Dim loLoc As ListObject
    Dim loStatus As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_projection_rebuild")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH70", "S10") Then GoTo CleanExit
    SetConfigWarehouseValue "WH70.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH70", Array("SKU-PR-001"))
    If wbInv Is Nothing Then GoTo CleanExit

    Set evt = CreateReceiveEventForTest("EVT-PR-001", "WH70", "S10", "user1", "SKU-PR-001", 5, "A1", "seed projection")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-PR-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    DeleteTableSurfaceForTest wbInv.Worksheets("SkuBalance"), "tblSkuBalance"
    DeleteTableSurfaceForTest wbInv.Worksheets("LocationBalance"), "tblLocationBalance"
    wbInv.Save

    Set evt = CreateReceiveEventForTest("EVT-PR-002", "WH70", "S10", "user1", "SKU-PR-001", 2, "A1", "rebuild after delete")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-PR-002", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loSku = wbInv.Worksheets("SkuBalance").ListObjects("tblSkuBalance")
    Set loLoc = wbInv.Worksheets("LocationBalance").ListObjects("tblLocationBalance")
    Set loStatus = wbInv.Worksheets("LedgerStatus").ListObjects("tblInventoryLedgerStatus")

    If loSku.ListRows.Count <> 1 Then GoTo CleanExit
    If loLoc.ListRows.Count <> 1 Then GoTo CleanExit
    If loStatus.ListRows.Count <> 1 Then GoTo CleanExit

    If StrComp(CStr(GetTableValue(loSku, 1, "SKU")), "SKU-PR-001", vbTextCompare) <> 0 Then GoTo CleanExit
    If CDbl(GetTableValue(loSku, 1, "QtyOnHand")) <> 7 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loLoc, 1, "Location")), "A1", vbTextCompare) <> 0 Then GoTo CleanExit
    If CDbl(GetTableValue(loLoc, 1, "QtyOnHand")) <> 7 Then GoTo CleanExit
    If CLng(GetTableValue(loStatus, 1, "TotalEventRows")) <> 2 Then GoTo CleanExit
    If CLng(GetTableValue(loStatus, 1, "TotalAppliedEvents")) <> 2 Then GoTo CleanExit
    If StrComp(CStr(GetTableValue(loStatus, 1, "LastEventId")), "EVT-PR-002", vbTextCompare) <> 0 Then GoTo CleanExit
    If CLng(GetTableValue(loStatus, 1, "DistinctSkuCount")) <> 1 Then GoTo CleanExit
    If CLng(GetTableValue(loStatus, 1, "DistinctLocationCount")) <> 1 Then GoTo CleanExit
    If Not IsDate(GetTableValue(loStatus, 1, "ProjectionRebuiltAtUTC")) Then GoTo CleanExit

    TestApplyReceive_RebuildsDeletedProjectionTablesInCanonicalWorkbook = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function GetTableValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    GetTableValue = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value
End Function

Private Function AssertLanWorkbookState(ByVal wbOps As Workbook, _
                                        ByVal expectedPath As String, _
                                        ByVal expectedRef As String, _
                                        ByVal expectedSnapshotLogId As String, _
                                        ByVal expectedTotalInv As Double, _
                                        ByVal expectedSku As String, _
                                        ByVal expectedSnapshotPrefix As String) As Boolean
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject

    If wbOps Is Nothing Then Exit Function
    If StrComp(wbOps.FullName, expectedPath, vbTextCompare) <> 0 Then Exit Function

    Set loInv = FindTableByName(wbOps, "invSys")
    Set loRecv = FindTableByName(wbOps, "ReceivedTally")
    Set loLog = FindTableByName(wbOps, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then Exit Function

    If loRecv.ListRows.Count <> 1 Then Exit Function
    If loLog.ListRows.Count <> 1 Then Exit Function
    If StrComp(CStr(GetTableValue(loRecv, 1, "REF_NUMBER")), expectedRef, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CStr(GetTableValue(loLog, 1, "REF_NUMBER")), expectedRef, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CStr(GetTableValue(loLog, 1, "SNAPSHOT_ID")), expectedSnapshotLogId, vbTextCompare) <> 0 Then Exit Function

    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> expectedTotalInv Then Exit Function
    If CDbl(GetTableValue(loInv, 1, "QtyAvailable")) <> expectedTotalInv Then Exit Function
    If StrComp(CStr(GetTableValue(loInv, 1, "ITEM_CODE")), expectedSku, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CStr(GetTableValue(loInv, 1, "LOCATION")), "A1", vbTextCompare) <> 0 Then Exit Function
    If InStr(1, CStr(GetTableValue(loInv, 1, "SnapshotId")), expectedSnapshotPrefix, vbTextCompare) <> 1 Then Exit Function
    If CBool(GetTableValue(loInv, 1, "IsStale")) <> False Then Exit Function
    If StrComp(CStr(GetTableValue(loInv, 1, "SourceType")), "LOCAL", vbTextCompare) <> 0 Then Exit Function
    If Not IsDate(GetTableValue(loInv, 1, "LastRefreshUTC")) Then Exit Function
    If Not IsDate(GetTableValue(loInv, 1, "LAST EDITED")) Then Exit Function

    AssertLanWorkbookState = True
End Function

Private Function ResolveCurrentTestUserId() As String
    ResolveCurrentTestUserId = Trim$(Environ$("USERNAME"))
    If ResolveCurrentTestUserId = "" Then ResolveCurrentTestUserId = Trim$(Application.UserName)
    If ResolveCurrentTestUserId = "" Then ResolveCurrentTestUserId = "user1"
End Function

Private Sub EnsureAuthCapabilityForTest(ByVal warehouseId As String, _
                                        ByVal userId As String, _
                                        ByVal capability As String, _
                                        ByVal capabilityWarehouseId As String, _
                                        ByVal stationId As String)
    Dim wbAuth As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim rowIndex As Long
    Dim lr As ListRow
    Dim usersWasProtected As Boolean
    Dim capsWasProtected As Boolean
    Dim report As String
    Dim openedTransient As Boolean

    Set wbAuth = FindWorkbookByName(warehouseId & ".invSys.Auth.xlsb")
    If wbAuth Is Nothing Then
        Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime(warehouseId, "svc_processor", "", report)
        openedTransient = Not wbAuth Is Nothing
    End If
    If wbAuth Is Nothing Then Exit Sub

    Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    Set loCaps = wbAuth.Worksheets("Capabilities").ListObjects("tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then GoTo CleanExit

    usersWasProtected = BeginEditableSheetForTest(loUsers.Parent)
    capsWasProtected = BeginEditableSheetForTest(loCaps.Parent)

    On Error GoTo CleanFail
    rowIndex = FindRowByColumnValueInTable(loUsers, "UserId", userId)
    If rowIndex = 0 Then
        Set lr = loUsers.ListRows.Add
        SetTableCell loUsers, lr.Index, "UserId", userId
        SetTableCell loUsers, lr.Index, "DisplayName", userId
        SetTableCell loUsers, lr.Index, "Status", "Active"
    Else
        SetTableCell loUsers, rowIndex, "Status", "Active"
    End If

    rowIndex = FindCapabilityRowForTest(loCaps, userId, capability, capabilityWarehouseId, stationId)
    If rowIndex = 0 Then
        Set lr = loCaps.ListRows.Add
        rowIndex = lr.Index
    End If
    SetTableCell loCaps, rowIndex, "UserId", userId
    SetTableCell loCaps, rowIndex, "Capability", capability
    SetTableCell loCaps, rowIndex, "WarehouseId", capabilityWarehouseId
    SetTableCell loCaps, rowIndex, "StationId", stationId
    SetTableCell loCaps, rowIndex, "Status", "ACTIVE"
    wbAuth.Save
CleanExit:
    RestoreSheetProtectionForTest loCaps.Parent, capsWasProtected
    RestoreSheetProtectionForTest loUsers.Parent, usersWasProtected
    CloseTransientWorkbookForTest wbAuth, openedTransient
    Exit Sub
CleanFail:
    RestoreSheetProtectionForTest loCaps.Parent, capsWasProtected
    RestoreSheetProtectionForTest loUsers.Parent, usersWasProtected
    CloseTransientWorkbookForTest wbAuth, openedTransient
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function FindCapabilityRowForTest(ByVal lo As ListObject, _
                                          ByVal userId As String, _
                                          ByVal capability As String, _
                                          ByVal warehouseId As String, _
                                          ByVal stationId As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(GetTableValue(lo, i, "UserId")), userId, vbTextCompare) = 0 _
           And StrComp(CStr(GetTableValue(lo, i, "Capability")), capability, vbTextCompare) = 0 _
           And StrComp(CStr(GetTableValue(lo, i, "WarehouseId")), warehouseId, vbTextCompare) = 0 _
           And StrComp(CStr(GetTableValue(lo, i, "StationId")), stationId, vbTextCompare) = 0 Then
            FindCapabilityRowForTest = i
            Exit Function
        End If
    Next i
End Function

Private Function CreateCanonicalReceiveInboxWorkbookForTest(ByVal rootPath As String, ByVal stationId As String) As Workbook
    Dim wb As Workbook
    Dim targetPath As String
    Dim report As String

    targetPath = rootPath & "\invSys.Inbox.Receiving." & stationId & ".xlsb"
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "InboxReceive"
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    If Not modProcessor.EnsureReceiveInboxSchema(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Function
    End If
    wb.Save
    Set CreateCanonicalReceiveInboxWorkbookForTest = wb
End Function

Private Function CreateCanonicalShipInboxWorkbookForTest(ByVal rootPath As String, ByVal stationId As String) As Workbook
    Dim wb As Workbook
    Dim targetPath As String
    Dim report As String

    targetPath = rootPath & "\invSys.Inbox.Shipping." & stationId & ".xlsb"
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "InboxShip"
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    If Not modProcessor.EnsureShipInboxSchema(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Function
    End If
    wb.Save
    Set CreateCanonicalShipInboxWorkbookForTest = wb
End Function

Private Function CreateCanonicalProductionInboxWorkbookForTest(ByVal rootPath As String, ByVal stationId As String) As Workbook
    Dim wb As Workbook
    Dim targetPath As String
    Dim report As String

    targetPath = rootPath & "\invSys.Inbox.Production." & stationId & ".xlsb"
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "InboxProd"
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    If Not modProcessor.EnsureProductionInboxSchema(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Function
    End If
    wb.Save
    Set CreateCanonicalProductionInboxWorkbookForTest = wb
End Function

Private Sub AddInboxReceiveEventRowForTest(ByVal lo As ListObject, _
                                           ByVal eventId As String, _
                                           ByVal warehouseId As String, _
                                           ByVal stationId As String, _
                                           ByVal userId As String, _
                                           ByVal sku As String, _
                                           ByVal qty As Double, _
                                           ByVal locationVal As String, _
                                           ByVal noteVal As String)
    Dim lr As ListRow
    Dim sheetWasProtected As Boolean

    If lo Is Nothing Then Exit Sub
    sheetWasProtected = BeginEditableSheetForTest(lo.Parent)

    On Error GoTo CleanFail
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValue(lo, 1, "EventID"))) = "" _
        And Trim$(CStr(GetTableValue(lo, 1, "SKU"))) = "" Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "EventID", eventId
    SetTableCell lo, lr.Index, "EventType", "RECEIVE"
    SetTableCell lo, lr.Index, "CreatedAtUTC", Now
    SetTableCell lo, lr.Index, "WarehouseId", warehouseId
    SetTableCell lo, lr.Index, "StationId", stationId
    SetTableCell lo, lr.Index, "UserId", userId
    SetTableCell lo, lr.Index, "SKU", sku
    SetTableCell lo, lr.Index, "Qty", qty
    SetTableCell lo, lr.Index, "Location", locationVal
    SetTableCell lo, lr.Index, "Note", noteVal
    SetTableCell lo, lr.Index, "Status", "NEW"
CleanExit:
    RestoreSheetProtectionForTest lo.Parent, sheetWasProtected
    Exit Sub
CleanFail:
    RestoreSheetProtectionForTest lo.Parent, sheetWasProtected
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function AssertInboxRowStatusForTest(ByVal wb As Workbook, ByVal eventId As String, ByVal expectedStatus As String) As Boolean
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = FindInboxTableForTest(wb)
    If lo Is Nothing Then Exit Function
    rowIndex = FindRowByColumnValueInTable(lo, "EventID", eventId)
    If rowIndex = 0 Then Exit Function
    If StrComp(CStr(GetTableValue(lo, rowIndex, "Status")), expectedStatus, vbTextCompare) <> 0 Then Exit Function
    AssertInboxRowStatusForTest = True
End Function

Private Function DescribeInboxRowStateForTest(ByVal wb As Workbook, ByVal eventId As String) As String
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = FindInboxTableForTest(wb)
    If lo Is Nothing Then
        DescribeInboxRowStateForTest = "missing-table"
        Exit Function
    End If

    rowIndex = FindRowByColumnValueInTable(lo, "EventID", eventId)
    If rowIndex = 0 Then
        DescribeInboxRowStateForTest = "missing-row"
        Exit Function
    End If

    DescribeInboxRowStateForTest = _
        "Status=" & CStr(GetTableValue(lo, rowIndex, "Status")) & _
        ", ErrorCode=" & CStr(GetTableValue(lo, rowIndex, "ErrorCode")) & _
        ", ErrorMessage=" & CStr(GetTableValue(lo, rowIndex, "ErrorMessage"))
End Function

Private Function FindInboxTableForTest(ByVal wb As Workbook) As ListObject
    Set FindInboxTableForTest = FindTableByName(wb, "tblInboxReceive")
    If Not FindInboxTableForTest Is Nothing Then Exit Function
    Set FindInboxTableForTest = FindTableByName(wb, "tblInboxShip")
    If Not FindInboxTableForTest Is Nothing Then Exit Function
    Set FindInboxTableForTest = FindTableByName(wb, "tblInboxProd")
End Function

Private Function FindRowByColumnValueInTable(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(GetTableValue(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByColumnValueInTable = i
            Exit Function
        End If
    Next i
End Function

Private Sub BuildSavedReceivingOperatorWorkbookForTest(ByVal targetPath As String, _
                                                       ByVal sku As String, _
                                                       ByVal refNumber As String, _
                                                       ByVal snapshotLogId As String, _
                                                       ByVal totalInv As Double, _
                                                       ByVal locationVal As String)
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loRecv As ListObject
    Dim loLog As ListObject

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Sub
    End If

    Set loInv = FindTableByName(wb, "invSys")
    Set loRecv = FindTableByName(wb, "ReceivedTally")
    Set loLog = FindTableByName(wb, "ReceivedLog")
    If loInv Is Nothing Or loRecv Is Nothing Or loLog Is Nothing Then
        CloseWorkbookIfOpen wb
        Exit Sub
    End If

    AddInvSysSeedRow loInv, 999, sku, "LAN Processor Item", "EA", locationVal, totalInv
    AddReceivedTallyRow loRecv, refNumber, "LAN Processor Item", 1, 999
    AddReceivedLogRow loLog, snapshotLogId, refNumber, "LAN Processor Item", 1, "EA", "Vendor", locationVal, sku, 999

    wb.SaveAs Filename:=targetPath, FileFormat:=50
    wb.Close SaveChanges:=False
End Sub

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

Private Function FindWorkbookByFullPathForTest(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindWorkbookByFullPathForTest = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub CloseWorkbookByNameIfOpen(ByVal workbookName As String)
    Dim wb As Workbook

    Set wb = FindWorkbookByName(workbookName)
    If wb Is Nothing Then Exit Sub
    CloseWorkbookIfOpen wb
End Sub

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

Private Sub AddReceivedLogRow(ByVal lo As ListObject, _
                              ByVal snapshotId As String, _
                              ByVal refNumber As String, _
                              ByVal itemName As String, _
                              ByVal qty As Double, _
                              ByVal uom As String, _
                              ByVal vendorName As String, _
                              ByVal locationVal As String, _
                              ByVal sku As String, _
                              ByVal rowValue As Long)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValue(lo, 1, "SNAPSHOT_ID"))) = "" _
        And Trim$(CStr(GetTableValue(lo, 1, "REF_NUMBER"))) = "" _
        And NzDblForTest(GetTableValue(lo, 1, "QUANTITY")) = 0 Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "SNAPSHOT_ID", snapshotId
    SetTableCell lo, lr.Index, "ENTRY_DATE", CDate("2026-03-25 08:00:00")
    SetTableCell lo, lr.Index, "REF_NUMBER", refNumber
    SetTableCell lo, lr.Index, "ITEMS", itemName
    SetTableCell lo, lr.Index, "QUANTITY", qty
    SetTableCell lo, lr.Index, "UOM", uom
    SetTableCell lo, lr.Index, "VENDOR", vendorName
    SetTableCell lo, lr.Index, "LOCATION", locationVal
    SetTableCell lo, lr.Index, "ITEM_CODE", sku
    SetTableCell lo, lr.Index, "ROW", rowValue
End Sub

Private Sub AddShippingTallyRow(ByVal lo As ListObject, _
                                ByVal refNumber As String, _
                                ByVal itemName As String, _
                                ByVal qty As Double, _
                                ByVal rowValue As Long, _
                                ByVal uom As String, _
                                ByVal locationVal As String, _
                                ByVal descriptionVal As String)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    Set lr = lo.ListRows(1)
    SetTableCell lo, lr.Index, "REF_NUMBER", refNumber
    SetTableCell lo, lr.Index, "ITEMS", itemName
    SetTableCell lo, lr.Index, "QUANTITY", qty
    SetTableCell lo, lr.Index, "ROW", rowValue
    SetTableCell lo, lr.Index, "UOM", uom
    SetTableCell lo, lr.Index, "LOCATION", locationVal
    SetTableCell lo, lr.Index, "DESCRIPTION", descriptionVal
End Sub

Private Sub AddAggregatePackagesLogRow(ByVal lo As ListObject, _
                                       ByVal guidVal As String, _
                                       ByVal userId As String, _
                                       ByVal actionVal As String, _
                                       ByVal rowValue As Long, _
                                       ByVal sku As String, _
                                       ByVal itemName As String, _
                                       ByVal qtyDelta As Double, _
                                       ByVal newValue As String)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValue(lo, 1, "GUID"))) = "" _
        And Trim$(CStr(GetTableValue(lo, 1, "USER"))) = "" Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "GUID", guidVal
    SetTableCell lo, lr.Index, "USER", userId
    SetTableCell lo, lr.Index, "ACTION", actionVal
    SetTableCell lo, lr.Index, "ROW", rowValue
    SetTableCell lo, lr.Index, "ITEM_CODE", sku
    SetTableCell lo, lr.Index, "ITEM", itemName
    SetTableCell lo, lr.Index, "QTY_DELTA", qtyDelta
    SetTableCell lo, lr.Index, "NEW_VALUE", newValue
    SetTableCell lo, lr.Index, "TIMESTAMP", CDate("2026-03-25 10:45:00")
End Sub

Private Sub AddProductionOutputRow(ByVal lo As ListObject, _
                                   ByVal processName As String, _
                                   ByVal outputName As String, _
                                   ByVal uom As String, _
                                   ByVal realOutput As Double, _
                                   ByVal batchVal As String, _
                                   ByVal recallCode As String, _
                                   ByVal rowValue As Long)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValue(lo, 1, "PROCESS"))) = "" _
        And Trim$(CStr(GetTableValue(lo, 1, "OUTPUT"))) = "" Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "PROCESS", processName
    SetTableCell lo, lr.Index, "OUTPUT", outputName
    SetTableCell lo, lr.Index, "UOM", uom
    SetTableCell lo, lr.Index, "REAL OUTPUT", realOutput
    SetTableCell lo, lr.Index, "BATCH", batchVal
    SetTableCell lo, lr.Index, "RECALL CODE", recallCode
    SetTableCell lo, lr.Index, "ROW", rowValue
End Sub

Private Sub AddProductionLogRow(ByVal lo As ListObject, _
                                ByVal recipeName As String, _
                                ByVal recipeId As String, _
                                ByVal itemName As String, _
                                ByVal uom As String, _
                                ByVal qty As Double, _
                                ByVal locationVal As String, _
                                ByVal rowValue As Long, _
                                ByVal sku As String, _
                                ByVal guidVal As String)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValue(lo, 1, "RECIPE"))) = "" _
        And Trim$(CStr(GetTableValue(lo, 1, "ITEM_CODE"))) = "" Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "TIMESTAMP", CDate("2026-03-25 11:10:00")
    SetTableCell lo, lr.Index, "RECIPE", recipeName
    SetTableCell lo, lr.Index, "RECIPE_ID", recipeId
    SetTableCell lo, lr.Index, "PROCESS", recipeName
    SetTableCell lo, lr.Index, "OUTPUT", itemName
    SetTableCell lo, lr.Index, "REAL OUTPUT", qty
    SetTableCell lo, lr.Index, "ITEM_CODE", sku
    SetTableCell lo, lr.Index, "ITEM", itemName
    SetTableCell lo, lr.Index, "UOM", uom
    SetTableCell lo, lr.Index, "QUANTITY", qty
    SetTableCell lo, lr.Index, "LOCATION", locationVal
    SetTableCell lo, lr.Index, "ROW", rowValue
    SetTableCell lo, lr.Index, "GUID", guidVal
End Sub

Private Sub AddAdminAuditRow(ByVal lo As ListObject, _
                             ByVal actionName As String, _
                             ByVal userId As String, _
                             ByVal warehouseId As String, _
                             ByVal stationId As String, _
                             ByVal targetType As String, _
                             ByVal targetId As String, _
                             ByVal reasonVal As String, _
                             ByVal detailVal As String, _
                             ByVal resultCode As String)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValue(lo, 1, "Action"))) = "" _
        And Trim$(CStr(GetTableValue(lo, 1, "UserId"))) = "" Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "LoggedAtUTC", CDate("2026-03-25 12:00:00")
    SetTableCell lo, lr.Index, "Action", actionName
    SetTableCell lo, lr.Index, "UserId", userId
    SetTableCell lo, lr.Index, "WarehouseId", warehouseId
    SetTableCell lo, lr.Index, "StationId", stationId
    SetTableCell lo, lr.Index, "TargetType", targetType
    SetTableCell lo, lr.Index, "TargetId", targetId
    SetTableCell lo, lr.Index, "Reason", reasonVal
    SetTableCell lo, lr.Index, "Detail", detailVal
    SetTableCell lo, lr.Index, "Result", resultCode
End Sub

Private Sub SetTableCell(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    If lo Is Nothing Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueIn
End Sub

Private Sub SetConfigWarehouseValue(ByVal workbookName As String, ByVal columnName As String, ByVal valueIn As Variant)
    Dim wb As Workbook
    Dim lo As ListObject
    Dim report As String
    Dim openedTransient As Boolean

    Set wb = FindWorkbookByName(workbookName)
    If wb Is Nothing Then
        Set wb = OpenConfigWorkbookForTest(workbookName, report, openedTransient)
    End If
    If wb Is Nothing Then Exit Sub
    Set lo = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    If lo Is Nothing Then GoTo CleanExit
    lo.DataBodyRange.Cells(1, lo.ListColumns(columnName).Index).Value = valueIn
    wb.Save
CleanExit:
    CloseTransientWorkbookForTest wb, openedTransient
End Sub

Private Sub EnsureConfigStationRowValue(ByVal workbookName As String, _
                                        ByVal stationId As String, _
                                        ByVal warehouseId As String, _
                                        ByVal columnName As String, _
                                        ByVal valueIn As Variant)
    Dim wb As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim lr As ListRow
    Dim report As String
    Dim openedTransient As Boolean

    Set wb = FindWorkbookByName(workbookName)
    If wb Is Nothing Then
        Set wb = OpenConfigWorkbookForTest(workbookName, report, openedTransient)
    End If
    If wb Is Nothing Then Exit Sub
    Set lo = wb.Worksheets("StationConfig").ListObjects("tblStationConfig")
    If lo Is Nothing Then GoTo CleanExit

    rowIndex = FindRowByColumnValueInTable(lo, "StationId", stationId)
    If rowIndex = 0 Then
        Set lr = lo.ListRows.Add
        rowIndex = lr.Index
        SetTableCell lo, rowIndex, "StationId", stationId
        SetTableCell lo, rowIndex, "WarehouseId", warehouseId
        SetTableCell lo, rowIndex, "StationName", stationId
        SetTableCell lo, rowIndex, "RoleDefault", "RECEIVE"
    End If

    SetTableCell lo, rowIndex, columnName, valueIn
    wb.Save
CleanExit:
    CloseTransientWorkbookForTest wb, openedTransient
End Sub

Private Function OpenConfigWorkbookForTest(ByVal workbookName As String, _
                                           ByRef report As String, _
                                           ByRef openedTransient As Boolean) As Workbook
    Dim warehouseId As String
    Dim alreadyOpen As Workbook

    Set alreadyOpen = FindWorkbookByName(workbookName)
    If Not alreadyOpen Is Nothing Then
        Set OpenConfigWorkbookForTest = alreadyOpen
        Exit Function
    End If

    warehouseId = InferWarehouseIdFromWorkbookNameForTest(workbookName)
    If warehouseId = "" Then Exit Function

    Set OpenConfigWorkbookForTest = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime(warehouseId, "", "", report)
    openedTransient = Not OpenConfigWorkbookForTest Is Nothing
End Function

Private Function InferWarehouseIdFromWorkbookNameForTest(ByVal workbookName As String) As String
    Dim dotPos As Long

    dotPos = InStr(1, workbookName, ".", vbTextCompare)
    If dotPos > 1 Then InferWarehouseIdFromWorkbookNameForTest = Left$(workbookName, dotPos - 1)
End Function

Private Sub CloseTransientWorkbookForTest(ByVal wb As Workbook, ByVal openedTransient As Boolean)
    If Not openedTransient Then Exit Sub
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    If Not wb.ReadOnly Then
        If wb.Saved = False Then wb.Save
    End If
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function CreateCanonicalInventoryWorkbookForTest(ByVal rootPath As String, ByVal warehouseId As String, ByVal skuList As Variant) As Workbook
    Dim wb As Workbook
    Dim targetPath As String
    Dim report As String

    targetPath = rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb"
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    If Not modInventorySchema.EnsureInventorySchema(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Function
    End If
    EnsureSkuCatalogForTest wb, skuList
    wb.Save
    Set CreateCanonicalInventoryWorkbookForTest = wb
End Function

Private Function CreateInventoryWorkbookForTestWithName(ByVal rootPath As String, ByVal workbookName As String, ByVal skuList As Variant) As Workbook
    Dim wb As Workbook
    Dim targetPath As String
    Dim report As String

    targetPath = rootPath & "\" & workbookName
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    If Not modInventorySchema.EnsureInventorySchema(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Function
    End If
    EnsureSkuCatalogForTest wb, skuList
    wb.Save
    Set CreateInventoryWorkbookForTestWithName = wb
End Function

Private Function CreateManagedInventoryDonorWorkbookForTest(ByVal rootPath As String, ByVal workbookName As String) As Workbook
    Dim wb As Workbook
    Dim report As String
    Dim targetPath As String

    targetPath = rootPath & "\" & workbookName
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Function
    End If
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Function
    End If
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then
        CloseWorkbookIfOpen wb
        Exit Function
    End If
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    Set CreateManagedInventoryDonorWorkbookForTest = wb
End Function

Private Sub EnsureSkuCatalogForTest(ByVal wb As Workbook, ByVal skuList As Variant)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim i As Long
    Dim sheetWasProtected As Boolean

    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = wb.Worksheets("SkuCatalog")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set lo = ws.ListObjects("tblSkuCatalog")
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    sheetWasProtected = BeginEditableSheetForTest(ws)
    Do While Not lo.DataBodyRange Is Nothing
        lo.ListRows(1).Delete
    Loop

    For i = LBound(skuList) To UBound(skuList)
        Set lr = lo.ListRows.Add
        SetTableCell lo, lr.Index, "SKU", CStr(skuList(i))
        SetTableCell lo, lr.Index, "ITEM_CODE", CStr(skuList(i))
        SetTableCell lo, lr.Index, "ITEM", CStr(skuList(i))
    Next i

    RestoreSheetProtectionForTest ws, sheetWasProtected
End Sub

Private Function CreateReceiveEventForTest(ByVal eventId As String, _
                                           ByVal warehouseId As String, _
                                           ByVal stationId As String, _
                                           ByVal userId As String, _
                                           ByVal sku As String, _
                                           ByVal qty As Double, _
                                           ByVal locationVal As String, _
                                           ByVal noteVal As String) As Object
    Dim evt As Object

    Set evt = CreateObject("Scripting.Dictionary")
    evt.CompareMode = vbTextCompare
    evt("EventID") = eventId
    evt("EventType") = "RECEIVE"
    evt("CreatedAtUTC") = Now
    evt("WarehouseId") = warehouseId
    evt("StationId") = stationId
    evt("UserId") = userId
    evt("SourceInbox") = "phase6-test-inbox"
    evt("SKU") = sku
    evt("Qty") = qty
    evt("Location") = locationVal
    evt("Note") = noteVal
    Set CreateReceiveEventForTest = evt
End Function

Private Sub DeleteTableSurfaceForTest(ByVal ws As Worksheet, ByVal tableName As String)
    Dim lo As ListObject

    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    ws.Unprotect
    lo.Delete
    ws.Cells.Clear
End Sub

Private Function NzDblForTest(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzDblForTest = CDbl(valueIn)
End Function

Private Function BeginEditableSheetForTest(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    BeginEditableSheetForTest = ws.ProtectContents
    If Not BeginEditableSheetForTest Then Exit Function

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 7103, "TestPhase6CoreSurfaces.BeginEditableSheetForTest", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected for test data setup."
    End If
End Function

Private Sub RestoreSheetProtectionForTest(ByVal ws As Worksheet, ByVal wasProtected As Boolean)
    If ws Is Nothing Then Exit Sub
    If Not wasProtected Then Exit Sub

    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0

    If Not ws.ProtectContents Then
        Err.Raise vbObjectError + 7104, "TestPhase6CoreSurfaces.RestoreSheetProtectionForTest", _
                  "Worksheet '" & ws.Name & "' could not be reprotected after test data setup."
    End If
End Sub

Private Function CreateSnapshotWorkbook(ByVal rootPath As String, _
                                        ByVal warehouseId As String, _
                                        ByVal sku As String, _
                                        ByVal qtyOnHand As Double, _
                                        ByVal lastAppliedUtc As Date, _
                                        Optional ByVal qtyAvailable As Variant, _
                                        Optional ByVal locationSummary As Variant, _
                                        Optional ByVal itemName As String = vbNullString, _
                                        Optional ByVal uom As String = vbNullString, _
                                        Optional ByVal locationVal As String = vbNullString, _
                                        Optional ByVal description As String = vbNullString, _
                                        Optional ByVal vendorName As String = vbNullString, _
                                        Optional ByVal vendorCode As String = vbNullString, _
                                        Optional ByVal category As String = vbNullString) As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim targetPath As String
    Dim resolvedQtyAvailable As Double
    Dim resolvedLocationSummary As String

    targetPath = rootPath & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "InventorySnapshot"
    ws.Range("A1").Value = "WarehouseId"
    ws.Range("B1").Value = "SKU"
    ws.Range("C1").Value = "ITEM"
    ws.Range("D1").Value = "UOM"
    ws.Range("E1").Value = "LOCATION"
    ws.Range("F1").Value = "DESCRIPTION"
    ws.Range("G1").Value = "VENDOR(s)"
    ws.Range("H1").Value = "VENDOR_CODE"
    ws.Range("I1").Value = "CATEGORY"
    ws.Range("J1").Value = "QtyOnHand"
    ws.Range("K1").Value = "QtyAvailable"
    ws.Range("L1").Value = "LocationSummary"
    ws.Range("M1").Value = "LastAppliedAtUTC"
    ws.Range("A2").Value = warehouseId
    ws.Range("B2").Value = sku
    If Trim$(itemName) = "" Then itemName = sku
    ws.Range("C2").Value = itemName
    ws.Range("D2").Value = uom
    ws.Range("E2").Value = locationVal
    ws.Range("F2").Value = description
    ws.Range("G2").Value = vendorName
    ws.Range("H2").Value = vendorCode
    ws.Range("I2").Value = category
    ws.Range("J2").Value = qtyOnHand
    If IsMissing(qtyAvailable) Or IsEmpty(qtyAvailable) Then
        resolvedQtyAvailable = qtyOnHand
    Else
        resolvedQtyAvailable = CDbl(qtyAvailable)
    End If
    ws.Range("K2").Value = resolvedQtyAvailable
    If IsMissing(locationSummary) Or IsEmpty(locationSummary) Then
        resolvedLocationSummary = "A1=" & CStr(CLng(qtyOnHand))
    Else
        resolvedLocationSummary = CStr(locationSummary)
    End If
    ws.Range("L2").Value = resolvedLocationSummary
    ws.Range("M2").Value = lastAppliedUtc
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:M2"), , xlYes)
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

Private Function NormalizeTestPath(ByVal pathText As String) As String
    pathText = Trim$(Replace$(pathText, "/", "\"))
    Do While Len(pathText) > 3 And Right$(pathText, 1) = "\"
        pathText = Left$(pathText, Len(pathText) - 1)
    Loop
    NormalizeTestPath = pathText
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
