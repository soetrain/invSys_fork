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
