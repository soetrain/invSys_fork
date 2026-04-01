Attribute VB_Name = "TestWarehouseBootstrap"
Option Explicit

Public Function TestValidateWarehouseSpec_TrimsFieldsAndAllowsBlankSharePoint() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim report As String

    spec.WarehouseId = "  WH2  "
    spec.WarehouseName = "  Warehouse Two  "
    spec.StationId = "  S1  "
    spec.AdminUser = "  justinwj  "
    spec.PathLocal = "  C:\invSys\WH2  "
    spec.PathSharePoint = "   "

    On Error GoTo CleanFail
    If Not modWarehouseBootstrap.ValidateWarehouseSpec(spec, report) Then GoTo CleanExit

    If StrComp(spec.WarehouseId, "WH2", vbTextCompare) = 0 _
       And StrComp(spec.WarehouseName, "Warehouse Two", vbTextCompare) = 0 _
       And StrComp(spec.StationId, "S1", vbTextCompare) = 0 _
       And StrComp(spec.AdminUser, "justinwj", vbTextCompare) = 0 _
       And StrComp(spec.PathLocal, "C:\invSys\WH2", vbTextCompare) = 0 _
       And spec.PathSharePoint = "" _
       And StrComp(report, "OK", vbTextCompare) = 0 Then
        TestValidateWarehouseSpec_TrimsFieldsAndAllowsBlankSharePoint = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateWarehouseSpec_RejectsEmptyWarehouseId() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim report As String

    spec.WarehouseId = "   "
    spec.WarehouseName = "Warehouse Two"
    spec.StationId = "S1"
    spec.AdminUser = "justinwj"

    On Error GoTo CleanFail
    If modWarehouseBootstrap.ValidateWarehouseSpec(spec, report) Then GoTo CleanExit

    If StrComp(spec.WarehouseId, "", vbTextCompare) = 0 _
       And InStr(1, report, "WarehouseId is required", vbTextCompare) > 0 Then
        TestValidateWarehouseSpec_RejectsEmptyWarehouseId = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateWarehouseSpec_RejectsWarehouseIdWithSpaces() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim report As String

    spec.WarehouseId = "WH 2"

    On Error GoTo CleanFail
    If modWarehouseBootstrap.ValidateWarehouseSpec(spec, report) Then GoTo CleanExit

    If InStr(1, report, "letters, digits, hyphens, and underscores", vbTextCompare) > 0 Then
        TestValidateWarehouseSpec_RejectsWarehouseIdWithSpaces = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateWarehouseSpec_AllowsWarehouseIdWithHyphenAndUnderscore() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim report As String

    spec.WarehouseId = "WH_2-A"

    On Error GoTo CleanFail
    If Not modWarehouseBootstrap.ValidateWarehouseSpec(spec, report) Then GoTo CleanExit

    If StrComp(spec.WarehouseId, "WH_2-A", vbTextCompare) = 0 _
       And StrComp(report, "OK", vbTextCompare) = 0 Then
        TestValidateWarehouseSpec_AllowsWarehouseIdWithHyphenAndUnderscore = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateWarehouseSpec_RejectsWarehouseIdWithOtherSpecialCharacters() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim report As String

    spec.WarehouseId = "WH.2"

    On Error GoTo CleanFail
    If modWarehouseBootstrap.ValidateWarehouseSpec(spec, report) Then GoTo CleanExit

    If InStr(1, report, "letters, digits, hyphens, and underscores", vbTextCompare) > 0 Then
        TestValidateWarehouseSpec_RejectsWarehouseIdWithOtherSpecialCharacters = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestWarehouseIdExists_LocalFolderExists() As Long
    Dim warehouseId As String
    Dim localPath As String

    warehouseId = "WHBOOTLOCAL01"
    localPath = "C:\invSys\" & warehouseId

    On Error GoTo CleanFail
    EnsureFolderBootstrapTest localPath

    If modWarehouseBootstrap.WarehouseIdExists(warehouseId) Then
        TestWarehouseIdExists_LocalFolderExists = 1
    End If

CleanExit:
    RemoveFolderIfEmptyBootstrapTest localPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestWarehouseIdExists_SharePointArtifactExists() As Long
    Dim warehouseId As String
    Dim rootPath As String
    Dim shareRoot As String
    Dim artifactPath As String

    warehouseId = "WHBOOTSP01"
    rootPath = BuildBootstrapTempRoot("cfg_sp_exists")
    shareRoot = BuildBootstrapTempRoot("share_sp_exists")
    artifactPath = shareRoot & "\Config\" & warehouseId & ".invSys.Config.xlsb"

    On Error GoTo CleanFail
    EnsureFolderBootstrapTest shareRoot & "\Config"
    WriteTextFileBootstrapTest artifactPath, "bootstrap-sharepoint-artifact"
    If Not LoadWarehouseBootstrapConfigTest(rootPath, warehouseId, shareRoot) Then GoTo CleanExit

    If modWarehouseBootstrap.WarehouseIdExists(warehouseId) Then
        TestWarehouseIdExists_SharePointArtifactExists = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    DeleteFolderRecursiveBootstrapTest shareRoot
    DeleteFolderRecursiveBootstrapTest rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestWarehouseIdExists_NeitherLocalNorSharePointExists() As Long
    Dim warehouseId As String
    Dim rootPath As String
    Dim shareRoot As String

    warehouseId = "WHBOOTNONE01"
    rootPath = BuildBootstrapTempRoot("cfg_none_exists")
    shareRoot = BuildBootstrapTempRoot("share_none_exists")

    On Error GoTo CleanFail
    EnsureFolderBootstrapTest shareRoot & "\Config"
    If Not LoadWarehouseBootstrapConfigTest(rootPath, warehouseId, shareRoot) Then GoTo CleanExit

    If Not modWarehouseBootstrap.WarehouseIdExists(warehouseId) Then
        TestWarehouseIdExists_NeitherLocalNorSharePointExists = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    DeleteFolderRecursiveBootstrapTest shareRoot
    DeleteFolderRecursiveBootstrapTest rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestWarehouseIdExists_SharePointUnavailableReturnsFalseAndLogsSkip() As Long
    Dim warehouseId As String
    Dim rootPath As String
    Dim shareRoot As String
    Dim logTextBefore As String
    Dim logTextAfter As String

    warehouseId = "WHBOOTSKIP01"
    rootPath = BuildBootstrapTempRoot("cfg_share_skip")
    shareRoot = "C:\Invalid<SharePointRoot"

    On Error GoTo CleanFail
    logTextBefore = ReadBootstrapPerfLogText()
    If Not LoadWarehouseBootstrapConfigTest(rootPath, warehouseId, shareRoot) Then GoTo CleanExit

    If modWarehouseBootstrap.WarehouseIdExists(warehouseId) Then GoTo CleanExit

    logTextAfter = ReadBootstrapPerfLogText()
    If InStr(1, logTextAfter, "SharePoint collision check skipped|WarehouseId=" & warehouseId, vbTextCompare) > 0 _
       And InStr(1, logTextAfter, shareRoot, vbTextCompare) > 0 _
       And (Len(logTextAfter) >= Len(logTextBefore)) Then
        TestWarehouseIdExists_SharePointUnavailableReturnsFalseAndLogsSkip = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    DeleteFolderRecursiveBootstrapTest rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestBootstrapWarehouseLocal_CreatesBootableLocalRuntime() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim rootPath As String
    Dim templateRoot As String
    Dim wbCfg As Workbook
    Dim report As String
    Dim loWh As ListObject
    Dim loSt As ListObject

    rootPath = BuildBootstrapTempRoot("warehouse_local_ok")
    templateRoot = BuildBootstrapTempRoot("warehouse_templates_ok")

    spec.WarehouseId = "WHBOOT-LOCAL_01"
    spec.WarehouseName = "Bootstrap Warehouse"
    spec.StationId = "ADM1"
    spec.AdminUser = "admin.bootstrap"
    spec.PathLocal = rootPath
    spec.PathSharePoint = "C:\ShareRoot\invSys"

    On Error GoTo CleanFail
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot
    If Not modWarehouseBootstrap.BootstrapWarehouseLocal(spec) Then GoTo CleanExit

    If Len(Dir$(rootPath, vbDirectory)) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\inbox", vbDirectory)) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\outbox", vbDirectory)) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\snapshots", vbDirectory)) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\config", vbDirectory)) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb")) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb")) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\" & spec.WarehouseId & ".invSys.Auth.xlsb")) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\" & spec.WarehouseId & ".Outbox.Events.xlsb")) = 0 Then GoTo CleanExit
    If Len(Dir$(rootPath & "\" & spec.WarehouseId & ".invSys.Snapshot.Inventory.xlsb")) = 0 Then GoTo CleanExit

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig(spec.WarehouseId, spec.StationId) Then GoTo CleanExit
    If Not modAuth.LoadAuth(spec.WarehouseId) Then GoTo CleanExit
    If Not modAuth.CanPerform("ADMIN_MAINT", spec.AdminUser, spec.WarehouseId, spec.StationId, "TEST", "BOOTSTRAP-TEST") Then GoTo CleanExit

    Set wbCfg = Application.Workbooks.Open(rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb")
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")

    If StrComp(CStr(GetBootstrapTableValue(loWh, 1, "WarehouseId")), spec.WarehouseId, vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetBootstrapTableValue(loWh, 1, "WarehouseName")), spec.WarehouseName, vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetBootstrapTableValue(loWh, 1, "PathDataRoot")), spec.PathLocal, vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetBootstrapTableValue(loWh, 1, "PathSharePointRoot")), spec.PathSharePoint, vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetBootstrapTableValue(loSt, 1, "StationId")), spec.StationId, vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetBootstrapTableValue(loSt, 1, "StationName")), spec.AdminUser, vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(GetBootstrapTableValue(loSt, 1, "RoleDefault")), "ADMIN", vbTextCompare) <> 0 Then GoTo CleanExit

    TestBootstrapWarehouseLocal_CreatesBootableLocalRuntime = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    CloseNoSaveBootstrapTest wbCfg
    DeleteFolderRecursiveBootstrapTest rootPath
    DeleteFolderRecursiveBootstrapTest templateRoot
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestBootstrapWarehouseLocal_FailureRollsBackPartialFolders() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim rootPath As String
    Dim templateRoot As String

    rootPath = BuildBootstrapTempRoot("warehouse_local_fail")
    templateRoot = BuildBootstrapTempRoot("warehouse_templates_fail")

    spec.WarehouseId = "WHBOOT-FAIL_01"
    spec.WarehouseName = "Bootstrap Failure"
    spec.StationId = "ADM1"
    spec.AdminUser = ""
    spec.PathLocal = rootPath
    spec.PathSharePoint = ""

    On Error GoTo CleanFail
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot
    If modWarehouseBootstrap.BootstrapWarehouseLocal(spec) Then GoTo CleanExit

    If Len(Dir$(rootPath, vbDirectory)) = 0 _
       And InStr(1, modWarehouseBootstrap.GetLastWarehouseBootstrapReport(), "AdminUser is required", vbTextCompare) > 0 Then
        TestBootstrapWarehouseLocal_FailureRollsBackPartialFolders = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteFolderRecursiveBootstrapTest rootPath
    DeleteFolderRecursiveBootstrapTest templateRoot
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestPublishInitialArtifacts_PublishSuccess() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim rootPath As String
    Dim templateRoot As String
    Dim shareRoot As String
    Dim publishedConfigPath As String
    Dim discoveryPath As String
    Dim discoveryText As String

    rootPath = BuildBootstrapTempRoot("warehouse_publish_ok")
    templateRoot = BuildBootstrapTempRoot("warehouse_templates_publish_ok")
    shareRoot = BuildBootstrapTempRoot("warehouse_share_publish_ok")

    spec.WarehouseId = "WHBOOT-PUBLISH_01"
    spec.WarehouseName = "Publish Warehouse"
    spec.StationId = "ADM1"
    spec.AdminUser = "admin.publish"
    spec.PathLocal = rootPath
    spec.PathSharePoint = shareRoot

    On Error GoTo CleanFail
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot
    If Not modWarehouseBootstrap.BootstrapWarehouseLocal(spec) Then GoTo CleanExit
    If Not modWarehouseBootstrap.PublishInitialArtifacts(spec) Then GoTo CleanExit

    publishedConfigPath = shareRoot & "\" & spec.WarehouseId & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    discoveryPath = shareRoot & "\" & spec.WarehouseId & ".config.json"
    discoveryText = ReadTextFileBootstrapTest(discoveryPath)

    If Len(Dir$(publishedConfigPath)) = 0 Then GoTo CleanExit
    If Len(Dir$(discoveryPath)) = 0 Then GoTo CleanExit
    If InStr(1, discoveryText, """" & spec.WarehouseId & """", vbTextCompare) = 0 Then GoTo CleanExit
    If InStr(1, modWarehouseBootstrap.GetLastWarehouseBootstrapReport(), "OK|Config=", vbTextCompare) = 0 Then GoTo CleanExit

    TestPublishInitialArtifacts_PublishSuccess = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteFolderRecursiveBootstrapTest shareRoot
    DeleteFolderRecursiveBootstrapTest rootPath
    DeleteFolderRecursiveBootstrapTest templateRoot
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestPublishInitialArtifacts_SharePointUnavailableReturnsFalse() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim rootPath As String
    Dim templateRoot As String
    Dim localConfigPath As String
    Dim logTextBefore As String
    Dim logTextAfter As String

    rootPath = BuildBootstrapTempRoot("warehouse_publish_offline")
    templateRoot = BuildBootstrapTempRoot("warehouse_templates_publish_offline")

    spec.WarehouseId = "WHBOOT-PUBLISH_02"
    spec.WarehouseName = "Publish Offline"
    spec.StationId = "ADM1"
    spec.AdminUser = "admin.offline"
    spec.PathLocal = rootPath
    spec.PathSharePoint = "C:\Invalid<SharePointRoot"

    On Error GoTo CleanFail
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot
    If Not modWarehouseBootstrap.BootstrapWarehouseLocal(spec) Then GoTo CleanExit

    logTextBefore = ReadBootstrapPerfLogText()
    If modWarehouseBootstrap.PublishInitialArtifacts(spec) Then GoTo CleanExit

    localConfigPath = rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    logTextAfter = ReadBootstrapPerfLogText()

    If Len(Dir$(localConfigPath)) = 0 Then GoTo CleanExit
    If InStr(1, modWarehouseBootstrap.GetLastWarehouseBootstrapReport(), "FAILED:", vbTextCompare) = 0 Then GoTo CleanExit
    If InStr(1, logTextAfter, "Initial publish failed|WarehouseId=" & spec.WarehouseId, vbTextCompare) = 0 Then GoTo CleanExit
    If Len(logTextAfter) < Len(logTextBefore) Then GoTo CleanExit

    TestPublishInitialArtifacts_SharePointUnavailableReturnsFalse = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteFolderRecursiveBootstrapTest rootPath
    DeleteFolderRecursiveBootstrapTest templateRoot
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestPublishInitialArtifacts_RepeatedPublishIsIdempotent() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim rootPath As String
    Dim templateRoot As String
    Dim shareRoot As String
    Dim wbLocalCfg As Workbook
    Dim wbPublishedCfg As Workbook
    Dim loWh As ListObject
    Dim report As String
    Dim publishedConfigPath As String

    rootPath = BuildBootstrapTempRoot("warehouse_publish_rerun")
    templateRoot = BuildBootstrapTempRoot("warehouse_templates_publish_rerun")
    shareRoot = BuildBootstrapTempRoot("warehouse_share_publish_rerun")

    spec.WarehouseId = "WHBOOT-PUBLISH_03"
    spec.WarehouseName = "Publish Rerun"
    spec.StationId = "ADM1"
    spec.AdminUser = "admin.rerun"
    spec.PathLocal = rootPath
    spec.PathSharePoint = shareRoot

    On Error GoTo CleanFail
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot
    If Not modWarehouseBootstrap.BootstrapWarehouseLocal(spec) Then GoTo CleanExit
    If Not modWarehouseBootstrap.PublishInitialArtifacts(spec) Then GoTo CleanExit

    Set wbLocalCfg = Application.Workbooks.Open(rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb")
    Set loWh = wbLocalCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    loWh.DataBodyRange.Cells(1, loWh.ListColumns("WarehouseName").Index).Value = "Publish Rerun v2"
    wbLocalCfg.Save
    CloseNoSaveBootstrapTest wbLocalCfg
    Set wbLocalCfg = Nothing

    If Not modWarehouseBootstrap.PublishInitialArtifacts(spec) Then GoTo CleanExit

    publishedConfigPath = shareRoot & "\" & spec.WarehouseId & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    Set wbPublishedCfg = Application.Workbooks.Open(publishedConfigPath)
    Set loWh = wbPublishedCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")

    If StrComp(CStr(GetBootstrapTableValue(loWh, 1, "WarehouseName")), "Publish Rerun v2", vbTextCompare) <> 0 Then GoTo CleanExit
    If InStr(1, modWarehouseBootstrap.GetLastWarehouseBootstrapReport(), "Discovery=REPLACED:", vbTextCompare) = 0 _
       And InStr(1, modWarehouseBootstrap.GetLastWarehouseBootstrapReport(), "Discovery=COPIED:", vbTextCompare) = 0 Then GoTo CleanExit

    TestPublishInitialArtifacts_RepeatedPublishIsIdempotent = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    CloseNoSaveBootstrapTest wbLocalCfg
    CloseNoSaveBootstrapTest wbPublishedCfg
    DeleteFolderRecursiveBootstrapTest shareRoot
    DeleteFolderRecursiveBootstrapTest rootPath
    DeleteFolderRecursiveBootstrapTest templateRoot
    Exit Function
CleanFail:
    report = Err.Description
    Resume CleanExit
End Function

Private Function LoadWarehouseBootstrapConfigTest(ByVal rootPath As String, _
                                                  ByVal warehouseId As String, _
                                                  ByVal shareRoot As String) As Boolean
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim report As String

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime(warehouseId, "S1", rootPath, report)
    If wbCfg Is Nothing Then GoTo CleanExit

    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    loWh.DataBodyRange.Cells(1, loWh.ListColumns("PathSharePointRoot").Index).Value = shareRoot
    wbCfg.Save

    LoadWarehouseBootstrapConfigTest = modConfig.LoadConfig(warehouseId, "S1")

CleanExit:
    CloseNoSaveBootstrapTest wbCfg
End Function

Private Function BuildBootstrapTempRoot(ByVal leafName As String) As String
    BuildBootstrapTempRoot = Environ$("TEMP") & "\invSys_" & leafName & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 1000))
End Function

Private Sub EnsureFolderBootstrapTest(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 1 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderBootstrapTest parentPath
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Sub RemoveFolderIfEmptyBootstrapTest(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then RmDir folderPath
    On Error GoTo 0
End Sub

Private Sub DeleteFolderRecursiveBootstrapTest(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

Private Sub WriteTextFileBootstrapTest(ByVal filePath As String, ByVal contents As String)
    Dim fileNum As Integer
    Dim parentPath As String

    parentPath = Left$(filePath, InStrRev(filePath, "\") - 1)
    EnsureFolderBootstrapTest parentPath

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, contents
    Close #fileNum
End Sub

Private Function ReadBootstrapPerfLogText() As String
    Dim logPath As String
    Dim fileNum As Integer

    logPath = Environ$("TEMP")
    If Right$(logPath, 1) <> "\" Then logPath = logPath & "\"
    logPath = logPath & "invSys.Inventory.Sync.log"
    If Len(Dir$(logPath)) = 0 Then Exit Function

    fileNum = FreeFile
    Open logPath For Input As #fileNum
    ReadBootstrapPerfLogText = Input$(LOF(fileNum), #fileNum)
    Close #fileNum
End Function

Private Function ReadTextFileBootstrapTest(ByVal filePath As String) As String
    Dim fileNum As Integer

    If Len(Dir$(filePath)) = 0 Then Exit Function
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    ReadTextFileBootstrapTest = Input$(LOF(fileNum), #fileNum)
    Close #fileNum
End Function

Private Sub CloseNoSaveBootstrapTest(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function GetBootstrapTableValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long

    If lo Is Nothing Then Exit Function
    idx = lo.ListColumns(columnName).Index
    GetBootstrapTableValue = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function
