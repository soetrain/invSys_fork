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
