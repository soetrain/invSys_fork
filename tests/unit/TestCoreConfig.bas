Attribute VB_Name = "TestCoreConfig"
Option Explicit

Public Sub RunConfigTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestLoad_ValidConfig(), passed, failed
    Tally TestLoad_MissingRequiredKey(), passed, failed
    Tally TestPrecedence_StationOverridesWarehouse(), passed, failed
    Tally TestGetRequired_MissingKey(), passed, failed
    Tally TestGetBool_TypeConversion(), passed, failed
    Tally TestReload_UpdatedValue(), passed, failed

    Debug.Print "Core.Config tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestLoad_ValidConfig() As Long
    Dim wb As Workbook
    Set wb = BuildConfigWorkbook("WH1", "S1", "RECEIVE")

    On Error GoTo CleanFail
    If modConfig.LoadConfig("WH1", "S1") Then
        If modConfig.GetLong("BatchSize", 0) = 500 Then
            TestLoad_ValidConfig = 1
        End If
    End If

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLoad_MissingRequiredKey() As Long
    Dim wb As Workbook
    Dim lo As ListObject
    Set wb = BuildConfigWorkbook("WH1", "S1", "RECEIVE")
    Set lo = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    lo.DataBodyRange.Cells(1, lo.ListColumns("WarehouseId").Index).Value = ""

    On Error GoTo CleanFail
    If Not modConfig.LoadConfig("WH1", "S1") Then
        TestLoad_MissingRequiredKey = 1
    End If

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestPrecedence_StationOverridesWarehouse() As Long
    Dim wb As Workbook
    Dim loWh As ListObject
    Dim loSt As ListObject

    Set wb = BuildConfigWorkbook("WH1", "S1", "SHIP")
    Set loWh = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wb.Worksheets("StationConfig").ListObjects("tblStationConfig")

    loWh.DataBodyRange.Cells(1, loWh.ListColumns("RoleDefault").Index).Value = "RECEIVE"
    loSt.DataBodyRange.Cells(1, loSt.ListColumns("RoleDefault").Index).Value = "SHIP"

    On Error GoTo CleanFail
    If modConfig.LoadConfig("WH1", "S1") Then
        If modConfig.GetString("RoleDefault", "") = "SHIP" Then
            TestPrecedence_StationOverridesWarehouse = 1
        End If
    End If

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestGetRequired_MissingKey() As Long
    Dim wb As Workbook
    Dim didRaise As Boolean
    Dim x As Variant

    Set wb = BuildConfigWorkbook("WH1", "S1", "RECEIVE")

    On Error GoTo CleanFail
    Call modConfig.LoadConfig("WH1", "S1")

    On Error Resume Next
    x = modConfig.GetRequired("NoSuchKey")
    didRaise = (Err.Number <> 0)
    Err.Clear
    On Error GoTo CleanFail

    If didRaise Then TestGetRequired_MissingKey = 1

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestGetBool_TypeConversion() As Long
    Dim wb As Workbook
    Dim loWh As ListObject

    Set wb = BuildConfigWorkbook("WH1", "S1", "RECEIVE")
    Set loWh = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    loWh.DataBodyRange.Cells(1, loWh.ListColumns("DesignsEnabled").Index).Value = "TRUE"

    On Error GoTo CleanFail
    If modConfig.LoadConfig("WH1", "S1") Then
        If modConfig.GetBool("DesignsEnabled", False) = True Then
            TestGetBool_TypeConversion = 1
        End If
    End If

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReload_UpdatedValue() As Long
    Dim wb As Workbook
    Dim loWh As ListObject

    Set wb = BuildConfigWorkbook("WH1", "S1", "RECEIVE")
    Set loWh = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")

    On Error GoTo CleanFail
    If Not modConfig.LoadConfig("WH1", "S1") Then GoTo CleanExit

    loWh.DataBodyRange.Cells(1, loWh.ListColumns("BatchSize").Index).Value = 250
    If modConfig.Reload() Then
        If modConfig.GetLong("BatchSize", 0) = 250 Then
            TestReload_UpdatedValue = 1
        End If
    End If

CleanExit:
    CloseNoSave wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Sub Tally(ByVal testResult As Long, ByRef passed As Long, ByRef failed As Long)
    If testResult = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub

Private Function BuildConfigWorkbook(ByVal whId As String, ByVal stId As String, ByVal roleDefault As String) As Workbook
    Dim wb As Workbook
    Dim wsWh As Worksheet
    Dim wsSt As Worksheet
    Dim loWh As ListObject
    Dim loSt As ListObject
    Dim p As String

    Set wb = Application.Workbooks.Add
    Set wsWh = wb.Worksheets(1)
    wsWh.Name = "WarehouseConfig"
    Set wsSt = wb.Worksheets.Add(After:=wsWh)
    wsSt.Name = "StationConfig"

    wsWh.Range("A1").Resize(1, 20).Value = Array( _
        "WarehouseId", "WarehouseName", "Timezone", "DefaultLocation", _
        "BatchSize", "LockTimeoutMinutes", "HeartbeatIntervalSeconds", "MaxLockHoldMinutes", _
        "SnapshotCadence", "BackupCadence", "PathDataRoot", "PathBackupRoot", "PathSharePointRoot", _
        "DesignsEnabled", "PoisonRetryMax", "AuthCacheTTLSeconds", _
        "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot", "RoleDefault")
    wsWh.Range("A2").Resize(1, 20).Value = Array( _
        whId, "Main Warehouse", "UTC", "A1", _
        500, 3, 30, 2, _
        "PER_BATCH", "DAILY", "C:\invSys\" & whId & "\", "C:\invSys\Backups\" & whId & "\", "", _
        False, 3, 300, _
        False, False, True, "RECEIVE")

    Set loWh = wsWh.ListObjects.Add(xlSrcRange, wsWh.Range("A1:T2"), , xlYes)
    loWh.Name = "tblWarehouseConfig"

    wsSt.Range("A1").Resize(1, 4).Value = Array("StationId", "WarehouseId", "StationName", "RoleDefault")
    wsSt.Range("A2").Resize(1, 4).Value = Array(stId, whId, Environ$("COMPUTERNAME"), roleDefault)
    Set loSt = wsSt.ListObjects.Add(xlSrcRange, wsSt.Range("A1:D2"), , xlYes)
    loSt.Name = "tblStationConfig"

    p = BuildUniqueConfigTestPath()
    wb.SaveAs Filename:=p, FileFormat:=51

    Set BuildConfigWorkbook = wb
End Function

Private Function BuildUniqueConfigTestPath() As String
    Dim stamp As String

    stamp = Format$(Now, "yyyymmdd_hhnnss") & "_" & Right$("000000" & CStr(CLng((Timer - Int(Timer)) * 1000000)), 6)
    BuildUniqueConfigTestPath = Environ$("TEMP") & "\WH1.invSys.Config.test." & stamp & ".xlsx"
End Function

Private Sub CloseNoSave(ByVal wb As Workbook)
    Dim p As String
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    p = wb.FullName
    wb.Close SaveChanges:=False
    If InStr(1, p, ".test.", vbTextCompare) > 0 Then
        If Len(Dir$(p)) > 0 Then Kill p
    End If
    On Error GoTo 0
End Sub
