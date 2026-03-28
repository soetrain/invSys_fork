Attribute VB_Name = "TestCoreAuth"
Option Explicit

Public Sub RunAuthTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestCanPerform_Allow(), passed, failed
    Tally TestCanPerform_Deny_MissingCapability(), passed, failed
    Tally TestCanPerform_WildcardStation(), passed, failed
    Tally TestCanPerform_DisabledUser(), passed, failed
    Tally TestCanPerform_ExpiredCapability(), passed, failed
    Tally TestRequire_RaisesOnDeny(), passed, failed

    Debug.Print "Core.Auth tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestCanPerform_Allow() As Long
    Dim wbAuth As Workbook
    Dim wbCfg As Workbook

    Set wbCfg = BuildConfigWorkbook("WH1", "S1")
    Set wbAuth = BuildAuthWorkbook("WH1")

    AddCapability wbAuth, "user1", "RECEIVE_POST", "WH1", "S1", "ACTIVE", "", ""

    On Error GoTo CleanFail
    Call modConfig.LoadConfig("WH1", "S1")
    If modAuth.LoadAuth("WH1") Then
        If modAuth.CanPerform("RECEIVE_POST", "user1", "WH1", "S1", "TEST", "REQ-1") Then
            TestCanPerform_Allow = 1
        End If
    End If

CleanExit:
    CloseNoSave wbAuth
    CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCanPerform_Deny_MissingCapability() As Long
    Dim wbAuth As Workbook
    Dim wbCfg As Workbook

    Set wbCfg = BuildConfigWorkbook("WH1", "S1")
    Set wbAuth = BuildAuthWorkbook("WH1")
    AddCapability wbAuth, "user1", "RECEIVE_POST", "WH1", "S1", "ACTIVE", "", ""

    On Error GoTo CleanFail
    Call modConfig.LoadConfig("WH1", "S1")
    Call modAuth.LoadAuth("WH1")

    If Not modAuth.CanPerform("SHIP_POST", "user1", "WH1", "S1", "TEST", "REQ-2") Then
        TestCanPerform_Deny_MissingCapability = 1
    End If

CleanExit:
    CloseNoSave wbAuth
    CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCanPerform_WildcardStation() As Long
    Dim wbAuth As Workbook
    Dim wbCfg As Workbook

    Set wbCfg = BuildConfigWorkbook("WH1", "S1")
    Set wbAuth = BuildAuthWorkbook("WH1")
    AddCapability wbAuth, "user1", "INBOX_PROCESS", "WH1", "*", "ACTIVE", "", ""

    On Error GoTo CleanFail
    Call modConfig.LoadConfig("WH1", "S1")
    Call modAuth.LoadAuth("WH1")

    If modAuth.CanPerform("INBOX_PROCESS", "user1", "WH1", "S2", "TEST", "REQ-3") Then
        TestCanPerform_WildcardStation = 1
    End If

CleanExit:
    CloseNoSave wbAuth
    CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCanPerform_DisabledUser() As Long
    Dim wbAuth As Workbook
    Dim wbCfg As Workbook
    Dim loUsers As ListObject

    Set wbCfg = BuildConfigWorkbook("WH1", "S1")
    Set wbAuth = BuildAuthWorkbook("WH1")
    AddCapability wbAuth, "user2", "RECEIVE_POST", "WH1", "S1", "ACTIVE", "", ""

    Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    loUsers.DataBodyRange.Cells(2, loUsers.ListColumns("Status").Index).Value = "Disabled"

    On Error GoTo CleanFail
    Call modConfig.LoadConfig("WH1", "S1")
    Call modAuth.LoadAuth("WH1")

    If Not modAuth.CanPerform("RECEIVE_POST", "user2", "WH1", "S1", "TEST", "REQ-4") Then
        TestCanPerform_DisabledUser = 1
    End If

CleanExit:
    CloseNoSave wbAuth
    CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCanPerform_ExpiredCapability() As Long
    Dim wbAuth As Workbook
    Dim wbCfg As Workbook

    Set wbCfg = BuildConfigWorkbook("WH1", "S1")
    Set wbAuth = BuildAuthWorkbook("WH1")
    AddCapability wbAuth, "user1", "RECEIVE_POST", "WH1", "S1", "ACTIVE", "", "2000-01-01"

    On Error GoTo CleanFail
    Call modConfig.LoadConfig("WH1", "S1")
    Call modAuth.LoadAuth("WH1")

    If Not modAuth.CanPerform("RECEIVE_POST", "user1", "WH1", "S1", "TEST", "REQ-5") Then
        TestCanPerform_ExpiredCapability = 1
    End If

CleanExit:
    CloseNoSave wbAuth
    CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRequire_RaisesOnDeny() As Long
    Dim wbAuth As Workbook
    Dim wbCfg As Workbook
    Dim raised As Boolean

    Set wbCfg = BuildConfigWorkbook("WH1", "S1")
    Set wbAuth = BuildAuthWorkbook("WH1")
    AddCapability wbAuth, "user1", "RECEIVE_POST", "WH1", "S1", "ACTIVE", "", ""

    On Error GoTo CleanFail
    Call modConfig.LoadConfig("WH1", "S1")
    Call modAuth.LoadAuth("WH1")

    On Error Resume Next
    Call modAuth.Require("SHIP_POST", "user1", "WH1", "S1", "TEST", "REQ-6")
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo CleanFail

    If raised Then TestRequire_RaisesOnDeny = 1

CleanExit:
    CloseNoSave wbAuth
    CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function BuildConfigWorkbook(ByVal whId As String, ByVal stId As String) As Workbook
    Dim wb As Workbook
    Dim wsWh As Worksheet
    Dim wsSt As Worksheet
    Dim p As String

    Set wb = Application.Workbooks.Add
    Set wsWh = wb.Worksheets(1)
    wsWh.Name = "WarehouseConfig"
    Set wsSt = wb.Worksheets.Add(After:=wsWh)
    wsSt.Name = "StationConfig"

    wsWh.Range("A1").Resize(1, 21).Value = Array( _
        "WarehouseId", "WarehouseName", "Timezone", "DefaultLocation", _
        "BatchSize", "LockTimeoutMinutes", "HeartbeatIntervalSeconds", "MaxLockHoldMinutes", _
        "SnapshotCadence", "BackupCadence", "PathDataRoot", "PathBackupRoot", "PathSharePointRoot", _
        "DesignsEnabled", "PoisonRetryMax", "AuthCacheTTLSeconds", "ProcessorServiceUserId", _
        "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot", "AutoRefreshIntervalSeconds")
    wsWh.Range("A2").Resize(1, 21).Value = Array( _
        whId, "Main Warehouse", "UTC", "A1", _
        500, 3, 30, 2, _
        "PER_BATCH", "DAILY", "C:\invSys\" & whId & "\", "C:\invSys\Backups\" & whId & "\", "", _
        False, 3, 300, "svc_processor", _
        False, False, True, 0)
    wsWh.ListObjects.Add(xlSrcRange, wsWh.Range("A1:U2"), , xlYes).Name = "tblWarehouseConfig"

    wsSt.Range("A1").Resize(1, 4).Value = Array("StationId", "WarehouseId", "StationName", "RoleDefault")
    wsSt.Range("A2").Resize(1, 4).Value = Array(stId, whId, Environ$("COMPUTERNAME"), "RECEIVE")
    wsSt.ListObjects.Add(xlSrcRange, wsSt.Range("A1:D2"), , xlYes).Name = "tblStationConfig"

    p = Environ$("TEMP") & "\WH1.invSys.Config.test.xlsx"
    On Error Resume Next
    Kill p
    On Error GoTo 0
    wb.SaveAs Filename:=p, FileFormat:=51

    Set BuildConfigWorkbook = wb
End Function

Private Function BuildAuthWorkbook(ByVal whId As String) As Workbook
    Dim wb As Workbook
    Dim wsUsers As Worksheet
    Dim wsCaps As Worksheet
    Dim p As String

    Set wb = Application.Workbooks.Add
    Set wsUsers = wb.Worksheets(1)
    wsUsers.Name = "Users"
    Set wsCaps = wb.Worksheets.Add(After:=wsUsers)
    wsCaps.Name = "Capabilities"

    wsUsers.Range("A1").Resize(1, 6).Value = Array("UserId", "DisplayName", "PinHash", "Status", "ValidFrom", "ValidTo")
    wsUsers.Range("A2").Resize(1, 6).Value = Array("user1", "User One", "", "Active", "", "")
    wsUsers.Range("A3").Resize(1, 6).Value = Array("user2", "User Two", "", "Active", "", "")
    wsUsers.ListObjects.Add(xlSrcRange, wsUsers.Range("A1:F3"), , xlYes).Name = "tblUsers"

    wsCaps.Range("A1").Resize(1, 7).Value = Array("UserId", "Capability", "WarehouseId", "StationId", "Status", "ValidFrom", "ValidTo")
    wsCaps.Range("A2").Resize(1, 7).Value = Array("", "", "", "", "", "", "")
    wsCaps.ListObjects.Add(xlSrcRange, wsCaps.Range("A1:G2"), , xlYes).Name = "tblCapabilities"

    p = Environ$("TEMP") & "\WH1.invSys.Auth.test.xlsx"
    On Error Resume Next
    Kill p
    On Error GoTo 0
    wb.SaveAs Filename:=p, FileFormat:=51

    Set BuildAuthWorkbook = wb
End Function

Private Sub AddCapability(ByVal wb As Workbook, _
                          ByVal userId As String, _
                          ByVal capability As String, _
                          ByVal whId As String, _
                          ByVal stId As String, _
                          ByVal status As String, _
                          ByVal validFrom As String, _
                          ByVal validTo As String)
    Dim lo As ListObject
    Dim r As ListRow

    Set lo = wb.Worksheets("Capabilities").ListObjects("tblCapabilities")
    If lo.Parent.ProtectContents Then lo.Parent.Unprotect
    If lo.Parent.ProtectContents Then
        Err.Raise vbObjectError + 2602, "TestCoreAuth.AddCapability", _
                  "Worksheet '" & lo.Parent.Name & "' is protected and could not be unprotected before writing to tblCapabilities."
    End If
    Set r = lo.ListRows.Add
    r.Range.Cells(1, lo.ListColumns("UserId").Index).Value = userId
    r.Range.Cells(1, lo.ListColumns("Capability").Index).Value = capability
    r.Range.Cells(1, lo.ListColumns("WarehouseId").Index).Value = whId
    r.Range.Cells(1, lo.ListColumns("StationId").Index).Value = stId
    r.Range.Cells(1, lo.ListColumns("Status").Index).Value = status
    r.Range.Cells(1, lo.ListColumns("ValidFrom").Index).Value = validFrom
    r.Range.Cells(1, lo.ListColumns("ValidTo").Index).Value = validTo
End Sub

Private Sub Tally(ByVal testResult As Long, ByRef passed As Long, ByRef failed As Long)
    If testResult = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub

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
