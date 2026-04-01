Attribute VB_Name = "TestWarehouseRetireReAuth"
Option Explicit

Public Function TestValidateUserCredential_SucceedsWithCorrectPasswordAndRole() As Long
    Dim rootPath As String
    Dim warehouseId As String
    Dim stationId As String
    Dim adminUser As String
    Dim passwordText As String

    warehouseId = "WHRET2A"
    stationId = "ADM1"
    adminUser = "admin.reauth"
    passwordText = "654321"
    rootPath = BuildRetireReAuthTempRoot("success")

    On Error GoTo CleanFail
    If Not PrepareRetireReAuthFixture(rootPath, warehouseId, stationId, adminUser, passwordText) Then GoTo CleanExit

    If modAuth.ValidateUserCredential(adminUser, passwordText, "ADMIN_MAINT") Then
        TestValidateUserCredential_SucceedsWithCorrectPasswordAndRole = 1
    End If

CleanExit:
    CleanupRetireReAuthFixture rootPath, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReAuthGate_WrongPassword_ShowsInlineErrorAndDoesNotAuthenticate() As Long
    Dim rootPath As String
    Dim warehouseId As String
    Dim stationId As String
    Dim adminUser As String
    Dim gate As frmReAuthGate

    warehouseId = "WHRET2B"
    stationId = "ADM1"
    adminUser = "admin.reauth"
    rootPath = BuildRetireReAuthTempRoot("wrong_password")

    On Error GoTo CleanFail
    If Not PrepareRetireReAuthFixture(rootPath, warehouseId, stationId, adminUser, "654321") Then GoTo CleanExit

    modDiagnostics.ResetDiagnosticCapture
    Set gate = New frmReAuthGate
    gate.InitializeGate "ADMIN_MAINT", adminUser
    gate.SetPasswordTextForTest "bad-password"
    gate.SimulateSubmit

    If (Not gate.Authenticated) _
       And gate.FailureCount = 1 _
       And (Not gate.IsLockedOut) _
       And gate.IsSubmitEnabled _
       And InStr(1, gate.ErrorText, "Invalid credentials", vbTextCompare) > 0 _
       And modDiagnostics.GetDiagnosticEventCount() = 0 Then
        TestReAuthGate_WrongPassword_ShowsInlineErrorAndDoesNotAuthenticate = 1
    End If

CleanExit:
    On Error Resume Next
    If Not gate Is Nothing Then Unload gate
    On Error GoTo 0
    CleanupRetireReAuthFixture rootPath, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReAuthGate_ThreeFailures_LocksOutAndLogs() As Long
    Dim rootPath As String
    Dim warehouseId As String
    Dim stationId As String
    Dim adminUser As String
    Dim gate As frmReAuthGate

    warehouseId = "WHRET2C"
    stationId = "ADM1"
    adminUser = "admin.reauth"
    rootPath = BuildRetireReAuthTempRoot("lockout")

    On Error GoTo CleanFail
    If Not PrepareRetireReAuthFixture(rootPath, warehouseId, stationId, adminUser, "654321") Then GoTo CleanExit

    modDiagnostics.ResetDiagnosticCapture
    Set gate = New frmReAuthGate
    gate.InitializeGate "ADMIN_MAINT", adminUser

    gate.SetPasswordTextForTest "bad-1"
    gate.SimulateSubmit
    gate.SetPasswordTextForTest "bad-2"
    gate.SimulateSubmit
    gate.SetPasswordTextForTest "bad-3"
    gate.SimulateSubmit

    If (Not gate.Authenticated) _
       And gate.FailureCount = 3 _
       And gate.IsLockedOut _
       And (Not gate.IsSubmitEnabled) _
       And InStr(1, modDiagnostics.GetLastDiagnosticCategory(), "REAUTH", vbTextCompare) > 0 _
       And InStr(1, modDiagnostics.GetLastDiagnosticMessage(), "Lockout|UserId=" & adminUser, vbTextCompare) > 0 _
       And modDiagnostics.GetDiagnosticEventCount() = 1 Then
        TestReAuthGate_ThreeFailures_LocksOutAndLogs = 1
    End If

CleanExit:
    On Error Resume Next
    If Not gate Is Nothing Then Unload gate
    On Error GoTo 0
    CleanupRetireReAuthFixture rootPath, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReAuthGate_Cancel_LeavesUnauthenticatedWithoutLog() As Long
    Dim rootPath As String
    Dim warehouseId As String
    Dim stationId As String
    Dim adminUser As String
    Dim gate As frmReAuthGate

    warehouseId = "WHRET2D"
    stationId = "ADM1"
    adminUser = "admin.reauth"
    rootPath = BuildRetireReAuthTempRoot("cancel")

    On Error GoTo CleanFail
    If Not PrepareRetireReAuthFixture(rootPath, warehouseId, stationId, adminUser, "654321") Then GoTo CleanExit

    modDiagnostics.ResetDiagnosticCapture
    Set gate = New frmReAuthGate
    gate.InitializeGate "ADMIN_MAINT", adminUser
    gate.SimulateCancel

    If (Not gate.Authenticated) _
       And gate.FailureCount = 0 _
       And (Not gate.IsLockedOut) _
       And modDiagnostics.GetDiagnosticEventCount() = 0 Then
        TestReAuthGate_Cancel_LeavesUnauthenticatedWithoutLog = 1
    End If

CleanExit:
    On Error Resume Next
    If Not gate Is Nothing Then Unload gate
    On Error GoTo 0
    CleanupRetireReAuthFixture rootPath, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function PrepareRetireReAuthFixture(ByVal rootPath As String, _
                                            ByVal warehouseId As String, _
                                            ByVal stationId As String, _
                                            ByVal adminUser As String, _
                                            ByVal passwordText As String) As Boolean
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook

    On Error GoTo FailPrepare

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook(warehouseId, stationId, rootPath, "ADMIN")
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook(warehouseId, rootPath)

    TestPhase2Helpers.AddCapability wbAuth, adminUser, "ADMIN_MAINT", warehouseId, stationId, "ACTIVE"
    TestPhase2Helpers.SetUserPinHash wbAuth, adminUser, modAuth.HashUserCredential(passwordText)
    wbCfg.Save
    wbAuth.Save

    If Not modConfig.LoadConfig(warehouseId, stationId) Then GoTo CleanExit
    If Not modAuth.LoadAuth(warehouseId) Then GoTo CleanExit

    PrepareRetireReAuthFixture = True

CleanExit:
    On Error Resume Next
    If Not wbAuth Is Nothing Then wbAuth.Close SaveChanges:=False
    If Not wbCfg Is Nothing Then wbCfg.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

FailPrepare:
    Resume CleanExit
End Function

Private Sub CleanupRetireReAuthFixture(ByVal rootPath As String, ByVal warehouseId As String)
    On Error Resume Next
    CloseWorkbookByNameRetireReAuth warehouseId & ".invSys.Config.xlsb"
    CloseWorkbookByNameRetireReAuth warehouseId & ".invSys.Auth.xlsb"
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    DeleteFolderRecursiveRetireReAuth rootPath
    On Error GoTo 0
End Sub

Private Function BuildRetireReAuthTempRoot(ByVal suffix As String) As String
    BuildRetireReAuthTempRoot = Environ$("TEMP") & "\invSys_retire_reauth_" & suffix & "_" & Format$(Now, "yyyymmdd_hhnnss")
End Function

Private Sub DeleteFolderRecursiveRetireReAuth(ByVal folderPath As String)
    Dim childName As String
    Dim childPath As String
    Dim attrs As Long

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub

    childName = Dir$(folderPath & "\*", vbNormal Or vbHidden Or vbSystem Or vbDirectory)
    Do While childName <> ""
        If childName <> "." And childName <> ".." Then
            childPath = folderPath & "\" & childName
            attrs = GetAttr(childPath)
            If (attrs And vbDirectory) = vbDirectory Then
                DeleteFolderRecursiveRetireReAuth childPath
            Else
                SetAttr childPath, vbNormal
                Kill childPath
            End If
        End If
        childName = Dir$
    Loop

    RmDir folderPath
End Sub

Private Sub CloseWorkbookByNameRetireReAuth(ByVal workbookName As String)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next wb
End Sub
