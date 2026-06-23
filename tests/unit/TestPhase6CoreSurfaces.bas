Attribute VB_Name = "TestPhase6CoreSurfaces"
Option Explicit

Private mLastTestFailure As String

Public Sub ClearLastTestFailure()
    mLastTestFailure = vbNullString
End Sub

Public Function GetLastTestFailure() As String
    GetLastTestFailure = mLastTestFailure
End Function

Public Function TestNasSelectWarehouseTarget_ReadsWarehouseIdFromConfig() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_dnas_select_not_folder_WH")

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH77", "S3", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH77", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S3", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    If target Is Nothing Then GoTo CleanExit
    If StrComp(target.WarehouseId, "WH77", vbTextCompare) = 0 _
       And InStr(1, rootPath, "WH77", vbTextCompare) = 0 _
       And StrComp(target.StationId, "S3", vbTextCompare) = 0 Then
        TestNasSelectWarehouseTarget_ReadsWarehouseIdFromConfig = 1
    End If

CleanExit:
    modNasConnection.ForgetTarget "WH77"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasGetCurrentTarget_ReturnsDeepCopy() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim selectedTarget As WarehouseTarget
    Dim targetCopy As WarehouseTarget
    Dim secondCopy As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_dnas_copy")

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH78", "S4", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH78", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, selectedTarget, "S4", True)
    If statusCode <> NAS_OK Then GoTo CleanExit

    Set targetCopy = modNasConnection.GetCurrentTarget()
    If targetCopy Is Nothing Then GoTo CleanExit
    targetCopy.WarehouseId = "MUTATED"
    targetCopy.RuntimeRoot = "C:\mutated"

    Set secondCopy = modNasConnection.GetCurrentTarget()
    If Not secondCopy Is Nothing Then
        If StrComp(secondCopy.WarehouseId, "WH78", vbTextCompare) = 0 _
           And StrComp(secondCopy.RuntimeRoot, rootPath, vbTextCompare) = 0 Then
            TestNasGetCurrentTarget_ReturnsDeepCopy = 1
        End If
    End If

CleanExit:
    modNasConnection.ForgetTarget "WH78"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasSelectWarehouseTarget_RequiresStationInboxRejectsBlankStation() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_dnas_station_required")

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH79", "S5", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH79", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "", True)
    If statusCode = WH_TARGET_INCOMPLETE And target Is Nothing Then
        TestNasSelectWarehouseTarget_RequiresStationInboxRejectsBlankStation = 1
    End If

CleanExit:
    modNasConnection.ForgetTarget "WH79"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasSelectWarehouseTarget_AllowsRoamingBlankStationWithoutInboxRequirement() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_dnas_roaming_blank_station")

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH74", "S7", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH74", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "", False)
    If statusCode <> NAS_OK Then GoTo CleanExit
    If target Is Nothing Then GoTo CleanExit
    If StrComp(target.WarehouseId, "WH74", vbTextCompare) = 0 _
       And target.StationId = "" _
       And StrComp(target.InboxRoot, rootPath, vbTextCompare) = 0 _
       And target.SourceType = WH_SOURCE_NAS Then
        TestNasSelectWarehouseTarget_AllowsRoamingBlankStationWithoutInboxRequirement = 1
    End If

CleanExit:
    modNasConnection.ForgetTarget "WH74"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasSelectWarehouseTarget_TwoStationsHaveIndependentInboxRoots() As Long
    Dim rootPath As String
    Dim inboxRootA As String
    Dim inboxRootB As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim targetA As WarehouseTarget
    Dim targetB As WarehouseTarget
    Dim statusA As NasStatusCode
    Dim statusB As NasStatusCode
    Dim report As String
    Dim configPath As String

    rootPath = BuildRuntimeTestRoot("phase6_dnas_two_station_inbox")
    inboxRootA = rootPath & "\station_s21"
    inboxRootB = rootPath & "\station_s22"
    configPath = rootPath & "\WH75.invSys.Config.xlsb"

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH75", "S21", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH75", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modConfig.EnsureStationConfigEntry("WH75", "S21", "Station 21", inboxRootA, "RECEIVE", configPath, rootPath, report) Then GoTo CleanExit
    If Not modConfig.EnsureStationConfigEntry("WH75", "S22", "Station 22", inboxRootB, "RECEIVE", configPath, rootPath, report) Then GoTo CleanExit

    statusA = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, targetA, "S21", True)
    statusB = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, targetB, "S22", True)
    If statusA <> NAS_OK Or statusB <> NAS_OK Then GoTo CleanExit
    If targetA Is Nothing Or targetB Is Nothing Then GoTo CleanExit

    If StrComp(targetA.WarehouseId, "WH75", vbTextCompare) = 0 _
       And StrComp(targetB.WarehouseId, "WH75", vbTextCompare) = 0 _
       And StrComp(targetA.StationId, "S21", vbTextCompare) = 0 _
       And StrComp(targetB.StationId, "S22", vbTextCompare) = 0 _
       And StrComp(targetA.InboxRoot, inboxRootA, vbTextCompare) = 0 _
       And StrComp(targetB.InboxRoot, inboxRootB, vbTextCompare) = 0 Then
        TestNasSelectWarehouseTarget_TwoStationsHaveIndependentInboxRoots = 1
    End If

CleanExit:
    modNasConnection.ForgetTarget "WH75"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasScanRoot_ReturnsPathStringsWithoutWarehouseInference() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim targets As Collection
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_dnas_scan_invsys_Zenbook_WH")

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH80", "S6", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH80", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit

    Set targets = modNasConnection.ScanNasRoot(rootPath)
    If Not targets Is Nothing Then
        If targets.Count = 1 _
           And StrComp(CStr(targets(1)), rootPath, vbTextCompare) = 0 _
           And InStr(1, CStr(targets(1)), "WH80", vbTextCompare) = 0 Then
            TestNasScanRoot_ReturnsPathStringsWithoutWarehouseInference = 1
        End If
    End If

CleanExit:
    modNasConnection.ForgetTarget "WH80"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasScanRoot_RejectsMismatchedConfigAuthPair() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim targets As Collection
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim report As String

    rootPath = BuildRuntimeTestRoot("phase6_dnas_scan_mismatched_pair")

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("BADCFG", "S6", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH80", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit

    Set targets = modNasConnection.ScanNasRoot(rootPath)
    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S6", False)
    If Not targets Is Nothing Then
        If targets.Count = 0 And statusCode = WH_AUTH_NOT_FOUND And target Is Nothing Then
            TestNasScanRoot_RejectsMismatchedConfigAuthPair = 1
        End If
    End If

CleanExit:
    modNasConnection.ForgetTarget "BADCFG"
    modNasConnection.ForgetTarget "WH80"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasResolveRememberedTarget_UnreachableFailsClosed() As Long
    Dim rootPath As String
    Dim remembered As WarehouseTarget
    Dim resolved As WarehouseTarget
    Dim statusCode As NasStatusCode

    rootPath = BuildRuntimeTestRoot("phase6_dnas_fail_closed")

    On Error GoTo CleanFail
    DeleteRuntimeRoot rootPath
    Set remembered = New WarehouseTarget
    remembered.WarehouseId = "WH81"
    remembered.WarehouseName = "Warehouse WH81"
    remembered.StationId = "S7"
    remembered.HubRoot = rootPath
    remembered.RuntimeRoot = rootPath
    remembered.ConfigPath = rootPath & "\WH81.invSys.Config.xlsb"
    remembered.AuthPath = rootPath & "\WH81.invSys.Auth.xlsb"
    remembered.InboxRoot = rootPath
    remembered.SourceType = WH_SOURCE_REMEMBERED
    remembered.LastResolvedUTC = Now

    modNasConnection.RememberTarget remembered
    modNasConnection.ClearWarehouseTarget

    If Not modNasConnection.ResolveWarehouseTarget(resolved, statusCode) _
       And statusCode = NAS_TARGET_UNREACHABLE _
       And Not modNasConnection.IsTargetResolved() Then
        If Not resolved Is Nothing Then
            If StrComp(resolved.WarehouseId, "WH81", vbTextCompare) = 0 Then
                TestNasResolveRememberedTarget_UnreachableFailsClosed = 1
            End If
        End If
    End If

CleanExit:
    modNasConnection.ForgetTarget "WH81"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasResolveRememberedTarget_ReachableRecomputesCachedHints() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim remembered As WarehouseTarget
    Dim resolved As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim report As String
    Dim expectedConfigPath As String
    Dim expectedAuthPath As String

    rootPath = BuildRuntimeTestRoot("phase6_dnas_remembered_recompute")
    expectedConfigPath = rootPath & "\WH84.invSys.Config.xlsb"
    expectedAuthPath = rootPath & "\WH84.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH84", "S10", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH84", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit

    Set remembered = New WarehouseTarget
    remembered.WarehouseId = "WH84"
    remembered.WarehouseName = "Warehouse WH84"
    remembered.StationId = "S10"
    remembered.HubRoot = rootPath
    remembered.RuntimeRoot = rootPath
    remembered.ConfigPath = rootPath & "\stale\wrong.Config.xlsb"
    remembered.AuthPath = rootPath & "\stale\wrong.Auth.xlsb"
    remembered.InboxRoot = rootPath & "\stale\inbox"
    remembered.SourceType = WH_SOURCE_REMEMBERED
    remembered.LastResolvedUTC = Now

    modNasConnection.RememberTarget remembered
    modNasConnection.ClearWarehouseTarget

    If modNasConnection.ResolveWarehouseTarget(resolved, statusCode) _
       And statusCode = NAS_OK _
       And Not resolved Is Nothing Then
        If StrComp(resolved.WarehouseId, "WH84", vbTextCompare) = 0 _
           And StrComp(resolved.StationId, "S10", vbTextCompare) = 0 _
           And StrComp(resolved.ConfigPath, expectedConfigPath, vbTextCompare) = 0 _
           And StrComp(resolved.AuthPath, expectedAuthPath, vbTextCompare) = 0 _
           And resolved.SourceType <> WH_SOURCE_FALLBACK Then
            TestNasResolveRememberedTarget_ReachableRecomputesCachedHints = 1
        End If
    End If

CleanExit:
    modNasConnection.ForgetTarget "WH84"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestNasFallbackPolicy_RoleRejectsFallbackAdminAccepts() As Long
    Dim fallbackTarget As WarehouseTarget
    Dim nasTarget As WarehouseTarget

    On Error GoTo CleanFail
    Set fallbackTarget = New WarehouseTarget
    fallbackTarget.WarehouseId = "WH82"
    fallbackTarget.WarehouseName = "Warehouse WH82"
    fallbackTarget.StationId = "S8"
    fallbackTarget.HubRoot = "C:\invSys\WH82"
    fallbackTarget.RuntimeRoot = "C:\invSys\WH82"
    fallbackTarget.SourceType = WH_SOURCE_FALLBACK
    fallbackTarget.LastResolvedUTC = Now

    Set nasTarget = New WarehouseTarget
    nasTarget.WarehouseId = "WH83"
    nasTarget.WarehouseName = "Warehouse WH83"
    nasTarget.StationId = "S9"
    nasTarget.HubRoot = "\\server\invSysWH83"
    nasTarget.RuntimeRoot = "\\server\invSysWH83"
    nasTarget.SourceType = WH_SOURCE_NAS
    nasTarget.LastResolvedUTC = Now

    If Not modNasConnection.IsWarehouseTargetAllowed(fallbackTarget, True) _
       And modNasConnection.IsWarehouseTargetAllowed(fallbackTarget, False) _
       And modNasConnection.IsWarehouseTargetAllowed(nasTarget, True) Then
        TestNasFallbackPolicy_RoleRejectsFallbackAdminAccepts = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthValidateUserCredentialForTarget_SignsInAndStatusOk() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_target_signin")
    authPath = rootPath & "\WH85.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH85", "S11", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH85", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH85", "S11", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S11", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "123456", target, "RECEIVE_POST")

    If authStatus = AUTH_OK _
       And modAuth.IsSignedIn() _
       And StrComp(modAuth.GetCurrentUserId(), "dilbert", vbTextCompare) = 0 _
       And modAuth.GetAuthStatus() = AUTH_OK Then
        TestAuthValidateUserCredentialForTarget_SignsInAndStatusOk = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH85"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthValidateUserCredentialForTarget_AcceptsResetPinForUserId() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_reset_pin")
    authPath = rootPath & "\WH88.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH88", "S14", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH88", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH88", "S14", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("old-pin")
    wbAuth.Save
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("new-pin")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S14", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "new-pin", target, "RECEIVE_POST")

    If authStatus = AUTH_OK _
       And modAuth.IsSignedIn() _
       And StrComp(modAuth.GetCurrentUserId(), "dilbert", vbTextCompare) = 0 _
       And StrComp(modAuth.GetCurrentUserDisplayName(), "Dilbert", vbTextCompare) = 0 Then
        TestAuthValidateUserCredentialForTarget_AcceptsResetPinForUserId = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH88"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthValidateUserCredentialForTarget_RejectsDisplayNameAsUserId() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_reject_display")
    authPath = rootPath & "\WH82.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH82", "S15", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH82", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH82", "S15", "u1", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "u1", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S15", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("Dilbert", "123456", target, "RECEIVE_POST")

    If authStatus = AUTH_USER_NOT_FOUND _
       And Not modAuth.IsSignedIn() _
       And modAuth.GetCurrentUserId() = "" Then
        TestAuthValidateUserCredentialForTarget_RejectsDisplayNameAsUserId = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH82"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthValidateUserCredentialForTarget_RejectsMismatchedTargetWarehouse() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_target_mismatch")
    authPath = rootPath & "\WH83.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH83", "S16", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH83", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH83", "S16", "u1", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "u1", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S16", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    target.WarehouseId = "WH83_MUTATED"
    authStatus = modAuth.ValidateUserCredentialForTarget("u1", "123456", target, "RECEIVE_POST")

    If authStatus = AUTH_WAREHOUSE_MISMATCH _
       And Not modAuth.IsSignedIn() _
       And modAuth.GetCurrentUserId() = "" Then
        TestAuthValidateUserCredentialForTarget_RejectsMismatchedTargetWarehouse = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH83"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthCapabilityScope_AllowsSelectedRuntimeFolderAlias() As Long
    Dim parentRoot As String
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String

    parentRoot = BuildRuntimeTestRoot("phase6_auth_scope_alias")
    rootPath = parentRoot & "\invsys_Zenbook_WH"
    If Len(Dir$(rootPath, vbDirectory)) = 0 Then MkDir rootPath

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH98", "S33", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH98", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    TestPhase2Helpers.AddCapability wbAuth, "justin", "RECEIVE_POST", "invsys_Zenbook_WH", "S33", "ACTIVE"
    TestPhase2Helpers.SetUserPinHash wbAuth, "justin", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(parentRoot, rootPath, target, "S33", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("justin", "123456", target, "RECEIVE_POST")
    If authStatus <> AUTH_OK Then GoTo CleanExit
    If Not modAuth.CanPerform("RECEIVE_POST", "justin", "WH98", "S33", "TEST", "AUTH-SCOPE-ALIAS") Then GoTo CleanExit

    TestAuthCapabilityScope_AllowsSelectedRuntimeFolderAlias = 1

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH98"
    modNasConnection.ForgetRoot parentRoot
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot parentRoot
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthFailedCredential_DoesNotReplaceSignedInUser() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim firstStatus As AuthStatusCode
    Dim secondStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_failed_preserve")
    authPath = rootPath & "\WH86.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH86", "S12", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH86", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH86", "S12", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH86", "S12", "calvin", "Calvin", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    TestPhase2Helpers.SetUserPinHash wbAuth, "calvin", modAuth.HashUserCredential("654321")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S12", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    firstStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "123456", target, "RECEIVE_POST")
    secondStatus = modAuth.ValidateUserCredentialForTarget("calvin", "wrong", target, "RECEIVE_POST")

    If firstStatus = AUTH_OK _
       And secondStatus = AUTH_CREDENTIAL_REJECTED _
       And StrComp(modAuth.GetCurrentUserId(), "dilbert", vbTextCompare) = 0 Then
        TestAuthFailedCredential_DoesNotReplaceSignedInUser = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH86"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthCorrectCredentialWithoutCapability_ReturnsNoCapabilities() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_no_capability")
    authPath = rootPath & "\WH87.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH87", "S13", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH87", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH87", "S13", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S13", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "123456", target, "SHIP_POST")

    If authStatus = AUTH_NO_CAPABILITIES _
       And Not modAuth.IsSignedIn() _
       And modAuth.GetCurrentUserId() = "" Then
        TestAuthCorrectCredentialWithoutCapability_ReturnsNoCapabilities = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH87"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRuntimeStatusUserLabel_UnsignedShowsNotSignedIn() As Long
    On Error GoTo CleanFail
    modAuth.SignOut

    If StrComp(modRibbonRuntimeStatus.GetStatusLabel("btnRuntimeUser"), "User ID: <not signed in>", vbTextCompare) = 0 Then
        TestRuntimeStatusUserLabel_UnsignedShowsNotSignedIn = 1
    End If

CleanExit:
    modAuth.SignOut
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRuntimeStatusUserLabel_TracksAuthSignIn() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_runtime_user_status")
    authPath = rootPath & "\WH88.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH88", "S14", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH88", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH88", "S14", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S14", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "123456", target, "RECEIVE_POST")

    If authStatus = AUTH_OK _
       And StrComp(modRibbonRuntimeStatus.GetStatusLabel("btnRuntimeUser"), "User ID: dilbert", vbTextCompare) = 0 Then
        TestRuntimeStatusUserLabel_TracksAuthSignIn = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH88"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRoleWriteCurrent_RejectsUnsignedUser() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim report As String
    Dim authPath As String
    Dim eventIdOut As String
    Dim queued As Boolean

    rootPath = BuildRuntimeTestRoot("phase6_role_write_unsigned")
    authPath = rootPath & "\WH89.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH89", "S15", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH89", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH89", "S15", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S15", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    If Not modNasConnection.SetCurrentTargetPathsForTest("\\test-nas\invSysWH1", "\\test-nas\invSysWH1\WH89") Then GoTo CleanExit
    report = ""
    queued = modRoleEventWriter.QueueReceiveEventCurrent("", "SKU-RM-UNSIGNED", 1, "A1", "unsigned", eventIdOut, report)

    If Not queued _
       And eventIdOut = "" _
       And InStr(1, report, "not signed in", vbTextCompare) > 0 Then
        TestRoleWriteCurrent_RejectsUnsignedUser = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH89"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRoleWriteCurrent_RejectsMissingCapability() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String
    Dim eventIdOut As String
    Dim queued As Boolean
    Dim payloadJson As String

    rootPath = BuildRuntimeTestRoot("phase6_role_write_no_cap")
    authPath = rootPath & "\WH90.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH90", "S16", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH90", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH90", "S16", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S16", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    If Not modNasConnection.SetCurrentTargetPathsForTest("\\test-nas\invSysWH1", "\\test-nas\invSysWH1\WH90") Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "123456", target)
    If authStatus <> AUTH_OK Then GoTo CleanExit

    payloadJson = modRoleEventWriter.BuildPayloadJson( _
        modRoleEventWriter.CreatePayloadItem(1, "SKU-RM-NOCAP", 1, "A1", "no-cap"))
    report = ""
    queued = modRoleEventWriter.QueuePayloadEventCurrent(CORE_EVENT_TYPE_SHIP, "", payloadJson, "no-cap", eventIdOut, report)

    If Not queued _
       And eventIdOut = "" _
       And InStr(1, report, "lacks SHIP_POST", vbTextCompare) > 0 Then
        TestRoleWriteCurrent_RejectsMissingCapability = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH90"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRoleWriteCurrent_RejectsFallbackTarget() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String
    Dim eventIdOut As String
    Dim queued As Boolean

    rootPath = BuildRuntimeTestRoot("phase6_role_write_fallback")
    authPath = rootPath & "\WH91.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH91", "S17", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH91", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH91", "S17", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S17", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "123456", target, "RECEIVE_POST")
    If authStatus <> AUTH_OK Then GoTo CleanExit
    If Not modNasConnection.SetCurrentTargetSourceTypeForTest(WH_SOURCE_FALLBACK) Then GoTo CleanExit

    report = ""
    queued = modRoleEventWriter.QueueReceiveEventCurrent("", "SKU-RM-FALLBACK", 1, "A1", "fallback", eventIdOut, report)

    If Not queued _
       And eventIdOut = "" _
       And InStr(1, report, "NAS warehouse target", vbTextCompare) > 0 Then
        TestRoleWriteCurrent_RejectsFallbackTarget = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH91"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRoleWriteCurrent_AllowsSignedInReceivePost() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String
    Dim eventIdOut As String
    Dim queued As Boolean

    rootPath = BuildRuntimeTestRoot("phase6_role_write_receive_ok")
    authPath = rootPath & "\WH92.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH92", "S18", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH92", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH92", "S18", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S18", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    If Not modNasConnection.SetCurrentTargetPathsForTest("\\test-nas\invSysWH1", "\\test-nas\invSysWH1\WH92") Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "123456", target, "RECEIVE_POST")
    If authStatus <> AUTH_OK Then GoTo CleanExit

    report = ""
    queued = modRoleEventWriter.QueueReceiveEventCurrent("", "SKU-RM-ALLOW", 2, "A1", "allowed", eventIdOut, report)

    If queued _
       And eventIdOut <> "" _
       And report = "" Then
        TestRoleWriteCurrent_AllowsSignedInReceivePost = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH92"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthSignOut_ClearsUserButKeepsWarehouseTarget() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim targetAfterSignOut As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_signout_keeps_target")
    authPath = rootPath & "\WH93.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH93", "S19", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH93", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH93", "S19", "calvin", "Calvin", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "calvin", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S19", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    If Not modNasConnection.SetCurrentTargetPathsForTest("\\test-nas\invSysWH1", "\\test-nas\invSysWH1\WH93") Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("calvin", "123456", target, "RECEIVE_POST")
    If authStatus <> AUTH_OK Then GoTo CleanExit

    modAuth.SignOut
    Set targetAfterSignOut = modNasConnection.GetCurrentTarget()
    If Not modAuth.IsSignedIn() _
       And modAuth.GetCurrentUserId() = "" _
       And Not targetAfterSignOut Is Nothing Then
        If StrComp(targetAfterSignOut.WarehouseId, "WH93", vbTextCompare) = 0 _
           And StrComp(targetAfterSignOut.StationId, "S19", vbTextCompare) = 0 _
           And modNasConnection.IsCurrentTargetAllowed(True) Then
            TestAuthSignOut_ClearsUserButKeepsWarehouseTarget = 1
        End If
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH93"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthCanPerform_SignedOutFailsClosedWithLoadedAuth() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_signedout_canperform")
    authPath = rootPath & "\WH94.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH94", "S20", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH94", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    If Not modAuth.EnsureStationRoleAuth("WH94", "S20", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S20", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH94") Then GoTo CleanExit
    modAuth.SignOut

    If Not modAuth.CanPerform("RECEIVE_POST", "dilbert", "WH94", "S20", "TEST", "AUTH-SIGNEDOUT") _
       And Not modAuth.IsSignedIn() Then
        TestAuthCanPerform_SignedOutFailsClosedWithLoadedAuth = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH94"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAuthTtlExpiry_FailsClosedForIsSignedInAndCanPerform() As Long
    Dim rootPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim loWh As ListObject
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim authPath As String

    rootPath = BuildRuntimeTestRoot("phase6_auth_ttl_expiry")
    authPath = rootPath & "\WH95.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modAuth.SignOut
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH95", "S21", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH95", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit
    Set loWh = FindTableByName(wbCfg, "tblWarehouseConfig")
    If loWh Is Nothing Then GoTo CleanExit
    SetTableCell loWh, 1, "AuthCacheTTLSeconds", 1
    wbCfg.Save
    If Not modAuth.EnsureStationRoleAuth("WH95", "S21", "dilbert", "Dilbert", "RECEIVE", authPath, "svc_processor", report:=report) Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, "dilbert", modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S21", True)
    If statusCode <> NAS_OK Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget("dilbert", "123456", target, "RECEIVE_POST")
    If authStatus <> AUTH_OK Then GoTo CleanExit
    Application.Wait Now + TimeSerial(0, 0, 2)

    If Not modAuth.IsSignedIn() _
       And modAuth.GetAuthStatus() = AUTH_REAUTH_REQUIRED _
       And Not modAuth.CanPerform("RECEIVE_POST", "dilbert", "WH95", "S21", "TEST", "AUTH-TTL") Then
        TestAuthTtlExpiry_FailsClosedForIsSignedInAndCanPerform = 1
    End If

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH95"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbCfg
    CloseWorkbookIfOpen wbAuth
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

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
    Dim configPath As String
    Dim wb As Workbook
    Dim openedForVerify As Boolean
    Dim loWh As ListObject
    Dim loSt As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_cfg_load")
    configPath = rootPath & "\WH62.invSys.Config.xlsb"

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH62", "S2") Then GoTo CleanExit

    If Len(Dir$(configPath)) = 0 Then GoTo CleanExit
    Set wb = FindWorkbookByFullPathForTest(configPath)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(configPath)
        openedForVerify = Not wb Is Nothing
    End If
    If wb Is Nothing Then GoTo CleanExit
    Set loWh = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wb.Worksheets("StationConfig").ListObjects("tblStationConfig")

    If modConfig.IsLoaded() _
       And StrComp(modConfig.GetResolvedWorkbookName(), "WH62.invSys.Config.xlsb", vbTextCompare) = 0 _
       And StrComp(modConfig.GetWarehouseId(), "WH62", vbTextCompare) = 0 _
       And StrComp(modConfig.GetStationId(), "S2", vbTextCompare) = 0 _
       And Not loWh Is Nothing _
       And Not loSt Is Nothing _
       And StrComp(CStr(GetTableValue(loWh, 1, "WarehouseId")), "WH62", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loSt, 1, "StationId")), "S2", vbTextCompare) = 0 Then
        TestLoadConfig_AutoBootstrapsCanonicalWorkbook = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    If openedForVerify Then
        CloseWorkbookIfOpen wb
    End If
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLoadConfig_BlankContextAutoBootstrapsDefaultRuntimeWorkbook() As Long
    Dim rootPath As String
    Dim configPath As String
    Dim wb As Workbook
    Dim openedForVerify As Boolean
    Dim loWh As ListObject
    Dim loSt As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_cfg_blank")
    configPath = rootPath & "\WH1.invSys.Config.xlsb"

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("", "") Then GoTo CleanExit

    If Len(Dir$(configPath)) = 0 Then GoTo CleanExit
    Set wb = FindWorkbookByFullPathForTest(configPath)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(configPath)
        openedForVerify = Not wb Is Nothing
    End If
    If wb Is Nothing Then GoTo CleanExit
    Set loWh = wb.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wb.Worksheets("StationConfig").ListObjects("tblStationConfig")

    If modConfig.IsLoaded() _
       And StrComp(modConfig.GetResolvedWorkbookName(), "WH1.invSys.Config.xlsb", vbTextCompare) = 0 _
       And StrComp(modConfig.GetWarehouseId(), "WH1", vbTextCompare) = 0 _
       And StrComp(modConfig.GetStationId(), "S1", vbTextCompare) = 0 _
       And Not loWh Is Nothing _
       And Not loSt Is Nothing _
       And StrComp(CStr(GetTableValue(loWh, 1, "WarehouseId")), "WH1", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loSt, 1, "StationId")), "S1", vbTextCompare) = 0 Then
        TestLoadConfig_BlankContextAutoBootstrapsDefaultRuntimeWorkbook = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    If openedForVerify Then
        CloseWorkbookIfOpen wb
    End If
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
    Dim configPath As String
    Dim authPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim openedCfgForVerify As Boolean
    Dim openedAuthForVerify As Boolean
    Dim loWh As ListObject
    Dim loSt As ListObject
    Dim loUsers As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_auth_load")
    configPath = rootPath & "\WH63.invSys.Config.xlsb"
    authPath = rootPath & "\WH63.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH63", "S3") Then GoTo CleanExit
    If Not modAuth.LoadAuth("WH63") Then GoTo CleanExit

    If Len(Dir$(configPath)) = 0 Then GoTo CleanExit
    If Len(Dir$(authPath)) = 0 Then GoTo CleanExit

    Set wbCfg = FindWorkbookByFullPathForTest(configPath)
    If wbCfg Is Nothing Then
        Set wbCfg = Application.Workbooks.Open(configPath)
        openedCfgForVerify = Not wbCfg Is Nothing
    End If
    Set wbAuth = FindWorkbookByFullPathForTest(authPath)
    If wbAuth Is Nothing Then
        Set wbAuth = Application.Workbooks.Open(authPath)
        openedAuthForVerify = Not wbAuth Is Nothing
    End If
    If wbCfg Is Nothing Or wbAuth Is Nothing Then GoTo CleanExit

    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    If modConfig.IsLoaded() _
       And modAuth.IsAuthLoaded() _
       And StrComp(modConfig.GetResolvedWorkbookName(), "WH63.invSys.Config.xlsb", vbTextCompare) = 0 _
       And StrComp(modAuth.GetResolvedAuthWorkbookName(), "WH63.invSys.Auth.xlsb", vbTextCompare) = 0 _
       And Not loWh Is Nothing _
       And Not loSt Is Nothing _
       And Not loUsers Is Nothing _
       And StrComp(CStr(GetTableValue(loWh, 1, "WarehouseId")), "WH63", vbTextCompare) = 0 _
       And StrComp(CStr(GetTableValue(loSt, 1, "StationId")), "S3", vbTextCompare) = 0 _
       And FindUserRow(loUsers, "svc_processor") > 0 Then
        TestLoadAuth_AutoBootstrapsCanonicalWorkbook = 1
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    If openedAuthForVerify Then CloseWorkbookIfOpen wbAuth
    If openedCfgForVerify Then CloseWorkbookIfOpen wbCfg
    DeleteRuntimeRoot rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestLoadAuth_BootstrapGrantsCurrentOperatorCapabilities() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim report As String
    Dim failureReason As String
    Dim authPath As String
    Dim capabilityOut As String
    Dim canReceive As Boolean
    Dim canShip As Boolean
    Dim canProd As Boolean
    Dim canProcess As Boolean

    rootPath = BuildRuntimeTestRoot("phase6_auth_caps")
    authPath = rootPath & "\WH65.invSys.Auth.xlsb"

    On Error GoTo CleanFail
    ClearLastTestFailure
    modAuth.SignOut
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookByNameIfOpen "WH65.invSys.Config.xlsb"
    CloseWorkbookByNameIfOpen "WH65.invSys.Auth.xlsb"
    CloseWorkbookByNameIfOpen "WH65.invSys.Data.Inventory.xlsb"
    CloseWorkbookByNameIfOpen "WH65.invSys.Snapshot.Inventory.xlsb"
    CloseWorkbookByNameIfOpen "WH65.Outbox.Events.xlsb"
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH65", "S5") Then
        failureReason = "LoadConfig failed for WH65/S5"
        GoTo CleanExit
    End If

    currentUser = Trim$(Environ$("USERNAME"))
    If currentUser = "" Then currentUser = Trim$(Application.UserName)
    If currentUser = "" Then
        failureReason = "Current test user could not be resolved"
        GoTo CleanExit
    End If
    If Not modAuth.EnsureStationRoleAuth("WH65", "S5", currentUser, currentUser, "RECEIVE", authPath, "svc_processor", capabilityOut, report) Then
        failureReason = "EnsureStationRoleAuth RECEIVE failed: " & report
        GoTo CleanExit
    End If
    If Not modAuth.EnsureStationRoleAuth("WH65", "S5", currentUser, currentUser, "SHIP", authPath, "svc_processor", capabilityOut, report) Then
        failureReason = "EnsureStationRoleAuth SHIP failed: " & report
        GoTo CleanExit
    End If
    If Not modAuth.EnsureStationRoleAuth("WH65", "S5", currentUser, currentUser, "PROD", authPath, "svc_processor", capabilityOut, report) Then
        failureReason = "EnsureStationRoleAuth PROD failed: " & report
        GoTo CleanExit
    End If
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH65", "svc_processor", rootPath, report)
    If wbAuth Is Nothing Then
        failureReason = "OpenOrCreateAuthWorkbookRuntime failed: " & report
        GoTo CleanExit
    End If
    TestPhase2Helpers.SetUserPinHash wbAuth, currentUser, modAuth.HashUserCredential("123456")
    wbAuth.Save
    If Not modAuth.LoadAuth("WH65") Then
        failureReason = "LoadAuth failed after provisioning: " & modAuth.ValidateAuth()
        GoTo CleanExit
    End If
    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S5", True)
    If statusCode <> NAS_OK Then
        failureReason = "SelectWarehouseTarget failed: " & CStr(statusCode)
        GoTo CleanExit
    End If
    authStatus = modAuth.ValidateUserCredentialForTarget(currentUser, "123456", target)
    If authStatus <> AUTH_OK Then
        failureReason = "ValidateUserCredentialForTarget failed: " & CStr(authStatus)
        GoTo CleanExit
    End If

    canReceive = modAuth.CanPerform("RECEIVE_POST", currentUser, "WH65", "S5", "TEST", "AUTH-RECV")
    canShip = modAuth.CanPerform("SHIP_POST", currentUser, "WH65", "S5", "TEST", "AUTH-SHIP")
    canProd = modAuth.CanPerform("PROD_POST", currentUser, "WH65", "S5", "TEST", "AUTH-PROD")
    canProcess = modAuth.HasProvisionedCapabilityForSystem("INBOX_PROCESS", "svc_processor", "WH65", "S5")
    If canReceive And canShip And canProd And canProcess Then
        TestLoadAuth_BootstrapGrantsCurrentOperatorCapabilities = 1
    Else
        failureReason = "Capability check failed: RECEIVE_POST=" & CStr(canReceive) _
                        & ", SHIP_POST=" & CStr(canShip) _
                        & ", PROD_POST=" & CStr(canProd) _
                        & ", INBOX_PROCESS=" & CStr(canProcess)
    End If

CleanExit:
    On Error Resume Next
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH65"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    CloseWorkbookIfOpen wbAuth
    CloseWorkbookByNameIfOpen "WH65.invSys.Config.xlsb"
    CloseWorkbookByNameIfOpen "WH65.invSys.Auth.xlsb"
    CloseWorkbookByNameIfOpen "WH65.invSys.Data.Inventory.xlsb"
    CloseWorkbookByNameIfOpen "WH65.invSys.Snapshot.Inventory.xlsb"
    CloseWorkbookByNameIfOpen "WH65.Outbox.Events.xlsb"
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    DeleteRuntimeRoot rootPath
    On Error GoTo 0
    If Len(failureReason) > 0 Then mLastTestFailure = failureReason
    Exit Function
CleanFail:
    If Len(failureReason) = 0 Then failureReason = "Unexpected error " & CStr(Err.Number) & ": " & Err.Description
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
       And CLng(GetTableValue(loInv, 1, "ROW")) = 1 _
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
                                        0, "", "Catalog Item", "CS", "R9", "Catalog Desc", "Vendor C", "VC-9", "raw", "4321")
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
       And StrComp(CStr(GetTableValue(loInv, 1, "ROW")), "4321", vbTextCompare) = 0 _
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

Public Function TestShippingEventCreator_QueuesSignedInCurrentTargetEvent() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim eventIdOut As String
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim processedCount As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbOps As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loInbox As ListObject
    Dim loInventoryLog As ListObject
    Dim invRow As Long
    Dim inboxRow As Long
    Dim logRow As Long
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String

    rootPath = BuildRuntimeTestRoot("phase6_shipping_event_creator")
    currentUser = "calvin"

    On Error GoTo CleanFail
    mLastTestFailure = vbNullString
    modAuth.SignOut
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH96", "S31", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH96", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then
        failureReason = "Config/auth runtime workbooks could not be created. " & report
        GoTo CleanExit
    End If
    SetConfigWarehouseValue "WH96.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.LoadConfig("WH96", "S31") Then
        failureReason = "LoadConfig failed: " & modConfig.Validate()
        GoTo CleanExit
    End If
    If Not modConfig.Reload() Then
        failureReason = "Config reload failed: " & modConfig.Validate()
        GoTo CleanExit
    End If

    EnsureAuthCapabilityForTest "WH96", currentUser, "SHIP_POST", "WH96", "*"
    EnsureAuthCapabilityForTest "WH96", "svc_processor", "INBOX_PROCESS", "WH96", "*"
    TestPhase2Helpers.SetUserPinHash wbAuth, currentUser, modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S31", True)
    If statusCode <> NAS_OK Then
        failureReason = "SelectWarehouseTarget failed: " & CStr(statusCode)
        GoTo CleanExit
    End If
    authStatus = modAuth.ValidateUserCredentialForTarget(currentUser, "123456", target, "SHIP_POST")
    If authStatus <> AUTH_OK Then
        failureReason = "ValidateUserCredentialForTarget failed: " & CStr(authStatus)
        GoTo CleanExit
    End If

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH96", Array("SKU-SHIP-CREATOR"))
    Set wbInbox = CreateCanonicalShipInboxWorkbookForTest(rootPath, "S31")
    If wbInv Is Nothing Or wbInbox Is Nothing Then
        failureReason = "Canonical shipping runtime workbooks could not be created."
        GoTo CleanExit
    End If
    Set evt = CreateReceiveEventForTest("EVT-SHIP-CREATOR-SEED", "WH96", "S31", currentUser, "SKU-SHIP-CREATOR", 12, "A1", "shipping creator seed")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-SHIP-CREATOR-SEED", statusOut, errorCode, errorMessage) Then
        failureReason = "Canonical shipping seed event failed: " & errorCode & "; " & errorMessage
        GoTo CleanExit
    End If

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then
        failureReason = "EnsureShippingWorkbookSurface failed: " & report
        GoTo CleanExit
    End If
    Set loInv = FindTableByName(wbOps, "invSys")
    If loInv Is Nothing Then
        failureReason = "Shipping operator invSys table was not created."
        GoTo CleanExit
    End If
    AddInvSysSeedRow loInv, 962, "SKU-SHIP-CREATOR", "Shipping Creator Item", "EA", "A1", 12
    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-SHIP-CREATOR")
    If invRow = 0 Then
        failureReason = "Shipping operator SKU row was not staged."
        GoTo CleanExit
    End If
    SetTableCell loInv, invRow, "SHIPMENTS", 4

    If Not modNasConnection.SetCurrentTargetPathsForTest("\\test-nas\invSysWH1", "\\test-nas\invSysWH1\WH96") Then GoTo CleanExit
    If Not modShippingEventCreator.QueueShipmentsSentEventFromWorkbook(wbOps, eventIdOut, report) Then
        failureReason = "QueueShipmentsSentEventFromWorkbook failed: " & report
        GoTo CleanExit
    End If
    If Not modNasConnection.SetCurrentTargetPathsForTest(rootPath, rootPath) Then
        failureReason = "Could not restore local processor target after NAS-gated shipping queue."
        GoTo CleanExit
    End If
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Trim$(eventIdOut) = "" Then
        failureReason = "Shipping event creator did not return an EventID."
        GoTo CleanExit
    End If
    Set loInbox = FindTableByName(wbInbox, "tblInboxShip")
    inboxRow = FindRowByColumnValueInTable(loInbox, "EventID", eventIdOut)
    If inboxRow = 0 Then
        failureReason = "Shipping event creator did not write the shipment event directly to the NAS shipping inbox before processor catch-up."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInbox, inboxRow, "EventType")), CORE_EVENT_TYPE_SHIP, vbTextCompare) <> 0 Then
        failureReason = "Shipping inbox recorded unexpected event type before processor catch-up."
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, eventIdOut, "NEW") Then
        failureReason = "Shipping inbox row was not NEW before processor catch-up."
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch("WH96", 500, report)
    If processedCount <> 1 Then
        failureReason = "RunBatch did not process the shipping creator event. " & report
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, eventIdOut, "PROCESSED") Then
        failureReason = "Shipping creator inbox row was not marked PROCESSED."
        GoTo CleanExit
    End If

    Set loInventoryLog = FindTableByName(wbInv, "tblInventoryLog")
    If loInventoryLog Is Nothing Then
        failureReason = "Canonical inventory log was missing after shipping creator process."
        GoTo CleanExit
    End If
    logRow = FindRowByColumnValueInTable(loInventoryLog, "EventID", eventIdOut)
    If logRow = 0 Then
        failureReason = "Canonical inventory log did not record the shipping creator event."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInventoryLog, logRow, "EventType")), CORE_EVENT_TYPE_SHIP, vbTextCompare) <> 0 Then
        failureReason = "Canonical inventory log recorded unexpected event type for shipping creator workflow."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInventoryLog, logRow, "QtyDelta")) <> -4 Then
        failureReason = "Canonical inventory log QtyDelta was not negative for shipping creator workflow."
        GoTo CleanExit
    End If

    TestShippingEventCreator_QueuesSignedInCurrentTargetEvent = 1

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH96"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    CloseWorkbookIfOpen wbAuth
    CloseWorkbookIfOpen wbCfg
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        mLastTestFailure = failureReason
        On Error GoTo 0
        Err.Raise vbObjectError + 7111, "TestShippingEventCreator_QueuesSignedInCurrentTargetEvent", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingState_TombstoneFiltersSentLineIdFromActiveCache() As Long
    Dim rootPath As String
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim activePath As String
    Dim sentPath As String
    Dim rows As Variant
    Dim loShip As ListObject

    rootPath = BuildRuntimeTestRoot("phase6_shipping_tombstone")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH91", "S21") Then GoTo CleanExit
    SetConfigWarehouseValue "WH91.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    activePath = LocalShippingStatePathForTest("active", "WH91")
    sentPath = LocalShippingStatePathForTest("sent", "WH91")
    DeleteFileIfExistsForTest activePath
    DeleteFileIfExistsForTest sentPath
    EnsureFolderForTest ParentFolderPathForTest(activePath)
    WriteTextFileForTest activePath, "REF-TOMB-001" & vbTab & "Tombstone Package" & vbTab & "2" & vbTab & "977" & vbTab & "EA" & vbTab & "A1" & vbTab & "v1" & vbTab & "Shipments" & vbTab & "Carrier" & vbTab & "SHIPLINE-TOMB-001" & vbTab & "RESERVE-TOMB-001"
    WriteTextFileForTest sentPath, "ID:SHIPLINE-TOMB-001"

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then
        failureReason = "EnsureShippingWorkbookSurface failed: " & report
        GoTo CleanExit
    End If

    wbOps.Activate
    rows = RunShippingMacro1ForTest("ShipmentsFormLoadLines", False)
    If Not IsEmpty(rows) Then
        failureReason = "Tombstoned active cache row was loaded back into the Shipments form."
        GoTo CleanExit
    End If

    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If Not loShip Is Nothing Then
        If Not loShip.DataBodyRange Is Nothing Then
            If loShip.ListRows.Count > 0 Then
                If Trim$(CStr(GetTableValue(loShip, 1, "REF_NUMBER"))) <> "" Then
                    failureReason = "Tombstoned row remained in ShipmentsTally projection after load."
                    GoTo CleanExit
                End If
            End If
        End If
    End If

    TestShippingState_TombstoneFiltersSentLineIdFromActiveCache = 1

CleanExit:
    DeleteFileIfExistsForTest activePath
    DeleteFileIfExistsForTest sentPath
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7114, "TestShippingState_TombstoneFiltersSentLineIdFromActiveCache", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingState_SentRowTombstoneFiltersLegacyActiveCache() As Long
    Dim activePath As String
    Dim sentPath As String
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim selectedRows(1 To 1) As Long
    Dim runResult As String
    Dim rows As Variant

    On Error GoTo CleanFail
    If Not modConfig.LoadConfig("WH92", "S21") Then GoTo CleanExit
    activePath = LocalShippingStatePathForTest("active", "WH92")
    sentPath = LocalShippingStatePathForTest("sent", "WH92")
    DeleteFileIfExistsForTest activePath
    DeleteFileIfExistsForTest sentPath

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then
        failureReason = "EnsureShippingWorkbookSurface failed: " & report
        GoTo CleanExit
    End If
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If loInv Is Nothing Or loShip Is Nothing Then
        failureReason = "Shipping workbook surface was incomplete."
        GoTo CleanExit
    End If

    AddInvSysSeedRow loInv, 976, "SKU-LEGACY-TOMB", "Legacy Tombstone Package", "EA", "A1", 20
    SetTableCell loInv, 1, "SHIPMENTS", 1
    AddShippingTallyRow loShip, "REF-LEGACY-TOMB-001", "Legacy Tombstone Package", 1, 976, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "UPS"
    SetTableCell loShip, 1, "LINE_ID", "SHIPLINE-LEGACY-TOMB-001"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-LEGACY-TOMB-001"

    wbOps.Activate
    selectedRows(1) = 1
    runResult = RunShippingSentRowsReportForTest(selectedRows, "UPS")
    If Left$(runResult, 3) <> "OK|" Then
        failureReason = "Shipments Sent did not complete for legacy tombstone setup: " & runResult
        GoTo CleanExit
    End If

    EnsureFolderForTest ParentFolderPathForTest(activePath)
    WriteTextFileForTest activePath, "REF-LEGACY-TOMB-001" & vbTab & "Legacy Tombstone Package" & vbTab & "1" & vbTab & "976" & vbTab & "EA" & vbTab & "A1" & vbTab & "v1" & vbTab & "Shipments" & vbTab & "UPS"

    rows = RunShippingMacro1ForTest("ShipmentsFormLoadLines", False)
    If Not IsEmpty(rows) Then
        failureReason = "Legacy active cache row without LINE_ID was loaded after the matching row was sent."
        GoTo CleanExit
    End If

    TestShippingState_SentRowTombstoneFiltersLegacyActiveCache = 1

CleanExit:
    DeleteFileIfExistsForTest activePath
    DeleteFileIfExistsForTest sentPath
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7153, "TestShippingState_SentRowTombstoneFiltersLegacyActiveCache", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingWorkflowGuard_ShipmentsSentWithZeroStagedFails() As Long
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim resultText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    If loInv Is Nothing Then GoTo CleanExit
    AddInvSysSeedRow loInv, 987, "SKU-SHIP-ZERO", "Zero Staged Package", "EA", "A1", 10
    SetTableCell loInv, 1, "SHIPMENTS", 0

    resultText = modShippingEventCreator.ValidateShipmentsSentStagingFromWorkbook(wbOps)
    If InStr(1, resultText, "No staged shipments found in invSys.SHIPMENTS", vbTextCompare) = 0 Then
        failureReason = "Unexpected Shipments Sent zero-staged validation result: " & resultText
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 0 Then
        failureReason = "Validation mutated invSys.SHIPMENTS."
        GoTo CleanExit
    End If

    TestShippingWorkflowGuard_ShipmentsSentWithZeroStagedFails = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7115, "TestShippingWorkflowGuard_ShipmentsSentWithZeroStagedFails", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingWorkflowGuard_ToShipmentsInsufficientInventoryFails() As Long
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim loAggPack As ListObject
    Dim resultText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loAggPack = FindTableByName(wbOps, "AggregatePackages")
    If loInv Is Nothing Or loAggPack Is Nothing Then GoTo CleanExit
    AddInvSysSeedRow loInv, 988, "SKU-SHIP-SHORT", "Short Package", "EA", "A1", 1
    AddAggregatePackagesRow loAggPack, 988, "Short Package", 2, "EA", "A1"

    resultText = modShippingEventCreator.ValidateToShipmentsFromWorkbook(wbOps)
    If InStr(1, resultText, "ROW 988 requires", vbTextCompare) = 0 _
       Or InStr(1, resultText, "only", vbTextCompare) = 0 _
       Or InStr(1, resultText, "TOTAL INV", vbTextCompare) = 0 Then
        failureReason = "Unexpected To Shipments shortage validation result: " & resultText
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 0 Then
        failureReason = "Validation staged inventory despite shortage."
        GoTo CleanExit
    End If

    TestShippingWorkflowGuard_ToShipmentsInsufficientInventoryFails = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7116, "TestShippingWorkflowGuard_ToShipmentsInsufficientInventoryFails", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingWorkflowGuard_BoxesMadeInsufficientComponentFails() As Long
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim loAggBom As ListObject
    Dim resultText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loAggBom = FindTableByName(wbOps, "AggregateBoxBOM")
    If loInv Is Nothing Or loAggBom Is Nothing Then GoTo CleanExit
    AddInvSysSeedRow loInv, 989, "SKU-COMP-SHORT", "Short Component", "EA", "A1", 3
    AddAggregateBomRow loAggBom, 989, "Short Component", 5, "EA", "A1"

    resultText = modShippingEventCreator.ValidateBoxesMadeFromWorkbook(wbOps)
    If InStr(1, resultText, "ROW 989 requires", vbTextCompare) = 0 _
       Or InStr(1, resultText, "only", vbTextCompare) = 0 _
       Or InStr(1, resultText, "available", vbTextCompare) = 0 Then
        failureReason = "Unexpected Boxes Made shortage validation result: " & resultText
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "USED")) <> 0 Then
        failureReason = "Validation staged component usage despite shortage."
        GoTo CleanExit
    End If

    TestShippingWorkflowGuard_BoxesMadeInsufficientComponentFails = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7117, "TestShippingWorkflowGuard_BoxesMadeInsufficientComponentFails", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingWorkflowGuard_ConfirmInventoryUseExistingWarns() As Long
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim resultText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    SetUseExistingInventoryForTest wbOps.Worksheets("ShipmentsTally"), True

    resultText = modShippingEventCreator.ValidateConfirmInventoryFromWorkbook(wbOps)
    If InStr(1, resultText, "Use existing inventory is enabled", vbTextCompare) = 0 Then
        failureReason = "Unexpected Confirm Inventory use-existing validation result: " & resultText
        GoTo CleanExit
    End If

    TestShippingWorkflowGuard_ConfirmInventoryUseExistingWarns = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7118, "TestShippingWorkflowGuard_ConfirmInventoryUseExistingWarns", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingAggregateBomMath_MultipliesComponentQtyByPackageQty() As Long
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loBomView As ListObject
    Dim loAggBom As ListObject
    Dim rowKraft As Long
    Dim rowTape As Long

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    Set loAggBom = FindTableByName(wbOps, "AggregateBoxBOM")
    If loInv Is Nothing Or loShip Is Nothing Or loBomView Is Nothing Or loAggBom Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 990, "SKU-T25", "T25", "EA", "A1", 10
    AddInvSysSeedRow loInv, 991, "SKU-KRAFT", "Kraft Paper", "EA", "A1", 100
    AddInvSysSeedRow loInv, 992, "SKU-TAPE", "Tape", "EA", "A1", 100
    AddShippingTallyRow loShip, "REF-BOM-001", "T25", 3, 990, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Warehouse"
    AddShippingBomViewRow loBomView, 990, "T25", 991, "Kraft Paper", 2, "EA"
    AddShippingBomViewRow loBomView, 990, "T25", 992, "Tape", 1, "EA"

    If Not modShippingEventCreator.RebuildShippingAggregatesForWorkbook(wbOps, report) Then
        failureReason = "RebuildShippingAggregatesForWorkbook failed: " & report
        GoTo CleanExit
    End If
    rowKraft = FindRowByColumnValueInTable(loAggBom, "ROW", "991")
    rowTape = FindRowByColumnValueInTable(loAggBom, "ROW", "992")
    If rowKraft = 0 Or rowTape = 0 Then
        failureReason = "AggregateBoxBOM did not include expected component rows."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loAggBom, rowKraft, "QUANTITY")) <> 6 Then
        failureReason = "Kraft Paper aggregate quantity was not 2 per box x 3 boxes."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loAggBom, rowTape, "QUANTITY")) <> 3 Then
        failureReason = "Tape aggregate quantity was not 1 per box x 3 boxes."
        GoTo CleanExit
    End If

    TestShippingAggregateBomMath_MultipliesComponentQtyByPackageQty = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7119, "TestShippingAggregateBomMath_MultipliesComponentQtyByPackageQty", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestBoxBuilderArchive_HidesArchivedBoxesUnlessRequested() As Long
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loBomView As ListObject
    Dim activeReport As String
    Dim archivedReport As String
    Dim allReport As String

    On Error GoTo CleanFail
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then
        failureReason = "EnsureShippingWorkbookSurface failed: " & report
        GoTo CleanExit
    End If
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loBomView Is Nothing Then
        failureReason = "ShippingBOMView was not created."
        GoTo CleanExit
    End If

    AddShippingBomViewRow loBomView, 990, "T25", 991, "Kraft Paper", 2, "EA"
    AddShippingBomViewRow loBomView, 991, "T26", 992, "Tape", 1, "EA"
    SetTableCell loBomView, 2, "IsActive", False
    SetOptionalTableCell loBomView, 2, "RetiredAtUTC", CDate("2026-06-18 12:00:00")

    wbOps.Activate
    activeReport = RunBoxBuilderSavedBoxesReportForTest(True, False)
    If InStr(1, activeReport, "COUNT=1", vbTextCompare) = 0 _
       Or InStr(1, activeReport, "BOX=T25", vbTextCompare) = 0 Then
        failureReason = "Active BoxBuilder list was wrong; expected only T25. " & activeReport
        GoTo CleanExit
    End If
    If InStr(1, activeReport, "BOX=T26", vbTextCompare) > 0 Then
        failureReason = "Active BoxBuilder list did not hide retired T26. " & activeReport
        GoTo CleanExit
    End If

    archivedReport = RunBoxBuilderSavedBoxesReportForTest(False, True)
    If InStr(1, archivedReport, "COUNT=1", vbTextCompare) = 0 _
       Or InStr(1, archivedReport, "BOX=T26", vbTextCompare) = 0 Then
        failureReason = "Archived BoxBuilder list was wrong; expected only retired T26. " & archivedReport
        GoTo CleanExit
    End If
    If InStr(1, archivedReport, "BOX=T25", vbTextCompare) > 0 Then
        failureReason = "Archived BoxBuilder list included active T25 when Show Active was off. " & archivedReport
        GoTo CleanExit
    End If

    allReport = RunBoxBuilderSavedBoxesReportForTest(True, True)
    If InStr(1, allReport, "COUNT=2", vbTextCompare) = 0 Then
        failureReason = "Archived BoxBuilder list did not include both active and retired designs. " & allReport
        GoTo CleanExit
    End If
    If InStr(1, allReport, "BOX=T26", vbTextCompare) = 0 Then
        failureReason = "Archived BoxBuilder list did not expose retired T26 for review/resurrection. " & allReport
        GoTo CleanExit
    End If

    TestBoxBuilderArchive_HidesArchivedBoxesUnlessRequested = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7142, "TestBoxBuilderArchive_HidesArchivedBoxesUnlessRequested", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestBoxBuilderForm_InitializesWithActiveArchiveFilters() As Long
    Dim report As String
    Dim failureReason As String

    On Error GoTo CleanFail
    If Not RunBoxBuilderInitializeSmokeForTest(report) Then
        failureReason = "BoxBuilder form initialization failed: " & report
        GoTo CleanExit
    End If

    TestBoxBuilderForm_InitializesWithActiveArchiveFilters = 1

CleanExit:
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7144, "TestBoxBuilderForm_InitializesWithActiveArchiveFilters", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingCommitLine_MergesPostedSameRefBoxVersionCarrier() As Long
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim ok As Boolean
    Dim matchCount As Long
    Dim rowIndex As Long

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If loInv Is Nothing Or loShip Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 987, "SKU-T25", "T25", "ea", "CLEARVIEW", 5
    AddShippingTallyRow loShip, "12", "T25", 1, 87, "ea", "CLEARVIEW", "v2"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "UPS"

    wbOps.Activate
    ok = RunShippingCommitLineForTest("SHIP", "ADD", 0, "12", "T25", 1, 87, "ea", "CLEARVIEW", "v2", "UPS", report)
    matchCount = CountShipmentRowsForTest(loShip, "12", "T25", "v2", "UPS")
    If matchCount <> 1 Then
        failureReason = "Expected same Ref/Box/Version/Carrier to merge into one row; found " & CStr(matchCount) & ". Commit result: " & CStr(ok) & "; " & report
        GoTo CleanExit
    End If

    rowIndex = FindShipmentRowForTest(loShip, "12", "T25", "v2", "UPS")
    If rowIndex = 0 Then
        failureReason = "Merged shipment row was not found."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loShip, rowIndex, "QUANTITY")) <> 2 Then
        failureReason = "Merged shipment quantity was not 2."
        GoTo CleanExit
    End If

    TestShippingCommitLine_MergesPostedSameRefBoxVersionCarrier = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7120, "TestShippingCommitLine_MergesPostedSameRefBoxVersionCarrier", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingBoard_TwoAddsSameRefBoxVersionCarrierShowOneRow() As Long
    Dim wbOps As Workbook
    Dim report As String
    Dim failureReason As String
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim ok As Boolean
    Dim matchCount As Long
    Dim rowIndex As Long

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If loInv Is Nothing Or loShip Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 989, "SKU-T25", "T25", "ea", "CLEARVIEW", 5
    AddShippingTallyRow loShip, "REF-BOARD-MERGE-001", "T25", 1, 989, "ea", "CLEARVIEW", "v2"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "UPS"

    wbOps.Activate
    ok = RunShippingCommitLineForTest("SHIP", "ADD", 0, "REF-BOARD-MERGE-001", "T25", 1, 989, "ea", "CLEARVIEW", "v2", "UPS", report)

    matchCount = CountShipmentRowsForTest(loShip, "REF-BOARD-MERGE-001", "T25", "v2", "UPS")
    If matchCount <> 1 Then
        failureReason = "Shipments board should show one row for same Ref/Box/Version/Carrier; found " & CStr(matchCount) & ". Commit result: " & CStr(ok) & "; " & report
        GoTo CleanExit
    End If

    rowIndex = FindShipmentRowForTest(loShip, "REF-BOARD-MERGE-001", "T25", "v2", "UPS")
    If rowIndex = 0 Then
        failureReason = "Merged shipment board row was not found."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loShip, rowIndex, "QUANTITY")) <> 2 Then
        failureReason = "Merged shipment board quantity was not 2. Commit result: " & CStr(ok) & "; " & report
        GoTo CleanExit
    End If

    TestShippingBoard_TwoAddsSameRefBoxVersionCarrierShowOneRow = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7123, "TestShippingBoard_TwoAddsSameRefBoxVersionCarrierShowOneRow", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingAdd_DefaultsOrderToWarehouseArea() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loBomView As ListObject
    Dim ok As Boolean

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loShip Is Nothing Or loBomView Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 984, "SKU-ADD-LOCKED", "Add Locked Item", "EA", "A1", 5
    AddShippingBomViewRow loBomView, 984, "Add Locked Item", 984, "Add Locked Item", 1, "EA"

    wbOps.Activate
    ok = RunShippingCommitLineForTest("SHIP", "ADD", 0, "REF-ADD-LOCKED", "Add Locked Item", 1, 984, "EA", "A1", "v1", "DHL", report)
    If Not ok Then
        If Trim$(CStr(GetTableValue(loShip, 1, "ITEMS"))) = "" Then
            failureReason = "Add did not create a visible order row. Result: " & CStr(ok) & "; " & report
            GoTo CleanExit
        End If
    End If
    If StrComp(Trim$(CStr(GetTableValue(loShip, 1, "AREA"))), "Warehouse", vbTextCompare) <> 0 Then
        failureReason = "Add should leave the order in Warehouse area until To Shipments."
        GoTo CleanExit
    End If

    TestShippingAdd_DefaultsOrderToWarehouseArea = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7125, "TestShippingAdd_DefaultsOrderToWarehouseArea", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingAdd_BlankCarrierRequiresCarrier() As Long
    Dim report As String
    Dim failureReason As String

    On Error GoTo CleanFail
    report = RunShippingValidateCommitInputsReportForTest("SHIP", "ADD", "Carrier Required Item", 1, 986, "")
    If InStr(1, report, "Select a Carrier", vbTextCompare) = 0 Then
        failureReason = "Blank Carrier did not return the expected user message. Report: " & report
        GoTo CleanExit
    End If

    TestShippingAdd_BlankCarrierRequiresCarrier = 1

CleanExit:
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7127, "TestShippingAdd_BlankCarrierRequiresCarrier", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingAdd_UsesDisplayedProjectedInventoryWhenVersionLedgerIsEmpty() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loBomView As ListObject
    Dim ok As Boolean

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loShip Is Nothing Or loBomView Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 89, "SKU-T28", "T28", "EA", "CLEARVIEW", 20
    AddShippingBomViewRow loBomView, 89, "T28", 900, "T28 component", 7, "EA"
    AddShippingBomViewRow loBomView, 89, "T28", 900, "T28 component", 7, "EA"
    SetOptionalTableCell loBomView, 2, "BomVersionLabel", "v2"

    wbOps.Activate
    ok = RunShippingCommitLineForTest("SHIP", _
                                      "ADD", _
                                      0, _
                                      "30", _
                                      "T28", _
                                      2, _
                                      89, _
                                      "ea", _
                                      "CLEARVIEW", _
                                      "v1", _
                                      "UPS", _
                                      report)
    If InStr(1, report, "requires 2", vbTextCompare) > 0 _
       Or InStr(1, report, "only 0", vbTextCompare) > 0 Then
        failureReason = "Shipping Add ignored displayed Projected Inv and reported a false version shortage. Result: " & CStr(ok) & "; " & report
        GoTo CleanExit
    End If
    If Trim$(CStr(GetTableValue(loShip, 1, "ITEMS"))) <> "T28" _
       Or CDbl(GetTableValue(loShip, 1, "QUANTITY")) <> 2 Then
        failureReason = "Shipping Add did not keep the visible T28 order row after using displayed Projected Inv. Result: " & CStr(ok) & "; " & report
        GoTo CleanExit
    End If

    TestShippingAdd_UsesDisplayedProjectedInventoryWhenVersionLedgerIsEmpty = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7150, "TestShippingAdd_UsesDisplayedProjectedInventoryWhenVersionLedgerIsEmpty", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingAdd_UsesDisplayedProjectedInventoryWhenTotalInvIsStaleZero() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loBomView As ListObject
    Dim resultText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loShip Is Nothing Or loBomView Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 89, "SKU-T28-STALE", "T28", "EA", "CLEARVIEW", 0
    AddShippingBomViewRow loBomView, 89, "T28", 900, "T28 component", 7, "EA"
    AddShippingBomViewRow loBomView, 89, "T28", 900, "T28 component", 7, "EA"
    SetOptionalTableCell loBomView, 2, "BomVersionLabel", "v2"
    AddShippingTallyRow loShip, "31", "T28", 1, 89, "ea", "CLEARVIEW", "v1"
    SetTableCell loShip, 1, "AREA", "Warehouse"
    SetTableCell loShip, 1, "CARRIER", "UPS"

    wbOps.Activate
    resultText = RunShippingProjectedAvailabilityOverrideForTest(1, 89, "v1", 20)
    If StrComp(resultText, "OK", vbTextCompare) <> 0 Then
        failureReason = "Shipping Add rejected visible Projected Inv when invSys TOTAL INV was stale zero. Report: " & resultText
        GoTo CleanExit
    End If
    If InStr(1, resultText, "only 0", vbTextCompare) > 0 _
       Or InStr(1, resultText, "TOTAL INV", vbTextCompare) > 0 Then
        failureReason = "Shipping Add still reported the stale TOTAL INV shortage. Report: " & resultText
        GoTo CleanExit
    End If
    If Trim$(CStr(GetTableValue(loShip, 1, "ITEMS"))) <> "T28" _
       Or CDbl(GetTableValue(loShip, 1, "QUANTITY")) <> 1 Then
        failureReason = "Shipping Add did not keep the visible T28 order row. Result: " & resultText
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 19 Then
        failureReason = "Local projected TOTAL INV was not repaired from displayed availability before locking; expected 19 but found " & CStr(GetTableValue(loInv, 1, "TOTAL INV")) & "."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 1 Then
        failureReason = "Local SHIPMENTS lock did not increase to 1."
        GoTo CleanExit
    End If

    TestShippingAdd_UsesDisplayedProjectedInventoryWhenTotalInvIsStaleZero = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7151, "TestShippingAdd_UsesDisplayedProjectedInventoryWhenTotalInvIsStaleZero", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingRemove_LockedRowReleasesInventory() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loBomView As ListObject
    Dim ok As Boolean
    Dim overlayPath As String
    Dim projectedText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loShip Is Nothing Or loBomView Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 987, "SKU-REMOVE-LOCKED", "Remove Locked Item", "EA", "A1", 4
    AddShippingBomViewRow loBomView, 987, "Remove Locked Item", 987, "Remove Locked Item", 1, "EA"
    SetTableCell loInv, 1, "SHIPMENTS", 1
    AddShippingTallyRow loShip, "REF-REMOVE-LOCKED", "Remove Locked Item", 1, 987, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Warehouse"
    SetTableCell loShip, 1, "CARRIER", "DHL"
    SetTableCell loShip, 1, "LINE_ID", "SHIPLINE-REMOVE-LOCKED-001"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-REMOVE-LOCKED-001"

    wbOps.Activate
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingRegisterProjectedOverlayForTest 987, "v1", 4, 5
    ok = RunShippingCommitLineForTest("SHIP", "DELETE", 1, "", "", 0, 0, "", "", "", "", report)
    If Not ok Then
        failureReason = "Remove locked row failed: " & report
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 5 Then
        failureReason = "Remove did not release locked inventory back to TOTAL INV."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 0 Then
        failureReason = "Remove did not clear locked SHIPMENTS staging."
        GoTo CleanExit
    End If
    If Not loShip.DataBodyRange Is Nothing Then
        If loShip.ListRows.Count > 0 Then
            If Trim$(CStr(GetTableValue(loShip, 1, "ITEMS"))) <> "" Then
                failureReason = "Remove left the locked shipment row visible."
                GoTo CleanExit
            End If
        End If
    End If
    projectedText = RunShippingProjectedOverlayTextForTest(987, "v1", "5")
    If CDbl(NzDblForTest(projectedText)) <> 5 Then
        failureReason = "Remove released inventory but left Projected Inv overlay deducted; expected 5 but found " & projectedText & "."
        GoTo CleanExit
    End If

    TestShippingRemove_LockedRowReleasesInventory = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7128, "TestShippingRemove_LockedRowReleasesInventory", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingRemove_StaleLockedRowClearsWithoutInflatingInventory() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loBomView As ListObject
    Dim ok As Boolean
    Dim overlayPath As String
    Dim projectedText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loShip Is Nothing Or loBomView Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 986, "SKU-REMOVE-STALE-LOCK", "Stale Locked Item", "EA", "A1", 5
    AddShippingBomViewRow loBomView, 986, "Stale Locked Item", 986, "Stale Locked Item", 1, "EA"
    SetTableCell loInv, 1, "SHIPMENTS", 0
    AddShippingTallyRow loShip, "REF-REMOVE-STALE-LOCK", "Stale Locked Item", 1, 986, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Warehouse"
    SetTableCell loShip, 1, "CARRIER", "DHL"
    SetTableCell loShip, 1, "LINE_ID", "SHIPLINE-REMOVE-STALE-LOCK-001"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-REMOVE-STALE-LOCK-001"

    wbOps.Activate
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingRegisterProjectedOverlayForTest 986, "v1", 4, 5
    ok = RunShippingCommitLineForTest("SHIP", "DELETE", 1, "", "", 0, 0, "", "", "", "", report)
    If Not ok Then
        failureReason = "Remove stale locked row failed: " & report
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 5 Then
        failureReason = "Remove stale locked row inflated TOTAL INV."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 0 Then
        failureReason = "Remove stale locked row changed SHIPMENTS staging."
        GoTo CleanExit
    End If
    If Not loShip.DataBodyRange Is Nothing Then
        If loShip.ListRows.Count > 0 Then
            If Trim$(CStr(GetTableValue(loShip, 1, "ITEMS"))) <> "" Then
                failureReason = "Remove left the stale locked shipment row visible."
                GoTo CleanExit
            End If
        End If
    End If
    projectedText = RunShippingProjectedOverlayTextForTest(986, "v1", "5")
    If CDbl(NzDblForTest(projectedText)) <> 5 Then
        failureReason = "Remove stale locked row left a deducted Projected Inv overlay; expected 5 but found " & projectedText & "."
        GoTo CleanExit
    End If

    TestShippingRemove_StaleLockedRowClearsWithoutInflatingInventory = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7129, "TestShippingRemove_StaleLockedRowClearsWithoutInflatingInventory", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingHold_PreservesReservationAndLocalDeduction() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loHold As ListObject
    Dim loBomView As ListObject
    Dim selectedRows(1 To 1) As Long
    Dim ok As Boolean

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loHold = FindTableByName(wbOps, "NotShipped")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loShip Is Nothing Or loHold Is Nothing Or loBomView Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 974, "SKU-HOLD-LOCKED", "Hold Locked Item", "EA", "A1", 12
    AddShippingBomViewRow loBomView, 974, "Hold Locked Item", 974, "Hold Locked Item", 1, "EA"
    SetTableCell loInv, 1, "SHIPMENTS", 1
    AddShippingTallyRow loShip, "REF-HOLD-LOCKED", "Hold Locked Item", 1, 974, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "UPS"
    SetTableCell loShip, 1, "LINE_ID", "SHIPLINE-HOLD-LOCKED-001"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-HOLD-LOCKED-001"

    wbOps.Activate
    selectedRows(1) = 1
    ok = RunShippingMoveHoldRowsForTest(selectedRows, True, report)
    If Not ok Then
        failureReason = "Send Hold failed: " & report
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 12 Then
        failureReason = "Send Hold restored TOTAL INV; hold should preserve the local deduction."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 1 Then
        failureReason = "Send Hold cleared SHIPMENTS staging; hold should remain reserved."
        GoTo CleanExit
    End If
    If loHold.ListRows.Count <> 1 Then
        failureReason = "Send Hold did not move exactly one row into NotShipped."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loHold, 1, "SERVER_RESERVE_EVENT_ID")), "RESERVE-HOLD-LOCKED-001", vbTextCompare) <> 0 Then
        failureReason = "Send Hold did not preserve SERVER_RESERVE_EVENT_ID."
        GoTo CleanExit
    End If

    ok = RunShippingMoveHoldRowsForTest(selectedRows, False, report)
    If Not ok Then
        failureReason = "Return from Hold failed: " & report
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 12 Then
        failureReason = "Return from Hold changed TOTAL INV; reservation should already be active."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 1 Then
        failureReason = "Return from Hold changed SHIPMENTS staging; reservation should already be active."
        GoTo CleanExit
    End If
    If loShip.ListRows.Count <> 1 Then
        failureReason = "Return from Hold did not move exactly one row back to Shipments."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loShip, 1, "SERVER_RESERVE_EVENT_ID")), "RESERVE-HOLD-LOCKED-001", vbTextCompare) <> 0 Then
        failureReason = "Return from Hold did not preserve SERVER_RESERVE_EVENT_ID."
        GoTo CleanExit
    End If

    TestShippingHold_PreservesReservationAndLocalDeduction = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7156, "TestShippingHold_PreservesReservationAndLocalDeduction", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingUpdate_PreservesExistingReservationWithoutDoubleDeducting() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim loBomView As ListObject
    Dim ok As Boolean
    Dim overlayPath As String
    Dim projectedText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loShip Is Nothing Or loBomView Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 973, "SKU-UPDATE-LOCKED", "Update Locked Item", "EA", "A1", 19
    AddShippingBomViewRow loBomView, 973, "Update Locked Item", 973, "Update Locked Item", 1, "EA"
    SetTableCell loInv, 1, "SHIPMENTS", 1
    AddShippingTallyRow loShip, "REF-UPDATE-LOCKED", "Update Locked Item", 1, 973, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "UPS"
    SetTableCell loShip, 1, "LINE_ID", "SHIPLINE-UPDATE-LOCKED-001"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-UPDATE-LOCKED-001"

    wbOps.Activate
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingRegisterProjectedOverlayForTest 973, "v1", 19, 20

    ok = RunShippingCommitLineForTest("SHIP", "UPDATE", 1, "REF-UPDATE-LOCKED", "Update Locked Item", 1, 973, "EA", "A1", "v1", "DHL", report, 19)
    If Not ok Then
        failureReason = "Update reserved row failed: " & report
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 19 Then
        failureReason = "Update reserved row changed TOTAL INV; expected 19 but found " & CStr(GetTableValue(loInv, 1, "TOTAL INV")) & "."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 1 Then
        failureReason = "Update reserved row changed SHIPMENTS staging; expected 1 but found " & CStr(GetTableValue(loInv, 1, "SHIPMENTS")) & "."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loShip, 1, "SERVER_RESERVE_EVENT_ID")), "RESERVE-UPDATE-LOCKED-001", vbTextCompare) <> 0 Then
        failureReason = "Update reserved row did not preserve SERVER_RESERVE_EVENT_ID."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loShip, 1, "AREA")), "Shipments", vbTextCompare) <> 0 Then
        failureReason = "Update reserved row did not preserve Shipments area."
        GoTo CleanExit
    End If
    projectedText = RunShippingProjectedOverlayTextForTest(973, "v1", "20")
    If CDbl(NzDblForTest(projectedText)) <> 19 Then
        failureReason = "Update reserved row double-deducted Projected Inv; expected 19 but found " & projectedText & "."
        GoTo CleanExit
    End If

    TestShippingUpdate_PreservesExistingReservationWithoutDoubleDeducting = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7158, "TestShippingUpdate_PreservesExistingReservationWithoutDoubleDeducting", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingSentRows_ReservedRowDoesNotAddBackTotalInv() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim selectedRows(1 To 1) As Long
    Dim ok As Boolean

    rootPath = BuildRuntimeTestRoot("phase6_ship_sent_reserved_no_addback")
    currentUser = "calvin"

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If loInv Is Nothing Or loShip Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 981, "SKU-SENT-RESERVED", "Reserved Sent Item", "EA", "A1", 4
    SetTableCell loInv, 1, "SHIPMENTS", 1
    AddShippingTallyRow loShip, "REF-SENT-RESERVED", "Reserved Sent Item", 1, 981, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "UPS"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-SENT-001"

    selectedRows(1) = 1
    wbOps.Activate
    ok = RunShippingApplySentRowsInventoryForTest(selectedRows, report)
    If Not ok Then
        failureReason = "Shipments Sent reserved-row macro failed: " & report
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 4 Then
        failureReason = "Reserved Shipments Sent changed TOTAL INV; expected it to stay at already-deducted 4."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 0 Then
        failureReason = "Reserved Shipments Sent did not clear SHIPMENTS staging."
        GoTo CleanExit
    End If

    TestShippingSentRows_ReservedRowDoesNotAddBackTotalInv = 1

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH98"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen FindWorkbookByName("WH98.invSys.Auth.xlsb")
    CloseWorkbookIfOpen FindWorkbookByName("WH98.invSys.Config.xlsb")
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7121, "TestShippingSentRows_ReservedRowDoesNotAddBackTotalInv", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingSentRows_UnreservedDirtyRowDeductsTotalInv() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim selectedRows(1 To 1) As Long
    Dim ok As Boolean

    rootPath = BuildRuntimeTestRoot("phase6_ship_sent_unreserved_deduct")
    currentUser = "calvin"

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If loInv Is Nothing Or loShip Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 982, "SKU-SENT-DIRTY", "Dirty Sent Item", "EA", "A1", 5
    SetTableCell loInv, 1, "SHIPMENTS", 1
    AddShippingTallyRow loShip, "REF-SENT-DIRTY", "Dirty Sent Item", 1, 982, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "UPS"

    selectedRows(1) = 1
    wbOps.Activate
    ok = RunShippingApplySentRowsInventoryForTest(selectedRows, report)
    If Not ok Then
        failureReason = "Shipments Sent unreserved-row macro failed: " & report
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "TOTAL INV")) <> 4 Then
        failureReason = "Unreserved Shipments Sent did not deduct TOTAL INV; expected 4."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 0 Then
        failureReason = "Unreserved Shipments Sent did not clear SHIPMENTS staging."
        GoTo CleanExit
    End If

    TestShippingSentRows_UnreservedDirtyRowDeductsTotalInv = 1

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH99"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen FindWorkbookByName("WH99.invSys.Auth.xlsb")
    CloseWorkbookIfOpen FindWorkbookByName("WH99.invSys.Config.xlsb")
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7122, "TestShippingSentRows_UnreservedDirtyRowDeductsTotalInv", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingSentRows_ReservedRowClearsLockedReservationTotal() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim selectedRows(1 To 1) As Long
    Dim ok As Boolean
    Dim totals As Object
    Dim key As String
    Dim reserveEventId As String
    Dim lineId As String

    rootPath = BuildRuntimeTestRoot("phase6_ship_sent_clears_locked")
    currentUser = "calvin"
    reserveEventId = "RESERVE-SENT-LOCKED-001"
    lineId = "SHIPLINE-SENT-LOCKED-001"

    On Error GoTo CleanFail
    If Not PrepareShippingPostSessionForTest(rootPath, "WH100", "S31", currentUser, failureReason) Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If loInv Is Nothing Or loShip Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 983, "SKU-SENT-LOCKED", "Locked Sent Item", "EA", "A1", 8
    SetTableCell loInv, 1, "SHIPMENTS", 2
    AddShippingTallyRow loShip, "REF-SENT-LOCKED", "Locked Sent Item", 2, 983, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "UPS"
    SetTableCell loShip, 1, "LINE_ID", lineId
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", reserveEventId
    CreateShippingReservationLedgerForTest rootPath, "WH100", reserveEventId, lineId, "REF-SENT-LOCKED", "Locked Sent Item", 983, "v1", 2, "EA", "A1"

    selectedRows(1) = 1
    wbOps.Activate
    ok = RunShippingCompleteSentRowsForTest(selectedRows, report)
    If Not ok Then
        failureReason = "Shipments Sent reserved-row completion failed: " & report
        GoTo CleanExit
    End If

    Set totals = RunShippingReservationTotalsForTest()
    key = "983|v1"
    If Not totals Is Nothing Then
        If totals.Exists(key) Then
            If CDbl(totals(key)) <> 0 Then
                failureReason = "Reserved shipment still contributes to Locked after Shipments Sent; expected 0 but found " & CStr(totals(key)) & "."
                GoTo CleanExit
            End If
        End If
    End If
    If CDbl(GetTableValue(loInv, 1, "SHIPMENTS")) <> 0 Then
        failureReason = "Reserved Shipments Sent did not clear SHIPMENTS staging."
        GoTo CleanExit
    End If

    TestShippingSentRows_ReservedRowClearsLockedReservationTotal = 1

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH100"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen FindWorkbookByName("WH100.invSys.Data.ShippingReservations.xlsb")
    CloseWorkbookIfOpen FindWorkbookByName("WH100.invSys.Auth.xlsb")
    CloseWorkbookIfOpen FindWorkbookByName("WH100.invSys.Config.xlsb")
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7124, "TestShippingSentRows_ReservedRowClearsLockedReservationTotal", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim selectedRows(1 To 1) As Long
    Dim ok As Boolean
    Dim projectedText As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If loInv Is Nothing Or loShip Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 988, "SKU-PROJECTED-SENT", "Projected Sent Item", "EA", "A1", 4
    SetTableCell loInv, 1, "SHIPMENTS", 1
    AddShippingTallyRow loShip, "REF-PROJECTED-SENT", "Projected Sent Item", 1, 988, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "DHL"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-PROJECTED-SENT-001"

    wbOps.Activate
    RunShippingRegisterProjectedOverlayForTest 988, "v1", 3, 4
    selectedRows(1) = 1
    ok = RunShippingCompleteSentRowsForTest(selectedRows, report)
    If Not ok Then
        failureReason = "Shipments Sent completion failed: " & report
        GoTo CleanExit
    End If

    projectedText = RunShippingProjectedOverlayTextForTest(988, "v1", "4")
    If CDbl(NzDblForTest(projectedText)) > 3.0000001 Then
        failureReason = "Shipments Sent increased projected inventory overlay; expected 3 or less but found " & projectedText & "."
        GoTo CleanExit
    End If

    TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7129, "TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingSentRows_ReservedCompletionKeepsProjectedDeductionWhenNasStale() As Long
    Dim failureReason As String
    Dim projectedQty As Double

    On Error GoTo CleanFail
    projectedQty = RunShippingSentProjectedOverlayQtyForTest(19, 19, 1, True)
    If projectedQty <> 19 Then
        failureReason = "Reserved Shipments Sent double-subtracted an already-projected shipment; expected 19 but found " & CStr(projectedQty) & "."
        GoTo CleanExit
    End If
    projectedQty = RunShippingSentProjectedOverlayQtyForTest(19, 19, 1, False)
    If projectedQty <> 18 Then
        failureReason = "Unprojected Shipments Sent did not deduct once; expected 18 but found " & CStr(projectedQty) & "."
        GoTo CleanExit
    End If
    projectedQty = RunShippingSentProjectedOverlayQtyForTest(19, 19, 1, False, True)
    If projectedQty <> 19 Then
        failureReason = "Reserved Shipments Sent double-subtracted when the active overlay was missing; expected 19 but found " & CStr(projectedQty) & "."
        GoTo CleanExit
    End If
    projectedQty = RunShippingSentProjectedOverlayQtyForTest(0, 19, 1, True, True)
    If projectedQty <> 19 Then
        failureReason = "Reserved Shipments Sent rewrote an existing positive projection to zero when backend lookup was stale/blank; expected 19 but found " & CStr(projectedQty) & "."
        GoTo CleanExit
    End If

    TestShippingSentRows_ReservedCompletionKeepsProjectedDeductionWhenNasStale = 1

CleanExit:
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7152, "TestShippingSentRows_ReservedCompletionKeepsProjectedDeductionWhenNasStale", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingSentRows_FullRunNeverIncreasesProjectedInventory() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loShip As ListObject
    Dim selectedRows(1 To 1) As Long
    Dim runResult As String
    Dim projectedAfter As Double
    Dim projectedAfterPeerSent As Double
    Dim projectedAfterCatchup As Double
    Dim projectedAfterEviction As Double
    Dim overlayPath As String

    rootPath = BuildRuntimeTestRoot("phase6_ship_sent_full_never_adds_projected")
    currentUser = "calvin"

    On Error GoTo CleanFail
    If Not PrepareShippingPostSessionForTest(rootPath, "WH101", "S31", currentUser, failureReason) Then GoTo CleanExit

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loShip = FindTableByName(wbOps, "ShipmentsTally")
    If loInv Is Nothing Or loShip Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 985, "SKU-SENT-FULL-PROJECTED", "Full Sent Projected Item", "EA", "A1", 10
    SetTableCell loInv, 1, "TOTAL INV", 9
    SetTableCell loInv, 1, "SHIPMENTS", 1
    AddShippingTallyRow loShip, "REF-SENT-FULL-PROJECTED", "Full Sent Projected Item", 1, 985, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "USPS"
    SetTableCell loShip, 1, "LINE_ID", "SHIPLINE-SENT-FULL-PROJECTED-001"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-SENT-FULL-PROJECTED-001"

    wbOps.Activate
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    If CDbl(NzDblForTest(RunShippingProjectedOverlayTextForTest(985, "v1", "10"))) <> 10 Then
        failureReason = "Test setup expected no projected overlay before Shipments Sent."
        GoTo CleanExit
    End If
    selectedRows(1) = 1
    runResult = RunShippingSentRowsReportForTest(selectedRows, "USPS")
    If Left$(runResult, 3) <> "OK|" Then
        failureReason = "Full Shipments Sent run failed: " & Mid$(runResult, 6)
        GoTo CleanExit
    End If
    If InStr(1, runResult, "already reserved at To Shipments", vbTextCompare) > 0 Then
        failureReason = "Full Shipments Sent returned the old confusing already-reserved popup text: " & runResult
        GoTo CleanExit
    End If
    If InStr(1, runResult, "completed the reservation", vbTextCompare) = 0 Then
        failureReason = "Full Shipments Sent did not explain that reserved rows complete an existing server reservation: " & runResult
        GoTo CleanExit
    End If

    projectedAfter = NzDblForTest(RunShippingProjectedOverlayTextForTest(985, "v1", "10"))
    If projectedAfter <> 9 Then
        failureReason = "Full Shipments Sent did not preserve the user-side projected deduction against stale NAS inventory; expected 9 but found " & CStr(projectedAfter) & "."
        GoTo CleanExit
    End If
    AddInvSysSeedRow loInv, 984, "SKU-SENT-FULL-PEER", "Full Sent Peer Item", "EA", "A1", 10
    SetTableCell loInv, 2, "TOTAL INV", 8
    SetTableCell loInv, 2, "SHIPMENTS", 2
    AddShippingTallyRow loShip, "REF-SENT-FULL-PEER", "Full Sent Peer Item", 2, 984, "EA", "A1", "v1"
    SetTableCell loShip, 1, "AREA", "Shipments"
    SetTableCell loShip, 1, "CARRIER", "USPS"
    SetTableCell loShip, 1, "LINE_ID", "SHIPLINE-SENT-FULL-PEER-001"
    SetTableCell loShip, 1, "SERVER_RESERVE_EVENT_ID", "RESERVE-SENT-FULL-PEER-001"
    selectedRows(1) = 1
    runResult = RunShippingSentRowsReportForTest(selectedRows, "USPS")
    If Left$(runResult, 3) <> "OK|" Then
        failureReason = "Peer Shipments Sent run failed: " & Mid$(runResult, 6)
        GoTo CleanExit
    End If
    projectedAfterPeerSent = NzDblForTest(RunShippingProjectedOverlayTextForTest(985, "v1", "10"))
    If projectedAfterPeerSent <> 9 Then
        failureReason = "Completing peer shipment cleared the prior SENT overlay; expected T28 projected 9 but found " & CStr(projectedAfterPeerSent) & "."
        GoTo CleanExit
    End If
    projectedAfterCatchup = NzDblForTest(RunShippingProjectedOverlayTextForTest(985, "v1", "9"))
    If projectedAfterCatchup <> 9 Then
        failureReason = "Full Shipments Sent did not return NAS catch-up value after backend deducted to 9; found " & CStr(projectedAfterCatchup) & "."
        GoTo CleanExit
    End If
    projectedAfterEviction = NzDblForTest(RunShippingProjectedOverlayTextForTest(985, "v1", "10"))
    If projectedAfterEviction <> 10 Then
        failureReason = "Full Shipments Sent left an active-lock overlay after SENT overlay catch-up eviction; expected backend 10 but found " & CStr(projectedAfterEviction) & "."
        GoTo CleanExit
    End If

    TestShippingSentRows_FullRunNeverIncreasesProjectedInventory = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH101"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen FindWorkbookByName("WH101.invSys.Data.ShippingReservations.xlsb")
    CloseWorkbookIfOpen FindWorkbookByName("WH101.invSys.Auth.xlsb")
    CloseWorkbookIfOpen FindWorkbookByName("WH101.invSys.Config.xlsb")
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7130, "TestShippingSentRows_FullRunNeverIncreasesProjectedInventory", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingProjectedOverlay_PreservesNasBaselineAcrossSentReregister() As Long
    Dim failureReason As String
    Dim overlayPath As String
    Dim projectedText As String

    On Error GoTo CleanFail
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath

    RunShippingRegisterProjectedOverlayForTest 976, "v1", 19, 20
    RunShippingRegisterProjectedOverlayForTest 976, "v1", 19, 19
    RunShippingClearProjectedOverlayForTest
    projectedText = RunShippingProjectedOverlayTextForTest(976, "v1", "20")
    If CDbl(NzDblForTest(projectedText)) <> 19 Then
        failureReason = "Shipments Sent re-registered the overlay with the local 19 baseline and let stale NAS 20 inflate Projected Inv; found " & projectedText & "."
        GoTo CleanExit
    End If

    TestShippingProjectedOverlay_PreservesNasBaselineAcrossSentReregister = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7154, "TestShippingProjectedOverlay_PreservesNasBaselineAcrossSentReregister", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingProjectedOverlay_EvictsStaleZeroWhenBackendPositive() As Long
    Dim failureReason As String
    Dim overlayPath As String
    Dim projectedText As String

    On Error GoTo CleanFail
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath

    RunShippingRegisterProjectedOverlayForTest 975, "v1", 0, 0
    RunShippingClearProjectedOverlayForTest
    projectedText = RunShippingProjectedOverlayTextForTest(975, "v1", "20")
    If CDbl(NzDblForTest(projectedText)) <> 20 Then
        failureReason = "Stale zero projected overlay overrode positive backend inventory; expected 20 but found " & projectedText & "."
        GoTo CleanExit
    End If

    TestShippingProjectedOverlay_EvictsStaleZeroWhenBackendPositive = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7157, "TestShippingProjectedOverlay_EvictsStaleZeroWhenBackendPositive", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp() As Long
    Dim failureReason As String
    Dim overlayPath As String
    Dim projectedText As String

    On Error GoTo CleanFail
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath

    RunShippingRegisterProjectedOverlayForTest 979, "v1", 9, 10
    RunShippingClearProjectedOverlayForTest
    projectedText = RunShippingProjectedOverlayTextForTest(979, "v1", "10")
    If CDbl(NzDblForTest(projectedText)) <> 9 Then
        failureReason = "Projected overlay did not persist across restart; expected 9 with stale NAS 10 but found " & projectedText & "."
        GoTo CleanExit
    End If

    projectedText = RunShippingProjectedOverlayTextForTest(979, "v1", "9")
    If CDbl(NzDblForTest(projectedText)) <> 9 Then
        failureReason = "Projected overlay did not return NAS value after backend caught up."
        GoTo CleanExit
    End If
    RunShippingClearProjectedOverlayForTest
    projectedText = RunShippingProjectedOverlayTextForTest(979, "v1", "10")
    If CDbl(NzDblForTest(projectedText)) <> 9 Then
        failureReason = "Projected overlay was cleared by a local backend catch-up and allowed stale NAS 10 to inflate projected inventory; found " & projectedText & "."
        GoTo CleanExit
    End If

    TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7133, "TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas() As Long
    Dim failureReason As String
    Dim overlayPath As String
    Dim projectedText As String

    On Error GoTo CleanFail
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath

    RunShippingRegisterProjectedOverlayForTest 978, "v1", 9, 10
    projectedText = RunShippingProjectedOverlayTextForTest(978, "v1", "9")
    If CDbl(NzDblForTest(projectedText)) <> 9 Then
        failureReason = "Projected overlay did not show local deducted value when backend also read 9."
        GoTo CleanExit
    End If

    RunShippingClearProjectedOverlayForTest
    projectedText = RunShippingProjectedOverlayTextForTest(978, "v1", "10")
    If CDbl(NzDblForTest(projectedText)) <> 9 Then
        failureReason = "Projected overlay was cleared by local catch-up and allowed stale NAS 10 to inflate projected inventory; found " & projectedText & "."
        GoTo CleanExit
    End If

    TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7135, "TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingProjectedOverlay_ClearsWhenBackendRisesAboveBaseline() As Long
    Dim failureReason As String
    Dim overlayPath As String
    Dim projectedText As String

    On Error GoTo CleanFail
    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath

    RunShippingRegisterProjectedOverlayForTest 977, "v1", 4, 4
    RunShippingClearProjectedOverlayForTest
    projectedText = RunShippingProjectedOverlayTextForTest(977, "v1", "24")
    If CDbl(NzDblForTest(projectedText)) <> 24 Then
        failureReason = "Stale shipment projected overlay masked a backend inventory increase; expected 24 but found " & projectedText & "."
        GoTo CleanExit
    End If

    RunShippingClearProjectedOverlayForTest
    projectedText = RunShippingProjectedOverlayTextForTest(977, "v1", "24")
    If CDbl(NzDblForTest(projectedText)) <> 24 Then
        failureReason = "Cleared stale overlay did not persist after reload."
        GoTo CleanExit
    End If

    TestShippingProjectedOverlay_ClearsWhenBackendRisesAboveBaseline = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7145, "TestShippingProjectedOverlay_ClearsWhenBackendRisesAboveBaseline", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim eventIdOut As String
    Dim payloadJson As String
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim wbSnap As Workbook
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loBomView As ListObject
    Dim invRow As Long
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim savedBoxes(1 To 1, 1 To 7) As Variant
    Dim shippables As Variant

    rootPath = BuildRuntimeTestRoot("phase6_ship_reserve_refresh_nas")
    currentUser = "calvin"

    On Error GoTo CleanFail
    If Not PrepareShippingPostSessionForTest(rootPath, "WH102", "S31", currentUser, failureReason) Then GoTo CleanExit

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH102", Array("SKU-SHIP-RESERVE-CATCHUP"))
    Set wbInbox = CreateCanonicalShipInboxWorkbookForTest(rootPath, "S31")
    If wbInv Is Nothing Or wbInbox Is Nothing Then
        failureReason = "Canonical shipping runtime workbooks could not be created."
        GoTo CleanExit
    End If
    Set evt = CreateReceiveEventForTest("EVT-SHIP-RESERVE-CATCHUP-SEED", "WH102", "S31", currentUser, "SKU-SHIP-RESERVE-CATCHUP", 10, "A1", "shipping reserve catch-up seed")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-SHIP-RESERVE-CATCHUP-SEED", statusOut, errorCode, errorMessage) Then
        failureReason = "Canonical shipping seed event failed: " & errorCode & "; " & errorMessage
        GoTo CleanExit
    End If

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH102", "SKU-SHIP-RESERVE-CATCHUP", 10, CDate("2026-03-25 12:15:00"), 10, "A1=10", "T27", "EA", "A1", "", "", "", "", "986")
    If wbSnap Is Nothing Then
        failureReason = "Stale snapshot workbook could not be created."
        GoTo CleanExit
    End If
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    payloadJson = modRoleEventWriter.BuildPayloadJson( _
        modRoleEventWriter.CreatePayloadItem(986, _
                                             "SKU-SHIP-RESERVE-CATCHUP", _
                                             1, _
                                             "A1", _
                                             "v1"))
    If Not modRoleEventWriter.QueuePayloadEvent(CORE_EVENT_TYPE_SHIP_RESERVE, "WH102", "S31", currentUser, payloadJson, "reserve-refresh-nas", "", "", Now, wbInbox, eventIdOut, report) Then
        failureReason = "QueuePayloadEvent failed for shipping reserve: " & report
        GoTo CleanExit
    End If

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then
        failureReason = "EnsureShippingWorkbookSurface failed: " & report
        GoTo CleanExit
    End If
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loBomView Is Nothing Then
        failureReason = "Shipping operator surface tables were missing."
        GoTo CleanExit
    End If
    AddInvSysSeedRow loInv, 986, "SKU-SHIP-RESERVE-CATCHUP", "T27", "EA", "A1", 10
    AddShippingBomViewRow loBomView, 986, "T27", 986, "T27", 1, "EA"

    If Not modOperatorReadModel.RunBatchAndRefreshOperatorWorkbook(wbOps, "WH102", "LOCAL", report) Then
        failureReason = "RunBatchAndRefreshOperatorWorkbook failed after shipping reserve: " & report
        GoTo CleanExit
    End If
    Set loInv = FindTableByName(wbOps, "invSys")
    invRow = FindRowByColumnValueInTable(loInv, "ITEM_CODE", "SKU-SHIP-RESERVE-CATCHUP")
    If invRow = 0 Then
        failureReason = "Operator invSys row was missing after reserve refresh."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInv, invRow, "TOTAL INV")) <> 9 Then
        failureReason = "Reserve processor refresh left NAS inventory stale; expected 9 but found " & CStr(GetTableValue(loInv, invRow, "TOTAL INV")) & "."
        GoTo CleanExit
    End If

    savedBoxes(1, 1) = 986
    savedBoxes(1, 2) = "T27"
    savedBoxes(1, 3) = "T27"
    savedBoxes(1, 4) = "EA"
    savedBoxes(1, 5) = "A1"
    savedBoxes(1, 6) = ""
    savedBoxes(1, 7) = ""

    wbOps.Activate
    shippables = RunShippingMacro1ForTest("BoxMakerFormLoadShippableVersionInventory", savedBoxes)
    If IsEmpty(shippables) Then
        failureReason = "Shippable version inventory returned no rows after reserve refresh."
        GoTo CleanExit
    End If
    If CDbl(NzDblForTest(shippables(1, 4))) <> 9 Then
        failureReason = "Shipping form NAS Inv did not catch up to projected reserve; expected 9 but found " & CStr(shippables(1, 4)) & "."
        GoTo CleanExit
    End If

    TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected = 1

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH102"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    CloseWorkbookIfOpen FindWorkbookByName("WH102.invSys.Auth.xlsb")
    CloseWorkbookIfOpen FindWorkbookByName("WH102.invSys.Config.xlsb")
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7134, "TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingRefresh_MergesLocalBoxBuildStagingAndClearsStaleOverlay() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim eventIdOut As String
    Dim payloadJson As String
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim wbSnap As Workbook
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loBomView As ListObject
    Dim invRow As Long
    Dim targetText As String
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim savedBoxes(1 To 1, 1 To 7) As Variant
    Dim shippables As Variant
    Dim overlayPath As String

    rootPath = BuildRuntimeTestRoot("phase6_box_build_refresh_nas")
    currentUser = "calvin"

    On Error GoTo CleanFail
    If Not PrepareShippingPostSessionForTest(rootPath, "WH108", "S31", currentUser, failureReason) Then GoTo CleanExit
    EnsureConfigStationRowValue "WH108.invSys.Config.xlsb", "S31", "WH108", "PathInboxRoot", rootPath & "\stale_inbox\"
    If Not modConfig.Reload() Then
        failureReason = "Config reload failed after stale inbox route setup: " & modConfig.Validate()
        GoTo CleanExit
    End If

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH108", Array("SKU-T24-BOX-BUILD-CATCHUP"))
    Set wbInbox = CreateCanonicalShipInboxWorkbookForTest(rootPath, "S31")
    If wbInv Is Nothing Or wbInbox Is Nothing Then
        failureReason = "Canonical shipping runtime workbooks could not be created."
        GoTo CleanExit
    End If
    Set evt = CreateReceiveEventForTest("EVT-BOX-BUILD-CATCHUP-SEED", "WH108", "S31", currentUser, "SKU-T24-BOX-BUILD-CATCHUP", 4, "CLEARVIEW", "box build catch-up seed")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-BOX-BUILD-CATCHUP-SEED", statusOut, errorCode, errorMessage) Then
        failureReason = "Canonical seed event failed: " & errorCode & "; " & errorMessage
        GoTo CleanExit
    End If

    Set wbSnap = CreateSnapshotWorkbook(rootPath, "WH108", "SKU-T24-BOX-BUILD-CATCHUP", 4, CDate("2026-06-19 08:00:00"), 4, "CLEARVIEW=4", "T24", "EA", "CLEARVIEW", "", "", "", "", "986")
    If wbSnap Is Nothing Then
        failureReason = "Stale snapshot workbook could not be created."
        GoTo CleanExit
    End If
    wbSnap.Close SaveChanges:=False
    Set wbSnap = Nothing

    payloadJson = modRoleEventWriter.BuildPayloadJson( _
        modRoleEventWriter.CreatePayloadItem(986, _
                                             "SKU-T24-BOX-BUILD-CATCHUP", _
                                             20, _
                                             "CLEARVIEW", _
                                             "T24 VERSION=v1", _
                                             "MADE"))
    If Not modRoleEventWriter.QueuePayloadEvent(CORE_EVENT_TYPE_BOX_BUILD, "WH108", "S31", currentUser, payloadJson, "box-build-refresh-nas", "", "", Now, Nothing, eventIdOut, report) Then
        failureReason = "QueuePayloadEvent failed for local staged box build: " & report
        GoTo CleanExit
    End If

    RunShippingClearProjectedOverlayForTest
    overlayPath = RunShippingProjectedOverlayPathForTest()
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingRegisterProjectedOverlayForTest 986, "v1", 4, 4

    If Not modRoleEventWriter.SyncLocalStagedInboxRows(report, "WH108", "S31") Then
        failureReason = "Local staged Box Maker row did not merge to NAS inbox: " & report
        GoTo CleanExit
    End If
    If InStr(1, report, "LocalStagingMerged=1", vbTextCompare) = 0 Then
        failureReason = "Local staging sync did not report merging the Box Maker row: " & report
        GoTo CleanExit
    End If

    Set loInv = FindTableByName(wbInbox, "tblInboxShip")
    If loInv Is Nothing Then
        failureReason = "Shipping inbox table was missing after local staging sync."
        GoTo CleanExit
    End If
    If loInv.ListRows.Count <> 1 Then
        failureReason = "Expected one merged Box Maker inbox row, found " & CStr(loInv.ListRows.Count) & "."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInv, 1, "EventType")), CORE_EVENT_TYPE_BOX_BUILD, vbTextCompare) <> 0 Then
        failureReason = "Merged inbox row was not BOX_BUILD: " & CStr(GetTableValue(loInv, 1, "EventType"))
        GoTo CleanExit
    End If
    If InStr(1, CStr(GetTableValue(loInv, 1, "PayloadJson")), "SKU-T24-BOX-BUILD-CATCHUP", vbTextCompare) = 0 Then
        failureReason = "Merged BOX_BUILD payload did not include the shippable SKU."
        GoTo CleanExit
    End If

    wbInbox.Save
    wbInbox.Close SaveChanges:=False
    Set wbInbox = Nothing

    targetText = modProcessor.DescribeInboxTargetsForAutomation("WH108")
    If InStr(1, targetText, rootPath & "\invSys.Inbox.Shipping.S31.xlsb", vbTextCompare) = 0 _
       Or InStr(1, targetText, "Table=tblInboxShip", vbTextCompare) = 0 Then
        failureReason = "Processor inbox targets did not include the connected NAS shipping inbox after local staging merge. " & targetText
        GoTo CleanExit
    End If

    If CDbl(NzDblForTest(RunShippingProjectedOverlayTextForTest(986, "v1", "24"))) <> 24 Then
        failureReason = "Stale Projected Inv overlay masked box build catch-up after backend rose to 24."
        GoTo CleanExit
    End If

    TestShippingRefresh_MergesLocalBoxBuildStagingAndClearsStaleOverlay = 1

CleanExit:
    If Trim$(overlayPath) <> "" Then DeleteFileIfExistsForTest overlayPath
    RunShippingClearProjectedOverlayForTest
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH108"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen wbSnap
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    CloseWorkbookIfOpen FindWorkbookByName("WH108.invSys.Auth.xlsb")
    CloseWorkbookIfOpen FindWorkbookByName("WH108.invSys.Config.xlsb")
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7146, "TestShippingRefresh_MergesLocalBoxBuildStagingAndClearsStaleOverlay", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingRefresh_FindsBackendShippingBomViewWithoutInvSysSurface() As Long
    Dim wbOps As Workbook
    Dim wsShip As Worksheet
    Dim wsBackend As Worksheet
    Dim loBomView As ListObject
    Dim macroName As String
    Dim foundView As Boolean
    Dim failureReason As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsShip = wbOps.Worksheets(1)
    wsShip.Name = "ShipmentsTally"
    Set wsBackend = wbOps.Worksheets.Add(After:=wsShip)
    wsBackend.Name = "ShippingBackend"
    wsBackend.Range("A1:H1").Value = Array("PackageRow", "PackageName", "BomVersion", "BomVersionLabel", "IsActive", "ComponentRow", "ComponentQty", "ComponentUOM")
    wsBackend.Range("A2:H2").Value = Array(989, "T28", 1, "v1", True, 1001, 1, "ea")
    Set loBomView = wsBackend.ListObjects.Add(xlSrcRange, wsBackend.Range("A1:H2"), , xlYes)
    loBomView.Name = "ShippingBOMView"

    If HasTableByName(wbOps, "invSys") Then
        failureReason = "Test setup unexpectedly created the generic invSys InventoryManagement table."
        GoTo CleanExit
    End If

    macroName = ShippingMacroNameForTest("ShippingBomViewTableExistsForWorkbookForTest")
    wbOps.Activate
    foundView = CBool(Application.Run(macroName, wbOps.Name))
    If Not foundView Then
        failureReason = "Shipping BOM lookup did not find ShippingBOMView on the backend sheet."
        GoTo CleanExit
    End If
    If HasTableByName(wbOps, "invSys") Then
        failureReason = "Shipping BOM lookup created the generic invSys InventoryManagement table."
        GoTo CleanExit
    End If

    TestShippingRefresh_FindsBackendShippingBomViewWithoutInvSysSurface = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7147, "TestShippingRefresh_FindsBackendShippingBomViewWithoutInvSysSurface", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestBoxMakerUnbox_QtyGreaterThanInventoryFailsBeforeQueue() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim componentRows(1 To 1, 1 To 8) As Variant

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    If loInv Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 86, "SKU-T24", "T24", "EA", "CLEARVIEW", 26
    AddInvSysSeedRow loInv, 900, "SKU-T24-COMP", "T24 component", "EA", "CLEARVIEW", 100

    componentRows(1, 2) = "T24 component"
    componentRows(1, 3) = "SKU-T24-COMP"
    componentRows(1, 4) = 900
    componentRows(1, 5) = 7
    componentRows(1, 6) = "ea"
    componentRows(1, 7) = "CLEARVIEW"
    componentRows(1, 8) = "component"

    wbOps.Activate
    report = RunBoxMakerCommitActionReportForTest(86, _
                                                  "T24", _
                                                  "ea", _
                                                  "CLEARVIEW", _
                                                  "T24 test box", _
                                                  "v1", _
                                                  58, _
                                                  componentRows, _
                                                  "UNBOX")
    If InStr(1, report, "Posted=1", vbTextCompare) > 0 Then
        failureReason = "Box Maker unbox posted even though Qty exceeded inventory."
        GoTo CleanExit
    End If
    If InStr(1, report, "Qty exceeds inventory", vbTextCompare) = 0 Then
        failureReason = "Box Maker unbox failure did not tell the user Qty exceeds inventory: " & report
        GoTo CleanExit
    End If

    TestBoxMakerUnbox_QtyGreaterThanInventoryFailsBeforeQueue = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7148, "TestBoxMakerUnbox_QtyGreaterThanInventoryFailsBeforeQueue", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestBoxMakerUnbox_UsesShippingReadModelInventoryWhenInvSysMissing() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim componentRows(1 To 1, 1 To 8) As Variant

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSysData_Shipping")
    If loInv Is Nothing Then GoTo CleanExit

    DeleteTableSurfaceForTest wbOps.Worksheets("InventoryManagement"), "invSys"
    AddInvSysSeedRow loInv, 89, "SKU-T28", "T28", "EA", "CLEARVIEW", 77

    componentRows(1, 2) = "T28 component"
    componentRows(1, 3) = "SKU-T28-COMP"
    componentRows(1, 4) = 900
    componentRows(1, 5) = 7
    componentRows(1, 6) = "ea"
    componentRows(1, 7) = "CLEARVIEW"
    componentRows(1, 8) = "component"

    wbOps.Activate
    report = RunBoxMakerCommitActionReportForTest(89, _
                                                  "T28", _
                                                  "ea", _
                                                  "CLEARVIEW", _
                                                  "T28 test box", _
                                                  "v1", _
                                                  78, _
                                                  componentRows, _
                                                  "UNBOX")
    If InStr(1, report, "current inventory was not resolved", vbTextCompare) > 0 Then
        failureReason = "Box Maker unbox did not use invSysData_Shipping inventory: " & report
        GoTo CleanExit
    End If
    If InStr(1, report, "Qty exceeds inventory", vbTextCompare) = 0 Then
        failureReason = "Box Maker unbox did not compare against Shipping read-model inventory: " & report
        GoTo CleanExit
    End If

    report = RunBoxMakerCommitActionReportForTest(89, _
                                                  "T28", _
                                                  "ea", _
                                                  "CLEARVIEW", _
                                                  "T28 test box", _
                                                  "v1", _
                                                  79, _
                                                  componentRows, _
                                                  "UNBOX", _
                                                  "78")
    If InStr(1, report, "current inventory was not resolved", vbTextCompare) > 0 Then
        failureReason = "Box Maker unbox ignored displayed NAS inventory and fell back to unresolved workbook state: " & report
        GoTo CleanExit
    End If
    If InStr(1, report, "has 78 in inventory", vbTextCompare) = 0 Then
        failureReason = "Box Maker unbox did not use displayed NAS inventory for the over-qty guard: " & report
        GoTo CleanExit
    End If

    TestBoxMakerUnbox_UsesShippingReadModelInventoryWhenInvSysMissing = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7149, "TestBoxMakerUnbox_UsesShippingReadModelInventoryWhenInvSysMissing", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingShippables_NasInvPrefersCurrentInvSysForSingleActiveVersion() As Long
    Dim report As String
    Dim failureReason As String
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim loBomView As ListObject
    Dim savedBoxes(1 To 1, 1 To 7) As Variant
    Dim shippables As Variant

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loBomView = FindTableByName(wbOps, "ShippingBOMView")
    If loInv Is Nothing Or loBomView Is Nothing Then GoTo CleanExit

    AddInvSysSeedRow loInv, 980, "SKU-T27", "T27", "EA", "A1", 9
    AddShippingBomViewRow loBomView, 980, "T27", 980, "T27", 1, "EA"
    AddInventoryLogRowForTest wbOps, "BOX_BUILD", "SKU-T27", 10, "VERSION=v1"

    savedBoxes(1, 1) = 980
    savedBoxes(1, 2) = "T27"
    savedBoxes(1, 3) = "T27"
    savedBoxes(1, 4) = "EA"
    savedBoxes(1, 5) = "A1"
    savedBoxes(1, 6) = ""
    savedBoxes(1, 7) = ""

    wbOps.Activate
    shippables = RunShippingMacro1ForTest("BoxMakerFormLoadShippableVersionInventory", savedBoxes)
    If IsEmpty(shippables) Then
        failureReason = "Shippable version inventory returned no rows."
        GoTo CleanExit
    End If
    If CDbl(NzDblForTest(shippables(1, 4))) <> 9 Then
        failureReason = "NAS Inv used stale version-log total instead of current invSys TOTAL INV; expected 9 but found " & CStr(shippables(1, 4)) & "."
        GoTo CleanExit
    End If

    TestShippingShippables_NasInvPrefersCurrentInvSysForSingleActiveVersion = 1

CleanExit:
    CloseWorkbookIfOpen wbOps
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7131, "TestShippingShippables_NasInvPrefersCurrentInvSysForSingleActiveVersion", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingProjectedDisplay_SubtractsLockedAndUnreservedRows() As Long
    Dim failureReason As String

    On Error GoTo CleanFail

    If CDbl(RunShippingProjectedDisplayQtyForTest(10, 1, 0, 0, 10)) <> 9 Then
        failureReason = "Stale server lock was not deducted from Projected Inv."
        GoTo CleanExit
    End If
    If CDbl(RunShippingProjectedDisplayQtyForTest(20, 0, 2, 0, 20)) <> 18 Then
        failureReason = "New local unreserved shipment row was not deducted from Projected Inv."
        GoTo CleanExit
    End If
    If CDbl(RunShippingProjectedDisplayQtyForTest(20, 3, 0, 0, 20)) <> 17 Then
        failureReason = "Locked reservation total was not fully deducted from Projected Inv."
        GoTo CleanExit
    End If
    If CDbl(RunShippingProjectedDisplayQtyForTest(20, 3, 2, 0, 20)) <> 15 Then
        failureReason = "Projected Inv did not combine stale/server locks with new local unreserved rows."
        GoTo CleanExit
    End If
    If CDbl(RunShippingProjectedDisplayQtyForTest(1, 3, 2, 0, 1)) <> 0 Then
        failureReason = "Projected Inv went negative instead of clamping to zero."
        GoTo CleanExit
    End If
    If CDbl(RunShippingProjectedDisplayQtyForTest(20, 3, 0, 2, 18)) <> 17 Then
        failureReason = "Projected Inv double-counted or ignored stale locks when active local reservations were already represented by the overlay."
        GoTo CleanExit
    End If
    If CDbl(RunShippingProjectedDisplayQtyForTest(20, 1, 0, 0, 18)) <> 17 Then
        failureReason = "Projected Inv increased after Shipments Sent released the completed row lock while NAS was still stale."
        GoTo CleanExit
    End If

    TestShippingProjectedDisplay_SubtractsLockedAndUnreservedRows = 1

CleanExit:
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7143, "TestShippingProjectedDisplay_SubtractsLockedAndUnreservedRows", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingReservationTotals_IgnoreSameWorkbookStaleActiveReservationWithoutLocalLine() As Long
    Dim failureReason As String
    Dim wbReservations As Workbook
    Dim loReservations As ListObject
    Dim totals As Object
    Dim key As String
    Dim lineId As String
    Dim sourcePath As String

    lineId = "SHIPLINE-STALE-LOCAL-001"
    key = "984|v1"
    sourcePath = "C:\ops\ShippingStation.xlsm"

    On Error GoTo CleanFail
    Set wbReservations = Application.Workbooks.Add(xlWBATWorksheet)
    Set loReservations = CreateShippingReservationTableForTest(wbReservations, "WH103", "RESERVE-STALE-LOCAL-001", lineId, "REF-STALE-LOCAL", "Stale Local Lock Item", 984, "v1", 1, "EA", "A1")
    If loReservations Is Nothing Then
        failureReason = "Could not create reservation table for test."
        GoTo CleanExit
    End If
    SetTableCell loReservations, 1, "SourceWorkbook", sourcePath

    Set totals = RunShippingReservationTotalsForTableWithLocalLinesForTest(loReservations, "WH103", sourcePath, lineId)
    If totals Is Nothing Or Not totals.Exists(key) Or CDbl(totals(key)) <> 1 Then
        failureReason = "Active same-workbook reservation with a local shipment line should appear in Locked totals."
        GoTo CleanExit
    End If

    Set totals = RunShippingReservationTotalsForTableWithLocalLinesForTest(loReservations, "WH103", sourcePath, "")
    If Not totals Is Nothing Then
        If totals.Exists(key) Then
            failureReason = "Stale same-workbook active reservation without a local shipment line still appears in Locked totals."
            GoTo CleanExit
        End If
    End If

    SetTableCell loReservations, 1, "SourceWorkbook", ""
    Set totals = RunShippingReservationTotalsForTableWithLocalLinesForTest(loReservations, "WH103", sourcePath, "")
    If Not totals Is Nothing Then
        If totals.Exists(key) Then
            failureReason = "Stale blank-source active reservation without a local shipment line still appears in Locked totals."
            GoTo CleanExit
        End If
    End If

    SetTableCell loReservations, 1, "SourceWorkbook", "C:\ops\OldShippingStation.xlsm"
    Set totals = RunShippingReservationTotalsForTableWithLocalLinesForTest(loReservations, "WH103", sourcePath, "")
    If Not totals Is Nothing Then
        If totals.Exists(key) Then
            failureReason = "Stale same-station active reservation from an old workbook without a local shipment line still appears in Locked totals."
            GoTo CleanExit
        End If
    End If

    SetTableCell loReservations, 1, "LineID", ""
    Set totals = RunShippingReservationTotalsForTableWithLocalLinesForTest(loReservations, "WH103", sourcePath, "")
    If Not totals Is Nothing Then
        If totals.Exists(key) Then
            failureReason = "Stale same-station active reservation with blank LineID still appears in Locked totals."
            GoTo CleanExit
        End If
    End If

    TestShippingReservationTotals_IgnoreSameWorkbookStaleActiveReservationWithoutLocalLine = 1

CleanExit:
    CloseWorkbookIfOpen wbReservations
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7132, "TestShippingReservationTotals_IgnoreSameWorkbookStaleActiveReservationWithoutLocalLine", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestShippingReservationTotals_IgnoreLocallySentActiveLedgerRows() As Long
    Dim failureReason As String
    Dim wbReservations As Workbook
    Dim loReservations As ListObject
    Dim totals As Object
    Dim key As String
    Dim lineId As String
    Dim sentPath As String

    lineId = "SHIPLINE-LOCKED-SENT-001"
    key = "985|v2"
    sentPath = LocalShippingStatePathForTest("sent", "WH102")
    DeleteFileIfExistsForTest sentPath

    On Error GoTo CleanFail
    Set wbReservations = Application.Workbooks.Add(xlWBATWorksheet)
    Set loReservations = CreateShippingReservationTableForTest(wbReservations, "WH102", "RESERVE-LOCKED-SENT-001", lineId, "REF-LOCKED-SENT", "Locked Sent Item", 985, "v2", 2, "EA", "A1")
    If loReservations Is Nothing Then
        failureReason = "Could not create reservation table for test."
        GoTo CleanExit
    End If

    Set totals = RunShippingReservationTotalsForTableForTest(loReservations, "WH102")
    If totals Is Nothing Or Not totals.Exists(key) Or CDbl(totals(key)) <> 2 Then
        failureReason = "Active reservation seed did not appear in Locked totals before sent tombstone."
        GoTo CleanExit
    End If

    EnsureFolderForTest ParentFolderPathForTest(sentPath)
    WriteTextFileForTest sentPath, "ID:" & lineId

    Set totals = RunShippingReservationTotalsForTableForTest(loReservations, "WH102")
    If Not totals Is Nothing Then
        If totals.Exists(key) Then
            failureReason = "Locally sent shipment line still appears in Locked totals after refresh/restart tombstone; found " & CStr(totals(key)) & "."
            GoTo CleanExit
        End If
    End If

    TestShippingReservationTotals_IgnoreLocallySentActiveLedgerRows = 1

CleanExit:
    DeleteFileIfExistsForTest sentPath
    CloseWorkbookIfOpen wbReservations
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7126, "TestShippingReservationTotals_IgnoreLocallySentActiveLedgerRows", failureReason
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

Public Function TestProductionEventCreator_QueuesSignedInCurrentTargetEvent() As Long
    Dim rootPath As String
    Dim currentUser As String
    Dim report As String
    Dim failureReason As String
    Dim eventIdOut As String
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim processedCount As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbOps As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loProd As ListObject
    Dim loInventoryLog As ListObject
    Dim logRow As Long

    rootPath = BuildRuntimeTestRoot("phase6_production_event_creator")
    currentUser = "calvin"

    On Error GoTo CleanFail
    mLastTestFailure = vbNullString
    modAuth.SignOut
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime("WH97", "S32", rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime("WH97", "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then
        failureReason = "Config/auth runtime workbooks could not be created. " & report
        GoTo CleanExit
    End If
    SetConfigWarehouseValue "WH97.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.LoadConfig("WH97", "S32") Then
        failureReason = "LoadConfig failed: " & modConfig.Validate()
        GoTo CleanExit
    End If
    If Not modConfig.Reload() Then
        failureReason = "Config reload failed: " & modConfig.Validate()
        GoTo CleanExit
    End If

    EnsureAuthCapabilityForTest "WH97", currentUser, "PROD_POST", "WH97", "*"
    EnsureAuthCapabilityForTest "WH97", "svc_processor", "INBOX_PROCESS", "WH97", "*"
    TestPhase2Helpers.SetUserPinHash wbAuth, currentUser, modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, "S32", True)
    If statusCode <> NAS_OK Then
        failureReason = "SelectWarehouseTarget failed: " & CStr(statusCode)
        GoTo CleanExit
    End If
    If Not modNasConnection.SetCurrentTargetPathsForTest("\\test-nas\invSysWH1", "\\test-nas\invSysWH1\WH97") Then GoTo CleanExit
    authStatus = modAuth.ValidateUserCredentialForTarget(currentUser, "123456", target, "PROD_POST")
    If authStatus <> AUTH_OK Then
        failureReason = "ValidateUserCredentialForTarget failed: " & CStr(authStatus)
        GoTo CleanExit
    End If

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH97", Array("SKU-PROD-CREATOR"))
    Set wbInbox = CreateCanonicalProductionInboxWorkbookForTest(rootPath, "S32")
    If wbInv Is Nothing Or wbInbox Is Nothing Then
        failureReason = "Canonical production runtime workbooks could not be created."
        GoTo CleanExit
    End If

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wbOps, report) Then
        failureReason = "EnsureProductionWorkbookSurface failed: " & report
        GoTo CleanExit
    End If
    Set loInv = FindTableByName(wbOps, "invSys")
    Set loProd = FindTableByName(wbOps, "ProductionOutput")
    If loInv Is Nothing Or loProd Is Nothing Then
        failureReason = "Production operator tables were not created."
        GoTo CleanExit
    End If
    AddInvSysSeedRow loInv, 963, "SKU-PROD-CREATOR", "Production Creator Item", "EA", "FG", 0
    AddProductionOutputRow loProd, "Blend", "Production Creator Item", "EA", 5, "BATCH-CREATOR-001", "RECALL-CREATOR-001", 963

    If Not modProductionEventCreator.QueueProductionCompleteEventFromWorkbook(wbOps, eventIdOut, report) Then
        failureReason = "QueueProductionCompleteEventFromWorkbook failed: " & report
        GoTo CleanExit
    End If
    If Not modNasConnection.SetCurrentTargetPathsForTest(rootPath, rootPath) Then
        failureReason = "Could not restore local processor target after NAS-gated production queue."
        GoTo CleanExit
    End If
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Trim$(eventIdOut) = "" Then
        failureReason = "Production event creator did not return an EventID."
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch("WH97", 500, report)
    If processedCount <> 1 Then
        failureReason = "RunBatch did not process the production creator event. " & report
        GoTo CleanExit
    End If
    If Not AssertInboxRowStatusForTest(wbInbox, eventIdOut, "PROCESSED") Then
        failureReason = "Production creator inbox row was not marked PROCESSED."
        GoTo CleanExit
    End If

    Set loInventoryLog = FindTableByName(wbInv, "tblInventoryLog")
    If loInventoryLog Is Nothing Then
        failureReason = "Canonical inventory log was missing after production creator process."
        GoTo CleanExit
    End If
    logRow = FindRowByColumnValueInTable(loInventoryLog, "EventID", eventIdOut)
    If logRow = 0 Then
        failureReason = "Canonical inventory log did not record the production creator event."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValue(loInventoryLog, logRow, "EventType")), CORE_EVENT_TYPE_PROD_COMPLETE, vbTextCompare) <> 0 Then
        failureReason = "Canonical inventory log recorded unexpected event type for production creator workflow."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loInventoryLog, logRow, "QtyDelta")) <> 5 Then
        failureReason = "Canonical inventory log QtyDelta was not positive for production creator workflow."
        GoTo CleanExit
    End If

    TestProductionEventCreator_QueuesSignedInCurrentTargetEvent = 1

CleanExit:
    modAuth.SignOut
    modNasConnection.ForgetTarget "WH97"
    modNasConnection.ForgetRoot rootPath
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbOps
    CloseWorkbookIfOpen wbInbox
    CloseWorkbookIfOpen wbInv
    CloseWorkbookIfOpen wbAuth
    CloseWorkbookIfOpen wbCfg
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        mLastTestFailure = failureReason
        On Error GoTo 0
        Err.Raise vbObjectError + 7112, "TestProductionEventCreator_QueuesSignedInCurrentTargetEvent", failureReason
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

Public Function TestAdminShipmentReconcile_AppliesSignedDeltaWithCorrectedShipEvidence() As Long
    Dim rootPath As String
    Dim wbInv As Workbook
    Dim evt As Object
    Dim item As Object
    Dim payloadJson As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject
    Dim loSku As ListObject
    Dim logRow As Long
    Dim skuRow As Long
    Dim failureReason As String

    rootPath = BuildRuntimeTestRoot("phase6_admin_ship_reconcile_apply")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH103", "S31") Then GoTo CleanExit
    SetConfigWarehouseValue "WH103.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH103", Array("SKU-ADMIN-RECON"))
    Set evt = CreateReceiveEventForTest("EVT-ADMIN-RECON-SEED", "WH103", "S31", "admin1", "SKU-ADMIN-RECON", 10, "A1", "seed")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-ADMIN-RECON-SEED", statusOut, errorCode, errorMessage) Then
        failureReason = "Seed receive failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    Set item = modRoleEventWriter.CreatePayloadItem(101, "SKU-ADMIN-RECON", 1, "A1", "shipment sent", "SHIPPED")
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)
    Set evt = CreatePayloadEventForTest("EVT-SHIP-ADMIN-RECON-001", CORE_EVENT_TYPE_SHIP, "WH103", "S31", "shipper1", payloadJson, "ship one")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-SHIP-ADMIN-RECON", statusOut, errorCode, errorMessage) Then
        failureReason = "Ship event failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    Set item = modRoleEventWriter.CreatePayloadItem(101, "SKU-ADMIN-RECON", 1, "A1", "dirty add-back", "RELEASED")
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)
    Set evt = CreatePayloadEventForTest("EVT-DIRTY-ADMIN-RECON-001", CORE_EVENT_TYPE_SHIP_RELEASE, "WH103", "S31", "shipper1", payloadJson, "dirty add-back")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-DIRTY-ADMIN-RECON", statusOut, errorCode, errorMessage) Then
        failureReason = "Dirty add-back setup failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    Set item = modRoleEventWriter.CreatePayloadItem(101, "SKU-ADMIN-RECON", -1, "A1", "Correct stale NAS value after shipment test", "RECONCILE")
    item("CorrectedShipEventId") = "EVT-SHIP-ADMIN-RECON-001"
    item("RepairNarrative") = "Correct stale NAS value after shipment test"
    item("MismatchFlag") = "NAS_INCREASED_AFTER_SHIP"
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)
    Set evt = CreatePayloadEventForTest("EVT-ADMIN-RECON-001", CORE_EVENT_TYPE_ADMIN_SHIPMENT_RECONCILE, "WH103", "S31", "admin1", payloadJson, "admin reconcile")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-ADMIN-RECON", statusOut, errorCode, errorMessage) Then
        failureReason = "Admin reconcile failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    Set loLog = FindTableByName(wbInv, "tblInventoryLog")
    Set loSku = FindTableByName(wbInv, "tblSkuBalance")
    logRow = FindRowByColumnValueInTable(loLog, "EventID", "EVT-ADMIN-RECON-001")
    skuRow = FindRowByColumnValueInTable(loSku, "SKU", "SKU-ADMIN-RECON")
    If logRow = 0 Then
        failureReason = "Admin reconcile event was not written to tblInventoryLog."
        GoTo CleanExit
    End If
    If CDbl(GetTableValue(loLog, logRow, "QtyDelta")) <> -1 Then
        failureReason = "Admin reconcile did not preserve signed correction delta."
        GoTo CleanExit
    End If
    If InStr(1, CStr(GetTableValue(loLog, logRow, "Note")), "CorrectsShipEventId=EVT-SHIP-ADMIN-RECON-001", vbTextCompare) = 0 Then
        failureReason = "Admin reconcile log note did not carry corrected Shipments Sent EventID evidence."
        GoTo CleanExit
    End If
    If InStr(1, CStr(GetTableValue(loLog, logRow, "Note")), "Correct stale NAS value after shipment test", vbTextCompare) = 0 Then
        failureReason = "Admin reconcile log note did not carry human repair narrative."
        GoTo CleanExit
    End If
    If skuRow = 0 Or CDbl(GetTableValue(loSku, skuRow, "QtyOnHand")) <> 9 Then
        failureReason = "Admin reconcile did not update derived NAS quantity back to 9."
        GoTo CleanExit
    End If

    TestAdminShipmentReconcile_AppliesSignedDeltaWithCorrectedShipEvidence = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7137, "TestAdminShipmentReconcile_AppliesSignedDeltaWithCorrectedShipEvidence", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestAdminShipmentReconcile_RejectsOrphanAndMissingNarrative() As Long
    Dim report As String
    Dim failureReason As String

    On Error GoTo CleanFail
    If modAdminShipmentReconcile.ValidateShipmentReconcileRequest("SKU-ADMIN-RECON-REQ", -1, "", "orphan correction", report) Then
        failureReason = "Blank CorrectedShipEventId was accepted."
        GoTo CleanExit
    End If
    If InStr(1, report, "CorrectedShipEventId", vbTextCompare) = 0 Then
        failureReason = "CorrectedShipEventId rejection did not explain the requirement: " & report
        GoTo CleanExit
    End If

    report = vbNullString
    If modAdminShipmentReconcile.ValidateShipmentReconcileRequest("SKU-ADMIN-RECON-REQ", -1, "EVT-SHIP-ADMIN-RECON-REQ", "", report) Then
        failureReason = "Missing RepairNarrative was accepted."
        GoTo CleanExit
    End If
    If InStr(1, report, "narrative", vbTextCompare) = 0 Then
        failureReason = "Missing narrative rejection was not explicit: " & report
        GoTo CleanExit
    End If

    TestAdminShipmentReconcile_RejectsOrphanAndMissingNarrative = 1

CleanExit:
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7138, "TestAdminShipmentReconcile_RejectsOrphanAndMissingNarrative", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestAdminShipmentReconcile_DetectsNasIncreaseAfterLatestShip() As Long
    Dim rootPath As String
    Dim wbInv As Workbook
    Dim evt As Object
    Dim item As Object
    Dim payloadJson As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim correctedShipEventId As String
    Dim currentNasQty As Double
    Dim qtyAfterShip As Double
    Dim flags As String
    Dim report As String
    Dim failureReason As String

    rootPath = BuildRuntimeTestRoot("phase6_admin_ship_reconcile_detect")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH105", "S31") Then GoTo CleanExit
    SetConfigWarehouseValue "WH105.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH105", Array("SKU-ADMIN-RECON-DETECT"))
    Set evt = CreateReceiveEventForTest("EVT-ADMIN-RECON-DETECT-SEED", "WH105", "S31", "admin1", "SKU-ADMIN-RECON-DETECT", 10, "A1", "seed")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-ADMIN-RECON-DETECT-SEED", statusOut, errorCode, errorMessage) Then
        failureReason = "Seed receive failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    Set item = modRoleEventWriter.CreatePayloadItem(101, "SKU-ADMIN-RECON-DETECT", 1, "A1", "shipment sent", "SHIPPED")
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)
    Set evt = CreatePayloadEventForTest("EVT-SHIP-ADMIN-RECON-DETECT", CORE_EVENT_TYPE_SHIP, "WH105", "S31", "shipper1", payloadJson, "ship one")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-SHIP-ADMIN-RECON-DETECT", statusOut, errorCode, errorMessage) Then
        failureReason = "Ship event failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    Set item = modRoleEventWriter.CreatePayloadItem(101, "SKU-ADMIN-RECON-DETECT", 1, "A1", "dirty add-back", "RELEASED")
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)
    Set evt = CreatePayloadEventForTest("EVT-DIRTY-ADMIN-RECON-DETECT", CORE_EVENT_TYPE_SHIP_RELEASE, "WH105", "S31", "shipper1", payloadJson, "dirty add-back")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-DIRTY-ADMIN-RECON-DETECT", statusOut, errorCode, errorMessage) Then
        failureReason = "Dirty add-back setup failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    If Not modAdminShipmentReconcile.DetectNasIncreaseAfterLastShip(wbInv, "SKU-ADMIN-RECON-DETECT", correctedShipEventId, currentNasQty, qtyAfterShip, report) Then
        failureReason = "NAS increase after SHIP was not detected: " & report
        GoTo CleanExit
    End If
    If StrComp(correctedShipEventId, "EVT-SHIP-ADMIN-RECON-DETECT", vbTextCompare) <> 0 Then
        failureReason = "Detector returned wrong corrected SHIP EventID: " & correctedShipEventId
        GoTo CleanExit
    End If
    If currentNasQty <> 10 Or qtyAfterShip <> 9 Then
        failureReason = "Detector compared wrong quantities. Current=" & CStr(currentNasQty) & "; AfterShip=" & CStr(qtyAfterShip)
        GoTo CleanExit
    End If

    currentNasQty = 0
    qtyAfterShip = 0
    report = vbNullString
    If Not modAdminShipmentReconcile.DetectNasIncreaseAfterShipEvent(wbInv, "EVT-SHIP-ADMIN-RECON-DETECT", "SKU-ADMIN-RECON-DETECT", currentNasQty, qtyAfterShip, report) Then
        failureReason = "Selected Shipments Sent EventID did not drive detection: " & report
        GoTo CleanExit
    End If
    If currentNasQty <> 10 Or qtyAfterShip <> 9 Then
        failureReason = "Selected EventID detector compared wrong quantities. Current=" & CStr(currentNasQty) & "; AfterShip=" & CStr(qtyAfterShip)
        GoTo CleanExit
    End If

    flags = modAdminShipmentReconcile.BuildShipmentMismatchFlags(9, 10, 0, False, False, True)
    If InStr(1, flags, "NAS_INCREASED_AFTER_SHIP", vbTextCompare) = 0 Then
        failureReason = "Mismatch flags did not include NAS_INCREASED_AFTER_SHIP."
        GoTo CleanExit
    End If

    TestAdminShipmentReconcile_DetectsNasIncreaseAfterLatestShip = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7139, "TestAdminShipmentReconcile_DetectsNasIncreaseAfterLatestShip", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestAdminShipmentReconcile_RecentShipmentsSentLogShowsLast20() As Long
    Dim rootPath As String
    Dim wbInv As Workbook
    Dim evt As Object
    Dim item As Object
    Dim payloadJson As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim recentText As String
    Dim i As Long
    Dim failureReason As String

    rootPath = BuildRuntimeTestRoot("phase6_admin_ship_reconcile_recent")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH106", "S31") Then GoTo CleanExit
    SetConfigWarehouseValue "WH106.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH106", Array("SKU-ADMIN-RECENT"))
    Set evt = CreateReceiveEventForTest("EVT-ADMIN-RECENT-SEED", "WH106", "S31", "admin1", "SKU-ADMIN-RECENT", 40, "A1", "seed")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-ADMIN-RECENT-SEED", statusOut, errorCode, errorMessage) Then
        failureReason = "Seed receive failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    For i = 1 To 25
        Set item = modRoleEventWriter.CreatePayloadItem(101, "SKU-ADMIN-RECENT", 1, "A1", "shipment sent", "SHIPPED")
        payloadJson = modRoleEventWriter.BuildPayloadJson(item)
        Set evt = CreatePayloadEventForTest("EVT-SHIP-RECENT-" & Format$(i, "000"), CORE_EVENT_TYPE_SHIP, "WH106", "S31", "shipper1", payloadJson, "ship one")
        If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-SHIP-RECENT-" & Format$(i, "000"), statusOut, errorCode, errorMessage) Then
            failureReason = "Ship event " & CStr(i) & " failed: " & errorCode & " " & errorMessage
            GoTo CleanExit
        End If
    Next i

    recentText = modAdminShipmentReconcile.BuildRecentShipmentSentLogText(wbInv, 20)
    If InStr(1, recentText, "EVT-SHIP-RECENT-025", vbTextCompare) = 0 Then
        failureReason = "Recent shipment list did not include newest Shipments Sent event."
        GoTo CleanExit
    End If
    If InStr(1, recentText, "EVT-SHIP-RECENT-006", vbTextCompare) = 0 Then
        failureReason = "Recent shipment list did not include the 20th newest Shipments Sent event."
        GoTo CleanExit
    End If
    If InStr(1, recentText, "EVT-SHIP-RECENT-005", vbTextCompare) > 0 Then
        failureReason = "Recent shipment list included an event older than the last 20."
        GoTo CleanExit
    End If

    TestAdminShipmentReconcile_RecentShipmentsSentLogShowsLast20 = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7140, "TestAdminShipmentReconcile_RecentShipmentsSentLogShowsLast20", failureReason
    End If
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestAdminShipmentReconcile_RecentLogIncludesShipReserveEvidence() As Long
    Dim rootPath As String
    Dim wbInv As Workbook
    Dim evt As Object
    Dim item As Object
    Dim payloadJson As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim recentText As String
    Dim currentNasQty As Double
    Dim qtyAfterShip As Double
    Dim report As String
    Dim diagnostics As String
    Dim failureReason As String

    rootPath = BuildRuntimeTestRoot("phase6_admin_ship_reconcile_reserve")

    On Error GoTo CleanFail
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    If Not modConfig.LoadConfig("WH107", "S31") Then GoTo CleanExit
    SetConfigWarehouseValue "WH107.invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.Reload() Then GoTo CleanExit

    Set wbInv = CreateCanonicalInventoryWorkbookForTest(rootPath, "WH107", Array("SKU-ADMIN-RESERVE-DETECT"))
    Set evt = CreateReceiveEventForTest("EVT-ADMIN-RESERVE-SEED", "WH107", "S31", "admin1", "SKU-ADMIN-RESERVE-DETECT", 10, "A1", "seed")
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-ADMIN-RESERVE-SEED", statusOut, errorCode, errorMessage) Then
        failureReason = "Seed receive failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    Set item = modRoleEventWriter.CreatePayloadItem(101, "SKU-ADMIN-RESERVE-DETECT", 1, "A1", "shipment reserved", "RESERVED")
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)
    Set evt = CreatePayloadEventForTest("EVT-RESERVE-ADMIN-RECON-001", CORE_EVENT_TYPE_SHIP_RESERVE, "WH107", "S31", "shipper1", payloadJson, "reserve one")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-RESERVE-ADMIN-RECON", statusOut, errorCode, errorMessage) Then
        failureReason = "Ship reserve event failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    Set item = modRoleEventWriter.CreatePayloadItem(101, "SKU-ADMIN-RESERVE-DETECT", 1, "A1", "dirty add-back", "RELEASED")
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)
    Set evt = CreatePayloadEventForTest("EVT-DIRTY-ADMIN-RESERVE-001", CORE_EVENT_TYPE_SHIP_RELEASE, "WH107", "S31", "shipper1", payloadJson, "dirty add-back")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-DIRTY-ADMIN-RESERVE", statusOut, errorCode, errorMessage) Then
        failureReason = "Dirty add-back setup failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If

    recentText = modAdminShipmentReconcile.BuildRecentShipmentSentLogText(wbInv, 20)
    If InStr(1, recentText, "EVT-RESERVE-ADMIN-RECON-001", vbTextCompare) = 0 Then
        failureReason = "Recent shipment evidence list did not include SHIP_RESERVE EventID: " & recentText
        GoTo CleanExit
    End If
    If InStr(1, recentText, "SHIP_RESERVE", vbTextCompare) = 0 Then
        failureReason = "Recent shipment evidence list did not label reserve event type: " & recentText
        GoTo CleanExit
    End If

    If Not modAdminShipmentReconcile.DetectNasIncreaseAfterShipEvent(wbInv, "EVT-RESERVE-ADMIN-RECON-001", "SKU-ADMIN-RESERVE-DETECT", currentNasQty, qtyAfterShip, report) Then
        failureReason = "Selected SHIP_RESERVE EventID did not drive stale NAS detection: " & report
        GoTo CleanExit
    End If
    If currentNasQty <> 10 Or qtyAfterShip <> 9 Then
        failureReason = "SHIP_RESERVE detector compared wrong quantities. Current=" & CStr(currentNasQty) & "; AfterReserve=" & CStr(qtyAfterShip)
        GoTo CleanExit
    End If

    diagnostics = modAdminShipmentReconcile.ShipmentLogDiagnosticsText(wbInv)
    If InStr(1, diagnostics, "SHIP_RESERVE rows: 1", vbTextCompare) = 0 Then
        failureReason = "Diagnostics did not count SHIP_RESERVE rows: " & diagnostics
        GoTo CleanExit
    End If

    TestAdminShipmentReconcile_RecentLogIncludesShipReserveEvidence = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseWorkbookIfOpen wbInv
    DeleteRuntimeRoot rootPath
    If failureReason <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7141, "TestAdminShipmentReconcile_RecentLogIncludesShipReserveEvidence", failureReason
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
    Dim colIndex As Long

    colIndex = GetTableColumnIndexForTest(lo, columnName)
    If colIndex = 0 Then Err.Raise vbObjectError + 7198, "GetTableValue", "Column '" & columnName & "' was not found in table '" & TableNameForTest(lo) & "'."
    If lo.DataBodyRange Is Nothing Or rowIndex < 1 Or rowIndex > lo.ListRows.Count Then _
        Err.Raise vbObjectError + 7198, "GetTableValue", "Row " & CStr(rowIndex) & " was not available in table '" & TableNameForTest(lo) & "'."
    GetTableValue = lo.DataBodyRange.Cells(rowIndex, colIndex).Value
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
    If lo.DataBodyRange Is Nothing Or lo.ListRows.Count = 0 Then
        Set lr = lo.ListRows.Add
    Else
        Set lr = lo.ListRows(1)
    End If
    SetTableCell lo, lr.Index, "REF_NUMBER", refNumber
    SetTableCell lo, lr.Index, "ITEMS", itemName
    SetTableCell lo, lr.Index, "QUANTITY", qty
    SetTableCell lo, lr.Index, "ROW", rowValue
    SetTableCell lo, lr.Index, "UOM", uom
    SetTableCell lo, lr.Index, "LOCATION", locationVal
    SetTableCell lo, lr.Index, "DESCRIPTION", descriptionVal
End Sub

Private Sub CreateShippingReservationLedgerForTest(ByVal rootPath As String, _
                                                   ByVal warehouseId As String, _
                                                   ByVal reserveEventId As String, _
                                                   ByVal lineId As String, _
                                                   ByVal refNumber As String, _
                                                   ByVal itemName As String, _
                                                   ByVal packageRow As Long, _
                                                   ByVal versionText As String, _
                                                   ByVal qty As Double, _
                                                   ByVal uom As String, _
                                                   ByVal locationVal As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim targetPath As String

    headers = Array("ReservationID", "Status", "WarehouseId", "StationId", "UserId", _
                    "LineID", "EventID", "RefNumber", "ItemName", "PackageRow", _
                    "Version", "Qty", "UOM", "Location", "SourceWorkbook", _
                    "CreatedAtUTC", "UpdatedAtUTC", "ReleasedAtUTC", "CompletedAtUTC", "ReleaseEventID")
    targetPath = NormalizeTestPath(rootPath) & "\" & warehouseId & ".invSys.Data.ShippingReservations.xlsb"
    CloseWorkbookIfOpen FindWorkbookByName(warehouseId & ".invSys.Data.ShippingReservations.xlsb")
    DeleteFileIfExistsForTest targetPath

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "ShippingReservations"
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Cells(2, 1).Value = lineId & "|" & reserveEventId
    ws.Cells(2, 2).Value = "ACTIVE"
    ws.Cells(2, 3).Value = warehouseId
    ws.Cells(2, 4).Value = "S31"
    ws.Cells(2, 5).Value = "calvin"
    ws.Cells(2, 6).Value = lineId
    ws.Cells(2, 7).Value = reserveEventId
    ws.Cells(2, 8).Value = refNumber
    ws.Cells(2, 9).Value = itemName
    ws.Cells(2, 10).Value = packageRow
    ws.Cells(2, 11).Value = versionText
    ws.Cells(2, 12).Value = qty
    ws.Cells(2, 13).Value = uom
    ws.Cells(2, 14).Value = locationVal
    ws.Cells(2, 15).Value = "phase6-test"
    ws.Cells(2, 16).Value = CDate("2026-03-25 13:00:00")
    ws.Cells(2, 17).Value = CDate("2026-03-25 13:00:00")

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1)), , xlYes)
    lo.Name = "tblShippingReservations"
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    wb.Close SaveChanges:=False
End Sub

Private Function CreateShippingReservationTableForTest(ByVal wb As Workbook, _
                                                       ByVal warehouseId As String, _
                                                       ByVal reserveEventId As String, _
                                                       ByVal lineId As String, _
                                                       ByVal refNumber As String, _
                                                       ByVal itemName As String, _
                                                       ByVal packageRow As Long, _
                                                       ByVal versionText As String, _
                                                       ByVal qty As Double, _
                                                       ByVal uom As String, _
                                                       ByVal locationVal As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim i As Long

    headers = Array("ReservationID", "Status", "WarehouseId", "StationId", "UserId", _
                    "LineID", "EventID", "RefNumber", "ItemName", "PackageRow", _
                    "Version", "Qty", "UOM", "Location", "SourceWorkbook", _
                    "CreatedAtUTC", "UpdatedAtUTC", "ReleasedAtUTC", "CompletedAtUTC", "ReleaseEventID")
    If wb Is Nothing Then Exit Function
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "ShippingReservations"
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Cells(2, 1).Value = lineId & "|" & reserveEventId
    ws.Cells(2, 2).Value = "ACTIVE"
    ws.Cells(2, 3).Value = warehouseId
    ws.Cells(2, 4).Value = "S31"
    ws.Cells(2, 5).Value = "calvin"
    ws.Cells(2, 6).Value = lineId
    ws.Cells(2, 7).Value = reserveEventId
    ws.Cells(2, 8).Value = refNumber
    ws.Cells(2, 9).Value = itemName
    ws.Cells(2, 10).Value = packageRow
    ws.Cells(2, 11).Value = versionText
    ws.Cells(2, 12).Value = qty
    ws.Cells(2, 13).Value = uom
    ws.Cells(2, 14).Value = locationVal
    ws.Cells(2, 15).Value = "phase6-test"
    ws.Cells(2, 16).Value = CDate("2026-03-25 13:00:00")
    ws.Cells(2, 17).Value = CDate("2026-03-25 13:00:00")

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1)), , xlYes)
    lo.Name = "tblShippingReservationsForTest"
    Set CreateShippingReservationTableForTest = lo
End Function

Private Sub AddInventoryLogRowForTest(ByVal wb As Workbook, _
                                      ByVal eventType As String, _
                                      ByVal sku As String, _
                                      ByVal qtyDelta As Double, _
                                      ByVal noteText As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim headers As Variant
    Dim i As Long

    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets("InventoryLog")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "InventoryLog"
    End If

    Set lo = FindTableByName(wb, "tblInventoryLog")
    If lo Is Nothing Then
        headers = Array("EventID", "EventType", "SKU", "QtyDelta", "Note")
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1)), , xlYes)
        lo.Name = "tblInventoryLog"
        If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    End If

    Set lr = lo.ListRows.Add
    SetTableCell lo, lr.Index, "EventID", "LOG-" & CStr(lr.Index)
    SetTableCell lo, lr.Index, "EventType", eventType
    SetTableCell lo, lr.Index, "SKU", sku
    SetTableCell lo, lr.Index, "QtyDelta", qtyDelta
    SetTableCell lo, lr.Index, "Note", noteText
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

Private Sub AddAggregatePackagesRow(ByVal lo As ListObject, _
                                    ByVal rowValue As Long, _
                                    ByVal itemName As String, _
                                    ByVal qty As Double, _
                                    ByVal uom As String, _
                                    ByVal locationVal As String)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValue(lo, 1, "ITEM"))) = "" _
        And NzDblForTest(GetTableValue(lo, 1, "QUANTITY")) = 0 Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "ROW", rowValue
    SetOptionalTableCell lo, lr.Index, "ITEM_CODE", "SKU-" & CStr(rowValue)
    SetTableCell lo, lr.Index, "ITEM", itemName
    SetTableCell lo, lr.Index, "QUANTITY", qty
    SetTableCell lo, lr.Index, "UOM", uom
    SetTableCell lo, lr.Index, "LOCATION", locationVal
End Sub

Private Sub AddAggregateBomRow(ByVal lo As ListObject, _
                               ByVal rowValue As Long, _
                               ByVal itemName As String, _
                               ByVal qty As Double, _
                               ByVal uom As String, _
                               ByVal locationVal As String)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And Trim$(CStr(GetTableValue(lo, 1, "ITEM"))) = "" _
        And NzDblForTest(GetTableValue(lo, 1, "QUANTITY")) = 0 Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "ROW", rowValue
    SetOptionalTableCell lo, lr.Index, "ITEM_CODE", "SKU-" & CStr(rowValue)
    SetTableCell lo, lr.Index, "ITEM", itemName
    SetTableCell lo, lr.Index, "QUANTITY", qty
    SetTableCell lo, lr.Index, "UOM", uom
    SetTableCell lo, lr.Index, "LOCATION", locationVal
End Sub

Private Sub AddShippingBomViewRow(ByVal lo As ListObject, _
                                  ByVal packageRow As Long, _
                                  ByVal packageItem As String, _
                                  ByVal componentRow As Long, _
                                  ByVal componentItem As String, _
                                  ByVal componentQty As Double, _
                                  ByVal componentUom As String)
    Dim lr As ListRow

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        Set lr = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 _
        And NzDblForTest(GetTableValue(lo, 1, "PackageRow")) = 0 _
        And Trim$(CStr(GetTableValue(lo, 1, "ComponentItem"))) = "" Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    SetTableCell lo, lr.Index, "PackageRow", packageRow
    SetTableCell lo, lr.Index, "PackageItem", packageItem
    SetOptionalTableCell lo, lr.Index, "PackageUOM", "EA"
    SetOptionalTableCell lo, lr.Index, "PackageLocation", "A1"
    SetOptionalTableCell lo, lr.Index, "BomVersionLabel", "v1"
    SetOptionalTableCell lo, lr.Index, "IsActive", True
    SetTableCell lo, lr.Index, "ComponentRow", componentRow
    SetTableCell lo, lr.Index, "ComponentItem", componentItem
    SetTableCell lo, lr.Index, "ComponentQty", componentQty
    SetTableCell lo, lr.Index, "ComponentUOM", componentUom
    SetOptionalTableCell lo, lr.Index, "ComponentLocation", "A1"
    SetOptionalTableCell lo, lr.Index, "UpdatedAtUTC", CDate("2026-03-25 10:50:00")
    SetOptionalTableCell lo, lr.Index, "UpdatedBy", "phase6-test"
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
    Dim colIndex As Long

    If lo Is Nothing Then Exit Sub
    colIndex = GetTableColumnIndexForTest(lo, columnName)
    If colIndex = 0 Then Err.Raise vbObjectError + 7198, "SetTableCell", "Column '" & columnName & "' was not found in table '" & TableNameForTest(lo) & "'."
    If lo.DataBodyRange Is Nothing Or rowIndex < 1 Or rowIndex > lo.ListRows.Count Then _
        Err.Raise vbObjectError + 7198, "SetTableCell", "Row " & CStr(rowIndex) & " was not available in table '" & TableNameForTest(lo) & "'."
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = valueIn
End Sub

Private Sub SetOptionalTableCell(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    Dim colIndex As Long

    If lo Is Nothing Then Exit Sub
    colIndex = GetTableColumnIndexForTest(lo, columnName)
    If colIndex = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Or rowIndex < 1 Or rowIndex > lo.ListRows.Count Then _
        Err.Raise vbObjectError + 7198, "SetOptionalTableCell", "Row " & CStr(rowIndex) & " was not available in table '" & TableNameForTest(lo) & "'."
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = valueIn
End Sub

Private Sub SetUseExistingInventoryForTest(ByVal ws As Worksheet, ByVal enabled As Boolean)
    Dim shp As Shape
    Dim expectedValue As Long

    If ws Is Nothing Then Exit Sub
    expectedValue = IIf(enabled, 1, xlOff)
    If ws.ProtectContents Then ws.Unprotect
    On Error Resume Next
    Set shp = ws.Shapes("CHK_USE_EXISTING")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlCheckBox, 12, 12, 160, 18)
        shp.Name = "CHK_USE_EXISTING"
    End If
    shp.ControlFormat.Value = expectedValue
    If shp.ControlFormat.Value <> expectedValue Then
        Err.Raise vbObjectError + 7197, "SetUseExistingInventoryForTest", "Could not set CHK_USE_EXISTING to the requested state."
    End If
End Sub

Private Function RunShippingMacro0ForTest(ByVal procedureName As String) As Variant
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest(procedureName)
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingMacro0ForTest = Application.Run(macroName)
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingMacro1ForTest(ByVal procedureName As String, ByVal arg1 As Variant) As Variant
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest(procedureName)
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingMacro1ForTest = Application.Run(macroName, arg1)
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingCommitLineForTest(ByVal targetName As String, _
                                              ByVal actionName As String, _
                                              ByVal tableRowIndex As Long, _
                                              ByVal refNumber As String, _
                                              ByVal itemName As String, _
                                              ByVal qtyValue As Double, _
                                              ByVal rowValue As Long, _
                                              ByVal uomValue As String, _
                                              ByVal locationValue As String, _
                                              ByVal descriptionValue As String, _
                                              ByVal carrierValue As String, _
                                              ByRef report As String, _
                                              Optional ByVal displayedAvailableQty As Variant) As Boolean
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ShipmentsFormCommitLine")
    If Not targetWb Is Nothing Then targetWb.Activate
    If IsMissing(displayedAvailableQty) Then
        RunShippingCommitLineForTest = CBool(Application.Run(macroName, _
                                                             targetName, _
                                                             actionName, _
                                                             tableRowIndex, _
                                                             refNumber, _
                                                             itemName, _
                                                             qtyValue, _
                                                             rowValue, _
                                                             uomValue, _
                                                             locationValue, _
                                                             descriptionValue, _
                                                             carrierValue, _
                                                             report))
    Else
        RunShippingCommitLineForTest = CBool(Application.Run(macroName, _
                                                             targetName, _
                                                             actionName, _
                                                             tableRowIndex, _
                                                             refNumber, _
                                                             itemName, _
                                                             qtyValue, _
                                                             rowValue, _
                                                             uomValue, _
                                                             locationValue, _
                                                             descriptionValue, _
                                                             carrierValue, _
                                                             report, _
                                                             displayedAvailableQty))
    End If
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingProjectedAvailabilityOverrideForTest(ByVal tableRowIndex As Long, _
                                                                 ByVal rowValue As Long, _
                                                                 ByVal versionLabel As String, _
                                                                 ByVal displayedAvailableQty As Variant) As String
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ValidateShippingAddProjectedAvailabilityOverrideForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingProjectedAvailabilityOverrideForTest = CStr(Application.Run(macroName, tableRowIndex, rowValue, versionLabel, displayedAvailableQty))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingSentProjectedOverlayQtyForTest(ByVal backendQty As Double, _
                                                           ByVal existingProjectedQty As Double, _
                                                           ByVal shippedQty As Double, _
                                                           Optional ByVal hasExistingOverlay As Boolean = False, _
                                                           Optional ByVal isReservedRow As Boolean = False) As Double
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ShipmentsSentProjectedOverlayQtyForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingSentProjectedOverlayQtyForTest = CDbl(Application.Run(macroName, backendQty, existingProjectedQty, shippedQty, hasExistingOverlay, isReservedRow))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingValidateCommitInputsReportForTest(ByVal targetName As String, _
                                                              ByVal actionName As String, _
                                                              ByVal itemName As String, _
                                                              ByVal qtyValue As Double, _
                                                              ByVal rowValue As Long, _
                                                              ByVal carrierValue As String) As String
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ValidateShipmentCommitInputsReportForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingValidateCommitInputsReportForTest = CStr(Application.Run(macroName, _
                                                                        targetName, _
                                                                        actionName, _
                                                                        itemName, _
                                                                        qtyValue, _
                                                                        rowValue, _
                                                                        carrierValue))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingSentRowsForTest(ByVal rowIndexes As Variant, ByVal carrierValue As String, ByRef report As String) As Boolean
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ShipmentsFormRunShipmentsSentRows")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingSentRowsForTest = CBool(Application.Run(macroName, rowIndexes, carrierValue, report))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingSentRowsReportForTest(ByVal rowIndexes As Variant, ByVal carrierValue As String) As String
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ShipmentsFormRunShipmentsSentRowsReportForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingSentRowsReportForTest = CStr(Application.Run(macroName, rowIndexes, carrierValue))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingMoveHoldRowsForTest(ByVal rowIndexes As Variant, ByVal moveToHold As Boolean, ByRef report As String) As Boolean
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ShipmentsFormMoveHoldRows")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingMoveHoldRowsForTest = CBool(Application.Run(macroName, rowIndexes, moveToHold, report))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunBoxBuilderSavedBoxesReportForTest(ByVal includeActive As Boolean, ByVal includeArchived As Boolean) As String
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("BoxBuilderFormLoadSavedBoxesReportForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunBoxBuilderSavedBoxesReportForTest = CStr(Application.Run(macroName, includeActive, includeArchived))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunBoxBuilderInitializeSmokeForTest(ByRef report As String) As Boolean
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("BoxBuilderFormInitializeSmokeForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunBoxBuilderInitializeSmokeForTest = CBool(Application.Run(macroName, report))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunBoxMakerCommitActionReportForTest(ByVal packageRow As Long, _
                                                      ByVal boxName As String, _
                                                      ByVal boxUom As String, _
                                                      ByVal boxLocation As String, _
                                                      ByVal boxDescription As String, _
                                                      ByVal versionLabel As String, _
                                                      ByVal boxQty As Double, _
                                                      ByVal componentRows As Variant, _
                                                      ByVal actionText As String, _
                                                      Optional ByVal displayedAvailableQty As Variant) As String
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("CommitBoxMakerFormActionReportForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    If IsMissing(displayedAvailableQty) Then
        RunBoxMakerCommitActionReportForTest = CStr(Application.Run(macroName, _
                                                                    packageRow, _
                                                                    boxName, _
                                                                    boxUom, _
                                                                    boxLocation, _
                                                                    boxDescription, _
                                                                    versionLabel, _
                                                                    boxQty, _
                                                                    componentRows, _
                                                                    actionText))
    Else
        RunBoxMakerCommitActionReportForTest = CStr(Application.Run(macroName, _
                                                                    packageRow, _
                                                                    boxName, _
                                                                    boxUom, _
                                                                    boxLocation, _
                                                                    boxDescription, _
                                                                    versionLabel, _
                                                                    boxQty, _
                                                                    componentRows, _
                                                                    actionText, _
                                                                    displayedAvailableQty))
    End If
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingApplySentRowsInventoryForTest(ByVal rowIndexes As Variant, ByRef report As String) As Boolean
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ValidateApplyShipmentsSentRowsInventoryFromCurrentWorkbook")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingApplySentRowsInventoryForTest = CBool(Application.Run(macroName, rowIndexes, report))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingCompleteSentRowsForTest(ByVal rowIndexes As Variant, ByRef report As String) As Boolean
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ValidateCompleteShipmentsSentRowsFromCurrentWorkbook")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingCompleteSentRowsForTest = CBool(Application.Run(macroName, rowIndexes, report))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingReservationTotalsForTest() As Object
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ShipmentsFormLoadNasReservationTotals")
    If Not targetWb Is Nothing Then targetWb.Activate
    Set RunShippingReservationTotalsForTest = Application.Run(macroName)
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingReservationTotalsForTableForTest(ByVal loReservations As ListObject, ByVal warehouseId As String) As Object
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ValidateShippingReservationTotalsFromTableForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    Set RunShippingReservationTotalsForTableForTest = Application.Run(macroName, loReservations, warehouseId)
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingReservationTotalsForTableWithLocalLinesForTest(ByVal loReservations As ListObject, _
                                                                           ByVal warehouseId As String, _
                                                                           ByVal localSourceWorkbook As String, _
                                                                           ByVal activeLineIdsCsv As String) As Object
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ValidateShippingReservationTotalsFromTableWithLocalLinesForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    Set RunShippingReservationTotalsForTableWithLocalLinesForTest = Application.Run(macroName, _
                                                                                   loReservations, _
                                                                                   warehouseId, _
                                                                                   localSourceWorkbook, _
                                                                                   activeLineIdsCsv)
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Sub RunShippingRegisterProjectedOverlayForTest(ByVal packageRow As Long, _
                                                       ByVal versionLabel As String, _
                                                       ByVal projectedQty As Double, _
                                                       Optional ByVal baselineQty As Variant)
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("RegisterPendingBoxVersionInventoryOverlay")
    If Not targetWb Is Nothing Then targetWb.Activate
    If IsMissing(baselineQty) Then
        Application.Run macroName, packageRow, versionLabel, projectedQty
    Else
        Application.Run macroName, packageRow, versionLabel, projectedQty, CDbl(baselineQty)
    End If
    If Not targetWb Is Nothing Then targetWb.Activate
End Sub

Private Sub RunShippingClearProjectedOverlayForTest()
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ClearPendingBoxVersionInventoryOverlayForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    Application.Run macroName
    If Not targetWb Is Nothing Then targetWb.Activate
End Sub

Private Function RunShippingProjectedOverlayPathForTest() As String
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("PendingBoxVersionInventoryOverlayPathForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingProjectedOverlayPathForTest = CStr(Application.Run(macroName))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingProjectedOverlayTextForTest(ByVal packageRow As Long, ByVal versionLabel As String, ByVal backendText As String) As String
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("PendingBoxVersionInventoryOverlayText")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingProjectedOverlayTextForTest = CStr(Application.Run(macroName, packageRow, versionLabel, backendText))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function RunShippingProjectedDisplayQtyForTest(ByVal nasQty As Double, _
                                                       ByVal lockedQty As Double, _
                                                       ByVal unreservedLocalQty As Double, _
                                                       ByVal reservedLocalQty As Double, _
                                                       ByVal pendingOverlayQty As Double) As Double
    Dim targetWb As Workbook
    Dim macroName As String

    Set targetWb = ActiveWorkbook
    macroName = ShippingMacroNameForTest("ShipmentsProjectedDisplayQtyForTest")
    If Not targetWb Is Nothing Then targetWb.Activate
    RunShippingProjectedDisplayQtyForTest = CDbl(Application.Run(macroName, nasQty, lockedQty, unreservedLocalQty, reservedLocalQty, pendingOverlayQty))
    If Not targetWb Is Nothing Then targetWb.Activate
End Function

Private Function ShippingMacroNameForTest(ByVal procedureName As String) As String
    Dim wb As Workbook

    Set wb = EnsureShippingAddinForTest()
    If wb Is Nothing Then Err.Raise vbObjectError + 7197, "ShippingMacroNameForTest", "Could not open deploy\current\invSys.Shipping.xlam for Shipping macro validation."
    ShippingMacroNameForTest = "'" & wb.Name & "'!modTS_Shipments." & procedureName
End Function

Private Function EnsureShippingAddinForTest() As Workbook
    Dim wb As Workbook
    Dim repoPath As String
    Dim addinPath As String

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, "invSys.Shipping.xlam", vbTextCompare) = 0 Then
            Set EnsureShippingAddinForTest = wb
            Exit Function
        End If
    Next wb

    repoPath = ParentFolderPathForTest(ParentFolderPathForTest(ThisWorkbook.Path))
    addinPath = repoPath & "\deploy\current\invSys.Shipping.xlam"
    If Len(Dir$(addinPath, vbNormal)) = 0 Then Exit Function
    Set EnsureShippingAddinForTest = Application.Workbooks.Open(Filename:=addinPath, ReadOnly:=True)
End Function

Private Function CountShipmentRowsForTest(ByVal lo As ListObject, _
                                          ByVal refNumber As String, _
                                          ByVal itemName As String, _
                                          ByVal versionText As String, _
                                          ByVal carrierText As String) As Long
    Dim r As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.ListRows.Count
        If ShipmentRowMatchesForTest(lo, r, refNumber, itemName, versionText, carrierText) Then
            CountShipmentRowsForTest = CountShipmentRowsForTest + 1
        End If
    Next r
End Function

Private Function FindShipmentRowForTest(ByVal lo As ListObject, _
                                        ByVal refNumber As String, _
                                        ByVal itemName As String, _
                                        ByVal versionText As String, _
                                        ByVal carrierText As String) As Long
    Dim r As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.ListRows.Count
        If ShipmentRowMatchesForTest(lo, r, refNumber, itemName, versionText, carrierText) Then
            FindShipmentRowForTest = r
            Exit Function
        End If
    Next r
End Function

Private Function ShipmentRowMatchesForTest(ByVal lo As ListObject, _
                                           ByVal rowIndex As Long, _
                                           ByVal refNumber As String, _
                                           ByVal itemName As String, _
                                           ByVal versionText As String, _
                                           ByVal carrierText As String) As Boolean
    If Trim$(CStr(GetTableValue(lo, rowIndex, "REF_NUMBER"))) <> refNumber Then Exit Function
    If Trim$(CStr(GetTableValue(lo, rowIndex, "ITEMS"))) <> itemName Then Exit Function
    If Trim$(CStr(GetTableValue(lo, rowIndex, "DESCRIPTION"))) <> versionText Then Exit Function
    If Trim$(CStr(GetTableValue(lo, rowIndex, "CARRIER"))) <> carrierText Then Exit Function
    ShipmentRowMatchesForTest = True
End Function

Private Function PrepareShippingPostSessionForTest(ByVal rootPath As String, _
                                                   ByVal warehouseId As String, _
                                                   ByVal stationId As String, _
                                                   ByVal userId As String, _
                                                   ByRef failureReason As String) As Boolean
    Dim report As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode

    modAuth.SignOut
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime(warehouseId, stationId, rootPath, report)
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime(warehouseId, "svc_processor", rootPath, report)
    If wbCfg Is Nothing Or wbAuth Is Nothing Then
        failureReason = "Config/auth runtime workbooks could not be created. " & report
        Exit Function
    End If
    SetConfigWarehouseValue warehouseId & ".invSys.Config.xlsb", "PathDataRoot", rootPath & "\"
    If Not modConfig.LoadConfig(warehouseId, stationId) Then
        failureReason = "LoadConfig failed: " & modConfig.Validate()
        Exit Function
    End If
    If Not modConfig.Reload() Then
        failureReason = "Config reload failed: " & modConfig.Validate()
        Exit Function
    End If

    EnsureAuthCapabilityForTest warehouseId, userId, "SHIP_POST", warehouseId, "*"
    EnsureAuthCapabilityForTest warehouseId, "svc_processor", "INBOX_PROCESS", warehouseId, "*"
    TestPhase2Helpers.SetUserPinHash wbAuth, userId, modAuth.HashUserCredential("123456")
    wbAuth.Save

    statusCode = modNasConnection.SelectWarehouseTarget(rootPath, rootPath, target, stationId, True)
    If statusCode <> NAS_OK Then
        failureReason = "SelectWarehouseTarget failed: " & CStr(statusCode)
        Exit Function
    End If
    authStatus = modAuth.ValidateUserCredentialForTarget(userId, "123456", target, "SHIP_POST")
    If authStatus <> AUTH_OK Then
        failureReason = "ValidateUserCredentialForTarget failed: " & CStr(authStatus)
        Exit Function
    End If

    PrepareShippingPostSessionForTest = True
End Function

Private Function GetTableColumnIndexForTest(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(CStr(lo.ListColumns(i).Name), columnName, vbTextCompare) = 0 Then
            GetTableColumnIndexForTest = i
            Exit Function
        End If
    Next i
End Function

Private Function TableNameForTest(ByVal lo As ListObject) As String
    If lo Is Nothing Then
        TableNameForTest = "<nothing>"
    Else
        TableNameForTest = lo.Name
    End If
End Function

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

Private Function CreatePayloadEventForTest(ByVal eventId As String, _
                                           ByVal eventType As String, _
                                           ByVal warehouseId As String, _
                                           ByVal stationId As String, _
                                           ByVal userId As String, _
                                           ByVal payloadJson As String, _
                                           ByVal noteVal As String) As Object
    Dim evt As Object

    Set evt = CreateObject("Scripting.Dictionary")
    evt.CompareMode = vbTextCompare
    evt("EventID") = eventId
    evt("EventType") = eventType
    evt("CreatedAtUTC") = Now
    evt("WarehouseId") = warehouseId
    evt("StationId") = stationId
    evt("UserId") = userId
    evt("SourceInbox") = "phase6-test-inbox"
    evt("PayloadJson") = payloadJson
    evt("Note") = noteVal
    Set CreatePayloadEventForTest = evt
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
                                        Optional ByVal category As String = vbNullString, _
                                        Optional ByVal rowKey As String = vbNullString) As Workbook
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
    ws.Range("C1").Value = "ROW"
    ws.Range("D1").Value = "ITEM"
    ws.Range("E1").Value = "UOM"
    ws.Range("F1").Value = "LOCATION"
    ws.Range("G1").Value = "DESCRIPTION"
    ws.Range("H1").Value = "VENDOR(s)"
    ws.Range("I1").Value = "VENDOR_CODE"
    ws.Range("J1").Value = "CATEGORY"
    ws.Range("K1").Value = "QtyOnHand"
    ws.Range("L1").Value = "QtyAvailable"
    ws.Range("M1").Value = "LocationSummary"
    ws.Range("N1").Value = "LastAppliedAtUTC"
    ws.Range("A2").Value = warehouseId
    ws.Range("B2").Value = sku
    ws.Range("C2").Value = rowKey
    If Trim$(itemName) = "" Then itemName = sku
    ws.Range("D2").Value = itemName
    ws.Range("E2").Value = uom
    ws.Range("F2").Value = locationVal
    ws.Range("G2").Value = description
    ws.Range("H2").Value = vendorName
    ws.Range("I2").Value = vendorCode
    ws.Range("J2").Value = category
    ws.Range("K2").Value = qtyOnHand
    If IsMissing(qtyAvailable) Or IsEmpty(qtyAvailable) Then
        resolvedQtyAvailable = qtyOnHand
    Else
        resolvedQtyAvailable = CDbl(qtyAvailable)
    End If
    ws.Range("L2").Value = resolvedQtyAvailable
    If IsMissing(locationSummary) Or IsEmpty(locationSummary) Then
        resolvedLocationSummary = "A1=" & CStr(CLng(qtyOnHand))
    Else
        resolvedLocationSummary = CStr(locationSummary)
    End If
    ws.Range("M2").Value = resolvedLocationSummary
    ws.Range("N2").Value = lastAppliedUtc
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:N2"), , xlYes)
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

Private Function LocalShippingStatePathForTest(ByVal kind As String, ByVal warehouseId As String) As String
    Dim rootPath As String

    rootPath = Environ$("LOCALAPPDATA")
    If Trim$(rootPath) = "" Then rootPath = Environ$("TEMP")
    LocalShippingStatePathForTest = NormalizeTestPath(rootPath) & "\invSys\shipping_" & kind & "_" & SafeFileTokenForTest(warehouseId) & ".tsv"
End Function

Private Function SafeFileTokenForTest(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String
    Dim outText As String

    valueText = Trim$(valueText)
    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then
            outText = outText & ch
        Else
            outText = outText & "_"
        End If
    Next i
    If outText = "" Then outText = "default"
    SafeFileTokenForTest = outText
End Function

Private Function ParentFolderPathForTest(ByVal filePath As String) As String
    Dim pos As Long

    pos = InStrRev(filePath, "\")
    If pos > 0 Then ParentFolderPathForTest = Left$(filePath, pos - 1)
End Function

Private Sub EnsureFolderForTest(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = NormalizeTestPath(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub
    parentPath = ParentFolderPathForTest(folderPath)
    If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderForTest parentPath
    MkDir folderPath
End Sub

Private Sub WriteTextFileForTest(ByVal filePath As String, ByVal textValue As String)
    Dim fileNo As Integer

    fileNo = FreeFile
    Open filePath For Output As #fileNo
    Print #fileNo, textValue
    Close #fileNo
End Sub

Private Sub DeleteFileIfExistsForTest(ByVal filePath As String)
    On Error Resume Next
    If Len(Dir$(filePath, vbNormal)) > 0 Then Kill filePath
    On Error GoTo 0
End Sub

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
