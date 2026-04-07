Attribute VB_Name = "TestReceivingReadiness"
Option Explicit

Private Type ReceivingFixture
    RootPath As String
    ShareRoot As String
    WarehouseId As String
    StationId As String
    UserId As String
    OperatorPath As String
    SnapshotPath As String
    ConfigPath As String
    AuthPath As String
End Type

Public Function TestCheckReceivingReadiness_AllReady_ReturnsReady() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("all_ready")
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)
    modTS_Received.ResetReceivingUiStub

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    modReceivingInit.ApplyReceivingReadinessForWorkbook wbOps, True

    If PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "SnapshotStatus") = "OK" _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "OK" _
       And PackedValueReadinessTest(readinessPacked, "RuntimeStatus") = "OK" _
       And modTS_Received.GetReceivingUiStubInitializeCount() = 1 _
       And modReceivingInit.GetReceivingReadinessPanelText(wbOps) = "" Then
        TestCheckReceivingReadiness_AllReady_ReturnsReady = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_SnapshotOk_WhenAuthMissingCapability() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim wbAuth As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("snapshot_ok")
    Set wbAuth = OpenWorkbookReadinessTest(fx.AuthPath)
    RemoveCapabilityReadinessTest wbAuth, fx.UserId, "RECEIVE_POST"
    wbAuth.Save
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "SnapshotStatus") = "OK" _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "MISSING_CAPABILITY" _
       And PackedValueReadinessTest(readinessPacked, "RuntimeStatus") = "OK" Then
        TestCheckReceivingReadiness_SnapshotOk_WhenAuthMissingCapability = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps, wbAuth
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_SnapshotStale_ReturnsStale() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("snapshot_stale")
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)
    SetOperatorReadModelStateReadinessTest wbOps, DateAdd("h", -4, Now), True
    wbOps.Save

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "SnapshotStatus") = "STALE" _
       And InStr(1, PackedValueReadinessTest(readinessPacked, "Messages"), "Refresh Inventory before posting", vbTextCompare) > 0 Then
        TestCheckReceivingReadiness_SnapshotStale_ReturnsStale = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_SnapshotMissing_ReturnsMissing() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("snapshot_missing")
    DeleteFileReadinessTest fx.SnapshotPath
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "SnapshotStatus") = "MISSING" _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "OK" _
       And PackedValueReadinessTest(readinessPacked, "RuntimeStatus") = "OK" Then
        TestCheckReceivingReadiness_SnapshotMissing_ReturnsMissing = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_SnapshotUnreadable_ReturnsUnreadable() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("snapshot_unreadable")
    CorruptSnapshotFileReadinessTest fx.SnapshotPath
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "SnapshotStatus") = "UNREADABLE" _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "OK" _
       And PackedValueReadinessTest(readinessPacked, "RuntimeStatus") = "OK" Then
        TestCheckReceivingReadiness_SnapshotUnreadable_ReturnsUnreadable = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_AuthOk_WhenSnapshotMissing() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("auth_ok")
    DeleteFileReadinessTest fx.SnapshotPath
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "SnapshotStatus") = "MISSING" _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "OK" _
       And PackedValueReadinessTest(readinessPacked, "RuntimeStatus") = "OK" Then
        TestCheckReceivingReadiness_AuthOk_WhenSnapshotMissing = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_AuthNoUser_ReturnsNoUser() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim wbAuth As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("auth_nouser")
    Set wbAuth = OpenWorkbookReadinessTest(fx.AuthPath)
    RemoveUserReadinessTest wbAuth, fx.UserId
    wbAuth.Save
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "NO_USER" _
       And InStr(1, PackedValueReadinessTest(readinessPacked, "Messages"), "not provisioned", vbTextCompare) > 0 Then
        TestCheckReceivingReadiness_AuthNoUser_ReturnsNoUser = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps, wbAuth
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_AuthMissingCapability_ReturnsMissingCapability() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim wbAuth As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("auth_missingcap")
    Set wbAuth = OpenWorkbookReadinessTest(fx.AuthPath)
    RemoveCapabilityReadinessTest wbAuth, fx.UserId, "RECEIVE_POST"
    wbAuth.Save
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "MISSING_CAPABILITY" _
       And InStr(1, PackedValueReadinessTest(readinessPacked, "Messages"), "does not have RECEIVE_POST", vbTextCompare) > 0 Then
        TestCheckReceivingReadiness_AuthMissingCapability_ReturnsMissingCapability = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps, wbAuth
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_AuthInactive_ReturnsInactive() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim wbAuth As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("auth_inactive")
    Set wbAuth = OpenWorkbookReadinessTest(fx.AuthPath)
    SetUserStatusReadinessTest wbAuth, fx.UserId, "Disabled"
    wbAuth.Save
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "INACTIVE" _
       And InStr(1, PackedValueReadinessTest(readinessPacked, "Messages"), "inactive", vbTextCompare) > 0 Then
        TestCheckReceivingReadiness_AuthInactive_ReturnsInactive = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps, wbAuth
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_RuntimeOk_WhenSnapshotMissingAndNoUser() As Long
    Dim fx As ReceivingFixture
    Dim wbOps As Workbook
    Dim wbAuth As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    fx = CreateReceivingFixture("runtime_ok")
    DeleteFileReadinessTest fx.SnapshotPath
    Set wbAuth = OpenWorkbookReadinessTest(fx.AuthPath)
    RemoveUserReadinessTest wbAuth, fx.UserId
    wbAuth.Save
    Set wbOps = OpenWorkbookReadinessTest(fx.OperatorPath)

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "SnapshotStatus") = "MISSING" _
       And PackedValueReadinessTest(readinessPacked, "AuthStatus") = "NO_USER" _
       And PackedValueReadinessTest(readinessPacked, "RuntimeStatus") = "OK" Then
        TestCheckReceivingReadiness_RuntimeOk_WhenSnapshotMissingAndNoUser = 1
    End If

CleanExit:
    CleanupReceivingFixture fx, wbOps, wbAuth
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_RuntimeMissingTables_ReturnsMissingTables() As Long
    Dim wbOps As Workbook
    Dim readinessPacked As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "RuntimeStatus") = "MISSING_TABLES" _
       And InStr(1, readinessPacked, "missing required tables", vbTextCompare) > 0 Then
        TestCheckReceivingReadiness_RuntimeMissingTables_ReturnsMissingTables = 1
    End If

CleanExit:
    CloseWorkbookNoSaveReadinessTest wbOps
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCheckReceivingReadiness_RuntimePathUnresolved_ReturnsPathUnresolved() As Long
    Dim wbOps As Workbook
    Dim readinessPacked As String
    Dim report As String

    On Error GoTo CleanFail
    Set wbOps = Application.Workbooks.Add
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then
        Err.Raise vbObjectError + 7423, "TestCheckReceivingReadiness_RuntimePathUnresolved_ReturnsPathUnresolved", report
    End If

    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    If Not PackedBoolReadinessTest(readinessPacked, "IsReady") _
       And PackedValueReadinessTest(readinessPacked, "RuntimeStatus") = "PATH_UNRESOLVED" _
       And InStr(1, readinessPacked, "Runtime path could not be resolved", vbTextCompare) > 0 Then
        TestCheckReceivingReadiness_RuntimePathUnresolved_ReturnsPathUnresolved = 1
    End If

CleanExit:
    CloseWorkbookNoSaveReadinessTest wbOps
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function CreateReceivingFixture(ByVal caseToken As String) As ReceivingFixture
    Dim fx As ReceivingFixture
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbSnap As Workbook
    Dim wbOps As Workbook
    Dim report As String

    fx.WarehouseId = "WHRD" & Right$("0000" & CStr(Int((Timer * 1000) Mod 10000)), 4)
    fx.StationId = "R1"
    fx.UserId = ResolveCurrentReadinessUserTest()
    fx.RootPath = BuildTempRootReadinessTest(caseToken)
    fx.ShareRoot = fx.RootPath & "\sharepoint"
    fx.ConfigPath = fx.RootPath & "\" & fx.WarehouseId & ".invSys.Config.xlsb"
    fx.AuthPath = fx.RootPath & "\" & fx.WarehouseId & ".invSys.Auth.xlsb"
    fx.SnapshotPath = fx.RootPath & "\" & fx.WarehouseId & ".invSys.Snapshot.Inventory.xlsb"
    fx.OperatorPath = fx.RootPath & "\" & fx.WarehouseId & ".Receiving.Operator.xlsm"

    EnsureFolderReadinessTest fx.RootPath
    EnsureFolderReadinessTest fx.ShareRoot

    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook(fx.WarehouseId, fx.StationId, fx.RootPath, "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", fx.RootPath
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathSharePointRoot", fx.ShareRoot
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "AutoRefreshIntervalSeconds", 3600
    wbCfg.Save

    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook(fx.WarehouseId, fx.RootPath)
    TestPhase2Helpers.SetUserPinHash wbAuth, fx.UserId, modAuth.HashUserCredential("123456")
    TestPhase2Helpers.AddCapability wbAuth, fx.UserId, "RECEIVE_POST", fx.WarehouseId, fx.StationId, "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, fx.UserId, "RECEIVE_VIEW", fx.WarehouseId, fx.StationId, "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, fx.UserId, "READMODEL_REFRESH", fx.WarehouseId, fx.StationId, "ACTIVE"
    wbAuth.Save

    Set wbSnap = Application.Workbooks.Add
    wbSnap.Worksheets(1).Name = "InventorySnapshot"
    wbSnap.Worksheets(1).Range("A1:B2").Value = Array(Array("SKU", "QtyOnHand"), Array("TEST-SKU-001", 100))
    wbSnap.SaveAs Filename:=fx.SnapshotPath, FileFormat:=50

    Set wbOps = Application.Workbooks.Add
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then Err.Raise vbObjectError + 7410, "CreateReceivingFixture", report
    SeedReadModelMetadataReadinessTest wbOps, Now, False
    wbOps.SaveAs Filename:=fx.OperatorPath, FileFormat:=52

    CloseWorkbookNoSaveReadinessTest wbCfg
    CloseWorkbookNoSaveReadinessTest wbAuth
    CloseWorkbookNoSaveReadinessTest wbSnap
    CloseWorkbookNoSaveReadinessTest wbOps

    modRuntimeWorkbooks.SetCoreDataRootOverride fx.RootPath
    CreateReceivingFixture = fx
End Function

Private Sub SeedReadModelMetadataReadinessTest(ByVal wb As Workbook, _
                                               ByVal refreshUtc As Date, _
                                               ByVal isStale As Boolean)
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = wb.Worksheets("InventoryManagement").ListObjects("invSys")
    If lo.ListRows.Count = 0 Then lo.ListRows.Add
    rowIndex = 1
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("ITEM_CODE").Index).Value = "TEST-SKU-001"
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("ITEM").Index).Value = "TEST-SKU-001"
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("TOTAL INV").Index).Value = 100
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("QtyAvailable").Index).Value = 100
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("LocationSummary").Index).Value = "A1=100"
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("LastRefreshUTC").Index).Value = refreshUtc
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("SnapshotId").Index).Value = "SNAP-READY-001"
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("SourceType").Index).Value = "LOCAL"
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("IsStale").Index).Value = IIf(isStale, "TRUE", "FALSE")
End Sub

Private Sub SetOperatorReadModelStateReadinessTest(ByVal wb As Workbook, ByVal refreshUtc As Date, ByVal isStale As Boolean)
    Dim lo As ListObject

    Set lo = wb.Worksheets("InventoryManagement").ListObjects("invSys")
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
    lo.DataBodyRange.Cells(1, lo.ListColumns("LastRefreshUTC").Index).Value = refreshUtc
    lo.DataBodyRange.Cells(1, lo.ListColumns("IsStale").Index).Value = IIf(isStale, "TRUE", "FALSE")
End Sub

Private Function ResolveCurrentReadinessUserTest() As String
    ResolveCurrentReadinessUserTest = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If ResolveCurrentReadinessUserTest = "" Then ResolveCurrentReadinessUserTest = Trim$(Application.UserName)
End Function

Private Function BuildTempRootReadinessTest(ByVal caseToken As String) As String
    BuildTempRootReadinessTest = Environ$("TEMP") & "\invSys_readiness_" & caseToken & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Right$("0000" & CStr(CLng(Timer * 1000) Mod 10000), 4)
End Function

Private Function PackedValueReadinessTest(ByVal packedText As String, ByVal keyName As String) As String
    Dim parts() As String
    Dim i As Long
    Dim prefix As String

    prefix = keyName & "="
    parts = Split(packedText, "|")
    For i = LBound(parts) To UBound(parts)
        If StrComp(Left$(parts(i), Len(prefix)), prefix, vbTextCompare) = 0 Then
            PackedValueReadinessTest = Mid$(parts(i), Len(prefix) + 1)
            Exit Function
        End If
    Next i
End Function

Private Function PackedBoolReadinessTest(ByVal packedText As String, ByVal keyName As String) As Boolean
    PackedBoolReadinessTest = (StrComp(PackedValueReadinessTest(packedText, keyName), "True", vbTextCompare) = 0)
End Function

Private Function OpenWorkbookReadinessTest(ByVal workbookPath As String) As Workbook
    Set OpenWorkbookReadinessTest = Application.Workbooks.Open(Filename:=workbookPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
End Function

Private Sub RemoveCapabilityReadinessTest(ByVal wbAuth As Workbook, ByVal userId As String, ByVal capabilityName As String)
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = wbAuth.Worksheets("Capabilities").ListObjects("tblCapabilities")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = lo.ListRows.Count To 1 Step -1
        If StrComp(CStr(lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 _
           And StrComp(CStr(lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("Capability").Index).Value), capabilityName, vbTextCompare) = 0 Then
            lo.ListRows(rowIndex).Delete
        End If
    Next rowIndex
End Sub

Private Sub RemoveUserReadinessTest(ByVal wbAuth As Workbook, ByVal userId As String)
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim rowIndex As Long

    Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    Set loCaps = wbAuth.Worksheets("Capabilities").ListObjects("tblCapabilities")

    If Not loUsers.DataBodyRange Is Nothing Then
        For rowIndex = loUsers.ListRows.Count To 1 Step -1
            If StrComp(CStr(loUsers.DataBodyRange.Cells(rowIndex, loUsers.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 Then
                loUsers.ListRows(rowIndex).Delete
            End If
        Next rowIndex
    End If

    If Not loCaps.DataBodyRange Is Nothing Then
        For rowIndex = loCaps.ListRows.Count To 1 Step -1
            If StrComp(CStr(loCaps.DataBodyRange.Cells(rowIndex, loCaps.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 Then
                loCaps.ListRows(rowIndex).Delete
            End If
        Next rowIndex
    End If
End Sub

Private Sub SetUserStatusReadinessTest(ByVal wbAuth As Workbook, ByVal userId As String, ByVal statusText As String)
    Dim loUsers As ListObject
    Dim rowIndex As Long

    Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    If loUsers.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To loUsers.ListRows.Count
        If StrComp(CStr(loUsers.DataBodyRange.Cells(rowIndex, loUsers.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 Then
            loUsers.DataBodyRange.Cells(rowIndex, loUsers.ListColumns("Status").Index).Value = statusText
            Exit Sub
        End If
    Next rowIndex
End Sub

Private Sub DeleteTableReadinessTest(ByVal wb As Workbook, ByVal tableName As String)
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            lo.Delete
            Exit Sub
        End If
    Next ws
End Sub

Private Sub DeleteWorksheetReadinessTest(ByVal wb As Workbook, ByVal sheetName As String)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub CorruptSnapshotFileReadinessTest(ByVal filePath As String)
    Dim fileNum As Integer

    DeleteFileReadinessTest filePath
    fileNum = FreeFile
    Open filePath For Binary Access Write As #fileNum
    Put #fileNum, , "NOTANEXCELFILE"
    Close #fileNum
End Sub

Private Sub DeleteFileReadinessTest(ByVal filePath As String)
    On Error Resume Next
    Kill filePath
    On Error GoTo 0
End Sub

Private Sub EnsureFolderReadinessTest(ByVal folderPath As String)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub
    fso.CreateFolder folderPath
End Sub

Private Sub CleanupReceivingFixture(ByRef fx As ReceivingFixture, ParamArray workbooksToClose() As Variant)
    Dim i As Long

    For i = LBound(workbooksToClose) To UBound(workbooksToClose)
        If IsObject(workbooksToClose(i)) Then CloseWorkbookNoSaveReadinessTest workbooksToClose(i)
    Next i

    modRuntimeWorkbooks.ClearCoreDataRootOverride
    DeleteFolderTreeReadinessTest fx.RootPath
End Sub

Private Sub CloseWorkbookNoSaveReadinessTest(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub DeleteFolderTreeReadinessTest(ByVal folderPath As String)
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFolder folderPath, True
    On Error GoTo 0
End Sub
