Attribute VB_Name = "test_ReceivingReadiness"
Option Explicit

Private mSummary As String
Private mRows As String

Public Function TestReceivingReadiness_StatusPanelRendersForKnownBadWorkbook() As Long
    Dim rootPath As String
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim configPath As String
    Dim authPath As String
    Dim snapshotPath As String
    Dim operatorPath As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbSnap As Workbook
    Dim wbOps As Workbook
    Dim report As String
    Dim readinessPacked As String
    Dim panelText As String

    On Error GoTo FailTest

    warehouseId = "WHRDINT1"
    stationId = "R1"
    userId = ResolveCurrentIntegrationUserReadiness()
    rootPath = Environ$("TEMP") & "\invSys_receiving_readiness_integration_" & Format$(Now, "yyyymmdd_hhnnss")
    configPath = rootPath & "\" & warehouseId & ".invSys.Config.xlsb"
    authPath = rootPath & "\" & warehouseId & ".invSys.Auth.xlsb"
    snapshotPath = rootPath & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    operatorPath = rootPath & "\" & warehouseId & ".Receiving.Operator.xlsm"

    EnsureFolderIntegrationReadiness rootPath

    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook(warehouseId, stationId, rootPath, "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", rootPath
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "AutoRefreshIntervalSeconds", 3600
    wbCfg.Save

    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook(warehouseId, rootPath)
    TestPhase2Helpers.SetUserPinHash wbAuth, userId, modAuth.HashUserCredential("123456")
    TestPhase2Helpers.AddCapability wbAuth, userId, "READMODEL_REFRESH", warehouseId, stationId, "ACTIVE"
    wbAuth.Save

    Set wbSnap = Application.Workbooks.Add
    wbSnap.Worksheets(1).Name = "InventorySnapshot"
    wbSnap.Worksheets(1).Range("A1:B2").Value = Array(Array("SKU", "QtyOnHand"), Array("TEST-SKU-001", 100))
    wbSnap.SaveAs Filename:=snapshotPath, FileFormat:=50

    Set wbOps = Application.Workbooks.Add
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wbOps, report) Then GoTo CleanExit
    SeedReadModelRowIntegration wbOps
    wbOps.SaveAs Filename:=operatorPath, FileFormat:=52
    wbOps.Close SaveChanges:=False

    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    modRoleEventWriter.SetCurrentUserId userId
    modTS_Received.ResetReceivingUiStub
    Set wbOps = Application.Workbooks.Open(Filename:=operatorPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)

    modReceivingInit.EnsureReceivingSurfaceForWorkbook wbOps
    readinessPacked = modReceivingInit.CheckReceivingReadinessPacked(wbOps)
    panelText = modReceivingInit.GetReceivingReadinessPanelText(wbOps)

    If PackedValueIntegrationReadiness(readinessPacked, "AuthStatus") = "MISSING_CAPABILITY" _
       And PackedValueIntegrationReadiness(readinessPacked, "RuntimeStatus") = "OK" _
       And PackedValueIntegrationReadiness(readinessPacked, "SnapshotStatus") = "OK" _
       And InStr(1, panelText, "Receiving post", vbTextCompare) > 0 _
       And modTS_Received.GetReceivingUiStubInitializeCount() = 0 Then
        mSummary = "Receiving readiness rendered an actionable sheet-level status panel for a known-bad operator workbook."
        mRows = "KnownBadWorkbook.MissingCapability" & vbTab & "PASS" & vbTab & panelText
        TestReceivingReadiness_StatusPanelRendersForKnownBadWorkbook = 1
    Else
        mSummary = "Receiving readiness did not render the expected status panel."
        mRows = "KnownBadWorkbook.MissingCapability" & vbTab & "FAIL" & vbTab & "Status=" & PackedValueIntegrationReadiness(readinessPacked, "AuthStatus") & "|" & panelText
    End If

CleanExit:
    On Error Resume Next
    If Not wbOps Is Nothing Then wbOps.Close SaveChanges:=False
    If Not wbSnap Is Nothing Then wbSnap.Close SaveChanges:=False
    If Not wbAuth Is Nothing Then wbAuth.Close SaveChanges:=False
    If Not wbCfg Is Nothing Then wbCfg.Close SaveChanges:=False
    modAuth.SignOut
    modRoleEventWriter.SetCurrentUserId vbNullString
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    DeleteFolderIntegrationReadiness rootPath
    Exit Function

FailTest:
    mSummary = "Receiving readiness integration raised an unexpected exception."
    mRows = "KnownBadWorkbook.MissingCapability" & vbTab & "FAIL" & vbTab & Err.Description
    Resume CleanExit
End Function

Public Function GetReceivingReadinessContextPacked() As String
    GetReceivingReadinessContextPacked = "Summary=" & mSummary
End Function

Public Function GetReceivingReadinessEvidenceRows() As String
    GetReceivingReadinessEvidenceRows = mRows
End Function

Private Function ResolveCurrentIntegrationUserReadiness() As String
    ResolveCurrentIntegrationUserReadiness = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If ResolveCurrentIntegrationUserReadiness = "" Then ResolveCurrentIntegrationUserReadiness = "readiness_user"
End Function

Private Sub SeedReadModelRowIntegration(ByVal wb As Workbook)
    Dim lo As ListObject

    Set lo = wb.Worksheets("InventoryManagement").ListObjects("invSys")
    If lo.ListRows.Count = 0 Then lo.ListRows.Add
    lo.DataBodyRange.Cells(1, lo.ListColumns("ITEM_CODE").Index).Value = "TEST-SKU-001"
    lo.DataBodyRange.Cells(1, lo.ListColumns("ITEM").Index).Value = "TEST-SKU-001"
    lo.DataBodyRange.Cells(1, lo.ListColumns("TOTAL INV").Index).Value = 100
    lo.DataBodyRange.Cells(1, lo.ListColumns("QtyAvailable").Index).Value = 100
    lo.DataBodyRange.Cells(1, lo.ListColumns("LocationSummary").Index).Value = "A1=100"
    lo.DataBodyRange.Cells(1, lo.ListColumns("LastRefreshUTC").Index).Value = Now
    lo.DataBodyRange.Cells(1, lo.ListColumns("SnapshotId").Index).Value = "SNAP-INT-001"
    lo.DataBodyRange.Cells(1, lo.ListColumns("SourceType").Index).Value = "LOCAL"
    lo.DataBodyRange.Cells(1, lo.ListColumns("IsStale").Index).Value = "FALSE"
End Sub

Private Sub EnsureFolderIntegrationReadiness(ByVal folderPath As String)
    If Trim$(folderPath) = "" Then Exit Sub
    If CreateObject("Scripting.FileSystemObject").FolderExists(folderPath) Then Exit Sub
    CreateObject("Scripting.FileSystemObject").CreateFolder folderPath
End Sub

Private Sub DeleteFolderIntegrationReadiness(ByVal folderPath As String)
    On Error Resume Next
    If Trim$(folderPath) <> "" Then CreateObject("Scripting.FileSystemObject").DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

Private Function PackedValueIntegrationReadiness(ByVal packedText As String, ByVal keyName As String) As String
    Dim parts() As String
    Dim i As Long
    Dim prefix As String

    prefix = keyName & "="
    parts = Split(packedText, "|")
    For i = LBound(parts) To UBound(parts)
        If StrComp(Left$(parts(i), Len(prefix)), prefix, vbTextCompare) = 0 Then
            PackedValueIntegrationReadiness = Mid$(parts(i), Len(prefix) + 1)
            Exit Function
        End If
    Next i
End Function
