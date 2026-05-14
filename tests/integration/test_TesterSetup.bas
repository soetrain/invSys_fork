Attribute VB_Name = "test_TesterSetup"
Option Explicit

Private mCaseNames() As String
Private mCaseResults() As String
Private mCaseDetails() As String
Private mCaseCount As Long
Private mSummary As String

Public Function TestTesterSetup_EndToEnd() As Long
    Dim detailText As String

    On Error GoTo FailTest

    ResetTesterSetupEvidence
    RecordTesterSetupCase "FreshMachine.CreatesRuntimeAndWorkbook", RunFreshMachineCase(detailText), detailText
    RecordTesterSetupCase "ExistingHub.CreatesNamespacedTesterRuntime", RunExistingHubCreatesTesterRuntimeCase(detailText), detailText
    RecordTesterSetupCase "IdempotentRerun.DoesNotDuplicateSeed", RunIdempotentRerunCase(detailText), detailText
    RecordTesterSetupCase "SharePointUnavailable.LocalSetupStillSucceeds", RunSharePointUnavailableCase(detailText), detailText
    RecordTesterSetupCase "ExistingAuth.HashPreservedCapabilitiesUpdated", RunExistingAuthCase(detailText), detailText

    If AllTesterSetupCasesPassed() Then
        mSummary = "Tester station setup passed fresh-machine, existing-hub, rerun-safe, offline-SharePoint, and existing-auth cases."
        TestTesterSetup_EndToEnd = 1
    Else
        mSummary = "One or more tester station setup cases failed."
    End If

    Exit Function

FailTest:
    RecordTesterSetupCase "Harness.Exception", False, Err.Description
    mSummary = "Tester station setup integration raised an unexpected exception."
End Function

Public Function GetTesterSetupContextPacked() As String
    GetTesterSetupContextPacked = "Summary=" & SafeTesterSetupText(mSummary)
End Function

Public Function GetTesterSetupEvidenceRows() As String
    Dim i As Long

    For i = 1 To mCaseCount
        If Len(GetTesterSetupEvidenceRows) > 0 Then GetTesterSetupEvidenceRows = GetTesterSetupEvidenceRows & vbLf
        GetTesterSetupEvidenceRows = GetTesterSetupEvidenceRows & _
            mCaseNames(i) & vbTab & mCaseResults(i) & vbTab & mCaseDetails(i)
    Next i
End Function

Private Function RunFreshMachineCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim shareRoot As String
    Dim templateRoot As String
    Dim spec As modTesterSetup.TesterSetupSpec
    Dim expectedHash As String

    warehouseId = "WHTSET_FRESH1"
    runtimeBase = BuildTesterSetupTempRoot("fresh")
    runtimeRoot = runtimeBase & "\runtime\" & warehouseId
    shareRoot = runtimeBase & "\sharepoint"
    templateRoot = runtimeBase & "\templates"
    expectedHash = modAuth.HashUserCredential("123456")

    On Error GoTo CleanFail
    EnsureFolderRecursiveTesterSetupCase shareRoot
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot

    spec = BuildTesterSetupSpecCase("tester.fresh", "123456", warehouseId, "R1", runtimeRoot, shareRoot)
    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If

    If Not AssertSetupArtifactsTesterSetupCase(spec, expectedHash, detailText) Then GoTo CleanExit

    RunFreshMachineCase = True
    detailText = "Fresh setup created the runtime tree, auth/config state, TEST-SKU-001 seed, and a valid receiving workbook."

CleanExit:
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    CleanupTesterSetupRoot runtimeBase
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunExistingHubCreatesTesterRuntimeCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim hubRoot As String
    Dim shareRoot As String
    Dim templateRoot As String
    Dim spec As modTesterSetup.TesterSetupSpec
    Dim expectedHash As String

    warehouseId = "TESTSTATION"
    runtimeBase = BuildTesterSetupTempRoot("existing_hub")
    hubRoot = runtimeBase & "\runtime\hub"
    shareRoot = runtimeBase & "\sharepoint"
    templateRoot = runtimeBase & "\templates"
    expectedHash = modAuth.HashUserCredential("111111")

    On Error GoTo CleanFail
    EnsureFolderRecursiveTesterSetupCase hubRoot
    EnsureFolderRecursiveTesterSetupCase shareRoot
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot

    spec = BuildTesterSetupSpecCase("tester.hub", "111111", warehouseId, "TS1", hubRoot, shareRoot)
    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If

    If InStr(1, modTesterSetup.GetLastTesterSetupReport(), "Runtime=CREATED", vbTextCompare) = 0 Then
        detailText = "Existing hub setup did not report tester runtime creation."
        GoTo CleanExit
    End If
    If Not AssertSetupArtifactsTesterSetupCase(spec, expectedHash, detailText) Then GoTo CleanExit

    RunExistingHubCreatesTesterRuntimeCase = True
    detailText = "Existing hub folder accepted a namespaced tester runtime without requiring matching warehouse artifacts first."

CleanExit:
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    CleanupTesterSetupRoot runtimeBase
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunIdempotentRerunCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim shareRoot As String
    Dim templateRoot As String
    Dim spec As modTesterSetup.TesterSetupSpec
    Dim qtyAfterFirst As Double
    Dim qtyAfterSecond As Double

    warehouseId = "WHTSET_RERUN1"
    runtimeBase = BuildTesterSetupTempRoot("rerun")
    runtimeRoot = runtimeBase & "\runtime\" & warehouseId
    shareRoot = runtimeBase & "\sharepoint"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    EnsureFolderRecursiveTesterSetupCase shareRoot
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot

    spec = BuildTesterSetupSpecCase("tester.rerun", "222222", warehouseId, "R1", runtimeRoot, shareRoot)
    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = "First setup failed: " & modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If
    qtyAfterFirst = ReadSkuQtyTesterSetupCase(runtimeRoot, warehouseId, "TEST-SKU-001")
    If qtyAfterFirst <> 100# Then
        detailText = "First setup did not seed QtyOnHand = 100."
        GoTo CleanExit
    End If

    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = "Second setup failed: " & modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If
    qtyAfterSecond = ReadSkuQtyTesterSetupCase(runtimeRoot, warehouseId, "TEST-SKU-001")
    If qtyAfterSecond <> 100# Then
        detailText = "Idempotent rerun duplicated the seed quantity."
        GoTo CleanExit
    End If
    If InStr(1, modTesterSetup.GetLastTesterSetupReport(), "Runtime=EXISTING", vbTextCompare) = 0 Then
        detailText = "Second setup did not report existing runtime reuse."
        GoTo CleanExit
    End If

    RunIdempotentRerunCase = True
    detailText = "Second setup reused the runtime and left TEST-SKU-001 at QtyOnHand = 100."

CleanExit:
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    CleanupTesterSetupRoot runtimeBase
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunSharePointUnavailableCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim templateRoot As String
    Dim spec As modTesterSetup.TesterSetupSpec
    Dim expectedHash As String

    warehouseId = "WHTSET_OFFLINE1"
    runtimeBase = BuildTesterSetupTempRoot("offline")
    runtimeRoot = runtimeBase & "\runtime\" & warehouseId
    templateRoot = runtimeBase & "\templates"
    expectedHash = modAuth.HashUserCredential("333333")

    On Error GoTo CleanFail
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot

    spec = BuildTesterSetupSpecCase("tester.offline", "333333", warehouseId, "R1", runtimeRoot, "C:\Invalid<SharePointRoot")
    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If

    If Not AssertSetupArtifactsTesterSetupCase(spec, expectedHash, detailText) Then GoTo CleanExit
    If Not AssertConfigSharePointTesterSetupCase(runtimeRoot, warehouseId, "C:\Invalid<SharePointRoot", detailText) Then GoTo CleanExit

    RunSharePointUnavailableCase = True
    detailText = "Local setup succeeded and recorded the unavailable SharePoint root without blocking runtime creation."

CleanExit:
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    CleanupTesterSetupRoot runtimeBase
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunExistingAuthCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim shareRoot As String
    Dim templateRoot As String
    Dim spec As modTesterSetup.TesterSetupSpec
    Dim originalHash As String

    warehouseId = "WHTSET_AUTH1"
    runtimeBase = BuildTesterSetupTempRoot("existing_auth")
    runtimeRoot = runtimeBase & "\runtime\" & warehouseId
    shareRoot = runtimeBase & "\sharepoint"
    templateRoot = runtimeBase & "\templates"
    originalHash = modAuth.HashUserCredential("444444")

    On Error GoTo CleanFail
    EnsureFolderRecursiveTesterSetupCase shareRoot
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot

    spec = BuildTesterSetupSpecCase("tester.auth", "444444", warehouseId, "R1", runtimeRoot, shareRoot)
    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = "First setup failed: " & modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If

    If Not MutateTesterAuthForRerun(runtimeRoot, warehouseId, "tester.auth", "R1", detailText) Then GoTo CleanExit

    spec.PinHash = modAuth.HashUserCredential("999999")
    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = "Second setup failed: " & modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If

    If Not AssertAuthStateTesterSetupCase(runtimeRoot, warehouseId, "tester.auth", "R1", originalHash, detailText) Then GoTo CleanExit

    RunExistingAuthCase = True
    detailText = "Existing tester auth kept the original hash and restored RECEIVE_POST, RECEIVE_VIEW, and READMODEL_REFRESH."

CleanExit:
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    CleanupTesterSetupRoot runtimeBase
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function BuildTesterSetupSpecCase(ByVal userId As String, _
                                          ByVal pinText As String, _
                                          ByVal warehouseId As String, _
                                          ByVal stationId As String, _
                                          ByVal pathLocal As String, _
                                          ByVal pathSharePointRoot As String) As modTesterSetup.TesterSetupSpec
    Dim spec As modTesterSetup.TesterSetupSpec

    spec.UserId = userId
    spec.PinHash = modAuth.HashUserCredential(pinText)
    spec.WarehouseId = warehouseId
    spec.StationId = stationId
    spec.PathLocal = pathLocal
    spec.PathSharePointRoot = pathSharePointRoot
    BuildTesterSetupSpecCase = spec
End Function

Private Function AssertSetupArtifactsTesterSetupCase(ByRef spec As modTesterSetup.TesterSetupSpec, _
                                                     ByVal expectedHash As String, _
                                                     ByRef detailText As String) As Boolean
    If Not FolderExistsTesterSetupCase(spec.PathLocal) Then
        detailText = "Runtime root missing."
        Exit Function
    End If
    If Not FolderExistsTesterSetupCase(spec.PathLocal & "\config") Then
        detailText = "config folder missing."
        Exit Function
    End If
    If Not FolderExistsTesterSetupCase(spec.PathLocal & "\auth") Then
        detailText = "auth folder missing."
        Exit Function
    End If
    If Not FolderExistsTesterSetupCase(spec.PathLocal & "\inbox") Then
        detailText = "inbox folder missing."
        Exit Function
    End If
    If Not FolderExistsTesterSetupCase(spec.PathLocal & "\outbox") Then
        detailText = "outbox folder missing."
        Exit Function
    End If
    If Not FolderExistsTesterSetupCase(spec.PathLocal & "\snapshots") Then
        detailText = "snapshots folder missing."
        Exit Function
    End If
    If Not FileExistsTesterSetupCase(spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb") Then
        detailText = "Inventory workbook missing."
        Exit Function
    End If
    If Not FileExistsTesterSetupCase(spec.PathLocal & "\" & spec.WarehouseId & ".Receiving.Operator.xlsm") Then
        detailText = "Receiving operator workbook missing."
        Exit Function
    End If
    If Not modTesterSetup.VerifyReceivingWorkbook(spec.PathLocal & "\" & spec.WarehouseId & ".Receiving.Operator.xlsm", detailText) Then Exit Function
    If Not AssertConfigSharePointTesterSetupCase(spec.PathLocal, spec.WarehouseId, spec.PathSharePointRoot, detailText) Then Exit Function
    If Not AssertAuthStateTesterSetupCase(spec.PathLocal, spec.WarehouseId, spec.UserId, spec.StationId, expectedHash, detailText) Then Exit Function
    If ReadSkuQtyTesterSetupCase(spec.PathLocal, spec.WarehouseId, "TEST-SKU-001") <> 100# Then
        detailText = "TEST-SKU-001 QtyOnHand was not 100."
        Exit Function
    End If

    AssertSetupArtifactsTesterSetupCase = True
End Function

Private Function AssertConfigSharePointTesterSetupCase(ByVal runtimeRoot As String, _
                                                       ByVal warehouseId As String, _
                                                       ByVal expectedShareRoot As String, _
                                                       ByRef detailText As String) As Boolean
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim actualValue As String

    On Error GoTo CleanFail
    Set wbCfg = OpenWorkbookTesterSetupCase(runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then
        detailText = "Config workbook could not be opened."
        GoTo CleanExit
    End If
    Set loWh = FindTableTesterSetupCase(wbCfg, "tblWarehouseConfig")
    If loWh Is Nothing Then
        detailText = "tblWarehouseConfig missing."
        GoTo CleanExit
    End If

    actualValue = CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "PathSharePointRoot"))
    If StrComp(Trim$(actualValue), Trim$(expectedShareRoot), vbTextCompare) <> 0 Then
        detailText = "PathSharePointRoot mismatch."
        GoTo CleanExit
    End If

    AssertConfigSharePointTesterSetupCase = True

CleanExit:
    CloseWorkbookTesterSetupCase wbCfg
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function AssertAuthStateTesterSetupCase(ByVal runtimeRoot As String, _
                                                ByVal warehouseId As String, _
                                                ByVal userId As String, _
                                                ByVal stationId As String, _
                                                ByVal expectedHash As String, _
                                                ByRef detailText As String) As Boolean
    Dim wbAuth As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim userRow As Long

    On Error GoTo CleanFail
    Set wbAuth = OpenWorkbookTesterSetupCase(runtimeRoot & "\" & warehouseId & ".invSys.Auth.xlsb")
    If wbAuth Is Nothing Then
        detailText = "Auth workbook could not be opened."
        GoTo CleanExit
    End If
    Set loUsers = FindTableTesterSetupCase(wbAuth, "tblUsers")
    Set loCaps = FindTableTesterSetupCase(wbAuth, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then
        detailText = "Auth tables missing."
        GoTo CleanExit
    End If

    userRow = FindRowByValueTesterSetupCase(loUsers, "UserId", userId)
    If userRow = 0 Then
        detailText = "Tester user row missing."
        GoTo CleanExit
    End If
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loUsers, userRow, "PinHash")), expectedHash, vbTextCompare) <> 0 Then
        detailText = "Tester PinHash did not match expected hash."
        GoTo CleanExit
    End If
    If StrComp(UCase$(CStr(TestPhase2Helpers.GetRowValue(loUsers, userRow, "Status"))), "ACTIVE", vbTextCompare) <> 0 Then
        detailText = "Tester user status was not ACTIVE."
        GoTo CleanExit
    End If
    If Not CapabilityIsActiveTesterSetupCase(loCaps, userId, "RECEIVE_POST", warehouseId, stationId) Then
        detailText = "RECEIVE_POST missing or inactive."
        GoTo CleanExit
    End If
    If Not CapabilityIsActiveTesterSetupCase(loCaps, userId, "RECEIVE_VIEW", warehouseId, stationId) Then
        detailText = "RECEIVE_VIEW missing or inactive."
        GoTo CleanExit
    End If
    If Not CapabilityIsActiveTesterSetupCase(loCaps, userId, "READMODEL_REFRESH", warehouseId, stationId) Then
        detailText = "READMODEL_REFRESH missing or inactive."
        GoTo CleanExit
    End If
    If CapabilityIsActiveTesterSetupCase(loCaps, userId, "ADMIN_MAINT", warehouseId, stationId) Then
        detailText = "ADMIN_MAINT should not remain active for tester user."
        GoTo CleanExit
    End If

    AssertAuthStateTesterSetupCase = True

CleanExit:
    CloseWorkbookTesterSetupCase wbAuth
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function MutateTesterAuthForRerun(ByVal runtimeRoot As String, _
                                          ByVal warehouseId As String, _
                                          ByVal userId As String, _
                                          ByVal stationId As String, _
                                          ByRef detailText As String) As Boolean
    Dim wbAuth As Workbook
    Dim loCaps As ListObject
    Dim rowIndex As Long

    On Error GoTo CleanFail
    Set wbAuth = OpenWorkbookTesterSetupCase(runtimeRoot & "\" & warehouseId & ".invSys.Auth.xlsb")
    If wbAuth Is Nothing Then
        detailText = "Auth workbook could not be opened for mutation."
        GoTo CleanExit
    End If
    Set loCaps = FindTableTesterSetupCase(wbAuth, "tblCapabilities")
    If loCaps Is Nothing Then
        detailText = "tblCapabilities missing."
        GoTo CleanExit
    End If

    rowIndex = FindCapabilityRowTesterSetupCase(loCaps, userId, "RECEIVE_POST", warehouseId, stationId)
    If rowIndex = 0 Then
        detailText = "RECEIVE_POST capability row missing."
        GoTo CleanExit
    End If
    loCaps.DataBodyRange.Cells(rowIndex, loCaps.ListColumns("Status").Index).Value = "INACTIVE"

    rowIndex = FindCapabilityRowTesterSetupCase(loCaps, userId, "READMODEL_REFRESH", warehouseId, stationId)
    If rowIndex > 0 Then loCaps.ListRows(rowIndex).Delete

    wbAuth.Save
    MutateTesterAuthForRerun = True

CleanExit:
    CloseWorkbookTesterSetupCase wbAuth
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function ReadSkuQtyTesterSetupCase(ByVal runtimeRoot As String, _
                                           ByVal warehouseId As String, _
                                           ByVal skuValue As String) As Double
    Dim wbInv As Workbook
    Dim loSku As ListObject
    Dim rowIndex As Long

    On Error GoTo CleanExit
    Set wbInv = OpenWorkbookTesterSetupCase(runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb")
    If wbInv Is Nothing Then GoTo CleanExit
    Set loSku = FindTableTesterSetupCase(wbInv, "tblSkuBalance")
    If loSku Is Nothing Then GoTo CleanExit
    rowIndex = FindRowByValueTesterSetupCase(loSku, "SKU", skuValue)
    If rowIndex = 0 Then GoTo CleanExit
    If IsNumeric(TestPhase2Helpers.GetRowValue(loSku, rowIndex, "QtyOnHand")) Then
        ReadSkuQtyTesterSetupCase = CDbl(TestPhase2Helpers.GetRowValue(loSku, rowIndex, "QtyOnHand"))
    End If

CleanExit:
    CloseWorkbookTesterSetupCase wbInv
End Function

Private Function CapabilityIsActiveTesterSetupCase(ByVal lo As ListObject, _
                                                   ByVal userId As String, _
                                                   ByVal capability As String, _
                                                   ByVal warehouseId As String, _
                                                   ByVal stationId As String) As Boolean
    Dim rowIndex As Long

    rowIndex = FindCapabilityRowTesterSetupCase(lo, userId, capability, warehouseId, stationId)
    If rowIndex = 0 Then Exit Function
    CapabilityIsActiveTesterSetupCase = (StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, rowIndex, "Status")), "ACTIVE", vbTextCompare) = 0)
End Function

Private Function FindCapabilityRowTesterSetupCase(ByVal lo As ListObject, _
                                                  ByVal userId As String, _
                                                  ByVal capability As String, _
                                                  ByVal warehouseId As String, _
                                                  ByVal stationId As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "UserId")), userId, vbTextCompare) = 0 _
           And StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "Capability")), capability, vbTextCompare) = 0 _
           And StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "WarehouseId")), warehouseId, vbTextCompare) = 0 _
           And StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, "StationId")), stationId, vbTextCompare) = 0 Then
            FindCapabilityRowTesterSetupCase = i
            Exit Function
        End If
    Next i
End Function

Private Function FindRowByValueTesterSetupCase(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByValueTesterSetupCase = i
            Exit Function
        End If
    Next i
End Function

Private Function OpenWorkbookTesterSetupCase(ByVal workbookPath As String) As Workbook
    If Len(Dir$(workbookPath, vbNormal)) = 0 Then Exit Function
    Set OpenWorkbookTesterSetupCase = Application.Workbooks.Open(workbookPath, False, False)
End Function

Private Function FindTableTesterSetupCase(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindTableTesterSetupCase = ws.ListObjects(tableName)
        If Not FindTableTesterSetupCase Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function BuildTesterSetupTempRoot(ByVal suffix As String) As String
    Randomize
    BuildTesterSetupTempRoot = Environ$("TEMP") & "\invSys_testersetup_" & suffix & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Rnd() * 10000), "0000")
End Function

Private Sub EnsureFolderRecursiveTesterSetupCase(ByVal folderPath As String)
    Dim parentPath As String
    Dim slashPos As Long

    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    slashPos = InStrRev(folderPath, "\")
    If slashPos > 3 Then
        parentPath = Left$(folderPath, slashPos - 1)
        If Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveTesterSetupCase parentPath
    End If
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function FolderExistsTesterSetupCase(ByVal folderPath As String) As Boolean
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    FolderExistsTesterSetupCase = (Len(Dir$(folderPath, vbDirectory)) > 0)
End Function

Private Function FileExistsTesterSetupCase(ByVal filePath As String) As Boolean
    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function
    FileExistsTesterSetupCase = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

Private Sub CleanupTesterSetupRoot(ByVal rootPath As String)
    Dim fso As Object

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FolderExists(rootPath) Then fso.DeleteFolder rootPath, True
    End If
    Set fso = Nothing
    On Error GoTo 0
End Sub

Private Sub CloseWorkbookTesterSetupCase(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub ResetTesterSetupEvidence()
    mCaseCount = 0
    Erase mCaseNames
    Erase mCaseResults
    Erase mCaseDetails
    mSummary = vbNullString
End Sub

Private Sub RecordTesterSetupCase(ByVal caseName As String, ByVal passed As Boolean, ByVal detailText As String)
    mCaseCount = mCaseCount + 1
    ReDim Preserve mCaseNames(1 To mCaseCount)
    ReDim Preserve mCaseResults(1 To mCaseCount)
    ReDim Preserve mCaseDetails(1 To mCaseCount)
    mCaseNames(mCaseCount) = caseName
    mCaseResults(mCaseCount) = IIf(passed, "PASS", "FAIL")
    mCaseDetails(mCaseCount) = SafeTesterSetupText(detailText)
End Sub

Private Function AllTesterSetupCasesPassed() As Boolean
    Dim i As Long

    If mCaseCount = 0 Then Exit Function
    For i = 1 To mCaseCount
        If mCaseResults(i) <> "PASS" Then Exit Function
    Next i
    AllTesterSetupCasesPassed = True
End Function

Private Function SafeTesterSetupText(ByVal textIn As String) As String
    SafeTesterSetupText = Replace$(Replace$(Trim$(textIn), vbCr, " "), vbLf, " ")
End Function
