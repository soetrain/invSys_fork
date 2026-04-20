Attribute VB_Name = "modTesterBundle"
Option Explicit

Private Const TESTER_BUNDLE_VERSION As String = "1.0"
Private Const TESTER_STATION_ID As String = "R1"
Private Const TESTER_SKU As String = "TEST-SKU-001"
Private Const TESTER_README_NAME As String = "TesterReadme.md"
Private Const TESTER_CONFIG_FILE As String = "tblWarehouseConfig.csv"
Private Const TESTER_SEED_FILE As String = "TEST-SKU-001.seed.json"
Private Const TESTER_AUTH_TEMPLATE_FILE As String = "tester-auth-template.csv"
Private Const TESTER_MANIFEST_FILE As String = "manifest.json"
Private Const TESTER_ADDINS_MANIFEST_FILE As String = "addins-manifest.json"
Private mLastTesterBundleReport As String
Private mLastTesterBundleZipPath As String
Private mLastTesterBundleReadmePath As String
Private mTesterBundleSharePointRootOverride As String

Public Function WriteTesterBundle(ByVal warehouseId As String, ByVal destDir As String) As Boolean
    Dim runtimeRoot As String
    Dim configPath As String
    Dim warehouseName As String
    Dim sharePointRoot As String
    Dim stageRoot As String
    Dim tempRoot As String
    Dim tempZipPath As String
    Dim finalZipPath As String
    Dim tempReadmePath As String
    Dim finalReadmePath As String
    Dim report As String
    Dim expectedEntries As Variant

    On Error GoTo FailWrite

    ResetLastTesterBundleState

    warehouseId = Trim$(warehouseId)
    destDir = NormalizeFolderPathTester(destDir)
    If warehouseId = "" Then
        report = "WarehouseId is required."
        GoTo FailSoft
    End If
    If destDir = "" Then
        report = "Destination folder is required."
        GoTo FailSoft
    End If

    runtimeRoot = ResolveRuntimeRootTester(warehouseId)
    If runtimeRoot = "" Then
        report = "Runtime root not found for warehouse " & warehouseId
        GoTo FailSoft
    End If

    configPath = runtimeRoot & warehouseId & ".invSys.Config.xlsb"
    If Not ReadBundleConfigTester(configPath, warehouseName, sharePointRoot, report) Then GoTo FailSoft

    EnsureFolderRecursiveTester destDir
    stageRoot = BuildWorkingFolderTester("stage")
    tempRoot = BuildWorkingFolderTester("temp")
    EnsureFolderRecursiveTester stageRoot
    EnsureFolderRecursiveTester tempRoot
    If Not WriteBundleStageFilesTester(stageRoot, warehouseId, warehouseName, report) Then GoTo FailSoft

    tempZipPath = tempRoot & warehouseId & "_TesterBundle_" & Format$(Date, "yyyymmdd") & ".zip"
    finalZipPath = destDir & warehouseId & "_TesterBundle_" & Format$(Date, "yyyymmdd") & ".zip"
    tempReadmePath = tempRoot & warehouseId & "_TesterReadme_" & Format$(Date, "yyyymmdd") & ".md"
    finalReadmePath = destDir & warehouseId & "_TesterReadme_" & Format$(Date, "yyyymmdd") & ".md"
    WriteTextFileTester tempReadmePath, GetTesterReadmeText(warehouseId)

    expectedEntries = GetExpectedBundleEntriesTester()
    If Not BuildZipFromFolderTester(stageRoot, tempZipPath, expectedEntries, report) Then GoTo FailSoft
    If Not VerifyTesterBundle(tempZipPath) Then
        report = GetLastTesterBundleReport()
        GoTo FailSoft
    End If

    If Not CommitFileAtomicallyTester(tempZipPath, finalZipPath, report) Then GoTo FailSoft
    If Not CommitFileAtomicallyTester(tempReadmePath, finalReadmePath, report) Then GoTo FailSoft

    WriteTesterBundle = True
    report = "OK|Zip=" & finalZipPath & "|Readme=" & finalReadmePath
    mLastTesterBundleZipPath = finalZipPath
    mLastTesterBundleReadmePath = finalReadmePath
    GoTo CleanExit

FailSoft:
    WriteTesterBundle = False
    If Len(report) = 0 Then report = "WriteTesterBundle failed."
    LogDiagnosticEvent "TESTER-BUNDLE", report
    GoTo CleanExit

FailWrite:
    report = "WriteTesterBundle failed: " & Err.Description
    Resume FailSoft

CleanExit:
    mLastTesterBundleReport = report
    DeleteFolderRecursiveTester stageRoot
    DeleteFolderRecursiveTester tempRoot
End Function

Public Function VerifyTesterBundle(ByVal zipPath As String) As Boolean
    Dim extractRoot As String
    Dim expectedEntries As Variant
    Dim i As Long
    Dim manifestPath As String
    Dim manifestText As String
    Dim report As String

    On Error GoTo FailVerify

    zipPath = Trim$(zipPath)
    If zipPath = "" Then
        report = "Bundle zip path is required."
        GoTo FailSoft
    End If
    If Not FileExistsTester(zipPath) Then
        report = "Bundle zip not found: " & zipPath
        GoTo FailSoft
    End If
    If SafeFileLenTester(zipPath) <= 0 Then
        report = "Bundle zip is zero-byte: " & zipPath
        GoTo FailSoft
    End If

    extractRoot = BuildWorkingFolderTester("verify")
    EnsureFolderRecursiveTester extractRoot
    If Not ExtractTesterBundleToFolder(zipPath, extractRoot) Then
        report = GetLastTesterBundleReport()
        GoTo FailSoft
    End If

    expectedEntries = GetExpectedBundleEntriesTester()
    For i = LBound(expectedEntries) To UBound(expectedEntries)
        If Not FileExistsTester(extractRoot & Replace$(CStr(expectedEntries(i)), "/", "\")) Then
            report = "Bundle entry missing: " & CStr(expectedEntries(i))
            GoTo FailSoft
        End If
    Next i

    manifestPath = extractRoot & TESTER_MANIFEST_FILE
    manifestText = ReadAllTextTester(manifestPath)
    If Not ManifestLooksValidTester(manifestText) Then
        report = "Bundle manifest is missing required keys. Text=" & Left$(Replace$(Replace$(manifestText, vbCr, " "), vbLf, " "), 240)
        GoTo FailSoft
    End If

    VerifyTesterBundle = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    VerifyTesterBundle = False
    If Len(report) = 0 Then report = "VerifyTesterBundle failed."
    LogDiagnosticEvent "TESTER-BUNDLE", report
    GoTo CleanExit

FailVerify:
    report = "VerifyTesterBundle failed: " & Err.Description
    Resume FailSoft

CleanExit:
    mLastTesterBundleReport = report
    DeleteFolderRecursiveTester extractRoot
End Function

Public Function ExtractTesterBundleToFolder(ByVal zipPath As String, ByVal outputRoot As String) As Boolean
    Dim expectedEntries As Variant
    Dim report As String
    Dim commandText As String

    On Error GoTo FailExtract

    zipPath = Trim$(zipPath)
    outputRoot = NormalizeFolderPathTester(outputRoot)
    If zipPath = "" Or outputRoot = "" Then
        report = "Zip path and output folder are required."
        GoTo FailSoft
    End If

    EnsureFolderRecursiveTester outputRoot
    commandText = "Expand-Archive -LiteralPath " & QuoteForPowerShellTester(zipPath) & _
                  " -DestinationPath " & QuoteForPowerShellTester(outputRoot) & " -Force"
    If Not RunPowerShellCommandTester(commandText, report) Then
        report = "Bundle extract failed: " & report
        GoTo FailSoft
    End If
    expectedEntries = GetExpectedBundleEntriesTester()
    If Not WaitForExpectedEntriesTester(outputRoot, expectedEntries, 10, report) Then GoTo FailSoft

    ExtractTesterBundleToFolder = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    ExtractTesterBundleToFolder = False
    If Len(report) = 0 Then report = "ExtractTesterBundleToFolder failed."
    LogDiagnosticEvent "TESTER-BUNDLE", report
    GoTo CleanExit

FailExtract:
    report = "ExtractTesterBundleToFolder failed: " & Err.Description
    Resume FailSoft

CleanExit:
    mLastTesterBundleReport = report
End Function

Public Function PublishTesterBundle(ByVal warehouseId As String) As Boolean
    Dim tempDest As String
    Dim sharePointRoot As String
    Dim localZipPath As String
    Dim localReadmePath As String
    Dim zipTargetPath As String
    Dim readmeTargetPath As String
    Dim zipStatus As String
    Dim readmeStatus As String
    Dim manifestTempPath As String
    Dim manifestTargetPath As String
    Dim manifestText As String
    Dim report As String

    On Error GoTo FailPublish

    ResetLastTesterBundleState
    warehouseId = Trim$(warehouseId)
    If warehouseId = "" Then
        report = "WarehouseId is required."
        GoTo FailSoft
    End If

    tempDest = BuildWorkingFolderTester("publish")
    EnsureFolderRecursiveTester tempDest
    If Not WriteTesterBundle(warehouseId, tempDest) Then
        report = GetLastTesterBundleReport()
        GoTo FailSoft
    End If

    localZipPath = mLastTesterBundleZipPath
    localReadmePath = mLastTesterBundleReadmePath
    sharePointRoot = ResolveSharePointRootTester(warehouseId)
    If sharePointRoot = "" Then
        report = "PathSharePointRoot not configured."
        GoTo FailSoft
    End If

    zipTargetPath = sharePointRoot & "TesterPackage\" & warehouseId & "\" & warehouseId & ".TesterBundle.zip"
    readmeTargetPath = sharePointRoot & "TesterPackage\" & warehouseId & "\" & warehouseId & ".TesterReadme.md"
    If Not modWarehouseSync.PublishFileToTargetPath(localZipPath, zipTargetPath, zipStatus) Then
        report = "Bundle publish failed: " & zipStatus
        GoTo FailSoft
    End If
    If Not modWarehouseSync.PublishFileToTargetPath(localReadmePath, readmeTargetPath, readmeStatus) Then
        report = "Readme publish failed: " & readmeStatus
        GoTo FailSoft
    End If

    manifestText = BuildAddinsManifestWithBundleTimestampTester(sharePointRoot, warehouseId)
    manifestTempPath = tempDest & "addins-manifest.json"
    manifestTargetPath = sharePointRoot & "Addins\" & TESTER_ADDINS_MANIFEST_FILE
    WriteTextFileTester manifestTempPath, manifestText
    If Not modWarehouseSync.PublishFileToTargetPath(manifestTempPath, manifestTargetPath, report) Then
        report = "Add-ins manifest update failed: " & report
        GoTo FailSoft
    End If

    PublishTesterBundle = True
    report = "OK|Bundle=" & zipStatus & "|Readme=" & readmeStatus
    GoTo CleanExit

FailSoft:
    PublishTesterBundle = False
    If Len(report) = 0 Then report = "PublishTesterBundle failed."
    LogDiagnosticEvent "TESTER-BUNDLE", report
    GoTo CleanExit

FailPublish:
    report = "PublishTesterBundle failed: " & Err.Description
    Resume FailSoft

CleanExit:
    mLastTesterBundleReport = report
    DeleteFolderRecursiveTester tempDest
End Function

Public Function GetLastTesterBundleReport() As String
    GetLastTesterBundleReport = mLastTesterBundleReport
End Function

Public Function GetLastTesterBundleZipPath() As String
    GetLastTesterBundleZipPath = mLastTesterBundleZipPath
End Function

Public Function GetLastTesterBundleReadmePath() As String
    GetLastTesterBundleReadmePath = mLastTesterBundleReadmePath
End Function

Public Sub SetTesterBundleSharePointRootOverride(ByVal rootPath As String)
    mTesterBundleSharePointRootOverride = Trim$(rootPath)
End Sub

Public Sub ClearTesterBundleSharePointRootOverride()
    mTesterBundleSharePointRootOverride = vbNullString
End Sub

Private Sub ResetLastTesterBundleState()
    mLastTesterBundleReport = vbNullString
    mLastTesterBundleZipPath = vbNullString
    mLastTesterBundleReadmePath = vbNullString
End Sub

Private Function ResolveRuntimeRootTester(ByVal warehouseId As String) As String
    Dim overrideRoot As String
    Dim defaultRoot As String

    overrideRoot = NormalizeFolderPathTester(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If overrideRoot <> "" And FolderExistsTester(overrideRoot) Then
        ResolveRuntimeRootTester = overrideRoot
        Exit Function
    End If

    defaultRoot = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(warehouseId, True)
    If FolderExistsTester(defaultRoot) Then ResolveRuntimeRootTester = defaultRoot
End Function

Private Function ResolveSharePointRootTester(ByVal warehouseId As String) As String
    Dim runtimeRoot As String
    Dim configPath As String
    Dim warehouseName As String
    Dim sharePointRoot As String
    Dim report As String

    ResolveSharePointRootTester = NormalizeFolderPathTester(mTesterBundleSharePointRootOverride)
    If ResolveSharePointRootTester <> "" Then Exit Function

    sharePointRoot = NormalizeFolderPathTester(modConfig.GetString("PathSharePointRoot", ""))
    If sharePointRoot <> "" Then
        ResolveSharePointRootTester = sharePointRoot
        Exit Function
    End If

    runtimeRoot = ResolveRuntimeRootTester(warehouseId)
    If runtimeRoot = "" Then Exit Function
    configPath = runtimeRoot & warehouseId & ".invSys.Config.xlsb"
    If Not ReadBundleConfigTester(configPath, warehouseName, sharePointRoot, report) Then Exit Function
    ResolveSharePointRootTester = NormalizeFolderPathTester(sharePointRoot)
End Function

Private Function ReadBundleConfigTester(ByVal configPath As String, _
                                        ByRef warehouseName As String, _
                                        ByRef sharePointRoot As String, _
                                        ByRef report As String) As Boolean
    Dim wbCfg As Workbook
    Dim openedTransient As Boolean
    Dim loWh As ListObject
    Dim loSt As ListObject

    On Error GoTo FailRead

    If Not FileExistsTester(configPath) Then
        report = "Config workbook not found: " & configPath
        Exit Function
    End If

    Set wbCfg = OpenWorkbookByPathTester(configPath, openedTransient)
    If wbCfg Is Nothing Then
        report = "Config workbook could not be opened: " & configPath
        GoTo CleanExit
    End If

    Set loWh = GetListObjectTester(wbCfg, "WarehouseConfig", "tblWarehouseConfig")
    Set loSt = GetListObjectTester(wbCfg, "StationConfig", "tblStationConfig")
    If loWh Is Nothing Or loSt Is Nothing Then
        report = "Required config tables were not found."
        GoTo CleanExit
    End If

    warehouseName = CStr(GetTableValueTester(loWh, 1, "WarehouseName"))
    sharePointRoot = CStr(GetTableValueTester(loWh, 1, "PathSharePointRoot"))
    If warehouseName = "" Then warehouseName = CStr(GetTableValueTester(loWh, 1, "WarehouseId"))

    ReadBundleConfigTester = True
    report = "OK"

CleanExit:
    If openedTransient Then CloseWorkbookQuietlyTester wbCfg
    Exit Function

FailRead:
    report = "ReadBundleConfigTester failed: " & Err.Description
    Resume CleanExit
End Function

Private Function WriteBundleStageFilesTester(ByVal stageRoot As String, _
                                             ByVal warehouseId As String, _
                                             ByVal warehouseName As String, _
                                             ByRef report As String) As Boolean
    Dim configDir As String
    Dim seedDir As String
    Dim authDir As String

    On Error GoTo FailStage

    configDir = stageRoot & "config\"
    seedDir = stageRoot & "seed\"
    authDir = stageRoot & "auth\"
    EnsureFolderRecursiveTester configDir
    EnsureFolderRecursiveTester seedDir
    EnsureFolderRecursiveTester authDir

    WriteTextFileTester stageRoot & TESTER_README_NAME, GetTesterReadmeText(warehouseId)
    WriteTextFileTester configDir & TESTER_CONFIG_FILE, BuildSanitizedConfigCsvTester(warehouseId, warehouseName)
    WriteTextFileTester seedDir & TESTER_SEED_FILE, BuildSeedJsonTester(warehouseId)
    WriteTextFileTester authDir & TESTER_AUTH_TEMPLATE_FILE, BuildAuthTemplateCsvTester()
    WriteTextFileTester stageRoot & TESTER_MANIFEST_FILE, BuildBundleManifestTester(warehouseId)

    WriteBundleStageFilesTester = True
    report = "OK"
    Exit Function

FailStage:
    report = "WriteBundleStageFilesTester failed: " & Err.Description
End Function

Private Function GetTesterReadmeText(ByVal warehouseId As String) As String
    Dim lines(0 To 9) As String

    lines(0) = "# invSys " & warehouseId & " Tester Bundle"
    lines(1) = "1. Download add-ins from Addins/ on SharePoint"
    lines(2) = "2. Open invSys.Admin.xlam"
    lines(3) = "3. Click Setup Tester Station"
    lines(4) = "4. Enter your UserId and a PIN when prompted"
    lines(5) = "5. Click Open Receiving Workbook"
    lines(6) = "6. Click Refresh Inventory - confirm TEST-SKU-001 shows QtyOnHand = 100"
    lines(7) = "7. Enter: SKU = TEST-SKU-001, Qty = 10"
    lines(8) = "8. Click Confirm Writes"
    lines(9) = "9. Expected result: row marked CONFIRMED, QtyOnHand updates to 110 after processor runs"
    GetTesterReadmeText = Join(lines, vbCrLf)
End Function

Private Function BuildSanitizedConfigCsvTester(ByVal warehouseId As String, ByVal warehouseName As String) As String
    BuildSanitizedConfigCsvTester = "WarehouseId,WarehouseName,StationId" & vbCrLf & _
        CsvValueTester(warehouseId) & "," & CsvValueTester(warehouseName) & "," & CsvValueTester(TESTER_STATION_ID)
End Function

Private Function BuildSeedJsonTester(ByVal warehouseId As String) As String
    BuildSeedJsonTester = "{" & vbCrLf & _
        "  ""SKU"": """ & EscapeJsonTester(TESTER_SKU) & """," & vbCrLf & _
        "  ""Description"": ""Test SKU for Confirm Writes""," & vbCrLf & _
        "  ""QtyOnHand"": 100," & vbCrLf & _
        "  ""WarehouseId"": """ & EscapeJsonTester(warehouseId) & """," & vbCrLf & _
        "  ""StationId"": """ & EscapeJsonTester(TESTER_STATION_ID) & """" & vbCrLf & _
        "}"
End Function

Private Function BuildAuthTemplateCsvTester() As String
    BuildAuthTemplateCsvTester = "UserId,WarehouseId,StationId,PasswordHash,Capabilities,Status" & vbCrLf
End Function

Private Function BuildBundleManifestTester(ByVal warehouseId As String) As String
    Dim requiredAddins As Variant
    Dim i As Long
    Dim lines() As String

    requiredAddins = GetRequiredAddinsTester()
    ReDim lines(0 To UBound(requiredAddins) + 7)
    lines(0) = "{"
    lines(1) = "  ""BundleVersion"": """ & TESTER_BUNDLE_VERSION & ""","
    lines(2) = "  ""WarehouseId"": """ & EscapeJsonTester(warehouseId) & ""","
    lines(3) = "  ""CreatedUTC"": """ & EscapeJsonTester(Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")) & ""","
    lines(4) = "  ""RequiredAddins"": ["
    For i = LBound(requiredAddins) To UBound(requiredAddins)
        lines(5 + i) = "    """ & EscapeJsonTester(CStr(requiredAddins(i))) & """" & IIf(i < UBound(requiredAddins), ",", "")
    Next i
    lines(UBound(requiredAddins) + 6) = "  ]"
    lines(UBound(requiredAddins) + 7) = "}"
    BuildBundleManifestTester = Join(lines, vbCrLf)
End Function

Private Function BuildAddinsManifestWithBundleTimestampTester(ByVal sharePointRoot As String, ByVal warehouseId As String) As String
    Dim addinNames As Variant
    Dim i As Long
    Dim addinsRoot As String
    Dim lines() As String
    Dim idx As Long
    Dim addinPath As String

    addinsRoot = NormalizeFolderPathTester(sharePointRoot) & "Addins\"
    addinNames = GetRequiredAddinsTester()
    ReDim lines(0 To UBound(addinNames) + 8)

    idx = 0
    lines(idx) = "{": idx = idx + 1
    lines(idx) = "  ""published_utc"": """ & EscapeJsonTester(Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")) & """,": idx = idx + 1
    lines(idx) = "  ""tester_bundle_published_utc"": """ & EscapeJsonTester(Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")) & """,": idx = idx + 1
    lines(idx) = "  ""tester_bundle_warehouse_id"": """ & EscapeJsonTester(warehouseId) & """,": idx = idx + 1
    lines(idx) = "  ""files"": [": idx = idx + 1
    For i = LBound(addinNames) To UBound(addinNames)
        addinPath = addinsRoot & CStr(addinNames(i))
        lines(idx) = "    { ""name"": """ & EscapeJsonTester(CStr(addinNames(i))) & """, ""size_bytes"": " & CStr(SafeFileLenTester(addinPath)) & " }" & IIf(i < UBound(addinNames), ",", "")
        idx = idx + 1
    Next i
    lines(idx) = "  ]": idx = idx + 1
    lines(idx) = "}": idx = idx + 1
    ReDim Preserve lines(0 To idx - 1)
    BuildAddinsManifestWithBundleTimestampTester = Join(lines, vbCrLf)
End Function

Private Function BuildZipFromFolderTester(ByVal sourceFolder As String, _
                                          ByVal zipPath As String, _
                                          ByVal expectedEntries As Variant, _
                                          ByRef report As String) As Boolean
    Dim commandText As String

    On Error GoTo FailZip

    EnsureFolderRecursiveTester ParentFolderTester(zipPath)
    DeleteFileIfPresentTester zipPath
    commandText = "Compress-Archive -Path " & QuoteForPowerShellTester(sourceFolder & "*") & _
                  " -DestinationPath " & QuoteForPowerShellTester(zipPath) & " -Force"
    If Not RunPowerShellCommandTester(commandText, report) Then
        report = "Bundle zip build failed: " & report
        GoTo CleanExit
    End If
    If Not WaitForZipReadyTester(zipPath, expectedEntries, 10, report) Then GoTo CleanExit

    BuildZipFromFolderTester = True
    report = "OK"

CleanExit:
    Exit Function

FailZip:
    report = "BuildZipFromFolderTester failed: " & Err.Description
    Resume CleanExit
End Function

Private Function WaitForZipReadyTester(ByVal zipPath As String, ByVal expectedEntries As Variant, ByVal timeoutSeconds As Single, ByRef report As String) As Boolean
    Dim startTime As Single

    startTime = Timer
    Do
        If VerifyTesterBundleFastTester(zipPath, expectedEntries) Then
            WaitForZipReadyTester = True
            report = "OK"
            Exit Function
        End If
        DoEvents
    Loop While TimerElapsedTester(startTime) < timeoutSeconds

    report = "Bundle zip did not become ready before timeout."
End Function

Private Function VerifyTesterBundleFastTester(ByVal zipPath As String, ByVal expectedEntries As Variant) As Boolean
    Dim extractRoot As String
    Dim i As Long

    extractRoot = BuildWorkingFolderTester("verifyfast")
    EnsureFolderRecursiveTester extractRoot
    If Not ExtractZipToFolderFastTester(zipPath, extractRoot, expectedEntries) Then GoTo CleanExit

    For i = LBound(expectedEntries) To UBound(expectedEntries)
        If Not FileExistsTester(extractRoot & Replace$(CStr(expectedEntries(i)), "/", "\")) Then GoTo CleanExit
    Next i
    VerifyTesterBundleFastTester = True

CleanExit:
    DeleteFolderRecursiveTester extractRoot
End Function

Private Function ExtractZipToFolderFastTester(ByVal zipPath As String, ByVal outputRoot As String, ByVal expectedEntries As Variant) As Boolean
    Dim report As String
    Dim commandText As String

    On Error GoTo FailExtract

    commandText = "Expand-Archive -LiteralPath " & QuoteForPowerShellTester(zipPath) & _
                  " -DestinationPath " & QuoteForPowerShellTester(outputRoot) & " -Force"
    If Not RunPowerShellCommandTester(commandText, report) Then GoTo CleanExit
    If Not WaitForExpectedEntriesTester(outputRoot, expectedEntries, 2, report) Then GoTo CleanExit
    ExtractZipToFolderFastTester = True

CleanExit:
    Exit Function

FailExtract:
    Resume CleanExit
End Function

Private Function WaitForExpectedEntriesTester(ByVal outputRoot As String, ByVal expectedEntries As Variant, ByVal timeoutSeconds As Single, ByRef report As String) As Boolean
    Dim startTime As Single
    Dim i As Long
    Dim allFound As Boolean

    startTime = Timer
    Do
        allFound = True
        For i = LBound(expectedEntries) To UBound(expectedEntries)
            If Not FileExistsTester(outputRoot & Replace$(CStr(expectedEntries(i)), "/", "\")) Then
                allFound = False
                Exit For
            End If
        Next i
        If allFound Then
            WaitForExpectedEntriesTester = True
            report = "OK"
            Exit Function
        End If
        DoEvents
    Loop While TimerElapsedTester(startTime) < timeoutSeconds

    report = "Expected bundle entries were not extracted before timeout."
End Function

Private Function ManifestLooksValidTester(ByVal manifestText As String) As Boolean
    Dim addinNames As Variant
    Dim i As Long

    manifestText = Trim$(manifestText)
    If manifestText = "" Then Exit Function
    If Left$(manifestText, 1) <> "{" Or Right$(manifestText, 1) <> "}" Then Exit Function
    If InStr(1, manifestText, """BundleVersion"": ""1.0""", vbTextCompare) = 0 Then Exit Function
    If InStr(1, manifestText, """WarehouseId"": """, vbTextCompare) = 0 Then Exit Function
    If InStr(1, manifestText, """CreatedUTC"": """, vbTextCompare) = 0 Then Exit Function
    If InStr(1, manifestText, """RequiredAddins"": [", vbTextCompare) = 0 Then Exit Function

    addinNames = GetRequiredAddinsTester()
    For i = LBound(addinNames) To UBound(addinNames)
        If InStr(1, manifestText, """" & CStr(addinNames(i)) & """", vbTextCompare) = 0 Then Exit Function
    Next i

    ManifestLooksValidTester = True
End Function

Private Function GetExpectedBundleEntriesTester() As Variant
    GetExpectedBundleEntriesTester = Array( _
        TESTER_README_NAME, _
        "config/" & TESTER_CONFIG_FILE, _
        "seed/" & TESTER_SEED_FILE, _
        "auth/" & TESTER_AUTH_TEMPLATE_FILE, _
        TESTER_MANIFEST_FILE)
End Function

Private Function GetRequiredAddinsTester() As Variant
    GetRequiredAddinsTester = Array( _
        "invSys.Core.xlam", _
        "invSys.Inventory.Domain.xlam", _
        "invSys.Receiving.xlam", _
        "invSys.Admin.xlam")
End Function

Private Function CommitFileAtomicallyTester(ByVal sourcePath As String, ByVal targetPath As String, ByRef report As String) As Boolean
    Dim backupPath As String
    Dim hadExistingTarget As Boolean

    On Error GoTo FailCommit

    If Not FileExistsTester(sourcePath) Then
        report = "Source file not found: " & sourcePath
        Exit Function
    End If

    EnsureFolderRecursiveTester ParentFolderTester(targetPath)
    backupPath = targetPath & ".previous"
    hadExistingTarget = FileExistsTester(targetPath)
    DeleteFileIfPresentTester backupPath
    If hadExistingTarget Then Name targetPath As backupPath
    Name sourcePath As targetPath
    DeleteFileIfPresentTester backupPath
    CommitFileAtomicallyTester = True
    report = "OK"
    Exit Function

FailCommit:
    On Error Resume Next
    If Not FileExistsTester(targetPath) And FileExistsTester(backupPath) Then Name backupPath As targetPath
    On Error GoTo 0
    report = "CommitFileAtomicallyTester failed: " & Err.Description
End Function

Private Function OpenWorkbookByPathTester(ByVal workbookPath As String, ByRef openedTransient As Boolean) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(NormalizeFilePathTester(wb.FullName), NormalizeFilePathTester(workbookPath), vbTextCompare) = 0 Then
            Set OpenWorkbookByPathTester = wb
            Exit Function
        End If
    Next wb

    openedTransient = True
    Set OpenWorkbookByPathTester = Application.Workbooks.Open(workbookPath, False, True)
End Function

Private Function RunPowerShellCommandTester(ByVal commandBody As String, ByRef report As String) As Boolean
    Dim shellObj As Object
    Dim exitCode As Long

    On Error GoTo FailRun

    Set shellObj = CreateObject("WScript.Shell")
    exitCode = shellObj.Run("powershell -NoProfile -ExecutionPolicy Bypass -Command " & QuoteForCmdTester(commandBody), 0, True)
    If exitCode <> 0 Then
        report = "PowerShell exit code " & CStr(exitCode)
        GoTo CleanExit
    End If

    RunPowerShellCommandTester = True
    report = "OK"

CleanExit:
    Set shellObj = Nothing
    Exit Function

FailRun:
    report = "RunPowerShellCommandTester failed: " & Err.Description
    Resume CleanExit
End Function

Private Function QuoteForPowerShellTester(ByVal textIn As String) As String
    QuoteForPowerShellTester = "'" & Replace$(textIn, "'", "''") & "'"
End Function

Private Function QuoteForCmdTester(ByVal textIn As String) As String
    QuoteForCmdTester = """" & Replace$(textIn, """", """""") & """"
End Function

Private Function ResolveShellNamespaceTester(ByVal shellApp As Object, ByVal targetPath As String, ByVal timeoutSeconds As Single) As Object
    Dim startTime As Single

    startTime = Timer
    Do
        Set ResolveShellNamespaceTester = shellApp.Namespace(targetPath)
        If Not ResolveShellNamespaceTester Is Nothing Then Exit Function
        DoEvents
    Loop While TimerElapsedTester(startTime) < timeoutSeconds
End Function

Private Function GetListObjectTester(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set GetListObjectTester = wb.Worksheets(sheetName).ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function GetTableValueTester(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetTableValueTester = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value
End Function

Private Sub CloseWorkbookQuietlyTester(ByVal wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function NormalizeFolderPathTester(ByVal folderPath As String) As String
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathTester = folderPath
End Function

Private Function NormalizeFilePathTester(ByVal filePath As String) As String
    NormalizeFilePathTester = Trim$(Replace$(filePath, "/", "\"))
End Function

Private Function FolderExistsTester(ByVal folderPath As String) As Boolean
    folderPath = NormalizeFilePathTester(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) = "\" And Len(folderPath) > 3 Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    FolderExistsTester = (Len(Dir$(folderPath, vbDirectory)) > 0)
End Function

Private Function FileExistsTester(ByVal filePath As String) As Boolean
    filePath = NormalizeFilePathTester(filePath)
    If filePath = "" Then Exit Function
    FileExistsTester = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

Private Function SafeFileLenTester(ByVal filePath As String) As Long
    On Error Resume Next
    SafeFileLenTester = FileLen(filePath)
    On Error GoTo 0
End Function

Private Sub EnsureFolderRecursiveTester(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String
    Dim slashPos As Long

    folderPath = NormalizeFolderPathTester(folderPath)
    If folderPath = "" Then Exit Sub
    folderPath = Left$(folderPath, Len(folderPath) - 1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then GoTo CleanExit
    slashPos = InStrRev(folderPath, "\")
    If slashPos > 3 Then
        parentPath = Left$(folderPath, slashPos - 1)
        If Not fso.FolderExists(parentPath) Then EnsureFolderRecursiveTester parentPath
    End If
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

CleanExit:
    Set fso = Nothing
End Sub

Private Sub DeleteFolderRecursiveTester(ByVal folderPath As String)
    Dim fso As Object

    folderPath = NormalizeFolderPathTester(folderPath)
    If folderPath = "" Then Exit Sub
    folderPath = Left$(folderPath, Len(folderPath) - 1)

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FolderExists(folderPath) Then fso.DeleteFolder folderPath, True
    End If
    Set fso = Nothing
    On Error GoTo 0
End Sub

Private Sub DeleteFileIfPresentTester(ByVal filePath As String)
    On Error Resume Next
    If FileExistsTester(filePath) Then Kill filePath
    On Error GoTo 0
End Sub

Private Function BuildWorkingFolderTester(ByVal roleName As String) As String
    Randomize
    BuildWorkingFolderTester = NormalizeFolderPathTester(Environ$("TEMP")) & _
        "invSys_testerbundle\" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Rnd() * 100000), "00000") & "\" & Trim$(roleName) & "\"
End Function

Private Sub CreateEmptyZipTester(ByVal zipPath As String)
    Dim fileNum As Integer

    EnsureFolderRecursiveTester ParentFolderTester(zipPath)
    DeleteFileIfPresentTester zipPath
    fileNum = FreeFile
    Open zipPath For Binary Access Write As #fileNum
    Put #fileNum, , Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String$(18, vbNullChar)
    Close #fileNum
End Sub

Private Function ParentFolderTester(ByVal fullPath As String) As String
    Dim slashPos As Long

    fullPath = NormalizeFilePathTester(fullPath)
    slashPos = InStrRev(fullPath, "\")
    If slashPos <= 0 Then Exit Function
    ParentFolderTester = Left$(fullPath, slashPos - 1)
End Function

Private Function ReadAllTextTester(ByVal filePath As String) As String
    Dim fileNum As Integer

    If Not FileExistsTester(filePath) Then Exit Function
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    ReadAllTextTester = Space$(LOF(fileNum))
    Get #fileNum, , ReadAllTextTester
    Close #fileNum
End Function

Private Sub WriteTextFileTester(ByVal filePath As String, ByVal textOut As String)
    Dim fileNum As Integer

    EnsureFolderRecursiveTester ParentFolderTester(filePath)
    DeleteFileIfPresentTester filePath
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, textOut;
    Close #fileNum
End Sub

Private Function CsvValueTester(ByVal textIn As String) As String
    textIn = Replace$(textIn, """", """""")
    CsvValueTester = """" & textIn & """"
End Function

Private Function EscapeJsonTester(ByVal textIn As String) As String
    textIn = Replace$(textIn, "\", "\\")
    textIn = Replace$(textIn, """", Chr$(92) & """")
    textIn = Replace$(textIn, vbCrLf, "\n")
    textIn = Replace$(textIn, vbCr, "\n")
    textIn = Replace$(textIn, vbLf, "\n")
    EscapeJsonTester = textIn
End Function

Private Function TimerElapsedTester(ByVal startTime As Single) As Single
    If Timer >= startTime Then
        TimerElapsedTester = Timer - startTime
    Else
        TimerElapsedTester = (86400! - startTime) + Timer
    End If
End Function
