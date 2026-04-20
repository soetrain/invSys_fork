Attribute VB_Name = "prove_wan_wh2_setup"
Option Explicit

Private Const WAREHOUSE_ID_WAN_WH2 As String = "WH2"
Private Const PEER_WAREHOUSE_ID_WAN_WH2 As String = "WH1"
Private Const RESULT_RELATIVE_PATH_WAN_WH2 As String = "tests\integration\wan-wh2-setup-proof.md"

Private mSummary As String
Private mResultPath As String
Private mMachineName As String
Private mStationId As String
Private mSharePointRoot As String
Private mStepRows As String
Private mPassed As Long
Private mFailed As Long
Private mPeerSnapshotStampBefore As Date
Private mPeerSnapshotSizeBefore As Double

Public Function SetupVerification_WH2() As Long
    Dim runtimeRoot As String
    Dim inventoryPath As String
    Dim outboxPath As String
    Dim snapshotPath As String
    Dim publishedSnapshotPath As String
    Dim peerSnapshotPath As String
    Dim sharePointRoot As String
    Dim stationId As String
    Dim report As String
    Dim note As String
    Dim processedCount As Long
    Dim okStep As Boolean

    On Error GoTo FailRun

    ResetWanWh2SetupState

    runtimeRoot = ResolveRuntimeRootWanWh2()
    inventoryPath = runtimeRoot & "\" & WAREHOUSE_ID_WAN_WH2 & ".invSys.Data.Inventory.xlsb"
    outboxPath = runtimeRoot & "\" & WAREHOUSE_ID_WAN_WH2 & ".Outbox.Events.xlsb"
    snapshotPath = runtimeRoot & "\" & WAREHOUSE_ID_WAN_WH2 & ".invSys.Snapshot.Inventory.xlsb"

    okStep = FolderExistsWanWh2(runtimeRoot)
    If okStep Then
        note = "Runtime root exists at " & runtimeRoot & "."
    Else
        note = "Missing runtime root: " & runtimeRoot
    End If
    RecordWanWh2Step 1, okStep, note

    okStep = FileExistsWanWh2(inventoryPath) And GetFileSizeWanWh2(inventoryPath) > 0
    If okStep Then
        note = "Inventory workbook exists and is non-zero at " & inventoryPath & "."
    ElseIf FileExistsWanWh2(inventoryPath) Then
        note = "Inventory workbook exists but is zero bytes: " & inventoryPath
    Else
        note = "Missing inventory workbook: " & inventoryPath
    End If
    RecordWanWh2Step 2, okStep, note

    okStep = FileExistsWanWh2(outboxPath)
    If okStep Then
        note = "Outbox workbook exists at " & outboxPath & "."
    Else
        note = "Missing outbox workbook: " & outboxPath
    End If
    RecordWanWh2Step 3, okStep, note

    okStep = FileExistsWanWh2(snapshotPath)
    If okStep Then
        note = "Local snapshot workbook exists at " & snapshotPath & "."
    Else
        note = "Missing local snapshot workbook: " & snapshotPath
    End If
    RecordWanWh2Step 4, okStep, note

    okStep = ResolveWarehouseContextWanWh2(stationId, sharePointRoot, note)
    If okStep Then
        mStationId = stationId
        mSharePointRoot = sharePointRoot
    End If
    RecordWanWh2Step 5, okStep, note

    okStep = False
    If mSharePointRoot <> "" Then okStep = FolderExistsWanWh2(BuildFolderPathWanWh2(mSharePointRoot, "Events"))
    If okStep Then
        note = "SharePoint Events folder exists at " & BuildFolderPathWanWh2(mSharePointRoot, "Events") & "."
    ElseIf mSharePointRoot <> "" Then
        note = "Missing SharePoint Events folder: " & BuildFolderPathWanWh2(mSharePointRoot, "Events")
    Else
        note = "SharePoint root was not resolved from config."
    End If
    RecordWanWh2Step 6, okStep, note

    okStep = False
    If mSharePointRoot <> "" Then okStep = FolderExistsWanWh2(BuildFolderPathWanWh2(mSharePointRoot, "Snapshots"))
    If okStep Then
        note = "SharePoint Snapshots folder exists at " & BuildFolderPathWanWh2(mSharePointRoot, "Snapshots") & "."
    ElseIf mSharePointRoot <> "" Then
        note = "Missing SharePoint Snapshots folder: " & BuildFolderPathWanWh2(mSharePointRoot, "Snapshots")
    Else
        note = "SharePoint root was not resolved from config."
    End If
    RecordWanWh2Step 7, okStep, note

    If mSharePointRoot <> "" Then
        peerSnapshotPath = BuildFolderPathWanWh2(BuildFolderPathWanWh2(mSharePointRoot, "Snapshots"), PEER_WAREHOUSE_ID_WAN_WH2 & ".invSys.Snapshot.Inventory.xlsb")
        CapturePeerSnapshotBaselineWanWh2 peerSnapshotPath
    Else
        peerSnapshotPath = vbNullString
    End If

    note = vbNullString
    okStep = False
    If mStationId = "" Then
        note = "StationId could not be resolved from config; RunBatch was not attempted."
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
        If Not modConfig.LoadConfig(WAREHOUSE_ID_WAN_WH2, mStationId) Then
            note = "Config load failed: " & modConfig.Validate()
        Else
            processedCount = modProcessor.RunBatch(WAREHOUSE_ID_WAN_WH2, 0, report)
            okStep = RunBatchResultIsHealthyWanWh2(report)
            If okStep Then
                note = "RunBatch completed without fatal errors. Processed=" & CStr(processedCount) & "; Report=" & report
            Else
                note = "RunBatch reported a fatal or degraded result. Processed=" & CStr(processedCount) & "; Report=" & report
            End If
        End If
    End If
    RecordWanWh2Step 8, okStep, note

    publishedSnapshotPath = BuildFolderPathWanWh2(BuildFolderPathWanWh2(mSharePointRoot, "Snapshots"), WAREHOUSE_ID_WAN_WH2 & ".invSys.Snapshot.Inventory.xlsb")
    okStep = False
    If mSharePointRoot = "" Then
        note = "Published snapshot check blocked because PathSharePointRoot was not resolved from config."
    Else
        okStep = FileExistsWanWh2(publishedSnapshotPath)
        If okStep Then
            note = "Published snapshot exists at " & publishedSnapshotPath & "."
        Else
            note = "Missing published snapshot: " & publishedSnapshotPath
        End If
    End If
    RecordWanWh2Step 9, okStep, note

    okStep = False
    If mSharePointRoot = "" Then
        note = "Publish temp-file check blocked because PathSharePointRoot was not resolved from config."
    Else
        okStep = Not FileExistsWanWh2(publishedSnapshotPath & ".uploading")
        If okStep Then
            note = "No publish temp file remains at " & publishedSnapshotPath & ".uploading."
        Else
            note = "Publish temp file still present: " & publishedSnapshotPath & ".uploading"
        End If
    End If
    RecordWanWh2Step 10, okStep, note

    okStep = False
    If mSharePointRoot = "" Then
        note = "Peer-snapshot cross-contamination check blocked because PathSharePointRoot was not resolved from config."
    Else
        okStep = ValidatePeerSnapshotUnmodifiedWanWh2(peerSnapshotPath, note)
    End If
    RecordWanWh2Step 11, okStep, note

    If mFailed = 0 And mPassed = 11 Then
        mSummary = "WH2 WAN setup proof passed all 11 real-machine steps."
        SetupVerification_WH2 = 1
    Else
        mSummary = "WH2 WAN setup proof did not pass all required steps."
    End If

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    Exit Function

FailRun:
    RecordWanWh2Step 0, False, "Harness exception: " & Err.Description
    mSummary = "WH2 WAN setup proof raised an unexpected exception."
    Resume CleanExit
End Function

Public Function GetWanWh2SetupContextPacked() As String
    GetWanWh2SetupContextPacked = _
        "Summary=" & SafeTextWanWh2(mSummary) & _
        "|Machine=" & SafeTextWanWh2(mMachineName) & _
        "|Warehouse=" & WAREHOUSE_ID_WAN_WH2 & _
        "|Station=" & SafeTextWanWh2(mStationId) & _
        "|SharePointRoot=" & SafeTextWanWh2(mSharePointRoot) & _
        "|ResultPath=" & SafeTextWanWh2(mResultPath) & _
        "|Passed=" & CStr(mPassed) & _
        "|Failed=" & CStr(mFailed)
End Function

Public Function GetWanWh2SetupEvidenceRows() As String
    GetWanWh2SetupEvidenceRows = mStepRows
End Function

Public Sub WriteProofResult(ByVal machineName As String, ByVal step As Integer, ByVal passed As Boolean, ByVal note As String)
    Dim ts As String
    Dim lineOut As String

    If mResultPath = "" Then mResultPath = ResolveResultPathWanWh2()
    If mResultPath = "" Then Exit Sub

    EnsureResultFileHeaderWanWh2 mResultPath

    ts = CurrentUtcStampWanWh2()
    lineOut = "| " & SanitizeMarkdownWanWh2(machineName) & _
              " | " & CStr(step) & _
              " | " & IIf(passed, "PASS", "FAIL") & _
              " | " & SanitizeMarkdownWanWh2(note) & _
              " | " & SanitizeMarkdownWanWh2(ts) & " |"

    AppendLineWanWh2 mResultPath, lineOut
End Sub

Private Sub ResetWanWh2SetupState()
    mSummary = vbNullString
    mResultPath = ResolveResultPathWanWh2()
    mMachineName = ResolveMachineNameWanWh2()
    mStationId = vbNullString
    mSharePointRoot = vbNullString
    mStepRows = vbNullString
    mPassed = 0
    mFailed = 0
    mPeerSnapshotStampBefore = 0
    mPeerSnapshotSizeBefore = 0
End Sub

Private Sub RecordWanWh2Step(ByVal stepNo As Long, ByVal passed As Boolean, ByVal note As String)
    If passed Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
    End If

    If Len(mStepRows) > 0 Then mStepRows = mStepRows & vbLf
    mStepRows = mStepRows & "Step " & CStr(stepNo) & vbTab & IIf(passed, "PASS", "FAIL") & vbTab & note

    WriteProofResult mMachineName, stepNo, passed, note
End Sub

Private Function ResolveWarehouseContextWanWh2(ByRef stationId As String, ByRef sharePointRoot As String, ByRef note As String) As Boolean
    Dim configPath As String
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim loSt As ListObject
    Dim openedHere As Boolean

    On Error GoTo FailResolve

    configPath = ResolveRuntimeRootWanWh2() & "\" & WAREHOUSE_ID_WAN_WH2 & ".invSys.Config.xlsb"
    If Not FileExistsWanWh2(configPath) Then
        note = "Config workbook missing: " & configPath
        GoTo CleanExit
    End If

    Set wbCfg = FindOpenWorkbookByPathWanWh2(configPath)
    If wbCfg Is Nothing Then
        Set wbCfg = Application.Workbooks.Open(Filename:=configPath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedHere = Not wbCfg Is Nothing
    End If
    If wbCfg Is Nothing Then
        note = "Config workbook could not be opened: " & configPath
        GoTo CleanExit
    End If

    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    If loWh Is Nothing Or loSt Is Nothing Then
        note = "Config tables were missing from " & configPath
        GoTo CleanExit
    End If
    If loWh.DataBodyRange Is Nothing Or loSt.DataBodyRange Is Nothing Then
        note = "Config tables did not contain any data rows."
        GoTo CleanExit
    End If

    sharePointRoot = Trim$(CStr(loWh.DataBodyRange.Cells(1, loWh.ListColumns("PathSharePointRoot").Index).Value))
    stationId = ResolveFirstStationIdWanWh2(loSt)
    If sharePointRoot = "" Then
        note = "PathSharePointRoot was blank in " & configPath
        GoTo CleanExit
    End If
    If stationId = "" Then
        note = "No StationId row was present in tblStationConfig."
        GoTo CleanExit
    End If

    sharePointRoot = NormalizeFolderPathWanWh2(sharePointRoot)
    If Not FolderExistsWanWh2(sharePointRoot) Then
        note = "PathSharePointRoot was set but unreachable: " & sharePointRoot
        GoTo CleanExit
    End If

    note = "PathSharePointRoot=" & sharePointRoot & "; StationId=" & stationId & "; SharePoint root is reachable."
    ResolveWarehouseContextWanWh2 = True

CleanExit:
    If openedHere And Not wbCfg Is Nothing Then
        On Error Resume Next
        wbCfg.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Exit Function

FailResolve:
    note = "ResolveWarehouseContext failed: " & Err.Description
    Resume CleanExit
End Function

Private Function ResolveRuntimeRootWanWh2() As String
    ResolveRuntimeRootWanWh2 = Trim$(modRuntimeWorkbooks.TryResolveExistingRuntimeRoot(WAREHOUSE_ID_WAN_WH2))
    If ResolveRuntimeRootWanWh2 = "" Then
        ResolveRuntimeRootWanWh2 = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(WAREHOUSE_ID_WAN_WH2, False)
    End If
End Function

Private Function ResolveFirstStationIdWanWh2(ByVal loSt As ListObject) As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim valueText As String

    If loSt Is Nothing Then Exit Function
    If loSt.DataBodyRange Is Nothing Then Exit Function

    colIndex = loSt.ListColumns("StationId").Index
    For rowIndex = 1 To loSt.ListRows.Count
        valueText = Trim$(CStr(loSt.DataBodyRange.Cells(rowIndex, colIndex).Value))
        If valueText <> "" Then
            ResolveFirstStationIdWanWh2 = valueText
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub CapturePeerSnapshotBaselineWanWh2(ByVal peerSnapshotPath As String)
    If FileExistsWanWh2(peerSnapshotPath) Then
        mPeerSnapshotStampBefore = FileDateTime(peerSnapshotPath)
        mPeerSnapshotSizeBefore = GetFileSizeWanWh2(peerSnapshotPath)
    Else
        mPeerSnapshotStampBefore = 0
        mPeerSnapshotSizeBefore = 0
    End If
End Sub

Private Function ValidatePeerSnapshotUnmodifiedWanWh2(ByVal peerSnapshotPath As String, ByRef note As String) As Boolean
    Dim currentStamp As Date
    Dim currentSize As Double

    If Not FileExistsWanWh2(peerSnapshotPath) Then
        note = "Peer WH1 published snapshot missing after WH2 publish: " & peerSnapshotPath
        Exit Function
    End If

    currentStamp = FileDateTime(peerSnapshotPath)
    currentSize = GetFileSizeWanWh2(peerSnapshotPath)

    If mPeerSnapshotStampBefore = 0 Then
        note = "Peer WH1 published snapshot exists after WH2 publish at " & peerSnapshotPath & ". No baseline timestamp was available before the run."
        ValidatePeerSnapshotUnmodifiedWanWh2 = True
        Exit Function
    End If

    If currentStamp = mPeerSnapshotStampBefore And currentSize = mPeerSnapshotSizeBefore Then
        note = "Peer WH1 published snapshot remained present and unmodified at " & peerSnapshotPath & "."
        ValidatePeerSnapshotUnmodifiedWanWh2 = True
    Else
        note = "Peer WH1 published snapshot changed during WH2 proof. BeforeStamp=" & Format$(mPeerSnapshotStampBefore, "yyyy-mm-dd hh:nn:ss") & _
               "; AfterStamp=" & Format$(currentStamp, "yyyy-mm-dd hh:nn:ss") & _
               "; BeforeSize=" & CStr(mPeerSnapshotSizeBefore) & _
               "; AfterSize=" & CStr(currentSize)
    End If
End Function

Private Function RunBatchResultIsHealthyWanWh2(ByVal report As String) As Boolean
    Dim upperReport As String

    upperReport = UCase$(Trim$(report))
    If Left$(upperReport, 15) = "RUNBATCH FAILED" Then Exit Function
    If InStr(1, upperReport, "SNAPSHOTERROR=", vbTextCompare) > 0 Then Exit Function
    If InStr(1, upperReport, "PUBLISHWARNING=", vbTextCompare) > 0 Then Exit Function
    If InStr(1, upperReport, "INVENTORY WORKBOOK IS READ-ONLY OR LOCKED", vbTextCompare) > 0 Then Exit Function
    If InStr(1, upperReport, "NOT FOUND", vbTextCompare) > 0 Then Exit Function
    RunBatchResultIsHealthyWanWh2 = True
End Function

Private Function ResolveResultPathWanWh2() As String
    Dim repoRoot As String

    repoRoot = FindRepoRootWanWh2(ThisWorkbook.Path)
    If repoRoot = "" Then repoRoot = FindRepoRootWanWh2(CurDir$)
    If repoRoot = "" Then Exit Function

    ResolveResultPathWanWh2 = repoRoot & "\" & RESULT_RELATIVE_PATH_WAN_WH2
End Function

Private Function FindRepoRootWanWh2(ByVal startPath As String) As String
    Dim fso As Object
    Dim probe As String
    Dim parentPath As String

    On Error GoTo FailFind

    Set fso = CreateObject("Scripting.FileSystemObject")
    probe = Trim$(startPath)
    If probe = "" Then Exit Function

    If Not fso.FolderExists(probe) Then
        If fso.FileExists(probe) Then
            probe = fso.GetParentFolderName(probe)
        Else
            Exit Function
        End If
    End If

    Do While probe <> ""
        If fso.FolderExists(probe & "\.git") And fso.FolderExists(probe & "\tests") Then
            FindRepoRootWanWh2 = probe
            Exit Function
        End If
        parentPath = fso.GetParentFolderName(probe)
        If parentPath = probe Then Exit Do
        probe = parentPath
    Loop
    Exit Function

FailFind:
    FindRepoRootWanWh2 = vbNullString
End Function

Private Sub EnsureResultFileHeaderWanWh2(ByVal resultPath As String)
    Dim folderPath As String
    Dim lines(0 To 6) As String

    If resultPath = "" Then Exit Sub
    If FileExistsWanWh2(resultPath) Then Exit Sub

    folderPath = GetParentFolderWanWh2(resultPath)
    If folderPath <> "" Then EnsureFolderRecursiveWanWh2 folderPath

    lines(0) = "# WAN WH2 Setup Proof"
    lines(1) = ""
    lines(2) = "- Warehouse: `WH2`"
    lines(3) = "- Scope note: real-machine setup, publish proof, and WH1 cross-contamination check for the WAN proving path Slice B."
    lines(4) = ""
    lines(5) = "| Machine | Step | Result | Note | UTC timestamp |"
    lines(6) = "|---|---|---|---|---|"
    WriteAllLinesWanWh2 resultPath, lines
End Sub

Private Sub AppendLineWanWh2(ByVal targetPath As String, ByVal lineOut As String)
    Dim fso As Object
    Dim ts As Object

    On Error GoTo FailAppend

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(targetPath, 8, True, 0)
    ts.WriteLine lineOut

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    On Error GoTo 0
    Exit Sub

FailAppend:
    Resume CleanExit
End Sub

Private Sub WriteAllLinesWanWh2(ByVal targetPath As String, ByRef lines() As String)
    Dim fso As Object
    Dim ts As Object
    Dim i As Long

    On Error GoTo CleanExit

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(targetPath, True, False)
    For i = LBound(lines) To UBound(lines)
        ts.WriteLine lines(i)
    Next i

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    On Error GoTo 0
End Sub

Private Function ResolveMachineNameWanWh2() As String
    On Error Resume Next
    ResolveMachineNameWanWh2 = Trim$(Environ$("COMPUTERNAME"))
    If ResolveMachineNameWanWh2 = "" Then ResolveMachineNameWanWh2 = Trim$(CreateObject("WScript.Network").ComputerName)
    If ResolveMachineNameWanWh2 = "" Then ResolveMachineNameWanWh2 = "UNKNOWN-MACHINE"
    On Error GoTo 0
End Function

Private Function CurrentUtcStampWanWh2() As String
    Dim shellObj As Object
    Dim execObj As Object
    Dim outputText As String

    On Error GoTo FallbackStamp

    Set shellObj = CreateObject("WScript.Shell")
    Set execObj = shellObj.Exec("powershell -NoProfile -Command ""[DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss')""")
    outputText = Trim$(execObj.StdOut.ReadAll)
    If outputText <> "" Then
        CurrentUtcStampWanWh2 = outputText
        Exit Function
    End If

FallbackStamp:
    CurrentUtcStampWanWh2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Private Function NormalizeFolderPathWanWh2(ByVal folderPath As String) As String
    NormalizeFolderPathWanWh2 = Trim$(Replace$(folderPath, "/", "\"))
    If Right$(NormalizeFolderPathWanWh2, 1) = "\" Then
        NormalizeFolderPathWanWh2 = Left$(NormalizeFolderPathWanWh2, Len(NormalizeFolderPathWanWh2) - 1)
    End If
End Function

Private Function BuildFolderPathWanWh2(ByVal rootPath As String, ByVal childName As String) As String
    rootPath = NormalizeFolderPathWanWh2(rootPath)
    childName = Trim$(Replace$(childName, "/", "\"))
    If rootPath = "" Then
        BuildFolderPathWanWh2 = childName
    ElseIf childName = "" Then
        BuildFolderPathWanWh2 = rootPath
    Else
        BuildFolderPathWanWh2 = rootPath & "\" & childName
    End If
End Function

Private Function FileExistsWanWh2(ByVal fullPath As String) As Boolean
    Dim fso As Object

    fullPath = Trim$(fullPath)
    If fullPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsWanWh2 = fso.FileExists(fullPath)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    FileExistsWanWh2 = (Len(Dir$(fullPath, vbNormal)) > 0)
    On Error GoTo 0
End Function

Private Function FolderExistsWanWh2(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsWanWh2 = fso.FolderExists(folderPath)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    FolderExistsWanWh2 = (Len(Dir$(folderPath, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function GetFileSizeWanWh2(ByVal fullPath As String) As Double
    On Error Resume Next
    If FileExistsWanWh2(fullPath) Then GetFileSizeWanWh2 = FileLen(fullPath)
    On Error GoTo 0
End Function

Private Function FindOpenWorkbookByPathWanWh2(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    fullPath = LCase$(Trim$(fullPath))
    If fullPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If LCase$(Trim$(wb.FullName)) = fullPath Then
            Set FindOpenWorkbookByPathWanWh2 = wb
            Exit Function
        End If
    Next wb
End Function

Private Function GetParentFolderWanWh2(ByVal fullPath As String) As String
    On Error Resume Next
    GetParentFolderWanWh2 = CreateObject("Scripting.FileSystemObject").GetParentFolderName(fullPath)
    On Error GoTo 0
End Function

Private Sub EnsureFolderRecursiveWanWh2(ByVal folderPath As String)
    Dim parentPath As String
    Dim fso As Object

    folderPath = NormalizeFolderPathWanWh2(folderPath)
    If folderPath = "" Then Exit Sub
    If FolderExistsWanWh2(folderPath) Then Exit Sub

    parentPath = GetParentFolderWanWh2(folderPath)
    If parentPath <> "" And Not FolderExistsWanWh2(parentPath) Then EnsureFolderRecursiveWanWh2 parentPath

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.CreateFolder folderPath
    On Error GoTo 0
End Sub

Private Function SanitizeMarkdownWanWh2(ByVal textIn As String) As String
    SanitizeMarkdownWanWh2 = Trim$(Replace$(Replace$(CStr(textIn), "|", " ; "), vbCr, " "))
    SanitizeMarkdownWanWh2 = Replace$(SanitizeMarkdownWanWh2, vbLf, " ")
End Function

Private Function SafeTextWanWh2(ByVal textIn As String) As String
    SafeTextWanWh2 = Replace$(Replace$(CStr(textIn), "|", "/"), vbCr, " ")
    SafeTextWanWh2 = Replace$(SafeTextWanWh2, vbLf, " ")
End Function
