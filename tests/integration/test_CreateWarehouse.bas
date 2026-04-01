Attribute VB_Name = "test_CreateWarehouse"
Option Explicit

Private mCheckNames() As String
Private mCheckResults() As String
Private mCheckDetails() As String
Private mCheckCount As Long

Private mWarehouseId As String
Private mStationId As String
Private mLocalRoot As String
Private mSharePointRoot As String
Private mSummary As String

Public Function TestCreateWarehouse_EndToEndLifecycle() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim duplicateSpec As modWarehouseBootstrap.WarehouseSpec
    Dim templateRoot As String
    Dim duplicateRoot As String
    Dim validSpec As Boolean
    Dim existsBefore As Boolean
    Dim bootstrapOk As Boolean
    Dim publishOk As Boolean
    Dim duplicateExists As Boolean
    Dim duplicateRejected As Boolean
    Dim detail As String
    Dim duplicateReport As String

    On Error GoTo FailTest

    ResetCreateWarehouseEvidence

    mWarehouseId = "WHBOOT-E2E_01"
    mStationId = "ADM1"
    mLocalRoot = BuildCreateWarehouseTempRoot("local")
    mSharePointRoot = BuildCreateWarehouseTempRoot("share")
    templateRoot = BuildCreateWarehouseTempRoot("templates")
    duplicateRoot = BuildCreateWarehouseTempRoot("duplicate")

    spec.WarehouseId = mWarehouseId
    spec.WarehouseName = "Create Warehouse Integration"
    spec.StationId = mStationId
    spec.AdminUser = "admin.integration"
    spec.PathLocal = mLocalRoot
    spec.PathSharePoint = mSharePointRoot

    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot

    validSpec = modWarehouseBootstrap.ValidateWarehouseSpec(spec, detail)
    RecordCreateWarehouseCheck "WarehouseSpec.Valid", validSpec, detail
    If Not validSpec Then GoTo CleanExit

    existsBefore = modWarehouseBootstrap.WarehouseIdExists(spec.WarehouseId)
    RecordCreateWarehouseCheck "CollisionCheck.InitialClear", Not existsBefore, _
        "WarehouseIdExists=" & CStr(existsBefore)
    If existsBefore Then GoTo CleanExit

    bootstrapOk = modWarehouseBootstrap.BootstrapWarehouseLocal(spec)
    RecordCreateWarehouseCheck "Bootstrap.Local", bootstrapOk, modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
    If Not bootstrapOk Then GoTo CleanExit

    RecordCreateWarehouseCheck "LocalStructure.Exists", _
        AssertLocalStructureCreateWarehouse(spec, detail), detail

    RecordCreateWarehouseCheck "ConfigSeeded.Correctly", _
        AssertConfigSeededCreateWarehouse(spec, detail), detail

    publishOk = modWarehouseBootstrap.PublishInitialArtifacts(spec)
    RecordCreateWarehouseCheck "SharePointPublish.Initial", publishOk, modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
    If Not publishOk Then GoTo CleanExit

    RecordCreateWarehouseCheck "SharePointArtifacts.Exists", _
        AssertSharePointArtifactsCreateWarehouse(spec, detail), detail

    duplicateExists = modWarehouseBootstrap.WarehouseIdExists(spec.WarehouseId)
    RecordCreateWarehouseCheck "CollisionCheck.DuplicateVisible", duplicateExists, _
        "WarehouseIdExists=" & CStr(duplicateExists)

    duplicateSpec = spec
    duplicateSpec.PathLocal = duplicateRoot
    duplicateRejected = Not modWarehouseBootstrap.BootstrapWarehouseLocal(duplicateSpec)
    duplicateReport = modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
    RecordCreateWarehouseCheck "DuplicateRun.Rejected", _
        duplicateRejected And InStr(1, duplicateReport, "already exists", vbTextCompare) > 0, _
        duplicateReport

    If AllCreateWarehouseChecksPassed() Then
        mSummary = "Create warehouse lifecycle completed, SharePoint artifacts were published, and duplicate rejection was proven."
        TestCreateWarehouse_EndToEndLifecycle = 1
    Else
        mSummary = "One or more create warehouse lifecycle checks failed."
    End If

CleanExit:
    On Error Resume Next
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteCreateWarehouseFolderRecursive duplicateRoot
    DeleteCreateWarehouseFolderRecursive mSharePointRoot
    DeleteCreateWarehouseFolderRecursive mLocalRoot
    DeleteCreateWarehouseFolderRecursive templateRoot
    On Error GoTo 0
    If mSummary = "" Then mSummary = "Create warehouse lifecycle did not complete."
    Exit Function

FailTest:
    RecordCreateWarehouseCheck "TestHarness.Exception", False, Err.Description
    mSummary = "Create warehouse lifecycle raised an unexpected exception."
    Resume CleanExit
End Function

Public Function GetCreateWarehouseContextPacked() As String
    GetCreateWarehouseContextPacked = _
        "WarehouseId=" & SafeCreateWarehouseText(mWarehouseId) & "|" & _
        "StationId=" & SafeCreateWarehouseText(mStationId) & "|" & _
        "LocalRoot=" & SafeCreateWarehouseText(mLocalRoot) & "|" & _
        "SharePointRoot=" & SafeCreateWarehouseText(mSharePointRoot) & "|" & _
        "Summary=" & SafeCreateWarehouseText(mSummary)
End Function

Public Function GetCreateWarehouseEvidenceRows() As String
    Dim i As Long

    For i = 1 To mCheckCount
        If Len(GetCreateWarehouseEvidenceRows) > 0 Then GetCreateWarehouseEvidenceRows = GetCreateWarehouseEvidenceRows & vbLf
        GetCreateWarehouseEvidenceRows = GetCreateWarehouseEvidenceRows & _
            mCheckNames(i) & vbTab & mCheckResults(i) & vbTab & mCheckDetails(i)
    Next i
End Function

Private Sub ResetCreateWarehouseEvidence()
    mCheckCount = 0
    Erase mCheckNames
    Erase mCheckResults
    Erase mCheckDetails
    mWarehouseId = vbNullString
    mStationId = vbNullString
    mLocalRoot = vbNullString
    mSharePointRoot = vbNullString
    mSummary = vbNullString
End Sub

Private Sub RecordCreateWarehouseCheck(ByVal checkName As String, _
                                       ByVal passed As Boolean, _
                                       ByVal detailText As String)
    mCheckCount = mCheckCount + 1
    ReDim Preserve mCheckNames(1 To mCheckCount)
    ReDim Preserve mCheckResults(1 To mCheckCount)
    ReDim Preserve mCheckDetails(1 To mCheckCount)

    mCheckNames(mCheckCount) = Trim$(checkName)
    mCheckResults(mCheckCount) = IIf(passed, "PASS", "FAIL")
    mCheckDetails(mCheckCount) = SafeCreateWarehouseText(detailText)
End Sub

Private Function AllCreateWarehouseChecksPassed() As Boolean
    Dim i As Long

    AllCreateWarehouseChecksPassed = (mCheckCount > 0)
    For i = 1 To mCheckCount
        If StrComp(mCheckResults(i), "PASS", vbTextCompare) <> 0 Then
            AllCreateWarehouseChecksPassed = False
            Exit Function
        End If
    Next i
End Function

Private Function AssertLocalStructureCreateWarehouse(ByRef spec As modWarehouseBootstrap.WarehouseSpec, _
                                                     ByRef detailText As String) As Boolean
    Dim requiredPaths As Variant
    Dim item As Variant

    requiredPaths = Array( _
        spec.PathLocal, _
        spec.PathLocal & "\inbox", _
        spec.PathLocal & "\outbox", _
        spec.PathLocal & "\snapshots", _
        spec.PathLocal & "\config", _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb", _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Config.xlsb", _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Auth.xlsb", _
        spec.PathLocal & "\" & spec.WarehouseId & ".Outbox.Events.xlsb", _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Snapshot.Inventory.xlsb")

    For Each item In requiredPaths
        If Not CreateWarehousePathExists(CStr(item)) Then
            detailText = "Missing path: " & CStr(item)
            Exit Function
        End If
    Next item

    detailText = "All required runtime folders and seeded artifacts were created under " & spec.PathLocal
    AssertLocalStructureCreateWarehouse = True
End Function

Private Function AssertConfigSeededCreateWarehouse(ByRef spec As modWarehouseBootstrap.WarehouseSpec, _
                                                   ByRef detailText As String) As Boolean
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim loSt As ListObject

    On Error GoTo FailAssert

    Set wbCfg = Application.Workbooks.Open(spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Config.xlsb")
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")

    If StrComp(CStr(GetCreateWarehouseTableValue(loWh, 1, "WarehouseId")), spec.WarehouseId, vbTextCompare) <> 0 Then
        detailText = "WarehouseId was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loWh, 1, "WarehouseName")), spec.WarehouseName, vbTextCompare) <> 0 Then
        detailText = "WarehouseName was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loWh, 1, "PathDataRoot")), spec.PathLocal, vbTextCompare) <> 0 Then
        detailText = "PathDataRoot was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loWh, 1, "PathSharePointRoot")), spec.PathSharePoint, vbTextCompare) <> 0 Then
        detailText = "PathSharePointRoot was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loSt, 1, "StationId")), spec.StationId, vbTextCompare) <> 0 Then
        detailText = "StationId row was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loSt, 1, "StationName")), spec.AdminUser, vbTextCompare) <> 0 Then
        detailText = "Admin user was not seeded into StationName."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loSt, 1, "RoleDefault")), "ADMIN", vbTextCompare) <> 0 Then
        detailText = "RoleDefault was not seeded as ADMIN."
        GoTo CleanExit
    End If

    detailText = "Config workbook seeded WarehouseId, WarehouseName, StationId, PathDataRoot, PathSharePointRoot, and ADMIN defaults."
    AssertConfigSeededCreateWarehouse = True

CleanExit:
    CloseCreateWarehouseWorkbook wbCfg
    Exit Function

FailAssert:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function AssertSharePointArtifactsCreateWarehouse(ByRef spec As modWarehouseBootstrap.WarehouseSpec, _
                                                          ByRef detailText As String) As Boolean
    Dim discoveryPath As String
    Dim publishedConfigPath As String

    discoveryPath = spec.PathSharePoint & "\" & spec.WarehouseId & ".config.json"
    publishedConfigPath = spec.PathSharePoint & "\" & spec.WarehouseId & "\" & spec.WarehouseId & ".invSys.Config.xlsb"

    If Not CreateWarehousePathExists(discoveryPath) Then
        detailText = "Discovery artifact missing: " & discoveryPath
        Exit Function
    End If
    If Not CreateWarehousePathExists(publishedConfigPath) Then
        detailText = "Published config artifact missing: " & publishedConfigPath
        Exit Function
    End If

    detailText = "Discovery artifact and published config workbook exist under " & spec.PathSharePoint
    AssertSharePointArtifactsCreateWarehouse = True
End Function

Private Function GetCreateWarehouseTableValue(ByVal lo As ListObject, _
                                              ByVal rowIndex As Long, _
                                              ByVal columnName As String) As Variant
    Dim idx As Long

    idx = lo.ListColumns(columnName).Index
    GetCreateWarehouseTableValue = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Function CreateWarehousePathExists(ByVal pathIn As String) As Boolean
    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function

    CreateWarehousePathExists = (Len(Dir$(pathIn, vbDirectory)) > 0)
    If Not CreateWarehousePathExists Then
        CreateWarehousePathExists = (Len(Dir$(pathIn, vbNormal)) > 0)
    End If
End Function

Private Function BuildCreateWarehouseTempRoot(ByVal leafName As String) As String
    BuildCreateWarehouseTempRoot = Environ$("TEMP") & "\invSys_createwarehouse_" & leafName & "_" & _
                                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 1000))
End Function

Private Sub DeleteCreateWarehouseFolderRecursive(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

Private Sub CloseCreateWarehouseWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function SafeCreateWarehouseText(ByVal textIn As String) As String
    SafeCreateWarehouseText = Replace$(Replace$(Trim$(textIn), vbCr, " "), vbLf, " ")
End Function
