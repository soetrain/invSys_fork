Attribute VB_Name = "modInventoryPublisher"
Option Explicit

Private Const TABLE_INVENTORY_LOG_PUBLISHER As String = "tblInventoryLog"
Private Const TABLE_APPLIED_EVENTS_PUBLISHER As String = "tblAppliedEvents"
Private Const TABLE_SKU_BALANCE_PUBLISHER As String = "tblSkuBalance"
Private Const TABLE_LOCATION_BALANCE_PUBLISHER As String = "tblLocationBalance"
Private Const TABLE_LEDGER_STATUS_PUBLISHER As String = "tblInventoryLedgerStatus"
Private Const MIN_RECENT_PUBLISH_SECONDS As Long = 5

Private mRecentPublishes As Object

Public Function PublishOpenInventorySnapshots(Optional ByRef report As String = "") As Long
    On Error GoTo FailPublish

    Dim wb As Workbook
    Dim publishReport As String
    Dim detail As String

    For Each wb In Application.Workbooks
        publishReport = vbNullString
        If IsInventorySourceWorkbookPublisher(wb) Then
            If EnsureSnapshotPublicationForWorkbook(wb, publishReport) Then
                If StrComp(Trim$(publishReport), "SKIPPED_RECENT", vbTextCompare) <> 0 Then
                    PublishOpenInventorySnapshots = PublishOpenInventorySnapshots + 1
                End If
            End If
            If publishReport <> "" Then
                If detail <> "" Then detail = detail & "; "
                detail = detail & wb.Name & "=" & publishReport
            End If
        End If
    Next wb

    report = detail
    Exit Function

FailPublish:
    report = "PublishOpenInventorySnapshots failed: " & Err.Description
End Function

Public Function EnsureSnapshotPublicationForWorkbook(Optional ByVal targetWb As Workbook = Nothing, _
                                                     Optional ByRef report As String = "") As Boolean
    On Error GoTo FailPublish

    Dim wb As Workbook
    Dim resolvedWarehouseId As String
    Dim snapshotPath As String
    Dim publishKey As String

    Set wb = targetWb
    If wb Is Nothing Then Set wb = ResolveCandidateInventoryWorkbookPublisher()
    If wb Is Nothing Then
        report = "Inventory source workbook not resolved."
        Exit Function
    End If
    If Not IsInventorySourceWorkbookPublisher(wb) Then
        report = "Workbook is not an inventory source workbook."
        Exit Function
    End If
    If Not TryResolveInventorySourceWarehouseIdPublisher(wb, resolvedWarehouseId) Then
        report = "Inventory source warehouse could not be resolved."
        Exit Function
    End If
    If Not EnsureWarehouseConfigLoadedPublisher(resolvedWarehouseId, report) Then Exit Function

    publishKey = BuildPublishKeyPublisher(wb, resolvedWarehouseId)
    If ShouldSkipRecentPublishPublisher(publishKey) Then
        report = "SKIPPED_RECENT"
        EnsureSnapshotPublicationForWorkbook = True
        Exit Function
    End If

    snapshotPath = vbNullString
    If Not modWarehouseSync.GenerateWarehouseSnapshot(resolvedWarehouseId, wb, "", Nothing, snapshotPath) Then
        report = snapshotPath
        Exit Function
    End If

    RecordRecentPublishPublisher publishKey
    report = snapshotPath
    EnsureSnapshotPublicationForWorkbook = True
    Exit Function

FailPublish:
    report = "EnsureSnapshotPublicationForWorkbook failed: " & Err.Description
End Function

Public Sub HandlePotentialInventoryWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    Dim report As String

    Call EnsureSnapshotPublicationForWorkbook(targetWb, report)
End Sub

Private Function ResolveCandidateInventoryWorkbookPublisher() As Workbook
    If Not Application.ActiveWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook.IsAddin Then
            If IsInventorySourceWorkbookPublisher(Application.ActiveWorkbook) Then
                Set ResolveCandidateInventoryWorkbookPublisher = Application.ActiveWorkbook
            End If
        End If
    End If
End Function

Private Function IsInventorySourceWorkbookPublisher(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function

    IsInventorySourceWorkbookPublisher = WorkbookHasTablePublisher(wb, TABLE_INVENTORY_LOG_PUBLISHER) _
        And WorkbookHasTablePublisher(wb, TABLE_APPLIED_EVENTS_PUBLISHER) _
        And WorkbookHasTablePublisher(wb, TABLE_SKU_BALANCE_PUBLISHER) _
        And WorkbookHasTablePublisher(wb, TABLE_LOCATION_BALANCE_PUBLISHER)
End Function

Private Function WorkbookHasTablePublisher(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasTablePublisher = Not (FindListObjectByNamePublisher(wb, tableName) Is Nothing)
End Function

Private Function FindListObjectByNamePublisher(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindListObjectByNamePublisher = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindListObjectByNamePublisher Is Nothing Then Exit Function
    Next ws
End Function

Private Function TryResolveInventorySourceWarehouseIdPublisher(ByVal wb As Workbook, ByRef warehouseId As String) As Boolean
    warehouseId = ResolveWarehouseIdFromLedgerStatusPublisher(wb)
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    warehouseId = ResolveWarehouseIdFromInventoryWorkbookNamePublisher(wb.Name)
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    warehouseId = ResolveWarehouseIdFromSiblingConfigPublisher(wb)
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    warehouseId = ResolveWarehouseIdFromOpenConfigPublisher()
    If warehouseId <> "" Then
        TryResolveInventorySourceWarehouseIdPublisher = True
        Exit Function
    End If

    If modConfig.IsLoaded() Then
        warehouseId = Trim$(modConfig.GetWarehouseId())
        TryResolveInventorySourceWarehouseIdPublisher = (warehouseId <> "")
    End If
End Function

Private Function ResolveWarehouseIdFromLedgerStatusPublisher(ByVal wb As Workbook) As String
    Dim lo As ListObject
    Dim idx As Long

    Set lo = FindListObjectByNamePublisher(wb, TABLE_LEDGER_STATUS_PUBLISHER)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idx = lo.ListColumns("WarehouseId").Index
    ResolveWarehouseIdFromLedgerStatusPublisher = Trim$(CStr(lo.DataBodyRange.Cells(1, idx).Value))
End Function

Private Function ResolveWarehouseIdFromInventoryWorkbookNamePublisher(ByVal wbName As String) As String
    Dim markerPos As Long

    markerPos = InStr(1, wbName, ".invSys.Data.Inventory.", vbTextCompare)
    If markerPos > 1 Then ResolveWarehouseIdFromInventoryWorkbookNamePublisher = Left$(wbName, markerPos - 1)
End Function

Private Function ResolveWarehouseIdFromSiblingConfigPublisher(ByVal wb As Workbook) As String
    Dim folderPath As String
    Dim fileName As String
    Dim candidate As String
    Dim current As String

    If wb Is Nothing Then Exit Function
    folderPath = GetParentFolderPublisher(wb.FullName)
    If folderPath = "" Then Exit Function

    fileName = Dir$(NormalizeFolderPathPublisher(folderPath) & "*.invSys.Config.xls*")
    Do While fileName <> ""
        candidate = ResolveWarehouseIdFromConfigWorkbookNamePublisher(fileName)
        If candidate <> "" Then
            If current = "" Then
                current = candidate
            ElseIf StrComp(current, candidate, vbTextCompare) <> 0 Then
                Exit Function
            End If
        End If
        fileName = Dir$
    Loop

    ResolveWarehouseIdFromSiblingConfigPublisher = current
End Function

Private Function ResolveWarehouseIdFromOpenConfigPublisher() As String
    Dim wb As Workbook
    Dim candidate As String
    Dim current As String

    For Each wb In Application.Workbooks
        candidate = ResolveWarehouseIdFromConfigWorkbookNamePublisher(wb.Name)
        If candidate <> "" Then
            If current = "" Then
                current = candidate
            ElseIf StrComp(current, candidate, vbTextCompare) <> 0 Then
                Exit Function
            End If
        End If
    Next wb

    ResolveWarehouseIdFromOpenConfigPublisher = current
End Function

Private Function ResolveWarehouseIdFromConfigWorkbookNamePublisher(ByVal wbName As String) As String
    Dim markerPos As Long

    markerPos = InStr(1, wbName, ".invSys.Config.", vbTextCompare)
    If markerPos > 1 Then ResolveWarehouseIdFromConfigWorkbookNamePublisher = Left$(wbName, markerPos - 1)
End Function

Private Function EnsureWarehouseConfigLoadedPublisher(ByVal warehouseId As String, ByRef report As String) As Boolean
    If Trim$(warehouseId) = "" Then
        report = "WarehouseId not resolved."
        Exit Function
    End If

    If modConfig.IsLoaded() Then
        If StrComp(Trim$(modConfig.GetWarehouseId()), Trim$(warehouseId), vbTextCompare) = 0 Then
            EnsureWarehouseConfigLoadedPublisher = True
            Exit Function
        End If
    End If

    EnsureWarehouseConfigLoadedPublisher = modConfig.LoadConfig(warehouseId, "")
    If Not EnsureWarehouseConfigLoadedPublisher Then report = "Config load failed for " & warehouseId & "."
End Function

Private Function BuildPublishKeyPublisher(ByVal wb As Workbook, ByVal warehouseId As String) As String
    If wb Is Nothing Then Exit Function
    BuildPublishKeyPublisher = LCase$(wb.FullName & "|" & warehouseId)
End Function

Private Sub EnsureRecentPublishRegistryPublisher()
    If mRecentPublishes Is Nothing Then
        Set mRecentPublishes = CreateObject("Scripting.Dictionary")
        mRecentPublishes.CompareMode = vbTextCompare
    End If
End Sub

Private Function ShouldSkipRecentPublishPublisher(ByVal publishKey As String) As Boolean
    EnsureRecentPublishRegistryPublisher
    If publishKey = "" Then Exit Function
    If Not mRecentPublishes.Exists(publishKey) Then Exit Function
    ShouldSkipRecentPublishPublisher = (DateDiff("s", CDate(mRecentPublishes(publishKey)), Now) < MIN_RECENT_PUBLISH_SECONDS)
End Function

Private Sub RecordRecentPublishPublisher(ByVal publishKey As String)
    EnsureRecentPublishRegistryPublisher
    If publishKey = "" Then Exit Sub
    mRecentPublishes(publishKey) = Now
End Sub

Private Function GetParentFolderPublisher(ByVal fullPath As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(fullPath, "\")
    If slashPos > 1 Then GetParentFolderPublisher = Left$(fullPath, slashPos - 1)
End Function

Private Function NormalizeFolderPathPublisher(ByVal folderPath As String) As String
    NormalizeFolderPathPublisher = Trim$(folderPath)
    If NormalizeFolderPathPublisher = "" Then Exit Function
    If Right$(NormalizeFolderPathPublisher, 1) <> "\" Then NormalizeFolderPathPublisher = NormalizeFolderPathPublisher & "\"
End Function
