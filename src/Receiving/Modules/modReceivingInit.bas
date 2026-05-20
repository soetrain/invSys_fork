Attribute VB_Name = "modReceivingInit"
Option Explicit

Private Type ReceivingContext
    WarehouseId As String
    StationId As String
    PathDataRoot As String
    PathSharePointRoot As String
    CurrentUserId As String
End Type

Public Type ReceivingReadinessResult
    IsReady As Boolean
    SnapshotStatus As String
    AuthStatus As String
    RuntimeStatus As String
    Messages As String
End Type

Private gAppEvents As cAppEvents

Private Const READINESS_STATUS_OK As String = "OK"
Private Const SNAPSHOT_STATUS_STALE As String = "STALE"
Private Const SNAPSHOT_STATUS_MISSING As String = "MISSING"
Private Const SNAPSHOT_STATUS_UNREADABLE As String = "UNREADABLE"
Private Const AUTH_STATUS_NO_USER As String = "NO_USER"
Private Const AUTH_STATUS_MISSING_CAPABILITY As String = "MISSING_CAPABILITY"
Private Const AUTH_STATUS_INACTIVE As String = "INACTIVE"
Private Const RUNTIME_STATUS_MISSING_TABLES As String = "MISSING_TABLES"
Private Const RUNTIME_STATUS_PATH_UNRESOLVED As String = "PATH_UNRESOLVED"

Private Const SHEET_RECEIVING_READY As String = "ReceivedTally"
Private Const SHEET_STATUS_ALIAS As String = "tblStatus"
Private Const SHEET_READMODEL_ALIAS As String = "tblReadModel"
Private Const SHEET_RECEIVING_ALIAS As String = "tblReceiving"
Private Const SHEET_INVENTORY_READY As String = "InventoryManagement"
Private Const SHEET_LOG_READY As String = "ReceivedLog"
Private Const TABLE_RECEIVING_READY As String = "ReceivedTally"
Private Const TABLE_READMODEL_READY As String = "invSys"
Private Const TABLE_STATUS_READY As String = "ReceivedLog"

Private Const SHAPE_RECEIVING_STATUS_ROW1 As String = "invSysReceivingReadinessRow1"
Private Const SHAPE_RECEIVING_STATUS_ROW2 As String = "invSysReceivingReadinessRow2"
Private Const SHAPE_RECEIVING_STATUS_ROW3 As String = "invSysReceivingReadinessRow3"
Private Const DEFAULT_STALE_THRESHOLD_SECONDS As Long = 3600

Public Sub InitReceivingAddin()
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim activeWb As Workbook

    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If
    Set activeWb = Application.ActiveWorkbook
    EnsureReceivingSurfaceForWorkbook activeWb
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEvents
End Sub

Public Sub Auto_Open()
    InitReceivingAddin
End Sub

Public Function CheckReceivingReadiness() As ReceivingReadinessResult
    CheckReceivingReadiness = CheckReceivingReadinessForWorkbook(Application.ActiveWorkbook)
End Function

Public Function CheckReceivingReadinessPacked(Optional ByVal targetWb As Workbook = Nothing) As String
    Dim readiness As ReceivingReadinessResult

    On Error GoTo FailPacked
    readiness = CheckReceivingReadinessForWorkbook(targetWb)
    CheckReceivingReadinessPacked = _
        "IsReady=" & CStr(readiness.IsReady) & _
        "|SnapshotStatus=" & readiness.SnapshotStatus & _
        "|AuthStatus=" & readiness.AuthStatus & _
        "|RuntimeStatus=" & readiness.RuntimeStatus & _
        "|Messages=" & readiness.Messages
    Exit Function

FailPacked:
    CheckReceivingReadinessPacked = _
        "IsReady=False" & _
        "|SnapshotStatus=ERROR" & _
        "|AuthStatus=ERROR" & _
        "|RuntimeStatus=ERROR" & _
        "|Messages=Readiness check failed: " & Err.Description
End Function

Public Function CheckReceivingReadinessForWorkbook(Optional ByVal targetWb As Workbook = Nothing) As ReceivingReadinessResult
    Dim wb As Workbook
    Dim ctx As ReceivingContext
    Dim result As ReceivingReadinessResult

    On Error GoTo FailReadiness
    Set wb = targetWb
    If wb Is Nothing Then Set wb = Application.ActiveWorkbook

    result.RuntimeStatus = ResolveRuntimeStatusReadiness(wb, ctx)
    result.SnapshotStatus = ResolveSnapshotStatusReadiness(wb, ctx)
    result.AuthStatus = ResolveAuthStatusReadiness(ctx)

    AppendReadinessMessage result.Messages, ResolveSnapshotMessageReadiness(result.SnapshotStatus, wb)
    AppendReadinessMessage result.Messages, ResolveAuthMessageReadiness(result.AuthStatus)
    AppendReadinessMessage result.Messages, ResolveRuntimeMessageReadiness(result.RuntimeStatus)

    result.IsReady = (StrComp(result.SnapshotStatus, READINESS_STATUS_OK, vbTextCompare) = 0 _
                      And StrComp(result.AuthStatus, READINESS_STATUS_OK, vbTextCompare) = 0 _
                      And StrComp(result.RuntimeStatus, READINESS_STATUS_OK, vbTextCompare) = 0)
    CheckReceivingReadinessForWorkbook = result
    Exit Function

FailReadiness:
    result.IsReady = False
    result.SnapshotStatus = "ERROR"
    result.AuthStatus = "ERROR"
    result.RuntimeStatus = "ERROR"
    result.Messages = "Readiness check failed: " & Err.Description
    CheckReceivingReadinessForWorkbook = result
End Function

Public Sub ApplyReceivingReadinessForWorkbook(Optional ByVal targetWb As Workbook = Nothing, _
                                              Optional ByVal initializeUiWhenReady As Boolean = True)
    Dim wb As Workbook
    Dim readiness As ReceivingReadinessResult

    Set wb = targetWb
    If wb Is Nothing Then Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If wb.IsAddin Then Exit Sub

    readiness = CheckReceivingReadinessForWorkbook(wb)
    If readiness.IsReady Then
        ClearReceivingReadinessPanel wb
        If initializeUiWhenReady Then modTS_Received.InitializeReceivingUiForWorkbook wb
    Else
        RenderReceivingReadinessPanel wb, readiness
    End If

    LogDiagnosticEvent "RECEIVING-READINESS", _
        "Workbook=" & SafeWorkbookNameReadiness(wb) & _
        "|SnapshotStatus=" & readiness.SnapshotStatus & _
        "|AuthStatus=" & readiness.AuthStatus & _
        "|RuntimeStatus=" & readiness.RuntimeStatus & _
        "|Messages=" & readiness.Messages
End Sub

Public Function GetReceivingReadinessPanelText(Optional ByVal targetWb As Workbook = Nothing) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim shapeNames As Variant
    Dim shp As Shape
    Dim lineOut As String

    Set wb = targetWb
    If wb Is Nothing Then Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then Exit Function

    Set ws = ResolveReceivingStatusSheet(wb)
    If ws Is Nothing Then Exit Function

    shapeNames = Array(SHAPE_RECEIVING_STATUS_ROW1, SHAPE_RECEIVING_STATUS_ROW2, SHAPE_RECEIVING_STATUS_ROW3)
    For i = LBound(shapeNames) To UBound(shapeNames)
        On Error Resume Next
        Set shp = ws.Shapes(CStr(shapeNames(i)))
        On Error GoTo 0
        If Not shp Is Nothing Then
            lineOut = Trim$(shp.TextFrame.Characters.Text)
            If lineOut <> "" Then
                If Len(GetReceivingReadinessPanelText) > 0 Then GetReceivingReadinessPanelText = GetReceivingReadinessPanelText & " | "
                GetReceivingReadinessPanelText = GetReceivingReadinessPanelText & lineOut
            End If
        End If
        Set shp = Nothing
    Next i
End Function

Public Sub EnsureReceivingSurfaceForWorkbook(ByVal wb As Workbook)
    Dim prevEvents As Boolean

    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    If Not IsLikelyReceivingWorkbookReadiness(wb) Then Exit Sub

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    ApplyReceivingReadinessForWorkbook wb, True
    Application.EnableEvents = prevEvents
End Sub

Private Function ResolveRuntimeStatusReadiness(ByVal wb As Workbook, _
                                               ByRef ctx As ReceivingContext) As String
    Dim configLoaded As Boolean
    Dim warehouseId As String
    Dim priorRootOverride As String
    Dim workbookRoot As String

    ctx.CurrentUserId = ResolveCurrentUserIdReadiness()

    If wb Is Nothing Or wb.IsAddin Then
        ResolveRuntimeStatusReadiness = RUNTIME_STATUS_PATH_UNRESOLVED
        Exit Function
    End If

    If Not WorkbookHasReceivingSurfacesReadiness(wb) Then
        ResolveRuntimeStatusReadiness = RUNTIME_STATUS_MISSING_TABLES
        Exit Function
    End If

    warehouseId = ResolveWarehouseIdFromWorkbookReadiness(wb)
    If warehouseId = "" Then
        ResolveRuntimeStatusReadiness = RUNTIME_STATUS_PATH_UNRESOLVED
        Exit Function
    End If

    priorRootOverride = modRuntimeWorkbooks.GetCoreDataRootOverride()
    workbookRoot = ResolveRuntimeRootFromWorkbookReadiness(wb, warehouseId)
    If workbookRoot <> "" Then modRuntimeWorkbooks.SetCoreDataRootOverride workbookRoot

    configLoaded = modConfig.LoadConfig(warehouseId, "")
    If Not configLoaded Then RestoreRuntimeRootOverrideReadiness priorRootOverride
    If Not configLoaded Then
        ResolveRuntimeStatusReadiness = RUNTIME_STATUS_PATH_UNRESOLVED
        Exit Function
    End If

    ctx.WarehouseId = Trim$(modConfig.GetWarehouseId())
    ctx.StationId = Trim$(modConfig.GetStationId())
    ctx.PathDataRoot = NormalizeFolderPathReadiness(modConfig.GetString("PathDataRoot", ""), False)
    ctx.PathSharePointRoot = NormalizeFolderPathReadiness(modConfig.GetString("PathSharePointRoot", ""), False)

    If ctx.WarehouseId = "" Or ctx.PathDataRoot = "" Or Not FolderExistsReadiness(ctx.PathDataRoot) Then
        RestoreRuntimeRootOverrideReadiness priorRootOverride
        ResolveRuntimeStatusReadiness = RUNTIME_STATUS_PATH_UNRESOLVED
        Exit Function
    End If

    modRuntimeWorkbooks.SetCoreDataRootOverride ctx.PathDataRoot
    ResolveRuntimeStatusReadiness = READINESS_STATUS_OK
End Function

Private Function ResolveSnapshotStatusReadiness(ByVal wb As Workbook, _
                                                ByRef ctx As ReceivingContext) As String
    Dim snapshotPath As String
    Dim snapshotAgeSeconds As Long
    Dim staleThresholdSeconds As Long

    If StrComp(ResolveRuntimeStatusReadinessCache(ctx), READINESS_STATUS_OK, vbTextCompare) <> 0 Then
        ResolveSnapshotStatusReadiness = SNAPSHOT_STATUS_MISSING
        Exit Function
    End If

    snapshotPath = ctx.PathDataRoot & "\" & ctx.WarehouseId & ".invSys.Snapshot.Inventory.xlsb"
    If Not FileExistsReadiness(snapshotPath) Then
        ResolveSnapshotStatusReadiness = SNAPSHOT_STATUS_MISSING
        Exit Function
    End If

    If Not CanOpenWorkbookReadiness(snapshotPath) Then
        ResolveSnapshotStatusReadiness = SNAPSHOT_STATUS_UNREADABLE
        Exit Function
    End If

    staleThresholdSeconds = ResolveStaleThresholdSecondsReadiness()
    snapshotAgeSeconds = ResolveSnapshotAgeSecondsReadiness(wb, snapshotPath)
    If SnapshotMarkedStaleReadiness(wb) Or snapshotAgeSeconds < 0 Or snapshotAgeSeconds > staleThresholdSeconds Then
        ResolveSnapshotStatusReadiness = SNAPSHOT_STATUS_STALE
        Exit Function
    End If

    ResolveSnapshotStatusReadiness = READINESS_STATUS_OK
End Function

Private Function ResolveAuthStatusReadiness(ByRef ctx As ReceivingContext) As String
    Dim authPath As String
    Dim wbAuth As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim openedTransient As Boolean
    Dim userRow As Long

    On Error GoTo FailAuth

    If Trim$(ctx.CurrentUserId) = "" Then
        ResolveAuthStatusReadiness = AUTH_STATUS_NO_USER
        Exit Function
    End If
    If Trim$(ctx.WarehouseId) = "" Or Trim$(ctx.PathDataRoot) = "" Then
        ResolveAuthStatusReadiness = AUTH_STATUS_NO_USER
        Exit Function
    End If

    authPath = ctx.PathDataRoot & "\" & ctx.WarehouseId & ".invSys.Auth.xlsb"
    If Not FileExistsReadiness(authPath) Then
        ResolveAuthStatusReadiness = AUTH_STATUS_NO_USER
        Exit Function
    End If

    Set wbAuth = FindOpenWorkbookByPathReadiness(authPath)
    If wbAuth Is Nothing Then
        Set wbAuth = Application.Workbooks.Open(Filename:=authPath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedTransient = Not wbAuth Is Nothing
    End If
    If wbAuth Is Nothing Then
        ResolveAuthStatusReadiness = AUTH_STATUS_NO_USER
        Exit Function
    End If

    Set loUsers = FindTableByNameReadiness(wbAuth, "tblUsers")
    Set loCaps = FindTableByNameReadiness(wbAuth, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then
        ResolveAuthStatusReadiness = AUTH_STATUS_NO_USER
        GoTo CleanExit
    End If

    userRow = FindRowByValueReadiness(loUsers, "UserId", ctx.CurrentUserId)
    If userRow = 0 Then
        ResolveAuthStatusReadiness = AUTH_STATUS_NO_USER
        GoTo CleanExit
    End If
    If StrComp(UCase$(SafeTrimReadiness(GetTableValueReadiness(loUsers, userRow, "Status"))), "ACTIVE", vbTextCompare) <> 0 Then
        ResolveAuthStatusReadiness = AUTH_STATUS_INACTIVE
        GoTo CleanExit
    End If
    If Not HasActiveCapabilityReadiness(loCaps, ctx.CurrentUserId, "RECEIVE_POST", ctx.WarehouseId, ctx.StationId) Then
        ResolveAuthStatusReadiness = AUTH_STATUS_MISSING_CAPABILITY
        GoTo CleanExit
    End If

    ResolveAuthStatusReadiness = READINESS_STATUS_OK

CleanExit:
    If openedTransient And Not wbAuth Is Nothing Then
        On Error Resume Next
        wbAuth.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Exit Function

FailAuth:
    ResolveAuthStatusReadiness = AUTH_STATUS_NO_USER
    Resume CleanExit
End Function

Private Function ResolveRuntimeStatusReadinessCache(ByRef ctx As ReceivingContext) As String
    If Trim$(ctx.WarehouseId) = "" Or Trim$(ctx.PathDataRoot) = "" Then
        ResolveRuntimeStatusReadinessCache = RUNTIME_STATUS_PATH_UNRESOLVED
    Else
        ResolveRuntimeStatusReadinessCache = READINESS_STATUS_OK
    End If
End Function

Private Function ResolveSnapshotMessageReadiness(ByVal statusCode As String, ByVal wb As Workbook) As String
    Dim snapshotPath As String
    Dim ageSeconds As Long

    Select Case UCase$(Trim$(statusCode))
        Case SNAPSHOT_STATUS_STALE
            snapshotPath = ResolveSnapshotPathForMessageReadiness(wb)
            ageSeconds = ResolveSnapshotAgeSecondsReadiness(wb, snapshotPath)
            If ageSeconds < 0 Then
                ResolveSnapshotMessageReadiness = "Snapshot freshness is unknown. Click Refresh Inventory before posting."
            Else
                ResolveSnapshotMessageReadiness = "Snapshot is " & FormatAgeReadiness(ageSeconds) & " old. Click Refresh Inventory before posting."
            End If
        Case SNAPSHOT_STATUS_MISSING
            ResolveSnapshotMessageReadiness = "Snapshot workbook is missing at the configured path. Click Refresh Inventory before posting."
        Case SNAPSHOT_STATUS_UNREADABLE
            ResolveSnapshotMessageReadiness = "Snapshot workbook could not be opened. Click Refresh Inventory or contact your admin."
    End Select
End Function

Private Function ResolveAuthMessageReadiness(ByVal statusCode As String) As String
    Select Case UCase$(Trim$(statusCode))
        Case AUTH_STATUS_NO_USER
            ResolveAuthMessageReadiness = "Your account is not provisioned for this warehouse. Run Setup Tester Station or contact your admin."
        Case AUTH_STATUS_MISSING_CAPABILITY
            ResolveAuthMessageReadiness = "Your account does not have RECEIVE_POST. Contact your admin."
        Case AUTH_STATUS_INACTIVE
            ResolveAuthMessageReadiness = "Your account is inactive. Contact your admin."
    End Select
End Function

Private Function ResolveRuntimeMessageReadiness(ByVal statusCode As String) As String
    Select Case UCase$(Trim$(statusCode))
        Case RUNTIME_STATUS_MISSING_TABLES
            ResolveRuntimeMessageReadiness = "Workbook is missing required tables. Run Setup Tester Station."
        Case RUNTIME_STATUS_PATH_UNRESOLVED
            ResolveRuntimeMessageReadiness = "Runtime path could not be resolved. Run Setup Tester Station."
    End Select
End Function

Private Sub RenderReceivingReadinessPanel(ByVal wb As Workbook, ByRef readiness As ReceivingReadinessResult)
    Dim ws As Worksheet
    Dim issues As Collection
    Dim issueText As Variant
    Dim rowIndex As Long
    Dim shapeName As String

    Set ws = ResolveReceivingStatusSheet(wb)
    If ws Is Nothing Then Exit Sub

    Set issues = BuildIssueCollectionReadiness(wb, readiness)
    ClearReceivingReadinessPanel wb

    rowIndex = 1
    For Each issueText In issues
        shapeName = ResolveStatusShapeNameReadiness(rowIndex)
        If shapeName <> "" Then
            ApplyStatusShapeReadiness ws, shapeName, rowIndex, CStr(issueText)
        End If
        rowIndex = rowIndex + 1
    Next issueText
End Sub

Private Sub ClearReceivingReadinessPanel(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim shapeNames As Variant
    Dim i As Long

    Set ws = ResolveReceivingStatusSheet(wb)
    If ws Is Nothing Then Exit Sub

    shapeNames = Array(SHAPE_RECEIVING_STATUS_ROW1, SHAPE_RECEIVING_STATUS_ROW2, SHAPE_RECEIVING_STATUS_ROW3)
    For i = LBound(shapeNames) To UBound(shapeNames)
        On Error Resume Next
        ws.Shapes(CStr(shapeNames(i))).Delete
        On Error GoTo 0
    Next i
End Sub

Private Function BuildIssueCollectionReadiness(ByVal wb As Workbook, ByRef readiness As ReceivingReadinessResult) As Collection
    Dim issues As New Collection

    If StrComp(readiness.SnapshotStatus, READINESS_STATUS_OK, vbTextCompare) <> 0 Then
        issues.Add "Snapshot " & readiness.SnapshotStatus & ": " & ResolveSnapshotMessageReadiness(readiness.SnapshotStatus, wb)
    End If
    If StrComp(readiness.AuthStatus, READINESS_STATUS_OK, vbTextCompare) <> 0 Then
        issues.Add "Auth " & readiness.AuthStatus & ": " & ResolveAuthMessageReadiness(readiness.AuthStatus)
    End If
    If StrComp(readiness.RuntimeStatus, READINESS_STATUS_OK, vbTextCompare) <> 0 Then
        issues.Add "Runtime " & readiness.RuntimeStatus & ": " & ResolveRuntimeMessageReadiness(readiness.RuntimeStatus)
    End If

    Set BuildIssueCollectionReadiness = issues
End Function

Private Sub ApplyStatusShapeReadiness(ByVal ws As Worksheet, _
                                      ByVal shapeName As String, _
                                      ByVal rowIndex As Long, _
                                      ByVal issueText As String)
    Dim targetRange As Range
    Dim shp As Shape
    Dim fillRgb As Long

    Set targetRange = ws.Range("A" & rowIndex & ":U" & rowIndex)
    Set shp = ws.Shapes.AddShape(5, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    shp.Name = shapeName
    shp.Placement = 1
    shp.TextFrame.Characters.Text = issueText
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.Line.Visible = False

    fillRgb = ResolveStatusColorReadiness(issueText)
    shp.Fill.Visible = True
    shp.Fill.ForeColor.RGB = fillRgb
End Sub

Private Function ResolveStatusColorReadiness(ByVal issueText As String) As Long
    Dim upperText As String

    upperText = UCase$(issueText)
    Select Case True
        Case InStr(1, upperText, SNAPSHOT_STATUS_STALE, vbTextCompare) > 0
            ResolveStatusColorReadiness = RGB(192, 123, 17)
        Case InStr(1, upperText, SNAPSHOT_STATUS_MISSING, vbTextCompare) > 0 _
             Or InStr(1, upperText, SNAPSHOT_STATUS_UNREADABLE, vbTextCompare) > 0 _
             Or InStr(1, upperText, RUNTIME_STATUS_MISSING_TABLES, vbTextCompare) > 0 _
             Or InStr(1, upperText, RUNTIME_STATUS_PATH_UNRESOLVED, vbTextCompare) > 0
            ResolveStatusColorReadiness = RGB(192, 57, 43)
        Case Else
            ResolveStatusColorReadiness = RGB(52, 73, 94)
    End Select
End Function

Private Function ResolveStatusShapeNameReadiness(ByVal rowIndex As Long) As String
    Select Case rowIndex
        Case 1
            ResolveStatusShapeNameReadiness = SHAPE_RECEIVING_STATUS_ROW1
        Case 2
            ResolveStatusShapeNameReadiness = SHAPE_RECEIVING_STATUS_ROW2
        Case 3
            ResolveStatusShapeNameReadiness = SHAPE_RECEIVING_STATUS_ROW3
    End Select
End Function

Private Function ResolveReceivingStatusSheet(ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Exit Function

    Set ResolveReceivingStatusSheet = WorkbookSheetByNameReadiness(wb, SHEET_RECEIVING_READY)
    If ResolveReceivingStatusSheet Is Nothing Then Set ResolveReceivingStatusSheet = WorkbookSheetByNameReadiness(wb, SHEET_RECEIVING_ALIAS)
    If ResolveReceivingStatusSheet Is Nothing Then
        On Error Resume Next
        Set ResolveReceivingStatusSheet = wb.Worksheets(1)
        On Error GoTo 0
    End If
End Function

Private Function WorkbookHasReceivingSurfacesReadiness(ByVal wb As Workbook) As Boolean
    WorkbookHasReceivingSurfacesReadiness = _
        (WorksheetExistsReadiness(wb, SHEET_RECEIVING_READY) Or WorksheetExistsReadiness(wb, SHEET_RECEIVING_ALIAS)) _
        And (WorksheetExistsReadiness(wb, SHEET_INVENTORY_READY) Or WorksheetExistsReadiness(wb, SHEET_READMODEL_ALIAS)) _
        And (WorksheetExistsReadiness(wb, SHEET_LOG_READY) Or WorksheetExistsReadiness(wb, SHEET_STATUS_ALIAS)) _
        And TableExistsReadiness(wb, TABLE_RECEIVING_READY) _
        And TableExistsReadiness(wb, TABLE_READMODEL_READY) _
        And TableExistsReadiness(wb, TABLE_STATUS_READY)
End Function

Private Function IsLikelyReceivingWorkbookReadiness(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Or wb.IsAddin Then Exit Function
    wbName = LCase$(Trim$(wb.Name))

    If WorkbookHasReceivingSurfacesReadiness(wb) Then
        IsLikelyReceivingWorkbookReadiness = True
        Exit Function
    End If

    If wbName Like "*.receiving.operator.xls*" Then
        IsLikelyReceivingWorkbookReadiness = True
        Exit Function
    End If

    If modConfig.IsLoaded() Then
        IsLikelyReceivingWorkbookReadiness = (Trim$(modConfig.GetWarehouseId()) <> "")
    End If
End Function

Private Function ResolveWarehouseIdFromWorkbookReadiness(ByVal wb As Workbook) As String
    Dim markerPos As Long
    Dim wbName As String
    Dim nameParts() As String

    If wb Is Nothing Then Exit Function
    wbName = Trim$(wb.Name)
    markerPos = InStr(1, wbName, ".Receiving.Operator.xls", vbTextCompare)
    If markerPos > 1 Then
        ResolveWarehouseIdFromWorkbookReadiness = Left$(wbName, markerPos - 1)
        Exit Function
    End If

    markerPos = InStr(1, wbName, "_Receiving_Operator.xls", vbTextCompare)
    If markerPos > 1 Then
        nameParts = Split(Left$(wbName, markerPos - 1), "_")
        If UBound(nameParts) >= 0 Then ResolveWarehouseIdFromWorkbookReadiness = nameParts(0)
        If ResolveWarehouseIdFromWorkbookReadiness <> "" Then Exit Function
    End If

    If modConfig.IsLoaded() Then ResolveWarehouseIdFromWorkbookReadiness = Trim$(modConfig.GetWarehouseId())
End Function

Private Function ResolveSnapshotPathForMessageReadiness(ByVal wb As Workbook) As String
    Dim warehouseId As String

    warehouseId = ResolveWarehouseIdFromWorkbookReadiness(wb)
    If warehouseId = "" Then Exit Function
    If Not modConfig.LoadConfig(warehouseId, "") Then Exit Function
    ResolveSnapshotPathForMessageReadiness = NormalizeFolderPathReadiness(modConfig.GetString("PathDataRoot", ""), False) & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
End Function

Private Function ResolveSnapshotAgeSecondsReadiness(ByVal wb As Workbook, ByVal snapshotPath As String) As Long
    Dim refreshUtc As Date

    refreshUtc = ResolveLastRefreshUtcReadiness(wb)
    If refreshUtc = 0 Then
        If snapshotPath <> "" And FileExistsReadiness(snapshotPath) Then
            On Error Resume Next
            refreshUtc = FileDateTime(snapshotPath)
            On Error GoTo 0
        End If
    End If
    If refreshUtc = 0 Then
        ResolveSnapshotAgeSecondsReadiness = -1
    Else
        ResolveSnapshotAgeSecondsReadiness = DateDiff("s", refreshUtc, Now)
    End If
End Function

Private Function ResolveLastRefreshUtcReadiness(ByVal wb As Workbook) As Date
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim valueIn As Variant

    Set lo = FindTableByNameReadiness(wb, TABLE_READMODEL_READY)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    colIndex = lo.ListColumns("LastRefreshUTC").Index
    For rowIndex = 1 To lo.ListRows.Count
        valueIn = lo.DataBodyRange.Cells(rowIndex, colIndex).Value
        If IsDate(valueIn) Then
            If ResolveLastRefreshUtcReadiness = 0 Or CDate(valueIn) > ResolveLastRefreshUtcReadiness Then
                ResolveLastRefreshUtcReadiness = CDate(valueIn)
            End If
        End If
    Next rowIndex
End Function

Private Function SnapshotMarkedStaleReadiness(ByVal wb As Workbook) As Boolean
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim valueText As String

    Set lo = FindTableByNameReadiness(wb, TABLE_READMODEL_READY)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    colIndex = lo.ListColumns("IsStale").Index
    For rowIndex = 1 To lo.ListRows.Count
        valueText = UCase$(SafeTrimReadiness(lo.DataBodyRange.Cells(rowIndex, colIndex).Value))
        If valueText = "TRUE" Or valueText = "YES" Or valueText = "1" Then
            SnapshotMarkedStaleReadiness = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function ResolveStaleThresholdSecondsReadiness() As Long
    ResolveStaleThresholdSecondsReadiness = modConfig.GetLong("AutoRefreshIntervalSeconds", DEFAULT_STALE_THRESHOLD_SECONDS)
    If ResolveStaleThresholdSecondsReadiness <= 0 Then ResolveStaleThresholdSecondsReadiness = DEFAULT_STALE_THRESHOLD_SECONDS
End Function

Private Function CanOpenWorkbookReadiness(ByVal workbookPath As String) As Boolean
    Dim wb As Workbook
    Dim alreadyOpen As Boolean

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then Exit Function

    Set wb = FindOpenWorkbookByPathReadiness(workbookPath)
    alreadyOpen = Not wb Is Nothing
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Application.Workbooks.Open(Filename:=workbookPath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        On Error GoTo 0
    End If
    If wb Is Nothing Then Exit Function

    CanOpenWorkbookReadiness = True
    If Not alreadyOpen Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Function

Private Function HasActiveCapabilityReadiness(ByVal loCaps As ListObject, _
                                              ByVal userId As String, _
                                              ByVal capabilityName As String, _
                                              ByVal warehouseId As String, _
                                              ByVal stationId As String) As Boolean
    Dim rowIndex As Long
    Dim capWarehouse As String
    Dim capStation As String

    If loCaps Is Nothing Then Exit Function
    If loCaps.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To loCaps.ListRows.Count
        If StrComp(SafeTrimReadiness(GetTableValueReadiness(loCaps, rowIndex, "UserId")), userId, vbTextCompare) <> 0 Then GoTo ContinueLoop
        If StrComp(UCase$(SafeTrimReadiness(GetTableValueReadiness(loCaps, rowIndex, "Capability"))), UCase$(capabilityName), vbTextCompare) <> 0 Then GoTo ContinueLoop
        If StrComp(UCase$(SafeTrimReadiness(GetTableValueReadiness(loCaps, rowIndex, "Status"))), "ACTIVE", vbTextCompare) <> 0 Then GoTo ContinueLoop

        capWarehouse = SafeTrimReadiness(GetTableValueReadiness(loCaps, rowIndex, "WarehouseId"))
        capStation = SafeTrimReadiness(GetTableValueReadiness(loCaps, rowIndex, "StationId"))
        If capWarehouse <> "" And StrComp(capWarehouse, warehouseId, vbTextCompare) <> 0 Then GoTo ContinueLoop
        If capStation <> "" And stationId <> "" And StrComp(capStation, stationId, vbTextCompare) <> 0 Then GoTo ContinueLoop

        HasActiveCapabilityReadiness = True
        Exit Function
ContinueLoop:
    Next rowIndex
End Function

Private Function FindOpenWorkbookByPathReadiness(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(SafeWorkbookPathReadiness(wb), workbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByPathReadiness = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindTableByNameReadiness(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTableByNameReadiness = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTableByNameReadiness Is Nothing Then Exit Function
    Next ws
End Function

Private Function WorkbookSheetByNameReadiness(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set WorkbookSheetByNameReadiness = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function WorksheetExistsReadiness(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    WorksheetExistsReadiness = Not WorkbookSheetByNameReadiness(wb, sheetName) Is Nothing
End Function

Private Function TableExistsReadiness(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    TableExistsReadiness = Not FindTableByNameReadiness(wb, tableName) Is Nothing
End Function

Private Function FindRowByValueReadiness(ByVal lo As ListObject, ByVal columnName As String, ByVal matchValue As String) As Long
    Dim rowIndex As Long
    Dim colIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    colIndex = lo.ListColumns(columnName).Index
    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(SafeTrimReadiness(lo.DataBodyRange.Cells(rowIndex, colIndex).Value), matchValue, vbTextCompare) = 0 Then
            FindRowByValueReadiness = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function GetTableValueReadiness(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > lo.ListRows.Count Then Exit Function

    GetTableValueReadiness = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value
End Function

Private Sub AppendReadinessMessage(ByRef messageText As String, ByVal nextMessage As String)
    nextMessage = Trim$(nextMessage)
    If nextMessage = "" Then Exit Sub
    If Len(messageText) > 0 Then messageText = messageText & "|"
    messageText = messageText & nextMessage
End Sub

Private Function ResolveCurrentUserIdReadiness() As String
    ResolveCurrentUserIdReadiness = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If ResolveCurrentUserIdReadiness = "" Then ResolveCurrentUserIdReadiness = Trim$(Application.UserName)
End Function

Private Function NormalizeFolderPathReadiness(ByVal pathIn As String, ByVal includeTrailingSlash As Boolean) As String
    NormalizeFolderPathReadiness = modConfig.NormalizeFolderPathForRuntime(pathIn, includeTrailingSlash)
End Function

Private Function ResolveRuntimeRootFromWorkbookReadiness(ByVal wb As Workbook, ByVal warehouseId As String) As String
    Dim workbookRoot As String
    Dim sepPos As Long

    If wb Is Nothing Then Exit Function
    If warehouseId = "" Then Exit Function

    workbookRoot = NormalizeFolderPathReadiness(modRuntimeWorkbooks.GetCoreDataRootOverride(), False)
    If RuntimeRootHasWarehouseArtifactsReadiness(workbookRoot, warehouseId) Then
        ResolveRuntimeRootFromWorkbookReadiness = workbookRoot
        Exit Function
    End If

    workbookRoot = NormalizeFolderPathReadiness(modRuntimeWorkbooks.TryResolveExistingRuntimeRoot(warehouseId), False)
    If RuntimeRootHasWarehouseArtifactsReadiness(workbookRoot, warehouseId) Then
        ResolveRuntimeRootFromWorkbookReadiness = workbookRoot
        Exit Function
    End If

    workbookRoot = NormalizeFolderPathReadiness(wb.Path, False)
    If workbookRoot = "" Then
        workbookRoot = NormalizeFolderPathReadiness(SafeWorkbookPathReadiness(wb), False)
        sepPos = InStrRev(workbookRoot, "\")
        If sepPos > 0 Then workbookRoot = Left$(workbookRoot, sepPos - 1)
    End If
    If workbookRoot = "" Then Exit Function

    If RuntimeRootHasWarehouseArtifactsReadiness(workbookRoot, warehouseId) Then
        ResolveRuntimeRootFromWorkbookReadiness = workbookRoot
    End If
End Function

Private Function RuntimeRootHasWarehouseArtifactsReadiness(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderPathReadiness(rootPath, False)
    warehouseId = Trim$(warehouseId)
    If rootPath = "" Or warehouseId = "" Then Exit Function

    RuntimeRootHasWarehouseArtifactsReadiness = _
        FileExistsReadiness(rootPath & "\" & warehouseId & ".invSys.Config.xlsb") _
        And FileExistsReadiness(rootPath & "\" & warehouseId & ".invSys.Auth.xlsb")
End Function

Private Sub RestoreRuntimeRootOverrideReadiness(ByVal priorRootOverride As String)
    If Trim$(priorRootOverride) = "" Then
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride priorRootOverride
    End If
End Sub

Private Function FolderExistsReadiness(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExistsReadiness = CreateObject("Scripting.FileSystemObject").FolderExists(folderPath)
    On Error GoTo 0
End Function

Private Function FileExistsReadiness(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExistsReadiness = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
    On Error GoTo 0
End Function

Private Function SafeTrimReadiness(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    SafeTrimReadiness = Trim$(CStr(valueIn))
End Function

Private Function SafeWorkbookPathReadiness(ByVal wb As Workbook) As String
    On Error Resume Next
    SafeWorkbookPathReadiness = Trim$(wb.FullName)
    On Error GoTo 0
End Function

Private Function SafeWorkbookNameReadiness(ByVal wb As Workbook) As String
    On Error Resume Next
    SafeWorkbookNameReadiness = Trim$(wb.Name)
    On Error GoTo 0
End Function

Private Function FormatAgeReadiness(ByVal ageSeconds As Long) As String
    If ageSeconds < 0 Then
        FormatAgeReadiness = "an unknown amount of time"
    ElseIf ageSeconds >= 7200 Then
        FormatAgeReadiness = CStr(ageSeconds \ 3600) & " hours"
    ElseIf ageSeconds >= 3600 Then
        FormatAgeReadiness = "1 hour"
    ElseIf ageSeconds >= 120 Then
        FormatAgeReadiness = CStr(ageSeconds \ 60) & " minutes"
    ElseIf ageSeconds >= 60 Then
        FormatAgeReadiness = "1 minute"
    Else
        FormatAgeReadiness = CStr(ageSeconds) & " seconds"
    End If
End Function
