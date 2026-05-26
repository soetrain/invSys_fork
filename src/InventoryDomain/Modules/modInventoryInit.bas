Attribute VB_Name = "modInventoryInit"
Option Explicit

Private gAppEvents As cInventoryAppEvents
Private gNextSourceSync As Date
Private gSourceSyncScheduled As Boolean
Private Const SOURCE_SYNC_INTERVAL_SECONDS As Long = 2
Private Const SOURCE_SYNC_IDLE_INTERVAL_SECONDS As Long = 2
Private Const SOURCE_SYNC_LOG_FILENAME As String = "invSys.Inventory.Sync.log"

Public Sub InitInventoryDomainAddin()
    Dim report As String

    If gAppEvents Is Nothing Then
        Set gAppEvents = New cInventoryAppEvents
        gAppEvents.Init
    End If
    Call modInventoryPublisher.PublishOpenInventorySnapshots(report)
    ScheduleSourceWorkbookSync
End Sub

Public Sub Auto_Open()
    InitInventoryDomainAddin
End Sub

Public Sub ScheduleSourceWorkbookSync(Optional ByVal delaySeconds As Long = 3)
    Dim procedureName As String

    If Not IsSourceSyncSchedulerHostInit() Then
        gSourceSyncScheduled = False
        AppendSyncLogEntry "SCHEDULE_SKIP", "Workbook=" & ThisWorkbook.Name & "|Reason=NotInventoryDomainAddin"
        Exit Sub
    End If

    procedureName = BuildSourceSyncProcedureInit()

    On Error Resume Next
    If gSourceSyncScheduled Then
        Application.OnTime EarliestTime:=gNextSourceSync, _
                           Procedure:=procedureName, _
                           Schedule:=False
    End If
    If Err.Number <> 0 Then
        AppendSyncLogEntry "CANCEL_WARN", "Workbook=" & ThisWorkbook.Name & "|Error=" & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    On Error GoTo ScheduleFailed
    If delaySeconds <= 0 Then delaySeconds = 3
    gNextSourceSync = Now + (CDbl(delaySeconds) / 86400#)
    Application.OnTime EarliestTime:=gNextSourceSync, _
                       Procedure:=procedureName, _
                       Schedule:=True
    gSourceSyncScheduled = True
    AppendSyncLogEntry "SCHEDULE", "NextRun=" & Format$(gNextSourceSync, "yyyy-mm-dd hh:nn:ss") & "|DelaySeconds=" & CStr(delaySeconds)
    Exit Sub

ScheduleFailed:
    gSourceSyncScheduled = False
    AppendSyncLogEntry "SCHEDULE_ERROR", "Workbook=" & ThisWorkbook.Name & "|Error=" & Err.Description
End Sub

Public Sub SyncSourceWorkbookFromCanonicalRuntime()
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevAlerts As Boolean
    Dim wb As Workbook
    Dim hasSyncTargets As Boolean
    Dim detectionLog As String
    Dim syncReport As String
    Dim pullReport As String

    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    gSourceSyncScheduled = False
    AppendSyncLogEntry "CANARY", "SchedulerFired=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    On Error GoTo CleanExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    detectionLog = "OpenWbs=" & CStr(Application.Workbooks.Count) & "|"
    For Each wb In Application.Workbooks
        detectionLog = detectionLog & wb.Name & "=" & CStr(ShouldSyncSourceWorkbookInit(wb)) & ";"
    Next wb
    AppendSyncLogEntry "DETECTION", detectionLog

    For Each wb In Application.Workbooks
        If ShouldSyncSourceWorkbookInit(wb) Then
            hasSyncTargets = True
            pullReport = vbNullString
            Call modInventoryApply.RefreshInvSysFromCanonicalRuntime(wb, "", pullReport)
            If syncReport <> "" Then syncReport = syncReport & " || "
            syncReport = syncReport & pullReport
        End If
    Next wb

CleanExit:
    If Err.Number <> 0 Then
        AppendSyncLogEntry "ERROR", "SyncSourceWorkbookFromCanonicalRuntime failed: " & Err.Description
    ElseIf hasSyncTargets Then
        AppendSyncLogEntry "SYNC", syncReport
    Else
        AppendSyncLogEntry "SYNC", "No source workbooks matched sync predicate."
    End If
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevAlerts
    If hasSyncTargets Then
        ScheduleSourceWorkbookSync SOURCE_SYNC_INTERVAL_SECONDS
    Else
        ScheduleSourceWorkbookSync SOURCE_SYNC_IDLE_INTERVAL_SECONDS
    End If
End Sub

Public Function GetSyncLogPath() As String
    GetSyncLogPath = ResolveSyncLogPathInit()
End Function

Public Sub ResetSyncLog()
    Dim logPath As String

    On Error Resume Next
    logPath = ResolveSyncLogPathInit()
    If Len(Dir$(logPath)) > 0 Then Kill logPath
    On Error GoTo 0
End Sub

Public Sub AppendSyncLogEntry(ByVal tag As String, ByVal valueText As String)
    Dim fileNum As Integer
    Dim logPath As String

    On Error Resume Next
    logPath = ResolveSyncLogPathInit()
    fileNum = FreeFile
    Open logPath For Append As #fileNum
    Print #fileNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & tag & " | " & valueText
    Close #fileNum
    On Error GoTo 0
End Sub

Private Function ShouldSyncSourceWorkbookInit(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function

    wbName = LCase$(Trim$(wb.Name))
    If wbName = "" Then Exit Function
    If Left$(wbName, 2) = "~$" Then Exit Function
    If wbName Like "*.xla" Or wbName Like "*.xlam" Then Exit Function
    If wbName Like "*.invsys.*.xls*" Then Exit Function
    If wbName Like "invsys.inbox.*.xls*" Then Exit Function
    If wbName Like "*.outbox.events.xls*" Then Exit Function
    If wbName Like "*.snapshot.inventory.xls*" Then Exit Function

    If wbName Like "*inventory_management*.xls*" Then
        ShouldSyncSourceWorkbookInit = True
        Exit Function
    End If

    ShouldSyncSourceWorkbookInit = WorkbookHasSyncTableInit(wb, "invSys") _
        And (WorkbookHasSyncTableInit(wb, "ReceivedTally") _
             Or WorkbookHasSyncTableInit(wb, "ShipmentsTally") _
             Or WorkbookHasSyncTableInit(wb, "ProductionOutput") _
             Or WorkbookHasSyncTableInit(wb, "Recipes"))
End Function

Private Function IsSourceSyncSchedulerHostInit() As Boolean
    Dim wbName As String

    On Error Resume Next
    wbName = LCase$(Trim$(ThisWorkbook.Name))
    If ThisWorkbook.IsAddin Then
        IsSourceSyncSchedulerHostInit = True
    ElseIf wbName Like "*.xla" Or wbName Like "*.xlam" Then
        IsSourceSyncSchedulerHostInit = True
    End If
    On Error GoTo 0
End Function

Private Function BuildSourceSyncProcedureInit() As String
    BuildSourceSyncProcedureInit = "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!modInventoryInit.SyncSourceWorkbookFromCanonicalRuntime"
End Function

Private Function WorkbookHasSyncTableInit(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function

    On Error Resume Next
    For Each ws In wb.Worksheets
        Set lo = ws.ListObjects(tableName)
        If Not lo Is Nothing Then
            WorkbookHasSyncTableInit = True
            Exit Function
        End If
        Set lo = Nothing
    Next ws
    On Error GoTo 0
End Function

Private Function ResolveSyncLogPathInit() As String
    Dim rootPath As String

    rootPath = Trim$(Environ$("TEMP"))
    If rootPath = "" Then rootPath = ThisWorkbook.Path
    If rootPath = "" Then rootPath = CurDir$
    If Right$(rootPath, 1) <> "\" Then rootPath = rootPath & "\"

    ResolveSyncLogPathInit = rootPath & SOURCE_SYNC_LOG_FILENAME
End Function
