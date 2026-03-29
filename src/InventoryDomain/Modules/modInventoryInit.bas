Attribute VB_Name = "modInventoryInit"
Option Explicit

Private gAppEvents As cInventoryAppEvents
Private gNextSourceSync As Date
Private gSourceSyncScheduled As Boolean

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
    On Error Resume Next
    If gSourceSyncScheduled Then
        Application.OnTime EarliestTime:=gNextSourceSync, _
                           Procedure:="'" & ThisWorkbook.Name & "'!modInventoryInit.SyncSourceWorkbookFromCanonicalRuntime", _
                           Schedule:=False
    End If
    On Error GoTo 0

    If delaySeconds <= 0 Then delaySeconds = 3
    gNextSourceSync = Now + (CDbl(delaySeconds) / 86400#)
    Application.OnTime EarliestTime:=gNextSourceSync, _
                       Procedure:="'" & ThisWorkbook.Name & "'!modInventoryInit.SyncSourceWorkbookFromCanonicalRuntime"
    gSourceSyncScheduled = True
End Sub

Public Sub SyncSourceWorkbookFromCanonicalRuntime()
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevAlerts As Boolean
    Dim prevCalculation As XlCalculation
    Dim wb As Workbook

    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    prevCalculation = Application.Calculation
    gSourceSyncScheduled = False

    On Error GoTo CleanExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    For Each wb In Application.Workbooks
        If ShouldSyncSourceWorkbookInit(wb) Then
            Call modInventoryApply.RefreshInvSysFromCanonicalRuntime(wb)
        End If
    Next wb

CleanExit:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevAlerts
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
