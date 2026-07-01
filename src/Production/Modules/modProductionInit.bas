Attribute VB_Name = "modProductionInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitProductionAddin()
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean

    ApplyRememberedRuntimeTargetProduction
    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If
    mProduction.InitializeProductionUiForWorkbook ThisWorkbook
    EnsureProductionSurfaceForWorkbook Application.ActiveWorkbook
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEvents
End Sub

Private Sub ApplyRememberedRuntimeTargetProduction()
    On Error Resume Next
    Call modRibbonRuntimeStatus.TryApplyRememberedWarehouseTarget
    On Error GoTo 0
End Sub

Public Sub Auto_Open()
    InitProductionAddin
End Sub

Public Sub EnsureProductionSurfaceForWorkbook(ByVal wb As Workbook)
    Dim prevEvents As Boolean

    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    If Not IsLikelyProductionWorkbook(wb) Then Exit Sub
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    mProduction.InitializeProductionUiForWorkbook wb
    Application.EnableEvents = prevEvents
End Sub

Private Function IsLikelyProductionWorkbook(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then Exit Function
    wbName = LCase$(Trim$(wb.Name))
    If wbName Like "*.production.operator.xls*" Then
        IsLikelyProductionWorkbook = True
        Exit Function
    End If
    If WorkbookSheetExistsProductionInit(wb, "Production") _
       And WorkbookSheetExistsProductionInit(wb, "Recipes") _
       And WorkbookTableExistsProductionInit(wb, "RecipeBuilder") Then
        IsLikelyProductionWorkbook = True
    End If
End Function

Private Function WorkbookSheetExistsProductionInit(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorkbookSheetExistsProductionInit = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Function WorkbookTableExistsProductionInit(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        WorkbookTableExistsProductionInit = Not ws.ListObjects(tableName) Is Nothing
        On Error GoTo 0
        If WorkbookTableExistsProductionInit Then Exit Function
    Next ws
End Function
