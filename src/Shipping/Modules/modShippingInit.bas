Attribute VB_Name = "modShippingInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitShippingAddin()
    Dim report As String
    Dim prevEvents As Boolean

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If
    Call modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(ThisWorkbook, report)

    Application.EnableEvents = prevEvents
    SetupAllHandlers
End Sub

Public Sub Auto_Open()
    InitShippingAddin
End Sub

Public Sub EnsureShippingSurfaceForWorkbook(ByVal wb As Workbook)
    Dim prevEvents As Boolean

    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    If Not IsLikelyShippingWorkbook(wb) Then Exit Sub
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    modTS_Shipments.InitializeShipmentsUiForWorkbook wb
    Application.EnableEvents = prevEvents
End Sub

Private Function IsLikelyShippingWorkbook(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then Exit Function
    wbName = LCase$(Trim$(wb.Name))
    If wbName Like "*.shipping.operator.xls*" Then
        IsLikelyShippingWorkbook = True
        Exit Function
    End If
    If WorkbookSheetExistsShippingInit(wb, "ShipmentsTally") _
       And WorkbookTableExistsShippingInit(wb, "ShipmentsTally") Then
        IsLikelyShippingWorkbook = True
    End If
End Function

Private Function WorkbookSheetExistsShippingInit(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorkbookSheetExistsShippingInit = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Function WorkbookTableExistsShippingInit(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        WorkbookTableExistsShippingInit = Not ws.ListObjects(tableName) Is Nothing
        On Error GoTo 0
        If WorkbookTableExistsShippingInit Then Exit Function
    Next ws
End Function
