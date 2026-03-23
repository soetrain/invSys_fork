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
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    modTS_Shipments.InitializeShipmentsUiForWorkbook wb
    Application.EnableEvents = prevEvents
End Sub
