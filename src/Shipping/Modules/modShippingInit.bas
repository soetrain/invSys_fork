Attribute VB_Name = "modShippingInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitShippingAddin()
    Dim report As String

    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If
    Call modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(ThisWorkbook, report)

    Application.EnableEvents = True
    SetupAllHandlers
End Sub

Public Sub Auto_Open()
    InitShippingAddin
End Sub

Public Sub EnsureShippingSurfaceForWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    modTS_Shipments.InitializeShipmentsUiForWorkbook wb
End Sub
