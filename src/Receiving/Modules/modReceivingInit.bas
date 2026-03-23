Attribute VB_Name = "modReceivingInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitReceivingAddin()
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If
    modTS_Received.InitializeReceivingUiForWorkbook ThisWorkbook
    modTS_Received.InitializeReceivingUiForWorkbook Application.ActiveWorkbook
    Application.EnableEvents = True
End Sub

Public Sub Auto_Open()
    InitReceivingAddin
End Sub

Public Sub EnsureReceivingSurfaceForWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    modTS_Received.InitializeReceivingUiForWorkbook wb
End Sub
