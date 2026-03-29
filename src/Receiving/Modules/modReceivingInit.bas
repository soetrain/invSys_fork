Attribute VB_Name = "modReceivingInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitReceivingAddin()
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean

    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If
    modTS_Received.InitializeReceivingUiForWorkbook ThisWorkbook
    EnsureReceivingSurfaceForWorkbook Application.ActiveWorkbook
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEvents
End Sub

Public Sub Auto_Open()
    InitReceivingAddin
End Sub

Public Sub EnsureReceivingSurfaceForWorkbook(ByVal wb As Workbook)
    Dim prevEvents As Boolean

    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    modTS_Received.InitializeReceivingUiForWorkbook wb
    Application.EnableEvents = prevEvents
End Sub
