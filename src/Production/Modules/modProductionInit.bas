Attribute VB_Name = "modProductionInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitProductionAddin()
    Dim prevEvents As Boolean

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If
    mProduction.InitializeProductionUiForWorkbook ThisWorkbook
    EnsureProductionSurfaceForWorkbook Application.ActiveWorkbook
    Application.EnableEvents = prevEvents
End Sub

Public Sub Auto_Open()
    InitProductionAddin
End Sub

Public Sub EnsureProductionSurfaceForWorkbook(ByVal wb As Workbook)
    Dim prevEvents As Boolean

    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    mProduction.InitializeProductionUiForWorkbook wb
    Application.EnableEvents = prevEvents
End Sub
