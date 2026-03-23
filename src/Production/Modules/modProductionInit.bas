Attribute VB_Name = "modProductionInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitProductionAddin()
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If
    mProduction.InitializeProductionUiForWorkbook ThisWorkbook
    EnsureProductionSurfaceForWorkbook Application.ActiveWorkbook
    Application.EnableEvents = True
End Sub

Public Sub Auto_Open()
    InitProductionAddin
End Sub

Public Sub EnsureProductionSurfaceForWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    mProduction.InitializeProductionUiForWorkbook wb
End Sub
