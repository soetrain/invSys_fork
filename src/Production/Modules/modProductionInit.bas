Attribute VB_Name = "modProductionInit"
Option Explicit

Public Sub InitProductionAddin()
    Dim report As String

    Call modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(ThisWorkbook, report)
    Application.EnableEvents = True
    InitializeProductionUI
End Sub

Public Sub Auto_Open()
    InitProductionAddin
End Sub
