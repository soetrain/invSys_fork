Attribute VB_Name = "modReceivingInit"
Option Explicit

Public Sub InitReceivingAddin()
    Dim report As String

    Call modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(ThisWorkbook, report)
    Application.EnableEvents = True
    EnsureGeneratedButtons
End Sub

Public Sub Auto_Open()
    InitReceivingAddin
End Sub
