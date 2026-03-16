Attribute VB_Name = "modAdminInit"
Option Explicit

Public Sub InitAdminAddin()
    Dim report As String

    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(ThisWorkbook, report)
    Call modAdminConsole.EnsureAdminSchema(ThisWorkbook, report)
End Sub

Public Sub Auto_Open()
    InitAdminAddin
End Sub
