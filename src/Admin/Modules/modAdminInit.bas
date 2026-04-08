Attribute VB_Name = "modAdminInit"
Option Explicit

Public Sub InitAdminAddin()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = modAdmin.ResolveInteractiveAdminWorkbook(False)
    If targetWb Is Nothing Then Exit Sub

    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    Call modAdminConsole.EnsureAdminSchema(targetWb, report)
End Sub

Public Sub Auto_Open()
    InitAdminAddin
End Sub
