Attribute VB_Name = "modAdminInit"
Option Explicit

Public Sub InitAdminAddin()
    Dim report As String
    Dim targetWb As Workbook

    ApplyRememberedRuntimeTargetAdmin
    Set targetWb = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(Nothing, ThisWorkbook, False)
    If targetWb Is Nothing Then Exit Sub

    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    Call modAdminConsole.EnsureAdminSchema(targetWb, report)
End Sub

Private Sub ApplyRememberedRuntimeTargetAdmin()
    On Error Resume Next
    Call modRibbonRuntimeStatus.TryApplyRememberedWarehouseTarget
    On Error GoTo 0
End Sub

Public Sub Auto_Open()
    InitAdminAddin
End Sub
