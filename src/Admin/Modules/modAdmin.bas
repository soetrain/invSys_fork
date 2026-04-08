Attribute VB_Name = "modAdmin"
Option Explicit

Sub Admin_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    Call modAdminConsole.OpenAdminConsole(targetWb, report)
End Sub

Sub Open_CreateDeleteUser()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    Call modAdminConsole.OpenUserManagement(targetWb, report)
End Sub

Sub Open_CreateWarehouse()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    frmCreateWarehouse.Show
End Sub

Sub Admin_SetupTesterStation_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    frmSetupTesterStation.Show
End Sub

Sub Open_SetupTesterStation()
    Admin_SetupTesterStation_Click
End Sub

Sub Verify_AddinsPublished()
    Dim report As String
    Dim detail As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modAddinsPublish.VerifyAddinsPublished() Then
        MsgBox "All required add-ins are published." & vbCrLf & modAddinsPublish.GetLastAddinsPublishReport(), vbInformation, "invSys Admin"
    Else
        detail = modAddinsPublish.GetLastAddinsPublishReport()
        If Len(detail) = 0 Then detail = "One or more required add-ins are missing or zero-byte."
        If InStr(1, detail, "PathSharePointRoot is not configured", vbTextCompare) > 0 Then
            detail = detail & vbCrLf & _
                     "Use Create New Warehouse or Setup Tester Station to choose the locally synced invSys SharePoint root first."
        End If
        MsgBox "Add-ins publish verification failed." & vbCrLf & detail, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Export_LoadedPackageReport()
    Dim report As String
    Dim pathOut As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modPackageDiagnostics.ExportLoadedPackageReport("", "", "", pathOut, report) Then
        MsgBox "Loaded package report written to:" & vbCrLf & pathOut, vbInformation, "invSys Admin"
    Else
        If Len(Trim$(report)) = 0 Then report = "Loaded package report export failed."
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Admin_RetireMigrateWarehouse_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    frmRetireMigrateWarehouse.Show
End Sub

Sub Open_RetireMigrateWarehouse()
    Admin_RetireMigrateWarehouse_Click
End Sub

Public Sub Scheduler_RunWarehouseBatch()
    PublishSchedulerResult modAdminConsole.RunScheduledWarehouseBatchForAutomation("", 0)
End Sub

Public Sub Scheduler_RunWarehousePublish()
    PublishSchedulerResult modAdminConsole.RunScheduledWarehousePublishForAutomation("", "")
End Sub

Public Sub Scheduler_RunHQAggregation()
    PublishSchedulerResult modAdminConsole.RunScheduledHQAggregationForAutomation("", "")
End Sub

Private Sub PublishSchedulerResult(ByVal resultText As String)
    Debug.Print resultText
    On Error Resume Next
    Application.StatusBar = resultText
    On Error GoTo 0
End Sub

Public Function ResolveInteractiveAdminWorkbook(Optional ByVal allowAddinFallback As Boolean = True) As Workbook
    Set ResolveInteractiveAdminWorkbook = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(Nothing, ThisWorkbook, allowAddinFallback)
End Function

''''''''''''''''''''''''''''''''''''
' This module contains administrative functions for the application.
' It includes functions to manage user accounts, roles, and permissions. yada yada
' It also includes functions to manage application settings and configurations.
' The functions in this module are used by the frmAdminControls form to perform administrative tasks.
''''''''''''''''''''''''''''''''''''







