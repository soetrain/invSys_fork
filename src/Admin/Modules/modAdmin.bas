Attribute VB_Name = "modAdmin"
Option Explicit

Sub Admin_Click()
    Dim report As String
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(ThisWorkbook, report)
    Call modAdminConsole.OpenAdminConsole(, report)
End Sub

Sub Open_CreateDeleteUser()
    Dim report As String
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(ThisWorkbook, report)
    Call modAdminConsole.OpenUserManagement(, report)
End Sub

Sub Open_CreateWarehouse()
    Dim report As String
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(ThisWorkbook, report)
    frmCreateWarehouse.Show
End Sub

Sub Admin_RetireMigrateWarehouse_Click()
    Dim report As String
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(ThisWorkbook, report)
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

''''''''''''''''''''''''''''''''''''
' This module contains administrative functions for the application.
' It includes functions to manage user accounts, roles, and permissions. yada yada
' It also includes functions to manage application settings and configurations.
' The functions in this module are used by the frmAdminControls form to perform administrative tasks.
''''''''''''''''''''''''''''''''''''







