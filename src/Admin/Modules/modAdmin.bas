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

''''''''''''''''''''''''''''''''''''
' This module contains administrative functions for the application.
' It includes functions to manage user accounts, roles, and permissions. yada yada
' It also includes functions to manage application settings and configurations.
' The functions in this module are used by the frmAdminControls form to perform administrative tasks.
''''''''''''''''''''''''''''''''''''







