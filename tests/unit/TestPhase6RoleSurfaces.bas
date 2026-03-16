Attribute VB_Name = "TestPhase6RoleSurfaces"
Option Explicit

Public Function TestEnsureReceivingWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "ReceivedTally") _
       And HasTable(wb, "AggregateReceived") _
       And HasTable(wb, "ReceivedLog") _
       And HasTable(wb, "invSys") Then
        TestEnsureReceivingWorkbookSurface_CreatesExpectedTables = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureShippingWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "ShipmentsTally") _
       And HasTable(wb, "BoxBuilder") _
       And HasTable(wb, "BoxBOM") _
       And HasTable(wb, "AggregatePackages") _
       And HasTable(wb, "Check_invSys") _
       And HasTable(wb, "invSys") Then
        TestEnsureShippingWorkbookSurface_CreatesExpectedTables = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureProductionWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "RB_AddRecipeName") _
       And HasTable(wb, "RecipeBuilder") _
       And HasTable(wb, "RC_RecipeChoose") _
       And HasTable(wb, "ProductionOutput") _
       And HasTable(wb, "Prod_invSys_Check") _
       And HasTable(wb, "Recipes") _
       And HasTable(wb, "TemplatesTable") _
       And HasTable(wb, "ProductionLog") _
       And HasTable(wb, "invSys") Then
        TestEnsureProductionWorkbookSurface_CreatesExpectedTables = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureAdminWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wb, report) Then GoTo CleanExit
    If Not modAdminConsole.EnsureAdminSchema(wb, report) Then GoTo CleanExit

    If HasTable(wb, "UserCredentials") _
       And HasTable(wb, "Emails") _
       And HasTable(wb, "tblAdminAudit") _
       And HasTable(wb, "tblAdminPoisonQueue") Then
        TestEnsureAdminWorkbookSurface_CreatesExpectedTables = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function HasTable(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        On Error Resume Next
        HasTable = Not ws.ListObjects(tableName) Is Nothing
        On Error GoTo 0
        If HasTable Then Exit Function
    Next ws
End Function

Private Sub CloseNoSavePhase6(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub
