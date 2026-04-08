Attribute VB_Name = "modAdminWorkbookTarget"
Option Explicit

Public Function ResolveAdminTargetWorkbook(Optional ByVal explicitWb As Workbook = Nothing, _
                                           Optional ByVal fallbackWb As Workbook = Nothing, _
                                           Optional ByVal allowAddinFallback As Boolean = True) As Workbook
    Dim wb As Workbook

    If Not explicitWb Is Nothing Then
        Set ResolveAdminTargetWorkbook = explicitWb
        Exit Function
    End If

    On Error Resume Next
    Set wb = Application.ActiveWorkbook
    On Error GoTo 0

    If Not wb Is Nothing Then
        If Not wb.IsAddin Then
            Set ResolveAdminTargetWorkbook = wb
            Exit Function
        End If
    End If

    For Each wb In Application.Workbooks
        If Not wb Is Nothing Then
            If Not wb.IsAddin Then
                Set ResolveAdminTargetWorkbook = wb
                Exit Function
            End If
        End If
    Next wb

    If Not allowAddinFallback Then Exit Function

    If Not fallbackWb Is Nothing Then
        Set ResolveAdminTargetWorkbook = fallbackWb
    Else
        Set ResolveAdminTargetWorkbook = ThisWorkbook
    End If
End Function
