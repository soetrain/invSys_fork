Attribute VB_Name = "modOperationsFormLifetime"
Option Explicit

' Excel can unload a child before its owner during workbook-window closure.
' Compare exact loaded instances without invoking a disconnected child reference.
Public Function IsLoaded(ByVal candidate As Object) As Boolean
    Dim instance As Object
    If candidate Is Nothing Then Exit Function
    For Each instance In VBA.UserForms
        If instance Is candidate Then
            IsLoaded = True
            Exit Function
        End If
    Next instance
End Function
