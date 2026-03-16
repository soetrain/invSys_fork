Attribute VB_Name = "modShippingInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitShippingAddin()
    Dim report As String

    Call modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(ThisWorkbook, report)
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If

    Application.EnableEvents = True
    SetupAllHandlers
End Sub

Public Sub Auto_Open()
    InitShippingAddin
End Sub
