Attribute VB_Name = "modShippingInit"
Option Explicit

Private gAppEvents As cAppEvents

Public Sub InitShippingAddin()
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cAppEvents
        gAppEvents.Init
    End If

    Application.EnableEvents = True
    modTS_Shipments.SetupAllHandlers
End Sub

Public Sub Auto_Open()
    InitShippingAddin
End Sub