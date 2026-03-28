Attribute VB_Name = "modInventoryInit"
Option Explicit

Private gAppEvents As cInventoryAppEvents

Public Sub InitInventoryDomainAddin()
    Dim report As String

    If gAppEvents Is Nothing Then
        Set gAppEvents = New cInventoryAppEvents
        gAppEvents.Init
    End If
    Call modInventoryPublisher.PublishOpenInventorySnapshots(report)
End Sub

Public Sub Auto_Open()
    InitInventoryDomainAddin
End Sub
