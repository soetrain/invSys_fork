Attribute VB_Name = "modShippingPostingService"
Option Explicit

Private Const EVENT_TYPE_SHIP_RELEASE As String = "SHIP_RELEASE"

Public Function ExecuteShipmentsSent(ByVal operatorWb As Workbook, _
                                     ByVal rowIndexes As Variant, _
                                     ByVal carrierValue As String, _
                                     ByRef report As String, _
                                     Optional ByVal skipAuthForTest As Boolean = False, _
                                     Optional ByVal ownerFacts As cShippingOwnerFacts = Nothing) As Boolean
    If Not CapturedWorkbookIsOpen(operatorWb) Then
        report = "The captured Shipping operator workbook is no longer open."
        Exit Function
    End If

    ExecuteShipmentsSent = modTS_Shipments.ShipmentsFormRunShipmentsSentRows( _
        rowIndexes, carrierValue, report, skipAuthForTest, operatorWb, ownerFacts)
End Function

Public Function ProjectedInventory(ByVal nasInventory As Double, _
                                   ByVal activeReservationQty As Double, _
                                   Optional ByVal appliedOverlayQty As Variant) As Double
    Dim projectedQty As Double

    If Not IsMissing(appliedOverlayQty) Then
        If IsNumeric(appliedOverlayQty) Then
            projectedQty = CDbl(appliedOverlayQty)
        Else
            projectedQty = nasInventory - activeReservationQty
        End If
    Else
        projectedQty = nasInventory - activeReservationQty
    End If
    If projectedQty < 0 Then projectedQty = 0
    ProjectedInventory = projectedQty
End Function

Public Function ReleaseExact(ByVal state As cShippingWorkflowState, _
                             ByVal shipmentLineId As String, _
                             ByVal reservationEventId As String, _
                             ByRef report As String) As Boolean
    If state Is Nothing Then
        report = "Shipping reservation state was not provided."
        Exit Function
    End If
    If Not state.ReleaseExact(shipmentLineId, reservationEventId) Then
        report = EVENT_TYPE_SHIP_RELEASE & _
                 " rejected: ShipmentLineId and ReservationEventId do not identify the exact active lock."
        Exit Function
    End If
    ReleaseExact = True
End Function

Public Function IdempotentReplayAccepted(ByVal originalEventId As String, _
                                         ByVal replayEventId As String, _
                                         ByVal processorStatus As String) As Boolean
    If Trim$(originalEventId) = "" Then Exit Function
    If StrComp(Trim$(originalEventId), Trim$(replayEventId), vbBinaryCompare) <> 0 Then Exit Function
    IdempotentReplayAccepted = _
        (StrComp(Trim$(processorStatus), "SKIP_DUP", vbTextCompare) = 0) _
        Or (StrComp(Trim$(processorStatus), "PROCESSED", vbTextCompare) = 0)
End Function

Public Function RestartCompletedLineMayRestore(ByVal completionTombstoneExists As Boolean, _
                                               ByVal persistedState As String) As Boolean
    If completionTombstoneExists Then Exit Function
    RestartCompletedLineMayRestore = _
        (StrComp(Trim$(persistedState), "COMPLETED", vbTextCompare) <> 0)
End Function

Public Function ClearCompletedStaging(ByVal completionTombstoneExists As Boolean, _
                                      ByVal persistedState As String) As Boolean
    ClearCompletedStaging = Not RestartCompletedLineMayRestore( _
        completionTombstoneExists, persistedState)
End Function

Private Function CapturedWorkbookIsOpen(ByVal operatorWb As Workbook) As Boolean
    On Error GoTo CleanExit

    Dim workbookName As String

    If operatorWb Is Nothing Then Exit Function
    If operatorWb.IsAddin Then Exit Function
    workbookName = operatorWb.Name
    CapturedWorkbookIsOpen = (Trim$(workbookName) <> "")

CleanExit:
End Function
