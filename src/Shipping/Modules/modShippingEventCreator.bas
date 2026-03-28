Attribute VB_Name = "modShippingEventCreator"
Option Explicit

Public Function QueueShipmentsSentEventFromWorkbook(ByVal wb As Workbook, Optional ByRef eventIdOut As String = "", Optional ByRef errNotes As String = "") As Boolean
    Dim invLo As ListObject
    Dim wsShip As Worksheet
    Dim deltas As Collection

    If wb Is Nothing Then
        errNotes = "Shipping workbook not provided."
        Exit Function
    End If
    If Not CanCurrentUserPerformCapability("SHIP_POST", "", "", "", errNotes) Then Exit Function

    Set invLo = GetInvSysTableShip(wb)
    If invLo Is Nothing Then
        errNotes = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    On Error Resume Next
    Set wsShip = wb.Worksheets("ShipmentsTally")
    On Error GoTo 0
    If wsShip Is Nothing Then
        errNotes = "ShipmentsTally sheet not found."
        Exit Function
    End If

    If Not BuildQueueableShipmentsSentDeltas(invLo, wsShip, deltas, errNotes) Then Exit Function
    QueueShipmentsSentEventFromWorkbook = QueueShipmentsSentEventCore(deltas, errNotes, eventIdOut)
End Function

Private Function QueueShipmentsSentEventCore(ByVal deltas As Collection, ByRef errNotes As String, ByRef eventIdOut As String) As Boolean
    Dim payloadJson As String

    payloadJson = BuildPayloadJsonFromDeltasShip(deltas)
    If payloadJson = "" Then
        If errNotes = "" Then errNotes = "No shipment payload rows were generated."
        Exit Function
    End If

    QueueShipmentsSentEventCore = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_SHIP, _
        modRoleEventWriter.ResolveCurrentUserId(), _
        payloadJson, _
        "BTN_SHIPMENTS_SENT", _
        eventIdOut, _
        errNotes)
End Function

Private Function BuildQueueableShipmentsSentDeltas(ByVal invLo As ListObject, ByVal ws As Worksheet, ByRef deltasOut As Collection, ByRef errNotes As String) As Boolean
    Dim aggPack As ListObject
    Dim rowFilter As Object
    Dim arrAgg As Variant
    Dim cRowAgg As Long
    Dim r As Long
    Dim filtered As Collection
    Dim delta As Variant

    Set deltasOut = BuildShipmentsSentDeltaPacket(invLo, errNotes)
    If deltasOut Is Nothing Then
        If errNotes = "" Then errNotes = "No staged shipments found in invSys.SHIPMENTS."
        Exit Function
    End If
    If deltasOut.Count = 0 Then
        If errNotes = "" Then errNotes = "No staged shipments found in invSys.SHIPMENTS."
        Exit Function
    End If

    On Error Resume Next
    Set aggPack = ws.ListObjects("AggregatePackages")
    On Error GoTo 0
    If Not aggPack Is Nothing Then
        If Not aggPack.DataBodyRange Is Nothing Then
            cRowAgg = ColumnIndexShip(aggPack, "ROW")
            If cRowAgg > 0 Then
                Set rowFilter = CreateObject("Scripting.Dictionary")
                arrAgg = aggPack.DataBodyRange.Value
                For r = 1 To UBound(arrAgg, 1)
                    If NzLngShip(arrAgg(r, cRowAgg)) > 0 Then rowFilter(CStr(NzLngShip(arrAgg(r, cRowAgg)))) = True
                Next r
            End If
        End If
    End If

    If Not rowFilter Is Nothing Then
        If rowFilter.Count > 0 Then
            Set filtered = New Collection
            For Each delta In deltasOut
                If rowFilter.Exists(CStr(delta("ROW"))) Then filtered.Add delta
            Next delta
            Set deltasOut = filtered
            If deltasOut.Count = 0 Then
                errNotes = "No staged shipments match the current AggregatePackages rows."
                Exit Function
            End If
        End If
    End If

    BuildQueueableShipmentsSentDeltas = True
End Function

Private Function BuildShipmentsSentDeltaPacket(ByVal invLo As ListObject, ByRef errNotes As String) As Collection
    Dim cShip As Long
    Dim cRow As Long
    Dim cItemCode As Long
    Dim cItemName As Long
    Dim result As Collection
    Dim arr As Variant
    Dim r As Long
    Dim delta As Object

    errNotes = ""
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function

    cShip = ColumnIndexShip(invLo, "SHIPMENTS")
    cRow = ColumnIndexShip(invLo, "ROW")
    cItemCode = ColumnIndexShip(invLo, "ITEM_CODE")
    cItemName = ColumnIndexShip(invLo, "ITEM")
    If cShip = 0 Or cRow = 0 Then
        errNotes = "invSys table missing SHIPMENTS/ROW columns."
        Exit Function
    End If

    Set result = New Collection
    arr = invLo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        If NzLngShip(arr(r, cRow)) = 0 Or NzDblShip(arr(r, cShip)) <= 0 Then GoTo NextRow
        Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = NzLngShip(arr(r, cRow))
        delta("QTY") = NzDblShip(arr(r, cShip))
        If cItemCode > 0 Then delta("ITEM_CODE") = NzStrShip(arr(r, cItemCode))
        If cItemName > 0 Then delta("ITEM_NAME") = NzStrShip(arr(r, cItemName))
        result.Add delta
NextRow:
    Next r

    If result.Count = 0 Then
        errNotes = "No staged shipments found in invSys.SHIPMENTS."
        Exit Function
    End If
    Set BuildShipmentsSentDeltaPacket = result
End Function

Private Function BuildPayloadJsonFromDeltasShip(ByVal deltas As Collection) As String
    Dim payloadItems As New Collection
    Dim delta As Variant

    If deltas Is Nothing Then Exit Function
    For Each delta In deltas
        payloadItems.Add modRoleEventWriter.CreatePayloadItem( _
            NzLngShip(delta("ROW")), _
            NzStrShip(delta("ITEM_CODE")), _
            NzDblShip(delta("QTY")), _
            "", _
            NzStrShip(delta("ITEM_NAME")))
    Next delta
    BuildPayloadJsonFromDeltasShip = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
End Function

Private Function GetInvSysTableShip(ByVal wb As Workbook) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("InventoryManagement")
    If Not ws Is Nothing Then Set GetInvSysTableShip = ws.ListObjects("invSys")
    On Error GoTo 0
End Function

Private Function ColumnIndexShip(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(colName), vbTextCompare) = 0 Then
            ColumnIndexShip = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function NzStrShip(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    NzStrShip = CStr(valueIn)
End Function

Private Function NzDblShip(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzDblShip = CDbl(valueIn)
End Function

Private Function NzLngShip(ByVal valueIn As Variant) As Long
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzLngShip = CLng(valueIn)
End Function
