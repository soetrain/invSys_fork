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
    If Not modRoleUiAccess.CanCurrentUserPerformCapability("SHIP_POST", "", "", "", errNotes) Then Exit Function

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

Public Function ValidateShipmentsSentStagingFromWorkbook(ByVal wb As Workbook) As String
    Dim invLo As ListObject
    Dim wsShip As Worksheet
    Dim deltas As Collection
    Dim errNotes As String

    If wb Is Nothing Then
        ValidateShipmentsSentStagingFromWorkbook = "Shipping workbook not provided."
        Exit Function
    End If

    Set invLo = GetInvSysTableShip(wb)
    Set wsShip = GetWorksheetShip(wb, "ShipmentsTally")
    If invLo Is Nothing Then
        ValidateShipmentsSentStagingFromWorkbook = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If wsShip Is Nothing Then
        ValidateShipmentsSentStagingFromWorkbook = "ShipmentsTally sheet not found."
        Exit Function
    End If

    If BuildQueueableShipmentsSentDeltas(invLo, wsShip, deltas, errNotes) Then
        ValidateShipmentsSentStagingFromWorkbook = "OK"
    Else
        If errNotes = "" Then errNotes = "No staged shipments found in invSys.SHIPMENTS."
        ValidateShipmentsSentStagingFromWorkbook = errNotes
    End If
End Function

Public Function ValidateToShipmentsFromWorkbook(ByVal wb As Workbook) As String
    Dim invLo As ListObject
    Dim aggPack As ListObject
    Dim deltas As Collection
    Dim errNotes As String

    If wb Is Nothing Then
        ValidateToShipmentsFromWorkbook = "Shipping workbook not provided."
        Exit Function
    End If

    Set invLo = GetInvSysTableShip(wb)
    Set aggPack = FindTableByNameShip(wb, "AggregatePackages")
    If invLo Is Nothing Then
        ValidateToShipmentsFromWorkbook = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If aggPack Is Nothing Then
        ValidateToShipmentsFromWorkbook = "AggregatePackages table not found."
        Exit Function
    End If
    If aggPack.DataBodyRange Is Nothing Then
        ValidateToShipmentsFromWorkbook = "AggregatePackages has no rows to stage."
        Exit Function
    End If

    Set deltas = BuildShipmentDeltaPacketShip(invLo, aggPack, errNotes)
    If deltas Is Nothing Then
        If errNotes = "" Then errNotes = "No additional shipments required; Shipments column already meets demand."
        ValidateToShipmentsFromWorkbook = errNotes
    ElseIf deltas.Count = 0 Then
        If errNotes = "" Then errNotes = "No additional shipments required; Shipments column already meets demand."
        ValidateToShipmentsFromWorkbook = errNotes
    Else
        ValidateToShipmentsFromWorkbook = "OK"
    End If
End Function

Public Function ValidateBoxesMadeFromWorkbook(ByVal wb As Workbook) As String
    Dim invLo As ListObject
    Dim aggBom As ListObject
    Dim shortageMsg As String

    If wb Is Nothing Then
        ValidateBoxesMadeFromWorkbook = "Shipping workbook not provided."
        Exit Function
    End If

    Set invLo = GetInvSysTableShip(wb)
    Set aggBom = FindTableByNameShip(wb, "AggregateBoxBOM")
    If invLo Is Nothing Then
        ValidateBoxesMadeFromWorkbook = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    If ValidateComponentInventoryShip(invLo, aggBom, shortageMsg) Then
        ValidateBoxesMadeFromWorkbook = "OK"
    Else
        ValidateBoxesMadeFromWorkbook = shortageMsg
    End If
End Function

Public Function ValidateConfirmInventoryFromWorkbook(ByVal wb As Workbook) As String
    Dim wsShip As Worksheet
    Dim shp As Shape

    If wb Is Nothing Then
        ValidateConfirmInventoryFromWorkbook = "Shipping workbook not provided."
        Exit Function
    End If
    Set wsShip = GetWorksheetShip(wb, "ShipmentsTally")
    If wsShip Is Nothing Then
        ValidateConfirmInventoryFromWorkbook = "ShipmentsTally sheet not found."
        Exit Function
    End If

    On Error Resume Next
    Set shp = wsShip.Shapes("CHK_USE_EXISTING")
    On Error GoTo 0
    If Not shp Is Nothing Then
        On Error Resume Next
        If shp.ControlFormat.Value = 1 Then
            ValidateConfirmInventoryFromWorkbook = "Use existing inventory is enabled. Skip Confirm inventory and go to 'To Shipments'."
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    ValidateConfirmInventoryFromWorkbook = "OK"
End Function

Public Function RebuildShippingAggregatesForWorkbook(ByVal wb As Workbook, Optional ByRef errNotes As String = "") As Boolean
    Dim loShip As ListObject
    Dim loBomView As ListObject
    Dim loAggBom As ListObject
    Dim packageQty As Object
    Dim componentQty As Object
    Dim componentInfo As Object
    Dim arrShip As Variant
    Dim arrBom As Variant
    Dim cShipRow As Long
    Dim cShipQty As Long
    Dim cShipArea As Long
    Dim cPkgRow As Long
    Dim cCompRow As Long
    Dim cCompItem As Long
    Dim cCompQty As Long
    Dim cCompUom As Long
    Dim cCompLoc As Long
    Dim r As Long
    Dim key As Variant
    Dim qty As Double
    Dim compRow As Long
    Dim info As Object
    Dim lr As ListRow

    If wb Is Nothing Then
        errNotes = "Shipping workbook not provided."
        Exit Function
    End If
    Set loShip = FindTableByNameShip(wb, "ShipmentsTally")
    Set loBomView = FindTableByNameShip(wb, "ShippingBOMView")
    Set loAggBom = FindTableByNameShip(wb, "AggregateBoxBOM")
    If loShip Is Nothing Or loBomView Is Nothing Or loAggBom Is Nothing Then
        errNotes = "Shipping aggregate source tables were not found."
        Exit Function
    End If
    ClearTableRowsShip loAggBom
    If loShip.DataBodyRange Is Nothing Or loBomView.DataBodyRange Is Nothing Then
        RebuildShippingAggregatesForWorkbook = True
        Exit Function
    End If

    cShipRow = ColumnIndexShip(loShip, "ROW")
    cShipQty = ColumnIndexShip(loShip, "QUANTITY")
    cShipArea = ColumnIndexShip(loShip, "AREA")
    If cShipRow = 0 Or cShipQty = 0 Then
        errNotes = "ShipmentsTally missing ROW/QUANTITY columns."
        Exit Function
    End If

    Set packageQty = CreateObject("Scripting.Dictionary")
    arrShip = loShip.DataBodyRange.Value
    For r = 1 To UBound(arrShip, 1)
        If cShipArea > 0 Then
            If StrComp(NzStrShip(arrShip(r, cShipArea)), "Hold", vbTextCompare) = 0 Then GoTo NextShip
        End If
        If NzLngShip(arrShip(r, cShipRow)) > 0 And NzDblShip(arrShip(r, cShipQty)) > 0 Then
            key = CStr(NzLngShip(arrShip(r, cShipRow)))
            If packageQty.Exists(key) Then
                packageQty(key) = NzDblShip(packageQty(key)) + NzDblShip(arrShip(r, cShipQty))
            Else
                packageQty.Add key, NzDblShip(arrShip(r, cShipQty))
            End If
        End If
NextShip:
    Next r
    If packageQty.Count = 0 Then
        RebuildShippingAggregatesForWorkbook = True
        Exit Function
    End If

    cPkgRow = ColumnIndexShip(loBomView, "PackageRow")
    cCompRow = ColumnIndexShip(loBomView, "ComponentRow")
    cCompItem = ColumnIndexShip(loBomView, "ComponentItem")
    cCompQty = ColumnIndexShip(loBomView, "ComponentQty")
    cCompUom = ColumnIndexShip(loBomView, "ComponentUOM")
    cCompLoc = ColumnIndexShip(loBomView, "ComponentLocation")
    If cPkgRow = 0 Or cCompRow = 0 Or cCompQty = 0 Then
        errNotes = "ShippingBOMView missing PackageRow/ComponentRow/ComponentQty columns."
        Exit Function
    End If

    Set componentQty = CreateObject("Scripting.Dictionary")
    Set componentInfo = CreateObject("Scripting.Dictionary")
    arrBom = loBomView.DataBodyRange.Value
    For r = 1 To UBound(arrBom, 1)
        key = CStr(NzLngShip(arrBom(r, cPkgRow)))
        If Not packageQty.Exists(key) Then GoTo NextBom
        compRow = NzLngShip(arrBom(r, cCompRow))
        qty = NzDblShip(arrBom(r, cCompQty)) * NzDblShip(packageQty(key))
        If compRow <= 0 Or qty <= 0 Then GoTo NextBom
        key = CStr(compRow)
        If componentQty.Exists(key) Then
            componentQty(key) = NzDblShip(componentQty(key)) + qty
        Else
            componentQty.Add key, qty
        End If
        If Not componentInfo.Exists(key) Then
            Set info = CreateObject("Scripting.Dictionary")
            If cCompItem > 0 Then
                info("ITEM") = NzStrShip(arrBom(r, cCompItem))
            Else
                info("ITEM") = ""
            End If
            If cCompUom > 0 Then
                info("UOM") = NzStrShip(arrBom(r, cCompUom))
            Else
                info("UOM") = ""
            End If
            If cCompLoc > 0 Then
                info("LOCATION") = NzStrShip(arrBom(r, cCompLoc))
            Else
                info("LOCATION") = ""
            End If
            componentInfo.Add key, info
        End If
NextBom:
    Next r

    For Each key In componentQty.Keys
        Set info = componentInfo(CStr(key))
        Set lr = loAggBom.ListRows.Add
        WriteTableCellShip loAggBom, lr.Index, "ROW", CLng(key)
        WriteTableCellShip loAggBom, lr.Index, "ITEM", info("ITEM")
        WriteTableCellShip loAggBom, lr.Index, "QUANTITY", NzDblShip(componentQty(key))
        WriteTableCellShip loAggBom, lr.Index, "UOM", info("UOM")
        WriteTableCellShip loAggBom, lr.Index, "LOCATION", info("LOCATION")
    Next key
    RebuildShippingAggregatesForWorkbook = True
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
        "", _
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

Private Function GetWorksheetShip(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set GetWorksheetShip = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindTableByNameShip(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTableByNameShip = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTableByNameShip Is Nothing Then Exit Function
    Next ws
End Function

Private Function FindInvListRowByRowValueShip(ByVal invLo As ListObject, ByVal rowValue As Long) As ListRow
    Dim cRow As Long
    Dim cel As Range

    If invLo Is Nothing Or rowValue <= 0 Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function
    cRow = ColumnIndexShip(invLo, "ROW")
    If cRow = 0 Then Exit Function
    For Each cel In invLo.ListColumns(cRow).DataBodyRange.Cells
        If NzLngShip(cel.Value) = rowValue Then
            Set FindInvListRowByRowValueShip = invLo.ListRows(cel.Row - invLo.DataBodyRange.Row + 1)
            Exit Function
        End If
    Next cel
End Function

Private Function BuildShipmentDeltaPacketShip(ByVal invLo As ListObject, ByVal aggPack As ListObject, ByRef errNotes As String) As Collection
    Dim cQtyAgg As Long
    Dim cRowAgg As Long
    Dim colTotalInv As Long
    Dim colShipments As Long
    Dim colRowInv As Long
    Dim colItemCode As Long
    Dim colItemName As Long
    Dim requirements As Object
    Dim arrAgg As Variant
    Dim r As Long
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim reqKeyStr As String
    Dim result As Collection
    Dim shipKey As Variant
    Dim invRow As ListRow
    Dim requiredQty As Double
    Dim alreadyStaged As Double
    Dim neededQty As Double
    Dim available As Double
    Dim delta As Object

    errNotes = ""
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    If aggPack Is Nothing Or aggPack.DataBodyRange Is Nothing Then Exit Function

    cQtyAgg = ColumnIndexShip(aggPack, "QUANTITY")
    cRowAgg = ColumnIndexShip(aggPack, "ROW")
    If cQtyAgg = 0 Or cRowAgg = 0 Then
        errNotes = "AggregatePackages missing QUANTITY/ROW columns."
        Exit Function
    End If

    colTotalInv = ColumnIndexShip(invLo, "TOTAL INV")
    colShipments = ColumnIndexShip(invLo, "SHIPMENTS")
    colRowInv = ColumnIndexShip(invLo, "ROW")
    colItemCode = ColumnIndexShip(invLo, "ITEM_CODE")
    colItemName = ColumnIndexShip(invLo, "ITEM")
    If colTotalInv = 0 Or colShipments = 0 Or colRowInv = 0 Then
        errNotes = "invSys table missing TOTAL INV/SHIPMENTS/ROW columns."
        Exit Function
    End If

    Set requirements = CreateObject("Scripting.Dictionary")
    arrAgg = aggPack.DataBodyRange.Value
    For r = 1 To UBound(arrAgg, 1)
        rowVal = NzLngShip(arrAgg(r, cRowAgg))
        qtyVal = NzDblShip(arrAgg(r, cQtyAgg))
        If rowVal = 0 Or qtyVal <= 0 Then GoTo NextAgg
        reqKeyStr = CStr(rowVal)
        If requirements.Exists(reqKeyStr) Then
            requirements(reqKeyStr) = NzDblShip(requirements(reqKeyStr)) + qtyVal
        Else
            requirements.Add reqKeyStr, qtyVal
        End If
NextAgg:
    Next r
    If requirements.Count = 0 Then Exit Function

    Set result = New Collection
    For Each shipKey In requirements.Keys
        Set invRow = FindInvListRowByRowValueShip(invLo, CLng(shipKey))
        If invRow Is Nothing Then
            AppendNoteShip errNotes, "Package ROW " & shipKey & " not found in invSys."
            Exit Function
        End If

        requiredQty = NzDblShip(requirements(shipKey))
        alreadyStaged = NzDblShip(invRow.Range.Cells(1, colShipments).Value)
        neededQty = requiredQty - alreadyStaged
        If neededQty <= 0 Then GoTo NextReq

        available = NzDblShip(invRow.Range.Cells(1, colTotalInv).Value)
        If neededQty > available + 0.0000001 Then
            AppendNoteShip errNotes, "ROW " & shipKey & " requires " & Format$(neededQty, "0.###") & " but only " & Format$(available, "0.###") & " in TOTAL INV."
            Exit Function
        End If

        Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = CLng(shipKey)
        delta("QTY") = neededQty
        If colItemCode > 0 Then delta("ITEM_CODE") = NzStrShip(invRow.Range.Cells(1, colItemCode).Value)
        If colItemName > 0 Then delta("ITEM_NAME") = NzStrShip(invRow.Range.Cells(1, colItemName).Value)
        result.Add delta
NextReq:
    Next shipKey

    If result.Count > 0 Then Set BuildShipmentDeltaPacketShip = result
End Function

Private Function ValidateComponentInventoryShip(ByVal invLo As ListObject, ByVal aggBom As ListObject, ByRef shortageMsg As String) As Boolean
    Dim cQty As Long
    Dim cRow As Long
    Dim arr As Variant
    Dim r As Long
    Dim rowVal As Long
    Dim requiredQty As Double
    Dim invRow As ListRow
    Dim colTotal As Long
    Dim colUsed As Long
    Dim available As Double

    shortageMsg = ""
    If invLo Is Nothing Then
        shortageMsg = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If aggBom Is Nothing Or aggBom.DataBodyRange Is Nothing Then
        shortageMsg = "AggregateBoxBOM has no component rows."
        Exit Function
    End If
    cQty = ColumnIndexShip(aggBom, "QUANTITY")
    cRow = ColumnIndexShip(aggBom, "ROW")
    If cQty = 0 Or cRow = 0 Then
        shortageMsg = "AggregateBoxBOM missing QUANTITY/ROW columns."
        Exit Function
    End If
    colTotal = ColumnIndexShip(invLo, "TOTAL INV")
    colUsed = ColumnIndexShip(invLo, "USED")
    If colTotal = 0 Then
        shortageMsg = "invSys table missing TOTAL INV column."
        Exit Function
    End If

    arr = aggBom.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        rowVal = NzLngShip(arr(r, cRow))
        requiredQty = NzDblShip(arr(r, cQty))
        If rowVal = 0 Or requiredQty <= 0 Then GoTo NextBom
        Set invRow = FindInvListRowByRowValueShip(invLo, rowVal)
        If invRow Is Nothing Then
            shortageMsg = "Component ROW " & rowVal & " not found in invSys."
            Exit Function
        End If
        available = NzDblShip(invRow.Range.Cells(1, colTotal).Value)
        If colUsed > 0 Then available = available - NzDblShip(invRow.Range.Cells(1, colUsed).Value)
        If requiredQty > available + 0.0000001 Then
            shortageMsg = "ROW " & rowVal & " requires " & Format$(requiredQty, "0.###") & " but only " & Format$(available, "0.###") & " available."
            Exit Function
        End If
NextBom:
    Next r
    ValidateComponentInventoryShip = True
End Function

Private Sub ClearTableRowsShip(ByVal lo As ListObject)
    On Error Resume Next
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    End If
    On Error GoTo 0
End Sub

Private Sub WriteTableCellShip(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    Dim colIndex As Long

    If lo Is Nothing Then Exit Sub
    colIndex = ColumnIndexShip(lo, columnName)
    If colIndex = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = valueIn
End Sub

Private Sub AppendNoteShip(ByRef target As String, ByVal text As String)
    If Len(text) = 0 Then Exit Sub
    If Len(target) > 0 Then
        target = target & vbCrLf & text
    Else
        target = text
    End If
End Sub

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
