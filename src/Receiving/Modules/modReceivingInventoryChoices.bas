Attribute VB_Name = "modReceivingInventoryChoices"
Option Explicit

Public Function FromTable(ByVal inventoryTable As ListObject, Optional ByVal filterText As String = "", Optional ByVal codeHeader As String = "ITEM_CODE") As Variant
    Dim sourceValues As Variant
    Dim outputValues() As Variant
    Dim trimmedValues() As Variant
    Dim recordIndex As Long
    Dim fieldIndex As Long
    Dim outputIndex As Long
    Dim matchedIndex As Long
    Dim searchText As String
    Dim searchableText As String
    Dim systemKey As String
    Dim itemCode As String
    Dim itemName As String
    Dim uomValue As String
    Dim qtyValue As Double
    Dim qtyText As String
    Dim locationValue As String
    Dim conditionValue As String
    Dim lotNumber As String
    Dim descriptionValue As String
    Dim vendorValue As String
    Dim groupKey As String
    Dim groupIndex As Object

    If inventoryTable Is Nothing Or inventoryTable.DataBodyRange Is Nothing Then Exit Function
    If modTS_Received.ReadInventoryColumn(inventoryTable, "System_Key") = 0 _
       Or modTS_Received.ReadInventoryColumn(inventoryTable, codeHeader) = 0 Then Exit Function

    searchText = LCase$(Trim$(filterText))
    sourceValues = inventoryTable.DataBodyRange.Value2
    ReDim outputValues(1 To UBound(sourceValues, 1), 1 To 10)
    Set groupIndex = CreateObject("Scripting.Dictionary")
    groupIndex.CompareMode = vbTextCompare
    For recordIndex = 1 To UBound(sourceValues, 1)
        systemKey = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "System_Key")
        itemCode = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, codeHeader)
        itemName = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "ITEM")
        If itemName = "" Then itemName = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "ItemName")
        uomValue = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "UOM")
        qtyText = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "QtyAvailable")
        If qtyText = "" Then qtyText = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "TOTAL INV")
        If IsNumeric(qtyText) Then qtyValue = CDbl(qtyText) Else qtyValue = 0
        locationValue = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "LOCATION")
        lotNumber = modTS_Received.ReceivingInventoryLotNumber(inventoryTable, recordIndex)
        conditionValue = UCase$(modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "Condition"))
        If conditionValue = "" Then conditionValue = "GOOD"
        descriptionValue = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "DESCRIPTION")
        vendorValue = modTS_Received.ReadInventoryCell(inventoryTable, recordIndex, "VENDOR(s)")
        If systemKey = "" Or itemCode = "" Then GoTo NextInventoryRecord

        groupKey = UCase$(itemCode) & Chr$(30) & UCase$(uomValue) & Chr$(30) & _
                   UCase$(locationValue) & Chr$(30) & UCase$(lotNumber) & Chr$(30) & conditionValue
        If groupIndex.Exists(groupKey) Then
            outputValues(CLng(groupIndex(groupKey)), 5) = _
                CDbl(outputValues(CLng(groupIndex(groupKey)), 5)) + qtyValue
        Else
            outputIndex = outputIndex + 1
            groupIndex.Add groupKey, outputIndex
            outputValues(outputIndex, 1) = systemKey
            outputValues(outputIndex, 2) = itemCode
            outputValues(outputIndex, 3) = itemName
            outputValues(outputIndex, 4) = uomValue
            outputValues(outputIndex, 5) = qtyValue
            outputValues(outputIndex, 6) = locationValue
            outputValues(outputIndex, 7) = lotNumber
            outputValues(outputIndex, 8) = conditionValue
            outputValues(outputIndex, 9) = descriptionValue
            outputValues(outputIndex, 10) = vendorValue
        End If
NextInventoryRecord:
    Next recordIndex

    If outputIndex = 0 Then Exit Function

    For recordIndex = 1 To outputIndex
        If CDbl(outputValues(recordIndex, 5)) <= 0 Then GoTo NextGroupedCount
        searchableText = LCase$(CStr(outputValues(recordIndex, 2)) & " " & _
                                    CStr(outputValues(recordIndex, 3)) & " " & _
                                     CStr(outputValues(recordIndex, 6)) & " " & _
                                     CStr(outputValues(recordIndex, 7)) & " " & _
                                     CStr(outputValues(recordIndex, 8)) & " " & _
                                     CStr(outputValues(recordIndex, 9)) & " " & _
                                     CStr(outputValues(recordIndex, 10)))
        If searchText = "" Or InStr(1, searchableText, searchText, vbTextCompare) > 0 Then _
            matchedIndex = matchedIndex + 1
NextGroupedCount:
    Next recordIndex
    If matchedIndex = 0 Then Exit Function

    ReDim trimmedValues(1 To matchedIndex, 1 To 10)
    matchedIndex = 0
    For recordIndex = 1 To outputIndex
        If CDbl(outputValues(recordIndex, 5)) <= 0 Then GoTo NextGroupedCopy
        searchableText = LCase$(CStr(outputValues(recordIndex, 2)) & " " & _
                                    CStr(outputValues(recordIndex, 3)) & " " & _
                                     CStr(outputValues(recordIndex, 6)) & " " & _
                                     CStr(outputValues(recordIndex, 7)) & " " & _
                                     CStr(outputValues(recordIndex, 8)) & " " & _
                                     CStr(outputValues(recordIndex, 9)) & " " & _
                                     CStr(outputValues(recordIndex, 10)))
        If searchText = "" Or InStr(1, searchableText, searchText, vbTextCompare) > 0 Then
            matchedIndex = matchedIndex + 1
            For fieldIndex = 1 To 10
                trimmedValues(matchedIndex, fieldIndex) = outputValues(recordIndex, fieldIndex)
            Next fieldIndex
        End If
NextGroupedCopy:
    Next recordIndex
    FromTable = trimmedValues
End Function
