Attribute VB_Name = "modProductionEventCreator"
Option Explicit

Public Function QueueProductionCompleteEventFromWorkbook(ByVal wb As Workbook, _
                                                          Optional ByRef eventIdOut As String = "", _
                                                          Optional ByRef errNotes As String = "") As Boolean
    Dim wsProd As Worksheet
    Dim loOut As ListObject
    Dim outputItems As Collection

    If wb Is Nothing Then
        errNotes = "Production workbook not provided."
        Exit Function
    End If
    If Not modRoleUiAccess.CanCurrentUserPerformCapability( _
        "PROD_POST", "", "", "", errNotes) Then Exit Function

    On Error Resume Next
    Set wsProd = wb.Worksheets("Production")
    If Not wsProd Is Nothing Then Set loOut = wsProd.ListObjects("ProductionOutput")
    On Error GoTo 0
    If loOut Is Nothing Then
        errNotes = "ProductionOutput table not found on Production sheet."
        Exit Function
    End If

    Set outputItems = BuildOutputItems(loOut, errNotes)
    If outputItems Is Nothing Then Exit Function
    If outputItems.Count = 0 Then
        errNotes = "No made quantities found in ProductionOutput."
        Exit Function
    End If

    QueueProductionCompleteEventFromWorkbook = _
        modRoleEventWriter.QueuePayloadEventCurrent( _
            modInventoryDomainBridge.CORE_EVENT_TYPE_PROD_COMPLETE, _
            "", _
            modProductionJson.BuildJsonArray(outputItems), _
            "PRODUCTION_OUTPUT_COMPLETE", _
            eventIdOut, _
            errNotes)
End Function

Private Function BuildOutputItems(ByVal loOut As ListObject, _
                                  ByRef errNotes As String) As Collection
    Dim cSystemKey As Long
    Dim cSku As Long
    Dim cOutput As Long
    Dim cReal As Long
    Dim cLocation As Long
    Dim cCondition As Long
    Dim arr As Variant
    Dim result As Collection
    Dim item As Object
    Dim systemKey As String
    Dim sku As String
    Dim qty As Double
    Dim r As Long

    errNotes = ""
    If loOut Is Nothing Or loOut.DataBodyRange Is Nothing Then Exit Function
    cSystemKey = ColumnIndexProd(loOut, "System_Key")
    cSku = ColumnIndexProd(loOut, "ITEM_CODE")
    If cSku = 0 Then cSku = ColumnIndexProd(loOut, "SKU")
    cOutput = ColumnIndexProd(loOut, "OUTPUT")
    cReal = ColumnIndexProd(loOut, "REAL OUTPUT")
    cLocation = ColumnIndexProd(loOut, "LOCATION")
    cCondition = ColumnIndexProd(loOut, "Condition")
    If cSystemKey = 0 Or cReal = 0 Then
        errNotes = "ProductionOutput requires System_Key and REAL OUTPUT columns."
        Exit Function
    End If

    Set result = New Collection
    arr = loOut.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        qty = NumberValue(arr(r, cReal))
        If qty <= 0 Then GoTo NextOutput

        systemKey = Trim$(TextValue(arr(r, cSystemKey)))
        If systemKey = "" Then
            systemKey = modRoleEventWriter.CreateSystemKey()
            loOut.DataBodyRange.Cells(r, cSystemKey).Value2 = systemKey
        End If
        sku = ""
        If cSku > 0 Then sku = Trim$(TextValue(arr(r, cSku)))
        If sku = "" And cOutput > 0 Then sku = Trim$(TextValue(arr(r, cOutput)))
        If sku = "" Then
            errNotes = "Production output with System_Key '" & systemKey & _
                       "' has no ITEM_CODE/SKU."
            Exit Function
        End If

        Set item = modProductionJson.CreateProductionInventoryEntityPayloadItem( _
            systemKey, sku, qty, _
            ColumnText(arr, r, cLocation), _
            ColumnText(arr, r, cCondition))
        item("IoType") = "MADE"
        result.Add item
NextOutput:
    Next r
    Set BuildOutputItems = result
End Function

Private Function ColumnIndexProd(ByVal lo As ListObject, _
                                 ByVal columnName As String) As Long
    Dim listColumn As ListColumn

    For Each listColumn In lo.ListColumns
        If StrComp(Trim$(listColumn.Name), Trim$(columnName), vbTextCompare) = 0 Then
            ColumnIndexProd = listColumn.Index
            Exit Function
        End If
    Next listColumn
End Function

Private Function ColumnText(ByVal values As Variant, _
                            ByVal rowIndex As Long, _
                            ByVal columnIndex As Long) As String
    If columnIndex <= 0 Then Exit Function
    ColumnText = TextValue(values(rowIndex, columnIndex))
End Function

Private Function TextValue(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    TextValue = CStr(valueIn)
End Function

Private Function NumberValue(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If Trim$(CStr(valueIn)) = "" Then Exit Function
    If IsNumeric(valueIn) Then NumberValue = CDbl(valueIn)
End Function
