Attribute VB_Name = "modProductionEventCreator"
Option Explicit

Public Function QueueProductionCompleteEventFromWorkbook(ByVal wb As Workbook, Optional ByRef eventIdOut As String = "", Optional ByRef errNotes As String = "") As Boolean
    Dim wsProd As Worksheet
    Dim invLo As ListObject
    Dim loOut As ListObject
    Dim madeDeltas As Collection

    If wb Is Nothing Then
        errNotes = "Production workbook not provided."
        Exit Function
    End If
    If Not modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST", "", "", "", errNotes) Then Exit Function

    Set invLo = GetInvSysTableProd(wb)
    If invLo Is Nothing Then
        errNotes = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    On Error Resume Next
    Set wsProd = wb.Worksheets("Production")
    If Not wsProd Is Nothing Then Set loOut = FindListObjectByNameOrHeadersProd(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    On Error GoTo 0
    If loOut Is Nothing Then
        errNotes = "ProductionOutput table not found on Production sheet."
        Exit Function
    End If

    Set madeDeltas = BuildMadeDeltasFromProductionOutput(loOut, invLo, errNotes)
    If madeDeltas Is Nothing Then
        If errNotes = "" Then errNotes = "No made quantities found in ProductionOutput."
        Exit Function
    End If
    If madeDeltas.Count = 0 Then
        If errNotes = "" Then errNotes = "No made quantities found in ProductionOutput."
        Exit Function
    End If

    QueueProductionCompleteEventFromWorkbook = QueueProductionCompleteEventCore(madeDeltas, errNotes, eventIdOut)
End Function

Private Function QueueProductionCompleteEventCore(ByVal madeDeltas As Collection, ByRef errNotes As String, ByRef eventIdOut As String) As Boolean
    Dim payloadItems As Collection
    Dim delta As Variant

    Set payloadItems = New Collection
    For Each delta In madeDeltas
        payloadItems.Add modRoleEventWriter.CreatePayloadItem( _
            NzLngProd(delta("ROW")), _
            NzStrProd(delta("ITEM_CODE")), _
            NzDblProd(delta("QTY")), _
            "", _
            NzStrProd(delta("ITEM_NAME")), _
            "MADE")
    Next delta

    QueueProductionCompleteEventCore = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_PROD_COMPLETE, _
        modRoleEventWriter.ResolveCurrentUserId(), _
        modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems), _
        "BTN_TO_TOTALINV", _
        eventIdOut, _
        errNotes)
End Function

Private Function BuildMadeDeltasFromProductionOutput(ByVal loOut As ListObject, ByVal invLo As ListObject, ByRef errNotes As String) As Collection
    Dim cReal As Long
    Dim cOutput As Long
    Dim cRowOut As Long
    Dim cItemCode As Long
    Dim cItemName As Long
    Dim rowIndex As Object
    Dim outputLookup As Object
    Dim agg As Object
    Dim arr As Variant
    Dim r As Long
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim key As String
    Dim delta As Object
    Dim result As Collection
    Dim k As Variant

    errNotes = ""
    If loOut Is Nothing Or loOut.DataBodyRange Is Nothing Then Exit Function
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then
        errNotes = "invSys table not found."
        Exit Function
    End If

    cReal = ColumnIndexProd(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLooseProd(loOut, "REALOUTPUT", "REAL_OUTPUT")
    cOutput = ColumnIndexProd(loOut, "OUTPUT")
    cRowOut = ColumnIndexProd(loOut, "ROW")
    If cRowOut = 0 Then cRowOut = ColumnIndexLooseProd(loOut, "ROW", "ROWID", "ROW#")
    If cReal = 0 Then
        errNotes = "ProductionOutput missing REAL OUTPUT column."
        Exit Function
    End If
    If cRowOut = 0 And cOutput = 0 Then
        errNotes = "ProductionOutput missing ROW/OUTPUT columns."
        Exit Function
    End If

    Set rowIndex = BuildInvSysRowIndexProd(invLo)
    If rowIndex Is Nothing Then
        errNotes = "invSys ROW index not available."
        Exit Function
    End If
    If rowIndex.Count = 0 Then
        errNotes = "invSys ROW index not available."
        Exit Function
    End If
    Set outputLookup = BuildInvSysOutputLookupProd(invLo)
    cItemCode = ColumnIndexProd(invLo, "ITEM_CODE")
    cItemName = ColumnIndexProd(invLo, "ITEM")

    Set agg = CreateObject("Scripting.Dictionary")
    arr = loOut.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        qtyVal = NzDblProd(arr(r, cReal))
        If qtyVal <= 0 Then GoTo NextRow

        If cRowOut > 0 Then rowVal = NzLngProd(arr(r, cRowOut))
        If rowVal = 0 And cOutput > 0 Then rowVal = LookupOutputRowProd(outputLookup, NzStrProd(arr(r, cOutput)))
        If rowVal = 0 Then GoTo NextRow
        If Not rowIndex.Exists(CStr(rowVal)) Then GoTo NextRow

        key = CStr(rowVal)
        If agg.Exists(key) Then
            agg(key)("QTY") = NzDblProd(agg(key)("QTY")) + qtyVal
        Else
            Set delta = CreateObject("Scripting.Dictionary")
            delta("ROW") = rowVal
            delta("QTY") = qtyVal
            If cItemCode > 0 Then delta("ITEM_CODE") = NzStrProd(invLo.DataBodyRange.Cells(CLng(rowIndex(key)), cItemCode).Value)
            If cItemName > 0 Then delta("ITEM_NAME") = NzStrProd(invLo.DataBodyRange.Cells(CLng(rowIndex(key)), cItemName).Value)
            agg.Add key, delta
        End If
NextRow:
        rowVal = 0
    Next r

    If agg.Count = 0 Then
        errNotes = "No made quantities found in ProductionOutput."
        Exit Function
    End If

    Set result = New Collection
    For Each k In agg.Keys
        result.Add agg(k)
    Next k
    Set BuildMadeDeltasFromProductionOutput = result
End Function

Private Function GetInvSysTableProd(ByVal wb As Workbook) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("InventoryManagement")
    If Not ws Is Nothing Then Set GetInvSysTableProd = ws.ListObjects("invSys")
    On Error GoTo 0
End Function

Private Function FindListObjectByNameOrHeadersProd(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not lo Is Nothing Then
        Set FindListObjectByNameOrHeadersProd = lo
        Exit Function
    End If
    For Each lo In ws.ListObjects
        If ListObjectHasHeadersProd(lo, headers) Then
            Set FindListObjectByNameOrHeadersProd = lo
            Exit Function
        End If
    Next lo
End Function

Private Function ListObjectHasHeadersProd(ByVal lo As ListObject, ByVal headers As Variant) As Boolean
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If ColumnIndexProd(lo, CStr(headers(i))) = 0 Then Exit Function
    Next i
    ListObjectHasHeadersProd = True
End Function

Private Function ColumnIndexProd(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(colName), vbTextCompare) = 0 Then
            ColumnIndexProd = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function ColumnIndexLooseProd(ByVal lo As ListObject, ParamArray names() As Variant) As Long
    Dim lc As ListColumn
    Dim hdr As String
    Dim i As Long
    For Each lc In lo.ListColumns
        hdr = NormalizeHeaderKeyProd(NzStrProd(lc.Name))
        For i = LBound(names) To UBound(names)
            If hdr = NormalizeHeaderKeyProd(CStr(names(i))) Then
                ColumnIndexLooseProd = lc.Index
                Exit Function
            End If
        Next i
    Next lc
End Function

Private Function NormalizeHeaderKeyProd(ByVal valueIn As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(valueIn)
        ch = Mid$(valueIn, i, 1)
        If ch Like "[A-Za-z0-9]" Then NormalizeHeaderKeyProd = NormalizeHeaderKeyProd & UCase$(ch)
    Next i
End Function

Private Function BuildInvSysRowIndexProd(ByVal invLo As ListObject) As Object
    Dim dict As Object
    Dim cRow As Long
    Dim arr As Variant
    Dim r As Long
    Dim rowVal As Long

    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    cRow = ColumnIndexProd(invLo, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLooseProd(invLo, "ROW", "ROWID", "ROW#")
    If cRow = 0 Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    arr = invLo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        rowVal = NzLngProd(arr(r, cRow))
        If rowVal <> 0 Then dict(CStr(rowVal)) = r
    Next r
    Set BuildInvSysRowIndexProd = dict
End Function

Private Function BuildInvSysOutputLookupProd(ByVal invLo As ListObject) As Object
    Dim dict As Object
    Dim cRow As Long
    Dim cItem As Long
    Dim cCode As Long
    Dim arr As Variant
    Dim r As Long

    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    cRow = ColumnIndexProd(invLo, "ROW")
    cItem = ColumnIndexProd(invLo, "ITEM")
    cCode = ColumnIndexProd(invLo, "ITEM_CODE")
    If cRow = 0 Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    arr = invLo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        If cItem > 0 Then AddLookupProd dict, NzStrProd(arr(r, cItem)), NzLngProd(arr(r, cRow))
        If cCode > 0 Then AddLookupProd dict, NzStrProd(arr(r, cCode)), NzLngProd(arr(r, cRow))
    Next r
    Set BuildInvSysOutputLookupProd = dict
End Function

Private Sub AddLookupProd(ByVal dict As Object, ByVal keyText As String, ByVal rowVal As Long)
    Dim norm As String
    norm = modItemSearch.NormalizeSearchText(keyText)
    If norm = "" Or rowVal = 0 Then Exit Sub
    If Not dict.Exists(norm) Then dict.Add norm, rowVal
End Sub

Private Function LookupOutputRowProd(ByVal outputLookup As Object, ByVal outputName As String) As Long
    Dim keyText As String
    keyText = modItemSearch.NormalizeSearchText(outputName)
    If keyText = "" Then Exit Function
    If outputLookup.Exists(keyText) Then LookupOutputRowProd = CLng(outputLookup(keyText))
End Function

Private Function NzStrProd(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    NzStrProd = CStr(valueIn)
End Function

Private Function NzDblProd(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzDblProd = CDbl(valueIn)
End Function

Private Function NzLngProd(ByVal valueIn As Variant) As Long
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzLngProd = CLng(valueIn)
End Function
