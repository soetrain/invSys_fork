Attribute VB_Name = "modProductionRecallReport"
Option Explicit
Option Private Module

Public Sub Execute(ByVal owner As frmProduction, ByVal context As String, ByVal operatorBook As Workbook)
    If Not modProductionRunBinding.RequireWorksheetContext(owner, context, operatorBook) Then Exit Sub
    Dim outcome As String, detail As String
    mProduction.BtnPrintRecallCodes outcome, detail
    owner.ShowStatus detail
End Sub

Public Sub Preview(ByRef outcome As String, ByRef detail As String)
    On Error GoTo Failed
    outcome = "FAILED": detail = ""
    Dim wsReport As Worksheet, rowCount As Long
    If Not mProduction.BuildRecallCodesReportFromCurrentWorkbook(wsReport, rowCount, detail) Then
        outcome = "REJECTED"
        MsgBox detail, vbInformation
        Exit Sub
    End If
    wsReport.Activate
    wsReport.PrintOut Preview:=True
    outcome = "PREVIEW_RETURNED"
    detail = "Print preview closed."
    Exit Sub
Failed:
    detail = "BTN_PRINT_CODES failed: " & Err.Description
    MsgBox detail, vbCritical
End Sub

' D14: mutate only managed fields; workbook-local columns stay in their cells.
Public Function WriteRows(ByVal sheet As Worksheet, ByRef values As Variant, _
                          ByRef detail As String) As ListObject
    Dim table As ListObject, candidate As ListObject, otherSheet As Worksheet
    Dim target As Range, growth As Range, columns() As Long
    Dim rowCount As Long, fieldCount As Long, c As Long, r As Long
    If sheet Is Nothing Then GoTo Unsafe
    If sheet.ProtectContents Then GoTo Unsafe
    rowCount = UBound(values, 1) - 1: fieldCount = UBound(values, 2)
    If rowCount < 1 Then GoTo Unsafe
    For Each otherSheet In sheet.Parent.Worksheets
        For Each candidate In otherSheet.ListObjects
            If StrComp(candidate.Name, "RecallCodesReport", vbTextCompare) = 0 Then
                If Not (otherSheet Is sheet) Then GoTo Unsafe
                Set table = candidate
            ElseIf otherSheet Is sheet Then
                If Not Application.Intersect(candidate.Range, sheet.Range("A1:D3")) Is Nothing Then GoTo Unsafe
            End If
        Next candidate
    Next otherSheet

    If table Is Nothing Then
        Set target = sheet.Range("A5").Resize(rowCount + 1, fieldCount)
        If Not AreaAvailable(target, columns, True) Then GoTo Unsafe
        If Application.WorksheetFunction.CountA(sheet.Range("A1:D3")) > 0 Then GoTo Unsafe
    Else
        If table.ShowTotals Or table.Range.Row < 4 Then GoTo Unsafe
        If Not ResolveColumns(table, values, columns) Then GoTo Unsafe
        Set target = table.HeaderRowRange.Resize(rowCount + 1, table.ListColumns.Count)
        If target.Rows.Count > table.Range.Rows.Count Then
            Set growth = target.Rows(table.Range.Rows.Count + 1).Resize( _
                target.Rows.Count - table.Range.Rows.Count, target.Columns.Count)
            If Not AreaAvailable(growth, columns, False) Then GoTo Unsafe
        End If
    End If

    ' All predictable refusal checks precede the first worksheet write.
    If table Is Nothing Then
        For c = 1 To fieldCount
            target.Cells(1, c).Value = values(1, c)
        Next c
        Set table = sheet.ListObjects.Add(xlSrcRange, target, , xlYes)
        table.Name = "RecallCodesReport"
        table.TableStyle = "TableStyleMedium2"
        If Not ResolveColumns(table, values, columns) Then GoTo Unsafe
    Else
        For c = 1 To fieldCount
            If Not table.ListColumns(columns(c)).DataBodyRange Is Nothing Then _
                table.ListColumns(columns(c)).DataBodyRange.ClearContents
        Next c
        table.Resize target
    End If

    Dim columnValues() As Variant
    ReDim columnValues(1 To rowCount, 1 To 1)
    For c = 1 To fieldCount
        For r = 1 To rowCount
            columnValues(r, 1) = values(r + 1, c)
        Next r
        table.ListColumns(columns(c)).DataBodyRange.Value = columnValues
        table.ListColumns(columns(c)).Range.EntireColumn.AutoFit
    Next c
    Set WriteRows = table
    Exit Function
Unsafe:
    detail = "Recall report cannot be refreshed safely. Check its managed headers and available space."
End Function

Private Function ResolveColumns(ByVal table As ListObject, ByRef values As Variant, _
                                ByRef columns() As Long) As Boolean
    Dim c As Long, column As ListColumn, header As String
    ReDim columns(1 To UBound(values, 2))
    For c = 1 To UBound(values, 2)
        header = UCase$(Trim$(CStr(values(1, c))))
        For Each column In table.ListColumns
            If UCase$(Trim$(column.Name)) = header Then
                If columns(c) <> 0 Then Exit Function
                columns(c) = column.Index
            End If
        Next column
        If columns(c) = 0 Then Exit Function
    Next c
    ResolveColumns = True
End Function

Private Function AreaAvailable(ByVal area As Range, ByRef columns() As Long, ByVal inspectAll As Boolean) As Boolean
    Dim table As ListObject, c As Long
    If inspectAll Then
        If Application.WorksheetFunction.CountA(area) > 0 Then Exit Function
    Else
        For c = LBound(columns) To UBound(columns)
            If Application.WorksheetFunction.CountA(area.Columns(columns(c))) > 0 Then Exit Function
        Next c
    End If
    For Each table In area.Worksheet.ListObjects
        If Not Application.Intersect(table.Range, area) Is Nothing Then Exit Function
    Next table
    AreaAvailable = True
End Function
