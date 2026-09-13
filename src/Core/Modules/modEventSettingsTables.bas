Attribute VB_Name = "modEventSettingsTables"
Option Explicit
Option Private Module

' Low-level table operations used only by authorized Core settings commands.
Public Function CreateTable(ByVal wb As Workbook, ByVal name As String, ByVal fields As Variant, _
                            ByVal created As Collection) As ListObject
    Dim sheet As Worksheet, table As ListObject, index As Long
    Set sheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    created.Add sheet
    For index = LBound(fields) To UBound(fields)
        sheet.Cells(1, index + 1).Value2 = fields(index)
    Next index
    Set table = sheet.ListObjects.Add(xlSrcRange, sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, UBound(fields) + 1)), , xlYes)
    table.Name = name
    Do While table.ListRows.Count > 0
        table.ListRows(table.ListRows.Count).Delete
    Loop
    Set CreateTable = table
End Function

Public Sub WriteCell(ByVal table As ListObject, ByVal row As ListRow, ByVal header As String, ByVal value As Variant)
    Dim cell As Range
    Set cell = row.Range.Cells(1, modTrackingPolicyModel.TableColumn(table, header))
    If VarType(value) = vbString Then cell.NumberFormat = "@"
    cell.Value2 = value
End Sub
