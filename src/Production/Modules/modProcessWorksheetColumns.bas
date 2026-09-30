Attribute VB_Name = "modProcessWorksheetColumns"
Option Explicit

' D14: the initial layout never determines the address of a managed field.
Public Function Field(ByVal table As ListObject, ByVal header As String) As ListColumn
    Dim column As ListColumn
    If table Is Nothing Then Exit Function
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), Trim$(header), vbTextCompare) = 0 Then
            If Not Field Is Nothing Then
                Set Field = Nothing
                Exit Function
            End If
            Set Field = column
        End If
    Next column
End Function

Public Function Cell(ByVal table As ListObject, ByVal rowIndex As Long, _
                     ByVal header As String) As Range
    Set Cell = Field(table, header).DataBodyRange.Cells(rowIndex, 1)
End Function

Public Function Text(ByVal table As ListObject, ByVal rowIndex As Long, _
                     ByVal header As String) As String
    Dim column As ListColumn, value As Variant
    Set column = Field(table, header)
    If column Is Nothing Then Exit Function
    value = column.DataBodyRange.Cells(rowIndex, 1).Value2
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    Text = CStr(value)
End Function

' Bind structured references without renaming/reordering the user's columns.
' Supported header normalization changes case and surrounding spaces only.
Public Function FormulaText(ByVal table As ListObject, ByVal canonical As String) As String
    Dim header As Variant, actual As String
    FormulaText = canonical
    For Each header In Array("Record Type", "ID", "Qty", "Basis Qty", "UOM", "Qty Mode")
        actual = Field(table, CStr(header)).Name
        FormulaText = Replace$(FormulaText, "[@" & CStr(header) & "]", "[@[" & actual & "]]")
        FormulaText = Replace$(FormulaText, "[" & CStr(header) & "]", "[" & actual & "]")
    Next header
End Function

Public Sub SetFormula(ByVal table As ListObject, ByVal header As String, ByVal canonical As String)
    Field(table, header).DataBodyRange.Formula = FormulaText(table, canonical)
End Sub

Public Sub ApplyTextIdentityFormats(ByVal table As ListObject, ByVal firstPair As Long, _
                                    ByVal pairCount As Long, ByVal headerOffset As Long)
    Dim pairNumber As Long
    If table Is Nothing Then Exit Sub
    If Not table.DataBodyRange Is Nothing Then
        Field(table, "ID").DataBodyRange.NumberFormat = "@"
        Field(table, "Output SKU").DataBodyRange.NumberFormat = "@"
        For pairNumber = firstPair To pairCount
            Field(table, "Accepted SKU " & CStr(pairNumber)).DataBodyRange.NumberFormat = "@"
        Next pairNumber
    End If
    table.Parent.Cells(table.HeaderRowRange.Row - headerOffset + 1, 5).NumberFormat = "@"
End Sub
