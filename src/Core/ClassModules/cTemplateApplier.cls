VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTemplateApplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SHEET_TEMPLATES As String = "TemplatesTable"
Private Const TABLE_TEMPLATES As String = "TemplatesTable"

Public Sub ApplyTemplates(ByVal targetLo As ListObject, ByVal templateScope As String, _
    Optional ByVal processKey As String = "", Optional ByVal tableNameOverride As String = "", _
    Optional ByVal recipeId As String = "", Optional ByVal ingredientId As String = "")

    If targetLo Is Nothing Then Exit Sub
    Dim templatesLo As ListObject
    Set templatesLo = GetTemplatesTable()
    If templatesLo Is Nothing Then Exit Sub
    If templatesLo.DataBodyRange Is Nothing Then Exit Sub

    Dim cScope As Long: cScope = ColumnIndex(templatesLo, "TEMPLATE_SCOPE")
    Dim cRecipe As Long: cRecipe = ColumnIndex(templatesLo, "RECIPE_ID")
    Dim cIngredient As Long: cIngredient = ColumnIndex(templatesLo, "INGREDIENT_ID")
    Dim cProcess As Long: cProcess = ColumnIndex(templatesLo, "PROCESS")
    Dim cTargetTable As Long: cTargetTable = ColumnIndex(templatesLo, "TARGET_TABLE")
    Dim cTargetCol As Long: cTargetCol = ColumnIndex(templatesLo, "TARGET_COLUMN")
    Dim cFormula As Long: cFormula = ColumnIndex(templatesLo, "FORMULA")
    Dim cActive As Long: cActive = ColumnIndex(templatesLo, "ACTIVE")

    Dim arr As Variant: arr = templatesLo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If cActive > 0 Then
            If LCase$(Trim$(NzStr(arr(r, cActive)))) = "false" Then GoTo NextRow
        End If
        If cScope > 0 Then
            If StrComp(NzStr(arr(r, cScope)), templateScope, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If cRecipe > 0 And recipeId <> "" Then
            If StrComp(NzStr(arr(r, cRecipe)), recipeId, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If cIngredient > 0 And ingredientId <> "" Then
            If StrComp(NzStr(arr(r, cIngredient)), ingredientId, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If cProcess > 0 And processKey <> "" Then
            If StrComp(NzStr(arr(r, cProcess)), processKey, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If cTargetTable > 0 Then
            Dim templateTable As String
            templateTable = NzStr(arr(r, cTargetTable))
            If templateTable <> "" Then
                Dim targetName As String
                If tableNameOverride <> "" Then
                    targetName = tableNameOverride
                Else
                    targetName = targetLo.Name
                End If
                If StrComp(templateTable, targetName, vbTextCompare) <> 0 Then GoTo NextRow
            End If
        End If

        Dim colName As String
        colName = NzStr(arr(r, cTargetCol))
        If colName = "" Then GoTo NextRow

        Dim formulaText As String
        formulaText = NzStr(arr(r, cFormula))
        If formulaText = "" Or Left$(formulaText, 1) <> "=" Then
            Dim cell As Range
            Set cell = templatesLo.DataBodyRange.Cells(r, cFormula)
            If Not cell Is Nothing Then
                If cell.HasFormula Then formulaText = CStr(cell.FormulaR1C1)
            End If
        End If
        If formulaText = "" Or Left$(formulaText, 1) <> "=" Then GoTo NextRow

        ApplyFormulaToColumn targetLo, colName, formulaText
NextRow:
    Next r
End Sub

Private Sub ApplyFormulaToColumn(ByVal targetLo As ListObject, ByVal colName As String, ByVal formulaText As String)
    If targetLo Is Nothing Then Exit Sub
    Dim colIdx As Long: colIdx = ColumnIndex(targetLo, colName)
    If colIdx = 0 Then Exit Sub
    If targetLo.DataBodyRange Is Nothing Then Exit Sub

    Dim rng As Range
    Set rng = targetLo.ListColumns(colIdx).DataBodyRange
    On Error Resume Next
    Err.Clear
    rng.FormulaR1C1 = formulaText
    If Err.Number <> 0 Then
        Err.Clear
        rng.Formula = formulaText
    End If
    On Error GoTo 0
End Sub

Private Function GetTemplatesTable() As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_TEMPLATES)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set GetTemplatesTable = ws.ListObjects(TABLE_TEMPLATES)
    On Error GoTo 0
End Function

Private Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    ColumnIndex = 0
End Function

Private Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function
