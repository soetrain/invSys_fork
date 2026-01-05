VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPickerRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSearch As cDynItemSearch

Public Sub HandleSelectionChange(ByVal Target As Range)
    ' Placeholder: no automatic picker on selection to avoid noisy popups.
End Sub

Public Function HandleBeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean) As Boolean
    If Target Is Nothing Then Exit Function
    Dim lo As ListObject
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function

    Dim colName As String
    colName = ColumnNameFromCell(lo, Target)
    If colName = "" Then Exit Function

    If LCase$(lo.Name) = "ip_chooserecipe" And (LCase$(colName) = "recipe_name" Or LCase$(colName) = "recipe") Then
        EnsureSearch
        mSearch.ShowForPaletteRecipeCell Target
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If

    If LCase$(lo.Name) = "ip_chooseingredient" And LCase$(colName) = "ingredient" Then
        Dim recipeId As String
        recipeId = mProduction.GetPaletteRecipeId()
        If Trim$(recipeId) = "" Then
            MsgBox "Select a RECIPE in IP_ChooseRecipe first.", vbInformation
            Cancel = True
            HandleBeforeDoubleClick = True
            Exit Function
        End If
        EnsureSearch
        mSearch.ShowForIngredientCell Target, recipeId
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If

    If LCase$(lo.Name) = "ip_chooseitem" And (LCase$(colName) = "item" Or LCase$(colName) = "items") Then
        EnsureSearch
        mSearch.ShowForCell Target
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If

    If (LCase$(lo.Name) = "inventorypalette_generated" Or LCase$(lo.Name) Like "proc_*_palette") _
        And (LCase$(colName) = "item" Or LCase$(colName) = "items") Then

        Dim ctxRec As String, ctxIng As String, ctxProc As String, ctxIO As String
        Dim ctxAmt As Variant
        If Not mProduction.GetPaletteTableContext(lo, ctxRec, ctxIng, ctxAmt, ctxProc, ctxIO) Then
            MsgBox "Palette table context not found. Load a recipe in Recipe Chooser first.", vbInformation
            Cancel = True
            HandleBeforeDoubleClick = True
            Exit Function
        End If

        Dim allowedRows As Object
        Set allowedRows = mProduction.GetAllowedInvRowsForIngredient(ctxRec, ctxIng)
        If allowedRows Is Nothing Or allowedRows.Count = 0 Then
            MsgBox "No IngredientPalette entries found for this ingredient. Add acceptable items first.", vbInformation
            Cancel = True
            HandleBeforeDoubleClick = True
            Exit Function
        End If

        EnsureSearch
        mSearch.ShowForPaletteItemCell Target, allowedRows
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If

    If LCase$(lo.Name) = "rc_recipechoose" And (LCase$(colName) = "recipe" Or LCase$(colName) = "recipe_name") Then
        EnsureSearch
        mSearch.ShowForRecipeChooserCell Target
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If
End Function

Private Sub EnsureSearch()
    If mSearch Is Nothing Then Set mSearch = New cDynItemSearch
End Sub

Private Function ColumnNameFromCell(lo As ListObject, targetCell As Range) As String
    If lo Is Nothing Then Exit Function
    If targetCell Is Nothing Then Exit Function
    Dim colIdx As Long
    colIdx = targetCell.Column - lo.Range.Column + 1
    If colIdx < 1 Or colIdx > lo.ListColumns.Count Then Exit Function
    ColumnNameFromCell = lo.ListColumns(colIdx).Name
End Function
