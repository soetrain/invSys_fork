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
    Dim allowedRows As Object
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function

    Dim colName As String
    colName = ColumnNameFromCell(lo, Target)
    If colName = "" Then Exit Function
    Dim colKey As String
    colKey = NormalizeHeaderKey(colName)

    If LCase$(lo.Name) = "ip_chooserecipe" And (colKey = "RECIPENAME" Or colKey = "RECIPE") Then
        EnsureSearch
        mSearch.ShowForPaletteRecipeCell Target
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If

    If LCase$(lo.Name) = "ip_chooseingredient" And colKey = "INGREDIENT" Then
        Dim recipeId As String
        recipeId = NzStr(RunRoleMacro("mProduction.GetPaletteRecipeId"))
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

    If LCase$(lo.Name) = "ip_chooseitem" And (colKey = "ITEM" Or colKey = "ITEMS") Then
        Dim palRecipeId As String
        Dim palIngId As String
        palRecipeId = NzStr(RunRoleMacro("mProduction.GetPaletteRecipeId"))
        palIngId = NzStr(RunRoleMacro("mProduction.GetPaletteIngredientId"))
        If palRecipeId = "" Or palIngId = "" Then
            MsgBox "Select a RECIPE and INGREDIENT before choosing an item.", vbInformation
            Cancel = True
            HandleBeforeDoubleClick = True
            Exit Function
        End If

        EnsureSearch
        mSearch.ShowForCell Target
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If

    If (LCase$(lo.Name) = "inventorypalette_generated" Or LCase$(lo.Name) Like "proc_*_palette") _
        And (colKey = "ITEM" Or colKey = "ITEMS" Or colKey = "ITEMCODE") Then

        If lo.DataBodyRange Is Nothing Then Exit Function
        If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Function

        Dim ctxRec As String, ctxIng As String, ctxProc As String, ctxIO As String
        Dim ctxAmt As Variant
        If Not ResolvePaletteTableContext(lo, ctxRec, ctxIng, ctxAmt, ctxProc, ctxIO) Then
            MsgBox "Palette table context not found. Load a recipe in Recipe Chooser first.", vbInformation
            Cancel = True
            HandleBeforeDoubleClick = True
            Exit Function
        End If

        Set allowedRows = RunRoleMacroObject("mProduction.GetAllowedInvRowsForIngredient", ctxRec, ctxIng)
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
    mSearch.UseTemplateForm "ufProductionItemSearch"
End Sub

Private Function ColumnNameFromCell(lo As ListObject, targetCell As Range) As String
    If lo Is Nothing Then Exit Function
    If targetCell Is Nothing Then Exit Function
    Dim colIdx As Long
    colIdx = targetCell.Column - lo.Range.Column + 1
    If colIdx < 1 Or colIdx > lo.ListColumns.Count Then Exit Function
    ColumnNameFromCell = lo.ListColumns(colIdx).Name
End Function

Private Function NormalizeHeaderKey(ByVal v As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(v)
        ch = Mid$(v, i, 1)
        If ch Like "[A-Za-z0-9]" Then out = out & UCase$(ch)
    Next i
    NormalizeHeaderKey = out
End Function

Private Function ResolvePaletteTableContext(ByVal lo As ListObject, ByRef recipeId As String, ByRef ingredientId As String, _
    ByRef amount As Variant, ByRef procName As String, ByRef ioVal As String) As Boolean

    Dim info As Variant
    info = RunRoleMacro("mProduction.GetPaletteTableContextInfo", lo)
    If IsEmpty(info) Then Exit Function
    On Error GoTo Fail
    recipeId = NzStr(info(1))
    ingredientId = NzStr(info(2))
    amount = info(3)
    procName = NzStr(info(4))
    ioVal = NzStr(info(5))
    ResolvePaletteTableContext = True
    Exit Function
Fail:
    ResolvePaletteTableContext = False
End Function

Private Function RunRoleMacro(ByVal macroName As String, ParamArray args() As Variant) As Variant
    Select Case ParamCount(args)
        Case 0: RunRoleMacro = Application.Run(macroName)
        Case 1: RunRoleMacro = Application.Run(macroName, args(0))
        Case 2: RunRoleMacro = Application.Run(macroName, args(0), args(1))
        Case 3: RunRoleMacro = Application.Run(macroName, args(0), args(1), args(2))
        Case 4: RunRoleMacro = Application.Run(macroName, args(0), args(1), args(2), args(3))
        Case Else: Err.Raise 5, "cPickerRouter.RunRoleMacro", "Too many macro arguments."
    End Select
End Function

Private Function RunRoleMacroObject(ByVal macroName As String, ParamArray args() As Variant) As Object
    Dim result As Variant
    result = RunRoleMacroForward(macroName, args)
    If IsObject(result) Then Set RunRoleMacroObject = result
End Function

Private Function RunRoleMacroForward(ByVal macroName As String, ByRef args() As Variant) As Variant
    Select Case ParamCount(args)
        Case 0: RunRoleMacroForward = Application.Run(macroName)
        Case 1: RunRoleMacroForward = Application.Run(macroName, args(0))
        Case 2: RunRoleMacroForward = Application.Run(macroName, args(0), args(1))
        Case 3: RunRoleMacroForward = Application.Run(macroName, args(0), args(1), args(2))
        Case 4: RunRoleMacroForward = Application.Run(macroName, args(0), args(1), args(2), args(3))
        Case Else: Err.Raise 5, "cPickerRouter.RunRoleMacroForward", "Too many macro arguments."
    End Select
End Function

Private Function ParamCount(ByRef args() As Variant) As Long
    On Error GoTo NoArgs
    ParamCount = UBound(args) - LBound(args) + 1
    Exit Function
NoArgs:
    ParamCount = 0
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function
