Attribute VB_Name = "mProduction"
Option Explicit

' Production system core module (wiring + helpers).

Private Const SHEET_PRODUCTION As String = "Production"
Private Const SHEET_TEMPLATES As String = "TemplatesTable"

Private Const TABLE_RECIPE_CHOOSER As String = "RC_RecipeChoose"
Private Const TABLE_INV_PALETTE_GENERATED As String = "InventoryPalette_generated"
Private Const TABLE_RECIPE_BUILDER_HEADER As String = "RB_AddRecipeName"
Private Const TABLE_RECIPE_BUILDER_LINES As String = "RecipeBuilder"

Private Const BTN_HIDE_SYSTEM As String = "BTN_HIDE_SYSTEM"
Private Const BTN_SHOW_SYSTEM As String = "BTN_SHOW_SYSTEM"
Private Const BTN_LOAD_RECIPE As String = "BTN_LOAD_RECIPE"
Private Const BTN_SAVE_RECIPE As String = "BTN_SAVE_RECIPE"
Private Const BTN_SAVE_PALETTE As String = "BTN_SAVE_PALETTE"
Private Const BTN_TO_USED As String = "BTN_TO_USED"
Private Const BTN_TO_MADE As String = "BTN_TO_MADE"
Private Const BTN_TO_TOTALINV As String = "BTN_TO_TOTALINV"
Private Const BTN_NEXT_BATCH As String = "BTN_NEXT_BATCH"
Private Const BTN_PRINT_CODES As String = "BTN_PRINT_CODES"

Private mRowCountCache As Object
Private mHiddenSystems As Collection
Private mSystemGroupsInit As Boolean
Private mSystemGroupNames(1 To 4) As String
Private mSystemGroupTables(1 To 4) As Variant

Public Sub InitializeProductionUI()
    EnsureProductionButtons
    EnsureSystemGroups
End Sub

' ===== Worksheet event entry points =====
Public Sub HandleProductionSelectionChange(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(Target) Then Exit Sub
    Dim router As New cPickerRouter
    router.HandleSelectionChange Target
End Sub

Public Sub HandleProductionBeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    If Target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(Target) Then Exit Sub
    Dim router As New cPickerRouter
    If router.HandleBeforeDoubleClick(Target, Cancel) Then Exit Sub
End Sub

Public Sub HandleProductionChange(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(Target) Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    If IsPaletteTable(lo) Then
        EnsureRowCountCache
        Dim key As String: key = lo.Name
        Dim newCount As Long: newCount = ListObjectRowCount(lo)
        Dim oldCount As Long
        If mRowCountCache.Exists(key) Then oldCount = CLng(mRowCountCache(key))
        If newCount > oldCount Then
            Dim bandMgr As New cTableBandManager
            bandMgr.Init lo.Parent
            bandMgr.ExpandBandForTable lo, (newCount - oldCount)
        End If
        mRowCountCache(key) = newCount
    End If
End Sub

' ===== Band/table helpers =====
Private Sub EnsureRowCountCache()
    If mRowCountCache Is Nothing Then
        Set mRowCountCache = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Function IsOnProductionSheet(ByVal Target As Range) As Boolean
    On Error Resume Next
    IsOnProductionSheet = (Target.Worksheet.Name = SHEET_PRODUCTION)
    On Error GoTo 0
End Function

Private Function IsPaletteTable(lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If nm = LCase$(TABLE_INV_PALETTE_GENERATED) Then
        IsPaletteTable = True
    ElseIf nm Like "proc_*_palette" Then
        IsPaletteTable = True
    End If
End Function

Private Function ListObjectRowCount(lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ListObjectRowCount = lo.DataBodyRange.Rows.Count
End Function

' ===== Generic helpers =====
Public Function GetProductionSheet() As Worksheet
    Set GetProductionSheet = SheetExists(SHEET_PRODUCTION)
End Function

Public Function SheetExists(nameOrCode As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set SheetExists = ws
            Exit Function
        End If
    Next ws
End Function

Public Function GetListObject(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function FindListObjectByNameOrHeaders(ws As Worksheet, tableName As String, headers As Variant) As ListObject
    Dim lo As ListObject
    Set lo = GetListObject(ws, tableName)
    If Not lo Is Nothing Then
        Set FindListObjectByNameOrHeaders = lo
        Exit Function
    End If
    For Each lo In ws.ListObjects
        If ListObjectHasHeaders(lo, headers) Then
            Set FindListObjectByNameOrHeaders = lo
            Exit Function
        End If
    Next lo
End Function

Private Function ListObjectHasHeaders(lo As ListObject, headers As Variant) As Boolean
    If lo Is Nothing Then Exit Function
    If lo.HeaderRowRange Is Nothing Then Exit Function
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If ColumnIndex(lo, CStr(headers(i))) = 0 Then Exit Function
    Next i
    ListObjectHasHeaders = True
End Function

Private Function TableColumnCount(lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.HeaderRowRange Is Nothing Then Exit Function
    TableColumnCount = lo.HeaderRowRange.Columns.Count
End Function

Private Sub ExpandSpanForTable(lo As ListObject, ByRef firstCol As Long, ByRef lastCol As Long)
    If lo Is Nothing Then Exit Sub
    If lo.HeaderRowRange Is Nothing Then Exit Sub
    Dim startCol As Long
    Dim endCol As Long
    If Not TableEffectiveSpan(lo, startCol, endCol) Then Exit Sub
    If firstCol = 0 Or startCol < firstCol Then firstCol = startCol
    If endCol > lastCol Then lastCol = endCol
End Sub

Private Function TableEffectiveSpan(lo As ListObject, ByRef startCol As Long, ByRef endCol As Long) As Boolean
    TableEffectiveSpan = False
    If lo Is Nothing Then Exit Function
    If lo.HeaderRowRange Is Nothing Then Exit Function

    Dim hdr As Range: Set hdr = lo.HeaderRowRange
    startCol = hdr.Column

    Dim lastIdx As Long
    Dim i As Long
    For i = hdr.Columns.Count To 1 Step -1
        Dim val As String
        val = Trim$(CStr(hdr.Cells(1, i).Value))
        If val <> "" Then
            lastIdx = i
            Exit For
        End If
    Next i
    If lastIdx = 0 Then lastIdx = hdr.Columns.Count
    endCol = startCol + lastIdx - 1
    TableEffectiveSpan = (endCol >= startCol)
End Function

Private Function ResolveListObject(ws As Worksheet, tableName As String) As ListObject
    Select Case tableName
        Case "RB_AddRecipeName", "RecipeBuilder", "IP_ChooseRecipe", "IP_ChooseIngredient", _
             "IP_ChooseItem", "RC_RecipeChoose", "RecipeChooser_generated", _
             "InventoryPalette_generated", "ProductionOutput", "Prod_invSys_Check"
            Set ResolveListObject = GetListObject(ws, tableName)
        Case Else
            Set ResolveListObject = GetListObject(ws, tableName)
    End Select
End Function

Private Sub GetSystemBounds(ws As Worksheet, ByRef startCols() As Long, ByRef endCols() As Long, ByRef topRows() As Long, ByRef bottomRows() As Long)
    Dim i As Long
    ReDim startCols(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))
    ReDim endCols(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))
    ReDim topRows(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))
    ReDim bottomRows(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))

    Dim maxEnd() As Long
    ReDim maxEnd(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))

    For i = LBound(mSystemGroupNames) To UBound(mSystemGroupNames)
        Dim tablesArr As Variant
        tablesArr = mSystemGroupTables(i)
        Dim j As Long
        For j = LBound(tablesArr) To UBound(tablesArr)
            Dim lo As ListObject
            Set lo = ResolveListObject(ws, CStr(tablesArr(j)))
            If Not lo Is Nothing Then
                Dim sCol As Long
                Dim eCol As Long
                Dim rTop As Long
                Dim rBottom As Long
                If TableEffectiveSpan(lo, sCol, eCol) Then
                    rTop = lo.Range.Row
                    rBottom = lo.Range.Row + lo.Range.Rows.Count - 1
                    If startCols(i) = 0 Or sCol < startCols(i) Then startCols(i) = sCol
                    If eCol > maxEnd(i) Then maxEnd(i) = eCol
                    If topRows(i) = 0 Or rTop < topRows(i) Then topRows(i) = rTop
                    If rBottom > bottomRows(i) Then bottomRows(i) = rBottom
                End If
            End If
        Next j
    Next i

    ' Define end bounds by the next group's start (keeps bands discrete).
    For i = LBound(mSystemGroupNames) To UBound(mSystemGroupNames)
        If startCols(i) = 0 Then GoTo NextGroup
        Dim nextStart As Long
        nextStart = 0
        Dim k As Long
        For k = i + 1 To UBound(mSystemGroupNames)
            If startCols(k) > 0 Then
                nextStart = startCols(k)
                Exit For
            End If
        Next k
        If nextStart > 0 Then
            endCols(i) = nextStart - 1
        Else
            endCols(i) = maxEnd(i)
            ' Rightmost system: extend to include any checkbox shapes to the right.
            Dim maxChkCol As Long
            maxChkCol = MaxCheckboxColumn(ws, startCols(i))
            If maxChkCol > endCols(i) Then endCols(i) = maxChkCol
        End If
NextGroup:
    Next i
End Sub

Private Function MaxCheckboxColumn(ws As Worksheet, startCol As Long) As Long
    If ws Is Nothing Then Exit Function
    If startCol = 0 Then Exit Function
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim isCheckbox As Boolean
        On Error Resume Next
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlCheckBox Then isCheckbox = True
        End If
        If Not isCheckbox Then
            If LCase$(shp.Name) Like "check box*" Then isCheckbox = True
        End If
        Dim c As Long
        c = shp.TopLeftCell.Column
        On Error GoTo 0
        If isCheckbox And c >= startCol Then
            If c > MaxCheckboxColumn Then MaxCheckboxColumn = c
        End If
    Next shp

    Dim ole As OLEObject
    For Each ole In ws.OLEObjects
        Dim isChk As Boolean
        On Error Resume Next
        Dim tName As String
        tName = TypeName(ole.Object)
        If LCase$(tName) Like "*checkbox*" Then isChk = True
        Dim cOle As Long
        cOle = ole.TopLeftCell.Column
        On Error GoTo 0
        If isChk And cOle >= startCol Then
            If cOle > MaxCheckboxColumn Then MaxCheckboxColumn = cOle
        End If
    Next ole
End Function

Private Function IsSystemVisible(ws As Worksheet, startCol As Long, endCol As Long) As Boolean
    If ws Is Nothing Then Exit Function
    If startCol = 0 Or endCol = 0 Then Exit Function
    Dim c As Long
    For c = startCol To endCol
        If Not ws.Columns(c).EntireColumn.Hidden Then
            IsSystemVisible = True
            Exit Function
        End If
    Next c
End Function

Public Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    ColumnIndex = 0
End Function

' ===== button scaffolding =====
Private Sub EnsureProductionButtons()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    Dim colA As Range: Set colA = ws.Columns("A")
    Dim leftA As Double: leftA = colA.Left + 2
    Dim colAWidth As Double
    colAWidth = colA.Width - 4
    If colAWidth < 40 Then colAWidth = 90

    Const BTN_STACK_SPACING As Double = 24
    Dim nextTop As Double: nextTop = ws.Rows(2).Top

    EnsureButtonCustom ws, BTN_HIDE_SYSTEM, "Hide system", "mProduction.BtnHideSystem", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_SHOW_SYSTEM, "Show system", "mProduction.BtnShowSystem", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING

    DeleteShapeIfExists ws, "BTN_TOGGLE_RECIPE_BUILDER"
    DeleteShapeIfExists ws, "BTN_TOGGLE_PALETTE_BUILDER"
    DeleteShapeIfExists ws, "BTN_TOGGLE_PRODUCTION"
    EnsureButtonCustom ws, BTN_LOAD_RECIPE, "Load Recipe", "mProduction.BtnLoadRecipe", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_SAVE_RECIPE, "Save Recipe", "mProduction.BtnSaveRecipe", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_SAVE_PALETTE, "Save IngredientPalette", "mProduction.BtnSavePalette", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_TO_USED, "To USED", "mProduction.BtnToUsed", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_TO_MADE, "Send to MADE", "mProduction.BtnToMade", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_TO_TOTALINV, "Send to TOTAL INV", "mProduction.BtnToTotalInv", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_NEXT_BATCH, "Next Batch", "mProduction.BtnNextBatch", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_PRINT_CODES, "Print recall codes", "mProduction.BtnPrintRecallCodes", leftA, nextTop, colAWidth
End Sub

Private Sub EnsureSystemGroups()
    If mSystemGroupsInit Then Exit Sub
    mSystemGroupNames(1) = "RecipeListBuilder"
    mSystemGroupTables(1) = Array("RecipeBuilder", "RB_AddRecipeName")

    mSystemGroupNames(2) = "InventoryPaletteBuilder"
    mSystemGroupTables(2) = Array("IP_ChooseIngredient", "IP_ChooseItem", "IP_ChooseRecipe")

    mSystemGroupNames(3) = "RecipeChooser"
    mSystemGroupTables(3) = Array("RC_RecipeChoose", "RecipeChooser_generated")

    mSystemGroupNames(4) = "ProductionInputOutput"
    mSystemGroupTables(4) = Array("InventoryPalette_generated", "ProductionOutput", "Prod_invSys_Check")

    Set mHiddenSystems = New Collection
    mSystemGroupsInit = True
End Sub

' ===== show/hide system bands =====
Public Sub BtnHideSystem()
    EnsureSystemGroups
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    Dim starts() As Long, ends() As Long, tops() As Long, bottoms() As Long
    GetSystemBounds ws, starts, ends, tops, bottoms

    Dim i As Long, nearestIdx As Long, bestStart As Long
    For i = LBound(mSystemGroupNames) To UBound(mSystemGroupNames)
        If starts(i) > 0 And ends(i) > 0 Then
            If IsSystemVisible(ws, starts(i), ends(i)) Then
                If bestStart = 0 Or starts(i) < bestStart Then
                    bestStart = starts(i)
                    nearestIdx = i
                End If
            End If
        End If
    Next i

    If nearestIdx = 0 Then Exit Sub
    ws.Range(ws.Columns(starts(nearestIdx)), ws.Columns(ends(nearestIdx))).EntireColumn.Hidden = True
    HideGroupShapes ws, starts(nearestIdx), ends(nearestIdx), tops(nearestIdx), bottoms(nearestIdx), True
    mHiddenSystems.Add nearestIdx
End Sub

Public Sub BtnShowSystem()
    EnsureSystemGroups
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub
    Dim starts() As Long, ends() As Long, tops() As Long, bottoms() As Long
    GetSystemBounds ws, starts, ends, tops, bottoms

    Dim idx As Long
    If Not mHiddenSystems Is Nothing And mHiddenSystems.Count > 0 Then
        idx = CLng(mHiddenSystems(mHiddenSystems.Count))
        mHiddenSystems.Remove mHiddenSystems.Count
    Else
        ' Fallback: show rightmost hidden system.
        Dim i As Long, bestStart As Long
        For i = LBound(mSystemGroupNames) To UBound(mSystemGroupNames)
            If starts(i) > 0 And ends(i) > 0 Then
                If Not IsSystemVisible(ws, starts(i), ends(i)) Then
                    If starts(i) > bestStart Then
                        bestStart = starts(i)
                        idx = i
                    End If
                End If
            End If
        Next i
        If idx = 0 Then Exit Sub
    End If

    If starts(idx) = 0 Or ends(idx) = 0 Then Exit Sub
    ws.Range(ws.Columns(starts(idx)), ws.Columns(ends(idx))).EntireColumn.Hidden = False
    HideGroupShapes ws, starts(idx), ends(idx), tops(idx), bottoms(idx), False
End Sub

Private Sub EnsureButtonCustom(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String, leftPos As Double, topPos As Double, Optional widthPts As Double = 118)
    Const BTN_HEIGHT As Double = 20
    If widthPts < 20 Then widthPts = 118
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, widthPts, BTN_HEIGHT)
        shp.Name = shapeName
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = onActionMacro
    Else
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = BTN_HEIGHT
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = onActionMacro
    End If
End Sub

Private Sub DeleteShapeIfExists(ws As Worksheet, shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub HideGroupShapes(ws As Worksheet, startCol As Long, endCol As Long, topRow As Long, bottomRow As Long, hideIt As Boolean)
    If ws Is Nothing Then Exit Sub
    If startCol = 0 Or endCol = 0 Then Exit Sub
    Dim endColAdj As Long
    endColAdj = endCol + 6 ' allow checkboxes just right of the table
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim c As Long
        Dim r As Long
        On Error Resume Next
        c = shp.TopLeftCell.Column
        r = shp.TopLeftCell.Row
        On Error GoTo 0
        If c >= startCol And c <= endColAdj Then
            shp.Visible = IIf(hideIt, msoFalse, msoTrue)
        End If
    Next shp

    Dim ole As OLEObject
    For Each ole In ws.OLEObjects
        Dim isChk As Boolean
        On Error Resume Next
        Dim tName As String
        tName = TypeName(ole.Object)
        If LCase$(tName) Like "*checkbox*" Then isChk = True
        Dim cOle As Long
        cOle = ole.TopLeftCell.Column
        On Error GoTo 0
        If isChk Then
            If cOle >= startCol And cOle <= endColAdj Then
                ole.Visible = Not hideIt
            End If
        End If
    Next ole
End Sub

' ===== button handlers (stubs for now) =====
Public Sub BtnLoadRecipe()
    LoadRecipeFromRecipes
End Sub

Public Sub BtnSaveRecipe()
    SaveRecipeToRecipes
End Sub

Public Sub BtnSavePalette()
    MsgBox "Save IngredientPalette not implemented yet.", vbInformation
End Sub

Public Sub BtnToUsed()
    MsgBox "To USED not implemented yet.", vbInformation
End Sub

Public Sub BtnToMade()
    MsgBox "Send to MADE not implemented yet.", vbInformation
End Sub

Public Sub BtnToTotalInv()
    MsgBox "Send to TOTAL INV not implemented yet.", vbInformation
End Sub

Public Sub BtnNextBatch()
    MsgBox "Next Batch not implemented yet.", vbInformation
End Sub

Public Sub BtnPrintRecallCodes()
    MsgBox "Print recall codes not implemented yet.", vbInformation
End Sub

' ===== Recipe Builder: Load / Save =====
Private Sub SaveRecipeToRecipes()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then
        MsgBox "Recipes sheet not found.", vbCritical
        Exit Sub
    End If

    Dim loHeader As ListObject: Set loHeader = GetListObject(wsProd, TABLE_RECIPE_BUILDER_HEADER)
    Dim loLines As ListObject: Set loLines = GetListObject(wsProd, TABLE_RECIPE_BUILDER_LINES)
    If loHeader Is Nothing Or loLines Is Nothing Then
        MsgBox "Recipe Builder tables not found on Production sheet.", vbExclamation
        Exit Sub
    End If
    If loHeader.DataBodyRange Is Nothing Then
        MsgBox "Enter a recipe name before saving.", vbExclamation
        Exit Sub
    End If
    If loLines.DataBodyRange Is Nothing Then
        MsgBox "Add at least one recipe line before saving.", vbExclamation
        Exit Sub
    End If

    Dim cName As Long: cName = ColumnIndex(loHeader, "RECIPE_NAME")
    Dim cDesc As Long: cDesc = ColumnIndex(loHeader, "DESCRIPTION")
    Dim cGuid As Long: cGuid = ColumnIndex(loHeader, "GUID")
    Dim cRecipeId As Long: cRecipeId = ColumnIndex(loHeader, "RECIPE_ID")
    If cName = 0 Or cRecipeId = 0 Then
        MsgBox "Recipe Builder header missing RECIPE_NAME or RECIPE_ID.", vbCritical
        Exit Sub
    End If

    Dim recipeName As String: recipeName = NzStr(loHeader.DataBodyRange.Cells(1, cName).Value)
    Dim recipeDesc As String
    If cDesc > 0 Then recipeDesc = NzStr(loHeader.DataBodyRange.Cells(1, cDesc).Value)
    If Trim$(recipeName) = "" Then
        MsgBox "Enter a RECIPE_NAME before saving.", vbExclamation
        Exit Sub
    End If

    Dim recipeId As String: recipeId = NzStr(loHeader.DataBodyRange.Cells(1, cRecipeId).Value)
    If recipeId = "" Then
        recipeId = modUR_Snapshot.GenerateGUID()
        loHeader.DataBodyRange.Cells(1, cRecipeId).Value = recipeId
    End If
    If cGuid > 0 Then
        Dim recipeGuid As String: recipeGuid = NzStr(loHeader.DataBodyRange.Cells(1, cGuid).Value)
        If recipeGuid = "" Then
            recipeGuid = modUR_Snapshot.GenerateGUID()
            loHeader.DataBodyRange.Cells(1, cGuid).Value = recipeGuid
        End If
    End If

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then
        MsgBox "Recipes table not found on Recipes sheet.", vbCritical
        Exit Sub
    End If
    Dim cRecRecipeId As Long: cRecRecipeId = ColumnIndex(loRecipes, "RECIPE_ID")
    If cRecRecipeId = 0 Then
        MsgBox "Recipes table missing RECIPE_ID column.", vbCritical
        Exit Sub
    End If

    ' Delete existing rows for this recipe ID (overwrite behavior).
    If Not loRecipes.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = loRecipes.DataBodyRange.Rows.Count To 1 Step -1
            If NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipeId).Value) = recipeId Then
                loRecipes.ListRows(r).Delete
            End If
        Next r
    End If

    ' Column indexes in Recipes table.
    Dim cRecRecipe As Long: cRecRecipe = ColumnIndex(loRecipes, "RECIPE")
    Dim cRecDesc As Long: cRecDesc = ColumnIndex(loRecipes, "DESCRIPTION")
    Dim cRecDept As Long: cRecDept = ColumnIndex(loRecipes, "DEPARTMENT")
    Dim cRecProcess As Long: cRecProcess = ColumnIndex(loRecipes, "PROCESS")
    Dim cRecDiagram As Long: cRecDiagram = ColumnIndex(loRecipes, "DIAGRAM_ID")
    Dim cRecIO As Long: cRecIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cRecIngredient As Long: cRecIngredient = ColumnIndex(loRecipes, "INGREDIENT")
    Dim cRecPercent As Long: cRecPercent = ColumnIndex(loRecipes, "PERCENT")
    Dim cRecUom As Long: cRecUom = ColumnIndex(loRecipes, "UOM")
    Dim cRecAmount As Long: cRecAmount = ColumnIndex(loRecipes, "AMOUNT")
    Dim cRecListRow As Long: cRecListRow = ColumnIndex(loRecipes, "RECIPE_LIST_ROW")
    Dim cRecIngId As Long: cRecIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cRecGuid As Long: cRecGuid = ColumnIndex(loRecipes, "GUID")

    ' Column indexes in RecipeBuilder lines.
    Dim cProc As Long: cProc = ColumnIndex(loLines, "PROCESS")
    Dim cDiag As Long: cDiag = ColumnIndex(loLines, "DIAGRAM_ID")
    Dim cIO As Long: cIO = ColumnIndex(loLines, "INPUT/OUTPUT")
    Dim cIng As Long: cIng = ColumnIndex(loLines, "INGREDIENT")
    Dim cPct As Long: cPct = ColumnIndex(loLines, "PERCENT")
    Dim cUomLine As Long: cUomLine = ColumnIndex(loLines, "UOM")
    Dim cAmt As Long: cAmt = ColumnIndex(loLines, "AMOUNT")
    Dim cListRow As Long: cListRow = ColumnIndex(loLines, "RECIPE_LIST_ROW")
    Dim cIngId As Long: cIngId = ColumnIndex(loLines, "INGREDIENT_ID")
    Dim cGuidLine As Long: cGuidLine = ColumnIndex(loLines, "GUID")

    Dim lineArr As Variant: lineArr = loLines.DataBodyRange.Value
    Dim rowCount As Long: rowCount = UBound(lineArr, 1)
    Dim savedCount As Long
    Dim seqRow As Long: seqRow = 1
    Dim i As Long
    For i = 1 To rowCount
        Dim hasData As Boolean
        If cIng > 0 Then
            hasData = (Trim$(NzStr(lineArr(i, cIng))) <> "")
        ElseIf cProc > 0 Then
            hasData = (Trim$(NzStr(lineArr(i, cProc))) <> "")
        End If
        If Not hasData Then GoTo NextLine

        Dim ingId As String
        If cIngId > 0 Then ingId = NzStr(lineArr(i, cIngId))
        If ingId = "" Then
            ingId = modUR_Snapshot.GenerateGUID()
            loLines.DataBodyRange.Cells(i, cIngId).Value = ingId
        End If

        Dim recListRow As Variant
        If cListRow > 0 Then recListRow = lineArr(i, cListRow)
        If NzStr(recListRow) = "" Then
            recListRow = seqRow
            loLines.DataBodyRange.Cells(i, cListRow).Value = recListRow
        End If

        Dim rowGuid As String
        If cGuidLine > 0 Then rowGuid = NzStr(lineArr(i, cGuidLine))
        If rowGuid = "" Then
            rowGuid = modUR_Snapshot.GenerateGUID()
            If cGuidLine > 0 Then loLines.DataBodyRange.Cells(i, cGuidLine).Value = rowGuid
        End If

        Dim lr As ListRow: Set lr = loRecipes.ListRows.Add
        If cRecRecipeId > 0 Then lr.Range.Cells(1, cRecRecipeId).Value = recipeId
        If cRecRecipe > 0 Then lr.Range.Cells(1, cRecRecipe).Value = recipeName
        If cRecDesc > 0 Then lr.Range.Cells(1, cRecDesc).Value = recipeDesc
        If cRecDept > 0 Then lr.Range.Cells(1, cRecDept).Value = "" ' optional for now
        If cRecProcess > 0 And cProc > 0 Then lr.Range.Cells(1, cRecProcess).Value = lineArr(i, cProc)
        If cRecDiagram > 0 And cDiag > 0 Then lr.Range.Cells(1, cRecDiagram).Value = lineArr(i, cDiag)
        If cRecIO > 0 And cIO > 0 Then lr.Range.Cells(1, cRecIO).Value = lineArr(i, cIO)
        If cRecIngredient > 0 And cIng > 0 Then lr.Range.Cells(1, cRecIngredient).Value = lineArr(i, cIng)
        If cRecPercent > 0 And cPct > 0 Then lr.Range.Cells(1, cRecPercent).Value = lineArr(i, cPct)
        If cRecUom > 0 And cUomLine > 0 Then lr.Range.Cells(1, cRecUom).Value = lineArr(i, cUomLine)
        If cRecAmount > 0 And cAmt > 0 Then lr.Range.Cells(1, cRecAmount).Value = lineArr(i, cAmt)
        If cRecListRow > 0 Then lr.Range.Cells(1, cRecListRow).Value = recListRow
        If cRecIngId > 0 Then lr.Range.Cells(1, cRecIngId).Value = ingId
        If cRecGuid > 0 Then lr.Range.Cells(1, cRecGuid).Value = rowGuid

        savedCount = savedCount + 1
        seqRow = seqRow + 1
NextLine:
    Next i

    If savedCount = 0 Then
        MsgBox "No recipe lines with data were found to save.", vbExclamation
    Else
        MsgBox "Saved recipe '" & recipeName & "' (" & savedCount & " lines).", vbInformation
    End If
    Exit Sub
ErrHandler:
    MsgBox "Save Recipe failed: " & Err.Description, vbCritical
End Sub

Private Sub LoadRecipeFromRecipes()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then
        MsgBox "Recipes sheet not found.", vbCritical
        Exit Sub
    End If

    Dim loHeader As ListObject: Set loHeader = GetListObject(wsProd, TABLE_RECIPE_BUILDER_HEADER)
    Dim loLines As ListObject: Set loLines = GetListObject(wsProd, TABLE_RECIPE_BUILDER_LINES)
    If loHeader Is Nothing Or loLines Is Nothing Then
        MsgBox "Recipe Builder tables not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim recipeId As String
    Dim recipeName As String

    Dim loSel As ListObject
    On Error Resume Next
    Set loSel = Application.ActiveCell.ListObject
    On Error GoTo 0
    If Not loSel Is Nothing Then
        Dim cSelRecipeId As Long: cSelRecipeId = ColumnIndex(loSel, "RECIPE_ID")
        Dim cSelRecipe As Long: cSelRecipe = ColumnIndex(loSel, "RECIPE")
        If cSelRecipeId > 0 Then
            recipeId = NzStr(loSel.DataBodyRange.Cells(Application.ActiveCell.Row - loSel.DataBodyRange.Row + 1, cSelRecipeId).Value)
        End If
        If recipeId = "" And cSelRecipe > 0 Then
            recipeName = NzStr(loSel.DataBodyRange.Cells(Application.ActiveCell.Row - loSel.DataBodyRange.Row + 1, cSelRecipe).Value)
        End If
    End If

    If recipeId = "" Then
        Dim cHeaderRecipeIdTmp As Long: cHeaderRecipeIdTmp = ColumnIndex(loHeader, "RECIPE_ID")
        If cHeaderRecipeIdTmp > 0 And Not loHeader.DataBodyRange Is Nothing Then
            recipeId = NzStr(loHeader.DataBodyRange.Cells(1, cHeaderRecipeIdTmp).Value)
        End If
    End If

    If recipeId = "" And recipeName = "" Then
        recipeId = InputBox("Enter RECIPE_ID to load:", "Load Recipe")
    End If

    If recipeId = "" And recipeName = "" Then Exit Sub

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then
        MsgBox "Recipes table not found on Recipes sheet.", vbCritical
        Exit Sub
    End If

    Dim cRecRecipeId As Long: cRecRecipeId = ColumnIndex(loRecipes, "RECIPE_ID")
    Dim cRecRecipe As Long: cRecRecipe = ColumnIndex(loRecipes, "RECIPE")
    Dim cRecDesc As Long: cRecDesc = ColumnIndex(loRecipes, "DESCRIPTION")
    Dim cRecProcess As Long: cRecProcess = ColumnIndex(loRecipes, "PROCESS")
    Dim cRecDiagram As Long: cRecDiagram = ColumnIndex(loRecipes, "DIAGRAM_ID")
    Dim cRecIO As Long: cRecIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cRecIngredient As Long: cRecIngredient = ColumnIndex(loRecipes, "INGREDIENT")
    Dim cRecPercent As Long: cRecPercent = ColumnIndex(loRecipes, "PERCENT")
    Dim cRecUom As Long: cRecUom = ColumnIndex(loRecipes, "UOM")
    Dim cRecAmount As Long: cRecAmount = ColumnIndex(loRecipes, "AMOUNT")
    Dim cRecListRow As Long: cRecListRow = ColumnIndex(loRecipes, "RECIPE_LIST_ROW")
    Dim cRecIngId As Long: cRecIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cRecGuid As Long: cRecGuid = ColumnIndex(loRecipes, "GUID")

    Dim matches As Collection: Set matches = New Collection
    If Not loRecipes.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To loRecipes.DataBodyRange.Rows.Count
            Dim rowRecipeId As String
            rowRecipeId = NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipeId).Value)
            Dim rowRecipeName As String
            If cRecRecipe > 0 Then rowRecipeName = NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipe).Value)
            If (recipeId <> "" And rowRecipeId = recipeId) Or (recipeId = "" And rowRecipeName = recipeName And rowRecipeName <> "") Then
                matches.Add r
                If recipeId = "" Then recipeId = rowRecipeId
                If recipeName = "" Then recipeName = rowRecipeName
            End If
        Next r
    End If

    If matches.Count = 0 Then
        MsgBox "No recipe rows found for the selected RECIPE_ID.", vbExclamation
        Exit Sub
    End If

    ' Update header table.
    Dim cHeaderName As Long: cHeaderName = ColumnIndex(loHeader, "RECIPE_NAME")
    Dim cHeaderDesc As Long: cHeaderDesc = ColumnIndex(loHeader, "DESCRIPTION")
    Dim cHeaderGuid As Long: cHeaderGuid = ColumnIndex(loHeader, "GUID")
    Dim cHeaderRecipeId As Long: cHeaderRecipeId = ColumnIndex(loHeader, "RECIPE_ID")
    EnsureTableHasRow loHeader
    If cHeaderName > 0 Then loHeader.DataBodyRange.Cells(1, cHeaderName).Value = recipeName
    If cHeaderRecipeId > 0 Then loHeader.DataBodyRange.Cells(1, cHeaderRecipeId).Value = recipeId
    If cHeaderDesc > 0 Then
        loHeader.DataBodyRange.Cells(1, cHeaderDesc).Value = NzStr(loRecipes.DataBodyRange.Cells(matches(1), cRecDesc).Value)
    End If
    If cHeaderGuid > 0 Then
        loHeader.DataBodyRange.Cells(1, cHeaderGuid).Value = NzStr(loRecipes.DataBodyRange.Cells(matches(1), cRecGuid).Value)
    End If

    ' Clear and rebuild RecipeBuilder lines.
    ClearListObjectData loLines
    Dim idx As Long
    For idx = 1 To matches.Count
        Dim rr As Long: rr = CLng(matches(idx))
        Dim lr As ListRow: Set lr = loLines.ListRows.Add
        Dim cProc As Long: cProc = ColumnIndex(loLines, "PROCESS")
        Dim cDiag As Long: cDiag = ColumnIndex(loLines, "DIAGRAM_ID")
        Dim cIO As Long: cIO = ColumnIndex(loLines, "INPUT/OUTPUT")
        Dim cIng As Long: cIng = ColumnIndex(loLines, "INGREDIENT")
        Dim cPct As Long: cPct = ColumnIndex(loLines, "PERCENT")
        Dim cUomLine As Long: cUomLine = ColumnIndex(loLines, "UOM")
        Dim cAmt As Long: cAmt = ColumnIndex(loLines, "AMOUNT")
        Dim cListRow As Long: cListRow = ColumnIndex(loLines, "RECIPE_LIST_ROW")
        Dim cIngId As Long: cIngId = ColumnIndex(loLines, "INGREDIENT_ID")
        Dim cGuidLine As Long: cGuidLine = ColumnIndex(loLines, "GUID")

        If cProc > 0 Then lr.Range.Cells(1, cProc).Value = loRecipes.DataBodyRange.Cells(rr, cRecProcess).Value
        If cDiag > 0 Then lr.Range.Cells(1, cDiag).Value = loRecipes.DataBodyRange.Cells(rr, cRecDiagram).Value
        If cIO > 0 Then lr.Range.Cells(1, cIO).Value = loRecipes.DataBodyRange.Cells(rr, cRecIO).Value
        If cIng > 0 Then lr.Range.Cells(1, cIng).Value = loRecipes.DataBodyRange.Cells(rr, cRecIngredient).Value
        If cPct > 0 Then lr.Range.Cells(1, cPct).Value = loRecipes.DataBodyRange.Cells(rr, cRecPercent).Value
        If cUomLine > 0 Then lr.Range.Cells(1, cUomLine).Value = loRecipes.DataBodyRange.Cells(rr, cRecUom).Value
        If cAmt > 0 Then lr.Range.Cells(1, cAmt).Value = loRecipes.DataBodyRange.Cells(rr, cRecAmount).Value
        If cListRow > 0 Then lr.Range.Cells(1, cListRow).Value = loRecipes.DataBodyRange.Cells(rr, cRecListRow).Value
        If cIngId > 0 Then lr.Range.Cells(1, cIngId).Value = loRecipes.DataBodyRange.Cells(rr, cRecIngId).Value
        If cGuidLine > 0 Then lr.Range.Cells(1, cGuidLine).Value = loRecipes.DataBodyRange.Cells(rr, cRecGuid).Value
    Next idx

    MsgBox "Loaded recipe '" & recipeName & "' (" & matches.Count & " lines).", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Load Recipe failed: " & Err.Description, vbCritical
End Sub

Private Sub EnsureTableHasRow(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Private Sub ClearListObjectData(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    On Error GoTo 0
End Sub

Private Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function
