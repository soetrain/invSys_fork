Attribute VB_Name = "mProduction"
' run "mProduction.InitializeProductionUI" in immediate window to clean up UI
Option Explicit

' Production system core module (wiring + helpers).

Private Const SHEET_PRODUCTION As String = "Production"
Private Const SHEET_TEMPLATES As String = "TemplatesTable"

Private Const TABLE_RECIPE_CHOOSER As String = "RC_RecipeChoose"
Private Const TABLE_RECIPE_CHOOSER_GENERATED As String = "RecipeChooser_generated"
Private Const TABLE_INV_PALETTE_GENERATED As String = "InventoryPalette_generated"
' System 1: Recipe List Builder tables.
Private Const TABLE_RECIPE_BUILDER_HEADER As String = "RB_AddRecipeName"
Private Const TABLE_RECIPE_BUILDER_LINES As String = "RecipeBuilder"

Private Const BTN_HIDE_SYSTEM As String = "BTN_HIDE_SYSTEM"
Private Const BTN_SHOW_SYSTEM As String = "BTN_SHOW_SYSTEM"
Private Const BTN_LOAD_RECIPE As String = "BTN_LOAD_RECIPE"             ' System 1: Recipe List Builder
Private Const BTN_SAVE_RECIPE As String = "BTN_SAVE_RECIPE"             ' System 1: Recipe List Builder
Private Const BTN_BUILD_RECIPE_TABLES As String = "BTN_BUILD_RECIPE_TABLES" ' System 1: Recipe List Builder
Private Const BTN_REMOVE_RECIPE_TABLES As String = "BTN_REMOVE_RECIPE_TABLES" ' System 1: Recipe List Builder
Private Const BTN_CLEAR_RECIPE_BUILDER As String = "BTN_CLEAR_RECIPE_BUILDER" ' System 1: Recipe List Builder
Private Const BTN_CLEAR_RECIPE_CHOOSER As String = "BTN_CLEAR_RECIPE_CHOOSER" ' System 3: Recipe Chooser
Private Const BTN_CLEAR_PALETTE_BUILDER As String = "BTN_CLEAR_PALETTE_BUILDER" ' System 2: Inventory Palette Builder
Private Const BTN_SAVE_PALETTE As String = "BTN_SAVE_PALETTE"
Private Const BTN_TO_USED As String = "BTN_TO_USED"
Private Const BTN_TO_MADE As String = "BTN_TO_MADE"
Private Const BTN_TO_TOTALINV As String = "BTN_TO_TOTALINV"
Private Const BTN_NEXT_BATCH As String = "BTN_NEXT_BATCH"
Private Const BTN_PRINT_CODES As String = "BTN_PRINT_CODES"

Private Const TEMPLATE_SCOPE_RECIPE_PROCESS As String = "RECIPE_PROCESS"
Private Const RECIPE_PROC_TABLE_SUFFIX As String = "rbuilder"
Private Const RECIPE_CHOOSER_TABLE_SUFFIX As String = "rchooser"
Private Const RECIPE_LINES_STAGING_ROW As Long = 500000 ' System 1: staging for RecipeBuilder lines during load
Private Const PALETTE_LINES_STAGING_ROW As Long = 500000 ' System 4: staging for InventoryPalette lines table

Private mRowCountCache As Object
Private mPaletteTableMeta As Object
Private mHiddenSystems As Collection
Private mRecipePicker As cDynItemSearch
Private mPickerRouter As cPickerRouter
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
    EnsurePickerRouter
    mPickerRouter.HandleSelectionChange Target
End Sub

Public Sub HandleProductionBeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    If Target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(Target) Then Exit Sub
    EnsurePickerRouter
    If mPickerRouter.HandleBeforeDoubleClick(Target, Cancel) Then Exit Sub
End Sub

Private Sub EnsurePickerRouter()
    If mPickerRouter Is Nothing Then Set mPickerRouter = New cPickerRouter
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
        If Not mRowCountCache.Exists(key) Then
            mRowCountCache(key) = newCount
            Exit Sub
        End If
        Dim oldCount As Long: oldCount = CLng(mRowCountCache(key))
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

Private Sub EnsurePaletteTableMeta()
    If mPaletteTableMeta Is Nothing Then
        Set mPaletteTableMeta = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Sub ClearPaletteTableMeta()
    If Not mPaletteTableMeta Is Nothing Then mPaletteTableMeta.RemoveAll
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

Public Function LoadRecipeList() As Variant
    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then Exit Function
    Dim lo As ListObject: Set lo = GetListObject(wsRec, "Recipes")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim cId As Long: cId = ColumnIndex(lo, "RECIPE_ID")
    Dim cName As Long: cName = ColumnIndex(lo, "RECIPE")
    Dim cDesc As Long: cDesc = ColumnIndex(lo, "DESCRIPTION")
    If cId = 0 Or cName = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = lo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rid As String: rid = NzStr(arr(r, cId))
        Dim rname As String: rname = NzStr(arr(r, cName))
        If rid = "" Or rname = "" Then GoTo NextRow
        If Not dict.Exists(rid) Then
            Dim info(1 To 3) As Variant
            info(1) = rid
            info(2) = rname
            If cDesc > 0 Then info(3) = NzStr(arr(r, cDesc)) Else info(3) = ""
            dict.Add rid, info
        End If
NextRow:
    Next r

    If dict.Count = 0 Then Exit Function
    Dim result() As Variant
    ReDim result(1 To dict.Count, 1 To 3)
    Dim i As Long: i = 1
    Dim key As Variant
    For Each key In dict.Keys
        Dim infoArr As Variant
        infoArr = dict(key)
        result(i, 1) = infoArr(1)
        result(i, 2) = infoArr(2)
        result(i, 3) = infoArr(3)
        i = i + 1
    Next key
    LoadRecipeList = result
End Function

' ===== System 3: Recipe Chooser =====
Public Sub LoadRecipeChooser(ByVal recipeId As String)
    On Error GoTo ErrHandler
    If Trim$(recipeId) = "" Then Exit Sub

    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then
        MsgBox "Recipes sheet not found.", vbCritical
        Exit Sub
    End If

    Dim loChooser As ListObject
    Set loChooser = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If loChooser Is Nothing Then
        MsgBox "RC_RecipeChoose table not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    EnsureTableHasRow loChooser

    Dim recipeName As String
    Dim recipeDesc As String
    Dim recipeDept As String
    GetRecipeSummary wsRec, recipeId, recipeName, recipeDesc, recipeDept

    Dim cRec As Long: cRec = ColumnIndex(loChooser, "RECIPE")
    If cRec = 0 Then cRec = ColumnIndex(loChooser, "RECIPE_NAME")
    Dim cRecId As Long: cRecId = ColumnIndex(loChooser, "RECIPE_ID")
    Dim cDesc As Long: cDesc = ColumnIndex(loChooser, "DESCRIPTION")
    Dim cDept As Long: cDept = ColumnIndex(loChooser, "DEPARTMENT")
    Dim cProc As Long: cProc = ColumnIndex(loChooser, "PROCESS")

    If Not loChooser.DataBodyRange Is Nothing Then
        If cRec > 0 Then loChooser.DataBodyRange.Cells(1, cRec).Value = recipeName
        If cRecId > 0 Then loChooser.DataBodyRange.Cells(1, cRecId).Value = recipeId
        If cDesc > 0 Then loChooser.DataBodyRange.Cells(1, cDesc).Value = recipeDesc
        If cDept > 0 Then loChooser.DataBodyRange.Cells(1, cDept).Value = recipeDept
        If cProc > 0 Then loChooser.DataBodyRange.Cells(1, cProc).Value = ""
    End If

    Dim chooserStyle As String
    Dim loStyle As ListObject
    Set loStyle = GetListObject(wsProd, TABLE_RECIPE_CHOOSER_GENERATED)
    If Not loStyle Is Nothing Then
        On Error Resume Next
        chooserStyle = loStyle.TableStyle
        On Error GoTo 0
    End If

    Dim paletteStyle As String
    Dim loPalette As ListObject
    Set loPalette = GetListObject(wsProd, TABLE_INV_PALETTE_GENERATED)
    If Not loPalette Is Nothing Then
        On Error Resume Next
        paletteStyle = loPalette.TableStyle
        On Error GoTo 0
    End If

    DeleteRecipeChooserProcessTables wsProd
    DeleteInventoryPaletteTables wsProd

    Dim procTables As Collection
    Set procTables = BuildRecipeChooserProcessTablesFromRecipes(recipeId, wsProd, wsRec, chooserStyle)
    BuildPaletteTablesForRecipeChooser recipeId, wsProd, wsRec, procTables, paletteStyle

    Exit Sub
ErrHandler:
    MsgBox "Load Recipe Chooser failed: " & Err.Description, vbCritical
End Sub

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
        If StrComp(Trim$(lc.Name), Trim$(colName), vbTextCompare) = 0 Then
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
    EnsureButtonCustom ws, BTN_BUILD_RECIPE_TABLES, "Add Recipe Process Table", "mProduction.BtnBuildRecipeProcessTables", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_REMOVE_RECIPE_TABLES, "Remove Recipe Process Table", "mProduction.BtnRemoveRecipeProcessTables", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_CLEAR_RECIPE_BUILDER, "Clear Recipe List Builder", "mProduction.BtnClearRecipeBuilder", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_SAVE_PALETTE, "Save IngredientPalette", "mProduction.BtnSavePalette", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_CLEAR_PALETTE_BUILDER, "Clear Inventory Palette Builder", "mProduction.BtnClearPaletteBuilder", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_CLEAR_RECIPE_CHOOSER, "Clear Chosen Recipe", "mProduction.BtnClearRecipeChooser", leftA, nextTop, colAWidth
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
    ' System 1: Recipe List Builder.
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
' System 1: Recipe List Builder actions (Load/Save/Add/Remove/Clear).
Public Sub BtnLoadRecipe()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim loHeader As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    If loHeader Is Nothing Then
        MsgBox "Recipe Builder header table not found on Production sheet.", vbExclamation
        Exit Sub
    End If
    Dim targetCell As Range
    Set targetCell = GetHeaderDataCell(loHeader, "RECIPE_NAME")
    If targetCell Is Nothing Then
        MsgBox "Recipe Builder header missing RECIPE_NAME column.", vbCritical
        Exit Sub
    End If
    If mRecipePicker Is Nothing Then Set mRecipePicker = New cDynItemSearch
    mRecipePicker.ShowForRecipeCell targetCell
End Sub

Public Sub BtnSaveRecipe()
    SaveRecipeToRecipes
End Sub

Public Sub BtnBuildRecipeProcessTables()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim loHeader As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Dim procTables As Collection
    Set procTables = GetRecipeBuilderProcessTables(wsProd)
    Dim recipeId As String
    If Not loHeader Is Nothing Then
        Dim idCell As Range: Set idCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
        If Not idCell Is Nothing Then recipeId = NzStr(idCell.Value)
    End If
    Dim builtCount As Long
    If procTables.Count = 0 Then
        builtCount = BuildRecipeProcessTablesFromLines(recipeId, True)
    End If
    If builtCount = 0 Then
        Dim newLo As ListObject
        Set newLo = CreateRecipeProcessTable(wsProd, "", 1)
        If newLo Is Nothing Then
            MsgBox "No PROCESS rows found to build process tables.", vbInformation
        Else
            FocusRecipeProcessTable newLo
            MsgBox "Created process table '" & newLo.Name & "'.", vbInformation
        End If
    End If
End Sub

Public Sub BtnRemoveRecipeProcessTables()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim sel As Range
    On Error Resume Next
    Set sel = Application.Selection
    On Error GoTo 0
    If sel Is Nothing Then
        MsgBox "Select one or more Recipe Process tables to remove.", vbInformation
        Exit Sub
    End If

    Dim targets As Object: Set targets = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject
    For Each lo In wsProd.ListObjects
        If IsRecipeProcessTable(lo) Then
            If Not Intersect(lo.Range, sel) Is Nothing Then
                targets(lo.Name) = lo.Range.Address
            End If
        End If
    Next lo

    If targets.Count = 0 Then
        MsgBox "No Recipe Process tables selected.", vbInformation
        Exit Sub
    End If

    Dim key As Variant
    For Each key In targets.Keys
        On Error Resume Next
        wsProd.ListObjects(CStr(key)).Delete
        wsProd.Range(CStr(targets(key))).Clear
        On Error GoTo 0
    Next key

    MsgBox "Removed " & targets.Count & " Recipe Process table(s).", vbInformation
End Sub

' System 1: Recipe List Builder actions (Load/Save/Add/Remove).
Public Sub BtnClearRecipeBuilder()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim loHeader As ListObject
    Dim loLines As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Set loLines = GetRecipeBuilderLinesTable(wsProd, loHeader)

    DeleteRecipeProcessTables wsProd

    If Not loLines Is Nothing Then
        RemoveRecipeBuilderLinesTable loLines
    End If

    If Not loHeader Is Nothing Then
        ClearListObjectData loHeader
    End If

    MsgBox "Recipe List Builder cleared.", vbInformation
End Sub

' System 2+: Inventory Palette / Production actions.

Public Sub BtnSavePalette()
    SaveIngredientPalette
End Sub

Public Sub BtnClearPaletteBuilder()
    ClearInventoryPaletteBuilder
End Sub

Public Sub BtnClearRecipeChooser()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim loChooser As ListObject
    Set loChooser = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If Not loChooser Is Nothing Then
        EnsureTableHasRow loChooser
        If Not loChooser.DataBodyRange Is Nothing Then
            loChooser.DataBodyRange.ClearContents
        End If
    End If

    DeleteRecipeChooserProcessTables wsProd
    DeleteInventoryPaletteTables wsProd

    MsgBox "Recipe Chooser cleared.", vbInformation
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

' ===== System 2: Inventory Palette Builder =====
Private Sub SaveIngredientPalette()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim loRecipe As ListObject
    Dim loIng As ListObject
    Dim loItems As ListObject
    Set loRecipe = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseRecipe", Array("RECIPE_NAME", "RECIPE_ID"))
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    If loRecipe Is Nothing Or loIng Is Nothing Or loItems Is Nothing Then
        MsgBox "Inventory Palette Builder tables not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim recipeId As String: recipeId = GetPaletteRecipeId()
    If recipeId = "" Then
        MsgBox "Select a RECIPE in IP_ChooseRecipe before saving.", vbInformation
        Exit Sub
    End If

    Dim ingredientId As String: ingredientId = GetPaletteIngredientId()
    If ingredientId = "" Then
        MsgBox "Select an INGREDIENT in IP_ChooseIngredient before saving.", vbInformation
        Exit Sub
    End If

    If loItems.DataBodyRange Is Nothing Then
        MsgBox "Add at least one acceptable item before saving.", vbInformation
        Exit Sub
    End If

    Dim wsPal As Worksheet: Set wsPal = SheetExists("IngredientPalette")
    If wsPal Is Nothing Then
        MsgBox "IngredientPalette sheet not found.", vbCritical
        Exit Sub
    End If

    Dim loPal As ListObject
    Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "ITEM"))
    If loPal Is Nothing Then
        Set loPal = FindListObjectByNameOrHeaders(wsPal, "Table40", Array("RECIPE_ID", "INGREDIENT_ID", "ITEM"))
    End If
    If loPal Is Nothing Then
        MsgBox "IngredientPalette table not found on IngredientPalette sheet.", vbCritical
        Exit Sub
    End If

    Dim ioVal As String
    Dim pctVal As Variant
    Dim uomVal As String
    Dim amtVal As Variant
    FindRecipeIngredientInfo recipeId, ingredientId, ioVal, pctVal, uomVal, amtVal

    ' Remove existing palette rows for this recipe + ingredient.
    If Not loPal.DataBodyRange Is Nothing Then
        Dim cPalRec As Long: cPalRec = ColumnIndex(loPal, "RECIPE_ID")
        Dim cPalIng As Long: cPalIng = ColumnIndex(loPal, "INGREDIENT_ID")
        If cPalRec > 0 And cPalIng > 0 Then
            Dim r As Long
            For r = loPal.DataBodyRange.Rows.Count To 1 Step -1
                If NzStr(loPal.DataBodyRange.Cells(r, cPalRec).Value) = recipeId _
                   And NzStr(loPal.DataBodyRange.Cells(r, cPalIng).Value) = ingredientId Then
                    loPal.ListRows(r).Delete
                End If
            Next r
        End If
    End If

    Dim cItem As Long: cItem = ColumnIndex(loItems, "ITEMS")
    If cItem = 0 Then cItem = ColumnIndex(loItems, "ITEM")
    Dim cUom As Long: cUom = ColumnIndex(loItems, "UOM")
    Dim cRow As Long: cRow = ColumnIndex(loItems, "ROW")

    Dim cOutRec As Long: cOutRec = ColumnIndex(loPal, "RECIPE_ID")
    Dim cOutIng As Long: cOutIng = ColumnIndex(loPal, "INGREDIENT_ID")
    Dim cOutIO As Long: cOutIO = ColumnIndex(loPal, "INPUT/OUTPUT")
    Dim cOutItem As Long: cOutItem = ColumnIndex(loPal, "ITEM")
    Dim cOutPct As Long: cOutPct = ColumnIndex(loPal, "PERCENT")
    Dim cOutUom As Long: cOutUom = ColumnIndex(loPal, "UOM")
    Dim cOutAmt As Long: cOutAmt = ColumnIndex(loPal, "AMOUNT")
    Dim cOutRow As Long: cOutRow = ColumnIndex(loPal, "ROW")
    Dim cOutGuid As Long: cOutGuid = ColumnIndex(loPal, "GUID")

    Dim added As Long
    Dim arr As Variant: arr = loItems.DataBodyRange.Value
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        Dim itemVal As String
        If cItem > 0 Then itemVal = NzStr(arr(i, cItem))
        If Trim$(itemVal) = "" Then GoTo NextItem

        Dim lr As ListRow: Set lr = loPal.ListRows.Add
        If cOutRec > 0 Then lr.Range.Cells(1, cOutRec).Value = recipeId
        If cOutIng > 0 Then lr.Range.Cells(1, cOutIng).Value = ingredientId
        If cOutIO > 0 Then lr.Range.Cells(1, cOutIO).Value = ioVal
        If cOutItem > 0 Then lr.Range.Cells(1, cOutItem).Value = itemVal
        If cOutPct > 0 Then lr.Range.Cells(1, cOutPct).Value = pctVal
        If cOutUom > 0 Then
            Dim itemUom As String
            If cUom > 0 Then itemUom = NzStr(arr(i, cUom))
            If itemUom <> "" Then
                lr.Range.Cells(1, cOutUom).Value = itemUom
            Else
                lr.Range.Cells(1, cOutUom).Value = uomVal
            End If
        End If
        If cOutAmt > 0 Then lr.Range.Cells(1, cOutAmt).Value = amtVal
        If cOutRow > 0 And cRow > 0 Then lr.Range.Cells(1, cOutRow).Value = arr(i, cRow)
        If cOutGuid > 0 Then lr.Range.Cells(1, cOutGuid).Value = modUR_Snapshot.GenerateGUID()
        added = added + 1
NextItem:
    Next i

    MsgBox "Saved IngredientPalette rows: " & added & ".", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Save IngredientPalette failed: " & Err.Description, vbCritical
End Sub

Private Sub ClearInventoryPaletteBuilder()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim loRecipe As ListObject
    Dim loIng As ListObject
    Dim loItems As ListObject
    Set loRecipe = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseRecipe", Array("RECIPE_NAME", "RECIPE_ID"))
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))

    ResetPaletteTable loItems
    ResetPaletteTable loIng
    ResetPaletteTable loRecipe

    MsgBox "Inventory Palette Builder cleared.", vbInformation
End Sub

Public Sub HandlePaletteRecipeSelected(ByVal recipeId As String)
    ' System 2: Inventory Palette Builder - clear ingredient/items when recipe changes.
    If Trim$(recipeId) = "" Then Exit Sub
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim loIng As ListObject
    Dim loItems As ListObject
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    ResetPaletteTable loItems
    If Not loIng Is Nothing Then
        ResetPaletteTable loIng
        Dim cRec As Long: cRec = ColumnIndex(loIng, "RECIPE_ID")
        If cRec > 0 Then
            Dim recCell As Range
            Set recCell = GetHeaderDataCell(loIng, "RECIPE_ID")
            If Not recCell Is Nothing Then recCell.Value = recipeId
        End If
    End If
End Sub

Public Sub HandlePaletteIngredientSelected(ByVal recipeId As String, ByVal ingredientId As String)
    ' System 2: Inventory Palette Builder - clear items when ingredient changes.
    If Trim$(ingredientId) = "" Then Exit Sub
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim loItems As ListObject
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    If Not loItems Is Nothing Then
        ResetPaletteTable loItems
        Dim cRec As Long: cRec = ColumnIndex(loItems, "RECIPE_ID")
        Dim cIng As Long: cIng = ColumnIndex(loItems, "INGREDIENT_ID")
        If cRec > 0 Then
            Dim recCell As Range
            Set recCell = GetHeaderDataCell(loItems, "RECIPE_ID")
            If Not recCell Is Nothing Then recCell.Value = recipeId
        End If
        If cIng > 0 Then
            Dim ingCell As Range
            Set ingCell = GetHeaderDataCell(loItems, "INGREDIENT_ID")
            If Not ingCell Is Nothing Then ingCell.Value = ingredientId
        End If
    End If
End Sub

Public Function GetPaletteRecipeId() As String
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Function
    Dim loRecipe As ListObject
    Set loRecipe = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseRecipe", Array("RECIPE_NAME", "RECIPE_ID"))
    If loRecipe Is Nothing Then Exit Function
    GetPaletteRecipeId = FirstNonEmptyColumnValue(loRecipe, "RECIPE_ID")
End Function

Public Function GetPaletteIngredientId() As String
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Function
    Dim loIng As ListObject
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    If loIng Is Nothing Then Exit Function
    GetPaletteIngredientId = FirstNonEmptyColumnValue(loIng, "INGREDIENT_ID")
End Function

Public Function LoadIngredientListForRecipe(ByVal recipeId As String) As Variant
    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then Exit Function
    Dim lo As ListObject: Set lo = GetListObject(wsRec, "Recipes")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Trim$(recipeId) = "" Then Exit Function

    Dim cRecId As Long: cRecId = ColumnIndex(lo, "RECIPE_ID")
    Dim cIngId As Long: cIngId = ColumnIndex(lo, "INGREDIENT_ID")
    Dim cIng As Long: cIng = ColumnIndex(lo, "INGREDIENT")
    Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
    Dim cProc As Long: cProc = ColumnIndex(lo, "PROCESS")
    Dim cIO As Long: cIO = ColumnIndex(lo, "INPUT/OUTPUT")
    Dim cAmt As Long: cAmt = ColumnIndex(lo, "AMOUNT")
    Dim cPct As Long: cPct = ColumnIndex(lo, "PERCENT")
    If cRecId = 0 Or cIngId = 0 Or cIng = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = lo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            Dim key As String
            key = NzStr(arr(r, cIngId)) & "|" & NzStr(arr(r, cProc))
            If Not dict.Exists(key) Then
                Dim info(1 To 7) As Variant
                info(1) = NzStr(arr(r, cIngId))
                info(2) = NzStr(arr(r, cIng))
                If cUom > 0 Then info(3) = NzStr(arr(r, cUom)) Else info(3) = ""
                If cProc > 0 Then info(4) = NzStr(arr(r, cProc)) Else info(4) = ""
                If cIO > 0 Then info(5) = NzStr(arr(r, cIO)) Else info(5) = ""
                If cAmt > 0 Then info(6) = arr(r, cAmt) Else info(6) = ""
                If cPct > 0 Then info(7) = arr(r, cPct) Else info(7) = ""
                dict.Add key, info
            End If
        End If
    Next r

    If dict.Count = 0 Then Exit Function
    Dim result() As Variant
    ReDim result(1 To dict.Count, 1 To 7)
    Dim i As Long: i = 1
    Dim k As Variant
    For Each k In dict.Keys
        Dim infoArr As Variant
        infoArr = dict(k)
        result(i, 1) = infoArr(1)
        result(i, 2) = infoArr(2)
        result(i, 3) = infoArr(3)
        result(i, 4) = infoArr(4)
        result(i, 5) = infoArr(5)
        result(i, 6) = infoArr(6)
        result(i, 7) = infoArr(7)
        i = i + 1
    Next k
    LoadIngredientListForRecipe = result
End Function

' ===== System 3: Recipe Chooser - data helpers =====
Private Sub GetRecipeSummary(ByVal wsRec As Worksheet, ByVal recipeId As String, _
    ByRef recipeName As String, ByRef recipeDesc As String, ByRef recipeDept As String)

    recipeName = ""
    recipeDesc = ""
    recipeDept = ""
    If wsRec Is Nothing Then Exit Sub

    Dim lo As ListObject: Set lo = GetListObject(wsRec, "Recipes")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cRecId As Long: cRecId = ColumnIndex(lo, "RECIPE_ID")
    Dim cRec As Long: cRec = ColumnIndex(lo, "RECIPE")
    Dim cDesc As Long: cDesc = ColumnIndex(lo, "DESCRIPTION")
    Dim cDept As Long: cDept = ColumnIndex(lo, "DEPARTMENT")
    If cRecId = 0 Or cRec = 0 Then Exit Sub

    Dim arr As Variant: arr = lo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            recipeName = NzStr(arr(r, cRec))
            If cDesc > 0 Then recipeDesc = NzStr(arr(r, cDesc))
            If cDept > 0 Then recipeDept = NzStr(arr(r, cDept))
            Exit Sub
        End If
    Next r
End Sub

Private Function BuildRecipeChooserProcessTablesFromRecipes(ByVal recipeId As String, _
    ByVal wsProd As Worksheet, ByVal wsRec As Worksheet, Optional ByVal baseStyle As String = "") As Collection

    Dim created As New Collection
    If wsProd Is Nothing Or wsRec Is Nothing Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If
    If Trim$(recipeId) = "" Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If
    If loRecipes.DataBodyRange Is Nothing Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim cRecId As Long: cRecId = ColumnIndex(loRecipes, "RECIPE_ID")
    Dim cProc As Long: cProc = ColumnIndex(loRecipes, "PROCESS")
    Dim cDiag As Long: cDiag = ColumnIndex(loRecipes, "DIAGRAM_ID")
    Dim cIO As Long: cIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cIng As Long: cIng = ColumnIndex(loRecipes, "INGREDIENT")
    Dim cPct As Long: cPct = ColumnIndex(loRecipes, "PERCENT")
    Dim cUom As Long: cUom = ColumnIndex(loRecipes, "UOM")
    Dim cAmt As Long: cAmt = ColumnIndex(loRecipes, "AMOUNT")
    Dim cIngId As Long: cIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cListRow As Long: cListRow = ColumnIndex(loRecipes, "RECIPE_LIST_ROW")

    If cRecId = 0 Or cProc = 0 Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim arr As Variant: arr = loRecipes.DataBodyRange.Value
    Dim procMap As Object: Set procMap = CreateObject("Scripting.Dictionary")
    Dim procOrder As Collection: Set procOrder = New Collection

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            Dim procName As String: procName = Trim$(NzStr(arr(r, cProc)))
            If procName <> "" Then
                If Not procMap.Exists(procName) Then
                    procMap.Add procName, New Collection
                    procOrder.Add procName
                End If
                procMap(procName).Add r
            End If
        End If
    Next r

    If procOrder.Count = 0 Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim startRow As Long, startCol As Long
    If Not GetRecipeChooserAnchor(wsProd, startRow, startCol) Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim headerNames As Variant
    headerNames = RecipeChooserHeaderList()
    Dim colCount As Long: colCount = UBound(headerNames) - LBound(headerNames) + 1

    Dim procKey As Variant
    Dim nextSeq As Long
    nextSeq = NextRecipeChooserSequence(wsProd)
    Dim idxProc As Long
    idxProc = 0
    For Each procKey In procOrder
        idxProc = idxProc + 1
        Dim rowsColl As Collection: Set rowsColl = procMap(procKey)
        Dim dataCount As Long: dataCount = rowsColl.Count
        If dataCount = 0 Then GoTo NextProc

        Dim tableRange As Range
        Set tableRange = wsProd.Range(wsProd.Cells(startRow, startCol), wsProd.Cells(startRow + dataCount, startCol + colCount - 1))
        If RangeHasListObjectCollisionStrict(wsProd, tableRange) Then
            Set tableRange = FindAvailableRecipeChooserRange(wsProd, startRow, startCol, dataCount + 1, colCount)
            If tableRange Is Nothing Then Exit For
        End If

        tableRange.Clear
        tableRange.Rows(1).Value = HeaderRowArray(headerNames)

        Dim dataArr() As Variant
        ReDim dataArr(1 To dataCount, 1 To colCount)
        Dim i As Long, c As Long
        For i = 1 To dataCount
            Dim srcRow As Long: srcRow = rowsColl(i)
            For c = 1 To colCount
                Dim hdr As String
                hdr = CStr(headerNames(LBound(headerNames) + c - 1))
                Select Case UCase$(hdr)
                    Case "PROCESS"
                        dataArr(i, c) = procKey
                    Case "DIAGRAM_ID"
                        If cDiag > 0 Then dataArr(i, c) = arr(srcRow, cDiag)
                    Case "INPUT/OUTPUT"
                        If cIO > 0 Then dataArr(i, c) = arr(srcRow, cIO)
                    Case "INGREDIENT"
                        If cIng > 0 Then dataArr(i, c) = arr(srcRow, cIng)
                    Case "PERCENT"
                        If cPct > 0 Then dataArr(i, c) = arr(srcRow, cPct)
                    Case "UOM"
                        If cUom > 0 Then dataArr(i, c) = arr(srcRow, cUom)
                    Case "AMOUNT NEEDED"
                        If cAmt > 0 Then dataArr(i, c) = arr(srcRow, cAmt)
                    Case "INGREDIENT_ID"
                        If cIngId > 0 Then dataArr(i, c) = arr(srcRow, cIngId)
                    Case "RECIPE_LIST_ROW"
                        If cListRow > 0 Then dataArr(i, c) = arr(srcRow, cListRow)
                End Select
            Next c
        Next i

        tableRange.Offset(1, 0).Resize(dataCount, colCount).Value = dataArr

        Dim newLo As ListObject
        Set newLo = wsProd.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        If idxProc = 1 Then
            On Error Resume Next
            newLo.Name = TABLE_RECIPE_CHOOSER_GENERATED
            If Err.Number <> 0 Then
                newLo.Name = UniqueListObjectName(wsProd, TABLE_RECIPE_CHOOSER_GENERATED)
            End If
            Err.Clear
            On Error GoTo 0
        Else
            newLo.Name = UniqueListObjectName(wsProd, BuildRecipeChooserProcessTableName(CStr(nextSeq)))
            nextSeq = nextSeq + 1
        End If
        If baseStyle <> "" Then
            On Error Resume Next
            newLo.TableStyle = baseStyle
            On Error GoTo 0
        End If
        created.Add newLo

        startRow = tableRange.Row + tableRange.Rows.Count + 2 ' keep 2 blank rows
NextProc:
    Next procKey

    If created.Count > 0 Then
        Dim tpl As New cTemplateApplier
        Dim loProc As ListObject
        For Each loProc In created
            Dim procNameTpl As String: procNameTpl = ProcessNameFromTable(loProc)
            tpl.ApplyTemplates loProc, TEMPLATE_SCOPE_RECIPE_PROCESS, procNameTpl, ""
        Next loProc
    End If

    Set BuildRecipeChooserProcessTablesFromRecipes = created
End Function

Private Sub BuildPaletteTablesForRecipeChooser(ByVal recipeId As String, ByVal wsProd As Worksheet, ByVal wsRec As Worksheet, _
    ByVal procTables As Collection, Optional ByVal baseStyle As String = "")

    If wsProd Is Nothing Or wsRec Is Nothing Then Exit Sub
    If Trim$(recipeId) = "" Then Exit Sub

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then Exit Sub
    If loRecipes.DataBodyRange Is Nothing Then Exit Sub

    Dim cRecId As Long: cRecId = ColumnIndex(loRecipes, "RECIPE_ID")
    Dim cProc As Long: cProc = ColumnIndex(loRecipes, "PROCESS")
    Dim cIO As Long: cIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cIngId As Long: cIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cAmt As Long: cAmt = ColumnIndex(loRecipes, "AMOUNT")
    If cRecId = 0 Or cProc = 0 Or cIO = 0 Or cIngId = 0 Then Exit Sub

    Dim arr As Variant: arr = loRecipes.DataBodyRange.Value
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim entries As Collection: Set entries = New Collection

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            Dim ioVal As String: ioVal = UCase$(Trim$(NzStr(arr(r, cIO))))
            If ioVal = "USED" Then
                Dim ingId As String: ingId = NzStr(arr(r, cIngId))
                Dim procName As String: procName = NzStr(arr(r, cProc))
                If ingId <> "" And procName <> "" Then
                    Dim key As String: key = procName & "|" & ingId
                    Dim amtVal As Variant
                    If cAmt > 0 Then amtVal = arr(r, cAmt)
                    If Not seen.Exists(key) Then
                        Dim info(0 To 4) As Variant
                        info(0) = recipeId
                        info(1) = ingId
                        info(2) = amtVal
                        info(3) = procName
                        info(4) = "USED"
                        seen.Add key, info
                        entries.Add info
                    Else
                        If IsNumeric(amtVal) Then
                            Dim curInfo As Variant
                            curInfo = seen(key)
                            If IsNumeric(curInfo(2)) Then
                                curInfo(2) = CDbl(curInfo(2)) + CDbl(amtVal)
                                seen(key) = curInfo
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If entries.Count = 0 Then Exit Sub

    Dim startRow As Long, startCol As Long
    Dim anchorStyle As String
    If Not GetInventoryPaletteAnchor(wsProd, startRow, startCol, anchorStyle) Then Exit Sub
    If baseStyle = "" Then baseStyle = anchorStyle

    EnsurePaletteTableMeta
    ClearPaletteTableMeta

    EnsureInventoryPaletteLinesTable wsProd, baseStyle

    Dim headerNames As Variant
    headerNames = InventoryPaletteHeaderList()
    Dim colCount As Long: colCount = UBound(headerNames) - LBound(headerNames) + 1

    Dim idx As Long
    Dim nextSeq As Long: nextSeq = 1
    For idx = 1 To entries.Count
        Dim infoArr As Variant
        infoArr = entries(idx)

        Dim tableRange As Range
        Set tableRange = wsProd.Range(wsProd.Cells(startRow, startCol), wsProd.Cells(startRow + 1, startCol + colCount - 1))
        If RangeHasListObjectCollisionStrict(wsProd, tableRange) Then
            Set tableRange = FindAvailablePaletteRange(wsProd, startRow, startCol, 2, colCount)
            If tableRange Is Nothing Then Exit For
        End If

        tableRange.Clear
        tableRange.Rows(1).Value = HeaderRowArray(headerNames)

        Dim newLo As ListObject
        Set newLo = wsProd.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        newLo.Name = UniqueListObjectName(wsProd, "proc_" & CStr(nextSeq) & "_palette")
        nextSeq = nextSeq + 1
        If baseStyle <> "" Then
            On Error Resume Next
            newLo.TableStyle = baseStyle
            On Error GoTo 0
        End If

        mPaletteTableMeta(newLo.Name) = infoArr

        If Not newLo.DataBodyRange Is Nothing Then
            Dim cProcColFill As Long: cProcColFill = ColumnIndex(newLo, "PROCESS")
            Dim cIOColFill As Long: cIOColFill = ColumnIndex(newLo, "INPUT/OUTPUT")
            Dim cQtyColFill As Long: cQtyColFill = ColumnIndex(newLo, "QUANTITY")
            If cProcColFill > 0 Then newLo.DataBodyRange.Cells(1, cProcColFill).Value = NzStr(infoArr(3))
            If cIOColFill > 0 Then newLo.DataBodyRange.Cells(1, cIOColFill).Value = NzStr(infoArr(4))
            If cQtyColFill > 0 Then newLo.DataBodyRange.Cells(1, cQtyColFill).Value = infoArr(2)
        End If

        startRow = tableRange.Row + tableRange.Rows.Count + 2
    Next idx
End Sub

Private Sub DeleteRecipeChooserProcessTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.Count To 1 Step -1
        Dim lo As ListObject
        Set lo = ws.ListObjects(i)
        If IsRecipeChooserProcessTable(lo) Or LCase$(lo.Name) = LCase$(TABLE_RECIPE_CHOOSER_GENERATED) Then
            Dim addr As String: addr = lo.Range.Address
            On Error Resume Next
            lo.Delete
            ws.Range(addr).Clear
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function IsRecipeChooserProcessTable(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    If Right$(nm, Len(RECIPE_CHOOSER_TABLE_SUFFIX) + 1) = "_" & LCase$(RECIPE_CHOOSER_TABLE_SUFFIX) Then
        IsRecipeChooserProcessTable = True
    End If
End Function

Private Function BuildRecipeChooserProcessTableName(ByVal processKey As String) As String
    Dim key As String: key = Trim$(processKey)
    If key <> "" And IsNumeric(key) Then
        BuildRecipeChooserProcessTableName = "proc_" & CLng(key) & "_" & RECIPE_CHOOSER_TABLE_SUFFIX
    Else
        key = SafeProcessKey(processKey)
        BuildRecipeChooserProcessTableName = "proc_" & key & "_" & RECIPE_CHOOSER_TABLE_SUFFIX
    End If
End Function

Private Function NextRecipeChooserSequence(ByVal ws As Worksheet) As Long
    Dim maxSeq As Long
    If ws Is Nothing Then
        NextRecipeChooserSequence = 1
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeChooserProcessTable(lo) Then
            Dim seq As Long
            seq = RecipeChooserSequenceFromName(lo.Name)
            If seq > maxSeq Then maxSeq = seq
        End If
    Next lo
    NextRecipeChooserSequence = maxSeq + 1
End Function

Private Function RecipeChooserSequenceFromName(ByVal tableName As String) As Long
    Dim nm As String: nm = LCase$(tableName)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    If Right$(nm, Len(RECIPE_CHOOSER_TABLE_SUFFIX) + 1) <> "_" & LCase$(RECIPE_CHOOSER_TABLE_SUFFIX) Then Exit Function
    Dim core As String
    core = Mid$(nm, 6, Len(nm) - 5 - (Len(RECIPE_CHOOSER_TABLE_SUFFIX) + 1))
    If core = "" Then Exit Function
    If Left$(core, 2) = "p_" Then core = Mid$(core, 3)
    RecipeChooserSequenceFromName = CLng(Val(core))
End Function

Private Function RecipeChooserHeaderList() As Variant
    RecipeChooserHeaderList = Array( _
        "PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", _
        "AMOUNT NEEDED", "INGREDIENT_ID", "RECIPE_LIST_ROW")
End Function

Private Function GetRecipeChooserAnchor(ByVal ws As Worksheet, ByRef startRow As Long, ByRef startCol As Long) As Boolean
    GetRecipeChooserAnchor = False
    If ws Is Nothing Then Exit Function
    Dim loChooser As ListObject
    Set loChooser = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If loChooser Is Nothing Then Exit Function

    startCol = loChooser.Range.Column
    startRow = loChooser.Range.Row + loChooser.Range.Rows.Count + 2 ' one blank row
    If startRow > 0 And startCol > 0 Then GetRecipeChooserAnchor = True
End Function

Private Function FindAvailableRecipeChooserRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
    ByVal totalRows As Long, ByVal totalCols As Long) As Range

    If ws Is Nothing Then Exit Function
    If totalRows < 1 Or totalCols < 1 Then Exit Function
    If startRow < 1 Then startRow = 1
    If startCol < 1 Then startCol = 1

    Dim maxRow As Long: maxRow = ws.Rows.Count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range
    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate) Then
            Set FindAvailableRecipeChooserRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 2
    Loop
End Function

Private Sub DeleteInventoryPaletteTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.Count To 1 Step -1
        Dim lo As ListObject
        Set lo = ws.ListObjects(i)
        If lo Is Nothing Then GoTo NextLo
        If LCase$(lo.Name) = LCase$(TABLE_INV_PALETTE_GENERATED) Or LCase$(lo.Name) Like "proc_*_palette" Then
            Dim addr As String: addr = lo.Range.Address
            On Error Resume Next
            lo.Delete
            ws.Range(addr).Clear
            On Error GoTo 0
        End If
NextLo:
    Next i
    ClearPaletteTableMeta
End Sub

Private Function InventoryPaletteHeaderList() As Variant
    InventoryPaletteHeaderList = Array( _
        "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", _
        "QUANTITY", "PROCESS", "LOCATION", "ROW", "INPUT/OUTPUT")
End Function

Private Function GetInventoryPaletteAnchor(ByVal ws As Worksheet, ByRef startRow As Long, ByRef startCol As Long, ByRef baseStyle As String) As Boolean
    GetInventoryPaletteAnchor = False
    If ws Is Nothing Then Exit Function
    Dim lo As ListObject
    Set lo = GetListObject(ws, TABLE_INV_PALETTE_GENERATED)
    If lo Is Nothing Then
        Set lo = FindListObjectByNameOrHeaders(ws, TABLE_INV_PALETTE_GENERATED, Array("ITEM_CODE", "ITEM", "ROW"))
    End If
    If Not lo Is Nothing Then
        On Error Resume Next
        baseStyle = lo.TableStyle
        On Error GoTo 0
        If lo.Range.Row < PALETTE_LINES_STAGING_ROW Then
            startRow = lo.Range.Row
            startCol = lo.Range.Column
            GetInventoryPaletteAnchor = True
            Exit Function
        End If
    End If

    Dim loProd As ListObject
    Dim loCheck As ListObject
    Set loProd = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    Set loCheck = FindListObjectByNameOrHeaders(ws, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loProd Is Nothing Then
        startCol = loProd.Range.Column
        Dim bottom As Long
        bottom = loProd.Range.Row + loProd.Range.Rows.Count - 1
        If Not loCheck Is Nothing Then
            Dim chkBottom As Long
            chkBottom = loCheck.Range.Row + loCheck.Range.Rows.Count - 1
            If chkBottom > bottom Then bottom = chkBottom
        End If
        startRow = bottom + 2
        GetInventoryPaletteAnchor = True
    End If
End Function

Private Function FindAvailablePaletteRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
    ByVal totalRows As Long, ByVal totalCols As Long) As Range

    If ws Is Nothing Then Exit Function
    If totalRows < 1 Or totalCols < 1 Then Exit Function
    If startRow < 1 Then startRow = 1
    If startCol < 1 Then startCol = 1

    Dim maxRow As Long: maxRow = ws.Rows.Count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range
    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate) Then
            Set FindAvailablePaletteRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 2
    Loop
End Function

Public Function GetPaletteTableContext(ByVal lo As ListObject, ByRef recipeId As String, ByRef ingredientId As String, _
    ByRef amount As Variant, ByRef procName As String, ByRef ioVal As String) As Boolean

    GetPaletteTableContext = False
    If lo Is Nothing Then Exit Function
    If mPaletteTableMeta Is Nothing Then Exit Function
    If Not mPaletteTableMeta.Exists(lo.Name) Then Exit Function

    Dim info As Variant
    info = mPaletteTableMeta(lo.Name)
    recipeId = NzStr(info(0))
    ingredientId = NzStr(info(1))
    amount = info(2)
    procName = NzStr(info(3))
    ioVal = NzStr(info(4))
    GetPaletteTableContext = True
End Function

Public Function GetAllowedInvRowsForIngredient(ByVal recipeId As String, ByVal ingredientId As String) As Object
    Set GetAllowedInvRowsForIngredient = Nothing
    If Trim$(recipeId) = "" Or Trim$(ingredientId) = "" Then Exit Function

    Dim wsPal As Worksheet: Set wsPal = SheetExists("IngredientPalette")
    If wsPal Is Nothing Then Exit Function

    Dim loPal As ListObject
    Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "ROW"))
    If loPal Is Nothing Then
        Set loPal = FindListObjectByNameOrHeaders(wsPal, "Table40", Array("RECIPE_ID", "INGREDIENT_ID", "ROW"))
    End If
    If loPal Is Nothing Then Exit Function
    If loPal.DataBodyRange Is Nothing Then Exit Function

    Dim cRec As Long: cRec = ColumnIndex(loPal, "RECIPE_ID")
    Dim cIng As Long: cIng = ColumnIndex(loPal, "INGREDIENT_ID")
    Dim cRow As Long: cRow = ColumnIndex(loPal, "ROW")
    If cRec = 0 Or cIng = 0 Or cRow = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loPal.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRec)) = recipeId And NzStr(arr(r, cIng)) = ingredientId Then
            Dim rowVal As String
            rowVal = NzStr(arr(r, cRow))
            If Trim$(rowVal) <> "" Then
                If Not dict.Exists(rowVal) Then dict.Add rowVal, True
            End If
        End If
    Next r

    If dict.Count = 0 Then Exit Function
    Set GetAllowedInvRowsForIngredient = dict
End Function

Private Sub FindRecipeIngredientInfo(ByVal recipeId As String, ByVal ingredientId As String, _
    ByRef ioVal As String, ByRef pctVal As Variant, ByRef uomVal As String, ByRef amtVal As Variant)

    ioVal = ""
    pctVal = ""
    uomVal = ""
    amtVal = ""

    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = GetListObject(wsRec, "Recipes")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cRecId As Long: cRecId = ColumnIndex(lo, "RECIPE_ID")
    Dim cIngId As Long: cIngId = ColumnIndex(lo, "INGREDIENT_ID")
    Dim cIO As Long: cIO = ColumnIndex(lo, "INPUT/OUTPUT")
    Dim cPct As Long: cPct = ColumnIndex(lo, "PERCENT")
    Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
    Dim cAmt As Long: cAmt = ColumnIndex(lo, "AMOUNT")
    If cRecId = 0 Or cIngId = 0 Then Exit Sub

    Dim arr As Variant: arr = lo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId And NzStr(arr(r, cIngId)) = ingredientId Then
            If cIO > 0 Then ioVal = NzStr(arr(r, cIO))
            If cPct > 0 Then pctVal = arr(r, cPct)
            If cUom > 0 Then uomVal = NzStr(arr(r, cUom))
            If cAmt > 0 Then amtVal = arr(r, cAmt)
            Exit Sub
        End If
    Next r
End Sub

Private Function FirstNonEmptyColumnValue(ByVal lo As ListObject, ByVal colName As String) As String
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim c As Long: c = ColumnIndex(lo, colName)
    If c = 0 Then Exit Function
    Dim arr As Variant: arr = lo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Trim$(NzStr(arr(r, c))) <> "" Then
            FirstNonEmptyColumnValue = NzStr(arr(r, c))
            Exit Function
        End If
    Next r
End Function

' ===== System 1: Recipe List Builder - Load / Save =====
' System 1: Recipe List Builder - save recipe to Recipes sheet.
Private Sub SaveRecipeToRecipes()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then
        MsgBox "Recipes sheet not found.", vbCritical
        Exit Sub
    End If

    Dim loHeader As ListObject
    Dim loLines As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Set loLines = GetRecipeBuilderLinesTable(wsProd, loHeader)
    If loLines Is Nothing Then Set loLines = EnsureRecipeBuilderLinesTable(wsProd, loHeader)
    If loHeader Is Nothing Or loLines Is Nothing Then
        MsgBox "Recipe Builder tables not found on Production sheet.", vbExclamation
        Exit Sub
    End If
    Dim nameCell As Range: Set nameCell = GetHeaderDataCell(loHeader, "RECIPE_NAME")
    If nameCell Is Nothing Then
        MsgBox "Recipe Builder header missing RECIPE_NAME column.", vbCritical
        Exit Sub
    End If
    Dim processTables As Collection
    Set processTables = GetRecipeBuilderProcessTables(wsProd)
    Dim sourceTables As Collection
    Set sourceTables = New Collection
    If Not processTables Is Nothing Then
        Dim loProc As ListObject
        For Each loProc In processTables
            If Not loProc.DataBodyRange Is Nothing Then sourceTables.Add loProc
        Next loProc
    End If
    If sourceTables.Count = 0 Then
        If loLines.DataBodyRange Is Nothing Then
            MsgBox "Add at least one recipe line before saving.", vbExclamation
            Exit Sub
        End If
        sourceTables.Add loLines
    End If

    Dim cDesc As Long: cDesc = ColumnIndex(loHeader, "DESCRIPTION")
    Dim cGuid As Long: cGuid = ColumnIndex(loHeader, "GUID")
    Dim cRecipeId As Long: cRecipeId = ColumnIndex(loHeader, "RECIPE_ID")
    If cRecipeId = 0 Then
        MsgBox "Recipe Builder header missing RECIPE_NAME or RECIPE_ID.", vbCritical
        Exit Sub
    End If

    Dim recipeName As String: recipeName = NzStr(nameCell.Value)
    Dim recipeDesc As String
    If cDesc > 0 Then
        Dim descCell As Range
        Set descCell = GetHeaderDataCell(loHeader, "DESCRIPTION")
        If Not descCell Is Nothing Then recipeDesc = NzStr(descCell.Value)
    End If
    If Trim$(recipeName) = "" Then
        MsgBox "Fill RB_AddRecipeName (RECIPE_NAME) or load a recipe before saving.", vbExclamation
        Exit Sub
    End If

    Dim recipeIdCell As Range: Set recipeIdCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
    Dim recipeId As String: recipeId = NzStr(recipeIdCell.Value)
    If recipeId = "" Then
        recipeId = modUR_Snapshot.GenerateGUID()
        recipeIdCell.Value = recipeId
    End If
    If cGuid > 0 Then
        Dim recipeGuidCell As Range: Set recipeGuidCell = GetHeaderDataCell(loHeader, "GUID")
        Dim recipeGuid As String: recipeGuid = NzStr(recipeGuidCell.Value)
        If recipeGuid = "" Then
            recipeGuid = modUR_Snapshot.GenerateGUID()
            recipeGuidCell.Value = recipeGuid
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

    Dim savedCount As Long
    Dim seqRow As Long: seqRow = 1
    Dim src As Variant
    For Each src In sourceTables
        AppendRecipeRowsFromTable src, recipeId, recipeName, recipeDesc, loRecipes, _
            cRecRecipeId, cRecRecipe, cRecDesc, cRecDept, cRecProcess, cRecDiagram, cRecIO, _
            cRecIngredient, cRecPercent, cRecUom, cRecAmount, cRecListRow, cRecIngId, cRecGuid, _
            seqRow, savedCount
    Next src

    Dim templateCount As Long
    If Not processTables Is Nothing Then
        If processTables.Count > 0 Then templateCount = RegisterRecipeTemplates(recipeId, processTables)
    End If

    If savedCount = 0 Then
        MsgBox "No recipe lines with data were found to save.", vbExclamation
    Else
        Dim msg As String
        msg = "Saved recipe '" & recipeName & "' (" & savedCount & " lines)."
        If templateCount > 0 Then msg = msg & vbCrLf & "Templates saved: " & templateCount & "."
        MsgBox msg, vbInformation
    End If
    Exit Sub
ErrHandler:
    MsgBox "Save Recipe failed: " & Err.Description, vbCritical
End Sub

' System 1: Recipe List Builder - load recipe into builder tables.
Public Sub LoadRecipeFromRecipes(Optional ByVal forceRecipeId As String = "")
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then
        MsgBox "Recipes sheet not found.", vbCritical
        Exit Sub
    End If

    Dim loHeader As ListObject
    Dim loLines As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Set loLines = GetRecipeBuilderLinesTable(wsProd, loHeader)
    If loLines Is Nothing Then Set loLines = EnsureRecipeBuilderLinesTable(wsProd, loHeader)
    If loHeader Is Nothing Or loLines Is Nothing Then
        MsgBox "Recipe Builder tables not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim recipeId As String
    Dim recipeName As String
    recipeId = forceRecipeId

    If recipeId = "" Then
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
    End If

    If recipeId = "" Then
        Dim cHeaderRecipeIdTmp As Long: cHeaderRecipeIdTmp = ColumnIndex(loHeader, "RECIPE_ID")
    If cHeaderRecipeIdTmp > 0 Then
        Dim hdrRecipeIdCell As Range: Set hdrRecipeIdCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
        If Not hdrRecipeIdCell Is Nothing Then recipeId = NzStr(hdrRecipeIdCell.Value)
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
    Dim hdrNameCell As Range: Set hdrNameCell = GetHeaderDataCell(loHeader, "RECIPE_NAME")
    Dim hdrIdCell As Range: Set hdrIdCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
    Dim hdrDescCell As Range: Set hdrDescCell = GetHeaderDataCell(loHeader, "DESCRIPTION")
    Dim hdrGuidCell As Range: Set hdrGuidCell = GetHeaderDataCell(loHeader, "GUID")
    If Not hdrNameCell Is Nothing Then hdrNameCell.Value = recipeName
    If Not hdrIdCell Is Nothing Then hdrIdCell.Value = recipeId
    If Not hdrDescCell Is Nothing And cRecDesc > 0 Then
        hdrDescCell.Value = NzStr(loRecipes.DataBodyRange.Cells(matches(1), cRecDesc).Value)
    End If
    If Not hdrGuidCell Is Nothing And cRecGuid > 0 Then
        hdrGuidCell.Value = NzStr(loRecipes.DataBodyRange.Cells(matches(1), cRecGuid).Value)
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

    Dim procCount As Long
    Dim hasProc As Boolean
    hasProc = RecipeLinesHasProcess(loLines)
    If hasProc Then
        Dim staged As Boolean
        staged = MoveRecipeBuilderLinesToStaging(loLines)
        procCount = BuildRecipeProcessTablesFromLines(recipeId, True, Not staged)
        ' Keep RecipeBuilder lines table staged until Clear Recipe List Builder.
    End If

    Dim loadMsg As String
    loadMsg = "Loaded recipe '" & recipeName & "' (" & matches.Count & " lines)."
    If procCount > 0 Then loadMsg = loadMsg & vbCrLf & "Process tables built: " & procCount & "."
    MsgBox loadMsg, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Load Recipe failed: " & Err.Description, vbCritical
End Sub

' System 1: Recipe List Builder - write recipe rows to Recipes table.
Private Sub AppendRecipeRowsFromTable(ByVal loSource As ListObject, ByVal recipeId As String, _
    ByVal recipeName As String, ByVal recipeDesc As String, ByVal loRecipes As ListObject, _
    ByVal cRecRecipeId As Long, ByVal cRecRecipe As Long, ByVal cRecDesc As Long, ByVal cRecDept As Long, _
    ByVal cRecProcess As Long, ByVal cRecDiagram As Long, ByVal cRecIO As Long, ByVal cRecIngredient As Long, _
    ByVal cRecPercent As Long, ByVal cRecUom As Long, ByVal cRecAmount As Long, ByVal cRecListRow As Long, _
    ByVal cRecIngId As Long, ByVal cRecGuid As Long, ByRef seqRow As Long, ByRef savedCount As Long)

    If loSource Is Nothing Then Exit Sub
    If loSource.DataBodyRange Is Nothing Then Exit Sub

    Dim cProc As Long: cProc = ColumnIndex(loSource, "PROCESS")
    Dim cDiag As Long: cDiag = ColumnIndex(loSource, "DIAGRAM_ID")
    Dim cIO As Long: cIO = ColumnIndex(loSource, "INPUT/OUTPUT")
    Dim cIng As Long: cIng = ColumnIndex(loSource, "INGREDIENT")
    Dim cPct As Long: cPct = ColumnIndex(loSource, "PERCENT")
    Dim cUomLine As Long: cUomLine = ColumnIndex(loSource, "UOM")
    Dim cAmt As Long: cAmt = ColumnIndex(loSource, "AMOUNT")
    Dim cListRow As Long: cListRow = ColumnIndex(loSource, "RECIPE_LIST_ROW")
    Dim cIngId As Long: cIngId = ColumnIndex(loSource, "INGREDIENT_ID")
    Dim cGuidLine As Long: cGuidLine = ColumnIndex(loSource, "GUID")

    Dim lineArr As Variant: lineArr = loSource.DataBodyRange.Value
    Dim rowCount As Long: rowCount = UBound(lineArr, 1)
    Dim processFallback As String: processFallback = ProcessNameFromTable(loSource)

    Dim i As Long
    For i = 1 To rowCount
        Dim hasData As Boolean
        If cIng > 0 Then
            hasData = (Trim$(NzStr(lineArr(i, cIng))) <> "")
        ElseIf cProc > 0 Then
            hasData = (Trim$(NzStr(lineArr(i, cProc))) <> "")
        End If
        If Not hasData Then GoTo NextLine

        Dim processVal As String
        If cProc > 0 Then processVal = NzStr(lineArr(i, cProc))
        If processVal = "" Then processVal = processFallback

        Dim ingId As String
        If cIngId > 0 Then ingId = NzStr(lineArr(i, cIngId))
        If ingId = "" Then
            ingId = modUR_Snapshot.GenerateGUID()
            If cIngId > 0 Then loSource.DataBodyRange.Cells(i, cIngId).Value = ingId
        End If

        Dim recListRow As Variant
        If cListRow > 0 Then recListRow = lineArr(i, cListRow)
        If NzStr(recListRow) = "" Then
            recListRow = seqRow
            If cListRow > 0 Then loSource.DataBodyRange.Cells(i, cListRow).Value = recListRow
        End If

        Dim rowGuid As String
        If cGuidLine > 0 Then rowGuid = NzStr(lineArr(i, cGuidLine))
        If rowGuid = "" Then
            rowGuid = modUR_Snapshot.GenerateGUID()
            If cGuidLine > 0 Then loSource.DataBodyRange.Cells(i, cGuidLine).Value = rowGuid
        End If

        Dim lr As ListRow: Set lr = loRecipes.ListRows.Add
        If cRecRecipeId > 0 Then lr.Range.Cells(1, cRecRecipeId).Value = recipeId
        If cRecRecipe > 0 Then lr.Range.Cells(1, cRecRecipe).Value = recipeName
        If cRecDesc > 0 Then lr.Range.Cells(1, cRecDesc).Value = recipeDesc
        If cRecDept > 0 Then lr.Range.Cells(1, cRecDept).Value = "" ' optional for now
        If cRecProcess > 0 Then lr.Range.Cells(1, cRecProcess).Value = processVal
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
End Sub

Private Function BuildRecipeProcessTablesFromLines(ByVal recipeId As String, Optional ByVal applyTemplates As Boolean = False, Optional ByVal anchorBelowLines As Boolean = True) As Long
    ' System 1: Recipe List Builder - build process tables under RB_AddRecipeName.
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Function

    Dim loLines As ListObject
    Set loLines = GetRecipeBuilderLinesTable(wsProd)
    If loLines Is Nothing Then
        MsgBox "Recipe Builder lines table not found on Production sheet.", vbExclamation
        Exit Function
    End If
    If loLines.DataBodyRange Is Nothing Then Exit Function

    Dim cProc As Long: cProc = ColumnIndex(loLines, "PROCESS")
    If cProc = 0 Then
        MsgBox "Recipe Builder lines missing PROCESS column.", vbCritical
        Exit Function
    End If

    Dim startRow As Long
    Dim startCol As Long
    Dim includeLines As Boolean
    includeLines = anchorBelowLines
    If includeLines Then
        If IsRecipeLinesStaged(loLines) Then includeLines = False
    End If
    If Not GetRecipeBuilderAnchor(wsProd, startRow, startCol, includeLines) Then
        MsgBox "Recipe Builder header table (RB_AddRecipeName) not found on Production sheet.", vbExclamation
        Exit Function
    End If

    Dim headerNames As Variant
    headerNames = RecipeProcessHeaderList()
    Dim colCount As Long: colCount = UBound(headerNames) - LBound(headerNames) + 1
    Dim srcIdx() As Long
    ReDim srcIdx(1 To colCount)
    Dim c As Long
    For c = 1 To colCount
        srcIdx(c) = ColumnIndex(loLines, CStr(headerNames(LBound(headerNames) + c - 1)))
    Next c

    Dim lineArr As Variant: lineArr = loLines.DataBodyRange.Value
    Dim procMap As Object: Set procMap = CreateObject("Scripting.Dictionary")
    Dim procOrder As Collection: Set procOrder = New Collection

    Dim r As Long
    For r = 1 To UBound(lineArr, 1)
        Dim procName As String: procName = Trim$(NzStr(lineArr(r, cProc)))
        If procName <> "" Then
            If Not procMap.Exists(procName) Then
                procMap.Add procName, New Collection
                procOrder.Add procName
            End If
            procMap(procName).Add r
        End If
    Next r

    If procOrder.Count = 0 Then Exit Function

    DeleteRecipeProcessTables wsProd

    Dim created As New Collection

    Dim procKey As Variant
    Dim nextSeq As Long
    nextSeq = NextRecipeProcessSequence(wsProd)
    For Each procKey In procOrder
        Dim rowsColl As Collection: Set rowsColl = procMap(procKey)
        Dim dataCount As Long: dataCount = rowsColl.Count
        If dataCount = 0 Then GoTo NextProc

        Dim tableRange As Range
        Set tableRange = wsProd.Range(wsProd.Cells(startRow, startCol), wsProd.Cells(startRow + dataCount, startCol + colCount - 1))

        If RangeHasListObjectCollision(wsProd, tableRange, loLines) Then
            MsgBox "Not enough space below Recipe Builder to create process tables. Clear space and try again.", vbExclamation
            Exit Function
        End If

        tableRange.Clear
        tableRange.Rows(1).Value = HeaderRowArray(headerNames)

        Dim dataArr() As Variant
        ReDim dataArr(1 To dataCount, 1 To colCount)
        Dim i As Long
        For i = 1 To dataCount
            Dim srcRow As Long: srcRow = rowsColl(i)
            For c = 1 To colCount
                Dim hdrName As String
                hdrName = CStr(headerNames(LBound(headerNames) + c - 1))
                If StrComp(hdrName, "PROCESS", vbTextCompare) = 0 Then
                    dataArr(i, c) = procKey
                ElseIf srcIdx(c) > 0 Then
                    dataArr(i, c) = lineArr(srcRow, srcIdx(c))
                End If
            Next c
        Next i

        tableRange.Offset(1, 0).Resize(dataCount, colCount).Value = dataArr

        Dim newLo As ListObject
        Set newLo = wsProd.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        newLo.Name = UniqueListObjectName(wsProd, BuildRecipeProcessTableName(CStr(nextSeq)))
        On Error Resume Next
        newLo.TableStyle = loLines.TableStyle
        On Error GoTo 0
        created.Add newLo
        nextSeq = nextSeq + 1

        startRow = startRow + dataCount + 3 ' keep 2 blank rows between process tables
NextProc:
    Next procKey

    BuildRecipeProcessTablesFromLines = created.Count

    If applyTemplates And created.Count > 0 And recipeId <> "" Then
        Dim tpl As New cTemplateApplier
        Dim loProc As ListObject
        For Each loProc In created
            Dim procNameTpl As String: procNameTpl = ProcessNameFromTable(loProc)
            tpl.ApplyTemplates loProc, TEMPLATE_SCOPE_RECIPE_PROCESS, procNameTpl, ""
        Next loProc
    End If
End Function

Private Function CreateRecipeProcessTable(ByVal ws As Worksheet, ByVal processName As String, Optional ByVal dataRows As Long = 1) As ListObject
    ' System 1: Recipe List Builder - add a blank process table under RB_AddRecipeName.
    If ws Is Nothing Then Exit Function
    If dataRows < 1 Then dataRows = 1

    Dim loLines As ListObject
    Set loLines = GetRecipeBuilderLinesTable(ws)
    If loLines Is Nothing Then Exit Function

    Dim headers As Variant
    headers = RecipeProcessHeaderList()
    Dim colCount As Long: colCount = UBound(headers) - LBound(headers) + 1
    Dim startRow As Long
    Dim startCol As Long
    Dim includeLines As Boolean
    includeLines = Not IsRecipeLinesStaged(loLines)
    If Not GetRecipeBuilderAnchor(ws, startRow, startCol, includeLines) Then Exit Function

    Dim tableRange As Range
    startRow = NextRecipeBuilderStartRow(ws, startRow)
    Set tableRange = FindAvailableRecipeProcessRange(ws, startRow, startCol, dataRows + 1, colCount, loLines)
    If tableRange Is Nothing Then Exit Function

    Dim seq As Long
    seq = NextRecipeProcessSequence(ws)
    If Trim$(processName) = "" Then processName = CStr(seq)

    tableRange.Clear
    tableRange.Rows(1).Value = HeaderRowArray(headers)

    Dim cProc As Long
    cProc = HeaderIndex(headers, "PROCESS")
    If cProc > 0 Then
        tableRange.Offset(1, cProc - 1).Value = processName
    End If

    Dim newLo As ListObject
    Set newLo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    newLo.Name = UniqueListObjectName(ws, BuildRecipeProcessTableName(CStr(seq)))
    On Error Resume Next
    newLo.TableStyle = loLines.TableStyle
    On Error GoTo 0

    FocusRecipeProcessTable newLo
    Set CreateRecipeProcessTable = newLo
End Function

Private Sub FocusRecipeProcessTable(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    lo.Parent.Activate
    Application.Goto lo.Range, True
    On Error GoTo 0
End Sub

Private Function NextRecipeBuilderStartRow(ByVal ws As Worksheet, ByVal baseRow As Long) As Long
    ' System 1: Recipe List Builder - stack new process tables below the last one.
    Dim startRow As Long
    startRow = baseRow
    If ws Is Nothing Then
        NextRecipeBuilderStartRow = startRow
        Exit Function
    End If

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeProcessTable(lo) Then
            Dim bottom As Long
            bottom = lo.Range.Row + lo.Range.Rows.Count - 1
            If bottom + 3 > startRow Then startRow = bottom + 3 ' keep 2 blank rows
        End If
    Next lo

    NextRecipeBuilderStartRow = startRow
End Function

Private Function GetRecipeBuilderAnchor(ByVal ws As Worksheet, ByRef startRow As Long, ByRef startCol As Long, Optional ByVal includeLines As Boolean = True) As Boolean
    ' System 1: Recipe List Builder anchor (under RB_AddRecipeName).
    GetRecipeBuilderAnchor = False
    If ws Is Nothing Then Exit Function
    Dim loHeader As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    If loHeader Is Nothing Then Exit Function

    startCol = loHeader.Range.Column
    startRow = loHeader.Range.Row + loHeader.Range.Rows.Count + 3 ' keep 2 blank rows before first process table

    If includeLines Then
        Dim loLines As ListObject
        Set loLines = GetRecipeBuilderLinesTable(ws, loHeader)
        If Not loLines Is Nothing Then
            Dim linesBottom As Long
            linesBottom = loLines.Range.Row + loLines.Range.Rows.Count - 1
            If linesBottom + 3 > startRow Then startRow = linesBottom + 3
        End If
    End If
    If startRow > 0 And startCol > 0 Then GetRecipeBuilderAnchor = True
End Function

Private Function EnsureRecipeBuilderLinesTable(ByVal ws As Worksheet, ByVal loHeader As ListObject) As ListObject
    ' System 1: Recipe List Builder - create RecipeBuilder lines table if missing.
    If ws Is Nothing Then Exit Function
    If loHeader Is Nothing Then Exit Function

    Dim existing As ListObject
    Set existing = GetRecipeBuilderLinesTable(ws, loHeader)
    If Not existing Is Nothing Then
        Set EnsureRecipeBuilderLinesTable = existing
        Exit Function
    End If

    Dim headers As Variant
    headers = RecipeProcessHeaderList()
    Dim colCount As Long: colCount = UBound(headers) - LBound(headers) + 1

    Dim startRow As Long
    Dim startCol As Long
    startRow = loHeader.Range.Row + loHeader.Range.Rows.Count + 2 ' one blank row below header
    startCol = loHeader.Range.Column

    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 1, startCol + colCount - 1))
    If RangeHasListObjectCollisionStrict(ws, tableRange, loHeader) Then Exit Function

    tableRange.Clear
    tableRange.Rows(1).Value = HeaderRowArray(headers)

    Dim newLo As ListObject
    Set newLo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    newLo.Name = UniqueListObjectName(ws, TABLE_RECIPE_BUILDER_LINES)
    On Error Resume Next
    newLo.TableStyle = loHeader.TableStyle
    On Error GoTo 0

    Set EnsureRecipeBuilderLinesTable = newLo
End Function

Private Function GetRecipeBuilderLinesTable(ByVal ws As Worksheet, Optional ByVal loHeader As ListObject) As ListObject
    ' System 1: Recipe List Builder - locate RecipeBuilder lines table under RB_AddRecipeName.
    If ws Is Nothing Then Exit Function

    Dim lo As ListObject
    Set lo = GetListObject(ws, TABLE_RECIPE_BUILDER_LINES)
    If Not lo Is Nothing Then
        Set GetRecipeBuilderLinesTable = lo
        Exit Function
    End If

    Dim headerStartCol As Long
    Dim headerBottom As Long
    If loHeader Is Nothing Then
        Set loHeader = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    End If
    If Not loHeader Is Nothing Then
        headerStartCol = loHeader.Range.Column
        headerBottom = loHeader.Range.Row + loHeader.Range.Rows.Count - 1
    End If

    Dim candidate As ListObject
    Dim bestRow As Long
    For Each lo In ws.ListObjects
        If ListObjectHasHeaders(lo, Array("PROCESS", "INGREDIENT")) Then
            If IsRecipeProcessTable(lo) Then GoTo NextLo
            If headerStartCol > 0 Then
                If lo.Range.Column <> headerStartCol Then GoTo NextLo
                If lo.Range.Row < headerBottom Then GoTo NextLo
            End If
            If bestRow = 0 Or lo.Range.Row < bestRow Then
                Set candidate = lo
                bestRow = lo.Range.Row
            End If
        End If
NextLo:
    Next lo

    If Not candidate Is Nothing Then
        Set GetRecipeBuilderLinesTable = candidate
        Exit Function
    End If

    If headerStartCol = 0 Then
        Set GetRecipeBuilderLinesTable = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_BUILDER_LINES, Array("PROCESS", "INGREDIENT"))
    End If
End Function

Private Function RecipeLinesHasProcess(ByVal loLines As ListObject) As Boolean
    ' System 1: Recipe List Builder - detect any PROCESS rows.
    If loLines Is Nothing Then Exit Function
    If loLines.DataBodyRange Is Nothing Then Exit Function
    Dim cProc As Long: cProc = ColumnIndex(loLines, "PROCESS")
    If cProc = 0 Then Exit Function
    Dim arr As Variant: arr = loLines.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Trim$(NzStr(arr(r, cProc))) <> "" Then
            RecipeLinesHasProcess = True
            Exit Function
        End If
    Next r
End Function

Private Function IsRecipeLinesStaged(ByVal loLines As ListObject) As Boolean
    ' System 1: Recipe List Builder - check if lines table is staged off-screen.
    If loLines Is Nothing Then Exit Function
    IsRecipeLinesStaged = (loLines.Range.Row >= RECIPE_LINES_STAGING_ROW)
End Function

Private Function MoveRecipeBuilderLinesToStaging(ByVal loLines As ListObject) As Boolean
    ' System 1: Recipe List Builder - move lines table out of view before building process tables.
    If loLines Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = loLines.Parent
    Dim startRow As Long: startRow = RECIPE_LINES_STAGING_ROW
    If loLines.Range.Row >= startRow Then
        MoveRecipeBuilderLinesToStaging = True
        Exit Function
    End If

    Dim dest As Range
    Set dest = ws.Cells(startRow, loLines.Range.Column)
    On Error Resume Next
    loLines.Range.Cut Destination:=dest
    MoveRecipeBuilderLinesToStaging = (Err.Number = 0)
    If MoveRecipeBuilderLinesToStaging Then
        On Error Resume Next
        loLines.Name = TABLE_RECIPE_BUILDER_LINES
        On Error GoTo 0
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function EnsureInventoryPaletteLinesTable(ByVal ws As Worksheet, Optional ByVal baseStyle As String = "") As ListObject
    ' System 4: Production Input/Output - keep InventoryPalette lines table staged off-screen.
    If ws Is Nothing Then Exit Function

    Dim lo As ListObject
    Set lo = GetListObject(ws, TABLE_INV_PALETTE_GENERATED)

    Dim startRow As Long
    Dim startCol As Long
    startRow = PALETTE_LINES_STAGING_ROW
    startCol = 1

    Dim loProd As ListObject
    Set loProd = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If Not loProd Is Nothing Then startCol = loProd.Range.Column

    If Not lo Is Nothing Then
        If lo.Range.Row < startRow Then
            Dim dest As Range
            Set dest = ws.Cells(startRow, startCol)
            On Error Resume Next
            lo.Range.Cut Destination:=dest
            On Error GoTo 0
        End If
        On Error Resume Next
        If baseStyle <> "" Then lo.TableStyle = baseStyle
        On Error GoTo 0
        Set EnsureInventoryPaletteLinesTable = lo
        Exit Function
    End If

    Dim headers As Variant
    headers = InventoryPaletteHeaderList()
    Dim colCount As Long: colCount = UBound(headers) - LBound(headers) + 1

    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 1, startCol + colCount - 1))
    If RangeHasListObjectCollisionStrict(ws, tableRange) Then Exit Function

    tableRange.Clear
    tableRange.Rows(1).Value = HeaderRowArray(headers)

    Dim newLo As ListObject
    Set newLo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    newLo.Name = TABLE_INV_PALETTE_GENERATED
    On Error Resume Next
    If baseStyle <> "" Then newLo.TableStyle = baseStyle
    On Error GoTo 0

    Set EnsureInventoryPaletteLinesTable = newLo
End Function

Private Function HeaderIndex(ByVal headers As Variant, ByVal headerName As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(i)), headerName, vbTextCompare) = 0 Then
            HeaderIndex = i - LBound(headers) + 1
            Exit Function
        End If
    Next i
End Function

Private Function FindAvailableRecipeProcessRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
    ByVal totalRows As Long, ByVal totalCols As Long, ByVal loLines As ListObject) As Range

    If ws Is Nothing Then Exit Function
    If totalRows < 1 Or totalCols < 1 Then Exit Function
    If startRow < 1 Then startRow = 1
    If startCol < 1 Then startCol = 1

    Dim maxRow As Long
    maxRow = ws.Rows.Count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range

    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate, loLines) Then
            Set FindAvailableRecipeProcessRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 2 ' keep 2 blank rows between tables
    Loop
End Function

Private Sub DeleteRecipeProcessTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.Count To 1 Step -1
        Dim lo As ListObject
        Set lo = ws.ListObjects(i)
        If IsRecipeProcessTable(lo) Then
            Dim addr As String
            addr = lo.Range.Address
            On Error Resume Next
            lo.Delete
            ws.Range(addr).Clear
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function GetRecipeBuilderProcessTables(ByVal ws As Worksheet) As Collection
    Dim result As New Collection
    If ws Is Nothing Then
        Set GetRecipeBuilderProcessTables = result
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeProcessTable(lo) Then result.Add lo
    Next lo
    Set GetRecipeBuilderProcessTables = result
End Function

Private Function IsRecipeProcessTable(ByVal lo As ListObject) As Boolean
    ' System 1: Recipe List Builder - identify process tables.
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    If Right$(nm, Len(RECIPE_PROC_TABLE_SUFFIX) + 1) = "_" & LCase$(RECIPE_PROC_TABLE_SUFFIX) Then
        IsRecipeProcessTable = True
    End If
End Function

' System 1: Recipe List Builder - register process formulas as templates.
Private Function RegisterRecipeTemplates(ByVal recipeId As String, ByVal processTables As Collection) As Long
    If processTables Is Nothing Then Exit Function
    If processTables.Count = 0 Then Exit Function

    Dim wsTpl As Worksheet: Set wsTpl = SheetExists(SHEET_TEMPLATES)
    If wsTpl Is Nothing Then Exit Function
    Dim loTpl As ListObject: Set loTpl = GetListObject(wsTpl, "TemplatesTable")
    If loTpl Is Nothing Then Exit Function

    Dim cGuid As Long: cGuid = ColumnIndex(loTpl, "GUID")
    Dim cScope As Long: cScope = ColumnIndex(loTpl, "TEMPLATE_SCOPE")
    Dim cRecipe As Long: cRecipe = ColumnIndex(loTpl, "RECIPE_ID")
    Dim cIngredient As Long: cIngredient = ColumnIndex(loTpl, "INGREDIENT_ID")
    Dim cProcess As Long: cProcess = ColumnIndex(loTpl, "PROCESS")
    Dim cTargetTable As Long: cTargetTable = ColumnIndex(loTpl, "TARGET_TABLE")
    Dim cTargetCol As Long: cTargetCol = ColumnIndex(loTpl, "TARGET_COLUMN")
    Dim cFormula As Long: cFormula = ColumnIndex(loTpl, "FORMULA")
    Dim cNotes As Long: cNotes = ColumnIndex(loTpl, "NOTES")
    Dim cActive As Long: cActive = ColumnIndex(loTpl, "ACTIVE")
    Dim cCreated As Long: cCreated = ColumnIndex(loTpl, "CREATED_AT")
    Dim cUpdated As Long: cUpdated = ColumnIndex(loTpl, "UPDATED_AT")

    If Not loTpl.DataBodyRange Is Nothing And cScope > 0 And cRecipe > 0 Then
        Dim r As Long
        For r = loTpl.DataBodyRange.Rows.Count To 1 Step -1
            If StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cScope).Value), TEMPLATE_SCOPE_RECIPE_PROCESS, vbTextCompare) = 0 Then
                If recipeId = "" Or StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cRecipe).Value), recipeId, vbTextCompare) = 0 Then
                    loTpl.ListRows(r).Delete
                End If
            End If
        Next r
    End If

    Dim nowVal As Date: nowVal = Now
    Dim added As Long

    Dim loProc As ListObject
    For Each loProc In processTables
        If loProc.DataBodyRange Is Nothing Then GoTo NextProc
        Dim procName As String: procName = ProcessNameFromTable(loProc)
        Dim lc As ListColumn
        For Each lc In loProc.ListColumns
            Dim formulaText As String
            formulaText = GetColumnFormulaText(lc)
            If formulaText = "" Then GoTo NextCol

            Dim lr As ListRow: Set lr = loTpl.ListRows.Add
            If cGuid > 0 Then lr.Range.Cells(1, cGuid).Value = modUR_Snapshot.GenerateGUID()
            If cScope > 0 Then lr.Range.Cells(1, cScope).Value = TEMPLATE_SCOPE_RECIPE_PROCESS
            If cRecipe > 0 Then lr.Range.Cells(1, cRecipe).Value = recipeId
            If cIngredient > 0 Then lr.Range.Cells(1, cIngredient).Value = ""
            If cProcess > 0 Then lr.Range.Cells(1, cProcess).Value = procName
            If cTargetTable > 0 Then lr.Range.Cells(1, cTargetTable).Value = loProc.Name
            If cTargetCol > 0 Then lr.Range.Cells(1, cTargetCol).Value = lc.Name
            If cFormula > 0 Then lr.Range.Cells(1, cFormula).Value = formulaText
            If cNotes > 0 Then lr.Range.Cells(1, cNotes).Value = "Recipe builder"
            If cActive > 0 Then lr.Range.Cells(1, cActive).Value = True
            If cCreated > 0 Then lr.Range.Cells(1, cCreated).Value = nowVal
            If cUpdated > 0 Then lr.Range.Cells(1, cUpdated).Value = nowVal
            added = added + 1
NextCol:
        Next lc
NextProc:
    Next loProc

    RegisterRecipeTemplates = added
End Function

Private Function ProcessNameFromTable(ByVal lo As ListObject) As String
    If lo Is Nothing Then Exit Function
    Dim cProc As Long: cProc = ColumnIndex(lo, "PROCESS")
    If cProc > 0 And Not lo.DataBodyRange Is Nothing Then
        ProcessNameFromTable = NzStr(lo.DataBodyRange.Cells(1, cProc).Value)
    End If
    If ProcessNameFromTable = "" Then ProcessNameFromTable = ExtractProcessKeyFromTableName(lo.Name)
End Function

Private Function GetColumnFormulaText(ByVal lc As ListColumn) As String
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function
    Dim cell As Range
    Set cell = lc.DataBodyRange.Cells(1, 1)
    On Error Resume Next
    If cell.HasFormula Then GetColumnFormulaText = CStr(cell.Formula)
    On Error GoTo 0
    If Left$(GetColumnFormulaText, 1) <> "=" Then GetColumnFormulaText = ""
End Function

Private Function SafeProcessKey(ByVal rawKey As String) As String
    Dim cleaned As String
    cleaned = Trim$(rawKey)
    If cleaned = "" Then cleaned = "process"

    Dim i As Long, ch As String, key As String
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            key = key & LCase$(ch)
        Else
            key = key & "_"
        End If
    Next i

    Do While InStr(key, "__") > 0
        key = Replace(key, "__", "_")
    Loop
    key = Trim$(key)
    If key = "" Then key = "process"
    If Not key Like "[A-Za-z_]*" Then key = "p_" & key
    SafeProcessKey = key
End Function

Private Function BuildRecipeProcessTableName(ByVal processKey As String) As String
    ' System 1: Recipe List Builder - process table naming.
    Dim key As String: key = Trim$(processKey)
    If key <> "" And IsNumeric(key) Then
        BuildRecipeProcessTableName = "proc_" & CLng(key) & "_" & RECIPE_PROC_TABLE_SUFFIX
    Else
        key = SafeProcessKey(processKey)
        BuildRecipeProcessTableName = "proc_" & key & "_" & RECIPE_PROC_TABLE_SUFFIX
    End If
End Function

Private Function NextRecipeProcessSequence(ByVal ws As Worksheet) As Long
    ' System 1: Recipe List Builder - next numeric process table sequence.
    Dim maxSeq As Long
    If ws Is Nothing Then
        NextRecipeProcessSequence = 1
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeProcessTable(lo) Then
            Dim seq As Long
            seq = RecipeProcessSequenceFromName(lo.Name)
            If seq > maxSeq Then maxSeq = seq
        End If
    Next lo
    NextRecipeProcessSequence = maxSeq + 1
End Function

Private Function RecipeProcessSequenceFromName(ByVal tableName As String) As Long
    ' System 1: Recipe List Builder - parse numeric process table sequence.
    Dim nm As String: nm = LCase$(tableName)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    If Right$(nm, Len(RECIPE_PROC_TABLE_SUFFIX) + 1) <> "_" & LCase$(RECIPE_PROC_TABLE_SUFFIX) Then Exit Function
    Dim core As String
    core = Mid$(nm, 6, Len(nm) - 5 - (Len(RECIPE_PROC_TABLE_SUFFIX) + 1))
    If core = "" Then Exit Function
    If Left$(core, 2) = "p_" Then core = Mid$(core, 3)
    RecipeProcessSequenceFromName = CLng(Val(core))
End Function

Private Function RecipeProcessHeaderList() As Variant
    ' System 1: Recipe List Builder - process table headers.
    RecipeProcessHeaderList = Array( _
        "PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", _
        "OOO", "INSTRUCTION", "RECIPE_LIST_ROW", "INGREDIENT_ID", "GUID")
End Function

Private Function HeaderRowArray(ByVal headers As Variant) As Variant
    Dim cols As Long: cols = UBound(headers) - LBound(headers) + 1
    Dim arr() As Variant
    ReDim arr(1 To 1, 1 To cols)
    Dim i As Long
    For i = 1 To cols
        arr(1, i) = headers(LBound(headers) + i - 1)
    Next i
    HeaderRowArray = arr
End Function

Private Function UniqueListObjectName(ByVal ws As Worksheet, ByVal baseName As String) As String
    Dim nameTry As String: nameTry = baseName
    Dim idx As Long: idx = 1
    Do While Not GetListObject(ws, nameTry) Is Nothing
        nameTry = baseName & "_" & CStr(idx)
        idx = idx + 1
    Loop
    UniqueListObjectName = nameTry
End Function

Private Function ExtractProcessKeyFromTableName(ByVal tableName As String) As String
    Dim nm As String: nm = LCase$(tableName)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    Dim parts As Variant: parts = Split(nm, "_")
    If UBound(parts) < 2 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(parts) - 1
        If ExtractProcessKeyFromTableName <> "" Then ExtractProcessKeyFromTableName = ExtractProcessKeyFromTableName & "_"
        ExtractProcessKeyFromTableName = ExtractProcessKeyFromTableName & parts(i)
    Next i
End Function

Private Function RangeHasListObjectCollision(ByVal ws As Worksheet, ByVal targetRange As Range, ParamArray allowedTables() As Variant) As Boolean
    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo Is Nothing Then GoTo NextLo
        If IsListObjectAllowed(lo, allowedTables) Then GoTo NextLo
        If Not Intersect(lo.Range, targetRange) Is Nothing Then
            RangeHasListObjectCollision = True
            Exit Function
        End If
NextLo:
    Next lo
End Function

Private Function IsListObjectAllowed(ByVal lo As ListObject, ByVal allowedTables As Variant) As Boolean
    Dim v As Variant
    For Each v In allowedTables
        If TypeName(v) = "ListObject" Then
            If lo Is v Then
                IsListObjectAllowed = True
                Exit Function
            End If
        End If
    Next v
    If IsRecipeProcessTable(lo) Then IsListObjectAllowed = True
End Function

Private Function RangeHasListObjectCollisionStrict(ByVal ws As Worksheet, ByVal targetRange As Range, ParamArray allowedTables() As Variant) As Boolean
    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo Is Nothing Then GoTo NextLo
        If IsListObjectAllowedStrict(lo, False, allowedTables) Then GoTo NextLo
        If Not Intersect(lo.Range, targetRange) Is Nothing Then
            RangeHasListObjectCollisionStrict = True
            Exit Function
        End If
NextLo:
    Next lo
End Function

Private Function IsListObjectAllowedStrict(ByVal lo As ListObject, ByVal allowRecipeTables As Boolean, ByVal allowedTables As Variant) As Boolean
    Dim v As Variant
    For Each v In allowedTables
        If TypeName(v) = "ListObject" Then
            If lo Is v Then
                IsListObjectAllowedStrict = True
                Exit Function
            End If
        End If
    Next v
    If allowRecipeTables Then
        If IsRecipeProcessTable(lo) Then IsListObjectAllowedStrict = True
    End If
End Function

Private Sub EnsureTableHasRow(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    lo.ListRows.Add AlwaysInsert:=True
    On Error GoTo 0
End Sub

' System 2: Inventory Palette Builder - clear values but keep a single data row.
Private Sub ResetPaletteTable(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableHasRow lo
    If lo.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    lo.DataBodyRange.SpecialCells(xlCellTypeConstants).ClearContents
    On Error GoTo 0
End Sub

Private Function GetHeaderDataCell(lo As ListObject, colName As String) As Range
    If lo Is Nothing Then Exit Function
    Dim idx As Long: idx = ColumnIndex(lo, colName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then
        Set GetHeaderDataCell = lo.HeaderRowRange.Offset(1, 0).Cells(1, idx)
    Else
        Set GetHeaderDataCell = lo.DataBodyRange.Cells(1, idx)
    End If
End Function

Private Sub ClearListObjectData(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    On Error GoTo 0
End Sub

Private Sub RemoveRecipeBuilderLinesTable(ByVal lo As ListObject)
    ' System 1: Recipe List Builder - remove RecipeBuilder lines table after load.
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim addr As String: addr = lo.Range.Address
    On Error Resume Next
    lo.Delete
    ws.Range(addr).Clear
    On Error GoTo 0
End Sub

Private Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function
