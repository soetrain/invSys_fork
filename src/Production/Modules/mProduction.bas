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
Private Const BTN_SAVE_FORMULAS As String = "BTN_SAVE_FORMULAS"
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

Private Const CHK_PROC_PREFIX As String = "CHK_PROC_"
Private Const CHK_BATCH_PREFIX As String = "CHK_BATCH_"
Private Const CHK_RECALL_PREFIX As String = "CHK_RECALL_"

Private Const TEMPLATE_SCOPE_RECIPE_PROCESS As String = "RECIPE_PROCESS"
Private Const TEMPLATE_SCOPE_PALETTE_BUILDER As String = "PALETTE_BUILDER"
Private Const TEMPLATE_SCOPE_PROD_RUN As String = "PROD_RUN"
Private Const TEMPLATE_TABLEKEY_PALETTE As String = "proc_*_palette"
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
Public Sub HandleProductionSelectionChange(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(target) Then Exit Sub
    EnsurePickerRouter
    mPickerRouter.HandleSelectionChange target
End Sub

Public Sub HandleProductionBeforeDoubleClick(ByVal target As Range, ByRef Cancel As Boolean)
    If target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(target) Then Exit Sub
    EnsurePickerRouter
    If mPickerRouter.HandleBeforeDoubleClick(target, Cancel) Then Exit Sub
End Sub

Private Sub EnsurePickerRouter()
    If mPickerRouter Is Nothing Then Set mPickerRouter = New cPickerRouter
End Sub

Public Sub HandleProductionChange(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(target) Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    If IsBandManagedTable(lo) Then
        EnsureRowCountCache
        Dim key As String: key = lo.Name
        Dim newCount As Long: newCount = ListObjectRowCount(lo)
        If Not mRowCountCache.Exists(key) Then
            mRowCountCache(key) = newCount
            Exit Sub
        End If
        Dim oldCount As Long: oldCount = CLng(mRowCountCache(key))
        If newCount > oldCount Then
            If LCase$(lo.Name) <> "prod_invsys_check" Then
                Dim bandMgr As New cTableBandManager
                bandMgr.Init lo.Parent
                bandMgr.ExpandBandForTable lo, (newCount - oldCount)
            End If
        End If
        mRowCountCache(key) = newCount
    End If

    If LCase$(lo.Name) = "productionoutput" Then
        RenderOutputRowCheckboxes lo.Parent
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

Private Function IsOnProductionSheet(ByVal target As Range) As Boolean
    On Error Resume Next
    IsOnProductionSheet = (target.Worksheet.Name = SHEET_PRODUCTION)
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

Private Function IsBandManagedTable(lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If nm = LCase$(TABLE_INV_PALETTE_GENERATED) Then
        IsBandManagedTable = True
    ElseIf nm Like "proc_*_palette" Then
        IsBandManagedTable = True
    ElseIf nm = "productionoutput" Then
        IsBandManagedTable = True
    ElseIf nm = "prod_invsys_check" Then
        IsBandManagedTable = True
    End If
End Function

Private Function ListObjectRowCount(lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ListObjectRowCount = lo.DataBodyRange.rows.count
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
    Dim seenRows As Object: Set seenRows = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = lo.DataBodyRange.value
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

    If dict.count = 0 Then Exit Function
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 3)
    Dim i As Long: i = 1
    Dim key As Variant
    For Each key In dict.keys
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
        If cRec > 0 Then loChooser.DataBodyRange.Cells(1, cRec).value = recipeName
        If cRecId > 0 Then loChooser.DataBodyRange.Cells(1, cRecId).value = recipeId
        If cDesc > 0 Then loChooser.DataBodyRange.Cells(1, cDesc).value = recipeDesc
        If cDept > 0 Then loChooser.DataBodyRange.Cells(1, cDept).value = recipeDept
        If cProc > 0 Then loChooser.DataBodyRange.Cells(1, cProc).value = ""
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
    RenderProcessSelectorCheckboxes wsProd, procTables
    BuildPaletteTablesForRecipeChooser recipeId, wsProd, wsRec, procTables, paletteStyle
    RenderPaletteKeepCheckboxes wsProd
    ApplyProductionOutputTemplates recipeId, wsProd
    RenderOutputRowCheckboxes wsProd

    Exit Sub
ErrHandler:
    MsgBox "Load Recipe Chooser failed: " & Err.description, vbCritical
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
    TableColumnCount = lo.HeaderRowRange.Columns.count
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
    For i = hdr.Columns.count To 1 Step -1
        Dim val As String
        val = Trim$(CStr(hdr.Cells(1, i).value))
        If val <> "" Then
            lastIdx = i
            Exit For
        End If
    Next i
    If lastIdx = 0 Then lastIdx = hdr.Columns.count
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
                    rTop = lo.Range.row
                    rBottom = lo.Range.row + lo.Range.rows.count - 1
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
    For Each shp In ws.shapes
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

Private Function ColumnIndexLoose(lo As ListObject, ParamArray names() As Variant) As Long
    If lo Is Nothing Then Exit Function
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        Dim hdr As String
        hdr = NormalizeHeaderKey(NzStr(lc.Name))
        Dim i As Long
        For i = LBound(names) To UBound(names)
            If hdr = NormalizeHeaderKey(CStr(names(i))) Then
                ColumnIndexLoose = lc.Index
                Exit Function
            End If
        Next i
    Next lc
End Function

Private Function NormalizeHeaderKey(ByVal v As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(v)
        ch = Mid$(v, i, 1)
        If ch Like "[A-Za-z0-9]" Then out = out & UCase$(ch)
    Next i
    NormalizeHeaderKey = out
End Function

Private Function NormalizeRowKey(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    Dim s As String
    If IsNumeric(v) Then
        NormalizeRowKey = CStr(CLng(v))
        Exit Function
    End If
    s = Trim$(CStr(v))
    If s = "" Then Exit Function
    If IsNumeric(s) Then
        NormalizeRowKey = CStr(CLng(val(s)))
    Else
        NormalizeRowKey = s
    End If
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
    Dim nextTop As Double: nextTop = ws.rows(2).Top

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
    EnsureButtonCustom ws, BTN_SAVE_FORMULAS, "Save Formulas", "mProduction.BtnSaveFormulas", leftA, nextTop, colAWidth
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
    If Not mHiddenSystems Is Nothing And mHiddenSystems.count > 0 Then
        idx = CLng(mHiddenSystems(mHiddenSystems.count))
        mHiddenSystems.Remove mHiddenSystems.count
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
    Set shp = ws.shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.shapes.AddFormControl(xlButtonControl, leftPos, topPos, widthPts, BTN_HEIGHT)
        shp.Name = shapeName
        shp.TextFrame.Characters.text = caption
        shp.OnAction = onActionMacro
    Else
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = BTN_HEIGHT
        shp.TextFrame.Characters.text = caption
        shp.OnAction = onActionMacro
    End If
End Sub

Private Sub DeleteShapeIfExists(ws As Worksheet, shapeName As String)
    On Error Resume Next
    ws.shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub HideGroupShapes(ws As Worksheet, startCol As Long, endCol As Long, topRow As Long, bottomRow As Long, hideIt As Boolean)
    If ws Is Nothing Then Exit Sub
    If startCol = 0 Or endCol = 0 Then Exit Sub
    Dim endColAdj As Long
    endColAdj = endCol + 6 ' allow checkboxes just right of the table
    Dim shp As Shape
    For Each shp In ws.shapes
        Dim c As Long
        Dim r As Long
        On Error Resume Next
        c = shp.TopLeftCell.Column
        r = shp.TopLeftCell.row
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

Public Sub BtnSaveFormulas()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim recipeId As String
    recipeId = ResolveActiveRecipeId(wsProd, True)
    If recipeId = "" Then
        MsgBox "Select or load a RECIPE before saving formulas.", vbExclamation
        Exit Sub
    End If

    Dim saved As Long
    saved = SaveFormulaTemplatesForRecipe(recipeId, wsProd)
    MsgBox "Saved formulas: " & saved & ".", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Save Formulas failed: " & Err.description, vbCritical
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
        If Not idCell Is Nothing Then recipeId = NzStr(idCell.value)
    End If
    Dim builtCount As Long
    If procTables.count = 0 Then
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

    If targets.count = 0 Then
        MsgBox "No Recipe Process tables selected.", vbInformation
        Exit Sub
    End If

    Dim key As Variant
    For Each key In targets.keys
        On Error Resume Next
        wsProd.ListObjects(CStr(key)).Delete
        wsProd.Range(CStr(targets(key))).Clear
        On Error GoTo 0
    Next key

    MsgBox "Removed " & targets.count & " Recipe Process table(s).", vbInformation
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
    DeleteCheckboxesByPrefix wsProd, CHK_PROC_PREFIX
    DeleteCheckboxesByPrefix wsProd, CHK_BATCH_PREFIX
    DeleteCheckboxesByPrefix wsProd, CHK_RECALL_PREFIX

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If Not loOut Is Nothing Then
        ClearListObjectContents loOut
    End If

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        ClearListObjectContents loCheck
    End If

    MsgBox "Recipe Chooser cleared.", vbInformation
End Sub

Public Sub BtnToUsed()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim usedDict As Object
    Set usedDict = BuildUsedDeltasFromPalette(wsProd)
    If usedDict Is Nothing Then
        MsgBox "No USED quantities found in palette tables.", vbInformation
        Exit Sub
    ElseIf usedDict.count = 0 Then
        MsgBox "No USED quantities found in palette tables.", vbInformation
        Exit Sub
    End If

    Dim errNotes As String
    Dim priorUsed As Object
    Set priorUsed = BuildUsedSnapshotFromCheck(FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV")))

    Dim stagedTotal As Double
    stagedTotal = StageUsedToInvSys(invLo, usedDict, priorUsed, errNotes)
    If stagedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unknown staging failure."
        MsgBox "To USED cancelled: " & errNotes, vbCritical
        Exit Sub
    End If

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        WriteProdInvSysCheck loCheck, invLo, usedDict
    End If

    Dim msg As String
    msg = "Applied USED deltas: " & Format$(stagedTotal, "0.###") & " units."
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
    Exit Sub
ErrHandler:
    MsgBox "BTN_TO_USED failed: " & Err.description, vbCritical
End Sub

Public Sub BtnToMade()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        MsgBox "ProductionOutput table not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim errNotes As String
    Dim outputEntries As Collection
    Set outputEntries = BuildOutputEntriesFromProcessTables(wsProd)
    If outputEntries Is Nothing Then
        MsgBox "No OUTPUT items found in process tables.", vbInformation
        Exit Sub
    ElseIf outputEntries.count = 0 Then
        MsgBox "No OUTPUT items found in process tables.", vbInformation
        Exit Sub
    End If

    UpdateProductionOutputTable loOut, outputEntries, invLo, errNotes
    EnsureOutputBatchNumbers loOut
    RenderOutputRowCheckboxes wsProd
    ApplyRecallCodesForOutput wsProd, loOut, invLo, errNotes

    Dim usedNotes As String
    Dim usedDeltas As Collection
    Set usedDeltas = BuildUsedDeltaPacketFromInvSys(invLo, usedNotes)

    Dim madeNotes As String
    Dim madeDeltas As Collection
    Set madeDeltas = BuildMadeDeltasFromProductionOutput(loOut, invLo, madeNotes)
    If madeDeltas Is Nothing Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        MsgBox "Send to MADE cancelled: " & madeNotes, vbExclamation
        Exit Sub
    ElseIf madeDeltas.count = 0 Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        MsgBox "Send to MADE cancelled: " & madeNotes, vbExclamation
        Exit Sub
    End If

    Dim usedTotal As Double
    Dim madeTotal As Double

    If Not usedDeltas Is Nothing Then
        usedTotal = modInvMan.ApplyUsedDeltas(usedDeltas, errNotes, "BTN_TO_MADE - Components Used")
        If usedTotal < 0 Then
            If errNotes = "" Then errNotes = "Unable to deduct USED inventory."
            MsgBox "Send to MADE cancelled: " & errNotes, vbExclamation
            Exit Sub
        End If
    ElseIf usedNotes <> "" Then
        AppendNote errNotes, usedNotes
    End If

    madeTotal = modInvMan.ApplyMadeDeltas(madeDeltas, errNotes, "BTN_TO_MADE - Finished Goods Staged")
    If madeTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to stage MADE inventory."
        MsgBox "Send to MADE cancelled: " & errNotes, vbExclamation
        Exit Sub
    End If

    Dim logNotes As String
    LogProductionOutputToProductionLog wsProd, loOut, invLo, logNotes
    If logNotes <> "" Then AppendNote errNotes, logNotes

    Dim rowKeys As Object
    Set rowKeys = BuildRowKeySetFromDeltas(usedDeltas, madeDeltas)
    Dim usedSnapshot As Object
    Set usedSnapshot = BuildUsedSnapshotForRows(invLo, rowKeys)

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        If Not usedSnapshot Is Nothing Then
            WriteProdInvSysCheck loCheck, invLo, usedSnapshot
        End If
    End If

    Dim msg As String
    msg = "Recorded component usage: " & Format$(usedTotal, "0.###") & " units."
    msg = msg & vbCrLf & "Recorded finished goods (MADE): " & Format$(madeTotal, "0.###")
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
    Exit Sub
ErrHandler:
    MsgBox "BTN_TO_MADE failed: " & Err.description, vbCritical
End Sub

Public Sub BtnToTotalInv()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        MsgBox "ProductionOutput table not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim errNotes As String
    Dim madeNotes As String
    Dim madeDeltas As Collection
    Set madeDeltas = BuildMadeDeltasFromProductionOutput(loOut, invLo, madeNotes)
    If madeDeltas Is Nothing Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        MsgBox "Send to TOTAL INV cancelled: " & madeNotes, vbExclamation
        Exit Sub
    ElseIf madeDeltas.count = 0 Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        MsgBox "Send to TOTAL INV cancelled: " & madeNotes, vbExclamation
        Exit Sub
    End If

    Dim totalMoved As Double
    totalMoved = modInvMan.ApplyMadeToInventoryDeltas(madeDeltas, errNotes, "BTN_TO_TOTALINV - Move Made to Total Inv")
    If totalMoved < 0 Then
        If errNotes = "" Then errNotes = "Unable to move MADE to TOTAL INV."
        MsgBox "Send to TOTAL INV cancelled: " & errNotes, vbExclamation
        Exit Sub
    End If

    Dim rowKeys As Object
    Set rowKeys = BuildRowKeySetFromDeltas(Nothing, madeDeltas)
    Dim usedSnapshot As Object
    Set usedSnapshot = BuildUsedSnapshotForRows(invLo, rowKeys)

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        If Not usedSnapshot Is Nothing Then
            WriteProdInvSysCheck loCheck, invLo, usedSnapshot
        End If
    End If

    Dim msg As String
    msg = "Moved MADE to TOTAL INV: " & Format$(totalMoved, "0.###") & " units."
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
    Exit Sub
ErrHandler:
    MsgBox "BTN_TO_TOTALINV failed: " & Err.description, vbCritical
End Sub

Public Sub BtnNextBatch()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    EnsurePaletteTableMetaForExistingTables ws

    Dim invLo As ListObject
    Set invLo = GetInvSysTable()

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If Not loOut Is Nothing Then
        EnsureOutputBatchNumbers loOut
        ClearProductionOutputForNextBatch ws, loOut
    End If

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row >= PALETTE_LINES_STAGING_ROW Then GoTo NextLo
            Dim procName As String
            Dim recipeId As String
            Dim ingId As String
            Dim amtVal As Variant
            Dim ioVal As String
            If GetPaletteTableContext(lo, recipeId, ingId, amtVal, procName, ioVal) = False Then
                procName = ProcessNameFromTable(lo)
            End If
            If Trim$(procName) = "" Then procName = lo.Name

            If Not IsPaletteKeepSelected(ws, procName) Then
                ClearPaletteTableSelection lo
            End If
        End If
NextLo:
    Next lo

    MsgBox "Next Batch ready. Inventory selections cleared for unchecked processes.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "BTN_NEXT_BATCH failed: " & Err.description, vbCritical
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
    If wsPal Is Nothing Then Set wsPal = SheetExists("IngredientsPalette")
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
            For r = loPal.DataBodyRange.rows.count To 1 Step -1
                If NzStr(loPal.DataBodyRange.Cells(r, cPalRec).value) = recipeId _
                   And NzStr(loPal.DataBodyRange.Cells(r, cPalIng).value) = ingredientId Then
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
    Dim arr As Variant: arr = loItems.DataBodyRange.value
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        Dim itemVal As String
        If cItem > 0 Then itemVal = NzStr(arr(i, cItem))
        If Trim$(itemVal) = "" Then GoTo NextItem

        Dim lr As ListRow: Set lr = loPal.ListRows.Add
        If cOutRec > 0 Then lr.Range.Cells(1, cOutRec).value = recipeId
        If cOutIng > 0 Then lr.Range.Cells(1, cOutIng).value = ingredientId
        If cOutIO > 0 Then lr.Range.Cells(1, cOutIO).value = ioVal
        If cOutItem > 0 Then lr.Range.Cells(1, cOutItem).value = itemVal
        If cOutPct > 0 Then lr.Range.Cells(1, cOutPct).value = pctVal
        If cOutUom > 0 Then
            Dim itemUom As String
            If cUom > 0 Then itemUom = NzStr(arr(i, cUom))
            If itemUom <> "" Then
                lr.Range.Cells(1, cOutUom).value = itemUom
            Else
                lr.Range.Cells(1, cOutUom).value = uomVal
            End If
        End If
        If cOutAmt > 0 Then lr.Range.Cells(1, cOutAmt).value = amtVal
        If cOutRow > 0 And cRow > 0 Then lr.Range.Cells(1, cOutRow).value = arr(i, cRow)
        If cOutGuid > 0 Then lr.Range.Cells(1, cOutGuid).value = modUR_Snapshot.GenerateGUID()
        added = added + 1
NextItem:
    Next i

    MsgBox "Saved IngredientPalette rows: " & added & ".", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Save IngredientPalette failed: " & Err.description, vbCritical
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
            If Not recCell Is Nothing Then recCell.value = recipeId
        End If
    End If

    If Not loItems Is Nothing Then
        ClearListObjectFormulas loItems
    End If
    If Not loIng Is Nothing Then
        ClearListObjectFormulas loIng
    End If

    Dim tpl As New cTemplateApplier
    If Not loIng Is Nothing Then
        tpl.ApplyTemplates loIng, TEMPLATE_SCOPE_PALETTE_BUILDER, "", "IP_ChooseIngredient", recipeId
    End If
    If Not loItems Is Nothing Then
        tpl.ApplyTemplates loItems, TEMPLATE_SCOPE_PALETTE_BUILDER, "", "IP_ChooseItem", recipeId
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
            If Not recCell Is Nothing Then recCell.value = recipeId
        End If
        If cIng > 0 Then
            Dim ingCell As Range
            Set ingCell = GetHeaderDataCell(loItems, "INGREDIENT_ID")
            If Not ingCell Is Nothing Then ingCell.value = ingredientId
        End If
        PopulateChooseItemFromIngredientPalette recipeId, ingredientId, loItems
    End If
End Sub

Private Sub PopulateChooseItemFromIngredientPalette(ByVal recipeId As String, ByVal ingredientId As String, ByVal loItems As ListObject)
    If Trim$(recipeId) = "" Or Trim$(ingredientId) = "" Then Exit Sub
    If loItems Is Nothing Then Exit Sub

    Dim wsPal As Worksheet: Set wsPal = SheetExists("IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = SheetExists("IngredientsPalette")
    If wsPal Is Nothing Then Exit Sub

    Dim loPal As ListObject
    Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "ROW"))
    If loPal Is Nothing Then
        Set loPal = FindListObjectByNameOrHeaders(wsPal, "Table40", Array("RECIPE_ID", "INGREDIENT_ID", "ROW"))
    End If
    If loPal Is Nothing Then Exit Sub
    If loPal.DataBodyRange Is Nothing Then Exit Sub

    Dim cRec As Long: cRec = ColumnIndex(loPal, "RECIPE_ID")
    Dim cIng As Long: cIng = ColumnIndex(loPal, "INGREDIENT_ID")
    Dim cRow As Long: cRow = ColumnIndex(loPal, "ROW")
    Dim cItem As Long: cItem = ColumnIndex(loPal, "ITEM")
    Dim cUom As Long: cUom = ColumnIndex(loPal, "UOM")
    If cRec = 0 Or cIng = 0 Then Exit Sub

    Dim oItem As Long: oItem = ColumnIndex(loItems, "ITEMS")
    If oItem = 0 Then oItem = ColumnIndex(loItems, "ITEM")
    Dim oUom As Long: oUom = ColumnIndex(loItems, "UOM")
    Dim oDesc As Long: oDesc = ColumnIndex(loItems, "DESCRIPTION")
    Dim oRow As Long: oRow = ColumnIndex(loItems, "ROW")
    Dim oRec As Long: oRec = ColumnIndex(loItems, "RECIPE_ID")
    Dim oIng As Long: oIng = ColumnIndex(loItems, "INGREDIENT_ID")

    Dim wsInv As Worksheet: Set wsInv = SheetExists("InventoryManagement")
    Dim loInv As ListObject
    If Not wsInv Is Nothing Then Set loInv = GetListObject(wsInv, "invSys")

    Dim arr As Variant: arr = loPal.DataBodyRange.value
    Dim r As Long, writeRow As Long
    Dim normRec As String: normRec = NormalizeIdFirst(recipeId)
    Dim normIng As String: normIng = NormalizeIdLast(ingredientId)
    For r = 1 To UBound(arr, 1)
        If NormalizeIdFirst(NzStr(arr(r, cRec))) = normRec And NormalizeIdLast(NzStr(arr(r, cIng))) = normIng Then
            Dim rowVal As Long
            rowVal = CLng(NzLng(arr(r, cRow)))
            Dim itemName As String: itemName = IIf(cItem > 0, NzStr(arr(r, cItem)), "")
            Dim uomVal As String: uomVal = IIf(cUom > 0, NzStr(arr(r, cUom)), "")
            Dim descVal As String: descVal = ""

            If rowVal > 0 And Not loInv Is Nothing Then
                ResolveInvSysDetailsByRow loInv, rowVal, itemName, uomVal, descVal
            End If

            writeRow = writeRow + 1
            EnsureListObjectRowCount loItems, writeRow
            If oItem > 0 Then loItems.DataBodyRange.Cells(writeRow, oItem).value = itemName
            If oUom > 0 Then loItems.DataBodyRange.Cells(writeRow, oUom).value = uomVal
            If oDesc > 0 Then loItems.DataBodyRange.Cells(writeRow, oDesc).value = descVal
            If oRow > 0 Then loItems.DataBodyRange.Cells(writeRow, oRow).value = rowVal
            If oRec > 0 Then loItems.DataBodyRange.Cells(writeRow, oRec).value = recipeId
            If oIng > 0 Then loItems.DataBodyRange.Cells(writeRow, oIng).value = ingredientId
        End If
    Next r
End Sub

Private Sub EnsureListObjectRowCount(ByVal lo As ListObject, ByVal needed As Long)
    If lo Is Nothing Then Exit Sub
    If needed < 1 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add AlwaysInsert:=True
    End If
    Do While lo.ListRows.count < needed
        lo.ListRows.Add AlwaysInsert:=True
    Loop
End Sub

Private Function EnsureListObjectRowCountSafe(ByVal lo As ListObject, ByVal needed As Long) As Boolean
    EnsureListObjectRowCountSafe = True
    If lo Is Nothing Then Exit Function
    If needed < 1 Then Exit Function

    On Error Resume Next
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add AlwaysInsert:=True
        If Err.Number <> 0 Then
            EnsureListObjectRowCountSafe = False
            Err.Clear
            Exit Function
        End If
    End If
    Do While lo.ListRows.count < needed
        lo.ListRows.Add AlwaysInsert:=True
        If Err.Number <> 0 Then
            EnsureListObjectRowCountSafe = False
            Err.Clear
            Exit Function
        End If
    Loop
    On Error GoTo 0
End Function

Private Function ExpandProductionInputOutputBand(ByVal ws As Worksheet, ByVal loCheck As ListObject, ByVal rowsAdded As Long) As Boolean
    ExpandProductionInputOutputBand = False
    If ws Is Nothing Then Exit Function
    If loCheck Is Nothing Then Exit Function
    If rowsAdded <= 0 Then Exit Function

    Dim bandLeft As Long
    Dim bandRight As Long
    Dim lo As ListObject

    Dim sCol As Long
    Dim eCol As Long
    If TableEffectiveSpan(loCheck, sCol, eCol) Then
        bandLeft = sCol
        bandRight = eCol
    Else
        bandLeft = loCheck.Range.Column
        bandRight = loCheck.Range.Column + loCheck.Range.Columns.count - 1
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If Not loOut Is Nothing Then
        If TableEffectiveSpan(loOut, sCol, eCol) Then
            If sCol < bandLeft Then bandLeft = sCol
            If eCol > bandRight Then bandRight = eCol
        End If
    End If

    For Each lo In ws.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                If TableEffectiveSpan(lo, sCol, eCol) Then
                    If sCol < bandLeft Then bandLeft = sCol
                    If eCol > bandRight Then bandRight = eCol
                End If
            End If
        End If
    Next lo

    If bandLeft = 0 Or bandRight = 0 Then Exit Function

    Dim insertTop As Long
    insertTop = loCheck.Range.row + loCheck.Range.rows.count
    If insertTop <= 0 Then Exit Function
    If insertTop + rowsAdded - 1 > ws.rows.count Then Exit Function

    On Error Resume Next
    ws.rows(insertTop).Resize(rowsAdded).Insert Shift:=xlShiftDown
    If Err.Number = 0 Then ExpandProductionInputOutputBand = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function ExpandProductionOutputBand(ByVal ws As Worksheet, ByVal loOut As ListObject, ByVal rowsAdded As Long) As Boolean
    ExpandProductionOutputBand = False
    If ws Is Nothing Then Exit Function
    If loOut Is Nothing Then Exit Function
    If rowsAdded <= 0 Then Exit Function

    Dim bandLeft As Long
    Dim bandRight As Long
    Dim lo As ListObject

    Dim sCol As Long
    Dim eCol As Long
    If TableEffectiveSpan(loOut, sCol, eCol) Then
        bandLeft = sCol
        bandRight = eCol
    Else
        bandLeft = loOut.Range.Column
        bandRight = loOut.Range.Column + loOut.Range.Columns.count - 1
    End If

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(ws, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        If TableEffectiveSpan(loCheck, sCol, eCol) Then
            If sCol < bandLeft Then bandLeft = sCol
            If eCol > bandRight Then bandRight = eCol
        End If
    End If

    For Each lo In ws.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                If TableEffectiveSpan(lo, sCol, eCol) Then
                    If sCol < bandLeft Then bandLeft = sCol
                    If eCol > bandRight Then bandRight = eCol
                End If
            End If
        End If
    Next lo

    If bandLeft = 0 Or bandRight = 0 Then Exit Function

    Dim insertTop As Long
    insertTop = loOut.Range.row + loOut.Range.rows.count
    If insertTop <= 0 Then Exit Function
    If insertTop + rowsAdded - 1 > ws.rows.count Then Exit Function

    On Error Resume Next
    ws.rows(insertTop).Resize(rowsAdded).Insert Shift:=xlShiftDown
    If Err.Number = 0 Then ExpandProductionOutputBand = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function ExpandListObjectRows(ByVal lo As ListObject, ByVal addRows As Long) As Boolean
    ExpandListObjectRows = False
    If lo Is Nothing Then Exit Function
    If addRows <= 0 Then Exit Function

    Dim ws As Worksheet
    Set ws = lo.Parent

    Dim baseRange As Range
    Set baseRange = lo.Range
    If baseRange Is Nothing Then Exit Function

    Dim newRowCount As Long
    newRowCount = baseRange.rows.count + addRows
    If baseRange.row + newRowCount - 1 > ws.rows.count Then Exit Function

    Dim newRange As Range
    Set newRange = baseRange.Resize(newRowCount, baseRange.Columns.count)

    On Error Resume Next
    lo.Resize newRange
    If Err.Number = 0 Then ExpandListObjectRows = True
    Err.Clear
    On Error GoTo 0
End Function

Private Sub EnsureListObjectRowCountFullRow(ByVal lo As ListObject, ByVal needed As Long)
    ' Expand table by inserting full worksheet rows to avoid table-collision errors.
    If lo Is Nothing Then Exit Sub
    If needed < 1 Then Exit Sub

    Dim currentRows As Long
    If lo.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = lo.DataBodyRange.rows.count
    End If
    If currentRows >= needed Then Exit Sub

    Dim addRows As Long
    addRows = needed - currentRows

    Dim lastRow As Long
    lastRow = lo.Range.row + lo.Range.rows.count - 1

    Dim insertAt As Long
    insertAt = lastRow
    If lo.ShowTotals Then
        insertAt = lastRow - 1
    End If

    Dim ws As Worksheet
    Set ws = lo.Parent
    ws.rows(insertAt + 1).Resize(addRows).Insert Shift:=xlShiftDown

    Dim newRange As Range
    Set newRange = lo.Range.Resize(lo.Range.rows.count + addRows)
    lo.Resize newRange
End Sub

Private Sub ResolveInvSysDetailsByRow(ByVal loInv As ListObject, ByVal invRow As Long, _
    ByRef itemName As String, ByRef uomVal As String, ByRef descVal As String)

    If loInv Is Nothing Then Exit Sub
    If invRow <= 0 Then Exit Sub
    If loInv.DataBodyRange Is Nothing Then Exit Sub

    Dim cRow As Long: cRow = ColumnIndex(loInv, "ROW")
    Dim cItem As Long: cItem = ColumnIndex(loInv, "ITEM")
    Dim cUom As Long: cUom = ColumnIndex(loInv, "UOM")
    Dim cDesc As Long: cDesc = ColumnIndex(loInv, "DESCRIPTION")
    If cRow = 0 Then Exit Sub

    Dim cel As Range
    For Each cel In loInv.ListColumns(cRow).DataBodyRange.Cells
        If NzLng(cel.value) = invRow Then
            If itemName = "" And cItem > 0 Then itemName = NzStr(cel.Offset(0, cItem - cel.Column).value)
            If uomVal = "" And cUom > 0 Then uomVal = NzStr(cel.Offset(0, cUom - cel.Column).value)
            If descVal = "" And cDesc > 0 Then descVal = NzStr(cel.Offset(0, cDesc - cel.Column).value)
            Exit Sub
        End If
    Next cel
End Sub

Public Function GetPaletteRecipeId() As String
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Function
    Dim loRecipe As ListObject
    Set loRecipe = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseRecipe", Array("RECIPE_NAME", "RECIPE_ID"))
    If loRecipe Is Nothing Then Exit Function
    GetPaletteRecipeId = NormalizeIdFirst(FirstNonEmptyColumnValue(loRecipe, "RECIPE_ID"))
End Function

Private Function GetRecipeBuilderRecipeId(ByVal wsProd As Worksheet, Optional ByVal allowGenerate As Boolean = False) As String
    If wsProd Is Nothing Then Exit Function
    Dim loHeader As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    If loHeader Is Nothing Then Exit Function

    Dim idCell As Range: Set idCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
    Dim nameCell As Range: Set nameCell = GetHeaderDataCell(loHeader, "RECIPE_NAME")
    Dim recipeId As String
    If Not idCell Is Nothing Then recipeId = NzStr(idCell.value)

    If recipeId = "" And allowGenerate Then
        If Not nameCell Is Nothing Then
            If Trim$(NzStr(nameCell.value)) <> "" Then
                recipeId = modUR_Snapshot.GenerateGUID()
                If Not idCell Is Nothing Then idCell.value = recipeId
            End If
        End If
    End If

    GetRecipeBuilderRecipeId = NormalizeIdFirst(recipeId)
End Function

Private Function ResolveActiveRecipeId(ByVal wsProd As Worksheet, Optional ByVal allowGenerate As Boolean = False) As String
    Dim recipeId As String
    recipeId = GetRecipeBuilderRecipeId(wsProd, allowGenerate)
    If recipeId <> "" Then
        ResolveActiveRecipeId = recipeId
        Exit Function
    End If

    recipeId = GetPaletteRecipeId()
    If recipeId <> "" Then
        ResolveActiveRecipeId = recipeId
        Exit Function
    End If

    recipeId = GetRecipeChooserRecipeId(wsProd)
    ResolveActiveRecipeId = recipeId
End Function

Public Function GetPaletteIngredientId() As String
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Function
    Dim loIng As ListObject
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    If loIng Is Nothing Then Exit Function
    GetPaletteIngredientId = NormalizeIdLast(FirstNonEmptyColumnValue(loIng, "INGREDIENT_ID"))
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
    Dim arr As Variant: arr = lo.DataBodyRange.value
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

    If dict.count = 0 Then Exit Function
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 7)
    Dim i As Long: i = 1
    Dim k As Variant
    For Each k In dict.keys
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

    Dim arr As Variant: arr = lo.DataBodyRange.value
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

    Dim arr As Variant: arr = loRecipes.DataBodyRange.value
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

    If procOrder.count = 0 Then
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
        Dim dataCount As Long: dataCount = rowsColl.count
        If dataCount = 0 Then GoTo NextProc

        Dim tableRange As Range
        Set tableRange = wsProd.Range(wsProd.Cells(startRow, startCol), wsProd.Cells(startRow + dataCount, startCol + colCount - 1))
        If RangeHasListObjectCollisionStrict(wsProd, tableRange) Then
            Set tableRange = FindAvailableRecipeChooserRange(wsProd, startRow, startCol, dataCount + 1, colCount)
            If tableRange Is Nothing Then Exit For
        End If

        tableRange.Clear
        tableRange.rows(1).value = HeaderRowArray(headerNames)

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

        tableRange.Offset(1, 0).Resize(dataCount, colCount).value = dataArr

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

        startRow = tableRange.row + tableRange.rows.count + 3 ' keep 2 blank rows
NextProc:
    Next procKey

    If created.count > 0 Then
        Dim tpl As New cTemplateApplier
        Dim loProc As ListObject
        For Each loProc In created
            Dim procNameTpl As String: procNameTpl = ProcessNameFromTable(loProc)
            tpl.ApplyTemplates loProc, TEMPLATE_SCOPE_RECIPE_PROCESS, procNameTpl, "", recipeId
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

    Dim arr As Variant: arr = loRecipes.DataBodyRange.value
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
                    If Not IsProcessSelected(procName, wsProd) Then GoTo NextRecipeRow
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
NextRecipeRow:
    Next r

    If entries.count = 0 Then Exit Sub

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
    Dim invRowMap As Object
    Set invRowMap = BuildInvSysRowMap()

    Dim idx As Long
    Dim tpl As New cTemplateApplier
    Dim nextSeq As Long: nextSeq = 1
    Dim hdrProc As Long: hdrProc = HeaderIndex(headerNames, "PROCESS")
    Dim hdrIO As Long: hdrIO = HeaderIndex(headerNames, "INPUT/OUTPUT")
    Dim hdrQty As Long: hdrQty = HeaderIndex(headerNames, "QUANTITY")
    Dim hdrRow As Long: hdrRow = HeaderIndex(headerNames, "ROW")

    For idx = 1 To entries.count
        Dim infoArr As Variant
        infoArr = entries(idx)

        Dim rowList As Collection
        Set rowList = GetIngredientPaletteRows(infoArr(0), infoArr(1))

        Dim dataCount As Long
        If rowList Is Nothing Then
            dataCount = 1
        ElseIf rowList.count = 0 Then
            dataCount = 1
        Else
            dataCount = rowList.count
        End If

        Dim tableRange As Range
        Set tableRange = wsProd.Range(wsProd.Cells(startRow, startCol), wsProd.Cells(startRow + dataCount, startCol + colCount - 1))
        If RangeHasListObjectCollisionStrict(wsProd, tableRange) Then
            Set tableRange = FindAvailablePaletteRange(wsProd, startRow, startCol, dataCount + 1, colCount)
            If tableRange Is Nothing Then Exit For
        End If

        tableRange.Clear
        tableRange.rows(1).value = HeaderRowArray(headerNames)

        Dim dataArr() As Variant
        ReDim dataArr(1 To dataCount, 1 To colCount)
        Dim r2 As Long
        For r2 = 1 To dataCount
            If hdrProc > 0 Then dataArr(r2, hdrProc) = NzStr(infoArr(3))
            If hdrIO > 0 Then dataArr(r2, hdrIO) = NzStr(infoArr(4))
            If hdrQty > 0 Then dataArr(r2, hdrQty) = infoArr(2)
            If hdrRow > 0 Then
                If Not rowList Is Nothing And rowList.count > 0 Then
                    dataArr(r2, hdrRow) = rowList(r2)
                End If
            End If
        Next r2
        tableRange.Offset(1, 0).Resize(dataCount, colCount).value = dataArr

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

        ApplyProcessHeaderColor newLo, NzStr(infoArr(3))

        FillPaletteTableFromInvSys newLo, invRowMap

        tpl.ApplyTemplates newLo, TEMPLATE_SCOPE_PROD_RUN, NzStr(infoArr(3)), TEMPLATE_TABLEKEY_PALETTE, recipeId

        startRow = tableRange.row + tableRange.rows.count + 3
    Next idx
End Sub

Private Sub ApplyProductionOutputTemplates(ByVal recipeId As String, ByVal wsProd As Worksheet)
    If wsProd Is Nothing Then Exit Sub
    If Trim$(recipeId) = "" Then Exit Sub
    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then Exit Sub
    ClearListObjectFormulas loOut
    Dim tpl As New cTemplateApplier
    tpl.ApplyTemplates loOut, TEMPLATE_SCOPE_PROD_RUN, "", "ProductionOutput", recipeId
End Sub

Private Sub DeleteRecipeChooserProcessTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.count To 1 Step -1
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
    RecipeChooserSequenceFromName = CLng(val(core))
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
    startRow = loChooser.Range.row + loChooser.Range.rows.count + 2 ' one blank row
    If startRow > 0 And startCol > 0 Then GetRecipeChooserAnchor = True
End Function

Private Function FindAvailableRecipeChooserRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
    ByVal totalRows As Long, ByVal totalCols As Long) As Range

    If ws Is Nothing Then Exit Function
    If totalRows < 1 Or totalCols < 1 Then Exit Function
    If startRow < 1 Then startRow = 1
    If startCol < 1 Then startCol = 1

    Dim maxRow As Long: maxRow = ws.rows.count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range
    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate) Then
            Set FindAvailableRecipeChooserRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 3
    Loop
End Function

Private Sub DeleteInventoryPaletteTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.count To 1 Step -1
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
        If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
            startRow = lo.Range.row
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
        bottom = loProd.Range.row + loProd.Range.rows.count - 1
        If Not loCheck Is Nothing Then
            Dim chkBottom As Long
            chkBottom = loCheck.Range.row + loCheck.Range.rows.count - 1
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

    Dim maxRow As Long: maxRow = ws.rows.count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range
    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate) Then
            Set FindAvailablePaletteRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 3
    Loop
End Function

Private Function BuildInvSysRowMap() As Object
    Dim loInv As ListObject
    Set loInv = GetInvSysTable()
    If loInv Is Nothing Or loInv.DataBodyRange Is Nothing Then Exit Function

    Dim cRow As Long: cRow = ColumnIndex(loInv, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLoose(loInv, "ROW", "ROWID", "ROW#")
    If cRow = 0 Then Exit Function
    Dim cCode As Long: cCode = ColumnIndex(loInv, "ITEM_CODE")
    If cCode = 0 Then cCode = ColumnIndexLoose(loInv, "ITEM_CODE", "ITEMCODE", "ITEM CODE")
    Dim cVend As Long: cVend = ColumnIndex(loInv, "VENDOR(s)")
    If cVend = 0 Then cVend = ColumnIndexLoose(loInv, "VENDORS", "VENDOR", "VENDOR(S)")
    Dim cVendCode As Long: cVendCode = ColumnIndex(loInv, "VENDOR_CODE")
    If cVendCode = 0 Then cVendCode = ColumnIndexLoose(loInv, "VENDOR_CODE", "VENDORCODE", "VENDOR CODE")
    Dim cDesc As Long: cDesc = ColumnIndex(loInv, "DESCRIPTION")
    If cDesc = 0 Then cDesc = ColumnIndexLoose(loInv, "DESCRIPTION", "DESC")
    Dim cItem As Long: cItem = ColumnIndex(loInv, "ITEM")
    If cItem = 0 Then cItem = ColumnIndexLoose(loInv, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME")
    Dim cUom As Long: cUom = ColumnIndex(loInv, "UOM")
    If cUom = 0 Then cUom = ColumnIndexLoose(loInv, "UOM", "UNIT", "UNITOFMEASURE", "UNITOFMEASUREMENT")
    Dim cLoc As Long: cLoc = ColumnIndex(loInv, "LOCATION")
    If cLoc = 0 Then cLoc = ColumnIndexLoose(loInv, "LOCATION", "LOC")

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loInv.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowKey As String
        rowKey = NormalizeRowKey(arr(r, cRow))
        If rowKey <> "" Then
            If Not dict.Exists(rowKey) Then
                Dim info(1 To 7) As Variant
                If cCode > 0 Then info(1) = NzStr(arr(r, cCode)) Else info(1) = ""
                If cVend > 0 Then info(2) = NzStr(arr(r, cVend)) Else info(2) = ""
                If cVendCode > 0 Then info(3) = NzStr(arr(r, cVendCode)) Else info(3) = ""
                If cDesc > 0 Then info(4) = NzStr(arr(r, cDesc)) Else info(4) = ""
                If cItem > 0 Then info(5) = NzStr(arr(r, cItem)) Else info(5) = ""
                If cUom > 0 Then info(6) = NzStr(arr(r, cUom)) Else info(6) = ""
                If cLoc > 0 Then info(7) = NzStr(arr(r, cLoc)) Else info(7) = ""
                dict.Add rowKey, info
            End If
        End If
    Next r

    Set BuildInvSysRowMap = dict
End Function


Private Function GetInvSysTable() As ListObject
    Dim wsInv As Worksheet: Set wsInv = SheetExists("InventoryManagement")
    If wsInv Is Nothing Then Set wsInv = SheetExists("Inventory Management")
    If wsInv Is Nothing Then Set wsInv = SheetExists("INVENTORY MANAGEMENT")
    If wsInv Is Nothing Then Exit Function

    Dim loInv As ListObject: Set loInv = GetListObject(wsInv, "invSys")
    If Not loInv Is Nothing Then
        Set GetInvSysTable = loInv
        Exit Function
    End If

    Dim lo As ListObject
    For Each lo In wsInv.ListObjects
        If ColumnIndexLoose(lo, "ROW", "ROWID", "ROW#") > 0 Then
            If ColumnIndexLoose(lo, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME") > 0 _
                Or ColumnIndexLoose(lo, "ITEM_CODE", "ITEMCODE", "ITEM CODE") > 0 Then
                Set GetInvSysTable = lo
                Exit Function
            End If
        End If
    Next lo
End Function

Private Function BuildUsedDeltasFromPalette(ByVal wsProd As Worksheet) As Object
    If wsProd Is Nothing Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim seenRows As Object: Set seenRows = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject
    For Each lo In wsProd.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row >= PALETTE_LINES_STAGING_ROW Then GoTo NextLo
            If lo.DataBodyRange Is Nothing Then GoTo NextLo

            Dim cRow As Long: cRow = ColumnIndex(lo, "ROW")
            If cRow = 0 Then GoTo NextLo
            Dim cQty As Long: cQty = ColumnIndex(lo, "QUANTITY")
            If cQty = 0 Then GoTo NextLo
            Dim cIO As Long: cIO = ColumnIndex(lo, "INPUT/OUTPUT")
            Dim cProc As Long: cProc = ColumnIndex(lo, "PROCESS")

            Dim arr As Variant: arr = lo.DataBodyRange.value
            Dim r As Long
            For r = 1 To UBound(arr, 1)
                Dim rowKey As String
                rowKey = NormalizeRowKey(arr(r, cRow))
                If rowKey = "" Then GoTo NextRow

                If cProc > 0 Then
                    Dim procName As String
                    procName = NzStr(arr(r, cProc))
                    If procName <> "" Then
                        If Not IsProcessSelected(procName, wsProd) Then GoTo NextRow
                    End If
                End If

                If cIO > 0 Then
                    Dim ioVal As String
                    ioVal = LCase$(Trim$(NzStr(arr(r, cIO))))
                    If ioVal <> "" And ioVal <> "used" Then GoTo NextRow
                End If

                Dim qty As Double
                qty = NzDbl(arr(r, cQty))
                If qty = 0 Then GoTo NextRow

                If seenRows.Exists(rowKey) Then GoTo NextRow
                seenRows.Add rowKey, True
                dict.Add rowKey, qty
NextRow:
            Next r
        End If
NextLo:
    Next lo

    If dict.count = 0 Then Exit Function
    Set BuildUsedDeltasFromPalette = dict
End Function

Private Function BuildUsedSnapshotFromCheck(ByVal loCheck As ListObject) As Object
    If loCheck Is Nothing Then Exit Function
    If loCheck.DataBodyRange Is Nothing Then Exit Function

    Dim cUsed As Long: cUsed = ColumnIndex(loCheck, "USED")
    If cUsed = 0 Then cUsed = ColumnIndexLoose(loCheck, "USED")
    Dim cRow As Long: cRow = ColumnIndex(loCheck, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLoose(loCheck, "ROW", "ROWID", "ROW#")
    If cUsed = 0 Or cRow = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loCheck.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowKey As String
        rowKey = NormalizeRowKey(arr(r, cRow))
        If rowKey <> "" Then
            dict(rowKey) = NzDbl(arr(r, cUsed))
        End If
    Next r

    If dict.count = 0 Then Exit Function
    Set BuildUsedSnapshotFromCheck = dict
End Function

Private Function BuildInvSysRowIndex(ByVal invLo As ListObject) As Object
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLoose(invLo, "ROW", "ROWID", "ROW#")
    If cRow = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = invLo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowKey As String
        rowKey = NormalizeRowKey(arr(r, cRow))
        If rowKey <> "" Then
            If Not dict.Exists(rowKey) Then dict.Add rowKey, r
        End If
    Next r

    Set BuildInvSysRowIndex = dict
End Function

Private Function StageUsedToInvSys(ByVal invLo As ListObject, ByVal usedDict As Object, ByVal priorUsed As Object, ByRef errNotes As String) As Double
    StageUsedToInvSys = -1
    If invLo Is Nothing Then
        AppendNote errNotes, "invSys table not found."
        Exit Function
    End If
    If usedDict Is Nothing Then
        AppendNote errNotes, "No USED quantities to stage."
        Exit Function
    ElseIf usedDict.count = 0 Then
        AppendNote errNotes, "No USED quantities to stage."
        Exit Function
    End If
    If invLo.DataBodyRange Is Nothing Then
        AppendNote errNotes, "invSys table has no data rows."
        Exit Function
    End If

    Dim cUsed As Long: cUsed = ColumnIndex(invLo, "USED")
    If cUsed = 0 Then cUsed = ColumnIndexLoose(invLo, "USED")
    If cUsed = 0 Then
        AppendNote errNotes, "invSys USED column not found."
        Exit Function
    End If

    Dim rowIndex As Object
    Set rowIndex = BuildInvSysRowIndex(invLo)
    If rowIndex Is Nothing Then
        AppendNote errNotes, "invSys ROW index not available."
        Exit Function
    ElseIf rowIndex.count = 0 Then
        AppendNote errNotes, "invSys ROW index not available."
        Exit Function
    End If

    Dim key As Variant
    For Each key In usedDict.keys
        If Not rowIndex.Exists(CStr(key)) Then
            AppendNote errNotes, "invSys ROW " & CStr(key) & " not found; staging cancelled."
        End If
    Next key
    If errNotes <> "" Then Exit Function

    Dim total As Double
    For Each key In usedDict.keys
        Dim idx As Long
        idx = CLng(rowIndex(CStr(key)))
        Dim qty As Double
        qty = NzDbl(usedDict(key))
        Dim prevQty As Double
        If Not priorUsed Is Nothing Then
            If priorUsed.Exists(CStr(key)) Then prevQty = NzDbl(priorUsed(CStr(key)))
        End If
        Dim delta As Double
        delta = qty - prevQty
        If delta <> 0 Then
            invLo.DataBodyRange.Cells(idx, cUsed).value = NzDbl(invLo.DataBodyRange.Cells(idx, cUsed).value) + delta
            total = total + delta
        End If
    Next key

    StageUsedToInvSys = total
End Function

Private Sub WriteProdInvSysCheck(ByVal loCheck As ListObject, ByVal invLo As ListObject, ByVal usedDict As Object)
    If loCheck Is Nothing Then Exit Sub
    If usedDict Is Nothing Then
        ClearListObjectContents loCheck
        Exit Sub
    ElseIf usedDict.count = 0 Then
        ClearListObjectContents loCheck
        Exit Sub
    End If
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim rowIndex As Object
    Set rowIndex = BuildInvSysRowIndex(invLo)
    If rowIndex Is Nothing Then Exit Sub
    If rowIndex.count = 0 Then Exit Sub

    Dim cUsedChk As Long: cUsedChk = ColumnIndex(loCheck, "USED")
    If cUsedChk = 0 Then cUsedChk = ColumnIndexLoose(loCheck, "USED")
    Dim cMadeChk As Long: cMadeChk = ColumnIndex(loCheck, "MADE")
    If cMadeChk = 0 Then cMadeChk = ColumnIndexLoose(loCheck, "MADE")
    Dim cTotalChk As Long: cTotalChk = ColumnIndex(loCheck, "TOTAL INV")
    If cTotalChk = 0 Then cTotalChk = ColumnIndexLoose(loCheck, "TOTALINV", "TOTAL_INV", "TOTALINVENTORY")
    Dim cRowChk As Long: cRowChk = ColumnIndex(loCheck, "ROW")
    If cRowChk = 0 Then cRowChk = ColumnIndexLoose(loCheck, "ROW", "ROWID", "ROW#")

    Dim cUsedInv As Long: cUsedInv = ColumnIndex(invLo, "USED")
    If cUsedInv = 0 Then cUsedInv = ColumnIndexLoose(invLo, "USED")
    Dim cMadeInv As Long: cMadeInv = ColumnIndex(invLo, "MADE")
    If cMadeInv = 0 Then cMadeInv = ColumnIndexLoose(invLo, "MADE")
    Dim cTotalInv As Long: cTotalInv = ColumnIndex(invLo, "TOTAL INV")
    If cTotalInv = 0 Then cTotalInv = ColumnIndexLoose(invLo, "TOTALINV", "TOTAL_INV", "TOTALINVENTORY")

    Dim keys As Variant
    keys = SortedKeys(usedDict)
    If IsEmpty(keys) Then Exit Sub

    Dim rowsNeeded As Long
    If IsArray(keys) Then
        rowsNeeded = UBound(keys) - LBound(keys) + 1
    Else
        rowsNeeded = 1
    End If
    If rowsNeeded <= 0 Then
        ClearListObjectData loCheck
        Exit Sub
    End If

    Dim cols As Long
    cols = TableColumnCount(loCheck)
    If cols <= 0 Then Exit Sub

    Dim currentRows As Long
    If loCheck.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = loCheck.DataBodyRange.rows.count
    End If
    If currentRows < rowsNeeded Then
        Call EnsureListObjectRowCountFullRow(loCheck, rowsNeeded)
    End If
    If loCheck.DataBodyRange Is Nothing Then Exit Sub
    currentRows = loCheck.DataBodyRange.rows.count
    If rowsNeeded > currentRows Then rowsNeeded = currentRows
    If rowsNeeded <= 0 Then Exit Sub

    Dim i As Long
    For i = 1 To loCheck.DataBodyRange.rows.count
        If i > rowsNeeded Then
            loCheck.DataBodyRange.rows(i).ClearContents
        Else
            Dim rowKey As String
            If IsArray(keys) Then
                rowKey = CStr(keys(LBound(keys) + i - 1))
            Else
                rowKey = CStr(keys)
            End If

            If rowIndex.Exists(rowKey) Then
                Dim invIdx As Long
                invIdx = CLng(rowIndex(rowKey))
                If cUsedChk > 0 Then loCheck.DataBodyRange.Cells(i, cUsedChk).value = NzDbl(usedDict(rowKey))
                If cMadeChk > 0 And cMadeInv > 0 Then loCheck.DataBodyRange.Cells(i, cMadeChk).value = NzDbl(invLo.DataBodyRange.Cells(invIdx, cMadeInv).value)
                If cTotalChk > 0 And cTotalInv > 0 Then loCheck.DataBodyRange.Cells(i, cTotalChk).value = NzDbl(invLo.DataBodyRange.Cells(invIdx, cTotalInv).value)
                If cRowChk > 0 Then loCheck.DataBodyRange.Cells(i, cRowChk).value = rowKey
            End If
        End If
    Next i
End Sub

Private Function BuildOutputEntriesFromProcessTables(ByVal wsProd As Worksheet) As Collection
    If wsProd Is Nothing Then Exit Function

    Dim recipeId As String
    recipeId = GetRecipeChooserRecipeId(wsProd)

    Dim procTables As Collection
    Set procTables = GetRecipeChooserProcessTables(wsProd)
    If procTables Is Nothing Then Exit Function
    If procTables.count = 0 Then Exit Function

    Dim entryMap As Object: Set entryMap = CreateObject("Scripting.Dictionary")
    Dim order As New Collection

    Dim lo As ListObject
    For Each lo In procTables
        If lo Is Nothing Then GoTo NextLo
        If lo.DataBodyRange Is Nothing Then GoTo NextLo

        Dim cIO As Long: cIO = ColumnIndex(lo, "INPUT/OUTPUT")
        Dim cIng As Long: cIng = ColumnIndex(lo, "INGREDIENT")
        If cIO = 0 Or cIng = 0 Then GoTo NextLo

        Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
        Dim cAmt As Long: cAmt = ColumnIndex(lo, "AMOUNT NEEDED")
        If cAmt = 0 Then cAmt = ColumnIndex(lo, "AMOUNT")
        Dim cProc As Long: cProc = ColumnIndex(lo, "PROCESS")
        Dim cIngId As Long: cIngId = ColumnIndex(lo, "INGREDIENT_ID")

        Dim arr As Variant: arr = lo.DataBodyRange.value
        Dim r As Long
        For r = 1 To UBound(arr, 1)
            Dim ioVal As String
            ioVal = NzStr(arr(r, cIO))
            If Not IsOutputIoValue(ioVal) Then GoTo NextRow

            Dim procName As String
            If cProc > 0 Then procName = NzStr(arr(r, cProc))
            If procName = "" Then procName = ProcessNameFromTable(lo)
            If procName <> "" Then
                If Not IsProcessSelected(procName, wsProd) Then GoTo NextRow
            End If

            Dim outputName As String
            outputName = NzStr(arr(r, cIng))
            If outputName = "" Then GoTo NextRow

            Dim uomVal As String
            If cUom > 0 Then uomVal = NzStr(arr(r, cUom))
            Dim qtyVal As Double
            If cAmt > 0 Then qtyVal = NzDbl(arr(r, cAmt))
            Dim ingId As String
            If cIngId > 0 Then ingId = NzStr(arr(r, cIngId))

            Dim key As String
            key = BuildOutputKey(procName, outputName)
            If Not entryMap.Exists(key) Then
                Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
                entry("PROCESS") = procName
                entry("OUTPUT") = outputName
                entry("UOM") = uomVal
                entry("QTY") = qtyVal
                entry("INGREDIENT_ID") = ingId
                entry("RECIPE_ID") = recipeId
                entryMap.Add key, entry
                order.Add key
            Else
                Dim existing As Object
                Set existing = entryMap(key)
                existing("QTY") = NzDbl(existing("QTY")) + qtyVal
                If NzStr(existing("UOM")) = "" Then existing("UOM") = uomVal
                If NzStr(existing("INGREDIENT_ID")) = "" Then existing("INGREDIENT_ID") = ingId
                If NzStr(existing("RECIPE_ID")) = "" Then existing("RECIPE_ID") = recipeId
            End If
NextRow:
        Next r
NextLo:
    Next lo

    If order.count = 0 Then Exit Function

    Dim result As New Collection
    Dim k As Variant
    For Each k In order
        result.Add entryMap(k)
    Next k
    Set BuildOutputEntriesFromProcessTables = result
End Function

Private Sub EnsureProductionOutputHeaderOrder(ByVal loOut As ListObject)
    If loOut Is Nothing Then Exit Sub
    If loOut.HeaderRowRange Is Nothing Then Exit Sub

    Dim cUom As Long
    cUom = ColumnIndex(loOut, "UOM")
    If cUom = 0 Then Exit Sub
    If cUom + 2 > loOut.ListColumns.count Then Exit Sub

    Dim h1 As String
    Dim h2 As String
    h1 = Trim$(NzStr(loOut.HeaderRowRange.Cells(1, cUom + 1).value))
    h2 = Trim$(NzStr(loOut.HeaderRowRange.Cells(1, cUom + 2).value))

    If StrComp(h1, "BATCH", vbTextCompare) = 0 And StrComp(h2, "REAL OUTPUT", vbTextCompare) = 0 Then
        On Error Resume Next
        loOut.ListColumns(cUom + 1).Name = "REAL OUTPUT"
        loOut.ListColumns(cUom + 2).Name = "BATCH"
        On Error GoTo 0
    End If
End Sub

Private Sub UpdateProductionOutputTable(ByVal loOut As ListObject, ByVal entries As Collection, ByVal invLo As ListObject, ByRef errNotes As String)
    If loOut Is Nothing Then Exit Sub
    If entries Is Nothing Then Exit Sub
    If entries.count = 0 Then Exit Sub

    EnsureProductionOutputHeaderOrder loOut

    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cUom As Long: cUom = ColumnIndex(loOut, "UOM")
    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cRecall As Long: cRecall = ColumnIndex(loOut, "RECALL CODE")

    If cProc = 0 Or cOutput = 0 Then
        AppendNote errNotes, "ProductionOutput missing PROCESS/OUTPUT columns."
        Exit Sub
    End If

    Dim cRow As Long
    cRow = EnsureProductionOutputRowColumn(loOut)

    Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
    If Not loOut.DataBodyRange Is Nothing Then
        Dim arr As Variant: arr = loOut.DataBodyRange.value
        Dim r As Long
        For r = 1 To UBound(arr, 1)
            Dim key As String
            key = BuildOutputKey(NzStr(arr(r, cProc)), NzStr(arr(r, cOutput)))
            If key <> "|" Then
                If Not existing.Exists(key) Then existing.Add key, r
            End If
        Next r
    End If

    Dim outputLookup As Object
    If cRow > 0 Then Set outputLookup = BuildInvSysOutputLookup(invLo)

    Dim i As Long
    Dim currentRows As Long
    If loOut.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = loOut.DataBodyRange.rows.count
    End If

    Dim addCount As Long
    For i = 1 To entries.count
        Dim entryCount As Object
        Set entryCount = entries(i)
        Dim addKey As String
        addKey = BuildOutputKey(NzStr(entryCount("PROCESS")), NzStr(entryCount("OUTPUT")))
        If addKey <> "|" Then
            If Not existing.Exists(addKey) Then addCount = addCount + 1
        End If
    Next i
    If addCount > 0 Then
        Dim emptySlots As Long
        If Not loOut.DataBodyRange Is Nothing Then
            Dim rEmpty As Long
            For rEmpty = 1 To loOut.DataBodyRange.rows.count
                Dim procVal As String
                Dim outVal As String
                If cProc > 0 Then procVal = NzStr(loOut.DataBodyRange.Cells(rEmpty, cProc).value)
                If cOutput > 0 Then outVal = NzStr(loOut.DataBodyRange.Cells(rEmpty, cOutput).value)
                If Trim$(procVal) = "" And Trim$(outVal) = "" Then emptySlots = emptySlots + 1
            Next rEmpty
        End If
        Dim needRows As Long
        needRows = addCount - emptySlots
        If needRows > 0 Then
            Call EnsureListObjectRowCountFullRow(loOut, currentRows + needRows)
        End If
    End If

    Dim NextRow As Long
    NextRow = currentRows + 1

    For i = 1 To entries.count
        Dim entry As Object
        Set entry = entries(i)
        Dim procName As String: procName = NzStr(entry("PROCESS"))
        Dim outputName As String: outputName = NzStr(entry("OUTPUT"))
        Dim uomVal As String: uomVal = NzStr(entry("UOM"))

        Dim outKey As String
        outKey = BuildOutputKey(procName, outputName)
        If outKey = "|" Then GoTo NextEntry

        Dim targetRow As Long

        If existing.Exists(outKey) Then
            targetRow = CLng(existing(outKey))
            If cProc > 0 Then loOut.DataBodyRange.Cells(targetRow, cProc).value = procName
            If cOutput > 0 Then loOut.DataBodyRange.Cells(targetRow, cOutput).value = outputName
            If cUom > 0 Then
                If NzStr(loOut.DataBodyRange.Cells(targetRow, cUom).value) = "" Then
                    loOut.DataBodyRange.Cells(targetRow, cUom).value = uomVal
                End If
            End If
        Else
            targetRow = FindFirstEmptyOutputRow(loOut, cProc, cOutput)
            If targetRow = 0 Then
                targetRow = NextRow
                NextRow = NextRow + 1
            End If
            If cProc > 0 Then loOut.DataBodyRange.Cells(targetRow, cProc).value = procName
            If cOutput > 0 Then loOut.DataBodyRange.Cells(targetRow, cOutput).value = outputName
            If cUom > 0 Then loOut.DataBodyRange.Cells(targetRow, cUom).value = uomVal
            existing.Add outKey, targetRow
        End If

        If cRow > 0 Then
            Dim rowVal As Variant
            If Not loOut.DataBodyRange Is Nothing Then
                rowVal = loOut.DataBodyRange.Cells(targetRow, cRow).value
            End If
            If NzLng(rowVal) = 0 Then
                Dim recId As String
                Dim ingId As String
                recId = NzStr(entry("RECIPE_ID"))
                ingId = NzStr(entry("INGREDIENT_ID"))
                If recId <> "" And ingId <> "" Then
                    Dim allowed As Object
                    Set allowed = GetAllowedInvRowsForIngredient(recId, ingId)
                    If Not allowed Is Nothing Then
                        Dim kRow As Variant
                        For Each kRow In allowed.keys
                            rowVal = CLng(kRow)
                            Exit For
                        Next kRow
                    End If
                End If
            End If
            If NzLng(rowVal) = 0 And Not outputLookup Is Nothing Then
                rowVal = LookupOutputRow(outputLookup, outputName)
            End If
            If NzLng(rowVal) <> 0 Then
                If Not loOut.DataBodyRange Is Nothing Then
                    loOut.DataBodyRange.Cells(targetRow, cRow).value = rowVal
                End If
            End If
        End If
NextEntry:
    Next i
End Sub

Private Function FindFirstEmptyOutputRow(ByVal loOut As ListObject, ByVal cProc As Long, ByVal cOutput As Long) As Long
    FindFirstEmptyOutputRow = 0
    If loOut Is Nothing Then Exit Function
    If loOut.DataBodyRange Is Nothing Then
        FindFirstEmptyOutputRow = 1
        Exit Function
    End If
    If cProc = 0 And cOutput = 0 Then Exit Function

    Dim r As Long
    For r = 1 To loOut.DataBodyRange.rows.count
        Dim procVal As String
        Dim outVal As String
        If cProc > 0 Then procVal = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
        If cOutput > 0 Then outVal = NzStr(loOut.DataBodyRange.Cells(r, cOutput).value)
        If Trim$(procVal) = "" And Trim$(outVal) = "" Then
            FindFirstEmptyOutputRow = r
            Exit Function
        End If
    Next r
End Function

Private Function EnsureProductionOutputRowColumn(ByVal loOut As ListObject) As Long
    If loOut Is Nothing Then Exit Function
    Dim cRow As Long: cRow = ColumnIndex(loOut, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLoose(loOut, "ROW", "ROWID", "ROW#")
    If cRow = 0 Then
        On Error Resume Next
        Dim newCol As ListColumn
        Set newCol = loOut.ListColumns.Add
        If Not newCol Is Nothing Then
            newCol.Name = "ROW"
            cRow = newCol.Index
        End If
        On Error GoTo 0
    End If
    EnsureProductionOutputRowColumn = cRow
End Function

Private Function BuildOutputKey(ByVal procName As String, ByVal outputName As String) As String
    BuildOutputKey = NormalizeOutputKey(procName) & "|" & NormalizeOutputKey(outputName)
End Function

Private Function NormalizeOutputKey(ByVal v As String) As String
    NormalizeOutputKey = LCase$(Trim$(v))
End Function

Private Function NormalizeLookupKey(ByVal v As String) As String
    Dim s As String
    s = Trim$(v)
    If s = "" Then Exit Function
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    On Error Resume Next
    s = Application.WorksheetFunction.Trim(s)
    On Error GoTo 0
    NormalizeLookupKey = LCase$(s)
End Function

Private Function IsOutputIoValue(ByVal ioVal As String) As Boolean
    Dim v As String
    v = LCase$(Trim$(ioVal))
    If v = "" Then Exit Function
    If v = "made" Then IsOutputIoValue = True
End Function

Private Function BuildInvSysOutputLookup(ByVal invLo As ListObject) As Object
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLoose(invLo, "ROW", "ROWID", "ROW#")
    Dim cItem As Long: cItem = ColumnIndex(invLo, "ITEM")
    Dim cCode As Long: cCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim cDesc As Long: cDesc = ColumnIndex(invLo, "DESCRIPTION")
    If cRow = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = invLo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As Long: rowVal = NzLng(arr(r, cRow))
        If rowVal = 0 Then GoTo NextRow
        Dim itemName As String
        Dim itemCode As String
        Dim descVal As String
        If cItem > 0 Then itemName = NzStr(arr(r, cItem))
        If cCode > 0 Then itemCode = NzStr(arr(r, cCode))
        If cDesc > 0 Then descVal = NzStr(arr(r, cDesc))
        If itemName <> "" Then
            Dim keyName As String: keyName = NormalizeLookupKey(itemName)
            If keyName <> "" Then
                If Not dict.Exists(keyName) Then dict.Add keyName, rowVal
            End If
        End If
        If itemCode <> "" Then
            Dim keyCode As String: keyCode = NormalizeLookupKey(itemCode)
            If keyCode <> "" Then
                If Not dict.Exists(keyCode) Then dict.Add keyCode, rowVal
            End If
        End If
        If descVal <> "" Then
            Dim keyDesc As String: keyDesc = NormalizeLookupKey(descVal)
            If keyDesc <> "" Then
                If Not dict.Exists(keyDesc) Then dict.Add keyDesc, rowVal
            End If
        End If
NextRow:
    Next r

    If dict.count = 0 Then Exit Function
    Set BuildInvSysOutputLookup = dict
End Function

Private Function LookupOutputRow(ByVal outputLookup As Object, ByVal outputName As String) As Long
    If outputLookup Is Nothing Then Exit Function
    Dim key As String: key = NormalizeLookupKey(outputName)
    If key = "" Then Exit Function
    If outputLookup.Exists(key) Then LookupOutputRow = CLng(outputLookup(key))
End Function

Private Function BuildUsedDeltaPacketFromInvSys(ByVal invLo As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function

    Dim colUsed As Long: colUsed = ColumnIndex(invLo, "USED")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    If colUsed = 0 Or colRow = 0 Then
        errNotes = "invSys table missing USED/ROW columns."
        Exit Function
    End If

    Dim result As New Collection
    Dim arr As Variant: arr = invLo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim usedVal As Double: usedVal = NzDbl(arr(r, colUsed))
        Dim rowVal As Long: rowVal = NzLng(arr(r, colRow))
        If rowVal = 0 Or usedVal <= 0 Then GoTo NextRow
        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = rowVal
        delta("QTY") = usedVal
        If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(arr(r, colItemCode))
        If colItemName > 0 Then delta("ITEM_NAME") = NzStr(arr(r, colItemName))
        result.Add delta
NextRow:
    Next r

    If result.count = 0 Then
        errNotes = "No staged usage found in invSys.USED."
        Exit Function
    End If
    Set BuildUsedDeltaPacketFromInvSys = result
End Function

Private Function BuildMadeDeltasFromProductionOutput(ByVal loOut As ListObject, ByVal invLo As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If loOut Is Nothing Or loOut.DataBodyRange Is Nothing Then Exit Function
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then
        AppendNote errNotes, "invSys table not found."
        Exit Function
    End If

    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cRowOut As Long: cRowOut = ColumnIndex(loOut, "ROW")
    If cRowOut = 0 Then cRowOut = ColumnIndexLoose(loOut, "ROW", "ROWID", "ROW#")

    If cReal = 0 Then
        errNotes = "ProductionOutput missing REAL OUTPUT column."
        Exit Function
    End If
    If cRowOut = 0 And cOutput = 0 Then
        errNotes = "ProductionOutput missing ROW/OUTPUT columns."
        Exit Function
    End If

    Dim rowIndex As Object
    Set rowIndex = BuildInvSysRowIndex(invLo)
    If rowIndex Is Nothing Then
        AppendNote errNotes, "invSys ROW index not available."
        Exit Function
    ElseIf rowIndex.count = 0 Then
        AppendNote errNotes, "invSys ROW index not available."
        Exit Function
    End If

    Dim outputLookup As Object
    Set outputLookup = BuildInvSysOutputLookup(invLo)

    Dim cItemCode As Long: cItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim cItemName As Long: cItemName = ColumnIndex(invLo, "ITEM")

    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loOut.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim qtyVal As Double: qtyVal = NzDbl(arr(r, cReal))
        If qtyVal <= 0 Then GoTo NextRow

        Dim rowVal As Long
        If cRowOut > 0 Then rowVal = NzLng(arr(r, cRowOut))
        If rowVal = 0 And cOutput > 0 Then
            rowVal = LookupOutputRow(outputLookup, NzStr(arr(r, cOutput)))
        End If
        If rowVal = 0 Then
            AppendNote errNotes, "Output row missing ROW: " & NzStr(arr(r, cOutput))
            GoTo NextRow
        End If
        If Not rowIndex.Exists(CStr(rowVal)) Then
            AppendNote errNotes, "invSys row " & rowVal & " not found."
            GoTo NextRow
        End If

        Dim key As String: key = CStr(rowVal)
        If agg.Exists(key) Then
            Dim existing As Object
            Set existing = agg(key)
            existing("QTY") = NzDbl(existing("QTY")) + qtyVal
        Else
            Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
            delta("ROW") = rowVal
            delta("QTY") = qtyVal
            Dim invIdx As Long: invIdx = CLng(rowIndex(CStr(rowVal)))
            If cItemCode > 0 Then delta("ITEM_CODE") = NzStr(invLo.DataBodyRange.Cells(invIdx, cItemCode).value)
            If cItemName > 0 Then delta("ITEM_NAME") = NzStr(invLo.DataBodyRange.Cells(invIdx, cItemName).value)
            agg.Add key, delta
        End If
NextRow:
    Next r

    If agg.count = 0 Then
        If errNotes = "" Then errNotes = "No made quantities found in ProductionOutput."
        Exit Function
    End If

    Dim result As New Collection
    Dim k As Variant
    For Each k In agg.keys
        result.Add agg(k)
    Next k
    Set BuildMadeDeltasFromProductionOutput = result
End Function

Private Function BuildRowKeySetFromDeltas(ByVal usedDeltas As Collection, ByVal madeDeltas As Collection) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim delta As Variant

    If Not usedDeltas Is Nothing Then
        For Each delta In usedDeltas
            On Error Resume Next
            dict(CStr(delta("ROW"))) = True
            On Error GoTo 0
        Next delta
    End If

    If Not madeDeltas Is Nothing Then
        For Each delta In madeDeltas
            On Error Resume Next
            dict(CStr(delta("ROW"))) = True
            On Error GoTo 0
        Next delta
    End If

    If dict.count = 0 Then Exit Function
    Set BuildRowKeySetFromDeltas = dict
End Function

Private Function BuildUsedSnapshotForRows(ByVal invLo As ListObject, ByVal rowKeys As Object) As Object
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    If rowKeys Is Nothing Then Exit Function
    If rowKeys.count = 0 Then Exit Function

    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLoose(invLo, "ROW", "ROWID", "ROW#")
    Dim cUsed As Long: cUsed = ColumnIndex(invLo, "USED")
    If cRow = 0 Or cUsed = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = invLo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As String: rowVal = NzStr(arr(r, cRow))
        If rowVal <> "" Then
            If rowKeys.Exists(rowVal) Then dict(rowVal) = NzDbl(arr(r, cUsed))
        End If
    Next r

    If dict.count = 0 Then Exit Function
    Set BuildUsedSnapshotForRows = dict
End Function

Private Sub WriteArrayToTable(lo As ListObject, arr As Variant)
    If lo Is Nothing Then Exit Sub
    If IsEmpty(arr) Then Exit Sub
    Dim rowsNeeded As Long
    On Error Resume Next
    rowsNeeded = UBound(arr, 1)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    If rowsNeeded <= 0 Then
        ClearListObjectData lo
        Exit Sub
    End If
    Dim currentRows As Long
    If lo.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = lo.DataBodyRange.rows.count
    End If
    Dim diff As Long
    If currentRows < rowsNeeded Then
        For diff = 1 To rowsNeeded - currentRows
            lo.ListRows.Add
        Next diff
    ElseIf currentRows > rowsNeeded Then
        For diff = rowsNeeded + 1 To currentRows
            lo.ListRows(diff).Range.ClearContents
        Next diff
    End If
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.value = arr
End Sub

Private Sub ClearListObjectContents(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableHasRow lo
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.ClearContents
End Sub

Private Function SortedKeys(dict As Object) As Variant
    If dict Is Nothing Then Exit Function
    Dim keys As Variant: keys = dict.keys
    If Not IsArray(keys) Then
        SortedKeys = keys
        Exit Function
    End If
    Dim i As Long, j As Long
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(val(keys(j))) < CLng(val(keys(i))) Then
                Dim tmp As Variant
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
    SortedKeys = keys
End Function

Private Sub AppendNote(ByRef target As String, ByVal text As String)
    If Len(text) = 0 Then Exit Sub
    If Len(target) > 0 Then
        target = target & vbCrLf & text
    Else
        target = text
    End If
End Sub

Private Sub ApplyProcessHeaderColor(ByVal lo As ListObject, ByVal procName As String)
    If lo Is Nothing Then Exit Sub
    If lo.HeaderRowRange Is Nothing Then Exit Sub
    procName = Trim$(procName)
    If procName = "" Then Exit Sub

    Dim colorVal As Long
    colorVal = ProcessColorFromName(procName)
    On Error Resume Next
    lo.HeaderRowRange.Interior.Color = colorVal
    If IsColorDark(colorVal) Then
        lo.HeaderRowRange.Font.Color = vbWhite
    Else
        lo.HeaderRowRange.Font.Color = vbBlack
    End If
    On Error GoTo 0
End Sub

Private Function ProcessColorFromName(ByVal procName As String) As Long
    Static colorMap As Object
    Static usedMap As Object
    If colorMap Is Nothing Then Set colorMap = CreateObject("Scripting.Dictionary")
    If usedMap Is Nothing Then Set usedMap = CreateObject("Scripting.Dictionary")

    Dim key As String
    key = LCase$(Trim$(procName))
    If key = "" Then Exit Function
    If colorMap.Exists(key) Then
        ProcessColorFromName = colorMap(key)
        Exit Function
    End If

    Dim palette As Variant
    palette = ProcessColorPalette()
    Dim n As Long
    n = UBound(palette) - LBound(palette) + 1

    Dim startIdx As Long
    startIdx = HashProcessName(key) Mod n
    If startIdx < 0 Then startIdx = startIdx + n

    Dim idx As Long
    idx = startIdx
    Dim attempts As Long
    Do
        Dim c As Long
        c = palette(idx)
        If Not usedMap.Exists(CStr(c)) Then
            usedMap.Add CStr(c), True
            colorMap.Add key, c
            ProcessColorFromName = c
            Exit Function
        End If
        idx = idx + 1
        If idx >= n Then idx = 0
        attempts = attempts + 1
    Loop While attempts < n

    colorMap.Add key, palette(startIdx)
    ProcessColorFromName = palette(startIdx)
End Function

Private Function ProcessColorPalette() As Variant
    ProcessColorPalette = Array( _
        RGB(33, 150, 243), _
        RGB(233, 30, 99), _
        RGB(0, 150, 136), _
        RGB(255, 152, 0), _
        RGB(156, 39, 176), _
        RGB(76, 175, 80), _
        RGB(121, 85, 72), _
        RGB(63, 81, 181), _
        RGB(205, 220, 57), _
        RGB(0, 188, 212), _
        RGB(244, 67, 54), _
        RGB(255, 193, 7))
End Function

Private Function HashProcessName(ByVal procName As String) As Long
    Dim h As Double
    Dim i As Long
    For i = 1 To Len(procName)
        Dim ch As Long
        ch = AscW(Mid$(procName, i, 1))
        If ch < 0 Then ch = ch + 65536
        h = (h * 131#) + (ch * i)
        If h >= 2147483647# Then
            h = h - 2147483647# * Fix(h / 2147483647#)
        End If
    Next i
    HashProcessName = CLng(h)
End Function

Private Function HsvToRgb(ByVal h As Double, ByVal s As Double, ByVal v As Double) As Long
    Dim r As Double, g As Double, b As Double
    Dim i As Long
    Dim f As Double, p As Double, q As Double, t As Double

    i = Int(h * 6)
    f = h * 6 - i
    p = v * (1 - s)
    q = v * (1 - f * s)
    t = v * (1 - (1 - f) * s)

    Select Case (i Mod 6)
        Case 0
            r = v: g = t: b = p
        Case 1
            r = q: g = v: b = p
        Case 2
            r = p: g = v: b = t
        Case 3
            r = p: g = q: b = v
        Case 4
            r = t: g = p: b = v
        Case 5
            r = v: g = p: b = q
    End Select

    HsvToRgb = RGB(CLng(r * 255), CLng(g * 255), CLng(b * 255))
End Function

Private Function IsColorDark(ByVal colorVal As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = colorVal Mod 256
    g = (colorVal \ 256) Mod 256
    b = (colorVal \ 65536) Mod 256

    Dim luma As Double
    luma = (0.299 * r) + (0.587 * g) + (0.114 * b)
    IsColorDark = (luma < 140)
End Function

Private Sub RenderProcessSelectorCheckboxes(ByVal ws As Worksheet, ByVal procTables As Collection)
    If ws Is Nothing Then Exit Sub
    If procTables Is Nothing Then Exit Sub

    Dim prevStates As Object
    Set prevStates = CreateObject("Scripting.Dictionary")
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_PROC_PREFIX) & "*" Then
                Dim cap As String
                cap = LCase$(Trim$(GetCheckboxCaption(shp)))
                If cap <> "" Then
                    prevStates(cap) = (shp.ControlFormat.value = 1)
                End If
            End If
        End If
    Next shp

    DeleteCheckboxesByPrefix ws, CHK_PROC_PREFIX

    If procTables.count = 0 Then Exit Sub

    Dim maxCol As Long
    Dim lo As ListObject
    For Each lo In procTables
        If Not lo Is Nothing Then
            Dim endCol As Long
            endCol = lo.Range.Column + lo.Range.Columns.count - 1
            If endCol > maxCol Then maxCol = endCol
        End If
    Next lo
    If maxCol = 0 Then Exit Sub

    Dim leftPos As Double
    leftPos = ws.Columns(maxCol + 1).Left + 2
    Const CHK_HEIGHT As Double = 16
    Const CHK_WIDTH As Double = 140

    For Each lo In procTables
        If lo Is Nothing Then GoTo NextProc
        Dim procName As String
        procName = ProcessNameFromTable(lo)
        If Trim$(procName) = "" Then procName = lo.Name

        Dim topPos As Double
        topPos = lo.HeaderRowRange.Top + 2

        Dim baseName As String
        baseName = CHK_PROC_PREFIX & SafeProcessKey(procName)
        Dim shapeName As String
        shapeName = UniqueShapeName(ws, baseName)

        Dim chk As Shape
        Set chk = EnsureCheckboxShape(ws, shapeName, procName, "mProduction.ProcessCheckboxChanged", leftPos, topPos, CHK_WIDTH, CHK_HEIGHT)
        If Not chk Is Nothing Then
            chk.AlternativeText = procName
            Dim key As String
            key = LCase$(Trim$(procName))
            If prevStates.Exists(key) Then
                chk.ControlFormat.value = IIf(prevStates(key), 1, 0)
            Else
                chk.ControlFormat.value = 0
            End If
        End If
NextProc:
    Next lo
End Sub

Private Sub RenderPaletteKeepCheckboxes(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim prevStates As Object
    Set prevStates = CreateObject("Scripting.Dictionary")
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_BATCH_PREFIX) & "*" Then
                Dim cap As String
                cap = LCase$(Trim$(GetCheckboxCaption(shp)))
                If cap = "" Then cap = LCase$(Trim$(shp.AlternativeText))
                If cap <> "" Then
                    prevStates(cap) = (shp.ControlFormat.value = 1)
                End If
            End If
        End If
    Next shp

    DeleteCheckboxesByPrefix ws, CHK_BATCH_PREFIX

    Dim maxCol As Long
    Dim lo As ListObject
    Dim paletteTables As New Collection
    For Each lo In ws.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                paletteTables.Add lo
                Dim endCol As Long
                endCol = lo.Range.Column + lo.Range.Columns.count - 1
                If endCol > maxCol Then maxCol = endCol
            End If
        End If
    Next lo
    If paletteTables.count = 0 Then Exit Sub
    If maxCol = 0 Then Exit Sub

    Dim leftPos As Double
    leftPos = ws.Columns(maxCol + 1).Left + 2
    Const CHK_HEIGHT As Double = 14
    Const CHK_WIDTH As Double = 14

    For Each lo In paletteTables
        If lo Is Nothing Then GoTo NextPal
        Dim procName As String
        Dim recipeId As String
        Dim ingId As String
        Dim amtVal As Variant
        Dim ioVal As String
        If GetPaletteTableContext(lo, recipeId, ingId, amtVal, procName, ioVal) = False Then
            procName = ProcessNameFromTable(lo)
        End If
        If Trim$(procName) = "" Then procName = lo.Name

        Dim topPos As Double
        topPos = lo.HeaderRowRange.Top + 2

        Dim shapeName As String
        shapeName = CHK_BATCH_PREFIX & SafeProcessKey(procName)

        Dim chk As Shape
        Set chk = EnsureCheckboxShape(ws, shapeName, "", "mProduction.OutputCheckboxChanged", leftPos, topPos, CHK_WIDTH, CHK_HEIGHT)
        If Not chk Is Nothing Then
            chk.AlternativeText = procName
            Dim key As String
            key = LCase$(Trim$(procName))
            If prevStates.Exists(key) Then
                chk.ControlFormat.value = IIf(prevStates(key), 1, 0)
            Else
                chk.ControlFormat.value = 0
            End If
        End If
NextPal:
    Next lo
End Sub

Private Function IsPaletteKeepSelected(ByVal ws As Worksheet, ByVal procName As String) As Boolean
    If ws Is Nothing Then Exit Function
    procName = Trim$(procName)
    If procName = "" Then Exit Function

    Dim shapeName As String
    shapeName = CHK_BATCH_PREFIX & SafeProcessKey(procName)

    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Function
    If Not IsCheckboxShape(shp) Then Exit Function

    On Error Resume Next
    IsPaletteKeepSelected = (shp.ControlFormat.value = 1)
    On Error GoTo 0
End Function

Private Sub ClearPaletteTableSelection(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cCode As Long: cCode = ColumnIndex(lo, "ITEM_CODE")
    Dim cVend As Long: cVend = ColumnIndex(lo, "VENDORS")
    Dim cVendCode As Long: cVendCode = ColumnIndex(lo, "VENDOR_CODE")
    Dim cDesc As Long: cDesc = ColumnIndex(lo, "DESCRIPTION")
    Dim cItem As Long: cItem = ColumnIndex(lo, "ITEM")
    Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
    Dim cLoc As Long: cLoc = ColumnIndex(lo, "LOCATION")
    Dim cRow As Long: cRow = ColumnIndex(lo, "ROW")

    Dim r As Long
    For r = 1 To lo.DataBodyRange.rows.count
        If cCode > 0 Then lo.DataBodyRange.Cells(r, cCode).ClearContents
        If cVend > 0 Then lo.DataBodyRange.Cells(r, cVend).ClearContents
        If cVendCode > 0 Then lo.DataBodyRange.Cells(r, cVendCode).ClearContents
        If cDesc > 0 Then lo.DataBodyRange.Cells(r, cDesc).ClearContents
        If cItem > 0 Then lo.DataBodyRange.Cells(r, cItem).ClearContents
        If cUom > 0 Then lo.DataBodyRange.Cells(r, cUom).ClearContents
        If cLoc > 0 Then lo.DataBodyRange.Cells(r, cLoc).ClearContents
        If cRow > 0 Then lo.DataBodyRange.Cells(r, cRow).ClearContents
    Next r
End Sub

Private Sub EnsurePaletteTableMetaForExistingTables(ByVal wsProd As Worksheet)
    If wsProd Is Nothing Then Exit Sub
    EnsurePaletteTableMeta

    Dim palTables As Collection
    Set palTables = GetPaletteTablesInOrder(wsProd)
    If palTables Is Nothing Then Exit Sub
    If palTables.count = 0 Then Exit Sub

    Dim needsRebuild As Boolean
    Dim lo As ListObject
    If mPaletteTableMeta Is Nothing Then
        needsRebuild = True
    Else
        For Each lo In palTables
            If Not mPaletteTableMeta.Exists(lo.Name) Then
                needsRebuild = True
                Exit For
            End If
        Next lo
    End If

    If Not needsRebuild Then Exit Sub

    Dim entries As Collection
    Set entries = BuildPaletteMetaEntries(wsProd)
    If entries Is Nothing Then Exit Sub
    If entries.count = 0 Then Exit Sub

    ClearPaletteTableMeta
    Dim used() As Boolean
    ReDim used(1 To entries.count)

    For Each lo In palTables
        Dim procName As String
        procName = ProcessNameFromTable(lo)
        If Trim$(procName) = "" Then procName = lo.Name

        Dim matchIdx As Long
        matchIdx = FindPaletteEntryIndex(entries, used, procName)
        If matchIdx = 0 Then
            matchIdx = FindFirstUnusedEntryIndex(used)
        End If

        If matchIdx > 0 Then
            mPaletteTableMeta(lo.Name) = entries(matchIdx)
            used(matchIdx) = True
        End If
    Next lo
End Sub

Private Function GetPaletteTablesInOrder(ByVal wsProd As Worksheet) As Collection
    Dim result As New Collection
    If wsProd Is Nothing Then
        Set GetPaletteTablesInOrder = result
        Exit Function
    End If

    Dim countPal As Long
    Dim lo As ListObject
    For Each lo In wsProd.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                countPal = countPal + 1
            End If
        End If
    Next lo
    If countPal = 0 Then
        Set GetPaletteTablesInOrder = result
        Exit Function
    End If

    Dim arrLo() As ListObject
    Dim arrRow() As Long
    ReDim arrLo(1 To countPal)
    ReDim arrRow(1 To countPal)

    Dim i As Long
    i = 0
    For Each lo In wsProd.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                i = i + 1
                Set arrLo(i) = lo
                arrRow(i) = lo.Range.row
            End If
        End If
    Next lo

    Dim j As Long, k As Long
    For j = 1 To countPal - 1
        For k = j + 1 To countPal
            If arrRow(k) < arrRow(j) Then
                Dim tmpRow As Long
                Dim tmpLo As ListObject
                tmpRow = arrRow(j)
                arrRow(j) = arrRow(k)
                arrRow(k) = tmpRow
                Set tmpLo = arrLo(j)
                Set arrLo(j) = arrLo(k)
                Set arrLo(k) = tmpLo
            End If
        Next k
    Next j

    For i = 1 To countPal
        result.Add arrLo(i)
    Next i

    Set GetPaletteTablesInOrder = result
End Function

Private Function BuildPaletteMetaEntries(ByVal wsProd As Worksheet) As Collection
    Dim result As New Collection
    If wsProd Is Nothing Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim recipeId As String
    recipeId = GetRecipeChooserRecipeId(wsProd)
    If Trim$(recipeId) = "" Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim wsRec As Worksheet
    Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If
    If loRecipes.DataBodyRange Is Nothing Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim cRecId As Long: cRecId = ColumnIndex(loRecipes, "RECIPE_ID")
    Dim cProc As Long: cProc = ColumnIndex(loRecipes, "PROCESS")
    Dim cIO As Long: cIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cIngId As Long: cIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cAmt As Long: cAmt = ColumnIndex(loRecipes, "AMOUNT")
    If cRecId = 0 Or cProc = 0 Or cIO = 0 Or cIngId = 0 Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim arr As Variant: arr = loRecipes.DataBodyRange.value
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            Dim ioVal As String: ioVal = UCase$(Trim$(NzStr(arr(r, cIO))))
            If ioVal = "USED" Then
                Dim ingId As String: ingId = NzStr(arr(r, cIngId))
                Dim procName As String: procName = NzStr(arr(r, cProc))
                If ingId <> "" And procName <> "" Then
                    If Not IsProcessSelected(procName, wsProd) Then GoTo NextRow
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
                        result.Add info
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
NextRow:
    Next r

    Set BuildPaletteMetaEntries = result
End Function

Private Function FindPaletteEntryIndex(ByVal entries As Collection, ByRef used() As Boolean, ByVal procName As String) As Long
    If entries Is Nothing Then Exit Function
    If procName = "" Then Exit Function

    Dim i As Long
    For i = 1 To entries.count
        If Not used(i) Then
            Dim info As Variant
            info = entries(i)
            If StrComp(NzStr(info(3)), procName, vbTextCompare) = 0 Then
                FindPaletteEntryIndex = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FindFirstUnusedEntryIndex(ByRef used() As Boolean) As Long
    Dim i As Long
    For i = LBound(used) To UBound(used)
        If Not used(i) Then
            FindFirstUnusedEntryIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function IsProcessSelected(ByVal procName As String, ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then
        IsProcessSelected = True
        Exit Function
    End If
    procName = Trim$(procName)
    If procName = "" Then
        IsProcessSelected = True
        Exit Function
    End If

    Dim hasAny As Boolean
    Dim hasChecked As Boolean
    Dim hasMatch As Boolean
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_PROC_PREFIX) & "*" Then
                hasAny = True
                If shp.ControlFormat.value = 1 Then hasChecked = True
                Dim cap As String
                cap = Trim$(GetCheckboxCaption(shp))
                If cap = "" Then cap = Trim$(shp.AlternativeText)
                If StrComp(cap, procName, vbTextCompare) = 0 Then
                    hasMatch = True
                    If shp.ControlFormat.value = 1 Then
                        IsProcessSelected = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next shp

    If Not hasAny Then
        IsProcessSelected = True
    ElseIf hasMatch Then
        IsProcessSelected = False
    ElseIf hasChecked Then
        IsProcessSelected = False
    Else
        IsProcessSelected = False
    End If
End Function

Public Sub ProcessCheckboxChanged()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    Dim recipeId As String
    recipeId = GetRecipeChooserRecipeId(ws)
    If Trim$(recipeId) = "" Then Exit Sub

    Dim wsRec As Worksheet
    Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then Exit Sub

    DeleteInventoryPaletteTables ws

    Dim procTables As Collection
    Set procTables = GetRecipeChooserProcessTables(ws)

    BuildPaletteTablesForRecipeChooser recipeId, ws, wsRec, procTables, ""
    RenderPaletteKeepCheckboxes ws
    Exit Sub
ErrHandler:
    MsgBox "Process checkbox update failed: " & Err.description, vbExclamation
End Sub

Private Function GetRecipeChooserRecipeId(ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Function
    Dim lo As ListObject
    Set lo = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If lo Is Nothing Then Exit Function
    GetRecipeChooserRecipeId = NormalizeIdFirst(FirstNonEmptyColumnValue(lo, "RECIPE_ID"))
End Function

Private Function GetRecipeChooserProcessTables(ByVal ws As Worksheet) As Collection
    Dim result As New Collection
    If ws Is Nothing Then
        Set GetRecipeChooserProcessTables = result
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeChooserProcessTable(lo) Or LCase$(lo.Name) = LCase$(TABLE_RECIPE_CHOOSER_GENERATED) Then
            result.Add lo
        End If
    Next lo
    Set GetRecipeChooserProcessTables = result
End Function

Private Sub RenderOutputRowCheckboxes(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(ws, "Prod_invSys_Check", Array("USED", "TOTAL INV"))

    Dim prevRecall As Object
    Set prevRecall = CreateObject("Scripting.Dictionary")

    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            Dim nm As String
            nm = LCase$(shp.Name)
            If nm Like LCase$(CHK_RECALL_PREFIX) & "*" Then
                Dim idxR As Long
                idxR = ParseCheckboxIndex(shp.Name, CHK_RECALL_PREFIX)
                If idxR > 0 Then prevRecall(idxR) = (shp.ControlFormat.value = 1)
            End If
        End If
    Next shp

    DeleteCheckboxesByPrefix ws, CHK_RECALL_PREFIX

    Dim rightCol As Long
    rightCol = loOut.Range.Column + loOut.Range.Columns.count - 1
    Dim baseCol As Long
    baseCol = rightCol + 1
    Dim gapCols As Long
    If Not loCheck Is Nothing Then
        gapCols = loCheck.Range.Column - rightCol - 1
    Else
        gapCols = 2
    End If
    If gapCols < 1 Then gapCols = 1

    Dim leftRecall As Double
    Dim chkWidth As Double
    leftRecall = ws.Columns(baseCol).Left + 2
    chkWidth = ws.Columns(baseCol).Width - 4
    If chkWidth < 12 Then chkWidth = 12

    Dim r As Long
    For r = 1 To loOut.DataBodyRange.rows.count
        Dim topPos As Double
        Dim heightPts As Double
        topPos = loOut.DataBodyRange.rows(r).Top + 1
        heightPts = loOut.DataBodyRange.rows(r).Height - 2
        If heightPts < 12 Then heightPts = 12

        Dim shpRecall As Shape
        Set shpRecall = EnsureCheckboxShape(ws, CHK_RECALL_PREFIX & CStr(r), "", "mProduction.OutputCheckboxChanged", leftRecall, topPos, chkWidth, heightPts)
        If Not shpRecall Is Nothing Then
            shpRecall.AlternativeText = CStr(r)
            If prevRecall.Exists(r) Then shpRecall.ControlFormat.value = IIf(prevRecall(r), 1, 0)
        End If
    Next r
End Sub

Private Sub ClearProductionOutputForNextBatch(ByVal ws As Worksheet, ByVal loOut As ListObject)
    If ws Is Nothing Then Exit Sub
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cRecall As Long: cRecall = ColumnIndex(loOut, "RECALL CODE")
    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")

    Dim nextBatchMap As Object
    Dim maxBatchMap As Object
    If cBatch > 0 And cProc > 0 Then
        Set nextBatchMap = CreateObject("Scripting.Dictionary")
        Set maxBatchMap = CreateObject("Scripting.Dictionary")

        Dim batchVal As String
        Dim procName As String
        Dim key As String
        Dim curBatch As Long
        Dim r As Long
        For r = 1 To loOut.DataBodyRange.rows.count
            procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
            If procName <> "" Then
                batchVal = NzStr(loOut.DataBodyRange.Cells(r, cBatch).value)
                If IsNumeric(batchVal) Then
                    curBatch = CLng(val(batchVal))
                    key = LCase$(procName)
                    If Not maxBatchMap.Exists(key) Then
                        maxBatchMap.Add key, curBatch
                    ElseIf curBatch > CLng(maxBatchMap(key)) Then
                        maxBatchMap(key) = curBatch
                    End If
                End If
            End If
        Next r

        For r = 1 To loOut.DataBodyRange.rows.count
            procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
            If procName <> "" Then
                key = LCase$(procName)
                If Not nextBatchMap.Exists(key) Then
                    Dim nextBatch As Long
                    If Not maxBatchMap Is Nothing Then
                        If maxBatchMap.Exists(key) Then nextBatch = CLng(maxBatchMap(key)) + 1
                    End If
                    If nextBatch = 0 Then
                        nextBatch = NextBatchSequenceForProcess(ws, loOut, procName)
                    End If
                    If nextBatch > 0 Then nextBatchMap.Add key, nextBatch
                End If
            End If
        Next r
    End If

    For r = 1 To loOut.DataBodyRange.rows.count
        If cReal > 0 Then loOut.DataBodyRange.Cells(r, cReal).ClearContents
        If cBatch > 0 Then loOut.DataBodyRange.Cells(r, cBatch).ClearContents
        If cRecall > 0 Then loOut.DataBodyRange.Cells(r, cRecall).ClearContents
    Next r

    If Not nextBatchMap Is Nothing Then
        For r = 1 To loOut.DataBodyRange.rows.count
            procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
            If procName <> "" Then
                key = LCase$(procName)
                If nextBatchMap.Exists(key) Then
                    loOut.DataBodyRange.Cells(r, cBatch).value = nextBatchMap(key)
                End If
            End If
        Next r
    End If

    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_RECALL_PREFIX) & "*" Then
                On Error Resume Next
                shp.ControlFormat.value = 0
                On Error GoTo 0
            End If
        End If
    Next shp

    EnsureOutputBatchNumbers loOut
End Sub

Private Sub LogProductionOutputToProductionLog(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal invLo As ListObject, ByRef errNotes As String)
    If wsProd Is Nothing Then Exit Sub
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim wsLog As Worksheet
    Set wsLog = SheetExists("ProductionLog")
    If wsLog Is Nothing Then
        AppendNote errNotes, "ProductionLog sheet not found."
        Exit Sub
    End If

    Dim loLog As ListObject
    Set loLog = FindListObjectByNameOrHeaders(wsLog, "ProductionLog", Array("PROCESS", "BATCH", "TIMESTAMP"))
    If loLog Is Nothing Then
        Set loLog = FindListObjectByNameOrHeaders(wsLog, "Table46", Array("PROCESS", "BATCH", "TIMESTAMP"))
    End If
    If loLog Is Nothing Then
        AppendNote errNotes, "ProductionLog table not found."
        Exit Sub
    End If

    Dim cLogRecipe As Long: cLogRecipe = ColumnIndex(loLog, "RECIPE")
    Dim cLogRecipeId As Long: cLogRecipeId = ColumnIndex(loLog, "RECIPE_ID")
    Dim cLogDept As Long: cLogDept = ColumnIndex(loLog, "DEPARTMENT")
    Dim cLogDesc As Long: cLogDesc = ColumnIndex(loLog, "DESCRIPTION")
    Dim cLogPred As Long: cLogPred = ColumnIndex(loLog, "PREDICTED OUTPUT")
    Dim cLogProc As Long: cLogProc = ColumnIndex(loLog, "PROCESS")
    Dim cLogReal As Long: cLogReal = ColumnIndex(loLog, "REAL OUTPUT")
    If cLogReal = 0 Then cLogReal = ColumnIndexLoose(loLog, "REALOUTPUT", "REAL_OUTPUT")
    Dim cLogBatch As Long: cLogBatch = ColumnIndex(loLog, "BATCH")
    Dim cLogBatchId As Long: cLogBatchId = ColumnIndex(loLog, "BATCH_ID")
    Dim cLogItemCode As Long: cLogItemCode = ColumnIndex(loLog, "ITEM_CODE")
    Dim cLogVendors As Long: cLogVendors = ColumnIndex(loLog, "VENDORS")
    Dim cLogVendCode As Long: cLogVendCode = ColumnIndex(loLog, "VENDOR_CODE")
    Dim cLogItem As Long: cLogItem = ColumnIndex(loLog, "ITEM")
    Dim cLogUom As Long: cLogUom = ColumnIndex(loLog, "UOM")
    Dim cLogQty As Long: cLogQty = ColumnIndex(loLog, "QUANTITY")
    Dim cLogLoc As Long: cLogLoc = ColumnIndex(loLog, "LOCATION")
    Dim cLogRow As Long: cLogRow = ColumnIndex(loLog, "ROW")
    Dim cLogIO As Long: cLogIO = ColumnIndex(loLog, "INPUT/OUTPUT")
    Dim cLogTime As Long: cLogTime = ColumnIndex(loLog, "TIMESTAMP")
    Dim cLogIngId As Long: cLogIngId = ColumnIndex(loLog, "INGREDIENT_ID")
    Dim cLogGuid As Long: cLogGuid = ColumnIndex(loLog, "GUID")

    Dim recipeName As String
    Dim recipeId As String
    Dim recipeDept As String
    Dim recipeDesc As String
    Dim recipePred As String
    Dim loChooser As ListObject
    Set loChooser = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If Not loChooser Is Nothing Then
        recipeName = FirstNonEmptyColumnValue(loChooser, "RECIPE")
        recipeId = FirstNonEmptyColumnValue(loChooser, "RECIPE_ID")
        recipeDept = FirstNonEmptyColumnValue(loChooser, "DEPARTMENT")
        recipeDesc = FirstNonEmptyColumnValue(loChooser, "DESCRIPTION")
        recipePred = FirstNonEmptyColumnValue(loChooser, "PREDICTED OUTPUT")
    End If

    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cUom As Long: cUom = ColumnIndex(loOut, "UOM")
    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cRow As Long: cRow = ColumnIndex(loOut, "ROW")

    If cReal = 0 Or cProc = 0 Then Exit Sub

    Dim rowIndex As Object
    If Not invLo Is Nothing Then
        Set rowIndex = BuildInvSysRowIndex(invLo)
    End If
    Dim outputLookup As Object
    If Not invLo Is Nothing Then
        Set outputLookup = BuildInvSysOutputLookup(invLo)
    End If

    Dim cInvItemCode As Long, cInvVendors As Long, cInvVendCode As Long
    Dim cInvItem As Long, cInvUom As Long, cInvLoc As Long
    If Not invLo Is Nothing Then
        cInvItemCode = ColumnIndex(invLo, "ITEM_CODE")
        cInvVendors = ColumnIndexLoose(invLo, "VENDORS", "VENDOR", "VENDOR(S)")
        cInvVendCode = ColumnIndex(invLo, "VENDOR_CODE")
        cInvItem = ColumnIndex(invLo, "ITEM")
        cInvUom = ColumnIndex(invLo, "UOM")
        cInvLoc = ColumnIndex(invLo, "LOCATION")
    End If

    Dim r As Long
    For r = 1 To loOut.DataBodyRange.rows.count
        Dim realVal As Double
        realVal = NzDbl(loOut.DataBodyRange.Cells(r, cReal).value)
        If realVal <= 0 Then GoTo NextRow

        Dim procName As String
        procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
        If procName = "" Then GoTo NextRow

        Dim outputName As String
        If cOutput > 0 Then outputName = NzStr(loOut.DataBodyRange.Cells(r, cOutput).value)

        Dim batchVal As String
        If cBatch > 0 Then batchVal = NzStr(loOut.DataBodyRange.Cells(r, cBatch).value)

        Dim rowVal As Long
        If cRow > 0 Then rowVal = NzLng(loOut.DataBodyRange.Cells(r, cRow).value)
        If rowVal = 0 Then
            If Not outputLookup Is Nothing And outputName <> "" Then
                rowVal = LookupOutputRow(outputLookup, outputName)
            End If
        End If
        If rowVal = 0 Then
            AppendNote errNotes, "Output row missing ROW: " & outputName
        End If

        Dim itemCode As String
        Dim vendors As String
        Dim vendCode As String
        Dim itemName As String
        Dim uomVal As String
        Dim locVal As String

        If rowVal > 0 And Not rowIndex Is Nothing Then
            If rowIndex.Exists(CStr(rowVal)) Then
                Dim invIdx As Long
                invIdx = CLng(rowIndex(CStr(rowVal)))
                If cInvItemCode > 0 Then itemCode = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvItemCode).value)
                If cInvVendors > 0 Then vendors = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvVendors).value)
                If cInvVendCode > 0 Then vendCode = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvVendCode).value)
                If cInvItem > 0 Then itemName = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvItem).value)
                If cInvUom > 0 Then uomVal = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvUom).value)
                If cInvLoc > 0 Then locVal = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvLoc).value)
            End If
        End If
        If itemName = "" Then itemName = outputName
        If uomVal = "" And cUom > 0 Then uomVal = NzStr(loOut.DataBodyRange.Cells(r, cUom).value)

        Dim lr As ListRow
        Set lr = loLog.ListRows.Add
        If cLogRecipe > 0 Then lr.Range.Cells(1, cLogRecipe).value = recipeName
        If cLogRecipeId > 0 Then lr.Range.Cells(1, cLogRecipeId).value = recipeId
        If cLogDept > 0 Then lr.Range.Cells(1, cLogDept).value = recipeDept
        If cLogDesc > 0 Then lr.Range.Cells(1, cLogDesc).value = recipeDesc
        If cLogPred > 0 Then lr.Range.Cells(1, cLogPred).value = recipePred
        If cLogProc > 0 Then lr.Range.Cells(1, cLogProc).value = procName
        If cLogReal > 0 Then lr.Range.Cells(1, cLogReal).value = realVal
        If cLogBatch > 0 Then lr.Range.Cells(1, cLogBatch).value = batchVal
        If cLogBatchId > 0 Then lr.Range.Cells(1, cLogBatchId).value = Format$(Date, "yyyymmdd") & "-" & batchVal
        If cLogItemCode > 0 Then lr.Range.Cells(1, cLogItemCode).value = itemCode
        If cLogVendors > 0 Then lr.Range.Cells(1, cLogVendors).value = vendors
        If cLogVendCode > 0 Then lr.Range.Cells(1, cLogVendCode).value = vendCode
        If cLogItem > 0 Then lr.Range.Cells(1, cLogItem).value = itemName
        If cLogUom > 0 Then lr.Range.Cells(1, cLogUom).value = uomVal
        If cLogQty > 0 Then lr.Range.Cells(1, cLogQty).value = realVal
        If cLogLoc > 0 Then lr.Range.Cells(1, cLogLoc).value = locVal
        If cLogRow > 0 Then lr.Range.Cells(1, cLogRow).value = rowVal
        If cLogIO > 0 Then lr.Range.Cells(1, cLogIO).value = "MADE"
        If cLogTime > 0 Then lr.Range.Cells(1, cLogTime).value = Now
        If cLogIngId > 0 Then lr.Range.Cells(1, cLogIngId).value = ""
        If cLogGuid > 0 Then lr.Range.Cells(1, cLogGuid).value = modUR_Snapshot.GenerateGUID()
NextRow:
    Next r
End Sub

Private Sub ApplyRecallCodesForOutput(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal invLo As ListObject, ByRef errNotes As String)
    If wsProd Is Nothing Then Exit Sub
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim recallRows As Object
    Set recallRows = GetRecallCheckedRows(wsProd)
    If recallRows Is Nothing Then Exit Sub
    If recallRows.count = 0 Then Exit Sub

    Dim cRecall As Long: cRecall = ColumnIndex(loOut, "RECALL CODE")
    If cRecall = 0 Then Exit Sub

    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cUom As Long: cUom = ColumnIndex(loOut, "UOM")
    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cRow As Long: cRow = ColumnIndex(loOut, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLoose(loOut, "ROW", "ROWID", "ROW#")

    Dim recipeName As String
    Dim recipeId As String
    GetRecipeChooserInfo wsProd, recipeName, recipeId

    Dim wsLog As Worksheet
    Set wsLog = SheetExists("BatchCodesLog")
    If wsLog Is Nothing Then Set wsLog = SheetExists("BatchCodeLogs")

    Dim loLog As ListObject
    If Not wsLog Is Nothing Then
        Set loLog = FindListObjectByNameOrHeaders(wsLog, "Table48", Array("RECIPE", "RECIPE_ID", "PROCESS", "OUTPUT"))
    End If

    Dim cLogRec As Long, cLogRecId As Long, cLogProc As Long, cLogOut As Long
    Dim cLogUom As Long, cLogReal As Long, cLogBatch As Long, cLogRecall As Long
    Dim cLogTime As Long, cLogLoc As Long, cLogUser As Long, cLogGuid As Long
    If Not loLog Is Nothing Then
        cLogRec = ColumnIndex(loLog, "RECIPE")
        cLogRecId = ColumnIndex(loLog, "RECIPE_ID")
        cLogProc = ColumnIndex(loLog, "PROCESS")
        cLogOut = ColumnIndex(loLog, "OUTPUT")
        cLogUom = ColumnIndex(loLog, "UOM")
        cLogReal = ColumnIndex(loLog, "REAL OUTPUT")
        If cLogReal = 0 Then cLogReal = ColumnIndexLoose(loLog, "REALOUTPUT", "REAL_OUTPUT")
        cLogBatch = ColumnIndex(loLog, "BATCH")
        cLogRecall = ColumnIndex(loLog, "RECALL CODE")
        cLogTime = ColumnIndex(loLog, "TIMESTAMP")
        cLogLoc = ColumnIndex(loLog, "LOCATION")
        cLogUser = ColumnIndex(loLog, "USER")
        cLogGuid = ColumnIndex(loLog, "GUID")
    End If

    Dim key As Variant
    For Each key In recallRows.keys
        Dim idx As Long: idx = CLng(key)
        If idx < 1 Or idx > loOut.DataBodyRange.rows.count Then GoTo NextRow

        Dim codeVal As String
        codeVal = NzStr(loOut.DataBodyRange.Cells(idx, cRecall).value)
        If Trim$(codeVal) = "" Then
            codeVal = GenerateRecallCode()
            loOut.DataBodyRange.Cells(idx, cRecall).value = codeVal

            If Not loLog Is Nothing Then
                Dim lr As ListRow: Set lr = loLog.ListRows.Add
                If cLogRec > 0 Then lr.Range.Cells(1, cLogRec).value = recipeName
                If cLogRecId > 0 Then lr.Range.Cells(1, cLogRecId).value = recipeId
                If cLogProc > 0 And cProc > 0 Then lr.Range.Cells(1, cLogProc).value = loOut.DataBodyRange.Cells(idx, cProc).value
                If cLogOut > 0 And cOutput > 0 Then lr.Range.Cells(1, cLogOut).value = loOut.DataBodyRange.Cells(idx, cOutput).value
                If cLogUom > 0 And cUom > 0 Then lr.Range.Cells(1, cLogUom).value = loOut.DataBodyRange.Cells(idx, cUom).value
                If cLogReal > 0 And cReal > 0 Then lr.Range.Cells(1, cLogReal).value = loOut.DataBodyRange.Cells(idx, cReal).value
                If cLogBatch > 0 And cBatch > 0 Then lr.Range.Cells(1, cLogBatch).value = loOut.DataBodyRange.Cells(idx, cBatch).value
                If cLogRecall > 0 Then lr.Range.Cells(1, cLogRecall).value = codeVal
                If cLogTime > 0 Then lr.Range.Cells(1, cLogTime).value = Now
                If cLogUser > 0 Then lr.Range.Cells(1, cLogUser).value = Environ$("USERNAME")
                If cLogGuid > 0 Then lr.Range.Cells(1, cLogGuid).value = modUR_Snapshot.GenerateGUID()
                If cLogLoc > 0 Then
                    Dim locVal As String
                    If cRow > 0 Then locVal = ResolveInvSysLocationByRow(invLo, NzLng(loOut.DataBodyRange.Cells(idx, cRow).value))
                    If locVal <> "" Then lr.Range.Cells(1, cLogLoc).value = locVal
                End If
            End If
        End If
NextRow:
    Next key
End Sub

Private Function GetRecallCheckedRows(ByVal ws As Worksheet) As Object
    If ws Is Nothing Then Exit Function
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_RECALL_PREFIX) & "*" Then
                Dim idx As Long
                idx = ParseCheckboxIndex(shp.Name, CHK_RECALL_PREFIX)
                If idx > 0 Then
                    If shp.ControlFormat.value = 1 Then dict(CStr(idx)) = True
                End If
            End If
        End If
    Next shp
    If dict.count = 0 Then Exit Function
    Set GetRecallCheckedRows = dict
End Function

Private Function GenerateRecallCode() As String
    Dim guidVal As String
    guidVal = Replace(modUR_Snapshot.GenerateGUID(), "-", "")
    GenerateRecallCode = "RC-" & Left$(guidVal, 12)
End Function

Private Function GenerateBatchNumber(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal procName As String) As Long
    GenerateBatchNumber = NextBatchSequenceForProcess(wsProd, loOut, procName)
End Function

Private Sub EnsureOutputBatchNumbers(ByVal loOut As ListObject)
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim wsProd As Worksheet
    Set wsProd = loOut.Parent

    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    If cBatch = 0 Or cProc = 0 Then Exit Sub

    Dim batchMap As Object
    Set batchMap = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To loOut.DataBodyRange.rows.count
        Dim procName As String
        procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
        If procName <> "" Then
            Dim existingBatch As String
            existingBatch = NzStr(loOut.DataBodyRange.Cells(r, cBatch).value)
            If IsNumeric(existingBatch) Then
                batchMap(LCase$(procName)) = CStr(CLng(val(existingBatch)))
            End If
        End If
    Next r

    For r = 1 To loOut.DataBodyRange.rows.count
        Dim batchVal As String
        batchVal = NzStr(loOut.DataBodyRange.Cells(r, cBatch).value)
        If batchVal = "" Or Not IsNumeric(batchVal) Then
            Dim procName2 As String
            procName2 = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
            If procName2 <> "" Then
                Dim key As String
                key = LCase$(procName2)
                If batchMap.Exists(key) Then
                    loOut.DataBodyRange.Cells(r, cBatch).value = batchMap(key)
                Else
                    Dim newBatch As Long
                    newBatch = GenerateBatchNumber(wsProd, loOut, procName2)
                    If newBatch > 0 Then
                        loOut.DataBodyRange.Cells(r, cBatch).value = newBatch
                        batchMap(key) = CStr(newBatch)
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Function NextBatchSequenceForProcess(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal procName As String) As Long
    Dim maxBatch As Long
    maxBatch = MaxBatchFromOutput(loOut, procName)

    Dim wsLog As Worksheet
    Dim loLog As ListObject

    Set wsLog = SheetExists("BatchCodesLog")
    If wsLog Is Nothing Then Set wsLog = SheetExists("BatchCodeLogs")
    If Not wsLog Is Nothing Then
        Set loLog = FindListObjectByNameOrHeaders(wsLog, "Table48", Array("PROCESS", "BATCH", "TIMESTAMP"))
        If Not loLog Is Nothing Then
            AccumulateBatchMaxFromLog loLog, procName, maxBatch
        End If
    End If

    Dim wsProdLog As Worksheet
    Set wsProdLog = SheetExists("ProductionLog")
    If Not wsProdLog Is Nothing Then
        Dim loProdLog As ListObject
        Set loProdLog = FindListObjectByNameOrHeaders(wsProdLog, "ProductionLog", Array("PROCESS", "BATCH", "TIMESTAMP"))
        If loProdLog Is Nothing Then
            Set loProdLog = FindListObjectByNameOrHeaders(wsProdLog, "Table46", Array("PROCESS", "BATCH", "TIMESTAMP"))
        End If
        If Not loProdLog Is Nothing Then
            AccumulateBatchMaxFromLog loProdLog, procName, maxBatch
        End If
    End If

    NextBatchSequenceForProcess = maxBatch + 1
End Function

Private Function MaxBatchFromOutput(ByVal loOut As ListObject, ByVal procName As String) As Long
    If loOut Is Nothing Then Exit Function
    If loOut.DataBodyRange Is Nothing Then Exit Function

    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    If cBatch = 0 Or cProc = 0 Then Exit Function

    Dim arr As Variant: arr = loOut.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If StrComp(NzStr(arr(r, cProc)), procName, vbTextCompare) = 0 Then
            Dim b As Long
            b = CLng(val(arr(r, cBatch)))
            If b > MaxBatchFromOutput Then MaxBatchFromOutput = b
        End If
    Next r
End Function

Private Sub AccumulateBatchMaxFromLog(ByVal loLog As ListObject, ByVal procName As String, ByRef maxBatch As Long)
    If loLog Is Nothing Then Exit Sub
    If loLog.DataBodyRange Is Nothing Then Exit Sub

    Dim cBatch As Long: cBatch = ColumnIndex(loLog, "BATCH")
    Dim cProc As Long: cProc = ColumnIndex(loLog, "PROCESS")
    Dim cTime As Long: cTime = ColumnIndex(loLog, "TIMESTAMP")
    If cBatch = 0 Or cTime = 0 Then Exit Sub

    Dim arr As Variant: arr = loLog.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If cProc > 0 Then
            If StrComp(NzStr(arr(r, cProc)), procName, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        Dim tVal As Variant
        tVal = arr(r, cTime)
        If Not IsDate(tVal) Then GoTo NextRow
        If DateValue(tVal) <> Date Then GoTo NextRow

        Dim b As Long
        b = CLng(val(arr(r, cBatch)))
        If b > maxBatch Then maxBatch = b
NextRow:
    Next r
End Sub

Private Sub GetRecipeChooserInfo(ByVal ws As Worksheet, ByRef recipeName As String, ByRef recipeId As String)
    recipeName = ""
    recipeId = ""
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject
    Set lo = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If lo Is Nothing Then Exit Sub
    recipeName = FirstNonEmptyColumnValue(lo, "RECIPE")
    recipeId = FirstNonEmptyColumnValue(lo, "RECIPE_ID")
End Sub

Private Function ResolveInvSysLocationByRow(ByVal invLo As ListObject, ByVal invRow As Long) As String
    If invLo Is Nothing Then Exit Function
    If invRow <= 0 Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    Dim cLoc As Long: cLoc = ColumnIndex(invLo, "LOCATION")
    If cRow = 0 Or cLoc = 0 Then Exit Function

    Dim cel As Range
    For Each cel In invLo.ListColumns(cRow).DataBodyRange.Cells
        If NzLng(cel.value) = invRow Then
            ResolveInvSysLocationByRow = NzStr(cel.Offset(0, cLoc - cel.Column).value)
            Exit Function
        End If
    Next cel
End Function

Public Sub OutputCheckboxChanged()
    ' Placeholder for batch/recall checkbox behavior.
End Sub

Private Function ParseCheckboxIndex(ByVal shapeName As String, ByVal prefix As String) As Long
    If LCase$(Left$(shapeName, Len(prefix))) <> LCase$(prefix) Then Exit Function
    Dim tail As String
    tail = Mid$(shapeName, Len(prefix) + 1)
    If tail = "" Then Exit Function
    If IsNumeric(tail) Then ParseCheckboxIndex = CLng(val(tail))
End Function

Private Sub DeleteCheckboxesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    If ws Is Nothing Then Exit Sub
    Dim toDelete As Collection
    Set toDelete = New Collection
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(prefix) & "*" Then
                toDelete.Add shp.Name
            End If
        End If
    Next shp
    Dim nameVal As Variant
    For Each nameVal In toDelete
        On Error Resume Next
        ws.shapes(CStr(nameVal)).Delete
        On Error GoTo 0
    Next nameVal
End Sub

Private Function EnsureCheckboxShape(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal onActionMacro As String, _
    ByVal leftPos As Double, ByVal topPos As Double, ByVal widthPts As Double, ByVal heightPts As Double) As Shape

    If ws Is Nothing Then Exit Function
    If widthPts < 10 Then widthPts = 10
    If heightPts < 10 Then heightPts = 10

    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.shapes(shapeName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        If Not IsCheckboxShape(shp) Then Set shp = Nothing
    End If

    If shp Is Nothing Then
        Set shp = ws.shapes.AddFormControl(xlCheckBox, leftPos, topPos, widthPts, heightPts)
        shp.Name = shapeName
    Else
        shp.Name = shapeName
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = heightPts
    End If

    If onActionMacro <> "" Then shp.OnAction = onActionMacro
    ForceCheckboxCaption shp, caption
    Set EnsureCheckboxShape = shp
End Function

Private Sub ForceCheckboxCaption(ByVal shp As Shape, ByVal caption As String)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.ControlFormat.caption = caption
    shp.TextFrame.Characters.text = caption
    On Error GoTo 0
End Sub

Private Function IsCheckboxShape(ByVal shp As Shape) As Boolean
    If shp Is Nothing Then Exit Function
    If shp.Type <> msoFormControl Then Exit Function
    On Error Resume Next
    If shp.FormControlType = xlCheckBox Then IsCheckboxShape = True
    On Error GoTo 0
End Function

Private Function GetCheckboxCaption(ByVal shp As Shape) As String
    If shp Is Nothing Then Exit Function
    On Error Resume Next
    GetCheckboxCaption = shp.ControlFormat.caption
    If GetCheckboxCaption = "" Then GetCheckboxCaption = shp.TextFrame.Characters.text
    If GetCheckboxCaption = "" Then GetCheckboxCaption = shp.AlternativeText
    On Error GoTo 0
End Function

Private Function UniqueShapeName(ByVal ws As Worksheet, ByVal baseName As String) As String
    Dim nameTry As String
    nameTry = baseName
    Dim idx As Long
    idx = 1
    Do While ShapeExists(ws, nameTry)
        nameTry = baseName & "_" & CStr(idx)
        idx = idx + 1
    Loop
    UniqueShapeName = nameTry
End Function

Private Function ShapeExists(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.shapes(shapeName)
    ShapeExists = Not shp Is Nothing
    On Error GoTo 0
End Function


Private Sub FillPaletteTableFromInvSys(ByVal lo As ListObject, ByVal rowMap As Object)
    If lo Is Nothing Then Exit Sub
    If rowMap Is Nothing Then Exit Sub
    If rowMap.count = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cRow As Long: cRow = ColumnIndex(lo, "ROW")
    If cRow = 0 Then cRow = ColumnIndexLoose(lo, "ROW", "ROWID", "ROW#")
    If cRow = 0 Then Exit Sub
    Dim cCode As Long: cCode = ColumnIndex(lo, "ITEM_CODE")
    If cCode = 0 Then cCode = ColumnIndexLoose(lo, "ITEM_CODE", "ITEMCODE", "ITEM CODE")
    Dim cVend As Long: cVend = ColumnIndex(lo, "VENDORS")
    If cVend = 0 Then cVend = ColumnIndexLoose(lo, "VENDORS", "VENDOR", "VENDOR(S)")
    Dim cVendCode As Long: cVendCode = ColumnIndex(lo, "VENDOR_CODE")
    If cVendCode = 0 Then cVendCode = ColumnIndexLoose(lo, "VENDOR_CODE", "VENDORCODE", "VENDOR CODE")
    Dim cDesc As Long: cDesc = ColumnIndex(lo, "DESCRIPTION")
    If cDesc = 0 Then cDesc = ColumnIndexLoose(lo, "DESCRIPTION", "DESC")
    Dim cItem As Long: cItem = ColumnIndex(lo, "ITEM")
    If cItem = 0 Then cItem = ColumnIndexLoose(lo, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME")
    Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
    If cUom = 0 Then cUom = ColumnIndexLoose(lo, "UOM", "UNIT", "UNITOFMEASURE", "UNITOFMEASUREMENT")
    Dim cLoc As Long: cLoc = ColumnIndex(lo, "LOCATION")
    If cLoc = 0 Then cLoc = ColumnIndexLoose(lo, "LOCATION", "LOC")

    Dim r As Long
    For r = 1 To lo.DataBodyRange.rows.count
        Dim rowKey As String
        rowKey = NormalizeRowKey(lo.DataBodyRange.Cells(r, cRow).value)
        If rowKey <> "" Then
            If rowMap.Exists(rowKey) Then
                Dim info As Variant
                info = rowMap(rowKey)
                If cCode > 0 And NzStr(lo.DataBodyRange.Cells(r, cCode).value) = "" Then lo.DataBodyRange.Cells(r, cCode).value = info(1)
                If cVend > 0 And NzStr(lo.DataBodyRange.Cells(r, cVend).value) = "" Then lo.DataBodyRange.Cells(r, cVend).value = info(2)
                If cVendCode > 0 And NzStr(lo.DataBodyRange.Cells(r, cVendCode).value) = "" Then lo.DataBodyRange.Cells(r, cVendCode).value = info(3)
                If cDesc > 0 And NzStr(lo.DataBodyRange.Cells(r, cDesc).value) = "" Then lo.DataBodyRange.Cells(r, cDesc).value = info(4)
                If cItem > 0 And NzStr(lo.DataBodyRange.Cells(r, cItem).value) = "" Then lo.DataBodyRange.Cells(r, cItem).value = info(5)
                If cUom > 0 And NzStr(lo.DataBodyRange.Cells(r, cUom).value) = "" Then lo.DataBodyRange.Cells(r, cUom).value = info(6)
                If cLoc > 0 And NzStr(lo.DataBodyRange.Cells(r, cLoc).value) = "" Then lo.DataBodyRange.Cells(r, cLoc).value = info(7)
            End If
        End If
    Next r
End Sub

Private Function GetIngredientPaletteRows(ByVal recipeId As String, ByVal ingredientId As String) As Collection
    Dim wsPal As Worksheet: Set wsPal = SheetExists("IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = SheetExists("IngredientsPalette")
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

    Dim normRec As String: normRec = NormalizeIdFirst(recipeId)
    Dim normIng As String: normIng = NormalizeIdLast(ingredientId)

    Dim col As New Collection
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loPal.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NormalizeIdFirst(NzStr(arr(r, cRec))) = normRec _
            And NormalizeIdLast(NzStr(arr(r, cIng))) = normIng Then
            Dim rowKey As String
            rowKey = NormalizeRowKey(arr(r, cRow))
            If rowKey <> "" Then
                If Not seen.Exists(rowKey) Then
                    seen.Add rowKey, True
                    col.Add rowKey
                End If
            End If
        End If
    Next r

    If col.count = 0 Then Exit Function
    Set GetIngredientPaletteRows = col
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
    If wsPal Is Nothing Then Set wsPal = SheetExists("IngredientsPalette")
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
    Dim arr As Variant: arr = loPal.DataBodyRange.value
    Dim r As Long
    Dim normRec As String: normRec = NormalizeIdFirst(recipeId)
    Dim normIng As String: normIng = NormalizeIdLast(ingredientId)
    For r = 1 To UBound(arr, 1)
        If NormalizeIdFirst(NzStr(arr(r, cRec))) = normRec And NormalizeIdLast(NzStr(arr(r, cIng))) = normIng Then
            Dim rowVal As String
            rowVal = NzStr(arr(r, cRow))
            If Trim$(rowVal) <> "" Then
                If Not dict.Exists(rowVal) Then dict.Add rowVal, True
            End If
        End If
    Next r

    If dict.count = 0 Then Exit Function
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

    Dim arr As Variant: arr = lo.DataBodyRange.value
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
    Dim arr As Variant: arr = lo.DataBodyRange.value
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
    If sourceTables.count = 0 Then
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

    Dim recipeName As String: recipeName = NzStr(nameCell.value)
    Dim recipeDesc As String
    If cDesc > 0 Then
        Dim descCell As Range
        Set descCell = GetHeaderDataCell(loHeader, "DESCRIPTION")
        If Not descCell Is Nothing Then recipeDesc = NzStr(descCell.value)
    End If
    If Trim$(recipeName) = "" Then
        MsgBox "Fill RB_AddRecipeName (RECIPE_NAME) or load a recipe before saving.", vbExclamation
        Exit Sub
    End If

    Dim recipeIdCell As Range: Set recipeIdCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
    Dim recipeId As String: recipeId = NzStr(recipeIdCell.value)
    If recipeId = "" Then
        recipeId = modUR_Snapshot.GenerateGUID()
        recipeIdCell.value = recipeId
    End If
    If cGuid > 0 Then
        Dim recipeGuidCell As Range: Set recipeGuidCell = GetHeaderDataCell(loHeader, "GUID")
        Dim recipeGuid As String: recipeGuid = NzStr(recipeGuidCell.value)
        If recipeGuid = "" Then
            recipeGuid = modUR_Snapshot.GenerateGUID()
            recipeGuidCell.value = recipeGuid
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
        For r = loRecipes.DataBodyRange.rows.count To 1 Step -1
            If NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipeId).value) = recipeId Then
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
        If processTables.count > 0 Then templateCount = RegisterRecipeTemplates(recipeId, processTables)
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
    MsgBox "Save Recipe failed: " & Err.description, vbCritical
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
                recipeId = NzStr(loSel.DataBodyRange.Cells(Application.ActiveCell.row - loSel.DataBodyRange.row + 1, cSelRecipeId).value)
            End If
            If recipeId = "" And cSelRecipe > 0 Then
                recipeName = NzStr(loSel.DataBodyRange.Cells(Application.ActiveCell.row - loSel.DataBodyRange.row + 1, cSelRecipe).value)
            End If
        End If
    End If

    If recipeId = "" Then
        Dim cHeaderRecipeIdTmp As Long: cHeaderRecipeIdTmp = ColumnIndex(loHeader, "RECIPE_ID")
    If cHeaderRecipeIdTmp > 0 Then
        Dim hdrRecipeIdCell As Range: Set hdrRecipeIdCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
        If Not hdrRecipeIdCell Is Nothing Then recipeId = NzStr(hdrRecipeIdCell.value)
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
        For r = 1 To loRecipes.DataBodyRange.rows.count
            Dim rowRecipeId As String
            rowRecipeId = NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipeId).value)
            Dim rowRecipeName As String
            If cRecRecipe > 0 Then rowRecipeName = NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipe).value)
            If (recipeId <> "" And rowRecipeId = recipeId) Or (recipeId = "" And rowRecipeName = recipeName And rowRecipeName <> "") Then
                matches.Add r
                If recipeId = "" Then recipeId = rowRecipeId
                If recipeName = "" Then recipeName = rowRecipeName
            End If
        Next r
    End If

    If matches.count = 0 Then
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
    If Not hdrNameCell Is Nothing Then hdrNameCell.value = recipeName
    If Not hdrIdCell Is Nothing Then hdrIdCell.value = recipeId
    If Not hdrDescCell Is Nothing And cRecDesc > 0 Then
        hdrDescCell.value = NzStr(loRecipes.DataBodyRange.Cells(matches(1), cRecDesc).value)
    End If
    If Not hdrGuidCell Is Nothing And cRecGuid > 0 Then
        hdrGuidCell.value = NzStr(loRecipes.DataBodyRange.Cells(matches(1), cRecGuid).value)
    End If

    ' Clear and rebuild RecipeBuilder lines.
    ClearListObjectData loLines
    Dim idx As Long
    For idx = 1 To matches.count
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

        If cProc > 0 Then lr.Range.Cells(1, cProc).value = loRecipes.DataBodyRange.Cells(rr, cRecProcess).value
        If cDiag > 0 Then lr.Range.Cells(1, cDiag).value = loRecipes.DataBodyRange.Cells(rr, cRecDiagram).value
        If cIO > 0 Then lr.Range.Cells(1, cIO).value = loRecipes.DataBodyRange.Cells(rr, cRecIO).value
        If cIng > 0 Then lr.Range.Cells(1, cIng).value = loRecipes.DataBodyRange.Cells(rr, cRecIngredient).value
        If cPct > 0 Then lr.Range.Cells(1, cPct).value = loRecipes.DataBodyRange.Cells(rr, cRecPercent).value
        If cUomLine > 0 Then lr.Range.Cells(1, cUomLine).value = loRecipes.DataBodyRange.Cells(rr, cRecUom).value
        If cAmt > 0 Then lr.Range.Cells(1, cAmt).value = loRecipes.DataBodyRange.Cells(rr, cRecAmount).value
        If cListRow > 0 Then lr.Range.Cells(1, cListRow).value = loRecipes.DataBodyRange.Cells(rr, cRecListRow).value
        If cIngId > 0 Then lr.Range.Cells(1, cIngId).value = loRecipes.DataBodyRange.Cells(rr, cRecIngId).value
        If cGuidLine > 0 Then lr.Range.Cells(1, cGuidLine).value = loRecipes.DataBodyRange.Cells(rr, cRecGuid).value
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
    loadMsg = "Loaded recipe '" & recipeName & "' (" & matches.count & " lines)."
    If procCount > 0 Then loadMsg = loadMsg & vbCrLf & "Process tables built: " & procCount & "."
    MsgBox loadMsg, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Load Recipe failed: " & Err.description, vbCritical
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

    Dim lineArr As Variant: lineArr = loSource.DataBodyRange.value
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
            If cIngId > 0 Then loSource.DataBodyRange.Cells(i, cIngId).value = ingId
        End If

        Dim recListRow As Variant
        If cListRow > 0 Then recListRow = lineArr(i, cListRow)
        If NzStr(recListRow) = "" Then
            recListRow = seqRow
            If cListRow > 0 Then loSource.DataBodyRange.Cells(i, cListRow).value = recListRow
        End If

        Dim rowGuid As String
        If cGuidLine > 0 Then rowGuid = NzStr(lineArr(i, cGuidLine))
        If rowGuid = "" Then
            rowGuid = modUR_Snapshot.GenerateGUID()
            If cGuidLine > 0 Then loSource.DataBodyRange.Cells(i, cGuidLine).value = rowGuid
        End If

        Dim lr As ListRow: Set lr = loRecipes.ListRows.Add
        If cRecRecipeId > 0 Then lr.Range.Cells(1, cRecRecipeId).value = recipeId
        If cRecRecipe > 0 Then lr.Range.Cells(1, cRecRecipe).value = recipeName
        If cRecDesc > 0 Then lr.Range.Cells(1, cRecDesc).value = recipeDesc
        If cRecDept > 0 Then lr.Range.Cells(1, cRecDept).value = "" ' optional for now
        If cRecProcess > 0 Then lr.Range.Cells(1, cRecProcess).value = processVal
        If cRecDiagram > 0 And cDiag > 0 Then lr.Range.Cells(1, cRecDiagram).value = lineArr(i, cDiag)
        If cRecIO > 0 And cIO > 0 Then lr.Range.Cells(1, cRecIO).value = lineArr(i, cIO)
        If cRecIngredient > 0 And cIng > 0 Then lr.Range.Cells(1, cRecIngredient).value = lineArr(i, cIng)
        If cRecPercent > 0 And cPct > 0 Then lr.Range.Cells(1, cRecPercent).value = lineArr(i, cPct)
        If cRecUom > 0 And cUomLine > 0 Then lr.Range.Cells(1, cRecUom).value = lineArr(i, cUomLine)
        If cRecAmount > 0 And cAmt > 0 Then lr.Range.Cells(1, cRecAmount).value = lineArr(i, cAmt)
        If cRecListRow > 0 Then lr.Range.Cells(1, cRecListRow).value = recListRow
        If cRecIngId > 0 Then lr.Range.Cells(1, cRecIngId).value = ingId
        If cRecGuid > 0 Then lr.Range.Cells(1, cRecGuid).value = rowGuid

        savedCount = savedCount + 1
        seqRow = seqRow + 1
NextLine:
    Next i
End Sub

Private Function BuildRecipeProcessTablesFromLines(ByVal recipeId As String, Optional ByVal ApplyTemplates As Boolean = False, Optional ByVal anchorBelowLines As Boolean = True) As Long
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

    Dim lineArr As Variant: lineArr = loLines.DataBodyRange.value
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

    If procOrder.count = 0 Then Exit Function

    DeleteRecipeProcessTables wsProd

    Dim created As New Collection

    Dim procKey As Variant
    Dim nextSeq As Long
    nextSeq = NextRecipeProcessSequence(wsProd)
    For Each procKey In procOrder
        Dim rowsColl As Collection: Set rowsColl = procMap(procKey)
        Dim dataCount As Long: dataCount = rowsColl.count
        If dataCount = 0 Then GoTo NextProc

        Dim tableRange As Range
        Set tableRange = wsProd.Range(wsProd.Cells(startRow, startCol), wsProd.Cells(startRow + dataCount, startCol + colCount - 1))

        If RangeHasListObjectCollision(wsProd, tableRange, loLines) Then
            MsgBox "Not enough space below Recipe Builder to create process tables. Clear space and try again.", vbExclamation
            Exit Function
        End If

        tableRange.Clear
        tableRange.rows(1).value = HeaderRowArray(headerNames)

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

        tableRange.Offset(1, 0).Resize(dataCount, colCount).value = dataArr

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

    BuildRecipeProcessTablesFromLines = created.count

    If ApplyTemplates And created.count > 0 And recipeId <> "" Then
        Dim tpl As New cTemplateApplier
        Dim loProc As ListObject
        For Each loProc In created
            Dim procNameTpl As String: procNameTpl = ProcessNameFromTable(loProc)
            tpl.ApplyTemplates loProc, TEMPLATE_SCOPE_RECIPE_PROCESS, procNameTpl, "", recipeId
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
    tableRange.rows(1).value = HeaderRowArray(headers)

    Dim cProc As Long
    cProc = HeaderIndex(headers, "PROCESS")
    If cProc > 0 Then
        tableRange.Offset(1, cProc - 1).value = processName
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
            bottom = lo.Range.row + lo.Range.rows.count - 1
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
    startRow = loHeader.Range.row + loHeader.Range.rows.count + 3 ' keep 2 blank rows before first process table

    If includeLines Then
        Dim loLines As ListObject
        Set loLines = GetRecipeBuilderLinesTable(ws, loHeader)
        If Not loLines Is Nothing Then
            Dim linesBottom As Long
            linesBottom = loLines.Range.row + loLines.Range.rows.count - 1
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
    startRow = loHeader.Range.row + loHeader.Range.rows.count + 2 ' one blank row below header
    startCol = loHeader.Range.Column

    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 1, startCol + colCount - 1))
    If RangeHasListObjectCollisionStrict(ws, tableRange, loHeader) Then Exit Function

    tableRange.Clear
    tableRange.rows(1).value = HeaderRowArray(headers)

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
        headerBottom = loHeader.Range.row + loHeader.Range.rows.count - 1
    End If

    Dim candidate As ListObject
    Dim bestRow As Long
    For Each lo In ws.ListObjects
        If ListObjectHasHeaders(lo, Array("PROCESS", "INGREDIENT")) Then
            If IsRecipeProcessTable(lo) Then GoTo NextLo
            If headerStartCol > 0 Then
                If lo.Range.Column <> headerStartCol Then GoTo NextLo
                If lo.Range.row < headerBottom Then GoTo NextLo
            End If
            If bestRow = 0 Or lo.Range.row < bestRow Then
                Set candidate = lo
                bestRow = lo.Range.row
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
    Dim arr As Variant: arr = loLines.DataBodyRange.value
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
    IsRecipeLinesStaged = (loLines.Range.row >= RECIPE_LINES_STAGING_ROW)
End Function

Private Function MoveRecipeBuilderLinesToStaging(ByVal loLines As ListObject) As Boolean
    ' System 1: Recipe List Builder - move lines table out of view before building process tables.
    If loLines Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = loLines.Parent
    Dim startRow As Long: startRow = RECIPE_LINES_STAGING_ROW
    If loLines.Range.row >= startRow Then
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
        If lo.Range.row < startRow Then
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
    tableRange.rows(1).value = HeaderRowArray(headers)

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
    maxRow = ws.rows.count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range

    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate, loLines) Then
            Set FindAvailableRecipeProcessRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 3 ' keep 2 blank rows between tables
    Loop
End Function

Private Sub DeleteRecipeProcessTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.count To 1 Step -1
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

Private Function GetOrAddTemplateRow(ByVal loTpl As ListObject, ByVal cScope As Long, ByVal cRecipe As Long, _
    ByVal cTargetCol As Long, ByVal cFormula As Long) As ListRow

    If loTpl Is Nothing Then Exit Function
    If loTpl.DataBodyRange Is Nothing Then
        Set GetOrAddTemplateRow = loTpl.ListRows.Add
        Exit Function
    End If

    Dim r As Long
    For r = 1 To loTpl.DataBodyRange.Rows.Count
        If IsTemplateRowEmpty(loTpl, r, cScope, cRecipe, cTargetCol, cFormula) Then
            Set GetOrAddTemplateRow = loTpl.ListRows(r)
            Exit Function
        End If
    Next r

    Set GetOrAddTemplateRow = loTpl.ListRows.Add
End Function

Private Function IsTemplateRowEmpty(ByVal loTpl As ListObject, ByVal rowIdx As Long, ByVal cScope As Long, _
    ByVal cRecipe As Long, ByVal cTargetCol As Long, ByVal cFormula As Long) As Boolean

    If loTpl Is Nothing Then Exit Function
    If loTpl.DataBodyRange Is Nothing Then Exit Function
    If rowIdx < 1 Or rowIdx > loTpl.DataBodyRange.Rows.Count Then Exit Function

    Dim rowRange As Range
    Set rowRange = loTpl.DataBodyRange.Rows(rowIdx)

    Dim scopeVal As String
    Dim recipeVal As String
    Dim targetVal As String
    Dim formulaVal As String

    If cScope > 0 Then scopeVal = NzStr(rowRange.Cells(1, cScope).Value)
    If cRecipe > 0 Then recipeVal = NzStr(rowRange.Cells(1, cRecipe).Value)
    If cTargetCol > 0 Then targetVal = NzStr(rowRange.Cells(1, cTargetCol).Value)

    If cFormula > 0 Then
        Dim fCell As Range
        Set fCell = rowRange.Cells(1, cFormula)
        If Not fCell Is Nothing Then
            If fCell.HasFormula Then
                formulaVal = CStr(fCell.FormulaR1C1)
            Else
                formulaVal = NzStr(fCell.Value)
            End If
        End If
    End If

    If scopeVal <> "" Or recipeVal <> "" Or targetVal <> "" Then Exit Function
    If formulaVal <> "" And formulaVal <> "0" Then Exit Function

    IsTemplateRowEmpty = True
End Function

' System 1: Recipe List Builder - register process formulas as templates.
Private Function RegisterRecipeTemplates(ByVal recipeId As String, ByVal processTables As Collection) As Long
    If processTables Is Nothing Then Exit Function
    If processTables.count = 0 Then Exit Function

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

    NormalizeTemplateFormulaColumn loTpl, cFormula

    If Not loTpl.DataBodyRange Is Nothing And cScope > 0 And cRecipe > 0 Then
        Dim r As Long
        For r = loTpl.DataBodyRange.rows.count To 1 Step -1
            If StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cScope).value), TEMPLATE_SCOPE_RECIPE_PROCESS, vbTextCompare) = 0 Then
                If recipeId = "" Or StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cRecipe).value), recipeId, vbTextCompare) = 0 Then
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

            Dim lr As ListRow
            Set lr = GetOrAddTemplateRow(loTpl, cScope, cRecipe, cTargetCol, cFormula)
            If cGuid > 0 Then lr.Range.Cells(1, cGuid).Value = modUR_Snapshot.GenerateGUID()
            If cScope > 0 Then lr.Range.Cells(1, cScope).Value = TEMPLATE_SCOPE_RECIPE_PROCESS
            If cRecipe > 0 Then lr.Range.Cells(1, cRecipe).Value = recipeId
            If cIngredient > 0 Then lr.Range.Cells(1, cIngredient).Value = ""
            If cProcess > 0 Then lr.Range.Cells(1, cProcess).Value = procName
            If cTargetTable > 0 Then lr.Range.Cells(1, cTargetTable).Value = ""
            If cTargetCol > 0 Then lr.Range.Cells(1, cTargetCol).Value = lc.Name
            If cFormula > 0 Then WriteTemplateFormulaCell lr.Range.Cells(1, cFormula), formulaText
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
        ProcessNameFromTable = NzStr(lo.DataBodyRange.Cells(1, cProc).value)
    End If
    If ProcessNameFromTable = "" Then ProcessNameFromTable = ExtractProcessKeyFromTableName(lo.Name)
End Function

Private Function GetColumnFormulaText(ByVal lc As ListColumn) As String
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function
    Dim cell As Range
    Set cell = lc.DataBodyRange.Cells(1, 1)
    On Error Resume Next
    If cell.HasFormula Then GetColumnFormulaText = CStr(cell.FormulaR1C1)
    On Error GoTo 0
    If Left$(GetColumnFormulaText, 1) <> "=" Then GetColumnFormulaText = ""
End Function

Private Function SaveFormulaTemplatesForRecipe(ByVal recipeId As String, ByVal wsProd As Worksheet) As Long
    If Trim$(recipeId) = "" Then Exit Function
    If wsProd Is Nothing Then Exit Function

    Dim wsTpl As Worksheet: Set wsTpl = SheetExists(SHEET_TEMPLATES)
    If wsTpl Is Nothing Then
        MsgBox "TemplatesTable sheet not found.", vbCritical
        Exit Function
    End If
    Dim loTpl As ListObject: Set loTpl = GetListObject(wsTpl, "TemplatesTable")
    If loTpl Is Nothing Then
        MsgBox "TemplatesTable not found.", vbCritical
        Exit Function
    End If

    Dim cGuid As Long, cScope As Long, cRecipe As Long, cIngredient As Long, cProcess As Long
    Dim cTargetTable As Long, cTargetCol As Long, cFormula As Long, cNotes As Long
    Dim cActive As Long, cCreated As Long, cUpdated As Long
    If Not GetTemplateColumnIndexes(loTpl, cGuid, cScope, cRecipe, cIngredient, cProcess, cTargetTable, _
        cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated) Then Exit Function
    NormalizeTemplateFormulaColumn loTpl, cFormula

    Dim totalAdded As Long
    Dim nowVal As Date: nowVal = Now

    ' Scope: Recipe process tables (builder/chooser share formulas).
    ClearTemplatesForScope loTpl, recipeId, TEMPLATE_SCOPE_RECIPE_PROCESS, cScope, cRecipe
    Dim procTables As Collection
    Set procTables = GetRecipeBuilderProcessTables(wsProd)
    If Not procTables Is Nothing Then
        Dim loProc As ListObject
        For Each loProc In procTables
            totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, loProc, recipeId, TEMPLATE_SCOPE_RECIPE_PROCESS, _
                ProcessNameFromTable(loProc), "", nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
                cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Recipe process")
        Next loProc
    End If

    ' Scope: Inventory Palette Builder tables.
    ClearTemplatesForScope loTpl, recipeId, TEMPLATE_SCOPE_PALETTE_BUILDER, cScope, cRecipe
    Dim loIng As ListObject
    Dim loItems As ListObject
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, loIng, recipeId, TEMPLATE_SCOPE_PALETTE_BUILDER, _
        "", "IP_ChooseIngredient", nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
        cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Palette builder")
    totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, loItems, recipeId, TEMPLATE_SCOPE_PALETTE_BUILDER, _
        "", "IP_ChooseItem", nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
        cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Palette builder")

    ' Scope: Production run tables.
    ClearTemplatesForScope loTpl, recipeId, TEMPLATE_SCOPE_PROD_RUN, cScope, cRecipe
    Dim lo As ListObject
    For Each lo In wsProd.ListObjects
        If LCase$(lo.Name) Like "proc_*_palette" Then
            totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, lo, recipeId, TEMPLATE_SCOPE_PROD_RUN, _
                ProcessNameFromTable(lo), TEMPLATE_TABLEKEY_PALETTE, nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
                cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Production run")
        End If
    Next lo
    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, loOut, recipeId, TEMPLATE_SCOPE_PROD_RUN, _
        "", "ProductionOutput", nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
        cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Production run")

    SaveFormulaTemplatesForRecipe = totalAdded
End Function

Private Function GetTemplateColumnIndexes(ByVal loTpl As ListObject, ByRef cGuid As Long, ByRef cScope As Long, _
    ByRef cRecipe As Long, ByRef cIngredient As Long, ByRef cProcess As Long, ByRef cTargetTable As Long, _
    ByRef cTargetCol As Long, ByRef cFormula As Long, ByRef cNotes As Long, ByRef cActive As Long, _
    ByRef cCreated As Long, ByRef cUpdated As Long) As Boolean

    If loTpl Is Nothing Then Exit Function
    cGuid = ColumnIndex(loTpl, "GUID")
    cScope = ColumnIndex(loTpl, "TEMPLATE_SCOPE")
    cRecipe = ColumnIndex(loTpl, "RECIPE_ID")
    cIngredient = ColumnIndex(loTpl, "INGREDIENT_ID")
    cProcess = ColumnIndex(loTpl, "PROCESS")
    cTargetTable = ColumnIndex(loTpl, "TARGET_TABLE")
    cTargetCol = ColumnIndex(loTpl, "TARGET_COLUMN")
    cFormula = ColumnIndex(loTpl, "FORMULA")
    cNotes = ColumnIndex(loTpl, "NOTES")
    cActive = ColumnIndex(loTpl, "ACTIVE")
    cCreated = ColumnIndex(loTpl, "CREATED_AT")
    cUpdated = ColumnIndex(loTpl, "UPDATED_AT")

    If cScope = 0 Or cRecipe = 0 Or cTargetCol = 0 Or cFormula = 0 Then
        MsgBox "TemplatesTable is missing required columns (TEMPLATE_SCOPE, RECIPE_ID, TARGET_COLUMN, FORMULA).", vbCritical
        Exit Function
    End If
    GetTemplateColumnIndexes = True
End Function

Private Sub ClearTemplatesForScope(ByVal loTpl As ListObject, ByVal recipeId As String, ByVal scopeName As String, _
    ByVal cScope As Long, ByVal cRecipe As Long)

    If loTpl Is Nothing Then Exit Sub
    If cScope = 0 Or cRecipe = 0 Then Exit Sub
    If loTpl.DataBodyRange Is Nothing Then Exit Sub

    Dim r As Long
    For r = loTpl.DataBodyRange.Rows.Count To 1 Step -1
        If StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cScope).Value), scopeName, vbTextCompare) = 0 Then
            If StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cRecipe).Value), recipeId, vbTextCompare) = 0 Then
                loTpl.ListRows(r).Delete
            End If
        End If
    Next r
End Sub

Private Function AddTemplateRowsFromTable(ByVal loTpl As ListObject, ByVal loSource As ListObject, ByVal recipeId As String, _
    ByVal scopeName As String, ByVal processName As String, ByVal targetTableName As String, ByVal nowVal As Date, _
    ByVal cGuid As Long, ByVal cScope As Long, ByVal cRecipe As Long, ByVal cIngredient As Long, ByVal cProcess As Long, _
    ByVal cTargetTable As Long, ByVal cTargetCol As Long, ByVal cFormula As Long, ByVal cNotes As Long, _
    ByVal cActive As Long, ByVal cCreated As Long, ByVal cUpdated As Long, Optional ByVal noteText As String = "") As Long

    If loTpl Is Nothing Or loSource Is Nothing Then Exit Function
    If loSource.DataBodyRange Is Nothing Then Exit Function

    Dim added As Long
    Dim lc As ListColumn
    For Each lc In loSource.ListColumns
        Dim formulaText As String
        formulaText = GetColumnFormulaText(lc)
        If formulaText = "" Then GoTo NextCol

        Dim lr As ListRow
        Set lr = GetOrAddTemplateRow(loTpl, cScope, cRecipe, cTargetCol, cFormula)
        If cGuid > 0 Then lr.Range.Cells(1, cGuid).Value = modUR_Snapshot.GenerateGUID()
        If cScope > 0 Then lr.Range.Cells(1, cScope).Value = scopeName
        If cRecipe > 0 Then lr.Range.Cells(1, cRecipe).Value = recipeId
        If cIngredient > 0 Then lr.Range.Cells(1, cIngredient).Value = ""
        If cProcess > 0 Then lr.Range.Cells(1, cProcess).Value = processName
        If cTargetTable > 0 Then lr.Range.Cells(1, cTargetTable).Value = targetTableName
        If cTargetCol > 0 Then lr.Range.Cells(1, cTargetCol).Value = lc.Name
        If cFormula > 0 Then WriteTemplateFormulaCell lr.Range.Cells(1, cFormula), formulaText
        If cNotes > 0 Then lr.Range.Cells(1, cNotes).Value = noteText
        If cActive > 0 Then lr.Range.Cells(1, cActive).Value = True
        If cCreated > 0 Then lr.Range.Cells(1, cCreated).Value = nowVal
        If cUpdated > 0 Then lr.Range.Cells(1, cUpdated).Value = nowVal
        added = added + 1
NextCol:
    Next lc

    AddTemplateRowsFromTable = added
End Function

Private Sub NormalizeTemplateFormulaColumn(ByVal loTpl As ListObject, ByVal cFormula As Long)
    If loTpl Is Nothing Then Exit Sub
    If cFormula = 0 Then Exit Sub
    Dim lc As ListColumn
    Set lc = loTpl.ListColumns(cFormula)
    If lc Is Nothing Then Exit Sub
    On Error Resume Next
    lc.Range.NumberFormat = "@"
    On Error GoTo 0
    If lc.DataBodyRange Is Nothing Then Exit Sub

    Dim cell As Range
    For Each cell In lc.DataBodyRange.Cells
        Dim formulaText As String
        If cell.HasFormula Then
            formulaText = CStr(cell.FormulaR1C1)
        Else
            formulaText = NzStr(cell.Value)
        End If
        If Left$(formulaText, 1) = "=" Then
            WriteTemplateFormulaCell cell, formulaText
        End If
    Next cell
End Sub

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
    RecipeProcessSequenceFromName = CLng(val(core))
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

Private Sub ClearListObjectFormulas(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    lo.DataBodyRange.SpecialCells(xlCellTypeFormulas).ClearContents
    On Error GoTo 0
End Sub

Private Sub WriteTemplateFormulaCell(ByVal targetCell As Range, ByVal formulaText As String)
    If targetCell Is Nothing Then Exit Sub
    If formulaText = "" Then Exit Sub
    On Error Resume Next
    targetCell.NumberFormat = "@"
    targetCell.Value = "'" & formulaText
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

Private Function NzDbl(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzDbl = 0#
    Else
        NzDbl = CDbl(v)
    End If
End Function

Private Function NzLng(v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzLng = 0
    Else
        NzLng = CLng(v)
    End If
End Function

Private Function NormalizeIdFirst(ByVal v As String) As String
    Dim tokens As Variant
    tokens = SplitTokens(v)
    If IsEmpty(tokens) Then Exit Function
    NormalizeIdFirst = CStr(tokens(LBound(tokens)))
End Function

Private Function NormalizeIdLast(ByVal v As String) As String
    Dim tokens As Variant
    tokens = SplitTokens(v)
    If IsEmpty(tokens) Then Exit Function
    NormalizeIdLast = CStr(tokens(UBound(tokens)))
End Function

Private Function SplitTokens(ByVal v As String) As Variant
    Dim s As String
    s = Trim$(v)
    If s = "" Then Exit Function
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Application.WorksheetFunction.Trim(s)
    Dim parts As Variant
    parts = Split(s, " ")
    Dim cleaned() As String
    Dim i As Long, n As Long
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then
            n = n + 1
            ReDim Preserve cleaned(0 To n - 1)
            cleaned(n - 1) = Trim$(parts(i))
        End If
    Next i
    If n = 0 Then Exit Function
    SplitTokens = cleaned
End Function
