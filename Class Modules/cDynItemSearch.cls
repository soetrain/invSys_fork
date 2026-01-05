'========================
' Class: cDynItemSearch (late-bound)
' Purpose: Runtime-built Item Search form
'------------------------
Option Explicit

Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents lst As MSForms.ListBox
Private WithEvents chkShippable As MSForms.CheckBox
Private uf As Object
Private txtDesc As MSForms.TextBox
Private itemsArr As Variant
Private callbackCell As Range
Private mPickerMode As String
Private mCategoryCol As Long
Private mHasCategory As Boolean
Private mIngredientRecipeId As String
Private mAllowedRows As Object

Public Sub ShowForCell(ByVal targetCell As Range)
    Debug.Print "cDynItemSearch.ShowForCell entered"
    If targetCell Is Nothing Then Exit Sub
    mPickerMode = "item"
    Set callbackCell = targetCell
    Debug.Print "  Calling BuildForm"
    If Not BuildForm() Then
        Debug.Print "  BuildForm returned False"
        Exit Sub
    End If
    ConfigureForMode
    Debug.Print "  Calling LoadItems"
    If Not LoadItems() Then
        Debug.Print "  LoadItems returned False"
        Exit Sub
    End If
    Debug.Print "  Showing uf vbModeless"
    uf.Show vbModeless
    Debug.Print "  uf.Show returned (form should now be visible)"
End Sub

Public Sub ShowForRecipeCell(ByVal targetCell As Range)
    Debug.Print "cDynItemSearch.ShowForRecipeCell entered"
    If targetCell Is Nothing Then Exit Sub
    mPickerMode = "recipe"
    Set mAllowedRows = Nothing
    Set callbackCell = targetCell
    If Not BuildForm() Then Exit Sub
    ConfigureForMode
    If Not LoadItems() Then Exit Sub
    uf.Show vbModeless
End Sub

Public Sub ShowForRecipeChooserCell(ByVal targetCell As Range)
    If targetCell Is Nothing Then Exit Sub
    mPickerMode = "recipe_chooser"
    Set mAllowedRows = Nothing
    Set callbackCell = targetCell
    If Not BuildForm() Then Exit Sub
    ConfigureForMode
    If Not LoadItems() Then Exit Sub
    uf.Show vbModal
End Sub

Public Sub ShowForPaletteRecipeCell(ByVal targetCell As Range)
    If targetCell Is Nothing Then Exit Sub
    mPickerMode = "palette_recipe"
    Set mAllowedRows = Nothing
    Set callbackCell = targetCell
    If Not BuildForm() Then Exit Sub
    ConfigureForMode
    If Not LoadItems() Then Exit Sub
    uf.Show vbModal
End Sub

Public Sub ShowForIngredientCell(ByVal targetCell As Range, ByVal recipeId As String)
    If targetCell Is Nothing Then Exit Sub
    If Trim$(recipeId) = "" Then Exit Sub
    mPickerMode = "ingredient"
    mIngredientRecipeId = recipeId
    Set mAllowedRows = Nothing
    Set callbackCell = targetCell
    If Not BuildForm() Then Exit Sub
    ConfigureForMode
    If Not LoadItems() Then Exit Sub
    uf.Show vbModal
End Sub

Public Sub ShowForPaletteItemCell(ByVal targetCell As Range, ByVal allowedRows As Object)
    If targetCell Is Nothing Then Exit Sub
    mPickerMode = "palette_item"
    Set mAllowedRows = allowedRows
    Set callbackCell = targetCell
    If Not BuildForm() Then Exit Sub
    ConfigureForMode
    If Not LoadItems() Then Exit Sub
    uf.Show vbModal
End Sub

Private Function BuildForm() As Boolean
    Debug.Print "  BuildForm entered"
    Set uf = Nothing
    On Error Resume Next
    Err.Clear

    ' Use compiled template form to avoid FM20/COM issues
    Set uf = VBA.UserForms.Add("ufDynItemSearchTemplate")
    If Err.Number <> 0 Then
        Debug.Print "    Add('ufDynItemSearchTemplate') ERROR:", Err.Number, Err.Description
    ElseIf uf Is Nothing Then
        Debug.Print "    Add('ufDynItemSearchTemplate') returned Nothing"
    Else
        Debug.Print "    Add('ufDynItemSearchTemplate') SUCCESS, type:", TypeName(uf)
    End If
    On Error GoTo 0

    If uf Is Nothing Then
        Debug.Print "    BuildForm FAIL: uf is Nothing"
        BuildForm = False
        Exit Function
    End If

    ' Build controls and layout
    uf.Caption = "Item Search"
    uf.Width = 480
    uf.Height = 420

    ' layout metrics
    Dim m As Single: m = 10
    Dim innerW As Single: innerW = uf.InsideWidth - 2 * m
    Dim innerH As Single: innerH = uf.InsideHeight - 2 * m
    Dim listH As Single: listH = innerH * 0.62 ' allow space for shippable filter and description
    Dim chkH As Single: chkH = 16

    On Error Resume Next
    Set txtSearch = uf.Controls.Add("Forms.TextBox.1", "txtSearch", True)
    If Err.Number <> 0 Then Debug.Print "    ERROR adding txtSearch:", Err.Number, Err.Description
    Err.Clear

    Set chkShippable = uf.Controls.Add("Forms.CheckBox.1", "chkShippable", True)
    If Err.Number <> 0 Then Debug.Print "    ERROR adding chkShippable:", Err.Number, Err.Description
    Err.Clear

    Set lst = uf.Controls.Add("Forms.ListBox.1", "lst", True)
    If Err.Number <> 0 Then Debug.Print "    ERROR adding lst:", Err.Number, Err.Description
    Err.Clear

    Set txtDesc = uf.Controls.Add("Forms.TextBox.1", "txtDesc", True)
    If Err.Number <> 0 Then Debug.Print "    ERROR adding txtDesc:", Err.Number, Err.Description
    Err.Clear
    On Error GoTo 0

    ' Position search box and list to use most of the space
    txtSearch.Left = m: txtSearch.Top = m: txtSearch.Width = innerW
    chkShippable.Left = m
    chkShippable.Top = txtSearch.Top + txtSearch.Height + 4
    chkShippable.Width = innerW
    chkShippable.Height = chkH
    chkShippable.Caption = "Shippable only"
    chkShippable.Value = ShouldDefaultShippable()

    lst.Left = m: lst.Top = chkShippable.Top + chkShippable.Height + 6
    lst.Width = innerW: lst.Height = listH

    ' Configure listbox
    lst.ColumnCount = 7
    lst.ColumnHeads = False
    lst.BoundColumn = 3 ' ITEM (display)
    lst.TextColumn = 3
    lst.IntegralHeight = False
    ' ROW | ITEM_CODE | ITEM | UOM | LOCATION | DESCRIPTION | VENDORS
    lst.ColumnWidths = "45;110;260;60;110;0;0"

    ' Configure description box (remaining space)
    txtDesc.Left = m
    txtDesc.Top = lst.Top + lst.Height + 6
    txtDesc.Width = innerW
    txtDesc.Height = (uf.InsideHeight - txtDesc.Top - m)
    txtDesc.MultiLine = True
    txtDesc.Locked = True
    txtDesc.ScrollBars = 2 'fmScrollBarsVertical

    BuildForm = True
    Debug.Print "    BuildForm SUCCESS"
End Function

Private Sub ConfigureForMode()
    If uf Is Nothing Then Exit Sub
    If LCase$(mPickerMode) = "ingredient" Then
        uf.Caption = "Ingredient Search"
        If Not chkShippable Is Nothing Then chkShippable.Visible = False

        lst.ColumnCount = 7
        lst.ColumnHeads = False
        lst.BoundColumn = 2 ' INGREDIENT
        lst.TextColumn = 2
        lst.IntegralHeight = False
        ' INGREDIENT_ID | INGREDIENT | UOM | PROCESS | INPUT/OUTPUT | AMOUNT | PERCENT
        lst.ColumnWidths = "90;220;60;120;90;0;0"
        Exit Sub
    End If

    If LCase$(mPickerMode) = "palette_item" Then
        uf.Caption = "Palette Item Search"
        If Not chkShippable Is Nothing Then chkShippable.Visible = False
        Exit Sub
    End If

    If LCase$(mPickerMode) <> "recipe" And LCase$(mPickerMode) <> "palette_recipe" And LCase$(mPickerMode) <> "recipe_chooser" Then Exit Sub

    uf.Caption = "Recipe Search"

    If Not chkShippable Is Nothing Then
        chkShippable.Visible = False
    End If

    lst.ColumnCount = 3
    lst.ColumnHeads = False
    lst.BoundColumn = 2 ' RECIPE
    lst.TextColumn = 2
    lst.IntegralHeight = False
    ' RECIPE_ID | RECIPE | DESCRIPTION
    lst.ColumnWidths = "120;260;0"
End Sub

Private Function LoadItems() As Boolean
    Debug.Print "  LoadItems entered"
    If LCase$(mPickerMode) = "recipe" Or LCase$(mPickerMode) = "palette_recipe" Or LCase$(mPickerMode) = "recipe_chooser" Then
        itemsArr = mProduction.LoadRecipeList()
    ElseIf LCase$(mPickerMode) = "ingredient" Then
        itemsArr = mProduction.LoadIngredientListForRecipe(mIngredientRecipeId)
    Else
        itemsArr = modTS_Received.LoadItemList(True)
    End If
    lst.Clear
    If IsEmpty(itemsArr) Then
        Debug.Print "    LoadItems FAIL: itemsArr is Empty"
        LoadItems = False
        Exit Function
    End If
    Debug.Print "    LoadItems: UBound =", UBound(itemsArr, 1)
    If LCase$(mPickerMode) <> "recipe" And LCase$(mPickerMode) <> "palette_recipe" _
        And LCase$(mPickerMode) <> "recipe_chooser" And LCase$(mPickerMode) <> "ingredient" Then
        InitCategoryState
    End If
    ApplyFilter
    LoadItems = True
    Debug.Print "    LoadItems SUCCESS"
End Function

' ==== events ====
Private Sub txtSearch_Change()
    ApplyFilter
End Sub

Private Sub chkShippable_Click()
    ApplyFilter
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommitSelection
End Sub

Private Sub lst_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CommitSelection
        KeyCode = 0
    End If
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CommitSelection
        KeyCode = 0
    End If
End Sub

Private Sub lst_Click()
    UpdateDescription
End Sub

Private Sub UpdateDescription()
    On Error Resume Next
    If txtDesc Is Nothing Then Exit Sub
    txtDesc.Text = ""
    If lst.ListIndex >= 0 Then
    If LCase$(mPickerMode) = "recipe" Or LCase$(mPickerMode) = "palette_recipe" Or LCase$(mPickerMode) = "recipe_chooser" Then
        txtDesc.Text = lst.List(lst.ListIndex, 2) ' DESCRIPTION
    ElseIf LCase$(mPickerMode) = "ingredient" Then
        txtDesc.Text = "Process: " & lst.List(lst.ListIndex, 3) & vbCrLf & _
                       "Input/Output: " & lst.List(lst.ListIndex, 4)
        Else
            txtDesc.Text = lst.List(lst.ListIndex, 5) ' DESCRIPTION
            ' Optionally append vendors for clarity
            If lst.List(lst.ListIndex, 6) <> "" Then
                txtDesc.Text = txtDesc.Text & vbCrLf & "Vendors: " & lst.List(lst.ListIndex, 6)
            End If
        End If
    End If
End Sub

' Manual search filter you can call from external code (since WithEvents removed)
Public Sub FilterList(ByVal term As String)
    ApplyFilter term
End Sub

Private Sub ApplyFilter(Optional ByVal termOverride As Variant)
    Dim term As String
    If IsMissing(termOverride) Then
        term = NzStr(txtSearch.Text)
    Else
        term = NzStr(termOverride)
    End If
    term = LCase$(Trim$(term))

    Dim i As Long
    lst.Clear
    If IsEmpty(itemsArr) Then Exit Sub
    If LCase$(mPickerMode) = "recipe" Then
        For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
            Dim recId As String
            Dim recName As String
            recId = NzStr(itemsArr(i, 1))
            recName = NzStr(itemsArr(i, 2))
            If term = "" _
               Or InStr(1, LCase$(recId), term) > 0 _
               Or InStr(1, LCase$(recName), term) > 0 Then
                lst.AddItem
                lst.List(lst.ListCount - 1, 0) = recId
                lst.List(lst.ListCount - 1, 1) = recName
                lst.List(lst.ListCount - 1, 2) = NzStr(itemsArr(i, 3))
            End If
        Next i
        If lst.ListCount > 0 Then
            lst.ListIndex = 0
            UpdateDescription
        End If
        Exit Sub
    End If

    If LCase$(mPickerMode) = "palette_recipe" Or LCase$(mPickerMode) = "recipe_chooser" Then
        For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
            Dim pRecId As String
            Dim pRecName As String
            pRecId = NzStr(itemsArr(i, 1))
            pRecName = NzStr(itemsArr(i, 2))
            If term = "" _
               Or InStr(1, LCase$(pRecId), term) > 0 _
               Or InStr(1, LCase$(pRecName), term) > 0 Then
                lst.AddItem
                lst.List(lst.ListCount - 1, 0) = pRecId
                lst.List(lst.ListCount - 1, 1) = pRecName
                lst.List(lst.ListCount - 1, 2) = NzStr(itemsArr(i, 3))
            End If
        Next i
        If lst.ListCount > 0 Then
            lst.ListIndex = 0
            UpdateDescription
        End If
        Exit Sub
    End If

    If LCase$(mPickerMode) = "ingredient" Then
        For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
            Dim ingId As String
            Dim ingName As String
            ingId = NzStr(itemsArr(i, 1))
            ingName = NzStr(itemsArr(i, 2))
            If term = "" _
               Or InStr(1, LCase$(ingId), term) > 0 _
               Or InStr(1, LCase$(ingName), term) > 0 Then
                lst.AddItem
                lst.List(lst.ListCount - 1, 0) = ingId
                lst.List(lst.ListCount - 1, 1) = ingName
                lst.List(lst.ListCount - 1, 2) = NzStr(itemsArr(i, 3)) ' UOM
                lst.List(lst.ListCount - 1, 3) = NzStr(itemsArr(i, 4)) ' PROCESS
                lst.List(lst.ListCount - 1, 4) = NzStr(itemsArr(i, 5)) ' INPUT/OUTPUT
                lst.List(lst.ListCount - 1, 5) = NzStr(itemsArr(i, 6)) ' AMOUNT
                lst.List(lst.ListCount - 1, 6) = NzStr(itemsArr(i, 7)) ' PERCENT
            End If
        Next i
        If lst.ListCount > 0 Then
            lst.ListIndex = 0
            UpdateDescription
        End If
        Exit Sub
    End If

    If LCase$(mPickerMode) = "palette_item" Then
        If mAllowedRows Is Nothing Then Exit Sub
        If mAllowedRows.Count = 0 Then Exit Sub
    End If

    Dim shippableOnly As Boolean
    If Not chkShippable Is Nothing Then
        On Error Resume Next
        shippableOnly = (chkShippable.Value = True)
        On Error GoTo 0
    End If

    For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
        If LCase$(mPickerMode) = "palette_item" Then
            Dim rowKey As String
            rowKey = NzStr(itemsArr(i, 1))
            If Not mAllowedRows.Exists(rowKey) Then GoTo NextItem
        End If
        If shippableOnly And Not IsRowShippable(i) Then GoTo NextItem
        ' Match on ITEM (3) or ITEM_CODE (2)
        If term = "" _
           Or InStr(1, LCase$(NzStr(itemsArr(i, 3))), term) > 0 _
           Or InStr(1, LCase$(NzStr(itemsArr(i, 2))), term) > 0 Then
            lst.AddItem
            lst.List(lst.ListCount - 1, 0) = itemsArr(i, 1)
            lst.List(lst.ListCount - 1, 1) = itemsArr(i, 2)
            lst.List(lst.ListCount - 1, 2) = itemsArr(i, 3)
            lst.List(lst.ListCount - 1, 3) = itemsArr(i, 4)
            lst.List(lst.ListCount - 1, 4) = itemsArr(i, 5)
            lst.List(lst.ListCount - 1, 5) = itemsArr(i, 6)
            lst.List(lst.ListCount - 1, 6) = itemsArr(i, 7)
        End If
NextItem:
    Next i
    If lst.ListCount > 0 Then
        lst.ListIndex = 0
        UpdateDescription
    End If
End Sub

Private Sub InitCategoryState()
    mCategoryCol = 0
    mHasCategory = False
    If IsEmpty(itemsArr) Then Exit Sub
    On Error Resume Next
    If UBound(itemsArr, 2) >= 8 Then mCategoryCol = 8
    On Error GoTo 0
    If mCategoryCol = 0 Then Exit Sub

    Dim i As Long
    For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
        If Trim$(NzStr(itemsArr(i, mCategoryCol))) <> "" Then
            mHasCategory = True
            Exit For
        End If
    Next i
End Sub

Private Function IsRowShippable(ByVal rowIndex As Long) As Boolean
    If Not mHasCategory Or mCategoryCol = 0 Then
        IsRowShippable = True
        Exit Function
    End If
    IsRowShippable = (LCase$(Trim$(NzStr(itemsArr(rowIndex, mCategoryCol)))) = "shippable")
End Function

Private Function ShouldDefaultShippable() As Boolean
    Dim lo As ListObject
    If LCase$(mPickerMode) = "recipe" Or LCase$(mPickerMode) = "palette_recipe" _
        Or LCase$(mPickerMode) = "recipe_chooser" Or LCase$(mPickerMode) = "ingredient" _
        Or LCase$(mPickerMode) = "palette_item" Then Exit Function
    On Error Resume Next
    If Not callbackCell Is Nothing Then Set lo = callbackCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    ShouldDefaultShippable = (LCase$(lo.Name) = "shipmentstally")
End Function

Public Sub CommitSelection()
    If lst.ListIndex < 0 Then Exit Sub
    If callbackCell Is Nothing Then Exit Sub

    If LCase$(mPickerMode) = "recipe" Then
        Dim recId As String
        Dim recName As String
        Dim recDesc As String
        recId = NzStr(lst.List(lst.ListIndex, 0))
        recName = NzStr(lst.List(lst.ListIndex, 1))
        recDesc = NzStr(lst.List(lst.ListIndex, 2))

        Dim loRecipe As ListObject
        On Error Resume Next
        Set loRecipe = callbackCell.ListObject
        On Error GoTo 0
        If Not loRecipe Is Nothing Then
            Dim cRecNameCol As Long: cRecNameCol = ColumnIndex(loRecipe, "RECIPE_NAME")
            Dim cRecIdCol As Long: cRecIdCol = ColumnIndex(loRecipe, "RECIPE_ID")
            Dim cRecDescCol As Long: cRecDescCol = ColumnIndex(loRecipe, "DESCRIPTION")
            If cRecNameCol > 0 Then
                Dim recNameCell As Range
                Set recNameCell = EnsureListObjectCell(loRecipe, cRecNameCol)
                If Not recNameCell Is Nothing Then recNameCell.Value = recName
            End If
            If cRecIdCol > 0 Then
                Dim recIdCell As Range
                Set recIdCell = EnsureListObjectCell(loRecipe, cRecIdCol)
                If Not recIdCell Is Nothing Then recIdCell.Value = recId
            End If
            If cRecDescCol > 0 Then
                Dim recDescCell As Range
                Set recDescCell = EnsureListObjectCell(loRecipe, cRecDescCol)
                If Not recDescCell Is Nothing Then recDescCell.Value = recDesc
            End If
        Else
            callbackCell.Value = recName
        End If

        If recId <> "" Then
            mProduction.LoadRecipeFromRecipes recId
        End If
        uf.Hide
        Exit Sub
    End If

    If LCase$(mPickerMode) = "palette_recipe" Then
        Dim pRecId As String
        Dim pRecName As String
        Dim pRecDesc As String
        pRecId = NzStr(lst.List(lst.ListIndex, 0))
        pRecName = NzStr(lst.List(lst.ListIndex, 1))
        pRecDesc = NzStr(lst.List(lst.ListIndex, 2))

        Dim loPalRecipe As ListObject
        On Error Resume Next
        Set loPalRecipe = callbackCell.ListObject
        On Error GoTo 0
        If loPalRecipe Is Nothing Then
            MsgBox "IP_ChooseRecipe table not found for selection.", vbExclamation
            Exit Sub
        End If

        Dim palRecipeRow As Range
        Set palRecipeRow = ResolveListObjectRowRange(loPalRecipe, callbackCell)
        If palRecipeRow Is Nothing Then
            MsgBox "Unable to add a row to IP_ChooseRecipe.", vbExclamation
            Exit Sub
        End If

        Dim cPalNameCol As Long: cPalNameCol = ColumnIndex(loPalRecipe, "RECIPE_NAME")
        Dim cPalIdCol As Long: cPalIdCol = ColumnIndex(loPalRecipe, "RECIPE_ID")
        Dim cPalDescCol As Long: cPalDescCol = ColumnIndex(loPalRecipe, "DESCRIPTION")
        If cPalNameCol > 0 Then palRecipeRow.Cells(1, cPalNameCol).Value = pRecName
        If cPalIdCol > 0 Then palRecipeRow.Cells(1, cPalIdCol).Value = pRecId
        If cPalDescCol > 0 Then palRecipeRow.Cells(1, cPalDescCol).Value = pRecDesc

        mProduction.HandlePaletteRecipeSelected pRecId
        uf.Hide
        Exit Sub
    End If

    If LCase$(mPickerMode) = "recipe_chooser" Then
        Dim chRecId As String
        Dim chRecName As String
        Dim chRecDesc As String
        chRecId = NzStr(lst.List(lst.ListIndex, 0))
        chRecName = NzStr(lst.List(lst.ListIndex, 1))
        chRecDesc = NzStr(lst.List(lst.ListIndex, 2))

        Dim loChooser As ListObject
        On Error Resume Next
        Set loChooser = callbackCell.ListObject
        On Error GoTo 0
        If loChooser Is Nothing Then
            MsgBox "RC_RecipeChoose table not found for selection.", vbExclamation
            Exit Sub
        End If

        Dim chooserRow As Range
        Set chooserRow = ResolveListObjectRowRange(loChooser, callbackCell)
        If chooserRow Is Nothing Then
            MsgBox "Unable to add a row to RC_RecipeChoose.", vbExclamation
            Exit Sub
        End If

        Dim cChooserName As Long: cChooserName = ColumnIndex(loChooser, "RECIPE")
        If cChooserName = 0 Then cChooserName = ColumnIndex(loChooser, "RECIPE_NAME")
        Dim cChooserIdCol As Long: cChooserIdCol = ColumnIndex(loChooser, "RECIPE_ID")
        Dim cChooserDescCol As Long: cChooserDescCol = ColumnIndex(loChooser, "DESCRIPTION")
        If cChooserName > 0 Then chooserRow.Cells(1, cChooserName).Value = chRecName
        If cChooserIdCol > 0 Then chooserRow.Cells(1, cChooserIdCol).Value = chRecId
        If cChooserDescCol > 0 Then chooserRow.Cells(1, cChooserDescCol).Value = chRecDesc

        If chRecId <> "" Then
            mProduction.LoadRecipeChooser chRecId
        End If
        uf.Hide
        Exit Sub
    End If

    If LCase$(mPickerMode) = "ingredient" Then
        Dim selIngId As String
        Dim selIngName As String
        Dim selUom As String
        Dim selProc As String
        Dim selIO As String
        Dim selAmt As String
        selIngId = NzStr(lst.List(lst.ListIndex, 0))
        selIngName = NzStr(lst.List(lst.ListIndex, 1))
        selUom = NzStr(lst.List(lst.ListIndex, 2))
        selProc = NzStr(lst.List(lst.ListIndex, 3))
        selIO = NzStr(lst.List(lst.ListIndex, 4))
        selAmt = NzStr(lst.List(lst.ListIndex, 5))

        Dim loIng As ListObject
        On Error Resume Next
        Set loIng = callbackCell.ListObject
        On Error GoTo 0
        If loIng Is Nothing Then
            MsgBox "IP_ChooseIngredient table not found for selection.", vbExclamation
            Exit Sub
        End If

        Dim ingRowRange As Range
        Set ingRowRange = ResolveListObjectRowRange(loIng, callbackCell)
        If ingRowRange Is Nothing Then
            MsgBox "Unable to add a row to IP_ChooseIngredient.", vbExclamation
            Exit Sub
        End If

        Dim cIngNameCol As Long: cIngNameCol = ColumnIndex(loIng, "INGREDIENT")
        Dim cIngUomCol As Long: cIngUomCol = ColumnIndex(loIng, "UOM")
        Dim cIngQtyCol As Long: cIngQtyCol = ColumnIndex(loIng, "QUANTITY")
        Dim cIngDescCol As Long: cIngDescCol = ColumnIndex(loIng, "DESCRIPTION")
        Dim cIngGuidCol As Long: cIngGuidCol = ColumnIndex(loIng, "GUID")
        Dim cIngRecCol As Long: cIngRecCol = ColumnIndex(loIng, "RECIPE_ID")
        Dim cIngIdCol As Long: cIngIdCol = ColumnIndex(loIng, "INGREDIENT_ID")
        Dim cIngProcCol As Long: cIngProcCol = ColumnIndex(loIng, "PROCESS")

        If cIngNameCol > 0 Then ingRowRange.Cells(1, cIngNameCol).Value = selIngName
        If cIngUomCol > 0 Then ingRowRange.Cells(1, cIngUomCol).Value = selUom
        If cIngQtyCol > 0 Then ingRowRange.Cells(1, cIngQtyCol).Value = selAmt
        If cIngDescCol > 0 Then ingRowRange.Cells(1, cIngDescCol).Value = selIO
        If cIngGuidCol > 0 Then
            Dim gVal As String
            gVal = NzStr(ingRowRange.Cells(1, cIngGuidCol).Value)
            If Trim$(gVal) = "" Then ingRowRange.Cells(1, cIngGuidCol).Value = modUR_Snapshot.GenerateGUID()
        End If
        If cIngRecCol > 0 Then ingRowRange.Cells(1, cIngRecCol).Value = mIngredientRecipeId
        If cIngIdCol > 0 Then ingRowRange.Cells(1, cIngIdCol).Value = selIngId
        If cIngProcCol > 0 Then ingRowRange.Cells(1, cIngProcCol).Value = selProc

        mProduction.HandlePaletteIngredientSelected mIngredientRecipeId, selIngId
        uf.Hide
        Exit Sub
    End If

    callbackCell.Value = lst.List(lst.ListIndex, 2) ' ITEM text in active cell

    Dim lo As ListObject
    Dim rowIdx As Long
    Dim tblName As String
    On Error Resume Next
    Set lo = callbackCell.ListObject
    On Error GoTo 0
    If Not lo Is Nothing Then
        tblName = LCase$(lo.Name)
        If Not lo.DataBodyRange Is Nothing Then
            rowIdx = callbackCell.Row - lo.DataBodyRange.Row + 1
            If rowIdx < 1 Or rowIdx > lo.ListRows.Count Then rowIdx = 0
        End If
    End If

    Dim itemName As String, itemCode As String, itemRow As Long
    Dim uom As String, location As String, vendor As String
    itemName = lst.List(lst.ListIndex, 2)
    itemCode = lst.List(lst.ListIndex, 1)
    itemRow = CLng(NzLng(lst.List(lst.ListIndex, 0)))
    uom = lst.List(lst.ListIndex, 3)
    location = lst.List(lst.ListIndex, 4)
    vendor = lst.List(lst.ListIndex, 6)

    Select Case tblName
        Case "receivedtally"
            Dim refNum As String, qty As Double
            If rowIdx > 0 Then
                refNum = NzStr(lo.DataBodyRange.Cells(rowIdx, ColumnIndex(lo, "REF_NUMBER")).Value)
                qty = NzDbl(lo.DataBodyRange.Cells(rowIdx, ColumnIndex(lo, "QUANTITY")).Value)
            End If
            modTS_Received.AddOrMergeFromSearch _
                refNum, _
                itemName, _
                itemCode, _
                qty, _
                vendor, _
                "", _
                lst.List(lst.ListIndex, 5), _
                uom, _
                location, _
                itemRow, _
                stagingRow:=rowIdx
        Case "shipmentstally"
            modTS_Shipments.ApplyItemSelection callbackCell, lo, rowIdx, itemName, itemCode, itemRow, uom, location, vendor, lst.List(lst.ListIndex, 5)
        Case "boxbom"
            modTS_Shipments.ApplyItemToBoxBOM callbackCell, itemName, itemRow, uom, location, lst.List(lst.ListIndex, 5)
        Case "ip_chooseitem"
            Dim colItem As Long: colItem = ColumnIndex(lo, "ITEMS")
            If colItem = 0 Then colItem = ColumnIndex(lo, "ITEM")
            Dim colUom As Long: colUom = ColumnIndex(lo, "UOM")
            Dim colDesc As Long: colDesc = ColumnIndex(lo, "DESCRIPTION")
            Dim colRow As Long: colRow = ColumnIndex(lo, "ROW")
            Dim colRec As Long: colRec = ColumnIndex(lo, "RECIPE_ID")
            Dim colIng As Long: colIng = ColumnIndex(lo, "INGREDIENT_ID")

            If rowIdx = 0 Then rowIdx = EnsureListObjectRow(lo)
            If rowIdx > 0 Then
                If colItem > 0 Then lo.DataBodyRange.Cells(rowIdx, colItem).Value = itemName
                If colUom > 0 Then lo.DataBodyRange.Cells(rowIdx, colUom).Value = uom
                If colDesc > 0 Then lo.DataBodyRange.Cells(rowIdx, colDesc).Value = lst.List(lst.ListIndex, 5)
                If colRow > 0 Then lo.DataBodyRange.Cells(rowIdx, colRow).Value = itemRow
                If colRec > 0 Then lo.DataBodyRange.Cells(rowIdx, colRec).Value = mProduction.GetPaletteRecipeId()
                If colIng > 0 Then lo.DataBodyRange.Cells(rowIdx, colIng).Value = mProduction.GetPaletteIngredientId()
            End If
        Case Else
            If tblName = "inventorypalette_generated" Or tblName Like "proc_*_palette" Then
                Dim palItemRow As Long
                palItemRow = rowIdx
                If palItemRow = 0 Then palItemRow = EnsureListObjectRow(lo)

                Dim cPalItemCode As Long: cPalItemCode = ColumnIndex(lo, "ITEM_CODE")
                Dim cPalVend As Long: cPalVend = ColumnIndex(lo, "VENDORS")
                Dim cPalVendCode As Long: cPalVendCode = ColumnIndex(lo, "VENDOR_CODE")
                Dim cPalDescCol2 As Long: cPalDescCol2 = ColumnIndex(lo, "DESCRIPTION")
                Dim cPalItem As Long: cPalItem = ColumnIndex(lo, "ITEM")
                Dim cPalUomCol As Long: cPalUomCol = ColumnIndex(lo, "UOM")
                Dim cPalQtyCol As Long: cPalQtyCol = ColumnIndex(lo, "QUANTITY")
                Dim cPalProcCol As Long: cPalProcCol = ColumnIndex(lo, "PROCESS")
                Dim cPalLocCol As Long: cPalLocCol = ColumnIndex(lo, "LOCATION")
                Dim cPalRowCol As Long: cPalRowCol = ColumnIndex(lo, "ROW")
                Dim cPalIOCol As Long: cPalIOCol = ColumnIndex(lo, "INPUT/OUTPUT")

                If palItemRow > 0 Then
                    If cPalItemCode > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalItemCode).Value = itemCode
                    If cPalVend > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalVend).Value = vendor
                    If cPalVendCode > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalVendCode).Value = ""
                    If cPalDescCol2 > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalDescCol2).Value = lst.List(lst.ListIndex, 5)
                    If cPalItem > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalItem).Value = itemName
                    If cPalUomCol > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalUomCol).Value = uom
                    If cPalLocCol > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalLocCol).Value = location
                    If cPalRowCol > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalRowCol).Value = itemRow

                    Dim ctxRec As String, ctxIng As String, ctxProc As String, ctxIO As String
                    Dim ctxAmt As Variant
                    If mProduction.GetPaletteTableContext(lo, ctxRec, ctxIng, ctxAmt, ctxProc, ctxIO) Then
                        If cPalProcCol > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalProcCol).Value = ctxProc
                        If cPalIOCol > 0 Then lo.DataBodyRange.Cells(palItemRow, cPalIOCol).Value = ctxIO
                        If cPalQtyCol > 0 Then
                            If NzStr(lo.DataBodyRange.Cells(palItemRow, cPalQtyCol).Value) = "" Then
                                lo.DataBodyRange.Cells(palItemRow, cPalQtyCol).Value = ctxAmt
                            End If
                        End If
                    End If
                End If
            End If
    End Select

    uf.Hide
End Sub

Private Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function
Private Function NzDbl(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then NzDbl = 0# Else NzDbl = CDbl(v)
End Function
Private Function NzLng(v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then NzLng = 0 Else NzLng = CLng(v)
End Function
Private Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then ColumnIndex = lc.Index: Exit Function
    Next
    ColumnIndex = 0
End Function

Private Function EnsureListObjectCell(ByVal lo As ListObject, ByVal colIndex As Long) As Range
    If lo Is Nothing Then Exit Function
    If colIndex < 1 Then Exit Function
    On Error Resume Next
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
    On Error GoTo 0
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set EnsureListObjectCell = lo.DataBodyRange.Cells(1, colIndex)
End Function

Private Function EnsureListObjectRow(ByVal lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    On Error Resume Next
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add AlwaysInsert:=True
    On Error GoTo 0
    If lo.DataBodyRange Is Nothing Then Exit Function
    EnsureListObjectRow = 1
End Function

Private Function GetRowIndexForCell(ByVal lo As ListObject, ByVal targetCell As Range) As Long
    If lo Is Nothing Then Exit Function
    If targetCell Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim idx As Long
    idx = targetCell.Row - lo.DataBodyRange.Row + 1
    If idx < 1 Or idx > lo.ListRows.Count Then idx = 0
    GetRowIndexForCell = idx
End Function

Private Function ResolveListObjectRowRange(ByVal lo As ListObject, ByVal targetCell As Range) As Range
    If lo Is Nothing Then Exit Function
    If Not lo.DataBodyRange Is Nothing Then
        If Not targetCell Is Nothing Then
            Dim idx As Long
            idx = targetCell.Row - lo.DataBodyRange.Row + 1
            If idx >= 1 And idx <= lo.ListRows.Count Then
                Set ResolveListObjectRowRange = lo.DataBodyRange.Rows(idx)
                Exit Function
            End If
        End If
    End If
    If lo.HeaderRowRange Is Nothing Then Exit Function
    Set ResolveListObjectRowRange = lo.HeaderRowRange.Offset(1, 0).Resize(1, lo.ListColumns.Count)
End Function
