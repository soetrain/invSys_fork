'========================
' Class: cDynItemSearch (late-bound)
' Purpose: Runtime-built Item Search form for ReceivedTally
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
    Set callbackCell = targetCell
    If Not BuildForm() Then Exit Sub
    ConfigureForMode
    If Not LoadItems() Then Exit Sub
    uf.Show vbModeless
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
    If LCase$(mPickerMode) <> "recipe" Then Exit Sub

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
    If LCase$(mPickerMode) = "recipe" Then
        itemsArr = mProduction.LoadRecipeList()
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
    If LCase$(mPickerMode) <> "recipe" Then InitCategoryState
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
        If LCase$(mPickerMode) = "recipe" Then
            txtDesc.Text = lst.List(lst.ListIndex, 2) ' DESCRIPTION
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

    Dim shippableOnly As Boolean
    If Not chkShippable Is Nothing Then
        On Error Resume Next
        shippableOnly = (chkShippable.Value = True)
        On Error GoTo 0
    End If

    For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
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
    If LCase$(mPickerMode) = "recipe" Then Exit Function
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
            Dim cName As Long: cName = ColumnIndex(loRecipe, "RECIPE_NAME")
            Dim cId As Long: cId = ColumnIndex(loRecipe, "RECIPE_ID")
            Dim cDesc As Long: cDesc = ColumnIndex(loRecipe, "DESCRIPTION")
            If cName > 0 Then
                Dim nameCell As Range
                Set nameCell = EnsureListObjectCell(loRecipe, cName)
                If Not nameCell Is Nothing Then nameCell.Value = recName
            End If
            If cId > 0 Then
                Dim idCell As Range
                Set idCell = EnsureListObjectCell(loRecipe, cId)
                If Not idCell Is Nothing Then idCell.Value = recId
            End If
            If cDesc > 0 Then
                Dim descCell As Range
                Set descCell = EnsureListObjectCell(loRecipe, cDesc)
                If Not descCell Is Nothing Then descCell.Value = recDesc
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
        Case Else
            ' no special handling
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
