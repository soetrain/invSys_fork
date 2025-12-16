'========================
' Class: cDynItemSearch (late-bound)
' Purpose: Runtime-built Item Search form for ReceivedTally
'------------------------
Option Explicit

Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents lst As MSForms.ListBox
Private uf As Object
Private txtDesc As MSForms.TextBox
Private itemsArr As Variant
Private callbackCell As Range

Public Sub ShowForCell(ByVal targetCell As Range)
    Debug.Print "cDynItemSearch.ShowForCell entered"
    If targetCell Is Nothing Then Exit Sub
    Set callbackCell = targetCell
    Debug.Print "  Calling BuildForm"
    If Not BuildForm() Then
        Debug.Print "  BuildForm returned False"
        Exit Sub
    End If
    Debug.Print "  Calling LoadItems"
    If Not LoadItems() Then
        Debug.Print "  LoadItems returned False"
        Exit Sub
    End If
    Debug.Print "  Showing uf vbModeless"
    uf.Show vbModeless
    Debug.Print "  uf.Show returned (form should now be visible)"
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

    ' Build controls
    uf.Caption = "Item Search"
    uf.Width = 420
    uf.Height = 340

    On Error Resume Next
    Set txtSearch = uf.Controls.Add("Forms.TextBox.1", "txtSearch", True)
    If Err.Number <> 0 Then Debug.Print "    ERROR adding txtSearch:", Err.Number, Err.Description
    Err.Clear

    Set lst = uf.Controls.Add("Forms.ListBox.1", "lst", True)
    If Err.Number <> 0 Then Debug.Print "    ERROR adding lst:", Err.Number, Err.Description
    Err.Clear

    Set txtDesc = uf.Controls.Add("Forms.TextBox.1", "txtDesc", True)
    If Err.Number <> 0 Then Debug.Print "    ERROR adding txtDesc:", Err.Number, Err.Description
    Err.Clear
    On Error GoTo 0

    ' Configure listbox
    lst.ColumnCount = 7
    lst.ColumnHeads = False
    lst.BoundColumn = 0
    lst.TextColumn = 3 ' ITEM
    lst.ColumnWidths = "60;90;180;60;90;180;140" ' ROW | ITEM_CODE | ITEM | UOM | LOCATION | DESCRIPTION | VENDORS

    ' Configure description box
    txtDesc.Left = 10: txtDesc.Top = lst.Top + lst.Height + 5
    txtDesc.Width = uf.Width - 30: txtDesc.Height = 50
    txtDesc.MultiLine = True: txtDesc.Locked = True

    ' Position search box and list
    txtSearch.Left = 10: txtSearch.Top = 10: txtSearch.Width = uf.Width - 30
    lst.Left = 10: lst.Top = 35: lst.Width = uf.Width - 30: lst.Height = 230

    BuildForm = True
    Debug.Print "    BuildForm SUCCESS"
End Function

Private Function LoadItems() As Boolean
    Debug.Print "  LoadItems entered"
    itemsArr = modTS_Received.LoadItemList()
    lst.Clear
    If IsEmpty(itemsArr) Then
        Debug.Print "    LoadItems FAIL: itemsArr is Empty"
        LoadItems = False
        Exit Function
    End If
    Debug.Print "    LoadItems: UBound =", UBound(itemsArr, 1)
    Dim i As Long
    For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = itemsArr(i, 1) ' ROW
        lst.List(lst.ListCount - 1, 1) = itemsArr(i, 2) ' ITEM_CODE
        lst.List(lst.ListCount - 1, 2) = itemsArr(i, 3) ' ITEM
        lst.List(lst.ListCount - 1, 3) = itemsArr(i, 4) ' UOM
        lst.List(lst.ListCount - 1, 4) = itemsArr(i, 5) ' LOCATION
        lst.List(lst.ListCount - 1, 5) = itemsArr(i, 6) ' DESCRIPTION
        lst.List(lst.ListCount - 1, 6) = itemsArr(i, 7) ' VENDORS
    Next i
    LoadItems = True
    Debug.Print "    LoadItems SUCCESS"
End Function

' ==== events ====
Private Sub txtSearch_Change()
    Dim term As String: term = LCase$(Trim$(txtSearch.Text))
    Dim i As Long
    lst.Clear
    If IsEmpty(itemsArr) Then Exit Sub
    For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
        If term = "" Or InStr(1, LCase$(itemsArr(i, 3)), term) > 0 Then
            lst.AddItem
            lst.List(lst.ListCount - 1, 0) = itemsArr(i, 1)
            lst.List(lst.ListCount - 1, 1) = itemsArr(i, 2)
            lst.List(lst.ListCount - 1, 2) = itemsArr(i, 3)
            lst.List(lst.ListCount - 1, 3) = itemsArr(i, 4)
            lst.List(lst.ListCount - 1, 4) = itemsArr(i, 5)
            lst.List(lst.ListCount - 1, 5) = itemsArr(i, 6)
            lst.List(lst.ListCount - 1, 6) = itemsArr(i, 7)
        End If
    Next i
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
        txtDesc.Text = lst.List(lst.ListIndex, 5)
    End If
End Sub

' Manual search filter you can call from external code (since WithEvents removed)
Public Sub FilterList(ByVal term As String)
    Dim i As Long
    lst.Clear
    If IsEmpty(itemsArr) Then Exit Sub
    term = LCase$(Trim$(term))
    For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
        If term = "" Or InStr(1, LCase$(itemsArr(i, 3)), term) > 0 Then
            lst.AddItem
            lst.List(lst.ListCount - 1, 0) = itemsArr(i, 1)
            lst.List(lst.ListCount - 1, 1) = itemsArr(i, 2)
            lst.List(lst.ListCount - 1, 2) = itemsArr(i, 3)
            lst.List(lst.ListCount - 1, 3) = itemsArr(i, 4)
            lst.List(lst.ListCount - 1, 4) = itemsArr(i, 5)
            lst.List(lst.ListCount - 1, 5) = itemsArr(i, 6)
            lst.List(lst.ListCount - 1, 6) = itemsArr(i, 7)
        End If
    Next i
End Sub

Public Sub CommitSelection()
    If lst.ListIndex < 0 Then Exit Sub
    If callbackCell Is Nothing Then Exit Sub

    callbackCell.Value = lst.List(lst.ListIndex, 2) ' ITEM to staging cell

    Dim refNum As String, qty As Double
    Dim rt As ListObject
    On Error Resume Next
    Set rt = callbackCell.ListObject
    On Error GoTo 0
    If Not rt Is Nothing And Not rt.DataBodyRange Is Nothing Then
        Dim rowIdx As Long
        rowIdx = callbackCell.Row - rt.DataBodyRange.Row + 1
        If rowIdx >= 1 And rowIdx <= rt.DataBodyRange.Rows.Count Then
            refNum = NzStr(rt.DataBodyRange.Cells(rowIdx, ColumnIndex(rt, "REF_NUMBER")).Value)
            qty = NzDbl(rt.DataBodyRange.Cells(rowIdx, ColumnIndex(rt, "QUANTITY")).Value)
        End If
    End If

    modTS_Received.AddOrMergeFromSearch _
        refNum, _
        lst.List(lst.ListIndex, 2), _
        lst.List(lst.ListIndex, 1), _
        qty, _
        lst.List(lst.ListIndex, 6), _
        "", _
        lst.List(lst.ListIndex, 5), _
        lst.List(lst.ListIndex, 3), _
        lst.List(lst.ListIndex, 4), _
        CLng(NzLng(lst.List(lst.ListIndex, 0)))

    modTS_Received.RebuildAggregation
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
