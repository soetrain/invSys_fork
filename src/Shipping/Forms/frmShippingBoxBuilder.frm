VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShippingBoxBuilder
   Caption         =   "Shipping Box Builder"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShippingBoxBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtBoxName As MSForms.TextBox
Private WithEvents mTxtUom As MSForms.TextBox
Private WithEvents mTxtLocation As MSForms.TextBox
Private WithEvents mTxtDescription As MSForms.TextBox
Private WithEvents mTxtVersion As MSForms.TextBox
Private WithEvents mTxtSearch As MSForms.TextBox
Private WithEvents mTxtQty As MSForms.TextBox
Private WithEvents mLstInventory As MSForms.ListBox
Private WithEvents mLstBom As MSForms.ListBox
Private WithEvents mBtnAdd As MSForms.CommandButton
Private WithEvents mBtnRemove As MSForms.CommandButton
Private WithEvents mBtnSave As MSForms.CommandButton
Private WithEvents mBtnCancel As MSForms.CommandButton

Private mLblStatus As MSForms.Label
Private mInventoryData As Variant
Private mBuilt As Boolean

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Public Sub InitializeFromShipping()
    If Not mBuilt Then BuildLayout
    LoadCurrentBoxState
    LoadInventoryCache
    RenderInventoryList
End Sub

Private Sub BuildLayout()
    If mBuilt Then Exit Sub
    mBuilt = True

    Me.Caption = "Shipping Box Builder"
    Me.Width = 760
    Me.Height = 520

    AddLabel "lblTitle", "Box Builder", 12, 10, 180, 20, True
    AddLabel "lblBoxName", "Box Name", 12, 42, 70, 18, False
    AddLabel "lblUom", "UOM", 292, 42, 40, 18, False
    AddLabel "lblLocation", "Location", 388, 42, 70, 18, False
    AddLabel "lblVersion", "Version", 582, 42, 60, 18, False
    AddLabel "lblDesc", "Description", 12, 76, 86, 18, False
    AddLabel "lblSearch", "Managed Inventory", 12, 122, 140, 18, True
    AddLabel "lblQty", "Qty", 306, 122, 30, 18, False
    AddLabel "lblBom", "BOM Components", 390, 122, 160, 18, True

    Set mTxtBoxName = AddTextBox("txtBoxName", 86, 38, 190, 22)
    Set mTxtUom = AddTextBox("txtUom", 328, 38, 44, 22)
    Set mTxtLocation = AddTextBox("txtLocation", 452, 38, 110, 22)
    Set mTxtVersion = AddTextBox("txtVersion", 642, 38, 44, 22)
    Set mTxtDescription = AddTextBox("txtDescription", 100, 72, 586, 22)
    Set mTxtSearch = AddTextBox("txtSearch", 12, 144, 280, 22)
    Set mTxtQty = AddTextBox("txtQty", 336, 144, 38, 22)

    Set mLstInventory = AddListBox("lstInventory", 12, 174, 362, 250)
    With mLstInventory
        .ColumnCount = 6
        .ColumnWidths = "38 pt;70 pt;120 pt;38 pt;62 pt;120 pt"
    End With

    Set mLstBom = AddListBox("lstBom", 390, 144, 342, 280)
    With mLstBom
        .ColumnCount = 8
        .ColumnWidths = "28 pt;112 pt;70 pt;38 pt;38 pt;36 pt;62 pt;120 pt"
    End With

    Set mBtnAdd = AddButton("btnAdd", "Add", 300, 430, 74, 28)
    Set mBtnRemove = AddButton("btnRemove", "Remove", 390, 430, 74, 28)
    Set mBtnSave = AddButton("btnSave", "Save Box", 566, 430, 78, 28)
    Set mBtnCancel = AddButton("btnCancel", "Cancel", 654, 430, 78, 28)
    Set mLblStatus = AddLabel("lblStatus", "", 12, 466, 720, 28, False)

    mTxtVersion.Value = "v1"
    mTxtQty.Value = "1"
End Sub

Private Sub LoadCurrentBoxState()
    On Error GoTo FailSoft

    Dim meta As Variant
    Dim rowsData As Variant
    Dim r As Long

    meta = modTS_Shipments.BoxBuilderFormCurrentMeta()
    If Not IsEmpty(meta) Then
        mTxtBoxName.Value = NzText(meta(1))
        mTxtUom.Value = NzText(meta(2))
        mTxtLocation.Value = NzText(meta(3))
        mTxtDescription.Value = NzText(meta(4))
    End If
    If Trim$(CStr(mTxtUom.Value)) = "" Then mTxtUom.Value = "ea"
    If Trim$(CStr(mTxtVersion.Value)) = "" Then mTxtVersion.Value = "v1"

    mLstBom.Clear
    rowsData = modTS_Shipments.BoxBuilderFormCurrentComponents()
    If IsEmpty(rowsData) Then Exit Sub
    For r = 1 To UBound(rowsData, 1)
        AddBomListRow NzText(rowsData(r, 1)), _
                      NzText(rowsData(r, 2)), _
                      NzText(rowsData(r, 3)), _
                      NzText(rowsData(r, 4)), _
                      NzText(rowsData(r, 5)), _
                      NzText(rowsData(r, 6)), _
                      NzText(rowsData(r, 7)), _
                      NzText(rowsData(r, 8))
    Next r
    Exit Sub

FailSoft:
    ShowStatus "Could not load current BoxBuilder rows: " & Err.Description
End Sub

Private Sub LoadInventoryCache()
    On Error GoTo FailSoft
    mInventoryData = modTS_Shipments.LoadShippingComponentPickerItems()
    If IsEmpty(mInventoryData) Then
        ShowStatus "No managed inventory rows are available. Refresh Inventory or run Setup UI."
    Else
        ShowStatus "Loaded " & CStr(UBound(mInventoryData, 1)) & " managed inventory item(s)."
    End If
    Exit Sub

FailSoft:
    mInventoryData = Empty
    ShowStatus "Inventory load failed: " & Err.Description
End Sub

Private Sub RenderInventoryList()
    On Error GoTo FailSoft

    Dim filterText As String
    Dim haystack As String
    Dim r As Long
    Dim idx As Long

    mLstInventory.Clear
    If IsEmpty(mInventoryData) Then Exit Sub

    filterText = LCase$(Trim$(CStr(mTxtSearch.Value)))
    For r = 1 To UBound(mInventoryData, 1)
        haystack = LCase$(NzText(mInventoryData(r, 1)) & " " & _
                          NzText(mInventoryData(r, 2)) & " " & _
                          NzText(mInventoryData(r, 3)) & " " & _
                          NzText(mInventoryData(r, 6)))
        If filterText = "" Or InStr(1, haystack, filterText, vbTextCompare) > 0 Then
            mLstInventory.AddItem NzText(mInventoryData(r, 1))
            idx = mLstInventory.ListCount - 1
            mLstInventory.List(idx, 1) = NzText(mInventoryData(r, 2))
            mLstInventory.List(idx, 2) = NzText(mInventoryData(r, 3))
            mLstInventory.List(idx, 3) = NzText(mInventoryData(r, 4))
            mLstInventory.List(idx, 4) = NzText(mInventoryData(r, 5))
            mLstInventory.List(idx, 5) = NzText(mInventoryData(r, 6))
        End If
    Next r
    Exit Sub

FailSoft:
    ShowStatus "Inventory filter failed: " & Err.Description
End Sub

Private Sub mTxtSearch_Change()
    RenderInventoryList
End Sub

Private Sub mBtnAdd_Click()
    If mLstInventory.ListIndex < 0 Then
        ShowStatus "Select a managed inventory item to add."
        Exit Sub
    End If

    Dim qtyText As String
    qtyText = Trim$(CStr(mTxtQty.Value))
    If ParseNumber(qtyText) <= 0 Then
        ShowStatus "Enter a positive component quantity."
        Exit Sub
    End If

    AddBomListRow NormalizeVersionText(CStr(mTxtVersion.Value)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 2)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 1)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 0)), _
                  qtyText, _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 3)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 4)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 5))
    ShowStatus "Component added."
End Sub

Private Sub mBtnRemove_Click()
    If mLstBom.ListIndex < 0 Then
        ShowStatus "Select a BOM component to remove."
        Exit Sub
    End If
    mLstBom.RemoveItem mLstBom.ListIndex
    ShowStatus "Component removed."
End Sub

Private Sub mBtnSave_Click()
    On Error GoTo ErrHandler

    Dim bomRows As Variant
    Dim i As Long

    If Trim$(CStr(mTxtBoxName.Value)) = "" Then
        ShowStatus "Enter a Box Name."
        Exit Sub
    End If
    If Trim$(CStr(mTxtUom.Value)) = "" Then
        ShowStatus "Enter a UOM."
        Exit Sub
    End If
    If mLstBom.ListCount = 0 Then
        ShowStatus "Add at least one BOM component."
        Exit Sub
    End If

    ReDim bomRows(1 To mLstBom.ListCount, 1 To 8)
    For i = 0 To mLstBom.ListCount - 1
        bomRows(i + 1, 1) = CStr(mLstBom.List(i, 0))
        bomRows(i + 1, 2) = CStr(mLstBom.List(i, 1))
        bomRows(i + 1, 3) = CStr(mLstBom.List(i, 2))
        bomRows(i + 1, 4) = CStr(mLstBom.List(i, 3))
        bomRows(i + 1, 5) = CStr(mLstBom.List(i, 4))
        bomRows(i + 1, 6) = CStr(mLstBom.List(i, 5))
        bomRows(i + 1, 7) = CStr(mLstBom.List(i, 6))
        bomRows(i + 1, 8) = CStr(mLstBom.List(i, 7))
    Next i

    modTS_Shipments.CommitBoxBuilderFormState CStr(mTxtBoxName.Value), _
                                             CStr(mTxtUom.Value), _
                                             CStr(mTxtLocation.Value), _
                                             CStr(mTxtDescription.Value), _
                                             bomRows
    Me.Hide
    Exit Sub

ErrHandler:
    ShowStatus "Save failed: " & Err.Description
End Sub

Private Sub mBtnCancel_Click()
    Me.Hide
End Sub

Private Sub AddBomListRow(ByVal versionText As String, _
                          ByVal itemName As String, _
                          ByVal itemCode As String, _
                          ByVal rowText As String, _
                          ByVal qtyText As String, _
                          ByVal uomText As String, _
                          ByVal locationText As String, _
                          ByVal descriptionText As String)
    Dim idx As Long

    versionText = NormalizeVersionText(versionText)
    mLstBom.AddItem versionText
    idx = mLstBom.ListCount - 1
    mLstBom.List(idx, 1) = itemName
    mLstBom.List(idx, 2) = itemCode
    mLstBom.List(idx, 3) = rowText
    mLstBom.List(idx, 4) = qtyText
    mLstBom.List(idx, 5) = uomText
    mLstBom.List(idx, 6) = locationText
    mLstBom.List(idx, 7) = descriptionText
End Sub

Private Function AddLabel(ByVal controlName As String, _
                          ByVal captionText As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal boldText As Boolean) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    With AddLabel
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .WordWrap = True
        .Font.Bold = boldText
    End With
End Function

Private Function AddTextBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    With AddTextBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddListBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", controlName, True)
    With AddListBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddButton(ByVal controlName As String, _
                           ByVal captionText As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    With AddButton
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Sub ShowStatus(ByVal messageText As String)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = messageText
End Sub

Private Function NzText(ByVal value As Variant) As String
    On Error GoTo UseBlank
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        NzText = ""
    Else
        NzText = Trim$(CStr(value))
    End If
    Exit Function
UseBlank:
    NzText = ""
End Function

Private Function ParseNumber(ByVal value As String) As Double
    On Error GoTo UseZero
    value = Trim$(value)
    If value = "" Then Exit Function
    ParseNumber = CDbl(value)
    Exit Function
UseZero:
    ParseNumber = 0#
End Function

Private Function NormalizeVersionText(ByVal versionText As String) As String
    versionText = LCase$(Trim$(versionText))
    If versionText = "" Then
        NormalizeVersionText = "v1"
    ElseIf Left$(versionText, 1) = "v" Then
        NormalizeVersionText = versionText
    Else
        NormalizeVersionText = "v" & versionText
    End If
End Function
