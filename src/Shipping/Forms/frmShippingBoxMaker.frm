VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShippingBoxMaker
   Caption         =   "Shipping Box Maker"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShippingBoxMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mCboBoxes As MSForms.ComboBox
Private WithEvents mCboVersions As MSForms.ComboBox
Private WithEvents mTxtQty As MSForms.TextBox
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnMake As MSForms.CommandButton
Private WithEvents mBtnUnmake As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton

Private mTxtBoxName As MSForms.TextBox
Private mTxtUom As MSForms.TextBox
Private mTxtLocation As MSForms.TextBox
Private mTxtDescription As MSForms.TextBox
Private WithEvents mLstShippables As MSForms.ListBox
Private mLstComponents As MSForms.ListBox
Private mLblPackageInv As MSForms.Label
Private mLblStatus As MSForms.Label

Private mSavedBoxes As Variant
Private mComponents As Variant
Private mSelectedPackageRow As Long
Private mLoading As Boolean
Private mBuilt As Boolean

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Public Sub InitializeFromShipping()
    If Not mBuilt Then BuildLayout

    mLoading = True
    LoadSavedBoxes
    mLoading = False

    If mCboBoxes.ListCount > 0 Then
        mCboBoxes.ListIndex = 0
        LoadSelectedBox
    Else
        ClearBoxFields
        ShowStatus "No active saved box designs found."
    End If
End Sub

Private Sub BuildLayout()
    If mBuilt Then Exit Sub
    mBuilt = True

    Me.Caption = "Shipping Box Maker"
    Me.Width = 780
    Me.Height = 560

    AddLabel "lblTitle", "Box Maker", 12, 10, 140, 20, True
    AddLabel "lblBox", "Saved Box", 12, 42, 74, 18, False
    AddLabel "lblVersion", "Version", 432, 42, 58, 18, False
    AddLabel "lblQty", "Qty", 570, 42, 38, 18, False

    Set mCboBoxes = AddComboBox("cboBoxes", 86, 38, 320, 22)
    With mCboBoxes
        .ColumnCount = 7
        .ColumnWidths = "0 pt;150 pt;0 pt;42 pt;82 pt;130 pt;0 pt"
        .BoundColumn = 1
        .TextColumn = 2
        .Style = 2
    End With
    Set mCboVersions = AddComboBox("cboVersions", 492, 38, 64, 22)
    With mCboVersions
        .ColumnCount = 8
        .ColumnWidths = "42 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt"
        .Style = 2
    End With
    Set mTxtQty = AddTextBox("txtQty", 636, 38, 60, 22)
    mTxtQty.Value = "1"
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 706, 36, 58, 26)

    AddLabel "lblBoxName", "Box Name", 12, 76, 72, 18, False
    AddLabel "lblUom", "UOM", 302, 76, 38, 18, False
    AddLabel "lblLocation", "Location", 398, 76, 66, 18, False
    Set mTxtBoxName = AddTextBox("txtBoxName", 86, 72, 194, 22)
    Set mTxtUom = AddTextBox("txtUom", 340, 72, 42, 22)
    Set mTxtLocation = AddTextBox("txtLocation", 464, 72, 118, 22)
    AddLabel "lblDescription", "Description", 12, 108, 82, 18, False
    Set mTxtDescription = AddTextBox("txtDescription", 96, 104, 600, 22)

    LockTextBox mTxtBoxName
    LockTextBox mTxtUom
    LockTextBox mTxtLocation
    LockTextBox mTxtDescription

    AddLabel "lblShippables", "Shippable Inventory", 12, 142, 170, 18, True
    Set mLblPackageInv = AddLabel("lblPackageInv", "", 190, 142, 330, 18, False)
    AddShippableHeaders 12, 164
    Set mLstShippables = AddListBox("lstShippables", 12, 184, 752, 72)
    With mLstShippables
        .ColumnCount = 5
        .ColumnWidths = "220 pt;70 pt;44 pt;120 pt;56 pt"
    End With

    AddLabel "lblComponents", "Components To Deduct", 12, 270, 170, 18, True
    AddComponentHeaders 12, 292
    Set mLstComponents = AddListBox("lstComponents", 12, 312, 752, 142)
    With mLstComponents
        .ColumnCount = 9
        .ColumnWidths = "126 pt;42 pt;50 pt;58 pt;54 pt;38 pt;76 pt;62 pt;150 pt"
    End With

    Set mBtnMake = AddButton("btnMake", "Make Boxes", 506, 466, 86, 30)
    Set mBtnUnmake = AddButton("btnUnmake", "Unmake Boxes", 604, 466, 92, 30)
    Set mBtnClose = AddButton("btnClose", "Close", 706, 466, 58, 30)
    Set mLblStatus = AddLabel("lblStatus", "", 12, 464, 480, 44, False)
End Sub

Private Sub LoadSavedBoxes()
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long
    Dim c As Long
    Dim idx As Long

    rowsData = modTS_Shipments.BoxMakerFormLoadSavedBoxes()
    If IsEmpty(rowsData) Then
        If mCboBoxes.ListCount > 0 Then
            ShowStatus "Refresh did not return saved box designs; keeping current list."
        End If
        Exit Sub
    End If

    mSavedBoxes = rowsData
    mCboBoxes.Clear

    For r = 1 To UBound(mSavedBoxes, 1)
        mCboBoxes.AddItem NzText(mSavedBoxes(r, 1))
        idx = mCboBoxes.ListCount - 1
        For c = 2 To 7
            mCboBoxes.List(idx, c - 1) = NzText(mSavedBoxes(r, c))
        Next c
    Next r
    RenderShippableInventory
    ShowStatus "Loaded " & CStr(mCboBoxes.ListCount) & " active box design(s)."
    Exit Sub

FailSoft:
    ShowStatus "Could not load saved boxes: " & Err.Description
End Sub

Private Sub LoadSelectedBox()
    On Error GoTo FailSoft

    If mCboBoxes.ListIndex < 0 Then Exit Sub
    mSelectedPackageRow = CLng(Val(CStr(mCboBoxes.List(mCboBoxes.ListIndex, 0))))
    mTxtBoxName.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 1))
    mTxtUom.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 3))
    mTxtLocation.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 4))
    mTxtDescription.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 5))
    LoadVersionsForPackage mSelectedPackageRow
    LoadSelectedVersionComponents
    RenderPackageInventory
    SelectShippableInventoryRow
    Exit Sub

FailSoft:
    ShowStatus "Could not load selected box: " & Err.Description
End Sub

Private Sub LoadVersionsForPackage(ByVal packageRow As Long)
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long
    Dim c As Long
    Dim idx As Long

    mCboVersions.Clear
    rowsData = modTS_Shipments.BoxMakerFormLoadVersions(packageRow)
    If IsEmpty(rowsData) Then
        ShowStatus "No active versions found for selected box."
        Exit Sub
    End If

    For r = 1 To UBound(rowsData, 1)
        If UCase$(NzText(rowsData(r, 2))) <> "ACTIVE" Then GoTo NextRow
        mCboVersions.AddItem NzText(rowsData(r, 1))
        idx = mCboVersions.ListCount - 1
        For c = 2 To 8
            mCboVersions.List(idx, c - 1) = NzText(rowsData(r, c))
        Next c
NextRow:
    Next r
    If mCboVersions.ListCount > 0 Then mCboVersions.ListIndex = 0
    Exit Sub

FailSoft:
    ShowStatus "Could not load versions: " & Err.Description
End Sub

Private Sub LoadSelectedVersionComponents()
    On Error GoTo FailSoft

    If mSelectedPackageRow <= 0 Then Exit Sub
    If mCboVersions.ListIndex < 0 Then Exit Sub

    mComponents = modTS_Shipments.BoxMakerFormLoadVersionComponents(mSelectedPackageRow, SelectedVersionLabel())
    RenderComponents
    Exit Sub

FailSoft:
    ShowStatus "Could not load components: " & Err.Description
End Sub

Private Sub RenderComponents()
    On Error GoTo FailSoft

    Dim r As Long
    Dim idx As Long
    Dim qtyMade As Double
    Dim perBoxQty As Double
    Dim requiredQty As Double
    Dim currentInv As Variant

    mLstComponents.Clear
    If IsEmpty(mComponents) Then
        ShowStatus "Selected version has no components."
        Exit Sub
    End If

    qtyMade = ParseNumber(Trim$(CStr(mTxtQty.Value)))
    If qtyMade <= 0 Then qtyMade = 0

    For r = 1 To UBound(mComponents, 1)
        perBoxQty = ParseNumber(NzText(mComponents(r, 5)))
        requiredQty = perBoxQty * qtyMade
        currentInv = mComponents(r, 9)

        mLstComponents.AddItem NzText(mComponents(r, 2))
        idx = mLstComponents.ListCount - 1
        mLstComponents.List(idx, 1) = NzText(mComponents(r, 4))
        mLstComponents.List(idx, 2) = Format$(perBoxQty, "0.###")
        mLstComponents.List(idx, 3) = Format$(requiredQty, "0.###")
        If NzText(currentInv) = "" Then
            mLstComponents.List(idx, 4) = "unknown"
        Else
            mLstComponents.List(idx, 4) = NzText(currentInv)
        End If
        mLstComponents.List(idx, 5) = NzText(mComponents(r, 6))
        mLstComponents.List(idx, 6) = NzText(mComponents(r, 7))
        mLstComponents.List(idx, 7) = NzText(mComponents(r, 3))
        mLstComponents.List(idx, 8) = NzText(mComponents(r, 8))
    Next r
    ShowStatus "Loaded " & CStr(mLstComponents.ListCount) & " component row(s) for " & SelectedVersionLabel() & "."
    Exit Sub

FailSoft:
    ShowStatus "Component render failed: " & Err.Description
End Sub

Private Sub RenderPackageInventory()
    On Error GoTo FailSoft

    Dim currentInv As Variant

    If mSelectedPackageRow <= 0 Then
        mLblPackageInv.Caption = ""
        Exit Sub
    End If
    currentInv = modTS_Shipments.BoxMakerFormCurrentInventory(mSelectedPackageRow, CStr(mTxtBoxName.Value))
    If NzText(currentInv) = "" Then
        mLblPackageInv.Caption = "Current inventory for shippable box: unknown"
    Else
        mLblPackageInv.Caption = "Current inventory for shippable box: " & NzText(currentInv)
    End If
    Exit Sub

FailSoft:
    mLblPackageInv.Caption = ""
End Sub

Private Sub mCboBoxes_Change()
    If mLoading Then Exit Sub
    LoadSelectedBox
End Sub

Private Sub mCboVersions_Change()
    If mLoading Then Exit Sub
    LoadSelectedVersionComponents
End Sub

Private Sub mTxtQty_Change()
    RenderComponents
End Sub

Private Sub mBtnRefresh_Click()
    InitializeFromShipping
End Sub

Private Sub mBtnMake_Click()
    PostBoxMakerAction "MAKE"
End Sub

Private Sub mBtnUnmake_Click()
    PostBoxMakerAction "UNMAKE"
End Sub

Private Sub mBtnClose_Click()
    Me.Hide
End Sub

Private Sub mLstShippables_Click()
    On Error GoTo FailSoft

    Dim packageRow As Long
    Dim r As Long

    If mLoading Then Exit Sub
    If mLstShippables.ListIndex < 0 Then Exit Sub

    packageRow = CLng(Val(NzText(mLstShippables.List(mLstShippables.ListIndex, 4))))
    If packageRow <= 0 Then Exit Sub

    For r = 0 To mCboBoxes.ListCount - 1
        If CLng(Val(NzText(mCboBoxes.List(r, 0)))) = packageRow Then
            mCboBoxes.ListIndex = r
            Exit Sub
        End If
    Next r
    Exit Sub

FailSoft:
End Sub

Private Sub PostBoxMakerAction(ByVal actionText As String)
    On Error GoTo ErrHandler

    Dim qtyMade As Double
    Dim resultMessage As String

    qtyMade = ParseNumber(Trim$(CStr(mTxtQty.Value)))
    If qtyMade <= 0 Then
        ShowStatus "Enter a positive box quantity."
        Exit Sub
    End If
    If mSelectedPackageRow <= 0 Or mCboVersions.ListIndex < 0 Then
        ShowStatus "Select a saved box and active version."
        Exit Sub
    End If
    If IsEmpty(mComponents) Then
        ShowStatus "Selected version has no components."
        Exit Sub
    End If

    If modTS_Shipments.CommitBoxMakerFormAction(mSelectedPackageRow, _
                                                CStr(mTxtBoxName.Value), _
                                                CStr(mTxtUom.Value), _
                                                CStr(mTxtLocation.Value), _
                                                CStr(mTxtDescription.Value), _
                                                SelectedVersionLabel(), _
                                                qtyMade, _
                                                mComponents, _
                                                resultMessage, _
                                                actionText) Then
        MsgBox resultMessage, vbInformation
        LoadSelectedVersionComponents
        RenderPackageInventory
        RenderShippableInventory
    Else
        If resultMessage = "" Then resultMessage = "BoxMaker action did not complete."
        MsgBox resultMessage, vbExclamation
        ShowStatus resultMessage
    End If
    Exit Sub

ErrHandler:
    ShowStatus "BoxMaker action failed: " & Err.Description
End Sub

Private Sub RenderShippableInventory()
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long
    Dim idx As Long
    Dim currentInv As String
    Dim prevLoading As Boolean

    If mLstShippables Is Nothing Then Exit Sub
    prevLoading = mLoading
    mLoading = True
    mLstShippables.Clear
    If IsEmpty(mSavedBoxes) Then GoTo CleanExit

    rowsData = modTS_Shipments.BoxMakerFormLoadShippableInventory(mSavedBoxes)
    If IsEmpty(rowsData) Then GoTo CleanExit

    For r = 1 To UBound(rowsData, 1)
        currentInv = NzText(rowsData(r, 3))
        If currentInv = "" Then currentInv = "unknown"

        mLstShippables.AddItem NzText(rowsData(r, 2))
        idx = mLstShippables.ListCount - 1
        mLstShippables.List(idx, 1) = currentInv
        mLstShippables.List(idx, 2) = NzText(rowsData(r, 4))
        mLstShippables.List(idx, 3) = NzText(rowsData(r, 5))
        mLstShippables.List(idx, 4) = NzText(rowsData(r, 1))
        If CLng(Val(NzText(rowsData(r, 1)))) = mSelectedPackageRow Then mLstShippables.ListIndex = idx
    Next r

CleanExit:
    mLoading = prevLoading
    Exit Sub

FailSoft:
    If Not mLstShippables Is Nothing Then mLstShippables.Clear
    mLoading = prevLoading
End Sub

Private Sub SelectShippableInventoryRow()
    On Error GoTo FailSoft

    Dim r As Long

    If mLstShippables Is Nothing Then Exit Sub
    If mSelectedPackageRow <= 0 Then Exit Sub

    For r = 0 To mLstShippables.ListCount - 1
        If CLng(Val(NzText(mLstShippables.List(r, 4)))) = mSelectedPackageRow Then
            mLstShippables.ListIndex = r
            Exit Sub
        End If
    Next r
    Exit Sub

FailSoft:
End Sub

Private Function SelectedVersionLabel() As String
    If Not mCboVersions Is Nothing Then
        If mCboVersions.ListIndex >= 0 Then SelectedVersionLabel = NzText(mCboVersions.List(mCboVersions.ListIndex, 0))
    End If
    If SelectedVersionLabel = "" Then SelectedVersionLabel = "v1"
    SelectedVersionLabel = NormalizeVersionText(SelectedVersionLabel)
End Function

Private Sub ClearBoxFields()
    mSelectedPackageRow = 0
    mCboVersions.Clear
    mTxtBoxName.Value = ""
    mTxtUom.Value = ""
    mTxtLocation.Value = ""
    mTxtDescription.Value = ""
    mLstComponents.Clear
    If Not mLstShippables Is Nothing Then mLstShippables.Clear
    mLblPackageInv.Caption = ""
End Sub

Private Function AddLabel(ByVal name As String, _
                          ByVal caption As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal boldText As Boolean) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", name, True)
    With AddLabel
        .Caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = boldText
    End With
End Function

Private Sub AddComponentHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrItem", "Item", leftPos, topPos, 126
    AddHeaderLabel "hdrRow", "ROW", leftPos + 128, topPos, 40
    AddHeaderLabel "hdrPerBox", "Per Box", leftPos + 170, topPos, 48
    AddHeaderLabel "hdrRequired", "Required", leftPos + 222, topPos, 56
    AddHeaderLabel "hdrCurrent", "Current Inv", leftPos + 282, topPos, 58
    AddHeaderLabel "hdrUom", "UOM", leftPos + 342, topPos, 38
    AddHeaderLabel "hdrLocation", "Location", leftPos + 382, topPos, 74
    AddHeaderLabel "hdrCode", "Code", leftPos + 460, topPos, 60
    AddHeaderLabel "hdrDesc", "Description", leftPos + 524, topPos, 150
End Sub

Private Sub AddShippableHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrShipBox", "Box", leftPos, topPos, 220
    AddHeaderLabel "hdrShipCurrent", "Current Inv", leftPos + 222, topPos, 68
    AddHeaderLabel "hdrShipUom", "UOM", leftPos + 294, topPos, 42
    AddHeaderLabel "hdrShipLocation", "Location", leftPos + 340, topPos, 118
    AddHeaderLabel "hdrShipRow", "ROW", leftPos + 462, topPos, 56
End Sub

Private Sub AddHeaderLabel(ByVal name As String, _
                           ByVal caption As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single)
    Dim lbl As MSForms.Label

    Set lbl = AddLabel(name, caption, leftPos, topPos, widthVal, 14, True)
    lbl.Font.Size = 8
End Sub

Private Function AddTextBox(ByVal name As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", name, True)
    With AddTextBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddComboBox(ByVal name As String, _
                             ByVal leftPos As Single, _
                             ByVal topPos As Single, _
                             ByVal widthVal As Single, _
                             ByVal heightVal As Single) As MSForms.ComboBox
    Set AddComboBox = Me.Controls.Add("Forms.ComboBox.1", name, True)
    With AddComboBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddListBox(ByVal name As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", name, True)
    With AddListBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddButton(ByVal name As String, _
                           ByVal caption As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", name, True)
    With AddButton
        .Caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Sub LockTextBox(ByVal txt As MSForms.TextBox)
    If txt Is Nothing Then Exit Sub
    txt.Locked = True
    txt.BackColor = &H8000000F
End Sub

Private Sub ShowStatus(ByVal message As String)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = message
End Sub

Private Function NzText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        NzText = ""
    Else
        NzText = CStr(value)
    End If
End Function

Private Function ParseNumber(ByVal textValue As String) As Double
    On Error GoTo UseZero
    textValue = Trim$(textValue)
    If textValue = "" Then Exit Function
    ParseNumber = CDbl(textValue)
    Exit Function
UseZero:
    ParseNumber = 0#
End Function

Private Function NormalizeVersionText(ByVal valueIn As String) As String
    valueIn = Trim$(valueIn)
    If valueIn = "" Then
        NormalizeVersionText = "v1"
    ElseIf LCase$(Left$(valueIn, 1)) = "v" Then
        NormalizeVersionText = "v" & CStr(Val(Mid$(valueIn, 2)))
    ElseIf IsNumeric(valueIn) Then
        NormalizeVersionText = "v" & CStr(Val(valueIn))
    Else
        NormalizeVersionText = valueIn
    End If
End Function
