VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSeedInventory
   Caption         =   "invSys Admin - Seed Inventory"
   ClientHeight    =   2350
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSeedInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mCmbWarehouse As MSForms.ComboBox
Attribute mCmbWarehouse.VB_VarHelpID = -1
Private WithEvents mBtnOK As MSForms.CommandButton
Attribute mBtnOK.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1
Private mLblTitle As MSForms.Label
Private mLblWarehouse As MSForms.Label
Private mLblStation As MSForms.Label
Private mLblUser As MSForms.Label
Private mLblRoot As MSForms.Label
Private mLblRootValue As MSForms.Label
Private mLblStatus As MSForms.Label
Private mTxtStation As MSForms.TextBox
Private mTxtUser As MSForms.TextBox

Private mAccepted As Boolean
Private mSelectedWarehouseId As String
Private mSelectedStationId As String
Private mSelectedRuntimeRoot As String
Private mSelectedUserId As String

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get SelectedWarehouseId() As String
    SelectedWarehouseId = mSelectedWarehouseId
End Property

Public Property Get SelectedStationId() As String
    SelectedStationId = mSelectedStationId
End Property

Public Property Get SelectedRuntimeRoot() As String
    SelectedRuntimeRoot = mSelectedRuntimeRoot
End Property

Public Property Get SelectedUserId() As String
    SelectedUserId = mSelectedUserId
End Property

Public Sub Configure(ByVal warehouseOptions As Collection, _
                     ByVal defaultWarehouseId As String, _
                     ByVal defaultStationId As String, _
                     ByVal defaultUserId As String)
    Dim item As Variant
    Dim rowIndex As Long
    Dim matchIndex As Long

    EnsureControls
    mCmbWarehouse.Clear
    matchIndex = -1

    If Not warehouseOptions Is Nothing Then
        For Each item In warehouseOptions
            mCmbWarehouse.AddItem CStr(item(0))
            rowIndex = mCmbWarehouse.ListCount - 1
            mCmbWarehouse.List(rowIndex, 1) = CStr(item(1))
            mCmbWarehouse.List(rowIndex, 2) = CStr(item(2))
            mCmbWarehouse.List(rowIndex, 3) = CStr(item(3))
            mCmbWarehouse.List(rowIndex, 4) = CStr(item(4))
            If matchIndex < 0 _
               And StrComp(CStr(item(1)), defaultWarehouseId, vbTextCompare) = 0 _
               And (defaultStationId = "" Or StrComp(CStr(item(2)), defaultStationId, vbTextCompare) = 0) Then
                matchIndex = rowIndex
            End If
        Next item
    End If

    If mCmbWarehouse.ListCount > 0 Then
        If matchIndex < 0 Then matchIndex = 0
        mCmbWarehouse.ListIndex = matchIndex
    End If

    mTxtStation.Value = IIf(defaultStationId = "", "S1", defaultStationId)
    mTxtUser.Value = defaultUserId
    ApplyWarehouseSelection
End Sub

Private Sub UserForm_Initialize()
    EnsureControls
End Sub

Private Sub EnsureControls()
    If Not mCmbWarehouse Is Nothing Then Exit Sub

    Me.Caption = "invSys Admin - Seed Inventory"
    Me.Width = 500
    Me.Height = 250

    Set mLblTitle = AddLabel("lblTitle", 12, 12, 456, 22, "Seed demo inventory into which warehouse?")
    Set mLblWarehouse = AddLabel("lblWarehouse", 12, 48, 92, 18, "Warehouse")
    Set mCmbWarehouse = AddCombo("cmbWarehouse", 108, 44, 348, 24)
    mCmbWarehouse.ColumnCount = 5
    mCmbWarehouse.ColumnWidths = "340 pt;0 pt;0 pt;0 pt;0 pt"
    mCmbWarehouse.MatchRequired = True
    mCmbWarehouse.Style = fmStyleDropDownList

    Set mLblStation = AddLabel("lblStation", 12, 82, 92, 18, "Station")
    Set mTxtStation = AddTextBox("txtStation", 108, 78, 90, 22)
    Set mLblUser = AddLabel("lblUser", 220, 82, 84, 18, "Admin user")
    Set mTxtUser = AddTextBox("txtUser", 304, 78, 152, 22)

    Set mLblRoot = AddLabel("lblRoot", 12, 116, 92, 18, "Runtime root")
    Set mLblRootValue = AddLabel("lblRootValue", 108, 116, 348, 36, "")
    mLblRootValue.WordWrap = True

    Set mLblStatus = AddLabel("lblStatus", 108, 158, 348, 18, "")
    mLblStatus.ForeColor = 255

    Set mBtnOK = AddButton("btnOK", 284, 186, 82, 28, "OK")
    Set mBtnCancel = AddButton("btnCancel", 374, 186, 82, 28, "Cancel")
End Sub

Private Function AddLabel(ByVal controlName As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal captionText As String) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    AddLabel.Left = leftPos
    AddLabel.Top = topPos
    AddLabel.Width = widthVal
    AddLabel.Height = heightVal
    AddLabel.Caption = captionText
End Function

Private Function AddCombo(ByVal controlName As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single) As MSForms.ComboBox
    Set AddCombo = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    AddCombo.Left = leftPos
    AddCombo.Top = topPos
    AddCombo.Width = widthVal
    AddCombo.Height = heightVal
End Function

Private Function AddTextBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    AddTextBox.Left = leftPos
    AddTextBox.Top = topPos
    AddTextBox.Width = widthVal
    AddTextBox.Height = heightVal
End Function

Private Function AddButton(ByVal controlName As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single, _
                           ByVal captionText As String) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    AddButton.Left = leftPos
    AddButton.Top = topPos
    AddButton.Width = widthVal
    AddButton.Height = heightVal
    AddButton.Caption = captionText
End Function

Private Sub mCmbWarehouse_Change()
    ApplyWarehouseSelection
End Sub

Private Sub mBtnOK_Click()
    If mCmbWarehouse.ListIndex < 0 Then
        mLblStatus.Caption = "Choose a warehouse."
        Exit Sub
    End If
    If Trim$(CStr(mTxtStation.Value)) = "" Then
        mLblStatus.Caption = "Station is required."
        Exit Sub
    End If
    If Trim$(CStr(mTxtUser.Value)) = "" Then
        mLblStatus.Caption = "Admin user is required."
        Exit Sub
    End If

    mSelectedWarehouseId = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 1))
    mSelectedStationId = Trim$(CStr(mTxtStation.Value))
    mSelectedRuntimeRoot = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 3))
    mSelectedUserId = Trim$(CStr(mTxtUser.Value))
    mAccepted = True
    Me.Hide
End Sub

Private Sub mBtnCancel_Click()
    mAccepted = False
    Me.Hide
End Sub

Private Sub ApplyWarehouseSelection()
    If mCmbWarehouse.ListIndex < 0 Then
        mLblRootValue.Caption = ""
        Exit Sub
    End If

    mTxtStation.Value = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 2))
    mLblRootValue.Caption = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 3))
    If Trim$(CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 4))) = "" _
       Or StrComp(CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 4)), "Ready", vbTextCompare) = 0 Then
        mLblStatus.Caption = ""
    Else
        mLblStatus.Caption = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 4))
    End If
End Sub
