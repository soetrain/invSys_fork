VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdminSettings
   Caption         =   "invSys Settings"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdminSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtCarrier As MSForms.TextBox
Private WithEvents mLstCarriers As MSForms.ListBox
Private WithEvents mBtnAdd As MSForms.CommandButton
Private WithEvents mBtnRemove As MSForms.CommandButton
Private WithEvents mBtnReset As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton

Private mLblStatus As MSForms.Label
Private mLoading As Boolean

Private Sub UserForm_Initialize()
    BuildLayout
    LoadCarriers
End Sub

Private Sub BuildLayout()
    Me.Caption = "invSys Settings"
    Me.Width = 390
    Me.Height = 285

    AddLabel "lblTitle", "Settings", 12, 10, 120, 18, True
    AddLabel "lblSection", "Shipping Carriers", 12, 42, 150, 18, True
    AddLabel "lblCarrier", "Carrier", 12, 72, 60, 18, False

    Set mTxtCarrier = AddTextBox("txtCarrier", 76, 68, 170, 22)
    Set mBtnAdd = AddButton("btnAdd", "Add", 258, 66, 48, 26)
    Set mBtnRemove = AddButton("btnRemove", "Remove", 312, 66, 58, 26)

    Set mLstCarriers = AddListBox("lstCarriers", 12, 104, 358, 96)
    With mLstCarriers
        .ColumnCount = 1
        .ColumnWidths = "320 pt"
    End With

    Set mBtnReset = AddButton("btnReset", "Reset Defaults", 12, 214, 84, 28)
    Set mBtnClose = AddButton("btnClose", "Close", 314, 214, 56, 28)
    Set mLblStatus = AddLabel("lblStatus", "", 108, 218, 196, 22, False)
End Sub

Private Sub LoadCarriers()
    Dim carriers As Variant
    Dim displayRows As Variant
    Dim idx As Long

    mLoading = True
    mLstCarriers.Clear
    carriers = modCarrierSettings.GetConfiguredCarriers()
    If Not IsEmpty(carriers) Then
        ReDim displayRows(0 To UBound(carriers) - 1, 0 To 0)
        For idx = LBound(carriers) To UBound(carriers)
            displayRows(idx - 1, 0) = CStr(carriers(idx))
        Next idx
        mLstCarriers.List = displayRows
    End If
    mLoading = False
End Sub

Private Sub mBtnAdd_Click()
    Dim carrierName As String

    carrierName = Trim$(CStr(mTxtCarrier.Value))
    If carrierName = "" Then
        ShowStatus "Enter a carrier."
        Exit Sub
    End If

    If modCarrierSettings.AddConfiguredCarrier(carrierName) Then
        mTxtCarrier.Value = ""
        LoadCarriers
        ShowStatus "Carrier added."
    Else
        ShowStatus "Carrier was not added."
    End If
End Sub

Private Sub mBtnRemove_Click()
    Dim carrierName As String

    If mLstCarriers.ListIndex < 0 Then
        ShowStatus "Select a carrier."
        Exit Sub
    End If

    carrierName = CStr(mLstCarriers.List(mLstCarriers.ListIndex, 0))
    If modCarrierSettings.RemoveConfiguredCarrier(carrierName) Then
        LoadCarriers
        ShowStatus "Carrier removed."
    Else
        ShowStatus "Carrier was not removed."
    End If
End Sub

Private Sub mBtnReset_Click()
    If MsgBox("Reset shipping carriers to defaults?", vbQuestion + vbYesNo, "invSys Settings") <> vbYes Then Exit Sub
    modCarrierSettings.ResetConfiguredCarriers
    LoadCarriers
    ShowStatus "Defaults restored."
End Sub

Private Sub mBtnClose_Click()
    Me.Hide
End Sub

Private Sub mLstCarriers_Click()
    If mLoading Then Exit Sub
    If mLstCarriers.ListIndex >= 0 Then mTxtCarrier.Value = CStr(mLstCarriers.List(mLstCarriers.ListIndex, 0))
End Sub

Private Sub ShowStatus(ByVal message As String)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = message
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
