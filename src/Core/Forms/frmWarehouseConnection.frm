VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWarehouseConnection
   Caption         =   "Connect / Select Warehouse"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWarehouseConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtRoot As MSForms.TextBox
Private WithEvents mTxtUser As MSForms.TextBox
Private WithEvents mTxtPassword As MSForms.TextBox
Private WithEvents mTxtStation As MSForms.TextBox
Private WithEvents mChkRequireStation As MSForms.CheckBox
Private WithEvents mBtnConnect As MSForms.CommandButton
Private WithEvents mBtnScan As MSForms.CommandButton
Private WithEvents mBtnOK As MSForms.CommandButton
Private WithEvents mBtnCancel As MSForms.CommandButton
Private WithEvents mLstTargets As MSForms.ListBox

Private mLblStatus As MSForms.Label
Private mWasAccepted As Boolean
Private mReason As String

Private Const COLOR_INFO As Long = 0
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_WARNING As Long = 192
Private Const COLOR_ERROR As Long = 255

Private Sub UserForm_Initialize()
    Me.Caption = "Connect / Select Warehouse"
    Me.Width = 560
    Me.Height = 390
    BuildConnectionLayout
    mTxtRoot.Value = modNasConnection.GetPromptDefaultRoot()
    mTxtUser.Value = modNasConnection.GetRememberedNasUser()
    If mReason <> "" Then
        ShowStatus mReason, COLOR_INFO
    Else
        ShowStatus "Enter a warehouse root, connect if needed, then scan.", COLOR_INFO
    End If
End Sub

Public Sub InitializeConnectionPrompt(Optional ByVal reason As String = "")
    mReason = Trim$(reason)
    If Not mLblStatus Is Nothing Then
        If mReason <> "" Then
            ShowStatus mReason, COLOR_INFO
        Else
            ShowStatus "Enter a warehouse root, connect if needed, then scan.", COLOR_INFO
        End If
    End If
End Sub

Public Property Get WasAccepted() As Boolean
    WasAccepted = mWasAccepted
End Property

Private Sub BuildConnectionLayout()
    AddLabel "lblTitle", "Connect / Select Warehouse", 18, 14, 260, 18, True
    Set mLblStatus = AddLabel("lblStatus", "", 18, 40, 500, 34, False)

    AddLabel "lblRoot", "Root", 18, 88, 84, 18, False
    Set mTxtRoot = AddTextBox("txtRoot", 104, 84, 322, 22)
    Set mBtnScan = AddButton("btnScan", "Scan", 436, 83, 70, 24)

    AddLabel "lblUser", "NAS user", 18, 122, 84, 18, False
    Set mTxtUser = AddTextBox("txtUser", 104, 118, 160, 22)
    AddLabel "lblPassword", "Password", 276, 122, 66, 18, False
    Set mTxtPassword = AddTextBox("txtPassword", 346, 118, 80, 22)
    mTxtPassword.PasswordChar = "*"
    Set mBtnConnect = AddButton("btnConnect", "Connect", 436, 117, 70, 24)

    AddLabel "lblStation", "Station", 18, 156, 84, 18, False
    Set mTxtStation = AddTextBox("txtStation", 104, 152, 80, 22)
    Set mChkRequireStation = AddCheckBox("chkRequireStation", "Require station inbox", 198, 154, 160, 18)

    AddLabel "lblTargets", "Warehouse runtimes", 18, 194, 130, 18, False
    Set mLstTargets = AddListBox("lstTargets", 18, 216, 488, 92)

    Set mBtnOK = AddButton("btnOK", "OK", 350, 324, 74, 26)
    Set mBtnCancel = AddButton("btnCancel", "Cancel", 432, 324, 74, 26)
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

Private Function AddCheckBox(ByVal controlName As String, _
                             ByVal captionText As String, _
                             ByVal leftPos As Single, _
                             ByVal topPos As Single, _
                             ByVal widthVal As Single, _
                             ByVal heightVal As Single) As MSForms.CheckBox
    Set AddCheckBox = Me.Controls.Add("Forms.CheckBox.1", controlName, True)
    With AddCheckBox
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Sub mBtnConnect_Click()
    Dim statusCode As NasStatusCode
    Dim rootPath As String
    Dim userName As String
    Dim passwordText As String

    rootPath = Trim$(CStr(mTxtRoot.Value))
    userName = Trim$(CStr(mTxtUser.Value))
    passwordText = CStr(mTxtPassword.Value)
    If rootPath = "" Or userName = "" Or Len(passwordText) = 0 Then
        ShowStatus "Enter root, NAS user, and password before connecting.", COLOR_WARNING
        Exit Sub
    End If

    statusCode = modNasConnection.ConnectNasRootWithCredentials(rootPath, userName, passwordText)
    mTxtPassword.Value = vbNullString
    If statusCode = NAS_OK Then
        ShowStatus "Connected. Scan the root to choose a warehouse.", COLOR_SUCCESS
        ScanRoot
    Else
        ShowStatus modNasConnection.GetConnectionStatus(), COLOR_ERROR
    End If
End Sub

Private Sub mBtnScan_Click()
    ScanRoot
End Sub

Private Sub mBtnOK_Click()
    Dim selectedPath As String
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode

    If mLstTargets.ListIndex < 0 Then
        ShowStatus "Select a warehouse runtime first.", COLOR_WARNING
        Exit Sub
    End If

    selectedPath = CStr(mLstTargets.Value)
    statusCode = modNasConnection.SelectWarehouseTarget( _
        Trim$(CStr(mTxtRoot.Value)), _
        selectedPath, _
        target, _
        Trim$(CStr(mTxtStation.Value)), _
        CBool(mChkRequireStation.Value))

    If statusCode = NAS_OK Then
        mWasAccepted = True
        Me.Hide
    Else
        ShowStatus "Target selection failed. Status: " & CStr(statusCode), COLOR_ERROR
    End If
End Sub

Private Sub mBtnCancel_Click()
    mWasAccepted = False
    Me.Hide
End Sub

Private Sub ScanRoot()
    Dim targets As Collection
    Dim item As Variant
    Dim rootPath As String
    Dim statusCode As NasStatusCode

    rootPath = Trim$(CStr(mTxtRoot.Value))
    If rootPath = "" Then
        ShowStatus "Enter a warehouse root first.", COLOR_WARNING
        Exit Sub
    End If

    statusCode = modNasConnection.TryRevalidateRememberedRoot(rootPath)
    If statusCode <> NAS_OK Then
        ShowStatus modNasConnection.GetConnectionStatus(), COLOR_WARNING
        Exit Sub
    End If

    Set targets = modNasConnection.ScanNasRoot(rootPath)
    mLstTargets.Clear
    For Each item In targets
        mLstTargets.AddItem CStr(item)
    Next item

    If mLstTargets.ListCount > 0 Then
        mLstTargets.ListIndex = 0
        ShowStatus "Found " & CStr(mLstTargets.ListCount) & " warehouse runtime(s).", COLOR_SUCCESS
    Else
        ShowStatus "No warehouse runtime folders were found under this root.", COLOR_WARNING
    End If
End Sub

Private Sub ShowStatus(ByVal messageText As String, ByVal foreColor As Long)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = messageText
    mLblStatus.ForeColor = foreColor
End Sub
