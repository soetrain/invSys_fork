VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWarehouseConnection
   Caption         =   "Connect Warehouse Storage"
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
Private mTxtStation As MSForms.TextBox
Private WithEvents mBtnConnect As MSForms.CommandButton
Private WithEvents mBtnScan As MSForms.CommandButton
Private WithEvents mBtnOK As MSForms.CommandButton
Private WithEvents mBtnCancel As MSForms.CommandButton
Private WithEvents mLstRoots As MSForms.ListBox
Private WithEvents mLstTargets As MSForms.ListBox

Private mLblStatus As MSForms.Label
Private mWasAccepted As Boolean
Private mReason As String
Private mResizeInitialized As Boolean
Private mRequireStationInboxForRole As Boolean
Private mLblStationHelp As MSForms.Label

Private Const COLOR_INFO As Long = 0
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_WARNING As Long = 192
Private Const COLOR_ERROR As Long = 255

Private Sub UserForm_Initialize()
    Me.Caption = "Connect Warehouse Storage"
    Me.Width = 560
    Me.Height = 430
    BuildConnectionLayout
    mTxtRoot.Value = modNasConnection.GetPromptDefaultRoot()
    mTxtUser.Value = modNasConnection.GetRememberedNasUser()
    If mReason <> "" Then
        ShowStatus mReason, COLOR_INFO
    Else
        ShowStatus "Connect server storage first, then sign in with an invSys user account.", COLOR_INFO
    End If
End Sub

Private Sub UserForm_Activate()
    If mResizeInitialized Then Exit Sub
    On Error Resume Next
    modUserFormResizeWin.EnableResizableUserForm Me, True, True
    On Error GoTo 0
    mResizeInitialized = True
End Sub

Public Sub InitializeConnectionPrompt(Optional ByVal reason As String = "", _
                                      Optional ByVal requireStationInbox As Boolean = False)
    mReason = Trim$(reason)
    mRequireStationInboxForRole = requireStationInbox
    UpdateStationHelp
    If Not mLblStatus Is Nothing Then
        If mReason <> "" Then
            ShowStatus mReason, COLOR_INFO
        Else
            ShowStatus "Connect server storage first, then sign in with an invSys user account.", COLOR_INFO
        End If
    End If
End Sub

Public Property Get WasAccepted() As Boolean
    WasAccepted = mWasAccepted
End Property

Private Sub BuildConnectionLayout()
    AddLabel "lblTitle", "Connect Warehouse Storage", 18, 14, 260, 18, True
    Set mLblStatus = AddLabel("lblStatus", "", 18, 40, 500, 34, False)

    AddLabel "lblRoot", "Selected root", 18, 88, 84, 18, False
    Set mTxtRoot = AddTextBox("txtRoot", 104, 84, 322, 22)
    Set mBtnScan = AddButton("btnScan", "Scan Roots", 436, 83, 70, 24)

    AddLabel "lblUser", "Server user", 18, 122, 84, 18, False
    Set mTxtUser = AddTextBox("txtUser", 104, 118, 160, 22)
    AddLabel "lblPassword", "Password", 276, 122, 66, 18, False
    Set mTxtPassword = AddTextBox("txtPassword", 346, 118, 80, 22)
    mTxtPassword.PasswordChar = "*"
    Set mBtnConnect = AddButton("btnConnect", "Connect", 436, 117, 70, 24)

    AddLabel "lblRoots", "Discovered NAS roots", 18, 156, 180, 18, False
    Set mLstRoots = AddListBox("lstRoots", 18, 178, 488, 54)

    AddLabel "lblStation", "Station", 18, 240, 84, 18, False
    Set mTxtStation = AddTextBox("txtStation", 104, 236, 160, 22)
    mTxtStation.Value = modStationIdentity.CurrentComputerStationId()
    mTxtStation.Locked = True
    mTxtStation.BackColor = &HEFEFEF
    Set mLblStationHelp = AddLabel("lblStationHelp", "", 276, 237, 230, 32, False)
    UpdateStationHelp

    AddLabel "lblTargets", "Warehouse runtimes", 18, 276, 130, 18, False
    Set mLstTargets = AddListBox("lstTargets", 18, 298, 488, 72)

    Set mBtnOK = AddButton("btnOK", "OK", 350, 380, 74, 26)
    Set mBtnCancel = AddButton("btnCancel", "Cancel", 432, 380, 74, 26)
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

Private Function AddComboBox(ByVal controlName As String, _
                             ByVal leftPos As Single, _
                             ByVal topPos As Single, _
                             ByVal widthVal As Single, _
                             ByVal heightVal As Single) As MSForms.ComboBox
    Set AddComboBox = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    With AddComboBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Style = fmStyleDropDownList
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
    On Error GoTo ConnectFailed
    Dim statusCode As NasStatusCode
    Dim rootPath As String
    Dim userName As String
    Dim passwordText As String
    Dim previousCursor As XlMousePointer

    rootPath = Trim$(CStr(mTxtRoot.Value))
    userName = Trim$(CStr(mTxtUser.Value))
    passwordText = CStr(mTxtPassword.Value)
    If rootPath = "" Or userName = "" Or Len(passwordText) = 0 Then
        ShowStatus "Enter root, server user, and server password before connecting.", COLOR_WARNING
        Exit Sub
    End If

    ShowStatus "Connecting to warehouse storage. Windows may briefly show Excel as busy while the NAS authenticates...", COLOR_INFO
    previousCursor = Application.Cursor
    Application.Cursor = xlWait
    Me.Repaint
    DoEvents
    statusCode = modNasConnection.ConnectNasRootWithCredentials(rootPath, userName, passwordText)
    Application.Cursor = previousCursor
    mTxtPassword.Value = vbNullString
    If statusCode = NAS_OK Then
        ShowStatus "Storage connected. Scanning the selected root for warehouse runtimes...", COLOR_SUCCESS
        ScanConnectedRoot
    Else
        ShowStatus modNasConnection.GetConnectionStatus(), COLOR_ERROR
    End If
    Exit Sub

ConnectFailed:
    Application.Cursor = previousCursor
    ShowStatus "Warehouse storage connection failed: " & Err.Description, COLOR_ERROR
End Sub

Private Sub mBtnScan_Click()
    ScanRootCandidates
End Sub

Private Sub mBtnOK_Click()
    Dim selectedPath As String
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode

    If mLstTargets.ListIndex < 0 Then
        ShowStatus "Select a warehouse runtime first.", COLOR_WARNING
        Exit Sub
    End If
    If mRequireStationInboxForRole And SelectedStationValue() = "" Then
        ShowStatus "The Windows computer name could not be resolved for this station.", COLOR_WARNING
        Exit Sub
    End If

    selectedPath = CStr(mLstTargets.Value)
    statusCode = modNasConnection.SelectWarehouseTarget( _
        Trim$(CStr(mTxtRoot.Value)), _
        selectedPath, _
        target, _
        SelectedStationValue(), _
        mRequireStationInboxForRole)

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

Private Sub mLstTargets_Change()
    RefreshStationsForSelectedTarget
End Sub

Private Sub mLstRoots_Change()
    If mLstRoots.ListIndex < 0 Then Exit Sub
    mTxtRoot.Value = CStr(mLstRoots.Value)
    mLstTargets.Clear
End Sub

Private Sub ScanRootCandidates()
    Dim roots As Collection
    Dim item As Variant

    ShowStatus "Discovering visible NAS roots. Select one, enter server credentials, then connect.", COLOR_INFO
    Me.Repaint
    DoEvents
    Set roots = modNasConnection.DiscoverVisibleNasRoots()
    mLstRoots.Clear
    mLstTargets.Clear
    For Each item In roots
        mLstRoots.AddItem CStr(item)
    Next item

    If mLstRoots.ListCount > 0 Then
        mLstRoots.ListIndex = 0
        ShowStatus "Found " & CStr(mLstRoots.ListCount) & " NAS root(s). Select one, enter server credentials, then connect.", COLOR_SUCCESS
    Else
        ShowStatus "No visible NAS roots were found. Enter an authorized root as a fallback, then enter server credentials and connect.", COLOR_WARNING
    End If
End Sub

Private Sub ScanConnectedRoot()
    Dim targets As Collection
    Dim item As Variant
    Dim rootPath As String

    rootPath = Trim$(CStr(mTxtRoot.Value))
    If rootPath = "" Then
        ShowStatus "Select a discovered root or enter an authorized root first.", COLOR_WARNING
        Exit Sub
    End If

    ShowStatus "Scanning the connected root for warehouse runtimes...", COLOR_INFO
    Me.Repaint
    DoEvents
    Set targets = modNasConnection.ScanNasRoot(rootPath)
    mLstTargets.Clear
    For Each item In targets
        mLstTargets.AddItem CStr(item)
    Next item

    If mLstTargets.ListCount > 0 Then
        mLstTargets.ListIndex = 0
        RefreshStationsForSelectedTarget
        ShowStatus "Found " & CStr(mLstTargets.ListCount) & " warehouse runtime(s). Select one and continue to invSys sign-in.", COLOR_SUCCESS
    Else
        ShowStatus "No warehouse runtime folders were found under this root.", COLOR_WARNING
    End If
End Sub

Private Sub RefreshStationsForSelectedTarget()
    If mTxtStation Is Nothing Then Exit Sub
    mTxtStation.Value = modStationIdentity.CurrentComputerStationId()
End Sub

Private Function SelectedStationValue() As String
    SelectedStationValue = modStationIdentity.CurrentComputerStationId()
End Function

Private Sub UpdateStationHelp()
    If mLblStationHelp Is Nothing Then Exit Sub
    mLblStationHelp.Caption = "Windows computer name; used automatically after sign-in."
End Sub

Private Sub ShowStatus(ByVal messageText As String, ByVal foreColor As Long)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = messageText
    mLblStatus.ForeColor = foreColor
End Sub
