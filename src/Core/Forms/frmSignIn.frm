VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSignIn
   Caption         =   "invSys Sign In"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4620
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtUser As MSForms.TextBox
Private WithEvents mTxtSecret As MSForms.TextBox
Private WithEvents mBtnOK As MSForms.CommandButton
Private WithEvents mBtnCancel As MSForms.CommandButton

Private mLblTarget As MSForms.Label
Private mLblStatus As MSForms.Label
Private mTarget As WarehouseTarget
Private mRequiredCapability As String
Private mResultStatus As AuthStatusCode
Private mHadAttempt As Boolean

Private Const COLOR_INFO As Long = 0
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_WARNING As Long = 192
Private Const COLOR_ERROR As Long = 255

Private Sub UserForm_Initialize()
    Me.Caption = "invSys Sign In"
    Me.Width = 410
    Me.Height = 280
    mResultStatus = AUTH_CANCELLED
    BuildSignInLayout
    RenderTarget
End Sub

Public Sub InitializeSignIn(ByVal target As WarehouseTarget, Optional ByVal requiredCapability As String = "")
    Set mTarget = target
    mRequiredCapability = UCase$(Trim$(requiredCapability))
    If mTxtUser Is Nothing Then Exit Sub
    mTxtUser.Value = modAuth.GetCurrentUserId()
    RenderTarget
End Sub

Public Property Get ResultStatus() As AuthStatusCode
    ResultStatus = mResultStatus
End Property

Private Sub BuildSignInLayout()
    AddLabel "lblTitle", "Sign in to invSys", 18, 14, 180, 20, True
    Set mLblTarget = AddLabel("lblTarget", "", 18, 42, 350, 32, False)
    Set mLblStatus = AddLabel("lblStatus", "", 18, 78, 350, 32, False)

    AddLabel "lblUser", "Account", 18, 122, 76, 18, False
    Set mTxtUser = AddTextBox("txtUser", 102, 118, 180, 22)
    AddLabel "lblSecret", "PIN/password", 18, 156, 84, 18, False
    Set mTxtSecret = AddTextBox("txtSecret", 102, 152, 180, 22)
    mTxtSecret.PasswordChar = "*"

    Set mBtnOK = AddButton("btnOK", "Sign In", 206, 198, 76, 26)
    Set mBtnCancel = AddButton("btnCancel", "Cancel", 292, 198, 76, 26)
    ShowStatus "Enter your invSys user ID and PIN/password.", COLOR_INFO
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

Private Sub mBtnOK_Click()
    Dim userId As String
    Dim secretText As String

    userId = Trim$(CStr(mTxtUser.Value))
    secretText = CStr(mTxtSecret.Value)
    If userId = "" Or Len(secretText) = 0 Then
        ShowStatus "Enter user ID and PIN/password.", COLOR_WARNING
        Exit Sub
    End If

    mHadAttempt = True
    mResultStatus = modAuth.ValidateUserCredentialForTarget(userId, secretText, mTarget, mRequiredCapability)
    mTxtSecret.Value = vbNullString
    If mResultStatus = AUTH_OK Then
        ShowStatus "Signed in.", COLOR_SUCCESS
        Me.Hide
    Else
        ShowStatus AuthStatusMessage(mResultStatus), COLOR_ERROR
    End If
End Sub

Private Sub mBtnCancel_Click()
    If Not mHadAttempt Then mResultStatus = AUTH_CANCELLED
    Me.Hide
End Sub

Private Sub RenderTarget()
    If mLblTarget Is Nothing Then Exit Sub
    If mTarget Is Nothing Then
        mLblTarget.Caption = "No warehouse server connected."
    Else
        mLblTarget.Caption = "Warehouse: " & mTarget.WarehouseId & "    Station: " & IIf(mTarget.StationId = "", "<roaming>", mTarget.StationId)
    End If
End Sub

Private Sub ShowStatus(ByVal messageText As String, ByVal foreColor As Long)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = messageText
    mLblStatus.ForeColor = foreColor
End Sub

Private Function AuthStatusMessage(ByVal statusCode As AuthStatusCode) As String
    Select Case statusCode
        Case AUTH_USER_NOT_FOUND
            AuthStatusMessage = "User was not found for this warehouse."
        Case AUTH_CREDENTIAL_REJECTED
            AuthStatusMessage = "PIN/password was rejected."
        Case AUTH_NO_CAPABILITIES
            AuthStatusMessage = "User lacks the required capability."
        Case AUTH_WORKBOOK_UNREADABLE
            AuthStatusMessage = "Auth workbook could not be read."
        Case AUTH_WAREHOUSE_MISMATCH
            AuthStatusMessage = "Warehouse target mismatch."
        Case Else
            AuthStatusMessage = "Sign-in failed. Status: " & CStr(statusCode)
    End Select
End Function
