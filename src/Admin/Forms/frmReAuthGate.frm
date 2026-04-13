VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReAuthGate 
   Caption         =   "Re-Authenticate"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmReAuthGate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReAuthGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAuthenticated As Boolean
Private mFailureCount As Long
Private mRequiredRole As String
Private mAdminUser As String
Private mLockedOut As Boolean
Private mResizeInitialized As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Re-Authenticate"
    Me.StartUpPosition = 1
    Me.txtPassword.PasswordChar = "*"
    InitializeGate "ADMIN_MAINT", ResolveDefaultAdminUserReAuth()
End Sub

Private Sub UserForm_Activate()
    If mResizeInitialized Then Exit Sub
    modUserFormResizeWin.EnableResizableUserForm Me
    mResizeInitialized = True
End Sub

Public Sub InitializeGate(Optional ByVal RequiredRole As String = "ADMIN_MAINT", _
                          Optional ByVal AdminUser As String = "")
    mRequiredRole = Trim$(RequiredRole)
    If mRequiredRole = "" Then mRequiredRole = "ADMIN_MAINT"

    mAdminUser = Trim$(AdminUser)
    If mAdminUser = "" Then mAdminUser = ResolveDefaultAdminUserReAuth()

    mAuthenticated = False
    mFailureCount = 0
    mLockedOut = False

    Me.lblPrompt.Caption = "Re-enter your password to continue."
    Me.lblAdminUserValue.Caption = mAdminUser
    Me.lblRoleValue.Caption = mRequiredRole
    Me.txtPassword.Value = ""
    Me.btnOK.Enabled = True
    ClearErrorMessage
End Sub

Public Sub ShowGate(Optional ByVal RequiredRole As String = "ADMIN_MAINT", _
                    Optional ByVal AdminUser As String = "")
    InitializeGate RequiredRole, AdminUser
    Me.Show vbModal
End Sub

Public Property Get Authenticated() As Boolean
    Authenticated = mAuthenticated
End Property

Public Property Get FailureCount() As Long
    FailureCount = mFailureCount
End Property

Public Property Get IsLockedOut() As Boolean
    IsLockedOut = mLockedOut
End Property

Public Property Get ErrorText() As String
    ErrorText = Trim$(Me.lblError.Caption)
End Property

Public Property Get IsSubmitEnabled() As Boolean
    IsSubmitEnabled = CBool(Me.btnOK.Enabled)
End Property

Public Sub SetPasswordTextForTest(ByVal passwordText As String)
    Me.txtPassword.Value = passwordText
End Sub

Public Sub SimulateSubmit()
    btnOK_Click
End Sub

Public Sub SimulateCancel()
    btnCancel_Click
End Sub

Private Sub btnOK_Click()
    Dim passwordText As String

    If mLockedOut Then Exit Sub

    passwordText = CStr(Me.txtPassword.Value)
    If modAuth.ValidateUserCredential(mAdminUser, passwordText, mRequiredRole) Then
        mAuthenticated = True
        ClearErrorMessage
        Me.Hide
        Exit Sub
    End If

    mAuthenticated = False
    mFailureCount = mFailureCount + 1
    ShowErrorMessage "Invalid credentials - access denied"
    Me.txtPassword.Value = ""

    If mFailureCount >= 3 Then
        mLockedOut = True
        Me.btnOK.Enabled = False
        modDiagnostics.LogDiagnosticEvent "REAUTH", _
            "Lockout|UserId=" & mAdminUser & "|Role=" & mRequiredRole & "|Failures=" & CStr(mFailureCount)
        Me.Hide
    End If
End Sub

Private Sub btnCancel_Click()
    mAuthenticated = False
    Me.Hide
End Sub

Private Sub ClearErrorMessage()
    Me.lblError.Caption = ""
    Me.lblError.ForeColor = RGB(192, 0, 0)
End Sub

Private Sub ShowErrorMessage(ByVal messageText As String)
    Me.lblError.Caption = Trim$(messageText)
    Me.lblError.ForeColor = RGB(192, 0, 0)
End Sub

Private Function ResolveDefaultAdminUserReAuth() As String
    ResolveDefaultAdminUserReAuth = Trim$(Environ$("USERNAME"))
    If ResolveDefaultAdminUserReAuth = "" Then ResolveDefaultAdminUserReAuth = Trim$(Application.UserName)
End Function

