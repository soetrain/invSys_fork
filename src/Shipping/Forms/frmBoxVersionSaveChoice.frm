VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBoxVersionSaveChoice
   Caption         =   "Save Box Version"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBoxVersionSaveChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mBtnUpdate As MSForms.CommandButton
Attribute mBtnUpdate.VB_VarHelpID = -1
Private WithEvents mBtnNew As MSForms.CommandButton
Attribute mBtnNew.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1

Private mLblTitle As MSForms.Label
Private mLblBody As MSForms.Label
Private mChoice As Long
Private mBoxName As String
Private mVersionLabel As String
Private mAnchors As Object
Private mResizeInitialized As Boolean

Private Const CHOICE_CANCEL As Long = 0
Private Const CHOICE_UPDATE As Long = 1
Private Const CHOICE_NEW As Long = 2
Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8

Private Sub UserForm_Initialize()
    mChoice = CHOICE_CANCEL
    BuildChoiceLayout
    RenderChoiceText
End Sub

Private Sub UserForm_Activate()
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me
        mResizeInitialized = True
    End If
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub UserForm_Layout()
    If mAnchors Is Nothing Then Exit Sub
    mAnchors.ResizeControls
End Sub

Private Sub UserForm_Terminate()
    Set mAnchors = Nothing
End Sub

Public Sub InitializeChoice(ByVal boxName As String, ByVal versionLabel As String)
    mBoxName = Trim$(boxName)
    mVersionLabel = Trim$(versionLabel)
    If mLblBody Is Nothing Then Exit Sub
    RenderChoiceText
End Sub

Public Property Get Choice() As Long
    Choice = mChoice
End Property

Private Sub BuildChoiceLayout()
    Me.Caption = "Save Box Version"
    Me.Width = 430
    Me.Height = 230

    Set mLblTitle = AddLabel("lblTitle", "Save Box Version", 18, 14, 360, 22, True)
    Set mLblBody = AddLabel("lblBody", "", 18, 44, 370, 70, False)

    Set mBtnUpdate = AddButton("btnUpdate", "Update Version", 18, 132, 112, 28)
    Set mBtnNew = AddButton("btnNewVersion", "New Version", 144, 132, 112, 28)
    Set mBtnCancel = AddButton("btnCancel", "Cancel", 270, 132, 90, 28)
    InitializeChoiceAnchors
End Sub

Private Sub RenderChoiceText()
    Dim textValue As String

    If mLblBody Is Nothing Then Exit Sub
    textValue = "You are editing " & IIf(mBoxName <> "", mBoxName, "this box")
    If mVersionLabel <> "" Then textValue = textValue & " " & mVersionLabel
    textValue = textValue & "." & vbCrLf & vbCrLf & _
                "Update the selected version, or save these rows as a new version?"
    mLblBody.Caption = textValue
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

Private Sub mBtnUpdate_Click()
    mChoice = CHOICE_UPDATE
    Me.Hide
End Sub

Private Sub mBtnNew_Click()
    mChoice = CHOICE_NEW
    Me.Hide
End Sub

Private Sub mBtnCancel_Click()
    mChoice = CHOICE_CANCEL
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    mChoice = CHOICE_CANCEL
End Sub

Private Sub InitializeChoiceAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me

    mAnchors.Add mLblTitle, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblBody, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnUpdate, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnNew, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnCancel, ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub
