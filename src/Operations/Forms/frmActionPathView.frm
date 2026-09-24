VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActionPathView
   Caption         =   "Action Path view"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   StartUpPosition =   1
End
Attribute VB_Name = "frmActionPathView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mOwner As frmActionPaths
Private mContext As String, mPathId As String, mBinding As String, mKey As String
Private mLoading As Boolean, mLayingOut As Boolean, mResizeReady As Boolean
Private mLayoutWidth As Single, mLayoutHeight As Single
Private mLayout As cOperationsAnchorManager
Private WithEvents mMethod As MSForms.ComboBox
Private WithEvents mRefresh As MSForms.CommandButton
Private WithEvents mClose As MSForms.CommandButton
Private mHowTo As MSForms.TextBox, mDiagnostic As MSForms.TextBox
Private mPair As MSForms.Label, mStatus As MSForms.Label

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object, choice As Variant
    mLoading = True
    Me.Caption = "Action Path view": Me.Width = 960: Me.Height = 680
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 840, 600
    For Each definition In Array( _
        Array("Label", "lblActionPathMethod", "Presentation", 12, 12, 100, 20), _
        Array("ComboBox", "cboActionPathView", "", 120, 10, 220, 24), _
        Array("CommandButton", "btnRefreshActionPathView", "Refresh", 350, 10, 92, 26), _
        Array("Label", "lblActionPathPair", "", 12, 46, 928, 112), _
        Array("Label", "lblActionPathHowTo", "Authored instructions - guide order", 12, 166, 450, 20), _
        Array("Label", "lblActionPathDiagnostic", "Observed actions and saved diagnostic result", 482, 166, 458, 20), _
        Array("TextBox", "txtActionPathHowTo", "", 12, 194, 450, 350), _
        Array("TextBox", "txtActionPathDiagnostic", "", 482, 194, 458, 350), _
        Array("Label", "lblActionPathViewStatus", "", 12, 556, 928, 44), _
        Array("CommandButton", "btnCloseActionPathView", "Close", 844, 610, 96, 26))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        If definition(0) = "Label" Or definition(0) = "CommandButton" Then control.Caption = definition(2)
        If definition(0) = "Label" Then control.WordWrap = True
        If definition(0) = "TextBox" Then
            control.MultiLine = True: control.WordWrap = True: control.Locked = True: control.ScrollBars = fmScrollBarsVertical
        End If
    Next definition
    Set mMethod = Me.Controls("cboActionPathView")
    Set mRefresh = Me.Controls("btnRefreshActionPathView"): Set mClose = Me.Controls("btnCloseActionPathView")
    Set mHowTo = Me.Controls("txtActionPathHowTo"): Set mDiagnostic = Me.Controls("txtActionPathDiagnostic")
    Set mPair = Me.Controls("lblActionPathPair"): Set mStatus = Me.Controls("lblActionPathViewStatus")
    mMethod.Style = fmStyleDropDownList
    For Each choice In Array("How-To", "Diagnostic", "Compare both"): mMethod.AddItem CStr(choice): Next choice
    mLoading = False
    ApplyLayout
End Sub

Public Function BindPair(ByVal owner As frmActionPaths, ByVal context As String, ByVal pathId As String, _
                         ByVal binding As String, ByVal key As String) As Boolean
    Set mOwner = owner: mContext = context: mPathId = pathId: mBinding = binding: mKey = key
    BindPair = RefreshView(True)
End Function

Public Function MatchesPair(ByVal context As String, ByVal pathId As String, ByVal binding As String) As Boolean
    MatchesPair = (mContext = context And mPathId = pathId And mBinding = binding And binding <> "")
End Function

Public Function RefreshView(Optional ByVal fresh As Boolean = False) As Boolean
    Dim instructions As String, diagnostic As String, provenance As String, method As String, notice As String, evaluationId As String
    On Error GoTo Failed
    If mLoading Then Exit Function
    mLoading = True
    If mOwner Is Nothing Or mBinding = "" Then GoTo Failed
    If Not mOwner.ReadViewEvaluation(mPathId, mBinding, evaluationId) Then GoTo Failed
    If Not modPathPresentation.Read(mContext, mPathId, mBinding, mKey, evaluationId, instructions, diagnostic, provenance, method, notice) Then GoTo Failed
    If fresh Then mMethod.Value = method
    mHowTo.Value = instructions: mDiagnostic.Value = diagnostic: mPair.Caption = provenance: mStatus.Caption = notice
    mMethod.Enabled = True
    RefreshView = True
    mLoading = False
    ApplyLayout
    Exit Function
Failed:
    If notice = "" Then notice = "Unavailable: the selected guide, recording or session changed. Reopen View guide and run."
    mHowTo.Value = "": mDiagnostic.Value = "": mPair.Caption = "": mStatus.Caption = notice
    mMethod.Enabled = False
    mLoading = False
End Function

Private Sub ApplyLayout()
    Dim width As Single, height As Single, half As Single, both As Boolean, howTo As Boolean
    If mLayout Is Nothing Or mLayingOut Then Exit Sub
    mLayingOut = True
    On Error GoTo Done
    mLayout.ApplyAnchoredLayout
    width = Me.InsideWidth - 24: height = Me.InsideHeight: half = (width - 12) / 2
    both = (mMethod.Value = "Compare both"): howTo = (mMethod.Value <> "Diagnostic")
    mPair.Move 12, 46, width, 112
    mStatus.Move 12, height - 84, width, 44
    mClose.Move Me.InsideWidth - 108, height - 32, 96, 26
    mHowTo.Visible = howTo: Me.Controls("lblActionPathHowTo").Visible = howTo
    mDiagnostic.Visible = Not howTo Or both: Me.Controls("lblActionPathDiagnostic").Visible = mDiagnostic.Visible
    mHowTo.Move 12, 194, IIf(both, half, width), height - 290
    Me.Controls("lblActionPathHowTo").Move 12, 166, mHowTo.Width, 20
    mDiagnostic.Move IIf(both, 24 + half, 12), 194, IIf(both, half, width), height - 290
    Me.Controls("lblActionPathDiagnostic").Move mDiagnostic.Left, 166, mDiagnostic.Width, 20
    mLayoutWidth = Me.InsideWidth: mLayoutHeight = Me.InsideHeight
Done:
    mLayingOut = False
End Sub

Private Sub mMethod_Change()
    If Not mLoading Then RefreshView
End Sub

Private Sub mRefresh_Click()
    RefreshView
End Sub

Private Sub mClose_Click()
    If Not mOwner Is Nothing Then mOwner.CloseActionPathView
End Sub

Public Sub ReleaseView()
    Set mOwner = Nothing
    mContext = "": mPathId = "": mBinding = "": mKey = ""
End Sub

Private Sub UserForm_Activate()
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeReady = True
    End If
    RefreshView
End Sub

Private Sub UserForm_Layout()
    If mLoading Or mLayingOut Then Exit Sub
    ' Control updates can queue Layout again without an operator resize.
    If Me.InsideWidth = mLayoutWidth And Me.InsideHeight = mLayoutHeight Then Exit Sub
    If mBinding <> "" Then RefreshView
    ApplyLayout
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        If Not mOwner Is Nothing Then mOwner.CloseActionPathView
    End If
End Sub
