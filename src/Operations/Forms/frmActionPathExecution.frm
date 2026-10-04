VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActionPathExecution
   Caption         =   "Configure execution"
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   StartUpPosition =   1
End
Attribute VB_Name = "frmActionPathExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mOwner As frmActionPathLibrary
Private mContext As String
Private mKey As String
Private mToken As String
Private mLoading As Boolean
Private mClosing As Boolean
Private mLayout As cOperationsAnchorManager
Private mGuide As MSForms.Label
Private mProfile As MSForms.Label
Private mStatus As MSForms.Label
Private WithEvents mSteps As MSForms.ListBox
Private WithEvents mInput As MSForms.ComboBox
Private mBinding As MSForms.ComboBox
Private mValue As MSForms.TextBox
Private WithEvents mApply As MSForms.CommandButton
Private WithEvents mSave As MSForms.CommandButton
Private WithEvents mClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object
    Me.Caption = "Configure execution": Me.Width = 900: Me.Height = 620
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 760, 620
    For Each definition In Array( _
        Array("Label", "lblExecutionGuide", "", 12, 12, 868, 44, 7), _
        Array("Label", "lblExecutionProfile", "", 12, 62, 868, 44, 7), _
        Array("Label", "lblExecutionSteps", "Guide steps", 12, 114, 200, 20, 3), _
        Array("ListBox", "lstExecutionSteps", "", 12, 140, 868, 146, 7), _
        Array("Label", "lblExecutionInput", "Input", 12, 300, 220, 20, 3), _
        Array("ComboBox", "cboExecutionInput", "", 12, 324, 280, 24, 3), _
        Array("Label", "lblExecutionBinding", "Binding", 310, 300, 160, 20, 3), _
        Array("ComboBox", "cboExecutionBinding", "", 310, 324, 180, 24, 3), _
        Array("CommandButton", "btnApplyExecutionInput", "Apply input", 704, 324, 176, 26, 6), _
        Array("Label", "lblExecutionValue", "Training value or registered prompt", 12, 362, 450, 20, 3), _
        Array("TextBox", "txtExecutionValue", "", 12, 388, 868, 28, 7), _
        Array("Label", "lblExecutionHelp", "Source entity is chosen from the training warehouse when setting up a run. Saving inputs does not execute the guide.", 12, 428, 868, 42, 7), _
        Array("Label", "lblExecutionStatus", "", 12, 486, 868, 40, 13), _
        Array("CommandButton", "btnSaveExecutionProfile", "Save execution profile", 12, 548, 220, 28, 9), _
        Array("CommandButton", "btnCloseExecution", "Close", 782, 548, 98, 28, 12))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        If definition(0) = "Label" Or definition(0) = "CommandButton" Then control.Caption = definition(2)
        If definition(0) = "Label" Then control.WordWrap = True
        mLayout.RegisterControl control, CLng(definition(7))
    Next definition
    Set mGuide = Me.Controls("lblExecutionGuide"): Set mProfile = Me.Controls("lblExecutionProfile")
    Set mStatus = Me.Controls("lblExecutionStatus"): Set mSteps = Me.Controls("lstExecutionSteps")
    Set mInput = Me.Controls("cboExecutionInput"): Set mBinding = Me.Controls("cboExecutionBinding")
    Set mValue = Me.Controls("txtExecutionValue"): Set mApply = Me.Controls("btnApplyExecutionInput")
    Set mSave = Me.Controls("btnSaveExecutionProfile"): Set mClose = Me.Controls("btnCloseExecution")
    mSteps.ColumnCount = 2: mSteps.BoundColumn = 1: mSteps.ColumnWidths = "0 pt;820 pt": mSteps.IntegralHeight = False
    mInput.Style = fmStyleDropDownList: mBinding.Style = fmStyleDropDownList
    mBinding.AddItem "Literal": mBinding.AddItem "Prompt"
End Sub

Public Function BindContext(ByVal context As String, ByVal key As String, ByVal token As String, ByVal owner As frmActionPathLibrary, ByRef notice As String) As Boolean
    mContext = context: mKey = key: mToken = token: Set mOwner = owner
    BindContext = RefreshDraft()
End Function

Public Function Matches(ByVal context As String, ByVal key As String) As Boolean
    Matches = (mToken <> "" And context = mContext And key = mKey)
End Function

Private Function RefreshDraft() As Boolean
    Dim rows As String, guide As String, profile As String, notice As String, row As Variant, fields As Variant, selected As String
    If Not modExecutionProfile.ReadDraft(mContext, mToken, rows, guide, profile, notice) Then Invalidate notice: Exit Function
    If mSteps.ListIndex >= 0 Then selected = CStr(mSteps.Value)
    mLoading = True: mSteps.Clear
    For Each row In Split(rows, vbCrLf)
        If CStr(row) <> "" Then
            fields = Split(CStr(row), vbTab)
            mSteps.AddItem CStr(fields(0)): mSteps.List(mSteps.ListCount - 1, 1) = CStr(fields(1))
            If fields(0) = selected Then mSteps.ListIndex = mSteps.ListCount - 1
        End If
    Next row
    mGuide.Caption = guide: mProfile.Caption = profile: mStatus.Caption = notice
    mLoading = False: RefreshDraft = True
End Function

Private Sub mSteps_Change()
    Dim names As String, name As Variant
    If mLoading Then Exit Sub
    mLoading = True: mInput.Clear: mValue.Value = "": mBinding.ListIndex = -1
    If mSteps.ListIndex >= 0 Then
        names = modExecutionProfile.InputNames(mContext, mToken, CStr(mSteps.Value))
        If names <> "" Then
            For Each name In Split(names, "|"): mInput.AddItem CStr(name): Next name
        End If
    End If
    mLoading = False
End Sub

Private Sub mInput_Change()
    Dim kind As String, value As String, notice As String
    If mLoading Or mInput.ListIndex < 0 Or mSteps.ListIndex < 0 Then Exit Sub
    If modExecutionProfile.ReadInput(mContext, mToken, CStr(mSteps.Value), CStr(mInput.Value), kind, value, notice) Then
        mBinding.Value = kind: mValue.Value = value
    Else
        Invalidate notice
    End If
    mStatus.Caption = notice
End Sub

Private Sub mApply_Click()
    Dim notice As String, applied As Boolean
    If mSteps.ListIndex < 0 Or mInput.ListIndex < 0 Then mStatus.Caption = "Select a guide step and input.": Exit Sub
    applied = modExecutionProfile.WriteInput(mContext, mToken, CStr(mSteps.Value), CStr(mInput.Value), CStr(mBinding.Value), CStr(mValue.Value), notice)
    If Not applied Then applied = RefreshDraft()
    mStatus.Caption = notice
End Sub

Private Sub mSave_Click()
    Dim notice As String, saved As Boolean
    saved = modExecutionProfile.SaveDraft(mContext, mToken, notice)
    saved = RefreshDraft()
    mStatus.Caption = notice
End Sub

Private Sub Invalidate(ByVal notice As String)
    ReleaseDraft
    mStatus.Caption = notice
End Sub

Public Sub ReleaseDraft()
    mClosing = True: Set mLayout = Nothing
    modExecutionProfile.CloseDraft mContext, mToken
    mLoading = True: mSteps.Clear: mInput.Clear: mValue.Value = "": mBinding.ListIndex = -1
    mGuide.Caption = "": mProfile.Caption = "": mSave.Enabled = False: mApply.Enabled = False
    mContext = "": mKey = "": mToken = "": mLoading = False
End Sub

Private Sub mClose_Click()
    If Not mOwner Is Nothing Then mOwner.CloseExecution
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        If Not mOwner Is Nothing Then mOwner.CloseExecution
    Else
        Set mOwner = Nothing
    End If
End Sub

Private Sub UserForm_Activate()
    If mClosing Then Exit Sub
    If mToken <> "" Then
        If Not RefreshDraft() Then Exit Sub
    End If
    modUserFormResizeWin.EnableResizableUserForm Me, True, True
    mLayout.ApplyAnchoredLayout
End Sub

Private Sub UserForm_Layout()
    If mClosing Then Exit Sub
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub
