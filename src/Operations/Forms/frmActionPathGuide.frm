VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActionPathGuide
   Caption         =   "Action Path guide"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   StartUpPosition =   1
End
Attribute VB_Name = "frmActionPathGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mContext As String
Private mPublished As Boolean
Private mDraftId As String
Private mLoading As Boolean
Private mResizeReady As Boolean
Private mLayout As cOperationsAnchorManager
Private mSource As MSForms.Label
Private mStatus As MSForms.Label
Private mEvidence As MSForms.TextBox
Private WithEvents mName As MSForms.TextBox
Private WithEvents mTags As MSForms.TextBox
Private WithEvents mInstructions As MSForms.TextBox
Private WithEvents mStepInstruction As MSForms.TextBox
Private WithEvents mSteps As MSForms.ListBox
Private WithEvents mUp As MSForms.CommandButton
Private WithEvents mDown As MSForms.CommandButton
Private WithEvents mRemove As MSForms.CommandButton
Private WithEvents mCancel As MSForms.CommandButton
Private WithEvents mSave As MSForms.CommandButton
Private WithEvents mExpected As MSForms.CommandButton
Private mExpectationSummary As MSForms.Label

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object
    Me.Caption = "Action Path guide": Me.Width = 900: Me.Height = 650
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 760, 600
    For Each definition In Array( _
        Array("Label", "lblGuideSource", "", 12, 10, 650, 48, 7), _
        Array("CommandButton", "btnGuideExpectedConclusion", "Expected conclusion", 672, 10, 208, 26, 6), _
        Array("Label", "lblGuideExpectationSummary", "Guide expectation: None", 672, 40, 208, 18, 6), _
        Array("Label", "lblGuideName", "Guide name", 12, 66, 510, 18, 7), _
        Array("TextBox", "txtGuideName", "", 12, 86, 510, 24, 7), _
        Array("Label", "lblGuideTags", "Tags", 540, 66, 340, 18, 6), _
        Array("TextBox", "txtGuideTags", "", 540, 86, 340, 24, 6), _
        Array("Label", "lblGuideInstructions", "Authored instructions for this guide", 12, 120, 868, 18, 7), _
        Array("TextBox", "txtGuideInstructions", "", 12, 140, 868, 64, 7), _
        Array("Label", "lblGuideSteps", "Authored step order", 12, 214, 320, 18, 3), _
        Array("ListBox", "lstGuideSteps", "", 12, 238, 320, 140, 11), _
        Array("Label", "lblGuideEvidence", "Observed controls and original outcomes", 348, 214, 532, 18, 7), _
        Array("TextBox", "txtGuideEvidence", "", 348, 238, 532, 276, 15), _
        Array("CommandButton", "btnGuideStepUp", "Move up", 12, 388, 82, 26, 9), _
        Array("CommandButton", "btnGuideStepDown", "Move down", 102, 388, 96, 26, 9), _
        Array("CommandButton", "btnRemoveGuideStep", "Remove selected", 206, 388, 126, 26, 9), _
        Array("Label", "lblGuideStepInstruction", "Authored instruction for selected step", 12, 424, 320, 18, 9), _
        Array("TextBox", "txtGuideStepInstruction", "", 12, 444, 320, 70, 9), _
        Array("Label", "lblGuideStatus", "", 12, 528, 868, 48, 13), _
        Array("Label", "lblGuidePublication", "Save publishes a guide version for permitted Viewers in this warehouse.", 12, 586, 650, 28, 13), _
        Array("CommandButton", "btnSaveGuide", "Save guide", 672, 586, 98, 28, 12), _
        Array("CommandButton", "btnCancelGuide", "Cancel", 782, 586, 98, 28, 12))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        If definition(0) = "Label" Or definition(0) = "CommandButton" Then control.Caption = definition(2)
        If definition(0) = "Label" Then control.WordWrap = True
        mLayout.RegisterControl control, CLng(definition(7))
    Next definition
    Set mSource = Me.Controls("lblGuideSource"): Set mStatus = Me.Controls("lblGuideStatus")
    Set mName = Me.Controls("txtGuideName"): Set mTags = Me.Controls("txtGuideTags")
    Set mInstructions = Me.Controls("txtGuideInstructions"): Set mStepInstruction = Me.Controls("txtGuideStepInstruction")
    Set mSteps = Me.Controls("lstGuideSteps"): Set mEvidence = Me.Controls("txtGuideEvidence")
    Set mUp = Me.Controls("btnGuideStepUp"): Set mDown = Me.Controls("btnGuideStepDown")
    Set mRemove = Me.Controls("btnRemoveGuideStep"): Set mCancel = Me.Controls("btnCancelGuide")
    Set mSave = Me.Controls("btnSaveGuide")
    Set mExpected = Me.Controls("btnGuideExpectedConclusion")
    Set mExpectationSummary = Me.Controls("lblGuideExpectationSummary")
    For Each definition In Array("txtGuideInstructions", "txtGuideStepInstruction", "txtGuideEvidence")
        Set control = Me.Controls(CStr(definition))
        control.MultiLine = True: control.WordWrap = True: control.ScrollBars = fmScrollBarsVertical
    Next definition
    mEvidence.Locked = True
    mSteps.ColumnCount = 2: mSteps.BoundColumn = 1: mSteps.TextColumn = 2
    mSteps.ColumnWidths = "0 pt;300 pt": mSteps.IntegralHeight = False
End Sub

Public Function BindContext(ByVal context As String, ByVal pathId As String) As Boolean
    Dim notice As String
    mContext = context
    If Not modActionGuideDraft.OpenDraft(context, pathId, mDraftId, notice) Then Invalidate notice: Exit Function
    BindContext = LoadDraft()
End Function

Public Function BindPublishedContext(ByVal context As String, ByVal key As String) As Boolean
    Dim notice As String
    mContext = context: mPublished = True
    If Not modActionGuideDraft.OpenPublishedDraft(context, key, mDraftId, notice) Then Invalidate notice: Exit Function
    BindPublishedContext = LoadDraft()
End Function

Public Function BindSelectedActions(ByVal context As String, ByVal token As String) As Boolean
    Dim notice As String
    mContext = context
    If Not modActionGuideDraft.OpenDraft(context, "", mDraftId, notice, token) Then Invalidate notice: Exit Function
    BindSelectedActions = LoadDraft()
End Function

Public Function ValidateBinding() As Boolean
    Dim rows As String, source As String, evidence As String, notice As String
    If mDraftId = "" Then Exit Function
    ValidateBinding = modActionGuideDraft.ReadDraft(mContext, mDraftId, rows, source, evidence, notice)
    If Not ValidateBinding Then Invalidate notice
End Function

Private Function LoadDraft() As Boolean
    Dim rows As String, source As String, evidence As String, notice As String, selected As String, row As Variant, fields As Variant
    Dim field As Variant, value As String, fieldNotice As String
    selected = SelectedStep()
    If Not modActionGuideDraft.ReadDraft(mContext, mDraftId, rows, source, evidence, notice) Then Invalidate notice: Exit Function
    mLoading = True: mSteps.Clear
    For Each field In Array("Name", "Tags", "Instructions")
        If Not modActionGuideDraft.ReadText(mContext, mDraftId, CStr(field), "", value, fieldNotice) Then Invalidate fieldNotice: Exit Function
        Me.Controls("txtGuide" & CStr(field)).Value = value
    Next field
    For Each row In Split(rows, vbCrLf)
        If CStr(row) <> "" Then
            fields = Split(CStr(row), vbTab)
            If UBound(fields) <> 1 Then Invalidate "Unavailable: invalid guide step projection.": Exit Function
            mSteps.AddItem CStr(fields(0)): mSteps.List(mSteps.ListCount - 1, 1) = CStr(fields(1))
            If CStr(fields(0)) = selected Then mSteps.ListIndex = mSteps.ListCount - 1
        End If
    Next row
    If mSteps.ListIndex < 0 And mSteps.ListCount > 0 Then mSteps.ListIndex = 0
    mSource.Caption = source: mEvidence.Value = evidence: mStatus.Caption = notice
    mLoading = False
    ReadStepInstruction
    RefreshExpectation
    LoadDraft = (mDraftId <> "")
End Function

Public Sub RefreshExpectation()
    Dim sequenceId As String, definition As String, summary As String, notice As String
    If Not modActionGuideDraft.ReadExpectation(mContext, mDraftId, sequenceId, definition, summary, notice) Then Invalidate notice: Exit Sub
    mExpectationSummary.Caption = summary
End Sub

Private Sub mExpected_Click()
    If ValidateBinding() Then modExpectationEditor.OpenForGuide mContext, mDraftId, Me
End Sub

Private Function SelectedStep() As String
    If mSteps.ListIndex >= 0 Then SelectedStep = CStr(mSteps.List(mSteps.ListIndex, 0))
End Function

Private Sub ReadStepInstruction()
    Dim value As String, notice As String, id As String
    If mLoading Then Exit Sub
    id = SelectedStep()
    If id <> "" Then
        If Not modActionGuideDraft.ReadText(mContext, mDraftId, "Instruction", id, value, notice) Then Invalidate notice: Exit Sub
    End If
    mLoading = True: mStepInstruction.Value = value: mLoading = False
    mStepInstruction.Enabled = (id <> "" And mDraftId <> "")
End Sub

Private Sub StageText(ByVal field As String, ByVal value As String)
    Dim notice As String, stepId As String
    If mLoading Then Exit Sub
    If field = "Instruction" Then stepId = SelectedStep()
    If Not modActionGuideDraft.WriteText(mContext, mDraftId, field, stepId, value, notice) Then Invalidate notice: Exit Sub
    mStatus.Caption = notice
End Sub

Private Sub ChangeStep(ByVal command As String)
    Dim notice As String
    If mLoading Then Exit Sub
    If modActionGuideDraft.EditStep(mContext, mDraftId, SelectedStep(), command, notice) Then
        LoadDraft
    ElseIf ValidateBinding() Then
        mStatus.Caption = notice
    End If
End Sub

Private Sub Invalidate(ByVal notice As String)
    Dim control As Object, name As Variant
    ReleaseDraft
    mLoading = True
    mSource.Caption = "": mEvidence.Value = "": mSteps.Clear
    For Each name In Array("txtGuideName", "txtGuideTags", "txtGuideInstructions", "txtGuideStepInstruction")
        Set control = Me.Controls(CStr(name))
        control.Value = "": control.Enabled = False
    Next name
    mSteps.Enabled = False: mUp.Enabled = False: mDown.Enabled = False: mRemove.Enabled = False
    mSave.Enabled = False
    mExpected.Enabled = False: mExpectationSummary.Caption = ""
    mStatus.Caption = notice: mLoading = False
End Sub

Public Sub ReleaseDraft()
    modExpectationEditor.CloseGuide Me
    modActionGuideDraft.CloseDraft mContext, mDraftId
    mDraftId = "": mContext = ""
End Sub

Private Sub mName_Change()
    StageText "Name", CStr(mName.Value)
End Sub
Private Sub mTags_Change()
    StageText "Tags", CStr(mTags.Value)
End Sub
Private Sub mInstructions_Change()
    StageText "Instructions", CStr(mInstructions.Value)
End Sub
Private Sub mStepInstruction_Change()
    StageText "Instruction", CStr(mStepInstruction.Value)
End Sub
Private Sub mSteps_Change()
    ReadStepInstruction
End Sub
Private Sub mUp_Click()
    ChangeStep "Up"
End Sub
Private Sub mDown_Click()
    ChangeStep "Down"
End Sub
Private Sub mRemove_Click()
    ChangeStep "Remove"
End Sub
Private Sub mCancel_Click()
    modGuideEditor.CancelEditor Me
End Sub
Private Sub mSave_Click()
    Dim notice As String
    If modActionGuideDraft.SaveDraft(mContext, mDraftId, notice) Then
        mStatus.Caption = notice
    ElseIf ValidateBinding() Then
        mStatus.Caption = notice
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True: modGuideEditor.CancelEditor Me
    Else
        ReleaseDraft
    End If
End Sub

Private Sub UserForm_Activate()
    If Not ValidateBinding() Then Exit Sub
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeReady = True
    End If
    mLayout.ApplyAnchoredLayout
End Sub

Private Sub UserForm_Layout()
    Dim entry As String
    If mContext <> "" Then
        entry = "Create guide": If mPublished Then entry = "Edit guide"
        If mContext <> modActivity.CaptureContext() Then Invalidate "Unavailable: the session or warehouse changed. Reopen " & entry & "."
    End If
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub
