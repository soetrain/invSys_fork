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
Private mDiscarding As Boolean
Private mDirtyName As Boolean
Private mDirtyTags As Boolean
Private mDirtyInstructions As Boolean
Private mDirtyStepInstruction As Boolean
Private mEditingStepId As String
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
        Array("Label", "lblGuideSource", "", 12, 10, 650, 96, 7), _
        Array("CommandButton", "btnGuideExpectedConclusion", "Expected conclusion", 672, 10, 208, 26, 6), _
        Array("Label", "lblGuideExpectationSummary", "Guide expectation: None", 672, 40, 208, 18, 6), _
        Array("Label", "lblGuideName", "Guide name", 12, 114, 510, 18, 7), _
        Array("TextBox", "txtGuideName", "", 12, 134, 510, 24, 7), _
        Array("Label", "lblGuideTags", "Tags", 540, 114, 340, 18, 6), _
        Array("TextBox", "txtGuideTags", "", 540, 134, 340, 24, 6), _
        Array("Label", "lblGuideInstructions", "Authored instructions for this guide", 12, 168, 868, 18, 7), _
        Array("TextBox", "txtGuideInstructions", "", 12, 188, 868, 64, 7), _
        Array("Label", "lblGuideSteps", "Authored step order", 12, 262, 320, 18, 3), _
        Array("ListBox", "lstGuideSteps", "", 12, 286, 320, 92, 11), _
        Array("Label", "lblGuideEvidence", "Observed controls and original outcomes", 348, 262, 532, 18, 7), _
        Array("TextBox", "txtGuideEvidence", "", 348, 286, 532, 228, 15), _
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
    mCancel.TakeFocusOnClick = False
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
    If Not FlushDirty() Then Exit Function
    mContext = context
    If Not modActionGuideDraft.OpenDraft(context, pathId, mDraftId, notice) Then Invalidate notice: Exit Function
    BindContext = LoadDraft()
End Function

Public Function BindPublishedContext(ByVal context As String, ByVal key As String) As Boolean
    Dim notice As String
    If Not FlushDirty() Then Exit Function
    mContext = context: mPublished = True
    If Not modActionGuideDraft.OpenPublishedDraft(context, key, mDraftId, notice) Then Invalidate notice: Exit Function
    BindPublishedContext = LoadDraft()
End Function

Public Function BindSelectedActions(ByVal context As String, ByVal token As String) As Boolean
    Dim notice As String
    If Not FlushDirty() Then Exit Function
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
    Dim name As String, tags As String, instructions As String, instruction As String, summary As String
    If Not FlushDirty() Then Exit Function
    selected = SelectedStep()
    If Not modActionGuideDraft.ReadEditor(mContext, mDraftId, selected, rows, source, evidence, _
        name, tags, instructions, instruction, summary, notice) Then Invalidate notice: Exit Function
    mLoading = True: mSteps.Clear
    mName.Value = name: mTags.Value = tags: mInstructions.Value = instructions
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
    mStepInstruction.Value = instruction
    mEditingStepId = SelectedStep()
    ClearDirty
    mStepInstruction.Enabled = (SelectedStep() <> "" And mDraftId <> "")
    mExpectationSummary.Caption = summary
    mLoading = False
    LoadDraft = (mDraftId <> "")
End Function

Public Sub RefreshExpectation()
    Dim sequenceId As String, definition As String, summary As String, notice As String
    If Not FlushDirty() Then Exit Sub
    If Not modActionGuideDraft.ReadExpectation(mContext, mDraftId, sequenceId, definition, summary, notice) Then Invalidate notice: Exit Sub
    mExpectationSummary.Caption = summary
End Sub

Private Sub mExpected_Click()
    If Not FlushDirty() Then Exit Sub
    If ValidateBinding() Then modExpectationEditor.OpenForGuide mContext, mDraftId, Me
End Sub

Private Function SelectedStep() As String
    If mSteps.ListIndex >= 0 Then SelectedStep = CStr(mSteps.List(mSteps.ListIndex, 0))
End Function

Private Sub ReadStepInstruction()
    Dim value As String, notice As String, id As String
    If mLoading Then Exit Sub
    If Not FlushDirty() Then Exit Sub
    id = SelectedStep()
    If id <> "" Then
        If Not modActionGuideDraft.ReadText(mContext, mDraftId, "Instruction", id, value, notice) Then Invalidate notice: Exit Sub
    End If
    mLoading = True: mStepInstruction.Value = value: mLoading = False
    mEditingStepId = id
    mStepInstruction.Enabled = (id <> "" And mDraftId <> "")
End Sub

Private Sub ClearDirty()
    mDirtyName = False: mDirtyTags = False
    mDirtyInstructions = False: mDirtyStepInstruction = False
End Sub

Private Function FlushDirtyField(ByVal field As String) As Boolean
    Dim notice As String, stepId As String
    Dim value As String, dirty As Boolean
    If mDiscarding Then Exit Function
    Select Case field
        Case "Name": dirty = mDirtyName: value = CStr(mName.Value)
        Case "Tags": dirty = mDirtyTags: value = CStr(mTags.Value)
        Case "Instructions": dirty = mDirtyInstructions: value = CStr(mInstructions.Value)
        Case "Instruction"
            dirty = mDirtyStepInstruction: value = CStr(mStepInstruction.Value)
            stepId = mEditingStepId
        Case Else: Exit Function
    End Select
    If Not dirty Then FlushDirtyField = True: Exit Function
    If mDraftId = "" Or (field = "Instruction" And stepId = "") Then
        Invalidate "Unavailable: the guide draft changed. Reopen Create guide."
        Exit Function
    End If
    If Not modActionGuideDraft.WriteText(mContext, mDraftId, field, stepId, value, notice) Then
        If notice = "" Then notice = "Unavailable: authored text could not be staged."
        Invalidate notice
        Exit Function
    End If
    Select Case field
        Case "Name": mDirtyName = False
        Case "Tags": mDirtyTags = False
        Case "Instructions": mDirtyInstructions = False
        Case "Instruction": mDirtyStepInstruction = False
    End Select
    mStatus.Caption = notice
    FlushDirtyField = True
End Function

Private Function FlushDirty() As Boolean
    If Not FlushDirtyField("Name") Then Exit Function
    If Not FlushDirtyField("Tags") Then Exit Function
    If Not FlushDirtyField("Instructions") Then Exit Function
    If Not FlushDirtyField("Instruction") Then Exit Function
    FlushDirty = True
End Function

Public Function FlushPendingText() As Boolean
    FlushPendingText = FlushDirty()
End Function

Private Sub ChangeStep(ByVal command As String)
    Dim notice As String
    If mLoading Then Exit Sub
    If Not FlushDirty() Then Exit Sub
    If modActionGuideDraft.EditStep(mContext, mDraftId, SelectedStep(), command, notice) Then
        LoadDraft
    ElseIf ValidateBinding() Then
        mStatus.Caption = notice
    End If
End Sub

Private Sub Invalidate(ByVal notice As String)
    Dim control As Object, name As Variant
    mDiscarding = True
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
    mEditingStepId = "": ClearDirty
End Sub

Public Sub ReleaseDraft()
    mDiscarding = True: ClearDirty
    modExpectationEditor.CloseGuide Me
    modActionGuideDraft.CloseDraft mContext, mDraftId
    mDraftId = "": mContext = "": mEditingStepId = ""
End Sub

Private Sub mName_Change()
    If Not mLoading And Not mDiscarding Then mDirtyName = True
End Sub
Private Sub mTags_Change()
    If Not mLoading And Not mDiscarding Then mDirtyTags = True
End Sub
Private Sub mInstructions_Change()
    If Not mLoading And Not mDiscarding Then mDirtyInstructions = True
End Sub
Private Sub mStepInstruction_Change()
    If Not mLoading And Not mDiscarding Then mDirtyStepInstruction = True
End Sub
Private Sub mName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mDiscarding Then Call FlushDirtyField("Name")
End Sub
Private Sub mName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not mDiscarding And KeyCode = vbKeyTab Then
        If Not FlushDirtyField("Name") Then KeyCode = 0
    End If
End Sub
Private Sub mTags_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mDiscarding Then Call FlushDirtyField("Tags")
End Sub
Private Sub mTags_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not mDiscarding And KeyCode = vbKeyTab Then
        If Not FlushDirtyField("Tags") Then KeyCode = 0
    End If
End Sub
Private Sub mInstructions_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mDiscarding Then Call FlushDirtyField("Instructions")
End Sub
Private Sub mInstructions_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not mDiscarding And KeyCode = vbKeyTab Then
        If Not FlushDirtyField("Instructions") Then KeyCode = 0
    End If
End Sub
Private Sub mStepInstruction_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mDiscarding Then Call FlushDirtyField("Instruction")
End Sub
Private Sub mStepInstruction_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not mDiscarding And KeyCode = vbKeyTab Then
        If Not FlushDirtyField("Instruction") Then KeyCode = 0
    End If
End Sub
Private Sub FlushBeforePointerFocus(ByVal targetField As String)
    If mDiscarding Or mLoading Then Exit Sub
    If targetField <> "Name" Then
        If Not FlushDirtyField("Name") Then Exit Sub
    End If
    If targetField <> "Tags" Then
        If Not FlushDirtyField("Tags") Then Exit Sub
    End If
    If targetField <> "Instructions" Then
        If Not FlushDirtyField("Instructions") Then Exit Sub
    End If
    If targetField <> "Instruction" Then Call FlushDirtyField("Instruction")
End Sub
Private Sub mName_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FlushBeforePointerFocus "Name"
End Sub
Private Sub mTags_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FlushBeforePointerFocus "Tags"
End Sub
Private Sub mInstructions_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FlushBeforePointerFocus "Instructions"
End Sub
Private Sub mStepInstruction_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FlushBeforePointerFocus "Instruction"
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
    mDiscarding = True: ClearDirty
    modGuideEditor.CancelEditor Me
End Sub
Private Sub mSave_Click()
    Dim notice As String
    If mDraftId = "" Then Exit Sub
    If Not FlushDirty() Then Exit Sub
    If modActionGuideDraft.SaveDraft(mContext, mDraftId, notice) Then
        mStatus.Caption = notice
    ElseIf ValidateBinding() Then
        mStatus.Caption = notice
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    mDiscarding = True: ClearDirty
    If CloseMode = 0 Then
        ReleaseDraft
        modGuideEditor.ForgetWindowClosed Me
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
