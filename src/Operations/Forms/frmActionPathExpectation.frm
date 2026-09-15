VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActionPathExpectation
   Caption         =   "Expected steps and conclusion"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   StartUpPosition =   1
End
Attribute VB_Name = "frmActionPathExpectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mContext As String
Private mDraftId As String
Private mSequence As String
Private mLoading As Boolean
Private mResizeReady As Boolean
Private mLayout As cOperationsAnchorManager
Private WithEvents mControl As MSForms.ComboBox
Private mOutcome As MSForms.ComboBox
Private mRetry As MSForms.CheckBox
Private mSteps As MSForms.ListBox
Private mTerminal As MSForms.ComboBox
Private mKind As MSForms.ComboBox
Private mStatus As MSForms.Label
Private mEditBindings As Collection
Private WithEvents mUse As MSForms.CommandButton
Private WithEvents mCancel As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object, binding As cExpectationButton
    Me.Caption = "Expected steps and conclusion"
    Me.Width = 900: Me.Height = 630
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 760, 550
    For Each definition In Array( _
        Array("Label", "lblExpectationHelp", "Choose the controls and outcomes expected for this recording. These steps describe intent; they do not perform the work.", 12, 12, 868, 36, 7), _
        Array("Label", "lblExpectedControl", "Control to add", 12, 56, 660, 18, 7), _
        Array("ComboBox", "cboExpectedControl", "", 12, 76, 650, 26, 7), _
        Array("CommandButton", "btnAddExpectedStep", "Add expected step", 676, 76, 204, 26, 6), _
        Array("Label", "lblExpectedOutcome", "Required outcome", 12, 110, 280, 18, 3), _
        Array("ComboBox", "cboExpectedOutcome", "", 12, 130, 280, 26, 3), _
        Array("CheckBox", "chkExpectedRetry", "Allow retries for added step", 308, 130, 320, 26, 3), _
        Array("Label", "lblExpectedSteps", "Expected steps in order (maximum 256)", 12, 168, 868, 18, 7), _
        Array("ListBox", "lstExpectedSteps", "", 12, 188, 868, 180, 15), _
        Array("CommandButton", "btnRemoveExpectedStep", "Remove selected", 12, 380, 140, 26, 9), _
        Array("CommandButton", "btnExpectedStepUp", "Move up", 164, 380, 80, 26, 9), _
        Array("CommandButton", "btnExpectedStepDown", "Move down", 256, 380, 100, 26, 9), _
        Array("Label", "lblTerminalStep", "Conclusion step", 12, 420, 440, 18, 13), _
        Array("Label", "lblTerminalKind", "Conclusion to check", 464, 420, 416, 18, 12), _
        Array("ComboBox", "cboTerminalStep", "", 12, 440, 440, 26, 13), _
        Array("ComboBox", "cboTerminalKind", "", 464, 440, 416, 26, 12), _
        Array("Label", "lblExpectationStatus", "", 12, 480, 868, 48, 13), _
        Array("CommandButton", "btnUseExpectation", "Use for this recording", 540, 550, 230, 28, 12), _
        Array("CommandButton", "btnCancelExpectation", "Cancel", 782, 550, 98, 28, 12))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        If definition(0) = "Label" Or definition(0) = "CommandButton" Or definition(0) = "CheckBox" Then control.Caption = definition(2)
        If definition(0) = "Label" Then control.WordWrap = True
        If definition(0) = "ComboBox" Then
            control.Style = fmStyleDropDownList: control.ColumnCount = 2
            control.BoundColumn = 1: control.TextColumn = 2
            control.ColumnWidths = "0 pt;600 pt"
        End If
        mLayout.RegisterControl control, CLng(definition(7))
    Next definition
    Set mControl = Me.Controls("cboExpectedControl"): Set mOutcome = Me.Controls("cboExpectedOutcome")
    Set mRetry = Me.Controls("chkExpectedRetry"): Set mSteps = Me.Controls("lstExpectedSteps")
    Set mTerminal = Me.Controls("cboTerminalStep"): Set mKind = Me.Controls("cboTerminalKind")
    Set mStatus = Me.Controls("lblExpectationStatus")
    Set mEditBindings = New Collection
    For Each definition In Array(Array("Add", "btnAddExpectedStep"), Array("Remove", "btnRemoveExpectedStep"), _
                                 Array("Up", "btnExpectedStepUp"), Array("Down", "btnExpectedStepDown"))
        Set binding = New cExpectationButton
        binding.Connect Me.Controls(CStr(definition(1))), Me, CStr(definition(0))
        mEditBindings.Add binding
    Next definition
    Set mUse = Me.Controls("btnUseExpectation"): Set mCancel = Me.Controls("btnCancelExpectation")
    mSteps.ColumnCount = 5: mSteps.BoundColumn = 1: mSteps.TextColumn = 2: mSteps.IntegralHeight = False
    FillChoices mKind, "None" & vbTab & "None" & vbCrLf & "CommandCompleted" & vbTab & "Command completed" & vbCrLf & _
        "SourceEventsApplied" & vbTab & "Source events applied"
    ArrangeColumns
End Sub

Public Function BindContext(ByVal context As String, Optional ByVal pathId As String = "", Optional ByVal guideDraftId As String = "") As Boolean
    Dim projection As String, notice As String, scope As String
    ReleaseDraft
    mContext = context
    If Not modActionRecording.OpenExpectation(context, projection, notice, pathId, guideDraftId) Then mStatus.Caption = notice: Exit Function
    scope = IIf(pathId = "", "recording", "evaluation")
    If guideDraftId <> "" Then scope = "guide"
    mUse.Caption = "Use for this " & scope
    Me.Controls("lblExpectationHelp").Caption = "Choose the controls and outcomes expected for this " & _
        scope & ". These steps describe intent; they do not perform the work."
    LoadProjection projection, False
    mLoading = True
    FillChoices mControl, modActionRecording.ExpectationChoices(mContext, mDraftId, "")
    mRetry.Value = True
    mLoading = False
    mStatus.Caption = notice: mUse.Enabled = True
    BindContext = (mDraftId <> "")
End Function

Private Sub FillChoices(ByVal control As MSForms.ComboBox, ByVal rows As String)
    Dim row As Variant, fields As Variant
    control.Clear
    For Each row In Split(rows, vbCrLf)
        If CStr(row) <> "" Then
            fields = Split(CStr(row), vbTab)
            If UBound(fields) <> 1 Then Err.Raise 5, , "Invalid expectation choices."
            control.AddItem CStr(fields(0)): control.List(control.ListCount - 1, 1) = CStr(fields(1))
        End If
    Next row
End Sub

Private Function SelectedValue(ByVal control As Object) As String
    If control.ListIndex >= 0 Then SelectedValue = CStr(control.List(control.ListIndex, 0))
End Function

Private Sub LoadProjection(ByVal projection As String, ByVal preserveSelection As Boolean)
    Dim rows As Variant, fields As Variant, index As Long, selected As String, terminal As String, kind As String
    rows = Split(projection, vbCrLf): fields = Split(CStr(rows(0)), vbTab)
    If UBound(fields) <> 4 Then Err.Raise 5, , "Invalid expectation projection."
    If fields(0) <> "EXPECTATION1" Then Err.Raise 5, , "Unsupported expectation projection."
    If preserveSelection Then
        If fields(1) <> mDraftId Or fields(2) <> mSequence Then Err.Raise 5, , "Expectation binding changed."
        selected = SelectedValue(mSteps): terminal = SelectedValue(mTerminal): kind = SelectedValue(mKind)
    Else
        terminal = CStr(fields(3)): kind = CStr(fields(4))
    End If
    mDraftId = CStr(fields(1)): mSequence = CStr(fields(2))
    mLoading = True: mSteps.Clear: mTerminal.Clear
    For index = 1 To UBound(rows)
        fields = Split(CStr(rows(index)), vbTab)
        If UBound(fields) <> 4 Then Err.Raise 5, , "Invalid expectation step."
        mSteps.AddItem CStr(fields(0))
        mSteps.List(mSteps.ListCount - 1, 1) = CStr(index) & ". " & CStr(fields(4))
        mSteps.List(mSteps.ListCount - 1, 2) = CStr(fields(2))
        mSteps.List(mSteps.ListCount - 1, 3) = IIf(fields(3) = "True", "Retries allowed", "No retries")
        mSteps.List(mSteps.ListCount - 1, 4) = CStr(fields(1))
        If fields(0) = selected Then mSteps.ListIndex = mSteps.ListCount - 1
        mTerminal.AddItem CStr(fields(0))
        mTerminal.List(mTerminal.ListCount - 1, 1) = CStr(index) & ". " & CStr(fields(4))
        If fields(0) = terminal Then mTerminal.ListIndex = mTerminal.ListCount - 1
    Next index
    If kind = "" Then kind = "None"
    mKind.Value = kind: mLoading = False
End Sub

Public Function ValidateBinding() As Boolean
    Dim projection As String, notice As String
    ValidateBinding = modActionRecording.EditExpectation(mContext, mDraftId, "Read", "", "", "", False, projection, notice)
    If Not ValidateBinding Then ReleaseDraft: mStatus.Caption = notice
End Function

Public Sub ChangeDraft(ByVal command As String)
    Dim projection As String, notice As String, valid As Boolean
    If mLoading Then Exit Sub
    If modActionRecording.EditExpectation(mContext, mDraftId, command, SelectedValue(mSteps), _
        SelectedValue(mControl), SelectedValue(mOutcome), CBool(mRetry.Value), projection, notice) Then
        LoadProjection projection, True
        If command = "Add" Then mSteps.ListIndex = mSteps.ListCount - 1
    Else
        valid = ValidateBinding()
    End If
    mStatus.Caption = notice
End Sub

Private Sub mControl_Change()
    If mLoading Then Exit Sub
    If Not ValidateBinding() Then Exit Sub
    FillChoices mOutcome, modActionRecording.ExpectationChoices(mContext, mDraftId, SelectedValue(mControl))
End Sub

Private Sub mUse_Click()
    Dim notice As String, valid As Boolean
    If modActionRecording.UseExpectation(mContext, mDraftId, SelectedValue(mTerminal), SelectedValue(mKind), notice) Then
        modExpectationEditor.Staged
        ReleaseDraft: Me.Hide
    Else
        valid = ValidateBinding()
    End If
    mStatus.Caption = notice
End Sub

Private Sub mCancel_Click()
    ReleaseDraft: Me.Hide
End Sub

Public Sub ReleaseDraft()
    modActionRecording.CloseExpectation mContext, mDraftId
    mLoading = True
    mDraftId = "": mSequence = "": mContext = ""
    mSteps.Clear: mTerminal.Clear: mControl.Clear: mOutcome.Clear
    mRetry.Value = False: mKind.ListIndex = 0: mUse.Enabled = False
    mLoading = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim binding As cExpectationButton
    If CloseMode = 0 Then
        Cancel = True: ReleaseDraft: Me.Hide
    Else
        If Not mEditBindings Is Nothing Then
            For Each binding In mEditBindings: binding.Disconnect: Next binding
        End If
        Set mEditBindings = Nothing: Set mLayout = Nothing
    End If
End Sub

Private Sub UserForm_Activate()
    If Not ValidateBinding() Then Exit Sub
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeReady = True
    End If
    mLayout.ApplyAnchoredLayout
    ArrangeColumns
End Sub

Private Sub ArrangeColumns()
    Dim control As Variant
    If mSteps Is Nothing Then Exit Sub
    mSteps.ColumnWidths = "0 pt;" & CStr(mSteps.Width - 220) & " pt;110 pt;90 pt;0 pt"
    For Each control In Array(mControl, mOutcome, mTerminal, mKind)
        control.ColumnWidths = "0 pt;" & CStr(control.Width - 18) & " pt"
    Next control
End Sub

Private Sub UserForm_Layout()
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
    ArrangeColumns
End Sub
