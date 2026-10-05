VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActionPathRun
   Caption         =   "Run How-To"
   ClientHeight    =   10800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   StartUpPosition =   1
End
Attribute VB_Name = "frmActionPathRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mOwner As frmActionPathLibrary
Private mContext As String
Private mToken As String
Private mWorkbook As String
Private mCapturedWorkbook As Workbook
Private mExecuting As Boolean
Private mClosing As Boolean
Private mCloseRequested As Boolean
Private mLayout As cOperationsAnchorManager
Private WithEvents mEntities As MSForms.ListBox
Private mMode As MSForms.ComboBox
Private mStatus As MSForms.Label
Private mEvaluation As MSForms.TextBox
Private WithEvents mStart As MSForms.CommandButton
Private WithEvents mNext As MSForms.CommandButton
Private WithEvents mStop As MSForms.CommandButton
Private WithEvents mVerify As MSForms.CommandButton
Private WithEvents mClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object
    Me.Caption = "Run How-To": Me.Width = 900: Me.Height = 720
    Set mLayout = modOperationsLayout.OperationsAnchorManager(): mLayout.ConfigureForForm Me, 760, 720
    For Each definition In Array( _
        Array("Label", "lblRunGuide", "", 12, 12, 868, 44, 7), _
        Array("Label", "lblRunProfile", "", 12, 62, 868, 44, 7), _
        Array("Label", "lblRunTarget", "", 12, 112, 868, 36, 7), _
        Array("Label", "lblRunWorkbook", "Receiving will open in the captured Training target.", 12, 154, 868, 32, 7), _
        Array("Label", "lblRunSource", "Training item: exact entity key / item and location", 12, 196, 868, 20, 7), _
        Array("ListBox", "lstRunSourceEntities", "", 12, 222, 868, 140, 7), _
        Array("ComboBox", "cboRunMode", "", 12, 378, 186, 26, 3), _
        Array("CommandButton", "btnStartRun", "Start Run", 222, 378, 130, 28, 3), _
        Array("CommandButton", "btnNextRunStep", "Next Step", 376, 378, 130, 28, 3), _
        Array("CommandButton", "btnStopRun", "Stop", 550, 378, 130, 28, 3), _
        Array("Label", "lblRunStatus", "", 12, 424, 868, 56, 7), _
        Array("TextBox", "txtRunVerification", "", 12, 492, 868, 112, 15), _
        Array("CommandButton", "btnVerifyRun", "Verify run", 12, 648, 160, 28, 9), _
        Array("CommandButton", "btnCloseRun", "Close", 782, 648, 98, 28, 12))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        If definition(0) = "Label" Or definition(0) = "CommandButton" Then control.Caption = definition(2)
        If definition(0) = "Label" Then control.WordWrap = True
        mLayout.RegisterControl control, CLng(definition(7))
    Next definition
    Set mEntities = Me.Controls("lstRunSourceEntities"): mEntities.ColumnCount = 2: mEntities.BoundColumn = 1: mEntities.IntegralHeight = False
    Set mMode = Me.Controls("cboRunMode"): mMode.Style = fmStyleDropDownList: mMode.AddItem "Run all": mMode.AddItem "Step through": mMode.ListIndex = 0
    Set mStart = Me.Controls("btnStartRun"): Set mNext = Me.Controls("btnNextRunStep"): Set mStop = Me.Controls("btnStopRun")
    Set mVerify = Me.Controls("btnVerifyRun"): Set mClose = Me.Controls("btnCloseRun"): Set mStatus = Me.Controls("lblRunStatus")
    Set mEvaluation = Me.Controls("txtRunVerification"): mEvaluation.MultiLine = True: mEvaluation.Locked = True: mEvaluation.ScrollBars = fmScrollBarsVertical
    mStart.Enabled = False: mNext.Enabled = False: mStop.Enabled = False: mVerify.Enabled = False
End Sub

Public Function BindSetup(ByVal context As String, ByVal key As String, ByVal token As String, ByVal snapshot As String, _
                          ByVal guide As String, ByVal profile As String, ByVal target As String, ByVal owner As frmActionPathLibrary, ByRef notice As String) As Boolean
    Dim rows As Variant, index As Long
    mContext = context: mToken = token: Set mOwner = owner
    Me.Controls("lblRunGuide").Caption = guide: Me.Controls("lblRunProfile").Caption = profile: Me.Controls("lblRunTarget").Caption = target
    rows = modReceivingExecution.Choices(snapshot)
    If Not IsEmpty(rows) Then
        For index = 1 To UBound(rows, 1)
            mEntities.AddItem CStr(rows(index, 1))
            mEntities.List(mEntities.ListCount - 1, 1) = CStr(rows(index, 2)) & " | " & CStr(rows(index, 3)) & " | " & CStr(rows(index, 6)) & " | " & CStr(rows(index, 8))
        Next index
    End If
    mStatus.Caption = notice: UpdateState
    BindSetup = True
End Function

Private Sub UpdateState()
    Dim state As String, completed As Long, total As Long, current As Boolean
    state = modExecutionRun.State(mContext, mToken, completed, total)
    current = (mContext <> "" And mContext = modActivity.CaptureContext())
    mStart.Enabled = current And state = "Ready" And mEntities.ListIndex >= 0 And Not mExecuting
    mNext.Enabled = current And state = "Running" And Not mExecuting
    mStop.Enabled = (state = "Running")
    mVerify.Enabled = current And state <> "Ready" And state <> "Running" And state <> "Unavailable" And Not mExecuting
    mEntities.Enabled = (state = "Ready" And Not mExecuting): mMode.Enabled = mEntities.Enabled
    If mWorkbook <> "" Then Me.Controls("lblRunWorkbook").Caption = "Captured Receiving workbook: " & mWorkbook
End Sub

Private Sub mEntities_Change()
    If mToken <> "" Then UpdateState
End Sub

Private Sub mStart_Click()
    Dim mode As String, notice As String, started As Boolean
    On Error GoTo Failed
    If mExecuting Or mEntities.ListIndex < 0 Then Exit Sub
    mode = "RunAll": If mMode.Value = "Step through" Then mode = "StepThrough"
    mExecuting = True: UpdateState
    started = modExecutionRun.Start(mContext, mToken, CStr(mEntities.Value), mode, notice)
    mExecuting = False
    If started Then
        mStatus.Caption = notice
        If mode = "RunAll" Then DispatchSteps True
    Else
        mStatus.Caption = notice
    End If
    UpdateState
    If mCloseRequested And Not mOwner Is Nothing Then mOwner.CloseRunner
    Exit Sub
Failed:
    mExecuting = False
    modExecutionRun.RequestStop mContext, mToken, notice
    mStatus.Caption = notice: UpdateState
End Sub

Private Sub DispatchSteps(ByVal allSteps As Boolean)
    Dim controlId As String, inputs As String, notice As String, delivered As Boolean, completed As Boolean, state As String, count As Long, total As Long
    On Error GoTo Failed
    If mExecuting Then Exit Sub
    mExecuting = True: UpdateState
    Do
        If Not modExecutionRun.BeginStep(mContext, mToken, controlId, inputs, notice) Then Exit Do
        mStatus.Caption = notice: Me.Repaint
        delivered = modReceivingExecution.Dispatch(mContext, controlId, inputs, mCapturedWorkbook, mWorkbook, notice)
        completed = modExecutionRun.FinishStep(mContext, mToken, delivered, notice)
        mStatus.Caption = notice
        state = modExecutionRun.State(mContext, mToken, count, total)
        If Not completed Or state <> "Running" Or Not allSteps Then Exit Do
        DoEvents
    Loop
Done:
    mExecuting = False
    If notice <> "" Then mStatus.Caption = notice
    UpdateState
    If mCloseRequested And Not mOwner Is Nothing Then mOwner.CloseRunner
    Exit Sub
Failed:
    completed = modExecutionRun.FinishStep(mContext, mToken, False, notice)
    modExecutionRun.RequestStop mContext, mToken, notice
    Resume Done
End Sub

Private Sub mNext_Click()
    DispatchSteps False
End Sub

Private Sub mStop_Click()
    Dim notice As String
    modExecutionRun.RequestStop mContext, mToken, notice
    mStatus.Caption = notice: UpdateState
End Sub

Private Sub mVerify_Click()
    Dim text As String, notice As String, evaluated As Boolean
    If mExecuting Then Exit Sub
    evaluated = modExecutionRun.Verify(mContext, mToken, text, notice)
    mEvaluation.Value = text: mStatus.Caption = notice: UpdateState
End Sub

Public Function ReleaseRunner() As Boolean
    Dim notice As String
    mCloseRequested = True
    If mExecuting Then modExecutionRun.RequestStop mContext, mToken, notice: Exit Function
    If Not modExecutionRun.CloseSetup(mContext, mToken) Then Exit Function
    mClosing = True: Set mLayout = Nothing: Set mCapturedWorkbook = Nothing: mContext = "": mToken = ""
    ReleaseRunner = True
End Function

Private Sub mClose_Click()
    If Not mOwner Is Nothing Then mOwner.CloseRunner
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        If Not mOwner Is Nothing Then mOwner.CloseRunner
    Else
        Set mOwner = Nothing
    End If
End Sub

Private Sub UserForm_Activate()
    If mClosing Then Exit Sub
    modUserFormResizeWin.EnableResizableUserForm Me, True, True
    UpdateState: UserForm_Layout
End Sub

Private Sub UserForm_Layout()
    If mClosing Or mLayout Is Nothing Then Exit Sub
    mLayout.ApplyAnchoredLayout
    If Not mEntities Is Nothing Then mEntities.ColumnWidths = CStr(mEntities.Width * 0.45) & " pt;" & CStr(mEntities.Width * 0.55) & " pt"
End Sub
