VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActionPaths
   Caption         =   "Action Paths"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   StartUpPosition =   1
End
Attribute VB_Name = "frmActionPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mContext As String
Private mLoading As Boolean
Private mResizeReady As Boolean
Private mLayout As cOperationsAnchorManager
Private WithEvents mSearch As MSForms.TextBox
Private WithEvents mPaths As MSForms.ListBox
Private WithEvents mRefresh As MSForms.CommandButton
Private WithEvents mClose As MSForms.CommandButton
Private mEvidence As MSForms.TextBox
Private mStatus As MSForms.Label
Private mSummary As MSForms.Label
Private WithEvents mExpected As MSForms.CommandButton
Private mSelectedId As String

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object
    Me.Width = 820: Me.Height = 640
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 720, 520
    For Each definition In Array( _
        Array("Label", "lblPathSearch", "Search recordings", 12, 12, 130, 20, 3), _
        Array("TextBox", "txtPathSearch", "", 148, 10, 550, 24, 7), _
        Array("CommandButton", "btnPathRefresh", "Refresh", 708, 10, 92, 26, 6), _
        Array("Label", "lblPathList", "Saved recordings - select one to validate its evidence", 12, 44, 788, 20, 7), _
        Array("ListBox", "lstActionPaths", "", 12, 68, 788, 130, 7), _
        Array("Label", "lblPathEvidence", "Observed controls and outcomes", 12, 208, 600, 20, 7), _
        Array("CommandButton", "btnExpectedConclusion", "Expected conclusion", 630, 204, 170, 26, 6), _
        Array("TextBox", "txtPathEvidence", "", 12, 232, 788, 284, 15), _
        Array("Label", "lblExpectationSummary", "", 12, 528, 788, 24, 13), _
        Array("Label", "lblPathStatus", "Select a recording to inspect its evidence.", 12, 560, 680, 40, 13), _
        Array("CommandButton", "btnClose", "Close", 714, 564, 86, 28, 12))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        If definition(0) = "Label" Or definition(0) = "CommandButton" Then control.Caption = definition(2)
        mLayout.RegisterControl control, CLng(definition(7))
    Next definition
    Set mSearch = Me.Controls("txtPathSearch"): Set mRefresh = Me.Controls("btnPathRefresh")
    Set mPaths = Me.Controls("lstActionPaths"): Set mEvidence = Me.Controls("txtPathEvidence")
    Set mStatus = Me.Controls("lblPathStatus"): Set mClose = Me.Controls("btnClose")
    Set mSummary = Me.Controls("lblExpectationSummary"): Set mExpected = Me.Controls("btnExpectedConclusion")
    mExpected.Enabled = False: mSummary.WordWrap = True
    mPaths.ColumnCount = 3: mPaths.ColumnWidths = "0 pt;340 pt;300 pt": mPaths.IntegralHeight = False
    mEvidence.MultiLine = True: mEvidence.WordWrap = True: mEvidence.Locked = True
    mEvidence.ScrollBars = fmScrollBarsVertical
    mStatus.WordWrap = True
End Sub

Public Sub BindContext(ByVal context As String)
    If mContext <> context Then ClearContent "Loading recording library."
    mContext = context
    RefreshPaths
End Sub

Private Function ContextValid() As Boolean
    ContextValid = (mContext <> "" And mContext = modActivity.CaptureContext())
    If Not ContextValid Then ClearContent "Unavailable: the invSys session or warehouse changed. Reopen Viewer."
End Function

Private Sub RefreshPaths()
    Dim rows As String, notice As String, row As Variant, fields As Variant, selected As String, index As Long
    On Error GoTo Failed
    If mLoading Then Exit Sub
    If Not ContextValid() Then Exit Sub
    If mPaths.ListIndex >= 0 Then selected = CStr(mPaths.List(mPaths.ListIndex, 0))
    mLoading = True: mPaths.Clear: mEvidence.Value = ""
    If Not modActionPathRead.ListPaths(mContext, CStr(mSearch.Value), rows, notice) Then GoTo Done
    For Each row In Split(rows, vbCrLf)
        If CStr(row) <> "" Then
            fields = Split(CStr(row), vbTab)
            If UBound(fields) <> 2 Then GoTo Failed
            mPaths.AddItem CStr(fields(0))
            For index = 1 To 2: mPaths.List(mPaths.ListCount - 1, index) = CStr(fields(index)): Next index
            If CStr(fields(0)) = selected Then mPaths.ListIndex = mPaths.ListCount - 1
        End If
    Next row
Done:
    mLoading = False: mStatus.Caption = notice
    ReadSelection
    Exit Sub
Failed:
    mLoading = False
    ClearContent "Unavailable: the recording library could not be displayed."
End Sub

Private Sub ReadSelection()
    Dim evidence As String, notice As String, succeeded As Boolean, selected As String
    If mLoading Then Exit Sub
    mEvidence.Value = ""
    If Not ContextValid() Then Exit Sub
    If mPaths.ListIndex >= 0 Then selected = CStr(mPaths.List(mPaths.ListIndex, 0))
    If mSelectedId <> selected Then
        modExpectationEditor.CloseLibrary Me
        modActionPathRead.ClearSelection mContext
    End If
    mSelectedId = selected: mExpected.Enabled = False: mSummary.Caption = ""
    If selected = "" Then Exit Sub
    succeeded = modActionPathRead.ReadPath(mContext, selected, evidence, notice)
    mStatus.Caption = notice
    If succeeded Then
        mEvidence.Value = evidence: mExpected.Enabled = True
        RefreshExpectation
    Else
        modExpectationEditor.CloseLibrary Me
    End If
End Sub

Public Sub RefreshExpectation()
    If Not ContextValid() Or mSelectedId = "" Then Exit Sub
    mSummary.Caption = modActionPathRead.ExpectationSummary(mContext, mSelectedId)
End Sub

Private Sub mExpected_Click()
    If Not ContextValid() Or mSelectedId = "" Then Exit Sub
    modExpectationEditor.OpenForRun mContext, mSelectedId, Me
End Sub

Public Sub ClearContent(ByVal notice As String)
    modExpectationEditor.CloseLibrary Me
    modActionPathRead.ClearSelection mContext
    mLoading = True
    mPaths.Clear: mEvidence.Value = "": mStatus.Caption = notice
    mSelectedId = "": mSummary.Caption = "": mExpected.Enabled = False
    mLoading = False
End Sub

Private Sub mPaths_Change()
    ReadSelection
End Sub

Private Sub mSearch_Change()
    RefreshPaths
End Sub

Private Sub mRefresh_Click()
    RefreshPaths
End Sub

Private Sub mClose_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True: Me.Hide
    Else
        ClearContent ""
    End If
End Sub

Private Sub UserForm_Activate()
    If Not ContextValid() Then Exit Sub
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeReady = True
    End If
    mLayout.ApplyAnchoredLayout
    ReadSelection
End Sub

Private Sub UserForm_Layout()
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub
