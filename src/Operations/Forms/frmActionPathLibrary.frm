VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActionPathLibrary
   Caption         =   "Published guides"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   StartUpPosition =   1
End
Attribute VB_Name = "frmActionPathLibrary"
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
Private mOwner As frmActionPaths
Private WithEvents mSearch As MSForms.TextBox
Private WithEvents mGuides As MSForms.ListBox
Private WithEvents mRefresh As MSForms.CommandButton
Private WithEvents mClose As MSForms.CommandButton
Private mInstructions As MSForms.TextBox
Private mObservations As MSForms.TextBox
Private mSource As MSForms.Label
Private mStatus As MSForms.Label
Private mObserved As MSForms.Label
Private WithEvents mUse As MSForms.CommandButton
Private mObservedId As String
Private mObservedBinding As String

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object
    Me.Caption = "Published guides": Me.Width = 900: Me.Height = 650
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 760, 600
    For Each definition In Array( _
        Array("Label", "lblPublishedSearch", "Search name, tags or ID", 12, 12, 150, 20, 3), _
        Array("TextBox", "txtGuideSearch", "", 170, 10, 600, 24, 7), _
        Array("CommandButton", "btnRefreshGuides", "Refresh", 782, 10, 98, 26, 6), _
        Array("ListBox", "lstPublishedGuides", "", 12, 46, 868, 116, 7), _
        Array("Label", "lblPublishedGuideSource", "", 12, 172, 868, 82, 7), _
        Array("Label", "lblPublishedAuthored", "Authored instructions - reviewed guide order", 12, 264, 350, 20, 3), _
        Array("Label", "lblPublishedObserved", "Observed controls - original source order", 452, 264, 428, 20, 7), _
        Array("TextBox", "txtPublishedInstructions", "", 12, 292, 426, 238, 11), _
        Array("TextBox", "txtPublishedObservations", "", 452, 292, 428, 238, 15), _
        Array("Label", "lblGuideObservedRun", "", 12, 542, 650, 36, 13), _
        Array("CommandButton", "btnUseGuideForRun", "Use for selected run", 674, 546, 206, 28, 12), _
        Array("Label", "lblPublishedGuideStatus", "", 12, 582, 752, 32, 13), _
        Array("CommandButton", "btnCloseGuides", "Close", 782, 588, 98, 26, 12))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        If definition(0) = "Label" Or definition(0) = "CommandButton" Then control.Caption = definition(2)
        If definition(0) = "Label" Then control.WordWrap = True
        mLayout.RegisterControl control, CLng(definition(7))
    Next definition
    Set mSearch = Me.Controls("txtGuideSearch"): Set mGuides = Me.Controls("lstPublishedGuides")
    Set mRefresh = Me.Controls("btnRefreshGuides"): Set mClose = Me.Controls("btnCloseGuides")
    Set mInstructions = Me.Controls("txtPublishedInstructions"): Set mObservations = Me.Controls("txtPublishedObservations")
    Set mSource = Me.Controls("lblPublishedGuideSource"): Set mStatus = Me.Controls("lblPublishedGuideStatus")
    Set mObserved = Me.Controls("lblGuideObservedRun"): Set mUse = Me.Controls("btnUseGuideForRun")
    mUse.Enabled = False
    mGuides.ColumnCount = 3: mGuides.BoundColumn = 1
    mGuides.ColumnWidths = "0 pt;400 pt;420 pt": mGuides.IntegralHeight = False
    For Each definition In Array("txtPublishedInstructions", "txtPublishedObservations")
        Set control = Me.Controls(CStr(definition))
        control.MultiLine = True: control.WordWrap = True: control.Locked = True: control.ScrollBars = fmScrollBarsVertical
    Next definition
End Sub

Public Sub BindObservedRun(ByVal pathId As String, ByVal binding As String)
    mObservedId = "": mObservedBinding = "": mObserved.Caption = "": mUse.Enabled = False
    If Not ContextValid() Then Exit Sub
    mObserved.Caption = modPathEvaluation.SelectedRunCaption(mContext, pathId, binding)
    If mObserved.Caption <> "" Then
        mObservedId = pathId: mObservedBinding = binding
    Else
        mObserved.Caption = "Select a recording in Action Paths, then reopen Published guides to use an expectation."
    End If
    ValidateSelection
End Sub

Public Sub BindContext(ByVal context As String, ByVal owner As frmActionPaths)
    mContext = context: Set mOwner = owner
    RefreshGuides
End Sub

Private Function ContextValid() As Boolean
    ContextValid = (mContext <> "" And mContext = modActivity.CaptureContext())
    If Not ContextValid Then
        mObservedId = "": mObservedBinding = "": mObserved.Caption = ""
        mLoading = True: mGuides.Clear: mLoading = False
        ClearSelection "Unavailable: the invSys session or warehouse changed. Reopen Viewer."
    End If
End Function

Public Sub ValidateSelection()
    Dim instructions As String, observations As String, provenance As String, notice As String, key As String
    If mLoading Then Exit Sub
    If Not ContextValid() Then Exit Sub
    If mGuides.ListIndex < 0 Then ClearSelection "Select a published guide version.": Exit Sub
    key = CStr(mGuides.Value)
    If Not modGuideLibraryRead.ReadGuide(mContext, key, instructions, observations, provenance, notice) Then
        ClearSelection notice
        Exit Sub
    End If
    If Not ContextValid() Then Exit Sub
    If CStr(mGuides.Value) <> key Then Exit Sub
    mInstructions.Value = instructions: mObservations.Value = observations
    mSource.Caption = provenance: mStatus.Caption = notice
    mUse.Enabled = False
    If Not mOwner Is Nothing And mObservedId <> "" Then
        mUse.Enabled = mOwner.ObservedRunMatches(mObservedId, mObservedBinding)
        If Not mUse.Enabled Then mStatus.Caption = "The selected recording changed. Reopen Published guides from the intended recording."
    End If
End Sub

Private Sub RefreshGuides()
    Dim rows As String, notice As String, row As Variant, fields As Variant, selected As String, index As Long
    On Error GoTo Failed
    If mLoading Then Exit Sub
    If Not ContextValid() Then Exit Sub
    If mGuides.ListIndex >= 0 Then selected = CStr(mGuides.Value)
    mLoading = True: mGuides.Clear
    ClearSelection "Loading published guides."
    If Not modGuideLibraryRead.ListGuides(mContext, CStr(mSearch.Value), rows, notice) Then GoTo Done
    For Each row In Split(rows, vbCrLf)
        If CStr(row) <> "" Then
            fields = Split(CStr(row), vbTab)
            If UBound(fields) <> 2 Then GoTo Failed
            mGuides.AddItem CStr(fields(0))
            For index = 1 To 2: mGuides.List(mGuides.ListCount - 1, index) = CStr(fields(index)): Next index
            If CStr(fields(0)) = selected Then mGuides.ListIndex = mGuides.ListCount - 1
        End If
    Next row
Done:
    mLoading = False: mStatus.Caption = notice
    If mGuides.ListIndex >= 0 Then ValidateSelection
    Exit Sub
Failed:
    mLoading = False
    ClearSelection "Unavailable: the published-guide library could not be displayed."
End Sub

Private Sub ClearSelection(ByVal notice As String)
    mUse.Enabled = False
    mInstructions.Value = "": mObservations.Value = "": mSource.Caption = "": mStatus.Caption = notice
End Sub

Public Sub ReleaseReader()
    mObservedId = "": mObservedBinding = "": mObserved.Caption = ""
    Set mOwner = Nothing: mContext = ""
    mLoading = True: mGuides.Clear: mSearch.Value = "": mLoading = False
    ClearSelection ""
End Sub

Private Sub mUse_Click()
    Dim key As String, notice As String, applied As Boolean
    If Not ContextValid() Then Exit Sub
    If mOwner Is Nothing Or mGuides.ListIndex < 0 Then Exit Sub
    key = CStr(mGuides.Value)
    applied = mOwner.UseGuideForRun(mObservedId, mObservedBinding, key, notice)
    If Not ContextValid() Then Exit Sub
    mStatus.Caption = notice
End Sub

Private Sub mSearch_Change()
    mRefresh_Click
End Sub

Private Sub mRefresh_Click()
    RefreshGuides
End Sub

Private Sub mGuides_Change()
    ValidateSelection
End Sub

Private Sub mClose_Click()
    If Not mOwner Is Nothing Then mOwner.ClosePublishedGuides
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        If Not mOwner Is Nothing Then mOwner.ClosePublishedGuides
    End If
End Sub

Private Sub UserForm_Activate()
    If Not ContextValid() Then Exit Sub
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeReady = True
    End If
    mLayout.ApplyAnchoredLayout
    ValidateSelection
End Sub

Private Sub UserForm_Layout()
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub
