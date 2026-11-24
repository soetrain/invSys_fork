VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGuideActionPicker
   Caption         =   "Choose tracked actions"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   StartUpPosition =   1
End
Attribute VB_Name = "frmGuideActionPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mContext As String, mToken As String
Private mOwner As frmActionPaths
Private mLoading As Boolean, mResizeReady As Boolean
Private mRenderedSelected As Object
Private mSelectedIds As Object
Private mLoadedRows As String, mLoadedSource As String, mLoadedNotice As String
Private mSnapshotReady As Boolean
Private mLayout As cOperationsAnchorManager
Private mSource As MSForms.Label, mStatus As MSForms.Label
Private WithEvents mSearch As MSForms.TextBox
Private WithEvents mActions As MSForms.ListBox
Private WithEvents mCreate As MSForms.CommandButton
Private WithEvents mCancel As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object
    Me.Caption = "Choose tracked actions": Me.Width = 900: Me.Height = 650
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 760, 600
    For Each definition In Array( _
        Array("Label", "lblGuideActionSource", "", 12, 10, 868, 60, 7), _
        Array("Label", "lblGuideActionSearch", "Search actions", 12, 80, 124, 18, 3), _
        Array("TextBox", "txtGuideActionSearch", "", 148, 76, 732, 24, 7), _
        Array("ListBox", "lstGuideActions", "", 12, 112, 868, 310, 15), _
        Array("Label", "lblGuideActionStatus", "", 12, 436, 868, 100, 13), _
        Array("Label", "lblGuideActionHint", "Select original actions, then review their authored order in the guide editor.", 12, 550, 868, 20, 13), _
        Array("CommandButton", "btnCreateSelectedGuide", "Create guide", 636, 586, 136, 28, 12), _
        Array("CommandButton", "btnCancelGuideActions", "Cancel", 784, 586, 96, 28, 12))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        If definition(0) = "Label" Or definition(0) = "CommandButton" Then control.Caption = definition(2)
        If definition(0) = "Label" Then control.WordWrap = True
        mLayout.RegisterControl control, CLng(definition(7))
    Next definition
    Set mSource = Me.Controls("lblGuideActionSource"): Set mStatus = Me.Controls("lblGuideActionStatus")
    Set mSearch = Me.Controls("txtGuideActionSearch"): Set mActions = Me.Controls("lstGuideActions")
    Set mCreate = Me.Controls("btnCreateSelectedGuide"): Set mCancel = Me.Controls("btnCancelGuideActions")
    mActions.ColumnCount = 4: mActions.ColumnWidths = "0 pt;280 pt;180 pt;240 pt"
    mActions.IntegralHeight = False: mActions.MultiSelect = fmMultiSelectMulti
    mCreate.Enabled = False
End Sub

Public Function BindContext(ByVal context As String, ByVal owner As frmActionPaths) As Boolean
    Dim notice As String
    mContext = context: Set mOwner = owner
    If Not modGuideActionPicker.OpenSelection(context, mToken, notice) Then Invalidate notice: Exit Function
    BindContext = RefreshChoices()
End Function

Public Function RefreshChoices() As Boolean
    Dim rows As String, source As String, notice As String, count As Long, row As Variant, fields As Variant
    Dim selectedIds As Object, id As String
    On Error GoTo Failed
    If mLoading Or mToken = "" Then Exit Function
    If Not modGuideActionPicker.ReadRows(mContext, mToken, "", rows, source, count, notice) Then Invalidate notice: Exit Function
    Set selectedIds = CreateObject("Scripting.Dictionary")
    For Each row In Split(rows, vbCrLf)
        If CStr(row) <> "" Then
            fields = Split(CStr(row), vbTab)
            If UBound(fields) <> 5 Then GoTo Failed
            id = CStr(fields(0))
            If fields(4) = "1" Then selectedIds.Add id, True
        End If
    Next row
    If selectedIds.Count <> count Then GoTo Failed
    mLoadedRows = rows: mLoadedSource = source: mLoadedNotice = notice
    Set mSelectedIds = selectedIds: mSnapshotReady = True
    RefreshChoices = FilterLoadedChoices()
    Exit Function
Failed:
    Invalidate "Unavailable: the selected published actions could not be displayed."
End Function

Private Function FilterLoadedChoices() As Boolean
    Dim row As Variant, fields As Variant, id As String, query As String, column As Long
    Dim rendered As Object
    On Error GoTo Failed
    If Not mSnapshotReady Or mToken = "" Or mContext = "" Or mContext <> modActivity.CaptureContext() Then
        Invalidate "Unavailable: the invSys session or warehouse changed. Reopen Choose tracked actions."
        Exit Function
    End If
    Set rendered = CreateObject("Scripting.Dictionary")
    query = CStr(mSearch.Value)
    mLoading = True: mActions.Clear
    For Each row In Split(mLoadedRows, vbCrLf)
        If CStr(row) <> "" Then
            fields = Split(CStr(row), vbTab)
            If UBound(fields) <> 5 Then GoTo Failed
            id = CStr(fields(0))
            If query = "" Or InStr(1, CStr(fields(1)) & " " & CStr(fields(5)), query, vbTextCompare) > 0 Then
                If rendered.Exists(id) Then GoTo Failed
                rendered.Add id, mSelectedIds.Exists(id)
                mActions.AddItem id
                For column = 1 To 3: mActions.List(mActions.ListCount - 1, column) = CStr(fields(column)): Next column
                mActions.Selected(mActions.ListCount - 1) = mSelectedIds.Exists(id)
            End If
        End If
    Next row
    Set mRenderedSelected = rendered
    mActions.Enabled = True
    mSource.Caption = mLoadedSource: mStatus.Caption = mLoadedNotice
    mCreate.Enabled = (mSelectedIds.Count > 0)
    mLoading = False: FilterLoadedChoices = True
    Exit Function
Failed:
    Invalidate "Unavailable: the selected published actions could not be displayed."
End Function

Private Sub mActions_Change()
    Dim index As Long, notice As String, id As String, selected As Boolean, changed As Boolean
    If mLoading Or mToken = "" Then Exit Sub
    If mRenderedSelected Is Nothing Then Invalidate "Unavailable: the selected action list changed. Reopen Choose tracked actions.": Exit Sub
    mLoading = True
    For index = 0 To mActions.ListCount - 1
        id = CStr(mActions.List(index, 0)): selected = mActions.Selected(index)
        If Not mRenderedSelected.Exists(id) Then Invalidate "Unavailable: the selected action list changed. Reopen Choose tracked actions.": Exit Sub
        If selected <> CBool(mRenderedSelected(id)) Then
            If Not modGuideActionPicker.Choose(mContext, mToken, id, selected, notice) Then Invalidate notice: Exit Sub
            mRenderedSelected(id) = selected: changed = True
        End If
    Next index
    If changed Then modGuideEditor.CloseActionPicker Me
    mLoading = False
    If changed Then RefreshChoices
End Sub

Private Sub mSearch_Change()
    If Not mLoading Then FilterLoadedChoices
End Sub

Private Sub mCreate_Click()
    If RefreshChoices() And mCreate.Enabled Then modGuideEditor.OpenForSelectedActions mContext, mToken, Me
End Sub

Private Sub mCancel_Click()
    Unload Me
End Sub

Private Sub Invalidate(ByVal notice As String)
    ReleaseSelection
    mLoading = True: mActions.Clear: mActions.Enabled = False: mSearch.Value = "": mSearch.Enabled = False
    mSource.Caption = "": mStatus.Caption = notice: mCreate.Enabled = False: mLoading = False
End Sub

Private Sub ReleaseSelection()
    modGuideEditor.CloseActionPicker Me
    modGuideActionPicker.CloseSelection mContext, mToken
    Set mRenderedSelected = Nothing
    Set mSelectedIds = Nothing
    mLoadedRows = "": mLoadedSource = "": mLoadedNotice = "": mSnapshotReady = False
    mToken = "": mContext = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        If modGuideReaderClose.SchedulePicker(Me) Then
            Me.Hide
        Else
            mStatus.Caption = "Unable to close Choose tracked actions. Use Cancel or try again."
        End If
        Exit Sub
    End If
    ReleaseSelection
    If modOperationsFormLifetime.IsLoaded(mOwner) Then mOwner.ForgetClosedGuideActionPicker Me
    Set mOwner = Nothing
End Sub

Private Sub UserForm_Activate()
    If Not RefreshChoices() Then Exit Sub
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True: mResizeReady = True
    End If
    mLayout.ApplyAnchoredLayout
End Sub

Private Sub UserForm_Layout()
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub
