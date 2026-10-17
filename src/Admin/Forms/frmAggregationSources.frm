VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAggregationSources
   Caption         =   "invSys Aggregation Sources"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11100
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAggregationSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

' D17 source selection is session-only. This form connects and discovers
' published snapshots without changing the operational warehouse context.
Private WithEvents mTxtRoot As MSForms.TextBox
Private WithEvents mTxtUser As MSForms.TextBox
Private WithEvents mTxtPassword As MSForms.TextBox
Private WithEvents mLstSources As MSForms.ListBox
Private WithEvents mLstRejected As MSForms.ListBox
Private WithEvents mBtnDiscover As MSForms.CommandButton
Private WithEvents mBtnConnect As MSForms.CommandButton
Private WithEvents mBtnAddServer As MSForms.CommandButton
Private WithEvents mBtnAggregate As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private mLblSelected As MSForms.Label
Private mLblStatus As MSForms.Label
Private mBuilt As Boolean
Private mResizeInitialized As Boolean
Private mSourceContext As String
Private mSourceWorkbook As Workbook
Private mBusy As Boolean
Private mLoading As Boolean
Private mActivity As cProvisioningActivity

' Only the actual Admin callback starts an observation; loading/smoke does not.
Public Sub ObserveEntry()
    If Not mActivity Is Nothing Or Me.Visible Then Exit Sub
    If Not modAggregationCommands.CheckForm(mSourceContext, SelectedPaths()) Then
        ShowStatus "Source context changed. Reopen Aggregation.": Exit Sub
    End If
    Set mActivity = New cProvisioningActivity
    mActivity.Initialize mSourceWorkbook, "ADMIN_AGGREGATION_", "Aggregation"
End Sub

Private Sub UserForm_Initialize()
    Set mSourceWorkbook = modAdmin.ResolveInteractiveAdminWorkbook()
    If Not mSourceWorkbook Is Nothing Then mSourceContext = modAggregationCommands.CaptureForm(mSourceWorkbook.Name)
    BuildLayout
    DiscoverConnectedRoots
End Sub

Private Sub UserForm_Activate()
    If Not mActivity Is Nothing Then
        mActivity.Presented Me.Visible
        If Not modAggregationCommands.CheckForm(mSourceContext, SelectedPaths()) Then ShowStatus "Source context changed. Reopen Aggregation."
    End If
    If mResizeInitialized Then Exit Sub
    On Error Resume Next
    modUserFormResizeWin.EnableResizableUserForm Me, True, True
    On Error GoTo 0
    mResizeInitialized = True
End Sub

Private Sub BuildLayout()
    If mBuilt Then Exit Sub

    Me.Caption = "invSys Aggregation Sources"
    Me.Width = 840
    Me.Height = 500

    AddLabel "Advisory aggregation source set", 12, 10, 310, 18, True
    AddLabel "Select the published warehouse snapshots to include for this run. This never changes Send To or local warehouse authority.", 12, 30, 785, 28, False
    AddLabel "NAS root", 12, 66, 65, 18, True
    Set mTxtRoot = AddTextBox("txtRoot", 80, 62, 410, 22)
    Set mBtnDiscover = AddButton("btnDiscover", "Refresh Connected", 500, 60, 105, 26)
    AddLabel "Windows user", 590, 66, 80, 18, False
    Set mTxtUser = AddTextBox("txtUser", 670, 62, 125, 22)
    AddLabel "Password", 12, 94, 62, 18, False
    Set mTxtPassword = AddTextBox("txtPassword", 80, 90, 190, 22)
    mTxtPassword.PasswordChar = "*"
    Set mBtnAddServer = AddButton("btnAddServer", "Add Server", 610, 60, 85, 26)
    Set mBtnConnect = AddButton("btnConnect", "Connect + Discover", 280, 88, 125, 26)
    AddLabel "Credentials are used only for this Windows connection and are cleared from this form after use.", 414, 94, 380, 18, False

    AddLabel "Selected Sources", 12, 126, 150, 18, True
    Set mLblSelected = AddLabel("Selected Sources: 0", 655, 126, 140, 18, True)
    Set mLstSources = AddListBox("lstSources", 12, 146, 783, 180)
    With mLstSources
        .ColumnCount = 6
        .ColumnWidths = "68 pt;120 pt;245 pt;100 pt;150 pt;65 pt"
        .MultiSelect = fmMultiSelectMulti
    End With

    AddLabel "Warehouse", 16, 132, 65, 12, False
    AddLabel "Server root", 88, 132, 90, 12, False
    AddLabel "Snapshot path", 214, 132, 95, 12, False
    AddLabel "Freshness", 464, 132, 65, 12, False
    AddLabel "Source fingerprint", 568, 132, 105, 12, False
    AddLabel "State", 735, 132, 45, 12, False

    AddLabel "Rejected / Skipped", 12, 338, 180, 18, True
    Set mLstRejected = AddListBox("lstRejected", 12, 358, 783, 64)
    mLstRejected.ColumnCount = 1
    mLstRejected.ColumnWidths = "760 pt"

    Set mBtnAggregate = AddButton("btnAggregate", "Aggregate Selected Sources", 484, 432, 175, 27)
    Set mBtnClose = AddButton("btnClose", "Close", 670, 432, 125, 27)
    Set mLblStatus = AddLabel("Discover a connected NAS root to begin.", 12, 432, 460, 28, False)
    mBuilt = True
End Sub

Private Function DiscoverConnectedRoots() As Boolean
    Dim roots As Collection
    Dim rootPath As Variant
    Dim configRoot As String
    Dim discovered As Boolean

    On Error GoTo FailedDiscovery
    mLoading = True: discovered = True
    Set roots = modNasConnection.GetKnownWarehouseTargetRoots()
    For Each rootPath In roots
        If Trim$(CStr(rootPath)) <> "" Then
            If mTxtRoot.Value = "" Then
                mTxtRoot.Value = CStr(rootPath)
            End If
            If Not DiscoverRoot(CStr(rootPath)) Then discovered = False
        End If
    Next rootPath
    If mLstSources.ListCount = 0 Then ShowStatus "No connected warehouse sources were found. Use Add Server only for another server."
    mLoading = False: DiscoverConnectedRoots = discovered
    Exit Function
FailedDiscovery:
    mLoading = False
    ShowStatus "Source discovery is unavailable."
End Function

Private Sub mBtnDiscover_Click()
    Dim action As String, outcome As String, paths As String
    On Error GoTo FailedDiscovery
    paths = SelectedPaths(): outcome = "DENIED"
    If Not mActivity Is Nothing Then action = mActivity.BeginAction("REFRESH")
    If Not modAggregationCommands.CheckForm(mSourceContext, paths) Then GoTo FinishedDiscovery
    outcome = "FAILED"
    If DiscoverConnectedRoots() Then outcome = "DISCOVERED"
    If Not modAggregationCommands.CheckForm(mSourceContext, SelectedPaths()) Then outcome = "DENIED"
FinishedDiscovery:
    If outcome = "DENIED" Then ShowStatus "Source context changed. Reopen Aggregation."
    FinishActivity action, outcome, SelectedPaths()
    Exit Sub
FailedDiscovery:
    mLoading = False: outcome = "FAILED"
    ShowStatus "Source discovery is unavailable."
    Resume FinishedDiscovery
End Sub

Private Sub mBtnAddServer_Click()
    Dim action As String, paths As String
    paths = SelectedPaths()
    If Not mActivity Is Nothing Then action = mActivity.BeginAction("ADD_SERVER")
    mTxtRoot.SetFocus
    ShowStatus "Enter another NAS server root and Windows credentials, then choose Connect + Discover."
    FinishActivity action, "STAGED", paths
End Sub

Private Sub mBtnConnect_Click()
    Dim report As String
    Dim rootPath As String
    Dim connected As Boolean
    Dim action As String, outcome As String, paths As String

    On Error GoTo FailedConnection
    rootPath = Trim$(mTxtRoot.Value): paths = SelectedPaths(): outcome = "DENIED"
    If Not mActivity Is Nothing Then action = mActivity.BeginAction("CONNECT")
    If Not modAggregationCommands.CheckForm(mSourceContext, paths) Then ShowStatus "Source context changed. Reopen Aggregation.": GoTo FinishedConnection
    outcome = "FAILED"
    connected = modAdminConsole.ConnectAggregationServerForAdmin(rootPath, Trim$(mTxtUser.Value), CStr(mTxtPassword.Value), report)
    mTxtPassword.Value = vbNullString
    mTxtUser.Value = vbNullString
    If Not modAggregationCommands.CheckForm(mSourceContext, paths) Then outcome = "DENIED": ShowStatus "Connection returned; source context changed. Reopen Aggregation.": GoTo FinishedConnection
    If Not connected Then ShowStatus "The Windows storage connection was not established.": GoTo FinishedConnection
    outcome = "CONNECTED"
    ShowStatus "Windows storage connection established."
    DiscoverRoot rootPath
FinishedConnection:
    mTxtPassword.Value = vbNullString: mTxtUser.Value = vbNullString
    FinishActivity action, outcome, paths
    Exit Sub
FailedConnection:
    outcome = "FAILED": ShowStatus "The connection action did not finish. Review the current connection state."
    Resume FinishedConnection
End Sub

Private Function DiscoverRoot(ByVal rootPath As String) As Boolean
    Dim records As Collection
    Dim report As String
    Dim item As Variant
    Dim loading As Boolean

    On Error GoTo FailedDiscovery
    loading = mLoading
    rootPath = Trim$(rootPath)
    If rootPath = "" Then
        ShowStatus "Enter a connected NAS root."
        Exit Function
    End If
    If Not modAdminConsole.DiscoverAggregationSourcesForAdmin(rootPath, records, report) Then
        AddRejected "Rejected: source discovery is unavailable."
        ShowStatus "Source discovery is unavailable."
        Exit Function
    End If
    mLoading = True
    For Each item In records
        AddSourceRecord CStr(item)
    Next item
    UpdateSelectedState
    ShowStatus "Source discovery returned. Review READY and rejected sources."
    mLoading = loading: DiscoverRoot = True
    Exit Function
FailedDiscovery:
    mLoading = loading
    ShowStatus "Source discovery is unavailable."
End Function

Private Sub AddSourceRecord(ByVal recordText As String)
    Dim fields() As String
    Dim i As Long
    Dim fingerprint As String

    fields = Split(recordText, vbTab)
    If UBound(fields) < 5 Then
        AddRejected "Rejected: invalid discovery record."
        Exit Sub
    End If
    For i = 0 To mLstSources.ListCount - 1
        If StrComp(CStr(mLstSources.List(i, 0)), fields(0), vbTextCompare) = 0 Then
            If StrComp(CStr(mLstSources.List(i, 2)), fields(2), vbTextCompare) <> 0 Then
                AddRejected "Rejected: WarehouseId " & fields(0) & " is available from more than one source identity."
            End If
            Exit Sub
        End If
    Next i

    fingerprint = modAggregationCommands.RegisterSource(mSourceContext, fields(0), fields(5), fields(2))
    If fingerprint = "" Then
        AddRejected "Rejected: source identity, publication or schema is invalid."
        Exit Sub
    End If

    mLstSources.AddItem fields(0)
    i = mLstSources.ListCount - 1
    mLstSources.List(i, 1) = fields(1)
    mLstSources.List(i, 2) = fields(2)
    mLstSources.List(i, 3) = fields(3)
    mLstSources.List(i, 4) = fingerprint
    mLstSources.List(i, 5) = "READY"
    mLstSources.Selected(i) = True
End Sub

Private Sub AddRejected(ByVal message As String)
    mLstRejected.AddItem message
End Sub

Private Sub mLstSources_Click()
    Dim action As String, paths As String
    If mLoading Or mBusy Then Exit Sub
    paths = SelectedPaths()
    If Not mActivity Is Nothing Then action = mActivity.BeginAction("SELECT_SOURCE")
    UpdateSelectedState
    FinishActivity action, "SELECTED", paths
End Sub

Private Sub mBtnAggregate_Click()
    Dim selectedFiles As Collection
    Dim path As Variant
    Dim report As String
    Dim targetWb As Workbook
    Dim confirmed As Boolean, completed As Boolean
    Dim capturedPaths As String
    Dim action As String, outcome As String

    If mBusy Then Exit Sub
    mBusy = True: mBtnAggregate.Enabled = False
    On Error GoTo FailedAggregate

    capturedPaths = SelectedPaths(): outcome = "REJECTED"
    If Not mActivity Is Nothing Then action = mActivity.BeginAction("AGGREGATE")
    Set selectedFiles = New Collection
    If capturedPaths <> "" Then
        For Each path In Split(capturedPaths, vbLf): selectedFiles.Add CStr(path): Next path
    End If
    If selectedFiles.Count = 0 Then
        ShowStatus "Select at least one READY source."
        GoTo FinishedAggregate
    End If

    outcome = "DENIED"
    If Not modAggregationCommands.CheckForm(mSourceContext, capturedPaths) Then
        ShowStatus "Source context changed. Reopen Aggregation.": GoTo FinishedAggregate
    End If
    Set targetWb = mSourceWorkbook
    If modAggregationCommands.NeedsReview(mSourceContext) Then
        confirmed = (MsgBox("A previous attempt needs review. Verify its saved state and resolve its retained staging before starting a fresh attempt?", vbYesNo Or vbDefaultButton2, "invSys Aggregator") = vbYes)
        If Not modAggregationCommands.CheckForm(mSourceContext, capturedPaths) Then
            ShowStatus "Source session, feed or selected snapshots changed. Reopen Aggregation."
            GoTo FinishedAggregate
        End If
        If Not modAggregationCommands.ResolveRetry(mSourceContext, capturedPaths, confirmed, report) Then
            outcome = "INDETERMINATE": If Not confirmed Then outcome = "CANCELLED"
            ShowStatus report: GoTo FinishedAggregate
        End If
    End If
    completed = modAdminConsole.RunHQAggregationFromSourceSet(selectedFiles, "", "", targetWb, report, mSourceContext, outcome)
    If completed Then
        ShowStatus "Completed. " & report
        MsgBox "Advisory Global Inventory Snapshot updated." & vbCrLf & vbCrLf & report & vbCrLf & vbCrLf & _
               "The selected source set was used for this session only. Warehouse-local inventory remains authoritative.", _
               vbInformation, "invSys Aggregator"
    Else
        ShowStatus report
        MsgBox report, vbExclamation, "invSys Aggregator"
    End If
    If Not modAggregationCommands.CheckForm(mSourceContext, capturedPaths) Then
        If completed Then
            ShowStatus "Complete output was saved. Reopen Aggregation before another attempt."
        ElseIf InStr(1, report, "Reopen Aggregation", vbTextCompare) = 0 Then
            ShowStatus report & " Reopen Aggregation before another attempt."
        End If
    End If
FinishedAggregate:
    FinishActivity action, outcome, capturedPaths
    mBusy = False: mBtnAggregate.Enabled = True
    Exit Sub
FailedAggregate:
    outcome = "FAILED"
    ShowStatus "Aggregation did not complete. Review the current output and retained attempt before retrying."
    Resume FinishedAggregate
End Sub

Private Function SelectedPaths() As String
    Dim i As Long
    If mLstSources Is Nothing Then Exit Function
    For i = 0 To mLstSources.ListCount - 1
        If mLstSources.Selected(i) Then
            If SelectedPaths <> "" Then SelectedPaths = SelectedPaths & vbLf
            SelectedPaths = SelectedPaths & CStr(mLstSources.List(i, 2))
        End If
    Next i
End Function

Private Sub UserForm_Terminate()
    modAggregationCommands.ReleaseForm mSourceContext
    Set mSourceWorkbook = Nothing
    Set mActivity = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Cancel <> 0 Or mActivity Is Nothing Then Exit Sub
    FinishActivity mActivity.BeginAction("CLOSE"), "CLOSED", SelectedPaths()
End Sub

Private Sub FinishActivity(ByVal action As String, ByVal outcome As String, ByVal paths As String)
    If mActivity Is Nothing Then Exit Sub
    mActivity.Finish action, outcome
    ' Optional tracking notices yield just like the existing result dialogs.
    If Not modAggregationCommands.CheckForm(mSourceContext, paths) Then
        If InStr(1, mLblStatus.Caption, "Reopen Aggregation", vbTextCompare) = 0 Then _
            ShowStatus mLblStatus.Caption & " Reopen Aggregation before another attempt."
    End If
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub

' Packaged smoke entry point used by the D13 build validation. It initializes
' the same form instance opened by the public Admin callback without selecting
' a source, connecting to a server, or changing warehouse state.
Public Function AggregationSourcesSmokeForAutomation() As String
    On Error GoTo FailSmoke
    If mLstSources Is Nothing Or mLstRejected Is Nothing Then
        AggregationSourcesSmokeForAutomation = "FAIL|Source-set controls were not initialized."
    Else
        AggregationSourcesSmokeForAutomation = "OK|Aggregation source-set form initialized."
    End If
    Unload Me
    Exit Function

FailSmoke:
    AggregationSourcesSmokeForAutomation = "FAIL|" & Err.Description
    On Error Resume Next
    Unload Me
    On Error GoTo 0
End Function

Private Sub UpdateSelectedState()
    Dim i As Long
    Dim countSelected As Long

    For i = 0 To mLstSources.ListCount - 1
        If mLstSources.Selected(i) Then countSelected = countSelected + 1
    Next i
    mLblSelected.Caption = "Selected Sources: " & CStr(countSelected)
End Sub

Private Sub ShowStatus(ByVal message As String)
    mLblStatus.Caption = message
End Sub

Private Function AddLabel(ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single, ByVal boldText As Boolean) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", "lbl" & CStr(Me.Controls.Count + 1), True)
    With AddLabel
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = boldText
    End With
End Function

Private Function AddTextBox(ByVal name As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", name, True)
    With AddTextBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddListBox(ByVal name As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", name, True)
    With AddListBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddButton(ByVal name As String, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", name, True)
    With AddButton
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function
