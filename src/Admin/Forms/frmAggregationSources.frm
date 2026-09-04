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
Private WithEvents mBtnAggregate As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private mLblSelected As MSForms.Label
Private mLblStatus As MSForms.Label
Private mBuilt As Boolean
Private mResizeInitialized As Boolean

Private Sub UserForm_Initialize()
    BuildLayout
    LoadKnownRoots
End Sub

Private Sub UserForm_Activate()
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
    Set mBtnDiscover = AddButton("btnDiscover", "Discover", 500, 60, 80, 26)
    AddLabel "Windows user", 590, 66, 80, 18, False
    Set mTxtUser = AddTextBox("txtUser", 670, 62, 125, 22)
    AddLabel "Password", 12, 94, 62, 18, False
    Set mTxtPassword = AddTextBox("txtPassword", 80, 90, 190, 22)
    mTxtPassword.PasswordChar = "*"
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

Private Sub LoadKnownRoots()
    Dim roots As Collection
    Dim rootPath As Variant
    Dim configRoot As String

    On Error Resume Next
    configRoot = Trim$(modConfig.GetString("PathSharePointRoot", ""))
    On Error GoTo 0
    If configRoot <> "" Then
        mTxtRoot.Value = configRoot
    End If

    Set roots = modNasConnection.GetKnownWarehouseTargetRoots()
    For Each rootPath In roots
        If Trim$(CStr(rootPath)) <> "" Then
            If mTxtRoot.Value = "" Then
                mTxtRoot.Value = CStr(rootPath)
            End If
        End If
    Next rootPath
End Sub

Private Sub mBtnDiscover_Click()
    DiscoverRoot Trim$(mTxtRoot.Value)
End Sub

Private Sub mBtnConnect_Click()
    Dim report As String
    Dim rootPath As String

    rootPath = Trim$(mTxtRoot.Value)
    If Not modAdminConsole.ConnectAggregationServerForAdmin(rootPath, Trim$(mTxtUser.Value), CStr(mTxtPassword.Value), report) Then
        ShowStatus report
        mTxtPassword.Value = vbNullString
        Exit Sub
    End If
    mTxtPassword.Value = vbNullString
    ShowStatus report
    DiscoverRoot rootPath
End Sub

Private Sub DiscoverRoot(ByVal rootPath As String)
    Dim records As Collection
    Dim report As String
    Dim item As Variant

    rootPath = Trim$(rootPath)
    If rootPath = "" Then
        ShowStatus "Enter a connected NAS root."
        Exit Sub
    End If
    If Not modAdminConsole.DiscoverAggregationSourcesForAdmin(rootPath, records, report) Then
        AddRejected "Rejected: " & report
        ShowStatus report
        Exit Sub
    End If
    For Each item In records
        AddSourceRecord CStr(item)
    Next item
    UpdateSelectedState
    ShowStatus report
End Sub

Private Sub AddSourceRecord(ByVal recordText As String)
    Dim fields() As String
    Dim i As Long

    fields = Split(recordText, vbTab)
    If UBound(fields) < 4 Then
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

    mLstSources.AddItem fields(0)
    i = mLstSources.ListCount - 1
    mLstSources.List(i, 1) = fields(1)
    mLstSources.List(i, 2) = fields(2)
    mLstSources.List(i, 3) = fields(3)
    mLstSources.List(i, 4) = fields(4)
    mLstSources.List(i, 5) = "READY"
    mLstSources.Selected(i) = True
End Sub

Private Sub AddRejected(ByVal message As String)
    mLstRejected.AddItem message
End Sub

Private Sub mLstSources_Click()
    UpdateSelectedState
End Sub

Private Sub mBtnAggregate_Click()
    Dim selectedFiles As Collection
    Dim i As Long
    Dim report As String
    Dim targetWb As Workbook

    Set selectedFiles = New Collection
    For i = 0 To mLstSources.ListCount - 1
        If mLstSources.Selected(i) Then selectedFiles.Add CStr(mLstSources.List(i, 2))
    Next i
    If selectedFiles.Count = 0 Then
        ShowStatus "Select at least one READY source."
        Exit Sub
    End If

    Set targetWb = modAdmin.ResolveInteractiveAdminWorkbook()
    If modAdminConsole.RunHQAggregationFromSourceSet(selectedFiles, "", "", targetWb, report) Then
        ShowStatus "Completed. " & report
        MsgBox "Advisory Global Inventory Snapshot updated." & vbCrLf & vbCrLf & report & vbCrLf & vbCrLf & _
               "The selected source set was used for this session only. Warehouse-local inventory remains authoritative.", _
               vbInformation, "invSys Aggregator"
    Else
        ShowStatus report
        MsgBox report, vbExclamation, "invSys Aggregator"
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
