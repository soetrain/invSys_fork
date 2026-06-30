VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShippingBoxBuilder
   Caption         =   "Shipping Box Builder"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12480
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShippingBoxBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mCboBoxes As MSForms.ComboBox
Private WithEvents mCboVersions As MSForms.ComboBox
Private WithEvents mCboStatus As MSForms.ComboBox
Private WithEvents mChkShowActive As MSForms.CheckBox
Private WithEvents mChkShowArchived As MSForms.CheckBox
Private WithEvents mTxtBoxName As MSForms.TextBox
Private WithEvents mTxtUom As MSForms.TextBox
Private WithEvents mTxtLocation As MSForms.TextBox
Private WithEvents mTxtDescription As MSForms.TextBox
Private WithEvents mTxtSearch As MSForms.TextBox
Private WithEvents mTxtQty As MSForms.TextBox
Private WithEvents mLstInventory As MSForms.ListBox
Private WithEvents mLstBom As MSForms.ListBox
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnNewBox As MSForms.CommandButton
Private WithEvents mBtnAdd As MSForms.CommandButton
Private WithEvents mBtnRemove As MSForms.CommandButton
Private WithEvents mBtnSaveBox As MSForms.CommandButton
Private WithEvents mBtnUpdateVersion As MSForms.CommandButton
Private WithEvents mBtnNewVersion As MSForms.CommandButton
Private WithEvents mBtnDeleteVersion As MSForms.CommandButton
Private WithEvents mBtnArchiveBox As MSForms.CommandButton
Private WithEvents mBtnDeleteBox As MSForms.CommandButton
Private WithEvents mBtnCancel As MSForms.CommandButton

Private mLblStatus As MSForms.Label
Private mLblInventoryStatus As MSForms.Label
Private mInventoryData As Variant
Private mSavedBoxes As Variant
Private mVersionRows As Variant
Private mSelectedPackageRow As Long
Private mBuilt As Boolean
Private mLoading As Boolean
Private mAnchors As Object
Private mResizeInitialized As Boolean

Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Private Sub UserForm_Activate()
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me
        mResizeInitialized = True
    End If
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub UserForm_Layout()
    If mAnchors Is Nothing Then Exit Sub
    mAnchors.ResizeControls
End Sub

Private Sub UserForm_Terminate()
    Set mAnchors = Nothing
End Sub

Public Sub InitializeFromShipping()
    If Not mBuilt Then BuildLayout

    mLoading = True
    LoadSavedBoxes
    LoadCurrentBoxState
    LoadInventoryCache
    RenderInventoryList
    mLoading = False

    If mSelectedPackageRow > 0 Then
        LoadVersionsForPackage mSelectedPackageRow
        LoadSelectedVersionComponents
    End If
End Sub

Private Sub BuildLayout()
    Dim previousLoading As Boolean

    If mBuilt Then Exit Sub
    previousLoading = mLoading
    mLoading = True
    mBuilt = True

    Me.Caption = "Shipping Box Builder"
    Me.Width = 900
    Me.Height = 610

    AddLabel "lblTitle", "Box Builder", 12, 10, 180, 20, True
    Set mChkShowActive = AddCheckBox("chkShowActive", "Show Active", 220, 8, 92, 22)
    mChkShowActive.Value = True
    Set mChkShowArchived = AddCheckBox("chkShowArchived", "Show Archived", 320, 8, 110, 22)
    AddLabel "lblBoxSelect", "Saved Box", 12, 40, 76, 18, False
    AddLabel "lblVersionSelect", "Version", 472, 40, 70, 18, False
    AddLabel "lblStatusSelect", "Status", 606, 40, 54, 18, False

    Set mCboBoxes = AddComboBox("cboBoxes", 88, 36, 294, 22)
    With mCboBoxes
        .ColumnCount = 5
        .ColumnWidths = "0 pt;150 pt;42 pt;72 pt;120 pt"
        .BoundColumn = 1
        .TextColumn = 2
        .Style = 2
    End With
    Set mBtnNewBox = AddButton("btnNewBox", "New Box", 392, 34, 72, 26)
    Set mCboVersions = AddComboBox("cboVersions", 510, 36, 78, 22)
    With mCboVersions
        .ColumnCount = 8
        .ColumnWidths = "42 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt"
        .Style = 2
    End With
    Set mCboStatus = AddComboBox("cboStatus", 658, 36, 86, 22)
    With mCboStatus
        .Style = 2
        .AddItem "Active"
        .AddItem "Retired/Archived"
        .ListIndex = 0
    End With
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 758, 34, 78, 26)

    AddLabel "lblBoxName", "Box Name", 12, 78, 70, 18, False
    AddLabel "lblUom", "UOM", 294, 78, 40, 18, False
    AddLabel "lblLocation", "Location", 392, 78, 70, 18, False
    AddLabel "lblDesc", "Description", 12, 112, 86, 18, False

    Set mTxtBoxName = AddTextBox("txtBoxName", 88, 74, 190, 22)
    Set mTxtUom = AddTextBox("txtUom", 330, 74, 44, 22)
    Set mTxtLocation = AddTextBox("txtLocation", 456, 74, 118, 22)
    Set mTxtDescription = AddTextBox("txtDescription", 100, 108, 628, 22)

    AddLabel "lblSearch", "Managed Inventory", 12, 154, 140, 18, True
    Set mLblInventoryStatus = AddLabel("lblInventoryStatus", "", 154, 154, 256, 18, False)
    AddLabel "lblQty", "Qty", 328, 154, 30, 18, False
    AddLabel "lblBom", "BOM Components", 450, 154, 160, 18, True

    Set mTxtSearch = AddTextBox("txtSearch", 12, 176, 304, 22)
    Set mTxtQty = AddTextBox("txtQty", 360, 176, 50, 22)

    Set mLstInventory = AddListBox("lstInventory", 12, 206, 398, 284)
    With mLstInventory
        .ColumnCount = 7
        .ColumnWidths = "38 pt;62 pt;118 pt;36 pt;58 pt;106 pt;52 pt"
    End With

    Set mLstBom = AddListBox("lstBom", 450, 176, 386, 314)
    With mLstBom
        .ColumnCount = 8
        .ColumnWidths = "34 pt;126 pt;72 pt;40 pt;42 pt;38 pt;62 pt;124 pt"
    End With

    Set mBtnAdd = AddButton("btnAdd", "Add", 250, 498, 76, 28)
    Set mBtnRemove = AddButton("btnRemove", "Remove", 450, 498, 76, 28)
    Set mBtnSaveBox = AddButton("btnSaveBox", "Save Box", 538, 498, 86, 28)
    Set mBtnUpdateVersion = AddButton("btnUpdateVersion", "Update Version", 636, 498, 102, 28)
    Set mBtnNewVersion = AddButton("btnNewVersion", "Save New Version", 748, 498, 118, 28)
    Set mBtnArchiveBox = AddButton("btnArchiveBox", "Archive Box", 450, 534, 76, 28)
    Set mBtnDeleteVersion = AddButton("btnDeleteVersion", "Delete Version", 538, 534, 102, 28)
    Set mBtnDeleteBox = AddButton("btnDeleteBox", "Delete Box", 650, 534, 86, 28)
    Set mBtnCancel = AddButton("btnCancel", "Close", 760, 534, 76, 28)
    Set mLblStatus = AddLabel("lblStatus", "", 12, 536, 420, 36, False)

    mTxtUom.Value = "ea"
    mTxtQty.Value = "1"
    InitializeBoxBuilderAnchors
    mLoading = previousLoading
End Sub

Private Sub LoadSavedBoxes()
    On Error GoTo FailSoft

    Dim r As Long
    Dim idx As Long

    mSavedBoxes = modTS_Shipments.BoxBuilderFormLoadSavedBoxes(CBool(mChkShowActive.Value), CBool(mChkShowArchived.Value))
    mCboBoxes.Clear
    mSelectedPackageRow = 0
    If IsEmpty(mSavedBoxes) Then
        ResetForNewBox
        If Not CBool(mChkShowActive.Value) And Not CBool(mChkShowArchived.Value) Then
            ShowStatus "No boxes shown. Enable Show Active or Show Archived."
        ElseIf CBool(mChkShowArchived.Value) Then
            ShowStatus "No boxes match the current active/archive filters."
        Else
            ShowStatus "No active saved boxes found. Use Show Archived to view retired designs."
        End If
        Exit Sub
    End If

    For r = 1 To UBound(mSavedBoxes, 1)
        mCboBoxes.AddItem NzText(mSavedBoxes(r, 1))
        idx = mCboBoxes.ListCount - 1
        mCboBoxes.List(idx, 1) = NzText(mSavedBoxes(r, 2))
        mCboBoxes.List(idx, 2) = NzText(mSavedBoxes(r, 3))
        mCboBoxes.List(idx, 3) = NzText(mSavedBoxes(r, 4))
        mCboBoxes.List(idx, 4) = NzText(mSavedBoxes(r, 5))
    Next r
    ShowStatus "Loaded " & CStr(mCboBoxes.ListCount) & " box design(s)."
    Exit Sub

FailSoft:
    ShowStatus "Could not load saved boxes: " & Err.Description
End Sub

Private Sub mChkShowActive_Click()
    If mLoading Then Exit Sub
    If mCboBoxes Is Nothing Then Exit Sub
    ReloadSavedBoxesAfterFilterChange
End Sub

Private Sub mChkShowArchived_Click()
    If mLoading Then Exit Sub
    If mCboBoxes Is Nothing Then Exit Sub
    ReloadSavedBoxesAfterFilterChange
End Sub

Private Sub ReloadSavedBoxesAfterFilterChange()
    If mCboBoxes Is Nothing Then Exit Sub
    LoadSavedBoxes
    If mCboBoxes.ListCount > 0 Then
        mCboBoxes.ListIndex = 0
        LoadSelectedBox
    Else
        ResetForNewBox
    End If
End Sub

Private Sub LoadCurrentBoxState()
    On Error GoTo FailSoft

    Dim meta As Variant
    Dim currentName As String
    Dim i As Long

    meta = modTS_Shipments.BoxBuilderFormCurrentMeta()
    If Not IsEmpty(meta) Then
        mTxtBoxName.Value = NzText(meta(1))
        mTxtUom.Value = NzText(meta(2))
        mTxtLocation.Value = NzText(meta(3))
        mTxtDescription.Value = NzText(meta(4))
    End If
    If Trim$(CStr(mTxtUom.Value)) = "" Then mTxtUom.Value = "ea"

    currentName = LCase$(Trim$(CStr(mTxtBoxName.Value)))
    If currentName <> "" Then
        For i = 0 To mCboBoxes.ListCount - 1
            If LCase$(Trim$(CStr(mCboBoxes.List(i, 1)))) = currentName Then
                mCboBoxes.ListIndex = i
                mSelectedPackageRow = CLng(Val(CStr(mCboBoxes.List(i, 0))))
                Exit For
            End If
        Next i
    End If

    If mSelectedPackageRow <= 0 Then LoadCurrentDisplayedComponents
    Exit Sub

FailSoft:
    ShowStatus "Could not load current BoxBuilder rows: " & Err.Description
End Sub

Private Sub LoadCurrentDisplayedComponents()
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long

    mLstBom.Clear
    EnsureVersionDefaults
    rowsData = modTS_Shipments.BoxBuilderFormCurrentComponents()
    If IsEmpty(rowsData) Then Exit Sub
    For r = 1 To UBound(rowsData, 1)
        AddBomListRow NzText(rowsData(r, 1)), _
                      NzText(rowsData(r, 2)), _
                      NzText(rowsData(r, 3)), _
                      NzText(rowsData(r, 4)), _
                      NzText(rowsData(r, 5)), _
                      NzText(rowsData(r, 6)), _
                      NzText(rowsData(r, 7)), _
                      NzText(rowsData(r, 8))
    Next r
    Exit Sub

FailSoft:
    ShowStatus "Could not load displayed components: " & Err.Description
End Sub

Private Sub LoadVersionsForPackage(ByVal packageRow As Long)
    On Error GoTo FailSoft

    Dim r As Long
    Dim c As Long
    Dim idx As Long

    mVersionRows = modTS_Shipments.BoxBuilderFormLoadVersions(packageRow)
    mCboVersions.Clear
    If IsEmpty(mVersionRows) Then
        mCboVersions.AddItem "v1"
        mCboVersions.ListIndex = 0
        mCboStatus.Value = "Active"
        Exit Sub
    End If

    For r = 1 To UBound(mVersionRows, 1)
        mCboVersions.AddItem NzText(mVersionRows(r, 1))
        idx = mCboVersions.ListCount - 1
        For c = 2 To 8
            mCboVersions.List(idx, c - 1) = NzText(mVersionRows(r, c))
        Next c
    Next r
    If mCboVersions.ListCount > 0 Then mCboVersions.ListIndex = 0
    Exit Sub

FailSoft:
    ShowStatus "Could not load versions: " & Err.Description
End Sub

Private Sub LoadSelectedVersionComponents()
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long

    If mSelectedPackageRow <= 0 Then Exit Sub
    If mCboVersions.ListIndex < 0 Then Exit Sub

    mCboStatus.Value = SelectedVersionStatus()
    rowsData = modTS_Shipments.BoxBuilderFormLoadVersionComponents(mSelectedPackageRow, SelectedVersionLabel())
    mLstBom.Clear
    If IsEmpty(rowsData) Then
        ShowStatus "No components found for " & SelectedVersionLabel() & "."
        Exit Sub
    End If

    For r = 1 To UBound(rowsData, 1)
        AddBomListRow NzText(rowsData(r, 1)), _
                      NzText(rowsData(r, 2)), _
                      NzText(rowsData(r, 3)), _
                      NzText(rowsData(r, 4)), _
                      NzText(rowsData(r, 5)), _
                      NzText(rowsData(r, 6)), _
                      NzText(rowsData(r, 7)), _
                      NzText(rowsData(r, 8))
    Next r
    ShowStatus "Loaded " & CStr(mLstBom.ListCount) & " component row(s) for " & SelectedVersionLabel() & "."
    Exit Sub

FailSoft:
    ShowStatus "Could not load version components: " & Err.Description
End Sub

Private Sub LoadInventoryCache()
    On Error GoTo FailSoft
    mInventoryData = modTS_Shipments.LoadShippingComponentPickerItems()
    If IsEmpty(mInventoryData) Then
        ShowInventoryStatus modTS_Shipments.ShippingComponentPickerLastStatus()
    Else
        ShowInventoryStatus modTS_Shipments.ShippingComponentPickerLastStatus()
    End If
    Exit Sub

FailSoft:
    mInventoryData = Empty
    ShowStatus "Inventory load failed: " & Err.Description
End Sub

Private Sub RenderInventoryList()
    On Error GoTo FailSoft

    Dim filterText As String
    Dim haystack As String
    Dim r As Long
    Dim idx As Long

    mLstInventory.Clear
    If IsEmpty(mInventoryData) Then Exit Sub

    filterText = LCase$(Trim$(CStr(mTxtSearch.Value)))
    For r = 1 To UBound(mInventoryData, 1)
        haystack = LCase$(NzText(mInventoryData(r, 1)) & " " & _
                          NzText(mInventoryData(r, 2)) & " " & _
                          NzText(mInventoryData(r, 3)) & " " & _
                          NzText(mInventoryData(r, 6)))
        If filterText = "" Or InStr(1, haystack, filterText, vbTextCompare) > 0 Then
            mLstInventory.AddItem NzText(mInventoryData(r, 1))
            idx = mLstInventory.ListCount - 1
            mLstInventory.List(idx, 1) = NzText(mInventoryData(r, 2))
            mLstInventory.List(idx, 2) = NzText(mInventoryData(r, 3))
            mLstInventory.List(idx, 3) = NzText(mInventoryData(r, 4))
            mLstInventory.List(idx, 4) = NzText(mInventoryData(r, 5))
            mLstInventory.List(idx, 5) = NzText(mInventoryData(r, 6))
            mLstInventory.List(idx, 6) = NzText(mInventoryData(r, 7))
        End If
    Next r
    If mLstInventory.ListCount = 0 Then
        ShowInventoryStatus "Inventory loaded, but no rows match the search filter."
    Else
        ShowInventoryStatus CStr(mLstInventory.ListCount) & " inventory row(s) shown."
    End If
    Exit Sub

FailSoft:
    ShowStatus "Inventory filter failed: " & Err.Description
End Sub

Private Sub mCboBoxes_Change()
    If mLoading Then Exit Sub
    LoadSelectedBox
End Sub

Private Sub mCboVersions_Change()
    If mLoading Then Exit Sub
    LoadSelectedVersionComponents
End Sub

Private Sub mTxtSearch_Change()
    RenderInventoryList
End Sub

Private Sub mBtnRefresh_Click()
    InitializeFromShipping
End Sub

Private Sub mBtnNewBox_Click()
    ResetForNewBox
    ShowStatus "New box design ready. Enter metadata, add inventory components, then Save Box."
End Sub

Private Sub mBtnAdd_Click()
    If mLstInventory.ListIndex < 0 Then
        ShowStatus "Select a managed inventory item to add."
        Exit Sub
    End If
    If ParseNumber(Trim$(CStr(mTxtQty.Value))) <= 0 Then
        ShowStatus "Enter a positive component quantity."
        Exit Sub
    End If

    AddBomListRow SelectedVersionLabel(), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 2)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 1)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 0)), _
                  Trim$(CStr(mTxtQty.Value)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 3)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 4)), _
                  CStr(mLstInventory.List(mLstInventory.ListIndex, 5))
    ShowStatus "Component added."
End Sub

Private Sub mBtnRemove_Click()
    If mLstBom.ListIndex < 0 Then
        ShowStatus "Select a BOM component to remove."
        Exit Sub
    End If
    mLstBom.RemoveItem mLstBom.ListIndex
    ShowStatus "Component removed."
End Sub

Private Sub mBtnUpdateVersion_Click()
    If mSelectedPackageRow <= 0 Or mCboVersions.ListIndex < 0 Then
        ShowStatus "Select a saved box/version before updating."
        Exit Sub
    End If
    SaveWithAction "UPDATE"
End Sub

Private Sub mBtnSaveBox_Click()
    SaveWithAction "BOX"
End Sub

Private Sub mBtnNewVersion_Click()
    SaveWithAction "NEW"
End Sub

Private Sub mBtnDeleteVersion_Click()
    If mSelectedPackageRow <= 0 Or mCboVersions.ListIndex < 0 Then
        ShowStatus "Select a saved box/version before deleting."
        Exit Sub
    End If
    modTS_Shipments.BoxBuilderFormDeleteVersion mSelectedPackageRow, SelectedVersionLabel()
    InitializeFromShipping
End Sub

Private Sub mBtnArchiveBox_Click()
    Dim report As String

    If mSelectedPackageRow <= 0 Then
        ShowStatus "Select a saved box before archiving."
        Exit Sub
    End If
    If MsgBox("Archive active designs for Shipping BOM ROW " & CStr(mSelectedPackageRow) & "?" & vbCrLf & _
              "Existing inventory remains available. New box making will use only active designs.", _
              vbQuestion + vbYesNo, "Archive Box Design") <> vbYes Then Exit Sub

    If modTS_Shipments.BoxBuilderFormArchiveBox(mSelectedPackageRow, report) Then
        InitializeFromShipping
        ShowStatus report
    Else
        If report = "" Then report = "Could not archive selected box design."
        ShowStatus report
        MsgBox report, vbExclamation
    End If
End Sub

Private Sub mBtnDeleteBox_Click()
    If mSelectedPackageRow <= 0 Then
        ShowStatus "Select a saved box before deleting."
        Exit Sub
    End If
    modTS_Shipments.BoxBuilderFormDeleteBox mSelectedPackageRow
    InitializeFromShipping
End Sub

Private Sub mBtnCancel_Click()
    Me.Hide
End Sub

Private Sub LoadSelectedBox()
    On Error GoTo FailSoft

    If mCboBoxes.ListIndex < 0 Then Exit Sub
    mSelectedPackageRow = CLng(Val(CStr(mCboBoxes.List(mCboBoxes.ListIndex, 0))))
    mTxtBoxName.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 1))
    mTxtUom.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 2))
    mTxtLocation.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 3))
    mTxtDescription.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 4))
    If Trim$(CStr(mTxtUom.Value)) = "" Then mTxtUom.Value = "ea"

    LoadVersionsForPackage mSelectedPackageRow
    LoadSelectedVersionComponents
    Exit Sub

FailSoft:
    ShowStatus "Could not load selected box: " & Err.Description
End Sub

Private Sub ResetForNewBox()
    On Error Resume Next
    mSelectedPackageRow = 0
    mCboBoxes.ListIndex = -1
    mCboVersions.Clear
    mCboVersions.AddItem "v1"
    mCboVersions.ListIndex = 0
    mCboStatus.Value = "Active"
    mTxtBoxName.Value = ""
    mTxtUom.Value = "ea"
    mTxtLocation.Value = ""
    mTxtDescription.Value = ""
    mTxtSearch.Value = ""
    mTxtQty.Value = "1"
    mLstBom.Clear
    RenderInventoryList
    On Error GoTo 0
End Sub

Private Sub EnsureVersionDefaults()
    If mCboVersions.ListCount > 0 Then Exit Sub
    mCboVersions.AddItem "v1"
    mCboVersions.ListIndex = 0
    mCboStatus.Value = "Active"
End Sub

Private Sub SaveWithAction(ByVal saveAction As String)
    On Error GoTo ErrHandler

    Dim bomRows As Variant
    Dim i As Long
    Dim saveVersionLabel As String

    If Trim$(CStr(mTxtBoxName.Value)) = "" Then
        ShowStatus "Enter a Box Name."
        Exit Sub
    End If
    If Trim$(CStr(mTxtUom.Value)) = "" Then
        ShowStatus "Enter a UOM."
        Exit Sub
    End If
    If mLstBom.ListCount = 0 Then
        ShowStatus "Add at least one BOM component."
        Exit Sub
    End If

    saveAction = UCase$(Trim$(saveAction))
    If saveAction = "BOX" Then
        saveVersionLabel = "v1"
    Else
        saveVersionLabel = SelectedVersionLabel()
    End If

    ReDim bomRows(1 To mLstBom.ListCount, 1 To 8)
    For i = 0 To mLstBom.ListCount - 1
        bomRows(i + 1, 1) = saveVersionLabel
        bomRows(i + 1, 2) = CStr(mLstBom.List(i, 1))
        bomRows(i + 1, 3) = CStr(mLstBom.List(i, 2))
        bomRows(i + 1, 4) = CStr(mLstBom.List(i, 3))
        bomRows(i + 1, 5) = CStr(mLstBom.List(i, 4))
        bomRows(i + 1, 6) = CStr(mLstBom.List(i, 5))
        bomRows(i + 1, 7) = CStr(mLstBom.List(i, 6))
        bomRows(i + 1, 8) = CStr(mLstBom.List(i, 7))
    Next i

    modTS_Shipments.CommitBoxBuilderFormState CStr(mTxtBoxName.Value), _
                                             CStr(mTxtUom.Value), _
                                             CStr(mTxtLocation.Value), _
                                             CStr(mTxtDescription.Value), _
                                             bomRows, _
                                             saveAction, _
                                             saveVersionLabel, _
                                             CStr(mCboStatus.Value)
    InitializeFromShipping
    Exit Sub

ErrHandler:
    ShowStatus "Save failed: " & Err.Description
End Sub

Private Sub AddBomListRow(ByVal versionText As String, _
                          ByVal itemName As String, _
                          ByVal itemCode As String, _
                          ByVal rowText As String, _
                          ByVal qtyText As String, _
                          ByVal uomText As String, _
                          ByVal locationText As String, _
                          ByVal descriptionText As String)
    Dim idx As Long

    versionText = NormalizeVersionText(versionText)
    mLstBom.AddItem versionText
    idx = mLstBom.ListCount - 1
    mLstBom.List(idx, 1) = itemName
    mLstBom.List(idx, 2) = itemCode
    mLstBom.List(idx, 3) = rowText
    mLstBom.List(idx, 4) = qtyText
    mLstBom.List(idx, 5) = uomText
    mLstBom.List(idx, 6) = locationText
    mLstBom.List(idx, 7) = descriptionText
End Sub

Private Function SelectedVersionLabel() As String
    If Not mCboVersions Is Nothing Then
        If mCboVersions.ListIndex >= 0 Then SelectedVersionLabel = NzText(mCboVersions.List(mCboVersions.ListIndex, 0))
    End If
    If SelectedVersionLabel = "" Then SelectedVersionLabel = "v1"
    SelectedVersionLabel = NormalizeVersionText(SelectedVersionLabel)
End Function

Private Function SelectedVersionStatus() As String
    SelectedVersionStatus = "Active"
    If mCboVersions Is Nothing Then Exit Function
    If mCboVersions.ListIndex < 0 Then Exit Function
    If mCboVersions.ColumnCount < 2 Then Exit Function
    SelectedVersionStatus = NzText(mCboVersions.List(mCboVersions.ListIndex, 1))
    If SelectedVersionStatus = "" Then SelectedVersionStatus = "Active"
End Function

Private Sub InitializeBoxBuilderAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me

    mAnchors.Add mCboBoxes, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mChkShowActive, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mChkShowArchived, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mBtnRefresh, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtDescription, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblInventoryStatus, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtSearch, ANCHOR_LEFT Or ANCHOR_TOP

    mAnchors.Add mLstInventory, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_BOTTOM
    mAnchors.Add mLstBom, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM

    mAnchors.Add mBtnAdd, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnRemove, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnSaveBox, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnUpdateVersion, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnNewVersion, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnArchiveBox, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnDeleteVersion, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnDeleteBox, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnCancel, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mLblStatus, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub

Private Function AddLabel(ByVal controlName As String, _
                          ByVal captionText As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal boldText As Boolean) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    With AddLabel
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .WordWrap = True
        .Font.Bold = boldText
    End With
End Function

Private Function AddCheckBox(ByVal controlName As String, _
                             ByVal captionText As String, _
                             ByVal leftPos As Single, _
                             ByVal topPos As Single, _
                             ByVal widthVal As Single, _
                             ByVal heightVal As Single) As MSForms.CheckBox
    Set AddCheckBox = Me.Controls.Add("Forms.CheckBox.1", controlName, True)
    With AddCheckBox
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddTextBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    With AddTextBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddComboBox(ByVal controlName As String, _
                             ByVal leftPos As Single, _
                             ByVal topPos As Single, _
                             ByVal widthVal As Single, _
                             ByVal heightVal As Single) As MSForms.ComboBox
    Set AddComboBox = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    With AddComboBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddListBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", controlName, True)
    With AddListBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddButton(ByVal controlName As String, _
                           ByVal captionText As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    With AddButton
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Sub ShowStatus(ByVal messageText As String)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = messageText
End Sub

Private Sub ShowInventoryStatus(ByVal messageText As String)
    If mLblInventoryStatus Is Nothing Then Exit Sub
    mLblInventoryStatus.Caption = messageText
End Sub

Private Function NzText(ByVal value As Variant) As String
    On Error GoTo UseBlank
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        NzText = ""
    Else
        NzText = Trim$(CStr(value))
    End If
    Exit Function
UseBlank:
    NzText = ""
End Function

Private Function ParseNumber(ByVal value As String) As Double
    On Error GoTo UseZero
    value = Trim$(value)
    If value = "" Then Exit Function
    ParseNumber = CDbl(value)
    Exit Function
UseZero:
    ParseNumber = 0#
End Function

Private Function NormalizeVersionText(ByVal versionText As String) As String
    versionText = LCase$(Trim$(versionText))
    If versionText = "" Then
        NormalizeVersionText = "v1"
    ElseIf Left$(versionText, 1) = "v" Then
        NormalizeVersionText = versionText
    Else
        NormalizeVersionText = "v" & versionText
    End If
End Function
