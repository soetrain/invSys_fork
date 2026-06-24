VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShippingBoxMaker
   Caption         =   "Shipping Box Maker"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShippingBoxMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mCboBoxes As MSForms.ComboBox
Private WithEvents mCboVersions As MSForms.ComboBox
Private WithEvents mTxtQty As MSForms.TextBox
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnMake As MSForms.CommandButton
Private WithEvents mBtnUnmake As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mTxtBoxPicker As MSForms.TextBox

Private mTxtBoxName As MSForms.TextBox
Private mTxtUom As MSForms.TextBox
Private mTxtLocation As MSForms.TextBox
Private mTxtDescription As MSForms.TextBox
Private WithEvents mLstShippables As MSForms.ListBox
Private mLstComponents As MSForms.ListBox
Private mLblPackageInv As MSForms.Label
Private mLblSyncState As MSForms.Label
Private mLblStatus As MSForms.Label

Private mSavedBoxes As Variant
Private mComponents As Variant
Private mShippableRows As Variant
Private mVersionComponentCache As Object
Private mPendingShippableInv As Object
Private mPendingComponentInv As Object
Private mPendingVersionInv As Object
Private mSelectedPackageRow As Long
Private mLoading As Boolean
Private mBuilt As Boolean
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
    On Error GoTo FailInit

    Dim previousScreenUpdating As Boolean
    Dim quietStarted As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    modUiQuiet.BeginQuietUi ActiveWorkbook
    quietStarted = True

    If Not mBuilt Then BuildLayout

    mLoading = True
    LoadSavedBoxes
    mLoading = False

    If mCboBoxes.ListCount > 0 Then
        mLoading = True
        mCboBoxes.ListIndex = 0
        mLoading = False
        LoadSelectedBox
    Else
        ClearBoxFields
        ShowStatus "No active saved box designs found."
    End If

CleanExit:
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Exit Sub

FailInit:
    ShowStatus "BoxMaker form load failed: " & Err.Description
    Resume CleanExit
End Sub

Private Sub BuildLayout()
    If mBuilt Then Exit Sub
    mBuilt = True

    Me.Caption = "Shipping Box Maker"
    Me.Width = 780
    Me.Height = 670

    AddLabel "lblTitle", "Box Maker", 12, 10, 140, 20, True
    AddLabel "lblBox", "Saved Box", 12, 42, 74, 18, False
    AddLabel "lblVersion", "Version", 432, 42, 58, 18, False
    AddLabel "lblQty", "Qty", 570, 42, 38, 18, False

    Set mCboBoxes = AddComboBox("cboBoxes", 86, 38, 320, 22)
    With mCboBoxes
        .ColumnCount = 7
        .ColumnWidths = "0 pt;150 pt;0 pt;42 pt;82 pt;130 pt;0 pt"
        .BoundColumn = 1
        .TextColumn = 2
        .Style = 2
    End With
    Set mCboVersions = AddComboBox("cboVersions", 492, 38, 64, 22)
    With mCboVersions
        .ColumnCount = 8
        .ColumnWidths = "42 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt"
        .Style = 2
    End With
    Set mTxtQty = AddTextBox("txtQty", 636, 38, 60, 22)
    mTxtQty.Value = "1"
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 706, 36, 58, 26)

    AddLabel "lblBoxName", "Box Name", 12, 76, 72, 18, False
    AddLabel "lblUom", "UOM", 302, 76, 38, 18, False
    AddLabel "lblLocation", "Location", 398, 76, 66, 18, False
    Set mTxtBoxName = AddTextBox("txtBoxName", 86, 72, 194, 22)
    Set mTxtUom = AddTextBox("txtUom", 340, 72, 42, 22)
    Set mTxtLocation = AddTextBox("txtLocation", 464, 72, 118, 22)
    AddLabel "lblDescription", "Description", 12, 108, 82, 18, False
    Set mTxtDescription = AddTextBox("txtDescription", 96, 104, 600, 22)

    LockTextBox mTxtBoxName
    LockTextBox mTxtUom
    LockTextBox mTxtLocation
    LockTextBox mTxtDescription

    AddLabel "lblShippables", "Shippable Inventory", 12, 142, 170, 18, True
    Set mLblPackageInv = AddLabel("lblPackageInv", "", 190, 142, 330, 18, False)
    Set mLblSyncState = AddLabel("lblSyncState", "", 522, 142, 242, 18, False)
    AddLabel "lblBoxPicker", "Box Picker", 12, 166, 72, 18, False
    Set mTxtBoxPicker = AddTextBox("txtBoxPicker", 86, 162, 320, 22)
    AddShippableHeaders 12, 194
    Set mLstShippables = AddListBox("lstShippables", 12, 214, 752, 72)
    With mLstShippables
        .ColumnCount = 7
        .ColumnWidths = "174 pt;54 pt;58 pt;72 pt;42 pt;110 pt;48 pt"
    End With

    AddLabel "lblComponents", "Components To Deduct", 12, 300, 170, 18, True
    AddComponentHeaders 12, 322
    Set mLstComponents = AddListBox("lstComponents", 12, 342, 752, 232)
    With mLstComponents
        .ColumnCount = 10
        .ColumnWidths = "112 pt;38 pt;46 pt;54 pt;50 pt;68 pt;34 pt;70 pt;58 pt;132 pt"
    End With

    Set mBtnMake = AddButton("btnMake", "Make Boxes", 506, 586, 86, 30)
    Set mBtnUnmake = AddButton("btnUnmake", "Unmake Boxes", 604, 586, 92, 30)
    Set mBtnClose = AddButton("btnClose", "Close", 706, 586, 58, 30)
    Set mLblStatus = AddLabel("lblStatus", "", 12, 584, 480, 44, False)
    InitializeBoxMakerAnchors
End Sub

Private Sub LoadSavedBoxes()
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long
    Dim c As Long
    Dim idx As Long

    rowsData = modTS_Shipments.BoxMakerFormLoadSavedBoxes()
    If IsEmpty(rowsData) Then
        If mCboBoxes.ListCount > 0 Then
            ShowStatus "Refresh did not return saved box designs; keeping current list."
        End If
        Exit Sub
    End If

    mSavedBoxes = rowsData
    mShippableRows = Empty
    mCboBoxes.Clear

    For r = 1 To UBound(mSavedBoxes, 1)
        mCboBoxes.AddItem NzText(mSavedBoxes(r, 1))
        idx = mCboBoxes.ListCount - 1
        For c = 2 To 7
            mCboBoxes.List(idx, c - 1) = NzText(mSavedBoxes(r, c))
        Next c
    Next r
    RenderShippableInventory
    ShowStatus "Loaded " & CStr(mCboBoxes.ListCount) & " active box design(s)."
    Exit Sub

FailSoft:
    ShowStatus "Could not load saved boxes: " & Err.Description
End Sub

Private Sub LoadSelectedBox()
    On Error GoTo FailSoft

    Dim previousLoading As Boolean

    If mCboBoxes.ListIndex < 0 Then Exit Sub
    previousLoading = mLoading
    mLoading = True
    mSelectedPackageRow = CLng(Val(CStr(mCboBoxes.List(mCboBoxes.ListIndex, 0))))
    mTxtBoxName.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 1))
    mTxtUom.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 3))
    mTxtLocation.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 4))
    mTxtDescription.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 5))
    Set mVersionComponentCache = CreateObject("Scripting.Dictionary")
    mVersionComponentCache.CompareMode = vbTextCompare
    LoadVersionsForPackage mSelectedPackageRow
    mLoading = previousLoading
    LoadSelectedVersionComponents
    RenderPackageInventory
    SelectShippableInventoryRow
    Exit Sub

FailSoft:
    mLoading = previousLoading
    ShowStatus "Could not load selected box: " & Err.Description
End Sub

Private Sub LoadVersionsForPackage(ByVal packageRow As Long)
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long
    Dim c As Long
    Dim idx As Long

    mCboVersions.Clear
    If LoadVersionsForPackageFromCache(packageRow) Then Exit Sub

    rowsData = modTS_Shipments.BoxMakerFormLoadVersions(packageRow)
    If IsEmpty(rowsData) Then
        ShowStatus "No active versions found for selected box."
        Exit Sub
    End If

    For r = 1 To UBound(rowsData, 1)
        If UCase$(NzText(rowsData(r, 2))) <> "ACTIVE" Then GoTo NextRow
        mCboVersions.AddItem NzText(rowsData(r, 1))
        idx = mCboVersions.ListCount - 1
        For c = 2 To 8
            mCboVersions.List(idx, c - 1) = NzText(rowsData(r, c))
        Next c
NextRow:
    Next r
    If mCboVersions.ListCount > 0 Then mCboVersions.ListIndex = 0
    Exit Sub

FailSoft:
    ShowStatus "Could not load versions: " & Err.Description
End Sub

Private Function LoadVersionsForPackageFromCache(ByVal packageRow As Long) As Boolean
    On Error GoTo FailSoft

    Dim seen As Object
    Dim r As Long
    Dim idx As Long
    Dim versionLabel As String

    If packageRow <= 0 Then Exit Function
    If IsEmpty(mShippableRows) Then Exit Function

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For r = 1 To UBound(mShippableRows, 1)
        If CLng(Val(NzText(mShippableRows(r, 1)))) <> packageRow Then GoTo NextRow
        versionLabel = NormalizeVersionText(NzText(mShippableRows(r, 3)))
        If versionLabel = "" Then GoTo NextRow
        If seen.Exists(versionLabel) Then GoTo NextRow
        seen(versionLabel) = True

        mCboVersions.AddItem versionLabel
        idx = mCboVersions.ListCount - 1
        mCboVersions.List(idx, 1) = "Active"
NextRow:
    Next r

    If mCboVersions.ListCount > 0 Then
        mCboVersions.ListIndex = 0
        LoadVersionsForPackageFromCache = True
    End If
    Exit Function

FailSoft:
End Function

Private Sub LoadSelectedVersionComponents()
    On Error GoTo FailSoft

    Dim versionLabel As String

    If mSelectedPackageRow <= 0 Then Exit Sub
    If mCboVersions.ListIndex < 0 Then Exit Sub

    versionLabel = SelectedVersionLabel()
    mComponents = CachedVersionComponents(versionLabel)
    If IsEmpty(mComponents) Then
        mComponents = modTS_Shipments.BoxMakerFormLoadVersionComponents(mSelectedPackageRow, versionLabel)
        If Not IsEmpty(mComponents) Then
            If mVersionComponentCache Is Nothing Then
                Set mVersionComponentCache = CreateObject("Scripting.Dictionary")
                mVersionComponentCache.CompareMode = vbTextCompare
            End If
            mVersionComponentCache(NormalizeVersionText(versionLabel)) = mComponents
        End If
    End If
    RenderComponents
    Exit Sub

FailSoft:
    ShowStatus "Could not load components: " & Err.Description
End Sub

Private Sub RenderComponents()
    On Error GoTo FailSoft

    Dim r As Long
    Dim idx As Long
    Dim qtyMade As Double
    Dim perBoxQty As Double
    Dim requiredQty As Double
    Dim nasInv As String
    Dim projectedInv As String
    Dim rowValue As Long

    mLstComponents.Clear
    If IsEmpty(mComponents) Then
        ShowStatus "Selected version has no components."
        Exit Sub
    End If

    qtyMade = ParseNumber(Trim$(CStr(mTxtQty.Value)))
    If qtyMade <= 0 Then qtyMade = 0

    For r = 1 To UBound(mComponents, 1)
        perBoxQty = ParseNumber(NzText(mComponents(r, 5)))
        requiredQty = perBoxQty * qtyMade
        rowValue = CLng(Val(NzText(mComponents(r, 4))))
        nasInv = NzText(mComponents(r, 9))
        projectedInv = DisplayComponentInventoryText(rowValue, nasInv)

        mLstComponents.AddItem NzText(mComponents(r, 2))
        idx = mLstComponents.ListCount - 1
        mLstComponents.List(idx, 1) = NzText(mComponents(r, 4))
        mLstComponents.List(idx, 2) = FormatQuantityText(perBoxQty)
        mLstComponents.List(idx, 3) = FormatQuantityText(requiredQty)
        If nasInv = "" Then
            mLstComponents.List(idx, 4) = "unknown"
        Else
            mLstComponents.List(idx, 4) = nasInv
        End If
        If projectedInv = "" Then
            mLstComponents.List(idx, 5) = "unknown"
        Else
            mLstComponents.List(idx, 5) = projectedInv
        End If
        mLstComponents.List(idx, 6) = NzText(mComponents(r, 6))
        mLstComponents.List(idx, 7) = NzText(mComponents(r, 7))
        mLstComponents.List(idx, 8) = NzText(mComponents(r, 3))
        mLstComponents.List(idx, 9) = NzText(mComponents(r, 8))
    Next r
    ShowStatus "Loaded " & CStr(mLstComponents.ListCount) & " component row(s) for " & SelectedVersionLabel() & "."
    Exit Sub

FailSoft:
    ShowStatus "Component render failed: " & Err.Description
End Sub

Private Function CachedVersionComponents(ByVal versionLabel As String) As Variant
    versionLabel = NormalizeVersionText(versionLabel)
    If versionLabel = "" Then Exit Function
    If mVersionComponentCache Is Nothing Then Exit Function
    If Not mVersionComponentCache.Exists(versionLabel) Then Exit Function
    CachedVersionComponents = mVersionComponentCache(versionLabel)
End Function

Private Sub SelectBoxMakerVersion(ByVal versionLabel As String)
    On Error GoTo CleanExit

    Dim r As Long
    Dim previousLoading As Boolean

    previousLoading = mLoading
    versionLabel = NormalizeVersionText(versionLabel)
    If versionLabel = "" Then Exit Sub
    If mCboVersions Is Nothing Then Exit Sub

    For r = 0 To mCboVersions.ListCount - 1
        If StrComp(NormalizeVersionText(NzText(mCboVersions.List(r, 0))), versionLabel, vbTextCompare) = 0 Then
            mLoading = True
            If mCboVersions.ListIndex <> r Then mCboVersions.ListIndex = r
            mLoading = previousLoading
            LoadSelectedVersionComponents
            SelectShippableInventoryRow
            Exit For
        End If
    Next r

CleanExit:
    mLoading = previousLoading
End Sub

Private Sub RenderPackageInventory()
    On Error GoTo FailSoft

    Dim currentInv As Variant

    If mSelectedPackageRow <= 0 Then
        mLblPackageInv.Caption = ""
        Exit Sub
    End If
    currentInv = modTS_Shipments.BoxMakerFormCurrentInventory(mSelectedPackageRow, CStr(mTxtBoxName.Value))
    SetPackageInventoryCaption NzText(currentInv)
    Exit Sub

FailSoft:
    mLblPackageInv.Caption = ""
End Sub

Private Sub RenderPackageInventoryFromCache()
    SetPackageInventoryCaption CachedShippableInventoryText(mSelectedPackageRow)
End Sub

Private Function CachedShippableInventoryText(ByVal packageRow As Long) As String
    On Error GoTo CleanExit

    Dim r As Long
    Dim versionLabel As String

    If packageRow <= 0 Then Exit Function
    If IsEmpty(mShippableRows) Then Exit Function
    versionLabel = SelectedVersionLabel()
    For r = 1 To UBound(mShippableRows, 1)
        If CLng(Val(NzText(mShippableRows(r, 1)))) = packageRow _
           And StrComp(NormalizeVersionText(NzText(mShippableRows(r, 3))), versionLabel, vbTextCompare) = 0 Then
            CachedShippableInventoryText = NzText(mShippableRows(r, 4))
            Exit Function
        End If
    Next r

CleanExit:
End Function

Private Sub SetPackageInventoryCaption(ByVal backendInventoryText As String)
    Dim nasInv As String
    Dim projectedInv As String

    If mLblPackageInv Is Nothing Then Exit Sub
    If mSelectedPackageRow <= 0 Then
        mLblPackageInv.Caption = ""
        Exit Sub
    End If
    nasInv = Trim$(backendInventoryText)
    If nasInv = "" Then nasInv = "unknown"
    projectedInv = DisplayBoxVersionInventoryText(mSelectedPackageRow, SelectedVersionLabel(), backendInventoryText)
    If projectedInv = "" Then projectedInv = "unknown"
    mLblPackageInv.Caption = "NAS Inv: " & nasInv & "    Projected Inv: " & projectedInv
End Sub

Private Sub mCboBoxes_Change()
    If mLoading Then Exit Sub
    LoadSelectedBox
End Sub

Private Sub mCboVersions_Change()
    If mLoading Then Exit Sub
    LoadSelectedVersionComponents
    SelectShippableInventoryRow
End Sub

Private Sub mTxtQty_Change()
    RenderComponents
End Sub

Private Sub mTxtBoxPicker_Change()
    If mLoading Then Exit Sub
    RenderShippableInventoryFromCache
End Sub

Private Sub mBtnRefresh_Click()
    On Error GoTo FailSoft

    Dim previousPointer As Long
    Dim previousScreenUpdating As Boolean

    previousPointer = Me.MousePointer
    previousScreenUpdating = Application.ScreenUpdating
    Me.MousePointer = fmMousePointerHourGlass
    Application.ScreenUpdating = False
    ShowStatus "Refreshing BoxMaker inventory..."

    RefreshShippableInventoryCache True
    RenderShippableInventoryFromCache
    If mSelectedPackageRow > 0 And Not mCboVersions Is Nothing Then
        If mCboVersions.ListIndex >= 0 Then LoadSelectedVersionComponents
    End If
    RenderPackageInventory

CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = previousScreenUpdating
    Me.MousePointer = previousPointer
    On Error GoTo 0
    If Err.Number = 0 Then ShowStatus "BoxMaker inventory refreshed."
    UpdateSyncStateLabel
    Exit Sub

FailSoft:
    ShowStatus "BoxMaker refresh failed: " & Err.Description
    UpdateSyncStateLabel
    Resume CleanExit
End Sub

Private Sub mBtnMake_Click()
    PostBoxMakerAction "MAKE"
End Sub

Private Sub mBtnUnmake_Click()
    PostBoxMakerAction "UNMAKE"
End Sub

Private Sub mBtnClose_Click()
    Me.Hide
End Sub

Private Sub mLstShippables_Click()
    On Error GoTo FailSoft

    Dim packageRow As Long
    Dim versionLabel As String
    Dim r As Long
    Dim previousLoading As Boolean

    If mLoading Then Exit Sub
    If mLstShippables.ListIndex < 0 Then Exit Sub

    packageRow = CLng(Val(NzText(mLstShippables.List(mLstShippables.ListIndex, 6))))
    versionLabel = NormalizeVersionText(NzText(mLstShippables.List(mLstShippables.ListIndex, 1)))
    If packageRow <= 0 Then Exit Sub
    If packageRow = mSelectedPackageRow _
       And StrComp(versionLabel, SelectedVersionLabel(), vbTextCompare) = 0 Then
        Exit Sub
    End If

    If packageRow <> mSelectedPackageRow Then
        For r = 0 To mCboBoxes.ListCount - 1
            If CLng(Val(NzText(mCboBoxes.List(r, 0)))) = packageRow Then
                previousLoading = mLoading
                mLoading = True
                mCboBoxes.ListIndex = r
                mLoading = previousLoading
                LoadSelectedBox
                Exit For
            End If
        Next r
    End If

    SelectBoxMakerVersion versionLabel
    Exit Sub

FailSoft:
    mLoading = False
End Sub

Private Sub PostBoxMakerAction(ByVal actionText As String)
    On Error GoTo ErrHandler

    Dim qtyMade As Double
    Dim resultMessage As String
    Dim startedAt As Single
    Dim elapsedMs As Long
    Dim postedOk As Boolean
    Dim syncCompleted As Boolean
    Dim previousPointer As Long
    Dim previousScreenUpdating As Boolean
    Dim quietStarted As Boolean
    Dim selectedNasInv As String

    qtyMade = ParseNumber(Trim$(CStr(mTxtQty.Value)))
    If qtyMade <= 0 Then
        ShowStatus "Enter a positive box quantity."
        Exit Sub
    End If
    If mSelectedPackageRow <= 0 Or mCboVersions.ListIndex < 0 Then
        ShowStatus "Select a saved box and active version."
        Exit Sub
    End If
    If IsEmpty(mComponents) Then
        ShowStatus "Selected version has no components."
        Exit Sub
    End If

    previousPointer = Me.MousePointer
    previousScreenUpdating = Application.ScreenUpdating
    Me.MousePointer = fmMousePointerHourGlass
    ShowStatus "Posting BoxMaker action..."
    DoEvents
    Application.ScreenUpdating = False
    modUiQuiet.BeginQuietUi ActiveWorkbook
    quietStarted = True
    startedAt = Timer
    selectedNasInv = CachedShippableInventoryText(mSelectedPackageRow)

    postedOk = modTS_Shipments.CommitBoxMakerFormAction(mSelectedPackageRow, _
                                                        CStr(mTxtBoxName.Value), _
                                                        CStr(mTxtUom.Value), _
                                                        CStr(mTxtLocation.Value), _
                                                        CStr(mTxtDescription.Value), _
                                                        SelectedVersionLabel(), _
                                                        qtyMade, _
                                                        mComponents, _
                                                        resultMessage, _
                                                        actionText, _
                                                        syncCompleted, _
                                                        selectedNasInv)
    elapsedMs = ElapsedMillisecondsForm(startedAt)
    If quietStarted Then
        modUiQuiet.EndQuietUi
        quietStarted = False
    End If
    Application.ScreenUpdating = previousScreenUpdating
    Me.MousePointer = previousPointer

    If postedOk Then
        RecordPendingComponentInventory actionText, qtyMade
        RecordPendingVersionInventory actionText, qtyMade
        RefreshShippableInventoryCache True
        mTxtQty.Value = ""
        resultMessage = AppendCompletionTiming(resultMessage, elapsedMs)
        MsgBox resultMessage, vbInformation
        ShowStatus "Completed in " & Format$(elapsedMs, "#,##0") & " ms."
        RenderComponents
        RenderPackageInventoryFromCache
        RenderShippableInventoryFromCache
        UpdateSyncStateLabel
    Else
        If resultMessage = "" Then resultMessage = "BoxMaker action did not complete."
        resultMessage = AppendCompletionTiming(resultMessage, elapsedMs)
        MsgBox resultMessage, vbExclamation
        ShowStatus resultMessage
        UpdateSyncStateLabel
    End If
    Exit Sub

ErrHandler:
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    Application.ScreenUpdating = previousScreenUpdating
    Me.MousePointer = previousPointer
    On Error GoTo 0
    ShowStatus "BoxMaker action failed: " & Err.Description
End Sub

Private Function AppendCompletionTiming(ByVal messageText As String, ByVal elapsedMs As Long) As String
    If Trim$(messageText) <> "" Then
        AppendCompletionTiming = messageText & vbCrLf & vbCrLf
    End If
    AppendCompletionTiming = AppendCompletionTiming & "Completed in " & Format$(elapsedMs, "#,##0") & " ms."
End Function

Private Function ElapsedMillisecondsForm(ByVal startedAt As Single) As Long
    Dim deltaSeconds As Single

    deltaSeconds = Timer - startedAt
    If deltaSeconds < 0 Then deltaSeconds = deltaSeconds + 86400!
    ElapsedMillisecondsForm = CLng(deltaSeconds * 1000)
End Function

Private Sub RenderShippableInventory()
    If RefreshShippableInventoryCache(False) Then
        RenderShippableInventoryFromCache
    Else
        If Not mLstShippables Is Nothing Then mLstShippables.Clear
        UpdateSyncStateLabel
    End If
End Sub

Private Function RefreshShippableInventoryCache(Optional ByVal forceReload As Boolean = False) As Boolean
    On Error GoTo FailSoft

    If IsEmpty(mSavedBoxes) Then
        mShippableRows = Empty
        Exit Function
    End If

    If forceReload Or IsEmpty(mShippableRows) Then
        mShippableRows = modTS_Shipments.BoxMakerFormLoadShippableVersionInventory(mSavedBoxes)
    End If

    RefreshShippableInventoryCache = Not IsEmpty(mShippableRows)
    Exit Function

FailSoft:
    mShippableRows = Empty
    RefreshShippableInventoryCache = False
End Function

Private Sub RenderShippableInventoryFromCache()
    On Error GoTo FailSoft

    Dim r As Long
    Dim idx As Long
    Dim nasInv As String
    Dim projectedInv As String
    Dim rowValue As Long
    Dim prevLoading As Boolean
    Dim filterText As String
    Dim shownCount As Long
    Dim displayRows As Variant
    Dim selectedIndex As Long

    If mLstShippables Is Nothing Then Exit Sub
    prevLoading = mLoading
    mLoading = True
    mLstShippables.Clear
    If IsEmpty(mShippableRows) Then GoTo CleanExit

    filterText = BoxPickerText()
    For r = 1 To UBound(mShippableRows, 1)
        If ShippableRowMatchesPicker(mShippableRows, r, filterText) Then shownCount = shownCount + 1
    Next r
    If shownCount = 0 Then
        If filterText <> "" Then ShowStatus "No shippable boxes match picker."
        GoTo CleanExit
    End If

    ReDim displayRows(0 To shownCount - 1, 0 To 6)
    selectedIndex = -1
    idx = 0
    For r = 1 To UBound(mShippableRows, 1)
        If Not ShippableRowMatchesPicker(mShippableRows, r, filterText) Then GoTo NextShippable
        rowValue = CLng(Val(NzText(mShippableRows(r, 1))))
        nasInv = NzText(mShippableRows(r, 4))
        If nasInv = "" Then nasInv = "unknown"
        projectedInv = DisplayBoxVersionInventoryText(rowValue, NzText(mShippableRows(r, 3)), NzText(mShippableRows(r, 4)))
        If projectedInv = "" Then projectedInv = "unknown"

        displayRows(idx, 0) = NzText(mShippableRows(r, 2))
        displayRows(idx, 1) = NzText(mShippableRows(r, 3))
        displayRows(idx, 2) = nasInv
        displayRows(idx, 3) = projectedInv
        displayRows(idx, 4) = NzText(mShippableRows(r, 5))
        displayRows(idx, 5) = NzText(mShippableRows(r, 6))
        displayRows(idx, 6) = NzText(mShippableRows(r, 1))
        If rowValue = mSelectedPackageRow _
           And StrComp(NormalizeVersionText(NzText(mShippableRows(r, 3))), SelectedVersionLabel(), vbTextCompare) = 0 Then
            selectedIndex = idx
        End If
        idx = idx + 1
NextShippable:
    Next r
    mLstShippables.List = displayRows
    If selectedIndex >= 0 Then mLstShippables.ListIndex = selectedIndex

CleanExit:
    mLoading = prevLoading
    UpdateSyncStateLabel
    Exit Sub

FailSoft:
    If Not mLstShippables Is Nothing Then mLstShippables.Clear
    mLoading = prevLoading
    UpdateSyncStateLabel
End Sub

Private Sub RecordPendingShippableInventory(ByVal actionText As String, ByVal qtyMade As Double)
    On Error GoTo CleanExit

    Dim key As String
    Dim currentText As String
    Dim projectedQty As Double
    Dim r As Long

    If mSelectedPackageRow <= 0 Or qtyMade <= 0 Then Exit Sub
    If mPendingShippableInv Is Nothing Then Set mPendingShippableInv = CreateObject("Scripting.Dictionary")

    currentText = ""
    If Not mLstShippables Is Nothing Then
        For r = 0 To mLstShippables.ListCount - 1
            If CLng(Val(NzText(mLstShippables.List(r, 6)))) = mSelectedPackageRow _
               And StrComp(NormalizeVersionText(NzText(mLstShippables.List(r, 1))), SelectedVersionLabel(), vbTextCompare) = 0 Then
                currentText = NzText(mLstShippables.List(r, 3))
                Exit For
            End If
        Next r
    End If
    If LCase$(Trim$(currentText)) = "unknown" Then currentText = ""
    projectedQty = ParseNumber(currentText)
    Select Case UCase$(Trim$(actionText))
        Case "UNMAKE", "UNBOX"
            projectedQty = projectedQty - qtyMade
            If projectedQty < 0 Then projectedQty = 0
        Case Else
            projectedQty = projectedQty + qtyMade
    End Select

    key = CStr(mSelectedPackageRow)
    mPendingShippableInv(key) = projectedQty
    UpdateSyncStateLabel

CleanExit:
End Sub

Private Function DisplayShippableInventoryText(ByVal rowValue As Long, ByVal backendText As String) As String
    On Error GoTo CleanExit

    Dim key As String
    Dim pendingQty As Double
    Dim backendQty As Double

    DisplayShippableInventoryText = Trim$(backendText)
    If rowValue <= 0 Then Exit Function
    If mPendingShippableInv Is Nothing Then Exit Function

    key = CStr(rowValue)
    If Not mPendingShippableInv.Exists(key) Then Exit Function

    pendingQty = CDbl(mPendingShippableInv(key))
    If Trim$(backendText) <> "" And IsNumeric(Replace$(backendText, ",", "")) Then
        backendQty = CDbl(Replace$(backendText, ",", ""))
        If Abs(backendQty - pendingQty) < 0.0000001 Then
            mPendingShippableInv.Remove key
            DisplayShippableInventoryText = FormatQuantityText(backendQty)
            Exit Function
        End If
    End If

    DisplayShippableInventoryText = FormatQuantityText(pendingQty)

CleanExit:
End Function

Private Sub RecordPendingComponentInventory(ByVal actionText As String, ByVal qtyMade As Double)
    On Error GoTo CleanExit

    Dim r As Long
    Dim rowValue As Long
    Dim perBoxQty As Double
    Dim requiredQty As Double
    Dim currentText As String
    Dim projectedQty As Double
    Dim key As String

    If qtyMade <= 0 Then Exit Sub
    If IsEmpty(mComponents) Then Exit Sub
    If mPendingComponentInv Is Nothing Then Set mPendingComponentInv = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(mComponents, 1)
        rowValue = CLng(Val(NzText(mComponents(r, 4))))
        perBoxQty = ParseNumber(NzText(mComponents(r, 5)))
        requiredQty = perBoxQty * qtyMade
        If rowValue <= 0 Or requiredQty <= 0 Then GoTo NextComponent

        currentText = DisplayComponentInventoryText(rowValue, NzText(mComponents(r, 9)))
        If currentText = "" Or Not IsNumeric(Replace$(currentText, ",", "")) Then GoTo NextComponent

        projectedQty = CDbl(Replace$(currentText, ",", ""))
        Select Case UCase$(Trim$(actionText))
            Case "UNMAKE", "UNBOX"
                projectedQty = projectedQty + requiredQty
            Case Else
                projectedQty = projectedQty - requiredQty
                If projectedQty < 0 Then projectedQty = 0
        End Select

        key = CStr(rowValue)
        mPendingComponentInv(key) = projectedQty
NextComponent:
    Next r

CleanExit:
End Sub

Private Function DisplayComponentInventoryText(ByVal rowValue As Long, ByVal backendText As String) As String
    On Error GoTo CleanExit

    Dim key As String
    Dim pendingQty As Double
    Dim backendQty As Double

    DisplayComponentInventoryText = Trim$(backendText)
    If rowValue <= 0 Then Exit Function
    If mPendingComponentInv Is Nothing Then Exit Function

    key = CStr(rowValue)
    If Not mPendingComponentInv.Exists(key) Then Exit Function

    pendingQty = CDbl(mPendingComponentInv(key))
    If Trim$(backendText) <> "" And IsNumeric(backendText) Then
        backendQty = CDbl(backendText)
        If Abs(backendQty - pendingQty) < 0.0000001 Then
            mPendingComponentInv.Remove key
            DisplayComponentInventoryText = FormatQuantityText(backendQty)
            Exit Function
        End If
    End If

    DisplayComponentInventoryText = FormatQuantityText(pendingQty)

CleanExit:
End Function

Private Sub RecordPendingVersionInventory(ByVal actionText As String, ByVal qtyMade As Double)
    On Error GoTo CleanExit

    Dim versionLabel As String
    Dim currentText As String
    Dim projectedQty As Double
    Dim baselineQty As Double

    If qtyMade <= 0 Then Exit Sub
    versionLabel = SelectedVersionLabel()
    If versionLabel = "" Then Exit Sub
    If mPendingVersionInv Is Nothing Then Set mPendingVersionInv = CreateObject("Scripting.Dictionary")

    currentText = SelectedShippableVersionInventoryText()
    If currentText = "" Or LCase$(currentText) = "unknown" Or Not IsNumeric(Replace$(currentText, ",", "")) Then
        projectedQty = 0
    Else
        projectedQty = CDbl(Replace$(currentText, ",", ""))
    End If
    baselineQty = projectedQty

    Select Case UCase$(Trim$(actionText))
        Case "UNMAKE", "UNBOX"
            projectedQty = projectedQty - qtyMade
            If projectedQty < 0 Then projectedQty = 0
        Case Else
            projectedQty = projectedQty + qtyMade
    End Select

    mPendingVersionInv(VersionPendingKey(mSelectedPackageRow, versionLabel)) = projectedQty
    modTS_Shipments.RegisterPendingBoxVersionInventoryOverlay mSelectedPackageRow, versionLabel, projectedQty, baselineQty

CleanExit:
End Sub

Private Function SelectedShippableVersionInventoryText() As String
    On Error GoTo CleanExit

    Dim r As Long
    Dim versionLabel As String

    If mLstShippables Is Nothing Then Exit Function
    If mSelectedPackageRow <= 0 Then Exit Function
    versionLabel = SelectedVersionLabel()
    For r = 0 To mLstShippables.ListCount - 1
        If CLng(Val(NzText(mLstShippables.List(r, 6)))) = mSelectedPackageRow _
           And StrComp(NormalizeVersionText(NzText(mLstShippables.List(r, 1))), versionLabel, vbTextCompare) = 0 Then
            SelectedShippableVersionInventoryText = NzText(mLstShippables.List(r, 3))
            Exit Function
        End If
    Next r

CleanExit:
End Function

Private Function DisplayBoxVersionInventoryText(ByVal packageRow As Long, _
                                                ByVal versionLabel As String, _
                                                ByVal backendText As String) As String
    On Error GoTo CleanExit

    Dim key As String
    Dim pendingQty As Double
    Dim backendQty As Double

    versionLabel = NormalizeVersionText(versionLabel)
    DisplayBoxVersionInventoryText = Trim$(backendText)
    key = VersionPendingKey(packageRow, versionLabel)
    If key = "" Then Exit Function
    If versionLabel = "" Then Exit Function
    If mPendingVersionInv Is Nothing Then
        DisplayBoxVersionInventoryText = modTS_Shipments.PendingBoxVersionInventoryOverlayText(packageRow, versionLabel, backendText)
        Exit Function
    End If
    If Not mPendingVersionInv.Exists(key) Then
        DisplayBoxVersionInventoryText = modTS_Shipments.PendingBoxVersionInventoryOverlayText(packageRow, versionLabel, backendText)
        Exit Function
    End If

    pendingQty = CDbl(mPendingVersionInv(key))
    If Trim$(backendText) <> "" And IsNumeric(Replace$(backendText, ",", "")) Then
        backendQty = CDbl(Replace$(backendText, ",", ""))
        If Abs(backendQty - pendingQty) < 0.0000001 Then
            mPendingVersionInv.Remove key
            DisplayBoxVersionInventoryText = FormatQuantityText(backendQty)
            Exit Function
        End If
    End If

    DisplayBoxVersionInventoryText = FormatQuantityText(pendingQty)

CleanExit:
End Function

Private Function VersionPendingKey(ByVal packageRow As Long, ByVal versionLabel As String) As String
    versionLabel = NormalizeVersionText(versionLabel)
    If packageRow <= 0 Or versionLabel = "" Then Exit Function
    VersionPendingKey = CStr(packageRow) & "|" & versionLabel
End Function

Private Function FormatQuantityText(ByVal qtyValue As Double) As String
    If Abs(qtyValue - Fix(qtyValue)) < 0.0000001 Then
        FormatQuantityText = Format$(qtyValue, "0")
    Else
        FormatQuantityText = Format$(qtyValue, "0.###")
    End If
End Function

Private Function BoxPickerText() As String
    If mTxtBoxPicker Is Nothing Then Exit Function
    BoxPickerText = LCase$(Trim$(NzText(mTxtBoxPicker.Value)))
End Function

Private Function ShippableRowMatchesPicker(ByVal rowsData As Variant, _
                                           ByVal rowIndex As Long, _
                                           ByVal filterText As String) As Boolean
    If filterText = "" Then
        ShippableRowMatchesPicker = True
        Exit Function
    End If

    ShippableRowMatchesPicker = TextContainsPicker(rowsData(rowIndex, 1), filterText) _
                                Or TextContainsPicker(rowsData(rowIndex, 2), filterText) _
                                Or TextContainsPicker(rowsData(rowIndex, 3), filterText) _
                                Or TextContainsPicker(rowsData(rowIndex, 5), filterText) _
                                Or TextContainsPicker(rowsData(rowIndex, 6), filterText)
End Function

Private Function TextContainsPicker(ByVal value As Variant, ByVal filterText As String) As Boolean
    If filterText = "" Then
        TextContainsPicker = True
    Else
        TextContainsPicker = (InStr(1, LCase$(Trim$(NzText(value))), filterText, vbTextCompare) > 0)
    End If
End Function

Private Sub UpdateSyncStateLabel()
    On Error GoTo CleanExit

    Dim pendingCount As Long

    If mLblSyncState Is Nothing Then Exit Sub
    pendingCount = PendingShippableInventoryCount()
    If pendingCount > 0 Then
        mLblSyncState.Caption = "Sync: pending (" & CStr(pendingCount) & " inventory row(s))"
        mLblSyncState.ForeColor = &H80&
    Else
        mLblSyncState.Caption = "Sync: inventory synced"
        mLblSyncState.ForeColor = &H8000&
    End If

CleanExit:
End Sub

Private Function PendingShippableInventoryCount() As Long
    On Error GoTo CleanExit

    If Not mPendingComponentInv Is Nothing Then PendingShippableInventoryCount = PendingShippableInventoryCount + mPendingComponentInv.Count
    If Not mPendingVersionInv Is Nothing Then PendingShippableInventoryCount = PendingShippableInventoryCount + mPendingVersionInv.Count

CleanExit:
End Function

Private Sub SelectShippableInventoryRow()
    On Error GoTo FailSoft

    Dim r As Long
    Dim versionLabel As String

    If mLstShippables Is Nothing Then Exit Sub
    If mSelectedPackageRow <= 0 Then Exit Sub
    versionLabel = SelectedVersionLabel()

    For r = 0 To mLstShippables.ListCount - 1
        If CLng(Val(NzText(mLstShippables.List(r, 6)))) = mSelectedPackageRow _
           And StrComp(NormalizeVersionText(NzText(mLstShippables.List(r, 1))), versionLabel, vbTextCompare) = 0 Then
            mLstShippables.ListIndex = r
            Exit Sub
        End If
    Next r
    Exit Sub

FailSoft:
End Sub

Private Function SelectedVersionLabel() As String
    If Not mCboVersions Is Nothing Then
        If mCboVersions.ListIndex >= 0 Then SelectedVersionLabel = NzText(mCboVersions.List(mCboVersions.ListIndex, 0))
    End If
    If SelectedVersionLabel = "" Then SelectedVersionLabel = "v1"
    SelectedVersionLabel = NormalizeVersionText(SelectedVersionLabel)
End Function

Private Sub ClearBoxFields()
    mSelectedPackageRow = 0
    mCboVersions.Clear
    mTxtBoxName.Value = ""
    mTxtUom.Value = ""
    mTxtLocation.Value = ""
    mTxtDescription.Value = ""
    mLstComponents.Clear
    If Not mLstShippables Is Nothing Then mLstShippables.Clear
    Set mVersionComponentCache = Nothing
    mLblPackageInv.Caption = ""
    UpdateSyncStateLabel
End Sub

Private Sub InitializeBoxMakerAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me

    mAnchors.Add mCboBoxes, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mBtnRefresh, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtDescription, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblPackageInv, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mLblSyncState, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtBoxPicker, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT

    mAnchors.Add mLstShippables, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstComponents, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM

    mAnchors.Add mLblStatus, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnMake, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnUnmake, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnClose, ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub

Private Function AddLabel(ByVal name As String, _
                          ByVal caption As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal boldText As Boolean) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", name, True)
    With AddLabel
        .Caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = boldText
    End With
End Function

Private Sub AddComponentHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrItem", "Item", leftPos, topPos, 110
    AddHeaderLabel "hdrRow", "ROW", leftPos + 114, topPos, 36
    AddHeaderLabel "hdrPerBox", "Per Box", leftPos + 154, topPos, 44
    AddHeaderLabel "hdrRequired", "Required", leftPos + 204, topPos, 54
    AddHeaderLabel "hdrNas", "NAS Inv", leftPos + 264, topPos, 50
    AddHeaderLabel "hdrProjected", "Projected Inv", leftPos + 318, topPos, 72
    AddHeaderLabel "hdrUom", "UOM", leftPos + 394, topPos, 34
    AddHeaderLabel "hdrLocation", "Location", leftPos + 432, topPos, 68
    AddHeaderLabel "hdrCode", "Code", leftPos + 506, topPos, 56
    AddHeaderLabel "hdrDesc", "Description", leftPos + 568, topPos, 132
End Sub

Private Sub AddShippableHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrShipBox", "Box", leftPos, topPos, 168
    AddHeaderLabel "hdrShipVersion", "Version", leftPos + 176, topPos, 58
    AddHeaderLabel "hdrShipNas", "NAS Inv", leftPos + 238, topPos, 56
    AddHeaderLabel "hdrShipProjected", "Projected Inv", leftPos + 298, topPos, 76
    AddHeaderLabel "hdrShipUom", "UOM", leftPos + 378, topPos, 42
    AddHeaderLabel "hdrShipLocation", "Location", leftPos + 426, topPos, 110
    AddHeaderLabel "hdrShipRow", "ROW", leftPos + 542, topPos, 56
End Sub

Private Sub AddHeaderLabel(ByVal name As String, _
                           ByVal caption As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single)
    Dim lbl As MSForms.Label

    Set lbl = AddLabel(name, caption, leftPos, topPos, widthVal, 14, True)
    lbl.Font.Size = 8
End Sub

Private Function AddTextBox(ByVal name As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", name, True)
    With AddTextBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddComboBox(ByVal name As String, _
                             ByVal leftPos As Single, _
                             ByVal topPos As Single, _
                             ByVal widthVal As Single, _
                             ByVal heightVal As Single) As MSForms.ComboBox
    Set AddComboBox = Me.Controls.Add("Forms.ComboBox.1", name, True)
    With AddComboBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddListBox(ByVal name As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", name, True)
    With AddListBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddButton(ByVal name As String, _
                           ByVal caption As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", name, True)
    With AddButton
        .Caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Sub LockTextBox(ByVal txt As MSForms.TextBox)
    If txt Is Nothing Then Exit Sub
    txt.Locked = True
    txt.BackColor = &H8000000F
End Sub

Private Sub ShowStatus(ByVal message As String)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = message
End Sub

Private Function NzText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        NzText = ""
    Else
        NzText = CStr(value)
    End If
End Function

Private Function ParseNumber(ByVal textValue As String) As Double
    On Error GoTo UseZero
    textValue = Trim$(textValue)
    If textValue = "" Then Exit Function
    ParseNumber = CDbl(textValue)
    Exit Function
UseZero:
    ParseNumber = 0#
End Function

Private Function NormalizeVersionText(ByVal valueIn As String) As String
    valueIn = Trim$(valueIn)
    If valueIn = "" Then
        NormalizeVersionText = "v1"
    ElseIf LCase$(Left$(valueIn, 1)) = "v" Then
        NormalizeVersionText = "v" & CStr(Val(Mid$(valueIn, 2)))
    ElseIf IsNumeric(valueIn) Then
        NormalizeVersionText = "v" & CStr(Val(valueIn))
    Else
        NormalizeVersionText = valueIn
    End If
End Function
