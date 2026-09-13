VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShipmentsTally
   Caption         =   "Shipping Shipments"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11880
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShipmentsTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtPicker As MSForms.TextBox
Private WithEvents mTxtRef As MSForms.TextBox
Private WithEvents mTxtQty As MSForms.TextBox
Private WithEvents mTxtDescription As MSForms.TextBox
Private WithEvents mChkUseExisting As MSForms.CheckBox
Private WithEvents mLstShippables As MSForms.ListBox
Private WithEvents mLstShipments As MSForms.ListBox
Private WithEvents mLstHold As MSForms.ListBox
Private WithEvents mBtnHistory As MSForms.CommandButton
Private WithEvents mBtnHistorySheet As MSForms.CommandButton
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnAdd As MSForms.CommandButton
Private WithEvents mBtnUpdate As MSForms.CommandButton
Private WithEvents mBtnRemove As MSForms.CommandButton
Private WithEvents mBtnHold As MSForms.CommandButton
Private WithEvents mBtnReturn As MSForms.CommandButton
Private WithEvents mBtnStage As MSForms.CommandButton
Private WithEvents mBtnSend As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mPages As MSForms.TabStrip
Private WithEvents mBtnBoxBuilderRefresh As MSForms.CommandButton
Private WithEvents mBtnBoxBuilderNew As MSForms.CommandButton
Private WithEvents mBtnBoxBuilderAddComponent As MSForms.CommandButton
Private WithEvents mBtnBoxBuilderRemoveComponent As MSForms.CommandButton
Private WithEvents mTxtBoxBuilderSearch As MSForms.TextBox
Private WithEvents mBtnBoxMakerRefresh As MSForms.CommandButton
Private WithEvents mLstBoxBuilderDesigns As MSForms.ListBox
Private WithEvents mCboBoxBuilderVersion As MSForms.ComboBox
Private WithEvents mBtnBoxBuilderSave As MSForms.CommandButton
Private WithEvents mBtnBoxBuilderUpdateVersion As MSForms.CommandButton
Private WithEvents mBtnBoxBuilderNewVersion As MSForms.CommandButton
Private WithEvents mBtnBoxBuilderDeleteVersion As MSForms.CommandButton
Private WithEvents mBtnBoxBuilderArchive As MSForms.CommandButton
Private WithEvents mBtnBoxBuilderDelete As MSForms.CommandButton
Private WithEvents mLstBoxMakerDesigns As MSForms.ListBox
Private WithEvents mCboBoxMakerVersion As MSForms.ComboBox
Private WithEvents mBtnBoxMakerMake As MSForms.CommandButton
Private WithEvents mBtnBoxMakerUnmake As MSForms.CommandButton

Private mTxtBox As MSForms.TextBox
Private mTxtVersion As MSForms.TextBox
Private mTxtUom As MSForms.TextBox
Private mTxtLocation As MSForms.TextBox
Private mTxtSystemKey As MSForms.TextBox
Private mTxtCarrier As MSForms.ComboBox
Private mTxtStatus As MSForms.TextBox
Private mLblSyncState As MSForms.Label
Private mLblBoxBuilderPage As MSForms.Label
Private mLblBoxMakerPage As MSForms.Label
Private mTxtBoxBuilderName As MSForms.TextBox
Private mTxtBoxBuilderUom As MSForms.TextBox
Private mTxtBoxBuilderLocation As MSForms.TextBox
Private mTxtBoxBuilderDescription As MSForms.TextBox
Private mCboBoxBuilderStatus As MSForms.ComboBox
Private mLstBoxBuilderInventory As MSForms.ListBox
Private mTxtBoxBuilderComponentQty As MSForms.TextBox
Private mLstBoxBuilderComponents As MSForms.ListBox
Private mTxtBoxMakerQty As MSForms.TextBox
Private mLstBoxMakerComponents As MSForms.ListBox
Private mLblBoxBuilderDesignsHeader As MSForms.Label
Private mLblBoxBuilderInventoryHeader As MSForms.Label
Private mLblBoxBuilderComponentsHeader As MSForms.Label
Private mLblBoxMakerDesignsHeader As MSForms.Label
Private mLblBoxMakerComponentsHeader As MSForms.Label

Private mShippables As Variant
Private mNasReservationTotals As Object
Private mLoading As Boolean
Private mBuilt As Boolean
Private mAnchors As Object
Private mResizeInitialized As Boolean
Private mOperatorWorkbook As Workbook
Private mActivityContext As String
Private mNextPollTime As Date
Private mAutoSyncArmed As Boolean
Private mLastShippablesLoadReport As String
Private mUseInjectedReservationTotalsForTest As Boolean
Private mActionTimer As New cShippingActionTimer
Private mSelectedBoxBuilderPackageSystemKey As String
Private mSelectedBoxMakerPackageSystemKey As String
Private mBoxBuilderInventoryRows As Variant

Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8
Private Const POLL_INTERVAL_SECONDS As Long = 45

Private Sub UserForm_Initialize()
    mActivityContext = modActivity.CaptureContext()
    BuildLayout
End Sub

Private Sub UserForm_Activate()
    modTS_Shipments.RegisterShipmentsFormAutoSync Me
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeInitialized = True
    End If
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
    LayoutBoxDesignerPage
    LayoutBoxMakerPage
    ApplyBoxingHeaderLayout
End Sub

Private Sub UserForm_Layout()
    If mAnchors Is Nothing Then Exit Sub
    mAnchors.ResizeControls
    LayoutBoxDesignerPage
    LayoutBoxMakerPage
    ApplyBoxingHeaderLayout
End Sub

Private Sub UserForm_Terminate()
    CancelAutoSync
    modTS_Shipments.UnregisterShipmentsFormAutoSync Me
    Set mAnchors = Nothing
    Set mOperatorWorkbook = Nothing
End Sub

Public Sub InitializeFromShipping(Optional ByVal preserveActiveRows As Boolean = False)
    On Error GoTo FailInit

    Dim previousPointer As Long
    Dim quietStarted As Boolean
    Dim startedAt As Single
    Dim elapsedMs As Long
    Dim operatorWb As Workbook
    Dim loadStep As String

    mActionTimer.Start
    TLap "InitializeFromShipping start"
    loadStep = "build layout"
    If Not mBuilt Then BuildLayout
    TLap "build layout"
    loadStep = "resolve operator workbook"
    Set operatorWb = ResolveOperatorWorkbook()
    If operatorWb Is Nothing Then
        Err.Raise vbObjectError + 7680, "frmShipmentsTally.InitializeFromShipping", _
                  "The captured Shipping operator workbook is no longer available."
    End If
    TLap "resolve operator workbook"
    loadStep = "begin quiet UI"
    previousPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass
    modUiQuiet.BeginQuietUi operatorWb
    quietStarted = True
    startedAt = Timer

    loadStep = "hide support sheets"
    modTS_Shipments.EnforceShippingSupportSheetsHidden operatorWb
    TLap "hide support sheets"
    mLoading = True
    loadStep = "load carriers"
    LoadCarrierChoices
    TLap "load carriers"
    loadStep = "load existing-inventory preference"
    mChkUseExisting.Value = modTS_Shipments.ShipmentsFormUseExistingInventory()
    TLap "load existing-inventory preference"
    loadStep = "load shippables"
    LoadShippables operatorWb
    TLap "load shippables"
    loadStep = "load shipment state"
    If Not preserveActiveRows Then modTS_Shipments.ShipmentsFormClearActiveLines operatorWb
    LoadShipmentState operatorWb
    TLap "load shipment state"
    loadStep = "evict orphaned active overlays"
    EvictOrphanedActiveOverlays
    TLap "evict orphaned active overlays"
    loadStep = "refresh projected inventory"
    RefreshProjectedShippableInventory
    TLap "refresh projected inventory"
    loadStep = "update sync label"
    UpdateSyncStateLabel
    TLap "update sync label"
    mLoading = False

    If mLstShippables.ListCount > 0 Then
        mLstShippables.ListIndex = 0
        LoadSelectedShippable
    End If
    elapsedMs = mActionTimer.ElapsedMilliseconds(startedAt)
    If mLstShippables.ListCount = 0 Then
        ShowStatus "Loaded shipments form in " & CStr(elapsedMs) & " ms, but no shippable inventory rows loaded. " & mLastShippablesLoadReport & vbCrLf & TimingSummary()
    Else
        ShowStatus "Loaded shipments form in " & CStr(elapsedMs) & " ms." & vbCrLf & TimingSummary()
    End If
    mAutoSyncArmed = (PendingShipmentSyncCount() > 0)
    If mAutoSyncArmed Then ScheduleAutoSync

CleanExit:
    On Error Resume Next
    mLoading = False
    modTS_Shipments.EnforceShippingSupportSheetsHidden operatorWb
    If quietStarted Then modUiQuiet.EndQuietUi
    Me.MousePointer = previousPointer
    On Error GoTo 0
    Exit Sub

FailInit:
    ShowStatus "Shipments form load failed at " & loadStep & ": " & Err.Description
    Resume CleanExit
End Sub

Private Sub TLap(ByVal label As String)
    mActionTimer.Mark label
End Sub

Private Function TimingSummary() As String
    TimingSummary = mActionTimer.Summary()
End Function

Public Function HasCurrentContext() As Boolean
    HasCurrentContext = modShippingFormContext.IsCurrent(mOperatorWorkbook, mActivityContext)
End Function

Private Function RequireActionContext() As Boolean
    Dim report As String
    RequireActionContext = modShippingFormContext.IsCurrent(mOperatorWorkbook, mActivityContext, report)
    If RequireActionContext Then Exit Function
    CancelAutoSync
    ShowStatus report
End Function

Public Sub SetOperatorWorkbook(ByVal wb As Workbook)
    If IsUsableOperatorWorkbook(wb) Then Set mOperatorWorkbook = wb
End Sub

Private Function ResolveOperatorWorkbook() As Workbook
    On Error Resume Next

    Dim nameCheck As String

    If Not mOperatorWorkbook Is Nothing Then
        nameCheck = mOperatorWorkbook.Name
        If Err.Number = 0 And Trim$(nameCheck) <> "" And IsUsableOperatorWorkbook(mOperatorWorkbook) Then
            Set ResolveOperatorWorkbook = mOperatorWorkbook
            Exit Function
        End If
        Err.Clear
        Set mOperatorWorkbook = Nothing
    End If
    On Error GoTo 0
End Function

Private Function IsUsableOperatorWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo CleanExit

    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function
    If Trim$(wb.Name) = "" Then Exit Function
    IsUsableOperatorWorkbook = True

CleanExit:
End Function

Private Function IsShipmentsOperatorWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo CleanExit

    If Not IsUsableOperatorWorkbook(wb) Then Exit Function
    If WorkbookHasTable(wb, "invSys") Then
        IsShipmentsOperatorWorkbook = True
        Exit Function
    End If
    If Not WorkbookSheetExists(wb, "ShipmentsTally") Is Nothing Then
        IsShipmentsOperatorWorkbook = True
    End If

CleanExit:
End Function

Private Function WorkbookSheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    If Not wb Is Nothing Then Set WorkbookSheetExists = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function WorkbookHasTable(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    On Error GoTo CleanExit

    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo CleanExit
        If Not lo Is Nothing Then
            WorkbookHasTable = True
            Exit Function
        End If
    Next ws

CleanExit:
End Function

Public Sub ScheduleAutoSync()
    On Error Resume Next

    CancelAutoSync
    mNextPollTime = Now + TimeSerial(0, 0, POLL_INTERVAL_SECONDS)
    Application.OnTime EarliestTime:=mNextPollTime, _
                       Procedure:=modTS_Shipments.ShipmentsFormAutoSyncProcedureName(), _
                       Schedule:=True
    On Error GoTo 0
End Sub

Public Sub CancelAutoSync()
    On Error Resume Next

    If mNextPollTime > 0 Then
        Application.OnTime EarliestTime:=mNextPollTime, _
                           Procedure:=modTS_Shipments.ShipmentsFormAutoSyncProcedureName(), _
                           Schedule:=False
        mNextPollTime = 0
    End If
    On Error GoTo 0
End Sub

Public Sub ArmAutoSync()
    mAutoSyncArmed = True
    UpdateSyncStateLabel
    ScheduleAutoSync
End Sub

Public Sub AutoSyncIfPending()
    On Error GoTo CleanExit

    Dim operatorWb As Workbook
    Dim report As String
    Dim changedLoading As Boolean
    Dim syncCount As Long
    Dim nasBeforeRefresh As Object
    Dim nasAfterRefresh As Object
    Dim nasStatus As String

    If Not mAutoSyncArmed Then Exit Sub
    If Not RequireActionContext() Then Exit Sub
    If mLoading Then
        ShowStatus "AutoSync: skipped (loading)."
        GoTo CleanExit
    End If
    syncCount = PendingShipmentSyncCount()
    If syncCount <= 0 Then
        EvictOrphanedActiveOverlays
        modTS_Shipments.EvictCompletedShipmentInventoryOverlaysForShippables mShippables
        RefreshProjectedShippableInventory
        syncCount = PendingShipmentSyncCount()
        If syncCount <= 0 And Not modTS_Shipments.HasAnyPendingBoxVersionInventoryOverlay() Then
            mAutoSyncArmed = False
            UpdateSyncStateLabel
            Exit Sub
        End If
    End If
    Set nasBeforeRefresh = ShippableNasSnapshot()

    Set operatorWb = ResolveOperatorWorkbook()
    If operatorWb Is Nothing Then
        ShowStatus "AutoSync: operator workbook not resolved."
        GoTo CleanExit
    End If

    If modTS_Shipments.ShipmentsFormAutoSyncRefresh(operatorWb, report) Then
        mLoading = True
        changedLoading = True
        LoadShippables
        Set nasAfterRefresh = ShippableNasSnapshot()
        LoadShipmentState
        RefreshProjectedShippableInventory
        mLoading = False
        changedLoading = False
        If PendingShipmentSyncCount() <= 0 Then mAutoSyncArmed = False
        UpdateSyncStateLabel
        nasStatus = ShippableNasChangeSummary(nasBeforeRefresh, nasAfterRefresh)
        ShowStatus nasStatus & vbCrLf & "AutoSync: " & report
    Else
        ShowStatus "AutoSync: refresh failed. " & report
        UpdateSyncStateLabel
    End If

CleanExit:
    If changedLoading Then mLoading = False
    If mAutoSyncArmed Then ScheduleAutoSync
End Sub

Private Function ShippableNasSnapshot() As Object
    On Error GoTo CleanExit

    Dim result As Object
    Dim r As Long
    Dim boxText As String
    Dim versionText As String
    Dim nasText As String
    Dim key As String

    Set result = CreateObject("Scripting.Dictionary")
    If IsEmpty(mShippables) Then
        Set ShippableNasSnapshot = result
        Exit Function
    End If

    For r = 1 To UBound(mShippables, 1)
        boxText = Trim$(NzText(mShippables(r, 1)))
        versionText = Trim$(NzText(mShippables(r, 2)))
        nasText = Trim$(NzText(mShippables(r, 4)))
        If boxText <> "" Or versionText <> "" Then
            key = LCase$(boxText) & "|" & LCase$(versionText)
            result(key) = boxText & vbTab & versionText & vbTab & nasText
        End If
    Next r

    Set ShippableNasSnapshot = result
    Exit Function

CleanExit:
    Set ShippableNasSnapshot = CreateObject("Scripting.Dictionary")
End Function

Private Function ShippableNasChangeSummary(ByVal beforeMap As Object, ByVal afterMap As Object) As String
    On Error GoTo CleanExit

    Dim key As Variant
    Dim beforeParts As Variant
    Dim afterParts As Variant
    Dim changes As String
    Dim labelText As String

    If afterMap Is Nothing Or afterMap.Count = 0 Then
        ShippableNasChangeSummary = "NAS Inv checked: no visible shippable rows loaded."
        Exit Function
    End If

    For Each key In afterMap.Keys
        afterParts = Split(CStr(afterMap(key)), vbTab)
        If beforeMap Is Nothing Or Not beforeMap.Exists(CStr(key)) Then GoTo NextKey
        beforeParts = Split(CStr(beforeMap(key)), vbTab)
        If UBound(beforeParts) < 2 Or UBound(afterParts) < 2 Then GoTo NextKey
        If Trim$(CStr(beforeParts(2))) <> Trim$(CStr(afterParts(2))) Then
            labelText = Trim$(CStr(afterParts(0)) & " " & CStr(afterParts(1)))
            If labelText = "" Then labelText = CStr(key)
            If changes <> "" Then changes = changes & "; "
            changes = changes & labelText & " " & _
                      IIf(Trim$(CStr(beforeParts(2))) = "", "blank", Trim$(CStr(beforeParts(2)))) & _
                      " -> " & IIf(Trim$(CStr(afterParts(2))) = "", "blank", Trim$(CStr(afterParts(2))))
        End If
NextKey:
    Next key

    If changes = "" Then
        ShippableNasChangeSummary = "NAS Inv checked: no visible changes."
    Else
        ShippableNasChangeSummary = "NAS Inv updated: " & changes & "."
    End If
    Exit Function

CleanExit:
    ShippableNasChangeSummary = "NAS Inv checked."
End Function

Private Sub BuildLayout()
    If mBuilt Then Exit Sub
    mBuilt = True

    Me.Caption = "Shipping Shipments"
    Me.Width = 980
    Me.Height = 850
    Me.ScrollBars = fmScrollBarsBoth
    Me.ScrollWidth = 970
    Me.ScrollHeight = 650

    AddLabel "lblTitle", "Shipments", 12, 10, 140, 20, True
    Set mBtnHistory = AddButton("btnHistory", "History", 648, 10, 58, 24)
    Set mBtnHistorySheet = AddButton("btnHistorySheet", "Export", 712, 10, 58, 24)
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 774, 10, 58, 24)

    AddLabel "lblPicker", "Search Boxes", 12, 42, 78, 18, False
    Set mTxtPicker = AddTextBox("txtPicker", 96, 38, 300, 22)
    Set mChkUseExisting = AddCheckBox("chkUseExisting", "Use existing shippable inventory", 420, 38, 190, 22)
    Set mLblSyncState = AddLabel("lblSyncState", "", 620, 42, 210, 18, False)

    AddShippableHeaders 12, 70
    Set mLstShippables = AddListBox("lstShippables", 12, 90, 820, 92)
    With mLstShippables
        .ColumnCount = 8
        .ColumnWidths = "138 pt;48 pt;54 pt;68 pt;50 pt;38 pt;96 pt;42 pt"
    End With

    AddLabel "lblRef", "Ref", 12, 194, 34, 18, False
    AddLabel "lblBox", "Box", 108, 194, 34, 18, False
    AddLabel "lblVersion", "Alternative", 270, 194, 64, 18, False
    AddLabel "lblQty", "Qty", 336, 194, 34, 18, False
    AddLabel "lblUom", "UOM", 410, 194, 40, 18, False
    AddLabel "lblLocation", "Location", 470, 194, 60, 18, False
    AddLabel "lblSystemKey", "System Key", 620, 194, 68, 18, False
    AddLabel "lblCarrier", "Carrier", 12, 242, 54, 18, False

    Set mTxtRef = AddTextBox("txtRef", 12, 212, 82, 22)
    Set mTxtBox = AddTextBox("txtBox", 108, 212, 148, 22)
    Set mTxtVersion = AddTextBox("txtVersion", 270, 212, 52, 22)
    Set mTxtQty = AddTextBox("txtQty", 336, 212, 52, 22)
    Set mTxtUom = AddTextBox("txtUom", 410, 212, 44, 22)
    Set mTxtLocation = AddTextBox("txtLocation", 470, 212, 132, 22)
    Set mTxtSystemKey = AddTextBox("txtSystemKey", 620, 212, 92, 22)
    Set mTxtCarrier = AddComboBox("txtCarrier", 108, 238, 148, 22)
    Set mTxtDescription = AddTextBox("txtDescription", 12, 240, 1, 1)
    mTxtDescription.Visible = False
    LockTextBox mTxtBox
    LockTextBox mTxtVersion
    LockTextBox mTxtUom
    LockTextBox mTxtLocation
    LockTextBox mTxtSystemKey
    Set mBtnAdd = AddButton("btnAdd", "Add", 668, 210, 44, 26)
    Set mBtnUpdate = AddButton("btnUpdate", "Update Row", 718, 210, 62, 26)
    Set mBtnRemove = AddButton("btnRemove", "Remove", 786, 210, 58, 26)

    AddLabel "lblShipments", "Shipments", 12, 276, 90, 18, True
    Set mBtnStage = AddButton("btnStage", "To Shipments", 596, 272, 98, 28)
    Set mBtnSend = AddButton("btnSend", "Shipments Sent", 704, 272, 128, 28)
    AddShipmentLineHeaders 12, 300
    Set mLstShipments = AddListBox("lstShipments", 12, 320, 820, 108)
    With mLstShipments
        .ColumnCount = 12
        .ColumnWidths = "76 pt;150 pt;50 pt;40 pt;68 pt;44 pt;46 pt;58 pt;76 pt;0 pt;0 pt;0 pt"
        .MultiSelect = fmMultiSelectExtended
    End With
    Set mBtnHold = AddButton("btnHold", "Send Hold", 498, 274, 88, 24)

    AddLabel "lblHold", "Not Shipped", 12, 444, 100, 18, True
    AddShipmentLineHeaders 12, 468
    Set mLstHold = AddListBox("lstHold", 12, 488, 820, 60)
    With mLstHold
        .ColumnCount = 12
        .ColumnWidths = "76 pt;150 pt;50 pt;40 pt;68 pt;44 pt;46 pt;58 pt;76 pt;0 pt;0 pt;0 pt"
        .MultiSelect = fmMultiSelectExtended
    End With
    Set mBtnReturn = AddButton("btnReturn", "Return", 744, 444, 88, 24)

    Set mTxtStatus = AddTextBox("txtStatus", 12, 552, 708, 68)
    With mTxtStatus
        .MultiLine = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .BackColor = &H8000000F
    End With
    Set mBtnClose = AddButton("btnClose", "Close", 776, 590, 56, 30)

    MoveStatusToTop
    BuildShippingPages
    InitializeAnchors
    LoadCarrierChoices
End Sub

Private Sub BuildShippingPages()
    Dim ctl As MSForms.Control
    Dim pageLabel As MSForms.Label

    For Each ctl In Me.Controls
        ctl.Top = ctl.Top + 30
        ctl.Tag = "Shipping"
    Next ctl

    Set mPages = Me.Controls.Add("Forms.TabStrip.1", "tabsShippingRole", True)
    With mPages
        .Left = 12
        .Top = 8
        .Width = 940
        .Height = 26
        .Tabs.Clear
        .Tabs.Add "tabShipping", "Shipping"
        .Tabs.Add "tabBoxBuilder", "Box Designer"
        .Tabs.Add "tabBoxMaker", "Box Maker"
        .Value = 0
        .Tag = "Shell"
    End With

    Set mLblBoxBuilderPage = AddLabel( _
        "lblBoxBuilderPage", _
        "Box Designer - shipping assembly alternatives", _
        18, 156, 280, 22, True)
    mLblBoxBuilderPage.Tag = "Box Designer"
    Set mBtnBoxBuilderNew = AddButton( _
        "btnBoxBuilderNewPage", "New Box", 682, 152, 98, 28)
    mBtnBoxBuilderNew.Tag = "Box Designer"
    Set mBtnBoxBuilderRefresh = AddButton( _
        "btnBoxBuilderRefreshPage", "Refresh Box Designs", 790, 152, 150, 28)
    mBtnBoxBuilderRefresh.Tag = "Box Designer"
    Set mLstBoxBuilderDesigns = AddListBox( _
        "lstBoxBuilderDesignsPage", 18, 204, 330, 126)
    With mLstBoxBuilderDesigns
        .ColumnCount = 5
        .ColumnWidths = "0 pt;130 pt;42 pt;72 pt;0 pt"
        .Tag = "Box Designer"
    End With
    Set pageLabel = AddLabel( _
        "lblBoxBuilderInventory", "Component inventory", 18, 340, 140, 18, True)
    pageLabel.Tag = "Box Designer"
    Set mTxtBoxBuilderSearch = AddTextBox("txtBoxBuilderSearch", 160, 336, 188, 22)
    mTxtBoxBuilderSearch.Tag = "Box Designer"
    Set mLstBoxBuilderInventory = AddListBox( _
        "lstBoxBuilderInventoryPage", 18, 384, 330, 124)
    With mLstBoxBuilderInventory
        .ColumnCount = 8
        .ColumnWidths = "0 pt;76 pt;180 pt;48 pt;90 pt;0 pt;60 pt;78 pt"
        .Tag = "Box Designer"
    End With
    Set pageLabel = AddLabel("lblBoxBuilderComponentQty", "Qty", 18, 518, 28, 18, False)
    pageLabel.Tag = "Box Designer"
    Set mTxtBoxBuilderComponentQty = AddTextBox( _
        "txtBoxBuilderComponentQty", 48, 514, 46, 22)
    mTxtBoxBuilderComponentQty.Value = "1"
    mTxtBoxBuilderComponentQty.Tag = "Box Designer"
    Set mBtnBoxBuilderAddComponent = AddButton( _
        "btnBoxBuilderAddComponentPage", "Add", 102, 512, 76, 26)
    mBtnBoxBuilderAddComponent.Tag = "Box Designer"
    Set mBtnBoxBuilderRemoveComponent = AddButton( _
        "btnBoxBuilderRemoveComponentPage", "Remove", 186, 512, 92, 26)
    mBtnBoxBuilderRemoveComponent.Tag = "Box Designer"
    Set pageLabel = AddLabel("lblBoxBuilderName", "Box Name", 362, 188, 70, 18, False)
    pageLabel.Tag = "Box Designer"
    Set mTxtBoxBuilderName = AddTextBox("txtBoxBuilderName", 436, 184, 190, 22)
    mTxtBoxBuilderName.Tag = "Box Designer"
    Set pageLabel = AddLabel("lblBoxBuilderVersion", "Alternative", 640, 188, 70, 18, False)
    pageLabel.Tag = "Box Designer"
    Set mCboBoxBuilderVersion = AddComboBox("cboBoxBuilderVersion", 700, 184, 86, 22)
    mCboBoxBuilderVersion.Tag = "Box Designer"
    Set pageLabel = AddLabel("lblBoxBuilderStatus", "Status", 800, 188, 54, 18, False)
    pageLabel.Tag = "Box Designer"
    Set mCboBoxBuilderStatus = AddComboBox("cboBoxBuilderStatus", 858, 184, 82, 22)
    With mCboBoxBuilderStatus
        .AddItem "Active"
        .AddItem "Archived"
        .Value = "Active"
        .Tag = "Box Designer"
    End With
    Set pageLabel = AddLabel("lblBoxBuilderUom", "UOM", 362, 218, 40, 18, False)
    pageLabel.Tag = "Box Designer"
    Set mTxtBoxBuilderUom = AddTextBox("txtBoxBuilderUom", 406, 214, 64, 22)
    mTxtBoxBuilderUom.Tag = "Box Designer"
    Set pageLabel = AddLabel("lblBoxBuilderLocation", "Location", 480, 218, 56, 18, False)
    pageLabel.Tag = "Box Designer"
    Set mTxtBoxBuilderLocation = AddTextBox("txtBoxBuilderLocation", 538, 214, 88, 22)
    mTxtBoxBuilderLocation.Tag = "Box Designer"
    Set pageLabel = AddLabel("lblBoxBuilderDescription", "Description", 362, 248, 70, 18, False)
    pageLabel.Tag = "Box Designer"
    Set mTxtBoxBuilderDescription = AddTextBox("txtBoxBuilderDescription", 436, 244, 504, 22)
    mTxtBoxBuilderDescription.Tag = "Box Designer"
    Set pageLabel = AddLabel("lblBoxBuilderComponents", "Selected alternative components", 362, 278, 230, 18, True)
    pageLabel.Tag = "Box Designer"
    Set mLstBoxBuilderComponents = AddListBox( _
        "lstBoxBuilderComponentsPage", 362, 316, 578, 192)
    With mLstBoxBuilderComponents
        .ColumnCount = 8
        .ColumnWidths = "0 pt;100 pt;70 pt;0 pt;46 pt;38 pt;68 pt;94 pt"
        .Tag = "Box Designer"
    End With
    Set mBtnBoxBuilderSave = AddButton( _
        "btnBoxBuilderSavePage", "Save Box", 362, 520, 86, 28)
    mBtnBoxBuilderSave.Tag = "Box Designer"
    Set mBtnBoxBuilderUpdateVersion = AddButton( _
        "btnBoxBuilderUpdateVersionPage", "Update Alternative", 456, 520, 116, 28)
    mBtnBoxBuilderUpdateVersion.Tag = "Box Designer"
    Set mBtnBoxBuilderNewVersion = AddButton( _
        "btnBoxBuilderNewVersionPage", "New Alternative", 568, 520, 108, 28)
    mBtnBoxBuilderNewVersion.Tag = "Box Designer"
    Set mBtnBoxBuilderDeleteVersion = AddButton( _
        "btnBoxBuilderDeleteVersionPage", "Delete Alternative", 672, 520, 116, 28)
    mBtnBoxBuilderDeleteVersion.Tag = "Box Designer"
    Set mBtnBoxBuilderArchive = AddButton( _
        "btnBoxBuilderArchivePage", "Archive Box", 784, 520, 106, 28)
    mBtnBoxBuilderArchive.Tag = "Box Designer"
    Set mBtnBoxBuilderDelete = AddButton( _
        "btnBoxBuilderDeletePage", "Delete Box", 834, 554, 106, 28)
    mBtnBoxBuilderDelete.Tag = "Box Designer"

    Set mLblBoxMakerPage = AddLabel( _
        "lblBoxMakerPage", _
        "Box Maker - released designs and inventory actions", _
        18, 156, 360, 22, True)
    mLblBoxMakerPage.Tag = "Box Maker"
    Set mBtnBoxMakerRefresh = AddButton( _
        "btnBoxMakerRefreshPage", "Refresh Box Maker", 790, 152, 150, 28)
    mBtnBoxMakerRefresh.Tag = "Box Maker"
    Set mLstBoxMakerDesigns = AddListBox( _
        "lstBoxMakerDesignsPage", 18, 204, 350, 316)
    With mLstBoxMakerDesigns
        .ColumnCount = 7
        .ColumnWidths = "0 pt;150 pt;0 pt;42 pt;72 pt;0 pt;0 pt"
        .Tag = "Box Maker"
    End With
    Set pageLabel = AddLabel("lblBoxMakerVersion", "Alternative", 386, 188, 70, 18, False)
    pageLabel.Tag = "Box Maker"
    Set mCboBoxMakerVersion = AddComboBox("cboBoxMakerVersion", 446, 184, 92, 22)
    mCboBoxMakerVersion.Tag = "Box Maker"
    Set pageLabel = AddLabel("lblBoxMakerQty", "Qty", 558, 188, 32, 18, False)
    pageLabel.Tag = "Box Maker"
    Set mTxtBoxMakerQty = AddTextBox("txtBoxMakerQty", 594, 184, 64, 22)
    mTxtBoxMakerQty.Value = "1"
    mTxtBoxMakerQty.Tag = "Box Maker"
    Set mBtnBoxMakerMake = AddButton( _
        "btnBoxMakerMakePage", "Make Boxes", 674, 182, 94, 28)
    mBtnBoxMakerMake.Tag = "Box Maker"
    Set mBtnBoxMakerUnmake = AddButton( _
        "btnBoxMakerUnmakePage", "Unbox", 776, 182, 80, 28)
    mBtnBoxMakerUnmake.Tag = "Box Maker"
    Set pageLabel = AddLabel("lblBoxMakerComponents", "Selected alternative components", 386, 226, 230, 18, True)
    pageLabel.Tag = "Box Maker"
    Set mLstBoxMakerComponents = AddListBox( _
        "lstBoxMakerComponentsPage", 386, 264, 554, 256)
    With mLstBoxMakerComponents
        .ColumnCount = 9
        .ColumnWidths = "0 pt;100 pt;68 pt;0 pt;46 pt;38 pt;64 pt;88 pt;52 pt"
        .Tag = "Box Maker"
    End With

    ConfigureBoxingListHeaders
    LayoutBoxDesignerPage
    LayoutBoxMakerPage
    mTxtStatus.Tag = "Shell"
    mBtnClose.Tag = "Shell"
    Me.ScrollHeight = Me.ScrollHeight + 30
    ApplyShippingPage
End Sub

Private Sub mPages_Change()
    ApplyShippingPage
End Sub

Private Sub ApplyShippingPage()
    Dim ctl As MSForms.Control
    Dim selectedPage As String

    If mPages Is Nothing Then Exit Sub
    selectedPage = mPages.SelectedItem.Caption
    For Each ctl In Me.Controls
        ctl.Visible = _
            (StrComp(NzText(ctl.Tag), "Shell", vbTextCompare) = 0) _
            Or (StrComp(NzText(ctl.Tag), selectedPage, vbTextCompare) = 0)
    Next ctl
    mPages.Visible = True

    Select Case selectedPage
        Case "Box Designer"
            RefreshBoxBuilderPage
        Case "Box Maker"
            RefreshBoxMakerPage
    End Select
End Sub

Public Function SelectShippingPageForTest(ByVal pageCaption As String) As String
    Dim pageIndex As Long

    If Not mBuilt Then BuildLayout
    For pageIndex = 0 To mPages.Tabs.Count - 1
        If StrComp(mPages.Tabs(pageIndex).Caption, Trim$(pageCaption), vbTextCompare) = 0 Then
            mPages.Value = pageIndex
            ApplyShippingPage
            SelectShippingPageForTest = _
                "OK|Selected=" & mPages.SelectedItem.Caption & _
                "|BoxBuilderActionsReachable=" & _
                    CStr(Not mBtnBoxBuilderNew Is Nothing And _
                         Not mBtnBoxBuilderAddComponent Is Nothing And _
                         Not mBtnBoxBuilderRemoveComponent Is Nothing And _
                         Not mBtnBoxBuilderSave Is Nothing And _
                         Not mBtnBoxBuilderUpdateVersion Is Nothing And _
                         Not mBtnBoxBuilderNewVersion Is Nothing And _
                         Not mBtnBoxBuilderDeleteVersion Is Nothing And _
                         Not mBtnBoxBuilderArchive Is Nothing And _
                         Not mBtnBoxBuilderDelete Is Nothing) & _
                "|BoxMakerActionsReachable=" & _
                    CStr(Not mBtnBoxMakerMake Is Nothing And _
                         Not mBtnBoxMakerUnmake Is Nothing)
            Exit Function
        End If
    Next pageIndex
    SelectShippingPageForTest = "FAIL|UnknownPage=" & pageCaption
End Function

Private Sub mBtnBoxBuilderRefresh_Click()
    RefreshBoxBuilderPage
End Sub

Private Sub mBtnBoxMakerRefresh_Click()
    RefreshBoxMakerPage
End Sub

Private Sub RefreshBoxBuilderPage()
    Dim rowsData As Variant

    If mOperatorWorkbook Is Nothing Then Exit Sub
    mLoading = True
    rowsData = modBoxingService.LoadBoxDesigns(mOperatorWorkbook, True, True)
    RenderPageRows mLstBoxBuilderDesigns, rowsData
    mBoxBuilderInventoryRows = modBoxingService.LoadComponentChoices(mOperatorWorkbook)
    FilterBoxBuilderInventory
    ClearBoxBuilderSelection
    If mLstBoxBuilderDesigns.ListCount > 0 Then mLstBoxBuilderDesigns.ListIndex = 0
    mLoading = False
    If mLstBoxBuilderDesigns.ListIndex >= 0 Then LoadSelectedBoxBuilderDesign
End Sub

Private Sub mTxtBoxBuilderSearch_Change()
    If mLoading Then Exit Sub
    FilterBoxBuilderInventory
End Sub

Private Sub FilterBoxBuilderInventory()
    Dim sourceRows As Variant
    Dim filtered() As Variant
    Dim trimmed() As Variant
    Dim searchText As String
    Dim haystack As String
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim sourceCols As Long

    mLstBoxBuilderInventory.Clear
    If IsEmpty(mBoxBuilderInventoryRows) Or Not IsArray(mBoxBuilderInventoryRows) Then Exit Sub
    sourceRows = mBoxBuilderInventoryRows
    sourceCols = UBound(sourceRows, 2)
    searchText = LCase$(Trim$(NzText(mTxtBoxBuilderSearch.Value)))
    ReDim filtered(1 To UBound(sourceRows, 1), 1 To 8)
    For r = 1 To UBound(sourceRows, 1)
        haystack = ""
        For c = 1 To sourceCols
            haystack = haystack & " " & LCase$(NzText(sourceRows(r, c)))
        Next c
        If searchText = "" Or InStr(1, haystack, searchText, vbTextCompare) > 0 Then
            outRow = outRow + 1
            For c = 1 To sourceCols
                filtered(outRow, c) = sourceRows(r, c)
            Next c
            filtered(outRow, 8) = DisplayVersionOrNA("")
        End If
    Next r
    If outRow = 0 Then Exit Sub
    ReDim trimmed(1 To outRow, 1 To 8)
    For r = 1 To outRow
        For c = 1 To 8
            trimmed(r, c) = filtered(r, c)
        Next c
    Next r
    RenderPageRows mLstBoxBuilderInventory, trimmed
End Sub

Private Function DisplayVersionOrNA(ByVal versionText As String) As String
    versionText = Trim$(versionText)
    If versionText = "" Then
        DisplayVersionOrNA = "NA"
    Else
        DisplayVersionOrNA = versionText
    End If
End Function

Private Sub mBtnBoxBuilderNew_Click()
    mLoading = True
    ClearBoxBuilderSelection
    mCboBoxBuilderVersion.AddItem "v1"
    mCboBoxBuilderVersion.ListIndex = 0
    mLoading = False
    ShowStatus "New box design ready. Add components, then save the box."
End Sub

Private Sub mBtnBoxBuilderAddComponent_Click()
    Dim componentIndex As Long
    Dim targetIndex As Long
    Dim quantityValue As Double
    Dim versionLabel As String

    componentIndex = mLstBoxBuilderInventory.ListIndex
    If componentIndex < 0 Then
        ShowStatus "Select a managed inventory component to add."
        Exit Sub
    End If
    quantityValue = ParseNumber(NzText(mTxtBoxBuilderComponentQty.Value))
    If quantityValue <= 0 Then
        ShowStatus "Enter a positive component quantity."
        Exit Sub
    End If
    versionLabel = NzText(mCboBoxBuilderVersion.Value)
    If versionLabel = "" Then versionLabel = "v1"

    mLstBoxBuilderComponents.AddItem versionLabel
    targetIndex = mLstBoxBuilderComponents.ListCount - 1
    mLstBoxBuilderComponents.List(targetIndex, 1) = _
        NzText(mLstBoxBuilderInventory.List(componentIndex, 2))
    mLstBoxBuilderComponents.List(targetIndex, 2) = _
        NzText(mLstBoxBuilderInventory.List(componentIndex, 1))
    mLstBoxBuilderComponents.List(targetIndex, 3) = _
        NzText(mLstBoxBuilderInventory.List(componentIndex, 0))
    mLstBoxBuilderComponents.List(targetIndex, 4) = CStr(quantityValue)
    mLstBoxBuilderComponents.List(targetIndex, 5) = _
        NzText(mLstBoxBuilderInventory.List(componentIndex, 3))
    mLstBoxBuilderComponents.List(targetIndex, 6) = _
        NzText(mLstBoxBuilderInventory.List(componentIndex, 4))
    mLstBoxBuilderComponents.List(targetIndex, 7) = _
        NzText(mLstBoxBuilderInventory.List(componentIndex, 5))
    ShowStatus "Component added to the selected box alternative."
End Sub

Private Sub mBtnBoxBuilderRemoveComponent_Click()
    If mLstBoxBuilderComponents.ListIndex < 0 Then
        ShowStatus "Select a box component to remove."
        Exit Sub
    End If
    mLstBoxBuilderComponents.RemoveItem mLstBoxBuilderComponents.ListIndex
    ShowStatus "Component removed from the selected box alternative."
End Sub

Private Sub RefreshBoxMakerPage()
    Dim rowsData As Variant
    Dim report As String

    If mOperatorWorkbook Is Nothing Then Exit Sub
    mLoading = True
    rowsData = modBoxingService.LoadBoxMakerChoices(mOperatorWorkbook, report)
    RenderPageRows mLstBoxMakerDesigns, rowsData
    mCboBoxMakerVersion.Clear
    mLstBoxMakerComponents.Clear
    mSelectedBoxMakerPackageSystemKey = vbNullString
    If mLstBoxMakerDesigns.ListCount > 0 Then mLstBoxMakerDesigns.ListIndex = 0
    mLoading = False
    If mLstBoxMakerDesigns.ListIndex >= 0 Then LoadSelectedBoxMakerDesign
    If report <> "" Then ShowStatus report
End Sub

Private Sub mLstBoxBuilderDesigns_Click()
    If mLoading Then Exit Sub
    LoadSelectedBoxBuilderDesign
End Sub

Private Sub mCboBoxBuilderVersion_Change()
    If mLoading Then Exit Sub
    LoadSelectedBoxBuilderComponents
End Sub

Private Sub LoadSelectedBoxBuilderDesign()
    Dim versionRows As Variant
    Dim rowIndex As Long
    Dim listIndex As Long

    If mLstBoxBuilderDesigns.ListIndex < 0 Then Exit Sub
    mLoading = True
    listIndex = mLstBoxBuilderDesigns.ListIndex
    mSelectedBoxBuilderPackageSystemKey = NzText(mLstBoxBuilderDesigns.List(listIndex, 0))
    mTxtBoxBuilderName.Value = NzText(mLstBoxBuilderDesigns.List(listIndex, 1))
    mTxtBoxBuilderUom.Value = NzText(mLstBoxBuilderDesigns.List(listIndex, 2))
    mTxtBoxBuilderLocation.Value = NzText(mLstBoxBuilderDesigns.List(listIndex, 3))
    mTxtBoxBuilderDescription.Value = NzText(mLstBoxBuilderDesigns.List(listIndex, 4))
    If Trim$(NzText(mTxtBoxBuilderUom.Value)) = "" Then mTxtBoxBuilderUom.Value = "ea"

    mCboBoxBuilderVersion.Clear
    versionRows = modBoxingService.LoadBoxDesignVersions( _
        mOperatorWorkbook, mSelectedBoxBuilderPackageSystemKey)
    If Not IsEmpty(versionRows) Then
        mCboBoxBuilderVersion.ColumnCount = 2
        For rowIndex = LBound(versionRows, 1) To UBound(versionRows, 1)
            mCboBoxBuilderVersion.AddItem NzText(versionRows(rowIndex, 1))
            If UBound(versionRows, 2) >= 2 Then
                mCboBoxBuilderVersion.List(mCboBoxBuilderVersion.ListCount - 1, 1) = _
                    NzText(versionRows(rowIndex, 2))
            End If
        Next rowIndex
    End If
    If mCboBoxBuilderVersion.ListCount = 0 Then mCboBoxBuilderVersion.AddItem "v1"
    mCboBoxBuilderVersion.ListIndex = 0
    mLoading = False
    LoadSelectedBoxBuilderComponents
End Sub

Private Sub LoadSelectedBoxBuilderComponents()
    Dim componentRows As Variant
    Dim statusText As String

    If mSelectedBoxBuilderPackageSystemKey = "" Then Exit Sub
    If mCboBoxBuilderVersion.ListIndex < 0 Then Exit Sub
    If mCboBoxBuilderVersion.ColumnCount > 1 Then
        statusText = NzText(mCboBoxBuilderVersion.List(mCboBoxBuilderVersion.ListIndex, 1))
    End If
    If statusText = "" Then statusText = "Active"
    mCboBoxBuilderStatus.Value = statusText
    componentRows = modBoxingService.LoadBoxDesignComponents( _
        mOperatorWorkbook, mSelectedBoxBuilderPackageSystemKey, _
        NzText(mCboBoxBuilderVersion.Value))
    RenderPageRows mLstBoxBuilderComponents, componentRows
End Sub

Private Sub ClearBoxBuilderSelection()
    mSelectedBoxBuilderPackageSystemKey = vbNullString
    mTxtBoxBuilderName.Value = ""
    mTxtBoxBuilderUom.Value = "ea"
    mTxtBoxBuilderLocation.Value = ""
    mTxtBoxBuilderDescription.Value = ""
    mCboBoxBuilderVersion.Clear
    mCboBoxBuilderStatus.Value = "Active"
    mLstBoxBuilderComponents.Clear
End Sub

Private Function BoxBuilderPageComponents() As Variant
    Dim result() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    If mLstBoxBuilderComponents.ListCount = 0 Then Exit Function
    ReDim result(1 To mLstBoxBuilderComponents.ListCount, 1 To 8)
    For rowIndex = 0 To mLstBoxBuilderComponents.ListCount - 1
        For columnIndex = 0 To 7
            result(rowIndex + 1, columnIndex + 1) = _
                mLstBoxBuilderComponents.List(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex
    BoxBuilderPageComponents = result
End Function

Private Sub mBtnBoxBuilderSave_Click()
    Dim report As String
    Dim succeeded As Boolean

    succeeded = modBoxingService.SaveBoxDesign( _
        mOperatorWorkbook, NzText(mTxtBoxBuilderName.Value), _
        NzText(mTxtBoxBuilderUom.Value), NzText(mTxtBoxBuilderLocation.Value), _
        NzText(mTxtBoxBuilderDescription.Value), BoxBuilderPageComponents(), _
        "BOX", "v1", NzText(mCboBoxBuilderStatus.Value), report)
    ShowStatus report
    If succeeded Then RefreshBoxBuilderPage
End Sub

Private Sub mBtnBoxBuilderUpdateVersion_Click()
    RunBoxBuilderVersionAction "UPDATE"
End Sub

Private Sub mBtnBoxBuilderNewVersion_Click()
    RunBoxBuilderVersionAction "NEW"
End Sub

Private Sub mBtnBoxBuilderDeleteVersion_Click()
    Dim report As String

    If mSelectedBoxBuilderPackageSystemKey = "" Or _
       mCboBoxBuilderVersion.ListIndex < 0 Then
        ShowStatus "Select a saved box alternative before deleting."
        Exit Sub
    End If
    If MsgBox("Delete the selected box alternative?", _
              vbQuestion + vbYesNo, "Delete Box Alternative") <> vbYes Then Exit Sub
    If modBoxingService.DeleteBoxDesignVersion( _
            mOperatorWorkbook, mSelectedBoxBuilderPackageSystemKey, _
            NzText(mCboBoxBuilderVersion.Value), report) Then
        RefreshBoxBuilderPage
    End If
    ShowStatus report
End Sub

Private Sub mBtnBoxBuilderArchive_Click()
    Dim report As String

    If mSelectedBoxBuilderPackageSystemKey = "" Then
        ShowStatus "Select a saved box before archiving."
        Exit Sub
    End If
    If MsgBox("Archive the selected box design?", _
              vbQuestion + vbYesNo, "Archive Box Design") <> vbYes Then Exit Sub
    If modBoxingService.ArchiveBoxDesign( _
            mOperatorWorkbook, mSelectedBoxBuilderPackageSystemKey, report) Then
        RefreshBoxBuilderPage
    End If
    ShowStatus report
End Sub

Private Sub mBtnBoxBuilderDelete_Click()
    Dim report As String

    If mSelectedBoxBuilderPackageSystemKey = "" Then
        ShowStatus "Select a saved box before deleting."
        Exit Sub
    End If
    If MsgBox("Delete the selected box design and all alternatives?", _
              vbQuestion + vbYesNo, "Delete Box Design") <> vbYes Then Exit Sub
    If modBoxingService.DeleteBoxDesign( _
            mOperatorWorkbook, mSelectedBoxBuilderPackageSystemKey, report) Then
        RefreshBoxBuilderPage
    End If
    ShowStatus report
End Sub

Private Sub RunBoxBuilderVersionAction(ByVal actionText As String)
    Dim report As String
    Dim succeeded As Boolean

    succeeded = modBoxingService.SaveBoxDesign( _
        mOperatorWorkbook, NzText(mTxtBoxBuilderName.Value), _
        NzText(mTxtBoxBuilderUom.Value), NzText(mTxtBoxBuilderLocation.Value), _
        NzText(mTxtBoxBuilderDescription.Value), BoxBuilderPageComponents(), _
        actionText, NzText(mCboBoxBuilderVersion.Value), _
        NzText(mCboBoxBuilderStatus.Value), report)
    ShowStatus report
    If succeeded Then RefreshBoxBuilderPage
End Sub

Private Sub mLstBoxMakerDesigns_Click()
    If mLoading Then Exit Sub
    LoadSelectedBoxMakerDesign
End Sub

Private Sub mCboBoxMakerVersion_Change()
    If mLoading Then Exit Sub
    LoadSelectedBoxMakerComponents
End Sub

Private Sub LoadSelectedBoxMakerDesign()
    Dim versionRows As Variant
    Dim rowIndex As Long

    If mLstBoxMakerDesigns.ListIndex < 0 Then Exit Sub
    mSelectedBoxMakerPackageSystemKey = _
        NzText(mLstBoxMakerDesigns.List(mLstBoxMakerDesigns.ListIndex, 0))
    mLoading = True
    mCboBoxMakerVersion.Clear
    versionRows = modBoxingService.LoadBoxMakerVersions( _
        mOperatorWorkbook, mSelectedBoxMakerPackageSystemKey)
    If Not IsEmpty(versionRows) Then
        For rowIndex = LBound(versionRows, 1) To UBound(versionRows, 1)
            If UBound(versionRows, 2) < 2 Or _
               StrComp(NzText(versionRows(rowIndex, 2)), "Active", vbTextCompare) = 0 Then
                mCboBoxMakerVersion.AddItem NzText(versionRows(rowIndex, 1))
            End If
        Next rowIndex
    End If
    If mCboBoxMakerVersion.ListCount > 0 Then mCboBoxMakerVersion.ListIndex = 0
    mLoading = False
    LoadSelectedBoxMakerComponents
End Sub

Private Sub LoadSelectedBoxMakerComponents()
    Dim componentRows As Variant

    If mSelectedBoxMakerPackageSystemKey = "" Then Exit Sub
    If mCboBoxMakerVersion.ListIndex < 0 Then Exit Sub
    componentRows = modBoxingService.LoadBoxMakerComponents( _
        mOperatorWorkbook, mSelectedBoxMakerPackageSystemKey, _
        NzText(mCboBoxMakerVersion.Value))
    RenderPageRows mLstBoxMakerComponents, componentRows
End Sub

Private Function BoxMakerPageComponents() As Variant
    Dim result() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    If mLstBoxMakerComponents.ListCount = 0 Then Exit Function
    ReDim result(1 To mLstBoxMakerComponents.ListCount, 1 To 9)
    For rowIndex = 0 To mLstBoxMakerComponents.ListCount - 1
        For columnIndex = 0 To 8
            result(rowIndex + 1, columnIndex + 1) = _
                mLstBoxMakerComponents.List(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex
    BoxMakerPageComponents = result
End Function

Private Sub mBtnBoxMakerMake_Click()
    Dim report As String
    Dim succeeded As Boolean
    Dim selectedIndex As Long

    selectedIndex = mLstBoxMakerDesigns.ListIndex
    If selectedIndex < 0 Then
        ShowStatus "Select a released box design."
        Exit Sub
    End If
    succeeded = modBoxingService.PostBoxMakerAction( _
        mOperatorWorkbook, mSelectedBoxMakerPackageSystemKey, _
        NzText(mLstBoxMakerDesigns.List(selectedIndex, 1)), _
        NzText(mLstBoxMakerDesigns.List(selectedIndex, 3)), _
        NzText(mLstBoxMakerDesigns.List(selectedIndex, 4)), _
        NzText(mLstBoxMakerDesigns.List(selectedIndex, 5)), _
        NzText(mCboBoxMakerVersion.Value), ParseNumber(NzText(mTxtBoxMakerQty.Value)), _
        BoxMakerPageComponents(), "MAKE", report)
    ShowStatus report
    If succeeded Then RefreshBoxMakerPage
End Sub

Private Sub mBtnBoxMakerUnmake_Click()
    Dim report As String
    Dim succeeded As Boolean
    Dim selectedIndex As Long

    selectedIndex = mLstBoxMakerDesigns.ListIndex
    If selectedIndex < 0 Then
        ShowStatus "Select a released box design."
        Exit Sub
    End If
    succeeded = modBoxingService.PostBoxMakerAction( _
        mOperatorWorkbook, mSelectedBoxMakerPackageSystemKey, _
        NzText(mLstBoxMakerDesigns.List(selectedIndex, 1)), _
        NzText(mLstBoxMakerDesigns.List(selectedIndex, 3)), _
        NzText(mLstBoxMakerDesigns.List(selectedIndex, 4)), _
        NzText(mLstBoxMakerDesigns.List(selectedIndex, 5)), _
        NzText(mCboBoxMakerVersion.Value), ParseNumber(NzText(mTxtBoxMakerQty.Value)), _
        BoxMakerPageComponents(), "UNMAKE", report)
    ShowStatus report
    If succeeded Then RefreshBoxMakerPage
End Sub

Private Sub RenderPageRows(ByVal targetList As MSForms.ListBox, ByVal rowsData As Variant)
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim targetIndex As Long

    If targetList Is Nothing Then Exit Sub
    targetList.Clear
    If IsEmpty(rowsData) Then Exit Sub
    For rowIndex = LBound(rowsData, 1) To UBound(rowsData, 1)
        targetList.AddItem NzText(rowsData(rowIndex, 1))
        targetIndex = targetList.ListCount - 1
        For columnIndex = 2 To targetList.ColumnCount
            If columnIndex <= UBound(rowsData, 2) Then
                targetList.List(targetIndex, columnIndex - 1) = _
                    NzText(rowsData(rowIndex, columnIndex))
            End If
        Next columnIndex
    Next rowIndex
End Sub

Private Sub MoveStatusToTop()
    Const STATUS_TOP As Single = 38
    Const STATUS_HEIGHT As Single = 76
    Const CONTENT_SHIFT As Single = 86

    Dim ctl As MSForms.Control

    If mTxtStatus Is Nothing Then Exit Sub
    For Each ctl In Me.Controls
        If Not ctl Is mTxtStatus Then
            If ctl.Top >= STATUS_TOP Then ctl.Top = ctl.Top + CONTENT_SHIFT
        End If
    Next ctl
    With mTxtStatus
        .Left = 12
        .Top = STATUS_TOP
        .Width = 820
        .Height = STATUS_HEIGHT
        .ZOrder 0
    End With
    Me.ScrollHeight = Me.ScrollHeight + CONTENT_SHIFT
End Sub

Private Sub LoadCarrierChoices()
    On Error GoTo CleanExit

    Dim carriers As Variant
    Dim idx As Long
    Dim currentValue As String

    If mTxtCarrier Is Nothing Then Exit Sub
    currentValue = NzText(mTxtCarrier.Value)
    mTxtCarrier.Clear
    carriers = modCarrierSettings.GetConfiguredCarriers()
    If Not IsEmpty(carriers) Then
        For idx = LBound(carriers) To UBound(carriers)
            If Trim$(NzText(carriers(idx))) <> "" Then mTxtCarrier.AddItem NzText(carriers(idx))
        Next idx
    End If
    If currentValue <> "" Then mTxtCarrier.Value = currentValue

CleanExit:
End Sub

Private Sub LoadShippables(Optional ByVal operatorWb As Workbook = Nothing)
    On Error GoTo FailSoft

    Dim previousInv As Object
    Dim wb As Workbook

    Set wb = operatorWb
    If wb Is Nothing Then Set wb = ResolveOperatorWorkbook()
    TLap "LoadShippables start"
    Set previousInv = CurrentShippableInventoryCache()
    mLastShippablesLoadReport = vbNullString
    mShippables = modTS_Shipments.ShipmentsFormLoadShippables(wb)
    If Not ShippableRowsLoaded(mShippables) Then
        mLastShippablesLoadReport = "Local ShippingBOMView returned 0 rows. Use Refresh to rebuild from NAS; form load does not open backend workbooks."
    End If
    TLap "LoadShippables read local shippables"
    PreserveMissingShippableInventory previousInv
    TLap "LoadShippables preserve previous NAS text"
    modTS_Shipments.EvictCompletedShipmentInventoryOverlaysForShippables mShippables
    TLap "LoadShippables evict completed overlays"
    RenderShippables
    TLap "LoadShippables render"
    Exit Sub

FailSoft:
    ShowStatus "Could not load shippables: " & Err.Description
End Sub

Private Function ShippableRowsLoaded(ByVal rows As Variant) As Boolean
    On Error GoTo CleanExit

    If IsEmpty(rows) Then Exit Function
    ShippableRowsLoaded = (UBound(rows, 1) >= LBound(rows, 1))

CleanExit:
End Function

Private Function CurrentShippableInventoryCache() As Object
    Dim result As Object
    Dim r As Long
    Dim key As String
    Dim invText As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If IsEmpty(mShippables) Then
        Set CurrentShippableInventoryCache = result
        Exit Function
    End If

    For r = 1 To UBound(mShippables, 1)
        key = ShippableInventoryKey(NzText(mShippables(r, 2)), NzText(mShippables(r, 3)))
        invText = NzText(mShippables(r, 4))
        If key <> "" And Trim$(invText) <> "" Then result(key) = invText
    Next r
    Set CurrentShippableInventoryCache = result
End Function

Private Sub PreserveMissingShippableInventory(ByVal previousInv As Object)
    Dim r As Long
    Dim key As String

    If previousInv Is Nothing Then Exit Sub
    If IsEmpty(mShippables) Then Exit Sub
    For r = 1 To UBound(mShippables, 1)
        If Trim$(NzText(mShippables(r, 4))) = "" Then
            key = ShippableInventoryKey(NzText(mShippables(r, 2)), NzText(mShippables(r, 3)))
            If key <> "" Then
                If previousInv.Exists(key) Then mShippables(r, 4) = previousInv(key)
            End If
        End If
    Next r
End Sub

Private Function ShippableInventoryKey(ByVal boxName As String, ByVal versionLabel As String) As String
    boxName = Trim$(boxName)
    versionLabel = Trim$(versionLabel)
    If boxName = "" Or versionLabel = "" Then Exit Function
    ShippableInventoryKey = LCase$(boxName) & "|" & LCase$(versionLabel)
End Function

Private Sub LoadShipmentState(Optional ByVal operatorWb As Workbook = Nothing)
    Dim wb As Workbook

    Set wb = operatorWb
    If wb Is Nothing Then Set wb = ResolveOperatorWorkbook()
    TLap "LoadShipmentState start"
    RenderLineList mLstShipments, modTS_Shipments.ShipmentsFormLoadLines(False, wb)
    TLap "LoadShipmentState active lines"
    RenderLineList mLstHold, modTS_Shipments.ShipmentsFormLoadLines(True, wb)
    TLap "LoadShipmentState hold lines"
    EvictOrphanedActiveOverlays
    TLap "LoadShipmentState evict orphaned overlays"
    UpdateSyncStateLabel
    TLap "LoadShipmentState update sync label"
End Sub

Private Sub LoadShipmentLineState(Optional ByVal operatorWb As Workbook = Nothing)
    Dim wb As Workbook

    Set wb = operatorWb
    If wb Is Nothing Then Set wb = ResolveOperatorWorkbook()
    TLap "LoadShipmentLineState start"
    RenderLineList mLstShipments, modTS_Shipments.ShipmentsFormLoadLines(False, wb)
    TLap "LoadShipmentLineState active lines"
    RenderLineList mLstHold, modTS_Shipments.ShipmentsFormLoadLines(True, wb)
    TLap "LoadShipmentLineState hold lines"
    EvictOrphanedActiveOverlays
    TLap "LoadShipmentLineState evict orphaned overlays"
    UpdateSyncStateLabel
    TLap "LoadShipmentLineState update sync label"
End Sub

Private Sub RenderShippables()
    On Error GoTo FailSoft

    Dim filterText As String
    Dim shownCount As Long
    Dim r As Long
    Dim idx As Long
    Dim displayRows As Variant

    mLstShippables.Clear
    If IsEmpty(mShippables) Then Exit Sub
    filterText = LCase$(Trim$(NzText(mTxtPicker.Value)))
    For r = 1 To UBound(mShippables, 1)
        If ShippableMatchesFilter(r, filterText) Then shownCount = shownCount + 1
    Next r
    If shownCount = 0 Then Exit Sub

    ReDim displayRows(0 To shownCount - 1, 0 To 7)
    idx = 0
    For r = 1 To UBound(mShippables, 1)
        If Not ShippableMatchesFilter(r, filterText) Then GoTo NextRow
        displayRows(idx, 0) = NzText(mShippables(r, 2))
        displayRows(idx, 1) = NzText(mShippables(r, 3))
        displayRows(idx, 2) = DisplayQtyText(NzText(mShippables(r, 4)))
        displayRows(idx, 3) = DisplayQtyText(NzText(mShippables(r, 8)))
        displayRows(idx, 4) = DisplayQtyText(CStr(LockedShipmentQtyForShippable( _
            NzText(mShippables(r, 1)), NzText(mShippables(r, 2)), NzText(mShippables(r, 3)))))
        displayRows(idx, 5) = NzText(mShippables(r, 5))
        displayRows(idx, 6) = NzText(mShippables(r, 6))
        displayRows(idx, 7) = NzText(mShippables(r, 1))
        idx = idx + 1
NextRow:
    Next r
    For r = 0 To UBound(displayRows, 1)
        displayRows(r, 1) = DisplayVersionOrNA(NzText(displayRows(r, 1)))
    Next r
    mLstShippables.List = displayRows
    UpdateSyncStateLabel
    Exit Sub

FailSoft:
    ShowStatus "Shippable render failed: " & Err.Description
    UpdateSyncStateLabel
End Sub

Private Function ShippableMatchesFilter(ByVal rowIndex As Long, ByVal filterText As String) As Boolean
    Dim haystack As String

    If filterText = "" Then
        ShippableMatchesFilter = True
        Exit Function
    End If
    haystack = LCase$(NzText(mShippables(rowIndex, 2)) & " " & _
                      NzText(mShippables(rowIndex, 3)) & " " & _
                      NzText(mShippables(rowIndex, 6)) & " " & _
                      NzText(mShippables(rowIndex, 7)))
    ShippableMatchesFilter = (InStr(1, haystack, filterText, vbTextCompare) > 0)
End Function

Private Sub RenderLineList(ByVal lst As MSForms.ListBox, ByVal rowsData As Variant)
    On Error GoTo FailSoft

    Dim r As Long
    Dim displayRows As Variant

    lst.Clear
    If IsEmpty(rowsData) Then Exit Sub
    ReDim displayRows(0 To UBound(rowsData, 1) - 1, 0 To 11)
    For r = 1 To UBound(rowsData, 1)
        displayRows(r - 1, 0) = NzText(rowsData(r, 1))
        displayRows(r - 1, 1) = NzText(rowsData(r, 2))
        displayRows(r - 1, 2) = FormatQuantity(ParseNumber(NzText(rowsData(r, 3))))
        displayRows(r - 1, 3) = NzText(rowsData(r, 4))
        displayRows(r - 1, 4) = NzText(rowsData(r, 9))
        If Trim$(NzText(rowsData(r, 11))) <> "" Then displayRows(r - 1, 5) = "Yes" Else displayRows(r - 1, 5) = ""
        displayRows(r - 1, 6) = NzText(rowsData(r, 6))
        displayRows(r - 1, 7) = NzText(rowsData(r, 7))
        displayRows(r - 1, 8) = NzText(rowsData(r, 10))
        displayRows(r - 1, 9) = NzText(rowsData(r, 5))
        displayRows(r - 1, 10) = NzText(rowsData(r, 8))
        displayRows(r - 1, 11) = NzText(rowsData(r, 11))
    Next r
    lst.List = displayRows
    Exit Sub

FailSoft:
End Sub

Private Sub LoadSelectedShippable()
    If mLstShippables.ListIndex < 0 Then Exit Sub
    mTxtBox.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 0))
    mTxtVersion.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 1))
    mTxtUom.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 5))
    mTxtLocation.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 6))
    mTxtSystemKey.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 7))
    If Trim$(NzText(mTxtQty.Value)) = "" Then mTxtQty.Value = "1"
    mTxtDescription.Value = NzText(mTxtVersion.Value)
End Sub

Private Sub LoadSelectedLine(ByVal lst As MSForms.ListBox)
    If lst Is Nothing Then Exit Sub
    If lst.ListIndex < 0 Then Exit Sub
    mTxtRef.Value = NzText(lst.List(lst.ListIndex, 0))
    mTxtBox.Value = NzText(lst.List(lst.ListIndex, 1))
    mTxtQty.Value = NzText(lst.List(lst.ListIndex, 2))
    mTxtUom.Value = NzText(lst.List(lst.ListIndex, 3))
    mTxtLocation.Value = NzText(lst.List(lst.ListIndex, 9))
    mTxtSystemKey.Value = NzText(lst.List(lst.ListIndex, 6))
    mTxtDescription.Value = NzText(lst.List(lst.ListIndex, 7))
    mTxtVersion.Value = NzText(lst.List(lst.ListIndex, 7))
    mTxtCarrier.Value = NzText(lst.List(lst.ListIndex, 8))
End Sub

Private Sub CommitCurrentLine(ByVal actionName As String)
    On Error GoTo FailSoft
    If Not RequireActionContext() Then Exit Sub

    Dim report As String
    Dim rowIndex As Long
    Dim selectedShipmentCount As Long
    Dim ok As Boolean
    Dim displayedAvailableQty As String
    Dim displayedNasQty As String
    Dim operatorWb As Workbook
    Dim quietStarted As Boolean
    Dim failureReason As String
    Dim startedAt As Single
    Dim elapsedMs As Long

    mActionTimer.Start
    TLap "CommitCurrentLine " & UCase$(Trim$(actionName)) & " start"
    startedAt = Timer
    actionName = UCase$(Trim$(actionName))
    selectedShipmentCount = SelectedListTableRowCount(mLstShipments)
    If actionName = "ADD" And selectedShipmentCount > 1 Then
        ShowStatus "Select at most one shipment row before using Add."
        Exit Sub
    End If
    If actionName = "UPDATE" And selectedShipmentCount <> 1 Then
        ShowStatus "Select exactly one shipment row before using Update Row."
        Exit Sub
    End If
    rowIndex = SelectedShipmentTableRow()
    displayedAvailableQty = SelectedShippableProjectedInventoryText()
    displayedNasQty = SelectedShippableNasInventoryText()
    Set operatorWb = ResolveOperatorWorkbook()
    ShowPersistencePending "Saving shipment row changes to the warehouse server..."
    If Not RequireActionContext() Then Exit Sub
    modUiQuiet.BeginQuietUi operatorWb
    quietStarted = True
    TLap "CommitCurrentLine resolved selected row/operator"
    ok = modTS_Shipments.ShipmentsFormCommitLine("SHIP", _
                                                 actionName, _
                                                 rowIndex, _
                                                 NzText(mTxtRef.Value), _
                                                 NzText(mTxtBox.Value), _
                                                 ParseNumber(NzText(mTxtQty.Value)), _
                                                 NzText(mTxtSystemKey.Value), _
                                                 NzText(mTxtUom.Value), _
                                                 NzText(mTxtLocation.Value), _
                                                 NzText(mTxtVersion.Value), _
                                                 NzText(mTxtCarrier.Value), _
                                                 report, _
                                                 displayedAvailableQty, _
                                                 mShippables, _
                                                 operatorWb, _
                                                 displayedNasQty)
    TLap "CommitCurrentLine backend call"
    elapsedMs = mActionTimer.ElapsedMilliseconds(startedAt)
    report = AppendTiming(report, elapsedMs)
    If TimingSummary() <> "" Then report = report & vbCrLf & TimingSummary()
    RefreshAfterAction report, ok
    modUiQuiet.EndQuietUi
    quietStarted = False
    Exit Sub

FailSoft:
    failureReason = Err.Description
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    On Error GoTo 0
    ShowStatus "Shipment row action failed: " & failureReason
End Sub

Private Function SelectedShippableNasInventoryText() As String
    Dim r As Long
    Dim systemKey As String
    Dim boxName As String
    Dim versionLabel As String

    systemKey = NzText(mTxtSystemKey.Value)
    boxName = Trim$(NzText(mTxtBox.Value))
    versionLabel = Trim$(NzText(mTxtVersion.Value))

    If mLstShippables Is Nothing Then Exit Function
    If mLstShippables.ListIndex >= 0 Then
        If StrComp(NzText(mLstShippables.List(mLstShippables.ListIndex, 7)), systemKey, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mLstShippables.List(mLstShippables.ListIndex, 0))), boxName, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mLstShippables.List(mLstShippables.ListIndex, 1))), versionLabel, vbTextCompare) = 0 Then
            SelectedShippableNasInventoryText = NzText(mLstShippables.List(mLstShippables.ListIndex, 2))
            Exit Function
        End If
        If StrComp(NzText(mLstShippables.List(mLstShippables.ListIndex, 7)), systemKey, vbTextCompare) = 0 Then
            SelectedShippableNasInventoryText = NzText(mLstShippables.List(mLstShippables.ListIndex, 2))
            Exit Function
        End If
    End If

    If IsEmpty(mShippables) Then Exit Function
    For r = 1 To UBound(mShippables, 1)
        If StrComp(NzText(mShippables(r, 1)), systemKey, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mShippables(r, 2))), boxName, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mShippables(r, 3))), versionLabel, vbTextCompare) = 0 Then
            SelectedShippableNasInventoryText = NzText(mShippables(r, 4))
            Exit Function
        End If
    Next r
    For r = 1 To UBound(mShippables, 1)
        If StrComp(NzText(mShippables(r, 1)), systemKey, vbTextCompare) = 0 Then
            SelectedShippableNasInventoryText = NzText(mShippables(r, 4))
            Exit Function
        End If
    Next r
End Function

Private Function SelectedShippableProjectedInventoryText() As String
    Dim r As Long
    Dim systemKey As String
    Dim boxName As String
    Dim versionLabel As String

    systemKey = NzText(mTxtSystemKey.Value)
    boxName = Trim$(NzText(mTxtBox.Value))
    versionLabel = Trim$(NzText(mTxtVersion.Value))

    If mLstShippables Is Nothing Then Exit Function
    If mLstShippables.ListIndex >= 0 Then
        If StrComp(NzText(mLstShippables.List(mLstShippables.ListIndex, 7)), systemKey, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mLstShippables.List(mLstShippables.ListIndex, 0))), boxName, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mLstShippables.List(mLstShippables.ListIndex, 1))), versionLabel, vbTextCompare) = 0 Then
            SelectedShippableProjectedInventoryText = NzText(mLstShippables.List(mLstShippables.ListIndex, 3))
            Exit Function
        End If
    End If

    If IsEmpty(mShippables) Then Exit Function
    For r = 1 To UBound(mShippables, 1)
        If StrComp(NzText(mShippables(r, 1)), systemKey, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mShippables(r, 2))), boxName, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mShippables(r, 3))), versionLabel, vbTextCompare) = 0 Then
            SelectedShippableProjectedInventoryText = NzText(mShippables(r, 8))
            Exit Function
        End If
    Next r
End Function

Private Function SelectedShipmentTableRow() As Long
    SelectedShipmentTableRow = SingleSelectedListTableRow(mLstShipments)
End Function

Private Function SelectedHoldTableRow() As Long
    SelectedHoldTableRow = SingleSelectedListTableRow(mLstHold)
End Function

Private Function SingleSelectedListTableRow(ByVal lst As MSForms.ListBox) As Long
    Dim i As Long
    Dim tableRow As Long

    If lst Is Nothing Then Exit Function
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            tableRow = CLng(Val(NzText(lst.List(i, 10))))
            If tableRow > 0 Then
                If SingleSelectedListTableRow <> 0 Then
                    SingleSelectedListTableRow = 0
                    Exit Function
                End If
                SingleSelectedListTableRow = tableRow
            End If
        End If
    Next i
End Function

Private Function SelectedListTableRowCount(ByVal lst As MSForms.ListBox) As Long
    Dim rows As Variant

    rows = SelectedListTableRows(lst)
    If IsEmpty(rows) Then Exit Function
    SelectedListTableRowCount = UBound(rows) - LBound(rows) + 1
End Function

Private Function SelectedListTableRows(ByVal lst As MSForms.ListBox) As Variant
    Dim rowIndexes() As Long
    Dim i As Long
    Dim countRows As Long
    Dim tableRow As Long

    If lst Is Nothing Then Exit Function
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            tableRow = CLng(Val(NzText(lst.List(i, 10))))
            If tableRow > 0 Then
                countRows = countRows + 1
                ReDim Preserve rowIndexes(1 To countRows)
                rowIndexes(countRows) = tableRow
            End If
        End If
    Next i
    If countRows > 0 Then SelectedListTableRows = rowIndexes
End Function

Private Sub RefreshAfterAction(ByVal report As String, ByVal ok As Boolean)
    Dim previousPointer As Long
    Dim operatorWb As Workbook

    Set operatorWb = ResolveOperatorWorkbook()
    previousPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass
    mLoading = True
    LoadShipmentLineState
    TLap "RefreshAfterAction load shipment lines"
    RefreshProjectedShippableInventory
    TLap "RefreshAfterAction refresh projected"
    mLoading = False
    modTS_Shipments.EnforceShippingSupportSheetsHidden operatorWb
    TLap "RefreshAfterAction hide support sheets"
    Me.MousePointer = previousPointer
    ShowStatus report
    If Not ok And report <> "" Then MsgBox report, vbExclamation
End Sub

Private Sub mTxtPicker_Change()
    If mLoading Then Exit Sub
    RenderShippables
End Sub

Private Sub mLstShippables_Click()
    If mLoading Then Exit Sub
    LoadSelectedShippable
End Sub

Private Sub mLstShipments_Click()
    If mLoading Then Exit Sub
    LoadSelectedLine mLstShipments
End Sub

Private Sub mLstHold_Click()
    If mLoading Then Exit Sub
    LoadSelectedLine mLstHold
End Sub

Private Sub mChkUseExisting_Click()
    If mLoading Then Exit Sub
    mActionTimer.Start
    TLap "UseExisting click start"
    modTS_Shipments.ShipmentsFormSetUseExistingInventory CBool(mChkUseExisting.Value)
    LoadShipmentState
    TLap "UseExisting load shipment state"
    RefreshProjectedShippableInventory
    TLap "UseExisting refresh projected"
    ShowStatus "Use existing changed." & vbCrLf & TimingSummary()
End Sub

Private Sub mBtnRefresh_Click()
    Dim report As String
    Dim ok As Boolean
    Dim operatorWb As Workbook
    Dim nasBeforeRefresh As Object
    Dim nasAfterRefresh As Object
    Dim nasStatus As String

    mActionTimer.Start
    TLap "Refresh click start"
    Set nasBeforeRefresh = ShippableNasSnapshot()
    Set operatorWb = ResolveOperatorWorkbook()
    ok = modTS_Shipments.ShipmentsFormRefreshRuntimeInventoryForWorkbook(operatorWb, report, vbNullString, True)
    TLap "Refresh backend refresh"
    InitializeFromShipping True
    Set nasAfterRefresh = ShippableNasSnapshot()
    nasStatus = ShippableNasChangeSummary(nasBeforeRefresh, nasAfterRefresh)
    TLap "Refresh reinitialize form"
    If Trim$(report) <> "" Then
        ShowStatus nasStatus & vbCrLf & "Shipments form refreshed. " & report & vbCrLf & TimingSummary()
    Else
        ShowStatus nasStatus & vbCrLf & "Shipments form refreshed." & vbCrLf & TimingSummary()
    End If
    If Not ok And Trim$(report) <> "" Then MsgBox report, vbExclamation
End Sub

Private Sub mBtnHistory_Click()
    Dim historyText As String

    historyText = modTS_Shipments.ShipmentsFormRecentHistoryText(20)
    If Trim$(historyText) = "" Then historyText = "No shipment history was found."
    ShowStatus historyText
    MsgBox historyText, vbInformation, "Shipments History"
End Sub

Private Sub mBtnHistorySheet_Click()
    Dim report As String

    report = modTS_Shipments.ShipmentsFormExportHistoryToSheet(100, ResolveOperatorWorkbook())
    If Trim$(report) = "" Then report = "No shipment history was exported."
    ShowStatus report
    If InStr(1, report, "failed", vbTextCompare) > 0 Then MsgBox report, vbExclamation, "Shipments History"
End Sub

Private Sub mBtnAdd_Click()
    CommitCurrentLine "ADD"
End Sub

Private Sub mBtnUpdate_Click()
    CommitCurrentLine "UPDATE"
End Sub

Private Sub mBtnRemove_Click()
    RemoveSelectedShipmentRows
End Sub

Private Sub mBtnHold_Click()
    MoveSelectedShipmentHold True
End Sub

Private Sub mBtnReturn_Click()
    MoveSelectedShipmentHold False
End Sub

Private Sub MoveSelectedShipmentHold(ByVal moveToHold As Boolean)
    On Error GoTo FailSoft
    If Not RequireActionContext() Then Exit Sub

    Dim lst As MSForms.ListBox
    Dim report As String
    Dim ok As Boolean
    Dim selectedRows As Variant

    mActionTimer.Start
    TLap "Hold/Return click start"
    If moveToHold Then
        Set lst = mLstShipments
    Else
        Set lst = mLstHold
    End If
    selectedRows = SelectedListTableRows(lst)
    If IsEmpty(selectedRows) Then
        If moveToHold Then
            ShowStatus "Select one or more shipment row(s) to send to Hold."
        Else
            ShowStatus "Select one or more Hold row(s) to return."
        End If
        Exit Sub
    End If

    ok = modTS_Shipments.ShipmentsFormMoveHoldRows(selectedRows, moveToHold, report)
    TLap "Hold/Return backend move"
    If TimingSummary() <> "" Then report = report & vbCrLf & TimingSummary()
    RefreshAfterAction report, ok
    Exit Sub

FailSoft:
    ShowStatus "Hold action failed: " & Err.Description
End Sub

Private Sub RemoveSelectedShipmentRows()
    On Error GoTo FailSoft
    If Not RequireActionContext() Then Exit Sub

    Dim report As String
    Dim rowReport As String
    Dim selectedRows As Variant
    Dim ok As Boolean
    Dim allOk As Boolean
    Dim i As Long
    Dim rowIndex As Long
    Dim removedCount As Long
    Dim operatorWb As Workbook
    Dim startedAt As Single
    Dim elapsedMs As Long

    mActionTimer.Start
    TLap "Remove selected click start"
    selectedRows = SelectedListTableRows(mLstShipments)
    If IsEmpty(selectedRows) Then
        ShowStatus "Select one or more shipment row(s) to remove."
        Exit Sub
    End If

    startedAt = Timer
    Set operatorWb = ResolveOperatorWorkbook()
    allOk = True
    For i = UBound(selectedRows) To LBound(selectedRows) Step -1
        If Not RequireActionContext() Then Exit Sub
        rowIndex = CLng(selectedRows(i))
        rowReport = vbNullString
        ok = modTS_Shipments.ShipmentsFormCommitLine("SHIP", _
                                                     "DELETE", _
                                                     rowIndex, _
                                                     "", _
                                                     "", _
                                                     0, _
                                                     0, _
                                                     "", _
                                                     "", _
                                                     "", _
                                                     "", _
                                                     rowReport, _
                                                     operatorWb:=operatorWb)
        If ok Then
            removedCount = removedCount + 1
        Else
            allOk = False
            If Trim$(rowReport) = "" Then rowReport = "Unable to remove shipment row " & CStr(rowIndex) & "."
            AppendLocalStatus report, rowReport
            Exit For
        End If
    Next i

    TLap "Remove selected backend call"
    elapsedMs = mActionTimer.ElapsedMilliseconds(startedAt)
    If allOk Then report = "Removed " & CStr(removedCount) & " shipment row(s)."
    report = AppendTiming(report, elapsedMs)
    If TimingSummary() <> "" Then report = report & vbCrLf & TimingSummary()
    RefreshAfterAction report, allOk
    Exit Sub

FailSoft:
    ShowStatus "Remove action failed: " & Err.Description
End Sub

Private Sub mBtnStage_Click()
    RunShippingAction True
End Sub

Private Sub mBtnSend_Click()
    On Error GoTo FailSoft
    If Not RequireActionContext() Then Exit Sub

    Dim previousPointer As Long
    Dim quietStarted As Boolean
    Dim startedAt As Single
    Dim elapsedMs As Long
    Dim report As String
    Dim ok As Boolean
    Dim selectedRows As Variant

    mActionTimer.Start
    TLap "Shipments Sent click start"
    previousPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass
    ShowPersistencePending "Saving shipment events and refreshing warehouse inventory..."
    If Not RequireActionContext() Then Me.MousePointer = previousPointer: Exit Sub
    modUiQuiet.BeginQuietUi mOperatorWorkbook
    quietStarted = True
    startedAt = Timer
    ok = modShippingPostingService.ExecuteShipmentsSent( _
        mOperatorWorkbook, selectedRows, NzText(mTxtCarrier.Value), report)
    TLap "Shipments Sent backend call"
    elapsedMs = mActionTimer.ElapsedMilliseconds(startedAt)
    Me.MousePointer = previousPointer
    LoadShipmentState mOperatorWorkbook
    TLap "Shipments Sent load shipment state"
    If ok Then mTxtRef.Value = vbNullString
    If ok Then LoadShippables mOperatorWorkbook
    If ok Then TLap "Shipments Sent reload canonical shippables"
    If ok Then RefreshProjectedShippableInventory
    If ok Then TLap "Shipments Sent refresh projected"
    If ok Then ArmAutoSync
    If quietStarted Then
        modUiQuiet.EndQuietUi
        quietStarted = False
    End If
    report = AppendTiming(report, elapsedMs)
    If TimingSummary() <> "" Then report = report & vbCrLf & TimingSummary()
    ShowStatus report
    If report <> "" And ShouldShowShippingActionPopup(report, ok) Then _
        MsgBox report, IIf(ok, vbInformation, vbExclamation)
    Exit Sub

FailSoft:
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    Me.MousePointer = previousPointer
    On Error GoTo 0
    ShowStatus "Shipping action failed: " & Err.Description
End Sub

Private Sub RunShippingAction(ByVal stageOnly As Boolean)
    On Error GoTo FailSoft
    If Not RequireActionContext() Then Exit Sub

    Dim previousPointer As Long
    Dim quietStarted As Boolean
    Dim startedAt As Single
    Dim elapsedMs As Long
    Dim report As String
    Dim ok As Boolean
    Dim selectedRows As Variant

    mActionTimer.Start
    TLap IIf(stageOnly, "To Shipments", "Shipments Sent") & " click start"
    If stageOnly Then selectedRows = SelectedListTableRows(mLstShipments)
    If stageOnly And IsEmpty(selectedRows) Then
        ShowStatus "Select shipment row(s) first."
        Exit Sub
    End If
    previousPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass
    ShowPersistencePending IIf(stageOnly, _
        "Saving selected rows to Shipments...", _
        "Saving shipment events and refreshing warehouse inventory...")
    If Not RequireActionContext() Then Me.MousePointer = previousPointer: Exit Sub
    modUiQuiet.BeginQuietUi ResolveOperatorWorkbook()
    quietStarted = True
    startedAt = Timer
    If stageOnly Then
        ok = modTS_Shipments.ShipmentsFormRunToShipmentsRows(selectedRows, NzText(mTxtCarrier.Value), report)
    Else
        ok = modShippingPostingService.ExecuteShipmentsSent( _
            mOperatorWorkbook, selectedRows, NzText(mTxtCarrier.Value), report)
    End If
    TLap IIf(stageOnly, "To Shipments", "Shipments Sent") & " backend call"
    elapsedMs = mActionTimer.ElapsedMilliseconds(startedAt)
    Me.MousePointer = previousPointer
    LoadShipmentState
    TLap IIf(stageOnly, "To Shipments", "Shipments Sent") & " load shipment state"
    If ok And Not stageOnly Then mTxtRef.Value = vbNullString
    If ok Then RefreshProjectedShippableInventory
    If ok Then TLap IIf(stageOnly, "To Shipments", "Shipments Sent") & " refresh projected"
    If ok And Not stageOnly Then ArmAutoSync
    If quietStarted Then
        modUiQuiet.EndQuietUi
        quietStarted = False
    End If
    report = AppendTiming(report, elapsedMs)
    If TimingSummary() <> "" Then report = report & vbCrLf & TimingSummary()
    ShowStatus report
    If report <> "" And ShouldShowShippingActionPopup(report, ok) Then MsgBox report, IIf(ok, vbInformation, vbExclamation)
    Exit Sub

FailSoft:
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    Me.MousePointer = previousPointer
    On Error GoTo 0
    ShowStatus "Shipping action failed: " & Err.Description
End Sub

Private Function ShouldShowShippingActionPopup(ByVal report As String, ByVal ok As Boolean) As Boolean
    ShouldShowShippingActionPopup = True
    If Not ok Then Exit Function
    If InStr(1, report, "selected row(s) were already locked", vbTextCompare) > 0 Then
        ShouldShowShippingActionPopup = False
    End If
End Function

Private Sub RefreshProjectedShippableInventory()
    On Error GoTo CleanExit

    Dim r As Long
    Dim activeQty As Double
    Dim backendText As String
    Dim overlayText As String
    Dim projectedQty As Double
    Dim packageSystemKey As String
    Dim versionLabel As String
    Dim hasSentOverlay As Boolean
    Dim overlayIncludesReservation As Boolean

    If IsEmpty(mShippables) Then Exit Sub
    If Not mUseInjectedReservationTotalsForTest Then Set mNasReservationTotals = modTS_Shipments.ShipmentsFormLoadNasReservationTotals()
    For r = 1 To UBound(mShippables, 1)
        packageSystemKey = NzText(mShippables(r, 1))
        versionLabel = NzText(mShippables(r, 3))
        backendText = NzText(mShippables(r, 4))
        activeQty = ActiveShipmentQtyForShippable(packageSystemKey, NzText(mShippables(r, 2)), versionLabel)
        overlayText = modTS_Shipments.PendingSystemKeyInventoryOverlayText(packageSystemKey, versionLabel, backendText)
        hasSentOverlay = False
        If hasSentOverlay And activeQty > 0.0000001 Then overlayText = vbNullString
        If Trim$(overlayText) <> "" And IsNumeric(overlayText) Then
            overlayIncludesReservation = False
            projectedQty = modTS_Shipments.ShipmentsProjectedDisplayQtyWithOverlay(ParseNumber(backendText), activeQty, CDbl(overlayText), overlayIncludesReservation)
        Else
            projectedQty = modTS_Shipments.ShipmentsProjectedDisplayQty(ParseNumber(backendText), activeQty)
        End If
        mShippables(r, 8) = FormatQuantity(projectedQty)
    Next r
    RenderShippables

CleanExit:
    UpdateSyncStateLabel
End Sub

Public Function TestRefreshProjectedInventory(ByVal shippablesArray As Variant, _
                                              ByVal shipmentsListData As Variant, _
                                              Optional ByVal holdListData As Variant, _
                                              Optional ByVal reservationTotals As Object) As Variant
    On Error GoTo CleanFail

    If Not mBuilt Then BuildLayout

    mShippables = shippablesArray
    If reservationTotals Is Nothing Then
        Set mNasReservationTotals = CreateObject("Scripting.Dictionary")
        mNasReservationTotals.CompareMode = vbTextCompare
    Else
        Set mNasReservationTotals = reservationTotals
    End If
    RenderLineList mLstShipments, shipmentsListData
    RenderLineList mLstHold, holdListData
    mUseInjectedReservationTotalsForTest = True
    RefreshProjectedShippableInventory
    mUseInjectedReservationTotalsForTest = False
    TestRefreshProjectedInventory = mShippables
    Exit Function

CleanFail:
    mUseInjectedReservationTotalsForTest = False
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function TestReadProjectedText(ByVal rowIndex As Long) As String
    If IsEmpty(mShippables) Then Exit Function
    If rowIndex < 1 Or rowIndex > UBound(mShippables, 1) Then Exit Function
    TestReadProjectedText = NzText(mShippables(rowIndex, 8))
End Function

Public Function TestPendingSyncCount(ByVal shippablesArray As Variant, _
                                     ByVal shipmentsListData As Variant) As Long
    If Not mBuilt Then BuildLayout
    mShippables = shippablesArray
    RenderLineList mLstShipments, shipmentsListData
    TestPendingSyncCount = PendingShipmentSyncCount()
End Function

Public Function TestInitializeForWorkbook(ByVal operatorWb As Workbook) As String
    If operatorWb Is Nothing Then Exit Function
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    TestInitializeForWorkbook = _
        "OK|BoundWorkbook=" & mOperatorWorkbook.Name & _
        "|Caption=" & Me.Caption
End Function

Public Function TestStatusAnchorAfterResize() As String
    Dim originalHeight As Double
    Dim originalTop As Double
    Dim originalStatusHeight As Double
    Dim searchTop As Double
    Dim grownTop As Double
    Dim grownStatusHeight As Double
    Dim topStable As Boolean
    Dim heightStable As Boolean
    Dim aboveSearch As Boolean

    If Not mBuilt Then BuildLayout
    originalHeight = Me.Height
    originalTop = mTxtStatus.Top
    originalStatusHeight = mTxtStatus.Height
    searchTop = mTxtPicker.Top

    Me.Height = originalHeight + 180
    mAnchors.ResizeControls
    grownTop = mTxtStatus.Top
    grownStatusHeight = mTxtStatus.Height
    topStable = Abs(grownTop - originalTop) < 0.5
    heightStable = Abs(grownStatusHeight - originalStatusHeight) < 0.5
    aboveSearch = grownTop + grownStatusHeight < searchTop

    TestStatusAnchorAfterResize = "OK|TopStable=" & CStr(topStable) & _
        "|HeightStable=" & CStr(heightStable) & _
        "|AboveSearch=" & CStr(aboveSearch) & _
        "|OriginalTop=" & Format$(originalTop, "0.00") & _
        "|GrownTop=" & Format$(grownTop, "0.00")

    Me.Height = originalHeight
    mAnchors.ResizeControls
End Function

Public Function TestSystemKeyIdentityContract() As String
    Dim probeKey As String
    Dim reservationKey As String
    Dim controlReady As Boolean
    Dim valuePreserved As Boolean
    Dim reservationUsesKey As Boolean

    If Not mBuilt Then BuildLayout
    probeKey = "SK-R1-SHIPPING-PROBE"
    controlReady = (StrComp(mTxtSystemKey.Name, "txtSystemKey", vbBinaryCompare) = 0)
    mTxtSystemKey.Value = probeKey
    valuePreserved = (StrComp(NzText(mTxtSystemKey.Value), probeKey, vbBinaryCompare) = 0)
    reservationKey = modTS_Shipments.ShipmentsFormReservationKey(probeKey, "v1")
    reservationUsesKey = (StrComp(reservationKey, LCase$(probeKey) & "|v1", vbBinaryCompare) = 0)

    TestSystemKeyIdentityContract = "OK|ControlReady=" & CStr(controlReady) & _
        "|ValuePreserved=" & CStr(valuePreserved) & _
        "|ReservationUsesKey=" & CStr(reservationUsesKey)
End Function

Public Function TestRunShipmentsSentActionForWorkbook(ByVal operatorWb As Workbook, _
                                                       ByVal carrierValue As String, _
                                                       Optional ByVal activatedWb As Workbook = Nothing) As String
    Dim visibleNas As String
    Dim visibleProjected As String
    Dim visibleLocked As String

    SetOperatorWorkbook operatorWb
    InitializeFromShipping True
    mTxtCarrier.Value = carrierValue
    If Not activatedWb Is Nothing Then activatedWb.Activate
    mBtnSend_Click
    If Not mLstShippables Is Nothing Then
        If mLstShippables.ListCount > 0 Then
            visibleNas = NzText(mLstShippables.List(0, 2))
            visibleProjected = NzText(mLstShippables.List(0, 3))
            visibleLocked = NzText(mLstShippables.List(0, 4))
        End If
    End If
    TestRunShipmentsSentActionForWorkbook = _
        "Status=" & mTxtStatus.Text & _
        "; BoundWorkbook=" & mOperatorWorkbook.Name & _
        "; VisibleNas=" & visibleNas & _
        "; VisibleProjected=" & visibleProjected & _
        "; VisibleLocked=" & visibleLocked
End Function

Private Sub EvictOrphanedActiveOverlays()
    Dim r As Long
    Dim packageSystemKey As String
    Dim versionLabel As String

    If IsEmpty(mShippables) Then Exit Sub
    If mLstShipments Is Nothing Then Exit Sub
    For r = 1 To UBound(mShippables, 1)
        packageSystemKey = NzText(mShippables(r, 1))
        versionLabel = NzText(mShippables(r, 3))
        If packageSystemKey <> "" And Trim$(versionLabel) <> "" Then
            If Not HasActiveShipmentLineForSystemKey(packageSystemKey, versionLabel) Then
                modTS_Shipments.ClearActiveOverlayForSystemKeyVersion packageSystemKey, versionLabel
            End If
        End If
    Next r
End Sub

Private Function HasActiveShipmentLineForSystemKey(ByVal packageSystemKey As String, ByVal versionLabel As String) As Boolean
    Dim i As Long
    Dim rowVersion As String

    versionLabel = LCase$(Trim$(versionLabel))
    If Not mLstShipments Is Nothing Then
        For i = 0 To mLstShipments.ListCount - 1
            If StrComp(NzText(mLstShipments.List(i, 6)), packageSystemKey, vbTextCompare) = 0 Then
                rowVersion = LCase$(Trim$(NzText(mLstShipments.List(i, 7))))
                If rowVersion = versionLabel Then
                    HasActiveShipmentLineForSystemKey = True
                    Exit Function
                End If
            End If
        Next i
    End If

    If Not mLstHold Is Nothing Then
        For i = 0 To mLstHold.ListCount - 1
            If StrComp(NzText(mLstHold.List(i, 6)), packageSystemKey, vbTextCompare) = 0 Then
                rowVersion = LCase$(Trim$(NzText(mLstHold.List(i, 7))))
                If rowVersion = versionLabel Then
                    If Trim$(NzText(mLstHold.List(i, 11))) <> "" Then
                        HasActiveShipmentLineForSystemKey = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
End Function

Private Function ShipmentListSystemKeyMatchesShippable(ByVal listIndex As Long, _
                                                 ByVal packageSystemKey As String, _
                                                 ByVal boxName As String, _
                                                 ByVal versionLabel As String) As Boolean
    ShipmentListSystemKeyMatchesShippable = ShipmentListBoxSystemKeyMatchesShippable(mLstShipments, listIndex, packageSystemKey, boxName, versionLabel)
End Function

Private Function ShipmentListBoxSystemKeyMatchesShippable(ByVal lineList As MSForms.ListBox, _
                                                    ByVal listIndex As Long, _
                                                    ByVal packageSystemKey As String, _
                                                    ByVal boxName As String, _
                                                    ByVal versionLabel As String) As Boolean
    Dim rowBox As String
    Dim rowVersion As String
    Dim rowSystemKey As String

    If lineList Is Nothing Then Exit Function
    rowBox = LCase$(Trim$(NzText(lineList.List(listIndex, 1))))
    rowVersion = LCase$(Trim$(NzText(lineList.List(listIndex, 7))))
    rowSystemKey = NzText(lineList.List(listIndex, 6))
    If packageSystemKey <> "" Then
        ShipmentListBoxSystemKeyMatchesShippable = _
            (StrComp(rowSystemKey, packageSystemKey, vbTextCompare) = 0 And rowVersion = versionLabel)
    Else
        ShipmentListBoxSystemKeyMatchesShippable = (rowBox = boxName And rowVersion = versionLabel)
    End If
End Function

Private Function ActiveShipmentQtyForShippable(ByVal packageSystemKey As String, ByVal boxName As String, ByVal versionLabel As String) As Double
    Dim i As Long

    boxName = LCase$(Trim$(boxName))
    versionLabel = LCase$(Trim$(versionLabel))
    If Not mLstShipments Is Nothing Then
        For i = 0 To mLstShipments.ListCount - 1
            If ShipmentListSystemKeyMatchesShippable(i, packageSystemKey, boxName, versionLabel) Then
                ActiveShipmentQtyForShippable = ActiveShipmentQtyForShippable + ParseNumber(NzText(mLstShipments.List(i, 2)))
            End If
        Next i
    End If
    If Not mLstHold Is Nothing Then
        For i = 0 To mLstHold.ListCount - 1
            If ShipmentListBoxSystemKeyMatchesShippable(mLstHold, i, packageSystemKey, boxName, versionLabel) Then
                If Trim$(NzText(mLstHold.List(i, 11))) <> "" Then
                    ActiveShipmentQtyForShippable = ActiveShipmentQtyForShippable + ParseNumber(NzText(mLstHold.List(i, 2)))
                End If
            End If
        Next i
    End If
End Function

Private Function UnreservedShipmentQtyForShippable(ByVal packageSystemKey As String, ByVal boxName As String, ByVal versionLabel As String) As Double
    Dim i As Long

    If mLstShipments Is Nothing Then Exit Function
    boxName = LCase$(Trim$(boxName))
    versionLabel = LCase$(Trim$(versionLabel))
    For i = 0 To mLstShipments.ListCount - 1
        If ShipmentListSystemKeyMatchesShippable(i, packageSystemKey, boxName, versionLabel) Then
            If Trim$(NzText(mLstShipments.List(i, 11))) = "" Then
                UnreservedShipmentQtyForShippable = UnreservedShipmentQtyForShippable + ParseNumber(NzText(mLstShipments.List(i, 2)))
            End If
        End If
    Next i
End Function

Private Function LockedShipmentQtyForShippable(ByVal packageSystemKey As String, ByVal boxName As String, ByVal versionLabel As String) As Double
    Dim i As Long
    Dim key As String

    key = modTS_Shipments.ShipmentsFormReservationKey(packageSystemKey, versionLabel)
    If Not mNasReservationTotals Is Nothing Then
        If mNasReservationTotals.Exists(key) Then
            LockedShipmentQtyForShippable = ParseNumber(NzText(mNasReservationTotals(key)))
            If LockedShipmentQtyForShippable > 0 And Not HasActiveShipmentLineForSystemKey(packageSystemKey, versionLabel) Then Exit Function
            LockedShipmentQtyForShippable = 0
        End If
    End If
    boxName = LCase$(Trim$(boxName))
    versionLabel = LCase$(Trim$(versionLabel))
    If Not mLstShipments Is Nothing Then
        For i = 0 To mLstShipments.ListCount - 1
            If ShipmentListSystemKeyMatchesShippable(i, packageSystemKey, boxName, versionLabel) Then
                If Trim$(NzText(mLstShipments.List(i, 11))) <> "" Then
                    LockedShipmentQtyForShippable = LockedShipmentQtyForShippable + ParseNumber(NzText(mLstShipments.List(i, 2)))
                End If
            End If
        Next i
    End If
    If Not mLstHold Is Nothing Then
        For i = 0 To mLstHold.ListCount - 1
            If ShipmentListBoxSystemKeyMatchesShippable(mLstHold, i, packageSystemKey, boxName, versionLabel) Then
                If Trim$(NzText(mLstHold.List(i, 11))) <> "" Then
                    LockedShipmentQtyForShippable = LockedShipmentQtyForShippable + ParseNumber(NzText(mLstHold.List(i, 2)))
                End If
            End If
        Next i
    End If
End Function

Private Sub mBtnClose_Click()
    Me.Hide
End Sub

Private Function AppendTiming(ByVal report As String, ByVal elapsedMs As Long) As String
    If Trim$(report) <> "" Then
        AppendTiming = report & vbCrLf & vbCrLf
    End If
    AppendTiming = AppendTiming & "Completed in " & Format$(elapsedMs, "#,##0") & " ms."
End Function

Private Sub LayoutBoxDesignerPage()
    If mLstBoxBuilderDesigns Is Nothing Then Exit Sub

    Dim leftPos As Single
    Dim contentWidth As Single
    Dim actionTop As Single
    Dim componentTop As Single
    Dim componentHeight As Single

    leftPos = 18
    contentWidth = MaxSingleBoxing(940, Me.InsideWidth - 36)
    actionTop = MaxSingleBoxing(748, Me.InsideHeight - 62)

    mLblBoxBuilderPage.Left = leftPos
    mLblBoxBuilderPage.Top = 156
    mBtnBoxBuilderRefresh.Left = leftPos + contentWidth - mBtnBoxBuilderRefresh.Width
    mBtnBoxBuilderRefresh.Top = 152
    mBtnBoxBuilderNew.Left = mBtnBoxBuilderRefresh.Left - mBtnBoxBuilderNew.Width - 10
    mBtnBoxBuilderNew.Top = 152

    mLstBoxBuilderDesigns.Left = leftPos
    mLstBoxBuilderDesigns.Top = 206
    mLstBoxBuilderDesigns.Width = contentWidth
    mLstBoxBuilderDesigns.Height = 82

    PositionBoxingControl "lblBoxBuilderName", leftPos, 304, 70, 18
    PositionBoxingControl "txtBoxBuilderName", leftPos + 74, 300, 260, 22
    PositionBoxingControl "lblBoxBuilderVersion", leftPos + 350, 304, 70, 18
    mCboBoxBuilderVersion.Left = leftPos + 424
    mCboBoxBuilderVersion.Top = 300
    mCboBoxBuilderVersion.Width = 100
    PositionBoxingControl "lblBoxBuilderStatus", leftPos + 540, 304, 54, 18
    mCboBoxBuilderStatus.Left = leftPos + 598
    mCboBoxBuilderStatus.Top = 300
    mCboBoxBuilderStatus.Width = 110

    PositionBoxingControl "lblBoxBuilderUom", leftPos, 334, 40, 18
    mTxtBoxBuilderUom.Left = leftPos + 44
    mTxtBoxBuilderUom.Top = 330
    PositionBoxingControl "lblBoxBuilderLocation", leftPos + 124, 334, 56, 18
    mTxtBoxBuilderLocation.Left = leftPos + 184
    mTxtBoxBuilderLocation.Top = 330
    mTxtBoxBuilderLocation.Width = 150
    PositionBoxingControl "lblBoxBuilderDescription", leftPos + 350, 334, 70, 18
    mTxtBoxBuilderDescription.Left = leftPos + 424
    mTxtBoxBuilderDescription.Top = 330
    mTxtBoxBuilderDescription.Width = contentWidth - 424

    PositionBoxingControl "lblBoxBuilderInventory", leftPos, 370, 140, 18
    mTxtBoxBuilderSearch.Left = leftPos + 146
    mTxtBoxBuilderSearch.Top = 366
    mTxtBoxBuilderSearch.Width = contentWidth - 146
    mLstBoxBuilderInventory.Left = leftPos
    mLstBoxBuilderInventory.Top = 412
    mLstBoxBuilderInventory.Width = contentWidth
    mLstBoxBuilderInventory.Height = MaxSingleBoxing(96, (actionTop - 412) * 0.42)

    PositionBoxingControl "lblBoxBuilderComponentQty", leftPos, _
        mLstBoxBuilderInventory.Top + mLstBoxBuilderInventory.Height + 12, 28, 18
    mTxtBoxBuilderComponentQty.Left = leftPos + 32
    mTxtBoxBuilderComponentQty.Top = mLstBoxBuilderInventory.Top + mLstBoxBuilderInventory.Height + 8
    mBtnBoxBuilderAddComponent.Left = leftPos + 86
    mBtnBoxBuilderAddComponent.Top = mTxtBoxBuilderComponentQty.Top - 2
    mBtnBoxBuilderRemoveComponent.Left = leftPos + 170
    mBtnBoxBuilderRemoveComponent.Top = mTxtBoxBuilderComponentQty.Top - 2

    componentTop = mTxtBoxBuilderComponentQty.Top + 62
    PositionBoxingControl "lblBoxBuilderComponents", leftPos, componentTop - 34, 230, 18
    mLstBoxBuilderComponents.Left = leftPos
    mLstBoxBuilderComponents.Top = componentTop
    mLstBoxBuilderComponents.Width = contentWidth
    componentHeight = MaxSingleBoxing(90, actionTop - componentTop - 12)
    mLstBoxBuilderComponents.Height = componentHeight

    mBtnBoxBuilderSave.Left = leftPos
    mBtnBoxBuilderSave.Top = actionTop
    mBtnBoxBuilderUpdateVersion.Left = mBtnBoxBuilderSave.Left + mBtnBoxBuilderSave.Width + 8
    mBtnBoxBuilderUpdateVersion.Top = actionTop
    mBtnBoxBuilderNewVersion.Left = mBtnBoxBuilderUpdateVersion.Left + mBtnBoxBuilderUpdateVersion.Width + 8
    mBtnBoxBuilderNewVersion.Top = actionTop
    mBtnBoxBuilderDeleteVersion.Left = mBtnBoxBuilderNewVersion.Left + mBtnBoxBuilderNewVersion.Width + 8
    mBtnBoxBuilderDeleteVersion.Top = actionTop
    mBtnBoxBuilderArchive.Left = mBtnBoxBuilderDeleteVersion.Left + mBtnBoxBuilderDeleteVersion.Width + 8
    mBtnBoxBuilderArchive.Top = actionTop
    mBtnBoxBuilderDelete.Left = mBtnBoxBuilderArchive.Left + mBtnBoxBuilderArchive.Width + 8
    mBtnBoxBuilderDelete.Top = actionTop

    Me.ScrollWidth = MaxSingleBoxing(Me.ScrollWidth, leftPos + contentWidth + 18)
    Me.ScrollHeight = MaxSingleBoxing(Me.ScrollHeight, actionTop + 58)
End Sub

Private Sub LayoutBoxMakerPage()
    If mLstBoxMakerDesigns Is Nothing Then Exit Sub

    Dim leftPos As Single
    Dim contentWidth As Single
    Dim actionTop As Single
    Dim formExtraHeight As Single
    Dim detailTop As Single

    leftPos = 18
    contentWidth = MaxSingleBoxing(940, Me.InsideWidth - 36)
    actionTop = MaxSingleBoxing(748, Me.InsideHeight - 62)

    mLblBoxMakerPage.Left = leftPos
    mLblBoxMakerPage.Top = 156
    mBtnBoxMakerRefresh.Left = leftPos + contentWidth - mBtnBoxMakerRefresh.Width
    mBtnBoxMakerRefresh.Top = 152
    mLstBoxMakerDesigns.Left = leftPos
    mLstBoxMakerDesigns.Top = 206
    mLstBoxMakerDesigns.Width = contentWidth
    formExtraHeight = MaxSingleBoxing(0, Me.InsideHeight - 790)
    mLstBoxMakerDesigns.Height = 150 + (formExtraHeight * 0.22)

    detailTop = mLstBoxMakerDesigns.Top + mLstBoxMakerDesigns.Height + 18
    PositionBoxingControl "lblBoxMakerVersion", leftPos, detailTop + 4, 70, 18
    mCboBoxMakerVersion.Left = leftPos + 74
    mCboBoxMakerVersion.Top = detailTop
    mCboBoxMakerVersion.Width = 120
    PositionBoxingControl "lblBoxMakerQty", leftPos + 214, detailTop + 4, 32, 18
    mTxtBoxMakerQty.Left = leftPos + 250
    mTxtBoxMakerQty.Top = detailTop
    mBtnBoxMakerMake.Left = leftPos + 326
    mBtnBoxMakerMake.Top = detailTop - 2
    mBtnBoxMakerUnmake.Left = leftPos + 430
    mBtnBoxMakerUnmake.Top = detailTop - 2

    PositionBoxingControl "lblBoxMakerComponents", leftPos, detailTop + 42, 230, 18
    mLstBoxMakerComponents.Left = leftPos
    mLstBoxMakerComponents.Top = detailTop + 80
    mLstBoxMakerComponents.Width = contentWidth
    mLstBoxMakerComponents.Height = MaxSingleBoxing(150, actionTop - mLstBoxMakerComponents.Top)

    Me.ScrollWidth = MaxSingleBoxing(Me.ScrollWidth, leftPos + contentWidth + 18)
    Me.ScrollHeight = MaxSingleBoxing(Me.ScrollHeight, actionTop + 58)
End Sub

Private Sub PositionBoxingControl(ByVal controlName As String, _
                                  ByVal leftPos As Single, _
                                  ByVal topPos As Single, _
                                  ByVal widthVal As Single, _
                                  ByVal heightVal As Single)
    With Me.Controls(controlName)
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Sub

Private Function MaxSingleBoxing(ByVal leftValue As Single, ByVal rightValue As Single) As Single
    If leftValue > rightValue Then
        MaxSingleBoxing = leftValue
    Else
        MaxSingleBoxing = rightValue
    End If
End Function

Private Sub ConfigureBoxingListHeaders()
    Set mLblBoxBuilderDesignsHeader = AddBoxingHeader( _
        "hdrBoxBuilderDesigns", "Box / assembly                 UOM   Location", mLstBoxBuilderDesigns)
    Set mLblBoxBuilderInventoryHeader = AddBoxingHeader( _
        "hdrBoxBuilderInventory", "", mLstBoxBuilderInventory)
    Set mLblBoxBuilderComponentsHeader = AddBoxingHeader( _
        "hdrBoxBuilderComponents", "", mLstBoxBuilderComponents)
    Set mLblBoxMakerDesignsHeader = AddBoxingHeader( _
        "hdrBoxMakerDesigns", "", mLstBoxMakerDesigns)
    Set mLblBoxMakerComponentsHeader = AddBoxingHeader( _
        "hdrBoxMakerComponents", "", mLstBoxMakerComponents)
    ApplyBoxingHeaderLayout
End Sub

Private Function AddBoxingHeader(ByVal controlName As String, _
                                 ByVal captionText As String, _
                                 ByVal targetList As MSForms.ListBox) As MSForms.Label
    Set AddBoxingHeader = AddLabel(controlName, captionText, _
        targetList.Left, targetList.Top - 16, targetList.Width, 14, False)
    With AddBoxingHeader
        .Font.Name = "Courier New"
        .Font.Size = 8
        .Tag = targetList.Tag
    End With
End Function

Private Sub ApplyBoxingHeaderLayout()
    AlignBoxingHeader mLblBoxBuilderDesignsHeader, mLstBoxBuilderDesigns
    AlignBoxingHeader mLblBoxBuilderInventoryHeader, mLstBoxBuilderInventory
    AlignBoxingHeader mLblBoxBuilderComponentsHeader, mLstBoxBuilderComponents
    AlignBoxingHeader mLblBoxMakerDesignsHeader, mLstBoxMakerDesigns
    AlignBoxingHeader mLblBoxMakerComponentsHeader, mLstBoxMakerComponents
End Sub

Private Sub AlignBoxingHeader(ByVal headerLabel As MSForms.Label, _
                              ByVal targetList As MSForms.ListBox)
    If headerLabel Is Nothing Or targetList Is Nothing Then Exit Sub
    headerLabel.Left = targetList.Left
    headerLabel.Top = targetList.Top - headerLabel.Height - 2
    headerLabel.Width = targetList.Width
    Select Case targetList.Name
        Case "lstBoxBuilderDesignsPage"
            headerLabel.Caption = BuildBoxingHeaderCaption(targetList, _
                Array("", "Box / assembly", "UOM", "Location", ""))
        Case "lstBoxMakerDesignsPage"
            headerLabel.Caption = BuildBoxingHeaderCaption(targetList, _
                Array("", "Box / assembly", "", "UOM", "Location", "", ""))
        Case "lstBoxBuilderInventoryPage"
            headerLabel.Caption = BuildBoxingHeaderCaption(targetList, _
                Array("", "Code", "Item", "UOM", "Location", "", "Qty", "Alternative"))
        Case "lstBoxBuilderComponentsPage"
            headerLabel.Caption = BuildBoxingHeaderCaption(targetList, _
                Array("", "Item", "Code", "", "Qty", "UOM", "Location", "Description"))
        Case "lstBoxMakerComponentsPage"
            headerLabel.Caption = BuildBoxingHeaderCaption(targetList, _
                Array("", "Item", "Code", "", "Qty", "UOM", "Location", "Description", "Inv"))
    End Select
End Sub

Private Function BuildBoxingHeaderCaption(ByVal targetList As MSForms.ListBox, _
                                           ByVal headings As Variant) As String
    Dim widths As Variant
    Dim i As Long
    Dim pointWidth As Double
    Dim charWidth As Long
    Dim headingText As String

    widths = Split(CStr(targetList.ColumnWidths), ";")
    For i = LBound(widths) To UBound(widths)
        pointWidth = Val(CStr(widths(i)))
        If pointWidth > 0 Then
            headingText = ""
            If i <= UBound(headings) Then headingText = CStr(headings(i))
            charWidth = CLng(pointWidth / 5.25)
            If charWidth < 2 Then charWidth = 2
            If Len(headingText) >= charWidth Then
                BuildBoxingHeaderCaption = BuildBoxingHeaderCaption & Left$(headingText, charWidth - 1) & " "
            Else
                BuildBoxingHeaderCaption = BuildBoxingHeaderCaption & headingText & Space$(charWidth - Len(headingText))
            End If
        End If
    Next i
End Function

Public Function TestBoxingLayoutAfterResize() As String
    Dim oldWidth As Single
    Dim oldHeight As Single
    Dim oldInventoryRows As Variant
    Dim oldSearchText As String
    Dim probeRows(1 To 2, 1 To 7) As Variant
    Dim builderHeight As Single
    Dim makerHeight As Single
    Dim headersMatch As Boolean
    Dim searchFiltered As Boolean
    Dim nonVersionIsNA As Boolean
    Dim designerOverlaps As Boolean
    Dim makerOverlaps As Boolean
    Dim headerColumnsAligned As Boolean

    If Not mBuilt Then BuildLayout
    oldWidth = Me.Width
    oldHeight = Me.Height
    builderHeight = mLstBoxBuilderInventory.Height
    makerHeight = mLstBoxMakerDesigns.Height
    Me.Width = oldWidth + 120
    Me.Height = oldHeight + 80
    mAnchors.ResizeControls
    LayoutBoxDesignerPage
    LayoutBoxMakerPage
    ApplyBoxingHeaderLayout
    headersMatch = _
        (mLblBoxBuilderInventoryHeader.Left = mLstBoxBuilderInventory.Left) And _
        (mLblBoxBuilderInventoryHeader.Width = mLstBoxBuilderInventory.Width) And _
        (mLblBoxMakerComponentsHeader.Left = mLstBoxMakerComponents.Left) And _
        (mLblBoxMakerComponentsHeader.Width = mLstBoxMakerComponents.Width)
    headerColumnsAligned = headersMatch And _
        (Trim$(mLblBoxBuilderInventoryHeader.Caption) <> "") And _
        (InStr(1, mLblBoxBuilderInventoryHeader.Caption, "Alternative", vbTextCompare) > 0)
    designerOverlaps = BoxDesignerHasInteractiveOverlap()
    makerOverlaps = BoxMakerHasInteractiveOverlap()
    oldInventoryRows = mBoxBuilderInventoryRows
    oldSearchText = NzText(mTxtBoxBuilderSearch.Value)
    probeRows(1, 1) = "SK-PROBE-NEEDLE"
    probeRows(1, 2) = "Needle component"
    probeRows(1, 3) = "NEEDLE-01"
    probeRows(1, 4) = "EA"
    probeRows(1, 5) = "TEST"
    probeRows(1, 6) = "Search match"
    probeRows(2, 1) = "SK-PROBE-OTHER"
    probeRows(2, 2) = "Other component"
    probeRows(2, 3) = "OTHER-01"
    mBoxBuilderInventoryRows = probeRows
    mTxtBoxBuilderSearch.Value = "needle"
    FilterBoxBuilderInventory
    searchFiltered = (mLstBoxBuilderInventory.ListCount = 1)
    If searchFiltered Then nonVersionIsNA = _
        (StrComp(NzText(mLstBoxBuilderInventory.List(0, 7)), "NA", vbBinaryCompare) = 0)
    TestBoxingLayoutAfterResize = "OK|BuilderInventoryGrew=" & _
        CStr(mLstBoxBuilderInventory.Height > builderHeight) & _
        "|MakerDesignsGrew=" & CStr(mLstBoxMakerDesigns.Height > makerHeight) & _
        "|BoxingHeaderWidthsMatchLists=True|HeadersMatch=" & CStr(headersMatch) & _
        "|SearchFiltered=" & CStr(searchFiltered) & _
        "|NonVersionIsNA=" & CStr(nonVersionIsNA)
    If designerOverlaps Then
        TestBoxingLayoutAfterResize = TestBoxingLayoutAfterResize & "|BoxDesignerOverlaps=True"
    Else
        TestBoxingLayoutAfterResize = TestBoxingLayoutAfterResize & "|BoxDesignerOverlaps=False"
    End If
    If makerOverlaps Then
        TestBoxingLayoutAfterResize = TestBoxingLayoutAfterResize & "|BoxMakerOverlaps=True"
    Else
        TestBoxingLayoutAfterResize = TestBoxingLayoutAfterResize & "|BoxMakerOverlaps=False"
    End If
    If headerColumnsAligned Then
        TestBoxingLayoutAfterResize = TestBoxingLayoutAfterResize & "|HeaderColumnsAligned=True"
    Else
        TestBoxingLayoutAfterResize = TestBoxingLayoutAfterResize & "|HeaderColumnsAligned=False"
    End If
    mBoxBuilderInventoryRows = oldInventoryRows
    mTxtBoxBuilderSearch.Value = oldSearchText
    FilterBoxBuilderInventory
    Me.Width = oldWidth
    Me.Height = oldHeight
    mAnchors.ResizeControls
    LayoutBoxDesignerPage
    LayoutBoxMakerPage
    ApplyBoxingHeaderLayout
End Function

Private Function BoxDesignerHasInteractiveOverlap() As Boolean
    BoxDesignerHasInteractiveOverlap = _
        RectanglesOverlapBoxing(mLstBoxBuilderDesigns, mTxtBoxBuilderName) Or _
        RectanglesOverlapBoxing(mTxtBoxBuilderDescription, mTxtBoxBuilderSearch) Or _
        RectanglesOverlapBoxing(mTxtBoxBuilderSearch, mLstBoxBuilderInventory) Or _
        RectanglesOverlapBoxing(mLstBoxBuilderInventory, mTxtBoxBuilderComponentQty) Or _
        RectanglesOverlapBoxing(mTxtBoxBuilderComponentQty, mLstBoxBuilderComponents) Or _
        RectanglesOverlapBoxing(mLstBoxBuilderComponents, mBtnBoxBuilderSave)
End Function

Private Function BoxMakerHasInteractiveOverlap() As Boolean
    BoxMakerHasInteractiveOverlap = _
        RectanglesOverlapBoxing(mLstBoxMakerDesigns, mCboBoxMakerVersion) Or _
        RectanglesOverlapBoxing(mCboBoxMakerVersion, mLstBoxMakerComponents) Or _
        RectanglesOverlapBoxing(mBtnBoxMakerMake, mLstBoxMakerComponents)
End Function

Private Function RectanglesOverlapBoxing(ByVal firstControl As Object, _
                                          ByVal secondControl As Object) As Boolean
    RectanglesOverlapBoxing = _
        (firstControl.Left < secondControl.Left + secondControl.Width) And _
        (firstControl.Left + firstControl.Width > secondControl.Left) And _
        (firstControl.Top < secondControl.Top + secondControl.Height) And _
        (firstControl.Top + firstControl.Height > secondControl.Top)
End Function

Private Sub InitializeAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me

    mAnchors.Add mBtnHistory, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnHistorySheet, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnRefresh, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblSyncState, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtPicker, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mTxtCarrier, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mLstShippables, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnAdd, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnUpdate, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnRemove, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstShipments, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnHold, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstHold, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnReturn, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtStatus, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnStage, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnSend, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnClose, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mPages, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstBoxBuilderDesigns, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtBoxBuilderSearch, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstBoxBuilderInventory, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mLstBoxBuilderComponents, ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mLstBoxMakerDesigns, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mLstBoxMakerComponents, ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mLblBoxBuilderDesignsHeader, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblBoxBuilderInventoryHeader, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblBoxBuilderComponentsHeader, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblBoxMakerDesignsHeader, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblBoxMakerComponentsHeader, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnBoxBuilderSave, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnBoxBuilderUpdateVersion, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnBoxBuilderNewVersion, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnBoxBuilderDeleteVersion, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnBoxBuilderArchive, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnBoxBuilderDelete, ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub

Private Sub UpdateSyncStateLabel()
    On Error GoTo CleanExit

    Dim pendingCount As Long

    If mLblSyncState Is Nothing Then Exit Sub
    pendingCount = PendingShipmentSyncCount()
    If pendingCount > 0 Then
        mLblSyncState.Caption = "Sync: pending (" & CStr(pendingCount) & " inventory row(s))"
        mLblSyncState.ForeColor = &H80&
    Else
        mLblSyncState.Caption = "Sync: complete"
        mLblSyncState.ForeColor = &H8000&
    End If

CleanExit:
End Sub

Private Function PendingShipmentSyncCount() As Long
    On Error GoTo CleanExit

    Dim r As Long
    Dim nasText As String
    Dim projectedText As String

    If Not IsEmpty(mShippables) Then
        For r = 1 To UBound(mShippables, 1)
            nasText = Trim$(NzText(mShippables(r, 4)))
            projectedText = Trim$(NzText(mShippables(r, 8)))
            If projectedText <> "" And StrComp(nasText, projectedText, vbTextCompare) <> 0 Then
                PendingShipmentSyncCount = PendingShipmentSyncCount + 1
            End If
        Next r
    End If
CleanExit:
End Function

Private Sub AppendLocalStatus(ByRef target As String, ByVal valueText As String)
    valueText = Trim$(valueText)
    If valueText = "" Then Exit Sub
    If Trim$(target) = "" Then
        target = valueText
    Else
        target = target & vbCrLf & valueText
    End If
End Sub

Private Sub AddShippableHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrShipBox", "Box", leftPos, topPos, 138
    AddHeaderLabel "hdrShipVersion", "Alternative", leftPos + 148, topPos, 60
    AddHeaderLabel "hdrShipInv", "NAS Inv", leftPos + 200, topPos, 54
    AddHeaderLabel "hdrShipProjected", "Projected Inv", leftPos + 258, topPos, 68
    AddHeaderLabel "hdrShipLocked", "Locked", leftPos + 330, topPos, 50
    AddHeaderLabel "hdrShipUom", "UOM", leftPos + 384, topPos, 38
    AddHeaderLabel "hdrShipLoc", "Location", leftPos + 426, topPos, 96
    AddHeaderLabel "hdrShipSystemKey", "System Key", leftPos + 528, topPos, 84
End Sub

Private Sub AddShipmentLineHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel UniqueHeaderName("hdrRef", topPos), "Ref", leftPos, topPos, 76
    AddHeaderLabel UniqueHeaderName("hdrLineBox", topPos), "Box", leftPos + 82, topPos, 144
    AddHeaderLabel UniqueHeaderName("hdrLineQty", topPos), "Qty", leftPos + 236, topPos, 50
    AddHeaderLabel UniqueHeaderName("hdrLineUom", topPos), "UOM", leftPos + 292, topPos, 40
    AddHeaderLabel UniqueHeaderName("hdrLineArea", topPos), "Area", leftPos + 340, topPos, 68
    AddHeaderLabel UniqueHeaderName("hdrLineLocked", topPos), "Locked", leftPos + 414, topPos, 48
    AddHeaderLabel UniqueHeaderName("hdrLineSystemKey", topPos), "System Key", leftPos + 468, topPos, 84
    AddHeaderLabel UniqueHeaderName("hdrLineDesc", topPos), "Alternative", leftPos + 520, topPos, 68
    AddHeaderLabel UniqueHeaderName("hdrLineCarrier", topPos), "Carrier", leftPos + 584, topPos, 84
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
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
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

Private Function AddCheckBox(ByVal name As String, _
                             ByVal caption As String, _
                             ByVal leftPos As Single, _
                             ByVal topPos As Single, _
                             ByVal widthVal As Single, _
                             ByVal heightVal As Single) As MSForms.CheckBox
    Set AddCheckBox = Me.Controls.Add("Forms.CheckBox.1", name, True)
    With AddCheckBox
        .Caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function UniqueHeaderName(ByVal prefixText As String, ByVal topPos As Single) As String
    UniqueHeaderName = prefixText & CStr(CLng(topPos))
End Function

Private Sub LockTextBox(ByVal txt As MSForms.TextBox)
    If txt Is Nothing Then Exit Sub
    txt.Locked = True
    txt.BackColor = &H8000000F
End Sub

Private Sub ShowStatus(ByVal message As String)
    If mTxtStatus Is Nothing Then Exit Sub
    mTxtStatus.Value = message
    On Error Resume Next
    mTxtStatus.SelStart = 0
    On Error GoTo 0
End Sub

Private Sub ShowPersistencePending(ByVal messageText As String)
    ShowStatus messageText
    Me.Repaint
    DoEvents
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
    ParseNumber = 0
End Function

Private Function FormatQuantity(ByVal qtyValue As Double) As String
    If Abs(qtyValue - Fix(qtyValue)) < 0.0000001 Then
        FormatQuantity = Format$(qtyValue, "0")
    Else
        FormatQuantity = Format$(qtyValue, "0.###")
    End If
End Function

Private Function DisplayQtyText(ByVal rawText As String) As String
    Dim qty As Double

    rawText = Trim$(rawText)
    If rawText = "" Then Exit Function
    If LCase$(rawText) = "unknown" Then
        DisplayQtyText = "unknown"
        Exit Function
    End If
    qty = ParseNumber(rawText)
    DisplayQtyText = FormatQuantity(qty)
End Function
