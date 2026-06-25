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
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnAdd As MSForms.CommandButton
Private WithEvents mBtnUpdate As MSForms.CommandButton
Private WithEvents mBtnRemove As MSForms.CommandButton
Private WithEvents mBtnHold As MSForms.CommandButton
Private WithEvents mBtnReturn As MSForms.CommandButton
Private WithEvents mBtnStage As MSForms.CommandButton
Private WithEvents mBtnSend As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton

Private mTxtBox As MSForms.TextBox
Private mTxtVersion As MSForms.TextBox
Private mTxtUom As MSForms.TextBox
Private mTxtLocation As MSForms.TextBox
Private mTxtRow As MSForms.TextBox
Private mTxtCarrier As MSForms.ComboBox
Private mTxtStatus As MSForms.TextBox
Private mLstReadiness As MSForms.ListBox
Private mLblSyncState As MSForms.Label

Private mShippables As Variant
Private mNasReservationTotals As Object
Private mLoading As Boolean
Private mBuilt As Boolean
Private mAnchors As Object
Private mResizeInitialized As Boolean
Private mOperatorWorkbook As Workbook
Private mNextPollTime As Date
Private mAutoSyncArmed As Boolean
Private mLastShippablesLoadReport As String
Private mUseInjectedReservationTotalsForTest As Boolean
Private mTimerLog() As String
Private mTimerCount As Long
Private mTimerStart As Single

Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8
Private Const POLL_INTERVAL_SECONDS As Long = 45

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Private Sub UserForm_Activate()
    modTS_Shipments.RegisterShipmentsFormAutoSync Me
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
    CancelAutoSync
    modTS_Shipments.UnregisterShipmentsFormAutoSync Me
    Set mAnchors = Nothing
    Set mOperatorWorkbook = Nothing
End Sub

Public Sub InitializeFromShipping()
    On Error GoTo FailInit

    Dim previousPointer As Long
    Dim quietStarted As Boolean
    Dim startedAt As Single
    Dim elapsedMs As Long
    Dim operatorWb As Workbook
    Dim loadStep As String

    TimingStart
    TLap "InitializeFromShipping start"
    loadStep = "build layout"
    If Not mBuilt Then BuildLayout
    TLap "build layout"
    loadStep = "resolve operator workbook"
    If mOperatorWorkbook Is Nothing And IsUsableOperatorWorkbook(ActiveWorkbook) Then Set mOperatorWorkbook = ActiveWorkbook
    Set operatorWb = ResolveOperatorWorkbook()
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
    elapsedMs = ElapsedMilliseconds(startedAt)
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

Private Sub TimingStart()
    mTimerCount = 0
    Erase mTimerLog
    mTimerStart = Timer
End Sub

Private Sub TLap(ByVal label As String)
    Dim elapsedMs As Long

    If mTimerStart <= 0 Then mTimerStart = Timer
    elapsedMs = ElapsedMilliseconds(mTimerStart)
    mTimerCount = mTimerCount + 1
    ReDim Preserve mTimerLog(1 To mTimerCount)
    mTimerLog(mTimerCount) = Format$(elapsedMs, "00000") & " ms  " & label
End Sub

Private Function TimingSummary() As String
    Dim i As Long
    Dim lines As String

    If mTimerCount <= 0 Then Exit Function
    For i = 1 To mTimerCount
        If lines <> "" Then lines = lines & vbCrLf
        lines = lines & mTimerLog(i)
    Next i
    TimingSummary = lines
End Function

Public Sub SetOperatorWorkbook(ByVal wb As Workbook)
    If IsUsableOperatorWorkbook(wb) Then Set mOperatorWorkbook = wb
End Sub

Private Function ResolveOperatorWorkbook() As Workbook
    On Error Resume Next

    Dim nameCheck As String
    Dim wb As Workbook
    Dim candidateWb As Workbook
    Dim candidateCount As Long

    If Not mOperatorWorkbook Is Nothing Then
        nameCheck = mOperatorWorkbook.Name
        If Err.Number = 0 And Trim$(nameCheck) <> "" And IsUsableOperatorWorkbook(mOperatorWorkbook) Then
            Set ResolveOperatorWorkbook = mOperatorWorkbook
            Exit Function
        End If
        Err.Clear
        Set mOperatorWorkbook = Nothing
    End If

    If IsShipmentsOperatorWorkbook(ActiveWorkbook) Then
        Set mOperatorWorkbook = ActiveWorkbook
        Set ResolveOperatorWorkbook = mOperatorWorkbook
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            If IsShipmentsOperatorWorkbook(wb) Then
                Set mOperatorWorkbook = wb
                Set ResolveOperatorWorkbook = wb
                Exit Function
            End If
            candidateCount = candidateCount + 1
            Set candidateWb = wb
        End If
    Next wb
    If candidateCount = 1 Then
        Set mOperatorWorkbook = candidateWb
        Set ResolveOperatorWorkbook = candidateWb
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
    ScheduleAutoSync
End Sub

Public Sub AutoSyncIfPending()
    On Error GoTo CleanExit

    Dim operatorWb As Workbook
    Dim report As String
    Dim changedLoading As Boolean
    Dim syncCount As Long
    Dim nasBeforeRefresh As String
    Dim nasAfterRefresh As String

    If Not mAutoSyncArmed Then Exit Sub
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
    nasBeforeRefresh = FirstShippableNasText()

    Set operatorWb = ResolveOperatorWorkbook()
    If operatorWb Is Nothing Then
        ShowStatus "AutoSync: operator workbook not resolved."
        GoTo CleanExit
    End If

    If modTS_Shipments.ShipmentsFormAutoSyncRefresh(operatorWb, report) Then
        mLoading = True
        changedLoading = True
        LoadShippables
        nasAfterRefresh = FirstShippableNasText()
        LoadShipmentState
        RefreshProjectedShippableInventory
        mLoading = False
        changedLoading = False
        UpdateSyncStateLabel
        If PendingShipmentSyncCount() <= 0 Then mAutoSyncArmed = False
        ShowStatus "AutoSync: NAS was " & IIf(nasBeforeRefresh = "", "unknown", nasBeforeRefresh) & _
                   ", now " & IIf(nasAfterRefresh = "", "unknown", nasAfterRefresh) & ". " & report
    Else
        ShowStatus "AutoSync: refresh failed. " & report
        UpdateSyncStateLabel
    End If

CleanExit:
    If changedLoading Then mLoading = False
    If mAutoSyncArmed Then ScheduleAutoSync
End Sub

Private Function FirstShippableNasText() As String
    On Error GoTo CleanExit

    If IsEmpty(mShippables) Then Exit Function
    If UBound(mShippables, 1) < 1 Then Exit Function
    FirstShippableNasText = NzText(mShippables(1, 4))

CleanExit:
End Function

Private Sub BuildLayout()
    If mBuilt Then Exit Sub
    mBuilt = True

    Me.Caption = "Shipping Shipments"
    Me.Width = 860
    Me.Height = 675
    Me.ScrollBars = fmScrollBarsBoth
    Me.ScrollWidth = 850
    Me.ScrollHeight = 650

    AddLabel "lblTitle", "Shipments", 12, 10, 140, 20, True
    Set mBtnHistory = AddButton("btnHistory", "History", 708, 10, 58, 24)
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
    AddLabel "lblVersion", "Version", 270, 194, 52, 18, False
    AddLabel "lblQty", "Qty", 336, 194, 34, 18, False
    AddLabel "lblUom", "UOM", 410, 194, 40, 18, False
    AddLabel "lblLocation", "Location", 470, 194, 60, 18, False
    AddLabel "lblRow", "ROW", 620, 194, 40, 18, False
    AddLabel "lblCarrier", "Carrier", 12, 242, 54, 18, False

    Set mTxtRef = AddTextBox("txtRef", 12, 212, 82, 22)
    Set mTxtBox = AddTextBox("txtBox", 108, 212, 148, 22)
    Set mTxtVersion = AddTextBox("txtVersion", 270, 212, 52, 22)
    Set mTxtQty = AddTextBox("txtQty", 336, 212, 52, 22)
    Set mTxtUom = AddTextBox("txtUom", 410, 212, 44, 22)
    Set mTxtLocation = AddTextBox("txtLocation", 470, 212, 132, 22)
    Set mTxtRow = AddTextBox("txtRow", 620, 212, 52, 22)
    Set mTxtCarrier = AddComboBox("txtCarrier", 108, 238, 148, 22)
    Set mTxtDescription = AddTextBox("txtDescription", 12, 240, 1, 1)
    mTxtDescription.Visible = False
    LockTextBox mTxtBox
    LockTextBox mTxtVersion
    LockTextBox mTxtUom
    LockTextBox mTxtLocation
    LockTextBox mTxtRow
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

    InitializeAnchors
    LoadCarrierChoices
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
    RenderLineList mLstShipments, modTS_Shipments.ShipmentsFormLoadLines(False, wb)
    RenderLineList mLstHold, modTS_Shipments.ShipmentsFormLoadLines(True, wb)
    EvictOrphanedActiveOverlays
    UpdateSyncStateLabel
End Sub

Private Sub LoadShipmentLineState(Optional ByVal operatorWb As Workbook = Nothing)
    Dim wb As Workbook

    Set wb = operatorWb
    If wb Is Nothing Then Set wb = ResolveOperatorWorkbook()
    RenderLineList mLstShipments, modTS_Shipments.ShipmentsFormLoadLines(False, wb)
    RenderLineList mLstHold, modTS_Shipments.ShipmentsFormLoadLines(True, wb)
    EvictOrphanedActiveOverlays
    UpdateSyncStateLabel
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
        displayRows(idx, 4) = DisplayQtyText(CStr(LockedShipmentQtyForShippable(CLng(Val(NzText(mShippables(r, 1)))), NzText(mShippables(r, 2)), NzText(mShippables(r, 3)))))
        displayRows(idx, 5) = NzText(mShippables(r, 5))
        displayRows(idx, 6) = NzText(mShippables(r, 6))
        displayRows(idx, 7) = NzText(mShippables(r, 1))
        idx = idx + 1
NextRow:
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

Private Sub RenderReadiness(ByVal rowsData As Variant)
    On Error GoTo FailSoft

    Dim r As Long
    Dim displayRows As Variant

    mLstReadiness.Clear
    If IsEmpty(rowsData) Then Exit Sub
    ReDim displayRows(0 To UBound(rowsData, 1) - 1, 0 To 8)
    For r = 1 To UBound(rowsData, 1)
        displayRows(r - 1, 0) = NzText(rowsData(r, 1))
        displayRows(r - 1, 1) = NzText(rowsData(r, 2))
        displayRows(r - 1, 2) = FormatQuantity(ParseNumber(NzText(rowsData(r, 3))))
        displayRows(r - 1, 3) = FormatQuantity(ParseNumber(NzText(rowsData(r, 4))))
        displayRows(r - 1, 4) = FormatQuantity(ParseNumber(NzText(rowsData(r, 5))))
        displayRows(r - 1, 5) = NzText(rowsData(r, 6))
        displayRows(r - 1, 6) = NzText(rowsData(r, 7))
        displayRows(r - 1, 7) = NzText(rowsData(r, 8))
        displayRows(r - 1, 8) = NzText(rowsData(r, 9))
    Next r
    mLstReadiness.List = displayRows
    Exit Sub

FailSoft:
End Sub

Private Sub LoadSelectedShippable()
    If mLstShippables.ListIndex < 0 Then Exit Sub
    mTxtBox.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 0))
    mTxtVersion.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 1))
    mTxtUom.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 5))
    mTxtLocation.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 6))
    mTxtRow.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 7))
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
    mTxtRow.Value = NzText(lst.List(lst.ListIndex, 6))
    mTxtDescription.Value = NzText(lst.List(lst.ListIndex, 7))
    mTxtVersion.Value = NzText(lst.List(lst.ListIndex, 7))
    mTxtCarrier.Value = NzText(lst.List(lst.ListIndex, 8))
End Sub

Private Sub CommitCurrentLine(ByVal actionName As String)
    On Error GoTo FailSoft

    Dim report As String
    Dim rowIndex As Long
    Dim ok As Boolean
    Dim displayedAvailableQty As String
    Dim operatorWb As Workbook
    Dim startedAt As Single
    Dim elapsedMs As Long

    TimingStart
    TLap "CommitCurrentLine " & UCase$(Trim$(actionName)) & " start"
    startedAt = Timer
    rowIndex = SelectedShipmentTableRow()
    displayedAvailableQty = SelectedShippableProjectedInventoryText()
    Set operatorWb = ResolveOperatorWorkbook()
    TLap "CommitCurrentLine resolved selected row/operator"
    ok = modTS_Shipments.ShipmentsFormCommitLine("SHIP", _
                                                 actionName, _
                                                 rowIndex, _
                                                 NzText(mTxtRef.Value), _
                                                 NzText(mTxtBox.Value), _
                                                 ParseNumber(NzText(mTxtQty.Value)), _
                                                 CLng(Val(NzText(mTxtRow.Value))), _
                                                 NzText(mTxtUom.Value), _
                                                 NzText(mTxtLocation.Value), _
                                                 NzText(mTxtVersion.Value), _
                                                 NzText(mTxtCarrier.Value), _
                                                 report, _
                                                 displayedAvailableQty, _
                                                 mShippables, _
                                                 operatorWb)
    TLap "CommitCurrentLine backend call"
    elapsedMs = ElapsedMilliseconds(startedAt)
    report = AppendTiming(report, elapsedMs)
    If TimingSummary() <> "" Then report = report & vbCrLf & TimingSummary()
    RefreshAfterAction report, ok
    Exit Sub

FailSoft:
    ShowStatus "Shipment row action failed: " & Err.Description
End Sub

Private Function SelectedShippableProjectedInventoryText() As String
    Dim r As Long
    Dim rowValue As Long
    Dim boxName As String
    Dim versionLabel As String

    rowValue = CLng(Val(NzText(mTxtRow.Value)))
    boxName = Trim$(NzText(mTxtBox.Value))
    versionLabel = Trim$(NzText(mTxtVersion.Value))

    If mLstShippables Is Nothing Then Exit Function
    If mLstShippables.ListIndex >= 0 Then
        If CLng(Val(NzText(mLstShippables.List(mLstShippables.ListIndex, 7)))) = rowValue _
           And StrComp(Trim$(NzText(mLstShippables.List(mLstShippables.ListIndex, 0))), boxName, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mLstShippables.List(mLstShippables.ListIndex, 1))), versionLabel, vbTextCompare) = 0 Then
            SelectedShippableProjectedInventoryText = NzText(mLstShippables.List(mLstShippables.ListIndex, 3))
            Exit Function
        End If
    End If

    If IsEmpty(mShippables) Then Exit Function
    For r = 1 To UBound(mShippables, 1)
        If CLng(Val(NzText(mShippables(r, 1)))) = rowValue _
           And StrComp(Trim$(NzText(mShippables(r, 2))), boxName, vbTextCompare) = 0 _
           And StrComp(Trim$(NzText(mShippables(r, 3))), versionLabel, vbTextCompare) = 0 Then
            SelectedShippableProjectedInventoryText = NzText(mShippables(r, 8))
            Exit Function
        End If
    Next r
End Function

Private Function SelectedShipmentTableRow() As Long
    If mLstShipments Is Nothing Then Exit Function
    If mLstShipments.ListIndex < 0 Then Exit Function
    SelectedShipmentTableRow = CLng(Val(NzText(mLstShipments.List(mLstShipments.ListIndex, 10))))
End Function

Private Function SelectedHoldTableRow() As Long
    If mLstHold Is Nothing Then Exit Function
    If mLstHold.ListIndex < 0 Then Exit Function
    SelectedHoldTableRow = CLng(Val(NzText(mLstHold.List(mLstHold.ListIndex, 10))))
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
    If countRows = 0 And lst.ListIndex >= 0 Then
        tableRow = CLng(Val(NzText(lst.List(lst.ListIndex, 10))))
        If tableRow > 0 Then
            ReDim rowIndexes(1 To 1)
            rowIndexes(1) = tableRow
            countRows = 1
        End If
    End If
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
    TimingStart
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

    TimingStart
    TLap "Refresh click start"
    Set operatorWb = ResolveOperatorWorkbook()
    ok = modTS_Shipments.ShipmentsFormRefreshRuntimeInventoryForWorkbook(operatorWb, report)
    TLap "Refresh backend refresh"
    InitializeFromShipping
    TLap "Refresh reinitialize form"
    If Trim$(report) <> "" Then
        ShowStatus "Shipments form refreshed. " & report & vbCrLf & TimingSummary()
    Else
        ShowStatus "Shipments form refreshed." & vbCrLf & TimingSummary()
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

Private Sub mBtnAdd_Click()
    CommitCurrentLine "ADD"
End Sub

Private Sub mBtnUpdate_Click()
    CommitCurrentLine "UPDATE"
End Sub

Private Sub mBtnRemove_Click()
    CommitCurrentLine "DELETE"
End Sub

Private Sub mBtnHold_Click()
    MoveSelectedShipmentHold True
End Sub

Private Sub mBtnReturn_Click()
    MoveSelectedShipmentHold False
End Sub

Private Sub MoveSelectedShipmentHold(ByVal moveToHold As Boolean)
    On Error GoTo FailSoft

    Dim lst As MSForms.ListBox
    Dim report As String
    Dim ok As Boolean
    Dim selectedRows As Variant

    TimingStart
    TLap "Hold/Return click start"
    If moveToHold Then
        Set lst = mLstShipments
    Else
        Set lst = mLstHold
    End If
    selectedRows = SelectedListTableRows(lst)
    If IsEmpty(selectedRows) Then
        ShowStatus "Select a shipment row first."
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

Private Sub mBtnStage_Click()
    RunShippingAction True
End Sub

Private Sub mBtnSend_Click()
    RunShippingAction False
End Sub

Private Sub RunShippingAction(ByVal stageOnly As Boolean)
    On Error GoTo FailSoft

    Dim previousPointer As Long
    Dim quietStarted As Boolean
    Dim startedAt As Single
    Dim elapsedMs As Long
    Dim report As String
    Dim ok As Boolean
    Dim selectedRows As Variant

    TimingStart
    TLap IIf(stageOnly, "To Shipments", "Shipments Sent") & " click start"
    selectedRows = SelectedListTableRows(mLstShipments)
    If IsEmpty(selectedRows) Then
        ShowStatus "Select shipment row(s) first."
        Exit Sub
    End If
    previousPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass
    modUiQuiet.BeginQuietUi ResolveOperatorWorkbook()
    quietStarted = True
    startedAt = Timer
    If stageOnly Then
        ok = modTS_Shipments.ShipmentsFormRunToShipmentsRows(selectedRows, NzText(mTxtCarrier.Value), report)
    Else
        ok = modTS_Shipments.ShipmentsFormRunShipmentsSentRows(selectedRows, NzText(mTxtCarrier.Value), report)
    End If
    TLap IIf(stageOnly, "To Shipments", "Shipments Sent") & " backend call"
    elapsedMs = ElapsedMilliseconds(startedAt)
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
    Dim projectedQty As Double
    Dim packageRow As Long

    If IsEmpty(mShippables) Then Exit Sub
    If Not mUseInjectedReservationTotalsForTest Then Set mNasReservationTotals = modTS_Shipments.ShipmentsFormLoadNasReservationTotals()
    For r = 1 To UBound(mShippables, 1)
        packageRow = CLng(Val(NzText(mShippables(r, 1))))
        backendText = NzText(mShippables(r, 4))
        activeQty = ActiveShipmentQtyForShippable(packageRow, NzText(mShippables(r, 2)), NzText(mShippables(r, 3)))
        projectedQty = modTS_Shipments.ShipmentsProjectedDisplayQty(ParseNumber(backendText), activeQty)
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

Private Sub EvictOrphanedActiveOverlays()
    Dim r As Long
    Dim packageRow As Long
    Dim versionLabel As String

    If IsEmpty(mShippables) Then Exit Sub
    If mLstShipments Is Nothing Then Exit Sub
    For r = 1 To UBound(mShippables, 1)
        packageRow = CLng(Val(NzText(mShippables(r, 1))))
        versionLabel = NzText(mShippables(r, 3))
        If packageRow > 0 And Trim$(versionLabel) <> "" Then
            If Not HasActiveShipmentLineForRow(packageRow, versionLabel) Then
                modTS_Shipments.ClearActiveOverlayForRowVersion packageRow, versionLabel
            End If
        End If
    Next r
End Sub

Private Function HasActiveShipmentLineForRow(ByVal packageRow As Long, ByVal versionLabel As String) As Boolean
    Dim i As Long
    Dim rowVersion As String

    versionLabel = LCase$(Trim$(versionLabel))
    If Not mLstShipments Is Nothing Then
        For i = 0 To mLstShipments.ListCount - 1
            If CLng(Val(NzText(mLstShipments.List(i, 6)))) = packageRow Then
                rowVersion = LCase$(Trim$(NzText(mLstShipments.List(i, 7))))
                If rowVersion = versionLabel Then
                    HasActiveShipmentLineForRow = True
                    Exit Function
                End If
            End If
        Next i
    End If

    If Not mLstHold Is Nothing Then
        For i = 0 To mLstHold.ListCount - 1
            If CLng(Val(NzText(mLstHold.List(i, 6)))) = packageRow Then
                rowVersion = LCase$(Trim$(NzText(mLstHold.List(i, 7))))
                If rowVersion = versionLabel Then
                    If Trim$(NzText(mLstHold.List(i, 11))) <> "" Then
                        HasActiveShipmentLineForRow = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
End Function

Private Function ShipmentListRowMatchesShippable(ByVal listIndex As Long, _
                                                 ByVal packageRow As Long, _
                                                 ByVal boxName As String, _
                                                 ByVal versionLabel As String) As Boolean
    ShipmentListRowMatchesShippable = ShipmentListBoxRowMatchesShippable(mLstShipments, listIndex, packageRow, boxName, versionLabel)
End Function

Private Function ShipmentListBoxRowMatchesShippable(ByVal lineList As MSForms.ListBox, _
                                                    ByVal listIndex As Long, _
                                                    ByVal packageRow As Long, _
                                                    ByVal boxName As String, _
                                                    ByVal versionLabel As String) As Boolean
    Dim rowBox As String
    Dim rowVersion As String
    Dim rowPackage As Long

    If lineList Is Nothing Then Exit Function
    rowBox = LCase$(Trim$(NzText(lineList.List(listIndex, 1))))
    rowVersion = LCase$(Trim$(NzText(lineList.List(listIndex, 7))))
    rowPackage = CLng(Val(NzText(lineList.List(listIndex, 6))))
    If packageRow > 0 Then
        ShipmentListBoxRowMatchesShippable = (rowPackage = packageRow And rowVersion = versionLabel)
    Else
        ShipmentListBoxRowMatchesShippable = (rowBox = boxName And rowVersion = versionLabel)
    End If
End Function

Private Function ActiveShipmentQtyForShippable(ByVal packageRow As Long, ByVal boxName As String, ByVal versionLabel As String) As Double
    Dim i As Long

    boxName = LCase$(Trim$(boxName))
    versionLabel = LCase$(Trim$(versionLabel))
    If Not mLstShipments Is Nothing Then
        For i = 0 To mLstShipments.ListCount - 1
            If ShipmentListRowMatchesShippable(i, packageRow, boxName, versionLabel) Then
                ActiveShipmentQtyForShippable = ActiveShipmentQtyForShippable + ParseNumber(NzText(mLstShipments.List(i, 2)))
            End If
        Next i
    End If
    If Not mLstHold Is Nothing Then
        For i = 0 To mLstHold.ListCount - 1
            If ShipmentListBoxRowMatchesShippable(mLstHold, i, packageRow, boxName, versionLabel) Then
                If Trim$(NzText(mLstHold.List(i, 11))) <> "" Then
                    ActiveShipmentQtyForShippable = ActiveShipmentQtyForShippable + ParseNumber(NzText(mLstHold.List(i, 2)))
                End If
            End If
        Next i
    End If
End Function

Private Function UnreservedShipmentQtyForShippable(ByVal packageRow As Long, ByVal boxName As String, ByVal versionLabel As String) As Double
    Dim i As Long

    If mLstShipments Is Nothing Then Exit Function
    boxName = LCase$(Trim$(boxName))
    versionLabel = LCase$(Trim$(versionLabel))
    For i = 0 To mLstShipments.ListCount - 1
        If ShipmentListRowMatchesShippable(i, packageRow, boxName, versionLabel) Then
            If Trim$(NzText(mLstShipments.List(i, 11))) = "" Then
                UnreservedShipmentQtyForShippable = UnreservedShipmentQtyForShippable + ParseNumber(NzText(mLstShipments.List(i, 2)))
            End If
        End If
    Next i
End Function

Private Function LockedShipmentQtyForShippable(ByVal packageRow As Long, ByVal boxName As String, ByVal versionLabel As String) As Double
    Dim i As Long
    Dim key As String

    key = modTS_Shipments.ShipmentsFormReservationKey(packageRow, versionLabel)
    If Not mNasReservationTotals Is Nothing Then
        If mNasReservationTotals.Exists(key) Then
            LockedShipmentQtyForShippable = ParseNumber(NzText(mNasReservationTotals(key)))
            If LockedShipmentQtyForShippable > 0 And Not HasActiveShipmentLineForRow(packageRow, versionLabel) Then Exit Function
            LockedShipmentQtyForShippable = 0
        End If
    End If
    boxName = LCase$(Trim$(boxName))
    versionLabel = LCase$(Trim$(versionLabel))
    If Not mLstShipments Is Nothing Then
        For i = 0 To mLstShipments.ListCount - 1
            If ShipmentListRowMatchesShippable(i, packageRow, boxName, versionLabel) Then
                If Trim$(NzText(mLstShipments.List(i, 11))) <> "" Then
                    LockedShipmentQtyForShippable = LockedShipmentQtyForShippable + ParseNumber(NzText(mLstShipments.List(i, 2)))
                End If
            End If
        Next i
    End If
    If Not mLstHold Is Nothing Then
        For i = 0 To mLstHold.ListCount - 1
            If ShipmentListBoxRowMatchesShippable(mLstHold, i, packageRow, boxName, versionLabel) Then
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

Private Function ElapsedMilliseconds(ByVal startedAt As Single) As Long
    Dim deltaSeconds As Single

    deltaSeconds = Timer - startedAt
    If deltaSeconds < 0 Then deltaSeconds = deltaSeconds + 86400!
    ElapsedMilliseconds = CLng(deltaSeconds * 1000)
End Function

Private Sub InitializeAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me

    mAnchors.Add mBtnHistory, ANCHOR_TOP Or ANCHOR_RIGHT
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
    mAnchors.Add mTxtStatus, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnStage, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnSend, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnClose, ANCHOR_RIGHT Or ANCHOR_BOTTOM
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
    If Not mLstShipments Is Nothing Then PendingShipmentSyncCount = PendingShipmentSyncCount + mLstShipments.ListCount

CleanExit:
End Function

Private Sub AddShippableHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrShipBox", "Box", leftPos, topPos, 138
    AddHeaderLabel "hdrShipVersion", "Version", leftPos + 148, topPos, 48
    AddHeaderLabel "hdrShipInv", "NAS Inv", leftPos + 200, topPos, 54
    AddHeaderLabel "hdrShipProjected", "Projected Inv", leftPos + 258, topPos, 68
    AddHeaderLabel "hdrShipLocked", "Locked", leftPos + 330, topPos, 50
    AddHeaderLabel "hdrShipUom", "UOM", leftPos + 384, topPos, 38
    AddHeaderLabel "hdrShipLoc", "Location", leftPos + 426, topPos, 96
    AddHeaderLabel "hdrShipRow", "ROW", leftPos + 528, topPos, 42
End Sub

Private Sub AddShipmentLineHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel UniqueHeaderName("hdrRef", topPos), "Ref", leftPos, topPos, 76
    AddHeaderLabel UniqueHeaderName("hdrLineBox", topPos), "Box", leftPos + 82, topPos, 144
    AddHeaderLabel UniqueHeaderName("hdrLineQty", topPos), "Qty", leftPos + 236, topPos, 50
    AddHeaderLabel UniqueHeaderName("hdrLineUom", topPos), "UOM", leftPos + 292, topPos, 40
    AddHeaderLabel UniqueHeaderName("hdrLineArea", topPos), "Area", leftPos + 340, topPos, 68
    AddHeaderLabel UniqueHeaderName("hdrLineLocked", topPos), "Locked", leftPos + 414, topPos, 48
    AddHeaderLabel UniqueHeaderName("hdrLineRow", topPos), "ROW", leftPos + 468, topPos, 46
    AddHeaderLabel UniqueHeaderName("hdrLineDesc", topPos), "Version", leftPos + 520, topPos, 58
    AddHeaderLabel UniqueHeaderName("hdrLineCarrier", topPos), "Carrier", leftPos + 584, topPos, 84
End Sub

Private Sub AddReadinessHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrReadyType", "Type", leftPos, topPos, 60
    AddHeaderLabel "hdrReadyItem", "Item", leftPos + 66, topPos, 174
    AddHeaderLabel "hdrReadyReq", "Required", leftPos + 250, topPos, 58
    AddHeaderLabel "hdrReadyInv", "Current", leftPos + 316, topPos, 58
    AddHeaderLabel "hdrReadyStaged", "Staged", leftPos + 382, topPos, 58
    AddHeaderLabel "hdrReadyUom", "UOM", leftPos + 448, topPos, 38
    AddHeaderLabel "hdrReadyLoc", "Location", leftPos + 494, topPos, 90
    AddHeaderLabel "hdrReadyRow", "ROW", leftPos + 592, topPos, 44
    AddHeaderLabel "hdrReadyStatus", "Status", leftPos + 644, topPos, 76
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
