VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventoryViewer
   Caption         =   "Viewer"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInventoryViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private Const SETTINGS_APP As String = "invSys"
Private Const SETTINGS_SECTION_OPERATIONS As String = "Operations"
Private Const SETTINGS_EVENT_RANGE As String = "InventoryViewerEventRange"

Private WithEvents mTxtSearch As MSForms.TextBox
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mBtnSettings As MSForms.CommandButton
Private WithEvents mBtnEventsPrevious As MSForms.CommandButton
Private WithEvents mBtnEventsNext As MSForms.CommandButton
Private mLblEventPage As MSForms.Label
Private mEventGroups As cEventPageProjection
Private mEventsPage As Long
Private mCboEventsView As MSForms.ComboBox
Private mCboEventsFamily As MSForms.ComboBox
Private mCboEventsSource As MSForms.ComboBox
Private mCboEventsOutcome As MSForms.ComboBox
Private mFilterBindings As Collection
Private mChangingFilters As Boolean
Private mConfiguringFilters As Boolean
Private mAppliedEventRange As String
Private WithEvents mTabs As MSForms.TabStrip
Private WithEvents mBtnExportListBox As MSForms.CommandButton
Private mCboEventRange As MSForms.ComboBox
Private WithEvents mLstInventory As MSForms.ListBox
Private mLblTitle As MSForms.Label
Private mLblHeaders As MSForms.Label
Private mLblStatus As MSForms.Label
Private mLblEventRange As MSForms.Label
Private mLblEventRangeHelp As MSForms.Label
Private mLblExportListBox As MSForms.Label
Private mTxtExportListBox As MSForms.TextBox
Private mHeaderLabels As Collection
Private mLayout As cOperationsAnchorManager
Private mWarehouseId As String
Private mRows As Variant
Private mBuilt As Boolean
Private mResizeInitialized As Boolean
Private mGeneration As Long
Private mColumnCount As Long
Private mLoadStatus As String
Private mSettingsContext As String
Private mLoadedColumnCount As Long
Private mDetail As cEventDetailController
Private mVisibleIndexes As Collection
Private mRecording As cRecordingControls
Private mPathLibrary As frmActionPaths
Private WithEvents mBtnActionPaths As MSForms.CommandButton

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Private Sub UserForm_Activate()
    If Not mRecording Is Nothing Then mRecording.Render
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeInitialized = True
    End If
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
    ConfigureViewerHeaderGeometry
End Sub

Private Sub UserForm_Layout()
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
    ConfigureViewerHeaderGeometry
End Sub

Private Sub UserForm_Terminate()
    If Not mDetail Is Nothing Then mDetail.CloseDetail
    Set mDetail = Nothing
    modOperationsTrackingSettings.CloseSettings
    modInventoryViewer.UnregisterInventoryViewer Me
    Set mLayout = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim binding As cViewerFilterBinding
    ClosePathLibrary
    If Not mRecording Is Nothing Then mRecording.Disconnect
    Set mRecording = Nothing
    If mFilterBindings Is Nothing Then Exit Sub
    For Each binding In mFilterBindings: binding.Disconnect: Next binding
    Set mFilterBindings = Nothing
End Sub

Public Sub SetWarehouse(ByVal warehouseId As String)
    If mSettingsContext <> modActivity.CaptureContext() Then ClosePathLibrary
    If mSettingsContext <> modActivity.CaptureContext() Then ClearViewerContent
    mWarehouseId = Trim$(warehouseId)
    mSettingsContext = modActivity.CaptureContext()
    If Not mRecording Is Nothing Then mRecording.BindContext mSettingsContext
    Me.Caption = "Viewer - " & mWarehouseId
End Sub

Public Sub SetGeneration(ByVal generation As Long)
    mGeneration = generation
End Sub

Public Sub RefreshInventory()
    If Not mBuilt Then BuildLayout
    If Not mRecording Is Nothing Then mRecording.Render
    If Not ViewerContextValid() Then Exit Sub
    mColumnCount = 6
    LoadViewerPayload modInventoryViewerData.LoadCurrentInventoryViewerData(), "inventory level(s)"
End Sub

Public Sub RefreshEvents()
    Dim publishedPayload As String
    If Not mBuilt Then BuildLayout
    If Not mRecording Is Nothing Then mRecording.Render
    If Not ViewerContextValid() Then Exit Sub
    mColumnCount = 10
    publishedPayload = modInventoryViewer.LoadInventoryViewerEvents()
    If Not ViewerContextValid() Then Exit Sub
    LoadViewerPayload publishedPayload, "event(s)"
End Sub

Private Sub ClearViewerContent()
    mRows = Empty
    Set mEventGroups = Nothing
    RefreshEventFilterChoices
    mEventsPage = 0
    If Not mBtnEventsPrevious Is Nothing Then mBtnEventsPrevious.Enabled = False
    If Not mBtnEventsNext Is Nothing Then mBtnEventsNext.Enabled = False
    If Not mLblEventPage Is Nothing Then mLblEventPage.Caption = "No loaded Events."
    mLoadedColumnCount = 0
    If Not mDetail Is Nothing Then mDetail.Invalidate
    If Not mLstInventory Is Nothing Then mLstInventory.Clear
End Sub

Private Function ViewerContextValid() As Boolean
    ViewerContextValid = (mSettingsContext <> "" And mSettingsContext = modActivity.CaptureContext())
    If ViewerContextValid Then Exit Function
    ClosePathLibrary
    ClearViewerContent
    mLoadStatus = "Unavailable. The invSys session or warehouse changed. Reopen Viewer."
    mLblStatus.Caption = mLoadStatus
End Function

Private Sub EventRefreshUnavailable()
    If mLoadedColumnCount = 10 Then
        mLoadStatus = "Stale. Events refresh failed; displaying previously loaded data. Try Refresh after publication is restored."
        If Not mDetail Is Nothing Then mDetail.MarkStale
    Else
        ClearViewerContent
        mLoadStatus = "Unavailable. Published Events could not be loaded. Try Refresh after publication is restored."
    End If
    mLblStatus.Caption = mLoadStatus
End Sub

Private Sub LoadViewerPayload(ByVal payload As String, ByVal rowLabel As String)
    Dim lines As Variant
    Dim header As Variant
    Dim dataRows() As Variant
    Dim fields As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim dataIndex As Long
    Dim dataColumnCount As Long

    lines = Split(payload, vbCrLf)
    header = Split(CStr(lines(0)), vbTab)
    If UBound(header) < 1 Or StrComp(CStr(header(0)), "OK", vbTextCompare) <> 0 Then
        If mColumnCount = 10 Then
            EventRefreshUnavailable
            Exit Sub
        End If
        ClearViewerContent
        If UBound(header) >= 1 Then
            mLoadStatus = ViewerUnescape(CStr(header(1)))
        Else
            mLoadStatus = "Inventory snapshot could not be loaded."
        End If
        mLblStatus.Caption = mLoadStatus
        Exit Sub
    End If

    dataColumnCount = mColumnCount
    If mColumnCount = 10 Then dataColumnCount = 18
    If mColumnCount = 10 And UBound(header) = 11 Then
        If CStr(header(4)) = "EVENTS1" Then dataColumnCount = 19 + UBound(Split(CStr(header(9)), ","))
    End If
    If UBound(lines) >= 1 Then
        ReDim dataRows(1 To UBound(lines), 1 To dataColumnCount)
        For rowIndex = 1 To UBound(lines)
            If Trim$(CStr(lines(rowIndex))) <> "" Then
                fields = Split(CStr(lines(rowIndex)), vbTab)
                If UBound(fields) >= mColumnCount - 1 Then
                    dataIndex = dataIndex + 1
                    For columnIndex = 1 To dataColumnCount
                        If columnIndex <= UBound(fields) + 1 Then dataRows(dataIndex, columnIndex) = ViewerUnescape(CStr(fields(columnIndex - 1)))
                    Next columnIndex
                End If
            End If
        Next rowIndex
    End If
    If dataIndex = 0 Then
        mRows = Empty
    ElseIf dataIndex = UBound(dataRows, 1) Then
        mRows = dataRows
    Else
        mRows = TrimViewerRows(dataRows, dataIndex, dataColumnCount)
    End If
    If mColumnCount = 10 Then
        Set mEventGroups = New cEventPageProjection
        mEventGroups.LoadProjection mRows, header
        mAppliedEventRange = CStr(mCboEventRange.Value)
        RefreshEventFilterChoices
        If mDetail Is Nothing Then Set mDetail = New cEventDetailController
        mDetail.LoadProjection mRows, header, mSettingsContext, mEventGroups
    End If
    mLoadedColumnCount = mColumnCount
    mLoadStatus = CStr(dataIndex) & " " & rowLabel & ". Published data read at " & CStr(header(2)) & "."
    If mColumnCount = 10 And UBound(header) = 11 Then
        mLoadStatus = CStr(dataIndex) & " contributing line(s). Published " & Replace$(Replace$(CStr(header(2)), "T", " "), "Z", " UTC") & _
            "; Loaded " & Replace$(Replace$(CStr(header(5)), "T", " "), "Z", " UTC") & "."
    End If
    mLblStatus.Caption = mLoadStatus
    RenderRows Trim$(CStr(mTxtSearch.Value))
End Sub

Public Function TestReport() As String
    TestReport = "OK|Warehouse=" & mWarehouseId & _
        "|VisibleRows=" & CStr(mLstInventory.ListCount) & _
        "|Generation=" & CStr(mGeneration) & _
        "|Modeless=True|Status=" & mLblStatus.Caption
End Function

Public Function TestApplySearch(ByVal filterText As String) As String
    mTxtSearch.Value = filterText
    RenderRows filterText
    TestApplySearch = "OK|Filter=" & filterText & _
        "|VisibleRows=" & CStr(mLstInventory.ListCount) & _
        "|Generation=" & CStr(mGeneration)
End Function

Public Function TestEventsReport(Optional ByVal rangeText As String = "") As String
    Dim removeRows As Long
    Dim shipmentHeldRows As Long
    Dim productionInputRows As Long
    Dim productionOutputRows As Long
    Dim readableDates As Long
    Dim rowIndex As Long
    Dim firstDate As String
    Dim firstReference As String
    If Not mBuilt Then BuildLayout
    mTabs.Value = 1
    If Trim$(rangeText) <> "" Then mCboEventRange.Value = rangeText
    mBtnRefresh_Click
    For rowIndex = 0 To mLstInventory.ListCount - 1
        If StrComp(Trim$(CStr(mLstInventory.List(rowIndex, 1))), "Remove", vbTextCompare) = 0 Then
            removeRows = removeRows + 1
        End If
        If StrComp(Trim$(CStr(mLstInventory.List(rowIndex, 1))), "Shipment Held", vbTextCompare) = 0 Then
            shipmentHeldRows = shipmentHeldRows + 1
        End If
        If StrComp(Trim$(CStr(mLstInventory.List(rowIndex, 1))), "Production Input Consumed", vbTextCompare) = 0 Then
            productionInputRows = productionInputRows + 1
        End If
        If StrComp(Trim$(CStr(mLstInventory.List(rowIndex, 1))), "Production Output Created", vbTextCompare) = 0 Then
            productionOutputRows = productionOutputRows + 1
        End If
        If Trim$(CStr(mLstInventory.List(rowIndex, 0))) <> "" And _
           Not IsNumeric(Trim$(CStr(mLstInventory.List(rowIndex, 0)))) Then
            readableDates = readableDates + 1
        End If
    Next rowIndex
    If mLstInventory.ListCount > 0 Then
        firstDate = CStr(mLstInventory.List(0, 0))
        firstReference = CStr(mLstInventory.List(0, 2))
    End If
    TestEventsReport = "OK|Title=" & mLblTitle.Caption & _
        "|TabCount=" & CStr(mTabs.Tabs.Count) & _
        "|TabCaptions=" & mTabs.Tabs(0).Caption & "," & mTabs.Tabs(1).Caption & "," & mTabs.Tabs(2).Caption & _
        "|SelectedTab=" & mTabs.Tabs(mTabs.Value).Caption & _
        "|VisibleRows=" & CStr(mLstInventory.ListCount) & _
        "|ReadableDates=" & CStr(readableDates) & _
        "|FirstDate=" & firstDate & _
        "|FirstReference=" & firstReference & _
        "|RemoveRows=" & CStr(removeRows) & _
        "|ShipmentHeldRows=" & CStr(shipmentHeldRows) & _
        "|ProductionInputRows=" & CStr(productionInputRows) & _
        "|ProductionOutputRows=" & CStr(productionOutputRows) & _
        "|EventRange=" & CStr(mCboEventRange.Value) & _
        "|RangeControlVisible=" & CStr(mCboEventRange.Visible) & _
        "|Columns=" & CStr(mLstInventory.ColumnCount) & _
        "|EventHeaderAligned=" & CStr(ViewerEventHeadersAlignedForTest()) & _
        "|ReadOnly=True|Generation=" & CStr(mGeneration)
End Function

Public Function ExportViewerListBoxToTable(ByVal listBoxName As String, _
                                           ByRef report As String) As Boolean
    Dim headers As Variant

    If StrComp(Trim$(listBoxName), "lstInventory", vbTextCompare) <> 0 Then
        report = "Viewer declares lstInventory. Enter that ListBox name, or open another declared Operations list."
        Exit Function
    End If
    headers = ViewerVisibleHeaders()
    ExportViewerListBoxToTable = modListBoxTableExport.ExportVisibleListBoxToNewTable( _
        mLstInventory, headers, report)
End Function

Public Function TestListBoxTableAction(ByRef report As String) As Boolean
    If Not mBuilt Then BuildLayout
    mTabs.Value = 0
    ApplyViewerTab
    TestListBoxTableAction = ExportViewerListBoxToTable("lstInventory", report)
End Function

Private Function ViewerVisibleHeaders() As Variant
    If Not mTabs Is Nothing Then
        If mTabs.Value = 1 Then
            ViewerVisibleHeaders = Array("Date", "Event", "Reference", "Item", "Qty", "UOM", "Location", "Condition", "User", "Details")
            Exit Function
        End If
    End If
    ViewerVisibleHeaders = Array("Item Code", "Item", "UOM", "Quantity", "Location", "Condition")
End Function

Private Sub BuildLayout()
    Dim rememberedRange As String
    Dim numericRange As Double

    If mBuilt Then Exit Sub
    Me.Width = 860
    Me.Height = 535

    Set mTabs = Me.Controls.Add("Forms.TabStrip.1", "tabsInventoryViewer", True)
    With mTabs
        .Move 12, 8, 820, 24
        .Tabs(0).Caption = "Inventory"
        .Tabs(1).Caption = "Events"
        .Tabs.Add "tabListBoxTable", "ListBox->Table"
        .Value = 0
    End With
    Set mLblTitle = AddLabel("lblTitle", "Current inventory levels", 12, 40, 360, 22, True)
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 740, 38, 92, 28)
    Set mBtnSettings = AddButton("btnSettings", "Settings", 636, 38, 92, 28)
    Set mBtnActionPaths = AddButton("btnActionPaths", "Action Paths", 516, 38, 108, 28)
    AddLabel "lblSearch", "Search", 12, 78, 76, 18, True
    Set mTxtSearch = AddTextBox("txtSearch", 92, 74, 740, 24)
    Set mLblEventRange = AddLabel("lblEventRange", "Event range", 12, 110, 96, 18, True)
    Set mCboEventRange = Me.Controls.Add("Forms.ComboBox.1", "cboEventRange", True)
    On Error Resume Next
    rememberedRange = Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION_OPERATIONS, SETTINGS_EVENT_RANGE, "All"))
    On Error GoTo 0
    Select Case UCase$(rememberedRange)
        Case "ALL"
            rememberedRange = "All"
        Case "DAY"
            rememberedRange = "Day"
        Case "WEEK"
            rememberedRange = "Week"
        Case "MONTH"
            rememberedRange = "Month"
        Case Else
            If IsNumeric(rememberedRange) Then
                On Error Resume Next
                Err.Clear
                numericRange = CDbl(rememberedRange)
                If Err.Number <> 0 Then numericRange = 0
                On Error GoTo 0
                If numericRange > 0 And numericRange = Fix(numericRange) And numericRange <= 36500 Then
                    rememberedRange = CStr(CLng(numericRange))
                Else
                    rememberedRange = "All"
                End If
            Else
                rememberedRange = "All"
            End If
    End Select
    With mCboEventRange
        .Move 112, 104, 150, 24
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        .AddItem "All"
        .AddItem "Day"
        .AddItem "Week"
        .AddItem "Month"
        .Value = rememberedRange
    End With
    Set mLblEventRangeHelp = AddLabel("lblEventRangeHelp", _
        "Choose Day, Week, Month, or type a whole number of days; select Refresh to apply.", _
        276, 110, 556, 18, False)
    mAppliedEventRange = rememberedRange
    Set mCboEventsView = AddEventFilter("cboEventsView", "View", "Operator actions")
    mCboEventsView.AddItem "All published events": mCboEventsView.List(1, 1) = "all"
    Set mCboEventsFamily = AddEventFilter("cboEventsFamily", "Event family", "All families")
    Set mCboEventsSource = AddEventFilter("cboEventsSource", "Source", "All sources")
    Set mCboEventsOutcome = AddEventFilter("cboEventsOutcome", "Recorded outcome", "All outcomes")
    Set mLblExportListBox = AddLabel("lblExportListBox", "ListBox name", 12, 110, 96, 18, True)
    Set mTxtExportListBox = AddTextBox("txtExportListBox", 112, 104, 330, 24)
    Set mBtnExportListBox = AddButton("btnExportListBox", "Export ListBox to Table", 454, 104, 170, 24)
    Set mLblHeaders = AddLabel("lblHeaders", _
        "Item Code                         Item                                  UOM       Quantity       Location                  Condition", _
        12, 140, 820, 18, True)
    Set mLstInventory = AddListBox("lstInventory", 12, 162, 820, 252)
    Set mBtnEventsPrevious = AddButton("btnEventsPrevious", "Previous", 12, 424, 86, 26)
    Set mBtnEventsNext = AddButton("btnEventsNext", "Next", 108, 424, 86, 26)
    Set mLblEventPage = AddLabel("lblEventPage", "No loaded Events.", 208, 428, 624, 22, False)
    With mLstInventory
        .ColumnCount = 6
        .ColumnWidths = "135 pt;190 pt;52 pt;72 pt;120 pt;74 pt"
        .IntegralHeight = False
    End With
    Set mHeaderLabels = New Collection
    ConfigureViewerHeaderGeometry
    Set mLblStatus = AddLabel("lblStatus", "Select Refresh to load the current published snapshot.", 12, 466, 680, 32, False)
    Set mBtnClose = AddButton("btnClose", "Close", 740, 466, 92, 30)
    Set mRecording = New cRecordingControls
    mRecording.Initialize Me

    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 720, 430
    mLayout.RegisterControl mTabs, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLblTitle, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mBtnRefresh, OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mBtnSettings, OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mBtnActionPaths, OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mTxtSearch, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLblEventRange, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mCboEventRange, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mLblEventRangeHelp, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLblExportListBox, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mTxtExportListBox, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mBtnExportListBox, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mLblHeaders, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLblStatus, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mBtnClose, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mBtnEventsPrevious, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mBtnEventsNext, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mLblEventPage, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mBuilt = True
    ApplyViewerTab
End Sub

Private Sub mTabs_Change()
    ApplyViewerTab
End Sub

Private Sub mBtnSettings_Click()
    Dim report As String
    If Not modOperationsTrackingSettings.OpenSettings(mSettingsContext, report) Then mLblStatus.Caption = report
End Sub

Private Sub mBtnActionPaths_Click()
    If Not ViewerContextValid() Then Exit Sub
    If mPathLibrary Is Nothing Then Set mPathLibrary = New frmActionPaths
    mPathLibrary.BindContext mSettingsContext
    If Not mPathLibrary.Visible Then mPathLibrary.Show vbModeless
End Sub

Private Sub ClosePathLibrary()
    If Not mPathLibrary Is Nothing Then Unload mPathLibrary
    Set mPathLibrary = Nothing
End Sub

Private Sub ApplyViewerTab()
    If mTabs Is Nothing Then Exit Sub
    mBtnEventsPrevious.Visible = (mTabs.Value = 1)
    mBtnEventsNext.Visible = (mTabs.Value = 1)
    mLblEventPage.Visible = (mTabs.Value = 1)
    mTxtSearch.Value = vbNullString
    mLblHeaders.Visible = False
    mLblExportListBox.Visible = False
    mTxtExportListBox.Visible = False
    mBtnExportListBox.Visible = False
    If mTabs.Value = 1 Then
        mTxtSearch.Visible = True
        mBtnRefresh.Visible = True
        mLstInventory.Visible = True
        mLblEventRange.Visible = True
        mCboEventRange.Visible = True
        mLblEventRangeHelp.Visible = True
        mLblTitle.Caption = "Inventory and shipping events"
        mLstInventory.ColumnCount = 10
        mLstInventory.ColumnWidths = "105 pt;82 pt;92 pt;130 pt;52 pt;46 pt;82 pt;72 pt;72 pt;190 pt"
        RefreshEvents
    ElseIf mTabs.Value = 2 Then
        mTxtSearch.Visible = False
        mBtnRefresh.Visible = False
        mLstInventory.Visible = False
        mLblEventRange.Visible = False
        mCboEventRange.Visible = False
        mLblEventRangeHelp.Visible = False
        mLblTitle.Caption = "ListBox->Table"
        mLblExportListBox.Visible = True
        mTxtExportListBox.Visible = True
        mBtnExportListBox.Visible = True
        mLblStatus.Caption = "Enter a declared open ListBox name, then export its displayed columns to a new worksheet table."
    Else
        mTxtSearch.Visible = True
        mBtnRefresh.Visible = True
        mLstInventory.Visible = True
        mLblEventRange.Visible = False
        mCboEventRange.Visible = False
        mLblEventRangeHelp.Visible = False
        mLblTitle.Caption = "Current inventory levels"
        mLstInventory.ColumnCount = 6
        mLstInventory.ColumnWidths = "135 pt;190 pt;52 pt;72 pt;120 pt;74 pt"
        RefreshInventory
    End If
    ConfigureViewerHeaderGeometry
End Sub

Private Sub mLstInventory_Click()
    If mColumnCount <> 10 Or mLstInventory.ListIndex < 0 Then Exit Sub
    If Not ViewerContextValid() Then Exit Sub
    If mDetail Is Nothing Or mVisibleIndexes Is Nothing Then Exit Sub
    If mLstInventory.ListIndex >= mVisibleIndexes.Count Then Exit Sub
    mDetail.ShowLine CLng(mVisibleIndexes(mLstInventory.ListIndex + 1))
End Sub

Private Sub mTxtSearch_Change()
    RenderRows Trim$(CStr(mTxtSearch.Value))
End Sub

Public Sub ApplyEventFilters()
    If Not mBuilt Or mChangingFilters Then Exit Sub
    If mTabs.Value <> 1 Then Exit Sub
    If Not ViewerContextValid() Then Exit Sub
    RenderRows Trim$(CStr(mTxtSearch.Value))
End Sub

Private Function AddEventFilter(ByVal name As String, ByVal caption As String, ByVal allCaption As String) As MSForms.ComboBox
    Dim binding As cViewerFilterBinding
    AddLabel "lbl" & Mid$(name, 4), caption, 12, 136, 190, 18, True
    Set AddEventFilter = Me.Controls.Add("Forms.ComboBox.1", name, True)
    With AddEventFilter
        .Move 12, 154, 190, 24
        .Style = fmStyleDropDownList
        .ColumnCount = 2: .BoundColumn = 2: .TextColumn = 1
        .ColumnWidths = "190 pt;0 pt"
        .AddItem allCaption: .List(0, 1) = "": .ListIndex = 0
    End With
    If mFilterBindings Is Nothing Then Set mFilterBindings = New Collection
    Set binding = New cViewerFilterBinding
    binding.Bind AddEventFilter, Me
    mFilterBindings.Add binding
End Function

Private Sub RefreshEventFilterChoices()
    If mCboEventsFamily Is Nothing Then Exit Sub
    mChangingFilters = True
    FillEventFacet mCboEventsFamily, "EventFamily", "All families"
    FillEventFacet mCboEventsSource, "Source", "All sources"
    FillEventFacet mCboEventsOutcome, "Outcome", "All outcomes"
    mChangingFilters = False
End Sub

Private Sub FillEventFacet(ByVal control As MSForms.ComboBox, ByVal field As String, ByVal allCaption As String)
    Dim selected As String, value As Variant, caption As String
    If Not IsNull(control.Value) Then selected = CStr(control.Value)
    control.Clear: control.AddItem allCaption: control.List(0, 1) = "": control.ListIndex = 0
    If mEventGroups Is Nothing Then Exit Sub
    For Each value In mEventGroups.FacetChoices(field)
        caption = CStr(value)
        If field = "Source" Then
            Select Case caption
                Case "Activity": caption = "User activity"
                Case "ShippingBOM": caption = "Box designs"
                Case "ShippingHolds": caption = "Held shipments"
            End Select
        ElseIf field = "Outcome" Then
            caption = StrConv(Replace$(caption, "_", " "), vbProperCase)
        End If
        control.AddItem caption: control.List(control.ListCount - 1, 1) = CStr(value)
        If CStr(value) = selected Then control.ListIndex = control.ListCount - 1
    Next value
End Sub

Private Sub ConfigureEventFilterGeometry()
    Dim names As Variant, index As Long, control As Object, label As Object
    Dim width As Single, top As Single, height As Single, visible As Boolean
    If mConfiguringFilters Or mCboEventsView Is Nothing Or mBtnEventsPrevious Is Nothing Then Exit Sub
    On Error GoTo Done
    mConfiguringFilters = True
    visible = (mTabs.Value = 1)
    If Not mBtnActionPaths Is Nothing Then mBtnActionPaths.Visible = visible
    top = 162: If visible Then top = 254
    If Not mRecording Is Nothing Then mRecording.Arrange visible, mTabs.Width
    height = mBtnEventsPrevious.Top - top - 10
    If Not visible And Not mBtnClose Is Nothing Then height = mBtnClose.Top - top - 12
    If height < 20 Then height = 20
    mLstInventory.Move 12, top, mTabs.Width, height
    width = (mLstInventory.Width - 36) / 4
    names = Array("EventsView", "EventsFamily", "EventsSource", "EventsOutcome")
    For index = 0 To 3
        Set control = Me.Controls("cbo" & CStr(names(index)))
        Set label = Me.Controls("lbl" & CStr(names(index)))
        control.Visible = visible: label.Visible = visible
        control.Move 12 + index * (width + 12), 154, width, 24
        control.ColumnWidths = CStr(width - 16) & " pt;0 pt"
        label.Move control.Left, 136, width, 18
    Next index
Done:
    mConfiguringFilters = False
End Sub

Private Sub mBtnEventsPrevious_Click()
    If Not ViewerContextValid() Or Not mBtnEventsPrevious.Enabled Then Exit Sub
    mEventsPage = mEventsPage - 1
    RenderRows Trim$(CStr(mTxtSearch.Value)), True
End Sub

Private Sub mBtnEventsNext_Click()
    If Not ViewerContextValid() Or Not mBtnEventsNext.Enabled Then Exit Sub
    mEventsPage = mEventsPage + 1
    RenderRows Trim$(CStr(mTxtSearch.Value)), True
End Sub

Private Sub mBtnRefresh_Click()
    If mTabs.Value = 1 Then
        RefreshEvents
    Else
        RefreshInventory
    End If
End Sub

Private Sub mBtnExportListBox_Click()
    Dim report As String
    If Not ViewerContextValid() Then Exit Sub
    If modInventoryViewer.ExportDeclaredListBoxToTable(Trim$(mTxtExportListBox.Text), report) Then
        mLblStatus.Caption = report
    Else
        mLblStatus.Caption = report
    End If
End Sub

Private Sub ConfigureViewerHeaderGeometry()
    Dim captions As Variant
    Dim widths As Variant
    Dim idx As Long
    Dim widthValue As Single
    Dim leftValue As Single
    Dim header As MSForms.Label

    If mLstInventory Is Nothing Then Exit Sub
    ConfigureEventFilterGeometry
    If mHeaderLabels Is Nothing Then Set mHeaderLabels = New Collection
    If Not mTabs Is Nothing Then
        If mTabs.Value = 1 Then
            captions = Array("Date", "Event", "Reference", "Item", "Qty", "UOM", "Location", "Condition", "User", "Details")
        Else
            captions = Array("Item Code", "Item", "UOM", "Quantity", "Location", "Condition")
        End If
    Else
        captions = Array("Item Code", "Item", "UOM", "Quantity", "Location", "Condition")
    End If
    widths = Split(mLstInventory.ColumnWidths, ";")
    leftValue = mLstInventory.Left
    For idx = LBound(captions) To UBound(captions)
        If idx + 1 > mHeaderLabels.Count Then
            Set header = AddLabel("hdrViewerColumn" & CStr(idx + 1), "", leftValue, _
                mLstInventory.Top - 20, 20, 18, True)
            header.Font.Size = 8
            mHeaderLabels.Add header
        Else
            Set header = mHeaderLabels(idx + 1)
        End If
        widthValue = CSng(Val(Replace$(Trim$(CStr(widths(idx))), "pt", "")))
        header.Caption = CStr(captions(idx))
        header.Move leftValue, mLstInventory.Top - 20, widthValue, 18
        header.Visible = (mTabs Is Nothing Or mTabs.Value <> 2)
        leftValue = leftValue + widthValue
    Next idx
    For idx = UBound(captions) + 2 To mHeaderLabels.Count
        mHeaderLabels(idx).Visible = False
    Next idx
End Sub

Private Function ViewerEventHeadersAlignedForTest() As Boolean
    Dim widths As Variant
    Dim idx As Long
    Dim leftValue As Single

    If mTabs Is Nothing Or mTabs.Value <> 1 Then Exit Function
    If mHeaderLabels Is Nothing Then Exit Function
    widths = Split(mLstInventory.ColumnWidths, ";")
    leftValue = mLstInventory.Left
    For idx = 0 To 9
        If idx + 1 > mHeaderLabels.Count Then Exit Function
        If Abs(mHeaderLabels(idx + 1).Left - leftValue) > 0.5 Then Exit Function
        leftValue = leftValue + CSng(Val(Replace$(Trim$(CStr(widths(idx))), "pt", "")))
    Next idx
    ViewerEventHeadersAlignedForTest = True
End Function

Private Sub mBtnClose_Click()
    Unload Me
End Sub

Private Sub RenderRows(ByVal filterText As String, Optional ByVal keepEventPage As Boolean = False)
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim matches As Boolean
    Dim rangeText As String
    Dim eventDays As Long
    Dim eventCutoff As Date
    Dim hasEventDateFilter As Boolean
    Dim numericRange As Double
    Dim storedRange As String

    If Not ViewerContextValid() Then Exit Sub
    Set mVisibleIndexes = New Collection
    mLstInventory.Clear
    mLblStatus.Caption = mLoadStatus
    If mTabs.Value = 1 Then
        mBtnEventsPrevious.Enabled = False: mBtnEventsNext.Enabled = False
        mLblEventPage.Caption = "No matching Events."
        rangeText = UCase$(Trim$(mAppliedEventRange))
        Select Case rangeText
            Case "", "ALL"
                storedRange = "All"
            Case "DAY"
                eventDays = 1
                storedRange = "Day"
            Case "WEEK"
                eventDays = 7
                storedRange = "Week"
            Case "MONTH"
                eventDays = 30
                storedRange = "Month"
            Case Else
                If IsNumeric(rangeText) Then
                    On Error Resume Next
                    Err.Clear
                    numericRange = CDbl(rangeText)
                    If Err.Number <> 0 Then numericRange = 0
                    On Error GoTo 0
                    If numericRange > 0 And numericRange = Fix(numericRange) And numericRange <= 36500 Then
                        eventDays = CLng(numericRange)
                        storedRange = CStr(eventDays)
                    End If
                End If
                If eventDays = 0 Then
                    mLblStatus.Caption = "Enter All, Day, Week, Month, or a whole number from 1 to 36500, then select Refresh."
                    Exit Sub
                End If
        End Select
        On Error Resume Next
        SaveSetting SETTINGS_APP, SETTINGS_SECTION_OPERATIONS, SETTINGS_EVENT_RANGE, storedRange
        On Error GoTo 0
        If eventDays > 0 Then
            hasEventDateFilter = True
            eventCutoff = DateAdd("d", -eventDays, Now)
        End If
        RenderEventPage filterText, hasEventDateFilter, eventCutoff, keepEventPage
        Exit Sub
    End If
    If IsEmpty(mRows) Then Exit Sub
    filterText = LCase$(Trim$(filterText))
    For rowIndex = LBound(mRows, 1) To UBound(mRows, 1)
        matches = (filterText = "")
        If Not matches Then
            For columnIndex = 1 To UBound(mRows, 2)
                If InStr(1, LCase$(CStr(mRows(rowIndex, columnIndex))), filterText, vbTextCompare) > 0 Then
                    matches = True
                    Exit For
                End If
            Next columnIndex
        End If
        If matches Then
            mVisibleIndexes.Add rowIndex
            mLstInventory.AddItem CStr(mRows(rowIndex, 1))
            For columnIndex = 2 To mColumnCount
                mLstInventory.List(mLstInventory.ListCount - 1, columnIndex - 1) = CStr(mRows(rowIndex, columnIndex))
            Next columnIndex
        End If
    Next rowIndex
End Sub

Private Sub RenderEventPage(ByVal filterText As String, ByVal restrictDates As Boolean, _
                            ByVal cutoff As Date, ByVal keepPage As Boolean)
    Dim pages As Long, first As Long, last As Long, index As Long, column As Long, utcText As String, utcNow As Date
    If mEventGroups Is Nothing Then Exit Sub
    If restrictDates Then
        utcText = Replace$(Left$(modTrainingWire.UtcTimestamp(), 19), "T", " ")
        If Not IsDate(utcText) Then
            mLblStatus.Caption = "Verified clock unavailable. Use All dates or try again."
            Exit Sub
        End If
        utcNow = CDate(utcText)
    End If
    mEventGroups.Match filterText, restrictDates, cutoff, Now, utcNow, _
        CStr(mCboEventsFamily.Value), CStr(mCboEventsSource.Value), CStr(mCboEventsOutcome.Value), (CStr(mCboEventsView.Value) = "all")
    pages = (mEventGroups.MatchingCount + 99) \ 100
    If Not keepPage Or mEventsPage < 1 Then mEventsPage = 1
    If mEventsPage > pages Then mEventsPage = pages
    mBtnEventsPrevious.Enabled = (mEventsPage > 1)
    mBtnEventsNext.Enabled = (mEventsPage < pages)
    mLblEventPage.Caption = "Page " & CStr(mEventsPage) & " of " & CStr(pages) & ". " & _
        CStr(mEventGroups.MatchingCount) & " matching / " & CStr(mEventGroups.AvailableCount) & " available groups."
    If pages = 0 Then Exit Sub
    first = (mEventsPage - 1) * 100 + 1
    last = first + 99
    If last > mEventGroups.MatchingCount Then last = mEventGroups.MatchingCount
    For index = first To last
        mVisibleIndexes.Add mEventGroups.SourceIndex(index)
        mLstInventory.AddItem mEventGroups.SummaryValue(index, 1)
        For column = 2 To mColumnCount
            mLstInventory.List(mLstInventory.ListCount - 1, column - 1) = mEventGroups.SummaryValue(index, column)
        Next column
    Next index
End Sub

Private Function TrimViewerRows(ByVal sourceRows As Variant, ByVal rowCount As Long, ByVal columnCount As Long) As Variant
    Dim resultRows() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    ReDim resultRows(1 To rowCount, 1 To columnCount)
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            resultRows(rowIndex, columnIndex) = sourceRows(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex
    TrimViewerRows = resultRows
End Function

Private Function ViewerUnescape(ByVal valueIn As String) As String
    Dim position As Long, current As String, decoded As String
    position = 1
    Do While position <= Len(valueIn)
        current = Mid$(valueIn, position, 1)
        If current = "\" And position < Len(valueIn) Then
            ' Decode once: an escaped backslash must not start another escape.
            Select Case Mid$(valueIn, position + 1, 1)
                Case "n": current = vbLf
                Case "r": current = vbCr
                Case "t": current = vbTab
                Case "\": current = "\"
                Case Else: current = current & Mid$(valueIn, position + 1, 1)
            End Select
            position = position + 1
        End If
        decoded = decoded & current
        position = position + 1
    Loop
    ViewerUnescape = decoded
End Function

Private Function AddLabel(ByVal controlName As String, _
                          ByVal captionText As String, _
                          ByVal leftValue As Double, _
                          ByVal topValue As Double, _
                          ByVal widthValue As Double, _
                          ByVal heightValue As Double, _
                          ByVal boldValue As Boolean) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    With AddLabel
        .Caption = captionText
        .Move leftValue, topValue, widthValue, heightValue
        .Font.Bold = boldValue
    End With
End Function

Private Function AddTextBox(ByVal controlName As String, _
                            ByVal leftValue As Double, _
                            ByVal topValue As Double, _
                            ByVal widthValue As Double, _
                            ByVal heightValue As Double) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    AddTextBox.Move leftValue, topValue, widthValue, heightValue
End Function

Private Function AddListBox(ByVal controlName As String, _
                            ByVal leftValue As Double, _
                            ByVal topValue As Double, _
                            ByVal widthValue As Double, _
                            ByVal heightValue As Double) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", controlName, True)
    AddListBox.Move leftValue, topValue, widthValue, heightValue
End Function

Private Function AddButton(ByVal controlName As String, _
                           ByVal captionText As String, _
                           ByVal leftValue As Double, _
                           ByVal topValue As Double, _
                           ByVal widthValue As Double, _
                           ByVal heightValue As Double) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    With AddButton
        .Caption = captionText
        .Move leftValue, topValue, widthValue, heightValue
    End With
End Function
