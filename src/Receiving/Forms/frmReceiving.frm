VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceiving
   Caption         =   "Receiving"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtRef As MSForms.TextBox
Private WithEvents mTxtReceiptId As MSForms.TextBox
Private WithEvents mTxtSearch As MSForms.TextBox
Private WithEvents mTxtItemSearch As MSForms.TextBox
Private WithEvents mTxtQty As MSForms.TextBox
Private WithEvents mTxtReceiveLocation As MSForms.TextBox
Private WithEvents mTxtLotNumber As MSForms.TextBox
Private WithEvents mTxtReturnReason As MSForms.TextBox
Private WithEvents mCboCondition As MSForms.ComboBox
Private WithEvents mCboDisposition As MSForms.ComboBox
Private WithEvents mTabs As MSForms.TabStrip
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnAdd As MSForms.CommandButton
Private WithEvents mBtnConfirm As MSForms.CommandButton
Private WithEvents mBtnClear As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mLstInventory As MSForms.ListBox
Private WithEvents mLstReceiveItems As MSForms.ListBox
Private WithEvents mLstStaged As MSForms.ListBox
Private WithEvents mLstAggregate As MSForms.ListBox

Private mLblRef As MSForms.Label
Private mLblReceiptId As MSForms.Label
Private mLblSearch As MSForms.Label
Private mLblItemSearch As MSForms.Label
Private mLblReceiveItemsTitle As MSForms.Label
Private mLblReceiveItemsHeader As MSForms.Label
Private mLblQty As MSForms.Label
Private mLblReceiveLocation As MSForms.Label
Private mLblLotNumber As MSForms.Label
Private mLblCondition As MSForms.Label
Private mLblReturnReason As MSForms.Label
Private mLblDisposition As MSForms.Label
Private mLblInventoryTitle As MSForms.Label
Private mLblInventoryHeader As MSForms.Label
Private mLblStagedTitle As MSForms.Label
Private mLblStagedHeader As MSForms.Label
Private mLblAggregateTitle As MSForms.Label
Private mLblAggregateHeader As MSForms.Label
Private mLblAggregateReferences As MSForms.Label
Private mTxtAggregateReferences As MSForms.TextBox
Private mLblPurchasingStub As MSForms.Label
Private mTxtStatus As MSForms.TextBox
Private mOperatorWorkbook As Workbook, mActivityContext As String
Private mHistoryRows As Variant
Private mItemRows As Variant
' Hidden representative System_Key values aligned one-for-one with result rows.
Private mReceiveItemSystemKeys As Collection
Private mBuilt As Boolean
Private mLoading As Boolean
Private mNavigationInputs As Collection
Private mResizeInitialized As Boolean
Private mResizing As Boolean
Private mLastConfirmQuietActive As Boolean

Private Const RECEIVING_BASE_WIDTH As Double = 1020
Private Const RECEIVING_BASE_HEIGHT As Double = 900

Private Sub UserForm_Initialize()
    mActivityContext = modActivity.CaptureContext()
    BuildLayout
End Sub

Private Sub UserForm_Activate()
    If Not mResizeInitialized Then
        modReceivingFormWindow.EnableReceivingResizable Me, True, True
        mResizeInitialized = True
    End If
    ResizeReceivingLayout
End Sub

Private Sub UserForm_Resize()
    ResizeReceivingLayout
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then _
        modReceivingActivityAction.CloseForm Me, ResolveOperatorWorkbook(), mActivityContext, True
    On Error Resume Next
    modReceivingNavigation.Detach mNavigationInputs
    modTS_Received.NotifyReceivingLauncherFormTerminating Me
    On Error GoTo 0
End Sub

Private Sub UserForm_Terminate()
    Set mOperatorWorkbook = Nothing
End Sub

Public Sub SetOperatorWorkbook(ByVal wb As Workbook)
    If Not wb Is Nothing Then
        If Not wb.IsAddin Then Set mOperatorWorkbook = wb
    End If
End Sub

Public Sub InitializeFromReceiving()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    If Not mBuilt Then BuildLayout
    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then
        ShowStatus "Open a Receiving operator workbook before using the Receiving form."
        Exit Sub
    End If

    mLoading = True
    modTS_Received.InitializeReceivingUiForWorkbook wb
    If Trim$(CStr(mTxtReceiptId.Value)) = "" Then mTxtReceiptId.Value = DefaultReceiptId()
    mTxtRef.Value = ""
    RefreshAllViews
    mLoading = False
    ShowStatus "Receiving form loaded for " & wb.Name & "."
    Exit Sub

ErrHandler:
    mLoading = False
    ShowStatus "Receiving form load failed: " & Err.Description
End Sub

Public Function TestSearchInventoryCount(ByVal values As Variant, ByVal filterText As String) As Long
    If Not mBuilt Then BuildLayout
    mHistoryRows = values
    FillInventoryList FilterInventoryRows(filterText)
    TestSearchInventoryCount = mLstInventory.ListCount
End Function

Public Function TestReceiptIdSeparatedFromReference() As Long
    If Not mBuilt Then BuildLayout
    mTxtReceiptId.Value = DefaultReceiptId()
    mTxtRef.Value = ""
    If Left$(CStr(mTxtReceiptId.Value), 4) = "RCV-" _
       And Trim$(CStr(mTxtRef.Value)) = "" Then
        TestReceiptIdSeparatedFromReference = 1
    End If
End Function

Public Function TestInitializeForWorkbook(ByVal operatorWb As Workbook) As String
    If operatorWb Is Nothing Then Exit Function
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    TestInitializeForWorkbook = _
        "OK|BoundWorkbook=" & mOperatorWorkbook.Name & _
        "|Caption=" & Me.Caption
End Function

Public Function TestRefreshInventoryActionForWorkbook(ByVal operatorWb As Workbook, _
                                                       Optional ByVal filterText As String = "") As String
    On Error GoTo Failed

    If operatorWb Is Nothing Then
        TestRefreshInventoryActionForWorkbook = "FAIL|Operator workbook is required."
        Exit Function
    End If
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    InitializeFromReceiving
    mTxtSearch.Value = filterText
    mBtnRefresh_Click
    TestRefreshInventoryActionForWorkbook = _
        "OK|VisibleRows=" & CStr(CountManagedItemChoices(filterText)) & _
        "|HistoryRows=" & CStr(mLstInventory.ListCount) & _
        "|AggregateRows=" & CStr(mLstAggregate.ListCount) & _
        "|Status=" & CStr(mTxtStatus.Text)
    Exit Function

Failed:
    TestRefreshInventoryActionForWorkbook = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function CountManagedItemChoices(ByVal filterText As String) As Long
    Dim r As Long
    Dim haystack As String
    filterText = LCase$(Trim$(filterText))
    If IsEmpty(mItemRows) Or Not IsArray(mItemRows) Then Exit Function
    For r = 1 To UBound(mItemRows, 1)
        haystack = LCase$(NzText(mItemRows(r, 2)) & " " & NzText(mItemRows(r, 3)))
        If filterText = "" Or InStr(1, haystack, filterText, vbTextCompare) > 0 Then
            CountManagedItemChoices = CountManagedItemChoices + 1
        End If
    Next r
End Function

Public Function TestRunConfirmWritesActionForWorkbook(ByVal operatorWb As Workbook, _
                                                       Optional ByVal activatedWb As Workbook = Nothing) As String
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    mBtnConfirm_Click
    TestRunConfirmWritesActionForWorkbook = _
        "Succeeded=" & CStr(modTS_Received.LastConfirmWritesSucceeded()) & _
        "; Status=" & modTS_Received.LastConfirmWritesStatus() & _
        "; BoundWorkbook=" & mOperatorWorkbook.Name & _
        "; QuietDuring=" & CStr(mLastConfirmQuietActive) & _
        "; QuietRestored=" & CStr(Not modUiQuiet.QuietUiIsActive())
End Function

Public Function TestPurchasingTabContract(ByVal operatorWb As Workbook) As String
    If operatorWb Is Nothing Then Exit Function
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    mTabs.Value = 2
    ApplyReceivingTab
    TestPurchasingTabContract = _
        "OK|Selected=" & mTabs.SelectedItem.Caption & _
        "|StubVisible=" & CStr(mLblPurchasingStub.Visible) & _
        "|EnabledPurchasingActions=0|Writes=0|Events=0"
End Function

Private Sub BuildLayout()
    If mBuilt Then Exit Sub

    Me.Caption = "Receiving"
    Me.Width = RECEIVING_BASE_WIDTH
    Me.Height = RECEIVING_BASE_HEIGHT
    Me.ScrollBars = fmScrollBarsBoth
    Me.KeepScrollBarsVisible = fmScrollBarsNone
    Me.ScrollWidth = RECEIVING_BASE_WIDTH - 20
    Me.ScrollHeight = RECEIVING_BASE_HEIGHT - 35

    Set mTabs = Me.Controls.Add("Forms.TabStrip.1", "tabsReceiving", True)
    With mTabs
        .Left = 18
        .Top = 10
        .Width = 964
        .Height = 26
        .Tabs.Clear
        .Tabs.Add "tabReceiving", "Receiving"
        .Tabs.Add "tabReturns", "Returns"
        .Tabs.Add "tabPurchasing", "Purchasing"
        .Value = 0
    End With

    Set mLblReceiptId = AddLabel("lblReceiptId", "Receipt ID", 18, 48, 70, 18, True)
    Set mTxtReceiptId = AddTextBox("txtReceiptId", 90, 46, 150, 22)
    mTxtReceiptId.Locked = True
    mTxtReceiptId.BackColor = &HEFEFEF
    Set mLblRef = AddLabel("lblRef", "PO/BOL Ref", 258, 48, 78, 18, True)
    Set mTxtRef = AddTextBox("txtRef", 338, 46, 150, 22)
    Set mLblQty = AddLabel("lblQty", "Qty", 720, 48, 34, 18, True)
    Set mTxtQty = AddTextBox("txtQty", 756, 46, 80, 22)
    mTxtQty.Value = "1"

    Set mLblItemSearch = AddLabel("lblItemSearch", "Receive item search", 18, 80, 120, 18, True)
    Set mTxtItemSearch = AddTextBox("txtItemSearch", 142, 78, 560, 22)
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 778, 78, 96, 28)
    Set mBtnAdd = AddButton("btnAdd", "Add Selected", 884, 78, 98, 28)
    Set mLblReceiveItemsTitle = AddLabel("lblReceiveItemsTitle", "Receive Item Results", 18, 110, 180, 18, True)
    Set mLblReceiveItemsHeader = AddLabel("lblReceiveItemsHeader", "", 18, 132, 964, 16, False)
    Set mLstReceiveItems = AddListBox("lstReceiveItems", 18, 150, 964, 116, 10, _
        "86 pt;180 pt;46 pt;62 pt;82 pt;96 pt;54 pt;70 pt;180 pt;100 pt")

    Set mLblReceiveLocation = AddLabel("lblReceiveLocation", "Receive location *", 18, 280, 112, 18, True)
    Set mTxtReceiveLocation = AddTextBox("txtReceiveLocation", 134, 276, 170, 22)
    Set mLblLotNumber = AddLabel("lblLotNumber", "Lot number (optional)", 324, 280, 130, 18, False)
    Set mTxtLotNumber = AddTextBox("txtLotNumber", 458, 276, 180, 22)
    Set mLblCondition = AddLabel("lblCondition", "Condition *", 658, 280, 76, 18, True)
    Set mCboCondition = AddComboBox("cboCondition", 738, 276, 120, 22)
    mCboCondition.AddItem "GOOD"
    mCboCondition.AddItem "BAD"
    mCboCondition.AddItem "DAMAGED"
    mCboCondition.AddItem "EXPIRED"
    mCboCondition.AddItem "REJECTED"
    mCboCondition.ListIndex = 0
    Set mLblReturnReason = AddLabel("lblReturnReason", "Return reason *", 18, 316, 94, 18, True)
    Set mTxtReturnReason = AddTextBox("txtReturnReason", 116, 312, 522, 22)
    Set mLblDisposition = AddLabel("lblDisposition", "Disposition *", 658, 316, 76, 18, True)
    Set mCboDisposition = AddComboBox("cboDisposition", 738, 312, 120, 22)
    mCboDisposition.AddItem "RETURN"
    mCboDisposition.AddItem "DUMP"
    mCboDisposition.ListIndex = 0
    mLblReturnReason.Visible = False
    mTxtReturnReason.Visible = False
    mLblDisposition.Visible = False
    mCboDisposition.Visible = False

    Set mLblSearch = AddLabel("lblSearch", "Search history", 18, 314, 90, 18, True)
    Set mTxtSearch = AddTextBox("txtSearch", 110, 312, 300, 22)
    Set mLblInventoryTitle = AddLabel("lblInventoryTitle", "Receiving Entries History", 18, 344, 220, 18, True)
    Set mLblInventoryHeader = AddLabel("lblInventoryHeader", "", 18, 366, 930, 16, False)
    Set mLstInventory = AddListBox("lstInventory", 18, 384, 964, 132, 10, _
        "90 pt;60 pt;90 pt;150 pt;52 pt;42 pt;80 pt;70 pt;62 pt;160 pt")

    Set mLblStagedTitle = AddLabel("lblStagedTitle", "Received Tally", 18, 532, 150, 18, True)
    Set mLblStagedHeader = AddLabel("lblStagedHeader", "", 18, 554, 520, 16, False)
    Set mLstStaged = AddListBox("lstStaged", 18, 572, 520, 190, 10, _
        "90 pt;58 pt;140 pt;48 pt;42 pt;72 pt;62 pt;70 pt;62 pt;120 pt")

    Set mLblAggregateTitle = AddLabel("lblAggregateTitle", "Aggregate Received", 560, 532, 160, 18, True)
    Set mLblAggregateHeader = AddLabel("lblAggregateHeader", "", 560, 554, 420, 16, False)
    Set mLstAggregate = AddListBox("lstAggregate", 560, 572, 422, 190, 10, _
        "72 pt;54 pt;76 pt;130 pt;42 pt;50 pt;65 pt;60 pt;62 pt;120 pt")
    Set mLblAggregateReferences = AddLabel("lblAggregateReferences", "Selected references", 560, 704, 130, 16, True)
    Set mTxtAggregateReferences = AddTextBox("txtAggregateReferences", 560, 722, 422, 40)
    mTxtAggregateReferences.Locked = True
    mTxtAggregateReferences.MultiLine = True
    mTxtAggregateReferences.WordWrap = True
    mTxtAggregateReferences.ScrollBars = fmScrollBarsVertical
    mTxtAggregateReferences.BackColor = &HFFFFFF

    Set mBtnConfirm = AddButton("btnConfirm", "Confirm Writes", 18, 780, 110, 30)
    Set mBtnClear = AddButton("btnClear", "Clear", 136, 780, 72, 30)
    Set mBtnClose = AddButton("btnClose", "Close", 892, 780, 90, 30)

    Set mLblPurchasingStub = AddLabel("lblPurchasingStub", _
        "Purchasing is not yet operational. This tab is reserved for future work and contains no purchasing write actions.", _
        36, 76, 900, 56, True)
    mLblPurchasingStub.WordWrap = True
    mLblPurchasingStub.Visible = False

    Set mTxtStatus = AddTextBox("txtStatus", 18, 830, 964, 34)
    With mTxtStatus
        .Locked = True
        .MultiLine = True
        .EnterKeyBehavior = False
        .ScrollBars = fmScrollBarsVertical
        .BackColor = &HFFFFFF
    End With

    mBuilt = True
    Set mNavigationInputs = modReceivingNavigation.CreateInputs(Me)
    ResizeReceivingLayout
    ApplyReceivingTab
End Sub

Private Function AddLabel(ByVal name As String, ByVal captionText As String, ByVal leftPos As Double, _
                          ByVal topPos As Double, ByVal widthVal As Double, ByVal heightVal As Double, _
                          ByVal boldText As Boolean) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", name, True)
    With lbl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = boldText
    End With
    Set AddLabel = lbl
End Function

Private Function AddTextBox(ByVal name As String, ByVal leftPos As Double, ByVal topPos As Double, _
                            ByVal widthVal As Double, ByVal heightVal As Double) As MSForms.TextBox
    Dim txt As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", name, True)
    With txt
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
    Set AddTextBox = txt
End Function

Private Function AddComboBox(ByVal name As String, ByVal leftPos As Double, ByVal topPos As Double, _
                             ByVal widthVal As Double, ByVal heightVal As Double) As MSForms.ComboBox
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls.Add("Forms.ComboBox.1", name, True)
    With cbo
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Style = fmStyleDropDownList
    End With
    Set AddComboBox = cbo
End Function

Private Function AddButton(ByVal name As String, ByVal captionText As String, ByVal leftPos As Double, _
                           ByVal topPos As Double, ByVal widthVal As Double, ByVal heightVal As Double) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", name, True)
    With btn
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
    Set AddButton = btn
End Function

Private Function AddListBox(ByVal name As String, ByVal leftPos As Double, ByVal topPos As Double, _
                            ByVal widthVal As Double, ByVal heightVal As Double, ByVal colCount As Long, _
                            ByVal widths As String) As MSForms.ListBox
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls.Add("Forms.ListBox.1", name, True)
    With lst
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .ColumnCount = colCount
        .ColumnWidths = widths
        .IntegralHeight = False
    End With
    Set AddListBox = lst
End Function

Private Sub ResizeReceivingLayout()
    If mResizing Or Not mBuilt Then Exit Sub
    mResizing = True
    On Error GoTo Done

    Dim layoutW As Double
    Dim layoutH As Double
    Dim margin As Double
    Dim bottomTop As Double
    Dim bottomH As Double
    Dim leftW As Double
    Dim rightW As Double
    Dim buttonTop As Double
    Dim statusTop As Double
    Dim extraHeight As Double
    Dim locationTop As Double
    Dim historyTop As Double
    Dim aggregateListH As Double

    margin = 18
    layoutW = MaxDoubleReceiving(RECEIVING_BASE_WIDTH - 20, Me.InsideWidth)
    layoutH = MaxDoubleReceiving(RECEIVING_BASE_HEIGHT - 35, Me.InsideHeight)
    Me.ScrollWidth = layoutW
    Me.ScrollHeight = layoutH
    mTabs.Width = layoutW - (margin * 2)
    mLblPurchasingStub.Width = layoutW - 72

    mLblQty.Left = layoutW - 282
    mTxtQty.Left = mLblQty.Left + 36
    mBtnRefresh.Left = layoutW - 224
    mBtnAdd.Left = layoutW - 116
    mTxtItemSearch.Width = MaxDoubleReceiving(260, mBtnRefresh.Left - mTxtItemSearch.Left - 12)

    extraHeight = MaxDoubleReceiving(0, layoutH - (RECEIVING_BASE_HEIGHT - 35))
    mLblReceiveItemsHeader.Width = layoutW - (margin * 2)
    mLstReceiveItems.Width = layoutW - (margin * 2)
    mLstReceiveItems.Height = 116 + (extraHeight * 0.22)

    locationTop = mLstReceiveItems.Top + mLstReceiveItems.Height + 14
    mLblReceiveLocation.Top = locationTop + 4
    mTxtReceiveLocation.Top = locationTop
    mLblLotNumber.Top = locationTop + 4
    mTxtLotNumber.Top = locationTop
    mLblCondition.Top = locationTop + 4
    mCboCondition.Top = locationTop
    mLblReturnReason.Top = locationTop + 40
    mTxtReturnReason.Top = locationTop + 36
    mLblDisposition.Top = locationTop + 40
    mCboDisposition.Top = locationTop + 36

    If mTabs.Value = 1 Then
        historyTop = locationTop + 72
    Else
        historyTop = locationTop + 36
    End If
    mLblSearch.Top = historyTop + 4
    mTxtSearch.Top = historyTop
    mLblInventoryTitle.Top = historyTop + 32
    mLblInventoryHeader.Top = historyTop + 54
    mLstInventory.Top = historyTop + 72
    mLblInventoryHeader.Width = layoutW - (margin * 2)
    mLstInventory.Width = layoutW - (margin * 2)
    mLstInventory.Height = 132 + (extraHeight * 0.22)

    bottomTop = mLstInventory.Top + mLstInventory.Height + 56
    bottomH = MaxDoubleReceiving(150, layoutH - bottomTop - 128)
    leftW = MaxDoubleReceiving(410, (layoutW - (margin * 3)) * 0.54)
    rightW = layoutW - leftW - (margin * 3)

    mLblStagedTitle.Top = bottomTop - 40
    mLblStagedHeader.Top = bottomTop - 18
    mLblStagedHeader.Width = leftW
    mLstStaged.Top = bottomTop
    mLstStaged.Width = leftW
    mLstStaged.Height = bottomH

    mLblAggregateTitle.Left = margin + leftW + margin
    mLblAggregateTitle.Top = bottomTop - 40
    mLblAggregateHeader.Left = mLblAggregateTitle.Left
    mLblAggregateHeader.Top = bottomTop - 18
    mLblAggregateHeader.Width = rightW
    mLstAggregate.Left = mLblAggregateTitle.Left
    mLstAggregate.Top = bottomTop
    mLstAggregate.Width = rightW
    aggregateListH = MaxDoubleReceiving(86, bottomH - 60)
    mLstAggregate.Height = aggregateListH
    mLblAggregateReferences.Left = mLstAggregate.Left
    mLblAggregateReferences.Top = mLstAggregate.Top + mLstAggregate.Height + 4
    mTxtAggregateReferences.Left = mLstAggregate.Left
    mTxtAggregateReferences.Top = mLblAggregateReferences.Top + 16
    mTxtAggregateReferences.Width = rightW
    mTxtAggregateReferences.Height = MaxDoubleReceiving(34, bottomH - aggregateListH - 20)

    buttonTop = layoutH - 72
    statusTop = layoutH - 38
    mBtnConfirm.Top = buttonTop
    mBtnClear.Top = buttonTop
    mBtnClose.Left = layoutW - mBtnClose.Width - margin
    mBtnClose.Top = buttonTop

    mTxtStatus.Top = statusTop
    mTxtStatus.Width = layoutW - (margin * 2)
    ApplyReceivingHeaderLayout

Done:
    mResizing = False
End Sub

Private Function ResolveOperatorWorkbook() As Workbook
    Dim candidate As Workbook

    If mOperatorWorkbook Is Nothing Then Exit Function
    For Each candidate In Application.Workbooks
        If candidate Is mOperatorWorkbook Then
            If Not candidate.IsAddin Then Set ResolveOperatorWorkbook = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Sub RefreshAllViews()
    LoadInventoryCache
    LoadHistoryCache
    RefreshReceiveItems
    RefreshInventory
    RefreshStaging
    modTS_Received.EnforceReceivingSupportSheetsHidden ResolveOperatorWorkbook()
End Sub

Private Sub LoadInventoryCache()
    On Error GoTo ErrHandler
    mItemRows = modTS_Received.LoadReceivingFormInventoryForWorkbook( _
        ResolveOperatorWorkbook(), CStr(mTxtItemSearch.Value))
    Exit Sub
ErrHandler:
    Erase mItemRows
    ShowStatus "Inventory cache load failed: " & Err.Description
End Sub

Public Function CanReuseFor(ByVal operatorWb As Workbook) As Boolean
    Dim captured As Workbook
    If mActivityContext = "" Or mActivityContext <> modActivity.CaptureContext() Then Exit Function
    Set captured = ResolveOperatorWorkbook()
    If Not captured Is Nothing Then CanReuseFor = (captured Is operatorWb)
End Function

Private Sub RefreshReceiveItems()
    FillReceiveItemResults mItemRows
    If mLstReceiveItems.ListCount = 1 Then
        mLstReceiveItems.ListIndex = 0
        LoadSelectedReceiveItemDetails
    End If
End Sub

Private Sub FillReceiveItemResults(ByVal values As Variant)
    Dim r As Long
    Dim rowIndex As Long

    mLstReceiveItems.Clear
    mLstReceiveItems.ColumnCount = 10
    Set mReceiveItemSystemKeys = New Collection
    If IsEmpty(values) Or Not IsArray(values) Then Exit Sub
    For r = 1 To UBound(values, 1)
        mReceiveItemSystemKeys.Add NzText(values(r, 1))
        mLstReceiveItems.AddItem NzText(values(r, 2))
        rowIndex = mLstReceiveItems.ListCount - 1
        mLstReceiveItems.List(rowIndex, 1) = NzText(values(r, 3))
        mLstReceiveItems.List(rowIndex, 2) = NzText(values(r, 4))
        mLstReceiveItems.List(rowIndex, 3) = NzText(values(r, 5))
        mLstReceiveItems.List(rowIndex, 4) = NzText(values(r, 6))
        mLstReceiveItems.List(rowIndex, 5) = vbNullString
        mLstReceiveItems.List(rowIndex, 6) = NzText(values(r, 7))
        mLstReceiveItems.List(rowIndex, 7) = NzText(values(r, 8))
        mLstReceiveItems.List(rowIndex, 8) = NzText(values(r, 9))
        mLstReceiveItems.List(rowIndex, 9) = NzText(values(r, 10))
    Next r
End Sub

Private Function SelectedReceiveItemSystemKey(ByVal rowIndex As Long) As String
    If mReceiveItemSystemKeys Is Nothing Then Exit Function
    If rowIndex < 0 Or rowIndex + 1 > mReceiveItemSystemKeys.Count Then Exit Function
    SelectedReceiveItemSystemKey = NzText(mReceiveItemSystemKeys(rowIndex + 1))
End Function

Private Sub LoadHistoryCache()
    On Error GoTo ErrHandler
    mHistoryRows = modTS_Received.LoadReceivingEntriesHistoryForWorkbook(ResolveOperatorWorkbook())
    Exit Sub
ErrHandler:
    Erase mHistoryRows
    ShowStatus "Receiving history load failed: " & Err.Description
End Sub

Private Sub RefreshInventory()
    On Error GoTo ErrHandler
    FillInventoryList FilterInventoryRows(CStr(mTxtSearch.Value))
    Exit Sub
ErrHandler:
    ShowStatus "Inventory refresh failed: " & Err.Description
End Sub

Private Sub RefreshStaging()
    On Error GoTo ErrHandler
    Dim aggregateReport As String

    If Not modTS_Received.RebuildAggregationForWorkbook( _
        ResolveOperatorWorkbook(), aggregateReport) Then
        ShowStatus aggregateReport
        Exit Sub
    End If
    FillListBox mLstStaged, _
        modTS_Received.LoadReceivingStagingViewForWorkbook( _
            ResolveOperatorWorkbook()), 10
    FillListBox mLstAggregate, _
        modTS_Received.LoadReceivingAggregateViewForWorkbook( _
            ResolveOperatorWorkbook()), 10
    If mLstAggregate.ListCount > 0 Then
        mLstAggregate.ListIndex = 0
        ShowSelectedAggregateReferences
    Else
        ClearAggregateReferenceDetail
    End If
    Exit Sub
ErrHandler:
    ShowStatus "Receiving staging refresh failed: " & Err.Description
End Sub

Private Sub FillListBox(ByVal target As MSForms.ListBox, ByVal values As Variant, ByVal maxCols As Long)
    Dim r As Long
    Dim c As Long
    Dim rowIndex As Long
    Dim rows As Long
    Dim cols As Long

    target.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub

    rows = UBound(values, 1)
    cols = UBound(values, 2)
    If rows < 1 Or cols < 1 Then Exit Sub
    If maxCols > 10 Then maxCols = 10
    If cols > maxCols Then cols = maxCols
    target.ColumnCount = maxCols

    For r = 1 To rows
        target.AddItem NzText(values(r, 1))
        rowIndex = target.ListCount - 1
        For c = 2 To cols
            target.List(rowIndex, c - 1) = NzText(values(r, c))
        Next c
    Next r
End Sub

Private Function FilterInventoryRows(ByVal filterText As String) As Variant
    Dim tokens() As String
    Dim sourceRows As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim haystack As String
    Dim token As Variant
    Dim matched As Boolean

    If IsEmpty(mHistoryRows) Or Not IsArray(mHistoryRows) Then Exit Function
    sourceRows = mHistoryRows
    filterText = NormalizeSearchText(filterText)
    tokens = Split(filterText, " ")

    ReDim result(1 To UBound(sourceRows, 1), 1 To UBound(sourceRows, 2))
    For r = 1 To UBound(sourceRows, 1)
        haystack = ""
        For c = 1 To UBound(sourceRows, 2)
            haystack = haystack & " " & NormalizeSearchText(NzText(sourceRows(r, c)))
        Next c

        matched = True
        If filterText <> "" Then
            For Each token In tokens
                If Trim$(CStr(token)) <> "" Then
                    If InStr(1, haystack, CStr(token), vbTextCompare) = 0 Then
                        matched = False
                        Exit For
                    End If
                End If
            Next token
        End If

        If matched Then
            outRow = outRow + 1
            For c = 1 To UBound(sourceRows, 2)
                result(outRow, c) = sourceRows(r, c)
            Next c
        End If
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To UBound(sourceRows, 2))
    For r = 1 To outRow
        For c = 1 To UBound(sourceRows, 2)
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    FilterInventoryRows = trimmed
End Function

Private Sub FillInventoryList(ByVal values As Variant)
    FillListBox mLstInventory, values, 10
    If IsEmpty(values) Or Not IsArray(values) Then
        If Trim$(CStr(mTxtSearch.Value)) <> "" Then ShowStatus "No receiving history rows match: " & CStr(mTxtSearch.Value)
    Else
        ShowStatus "Receiving history rows shown: " & CStr(UBound(values, 1))
    End If
End Sub

Private Sub mTxtSearch_Change()
    If mLoading Then Exit Sub
    RefreshInventory
End Sub

Private Sub mTxtItemSearch_Change()
    If mLoading Then Exit Sub
    LoadInventoryCache
    RefreshReceiveItems
End Sub

Private Sub mLstReceiveItems_Click()
    NavigationSelection "lstReceiveItems"
End Sub

Private Sub LoadSelectedReceiveItemDetails()
    If mLstReceiveItems Is Nothing Then Exit Sub
    If mLstReceiveItems.ListIndex < 0 Then Exit Sub
    mTxtReceiveLocation.Value = NzText( _
        mLstReceiveItems.List(mLstReceiveItems.ListIndex, 4))
    If mTabs.Value = 1 Then
        mTxtLotNumber.Value = NzText( _
            mLstReceiveItems.List(mLstReceiveItems.ListIndex, 6))
        mCboCondition.Value = NzText( _
            mLstReceiveItems.List(mLstReceiveItems.ListIndex, 7))
    End If
End Sub

Private Sub mBtnRefresh_Click()
    RefreshClicked
End Sub

Private Sub mBtnAdd_Click()
    AddSelectedInventory
End Sub

Private Sub mBtnConfirm_Click()
    On Error GoTo ErrHandler
    Dim report As String, statusMessage As String
    Dim succeeded As Boolean, quietStarted As Boolean

    mLastConfirmQuietActive = False
    If mTabs.Value = 1 Then
        ShowPersistencePending "Saving inventory dispositions to the warehouse server..."
    Else
        ShowPersistencePending "Saving received inventory to the warehouse server..."
    End If
    modUiQuiet.BeginQuietUi mOperatorWorkbook
    quietStarted = True
    mLastConfirmQuietActive = modUiQuiet.QuietUiIsActive()
    succeeded = modReceivingActivityAction.ConfirmWrites( _
        mOperatorWorkbook, mActivityContext, mTabs.Value = 0, report)
    modTS_Received.RecordConfirmWritesResult succeeded, report
    If succeeded Then
        mTxtRef.Value = ""
        mTxtReceiptId.Value = DefaultReceiptId()
        RefreshAllViews
        statusMessage = "Confirm Writes finished; staged receiving rows cleared." & vbCrLf & report
    Else
        RefreshAllViews
        statusMessage = "Confirm Writes did not complete; staged receiving rows were kept. " & report
    End If
    GoTo CleanExit
ErrHandler:
    statusMessage = "Confirm Writes failed: " & Err.Description
CleanExit:
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    On Error GoTo 0
    ShowStatus statusMessage
End Sub

Private Sub mLstAggregate_Click()
    NavigationSelection "lstAggregate"
End Sub

Private Sub ShowSelectedAggregateReferences()
    If mTxtAggregateReferences Is Nothing Or mLstAggregate Is Nothing Then Exit Sub
    If mLstAggregate.ListIndex < 0 Then
        ClearAggregateReferenceDetail
    Else
        mTxtAggregateReferences.Text = CStr(mLstAggregate.List(mLstAggregate.ListIndex, 0))
    End If
End Sub

Private Sub ClearAggregateReferenceDetail()
    If Not mTxtAggregateReferences Is Nothing Then mTxtAggregateReferences.Text = vbNullString
End Sub

Private Sub mBtnClear_Click()
    On Error GoTo ErrHandler
    Dim report As String
    If modReceivingActivityAction.LocalAction(ResolveOperatorWorkbook(), mActivityContext, True, report) Then RefreshStaging
    ShowStatus report
    Exit Sub
ErrHandler:
    ShowStatus "Clear failed: " & Err.Description
End Sub

Private Sub mBtnClose_Click()
    modReceivingActivityAction.CloseForm Me, ResolveOperatorWorkbook(), mActivityContext
End Sub

Private Sub RefreshClicked()
    On Error GoTo ErrHandler
    Dim report As String
    If modReceivingActivityAction.LocalAction(ResolveOperatorWorkbook(), mActivityContext, False, report) Then RefreshAllViews
    ShowStatus report
    Exit Sub
ErrHandler:
    ShowStatus "Refresh failed: " & Err.Description
End Sub

Private Sub AddSelectedInventory()
    On Error GoTo ErrHandler

    Dim idx As Long
    Dim qtyVal As Double
    Dim refVal As String
    Dim report As String
    Dim itemCode As String
    Dim sourceSystemKey As String
    Dim receiptType As String
    Dim returnReason As String
    Dim activityId As String, notice As String, outcome As String, controlId As String

    report = "Session or warehouse changed. Reopen Receiving before staging."
    If mActivityContext = "" Or mActivityContext <> modActivity.CaptureContext() Then GoTo Done
    controlId = "RECEIVING_ADD_SELECTED"
    If mTabs.Value = 1 Then controlId = "DISPOSITION_ADD_SELECTED"
    activityId = modActivity.BeginAction(controlId, mActivityContext, notice)
    outcome = "REJECTED"
    If Not modReceivingAddInput.Valid(Me, receiptType, report) Then GoTo Done
    idx = mLstReceiveItems.ListIndex
    refVal = Trim$(CStr(mTxtRef.Value))
    sourceSystemKey = SelectedReceiveItemSystemKey(idx)
    itemCode = NzText(mLstReceiveItems.List(idx, 0))
    qtyVal = CDbl(Val(CStr(mTxtQty.Value)))
    returnReason = Trim$(CStr(mTxtReturnReason.Value))
    outcome = "FAILED"
    If modTS_Received.StageReceivingFormItemForWorkbook( _
        ResolveOperatorWorkbook(), refVal, sourceSystemKey, itemCode, qtyVal, report, _
        Trim$(CStr(mTxtReceiveLocation.Value)), Trim$(CStr(mTxtLotNumber.Value)), _
        CStr(mCboCondition.Value), receiptType, returnReason) Then
        outcome = "STAGED"
        RefreshStaging
    End If
Done:
    modReceivingAddInput.Finish activityId, outcome, notice, report
    ShowStatus report
    Exit Sub

ErrHandler:
    If outcome <> "STAGED" Then outcome = "FAILED"
    report = "Add failed: " & Err.Description
    Resume Done
End Sub

Private Sub mTabs_Change()
    If mLoading Then Exit Sub
    NavigationSelection "tabsReceiving"
End Sub

Private Sub mLstInventory_Click()
    NavigationSelection "lstInventory"
End Sub
Private Sub mLstStaged_Click()
    NavigationSelection "lstStaged"
End Sub
Private Sub mCboCondition_Click()
    NavigationSelection "cboCondition"
End Sub
Private Sub mCboDisposition_Click()
    NavigationSelection "cboDisposition"
End Sub

Private Sub NavigationSelection(ByVal controlName As String)
    Dim activityId As String, notice As String, report As String, outcome As String
    On Error GoTo Failed
    If Not modReceivingNavigation.BeginSelection(Me, mNavigationInputs, controlName, _
        ResolveOperatorWorkbook(), mActivityContext, activityId, notice, report) Then
        If report <> "" Then ShowStatus report
        Exit Sub
    End If
    outcome = "FAILED"
    Select Case controlName
        Case "tabsReceiving": ApplyReceivingTab
        Case "lstReceiveItems": LoadSelectedReceiveItemDetails
        Case "lstAggregate": ShowSelectedAggregateReferences
    End Select
    outcome = "SELECTED"
Done:
    modReceivingAddInput.Finish activityId, outcome, notice, report
    If report <> "" Then ShowStatus Trim$(CStr(mTxtStatus.Value) & " " & report)
    Exit Sub
Failed:
    report = "Selection failed: " & Err.Description
    Resume Done
End Sub

Private Sub ApplyReceivingTab()
    If mTabs Is Nothing Or Not mBuilt Then Exit Sub
    ShowStatus modReceivingNavigation.ApplyTab(Me, mTabs.Value)
    ResizeReceivingLayout
End Sub

Private Sub ApplyReceivingHeaderLayout()
    AlignReceivingHeader mLblReceiveItemsHeader, mLstReceiveItems, _
        Array("Code", "Item", "UOM", "Available", "Location", "Capacity (coming later)", "Lot", "Condition", "Description", "Vendor")
    AlignReceivingHeader mLblInventoryHeader, mLstInventory, _
        Array("Date", "Type", "Reference", "Item", "Qty", "UOM", "Location", "Lot", "Condition", "Return reason")
    AlignReceivingHeader mLblStagedHeader, mLstStaged, _
        Array("Reference", "Type", "Item", "Qty", "UOM", "Location", "Lot", "Vendor", "Condition", "Return reason")
    AlignReceivingHeader mLblAggregateHeader, mLstAggregate, _
        Array("Reference", "Type", "Code", "Item", "UOM", "Qty", "Location", "Lot", "Condition", "Return reason")
End Sub

Private Sub AlignReceivingHeader(ByVal headerLabel As MSForms.Label, _
                                 ByVal targetList As MSForms.ListBox, _
                                 ByVal headings As Variant)
    If headerLabel Is Nothing Or targetList Is Nothing Then Exit Sub
    headerLabel.Left = targetList.Left
    headerLabel.Top = targetList.Top - headerLabel.Height - 2
    headerLabel.Width = targetList.Width
    headerLabel.Font.Name = "Courier New"
    headerLabel.Font.Size = 8
    headerLabel.WordWrap = False
    headerLabel.AutoSize = False
    headerLabel.Caption = modReceivingFormWindow.BuildReceivingHeaderCaption(targetList, headings)
End Sub

Public Function TestReturnsTabContract(ByVal operatorWb As Workbook) As String
    If operatorWb Is Nothing Then Exit Function
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    mTabs.Value = 1
    ApplyReceivingTab
    TestReturnsTabContract = _
        "OK|Selected=" & mTabs.SelectedItem.Caption & _
        "|AddCaption=" & mBtnAdd.Caption & _
        "|ConditionVisible=" & CStr(mCboCondition.Visible) & _
        "|ReturnReasonVisible=" & CStr(mTxtReturnReason.Visible) & _
        "|DispositionVisible=" & CStr(mCboDisposition.Visible) & _
        "|DispositionDefault=" & CStr(mCboDisposition.Value) & _
        "|DispositionOptions=RETURN,DUMP" & _
        "|HistoryTitle=" & mLblInventoryTitle.Caption & _
        "|TallyTitle=" & mLblStagedTitle.Caption & _
        "|AggregateTitle=" & mLblAggregateTitle.Caption & _
        "|ItemConditionColumn=True" & _
        "|ReceiptEventType=" & CStr(mCboDisposition.Value)
End Function

Public Function TestStageInboundReturnActionForWorkbook(ByVal operatorWb As Workbook) As String
    On Error GoTo Failed

    Dim protectedReceiptStatus As String
    Dim receiptStatus As String

    If operatorWb Is Nothing Then
        TestStageInboundReturnActionForWorkbook = "FAIL|Operator workbook is required."
        Exit Function
    End If
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    InitializeFromReceiving

    mTabs.Value = 0
    ApplyReceivingTab
    If mLstReceiveItems.ListCount = 0 Then
        TestStageInboundReturnActionForWorkbook = "FAIL|No receipt item choices."
        Exit Function
    End If
    mLstReceiveItems.ListIndex = 0
    LoadSelectedReceiveItemDetails
    mTxtRef.Value = "RECEIPT-TEST"
    mTxtQty.Value = "1"
    mCboCondition.Value = "GOOD"
    operatorWb.Worksheets("ReceivedTally").Protect
    mBtnAdd_Click
    operatorWb.Worksheets("ReceivedTally").Unprotect
    protectedReceiptStatus = CStr(mTxtStatus.Text)
    If InStr(1, protectedReceiptStatus, _
             "Receiving staging failed: Stage=", vbBinaryCompare) = 0 _
       Or InStr(1, protectedReceiptStatus, "; Error=", vbBinaryCompare) = 0 _
       Or InStr(1, protectedReceiptStatus, "; Source=", vbBinaryCompare) = 0 Then
        TestStageInboundReturnActionForWorkbook = _
            "FAIL|ProtectedReceipt=" & protectedReceiptStatus
        Exit Function
    End If
    modTS_Received.ClearReceivingFormStagingForWorkbook operatorWb
    RefreshStaging

    mBtnAdd_Click
    receiptStatus = CStr(mTxtStatus.Text)
    If mLstStaged.ListCount = 0 Then
        TestStageInboundReturnActionForWorkbook = "FAIL|ReceiptAction=" & receiptStatus
        Exit Function
    End If
    modTS_Received.ClearReceivingFormStagingForWorkbook operatorWb
    RefreshStaging

    mTabs.Value = 1
    ApplyReceivingTab
    If mLstReceiveItems.ListCount = 0 Then
        TestStageInboundReturnActionForWorkbook = "FAIL|No return item choices."
        Exit Function
    End If
    mLstReceiveItems.ListIndex = 0
    LoadSelectedReceiveItemDetails
    mTxtRef.Value = "RETURN-TEST"
    mTxtQty.Value = "1"
    mCboDisposition.Value = "RETURN"
    mTxtReturnReason.Value = "TEST RETURN"
    mBtnAdd_Click
    If mLstStaged.ListCount = 0 Then
        TestStageInboundReturnActionForWorkbook = "FAIL|" & CStr(mTxtStatus.Text)
    Else
        TestStageInboundReturnActionForWorkbook = _
            "OK|ReceiptAction=True|StagedRows=" & CStr(mLstStaged.ListCount) & _
            "|ReceiptType=" & NzText(mLstStaged.List(0, 1)) & _
            "|Condition=" & NzText(mLstStaged.List(0, 8)) & _
            "|Reason=" & NzText(mLstStaged.List(0, 9))
    End If
    Exit Function
Failed:
    TestStageInboundReturnActionForWorkbook = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    On Error Resume Next
    operatorWb.Worksheets("ReceivedTally").Unprotect
    On Error GoTo 0
End Function

Public Function TestStageProtectedDispositionActionForWorkbook(ByVal operatorWb As Workbook) As String
    On Error GoTo Failed

    If operatorWb Is Nothing Then
        TestStageProtectedDispositionActionForWorkbook = "FAIL|Operator workbook is required."
        Exit Function
    End If
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    InitializeFromReceiving
    mTabs.Value = 1
    ApplyReceivingTab
    If mLstReceiveItems.ListCount = 0 Then
        TestStageProtectedDispositionActionForWorkbook = "FAIL|No return item choices."
        Exit Function
    End If
    mLstReceiveItems.ListIndex = 0
    LoadSelectedReceiveItemDetails
    mTxtRef.Value = "RETURN-TEST"
    mTxtQty.Value = "1"
    mCboDisposition.Value = "RETURN"
    mTxtReturnReason.Value = "TEST RETURN"
    operatorWb.Worksheets("ReceivedTally").Protect
    mBtnAdd_Click
    operatorWb.Worksheets("ReceivedTally").Unprotect
    TestStageProtectedDispositionActionForWorkbook = "FAIL|" & CStr(mTxtStatus.Text)
    Exit Function
Failed:
    On Error Resume Next
    operatorWb.Worksheets("ReceivedTally").Unprotect
    On Error GoTo 0
    TestStageProtectedDispositionActionForWorkbook = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestReceivingSearchAndHeaderContract() As String
    Dim aligned As Boolean
    Dim capacityStub As Boolean
    Dim headersSingleLine As Boolean
    Dim hiddenSystemKeyMap As Boolean
    Dim searchRowsLoaded As Boolean
    Dim tenColumnItemResults As Boolean
    Dim receiveWidths As Variant
    Dim sampleRows(1 To 1, 1 To 10) As Variant

    If Not mBuilt Then BuildLayout
    If Not mOperatorWorkbook Is Nothing Then mTxtItemSearch_Change
    sampleRows(1, 1) = "TEST-SYSTEM-KEY"
    sampleRows(1, 2) = "TEST-CODE"
    sampleRows(1, 3) = "Test Item"
    sampleRows(1, 4) = "EA"
    sampleRows(1, 5) = 1
    sampleRows(1, 6) = "TEST"
    sampleRows(1, 7) = ""
    sampleRows(1, 8) = "GOOD"
    sampleRows(1, 9) = "Test description"
    sampleRows(1, 10) = "Test vendor"
    FillReceiveItemResults sampleRows
    ApplyReceivingHeaderLayout
    aligned = _
        (mLblReceiveItemsHeader.Left = mLstReceiveItems.Left) And _
        (mLblReceiveItemsHeader.Width = mLstReceiveItems.Width) And _
        (mLblInventoryHeader.Left = mLstInventory.Left) And _
        (mLblStagedHeader.Left = mLstStaged.Left) And _
        (mLblAggregateHeader.Left = mLstAggregate.Left)
    receiveWidths = Split(CStr(mLstReceiveItems.ColumnWidths), ";")
    tenColumnItemResults = (mLstReceiveItems.ColumnCount = 10) And _
        (UBound(receiveWidths) = 9)
    capacityStub = tenColumnItemResults And _
        (Val(CStr(receiveWidths(5))) >= 80) And _
        (mLstReceiveItems.ListCount > 0) And _
        (NzText(mLstReceiveItems.List(0, 5)) = "")
    hiddenSystemKeyMap = (mLstReceiveItems.ListCount > 0) And _
        (SelectedReceiveItemSystemKey(0) <> "")
    searchRowsLoaded = (mLstReceiveItems.ListCount > 0)
    headersSingleLine = (Not mLblReceiveItemsHeader.WordWrap) And _
        (Not mLblInventoryHeader.WordWrap) And _
        (Not mLblStagedHeader.WordWrap) And _
        (Not mLblAggregateHeader.WordWrap)
    If aligned And capacityStub And hiddenSystemKeyMap And _
       searchRowsLoaded And headersSingleLine Then
        TestReceivingSearchAndHeaderContract = _
            "OK|DedicatedItemResults=True|Location=True|OptionalLot=True|Condition=True|Returns=True|ViewerReadOnly=True|ReceivingHeaderColumnsAligned=True|CapacityStub=" & CStr(capacityStub) & _
            "|SearchRowsLoaded=" & CStr(searchRowsLoaded) & _
            "|HiddenSystemKeyMap=" & CStr(hiddenSystemKeyMap) & _
            "|TenColumnItemResults=" & CStr(tenColumnItemResults) & _
            "|HeadersSingleLine=" & CStr(headersSingleLine)
    Else
        TestReceivingSearchAndHeaderContract = _
            "FAIL|ReceivingHeaderColumnsAligned=" & CStr(aligned) & _
            "|CapacityStub=" & CStr(capacityStub) & _
            "|SearchRowsLoaded=" & CStr(searchRowsLoaded) & _
            "|HiddenSystemKeyMap=" & CStr(hiddenSystemKeyMap) & _
            "|TenColumnItemResults=" & CStr(tenColumnItemResults) & _
            "|HeadersSingleLine=" & CStr(headersSingleLine)
    End If
End Function

Private Sub ShowStatus(ByVal messageText As String)
    If mTxtStatus Is Nothing Then Exit Sub
    mTxtStatus.Text = messageText
End Sub

Private Sub ShowPersistencePending(ByVal messageText As String)
    ShowStatus messageText
    Me.Repaint
    DoEvents
End Sub

Private Function DefaultReceiptId() As String
    DefaultReceiptId = "RCV-" & Format$(Now, "yyyymmdd-hhnnss")
End Function

Private Function NormalizeSearchText(ByVal valueIn As String) As String
    valueIn = LCase$(Trim$(valueIn))
    Do While InStr(1, valueIn, "  ", vbBinaryCompare) > 0
        valueIn = Replace$(valueIn, "  ", " ")
    Loop
    NormalizeSearchText = valueIn
End Function

Private Function NzText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then
        NzText = ""
    Else
        NzText = Trim$(CStr(valueIn))
    End If
End Function

Private Function MaxDoubleReceiving(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then
        MaxDoubleReceiving = a
    Else
        MaxDoubleReceiving = b
    End If
End Function
