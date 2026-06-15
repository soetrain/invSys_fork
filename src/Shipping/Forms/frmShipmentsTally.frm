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
Private mLstReadiness As MSForms.ListBox
Private mLblStatus As MSForms.Label

Private mShippables As Variant
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

    Dim previousPointer As Long
    Dim quietStarted As Boolean

    If Not mBuilt Then BuildLayout
    previousPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass
    modUiQuiet.BeginQuietUi ActiveWorkbook
    quietStarted = True

    mLoading = True
    mChkUseExisting.Value = modTS_Shipments.ShipmentsFormUseExistingInventory()
    LoadShippables
    LoadShipmentState
    mLoading = False

    If mLstShippables.ListCount > 0 Then
        mLstShippables.ListIndex = 0
        LoadSelectedShippable
    End If

CleanExit:
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    Me.MousePointer = previousPointer
    On Error GoTo 0
    Exit Sub

FailInit:
    ShowStatus "Shipments form load failed: " & Err.Description
    Resume CleanExit
End Sub

Private Sub BuildLayout()
    If mBuilt Then Exit Sub
    mBuilt = True

    Me.Caption = "Shipping Shipments"
    Me.Width = 860
    Me.Height = 675
    Me.ScrollBars = fmScrollBarsBoth
    Me.ScrollWidth = 850
    Me.ScrollHeight = 780

    AddLabel "lblTitle", "Shipments", 12, 10, 140, 20, True
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 774, 10, 58, 24)

    AddLabel "lblPicker", "Box Picker", 12, 42, 70, 18, False
    Set mTxtPicker = AddTextBox("txtPicker", 86, 38, 250, 22)
    Set mChkUseExisting = AddCheckBox("chkUseExisting", "Use existing shippable inventory", 356, 38, 190, 22)

    AddShippableHeaders 12, 70
    Set mLstShippables = AddListBox("lstShippables", 12, 90, 820, 92)
    With mLstShippables
        .ColumnCount = 6
        .ColumnWidths = "190 pt;52 pt;70 pt;42 pt;116 pt;52 pt"
    End With

    AddLabel "lblRef", "Ref", 12, 198, 34, 18, False
    AddLabel "lblBox", "Box", 108, 198, 34, 18, False
    AddLabel "lblVersion", "Version", 270, 198, 52, 18, False
    AddLabel "lblQty", "Qty", 336, 198, 34, 18, False
    AddLabel "lblUom", "UOM", 410, 198, 40, 18, False
    AddLabel "lblLocation", "Location", 470, 198, 60, 18, False
    AddLabel "lblRow", "ROW", 620, 198, 40, 18, False

    Set mTxtRef = AddTextBox("txtRef", 12, 216, 82, 22)
    Set mTxtBox = AddTextBox("txtBox", 108, 216, 148, 22)
    Set mTxtVersion = AddTextBox("txtVersion", 270, 216, 52, 22)
    Set mTxtQty = AddTextBox("txtQty", 336, 216, 52, 22)
    Set mTxtUom = AddTextBox("txtUom", 410, 216, 44, 22)
    Set mTxtLocation = AddTextBox("txtLocation", 470, 216, 132, 22)
    Set mTxtRow = AddTextBox("txtRow", 620, 216, 52, 22)
    Set mTxtDescription = AddTextBox("txtDescription", 12, 248, 660, 22)
    LockTextBox mTxtBox
    LockTextBox mTxtVersion
    LockTextBox mTxtUom
    LockTextBox mTxtLocation
    LockTextBox mTxtRow
    Set mBtnAdd = AddButton("btnAdd", "Add", 690, 214, 44, 26)
    Set mBtnUpdate = AddButton("btnUpdate", "Update", 740, 214, 52, 26)
    Set mBtnRemove = AddButton("btnRemove", "Remove", 798, 214, 58, 26)

    AddLabel "lblShipments", "Shipments", 12, 286, 90, 18, True
    AddShipmentLineHeaders 12, 310
    Set mLstShipments = AddListBox("lstShipments", 12, 330, 820, 92)
    With mLstShipments
        .ColumnCount = 8
        .ColumnWidths = "78 pt;160 pt;54 pt;44 pt;102 pt;48 pt;220 pt;0 pt"
    End With
    Set mBtnHold = AddButton("btnHold", "Send Hold", 744, 286, 88, 24)

    AddLabel "lblHold", "Not Shipped", 12, 438, 100, 18, True
    AddShipmentLineHeaders 12, 462
    Set mLstHold = AddListBox("lstHold", 12, 482, 820, 72)
    With mLstHold
        .ColumnCount = 8
        .ColumnWidths = "78 pt;160 pt;54 pt;44 pt;102 pt;48 pt;220 pt;0 pt"
    End With
    Set mBtnReturn = AddButton("btnReturn", "Return", 744, 438, 88, 24)

    AddLabel "lblReadiness", "Readiness", 12, 568, 90, 18, True
    AddReadinessHeaders 12, 592
    Set mLstReadiness = AddListBox("lstReadiness", 12, 612, 820, 104)
    With mLstReadiness
        .ColumnCount = 9
        .ColumnWidths = "62 pt;180 pt;62 pt;62 pt;62 pt;42 pt;96 pt;48 pt;78 pt"
    End With

    Set mLblStatus = AddLabel("lblStatus", "", 12, 732, 500, 34, False)
    Set mBtnStage = AddButton("btnStage", "To Shipments", 560, 730, 92, 30)
    Set mBtnSend = AddButton("btnSend", "Shipments Sent", 662, 730, 104, 30)
    Set mBtnClose = AddButton("btnClose", "Close", 776, 730, 56, 30)

    InitializeAnchors
End Sub

Private Sub LoadShippables()
    On Error GoTo FailSoft

    mShippables = modTS_Shipments.ShipmentsFormLoadShippables()
    RenderShippables
    Exit Sub

FailSoft:
    ShowStatus "Could not load shippables: " & Err.Description
End Sub

Private Sub LoadShipmentState()
    RenderLineList mLstShipments, modTS_Shipments.ShipmentsFormLoadLines(False)
    RenderLineList mLstHold, modTS_Shipments.ShipmentsFormLoadLines(True)
    RenderReadiness modTS_Shipments.ShipmentsFormLoadReadiness()
End Sub

Private Sub LoadShipmentLineState()
    RenderLineList mLstShipments, modTS_Shipments.ShipmentsFormLoadLines(False)
    RenderLineList mLstHold, modTS_Shipments.ShipmentsFormLoadLines(True)
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

    ReDim displayRows(0 To shownCount - 1, 0 To 5)
    idx = 0
    For r = 1 To UBound(mShippables, 1)
        If Not ShippableMatchesFilter(r, filterText) Then GoTo NextRow
        displayRows(idx, 0) = NzText(mShippables(r, 2))
        displayRows(idx, 1) = NzText(mShippables(r, 3))
        displayRows(idx, 2) = DisplayQtyText(NzText(mShippables(r, 4)))
        displayRows(idx, 3) = NzText(mShippables(r, 5))
        displayRows(idx, 4) = NzText(mShippables(r, 6))
        displayRows(idx, 5) = NzText(mShippables(r, 1))
        idx = idx + 1
NextRow:
    Next r
    mLstShippables.List = displayRows
    Exit Sub

FailSoft:
    ShowStatus "Shippable render failed: " & Err.Description
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
    ReDim displayRows(0 To UBound(rowsData, 1) - 1, 0 To 7)
    For r = 1 To UBound(rowsData, 1)
        displayRows(r - 1, 0) = NzText(rowsData(r, 1))
        displayRows(r - 1, 1) = NzText(rowsData(r, 2))
        displayRows(r - 1, 2) = FormatQuantity(ParseNumber(NzText(rowsData(r, 3))))
        displayRows(r - 1, 3) = NzText(rowsData(r, 4))
        displayRows(r - 1, 4) = NzText(rowsData(r, 5))
        displayRows(r - 1, 5) = NzText(rowsData(r, 6))
        displayRows(r - 1, 6) = NzText(rowsData(r, 7))
        displayRows(r - 1, 7) = NzText(rowsData(r, 8))
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
    mTxtUom.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 3))
    mTxtLocation.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 4))
    mTxtRow.Value = NzText(mLstShippables.List(mLstShippables.ListIndex, 5))
    If Trim$(NzText(mTxtQty.Value)) = "" Then mTxtQty.Value = "1"
    If Trim$(NzText(mTxtDescription.Value)) = "" Then mTxtDescription.Value = NzText(mTxtVersion.Value)
End Sub

Private Sub LoadSelectedLine(ByVal lst As MSForms.ListBox)
    If lst Is Nothing Then Exit Sub
    If lst.ListIndex < 0 Then Exit Sub
    mTxtRef.Value = NzText(lst.List(lst.ListIndex, 0))
    mTxtBox.Value = NzText(lst.List(lst.ListIndex, 1))
    mTxtQty.Value = NzText(lst.List(lst.ListIndex, 2))
    mTxtUom.Value = NzText(lst.List(lst.ListIndex, 3))
    mTxtLocation.Value = NzText(lst.List(lst.ListIndex, 4))
    mTxtRow.Value = NzText(lst.List(lst.ListIndex, 5))
    mTxtDescription.Value = NzText(lst.List(lst.ListIndex, 6))
    mTxtVersion.Value = ""
End Sub

Private Sub CommitCurrentLine(ByVal actionName As String)
    On Error GoTo FailSoft

    Dim report As String
    Dim rowIndex As Long
    Dim ok As Boolean

    rowIndex = SelectedShipmentTableRow()
    ok = modTS_Shipments.ShipmentsFormCommitLine("SHIP", _
                                                 actionName, _
                                                 rowIndex, _
                                                 NzText(mTxtRef.Value), _
                                                 NzText(mTxtBox.Value), _
                                                 ParseNumber(NzText(mTxtQty.Value)), _
                                                 CLng(Val(NzText(mTxtRow.Value))), _
                                                 NzText(mTxtUom.Value), _
                                                 NzText(mTxtLocation.Value), _
                                                 NzText(mTxtDescription.Value), _
                                                 report)
    RefreshAfterAction report, ok
    Exit Sub

FailSoft:
    ShowStatus "Shipment row action failed: " & Err.Description
End Sub

Private Function SelectedShipmentTableRow() As Long
    If mLstShipments Is Nothing Then Exit Function
    If mLstShipments.ListIndex < 0 Then Exit Function
    SelectedShipmentTableRow = CLng(Val(NzText(mLstShipments.List(mLstShipments.ListIndex, 7))))
End Function

Private Function SelectedHoldTableRow() As Long
    If mLstHold Is Nothing Then Exit Function
    If mLstHold.ListIndex < 0 Then Exit Function
    SelectedHoldTableRow = CLng(Val(NzText(mLstHold.List(mLstHold.ListIndex, 7))))
End Function

Private Sub RefreshAfterAction(ByVal report As String, ByVal ok As Boolean)
    Dim previousPointer As Long

    previousPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass
    mLoading = True
    LoadShipmentLineState
    mLoading = False
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
    modTS_Shipments.ShipmentsFormSetUseExistingInventory CBool(mChkUseExisting.Value)
    LoadShipmentState
End Sub

Private Sub mBtnRefresh_Click()
    InitializeFromShipping
    ShowStatus "Shipments form refreshed."
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

    If moveToHold Then
        Set lst = mLstShipments
    Else
        Set lst = mLstHold
    End If
    If lst.ListIndex < 0 Then
        ShowStatus "Select a shipment row first."
        Exit Sub
    End If

    ok = modTS_Shipments.ShipmentsFormMoveHold(NzText(lst.List(lst.ListIndex, 0)), _
                                               NzText(lst.List(lst.ListIndex, 1)), _
                                               ParseNumber(NzText(lst.List(lst.ListIndex, 2))), _
                                               moveToHold, _
                                               report)
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

    previousPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass
    modUiQuiet.BeginQuietUi ActiveWorkbook
    quietStarted = True
    startedAt = Timer
    If stageOnly Then
        ok = modTS_Shipments.ShipmentsFormRunToShipments(report)
    Else
        ok = modTS_Shipments.ShipmentsFormRunShipmentsSent(report)
    End If
    elapsedMs = ElapsedMilliseconds(startedAt)
    If quietStarted Then
        modUiQuiet.EndQuietUi
        quietStarted = False
    End If
    Me.MousePointer = previousPointer
    LoadShipmentState
    report = AppendTiming(report, elapsedMs)
    ShowStatus report
    If report <> "" Then MsgBox report, IIf(ok, vbInformation, vbExclamation)
    Exit Sub

FailSoft:
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    Me.MousePointer = previousPointer
    On Error GoTo 0
    ShowStatus "Shipping action failed: " & Err.Description
End Sub

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

    mAnchors.Add mBtnRefresh, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtPicker, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstShippables, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtDescription, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnAdd, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnUpdate, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnRemove, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstShipments, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnHold, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstHold, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnReturn, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstReadiness, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mLblStatus, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnStage, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnSend, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnClose, ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub

Private Sub AddShippableHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrShipBox", "Box", leftPos, topPos, 184
    AddHeaderLabel "hdrShipVersion", "Version", leftPos + 194, topPos, 54
    AddHeaderLabel "hdrShipInv", "Current Inv", leftPos + 252, topPos, 72
    AddHeaderLabel "hdrShipUom", "UOM", leftPos + 330, topPos, 42
    AddHeaderLabel "hdrShipLoc", "Location", leftPos + 378, topPos, 116
    AddHeaderLabel "hdrShipRow", "ROW", leftPos + 500, topPos, 54
End Sub

Private Sub AddShipmentLineHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel UniqueHeaderName("hdrRef", topPos), "Ref", leftPos, topPos, 76
    AddHeaderLabel UniqueHeaderName("hdrLineBox", topPos), "Box", leftPos + 82, topPos, 154
    AddHeaderLabel UniqueHeaderName("hdrLineQty", topPos), "Qty", leftPos + 246, topPos, 50
    AddHeaderLabel UniqueHeaderName("hdrLineUom", topPos), "UOM", leftPos + 302, topPos, 40
    AddHeaderLabel UniqueHeaderName("hdrLineLoc", topPos), "Location", leftPos + 350, topPos, 98
    AddHeaderLabel UniqueHeaderName("hdrLineRow", topPos), "ROW", leftPos + 456, topPos, 46
    AddHeaderLabel UniqueHeaderName("hdrLineDesc", topPos), "Description", leftPos + 508, topPos, 210
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
