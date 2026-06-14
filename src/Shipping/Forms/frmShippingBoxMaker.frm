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
Private mPendingShippableInv As Object
Private mSelectedPackageRow As Long
Private mLoading As Boolean
Private mBuilt As Boolean

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Public Sub InitializeFromShipping()
    If Not mBuilt Then BuildLayout

    mLoading = True
    LoadSavedBoxes
    mLoading = False

    If mCboBoxes.ListCount > 0 Then
        mCboBoxes.ListIndex = 0
        LoadSelectedBox
    Else
        ClearBoxFields
        ShowStatus "No active saved box designs found."
    End If
End Sub

Private Sub BuildLayout()
    If mBuilt Then Exit Sub
    mBuilt = True

    Me.Caption = "Shipping Box Maker"
    Me.Width = 780
    Me.Height = 590

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
        .ColumnCount = 5
        .ColumnWidths = "220 pt;70 pt;44 pt;120 pt;56 pt"
    End With

    AddLabel "lblComponents", "Components To Deduct", 12, 300, 170, 18, True
    AddComponentHeaders 12, 322
    Set mLstComponents = AddListBox("lstComponents", 12, 342, 752, 142)
    With mLstComponents
        .ColumnCount = 9
        .ColumnWidths = "126 pt;42 pt;50 pt;58 pt;54 pt;38 pt;76 pt;62 pt;150 pt"
    End With

    Set mBtnMake = AddButton("btnMake", "Make Boxes", 506, 496, 86, 30)
    Set mBtnUnmake = AddButton("btnUnmake", "Unmake Boxes", 604, 496, 92, 30)
    Set mBtnClose = AddButton("btnClose", "Close", 706, 496, 58, 30)
    Set mLblStatus = AddLabel("lblStatus", "", 12, 494, 480, 44, False)
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

    If mCboBoxes.ListIndex < 0 Then Exit Sub
    mSelectedPackageRow = CLng(Val(CStr(mCboBoxes.List(mCboBoxes.ListIndex, 0))))
    mTxtBoxName.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 1))
    mTxtUom.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 3))
    mTxtLocation.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 4))
    mTxtDescription.Value = CStr(mCboBoxes.List(mCboBoxes.ListIndex, 5))
    LoadVersionsForPackage mSelectedPackageRow
    LoadSelectedVersionComponents
    RenderPackageInventory
    SelectShippableInventoryRow
    Exit Sub

FailSoft:
    ShowStatus "Could not load selected box: " & Err.Description
End Sub

Private Sub LoadVersionsForPackage(ByVal packageRow As Long)
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long
    Dim c As Long
    Dim idx As Long

    mCboVersions.Clear
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

Private Sub LoadSelectedVersionComponents()
    On Error GoTo FailSoft

    If mSelectedPackageRow <= 0 Then Exit Sub
    If mCboVersions.ListIndex < 0 Then Exit Sub

    mComponents = modTS_Shipments.BoxMakerFormLoadVersionComponents(mSelectedPackageRow, SelectedVersionLabel())
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
    Dim currentInv As Variant

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
        currentInv = mComponents(r, 9)

        mLstComponents.AddItem NzText(mComponents(r, 2))
        idx = mLstComponents.ListCount - 1
        mLstComponents.List(idx, 1) = NzText(mComponents(r, 4))
        mLstComponents.List(idx, 2) = FormatQuantityText(perBoxQty)
        mLstComponents.List(idx, 3) = FormatQuantityText(requiredQty)
        If NzText(currentInv) = "" Then
            mLstComponents.List(idx, 4) = "unknown"
        Else
            mLstComponents.List(idx, 4) = NzText(currentInv)
        End If
        mLstComponents.List(idx, 5) = NzText(mComponents(r, 6))
        mLstComponents.List(idx, 6) = NzText(mComponents(r, 7))
        mLstComponents.List(idx, 7) = NzText(mComponents(r, 3))
        mLstComponents.List(idx, 8) = NzText(mComponents(r, 8))
    Next r
    ShowStatus "Loaded " & CStr(mLstComponents.ListCount) & " component row(s) for " & SelectedVersionLabel() & "."
    Exit Sub

FailSoft:
    ShowStatus "Component render failed: " & Err.Description
End Sub

Private Sub RenderPackageInventory()
    On Error GoTo FailSoft

    Dim currentInv As Variant
    Dim displayInv As String

    If mSelectedPackageRow <= 0 Then
        mLblPackageInv.Caption = ""
        Exit Sub
    End If
    currentInv = modTS_Shipments.BoxMakerFormCurrentInventory(mSelectedPackageRow, CStr(mTxtBoxName.Value))
    displayInv = DisplayShippableInventoryText(mSelectedPackageRow, NzText(currentInv))
    If displayInv = "" Then
        mLblPackageInv.Caption = "Current inventory for shippable box: unknown"
    Else
        mLblPackageInv.Caption = "Current inventory for shippable box: " & displayInv
    End If
    Exit Sub

FailSoft:
    mLblPackageInv.Caption = ""
End Sub

Private Sub mCboBoxes_Change()
    If mLoading Then Exit Sub
    LoadSelectedBox
End Sub

Private Sub mCboVersions_Change()
    If mLoading Then Exit Sub
    LoadSelectedVersionComponents
End Sub

Private Sub mTxtQty_Change()
    RenderComponents
End Sub

Private Sub mTxtBoxPicker_Change()
    If mLoading Then Exit Sub
    RenderShippableInventory
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

    RenderShippableInventory
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
    Dim r As Long

    If mLoading Then Exit Sub
    If mLstShippables.ListIndex < 0 Then Exit Sub

    packageRow = CLng(Val(NzText(mLstShippables.List(mLstShippables.ListIndex, 4))))
    If packageRow <= 0 Then Exit Sub

    For r = 0 To mCboBoxes.ListCount - 1
        If CLng(Val(NzText(mCboBoxes.List(r, 0)))) = packageRow Then
            mCboBoxes.ListIndex = r
            Exit Sub
        End If
    Next r
    Exit Sub

FailSoft:
End Sub

Private Sub PostBoxMakerAction(ByVal actionText As String)
    On Error GoTo ErrHandler

    Dim qtyMade As Double
    Dim resultMessage As String
    Dim startedAt As Single
    Dim elapsedMs As Long
    Dim postedOk As Boolean
    Dim previousPointer As Long
    Dim previousScreenUpdating As Boolean
    Dim quietStarted As Boolean

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

    postedOk = modTS_Shipments.CommitBoxMakerFormAction(mSelectedPackageRow, _
                                                        CStr(mTxtBoxName.Value), _
                                                        CStr(mTxtUom.Value), _
                                                        CStr(mTxtLocation.Value), _
                                                        CStr(mTxtDescription.Value), _
                                                        SelectedVersionLabel(), _
                                                        qtyMade, _
                                                        mComponents, _
                                                        resultMessage, _
                                                        actionText)
    elapsedMs = ElapsedMillisecondsForm(startedAt)
    If quietStarted Then
        modUiQuiet.EndQuietUi
        quietStarted = False
    End If
    Application.ScreenUpdating = previousScreenUpdating
    Me.MousePointer = previousPointer

    If postedOk Then
        RecordPendingShippableInventory actionText, qtyMade
        resultMessage = AppendCompletionTiming(resultMessage, elapsedMs)
        MsgBox resultMessage, vbInformation
        ShowStatus "Completed in " & Format$(elapsedMs, "#,##0") & " ms."
        RenderPackageInventory
        RenderShippableInventory
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
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long
    Dim idx As Long
    Dim currentInv As String
    Dim rowValue As Long
    Dim prevLoading As Boolean
    Dim filterText As String
    Dim shownCount As Long

    If mLstShippables Is Nothing Then Exit Sub
    prevLoading = mLoading
    mLoading = True
    mLstShippables.Clear
    If IsEmpty(mSavedBoxes) Then GoTo CleanExit

    rowsData = modTS_Shipments.BoxMakerFormLoadShippableInventory(mSavedBoxes)
    If IsEmpty(rowsData) Then GoTo CleanExit

    filterText = BoxPickerText()
    For r = 1 To UBound(rowsData, 1)
        If Not ShippableRowMatchesPicker(rowsData, r, filterText) Then GoTo NextShippable
        rowValue = CLng(Val(NzText(rowsData(r, 1))))
        currentInv = DisplayShippableInventoryText(rowValue, NzText(rowsData(r, 3)))
        If currentInv = "" Then currentInv = "unknown"

        mLstShippables.AddItem NzText(rowsData(r, 2))
        idx = mLstShippables.ListCount - 1
        mLstShippables.List(idx, 1) = currentInv
        mLstShippables.List(idx, 2) = NzText(rowsData(r, 4))
        mLstShippables.List(idx, 3) = NzText(rowsData(r, 5))
        mLstShippables.List(idx, 4) = NzText(rowsData(r, 1))
        If CLng(Val(NzText(rowsData(r, 1)))) = mSelectedPackageRow Then mLstShippables.ListIndex = idx
        shownCount = shownCount + 1
NextShippable:
    Next r
    If filterText <> "" And shownCount = 0 Then ShowStatus "No shippable boxes match picker."

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
            If CLng(Val(NzText(mLstShippables.List(r, 4)))) = mSelectedPackageRow Then
                currentText = NzText(mLstShippables.List(r, 1))
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
    If Trim$(backendText) <> "" And IsNumeric(backendText) Then
        backendQty = CDbl(backendText)
        If Abs(backendQty - pendingQty) < 0.0000001 Then
            mPendingShippableInv.Remove key
            DisplayShippableInventoryText = FormatQuantityText(backendQty)
            Exit Function
        End If
    End If

    DisplayShippableInventoryText = FormatQuantityText(pendingQty)

CleanExit:
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
                                Or TextContainsPicker(rowsData(rowIndex, 4), filterText) _
                                Or TextContainsPicker(rowsData(rowIndex, 5), filterText)
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
        mLblSyncState.Caption = "Sync: pending (" & CStr(pendingCount) & " shippable row(s))"
        mLblSyncState.ForeColor = &H80&
    Else
        mLblSyncState.Caption = "Sync: inventory synced"
        mLblSyncState.ForeColor = &H8000&
    End If

CleanExit:
End Sub

Private Function PendingShippableInventoryCount() As Long
    On Error GoTo CleanExit

    If mPendingShippableInv Is Nothing Then Exit Function
    PendingShippableInventoryCount = mPendingShippableInv.Count

CleanExit:
End Function

Private Sub SelectShippableInventoryRow()
    On Error GoTo FailSoft

    Dim r As Long

    If mLstShippables Is Nothing Then Exit Sub
    If mSelectedPackageRow <= 0 Then Exit Sub

    For r = 0 To mLstShippables.ListCount - 1
        If CLng(Val(NzText(mLstShippables.List(r, 4)))) = mSelectedPackageRow Then
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
    mLblPackageInv.Caption = ""
    UpdateSyncStateLabel
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
    AddHeaderLabel "hdrItem", "Item", leftPos, topPos, 126
    AddHeaderLabel "hdrRow", "ROW", leftPos + 128, topPos, 40
    AddHeaderLabel "hdrPerBox", "Per Box", leftPos + 170, topPos, 48
    AddHeaderLabel "hdrRequired", "Required", leftPos + 222, topPos, 56
    AddHeaderLabel "hdrCurrent", "Current Inv", leftPos + 282, topPos, 58
    AddHeaderLabel "hdrUom", "UOM", leftPos + 342, topPos, 38
    AddHeaderLabel "hdrLocation", "Location", leftPos + 382, topPos, 74
    AddHeaderLabel "hdrCode", "Code", leftPos + 460, topPos, 60
    AddHeaderLabel "hdrDesc", "Description", leftPos + 524, topPos, 150
End Sub

Private Sub AddShippableHeaders(ByVal leftPos As Single, ByVal topPos As Single)
    AddHeaderLabel "hdrShipBox", "Box", leftPos, topPos, 220
    AddHeaderLabel "hdrShipCurrent", "Current Inv", leftPos + 222, topPos, 68
    AddHeaderLabel "hdrShipUom", "UOM", leftPos + 294, topPos, 42
    AddHeaderLabel "hdrShipLocation", "Location", leftPos + 340, topPos, 118
    AddHeaderLabel "hdrShipRow", "ROW", leftPos + 462, topPos, 56
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
