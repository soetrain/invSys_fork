Attribute VB_Name = "modTS_Shipments"
Option Explicit

' =============================================================
' Module: modTS_Shipments
' Purpose: All logic for the ShippingTally system (box builder,
'          holding subsystem, confirm/build/ship macros, logging).
' Notes:
'   - Buttons are generated dynamically (similar to modTS_Received).
'   - Shipping BOM authority is stored in the warehouse runtime workbook.
'   - Operator workbooks carry ShippingBOMView as a read/view projection.
'   - Hold subsystem keeps packages on NotShipped until released.
'   - Additional confirm/build/ship routines will be implemented in
'     subsequent iterations (placeholders provided below).
' =============================================================

' ===== constants =====
Private Const SHEET_SHIPMENTS As String = "ShipmentsTally"
Private Const SHEET_INV As String = "InventoryManagement"
Private Const SHEET_BOM As String = "ShippingBOM"
Private Const SHEET_BOM_TABLES As String = "ShippingBOMTables"

Private Const TABLE_SHIPMENTS As String = "ShipmentsTally"
Private Const TABLE_NOTSHIPPED As String = "NotShipped"
Private Const TABLE_AGG_BOM As String = "AggregateBoxBOM"
Private Const TABLE_AGG_PACK As String = "AggregatePackages"
Private Const TABLE_BOX_BUILDER As String = "BoxBuilder"
Private Const TABLE_BOX_BOM As String = "BoxBOM"
Private Const TABLE_CHECK_INV As String = "Check_invSys"
Private Const TABLE_SHIPPING_BOM_VIEW As String = "ShippingBOMView"
Private Const TABLE_CANONICAL_SHIPPING_BOM As String = "tblShippingBOM"
Private Const COL_BOXBOM_ITEM As String = "ITEM"
Private Const COL_CURRENT_INV As String = "CURRENT INV"
Private Const EVENT_TYPE_BOX_BUILD As String = "BOX_BUILD"

Private Const BTN_TOGGLE_BUILDER As String = "BTN_TOGGLE_BUILDER"
Private Const BTN_SAVE_BOX As String = "BTN_SAVE_BOX"
Private Const BTN_SWITCH_BOXMAKER As String = "BTN_SWITCH_BOXMAKER"
Private Const BTN_BOX_CREATED As String = "BTN_BOX_CREATED"
Private Const BTN_UNSHIP As String = "BTN_UNSHIP"
Private Const BTN_SEND_HOLD As String = "BTN_SEND_HOLD"
Private Const BTN_RETURN_HOLD As String = "BTN_RETURN_HOLD"
Private Const BTN_CONFIRM_INV As String = "BTN_CONFIRM_INV"
Private Const BTN_BOXES_MADE As String = "BTN_BOXES_MADE"
Private Const BTN_TO_TOTALINV As String = "BTN_TO_TOTALINV"
Private Const BTN_TO_SHIPMENTS As String = "BTN_TO_SHIPMENTS"
Private Const BTN_SHIPMENTS_SENT As String = "BTN_SHIPMENTS_SENT"
Private Const CHK_USE_EXISTING As String = "CHK_USE_EXISTING"
Private Const SHAPE_TYPE_FORM_CONTROL As Long = 8

Private Const SHIPPING_BOM_BLOCK_ROWS As Long = 52
Private Const SHIPPING_BOM_DATA_ROWS As Long = 50
Private Const SHIPPING_BOM_COLS As Long = 3 ' ROW, QUANTITY, UOM
Private Const SHIPMENTS_SENT_DEDUCTS_TOTALINV As Boolean = False
Private Const SHIP_LAYOUT_BUILDER_ADDR As String = "A3"
Private Const SHIP_LAYOUT_BOM_ADDR As String = "A7"
Private Const SHIP_LAYOUT_SHIPMENTS_ADDR As String = "H3"
Private Const SHIP_LAYOUT_NOTSHIPPED_ADDR As String = "P3"
Private Const SHIP_LAYOUT_AGG_BOM_ADDR As String = "X3"
Private Const SHIP_LAYOUT_AGG_PACK_ADDR As String = "AE3"
Private Const SHIP_LAYOUT_CHECK_ADDR As String = "AL3"
Private Const SHIP_LAYOUT_INV_ADDR As String = "AV3"
Private Const SHIP_LAYOUT_BOM_VIEW_ADDR As String = "BO3"
Private Const SHIP_LAYOUT_GAP_COLUMNS As Long = 1
Private Const SHIP_LAYOUT_GAP_ROWS As Long = 2

Private mDynSearch As Object
Private mNextInvSysRow As Long
Private mAggDirty As Boolean

' ===== public entry points =====
Public Sub InitializeShipmentsUI()
    InitializeShipmentsUiForWorkbook Application.ActiveWorkbook
End Sub

Public Sub InitializeShipmentsUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    Dim surfaceReport As String
    Dim wb As Workbook

    Set wb = ResolveShippingWorkbook(targetWb, SHEET_SHIPMENTS)
    If wb Is Nothing Then Set wb = ThisWorkbook

    Call modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wb, surfaceReport)
    ArrangeShippingSurface wb
    NormalizeShippingBootstrapArtifacts wb
    EnsureShipmentsButtons wb
    EnsureBuilderTablesReady wb
    ArrangeShippingSurface wb
    RefreshShippingBomViewForWorkbook wb, surfaceReport
    modOperatorReadModel.InitializeAutoSnapshotForWorkbook wb
    If mAggDirty Then RebuildShippingAggregates
End Sub

Public Sub BtnSwitchToBoxMaker()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    SetBoxMakerMode ws, Not IsBoxMakerMode(ws)
    InvalidateShippingRibbonLabels
    If IsBoxMakerMode(ws) Then
        ShowShippingStatus "BoxMaker mode ready. Enter box quantity in BoxBuilder and aggregate component usage in BoxBOM."
    Else
        ShowShippingStatus "BoxBuilder mode ready. Define the shippable and save its BOM."
    End If
End Sub

Public Sub BtnBoxCreated()
    On Error GoTo ErrHandler
    Dim stepName As String

    stepName = "require SHIP_POST"
    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then Exit Sub

    stepName = "resolve ShipmentsTally sheet"
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    stepName = "switch to BoxMaker mode"
    SetBoxMakerMode ws, True

    stepName = "resolve BoxMaker tables"
    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    Dim loBuilder As ListObject: Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If invLo Is Nothing Or loBuilder Is Nothing Or loBom Is Nothing Then
        MsgBox "BoxMaker requires invSys, BoxBuilder, and BoxBOM tables.", vbExclamation
        Exit Sub
    End If

    stepName = "recalculate BoxBOM from BoxBuilder"
    RecalculateBoxMakerBomFromBuilder ws, invLo, loBuilder

    Dim errNotes As String
    Dim usedTotal As Double
    Dim madeTotal As Double
    stepName = "apply BoxMaker inventory deltas"
    If Not ApplyBoxCreatedFromBuilder(loBuilder, loBom, invLo, usedTotal, madeTotal, errNotes) Then
        If errNotes = "" Then errNotes = "Box creation could not be posted."
        MsgBox errNotes, vbExclamation
        Exit Sub
    End If

    stepName = "reset BoxMaker quantities"
    ResetBoxMakerQuantities loBuilder, loBom
    stepName = "refresh BoxMaker current inventory"
    RefreshBoxMakerCurrentInventory ws
    stepName = "invalidate aggregates"
    InvalidateAggregates True, True

    stepName = "show Box Created status"
    Dim msg As String
    msg = "Box created. Used " & Format$(usedTotal, "0.###") & " component units; added " & Format$(madeTotal, "0.###") & " shippable units to TOTAL INV."
    If errNotes <> "" Then msg = msg & vbCrLf & vbCrLf & errNotes
    ShowShippingStatus msg
    Exit Sub

ErrHandler:
    MsgBox "BTN_BOX_CREATED failed during " & stepName & ": " & Err.Description, vbCritical
End Sub

Public Sub BtnBoxUnboxed()
    On Error GoTo ErrHandler

    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then Exit Sub

    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    SetBoxMakerMode ws, True

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    Dim loBuilder As ListObject: Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If invLo Is Nothing Or loBuilder Is Nothing Or loBom Is Nothing Then
        MsgBox "Box Unboxed requires invSys, BoxBuilder, and BoxBOM tables.", vbExclamation
        Exit Sub
    End If

    RecalculateBoxMakerBomFromBuilder ws, invLo, loBuilder

    Dim packageReturned As Double
    Dim componentsReturned As Double
    Dim errNotes As String
    If Not ApplyBoxUnboxedFromBuilder(loBuilder, loBom, invLo, packageReturned, componentsReturned, errNotes) Then
        If errNotes = "" Then errNotes = "Box could not be unboxed."
        MsgBox errNotes, vbExclamation
        Exit Sub
    End If

    ResetBoxMakerQuantities loBuilder, loBom
    RefreshBoxMakerCurrentInventory ws
    InvalidateAggregates True, True

    ShowShippingStatus "Box unboxed. Removed " & Format$(packageReturned, "0.###") & " shippable units; returned " & Format$(componentsReturned, "0.###") & " component units to TOTAL INV."
    Exit Sub

ErrHandler:
    MsgBox "BTN_BOX_UNBOXED failed: " & Err.Description, vbCritical
End Sub

Public Sub RibbonBoxMakerModeGetLabel(control As IRibbonControl, ByRef returnedVal)
    Dim ws As Worksheet

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If IsBoxMakerMode(ws) Then
        returnedVal = "BoxMaker Mode"
    Else
        returnedVal = "BoxBuilder Mode"
    End If
End Sub

Private Sub NormalizeShippingBootstrapArtifacts(ByVal wb As Workbook)
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    NormalizeShippingArtifactTable GetListObject(ws, TABLE_SHIPMENTS)
    NormalizeShippingArtifactTable GetListObject(ws, TABLE_NOTSHIPPED)
End Sub

Private Sub NormalizeShippingArtifactTable(ByVal lo As ListObject)
    Dim cRef As Long
    Dim cItems As Long
    Dim cQty As Long
    Dim cRow As Long
    Dim i As Long
    Dim itemText As String
    Dim qtyText As String
    Dim rowText As String

    If lo Is Nothing Then Exit Sub
    RemoveAutoGeneratedColumnsShipping lo
    If lo.DataBodyRange Is Nothing Then Exit Sub

    cRef = ColumnIndex(lo, "REF_NUMBER")
    cItems = ColumnIndex(lo, "ITEMS")
    cQty = ColumnIndex(lo, "QUANTITY")
    cRow = ColumnIndex(lo, "ROW")
    If cItems = 0 Or cQty = 0 Or cRow = 0 Then Exit Sub

    For i = lo.ListRows.Count To 1 Step -1
        itemText = UCase$(Trim$(NzStr(lo.DataBodyRange.Cells(i, cItems).Value)))
        qtyText = UCase$(Trim$(NzStr(lo.DataBodyRange.Cells(i, cQty).Value)))
        rowText = UCase$(Trim$(NzStr(lo.DataBodyRange.Cells(i, cRow).Value)))

        If cRef > 0 Then
            If Trim$(NzStr(lo.DataBodyRange.Cells(i, cRef).Value)) <> "" Then GoTo NextRow
        End If

        If (itemText = "ITEM" Or itemText = "ITEMS") _
            And (qtyText = "ROW" Or qtyText = "QUANTITY") _
            And (rowText = "ROW" Or rowText = "QUANTITY") Then
            lo.ListRows(i).Delete
        End If
NextRow:
    Next i
End Sub

Private Sub RemoveAutoGeneratedColumnsShipping(ByVal lo As ListObject)
    Dim i As Long

    If lo Is Nothing Then Exit Sub
    For i = lo.ListColumns.Count To 1 Step -1
        If LCase$(Left$(Trim$(lo.ListColumns(i).Name), 6)) = "column" Then
            lo.ListColumns(i).Delete
        End If
    Next i
End Sub

Private Function ResolveShippingWorkbook(Optional ByVal preferredWb As Workbook = Nothing, Optional ByVal requiredSheet As String = "") As Workbook
    If Not preferredWb Is Nothing Then
        Set ResolveShippingWorkbook = preferredWb
        Exit Function
    End If

    If Not Application.ActiveWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook.IsAddin Then
            If requiredSheet = "" Then
                Set ResolveShippingWorkbook = Application.ActiveWorkbook
                Exit Function
            ElseIf Not WorkbookSheetExistsShipping(Application.ActiveWorkbook, requiredSheet) Is Nothing Then
                Set ResolveShippingWorkbook = Application.ActiveWorkbook
                Exit Function
            End If
        End If
    End If

    If requiredSheet = "" Then
        Set ResolveShippingWorkbook = ThisWorkbook
    ElseIf Not WorkbookSheetExistsShipping(ThisWorkbook, requiredSheet) Is Nothing Then
        Set ResolveShippingWorkbook = ThisWorkbook
    End If
End Function

Private Function WorkbookSheetExistsShipping(ByVal wb As Workbook, ByVal nameOrCode As String) As Worksheet
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set WorkbookSheetExistsShipping = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub ArrangeShippingSurface(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim loBuilder As ListObject
    Dim loBom As ListObject
    Dim lo As ListObject
    Dim anchorRow As Long
    Dim anchorCol As Long
    Dim nextCol As Long
    Dim leftBandRight As Long

    If wb Is Nothing Then Exit Sub
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    anchorRow = ws.Range(SHIP_LAYOUT_BUILDER_ADDR).Row
    anchorCol = ws.Range(SHIP_LAYOUT_BUILDER_ADDR).Column

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)

    MoveListObjectToRowColShipping loBuilder, anchorRow, anchorCol
    ArrangeBoxBuilderBandShipping loBuilder, loBom

    leftBandRight = MaxLongShipping(ListObjectRightColumnShipping(loBuilder), ListObjectRightColumnShipping(loBom))
    If leftBandRight < anchorCol Then leftBandRight = anchorCol
    nextCol = leftBandRight + SHIP_LAYOUT_GAP_COLUMNS + 1

    Set lo = GetListObject(ws, TABLE_SHIPMENTS)
    nextCol = MoveListObjectAndNextColumnShipping(lo, anchorRow, nextCol)

    Set lo = GetListObject(ws, TABLE_NOTSHIPPED)
    nextCol = MoveListObjectAndNextColumnShipping(lo, anchorRow, nextCol)

    Set lo = GetListObject(ws, TABLE_AGG_BOM)
    nextCol = MoveListObjectAndNextColumnShipping(lo, anchorRow, nextCol)

    Set lo = GetListObject(ws, TABLE_AGG_PACK)
    nextCol = MoveListObjectAndNextColumnShipping(lo, anchorRow, nextCol)

    Set lo = GetListObject(ws, TABLE_CHECK_INV)
    nextCol = MoveListObjectAndNextColumnShipping(lo, anchorRow, nextCol)

    Set lo = GetListObject(ws, "invSysData_Shipping")
    nextCol = MoveListObjectAndNextColumnShipping(lo, anchorRow, nextCol)

    Set lo = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    nextCol = MoveListObjectAndNextColumnShipping(lo, anchorRow, nextCol)
End Sub

Private Sub MoveListObjectToAddressShipping(ByVal lo As ListObject, ByVal addressText As String)
    If lo Is Nothing Then Exit Sub
    MoveListObjectToRowColShipping lo, lo.Parent.Range(addressText).Row, lo.Parent.Range(addressText).Column
End Sub

Private Sub MoveListObjectToRowColShipping(ByVal lo As ListObject, ByVal targetRow As Long, ByVal targetCol As Long)
    Dim dest As Range
    Dim previousAlerts As Boolean

    If lo Is Nothing Then Exit Sub
    If targetRow < 1 Or targetCol < 1 Then Exit Sub
    If lo.Range.Row = targetRow And lo.Range.Column = targetCol Then Exit Sub

    Set dest = lo.Parent.Cells(targetRow, targetCol)

    On Error Resume Next
    ClearExcelClipboardStateShipping
    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    lo.Range.Cut Destination:=dest
    ClearExcelClipboardStateShipping
    Application.DisplayAlerts = previousAlerts
    Err.Clear
    On Error GoTo 0
End Sub

Private Function MoveListObjectAndNextColumnShipping(ByVal lo As ListObject, ByVal targetRow As Long, ByVal targetCol As Long) As Long
    Dim rightCol As Long

    If lo Is Nothing Then
        MoveListObjectAndNextColumnShipping = targetCol
        Exit Function
    End If

    MoveListObjectToRowColShipping lo, targetRow, targetCol
    rightCol = ListObjectRightColumnShipping(lo)
    If rightCol <= 0 Then
        MoveListObjectAndNextColumnShipping = targetCol
    Else
        MoveListObjectAndNextColumnShipping = rightCol + SHIP_LAYOUT_GAP_COLUMNS + 1
    End If
End Function

Private Function ListObjectRightColumnShipping(ByVal lo As ListObject) As Long
    Dim tableRange As Range

    On Error GoTo CleanFail
    If lo Is Nothing Then Exit Function

    Set tableRange = lo.Range
    If tableRange Is Nothing Then Set tableRange = lo.HeaderRowRange
    If tableRange Is Nothing Then Exit Function

    ListObjectRightColumnShipping = tableRange.Column + tableRange.Columns.Count - 1
    Exit Function

CleanFail:
    ListObjectRightColumnShipping = 0
End Function

Private Function ListObjectsOverlapRowsShipping(ByVal leftTable As ListObject, ByVal rightTable As ListObject) As Boolean
    Dim leftRange As Range
    Dim rightRange As Range
    Dim leftBottom As Long
    Dim rightBottom As Long

    On Error GoTo CleanFail
    If leftTable Is Nothing Then Exit Function
    If rightTable Is Nothing Then Exit Function

    Set leftRange = leftTable.Range
    Set rightRange = rightTable.Range
    If leftRange Is Nothing Or rightRange Is Nothing Then Exit Function

    leftBottom = leftRange.Row + leftRange.Rows.Count - 1
    rightBottom = rightRange.Row + rightRange.Rows.Count - 1
    ListObjectsOverlapRowsShipping = (leftRange.Row <= rightBottom And rightRange.Row <= leftBottom)
    Exit Function

CleanFail:
    ListObjectsOverlapRowsShipping = False
End Function

Private Sub EnsureListObjectColumnInsertSpaceShipping(ByVal lo As ListObject, ByVal extraColumns As Long)
    Dim ws As Worksheet
    Dim baseRight As Long
    Dim shiftBy As Long
    Dim maxLeft As Long
    Dim scanCol As Long
    Dim other As ListObject
    Dim otherLeft As Long

    On Error GoTo CleanExit
    If lo Is Nothing Then Exit Sub
    If extraColumns <= 0 Then Exit Sub

    Set ws = lo.Parent
    If ws Is Nothing Then Exit Sub

    baseRight = ListObjectRightColumnShipping(lo)
    If baseRight <= 0 Then Exit Sub

    shiftBy = extraColumns + SHIP_LAYOUT_GAP_COLUMNS

    For Each other In ws.ListObjects
        If Not other Is lo Then
            otherLeft = 0
            On Error Resume Next
            otherLeft = other.Range.Column
            On Error GoTo CleanExit
            If otherLeft > maxLeft Then maxLeft = otherLeft
        End If
    Next other

    For scanCol = maxLeft To baseRight + 1 Step -1
        For Each other In ws.ListObjects
            If Not other Is lo Then
                otherLeft = 0
                On Error Resume Next
                otherLeft = other.Range.Column
                On Error GoTo CleanExit
                If otherLeft = scanCol Then
                    If ListObjectsOverlapRowsShipping(lo, other) Then
                        MoveListObjectToRowColShipping other, other.Range.Row, other.Range.Column + shiftBy
                    End If
                End If
            End If
        Next other
    Next scanCol

CleanExit:
End Sub

Private Function ListObjectsOverlapColumnsShipping(ByVal leftTable As ListObject, ByVal rightTable As ListObject) As Boolean
    Dim leftRange As Range
    Dim rightRange As Range
    Dim leftRight As Long
    Dim rightRight As Long

    On Error GoTo CleanFail
    If leftTable Is Nothing Then Exit Function
    If rightTable Is Nothing Then Exit Function

    Set leftRange = leftTable.Range
    Set rightRange = rightTable.Range
    If leftRange Is Nothing Or rightRange Is Nothing Then Exit Function

    leftRight = leftRange.Column + leftRange.Columns.Count - 1
    rightRight = rightRange.Column + rightRange.Columns.Count - 1
    ListObjectsOverlapColumnsShipping = (leftRange.Column <= rightRight And rightRange.Column <= leftRight)
    Exit Function

CleanFail:
    ListObjectsOverlapColumnsShipping = False
End Function

Private Sub MoveListObjectColumnAddBlockersShipping(ByVal lo As ListObject, ByVal extraColumns As Long)
    Dim ws As Worksheet
    Dim other As ListObject
    Dim maxRight As Long
    Dim rightCol As Long
    Dim nextCol As Long

    On Error GoTo CleanExit
    If lo Is Nothing Then Exit Sub
    If extraColumns <= 0 Then Exit Sub

    Set ws = lo.Parent
    If ws Is Nothing Then Exit Sub

    For Each other In ws.ListObjects
        rightCol = ListObjectRightColumnShipping(other)
        If rightCol > maxRight Then maxRight = rightCol
    Next other

    nextCol = maxRight + extraColumns + SHIP_LAYOUT_GAP_COLUMNS + 1
    If nextCol < 1 Then Exit Sub

    For Each other In ws.ListObjects
        If Not other Is lo Then
            If ListObjectsOverlapRowsShipping(lo, other) _
               Or ListObjectsOverlapColumnsShipping(lo, other) Then
                MoveListObjectToRowColShipping other, other.Range.Row, nextCol
                rightCol = ListObjectRightColumnShipping(other)
                If rightCol > 0 Then
                    nextCol = rightCol + SHIP_LAYOUT_GAP_COLUMNS + 1
                Else
                    nextCol = nextCol + extraColumns + SHIP_LAYOUT_GAP_COLUMNS + 1
                End If
            End If
        End If
    Next other

CleanExit:
End Sub

Private Function MaxLongShipping(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLongShipping = leftValue
    Else
        MaxLongShipping = rightValue
    End If
End Function

Public Sub BtnToggleBuilder()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = GetListObject(ws, TABLE_BOX_BOM)
    Dim makeVisible As Boolean
    If lo Is Nothing Then
        makeVisible = True
    Else
        Dim firstCol As Long
        firstCol = lo.HeaderRowRange.Column
        makeVisible = ws.Columns(firstCol).EntireColumn.Hidden
    End If
    ToggleBuilderTables makeVisible
End Sub

Public Sub BtnSaveBox()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim loMeta As ListObject: Set loMeta = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loMeta Is Nothing Or loBom Is Nothing Then
        MsgBox "Box Builder tables not found on ShipmentsTally sheet.", vbExclamation
        Exit Sub
    End If

    EnsureTableHasRow loMeta
    RemoveColumnIfExistsShipping loMeta, "ROW"
    EnsureBoxBomEntryColumns loBom

    Dim boxName As String
    boxName = Trim$(NzStr(ValueFromTable(loMeta, "Box Name")))
    If boxName = "" Then
        MsgBox "Enter a Box Name before saving.", vbExclamation
        Exit Sub
    End If
    Dim boxUOM As String: boxUOM = Trim$(NzStr(ValueFromTable(loMeta, "UOM")))
    Dim boxLoc As String: boxLoc = Trim$(NzStr(ValueFromTable(loMeta, "LOCATION")))
    Dim boxDesc As String: boxDesc = Trim$(NzStr(ValueFromTable(loMeta, "DESCRIPTION")))
    If boxUOM = "" Then
        MsgBox "Box Builder UOM is required.", vbExclamation
        Exit Sub
    End If

    EnsureTableHasRow loBom

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim components As Collection
    Dim syncNotes As String
    Set components = CollectBomComponents(loBom, invLo, syncNotes)
    If components.count = 0 Then
        MsgBox "Add at least one valid component row (ROW/QUANTITY) to the BoxBOM table.", vbExclamation
        Exit Sub
    End If
    If components.count > SHIPPING_BOM_DATA_ROWS Then
        MsgBox "BOM exceeds the 50-row limit. Remove extra rows and try again.", vbExclamation
        Exit Sub
    End If

    Dim boxRowValue As Long
    boxRowValue = ResolveBoxPackageRowValue(ws.Parent, boxName, invLo)
    If boxRowValue = 0 Then Exit Sub
    boxRowValue = EnsureInvSysItem(boxName, boxUOM, boxLoc, boxDesc, invLo, boxRowValue)
    If boxRowValue = 0 Then Exit Sub

    Dim bomReport As String
    If Not SaveShippingBomToRuntime(ws.Parent, boxRowValue, boxName, boxUOM, boxLoc, boxDesc, components, bomReport) Then
        If bomReport = "" Then bomReport = "Unable to save Shipping BOM to the selected warehouse runtime."
        MsgBox bomReport, vbCritical
        Exit Sub
    End If
    RefreshShippingBomViewForWorkbook ws.Parent, bomReport

    Dim finalMsg As String
    finalMsg = "Saved BOM '" & boxName & "' to warehouse runtime (invSys ROW " & boxRowValue & ", " & components.count & " components)."
    If Len(syncNotes) > 0 Then
        finalMsg = finalMsg & vbCrLf & syncNotes
    End If
    If Len(bomReport) > 0 Then
        finalMsg = finalMsg & vbCrLf & bomReport
    End If
    MsgBox finalMsg, vbInformation

    ClearListObjectData loMeta
    ClearListObjectData loBom
    EnsureTableHasRow loMeta
    EnsureTableHasRow loBom
    InvalidateAggregates True
    Exit Sub

ErrHandler:
    MsgBox "BTN_SAVE_BOX failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnUnship()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = GetListObject(ws, TABLE_NOTSHIPPED)
    If lo Is Nothing Then
        MsgBox "NotShipped table not found.", vbExclamation
        Exit Sub
    End If
    Dim isHidden As Boolean
    isHidden = lo.Range.EntireColumn.Hidden
    lo.Range.EntireColumn.Hidden = Not isHidden
End Sub

Public Sub BtnSendHold()
    MoveSelectionToHold True
End Sub

Public Sub BtnReturnHold()
    MoveSelectionToHold False
End Sub

Public Function MoveShipmentHoldForAutomation(ByVal refNumber As String, _
                                              ByVal itemText As String, _
                                              ByVal qtyMove As Double, _
                                              ByVal moveToHold As Boolean) As String
    On Error GoTo Fail

    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        MoveShipmentHoldForAutomation = "ERR|ShipmentsTally sheet not found."
        Exit Function
    End If

    Dim loShip As ListObject: Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    Dim loHold As ListObject: Set loHold = GetListObject(ws, TABLE_NOTSHIPPED)
    If loShip Is Nothing Or loHold Is Nothing Then
        MoveShipmentHoldForAutomation = "ERR|ShipmentsTally or NotShipped table not found."
        Exit Function
    End If

    Dim sourceTable As ListObject
    Dim targetTable As ListObject
    If moveToHold Then
        Set sourceTable = loShip
        Set targetTable = loHold
    Else
        Set sourceTable = loHold
        Set targetTable = loShip
    End If

    If qtyMove <= 0 Then
        MoveShipmentHoldForAutomation = "ERR|Quantity must be positive."
        Exit Function
    End If

    Dim rowIndex As Long
    rowIndex = FindHoldRowIndex(sourceTable, refNumber, itemText)
    If rowIndex <= 0 Then
        MoveShipmentHoldForAutomation = "ERR|Source row not found."
        Exit Function
    End If

    Dim availableQty As Double
    availableQty = HoldRowQty(sourceTable, rowIndex)
    If availableQty <= 0 Then
        MoveShipmentHoldForAutomation = "ERR|Source row quantity is empty."
        Exit Function
    End If
    If qtyMove > availableQty Then qtyMove = availableQty

    MoveHoldRowQuantity sourceTable, targetTable, rowIndex, qtyMove
    InvalidateAggregates True

    MoveShipmentHoldForAutomation = "OK|Moved=" & CStr(qtyMove) _
        & "|SourceQty=" & CStr(HoldRowQtyByKey(sourceTable, refNumber, itemText)) _
        & "|TargetQty=" & CStr(HoldRowQtyByKey(targetTable, refNumber, itemText))
    Exit Function

Fail:
    MoveShipmentHoldForAutomation = "ERR|" & Err.Description
End Function

Public Sub BtnConfirmInventory()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    If UseExistingInventoryEnabled(ws) Then
        MsgBox "Use existing inventory is enabled. Skip Confirm inventory and go to 'To Shipments'.", vbInformation
        Exit Sub
    End If

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    Dim aggBom As ListObject: Set aggBom = GetListObject(ws, TABLE_AGG_BOM)

    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If
    If aggBom Is Nothing Or aggBom.DataBodyRange Is Nothing Then
        MsgBox "AggregateBoxBOM has no rows to confirm. Enter package quantities first.", vbInformation
        Exit Sub
    End If

    Dim shortage As String
    If Not ValidateComponentInventory(invLo, aggBom, shortage) Then
        MsgBox "Cannot confirm shipments:" & vbCrLf & shortage, vbExclamation
        Exit Sub
    End If

    Dim stageLogs As New Collection
    Dim errNotes As String
    Dim stagedTotal As Double
    stagedTotal = StageComponentsToUsed(invLo, aggBom, errNotes, stageLogs)
    If stagedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unknown staging failure."
        MsgBox "BTN_CONFIRM_INV cancelled: " & errNotes, vbCritical
        Exit Sub
    End If

    If stageLogs.Count > 0 Then LogShippingChanges "AggregateBoxBOM_Log", stageLogs

    InvalidateAggregates True

    Dim msg As String
    msg = "Confirmed component demand: " & Format$(stagedTotal, "0.###") & " units staged into invSys.USED."
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        MsgBox msg, vbExclamation
    Else
        ShowShippingStatus msg
    End If
    Exit Sub
ErrHandler:
    MsgBox "BTN_CONFIRM_INV failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnBoxesMade()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    If UseExistingInventoryEnabled(ws) Then
        MsgBox "Use existing inventory is enabled. Skip Boxes made and go to 'To Shipments'.", vbInformation
        Exit Sub
    End If

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    Dim aggBom As ListObject: Set aggBom = GetListObject(ws, TABLE_AGG_BOM)
    Dim aggPack As ListObject: Set aggPack = GetListObject(ws, TABLE_AGG_PACK)

    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If
    If aggBom Is Nothing Or aggBom.DataBodyRange Is Nothing Then
        MsgBox "AggregateBoxBOM has no rows. Enter package quantities in ShipmentsTally first.", vbInformation
        Exit Sub
    End If

    Dim errNotes As String, shortage As String
    Dim usedTotal As Double
    Dim madeTotal As Double
    Dim compLogs As New Collection
    Dim pkgLogs As New Collection
    Dim usedDeltas As Collection
    Dim madeDeltas As Collection

    If Not ValidateComponentInventory(invLo, aggBom, shortage) Then
        MsgBox "Cannot make boxes:" & vbCrLf & shortage, vbExclamation
        Exit Sub
    End If

    Set usedDeltas = BuildUsedDeltaPacket(invLo, aggBom, errNotes)
    If usedDeltas Is Nothing Then
        MsgBox "Boxes made cancelled: " & errNotes, vbExclamation
        Exit Sub
    End If

    Set madeDeltas = BuildMadeDeltaPacket(invLo, aggPack, errNotes)
    If madeDeltas Is Nothing Then
        MsgBox "Boxes made cancelled: " & errNotes, vbExclamation
        Exit Sub
    End If

    PrepareComponentLogEntries invLo, usedDeltas, compLogs
    PreparePackageLogEntries invLo, madeDeltas, pkgLogs

    usedTotal = ApplyUsedDeltasLocal(invLo, usedDeltas, errNotes)
    If usedTotal < 0 Then
        MsgBox "Boxes made cancelled: insufficient inventory to cover all BOM components." & vbCrLf & vbCrLf & errNotes, vbExclamation
        Exit Sub
    End If

    madeTotal = ApplyMadeDeltasLocal(invLo, madeDeltas, errNotes)
    If madeTotal < 0 Then
        MsgBox "Boxes made cancelled: " & errNotes, vbExclamation
        Exit Sub
    End If

    ResetShippingStaging clearShipments:=False, clearPackages:=True
    InvalidateAggregates True, True

    If compLogs.Count > 0 Then LogShippingChanges "AggregateBoxBOM_Log", compLogs
    If pkgLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", pkgLogs

    Dim msg As String
    msg = "Recorded component usage: " & Format$(usedTotal, "0.###") & " units."
    msg = msg & vbCrLf & "Recorded finished packages (MADE): " & Format$(madeTotal, "0.###")
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        If HasActionableShippingWarning(errNotes) Then
            MsgBox msg, vbExclamation
        Else
            ShowShippingStatus msg
        End If
    Else
        ShowShippingStatus msg
    End If
    Exit Sub
ErrHandler:
    MsgBox "BTN_BOXES_MADE failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnToTotalInv()
    On Error GoTo ErrHandler
    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim errNotes As String
    Dim deltas As Collection
    Set deltas = BuildTotalInventoryDeltaPacket(invLo, errNotes)
    If deltas Is Nothing Then
        If errNotes <> "" Then
            MsgBox errNotes, vbInformation
        Else
            MsgBox "No staged packages found in invSys.MADE. Run Boxes made before sending to TotalInv.", vbInformation
        End If
        Exit Sub
    End If
    If deltas.Count = 0 Then
        If errNotes <> "" Then
            MsgBox errNotes, vbInformation
        Else
            MsgBox "No staged packages found in invSys.MADE. Run Boxes made before sending to TotalInv.", vbInformation
        End If
        Exit Sub
    End If

    Dim shipLogs As New Collection
    PrepareTotalInventoryLogEntries invLo, deltas, shipLogs

    Dim movedTotal As Double
    movedTotal = ApplyMadeToInventoryDeltasLocal(invLo, deltas, errNotes)
    If movedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to move packages into TOTAL INV."
        MsgBox errNotes, vbCritical
        Exit Sub
    End If

    ClearShipmentEntryTables
    ResetShippingStaging clearShipments:=True, clearPackages:=True
    InvalidateAggregates True
    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs

    Dim msg As String
    msg = "Moved " & Format$(movedTotal, "0.###") & " packages from MADE into TOTAL INV."
    ShowShippingStatus msg
    Exit Sub
ErrHandler:
    MsgBox "BTN_TO_TOTALINV failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnToShipments()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    Dim aggPack As ListObject: Set aggPack = GetListObject(ws, TABLE_AGG_PACK)
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If
    If aggPack Is Nothing Or aggPack.DataBodyRange Is Nothing Then
        MsgBox "AggregatePackages has no rows to stage.", vbInformation
        Exit Sub
    End If

    Dim errNotes As String
    Dim deltas As Collection
    Set deltas = BuildShipmentDeltaPacket(invLo, aggPack, errNotes)
    If deltas Is Nothing Then
        If errNotes <> "" Then
            MsgBox errNotes, vbInformation
        Else
            MsgBox "No additional shipments required; Shipments column already meets demand.", vbInformation
        End If
        Exit Sub
    End If
    If deltas.Count = 0 Then
        If errNotes <> "" Then
            MsgBox errNotes, vbInformation
        Else
            MsgBox "No additional shipments required; Shipments column already meets demand.", vbInformation
        End If
        Exit Sub
    End If

    Dim shipLogs As New Collection
    PrepareShipmentStageLogEntries invLo, deltas, shipLogs

    Dim stagedTotal As Double
    stagedTotal = ApplyShipmentDeltasLocal(invLo, deltas, errNotes)
    If stagedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to stage shipments due to inventory shortage."
        MsgBox errNotes, vbCritical
        Exit Sub
    End If

    InvalidateAggregates True
    RestoreShipmentStageColumns invLo, deltas
    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs

    Dim msg As String
    msg = "Staged " & Format$(stagedTotal, "0.###") & " packages into invSys.SHIPMENTS."
    ShowShippingStatus msg
    Exit Sub
ErrHandler:
    Dim errMsg As String
    If errNotes <> "" Then
        errMsg = errNotes
    ElseIf Err.Number = 91 Then
        errMsg = "Cannot stage shipments: requested quantity exceeds available TOTAL INV or a package row is missing in invSys."
    Else
        errMsg = "BTN_TO_SHIPMENTS failed: " & Err.Description
    End If
    MsgBox errMsg, vbCritical
End Sub

Public Sub BtnShipmentsSent()
    On Error GoTo ErrHandler
    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then Exit Sub
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim queuedEventId As String
    Dim errNotes As String
    Dim deltas As Collection
    Dim runtimeReport As String
    If Not BuildQueueableShipmentsSentDeltas(invLo, ws, deltas, errNotes) Then
        If errNotes = "" Then errNotes = "Unable to build shipment event."
        MsgBox errNotes, vbInformation
        Exit Sub
    End If

    If Not QueueShipmentsSentEvent(deltas, errNotes, queuedEventId) Then
        If errNotes = "" Then errNotes = "Unable to queue shipment event."
        MsgBox errNotes, vbCritical
        Exit Sub
    End If

    Dim shipLogs As New Collection
    PrepareShipmentsSentLogEntries invLo, deltas, shipLogs, SHIPMENTS_SENT_DEDUCTS_TOTALINV

    Dim shippedTotal As Double
    shippedTotal = ApplyShipmentsSentDeltas(invLo, deltas, errNotes, SHIPMENTS_SENT_DEDUCTS_TOTALINV)
    If shippedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to finalize shipments."
        MsgBox errNotes, vbCritical
        Exit Sub
    End If

    ClearShipmentEntryTables
    InvalidateAggregates True
    ClearInstructionStaging ws

    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs
    If Not modOperatorReadModel.RunBatchAndRefreshOperatorWorkbook(ws.Parent, "", "LOCAL", runtimeReport) Then
        If runtimeReport = "" Then runtimeReport = "Local shipment post succeeded, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If
    ClearShipmentStageAfterRefresh ws.Parent, deltas

    Dim msg As String
    msg = "Finalized " & Format$(shippedTotal, "0.###") & " shipments."
    If SHIPMENTS_SENT_DEDUCTS_TOTALINV Then
        msg = msg & vbCrLf & "TOTAL INV reduced; SHIPMENTS cleared."
    Else
        msg = msg & vbCrLf & "SHIPMENTS cleared."
    End If
    If queuedEventId <> "" Then msg = msg & vbCrLf & "Inbox EventID: " & queuedEventId
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        If HasActionableShippingWarning(errNotes) Then
            MsgBox msg, vbExclamation
        Else
            ShowShippingStatus msg
        End If
    Else
        ShowShippingStatus msg
    End If
    ClearShipmentStageAfterRefresh ws.Parent, deltas
    Exit Sub
ErrHandler:
    Dim errMsg As String
    errMsg = "BTN_SHIPMENTS_SENT failed: " & Err.Description
    If queuedEventId <> "" Then errMsg = errMsg & vbCrLf & vbCrLf & "Inbox EventID already queued: " & queuedEventId
    MsgBox errMsg, vbCritical
End Sub

Private Sub ClearShipmentStageAfterRefresh(ByVal wb As Workbook, ByVal deltas As Collection)
    Dim attempt As Long
    Dim invLo As ListObject

    If wb Is Nothing Then Exit Sub

    For attempt = 1 To 3
        AllowExcelRefreshToSettle
        Set invLo = GetInvSysTableFromWorkbook(wb)
        ClearShipmentStageColumns invLo, deltas
        ClearAllShipmentStageColumns invLo
        If attempt > 1 Then
            If Not HasAnyShipmentStage(invLo) Then Exit For
        End If
    Next attempt
End Sub

Private Sub ShowShippingStatus(ByVal messageText As String)
    On Error Resume Next
    Application.StatusBar = FlattenStatusText(messageText)
    On Error GoTo 0
End Sub

Private Function FlattenStatusText(ByVal messageText As String) As String
    Dim result As String

    result = Replace(messageText, vbCrLf, "  ")
    result = Replace(result, vbCr, "  ")
    result = Replace(result, vbLf, "  ")
    Do While InStr(result, "   ") > 0
        result = Replace(result, "   ", "  ")
    Loop
    If Len(result) > 240 Then result = Left$(result, 237) & "..."
    FlattenStatusText = result
End Function

Private Function HasActionableShippingWarning(ByVal notes As String) As Boolean
    Dim lowered As String

    lowered = LCase$(Trim$(notes))
    If lowered = "" Then Exit Function

    HasActionableShippingWarning = _
        (InStr(1, lowered, "failed", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "error", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "poison", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "cancel", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "insufficient", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "unable", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "did not complete", vbTextCompare) > 0)
End Function

Private Function HasAnyShipmentStage(ByVal invLo As ListObject) As Boolean
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    Dim colShip As Long: colShip = ColumnIndex(invLo, "SHIPMENTS")
    If colShip = 0 Then Exit Function

    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If NzDbl(invLo.DataBodyRange.Cells(r, colShip).Value) <> 0 Then
            HasAnyShipmentStage = True
            Exit Function
        End If
    Next r
End Function

Private Sub ClearAllShipmentStageColumns(ByVal invLo As ListObject)
    If invLo Is Nothing Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim colShip As Long: colShip = ColumnIndex(invLo, "SHIPMENTS")
    If colShip = 0 Then Exit Sub

    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If NzDbl(invLo.DataBodyRange.Cells(r, colShip).Value) <> 0 Then
            invLo.DataBodyRange.Cells(r, colShip).Value = 0
        End If
    Next r
End Sub

Private Sub AllowExcelRefreshToSettle()
    On Error Resume Next
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents
    On Error GoTo 0
End Sub

Public Function QueueShipmentsSentEventFromCurrentWorkbook(ByRef eventIdOut As String, ByRef errNotes As String) As Boolean
    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim deltas As Collection

    If Not modRoleUiAccess.CanCurrentUserPerformCapability("SHIP_POST", "", "", "", errNotes) Then Exit Function

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        errNotes = "ShipmentsTally sheet not found."
        Exit Function
    End If

    Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        errNotes = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    If Not BuildQueueableShipmentsSentDeltas(invLo, ws, deltas, errNotes) Then Exit Function
    QueueShipmentsSentEventFromCurrentWorkbook = QueueShipmentsSentEvent(deltas, errNotes, eventIdOut)
End Function

Public Function ValidateQueueShipmentsSentEventFromCurrentWorkbook() As String
    Dim eventIdOut As String
    Dim errNotes As String

    If QueueShipmentsSentEventFromCurrentWorkbook(eventIdOut, errNotes) Then
        ValidateQueueShipmentsSentEventFromCurrentWorkbook = "OK"
    Else
        ValidateQueueShipmentsSentEventFromCurrentWorkbook = errNotes
    End If
End Function

Private Function QueueShipmentsSentEvent(ByVal deltas As Collection, ByRef errNotes As String, ByRef eventIdOut As String) As Boolean
    Dim payloadJson As String

    payloadJson = BuildPayloadJsonFromDeltas(deltas, "")
    If payloadJson = "" Then
        If errNotes = "" Then errNotes = "No shipment payload rows were generated."
        Exit Function
    End If

    QueueShipmentsSentEvent = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_SHIP, _
        "", _
        payloadJson, _
        "BTN_SHIPMENTS_SENT", _
        eventIdOut, _
        errNotes)
End Function

Private Function QueueBoxBuildEventFromBuilder(ByVal loBuilder As ListObject, _
                                               ByVal loBom As ListObject, _
                                               ByVal invLo As ListObject, _
                                               ByRef usedTotal As Double, _
                                               ByRef madeTotal As Double, _
                                               ByRef eventIdOut As String, _
                                               ByRef errNotes As String) As Boolean
    Dim payloadItems As Collection
    Dim payloadJson As String
    Dim boxName As String
    Dim boxQty As Double
    Dim packageRow As Long
    Dim packageIdx As Long
    Dim runtimeMax As Long
    Dim itemCode As String
    Dim uomVal As String
    Dim locationVal As String
    Dim descrVal As String

    errNotes = ""
    usedTotal = 0
    madeTotal = 0
    If loBuilder Is Nothing Or loBom Is Nothing Then
        errNotes = "BoxMaker required tables are missing."
        Exit Function
    End If

    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    boxQty = NzDbl(ValueFromTable(loBuilder, "Quantity"))
    If boxName = "" Then
        errNotes = "BoxBuilder Box Name is required."
        Exit Function
    End If
    If boxQty <= 0 Then
        errNotes = "BoxBuilder Quantity must be greater than zero."
        Exit Function
    End If

    If Not invLo Is Nothing Then
        packageIdx = FindInvRowIndexByItem(invLo, boxName)
        If packageIdx > 0 Then
            packageRow = NzLng(GetInvSysValueByIndex(invLo, packageIdx, "ROW"))
            itemCode = NzStr(GetInvSysValueByIndex(invLo, packageIdx, "ITEM_CODE"))
        End If
    End If
    If packageRow <= 0 Then packageRow = FindShippingBomPackageRowByName(loBuilder.Parent.Parent, boxName, runtimeMax)
    If packageRow <= 0 Then
        errNotes = "Box '" & boxName & "' was not found in ShippingBOM runtime."
        Exit Function
    End If
    If itemCode = "" Then itemCode = boxName
    uomVal = Trim$(NzStr(ValueFromTable(loBuilder, "UOM")))
    locationVal = Trim$(NzStr(ValueFromTable(loBuilder, "LOCATION")))
    descrVal = Trim$(NzStr(ValueFromTable(loBuilder, "DESCRIPTION")))

    Set payloadItems = New Collection
    If Not AddBoxBuildComponentPayloadItems(loBom, invLo, payloadItems, usedTotal, errNotes) Then Exit Function
    AddBoxBuildPayloadItem payloadItems, packageRow, itemCode, boxName, boxQty, uomVal, locationVal, descrVal, "MADE"
    madeTotal = boxQty

    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        errNotes = "No BoxMaker payload rows were generated."
        Exit Function
    End If

    QueueBoxBuildEventFromBuilder = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_BOX_BUILD, _
        "", _
        payloadJson, _
        "BTN_BOX_CREATED", _
        eventIdOut, _
        errNotes)
End Function

Private Function AddBoxBuildComponentPayloadItems(ByVal loBom As ListObject, _
                                                  ByVal invLo As ListObject, _
                                                  ByVal payloadItems As Collection, _
                                                  ByRef usedTotal As Double, _
                                                  ByRef errNotes As String) As Boolean
    Dim cItem As Long
    Dim cCode As Long
    Dim cRow As Long
    Dim cQty As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long
    Dim r As Long
    Dim itemName As String
    Dim itemCode As String
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim uomVal As String
    Dim locVal As String
    Dim descVal As String

    If payloadItems Is Nothing Then Exit Function
    If loBom Is Nothing Or loBom.DataBodyRange Is Nothing Then
        errNotes = "BoxBOM has no component rows."
        Exit Function
    End If

    EnsureBoxBomEntryColumns loBom
    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    cCode = ColumnIndex(loBom, "ITEM_CODE")
    cRow = ColumnIndex(loBom, "ROW")
    cQty = ColumnIndex(loBom, "QUANTITY")
    cUom = ColumnIndex(loBom, "UOM")
    cLoc = ColumnIndex(loBom, "LOCATION")
    cDesc = ColumnIndex(loBom, "DESCRIPTION")
    If cItem = 0 Or cRow = 0 Or cQty = 0 Then
        errNotes = "BoxBOM must include ITEM, ROW, and QUANTITY columns."
        Exit Function
    End If

    For r = 1 To loBom.ListRows.Count
        itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))
        itemCode = ""
        If cCode > 0 Then itemCode = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cCode).Value))
        rowVal = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        qtyVal = NzDbl(loBom.DataBodyRange.Cells(r, cQty).Value)
        If BoxMakerComponentRowIsBlank(itemName, rowVal, qtyVal) Then GoTo NextRow
        If qtyVal <= 0 Then
            errNotes = "BoxBOM row " & CStr(r) & " needs a component Quantity greater than zero."
            Exit Function
        End If

        If cUom > 0 Then uomVal = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cUom).Value)) Else uomVal = ""
        If cLoc > 0 Then locVal = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cLoc).Value)) Else locVal = ""
        If cDesc > 0 Then descVal = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cDesc).Value)) Else descVal = ""
        If itemCode = "" Then itemCode = ResolveItemCodeForBoxBuildPayload(invLo, rowVal, itemName)
        If itemCode = "" And itemName <> "" Then itemCode = itemName

        AddBoxBuildPayloadItem payloadItems, rowVal, itemCode, itemName, qtyVal, uomVal, locVal, descVal, "USED"
        usedTotal = usedTotal + qtyVal
NextRow:
    Next r

    If usedTotal <= 0 Then
        errNotes = "No component quantities were found in BoxBOM."
        Exit Function
    End If
    AddBoxBuildComponentPayloadItems = True
End Function

Private Sub AddBoxBuildPayloadItem(ByVal payloadItems As Collection, _
                                   ByVal rowVal As Long, _
                                   ByVal itemCode As String, _
                                   ByVal itemName As String, _
                                   ByVal qtyVal As Double, _
                                   ByVal uomVal As String, _
                                   ByVal locationVal As String, _
                                   ByVal descriptionVal As String, _
                                   ByVal ioType As String)
    Dim payloadItem As Object

    If payloadItems Is Nothing Then Exit Sub
    Set payloadItem = modRoleEventWriter.CreatePayloadItem(rowVal, itemCode, qtyVal, locationVal, itemName, ioType)
    payloadItem("ROW") = rowVal
    payloadItem("ITEM_CODE") = itemCode
    payloadItem("ITEM") = itemName
    payloadItem("UOM") = uomVal
    payloadItem("DESCRIPTION") = descriptionVal
    payloadItem("LOCATION") = locationVal
    payloadItems.Add payloadItem
End Sub

Private Function ResolveItemCodeForBoxBuildPayload(ByVal invLo As ListObject, ByVal rowVal As Long, ByVal itemName As String) As String
    Dim invIdx As Long

    If invLo Is Nothing Then Exit Function
    If rowVal > 0 Then invIdx = FindInvRowIndexByRow(invLo, rowVal)
    If invIdx <= 0 And itemName <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
    If invIdx <= 0 Then Exit Function
    ResolveItemCodeForBoxBuildPayload = NzStr(GetInvSysValueByIndex(invLo, invIdx, "ITEM_CODE"))
End Function

Private Function BuildQueueableShipmentsSentDeltas(ByVal invLo As ListObject, ByVal ws As Worksheet, ByRef deltasOut As Collection, ByRef errNotes As String) As Boolean
    Dim aggPack As ListObject
    Dim rowFilter As Object
    Dim arrAgg As Variant
    Dim cRowAgg As Long
    Dim r As Long
    Dim filtered As Collection
    Dim delta As Variant

    Set deltasOut = BuildShipmentsSentDeltaPacket(invLo, errNotes)
    If deltasOut Is Nothing Then
        If errNotes = "" Then errNotes = "No staged shipments found in invSys.SHIPMENTS."
        Exit Function
    End If
    If deltasOut.Count = 0 Then
        If errNotes = "" Then errNotes = "No staged shipments found in invSys.SHIPMENTS."
        Exit Function
    End If

    Set aggPack = GetListObject(ws, TABLE_AGG_PACK)
    If Not aggPack Is Nothing Then
        If Not aggPack.DataBodyRange Is Nothing Then
            cRowAgg = ColumnIndex(aggPack, "ROW")
            If cRowAgg > 0 Then
                Set rowFilter = CreateObject("Scripting.Dictionary")
                arrAgg = aggPack.DataBodyRange.Value
                For r = 1 To UBound(arrAgg, 1)
                    If NzLng(arrAgg(r, cRowAgg)) > 0 Then rowFilter(CStr(NzLng(arrAgg(r, cRowAgg)))) = True
                Next r
            End If
        End If
    End If

    If Not rowFilter Is Nothing Then
        If rowFilter.Count > 0 Then
            Set filtered = New Collection
            For Each delta In deltasOut
                If rowFilter.Exists(CStr(delta("ROW"))) Then filtered.Add delta
            Next delta
            Set deltasOut = filtered
            If deltasOut.Count = 0 Then
                errNotes = "No staged shipments match the current AggregatePackages rows."
                Exit Function
            End If
        End If
    End If

    BuildQueueableShipmentsSentDeltas = True
End Function

Public Sub ShowDynamicItemSearch(ByVal targetCell As Range)
    On Error GoTo ErrHandler
    Dim refreshReport As String

    If targetCell Is Nothing Then Exit Sub
    If ShouldRefreshShippingBomBeforePicker(targetCell) Then
        RefreshShippingBomViewForWorkbook targetCell.Worksheet.Parent, refreshReport
    End If
    If mDynSearch Is Nothing Then Set mDynSearch = CreateDynItemSearch()
    mDynSearch.UseTemplateForm "ufShippingItemSearch"
    mDynSearch.ShowForCell targetCell
    Exit Sub
ErrHandler:
    MsgBox "Shipping item picker is unavailable: " & Err.Description, vbExclamation
End Sub

Private Function ShouldRefreshShippingBomBeforePicker(ByVal targetCell As Range) As Boolean
    Dim lo As ListObject
    Dim tableName As String

    If targetCell Is Nothing Then Exit Function
    On Error Resume Next
    Set lo = targetCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function

    tableName = LCase$(Trim$(lo.Name))
    ShouldRefreshShippingBomBeforePicker = (tableName = "shipmentstally" Or tableName = "boxbuilder")
End Function

Public Sub HandleShippingSelectionChange(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    If target.Cells.CountLarge > 1 Then Exit Sub
    If target.Worksheet Is Nothing Then Exit Sub
    If target.Worksheet.Parent Is Nothing Then Exit Sub
    If target.Worksheet.Parent.IsAddin Then Exit Sub
    If StrComp(target.Worksheet.Name, SHEET_SHIPMENTS, vbTextCompare) <> 0 Then Exit Sub

    Dim lo As ListObject
    Dim loName As String
    Dim targetCol As ListColumn

    On Error Resume Next
    Set lo = target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    loName = LCase$(lo.Name)
    Select Case loName
        Case "shipmentstally"
            On Error Resume Next
            Set targetCol = lo.ListColumns("ITEMS")
            On Error GoTo 0
        Case "boxbom"
            On Error Resume Next
            Set targetCol = lo.ListColumns("ITEM")
            On Error GoTo 0
        Case "boxbuilder"
            If Not IsBoxMakerMode(target.Worksheet) Then Exit Sub
            On Error Resume Next
            Set targetCol = lo.ListColumns("Box Name")
            On Error GoTo 0
        Case Else
            Exit Sub
    End Select

    If targetCol Is Nothing Then Exit Sub
    If target.Column <> targetCol.Range.Column Then Exit Sub
    If target.Row <= lo.HeaderRowRange.Row Then Exit Sub

    Set gSelectedCell = target
    ShowDynamicItemSearch target
End Sub

Public Sub HandleShippingSheetChange(ByVal target As Range)
    On Error GoTo ExitHandler
    If target Is Nothing Then Exit Sub
    If target.Cells.CountLarge > 50 Then Exit Sub
    If target.Worksheet Is Nothing Then Exit Sub
    If target.Worksheet.Parent Is Nothing Then Exit Sub
    If target.Worksheet.Parent.IsAddin Then Exit Sub
    If StrComp(target.Worksheet.Name, SHEET_SHIPMENTS, vbTextCompare) <> 0 Then Exit Sub

    Dim lo As ListObject
    Dim qtyCol As ListColumn
    Dim hit As Range

    On Error Resume Next
    Set lo = target.Worksheet.ListObjects(TABLE_SHIPMENTS)
    On Error GoTo 0
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            On Error Resume Next
            Set qtyCol = lo.ListColumns("QUANTITY")
            On Error GoTo 0
            If Not qtyCol Is Nothing Then
                Set hit = Application.Intersect(target, qtyCol.DataBodyRange)
                If Not hit Is Nothing Then
                    Application.EnableEvents = False
                    InvalidateAggregates True
                    GoTo ExitHandler
                End If
            End If
        End If
    End If

    If IsBoxMakerMode(target.Worksheet) Then
        If BoxMakerCurrentInventoryWasEdited(target.Worksheet, target) Then
            Application.EnableEvents = False
            RefreshBoxMakerCurrentInventory target.Worksheet
            GoTo ExitHandler
        End If
    End If

    Set lo = Nothing
    Set qtyCol = Nothing
    Set hit = Nothing
    On Error Resume Next
    Set lo = target.Worksheet.ListObjects(TABLE_BOX_BUILDER)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not IsBoxMakerMode(target.Worksheet) Then Exit Sub

    On Error Resume Next
    Set qtyCol = lo.ListColumns("Quantity")
    On Error GoTo 0
    If qtyCol Is Nothing Then Exit Sub

    Set hit = Application.Intersect(target, qtyCol.DataBodyRange)
    If hit Is Nothing Then Exit Sub

    Application.EnableEvents = False
    ReloadBoxMakerBomFromBuilder lo.Parent
ExitHandler:
    Application.EnableEvents = True
End Sub

' ===== button scaffolding =====
Private Sub EnsureShipmentsButtons(Optional ByVal targetWb As Workbook = Nothing)
    Dim ws As Worksheet
    Dim wb As Workbook

    Set wb = ResolveShippingWorkbook(targetWb, SHEET_SHIPMENTS)
    If wb Is Nothing Then Exit Sub
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    DeleteLegacyShippingButtons ws
    DeleteLegacyCheckBoxes ws

    Dim colA As Range: Set colA = ws.Columns("A")
    Dim leftA As Double: leftA = colA.Left + 2
    Dim colAWidth As Double
    colAWidth = colA.Width - 4
    If colAWidth < 40 Then colAWidth = 60

    Const BTN_STACK_SPACING As Double = 24
    Const CHK_STACK_SPACING As Double = 28
    Dim chkTop As Double: chkTop = ws.Rows(1).Top + 2
    EnsureCheckbox ws, CHK_USE_EXISTING, "Use existing shippable inventory", "modTS_Shipments.ToggleUseExistingInventory", leftA, chkTop, colAWidth
End Sub

Private Sub RefreshShipmentsUiAccess(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    modRoleUiAccess.ApplyShapeCapability ws, BTN_SHIPMENTS_SENT, "SHIP_POST"
End Sub

Private Sub DeleteLegacyShippingButtons(ByVal ws As Worksheet)
    DeleteShapeIfExists ws, "BTN_SHOW_BUILDER"
    DeleteShapeIfExists ws, "BTN_HIDE_BUILDER"
    DeleteShapeIfExists ws, BTN_TOGGLE_BUILDER
    DeleteShapeIfExists ws, BTN_SAVE_BOX
    DeleteShapeIfExists ws, BTN_SWITCH_BOXMAKER
    DeleteShapeIfExists ws, BTN_BOX_CREATED
    DeleteShapeIfExists ws, BTN_UNSHIP
    DeleteShapeIfExists ws, BTN_SEND_HOLD
    DeleteShapeIfExists ws, BTN_RETURN_HOLD
    DeleteShapeIfExists ws, BTN_CONFIRM_INV
    DeleteShapeIfExists ws, BTN_BOXES_MADE
    DeleteShapeIfExists ws, BTN_TO_TOTALINV
    DeleteShapeIfExists ws, BTN_TO_SHIPMENTS
    DeleteShapeIfExists ws, BTN_SHIPMENTS_SENT
End Sub

Private Sub ClearExcelClipboardStateShipping()
    On Error Resume Next
    If Application.CutCopyMode <> False Then Application.CutCopyMode = False
    On Error GoTo 0
End Sub

Public Sub ToggleUseExistingInventory()
    InvalidateAggregates True
End Sub

Private Sub EnsureButtonCustom(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String, leftPos As Double, topPos As Double, Optional widthPts As Double = 118)
    Const BTN_HEIGHT As Double = 20
    Dim resolvedOnAction As String

    If widthPts < 20 Then widthPts = 118
    resolvedOnAction = ResolveOnActionMacroShipping(onActionMacro)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, widthPts, BTN_HEIGHT)
        shp.Name = shapeName
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = resolvedOnAction
    Else
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = BTN_HEIGHT
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = resolvedOnAction
    End If
End Sub

Private Sub EnsureCheckbox(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String, leftPos As Double, topPos As Double, Optional widthPts As Double = 118)
    Const CHK_HEIGHT As Double = 26
    Dim resolvedOnAction As String

    If widthPts < 20 Then widthPts = 118
    resolvedOnAction = ResolveOnActionMacroShipping(onActionMacro)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        On Error Resume Next
        If shp.Type <> SHAPE_TYPE_FORM_CONTROL Or shp.FormControlType <> xlCheckBox Then
            Set shp = Nothing
        End If
        On Error GoTo 0
    End If
    If shp Is Nothing Then
        Dim candidate As Shape
        Dim bestMatch As Shape
        Dim bestTop As Double: bestTop = 1E+30
        For Each candidate In ws.Shapes
            If candidate.Type = SHAPE_TYPE_FORM_CONTROL Then
                On Error Resume Next
                If candidate.FormControlType = xlCheckBox Then
                    Dim cap As String: cap = candidate.ControlFormat.Caption
                    If LCase$(candidate.Name) Like "check box*" Or LCase$(cap) Like "check box*" Then
                        If candidate.Top < bestTop Then
                            Set bestMatch = candidate
                            bestTop = candidate.Top
                        End If
                    End If
                End If
                On Error GoTo 0
            End If
        Next candidate
        If Not bestMatch Is Nothing Then Set shp = bestMatch
    End If
    If Not shp Is Nothing Then
        Dim existingCap As String
        On Error Resume Next
        existingCap = shp.ControlFormat.Caption
        On Error GoTo 0
        If LCase$(existingCap) Like "check box*" Then
            On Error Resume Next
            shp.Delete
            On Error GoTo 0
            Set shp = Nothing
        End If
    End If
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlCheckBox, leftPos, topPos, widthPts, CHK_HEIGHT)
        shp.Name = shapeName
        shp.OnAction = resolvedOnAction
    Else
        shp.Name = shapeName
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = CHK_HEIGHT
        shp.OnAction = resolvedOnAction
    End If
    ForceCheckboxCaption shp, caption
End Sub

Private Function ResolveOnActionMacroShipping(ByVal onActionMacro As String) As String
    onActionMacro = Trim$(onActionMacro)
    If onActionMacro = "" Then Exit Function
    If InStr(1, onActionMacro, "!", vbTextCompare) > 0 Then
        ResolveOnActionMacroShipping = onActionMacro
    Else
        ResolveOnActionMacroShipping = "'" & ThisWorkbook.Name & "'!" & onActionMacro
    End If
End Function

Private Sub DeleteLegacyCheckBoxes(ws As Worksheet)
    Dim shp As Shape
    Dim toDelete As Collection: Set toDelete = New Collection
    For Each shp In ws.Shapes
        If shp.Type = SHAPE_TYPE_FORM_CONTROL Then
            On Error Resume Next
            If shp.FormControlType = xlCheckBox Then
                Dim cap As String: cap = shp.ControlFormat.Caption
                If (LCase$(shp.Name) Like "check box*" Or LCase$(cap) Like "check box*") _
                   And StrComp(shp.Name, CHK_USE_EXISTING, vbTextCompare) <> 0 Then
                    toDelete.Add shp.Name
                End If
            End If
            On Error GoTo 0
        End If
    Next shp
    Dim nameVal As Variant
    For Each nameVal In toDelete
        On Error Resume Next
        ws.Shapes(CStr(nameVal)).Delete
        On Error GoTo 0
    Next nameVal
End Sub

Private Sub ForceCheckboxCaption(shp As Shape, caption As String)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.ControlFormat.Caption = caption
    shp.TextFrame.Characters.Text = caption
    On Error GoTo 0
End Sub

Private Sub DeleteShapeIfExists(ws As Worksheet, shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Function ResolveRow(lo As ListObject, targetCell As Range) As ListRow
    If lo Is Nothing Then Exit Function
    
    Dim rowIdx As Long: rowIdx = 0
    
    If Not targetCell Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            If targetCell.Row >= lo.DataBodyRange.Row _
               And targetCell.Row <= lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1 Then
                rowIdx = targetCell.Row - lo.DataBodyRange.Row + 1
            End If
        End If
    End If
    
    If rowIdx >= 1 And rowIdx <= lo.ListRows.Count Then
        Set ResolveRow = lo.ListRows(rowIdx)
    End If
End Function

Private Sub WriteValue(lr As ListRow, columnName As String, value As Variant)
    If lr Is Nothing Then Exit Sub
    Dim colIdx As Long: colIdx = ColumnIndex(lr.Parent, columnName)
    If colIdx = 0 Then Exit Sub
    lr.Range.Cells(1, colIdx).Value = value
End Sub

' ===== builder helpers =====
Private Sub ToggleBuilderTables(ByVal makeVisible As Boolean)
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo1 As ListObject: Set lo1 = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim lo2 As ListObject: Set lo2 = GetListObject(ws, TABLE_BOX_BOM)
    If lo1 Is Nothing And lo2 Is Nothing Then Exit Sub

    Dim firstCol As Long: firstCol = 0
    Dim lastCol As Long: lastCol = 0

    Dim arrTables As Variant
    arrTables = Array(lo1, lo2)
    Dim idx As Long
    Dim lo As ListObject
    For idx = LBound(arrTables) To UBound(arrTables)
        Set lo = arrTables(idx)
        If Not lo Is Nothing Then
            Dim startCol As Long
            Dim endCol As Long
            startCol = lo.HeaderRowRange.Column
            endCol = startCol + lo.HeaderRowRange.Columns.Count - 1
            If firstCol = 0 Or startCol < firstCol Then firstCol = startCol
            If endCol > lastCol Then lastCol = endCol
        End If
    Next idx

    If firstCol = 0 Or lastCol = 0 Then Exit Sub

    ws.Range(ws.Columns(firstCol), ws.Columns(lastCol)).EntireColumn.Hidden = Not makeVisible
End Sub

Private Sub EnsureBuilderTablesReady(Optional ByVal targetWb As Workbook = Nothing)
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ResolveShippingWorkbook(targetWb, SHEET_SHIPMENTS)
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim loBuilder As ListObject: Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    UnhideListObjectWorksheetColumnsShipping loBuilder
    UnhideListObjectWorksheetColumnsShipping loBom
    If Not loBuilder Is Nothing Then NormalizeBoxBuilderTable loBuilder
    If loBom Is Nothing Then
        Set loBom = CreateBoxBomTable(ws, loBuilder)
    End If
    If Not loBom Is Nothing Then
        EnsureBoxBomEntryColumns loBom
        EnsureBoxBomStarterRows loBom
        RepairBoxBomRowsFromInventory loBom
    End If
    ArrangeBoxBuilderBandShipping loBuilder, loBom
End Sub

Private Sub NormalizeBoxBuilderTable(ByVal loBuilder As ListObject)
    If loBuilder Is Nothing Then Exit Sub
    UnhideListObjectWorksheetColumnsShipping loBuilder
    EnsureColumnExists loBuilder, "Box Name"
    EnsureColumnExists loBuilder, "UOM"
    EnsureColumnExists loBuilder, "LOCATION"
    EnsureColumnExists loBuilder, "DESCRIPTION"
    RemoveColumnIfExistsShipping loBuilder, "ROW"
    EnsureTableHasRow loBuilder
End Sub

Private Sub SetBoxMakerMode(ByVal ws As Worksheet, Optional ByVal enabled As Boolean = True)
    Dim loBuilder As ListObject
    Dim loBom As ListObject

    If ws Is Nothing Then Exit Sub
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBuilder Is Nothing Then Exit Sub

    NormalizeBoxBuilderTable loBuilder
    If enabled Then
        EnsureColumnExists loBuilder, "Quantity", "Box Name"
    Else
        RemoveColumnIfExistsShipping loBuilder, "Quantity"
        RemoveBoxMakerInventoryColumns loBuilder, loBom
    End If

    If Not loBom Is Nothing Then
        EnsureBoxBomEntryColumns loBom
        If enabled Then EnsureBoxMakerInventoryColumns loBuilder, loBom
        EnsureBoxBomStarterRows loBom
        RepairBoxBomRowsFromInventory loBom
    ElseIf enabled Then
        EnsureBoxMakerInventoryColumns loBuilder, loBom
    End If
    ArrangeShippingSurface ws.Parent
    If enabled Then RefreshBoxMakerCurrentInventory ws
End Sub

Private Sub EnsureBoxMakerInventoryColumns(ByVal loBuilder As ListObject, ByVal loBom As ListObject)
    If Not loBuilder Is Nothing Then
        EnsureColumnExists loBuilder, COL_CURRENT_INV, "Quantity"
        FormatCurrentInventoryColumn loBuilder
    End If
    If Not loBom Is Nothing Then
        EnsureColumnExists loBom, COL_CURRENT_INV, "QUANTITY"
        FormatCurrentInventoryColumn loBom
    End If
End Sub

Private Sub RemoveBoxMakerInventoryColumns(ByVal loBuilder As ListObject, ByVal loBom As ListObject)
    RemoveColumnIfExistsShipping loBuilder, COL_CURRENT_INV
    RemoveColumnIfExistsShipping loBom, COL_CURRENT_INV
End Sub

Private Sub FormatCurrentInventoryColumn(ByVal lo As ListObject)
    Dim colIdx As Long

    If lo Is Nothing Then Exit Sub
    colIdx = ColumnIndex(lo, COL_CURRENT_INV)
    If colIdx = 0 Then Exit Sub

    On Error Resume Next
    lo.ListColumns(colIdx).Range.Locked = True
    lo.ListColumns(colIdx).Range.Interior.Color = RGB(242, 242, 242)
    lo.ListColumns(colIdx).Range.Font.Color = RGB(96, 96, 96)
    On Error GoTo 0
End Sub

Private Function BoxMakerCurrentInventoryWasEdited(ByVal ws As Worksheet, ByVal target As Range) As Boolean
    Dim lo As ListObject
    Dim colIdx As Long
    Dim hit As Range

    If ws Is Nothing Or target Is Nothing Then Exit Function

    Set lo = GetListObject(ws, TABLE_BOX_BUILDER)
    If Not lo Is Nothing Then
        colIdx = ColumnIndex(lo, COL_CURRENT_INV)
        If colIdx > 0 Then
            If Not lo.DataBodyRange Is Nothing Then
                Set hit = Application.Intersect(target, lo.DataBodyRange.Columns(colIdx))
                If Not hit Is Nothing Then
                    BoxMakerCurrentInventoryWasEdited = True
                    Exit Function
                End If
            End If
        End If
    End If

    Set lo = GetListObject(ws, TABLE_BOX_BOM)
    If Not lo Is Nothing Then
        colIdx = ColumnIndex(lo, COL_CURRENT_INV)
        If colIdx > 0 Then
            If Not lo.DataBodyRange Is Nothing Then
                Set hit = Application.Intersect(target, lo.DataBodyRange.Columns(colIdx))
                If Not hit Is Nothing Then
                    BoxMakerCurrentInventoryWasEdited = True
                    Exit Function
                End If
            End If
        End If
    End If
End Function

Private Sub RefreshBoxMakerCurrentInventory(ByVal ws As Worksheet)
    On Error GoTo CleanExit

    Dim loBuilder As ListObject
    Dim loBom As ListObject
    Dim invLo As ListObject
    Dim snapshotCache As Object
    Dim unresolvedCount As Long
    Dim refreshReport As String
    Dim eventsState As Boolean
    Dim restoreEvents As Boolean

    If ws Is Nothing Then Exit Sub
    If Not IsBoxMakerMode(ws) Then Exit Sub

    eventsState = Application.EnableEvents
    restoreEvents = True
    Application.EnableEvents = False

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBuilder Is Nothing And loBom Is Nothing Then GoTo CleanExit

    EnsureBoxMakerInventoryColumns loBuilder, loBom

    Set invLo = GetInvSysTableFromWorkbook(ws.Parent)
    If invLo Is Nothing Then Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        ClearBoxMakerCurrentInventory loBuilder
        ClearBoxMakerCurrentInventory loBom
        GoTo CleanExit
    End If

    unresolvedCount = RefreshBoxBuilderCurrentInventory(ws, loBuilder, invLo, snapshotCache)
    unresolvedCount = unresolvedCount + RefreshBoxBomCurrentInventory(ws, loBom, invLo, snapshotCache)

    If unresolvedCount > 0 Then
        On Error Resume Next
        modOperatorReadModel.RefreshInventoryReadModelForWorkbook ws.Parent, "", "LOCAL", refreshReport
        On Error GoTo CleanExit
        Set invLo = GetInvSysTableFromWorkbook(ws.Parent)
        If invLo Is Nothing Then Set invLo = GetInvSysTable()
        If Not invLo Is Nothing Then
            Set snapshotCache = Nothing
            RefreshBoxBuilderCurrentInventory ws, loBuilder, invLo, snapshotCache
            RefreshBoxBomCurrentInventory ws, loBom, invLo, snapshotCache
        End If
    End If

CleanExit:
    If restoreEvents Then Application.EnableEvents = eventsState
End Sub

Private Sub ClearBoxMakerCurrentInventory(ByVal lo As ListObject)
    Dim colIdx As Long

    If lo Is Nothing Then Exit Sub
    colIdx = ColumnIndex(lo, COL_CURRENT_INV)
    If colIdx = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.Columns(colIdx).ClearContents
    FormatCurrentInventoryColumn lo
End Sub

Private Function RefreshBoxBuilderCurrentInventory(ByVal ws As Worksheet, _
                                                   ByVal loBuilder As ListObject, _
                                                   ByVal invLo As ListObject, _
                                                   ByRef snapshotCache As Object) As Long
    Dim colInv As Long
    Dim boxName As String
    Dim invIdx As Long
    Dim packageRow As Long
    Dim currentInv As Variant
    Dim found As Boolean

    If loBuilder Is Nothing Then Exit Function
    colInv = ColumnIndex(loBuilder, COL_CURRENT_INV)
    If colInv = 0 Then Exit Function
    If loBuilder.DataBodyRange Is Nothing Then Exit Function

    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    If Not invLo Is Nothing Then
        If boxName <> "" Then invIdx = FindInvRowIndexByItem(invLo, boxName)
    End If
    If invIdx <= 0 Then
        packageRow = ResolveBoxMakerPackageRow(loBuilder, invLo)
        If packageRow > 0 And Not invLo Is Nothing Then invIdx = FindInvRowIndexByRow(invLo, packageRow)
    End If

    currentInv = ResolveCurrentInventoryValue(ws, invLo, packageRow, boxName, found, snapshotCache)
    If found Then
        loBuilder.DataBodyRange.Cells(1, colInv).Value = currentInv
    Else
        loBuilder.DataBodyRange.Cells(1, colInv).ClearContents
        If boxName <> "" Or packageRow > 0 Then RefreshBoxBuilderCurrentInventory = 1
    End If
    FormatCurrentInventoryColumn loBuilder
End Function

Private Function RefreshBoxBomCurrentInventory(ByVal ws As Worksheet, _
                                               ByVal loBom As ListObject, _
                                               ByVal invLo As ListObject, _
                                               ByRef snapshotCache As Object) As Long
    Dim cInv As Long
    Dim cRow As Long
    Dim cItem As Long
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String
    Dim invIdx As Long
    Dim currentInv As Variant
    Dim found As Boolean

    If loBom Is Nothing Then Exit Function
    cInv = ColumnIndex(loBom, COL_CURRENT_INV)
    If cInv = 0 Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function

    cRow = ColumnIndex(loBom, "ROW")
    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)

    For r = 1 To loBom.ListRows.Count
        rowVal = 0
        itemName = ""
        invIdx = 0
        If cRow > 0 Then rowVal = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        If cItem > 0 Then itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))
        If Not invLo Is Nothing Then
            If rowVal > 0 Then invIdx = FindInvRowIndexByRow(invLo, rowVal)
            If invIdx <= 0 And itemName <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
        End If

        found = False
        currentInv = ResolveCurrentInventoryValue(ws, invLo, rowVal, itemName, found, snapshotCache)
        If found Then
            loBom.DataBodyRange.Cells(r, cInv).Value = currentInv
        Else
            loBom.DataBodyRange.Cells(r, cInv).ClearContents
            If rowVal > 0 Or itemName <> "" Then RefreshBoxBomCurrentInventory = RefreshBoxBomCurrentInventory + 1
        End If
    Next r
    FormatCurrentInventoryColumn loBom
End Function

Private Function ResolveCurrentInventoryValue(ByVal ws As Worksheet, _
                                              ByVal invLo As ListObject, _
                                              ByVal rowVal As Long, _
                                              ByVal itemName As String, _
                                              ByRef found As Boolean, _
                                              ByRef snapshotCache As Object) As Variant
    Dim invIdx As Long
    Dim totalVal As Variant
    Dim cacheKey As String

    found = False
    If snapshotCache Is Nothing Then Set snapshotCache = BuildRuntimeSnapshotInventoryCache()
    If Not snapshotCache Is Nothing Then
        If rowVal > 0 Then
            cacheKey = "ROW:" & CStr(rowVal)
            If snapshotCache.Exists(cacheKey) Then
                found = True
                ResolveCurrentInventoryValue = snapshotCache(cacheKey)
                Exit Function
            End If
        End If
        If Trim$(itemName) <> "" Then
            cacheKey = "ITEM:" & LCase$(Trim$(itemName))
            If snapshotCache.Exists(cacheKey) Then
                found = True
                ResolveCurrentInventoryValue = snapshotCache(cacheKey)
                Exit Function
            End If
        End If
    End If

    If Not invLo Is Nothing Then
        If rowVal > 0 Then invIdx = FindInvRowIndexByRow(invLo, rowVal)
        If invIdx <= 0 And Trim$(itemName) <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
        If invIdx > 0 Then
            totalVal = GetInvSysValueByIndex(invLo, invIdx, "TOTAL INV")
            If Not IsBlankInventoryValue(totalVal) Then
                found = True
                ResolveCurrentInventoryValue = totalVal
                Exit Function
            End If
        End If
    End If

    ResolveCurrentInventoryValue = ResolveCurrentInventoryFromTable(GetListObject(ws, "invSysData_Shipping"), rowVal, itemName, found)
    If found Then Exit Function

    ResolveCurrentInventoryValue = ResolveCurrentInventoryFromTable(GetListObject(ws, TABLE_CHECK_INV), rowVal, itemName, found)
    If found Then Exit Function
End Function

Private Function ResolveCurrentInventoryFromTable(ByVal lo As ListObject, _
                                                  ByVal rowVal As Long, _
                                                  ByVal itemName As String, _
                                                  ByRef found As Boolean) As Variant
    Dim cRow As Long
    Dim cItem As Long
    Dim cTotal As Long
    Dim r As Long
    Dim totalVal As Variant

    found = False
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    cRow = ColumnIndex(lo, "ROW")
    cItem = ColumnIndex(lo, "ITEM")
    cTotal = ColumnIndex(lo, "TOTAL INV")
    If cTotal = 0 Then cTotal = ColumnIndex(lo, "QtyAvailable")
    If cTotal = 0 Then Exit Function

    For r = 1 To lo.ListRows.Count
        If rowVal > 0 And cRow > 0 Then
            If NzLng(lo.DataBodyRange.Cells(r, cRow).Value) <> rowVal Then GoTo NextRow
        ElseIf Trim$(itemName) <> "" And cItem > 0 Then
            If StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cItem).Value)), Trim$(itemName), vbTextCompare) <> 0 Then GoTo NextRow
        Else
            Exit Function
        End If

        totalVal = lo.DataBodyRange.Cells(r, cTotal).Value
        If Not IsBlankInventoryValue(totalVal) Then
            found = True
            ResolveCurrentInventoryFromTable = totalVal
            Exit Function
        End If
NextRow:
    Next r
End Function

Private Function IsBlankInventoryValue(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then
        IsBlankInventoryValue = True
    ElseIf IsError(value) Then
        IsBlankInventoryValue = True
    Else
        IsBlankInventoryValue = (Trim$(NzStr(value)) = "")
    End If
End Function

Private Function BuildRuntimeSnapshotInventoryCache() As Object
    On Error GoTo CleanExit

    Dim result As Object
    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim snapshotPath As String
    Dim wbSnap As Workbook
    Dim loSnap As ListObject
    Dim openedTransient As Boolean
    Dim cRow As Long
    Dim cItem As Long
    Dim cSku As Long
    Dim cQty As Long
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String
    Dim sku As String
    Dim qtyVal As Variant

    Set result = CreateObject("Scripting.Dictionary")
    Set BuildRuntimeSnapshotInventoryCache = result

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then Exit Function

    AddCurrentInventoryCacheFromWorkbookPath result, _
                                             rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb", _
                                             "invSys", _
                                             "TOTAL INV", _
                                             "QtyAvailable"
    AddSkuBalanceInventoryCacheFromWorkbookPath result, _
                                                rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb"

    snapshotPath = rootPath & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    Set wbSnap = FindOpenWorkbookByFullNameShipping(snapshotPath)
    If wbSnap Is Nothing Then
        If Len(Dir$(snapshotPath)) = 0 Then Exit Function
        Set wbSnap = Application.Workbooks.Open(Filename:=snapshotPath, UpdateLinks:=False, ReadOnly:=True)
        openedTransient = True
    End If

    Set loSnap = FindSnapshotListObjectShipping(wbSnap)
    If loSnap Is Nothing Then GoTo CleanExit
    If loSnap.DataBodyRange Is Nothing Then GoTo CleanExit

    cRow = ColumnIndex(loSnap, "ROW")
    cItem = ColumnIndex(loSnap, "ITEM")
    cSku = ColumnIndex(loSnap, "SKU")
    cQty = ColumnIndex(loSnap, "QtyOnHand")
    If cQty = 0 Then cQty = ColumnIndex(loSnap, "QtyAvailable")
    If cQty = 0 Then GoTo CleanExit

    For r = 1 To loSnap.ListRows.Count
        qtyVal = loSnap.DataBodyRange.Cells(r, cQty).Value
        If IsBlankInventoryValue(qtyVal) Then GoTo NextSnapshotRow

        rowVal = 0
        itemName = ""
        sku = ""
        If cRow > 0 Then rowVal = NzLng(loSnap.DataBodyRange.Cells(r, cRow).Value)
        If cItem > 0 Then itemName = Trim$(NzStr(loSnap.DataBodyRange.Cells(r, cItem).Value))
        If cSku > 0 Then sku = Trim$(NzStr(loSnap.DataBodyRange.Cells(r, cSku).Value))

        If rowVal > 0 Then result("ROW:" & CStr(rowVal)) = qtyVal
        If itemName <> "" Then result("ITEM:" & LCase$(itemName)) = qtyVal
        If sku <> "" Then result("ITEM:" & LCase$(sku)) = qtyVal
NextSnapshotRow:
    Next r

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbSnap
End Function

Private Sub AddCurrentInventoryCacheFromWorkbookPath(ByVal cache As Object, _
                                                     ByVal workbookPath As String, _
                                                     ByVal preferredTableName As String, _
                                                     ByVal primaryQtyColumn As String, _
                                                     ByVal fallbackQtyColumn As String)
    On Error GoTo CleanExit

    Dim wb As Workbook
    Dim lo As ListObject
    Dim openedTransient As Boolean
    Dim cRow As Long
    Dim cItem As Long
    Dim cSku As Long
    Dim cQty As Long
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String
    Dim sku As String
    Dim qtyVal As Variant

    If cache Is Nothing Then Exit Sub
    If Trim$(workbookPath) = "" Then Exit Sub
    If Len(Dir$(workbookPath)) = 0 Then Exit Sub

    Set wb = FindOpenWorkbookByFullNameShipping(workbookPath)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(Filename:=workbookPath, UpdateLinks:=False, ReadOnly:=True)
        openedTransient = True
    End If
    If wb Is Nothing Then GoTo CleanExit

    Set lo = FindListObjectByNameShipping(wb, preferredTableName)
    If lo Is Nothing Then GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit

    cRow = ColumnIndex(lo, "ROW")
    cItem = ColumnIndex(lo, "ITEM")
    cSku = ColumnIndex(lo, "ITEM_CODE")
    If cSku = 0 Then cSku = ColumnIndex(lo, "SKU")
    cQty = ColumnIndex(lo, primaryQtyColumn)
    If cQty = 0 Then cQty = ColumnIndex(lo, fallbackQtyColumn)
    If cQty = 0 Then GoTo CleanExit

    For r = 1 To lo.ListRows.Count
        qtyVal = lo.DataBodyRange.Cells(r, cQty).Value
        If IsBlankInventoryValue(qtyVal) Then GoTo NextRow

        rowVal = 0
        itemName = ""
        sku = ""
        If cRow > 0 Then rowVal = NzLng(lo.DataBodyRange.Cells(r, cRow).Value)
        If cItem > 0 Then itemName = Trim$(NzStr(lo.DataBodyRange.Cells(r, cItem).Value))
        If cSku > 0 Then sku = Trim$(NzStr(lo.DataBodyRange.Cells(r, cSku).Value))

        If rowVal > 0 Then cache("ROW:" & CStr(rowVal)) = qtyVal
        If itemName <> "" Then cache("ITEM:" & LCase$(itemName)) = qtyVal
        If sku <> "" Then cache("ITEM:" & LCase$(sku)) = qtyVal
NextRow:
    Next r

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wb
End Sub

Private Sub AddSkuBalanceInventoryCacheFromWorkbookPath(ByVal cache As Object, _
                                                        ByVal workbookPath As String)
    On Error GoTo CleanExit

    Dim wb As Workbook
    Dim loBalance As ListObject
    Dim openedTransient As Boolean
    Dim skuMeta As Object
    Dim cSku As Long
    Dim cQty As Long
    Dim r As Long
    Dim sku As String
    Dim qtyVal As Variant

    If cache Is Nothing Then Exit Sub
    If Trim$(workbookPath) = "" Then Exit Sub
    If Len(Dir$(workbookPath)) = 0 Then Exit Sub

    Set wb = FindOpenWorkbookByFullNameShipping(workbookPath)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(Filename:=workbookPath, UpdateLinks:=False, ReadOnly:=True)
        openedTransient = True
    End If
    If wb Is Nothing Then GoTo CleanExit

    Set loBalance = FindListObjectByNameShipping(wb, "tblSkuBalance")
    If loBalance Is Nothing Then GoTo CleanExit
    If loBalance.DataBodyRange Is Nothing Then GoTo CleanExit

    Set skuMeta = BuildSkuMetadataCacheShipping(wb)

    cSku = ColumnIndex(loBalance, "SKU")
    If cSku = 0 Then cSku = ColumnIndex(loBalance, "ITEM_CODE")
    cQty = ColumnIndex(loBalance, "QtyOnHand")
    If cQty = 0 Then cQty = ColumnIndex(loBalance, "TOTAL INV")
    If cSku = 0 Or cQty = 0 Then GoTo CleanExit

    For r = 1 To loBalance.ListRows.Count
        sku = Trim$(NzStr(loBalance.DataBodyRange.Cells(r, cSku).Value))
        If sku = "" Then GoTo NextRow
        qtyVal = loBalance.DataBodyRange.Cells(r, cQty).Value
        If IsBlankInventoryValue(qtyVal) Then GoTo NextRow

        AddSkuBalanceCacheEntryShipping cache, skuMeta, sku, qtyVal
NextRow:
    Next r

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wb
End Sub

Private Function BuildSkuMetadataCacheShipping(ByVal wb As Workbook) As Object
    Dim result As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    AddSkuMetadataFromTableShipping result, FindListObjectByNameShipping(wb, "tblSkuCatalog")
    AddSkuMetadataFromTableShipping result, FindListObjectByNameShipping(wb, "invSys")
    AddSkuMetadataFromTableShipping result, FindListObjectByNameShipping(wb, "tblItemSearchIndex")

    Set BuildSkuMetadataCacheShipping = result
End Function

Private Sub AddSkuMetadataFromTableShipping(ByVal cache As Object, ByVal lo As ListObject)
    Dim cSku As Long
    Dim cRow As Long
    Dim cItem As Long
    Dim r As Long
    Dim sku As String
    Dim info As Object

    If cache Is Nothing Then Exit Sub
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    cSku = ColumnIndex(lo, "SKU")
    If cSku = 0 Then cSku = ColumnIndex(lo, "ITEM_CODE")
    If cSku = 0 Then Exit Sub
    cRow = ColumnIndex(lo, "ROW")
    cItem = ColumnIndex(lo, "ITEM")

    For r = 1 To lo.ListRows.Count
        sku = Trim$(NzStr(lo.DataBodyRange.Cells(r, cSku).Value))
        If sku = "" Then GoTo NextRow
        Set info = CreateObject("Scripting.Dictionary")
        info.CompareMode = vbTextCompare
        info("SKU") = sku
        If cRow > 0 Then info("ROW") = NzLng(lo.DataBodyRange.Cells(r, cRow).Value)
        If cItem > 0 Then info("ITEM") = Trim$(NzStr(lo.DataBodyRange.Cells(r, cItem).Value))
        If cache.Exists(sku) Then cache.Remove sku
        cache.Add sku, info
NextRow:
    Next r
End Sub

Private Sub AddSkuBalanceCacheEntryShipping(ByVal cache As Object, _
                                            ByVal skuMeta As Object, _
                                            ByVal sku As String, _
                                            ByVal qtyVal As Variant)
    Dim info As Object
    Dim rowVal As Long
    Dim itemName As String

    If cache Is Nothing Then Exit Sub
    sku = Trim$(sku)
    If sku = "" Then Exit Sub

    cache("ITEM:" & LCase$(sku)) = qtyVal
    If skuMeta Is Nothing Then Exit Sub
    If Not skuMeta.Exists(sku) Then Exit Sub

    Set info = skuMeta(sku)
    On Error Resume Next
    rowVal = NzLng(info("ROW"))
    itemName = Trim$(NzStr(info("ITEM")))
    On Error GoTo 0

    If rowVal > 0 Then cache("ROW:" & CStr(rowVal)) = qtyVal
    If itemName <> "" Then cache("ITEM:" & LCase$(itemName)) = qtyVal
End Sub

Private Function FindSnapshotListObjectShipping(ByVal wb As Workbook) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set ws = wb.Worksheets("InventorySnapshot")
    If Not ws Is Nothing Then Set FindSnapshotListObjectShipping = ws.ListObjects("tblInventorySnapshot")
    On Error GoTo 0
    If Not FindSnapshotListObjectShipping Is Nothing Then Exit Function

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, "tblInventorySnapshot", vbTextCompare) = 0 Then
                Set FindSnapshotListObjectShipping = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function FindListObjectByNameShipping(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindListObjectByNameShipping = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function IsBoxMakerMode(ByVal ws As Worksheet) As Boolean
    Dim loBuilder As ListObject

    If ws Is Nothing Then Exit Function
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then Exit Function
    IsBoxMakerMode = (ColumnIndex(loBuilder, "Quantity") > 0)
End Function

Private Sub InvalidateShippingRibbonLabels()
    On Error Resume Next
    modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
    On Error GoTo 0
End Sub

Private Sub ArrangeBoxBuilderBandShipping(ByVal loBuilder As ListObject, ByVal loBom As ListObject)
    Dim targetRow As Long
    Dim targetCol As Long

    If loBuilder Is Nothing Or loBom Is Nothing Then Exit Sub
    targetRow = loBuilder.Range.Row + loBuilder.Range.Rows.Count + SHIP_LAYOUT_GAP_ROWS + 1
    targetCol = loBuilder.Range.Column
    If targetRow < 1 Or targetCol < 1 Then Exit Sub
    MoveListObjectToRowColShipping loBom, targetRow, targetCol
End Sub

Private Function CreateBoxBomTable(ByVal ws As Worksheet, ByVal loBuilder As ListObject) As ListObject
    Dim startCell As Range
    Dim startRow As Long
    Dim startCol As Long
    Dim headers As Variant
    Dim i As Long
    Dim dataRange As Range
    Dim lo As ListObject

    If ws Is Nothing Then Exit Function
    startRow = ws.Range(SHIP_LAYOUT_BOM_ADDR).Row
    startCol = ws.Range(SHIP_LAYOUT_BOM_ADDR).Column
    If Not loBuilder Is Nothing Then
        startRow = Application.WorksheetFunction.Max(startRow, loBuilder.Range.Row + loBuilder.Range.Rows.Count + SHIP_LAYOUT_GAP_ROWS + 1)
        startCol = loBuilder.Range.Column
    End If

    headers = Array(COL_BOXBOM_ITEM, "ITEM_CODE", "ROW", "QUANTITY", "UOM", "LOCATION", "DESCRIPTION")
    Set startCell = ws.Cells(startRow, startCol)
    For i = LBound(headers) To UBound(headers)
        startCell.Offset(0, i - LBound(headers)).Value = headers(i)
    Next i

    Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
    Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    lo.Name = TABLE_BOX_BOM
    Set CreateBoxBomTable = lo
End Function

Private Sub EnsureBoxBomStarterRows(ByVal loBom As ListObject)
    Const STARTER_ROWS As Long = 10
    Dim rowCount As Long

    If loBom Is Nothing Then Exit Sub
    If loBom.DataBodyRange Is Nothing Then
        rowCount = 0
    Else
        rowCount = loBom.ListRows.Count
    End If
    Do While rowCount < STARTER_ROWS
        loBom.ListRows.Add
        rowCount = rowCount + 1
    Loop
End Sub

Private Function TableRowIsBlankShipping(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim cell As Range

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then
        TableRowIsBlankShipping = True
        Exit Function
    End If
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function

    For Each cell In lo.ListRows(rowIndex).Range.Cells
        If Trim$(NzStr(cell.Value)) <> "" Then Exit Function
    Next cell
    TableRowIsBlankShipping = True
End Function

Public Sub ApplyItemSelection(targetCell As Range, lo As ListObject, rowIndex As Long, _
    ByVal itemName As String, ByVal itemCode As String, ByVal itemRow As Long, _
    ByVal uom As String, ByVal location As String, ByVal vendor As String, _
    Optional ByVal description As String = "")

    If lo Is Nothing Then Exit Sub
    
    Dim tableName As String
    tableName = LCase$(lo.Name)
    Dim targetRowIndex As Long

    Select Case tableName
        Case "shipmentstally"
            targetRowIndex = rowIndex
            If targetRowIndex <= 0 Then
                If Not targetCell Is Nothing Then
                    If Not lo.DataBodyRange Is Nothing Then
                        targetRowIndex = targetCell.Row - lo.DataBodyRange.Row + 1
                    End If
                End If
            End If
            If targetRowIndex <= 0 Then
                On Error Resume Next
                lo.ListRows.Add AlwaysInsert:=True
                On Error GoTo 0
                If Not lo.DataBodyRange Is Nothing Then targetRowIndex = lo.ListRows.Count
            End If
            If targetRowIndex <= 0 Then Exit Sub

            WriteValue lo.ListRows(targetRowIndex), "ITEMS", itemName
            WriteValue lo.ListRows(targetRowIndex), "ROW", itemRow
            WriteValue lo.ListRows(targetRowIndex), "UOM", uom
            WriteValue lo.ListRows(targetRowIndex), "LOCATION", location
            WriteValue lo.ListRows(targetRowIndex), "DESCRIPTION", description
            InvalidateAggregates True
            
        Case Else
            ' no-op
    End Select
End Sub

Public Sub ApplyItemToBoxBOM(targetCell As Range, ByVal itemName As String, ByVal itemRow As Long, _
    ByVal uom As String, ByVal location As String, ByVal description As String, Optional ByVal itemCode As String = "")

    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Sub
    EnsureBoxBomEntryColumns loBom

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then Exit Sub

    Dim invIdx As Long
    If itemRow > 0 Then invIdx = FindInvRowIndexByRow(invLo, itemRow)
    If invIdx = 0 And Len(Trim$(itemName)) > 0 Then
        invIdx = FindInvRowIndexByItem(invLo, itemName)
    End If
    If invIdx = 0 Then
        MsgBox "Item '" & itemName & "' not found in invSys.", vbExclamation
        Exit Sub
    End If

    Dim actualRow As Long
    Dim actualItem As String
    Dim actualUom As String, actualLoc As String, actualDesc As String

    Dim colRowInv As Long: colRowInv = ColumnIndex(invLo, "ROW")
    Dim colItemCodeInv As Long: colItemCodeInv = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemInv As Long: colItemInv = ColumnIndex(invLo, "ITEM")
    Dim colUomInv As Long: colUomInv = ColumnIndex(invLo, "UOM")
    Dim colLocInv As Long: colLocInv = ColumnIndex(invLo, "LOCATION")
    Dim colDescInv As Long: colDescInv = ColumnIndex(invLo, "DESCRIPTION")

    If colRowInv > 0 Then actualRow = NzLng(invLo.DataBodyRange.Cells(invIdx, colRowInv).Value)
    If actualRow <= 0 Then actualRow = RepairInvSysRowKeyShipping(invLo, invIdx, colRowInv)
    If colItemCodeInv > 0 And itemCode = "" Then itemCode = NzStr(invLo.DataBodyRange.Cells(invIdx, colItemCodeInv).Value)
    If colItemInv > 0 Then actualItem = NzStr(invLo.DataBodyRange.Cells(invIdx, colItemInv).Value)
    If colUomInv > 0 Then actualUom = NzStr(invLo.DataBodyRange.Cells(invIdx, colUomInv).Value)
    If colLocInv > 0 Then actualLoc = NzStr(invLo.DataBodyRange.Cells(invIdx, colLocInv).Value)
    If colDescInv > 0 Then actualDesc = NzStr(invLo.DataBodyRange.Cells(invIdx, colDescInv).Value)

    If Len(actualItem) = 0 Then actualItem = itemName
    If actualRow = 0 Then actualRow = itemRow
    If Len(actualUom) = 0 Then actualUom = uom
    If Len(actualLoc) = 0 Then actualLoc = location
    If Len(actualDesc) = 0 Then actualDesc = description

    Dim lr As ListRow
    Dim rowIdxResolved As Long: rowIdxResolved = 0
    If Not loBom.DataBodyRange Is Nothing And Not targetCell Is Nothing Then
        If targetCell.Row >= loBom.DataBodyRange.Row _
           And targetCell.Row <= loBom.DataBodyRange.Row + loBom.DataBodyRange.Rows.Count - 1 Then
            rowIdxResolved = targetCell.Row - loBom.DataBodyRange.Row + 1
        End If
    End If
    If rowIdxResolved >= 1 And rowIdxResolved <= loBom.ListRows.Count Then
        Set lr = loBom.ListRows(rowIdxResolved)
    Else
        Set lr = loBom.ListRows.Add
    End If

    WriteValue lr, COL_BOXBOM_ITEM, actualItem
    WriteValue lr, "ITEM_CODE", itemCode
    WriteValue lr, "ROW", actualRow
    WriteValue lr, "UOM", actualUom
    WriteValue lr, "LOCATION", actualLoc
    WriteValue lr, "DESCRIPTION", actualDesc
    Exit Sub

ErrHandler:
    MsgBox "ApplyItemToBoxBOM error: " & Err.Description, vbCritical
End Sub

Public Function LoadShippingBomPackagePickerItems() As Variant
    On Error GoTo FailSoft

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim openedTransient As Boolean
    Dim report As String

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then Exit Function

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbBom Is Nothing Then GoTo CleanExit

    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit

    LoadShippingBomPackagePickerItems = BuildPackagePickerItemsFromShippingBom(loBom)

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    Resume CleanExit
End Function

Private Function BuildPackagePickerItemsFromShippingBom(ByVal loBom As ListObject) As Variant
    Dim cPackageRow As Long
    Dim cPackageItem As Long
    Dim cPackageUom As Long
    Dim cPackageLocation As Long
    Dim cPackageDescription As Long
    Dim dict As Object
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim rowKey As String
    Dim itemName As String
    Dim uniqueKey As String

    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function

    cPackageRow = ColumnIndex(loBom, "PackageRow")
    cPackageItem = ColumnIndex(loBom, "PackageItem")
    cPackageUom = ColumnIndex(loBom, "PackageUOM")
    cPackageLocation = ColumnIndex(loBom, "PackageLocation")
    cPackageDescription = ColumnIndex(loBom, "PackageDescription")
    If cPackageRow = 0 Or cPackageItem = 0 Then Exit Function

    src = loBom.DataBodyRange.Value
    Set dict = CreateObject("Scripting.Dictionary")
    ReDim result(1 To UBound(src, 1), 1 To 7)

    For r = 1 To UBound(src, 1)
        rowKey = Trim$(NzStr(src(r, cPackageRow)))
        itemName = Trim$(NzStr(src(r, cPackageItem)))
        If rowKey = "" And itemName = "" Then GoTo NextPackage

        If rowKey <> "" Then
            uniqueKey = "ROW:" & rowKey
        Else
            uniqueKey = "ITEM:" & LCase$(itemName)
        End If
        If dict.Exists(uniqueKey) Then GoTo NextPackage
        dict.Add uniqueKey, True

        outRow = outRow + 1
        result(outRow, 1) = rowKey
        result(outRow, 2) = itemName
        result(outRow, 3) = itemName
        If cPackageUom > 0 Then result(outRow, 4) = NzStr(src(r, cPackageUom))
        If cPackageLocation > 0 Then result(outRow, 5) = NzStr(src(r, cPackageLocation))
        If cPackageDescription > 0 Then result(outRow, 6) = NzStr(src(r, cPackageDescription))
        result(outRow, 7) = ""
NextPackage:
    Next r

    If outRow = 0 Then Exit Function
    If outRow = UBound(src, 1) Then
        BuildPackagePickerItemsFromShippingBom = result
        Exit Function
    End If

    ReDim trimmed(1 To outRow, 1 To 7)
    For r = 1 To outRow
        For c = 1 To 7
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    BuildPackagePickerItemsFromShippingBom = trimmed
End Function

Public Sub ApplyItemToBoxBuilder(targetCell As Range, ByVal itemName As String, ByVal itemRow As Long, _
    ByVal uom As String, ByVal location As String, ByVal description As String)

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim loBuilder As ListObject
    Dim invLo As ListObject
    Dim invIdx As Long
    Dim actualRow As Long
    Dim actualItem As String
    Dim actualUom As String
    Dim actualLoc As String
    Dim actualDesc As String
    Dim report As String

    If targetCell Is Nothing Then
        Set ws = SheetExists(SHEET_SHIPMENTS)
    Else
        Set ws = targetCell.Worksheet
    End If
    If ws Is Nothing Then Exit Sub
    If Not IsBoxMakerMode(ws) Then Exit Sub

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then Exit Sub
    NormalizeBoxBuilderTable loBuilder
    EnsureColumnExists loBuilder, "Quantity", "Box Name"
    EnsureTableHasRow loBuilder

    Set invLo = GetInvSysTableFromWorkbook(ws.Parent)
    If invLo Is Nothing Then Set invLo = GetInvSysTable()
    If Not invLo Is Nothing Then
        If itemRow > 0 Then invIdx = FindInvRowIndexByRow(invLo, itemRow)
        If invIdx = 0 And Len(Trim$(itemName)) > 0 Then invIdx = FindInvRowIndexByItem(invLo, itemName)
    End If

    actualRow = itemRow
    actualItem = itemName
    actualUom = uom
    actualLoc = location
    actualDesc = description

    If invIdx > 0 Then
        actualRow = NzLng(GetInvSysValueByIndex(invLo, invIdx, "ROW"))
        actualItem = NzStr(GetInvSysValueByIndex(invLo, invIdx, "ITEM"))
        actualUom = NzStr(GetInvSysValueByIndex(invLo, invIdx, "UOM"))
        actualLoc = NzStr(GetInvSysValueByIndex(invLo, invIdx, "LOCATION"))
        actualDesc = NzStr(GetInvSysValueByIndex(invLo, invIdx, "DESCRIPTION"))
    End If

    If actualRow <= 0 Then actualRow = itemRow
    If Len(actualItem) = 0 Then actualItem = itemName
    If Len(actualUom) = 0 Then actualUom = uom
    If Len(actualLoc) = 0 Then actualLoc = location
    If Len(actualDesc) = 0 Then actualDesc = description

    If Not invLo Is Nothing Then
        If actualRow > 0 And actualItem <> "" Then
            actualRow = EnsureInvSysItem(actualItem, actualUom, actualLoc, actualDesc, invLo, actualRow)
        End If
    End If

    Application.EnableEvents = False
    WriteValue loBuilder.ListRows(1), "Box Name", actualItem
    WriteValue loBuilder.ListRows(1), "UOM", actualUom
    WriteValue loBuilder.ListRows(1), "LOCATION", actualLoc
    WriteValue loBuilder.ListRows(1), "DESCRIPTION", actualDesc
    Application.EnableEvents = True

    If actualRow > 0 Then
        If LoadBoxMakerBomForPackage(ws, actualRow, NzDbl(ValueFromTable(loBuilder, "Quantity")), report) Then
            ShowShippingStatus report
        ElseIf report <> "" Then
            ShowShippingStatus report
        End If
    End If
    RefreshBoxMakerCurrentInventory ws
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    MsgBox "ApplyItemToBoxBuilder error: " & Err.Description, vbCritical
End Sub

Private Sub ReloadBoxMakerBomFromBuilder(ByVal ws As Worksheet)
    On Error GoTo CleanExit

    Dim loBuilder As ListObject
    Dim invLo As ListObject
    Dim boxQty As Double
    Dim packageRow As Long
    Dim report As String

    If ws Is Nothing Then Exit Sub
    If Not IsBoxMakerMode(ws) Then Exit Sub

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then Exit Sub
    If loBuilder.DataBodyRange Is Nothing Then Exit Sub

    boxQty = NzDbl(ValueFromTable(loBuilder, "Quantity"))

    Set invLo = GetInvSysTableFromWorkbook(ws.Parent)
    If invLo Is Nothing Then Set invLo = GetInvSysTable()

    packageRow = ResolveBoxMakerPackageRow(loBuilder, invLo)
    If packageRow <= 0 Then Exit Sub

    If LoadBoxMakerBomForPackage(ws, packageRow, boxQty, report) Then
        ShowShippingStatus report
    ElseIf report <> "" Then
        ShowShippingStatus report
    End If
    RefreshBoxMakerCurrentInventory ws

CleanExit:
End Sub

Private Sub RecalculateBoxMakerBomFromBuilder(ByVal ws As Worksheet, _
                                              ByVal invLo As ListObject, _
                                              ByVal loBuilder As ListObject)
    On Error GoTo CleanExit

    Dim packageRow As Long
    Dim boxQty As Double
    Dim report As String

    If ws Is Nothing Or loBuilder Is Nothing Then Exit Sub
    If Not IsBoxMakerMode(ws) Then Exit Sub
    If loBuilder.DataBodyRange Is Nothing Then Exit Sub

    boxQty = NzDbl(ValueFromTable(loBuilder, "Quantity"))
    packageRow = ResolveBoxMakerPackageRow(loBuilder, invLo)
    If packageRow <= 0 Then Exit Sub

    LoadBoxMakerBomForPackage ws, packageRow, boxQty, report
    RefreshBoxMakerCurrentInventory ws

CleanExit:
End Sub

Private Function ResolveBoxMakerPackageRow(ByVal loBuilder As ListObject, ByVal invLo As ListObject) As Long
    Dim boxName As String
    Dim invIdx As Long
    Dim runtimeMax As Long

    If loBuilder Is Nothing Then Exit Function
    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    If boxName = "" Then Exit Function

    If Not invLo Is Nothing Then
        invIdx = FindInvRowIndexByItem(invLo, boxName)
        If invIdx > 0 Then
            ResolveBoxMakerPackageRow = NzLng(GetInvSysValueByIndex(invLo, invIdx, "ROW"))
            If ResolveBoxMakerPackageRow > 0 Then Exit Function
        End If
    End If

    ResolveBoxMakerPackageRow = FindShippingBomPackageRowByName(loBuilder.Parent.Parent, boxName, runtimeMax)
End Function

Private Function LoadBoxMakerBomForPackage(ByVal ws As Worksheet, _
                                           ByVal packageRow As Long, _
                                           ByVal packageQty As Double, _
                                           ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim loBom As ListObject
    Dim cPackageRow As Long
    Dim cComponentRow As Long
    Dim cComponentItem As Long
    Dim cComponentQty As Long
    Dim cComponentUom As Long
    Dim cComponentLocation As Long
    Dim cComponentDescription As Long
    Dim scaleQty As Double
    Dim arr As Variant
    Dim r As Long
    Dim outRow As Long
    Dim loView As ListObject
    Dim refreshReport As String
    Dim preservedCurrentInv As Object

    report = ""
    If ws Is Nothing Then Exit Function
    If packageRow <= 0 Then
        report = "Selected shippable has no invSys ROW."
        Exit Function
    End If

    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then
        report = "BoxBOM table was not found."
        Exit Function
    End If
    EnsureBoxBomEntryColumns loBom

    If Not TryLoadRuntimeShippingBomRows(arr, _
                                         cPackageRow, _
                                         cComponentRow, _
                                         cComponentItem, _
                                         cComponentQty, _
                                         cComponentUom, _
                                         cComponentLocation, _
                                         cComponentDescription, _
                                         report) Then
        Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
        If loView Is Nothing Then
            If report = "" Then report = "ShippingBOMView table was not found."
            Exit Function
        End If
        If loView.DataBodyRange Is Nothing Then RefreshShippingBomViewForWorkbook ws.Parent, refreshReport
        If loView.DataBodyRange Is Nothing Then
            If report = "" Then report = "No saved Shipping BOM rows are available for the selected warehouse."
            Exit Function
        End If

        cPackageRow = ColumnIndex(loView, "PackageRow")
        cComponentRow = ColumnIndex(loView, "ComponentRow")
        cComponentItem = ColumnIndex(loView, "ComponentItem")
        cComponentQty = ColumnIndex(loView, "ComponentQty")
        cComponentUom = ColumnIndex(loView, "ComponentUOM")
        cComponentLocation = ColumnIndex(loView, "ComponentLocation")
        cComponentDescription = ColumnIndex(loView, "ComponentDescription")
        If cPackageRow = 0 Or cComponentRow = 0 Or cComponentQty = 0 Then
            report = "ShippingBOMView is missing required PackageRow/Component columns."
            Exit Function
        End If
        arr = loView.DataBodyRange.Value
    End If

    scaleQty = packageQty
    If scaleQty <= 0 Then scaleQty = 1#

    Set preservedCurrentInv = CaptureBoxBomCurrentInventory(loBom)

    ClearListObjectData loBom
    EnsureBoxBomStarterRows loBom

    For r = 1 To UBound(arr, 1)
        If NzLng(arr(r, cPackageRow)) <> packageRow Then GoTo NextBomRow

        outRow = outRow + 1
        Do While loBom.ListRows.Count < outRow
            loBom.ListRows.Add
        Loop

        SetTableCellShipping loBom, outRow, COL_BOXBOM_ITEM, ValueFromArrayColumn(arr, r, cComponentItem)
        SetTableCellShipping loBom, outRow, "ROW", NzLng(arr(r, cComponentRow))
        SetTableCellShipping loBom, outRow, "QUANTITY", NzDbl(arr(r, cComponentQty)) * scaleQty
        SetTableCellShipping loBom, outRow, "UOM", ValueFromArrayColumn(arr, r, cComponentUom)
        SetTableCellShipping loBom, outRow, "LOCATION", ValueFromArrayColumn(arr, r, cComponentLocation)
        SetTableCellShipping loBom, outRow, "DESCRIPTION", ValueFromArrayColumn(arr, r, cComponentDescription)
NextBomRow:
    Next r

    EnsureBoxBomStarterRows loBom
    If outRow = 0 Then
        report = "No saved BoxBOM components were found for invSys ROW " & CStr(packageRow) & "."
        Exit Function
    End If

    LoadBoxMakerBomForPackage = True
    report = "Loaded BoxBOM for invSys ROW " & CStr(packageRow) & " (" & CStr(outRow) & " component row(s))."
    RefreshBoxMakerCurrentInventory ws
    RestorePreservedBoxBomCurrentInventory loBom, preservedCurrentInv
    Exit Function

FailSoft:
    report = "LoadBoxMakerBomForPackage failed: " & Err.Description
End Function

Private Function CaptureBoxBomCurrentInventory(ByVal loBom As ListObject) As Object
    Dim result As Object
    Dim cInv As Long
    Dim cRow As Long
    Dim cItem As Long
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String
    Dim currentVal As Variant

    Set result = CreateObject("Scripting.Dictionary")
    Set CaptureBoxBomCurrentInventory = result

    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function
    cInv = ColumnIndex(loBom, COL_CURRENT_INV)
    If cInv = 0 Then Exit Function

    cRow = ColumnIndex(loBom, "ROW")
    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    For r = 1 To loBom.ListRows.Count
        currentVal = loBom.DataBodyRange.Cells(r, cInv).Value
        If IsBlankInventoryValue(currentVal) Then GoTo NextRow

        rowVal = 0
        itemName = ""
        If cRow > 0 Then rowVal = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        If cItem > 0 Then itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))

        If rowVal > 0 Then result("ROW:" & CStr(rowVal)) = currentVal
        If itemName <> "" Then result("ITEM:" & LCase$(itemName)) = currentVal
NextRow:
    Next r
End Function

Private Sub RestorePreservedBoxBomCurrentInventory(ByVal loBom As ListObject, ByVal preserved As Object)
    Dim cInv As Long
    Dim cRow As Long
    Dim cItem As Long
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String
    Dim key As String

    If loBom Is Nothing Then Exit Sub
    If preserved Is Nothing Then Exit Sub
    If preserved.Count = 0 Then Exit Sub
    If loBom.DataBodyRange Is Nothing Then Exit Sub

    cInv = ColumnIndex(loBom, COL_CURRENT_INV)
    If cInv = 0 Then Exit Sub
    cRow = ColumnIndex(loBom, "ROW")
    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)

    For r = 1 To loBom.ListRows.Count
        If Not IsBlankInventoryValue(loBom.DataBodyRange.Cells(r, cInv).Value) Then GoTo NextRow

        rowVal = 0
        itemName = ""
        If cRow > 0 Then rowVal = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        If cItem > 0 Then itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))

        If rowVal > 0 Then
            key = "ROW:" & CStr(rowVal)
            If preserved.Exists(key) Then
                loBom.DataBodyRange.Cells(r, cInv).Value = preserved(key)
                GoTo NextRow
            End If
        End If
        If itemName <> "" Then
            key = "ITEM:" & LCase$(itemName)
            If preserved.Exists(key) Then loBom.DataBodyRange.Cells(r, cInv).Value = preserved(key)
        End If
NextRow:
    Next r
    FormatCurrentInventoryColumn loBom
End Sub

Private Function ValueFromArrayColumn(ByRef arr As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    If colIndex <= 0 Then Exit Function
    ValueFromArrayColumn = arr(rowIndex, colIndex)
End Function

Private Function TryLoadRuntimeShippingBomRows(ByRef arr As Variant, _
                                               ByRef cPackageRow As Long, _
                                               ByRef cComponentRow As Long, _
                                               ByRef cComponentItem As Long, _
                                               ByRef cComponentQty As Long, _
                                               ByRef cComponentUom As Long, _
                                               ByRef cComponentLocation As Long, _
                                               ByRef cComponentDescription As Long, _
                                               ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim openedTransient As Boolean

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "No connected warehouse target was available for ShippingBOM runtime lookup."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        report = "Connected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbBom Is Nothing Then GoTo CleanExit

    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit
    If loBom.DataBodyRange Is Nothing Then
        report = "Shipping BOM runtime workbook has no saved package rows."
        GoTo CleanExit
    End If

    cPackageRow = ColumnIndex(loBom, "PackageRow")
    cComponentRow = ColumnIndex(loBom, "ComponentRow")
    cComponentItem = ColumnIndex(loBom, "ComponentItem")
    cComponentQty = ColumnIndex(loBom, "ComponentQty")
    cComponentUom = ColumnIndex(loBom, "ComponentUOM")
    cComponentLocation = ColumnIndex(loBom, "ComponentLocation")
    cComponentDescription = ColumnIndex(loBom, "ComponentDescription")
    If cPackageRow = 0 Or cComponentRow = 0 Or cComponentQty = 0 Then
        report = "Shipping BOM runtime workbook is missing required PackageRow/Component columns."
        GoTo CleanExit
    End If

    arr = loBom.DataBodyRange.Value
    TryLoadRuntimeShippingBomRows = True
    report = ""

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "Shipping BOM runtime lookup failed: " & Err.Description
    Resume CleanExit
End Function

Private Function RepairInvSysRowKeyShipping(ByVal invLo As ListObject, ByVal invIdx As Long, ByVal colRowInv As Long) As Long
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function
    If invIdx <= 0 Or invIdx > invLo.ListRows.Count Then Exit Function

    RepairInvSysRowKeyShipping = invIdx
    If colRowInv > 0 Then
        On Error Resume Next
        invLo.DataBodyRange.Cells(invIdx, colRowInv).Value = RepairInvSysRowKeyShipping
        On Error GoTo 0
    End If
End Function

Private Sub RepairBoxBomRowsFromInventory(ByVal loBom As ListObject)
    Dim invLo As ListObject
    Dim cItem As Long
    Dim cRow As Long
    Dim r As Long
    Dim itemName As String
    Dim rowValue As Long
    Dim invIdx As Long
    Dim invRowCol As Long

    If loBom Is Nothing Then Exit Sub
    If loBom.DataBodyRange Is Nothing Then Exit Sub

    Set invLo = GetInvSysTable()
    If invLo Is Nothing Then Exit Sub

    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    cRow = ColumnIndex(loBom, "ROW")
    invRowCol = ColumnIndex(invLo, "ROW")
    If cItem = 0 Or cRow = 0 Then Exit Sub

    For r = 1 To loBom.ListRows.Count
        rowValue = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        If rowValue <= 0 Then
            itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))
            If itemName <> "" Then
                invIdx = FindInvRowIndexByItem(invLo, itemName)
                If invIdx > 0 Then
                    rowValue = RepairInvSysRowKeyShipping(invLo, invIdx, invRowCol)
                    If rowValue > 0 Then loBom.DataBodyRange.Cells(r, cRow).Value = rowValue
                End If
            End If
        End If
    Next r
End Sub

Private Function BuildBoxMakerAggregateTables(ByVal loBuilder As ListObject, _
                                             ByVal loBom As ListObject, _
                                             ByVal invLo As ListObject, _
                                             ByVal loAggBom As ListObject, _
                                             ByVal loAggPack As ListObject, _
                                             ByRef errNotes As String) As Boolean
    On Error GoTo FailBuild

    Dim boxName As String
    Dim boxQty As Double
    Dim packageIdx As Long
    Dim componentCount As Long
    Dim aggBomRow As Long
    Dim aggPackRow As Long
    Dim cItem As Long
    Dim cRow As Long
    Dim cQty As Long
    Dim r As Long
    Dim itemName As String
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim invIdx As Long
    Dim runtimeMax As Long
    Dim packageRow As Long
    Dim stepName As String

    errNotes = ""
    stepName = "check required tables"
    If loBuilder Is Nothing Or loBom Is Nothing Or invLo Is Nothing Or loAggBom Is Nothing Or loAggPack Is Nothing Then
        errNotes = "BoxMaker required tables are missing."
        Exit Function
    End If

    stepName = "ensure BoxMaker columns"
    EnsureColumnExists loBuilder, "Quantity", "Box Name"
    EnsureBoxBomEntryColumns loBom

    stepName = "read BoxBuilder values"
    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    boxQty = NzDbl(ValueFromTable(loBuilder, "Quantity"))
    If boxName = "" Then
        errNotes = "BoxBuilder Box Name is required."
        Exit Function
    End If
    If boxQty <= 0 Then
        errNotes = "BoxBuilder Quantity must be greater than zero."
        Exit Function
    End If

    stepName = "resolve package row in invSys"
    packageIdx = FindInvRowIndexByItem(invLo, boxName)
    If packageIdx <= 0 Then
        stepName = "resolve package row in ShippingBOM runtime"
        packageRow = FindShippingBomPackageRowByName(loBuilder.Parent.Parent, boxName, runtimeMax)
        If packageRow > 0 Then
            stepName = "ensure package row in invSys"
            EnsureInvSysItem boxName, _
                             Trim$(NzStr(ValueFromTable(loBuilder, "UOM"))), _
                             Trim$(NzStr(ValueFromTable(loBuilder, "LOCATION"))), _
                             Trim$(NzStr(ValueFromTable(loBuilder, "DESCRIPTION"))), _
                             invLo, _
                             packageRow
            packageIdx = FindInvRowIndexByItem(invLo, boxName)
        End If
    End If
    If packageIdx <= 0 Then
        errNotes = "Box '" & boxName & "' was not found in invSys or ShippingBOM runtime."
        Exit Function
    End If

    stepName = "clear aggregate tables"
    ClearListObjectData loAggBom
    ClearListObjectData loAggPack
    stepName = "append package aggregate row"
    AppendAggregateRowFromInventory loAggPack, invLo, packageIdx, boxQty, aggPackRow

    stepName = "resolve BoxBOM columns"
    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    cRow = ColumnIndex(loBom, "ROW")
    cQty = ColumnIndex(loBom, "QUANTITY")
    If cItem = 0 Or cRow = 0 Or cQty = 0 Then
        errNotes = "BoxBOM must include ITEM, ROW, and QUANTITY columns."
        Exit Function
    End If
    If loBom.DataBodyRange Is Nothing Then
        errNotes = "BoxBOM has no component rows."
        Exit Function
    End If

    For r = 1 To loBom.ListRows.Count
        stepName = "read BoxBOM row " & CStr(r)
        itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))
        rowVal = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        qtyVal = NzDbl(loBom.DataBodyRange.Cells(r, cQty).Value)
        If BoxMakerComponentRowIsBlank(itemName, rowVal, qtyVal) Then GoTo NextComponent
        If qtyVal <= 0 Then
            errNotes = "BoxBOM row " & CStr(r) & " needs a component Quantity greater than zero."
            Exit Function
        End If

        stepName = "resolve BoxBOM component row " & CStr(r)
        invIdx = 0
        If rowVal > 0 Then invIdx = FindInvRowIndexByRow(invLo, rowVal)
        If invIdx <= 0 And itemName <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
        If invIdx <= 0 Then
            errNotes = "BoxBOM component row " & CStr(r) & " was not found in invSys."
            Exit Function
        End If

        stepName = "append component aggregate row " & CStr(r)
        AppendAggregateRowFromInventory loAggBom, invLo, invIdx, qtyVal, aggBomRow
        componentCount = componentCount + 1
NextComponent:
    Next r

    If componentCount = 0 Then
        errNotes = "BoxBOM has no component quantities to deduct."
        Exit Function
    End If

    BuildBoxMakerAggregateTables = True
    Exit Function

FailBuild:
    errNotes = "BuildBoxMakerAggregateTables failed during " & stepName
    If r > 0 Then errNotes = errNotes & " (BoxBOM row " & CStr(r) & ")"
    errNotes = errNotes & ": " & Err.Description
End Function

Private Function BoxMakerComponentRowIsBlank(ByVal itemName As String, _
                                             ByVal rowVal As Long, _
                                             ByVal qtyVal As Double) As Boolean
    If itemName <> "" Then Exit Function
    If rowVal <> 0 Then Exit Function
    If qtyVal <> 0 Then Exit Function
    BoxMakerComponentRowIsBlank = True
End Function

Private Sub AppendAggregateRowFromInventory(ByVal loTarget As ListObject, _
                                           ByVal invLo As ListObject, _
                                           ByVal invIdx As Long, _
                                           ByVal qtyVal As Double, _
                                           ByRef nextRow As Long)
    Dim lr As ListRow

    If loTarget Is Nothing Or invLo Is Nothing Then Exit Sub
    If invIdx <= 0 Or invLo.DataBodyRange Is Nothing Then Exit Sub
    EnsureColumnExists loTarget, "ROW"
    EnsureColumnExists loTarget, "ITEM_CODE"
    EnsureColumnExists loTarget, "ITEM"
    EnsureColumnExists loTarget, "QUANTITY"
    EnsureColumnExists loTarget, "UOM"
    EnsureColumnExists loTarget, "LOCATION"

    nextRow = nextRow + 1
    Do While loTarget.ListRows.Count < nextRow
        loTarget.ListRows.Add
    Loop
    Set lr = loTarget.ListRows(nextRow)

    SetTableCellShipping loTarget, lr.Index, "ROW", GetInvSysValueByIndex(invLo, invIdx, "ROW")
    SetTableCellShipping loTarget, lr.Index, "ITEM_CODE", GetInvSysValueByIndex(invLo, invIdx, "ITEM_CODE")
    SetTableCellShipping loTarget, lr.Index, "ITEM", GetInvSysValueByIndex(invLo, invIdx, "ITEM")
    SetTableCellShipping loTarget, lr.Index, "QUANTITY", qtyVal
    SetTableCellShipping loTarget, lr.Index, "UOM", GetInvSysValueByIndex(invLo, invIdx, "UOM")
    SetTableCellShipping loTarget, lr.Index, "LOCATION", GetInvSysValueByIndex(invLo, invIdx, "LOCATION")
End Sub

Private Function GetInvSysValueByIndex(ByVal invLo As ListObject, ByVal invIdx As Long, ByVal columnName As String) As Variant
    Dim colIdx As Long

    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function
    If invIdx <= 0 Or invIdx > invLo.ListRows.Count Then Exit Function
    colIdx = ColumnIndex(invLo, columnName)
    If colIdx = 0 Then Exit Function
    GetInvSysValueByIndex = invLo.DataBodyRange.Cells(invIdx, colIdx).Value
End Function

Private Function ApplyBoxCreatedFromAggregates(ByVal invLo As ListObject, _
                                              ByVal loAggBom As ListObject, _
                                              ByVal loAggPack As ListObject, _
                                              ByRef usedTotal As Double, _
                                              ByRef madeTotal As Double, _
                                              ByRef errNotes As String) As Boolean
    Dim shortage As String
    Dim compLogs As New Collection
    Dim pkgLogs As New Collection
    Dim usedDeltas As Collection
    Dim madeDeltas As Collection
    Dim madeStaged As Double

    errNotes = ""
    If Not ValidateComponentInventory(invLo, loAggBom, shortage) Then
        errNotes = shortage
        Exit Function
    End If

    Set usedDeltas = BuildComponentDeltaPacketFromAggregate(invLo, loAggBom, errNotes)
    If usedDeltas Is Nothing Then Exit Function
    Set madeDeltas = BuildMadeDeltaPacket(invLo, loAggPack, errNotes)
    If madeDeltas Is Nothing Then Exit Function

    PrepareComponentLogEntries invLo, usedDeltas, compLogs
    PreparePackageLogEntries invLo, madeDeltas, pkgLogs

    usedTotal = ApplyUsedDeltasLocal(invLo, usedDeltas, errNotes)
    If usedTotal < 0 Then Exit Function

    madeStaged = ApplyMadeDeltasLocal(invLo, madeDeltas, errNotes)
    If madeStaged < 0 Then Exit Function

    madeTotal = ApplyMadeToInventoryDeltasLocal(invLo, madeDeltas, errNotes)
    If madeTotal < 0 Then Exit Function

    If compLogs.Count > 0 Then LogShippingChanges "AggregateBoxBOM_Log", compLogs
    If pkgLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", pkgLogs
    ApplyBoxCreatedFromAggregates = True
End Function

Private Function ApplyBoxCreatedFromBuilder(ByVal loBuilder As ListObject, _
                                           ByVal loBom As ListObject, _
                                           ByVal invLo As ListObject, _
                                           ByRef usedTotal As Double, _
                                           ByRef madeTotal As Double, _
                                           ByRef errNotes As String) As Boolean
    Dim eventIdOut As String
    Dim runtimeReport As String

    errNotes = ""
    If loBuilder Is Nothing Or loBom Is Nothing Or invLo Is Nothing Then
        errNotes = "Box Created required tables are missing."
        Exit Function
    End If
    If invLo.DataBodyRange Is Nothing Then
        errNotes = "invSys has no inventory rows."
        Exit Function
    End If

    EnsureColumnExists loBuilder, "Quantity", "Box Name"
    EnsureBoxBomEntryColumns loBom

    If Not QueueBoxBuildEventFromBuilder(loBuilder, loBom, invLo, usedTotal, madeTotal, eventIdOut, errNotes) Then Exit Function

    If Not modOperatorReadModel.RunBatchAndRefreshOperatorWorkbook(loBuilder.Parent.Parent, "", "LOCAL", runtimeReport) Then
        If runtimeReport = "" Then runtimeReport = "Box build event queued, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If
    If eventIdOut <> "" Then AppendNote errNotes, "Inbox EventID: " & eventIdOut

    ApplyBoxCreatedFromBuilder = True
End Function

Private Function ApplyBoxUnboxedFromBuilder(ByVal loBuilder As ListObject, _
                                           ByVal loBom As ListObject, _
                                           ByVal invLo As ListObject, _
                                           ByRef packageReturned As Double, _
                                           ByRef componentsReturned As Double, _
                                           ByRef errNotes As String) As Boolean
    Dim boxName As String
    Dim boxQty As Double
    Dim packageIdx As Long
    Dim componentCount As Long
    Dim cItem As Long
    Dim cRow As Long
    Dim cQty As Long
    Dim cTotal As Long
    Dim cLastEdited As Long
    Dim cTotalLastEdit As Long
    Dim r As Long
    Dim itemName As String
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim invIdx As Long
    Dim packageTotal As Double
    Dim runtimeMax As Long
    Dim packageRow As Long

    errNotes = ""
    If loBuilder Is Nothing Or loBom Is Nothing Or invLo Is Nothing Then
        errNotes = "Box Unboxed required tables are missing."
        Exit Function
    End If
    If invLo.DataBodyRange Is Nothing Then
        errNotes = "invSys has no inventory rows."
        Exit Function
    End If

    EnsureColumnExists loBuilder, "Quantity", "Box Name"
    EnsureBoxBomEntryColumns loBom

    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    boxQty = NzDbl(ValueFromTable(loBuilder, "Quantity"))
    If boxName = "" Then
        errNotes = "BoxBuilder Box Name is required."
        Exit Function
    End If
    If boxQty <= 0 Then
        errNotes = "BoxBuilder Quantity must be greater than zero."
        Exit Function
    End If

    packageIdx = FindInvRowIndexByItem(invLo, boxName)
    If packageIdx <= 0 Then
        packageRow = FindShippingBomPackageRowByName(loBuilder.Parent.Parent, boxName, runtimeMax)
        If packageRow > 0 Then
            EnsureInvSysItem boxName, _
                             Trim$(NzStr(ValueFromTable(loBuilder, "UOM"))), _
                             Trim$(NzStr(ValueFromTable(loBuilder, "LOCATION"))), _
                             Trim$(NzStr(ValueFromTable(loBuilder, "DESCRIPTION"))), _
                             invLo, _
                             packageRow
            packageIdx = FindInvRowIndexByItem(invLo, boxName)
        End If
    End If
    If packageIdx <= 0 Then
        errNotes = "Box '" & boxName & "' was not found in invSys or ShippingBOM runtime."
        Exit Function
    End If

    cTotal = ColumnIndex(invLo, "TOTAL INV")
    cLastEdited = ColumnIndex(invLo, "LAST EDITED")
    cTotalLastEdit = ColumnIndex(invLo, "TOTAL INV LAST EDIT")
    If cTotal = 0 Then
        errNotes = "invSys table missing TOTAL INV column."
        Exit Function
    End If

    packageTotal = NzDbl(invLo.DataBodyRange.Cells(packageIdx, cTotal).Value)
    If boxQty > packageTotal + 0.0000001 Then
        errNotes = "Box '" & boxName & "' only has " & Format$(packageTotal, "0.###") & " in TOTAL INV but needs " & Format$(boxQty, "0.###") & "."
        Exit Function
    End If

    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    cRow = ColumnIndex(loBom, "ROW")
    cQty = ColumnIndex(loBom, "QUANTITY")
    If cItem = 0 Or cRow = 0 Or cQty = 0 Then
        errNotes = "BoxBOM must include ITEM, ROW, and QUANTITY columns."
        Exit Function
    End If
    If loBom.DataBodyRange Is Nothing Then
        errNotes = "BoxBOM has no component rows."
        Exit Function
    End If

    For r = 1 To loBom.ListRows.Count
        itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))
        rowVal = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        qtyVal = NzDbl(loBom.DataBodyRange.Cells(r, cQty).Value)
        If itemName = "" And rowVal = 0 And qtyVal = 0 Then GoTo NextValidate
        If qtyVal <= 0 Then
            errNotes = "BoxBOM row " & CStr(r) & " needs a component Quantity greater than zero."
            Exit Function
        End If
        invIdx = 0
        If rowVal > 0 Then invIdx = FindInvRowIndexByRow(invLo, rowVal)
        If invIdx <= 0 And itemName <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
        If invIdx <= 0 Then
            errNotes = "BoxBOM component row " & CStr(r) & " was not found in invSys."
            Exit Function
        End If
        componentCount = componentCount + 1
NextValidate:
    Next r

    If componentCount = 0 Then
        errNotes = "BoxBOM has no component quantities to return."
        Exit Function
    End If

    invLo.DataBodyRange.Cells(packageIdx, cTotal).Value = packageTotal - boxQty
    If cLastEdited > 0 Then invLo.DataBodyRange.Cells(packageIdx, cLastEdited).Value = Now
    If cTotalLastEdit > 0 Then invLo.DataBodyRange.Cells(packageIdx, cTotalLastEdit).Value = Now
    packageReturned = boxQty

    For r = 1 To loBom.ListRows.Count
        itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))
        rowVal = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        qtyVal = NzDbl(loBom.DataBodyRange.Cells(r, cQty).Value)
        If itemName = "" And rowVal = 0 And qtyVal = 0 Then GoTo NextApply
        invIdx = 0
        If rowVal > 0 Then invIdx = FindInvRowIndexByRow(invLo, rowVal)
        If invIdx <= 0 And itemName <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
        If invIdx <= 0 Then GoTo NextApply

        invLo.DataBodyRange.Cells(invIdx, cTotal).Value = NzDbl(invLo.DataBodyRange.Cells(invIdx, cTotal).Value) + qtyVal
        If cLastEdited > 0 Then invLo.DataBodyRange.Cells(invIdx, cLastEdited).Value = Now
        If cTotalLastEdit > 0 Then invLo.DataBodyRange.Cells(invIdx, cTotalLastEdit).Value = Now
        componentsReturned = componentsReturned + qtyVal
NextApply:
    Next r

    ApplyBoxUnboxedFromBuilder = True
End Function

Private Sub ResetBoxMakerQuantities(ByVal loBuilder As ListObject, ByVal loBom As ListObject)
    Dim cQty As Long

    If Not loBuilder Is Nothing Then
        cQty = ColumnIndex(loBuilder, "Quantity")
        If cQty > 0 Then
            If Not loBuilder.DataBodyRange Is Nothing Then loBuilder.DataBodyRange.Columns(cQty).ClearContents
        End If
    End If

    If Not loBom Is Nothing Then
        cQty = ColumnIndex(loBom, "QUANTITY")
        If cQty > 0 Then
            If Not loBom.DataBodyRange Is Nothing Then loBom.DataBodyRange.Columns(cQty).ClearContents
        End If
    End If
End Sub

Private Function CollectBomComponents(loBom As ListObject, invLo As ListObject, ByRef syncNotes As String) As Collection
    Dim result As New Collection
    If loBom Is Nothing Or invLo Is Nothing Then
        Set CollectBomComponents = result
        Exit Function
    End If

    Dim cName As Long: cName = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    Dim cRow As Long: cRow = ColumnIndex(loBom, "ROW")
    Dim cQty As Long: cQty = ColumnIndex(loBom, "QUANTITY")
    Dim cUom As Long: cUom = ColumnIndex(loBom, "UOM")
    Dim cLoc As Long: cLoc = ColumnIndex(loBom, "LOCATION")
    Dim cDesc As Long: cDesc = ColumnIndex(loBom, "DESCRIPTION")
    If cName = 0 Or cRow = 0 Or cQty = 0 Or cUom = 0 Then
        MsgBox "BoxBOM table must include ITEM, ROW, QUANTITY, and UOM columns.", vbExclamation
        Exit Function
    End If

    If loBom.DataBodyRange Is Nothing Then
        Set CollectBomComponents = result
        Exit Function
    End If

    Dim invRowCol As Long: invRowCol = ColumnIndex(invLo, "ROW")
    Dim invItemCol As Long: invItemCol = ColumnIndex(invLo, "ITEM")
    Dim invUomCol As Long: invUomCol = ColumnIndex(invLo, "UOM")
    Dim invLocCol As Long: invLocCol = ColumnIndex(invLo, "LOCATION")
    Dim invDescCol As Long: invDescCol = ColumnIndex(invLo, "DESCRIPTION")
    If invRowCol = 0 Then
        MsgBox "invSys table must contain a ROW column.", vbCritical
        Exit Function
    End If

    Dim arr As Variant: arr = loBom.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim partName As String: partName = Trim$(NzStr(arr(r, cName)))
        Dim partRow As Long: partRow = NzLng(arr(r, cRow))
        Dim qty As Double: qty = NzDbl(arr(r, cQty))
        Dim uomVal As String: uomVal = Trim$(NzStr(arr(r, cUom)))

        If partName = "" And partRow = 0 And qty = 0 Then GoTo NextComponent
        If qty <= 0 Then
            Err.Raise vbObjectError + 1, , "Component row " & r & " has no quantity."
        End If

        Dim invIdx As Long
        Dim partResolvedName As String
        If partRow > 0 Then
            invIdx = FindInvRowIndexByRow(invLo, partRow)
            If invIdx = 0 And partName <> "" Then
                invIdx = FindInvRowIndexByItem(invLo, partName)
                If invIdx > 0 Then
                    Dim resolvedRow As Long
                    resolvedRow = NzLng(invLo.DataBodyRange.Cells(invIdx, invRowCol).Value)
                    If resolvedRow <> partRow Then
                        partRow = resolvedRow
                        AppendSyncMessage syncNotes, "Updated ROW for '" & partName & "' to " & resolvedRow & "."
                    End If
                End If
            End If
            If invIdx = 0 Then
                Err.Raise vbObjectError + 2, , "Component row " & partRow & " not found in invSys. Update BOM before saving."
            End If
        ElseIf partName <> "" Then
            invIdx = FindInvRowIndexByItem(invLo, partName)
            If invIdx = 0 Then
                Err.Raise vbObjectError + 3, , "Component '" & partName & "' not found in invSys."
            End If
            partRow = NzLng(invLo.DataBodyRange.Cells(invIdx, invRowCol).Value)
        Else
            Err.Raise vbObjectError + 4, , "Component row " & r & " is missing both item name and ROW."
        End If

        Dim actualUom As String, actualLoc As String, actualDesc As String
        Dim actualItem As String
        If invItemCol > 0 Then actualItem = NzStr(invLo.DataBodyRange.Cells(invIdx, invItemCol).Value)
        If invUomCol > 0 Then actualUom = NzStr(invLo.DataBodyRange.Cells(invIdx, invUomCol).Value)
        If invLocCol > 0 Then actualLoc = NzStr(invLo.DataBodyRange.Cells(invIdx, invLocCol).Value)
        If invDescCol > 0 Then actualDesc = NzStr(invLo.DataBodyRange.Cells(invIdx, invDescCol).Value)
        If actualItem <> "" Then partResolvedName = actualItem Else partResolvedName = partName
        If actualUom = "" Then actualUom = uomVal
        If StrComp(uomVal, actualUom, vbTextCompare) <> 0 Then
            AppendSyncMessage syncNotes, "UOM for '" & partResolvedName & "' reset to " & actualUom & "."
        End If
        uomVal = actualUom

        If cName > 0 And partResolvedName <> "" Then
            loBom.DataBodyRange.Cells(r, cName).Value = partResolvedName
        End If
        loBom.DataBodyRange.Cells(r, cRow).Value = partRow
        loBom.DataBodyRange.Cells(r, cUom).Value = uomVal
        If cLoc > 0 Then loBom.DataBodyRange.Cells(r, cLoc).Value = actualLoc
        If cDesc > 0 Then loBom.DataBodyRange.Cells(r, cDesc).Value = actualDesc

        Dim entry(1 To 6) As Variant
        entry(1) = partRow
        entry(2) = qty
        entry(3) = uomVal
        entry(4) = partResolvedName
        entry(5) = actualLoc
        entry(6) = actualDesc
        result.Add entry
NextComponent:
    Next

    Set CollectBomComponents = result
End Function

Private Sub EnsureBoxBomEntryColumns(loBom As ListObject)
    If loBom Is Nothing Then Exit Sub
    UnhideListObjectWorksheetColumnsShipping loBom
    RemoveBlankUnexpectedBoxBomColumnsShipping loBom
    Dim idxItem As Long
    idxItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    If idxItem = 0 Then
        idxItem = ColumnIndex(loBom, "BoxBOM")
        If idxItem > 0 Then
            loBom.ListColumns(idxItem).Name = COL_BOXBOM_ITEM
        Else
            EnsureColumnExists loBom, COL_BOXBOM_ITEM
        End If
    End If
    EnsureColumnExists loBom, "ITEM_CODE", COL_BOXBOM_ITEM
    EnsureColumnExists loBom, "QUANTITY", "ITEM_CODE"
    EnsureColumnExists loBom, "ROW"
    EnsureColumnExists loBom, "UOM"
    EnsureColumnExists loBom, "LOCATION"
    EnsureColumnExists loBom, "DESCRIPTION"
    UnhideListObjectWorksheetColumnsShipping loBom
End Sub

Private Sub HideListColumnShipping(ByVal lo As ListObject, ByVal columnName As String, ByVal hidden As Boolean)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    idx = ColumnIndex(lo, columnName)
    If idx = 0 Then Exit Sub
    If hidden Then Exit Sub

    On Error Resume Next
    lo.ListColumns(idx).Range.EntireColumn.Hidden = hidden
    On Error GoTo 0
End Sub

Private Sub UnhideListObjectWorksheetColumnsShipping(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    lo.Range.EntireColumn.Hidden = False
    On Error GoTo 0
End Sub

Private Sub RemoveBlankUnexpectedBoxBomColumnsShipping(ByVal loBom As ListObject)
    Dim i As Long
    Dim headerText As String

    If loBom Is Nothing Then Exit Sub
    For i = loBom.ListColumns.Count To 1 Step -1
        headerText = Trim$(NzStr(loBom.ListColumns(i).Name))
        If Not IsAllowedBoxBomHeaderShipping(headerText) Then
            If ListColumnDataIsBlankShipping(loBom.ListColumns(i)) Then
                loBom.ListColumns(i).Delete
            End If
        End If
    Next i
End Sub

Private Function IsAllowedBoxBomHeaderShipping(ByVal headerText As String) As Boolean
    Select Case UCase$(Trim$(headerText))
        Case "ITEM", "BOXBOM", "ITEM_CODE", "ROW", "QUANTITY", "CURRENT INV", "UOM", "LOCATION", "DESCRIPTION"
            IsAllowedBoxBomHeaderShipping = True
    End Select
End Function

Private Function ListColumnDataIsBlankShipping(ByVal listCol As ListColumn) As Boolean
    Dim cell As Range

    On Error GoTo CleanFail
    If listCol Is Nothing Then Exit Function
    If listCol.DataBodyRange Is Nothing Then
        ListColumnDataIsBlankShipping = True
        Exit Function
    End If

    For Each cell In listCol.DataBodyRange.Cells
        If Trim$(NzStr(cell.Value)) <> "" Then Exit Function
    Next cell
    ListColumnDataIsBlankShipping = True
    Exit Function

CleanFail:
    ListColumnDataIsBlankShipping = False
End Function

Private Sub RemoveColumnIfExistsShipping(ByVal lo As ListObject, ByVal colName As String)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    idx = ColumnIndex(lo, colName)
    If idx > 0 Then lo.ListColumns(idx).Delete
End Sub

Private Sub EnsureColumnExists(lo As ListObject, colName As String, Optional afterColumn As String = "")
    If lo Is Nothing Then Exit Sub
    If ColumnIndex(lo, colName) > 0 Then Exit Sub
    Dim insertPos As Long
    If afterColumn <> "" Then insertPos = ColumnIndex(lo, afterColumn)
    Dim newCol As ListColumn

    EnsureListObjectColumnInsertSpaceShipping lo, 1

    On Error Resume Next
    If insertPos > 0 Then
        Set newCol = lo.ListColumns.Add(insertPos + 1)
    Else
        Set newCol = lo.ListColumns.Add
    End If
    If Err.Number <> 0 Or newCol Is Nothing Then
        Err.Clear
        MoveListObjectColumnAddBlockersShipping lo, 1
        EnsureListObjectColumnInsertSpaceShipping lo, 1
        If insertPos > 0 Then
            Set newCol = lo.ListColumns.Add(insertPos + 1)
        Else
            Set newCol = lo.ListColumns.Add
        End If
    End If
    If Err.Number <> 0 Or newCol Is Nothing Then
        Err.Clear
        MoveListObjectColumnAddBlockersShipping lo, 1
        Set newCol = lo.ListColumns.Add
    End If
    On Error GoTo 0

    If newCol Is Nothing Then
        Err.Raise vbObjectError + 7142, "EnsureColumnExists", _
                  "Could not add column '" & colName & "' to " & lo.Name & _
                  ". Move adjacent worksheet tables farther away and retry Setup UI."
    End If

    newCol.Name = colName
End Sub

Private Sub PropagateBomMetadata(ws As Worksheet, comps As Collection)
    If ws Is Nothing Then Exit Sub
    If comps Is Nothing Then Exit Sub
    If comps.count = 0 Then Exit Sub
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To comps.count
        Dim info As Variant
        info = comps(i)
        Dim rowVal As Long: rowVal = NzLng(info(1))
        Dim uomVal As String: uomVal = NzStr(info(3))
        If rowVal > 0 Then
            If Not seen.Exists(rowVal) Then
                seen(rowVal) = True
                SyncSavedBomRows ws, rowVal, uomVal
            End If
        End If
    Next
End Sub

Private Sub SyncSavedBomRows(ws As Worksheet, ByVal rowValue As Long, ByVal uomValue As String)
    If ws Is Nothing Or rowValue = 0 Then Exit Sub
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        Dim cRow As Long: cRow = ColumnIndex(lo, "ROW")
        Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
        If cRow = 0 Or cUom = 0 Then GoTo NextTable
        If lo.DataBodyRange Is Nothing Then GoTo NextTable
        Dim lr As ListRow
        For Each lr In lo.ListRows
            If NzLng(lr.Range.Cells(1, cRow).Value) = rowValue Then
                lr.Range.Cells(1, cUom).Value = uomValue
            End If
        Next lr
NextTable:
    Next lo
End Sub

Private Sub AppendSyncMessage(ByRef target As String, ByVal text As String)
    If Len(text) = 0 Then Exit Sub
    If Len(target) = 0 Then
        target = text
    Else
        target = target & vbCrLf & text
    End If
End Sub

Private Function ShippingBomHeaders() As Variant
    ShippingBomHeaders = Array( _
        "PackageRow", "PackageItem", "PackageUOM", "PackageLocation", "PackageDescription", _
        "ComponentRow", "ComponentItem", "ComponentQty", "ComponentUOM", "ComponentLocation", "ComponentDescription", _
        "UpdatedAtUTC", "UpdatedBy")
End Function

Private Function ShippingBomPackageTableHeaders() As Variant
    ShippingBomPackageTableHeaders = Array( _
        "ComponentRow", "ComponentItem", "ComponentQty", "ComponentUOM", "ComponentLocation", "ComponentDescription", _
        "UpdatedAtUTC", "UpdatedBy")
End Function

Private Function SaveShippingBomToRuntime(ByVal operatorWb As Workbook, _
                                          ByVal packageRow As Long, _
                                          ByVal packageItem As String, _
                                          ByVal packageUom As String, _
                                          ByVal packageLocation As String, _
                                          ByVal packageDescription As String, _
                                          ByVal components As Collection, _
                                          ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim openedTransient As Boolean
    Dim info As Variant
    Dim lr As ListRow
    Dim i As Long
    Dim updatedAt As Date
    Dim updatedBy As String

    If packageRow <= 0 Then
        report = "Package ROW is required before saving Shipping BOM."
        Exit Function
    End If
    If components Is Nothing Then
        report = "At least one Shipping BOM component is required."
        Exit Function
    End If
    If components.Count = 0 Then
        report = "At least one Shipping BOM component is required."
        Exit Function
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "A connected warehouse target is required before saving Shipping BOM to the server."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        report = "Selected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, True, openedTransient, report)
    If wbBom Is Nothing Then Exit Function
    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit

    DeleteShippingBomPackageRows loBom, packageRow
    updatedAt = Now
    updatedBy = modRoleEventWriter.ResolveCurrentUserId()

    For i = 1 To components.Count
        info = components(i)
        Set lr = loBom.ListRows.Add
        SetTableCellShipping loBom, lr.Index, "PackageRow", packageRow
        SetTableCellShipping loBom, lr.Index, "PackageItem", packageItem
        SetTableCellShipping loBom, lr.Index, "PackageUOM", packageUom
        SetTableCellShipping loBom, lr.Index, "PackageLocation", packageLocation
        SetTableCellShipping loBom, lr.Index, "PackageDescription", packageDescription
        SetTableCellShipping loBom, lr.Index, "ComponentRow", NzLng(info(1))
        SetTableCellShipping loBom, lr.Index, "ComponentQty", NzDbl(info(2))
        SetTableCellShipping loBom, lr.Index, "ComponentUOM", NzStr(info(3))
        If UBound(info) >= 4 Then SetTableCellShipping loBom, lr.Index, "ComponentItem", NzStr(info(4))
        If UBound(info) >= 5 Then SetTableCellShipping loBom, lr.Index, "ComponentLocation", NzStr(info(5))
        If UBound(info) >= 6 Then SetTableCellShipping loBom, lr.Index, "ComponentDescription", NzStr(info(6))
        SetTableCellShipping loBom, lr.Index, "UpdatedAtUTC", updatedAt
        SetTableCellShipping loBom, lr.Index, "UpdatedBy", updatedBy
    Next i

    WriteShippingBomPackageTable wbBom, packageRow, packageItem, components, updatedAt, updatedBy
    wbBom.Save
    SaveShippingBomToRuntime = True
    report = "Shipping BOM runtime updated: " & wbBom.FullName

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "SaveShippingBomToRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Private Function RefreshShippingBomViewForWorkbook(ByVal operatorWb As Workbook, ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim loView As ListObject
    Dim openedTransient As Boolean

    If operatorWb Is Nothing Then Exit Function
    Set loView = GetShippingBomViewTable(operatorWb)
    If loView Is Nothing Then
        report = "ShippingBOMView table was not found in the operator workbook."
        Exit Function
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        ClearListObjectData loView
        report = "Shipping BOM view not refreshed: no connected warehouse target."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        ClearListObjectData loView
        report = "Shipping BOM view not refreshed: selected warehouse target is incomplete."
        Exit Function
    End If

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbBom Is Nothing Then
        ClearListObjectData loView
        Exit Function
    End If

    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit
    CopyShippingBomTable loBom, loView
    RefreshShippingBomViewForWorkbook = True
    report = "Shipping BOM view refreshed from " & wbBom.FullName

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "RefreshShippingBomViewForWorkbook failed: " & Err.Description
    Resume CleanExit
End Function

Private Function GetShippingBomViewTable(ByVal wb As Workbook) As ListObject
    Dim ws As Worksheet

    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    Set GetShippingBomViewTable = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
End Function

Private Function OpenShippingBomWorkbook(ByVal warehouseId As String, _
                                         ByVal rootPath As String, _
                                         ByVal createIfMissing As Boolean, _
                                         ByRef openedTransient As Boolean, _
                                         ByRef report As String) As Workbook
    On Error GoTo FailSoft

    Dim targetPath As String
    Dim wb As Workbook

    targetPath = ShippingBomWorkbookPath(warehouseId, rootPath)
    If targetPath = "" Then
        report = "Shipping BOM workbook path could not be resolved."
        Exit Function
    End If

    Set wb = FindOpenWorkbookByFullNameShipping(targetPath)
    If Not wb Is Nothing Then
        Set OpenShippingBomWorkbook = wb
        Exit Function
    End If

    If Len(Dir$(targetPath)) > 0 Then
        Set wb = Application.Workbooks.Open(Filename:=targetPath, UpdateLinks:=False, ReadOnly:=False)
        openedTransient = True
    ElseIf createIfMissing Then
        EnsureFolderRecursiveShipping GetParentFolderShipping(targetPath)
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        wb.Worksheets(1).Name = SHEET_BOM
        If EnsureShippingBomSchema(wb, report) Is Nothing Then
            CloseWorkbookNoSaveShipping wb
            Exit Function
        End If
        wb.SaveAs Filename:=targetPath, FileFormat:=50
        openedTransient = True
    Else
        report = "Shipping BOM runtime workbook was not found: " & targetPath
        Exit Function
    End If

    Set OpenShippingBomWorkbook = wb
    Exit Function

FailSoft:
    report = "OpenShippingBomWorkbook failed: " & Err.Description
End Function

Private Function EnsureShippingBomSchema(ByVal wb As Workbook, ByRef report As String) As ListObject
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim startCell As Range
    Dim dataRange As Range

    If wb Is Nothing Then Exit Function
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_BOM)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_BOM
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_CANONICAL_SHIPPING_BOM)
    On Error GoTo FailSoft

    headers = ShippingBomHeaders()
    If lo Is Nothing Then
        Set startCell = ws.Range("A1")
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = TABLE_CANONICAL_SHIPPING_BOM
        If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureColumnExists lo, CStr(headers(i))
    Next i
    Set EnsureShippingBomSchema = lo
    Exit Function

FailSoft:
    report = "EnsureShippingBomSchema failed: " & Err.Description
End Function

Private Sub DeleteShippingBomPackageRows(ByVal lo As ListObject, ByVal packageRow As Long)
    Dim cPackageRow As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    cPackageRow = ColumnIndex(lo, "PackageRow")
    If cPackageRow = 0 Then Exit Sub

    For i = lo.ListRows.Count To 1 Step -1
        If NzLng(lo.DataBodyRange.Cells(i, cPackageRow).Value) = packageRow Then
            lo.ListRows(i).Delete
        End If
    Next i
End Sub

Private Sub CopyShippingBomTable(ByVal loSource As ListObject, ByVal loTarget As ListObject)
    Dim headers As Variant
    Dim arr() As Variant
    Dim r As Long
    Dim c As Long
    Dim sourceCol As Long

    If loTarget Is Nothing Then Exit Sub
    ClearListObjectData loTarget
    If loSource Is Nothing Then Exit Sub
    If loSource.DataBodyRange Is Nothing Then Exit Sub

    headers = ShippingBomHeaders()
    ReDim arr(1 To loSource.DataBodyRange.Rows.Count, 1 To UBound(headers) - LBound(headers) + 1)
    For r = 1 To loSource.DataBodyRange.Rows.Count
        For c = LBound(headers) To UBound(headers)
            sourceCol = ColumnIndex(loSource, CStr(headers(c)))
            If sourceCol > 0 Then arr(r, c - LBound(headers) + 1) = loSource.DataBodyRange.Cells(r, sourceCol).Value
        Next c
    Next r
    WriteArrayToTable loTarget, arr
End Sub

Private Sub WriteShippingBomPackageTable(ByVal wbBom As Workbook, _
                                         ByVal packageRow As Long, _
                                         ByVal packageItem As String, _
                                         ByVal components As Collection, _
                                         ByVal updatedAt As Date, _
                                         ByVal updatedBy As String)
    On Error GoTo CleanExit

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim c As Long
    Dim info As Variant

    If wbBom Is Nothing Then Exit Sub
    If packageRow <= 0 Then Exit Sub
    If components Is Nothing Then Exit Sub
    If components.Count = 0 Then Exit Sub

    Set ws = WorkbookSheetExistsShipping(wbBom, SHEET_BOM_TABLES)
    If ws Is Nothing Then
        Set ws = wbBom.Worksheets.Add(After:=wbBom.Worksheets(wbBom.Worksheets.Count))
        ws.Name = SHEET_BOM_TABLES
    End If

    Set lo = EnsureShippingBomPackageTable(ws, packageRow, packageItem)
    If lo Is Nothing Then Exit Sub

    headers = ShippingBomPackageTableHeaders()
    ReDim arr(1 To components.Count, 1 To UBound(headers) - LBound(headers) + 1)
    For i = 1 To components.Count
        info = components(i)
        arr(i, 1) = NzLng(info(1))
        If UBound(info) >= 4 Then arr(i, 2) = NzStr(info(4))
        arr(i, 3) = NzDbl(info(2))
        arr(i, 4) = NzStr(info(3))
        If UBound(info) >= 5 Then arr(i, 5) = NzStr(info(5))
        If UBound(info) >= 6 Then arr(i, 6) = NzStr(info(6))
        arr(i, 7) = updatedAt
        arr(i, 8) = updatedBy
    Next i

    For c = LBound(headers) To UBound(headers)
        EnsureColumnExists lo, CStr(headers(c))
    Next c
    WriteArrayToTable lo, arr

CleanExit:
End Sub

Private Function EnsureShippingBomPackageTable(ByVal ws As Worksheet, ByVal packageRow As Long, ByVal packageItem As String) As ListObject
    On Error GoTo FailSoft

    Dim tableName As String
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim dataRange As Range
    Dim i As Long

    If ws Is Nothing Then Exit Function
    tableName = BomTableNameFromRow(packageRow)

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo FailSoft

    headers = ShippingBomPackageTableHeaders()
    If lo Is Nothing Then
        Set startCell = NextShippingBomPackageTableStartCell(ws)
        startCell.Offset(-1, 0).Value = "PackageRow"
        startCell.Offset(-1, 1).Value = packageRow
        startCell.Offset(-1, 2).Value = packageItem
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = tableName
        If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    Else
        lo.Range.Cells(1, 1).Offset(-1, 0).Value = "PackageRow"
        lo.Range.Cells(1, 1).Offset(-1, 1).Value = packageRow
        lo.Range.Cells(1, 1).Offset(-1, 2).Value = packageItem
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureColumnExists lo, CStr(headers(i))
    Next i
    Set EnsureShippingBomPackageTable = lo
    Exit Function

FailSoft:
End Function

Private Function NextShippingBomPackageTableStartCell(ByVal ws As Worksheet) As Range
    Dim lo As ListObject
    Dim maxRow As Long

    If ws Is Nothing Then Exit Function
    maxRow = 0
    For Each lo In ws.ListObjects
        If lo.Range.Row + lo.Range.Rows.Count + 1 > maxRow Then maxRow = lo.Range.Row + lo.Range.Rows.Count + 1
    Next lo
    If maxRow < 2 Then maxRow = 2
    Set NextShippingBomPackageTableStartCell = ws.Cells(maxRow + 2, 1)
End Function

Private Function ShippingBomWorkbookPath(ByVal warehouseId As String, ByVal rootPath As String) As String
    rootPath = NormalizeFolderPathShipping(rootPath)
    warehouseId = Trim$(warehouseId)
    If rootPath = "" Or warehouseId = "" Then Exit Function
    ShippingBomWorkbookPath = rootPath & "\" & warehouseId & ".invSys.Data.ShippingBOM.xlsb"
End Function

Private Function NormalizeFolderPathShipping(ByVal folderPath As String) As String
    NormalizeFolderPathShipping = Trim$(folderPath)
    Do While Len(NormalizeFolderPathShipping) > 1 And Right$(NormalizeFolderPathShipping, 1) = "\"
        NormalizeFolderPathShipping = Left$(NormalizeFolderPathShipping, Len(NormalizeFolderPathShipping) - 1)
    Loop
End Function

Private Function GetParentFolderShipping(ByVal fullPath As String) As String
    Dim pos As Long

    pos = InStrRev(fullPath, "\")
    If pos > 0 Then GetParentFolderShipping = Left$(fullPath, pos - 1)
End Function

Private Sub EnsureFolderRecursiveShipping(ByVal folderPath As String)
    If Trim$(folderPath) = "" Then Exit Sub
    modDeploymentPaths.EnsureFolderRecursiveManaged folderPath
End Sub

Private Function FindOpenWorkbookByFullNameShipping(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByFullNameShipping = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub CloseWorkbookNoSaveShipping(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub SetTableCellShipping(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal value As Variant)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    If rowIndex <= 0 Or rowIndex > lo.ListRows.Count Then Exit Sub
    idx = ColumnIndex(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = value
End Sub

Private Function EnsureBomTable(ws As Worksheet, ByVal boxName As String, ByVal boxRow As Long, ByRef blockRange As Range) As ListObject
    Dim targetName As String: targetName = BomTableNameFromRow(boxRow)
    Dim lo As ListObject

    ' try new naming scheme first
    On Error Resume Next
    Set lo = ws.ListObjects(targetName)
    On Error GoTo 0

    ' if not found, look for legacy table named by box name and rename it
    If lo Is Nothing Then
        Dim legacyName As String: legacyName = SafeTableName(boxName)
        If StrComp(legacyName, targetName, vbTextCompare) <> 0 Then
            On Error Resume Next
            Set lo = ws.ListObjects(legacyName)
            On Error GoTo 0
            If Not lo Is Nothing Then
                lo.Name = targetName
            End If
        End If
    End If

    If Not lo Is Nothing Then
        Set blockRange = BlockRangeFromHeader(ws, lo.HeaderRowRange.Row)
        If blockRange Is Nothing Then
            Set blockRange = lo.Range
        End If
        lo.Resize blockRange
        lo.HeaderRowRange.Cells(1, 1).Value = "ROW"
        lo.HeaderRowRange.Cells(1, 2).Value = "QUANTITY"
        lo.HeaderRowRange.Cells(1, 3).Value = "UOM"
        Set EnsureBomTable = lo
        Exit Function
    End If

    Dim startRow As Long: startRow = NextAvailableBomRow(ws)
    If startRow = 0 Then
        MsgBox "ShippingBOM sheet has no space for additional BOMs.", vbCritical
        Exit Function
    End If
    Set blockRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + SHIPPING_BOM_DATA_ROWS, SHIPPING_BOM_COLS))
    blockRange.Clear
    blockRange.Rows(1).Cells(1, 1).Value = "ROW"
    blockRange.Rows(1).Cells(1, 2).Value = "QUANTITY"
    blockRange.Rows(1).Cells(1, 3).Value = "UOM"
    Set lo = ws.ListObjects.Add(xlSrcRange, blockRange, , xlYes)
    lo.Name = targetName
    Set EnsureBomTable = lo
End Function

Private Sub WriteBomData(lo As ListObject, blockRange As Range, comps As Collection)
    If lo Is Nothing Then Exit Sub
    If blockRange Is Nothing Then Set blockRange = lo.Range
    lo.Resize blockRange
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    If comps.count = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To comps.count
        Dim info As Variant
        info = comps(i)
        lo.DataBodyRange.Cells(i, 1).Value = info(1)
        lo.DataBodyRange.Cells(i, 2).Value = info(2)
        lo.DataBodyRange.Cells(i, 3).Value = info(3)
    Next
End Sub

Private Function NextAvailableBomRow(ws As Worksheet) As Long
    Dim totalRows As Long: totalRows = ws.Rows.Count
    Dim startRow As Long
    startRow = 1
    Do
        If startRow + SHIPPING_BOM_BLOCK_ROWS - 1 > totalRows Then
            NextAvailableBomRow = 0
            Exit Function
        End If
        If IsBlockFree(ws, startRow) Then
            NextAvailableBomRow = startRow
            Exit Function
        End If
        startRow = startRow + SHIPPING_BOM_BLOCK_ROWS
    Loop
End Function

Private Function IsBlockFree(ws As Worksheet, startRow As Long) As Boolean
    Dim rg As Range
    Set rg = BlockRangeFromHeader(ws, startRow)
    If rg Is Nothing Then
        IsBlockFree = False
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If Not Intersect(lo.Range, rg) Is Nothing Then
            IsBlockFree = False
            Exit Function
        End If
    Next
    IsBlockFree = True
End Function

Private Function BlockRangeFromHeader(ws As Worksheet, startRow As Long) As Range
    On Error Resume Next
    Set BlockRangeFromHeader = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + SHIPPING_BOM_DATA_ROWS, SHIPPING_BOM_COLS))
    On Error GoTo 0
End Function

Private Function SafeTableName(ByVal sourceName As String) As String
    Dim cleaned As String
    cleaned = Trim$(sourceName)
    If cleaned = "" Then cleaned = "BOM_" & Format(Now, "yyyymmdd_hhnnss")
    Dim i As Long, ch As String, kept As String
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            kept = kept & ch
        Else
            kept = kept & "_"
        End If
    Next
    If kept = "" Then kept = "BOM_" & Format(Now, "yyyymmdd_hhnnss")
    If Not kept Like "[A-Za-z_]*" Then kept = "BOM_" & kept
    SafeTableName = kept
End Function

Private Function BomTableNameFromRow(ByVal rowValue As Long) As String
    If rowValue <= 0 Then
        BomTableNameFromRow = SafeTableName("BOM_" & Format$(Now, "yyyymmdd_hhnnss"))
    Else
        BomTableNameFromRow = "ROW_" & CStr(rowValue)
    End If
End Function

Private Function ValueFromTable(lo As ListObject, headerName As String) As Variant
    Dim colIdx As Long: colIdx = ColumnIndex(lo, headerName)
    If colIdx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ValueFromTable = lo.DataBodyRange.Cells(1, colIdx).Value
End Function

' ===== hold helpers =====
Private Sub MoveSelectionToHold(ByVal moveToHold As Boolean)
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim loShip As ListObject: Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    Dim loHold As ListObject: Set loHold = GetListObject(ws, TABLE_NOTSHIPPED)
    If loShip Is Nothing Or loHold Is Nothing Then
        MsgBox "ShipmentsTally or NotShipped table not found on ShipmentsTally sheet.", vbCritical
        Exit Sub
    End If

    Dim targetTable As ListObject
    Dim sourceTable As ListObject
    If moveToHold Then
        Set sourceTable = loShip
        Set targetTable = loHold
    Else
        Set sourceTable = loHold
        Set targetTable = loShip
    End If

    Dim rngSel As Range
    On Error Resume Next
    If Not sourceTable.DataBodyRange Is Nothing Then
        Set rngSel = Application.Intersect(Application.Selection, sourceTable.DataBodyRange)
    Else
        Set rngSel = Application.Intersect(Application.Selection, sourceTable.Range)
        If Not rngSel Is Nothing Then
            ExpandTableToSelection sourceTable, rngSel
            Set rngSel = Application.Intersect(Application.Selection, sourceTable.DataBodyRange)
        End If
    End If
    On Error GoTo 0
    If rngSel Is Nothing Then
        MsgBox "Select rows inside the " & sourceTable.Name & " table first.", vbInformation
        Exit Sub
    End If

    Dim processed As Object: Set processed = CreateObject("Scripting.Dictionary")
    Dim rowQueue As New Collection
    Dim cell As Range
    For Each cell In rngSel.Areas
        Dim r As Range
        For Each r In cell.Rows
            Dim rowIndex As Long
            rowIndex = r.Row - sourceTable.DataBodyRange.Row + 1
            If rowIndex >= 1 And rowIndex <= sourceTable.ListRows.Count Then
                If Not processed.Exists(CStr(rowIndex)) Then
                    processed.Add CStr(rowIndex), True
                    rowQueue.Add rowIndex
                End If
            End If
        Next r
    Next cell

    If rowQueue.Count = 0 Then
        MsgBox "Select rows inside the " & sourceTable.Name & " table first.", vbInformation
        Exit Sub
    End If

    Dim idx As Long
    For idx = rowQueue.Count To 1 Step -1
        HandleHoldRow sourceTable, targetTable, CLng(rowQueue(idx)), moveToHold
    Next idx

    InvalidateAggregates True
End Sub

Private Sub ExpandTableToSelection(lo As ListObject, sel As Range)
    If lo Is Nothing Then Exit Sub
    If sel Is Nothing Then Exit Sub
    If lo.HeaderRowRange Is Nothing Then Exit Sub

    Dim headerRow As Long: headerRow = lo.HeaderRowRange.Row
    Dim firstCol As Long: firstCol = lo.Range.Column
    Dim lastCol As Long: lastCol = firstCol + lo.Range.Columns.Count - 1

    Dim area As Range
    Dim lastRow As Long: lastRow = headerRow
    For Each area In sel.Areas
        Dim areaLast As Long
        areaLast = area.Row + area.Rows.Count - 1
        If areaLast > lastRow Then lastRow = areaLast
    Next area

    If lastRow <= headerRow Then Exit Sub
    On Error Resume Next
    lo.Resize lo.Parent.Range(lo.Parent.Cells(headerRow, firstCol), lo.Parent.Cells(lastRow, lastCol))
    On Error GoTo 0
End Sub

Private Sub HandleHoldRow(sourceTable As ListObject, targetTable As ListObject, rowIndex As Long, moveToHold As Boolean)
    Dim cRef As Long: cRef = ColumnIndex(sourceTable, "REF_NUMBER")
    Dim cItems As Long: cItems = ColumnIndex(sourceTable, "ITEMS")
    Dim cQty As Long: cQty = ColumnIndex(sourceTable, "QUANTITY")
    If cRef = 0 Or cItems = 0 Then
        MsgBox sourceTable.Name & " table needs REF_NUMBER and ITEMS columns.", vbCritical
        Exit Sub
    End If
    If cQty = 0 Then
        MsgBox sourceTable.Name & " table needs a QUANTITY column.", vbCritical
        Exit Sub
    End If

    Dim refVal As String: refVal = NzStr(sourceTable.DataBodyRange.Cells(rowIndex, cRef).Value)
    Dim itemVal As String: itemVal = NzStr(sourceTable.DataBodyRange.Cells(rowIndex, cItems).Value)
    Dim qtyVal As Double: qtyVal = NzDbl(sourceTable.DataBodyRange.Cells(rowIndex, cQty).Value)
    If qtyVal <= 0 Then Exit Sub

    Dim prompt As String
    Dim titleText As String
    Dim itemLabel As String
    itemLabel = itemVal
    If itemLabel = "" Then itemLabel = "(no item name)"
    If refVal <> "" Then
        titleText = "Hold quantity - REF " & refVal & " | " & itemLabel
    Else
        titleText = "Hold quantity - " & itemLabel
    End If
    If moveToHold Then
        prompt = "Enter quantity to hold for '" & itemLabel & "' (REF " & refVal & ", available " & qtyVal & "):"
    Else
        prompt = "Enter quantity to return to shipments for '" & itemLabel & "' (REF " & refVal & ", available " & qtyVal & "):"
    End If
    Dim qtyInput As Variant
    qtyInput = Application.InputBox(prompt, titleText, qtyVal, Type:=1)
    If qtyInput = False Then Exit Sub
    Dim qtyMove As Double: qtyMove = CDbl(qtyInput)
    If qtyMove <= 0 Then Exit Sub
    If qtyMove > qtyVal Then qtyMove = qtyVal

    MoveHoldRowQuantity sourceTable, targetTable, rowIndex, qtyMove
End Sub

Private Sub MoveHoldRowQuantity(sourceTable As ListObject, targetTable As ListObject, rowIndex As Long, qtyMove As Double)
    If sourceTable Is Nothing Or targetTable Is Nothing Then Exit Sub
    If rowIndex <= 0 Or rowIndex > sourceTable.ListRows.count Then Exit Sub
    If qtyMove <= 0 Then Exit Sub

    Dim cQty As Long: cQty = ColumnIndex(sourceTable, "QUANTITY")
    If cQty = 0 Then Exit Sub

    Dim qtyVal As Double: qtyVal = NzDbl(sourceTable.DataBodyRange.Cells(rowIndex, cQty).Value)
    If qtyVal <= 0 Then Exit Sub
    If qtyMove > qtyVal Then qtyMove = qtyVal

    AddOrMergeHoldRowFromSource sourceTable, targetTable, rowIndex, qtyMove

    Dim newQty As Double
    newQty = qtyVal - qtyMove
    If newQty <= 0 Then
        sourceTable.ListRows(rowIndex).Range.ClearContents
    Else
        sourceTable.DataBodyRange.Cells(rowIndex, cQty).Value = newQty
    End If
End Sub

Private Sub AddOrMergeHoldRowFromSource(sourceTable As ListObject, targetTable As ListObject, sourceRowIndex As Long, qtyMove As Double)
    If targetTable Is Nothing Then Exit Sub
    If qtyMove <= 0 Then Exit Sub
    If sourceTable Is Nothing Then Exit Sub
    If sourceRowIndex <= 0 Or sourceRowIndex > sourceTable.ListRows.count Then Exit Sub

    Dim sourceRef As Long: sourceRef = ColumnIndex(sourceTable, "REF_NUMBER")
    Dim sourceItems As Long: sourceItems = ColumnIndex(sourceTable, "ITEMS")
    If sourceRef = 0 Or sourceItems = 0 Then Exit Sub

    Dim refVal As String: refVal = NzStr(sourceTable.DataBodyRange.Cells(sourceRowIndex, sourceRef).Value)
    Dim itemVal As String: itemVal = NzStr(sourceTable.DataBodyRange.Cells(sourceRowIndex, sourceItems).Value)

    Dim cRef As Long: cRef = ColumnIndex(targetTable, "REF_NUMBER")
    Dim cItems As Long: cItems = ColumnIndex(targetTable, "ITEMS")
    Dim cQty As Long: cQty = ColumnIndex(targetTable, "QUANTITY")
    If cRef = 0 Or cItems = 0 Or cQty = 0 Then
        MsgBox targetTable.Name & " table needs REF_NUMBER, ITEMS, and QUANTITY columns.", vbCritical
        Exit Sub
    End If

    Dim lr As ListRow
    If Not targetTable.DataBodyRange Is Nothing Then
        For Each lr In targetTable.ListRows
            If StrComp(NzStr(lr.Range.Cells(1, cRef).Value), refVal, vbTextCompare) = 0 _
               And StrComp(NzStr(lr.Range.Cells(1, cItems).Value), itemVal, vbTextCompare) = 0 Then
                lr.Range.Cells(1, cQty).Value = NzDbl(lr.Range.Cells(1, cQty).Value) + qtyMove
                Exit Sub
            End If
        Next lr
    End If

    Set lr = targetTable.ListRows.Add
    Dim lc As ListColumn
    For Each lc In sourceTable.ListColumns
        Dim targetIdx As Long
        targetIdx = ColumnIndex(targetTable, CStr(lc.Name))
        If targetIdx > 0 Then
            lr.Range.Cells(1, targetIdx).Value = sourceTable.DataBodyRange.Cells(sourceRowIndex, lc.Index).Value
        End If
    Next lc
    lr.Range.Cells(1, cQty).Value = qtyMove
End Sub

Private Function FindHoldRowIndex(lo As ListObject, ByVal refNumber As String, ByVal itemText As String) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim cRef As Long: cRef = ColumnIndex(lo, "REF_NUMBER")
    Dim cItems As Long: cItems = ColumnIndex(lo, "ITEMS")
    If cRef = 0 Or cItems = 0 Then Exit Function

    Dim i As Long
    For i = 1 To lo.ListRows.count
        If StrComp(NzStr(lo.DataBodyRange.Cells(i, cRef).Value), refNumber, vbTextCompare) = 0 _
           And StrComp(NzStr(lo.DataBodyRange.Cells(i, cItems).Value), itemText, vbTextCompare) = 0 Then
            FindHoldRowIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function HoldRowQty(lo As ListObject, ByVal rowIndex As Long) As Double
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > lo.ListRows.count Then Exit Function

    Dim cQty As Long: cQty = ColumnIndex(lo, "QUANTITY")
    If cQty = 0 Then Exit Function
    HoldRowQty = NzDbl(lo.DataBodyRange.Cells(rowIndex, cQty).Value)
End Function

Private Function HoldRowQtyByKey(lo As ListObject, ByVal refNumber As String, ByVal itemText As String) As Double
    Dim rowIndex As Long
    rowIndex = FindHoldRowIndex(lo, refNumber, itemText)
    If rowIndex > 0 Then HoldRowQtyByKey = HoldRowQty(lo, rowIndex)
End Function

' ===== helpers reused from modTS_Received =====
Private Function SheetExists(nameOrCode As String) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ResolveShippingWorkbook(, nameOrCode)
    If wb Is Nothing Then Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set SheetExists = ws
            Exit Function
        End If
    Next ws
End Function

Private Function GetListObject(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function GetInvSysTable() As ListObject
    Dim wsInv As Worksheet: Set wsInv = GetInventoryWorksheetShipping()
    If wsInv Is Nothing Then Exit Function
    On Error Resume Next
    Set GetInvSysTable = wsInv.ListObjects("invSys")
    On Error GoTo 0
End Function

Private Function GetInvSysTableFromWorkbook(ByVal wb As Workbook) As ListObject
    Dim wsInv As Worksheet

    If wb Is Nothing Then Exit Function
    Set wsInv = WorkbookSheetExistsShipping(wb, SHEET_INV)
    If wsInv Is Nothing Then Set wsInv = WorkbookSheetExistsShipping(wb, "Inventory Management")
    If wsInv Is Nothing Then Set wsInv = WorkbookSheetExistsShipping(wb, "INVENTORY MANAGEMENT")
    If wsInv Is Nothing Then Exit Function

    Set GetInvSysTableFromWorkbook = GetListObject(wsInv, "invSys")
End Function

Private Function GetInventoryWorksheetShipping() As Worksheet
    Set GetInventoryWorksheetShipping = SheetExists(SHEET_INV)
    If GetInventoryWorksheetShipping Is Nothing Then Set GetInventoryWorksheetShipping = SheetExists("Inventory Management")
    If GetInventoryWorksheetShipping Is Nothing Then Set GetInventoryWorksheetShipping = SheetExists("INVENTORY MANAGEMENT")
End Function

Private Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    ColumnIndex = 0
End Function

Private Function FindInvListRowByRowValue(invLo As ListObject, ByVal rowValue As Long) As ListRow
    If invLo Is Nothing Or rowValue <= 0 Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then Exit Function
    Dim cel As Range
    For Each cel In invLo.ListColumns(cRow).DataBodyRange.Cells
        If NzLng(cel.Value) = rowValue Then
            Set FindInvListRowByRowValue = invLo.ListRows(cel.Row - invLo.DataBodyRange.Row + 1)
            Exit Function
        End If
    Next cel
End Function

Private Function ValidateComponentInventory(invLo As ListObject, aggBom As ListObject, ByRef shortageMsg As String) As Boolean
    shortageMsg = ""
    ValidateComponentInventory = False
    If invLo Is Nothing Then
        shortageMsg = "invSys table not found."
        Exit Function
    End If
    If aggBom Is Nothing Or aggBom.DataBodyRange Is Nothing Then
        ValidateComponentInventory = True
        Exit Function
    End If

    Dim cQtyAgg As Long: cQtyAgg = ColumnIndex(aggBom, "QUANTITY")
    Dim cRowAgg As Long: cRowAgg = ColumnIndex(aggBom, "ROW")
    If cQtyAgg = 0 Or cRowAgg = 0 Then
        shortageMsg = "AggregateBoxBOM is missing QUANTITY or ROW columns."
        Exit Function
    End If

    Dim colTotalInv As Long: colTotalInv = ColumnIndex(invLo, "TOTAL INV")
    If colTotalInv = 0 Then
        shortageMsg = "invSys table must contain TOTAL INV column."
        Exit Function
    End If

    Dim arr As Variant
    arr = aggBom.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As Long: rowVal = NzLng(arr(r, cRowAgg))
        Dim qtyNeeded As Double: qtyNeeded = NzDbl(arr(r, cQtyAgg))
        If rowVal = 0 Or qtyNeeded <= 0 Then GoTo NextComponent
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then
            AppendNote shortageMsg, "invSys ROW " & rowVal & " not found."
            GoTo NextComponent
        End If
        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotalInv)
        Dim available As Double: available = NzDbl(totalCell.Value)
        If available < qtyNeeded Then
            AppendNote shortageMsg, "ROW " & rowVal & " requires " & Format$(qtyNeeded, "0.###") & " but only " & Format$(available, "0.###") & " available."
        End If
NextComponent:
    Next r

    ValidateComponentInventory = (Len(shortageMsg) = 0)
End Function

Private Function FindInvRowIndexByRow(invLo As ListObject, ByVal rowValue As Long) As Long
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then Exit Function
    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If NzLng(invLo.DataBodyRange.Cells(r, cRow).Value) = rowValue Then
            FindInvRowIndexByRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindInvRowIndexByItem(invLo As ListObject, ByVal itemName As String) As Long
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    Dim cItem As Long: cItem = ColumnIndex(invLo, "ITEM")
    If cItem = 0 Then Exit Function
    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If StrComp(Trim$(NzStr(invLo.DataBodyRange.Cells(r, cItem).Value)), Trim$(itemName), vbTextCompare) = 0 Then
            FindInvRowIndexByItem = r
            Exit Function
        End If
    Next r
End Function

Private Sub EnsureInvSysRowSeed(invLo As ListObject)
    If mNextInvSysRow > 0 Then Exit Sub
    mNextInvSysRow = CurrentInvSysMaxRow(invLo) + 1
    If mNextInvSysRow <= 0 Then mNextInvSysRow = 1
End Sub

Private Function CurrentInvSysMaxRow(invLo As ListObject) As Long
    If invLo Is Nothing Then Exit Function
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function
    Dim maxVal As Long
    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        Dim v As Variant: v = invLo.DataBodyRange.Cells(r, cRow).Value
        If IsNumeric(v) Then
            If CLng(v) > maxVal Then maxVal = CLng(v)
        End If
    Next r
    CurrentInvSysMaxRow = maxVal
End Function

Private Function NextInvSysRowValue(invLo As ListObject) As Long
    EnsureInvSysRowSeed invLo
    If mNextInvSysRow <= 0 Then mNextInvSysRow = 1
    NextInvSysRowValue = mNextInvSysRow
    mNextInvSysRow = mNextInvSysRow + 1
End Function

Private Function ResolveBoxPackageRowValue(ByVal operatorWb As Workbook, ByVal boxName As String, ByVal invLo As ListObject) As Long
    Dim existingIdx As Long
    Dim cRow As Long
    Dim localRow As Long
    Dim maxRow As Long
    Dim runtimeRow As Long
    Dim runtimeMax As Long

    existingIdx = FindInvRowIndexByItem(invLo, boxName)
    cRow = ColumnIndex(invLo, "ROW")
    If existingIdx > 0 And cRow > 0 Then localRow = NzLng(invLo.DataBodyRange.Cells(existingIdx, cRow).Value)
    If localRow > 0 Then
        ResolveBoxPackageRowValue = localRow
        Exit Function
    End If

    runtimeRow = FindShippingBomPackageRowByName(operatorWb, boxName, runtimeMax)
    If runtimeRow > 0 Then
        ResolveBoxPackageRowValue = runtimeRow
        Exit Function
    End If

    maxRow = CurrentInvSysMaxRow(invLo)
    If runtimeMax > maxRow Then maxRow = runtimeMax
    ResolveBoxPackageRowValue = maxRow + 1
    If ResolveBoxPackageRowValue <= 0 Then ResolveBoxPackageRowValue = 1
End Function

Private Function FindShippingBomPackageRowByName(ByVal operatorWb As Workbook, _
                                                 ByVal boxName As String, _
                                                 ByRef maxPackageRow As Long) As Long
    On Error GoTo CleanExit

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim openedTransient As Boolean
    Dim report As String
    Dim cPackageRow As Long
    Dim cPackageItem As Long
    Dim i As Long
    Dim rowValue As Long

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then GoTo CleanExit

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then GoTo CleanExit

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbBom Is Nothing Then GoTo CleanExit

    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit

    cPackageRow = ColumnIndex(loBom, "PackageRow")
    cPackageItem = ColumnIndex(loBom, "PackageItem")
    If cPackageRow = 0 Then GoTo CleanExit
    If loBom.DataBodyRange Is Nothing Then GoTo CleanExit

    For i = 1 To loBom.ListRows.Count
        rowValue = NzLng(loBom.DataBodyRange.Cells(i, cPackageRow).Value)
        If rowValue > maxPackageRow Then maxPackageRow = rowValue
        If cPackageItem > 0 Then
            If StrComp(Trim$(NzStr(loBom.DataBodyRange.Cells(i, cPackageItem).Value)), Trim$(boxName), vbTextCompare) = 0 Then
                FindShippingBomPackageRowByName = rowValue
            End If
        End If
    Next i

CleanExit:
    On Error Resume Next
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    On Error GoTo 0
End Function

Private Function EnsureInvSysItem(boxName As String, uom As String, location As String, descr As String, invLo As ListObject, Optional ByVal preferredRowValue As Long = 0) As Long
    If invLo Is Nothing Then Exit Function
    EnsureInvSysRowSeed invLo
    Dim existingIdx As Long
    existingIdx = FindInvRowIndexByItem(invLo, boxName)
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If existingIdx > 0 Then
        If cRow > 0 Then EnsureInvSysItem = NzLng(invLo.DataBodyRange.Cells(existingIdx, cRow).Value)
        If EnsureInvSysItem <= 0 And preferredRowValue > 0 Then EnsureInvSysItem = preferredRowValue
        If EnsureInvSysItem >= mNextInvSysRow Then
            mNextInvSysRow = EnsureInvSysItem + 1
        End If
        UpdateInvSysRow invLo.ListRows(existingIdx), boxName, uom, location, descr, EnsureInvSysItem
        Exit Function
    End If

    Dim lr As ListRow: Set lr = invLo.ListRows.Add
    Dim newRowVal As Long
    If preferredRowValue > 0 Then
        newRowVal = preferredRowValue
        If newRowVal >= mNextInvSysRow Then mNextInvSysRow = newRowVal + 1
    Else
        newRowVal = NextInvSysRowValue(invLo)
    End If
    EnsureInvSysItem = newRowVal
    UpdateInvSysRow lr, boxName, uom, location, descr, newRowVal
End Function

Private Function EnsureInvSysItemByRow(ByVal rowValue As Long, _
                                       ByVal itemName As String, _
                                       ByVal uom As String, _
                                       ByVal location As String, _
                                       ByVal descr As String, _
                                       ByVal invLo As ListObject) As Long
    Dim existingIdx As Long
    Dim lr As ListRow

    If rowValue <= 0 Then
        EnsureInvSysItemByRow = EnsureInvSysItem(itemName, uom, location, descr, invLo)
        Exit Function
    End If
    If invLo Is Nothing Then Exit Function

    EnsureInvSysRowSeed invLo
    existingIdx = FindInvRowIndexByRow(invLo, rowValue)
    If existingIdx > 0 Then
        UpdateInvSysRow invLo.ListRows(existingIdx), itemName, uom, location, descr, rowValue
        EnsureInvSysItemByRow = rowValue
        Exit Function
    End If

    Set lr = invLo.ListRows.Add
    UpdateInvSysRow lr, itemName, uom, location, descr, rowValue
    If rowValue >= mNextInvSysRow Then mNextInvSysRow = rowValue + 1
    EnsureInvSysItemByRow = rowValue
End Function

Private Sub UpdateInvSysRow(lr As ListRow, boxName As String, uom As String, location As String, descr As String, Optional forceRowValue As Variant)
    If lr Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = lr.Parent
    Dim idx As Long
    If Not IsMissing(forceRowValue) Then
        idx = ColumnIndex(lo, "ROW")
        If idx > 0 Then lr.Range.Cells(1, idx).Value = forceRowValue
    End If
    idx = ColumnIndex(lo, "ITEM")
    If idx > 0 Then lr.Range.Cells(1, idx).Value = boxName
    idx = ColumnIndex(lo, "ITEM_CODE")
    If idx > 0 And Trim$(NzStr(lr.Range.Cells(1, idx).Value)) = "" Then
        lr.Range.Cells(1, idx).Value = boxName
    End If
    idx = ColumnIndex(lo, "UOM")
    If idx > 0 Then lr.Range.Cells(1, idx).Value = uom
    idx = ColumnIndex(lo, "LOCATION")
    If idx > 0 Then lr.Range.Cells(1, idx).Value = location
    idx = ColumnIndex(lo, "DESCRIPTION")
    If idx > 0 Then lr.Range.Cells(1, idx).Value = descr
End Sub

Private Sub EnsureTableHasRow(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Private Sub ClearListObjectData(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    On Error GoTo 0
End Sub

Public Sub InvalidateAggregates(Optional rebuildNow As Boolean = False, Optional skipAggRebuild As Boolean = False)
    mAggDirty = True
    If rebuildNow Then
        RebuildShippingAggregates skipAggRebuild
    End If
End Sub

Public Sub RebuildShippingAggregates(Optional skipAggRebuild As Boolean = False)
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim loShip As ListObject: Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    Dim loAggPack As ListObject: Set loAggPack = GetListObject(ws, TABLE_AGG_PACK)
    Dim loAggBom As ListObject: Set loAggBom = GetListObject(ws, TABLE_AGG_BOM)
    Dim loCheck As ListObject: Set loCheck = GetListObject(ws, TABLE_CHECK_INV)

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    Dim rowCache As Object, nameCache As Object
    BuildInvSysCaches invLo, rowCache, nameCache

    Dim pkgDict As Object
    Set pkgDict = BuildPackageSummary(loShip, rowCache, nameCache)

    WriteAggregatePackages loAggPack, pkgDict

    Dim bomDict As Object
    Dim useExisting As Boolean: useExisting = UseExistingInventoryEnabled(ws)
    If Not useExisting Then
        Set bomDict = BuildBomSummary(pkgDict, rowCache)
    End If

    If useExisting Then
        ClearListObjectData loAggBom
    ElseIf Not skipAggRebuild Then
        WriteAggregateBOM loAggBom, bomDict
    End If

    WriteCheckInv loCheck, rowCache, pkgDict, bomDict
    mAggDirty = False
End Sub

Private Function UseExistingInventoryEnabled(ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(CHK_USE_EXISTING)
    On Error GoTo 0
    If shp Is Nothing Then Exit Function
    On Error Resume Next
    UseExistingInventoryEnabled = (shp.ControlFormat.Value = 1)
    On Error GoTo 0
End Function

Private Sub BuildInvSysCaches(invLo As ListObject, ByRef rowCache As Object, ByRef nameCache As Object)
    Set rowCache = CreateObject("Scripting.Dictionary")
    Set nameCache = CreateObject("Scripting.Dictionary")
    If invLo Is Nothing Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If cRow = 0 Then Exit Sub
    Dim cItem As Long: cItem = ColumnIndex(invLo, "ITEM")
    Dim cItemCode As Long: cItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim cUom As Long: cUom = ColumnIndex(invLo, "UOM")
    Dim cLoc As Long: cLoc = ColumnIndex(invLo, "LOCATION")
    Dim cUsed As Long: cUsed = ColumnIndex(invLo, "USED")
    Dim cMade As Long: cMade = ColumnIndex(invLo, "MADE")
    Dim cShip As Long: cShip = ColumnIndex(invLo, "SHIPMENTS")
    Dim cTotal As Long: cTotal = ColumnIndex(invLo, "TOTAL INV")

    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        Dim rowVal As Long: rowVal = NzLng(invLo.DataBodyRange.Cells(r, cRow).Value)
        If rowVal = 0 Then GoTo NextRow
        Dim info As Object: Set info = CreateObject("Scripting.Dictionary")
        If cItem > 0 Then info("ITEM") = NzStr(invLo.DataBodyRange.Cells(r, cItem).Value)
        If cItemCode > 0 Then info("ITEM_CODE") = NzStr(invLo.DataBodyRange.Cells(r, cItemCode).Value)
        If cUom > 0 Then info("UOM") = NzStr(invLo.DataBodyRange.Cells(r, cUom).Value)
        If cLoc > 0 Then info("LOCATION") = NzStr(invLo.DataBodyRange.Cells(r, cLoc).Value)
        If cUsed > 0 Then info("USED") = NzDbl(invLo.DataBodyRange.Cells(r, cUsed).Value)
        If cMade > 0 Then info("MADE") = NzDbl(invLo.DataBodyRange.Cells(r, cMade).Value)
        If cShip > 0 Then info("SHIPMENTS") = NzDbl(invLo.DataBodyRange.Cells(r, cShip).Value)
        If cTotal > 0 Then info("TOTAL_INV") = NzDbl(invLo.DataBodyRange.Cells(r, cTotal).Value)
        Dim cacheKey As String
        cacheKey = CStr(rowVal)
        If rowCache.Exists(cacheKey) Then
            Set rowCache(cacheKey) = info
        Else
            rowCache.Add cacheKey, info
        End If
        Dim itemKey As String
        itemKey = LCase$(NzStr(infovalue(info, "ITEM")))
        If itemKey <> "" Then
            If Not nameCache.Exists(itemKey) Then nameCache(itemKey) = rowVal
        End If
NextRow:
    Next r
End Sub

Private Function BuildPackageSummary(loShip As ListObject, rowCache As Object, nameCache As Object) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If loShip Is Nothing Then
        Set BuildPackageSummary = dict
        Exit Function
    End If
    If loShip.DataBodyRange Is Nothing Then
        Set BuildPackageSummary = dict
        Exit Function
    End If

    Dim cItem As Long: cItem = ColumnIndex(loShip, "ITEMS")
    Dim cQty As Long: cQty = ColumnIndex(loShip, "QUANTITY")
    If cItem = 0 Or cQty = 0 Then
        Set BuildPackageSummary = dict
        Exit Function
    End If

    Dim data As Variant
    data = loShip.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim itemName As String: itemName = NzStr(data(r, cItem))
        Dim qty As Double: qty = NzDbl(data(r, cQty))
        If qty <= 0 Or itemName = "" Then GoTo NextRow
        Dim rowVal As Long
        rowVal = ResolveRowFromCaches(itemName, nameCache)
        If rowVal = 0 Then GoTo NextRow
        Dim key As String: key = CStr(rowVal)
        Dim info As Object
        If dict.Exists(key) Then
            Set info = dict(key)
        Else
            Set info = CreateObject("Scripting.Dictionary")
            info("ROW") = rowVal
            Dim invInfo As Object
            If Not rowCache Is Nothing Then
                If rowCache.Exists(key) Then
                    Set invInfo = rowCache(key)
                End If
            End If
            If Not invInfo Is Nothing Then
                info("ITEM") = NzStr(infovalue(invInfo, "ITEM"))
                info("UOM") = NzStr(infovalue(invInfo, "UOM"))
            Else
                info("ITEM") = itemName
                info("UOM") = ""
            End If
            info("QTY") = 0#
            dict.Add key, info
        End If
        info("QTY") = NzDbl(info("QTY")) + qty
NextRow:
    Next r
    Set BuildPackageSummary = dict
End Function

Private Function BuildBomSummary(pkgDict As Object, rowCache As Object) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If pkgDict Is Nothing Then
        Set BuildBomSummary = dict
        Exit Function
    End If
    If pkgDict.Count = 0 Then
        Set BuildBomSummary = dict
        Exit Function
    End If

    Dim wsShip As Worksheet
    Dim loView As ListObject
    Dim refreshReport As String

    Set wsShip = SheetExists(SHEET_SHIPMENTS)
    If Not wsShip Is Nothing Then
        Set loView = GetListObject(wsShip, TABLE_SHIPPING_BOM_VIEW)
        If Not loView Is Nothing Then
            If loView.DataBodyRange Is Nothing Then RefreshShippingBomViewForWorkbook wsShip.Parent, refreshReport
            If Not loView.DataBodyRange Is Nothing Then
                Set BuildBomSummary = BuildBomSummaryFromBomRows(pkgDict, rowCache, loView)
                Exit Function
            End If
        End If
    End If

    Dim key As Variant
    For Each key In pkgDict.Keys
        Dim pkgInfo As Object: Set pkgInfo = pkgDict(key)
        Dim pkgQty As Double: pkgQty = NzDbl(infovalue(pkgInfo, "QTY"))
        Dim pkgRow As Long: pkgRow = NzLng(infovalue(pkgInfo, "ROW"))
        If pkgRow = 0 Or pkgQty <= 0 Then GoTo NextPkg
        Dim bomLo As ListObject: Set bomLo = GetBomTableByRow(pkgRow)
        If bomLo Is Nothing Then GoTo NextPkg
        If bomLo.DataBodyRange Is Nothing Then GoTo NextPkg
        Dim cRow As Long: cRow = ColumnIndex(bomLo, "ROW")
        Dim cQty As Long: cQty = ColumnIndex(bomLo, "QUANTITY")
        Dim cUom As Long: cUom = ColumnIndex(bomLo, "UOM")
        If cRow = 0 Or cQty = 0 Then GoTo NextPkg
        Dim arr As Variant
        arr = bomLo.DataBodyRange.Value
        Dim r As Long
        For r = 1 To UBound(arr, 1)
            Dim compRow As Long: compRow = NzLng(arr(r, cRow))
            Dim bomQty As Double: bomQty = NzDbl(arr(r, cQty))
            If compRow = 0 Or bomQty = 0 Then GoTo NextComponent
            Dim totalUse As Double: totalUse = bomQty * pkgQty
            If totalUse = 0 Then GoTo NextComponent
            Dim compKey As String: compKey = CStr(compRow)
            Dim info As Object
            If dict.Exists(compKey) Then
                Set info = dict(compKey)
            Else
                Set info = CreateObject("Scripting.Dictionary")
                info("ROW") = compRow
                Dim invInfo As Object
                If Not rowCache Is Nothing Then
                    If rowCache.Exists(compKey) Then
                        Set invInfo = rowCache(compKey)
                    End If
                End If
                If Not invInfo Is Nothing Then
                    info("ITEM") = NzStr(infovalue(invInfo, "ITEM"))
                    info("UOM") = NzStr(infovalue(invInfo, "UOM"))
                Else
                    info("ITEM") = ""
                    info("UOM") = ""
                End If
                info("QTY") = 0#
                dict.Add compKey, info
            End If
            info("QTY") = NzDbl(info("QTY")) + totalUse
            If cUom > 0 Then
                Dim bomUom As String: bomUom = NzStr(arr(r, cUom))
                If bomUom <> "" Then info("UOM") = bomUom
            End If
NextComponent:
        Next r
NextPkg:
    Next key
    Set BuildBomSummary = dict
End Function

Private Function BuildBomSummaryFromBomRows(pkgDict As Object, rowCache As Object, loBomRows As ListObject) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If pkgDict Is Nothing Or loBomRows Is Nothing Then
        Set BuildBomSummaryFromBomRows = dict
        Exit Function
    End If
    If pkgDict.Count = 0 Or loBomRows.DataBodyRange Is Nothing Then
        Set BuildBomSummaryFromBomRows = dict
        Exit Function
    End If

    Dim cPackageRow As Long: cPackageRow = ColumnIndex(loBomRows, "PackageRow")
    Dim cComponentRow As Long: cComponentRow = ColumnIndex(loBomRows, "ComponentRow")
    Dim cComponentQty As Long: cComponentQty = ColumnIndex(loBomRows, "ComponentQty")
    Dim cComponentUom As Long: cComponentUom = ColumnIndex(loBomRows, "ComponentUOM")
    If cPackageRow = 0 Or cComponentRow = 0 Or cComponentQty = 0 Then
        Set BuildBomSummaryFromBomRows = dict
        Exit Function
    End If

    Dim arr As Variant
    arr = loBomRows.DataBodyRange.Value

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim pkgRow As Long: pkgRow = NzLng(arr(r, cPackageRow))
        Dim pkgKey As String: pkgKey = CStr(pkgRow)
        If pkgRow = 0 Then GoTo NextBomRow
        If Not pkgDict.Exists(pkgKey) Then GoTo NextBomRow

        Dim pkgInfo As Object: Set pkgInfo = pkgDict(pkgKey)
        Dim pkgQty As Double: pkgQty = NzDbl(infovalue(pkgInfo, "QTY"))
        If pkgQty <= 0 Then GoTo NextBomRow

        Dim compRow As Long: compRow = NzLng(arr(r, cComponentRow))
        Dim bomQty As Double: bomQty = NzDbl(arr(r, cComponentQty))
        If compRow = 0 Or bomQty = 0 Then GoTo NextBomRow

        Dim totalUse As Double: totalUse = bomQty * pkgQty
        If totalUse = 0 Then GoTo NextBomRow

        Dim compKey As String: compKey = CStr(compRow)
        Dim info As Object
        If dict.Exists(compKey) Then
            Set info = dict(compKey)
        Else
            Set info = CreateObject("Scripting.Dictionary")
            info("ROW") = compRow
            Dim invInfo As Object
            If Not rowCache Is Nothing Then
                If rowCache.Exists(compKey) Then Set invInfo = rowCache(compKey)
            End If
            If Not invInfo Is Nothing Then
                info("ITEM") = NzStr(infovalue(invInfo, "ITEM"))
                info("UOM") = NzStr(infovalue(invInfo, "UOM"))
            Else
                info("ITEM") = ""
                info("UOM") = ""
            End If
            info("QTY") = 0#
            dict.Add compKey, info
        End If
        info("QTY") = NzDbl(info("QTY")) + totalUse
        If cComponentUom > 0 Then
            Dim bomUom As String: bomUom = NzStr(arr(r, cComponentUom))
            If bomUom <> "" Then info("UOM") = bomUom
        End If
NextBomRow:
    Next r

    Set BuildBomSummaryFromBomRows = dict
End Function

Private Sub WriteAggregatePackages(lo As ListObject, pkgDict As Object)
    If lo Is Nothing Then Exit Sub
    ClearListObjectData lo
    If pkgDict Is Nothing Then Exit Sub
    If pkgDict.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = SortedKeys(pkgDict)
    Dim count As Long: count = UBound(keys) - LBound(keys) + 1
    ReDim arr(1 To count, 1 To 4)
    Dim i As Long
    For i = 1 To count
        Dim key As Variant: key = keys(LBound(keys) + i - 1)
        Dim info As Object: Set info = pkgDict(key)
        arr(i, 1) = NzDbl(infovalue(info, "QTY"))
        arr(i, 2) = NzStr(infovalue(info, "UOM"))
        Dim itemText As String: itemText = NzStr(infovalue(info, "ITEM"))
        If itemText = "" Then itemText = NzStr(key)
        arr(i, 3) = itemText
        arr(i, 4) = CLng(key)
    Next i
    WriteArrayToTable lo, arr
End Sub

Private Sub WriteAggregateBOM(lo As ListObject, bomDict As Object)
    If lo Is Nothing Then Exit Sub
    ClearListObjectData lo
    If bomDict Is Nothing Then Exit Sub
    If bomDict.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = SortedKeys(bomDict)
    Dim count As Long: count = UBound(keys) - LBound(keys) + 1
    ReDim arr(1 To count, 1 To 4)
    Dim i As Long
    For i = 1 To count
        Dim key As Variant: key = keys(LBound(keys) + i - 1)
        Dim info As Object: Set info = bomDict(key)
        arr(i, 1) = NzDbl(infovalue(info, "QTY"))
        arr(i, 2) = NzStr(infovalue(info, "UOM"))
        arr(i, 3) = NzStr(infovalue(info, "ITEM"))
        arr(i, 4) = CLng(key)
    Next i
    WriteArrayToTable lo, arr
End Sub

Private Sub WriteCheckInv(lo As ListObject, rowCache As Object, pkgDict As Object, bomDict As Object)
    If lo Is Nothing Then Exit Sub
    ClearListObjectData lo
    Dim rowsDict As Object: Set rowsDict = CreateObject("Scripting.Dictionary")
    Dim keyPkg As Variant, keyBom As Variant, rowKey As Variant
    If Not pkgDict Is Nothing Then
        For Each keyPkg In pkgDict.Keys
            rowsDict(CStr(keyPkg)) = True
        Next keyPkg
    End If
    If Not bomDict Is Nothing Then
        For Each keyBom In bomDict.Keys
            rowsDict(CStr(keyBom)) = True
        Next keyBom
    End If
    If rowsDict.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = SortedKeys(rowsDict)
    Dim count As Long: count = UBound(keys) - LBound(keys) + 1
    ReDim arr(1 To count, 1 To 5)
    Dim i As Long
    For i = 1 To count
        rowKey = keys(LBound(keys) + i - 1)
        Dim info As Object
        If Not rowCache Is Nothing Then
            If rowCache.Exists(CStr(rowKey)) Then Set info = rowCache(CStr(rowKey))
        End If
        arr(i, 1) = NzDbl(infovalue(info, "USED"))
        arr(i, 2) = NzDbl(infovalue(info, "MADE"))
        arr(i, 3) = NzDbl(infovalue(info, "SHIPMENTS"))
        arr(i, 4) = NzDbl(infovalue(info, "TOTAL_INV"))
        arr(i, 5) = CLng(rowKey)
    Next i
    WriteArrayToTable lo, arr
End Sub

Private Function BuildUsedDeltaPacket(invLo As ListObject, aggBom As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function

    Dim colUsed As Long: colUsed = ColumnIndex(invLo, "USED")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    If colUsed = 0 Or colRow = 0 Then
        errNotes = "invSys table missing USED/ROW columns."
        Exit Function
    End If

    Dim result As New Collection
    Dim arr As Variant: arr = invLo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim usedVal As Double: usedVal = NzDbl(arr(r, colUsed))
        Dim rowVal As Long: rowVal = NzLng(arr(r, colRow))
        If rowVal = 0 Or usedVal <= 0 Then GoTo NextRow
        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = rowVal
        delta("QTY") = usedVal
        If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(arr(r, colItemCode))
        If colItemName > 0 Then delta("ITEM_NAME") = NzStr(arr(r, colItemName))
        result.Add delta
NextRow:
    Next r

    If result.Count = 0 Then
        errNotes = "No staged usage found in invSys.USED."
        Exit Function
    End If
    Set BuildUsedDeltaPacket = result
End Function

Private Function BuildComponentDeltaPacketFromAggregate(invLo As ListObject, aggBom As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then
        errNotes = "invSys table not found."
        Exit Function
    End If
    If aggBom Is Nothing Or aggBom.DataBodyRange Is Nothing Then
        errNotes = "AggregateBoxBOM has no component rows."
        Exit Function
    End If

    Dim cQtyAgg As Long: cQtyAgg = ColumnIndex(aggBom, "QUANTITY")
    Dim cRowAgg As Long: cRowAgg = ColumnIndex(aggBom, "ROW")
    If cQtyAgg = 0 Or cRowAgg = 0 Then
        errNotes = "AggregateBoxBOM is missing QUANTITY or ROW columns."
        Exit Function
    End If

    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    Dim requirements As Object: Set requirements = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = aggBom.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As Long: rowVal = NzLng(arr(r, cRowAgg))
        Dim qtyVal As Double: qtyVal = NzDbl(arr(r, cQtyAgg))
        If rowVal = 0 Or qtyVal <= 0 Then GoTo NextAggRow
        Dim reqKey As String: reqKey = CStr(rowVal)
        If requirements.Exists(reqKey) Then
            requirements(reqKey) = NzDbl(requirements(reqKey)) + qtyVal
        Else
            requirements.Add reqKey, qtyVal
        End If
NextAggRow:
    Next r

    If requirements.Count = 0 Then
        errNotes = "No component quantities were found in AggregateBoxBOM."
        Exit Function
    End If

    Dim result As New Collection
    Dim key As Variant
    For Each key In requirements.Keys
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, CLng(key))
        If invRow Is Nothing Then
            AppendNote errNotes, "Component ROW " & CStr(key) & " not found in invSys."
        Else
            Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
            delta("ROW") = CLng(key)
            delta("QTY") = NzDbl(requirements(key))
            If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(invRow.Range.Cells(1, colItemCode).Value)
            If colItemName > 0 Then delta("ITEM_NAME") = NzStr(invRow.Range.Cells(1, colItemName).Value)
            result.Add delta
        End If
    Next key

    If result.Count = 0 Then
        If errNotes = "" Then errNotes = "No component deltas were available."
        Exit Function
    End If
    Set BuildComponentDeltaPacketFromAggregate = result
End Function

Private Function BuildComponentDeltaPacketFromBoxBom(invLo As ListObject, loBom As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then
        errNotes = "invSys table not found."
        Exit Function
    End If
    If loBom Is Nothing Or loBom.DataBodyRange Is Nothing Then
        errNotes = "BoxBOM has no component rows."
        Exit Function
    End If

    Dim cItem As Long: cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    Dim cQty As Long: cQty = ColumnIndex(loBom, "QUANTITY")
    Dim cRow As Long: cRow = ColumnIndex(loBom, "ROW")
    Dim cUom As Long: cUom = ColumnIndex(loBom, "UOM")
    Dim cLocation As Long: cLocation = ColumnIndex(loBom, "LOCATION")
    Dim cDescription As Long: cDescription = ColumnIndex(loBom, "DESCRIPTION")
    If cItem = 0 Or cQty = 0 Or cRow = 0 Then
        errNotes = "BoxBOM must include ITEM, ROW, and QUANTITY columns."
        Exit Function
    End If

    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    Dim colTotalInv As Long: colTotalInv = ColumnIndex(invLo, "TOTAL INV")
    Dim requirements As Object: Set requirements = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To loBom.ListRows.Count
        Dim itemName As String: itemName = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cItem).Value))
        Dim rowVal As Long: rowVal = NzLng(loBom.DataBodyRange.Cells(r, cRow).Value)
        Dim qtyVal As Double: qtyVal = NzDbl(loBom.DataBodyRange.Cells(r, cQty).Value)
        If BoxMakerComponentRowIsBlank(itemName, rowVal, qtyVal) Then GoTo NextBomRow
        If qtyVal <= 0 Then
            errNotes = "BoxBOM row " & CStr(r) & " needs a component Quantity greater than zero."
            Exit Function
        End If

        Dim invIdx As Long
        If rowVal > 0 Then
            invIdx = FindInvRowIndexByRow(invLo, rowVal)
        ElseIf itemName <> "" Then
            invIdx = FindInvRowIndexByItem(invLo, itemName)
        End If
        If invIdx <= 0 And (itemName <> "" Or rowVal > 0) Then
            Dim uomVal As String
            Dim locationVal As String
            Dim descrVal As String
            Dim currentInv As Variant
            Dim foundCurrent As Boolean
            Dim ensuredRow As Long
            Dim snapshotCache As Object

            If cUom > 0 Then uomVal = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cUom).Value))
            If cLocation > 0 Then locationVal = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cLocation).Value))
            If cDescription > 0 Then descrVal = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cDescription).Value))

            ensuredRow = EnsureInvSysItemByRow(rowVal, itemName, uomVal, locationVal, descrVal, invLo)
            If ensuredRow > 0 Then
                If rowVal > 0 Then invIdx = FindInvRowIndexByRow(invLo, rowVal)
                If invIdx <= 0 Then invIdx = FindInvRowIndexByRow(invLo, ensuredRow)
                If invIdx <= 0 And itemName <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
                If invIdx > 0 And colTotalInv > 0 Then
                    currentInv = ResolveCurrentInventoryValue(loBom.Parent, invLo, rowVal, itemName, foundCurrent, snapshotCache)
                    If foundCurrent Then invLo.DataBodyRange.Cells(invIdx, colTotalInv).Value = currentInv
                End If
            End If
        End If
        If invIdx <= 0 Then
            errNotes = "BoxBOM component row " & CStr(r) & " was not found in invSys."
            Exit Function
        End If

        Dim actualRow As Long
        actualRow = NzLng(GetInvSysValueByIndex(invLo, invIdx, "ROW"))
        If actualRow <= 0 Then actualRow = rowVal
        If actualRow <= 0 Then
            errNotes = "BoxBOM component row " & CStr(r) & " does not have an invSys ROW."
            Exit Function
        End If

        Dim reqKey As String: reqKey = CStr(actualRow)
        If requirements.Exists(reqKey) Then
            requirements(reqKey) = NzDbl(requirements(reqKey)) + qtyVal
        Else
            requirements.Add reqKey, qtyVal
        End If
NextBomRow:
    Next r

    If requirements.Count = 0 Then
        errNotes = "No component quantities were found in BoxBOM."
        Exit Function
    End If

    Dim result As New Collection
    Dim key As Variant
    For Each key In requirements.Keys
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, CLng(key))
        If invRow Is Nothing Then
            AppendNote errNotes, "Component ROW " & CStr(key) & " not found in invSys."
        Else
            Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
            delta("ROW") = CLng(key)
            delta("QTY") = NzDbl(requirements(key))
            If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(invRow.Range.Cells(1, colItemCode).Value)
            If colItemName > 0 Then delta("ITEM_NAME") = NzStr(invRow.Range.Cells(1, colItemName).Value)
            result.Add delta
        End If
    Next key

    If result.Count = 0 Then
        If errNotes = "" Then errNotes = "No component deltas were available from BoxBOM."
        Exit Function
    End If
    Set BuildComponentDeltaPacketFromBoxBom = result
End Function

Private Function BuildMadeDeltaPacket(invLo As ListObject, aggPack As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If aggPack Is Nothing Or aggPack.DataBodyRange Is Nothing Then Exit Function

    Dim cQtyAgg As Long: cQtyAgg = ColumnIndex(aggPack, "QUANTITY")
    Dim cRowAgg As Long: cRowAgg = ColumnIndex(aggPack, "ROW")
    If cQtyAgg = 0 Or cRowAgg = 0 Then
        errNotes = "AggregatePackages missing QUANTITY/ROW columns."
        Exit Function
    End If

    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")

    Dim result As New Collection
    Dim arr As Variant: arr = aggPack.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As Long: rowVal = NzLng(arr(r, cRowAgg))
        Dim qtyVal As Double: qtyVal = NzDbl(arr(r, cQtyAgg))
        If rowVal = 0 Or qtyVal <= 0 Then GoTo NextPkg
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then
            AppendNote errNotes, "Package ROW " & rowVal & " not found in invSys."
        Else
            Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
            delta("ROW") = rowVal
            delta("QTY") = qtyVal
            If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(invRow.Range.Cells(1, colItemCode).Value)
            If colItemName > 0 Then delta("ITEM_NAME") = NzStr(invRow.Range.Cells(1, colItemName).Value)
            result.Add delta
        End If
NextPkg:
    Next r

    If result.Count = 0 Then
        errNotes = IIf(errNotes = "", "No packages available to make.", errNotes)
        Exit Function
    End If
    Set BuildMadeDeltaPacket = result
End Function

Private Function BuildTotalInventoryDeltaPacket(invLo As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    Dim cMade As Long: cMade = ColumnIndex(invLo, "MADE")
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    Dim cItemCode As Long: cItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim cItemName As Long: cItemName = ColumnIndex(invLo, "ITEM")
    If cMade = 0 Or cRow = 0 Then
        errNotes = "invSys table missing MADE/ROW columns."
        Exit Function
    End If

    Dim result As New Collection
    Dim arr As Variant: arr = invLo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As Long: rowVal = NzLng(arr(r, cRow))
        Dim madeVal As Double: madeVal = NzDbl(arr(r, cMade))
        If rowVal = 0 Or madeVal <= 0 Then GoTo NextRow
        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = rowVal
        delta("QTY") = madeVal
        If cItemCode > 0 Then delta("ITEM_CODE") = NzStr(arr(r, cItemCode))
        If cItemName > 0 Then delta("ITEM_NAME") = NzStr(arr(r, cItemName))
        result.Add delta
NextRow:
    Next r

    If result.Count > 0 Then Set BuildTotalInventoryDeltaPacket = result
End Function

Private Function BuildShipmentsSentDeltaPacket(invLo As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    Dim cShip As Long: cShip = ColumnIndex(invLo, "SHIPMENTS")
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    Dim cItemCode As Long: cItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim cItemName As Long: cItemName = ColumnIndex(invLo, "ITEM")
    If cShip = 0 Or cRow = 0 Then
        errNotes = "invSys table missing SHIPMENTS/ROW columns."
        Exit Function
    End If

    Dim result As New Collection
    Dim arr As Variant: arr = invLo.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As Long: rowVal = NzLng(arr(r, cRow))
        Dim shipVal As Double: shipVal = NzDbl(arr(r, cShip))
        If rowVal = 0 Or shipVal <= 0 Then GoTo NextRow
        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = rowVal
        delta("QTY") = shipVal
        If cItemCode > 0 Then delta("ITEM_CODE") = NzStr(arr(r, cItemCode))
        If cItemName > 0 Then delta("ITEM_NAME") = NzStr(arr(r, cItemName))
        result.Add delta
NextRow:
    Next r

    If result.Count = 0 Then
        errNotes = "No staged shipments found in invSys.SHIPMENTS."
        Exit Function
    End If
    Set BuildShipmentsSentDeltaPacket = result
End Function

Private Function BuildShipmentDeltaPacket(invLo As ListObject, aggPack As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Then Exit Function
    If aggPack Is Nothing Or aggPack.DataBodyRange Is Nothing Then Exit Function

    Dim cQtyAgg As Long: cQtyAgg = ColumnIndex(aggPack, "QUANTITY")
    Dim cRowAgg As Long: cRowAgg = ColumnIndex(aggPack, "ROW")
    If cQtyAgg = 0 Or cRowAgg = 0 Then
        errNotes = "AggregatePackages missing QUANTITY/ROW columns."
        Exit Function
    End If

    Dim colTotalInv As Long: colTotalInv = ColumnIndex(invLo, "TOTAL INV")
    Dim colShipments As Long: colShipments = ColumnIndex(invLo, "SHIPMENTS")
    Dim colRowInv As Long: colRowInv = ColumnIndex(invLo, "ROW")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    If colTotalInv = 0 Or colShipments = 0 Or colRowInv = 0 Then
        errNotes = "invSys table missing TOTAL INV/SHIPMENTS/ROW columns."
        Exit Function
    End If

    Dim requirements As Object: Set requirements = CreateObject("Scripting.Dictionary")
    Dim arrAgg As Variant: arrAgg = aggPack.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arrAgg, 1)
        Dim rowVal As Long: rowVal = NzLng(arrAgg(r, cRowAgg))
        Dim qtyVal As Double: qtyVal = NzDbl(arrAgg(r, cQtyAgg))
        If rowVal = 0 Or qtyVal <= 0 Then GoTo NextAgg
        Dim reqKeyStr As String: reqKeyStr = CStr(rowVal)
        If requirements.Exists(reqKeyStr) Then
            requirements(reqKeyStr) = NzDbl(requirements(reqKeyStr)) + qtyVal
        Else
            requirements.Add reqKeyStr, qtyVal
        End If
NextAgg:
    Next r

    If requirements.Count = 0 Then Exit Function

    Dim result As New Collection
    Dim shipKey As Variant
    For Each shipKey In requirements.Keys
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, CLng(shipKey))
        If invRow Is Nothing Then
            AppendNote errNotes, "Package ROW " & shipKey & " not found in invSys."
            Exit Function
        End If

        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotalInv)
        Dim shipmentsCell As Range: Set shipmentsCell = invRow.Range.Cells(1, colShipments)
        Dim requiredQty As Double: requiredQty = NzDbl(requirements(shipKey))
        Dim alreadyStaged As Double: alreadyStaged = NzDbl(shipmentsCell.Value)
        Dim neededQty As Double: neededQty = requiredQty - alreadyStaged
        If neededQty <= 0 Then GoTo NextReq

        Dim available As Double: available = NzDbl(totalCell.Value)
        If neededQty > available + 0.0000001 Then
            AppendNote errNotes, "ROW " & shipKey & " requires " & Format$(neededQty, "0.###") & " but only " & Format$(available, "0.###") & " in TOTAL INV."
            Exit Function
        End If

        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = CLng(shipKey)
        delta("QTY") = neededQty
        If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(invRow.Range.Cells(1, colItemCode).Value)
        If colItemName > 0 Then delta("ITEM_NAME") = NzStr(invRow.Range.Cells(1, colItemName).Value)
        result.Add delta
NextReq:
    Next shipKey

    If result.Count > 0 Then Set BuildShipmentDeltaPacket = result
End Function

Private Function BuildPayloadJsonFromDeltas(ByVal deltas As Collection, Optional ByVal ioType As String = "") As String
    Dim payloadItems As New Collection
    Dim delta As Variant
    Dim payloadItem As Object

    If deltas Is Nothing Then Exit Function
    If deltas.Count = 0 Then Exit Function

    For Each delta In deltas
        Set payloadItem = modRoleEventWriter.CreatePayloadItem( _
            NzLng(delta("ROW")), _
            NzStr(delta("ITEM_CODE")), _
            NzDbl(delta("QTY")), _
            "", _
            NzStr(delta("ITEM_NAME")), _
            ioType)
        payloadItems.Add payloadItem
    Next delta

    BuildPayloadJsonFromDeltas = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
End Function

Private Sub PrepareTotalInventoryLogEntries(invLo As ListObject, deltas As Collection, logEntries As Collection)
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If logEntries Is Nothing Then Exit Sub

    Dim colTotalInv As Long: colTotalInv = ColumnIndex(invLo, "TOTAL INV")
    If colTotalInv = 0 Then Exit Sub

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then GoTo NextDelta
        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotalInv)
        Dim newTotal As Double: newTotal = NzDbl(totalCell.Value) + qtyVal
        logEntries.Add Array("BTN_TO_TOTALINV", rowVal, NzStr(delta("ITEM_CODE")), NzStr(delta("ITEM_NAME")), qtyVal, newTotal)
NextDelta:
    Next delta
End Sub

Private Sub PrepareShipmentStageLogEntries(invLo As ListObject, deltas As Collection, logEntries As Collection)
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If logEntries Is Nothing Then Exit Sub

    Dim colTotalInv As Long: colTotalInv = ColumnIndex(invLo, "TOTAL INV")
    If colTotalInv = 0 Then Exit Sub

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then GoTo NextShip
        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotalInv)
        Dim newTotal As Double: newTotal = NzDbl(totalCell.Value) - qtyVal
        logEntries.Add Array("BTN_TO_SHIPMENTS", rowVal, NzStr(delta("ITEM_CODE")), NzStr(delta("ITEM_NAME")), -qtyVal, newTotal)
NextShip:
    Next delta
End Sub

Private Sub PrepareShipmentsSentLogEntries(invLo As ListObject, deltas As Collection, logEntries As Collection, Optional deductTotalInv As Boolean = False)
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If logEntries Is Nothing Then Exit Sub

    Dim colShip As Long: colShip = ColumnIndex(invLo, "SHIPMENTS")
    Dim colTotal As Long: colTotal = ColumnIndex(invLo, "TOTAL INV")
    If colShip = 0 Then Exit Sub
    If deductTotalInv And colTotal = 0 Then Exit Sub

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then GoTo NextSent
        Dim newQty As Double
        If deductTotalInv Then
            Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotal)
            newQty = NzDbl(totalCell.Value) - qtyVal
        Else
            Dim shipCell As Range: Set shipCell = invRow.Range.Cells(1, colShip)
            newQty = NzDbl(shipCell.Value) - qtyVal
        End If
        If newQty < 0 Then newQty = 0
        logEntries.Add Array("BTN_SHIPMENTS_SENT", rowVal, NzStr(delta("ITEM_CODE")), NzStr(delta("ITEM_NAME")), -qtyVal, newQty)
NextSent:
    Next delta
End Sub

Private Sub PrepareComponentLogEntries(invLo As ListObject, deltas As Collection, logEntries As Collection)
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If logEntries Is Nothing Then Exit Sub
    Dim colTotalInv As Long: colTotalInv = ColumnIndex(invLo, "TOTAL INV")
    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If Not invRow Is Nothing Then
            Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotalInv)
            Dim newTotal As Double: newTotal = NzDbl(totalCell.Value) - qtyVal
            logEntries.Add Array("BTN_BOXES_MADE_COMPONENTS", rowVal, NzStr(delta("ITEM_CODE")), NzStr(delta("ITEM_NAME")), -qtyVal, newTotal)
        End If
    Next delta
End Sub

Private Sub PreparePackageLogEntries(invLo As ListObject, deltas As Collection, logEntries As Collection)
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If logEntries Is Nothing Then Exit Sub
    Dim colMade As Long: colMade = ColumnIndex(invLo, "MADE")
    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If Not invRow Is Nothing Then
            Dim madeCell As Range: Set madeCell = invRow.Range.Cells(1, colMade)
            Dim newMade As Double: newMade = NzDbl(madeCell.Value) + qtyVal
            logEntries.Add Array("BTN_BOXES_MADE_PACKAGES", rowVal, NzStr(delta("ITEM_CODE")), NzStr(delta("ITEM_NAME")), qtyVal, newMade)
        End If
    Next delta
End Sub

Private Sub ResetShippingStaging(Optional clearShipments As Boolean = False, Optional clearPackages As Boolean = False)
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject

    Set lo = GetListObject(ws, TABLE_AGG_BOM)
    ClearListObjectData lo

    If clearPackages Then
        Set lo = GetListObject(ws, TABLE_AGG_PACK)
        ClearListObjectData lo
    End If

    If clearShipments Then
        Set lo = GetListObject(ws, TABLE_SHIPMENTS)
        ClearListObjectData lo
        Set lo = GetListObject(ws, TABLE_NOTSHIPPED)
        ClearListObjectData lo
    End If
End Sub

Private Sub ClearShipmentEntryTables()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject

    Set lo = GetListObject(ws, TABLE_SHIPMENTS)
    ClearListObjectData lo

    Set lo = GetListObject(ws, TABLE_AGG_BOM)
    ClearListObjectData lo

    Set lo = GetListObject(ws, TABLE_AGG_PACK)
    ClearListObjectData lo
End Sub

Private Sub ClearInstructionStaging(Optional ws As Worksheet)
    If ws Is Nothing Then Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = GetListObject(ws, TABLE_CHECK_INV)
    If lo Is Nothing Then Exit Sub

    Dim instrCol As Long
    instrCol = lo.Range.Column + lo.Range.Columns.Count
    If instrCol <= 0 Then Exit Sub

    If Not lo.DataBodyRange Is Nothing Then
        Dim firstRow As Long: firstRow = lo.DataBodyRange.Row
        Dim lastRow As Long: lastRow = firstRow + lo.DataBodyRange.Rows.Count - 1
        ws.Range(ws.Cells(firstRow, instrCol), ws.Cells(lastRow, instrCol)).ClearContents
    End If
End Sub

Private Function ApplyShipmentsSentDeltas(invLo As ListObject, deltas As Collection, ByRef errNotes As String, Optional deductTotalInv As Boolean = False) As Double
    ApplyShipmentsSentDeltas = 0
    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        ApplyShipmentsSentDeltas = -1
        Exit Function
    End If
    If deltas Is Nothing Then Exit Function
    If deltas.Count = 0 Then Exit Function

    Dim colShip As Long: colShip = ColumnIndex(invLo, "SHIPMENTS")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colTotal As Long
    Dim colLastEdited As Long: colLastEdited = ColumnIndex(invLo, "LAST EDITED")
    Dim colTotalLastEdit As Long: colTotalLastEdit = ColumnIndex(invLo, "TOTAL INV LAST EDIT")
    If colShip = 0 Or colRow = 0 Then
        errNotes = "invSys table missing SHIPMENTS/ROW columns."
        ApplyShipmentsSentDeltas = -1
        Exit Function
    End If
    If deductTotalInv Then
        colTotal = ColumnIndex(invLo, "TOTAL INV")
        If colTotal = 0 Then
            errNotes = "invSys table missing TOTAL INV column."
            ApplyShipmentsSentDeltas = -1
            Exit Function
        End If
    End If

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextValidate
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then
            AppendNote errNotes, "invSys ROW " & rowVal & " not found."
            ApplyShipmentsSentDeltas = -1
            Exit Function
        End If
        Dim shipCell As Range: Set shipCell = invRow.Range.Cells(1, colShip)
        Dim currentShip As Double: currentShip = NzDbl(shipCell.Value)
        If qtyVal > currentShip + 0.0000001 Then
            AppendNote errNotes, "ROW " & rowVal & " only has " & Format$(currentShip, "0.###") & " staged but needs " & Format$(qtyVal, "0.###") & "."
            ApplyShipmentsSentDeltas = -1
            Exit Function
        End If
        If deductTotalInv Then
            Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotal)
            Dim currentTotal As Double: currentTotal = NzDbl(totalCell.Value)
            If qtyVal > currentTotal + 0.0000001 Then
                AppendNote errNotes, "ROW " & rowVal & " only has " & Format$(currentTotal, "0.###") & " in TOTAL INV but needs " & Format$(qtyVal, "0.###") & "."
                ApplyShipmentsSentDeltas = -1
                Exit Function
            End If
        End If
NextValidate:
    Next delta

    For Each delta In deltas
        rowVal = CLng(delta("ROW"))
        qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextApply
        Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then GoTo NextApply
        Set shipCell = invRow.Range.Cells(1, colShip)
        Dim newShip As Double: newShip = NzDbl(shipCell.Value) - qtyVal
        If newShip < 0 Then newShip = 0
        shipCell.Value = newShip
        If deductTotalInv Then
            Set totalCell = invRow.Range.Cells(1, colTotal)
            totalCell.Value = NzDbl(totalCell.Value) - qtyVal
            If colTotalLastEdit > 0 Then invRow.Range.Cells(1, colTotalLastEdit).Value = Now
        End If
        If colLastEdited > 0 Then invRow.Range.Cells(1, colLastEdited).Value = Now
        ApplyShipmentsSentDeltas = ApplyShipmentsSentDeltas + qtyVal
NextApply:
    Next delta
End Function

Private Function ApplyShipmentDeltasLocal(invLo As ListObject, deltas As Collection, ByRef errNotes As String) As Double
    ApplyShipmentDeltasLocal = 0
    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        ApplyShipmentDeltasLocal = -1
        Exit Function
    End If
    If deltas Is Nothing Then Exit Function
    If deltas.Count = 0 Then Exit Function

    Dim colTotal As Long: colTotal = ColumnIndex(invLo, "TOTAL INV")
    Dim colShip As Long: colShip = ColumnIndex(invLo, "SHIPMENTS")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colLastEdited As Long: colLastEdited = ColumnIndex(invLo, "LAST EDITED")
    Dim colTotalLastEdit As Long: colTotalLastEdit = ColumnIndex(invLo, "TOTAL INV LAST EDIT")
    If colTotal = 0 Or colShip = 0 Or colRow = 0 Then
        errNotes = "invSys table missing TOTAL INV/SHIPMENTS/ROW columns."
        ApplyShipmentDeltasLocal = -1
        Exit Function
    End If

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextValidate

        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then
            AppendNote errNotes, "invSys ROW " & rowVal & " not found."
            ApplyShipmentDeltasLocal = -1
            Exit Function
        End If

        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotal)
        Dim currentTotal As Double: currentTotal = NzDbl(totalCell.Value)
        If qtyVal > currentTotal + 0.0000001 Then
            AppendNote errNotes, "ROW " & rowVal & " only has " & Format$(currentTotal, "0.###") & " in TOTAL INV but needs " & Format$(qtyVal, "0.###") & "."
            ApplyShipmentDeltasLocal = -1
            Exit Function
        End If
NextValidate:
    Next delta

    For Each delta In deltas
        rowVal = CLng(delta("ROW"))
        qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextApply

        Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then GoTo NextApply
        Set totalCell = invRow.Range.Cells(1, colTotal)
        Dim shipCell As Range: Set shipCell = invRow.Range.Cells(1, colShip)
        totalCell.Value = NzDbl(totalCell.Value) - qtyVal
        shipCell.Value = NzDbl(shipCell.Value) + qtyVal
        If colLastEdited > 0 Then invRow.Range.Cells(1, colLastEdited).Value = Now
        If colTotalLastEdit > 0 Then invRow.Range.Cells(1, colTotalLastEdit).Value = Now
        ApplyShipmentDeltasLocal = ApplyShipmentDeltasLocal + qtyVal
NextApply:
    Next delta
End Function

Private Sub RestoreShipmentStageColumns(ByVal invLo As ListObject, ByVal deltas As Collection)
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If deltas.Count = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim colShip As Long: colShip = ColumnIndex(invLo, "SHIPMENTS")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    If colShip = 0 Then Exit Sub

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextDelta

        Dim invRow As ListRow
        If rowVal > 0 And colRow > 0 Then Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If Not invRow Is Nothing Then
            invRow.Range.Cells(1, colShip).Value = qtyVal
        Else
            RestoreShipmentStageByItemCode invLo, colShip, colItemCode, NzStr(delta("ITEM_CODE")), qtyVal
        End If
NextDelta:
    Next delta
End Sub

Private Sub RestoreShipmentStageByItemCode(ByVal invLo As ListObject, ByVal shipmentColumn As Long, ByVal itemCodeColumn As Long, ByVal itemCode As String, ByVal qtyVal As Double)
    If invLo Is Nothing Then Exit Sub
    If shipmentColumn <= 0 Or itemCodeColumn <= 0 Then Exit Sub
    If Len(Trim$(itemCode)) = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If StrComp(NzStr(invLo.DataBodyRange.Cells(r, itemCodeColumn).Value), itemCode, vbTextCompare) = 0 Then
            invLo.DataBodyRange.Cells(r, shipmentColumn).Value = qtyVal
            Exit Sub
        End If
    Next r
End Sub

Private Function ApplyUsedDeltasLocal(invLo As ListObject, deltas As Collection, ByRef errNotes As String) As Double
    ApplyUsedDeltasLocal = 0
    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        ApplyUsedDeltasLocal = -1
        Exit Function
    End If
    If deltas Is Nothing Then Exit Function
    If deltas.Count = 0 Then Exit Function

    Dim colUsed As Long: colUsed = ColumnIndex(invLo, "USED")
    Dim colTotal As Long: colTotal = ColumnIndex(invLo, "TOTAL INV")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colLastEdited As Long: colLastEdited = ColumnIndex(invLo, "LAST EDITED")
    Dim colTotalLastEdit As Long: colTotalLastEdit = ColumnIndex(invLo, "TOTAL INV LAST EDIT")
    If colUsed = 0 Or colTotal = 0 Or colRow = 0 Then
        errNotes = "invSys table missing USED/TOTAL INV/ROW columns."
        ApplyUsedDeltasLocal = -1
        Exit Function
    End If

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextValidate

        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then
            AppendNote errNotes, "invSys ROW " & rowVal & " not found."
            ApplyUsedDeltasLocal = -1
            Exit Function
        End If

        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotal)
        Dim available As Double: available = NzDbl(totalCell.Value)
        If qtyVal > available + 0.0000001 Then
            AppendNote errNotes, "ROW " & rowVal & " requires " & Format$(qtyVal, "0.###") & " but only " & Format$(available, "0.###") & " available."
            ApplyUsedDeltasLocal = -1
            Exit Function
        End If
NextValidate:
    Next delta

    For Each delta In deltas
        rowVal = CLng(delta("ROW"))
        qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextApply

        Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then GoTo NextApply
        Set totalCell = invRow.Range.Cells(1, colTotal)
        Dim usedCell As Range: Set usedCell = invRow.Range.Cells(1, colUsed)
        totalCell.Value = NzDbl(totalCell.Value) - qtyVal
        usedCell.Value = Application.WorksheetFunction.Max(0, NzDbl(usedCell.Value) - qtyVal)
        If colLastEdited > 0 Then invRow.Range.Cells(1, colLastEdited).Value = Now
        If colTotalLastEdit > 0 Then invRow.Range.Cells(1, colTotalLastEdit).Value = Now
        ApplyUsedDeltasLocal = ApplyUsedDeltasLocal + qtyVal
NextApply:
    Next delta
End Function

Private Function ApplyMadeDeltasLocal(invLo As ListObject, deltas As Collection, ByRef errNotes As String) As Double
    ApplyMadeDeltasLocal = 0
    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        ApplyMadeDeltasLocal = -1
        Exit Function
    End If
    If deltas Is Nothing Then Exit Function
    If deltas.Count = 0 Then Exit Function

    Dim colMade As Long: colMade = ColumnIndex(invLo, "MADE")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colLastEdited As Long: colLastEdited = ColumnIndex(invLo, "LAST EDITED")
    If colMade = 0 Or colRow = 0 Then
        errNotes = "invSys table missing MADE/ROW columns."
        ApplyMadeDeltasLocal = -1
        Exit Function
    End If

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextApply

        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then
            AppendNote errNotes, "invSys ROW " & rowVal & " not found."
            ApplyMadeDeltasLocal = -1
            Exit Function
        End If

        Dim madeCell As Range: Set madeCell = invRow.Range.Cells(1, colMade)
        madeCell.Value = NzDbl(madeCell.Value) + qtyVal
        If colLastEdited > 0 Then invRow.Range.Cells(1, colLastEdited).Value = Now
        ApplyMadeDeltasLocal = ApplyMadeDeltasLocal + qtyVal
NextApply:
    Next delta
End Function

Private Function ApplyMadeToInventoryDeltasLocal(invLo As ListObject, deltas As Collection, ByRef errNotes As String) As Double
    ApplyMadeToInventoryDeltasLocal = 0
    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        ApplyMadeToInventoryDeltasLocal = -1
        Exit Function
    End If
    If deltas Is Nothing Then Exit Function
    If deltas.Count = 0 Then Exit Function

    Dim colMade As Long: colMade = ColumnIndex(invLo, "MADE")
    Dim colTotal As Long: colTotal = ColumnIndex(invLo, "TOTAL INV")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colLastEdited As Long: colLastEdited = ColumnIndex(invLo, "LAST EDITED")
    Dim colTotalLastEdit As Long: colTotalLastEdit = ColumnIndex(invLo, "TOTAL INV LAST EDIT")
    If colMade = 0 Or colTotal = 0 Or colRow = 0 Then
        errNotes = "invSys table missing MADE/TOTAL INV/ROW columns."
        ApplyMadeToInventoryDeltasLocal = -1
        Exit Function
    End If

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextValidate

        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then
            AppendNote errNotes, "invSys ROW " & rowVal & " not found."
            ApplyMadeToInventoryDeltasLocal = -1
            Exit Function
        End If

        Dim madeCell As Range: Set madeCell = invRow.Range.Cells(1, colMade)
        Dim stagedQty As Double: stagedQty = NzDbl(madeCell.Value)
        If qtyVal > stagedQty + 0.0000001 Then
            AppendNote errNotes, "ROW " & rowVal & " only has " & Format$(stagedQty, "0.###") & " staged in MADE but requires " & Format$(qtyVal, "0.###") & "."
            ApplyMadeToInventoryDeltasLocal = -1
            Exit Function
        End If
NextValidate:
    Next delta

    For Each delta In deltas
        rowVal = CLng(delta("ROW"))
        qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextApply

        Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If invRow Is Nothing Then GoTo NextApply
        Set madeCell = invRow.Range.Cells(1, colMade)
        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotal)
        madeCell.Value = NzDbl(madeCell.Value) - qtyVal
        totalCell.Value = NzDbl(totalCell.Value) + qtyVal
        If colLastEdited > 0 Then invRow.Range.Cells(1, colLastEdited).Value = Now
        If colTotalLastEdit > 0 Then invRow.Range.Cells(1, colTotalLastEdit).Value = Now
        ApplyMadeToInventoryDeltasLocal = ApplyMadeToInventoryDeltasLocal + qtyVal
NextApply:
    Next delta
End Function

Private Sub ClearShipmentStageColumns(ByVal invLo As ListObject, ByVal deltas As Collection)
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If deltas.Count = 0 Then Exit Sub

    Dim colShip As Long: colShip = ColumnIndex(invLo, "SHIPMENTS")
    Dim colRow As Long: colRow = ColumnIndex(invLo, "ROW")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    If colShip = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim delta As Variant
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim invRow As ListRow
        If rowVal > 0 And colRow > 0 Then Set invRow = FindInvListRowByRowValue(invLo, rowVal)
        If Not invRow Is Nothing Then
            invRow.Range.Cells(1, colShip).Value = 0
        Else
            ClearShipmentStageByItemCode invLo, colShip, colItemCode, NzStr(delta("ITEM_CODE"))
        End If
NextDelta:
    Next delta
End Sub

Private Sub ClearShipmentStageByItemCode(ByVal invLo As ListObject, ByVal shipmentColumn As Long, ByVal itemCodeColumn As Long, ByVal itemCode As String)
    If invLo Is Nothing Then Exit Sub
    If shipmentColumn <= 0 Or itemCodeColumn <= 0 Then Exit Sub
    If Len(Trim$(itemCode)) = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If StrComp(NzStr(invLo.DataBodyRange.Cells(r, itemCodeColumn).Value), itemCode, vbTextCompare) = 0 Then
            invLo.DataBodyRange.Cells(r, shipmentColumn).Value = 0
            Exit Sub
        End If
    Next r
End Sub

Private Function StageComponentsToUsed(invLo As ListObject, aggBom As ListObject, ByRef errNotes As String, Optional ByRef logEntries As Collection) As Double
    StageComponentsToUsed = 0
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        StageComponentsToUsed = -1
        Exit Function
    End If
    If aggBom Is Nothing Or aggBom.DataBodyRange Is Nothing Then Exit Function

    Dim cQtyAgg As Long: cQtyAgg = ColumnIndex(aggBom, "QUANTITY")
    Dim cRowAgg As Long: cRowAgg = ColumnIndex(aggBom, "ROW")
    If cQtyAgg = 0 Or cRowAgg = 0 Then
        errNotes = "AggregateBoxBOM missing QUANTITY/ROW columns."
        StageComponentsToUsed = -1
        Exit Function
    End If

    Dim colUsedInv As Long: colUsedInv = ColumnIndex(invLo, "USED")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    If colUsedInv = 0 Then
        errNotes = "invSys table missing USED column."
        StageComponentsToUsed = -1
        Exit Function
    End If
    If logEntries Is Nothing Then Set logEntries = New Collection

    Dim arr As Variant
    arr = aggBom.DataBodyRange.Value
    Dim requirements As Object: Set requirements = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As Long: rowVal = NzLng(arr(r, cRowAgg))
        Dim qtyNeeded As Double: qtyNeeded = NzDbl(arr(r, cQtyAgg))
        If rowVal = 0 Or qtyNeeded <= 0 Then GoTo NextComponent
        Dim reqKey As String: reqKey = CStr(rowVal)
        If requirements.Exists(reqKey) Then
            requirements(reqKey) = NzDbl(requirements(reqKey)) + qtyNeeded
        Else
            requirements.Add reqKey, qtyNeeded
        End If
NextComponent:
    Next r
    If requirements.Count = 0 Then Exit Function

    Dim key As Variant
    For Each key In requirements.Keys
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, CLng(key))
        If invRow Is Nothing Then
            AppendNote errNotes, "invSys ROW " & key & " not found; staging aborted."
            StageComponentsToUsed = -1
            Exit Function
        End If
        Dim usedCell As Range: Set usedCell = invRow.Range.Cells(1, colUsedInv)
        Dim qtyStage As Double: qtyStage = NzDbl(requirements(key))
        Dim newUsed As Double: newUsed = NzDbl(usedCell.Value) + qtyStage
        usedCell.Value = newUsed
        StageComponentsToUsed = StageComponentsToUsed + qtyStage

        Dim itemCode As String, itemName As String
        If colItemCode > 0 Then itemCode = NzStr(invRow.Range.Cells(1, colItemCode).Value)
        If colItemName > 0 Then itemName = NzStr(invRow.Range.Cells(1, colItemName).Value)
        logEntries.Add Array("BTN_CONFIRM_INV_STAGE", CLng(key), itemCode, itemName, qtyStage, newUsed)
    Next key
End Function

Private Sub AppendNote(ByRef target As String, ByVal text As String)
    If Len(text) = 0 Then Exit Sub
    If Len(target) > 0 Then
        target = target & vbCrLf & text
    Else
        target = text
    End If
End Sub

Private Function infovalue(info As Object, field As String) As Variant
    If info Is Nothing Then Exit Function
    If info.Exists(field) Then infovalue = info(field)
End Function

Private Function SortedKeys(dict As Object) As Variant
    If dict Is Nothing Then Exit Function
    Dim keys As Variant: keys = dict.Keys
    If Not IsArray(keys) Then
        SortedKeys = keys
        Exit Function
    End If
    Dim i As Long, j As Long
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(keys(j)) < CLng(keys(i)) Then
                Dim tmp As Variant
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
    SortedKeys = keys
End Function

Private Function ResolveRowFromCaches(itemName As String, nameCache As Object) As Long
    If nameCache Is Nothing Then Exit Function
    Dim key As String: key = LCase$(Trim$(itemName))
    If key = "" Then Exit Function
    If nameCache.Exists(key) Then
        ResolveRowFromCaches = CLng(nameCache(key))
    End If
End Function

Private Sub WriteArrayToTable(lo As ListObject, arr As Variant)
    If lo Is Nothing Then Exit Sub
    If IsEmpty(arr) Then Exit Sub
    Dim rowsNeeded As Long
    On Error Resume Next
    rowsNeeded = UBound(arr, 1)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    If rowsNeeded <= 0 Then
        ClearListObjectData lo
        Exit Sub
    End If
    Dim currentRows As Long
    If lo.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = lo.DataBodyRange.Rows.Count
    End If
    Dim diff As Long
    If currentRows < rowsNeeded Then
        For diff = 1 To rowsNeeded - currentRows
            lo.ListRows.Add
        Next diff
    ElseIf currentRows > rowsNeeded Then
        For diff = currentRows To rowsNeeded + 1 Step -1
            lo.ListRows(diff).Delete
        Next diff
    End If
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.Value = arr
End Sub

Private Sub LogShippingChanges(ByVal logTableName As String, logEntries As Collection)
    On Error GoTo ErrHandler
    If logEntries Is Nothing Then Exit Sub
    If logEntries.Count = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = SheetExists(logTableName)
    If ws Is Nothing Then Exit Sub
    Dim tbl As ListObject: Set tbl = GetListObject(ws, logTableName)
    If tbl Is Nothing Then Exit Sub
    Dim entry As Variant
    Dim newRow As ListRow
    For Each entry In logEntries
        Set newRow = tbl.ListRows.Add
        With newRow.Range
            .Cells(1, 1).Value = modUR_Snapshot.GenerateGUID()
            .Cells(1, 2).Value = modRoleEventWriter.ResolveCurrentUserId()
            .Cells(1, 3).Value = entry(0)
            .Cells(1, 4).Value = entry(1)
            .Cells(1, 5).Value = entry(2)
            .Cells(1, 6).Value = entry(3)
            .Cells(1, 7).Value = entry(4)
            .Cells(1, 8).Value = entry(5)
            .Cells(1, 9).Value = Now
        End With
    Next entry
    Exit Sub
ErrHandler:
    Debug.Print "LogShippingChanges error (" & logTableName & "): " & Err.Description
End Sub

Private Function GetBomTableByRow(ByVal rowValue As Long) As ListObject
    Dim wsBOM As Worksheet: Set wsBOM = SheetExists(SHEET_BOM)
    If wsBOM Is Nothing Then Exit Function
    On Error Resume Next
    Set GetBomTableByRow = wsBOM.ListObjects(BomTableNameFromRow(rowValue))
    On Error GoTo 0
End Function

Public Function NzStr(v As Variant) As String
    If IsError(v) Then
        NzStr = ""
    ElseIf IsNull(v) Then
        NzStr = ""
    ElseIf IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Public Function NzDbl(v As Variant) As Double
    Dim textVal As String

    If IsError(v) Then
        NzDbl = 0#
    ElseIf IsNull(v) Then
        NzDbl = 0#
    ElseIf IsEmpty(v) Then
        NzDbl = 0#
    Else
        textVal = Trim$(CStr(v))
        If textVal = "" Then
            NzDbl = 0#
        ElseIf Not IsNumeric(textVal) Then
            NzDbl = 0#
        Else
            NzDbl = CDbl(textVal)
        End If
    End If
End Function

Public Function NzLng(v As Variant) As Long
    Dim textVal As String

    If IsError(v) Then
        NzLng = 0
    ElseIf IsNull(v) Then
        NzLng = 0
    ElseIf IsEmpty(v) Then
        NzLng = 0
    Else
        textVal = Trim$(CStr(v))
        If textVal = "" Then
            NzLng = 0
        ElseIf Not IsNumeric(textVal) Then
            NzLng = 0
        Else
            NzLng = CLng(CDbl(textVal))
        End If
    End If
End Function

' ===== Workbook/setup helpers (migrated from modTS_Data) =====
Public Sub SetupAllHandlers()
    On Error Resume Next
    mNextInvSysRow = 0
    mAggDirty = True
    ClearTableFilters
    modGlobals.InitializeGlobalVariables
    Application.OnKey "{F3}", "modGlobals.OpenItemSearchForCurrentCell"
    InitializeShipmentsUI
    On Error GoTo 0
End Sub

Public Sub GenerateRowNumbers()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowNumCol As Long
    Dim i As Long
    Dim maxRowNum As Long
    Dim newCol As ListColumn

    Set ws = GetInventoryWorksheetShipping()
    If ws Is Nothing Then GoTo ErrorHandler
    Set tbl = ws.ListObjects("invSys")

    On Error Resume Next
    rowNumCol = tbl.ListColumns("ROW").Index
    On Error GoTo ErrorHandler
    If rowNumCol = 0 Then
        Set newCol = tbl.ListColumns.Add
        newCol.Name = "ROW"
        rowNumCol = newCol.Index
    End If

    maxRowNum = 0
    For i = 1 To tbl.ListRows.Count
        If IsNumeric(tbl.DataBodyRange(i, rowNumCol).Value) Then
            maxRowNum = Application.WorksheetFunction.Max(maxRowNum, tbl.DataBodyRange(i, rowNumCol).Value)
        End If
    Next i

    For i = 1 To tbl.ListRows.Count
        If Trim$(tbl.DataBodyRange(i, rowNumCol).Value & "") = "" Then
            maxRowNum = maxRowNum + 1
            tbl.DataBodyRange(i, rowNumCol).Value = maxRowNum
        End If
    Next i
    MsgBox "Row numbers have been generated successfully.", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error generating row numbers: " & Err.Description, vbExclamation
End Sub

Public Function IsInItemsColumn(Target As Range) As Boolean
    IsInItemsColumn = False
    On Error Resume Next

    Dim lo As ListObject
    Set lo = Target.ListObject
    If lo Is Nothing Then Exit Function

    If lo.Name <> "ShipmentsTally" And lo.Name <> "ReceivedTally" Then Exit Function

    Dim itemsCol As ListColumn
    Set itemsCol = lo.ListColumns("ITEMS")
    On Error GoTo 0
    If itemsCol Is Nothing Then Exit Function

    If Target.Column = itemsCol.Range.Column Then
        If Target.Row > lo.HeaderRowRange.Row Then
            IsInItemsColumn = True
        End If
    End If
End Function

Public Sub ClearTableFilters()
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error Resume Next
    Set wb = ResolveShippingWorkbook()
    If wb Is Nothing Then Set wb = ThisWorkbook

    Set ws = WorkbookSheetExistsShipping(wb, "ShipmentsTally")
    If Not ws Is Nothing Then
        If Not ws.ListObjects("ShipmentsTally") Is Nothing Then
            ws.ListObjects("ShipmentsTally").AutoFilter.ShowAllData
        End If
        If Not ws.ListObjects("invSysData_Shipping") Is Nothing Then
            ws.ListObjects("invSysData_Shipping").AutoFilter.ShowAllData
        End If
    End If

    Set ws = WorkbookSheetExistsShipping(wb, "ReceivedTally")
    If Not ws Is Nothing Then
        If Not ws.ListObjects("ReceivedTally") Is Nothing Then
            ws.ListObjects("ReceivedTally").AutoFilter.ShowAllData
        End If
        If Not ws.ListObjects("invSysData_Receiving") Is Nothing Then
            ws.ListObjects("invSysData_Receiving").AutoFilter.ShowAllData
        End If
    End If
    On Error GoTo 0
End Sub

Public Function LoadItemList() As Variant
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowCount As Long
    Dim result As Variant
    Dim i As Long

    Set ws = GetInventoryWorksheetShipping()
    If ws Is Nothing Then GoTo ErrorHandler
    Set tbl = ws.ListObjects("invSys")
    If tbl Is Nothing Then GoTo ErrorHandler

    rowCount = tbl.ListRows.Count
    If rowCount = 0 Then GoTo ErrorHandler

    ReDim result(1 To rowCount, 0 To 4)

    Dim itemCodeCol As Integer, rowCol As Integer, itemCol As Integer
    Dim locCol As Integer
    On Error Resume Next
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    rowCol = tbl.ListColumns("ROW").Index
    itemCol = tbl.ListColumns("ITEM").Index
    locCol = tbl.ListColumns("LOCATION").Index
    On Error GoTo ErrorHandler
    If itemCodeCol = 0 Or rowCol = 0 Or itemCol = 0 Then GoTo ErrorHandler

    For i = 1 To rowCount
        result(i, 0) = tbl.DataBodyRange.Cells(i, rowCol).Value
        result(i, 1) = tbl.DataBodyRange.Cells(i, itemCodeCol).Value
        result(i, 2) = tbl.DataBodyRange.Cells(i, itemCol).Value
        If locCol > 0 Then
            result(i, 3) = tbl.DataBodyRange.Cells(i, locCol).Value
        End If
    Next i
    LoadItemList = result
    Exit Function
ErrorHandler:
    LoadItemList = Empty
End Function
