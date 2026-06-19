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
Private Const SHEET_SHIPPING_BACKEND As String = "ShippingBackend"
Private Const SHEET_BOM_TABLES As String = "ShippingBOMTables"

Private Const TABLE_SHIPMENTS As String = "ShipmentsTally"
Private Const TABLE_NOTSHIPPED As String = "NotShipped"
Private Const SHEET_SHIPPING_RESERVATIONS As String = "ShippingReservations"
Private Const TABLE_SHIPPING_RESERVATIONS As String = "tblShippingReservations"
Private Const TABLE_AGG_BOM As String = "AggregateBoxBOM"
Private Const TABLE_AGG_PACK As String = "AggregatePackages"
Private Const TABLE_BOX_BUILDER As String = "BoxBuilder"
Private Const TABLE_BOX_BOM As String = "BoxBOM"
Private Const TABLE_BOX_BOM_VERSIONS As String = "BoxBOMVersions"
Private Const TABLE_CHECK_INV As String = "Check_invSys"
Private Const TABLE_SHIPPING_BOM_VIEW As String = "ShippingBOMView"
Private Const TABLE_CANONICAL_SHIPPING_BOM As String = "tblShippingBOM"
Private Const COL_BOXBOM_ITEM As String = "ITEM"
Private Const COL_CURRENT_INV As String = "CURRENT INV"
Private Const COL_SHIPMENT_LINE_ID As String = "LINE_ID"
Private Const COL_SHIPMENT_RESERVE_EVENT_ID As String = "SERVER_RESERVE_EVENT_ID"
Private Const COL_BOM_VERSION As String = "BOM VERSION"
Private Const EVENT_TYPE_SHIP As String = "SHIP"
Private Const EVENT_TYPE_SHIP_RESERVE As String = "SHIP_RESERVE"
Private Const EVENT_TYPE_SHIP_RELEASE As String = "SHIP_RELEASE"
Private Const EVENT_TYPE_ADMIN_SHIPMENT_RECONCILE As String = "ADMIN_SHIPMENT_RECONCILE"
Private Const EVENT_TYPE_BOX_BUILD As String = "BOX_BUILD"
Private Const EVENT_TYPE_BOX_UNBOX As String = "BOX_UNBOX"
Private Const SHIP_RESERVATION_ACTIVE As String = "ACTIVE"
Private Const SHIP_RESERVATION_RELEASED As String = "RELEASED"
Private Const SHIP_RESERVATION_COMPLETED As String = "COMPLETED"

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
Private mSelectedBoxBomVersionLabel As String
Private mSelectedBoxBomVersionPackageRow As Long
Private mSelectedBoxBomVersionWorkbookName As String
Private mSelectedBoxBomVersionWorksheetName As String
Private mHandlingShippingSheetChange As Boolean
Private mLastBoxBomVersionSignature As String
Private mLastBoxBomVersionWorkbookName As String
Private mLastBoxBomVersionWorksheetName As String
Private mLastSavedShippingBomVersion As Long
Private mGeneratedIdentityEditWorkbookName As String
Private mGeneratedIdentityEditWorksheetName As String
Private mGeneratedIdentityEditAddress As String
Private mSuppressGeneratedIdentityEditGuard As Boolean
Private mLastComponentPickerStatus As String
Private mPendingBoxVersionInventoryOverlay As Object
Private mPendingBoxVersionInventoryOverlayBaseline As Object
Private mPendingBoxVersionInventoryOverlayPath As String

Private Const BOX_VERSION_SAVE_CANCEL As Long = 0
Private Const BOX_VERSION_SAVE_UPDATE As Long = 1
Private Const BOX_VERSION_SAVE_NEW As Long = 2

' ===== public entry points =====
Public Sub InitializeShipmentsUI()
    InitializeShipmentsUiForWorkbook Application.ActiveWorkbook
End Sub

Public Sub InitializeShipmentsUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    On Error GoTo CleanExit

    Dim surfaceReport As String
    Dim wb As Workbook
    Dim prevEvents As Boolean

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    mSuppressGeneratedIdentityEditGuard = True
    ClearGeneratedIdentityEditSelection

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

CleanExit:
    mSuppressGeneratedIdentityEditGuard = False
    Application.EnableEvents = prevEvents
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
    Dim inventoryState As Object
    Set inventoryState = CaptureBoxMakerCurrentInventoryState(loBuilder, loBom)

    Dim packageReturned As Double
    Dim componentsReturned As Double
    Dim errNotes As String
    If Not ApplyBoxUnboxedFromBuilder(loBuilder, loBom, invLo, packageReturned, componentsReturned, errNotes) Then
        If errNotes = "" Then errNotes = "Box could not be unboxed."
        MsgBox errNotes, vbExclamation
        Exit Sub
    End If

    RefreshBoxMakerCurrentInventory ws
    ApplyBoxUnboxExpectedCurrentInventoryDisplay loBuilder, loBom, inventoryState
    ResetBoxMakerQuantities loBuilder, loBom
    InvalidateAggregates True, True

    ShowShippingStatus "Box unboxed. Removed " & Format$(packageReturned, "0.###") & " shippable units; returned " & Format$(componentsReturned, "0.###") & " component units to TOTAL INV."
    Exit Sub

ErrHandler:
    MsgBox "BTN_BOX_UNBOXED failed: " & Err.Description, vbCritical
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
    EnsureColumnExists lo, COL_SHIPMENT_LINE_ID
    EnsureColumnExists lo, COL_SHIPMENT_RESERVE_EVENT_ID
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
    EnsureShipmentLineIds lo
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
    Dim anchorRow As Long
    Dim anchorCol As Long

    If wb Is Nothing Then Exit Sub
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    anchorRow = ws.Range(SHIP_LAYOUT_BUILDER_ADDR).Row
    anchorCol = ws.Range(SHIP_LAYOUT_BUILDER_ADDR).Column

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)

    MoveListObjectToRowColShipping loBuilder, anchorRow, anchorCol
    ArrangeBoxBuilderBandShipping loBuilder, loBom
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

Public Sub BtnOpenBoxBuilder()
    On Error GoTo ErrHandler

    Dim ws As Worksheet

    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then Exit Sub

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing _
       Or GetListObject(ws, TABLE_BOX_BUILDER) Is Nothing _
       Or GetListObject(ws, TABLE_BOX_BOM) Is Nothing Then
        InitializeShipmentsUiForWorkbook Application.ActiveWorkbook
    End If

    frmShippingBoxBuilder.InitializeFromShipping
    frmShippingBoxBuilder.Show
    Exit Sub

ErrHandler:
    MsgBox "BOX_BUILDER failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnOpenBoxMaker()
    On Error GoTo ErrHandler

    Dim ws As Worksheet

    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then Exit Sub

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing _
       Or GetListObject(ws, TABLE_BOX_BUILDER) Is Nothing _
       Or GetListObject(ws, TABLE_BOX_BOM) Is Nothing Then
        InitializeShipmentsUiForWorkbook Application.ActiveWorkbook
    End If

    frmShippingBoxMaker.InitializeFromShipping
    frmShippingBoxMaker.Show
    Exit Sub

ErrHandler:
    MsgBox "BOX_MAKER failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnOpenShipmentsForm()
    On Error GoTo ErrHandler

    Dim ws As Worksheet

    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then Exit Sub

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing _
       Or GetListObject(ws, TABLE_SHIPMENTS) Is Nothing _
       Or GetListObject(ws, TABLE_NOTSHIPPED) Is Nothing Then
        InitializeShipmentsUiForWorkbook Application.ActiveWorkbook
    End If

    frmShipmentsTally.InitializeFromShipping
    frmShipmentsTally.Show
    Exit Sub

ErrHandler:
    MsgBox "SHIPMENTS failed: " & Err.Description, vbCritical
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
    FillBlankBoxBomVersionShipping loBom

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
    FillBlankBoxBomVersionShipping loBom

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim components As Collection
    Dim syncNotes As String
    Dim saveVersionLabel As String
    Dim replaceVersion As Long
    Dim versionChoice As Long
    Dim forceNewVersion As Boolean

    saveVersionLabel = ResolveBoxBomSaveVersionLabel(ws, loBom)
    If saveVersionLabel <> "" Then
        versionChoice = PromptBoxVersionSaveChoiceShipping(boxName, saveVersionLabel)
        If versionChoice = BOX_VERSION_SAVE_CANCEL Then Exit Sub
        If versionChoice = BOX_VERSION_SAVE_UPDATE Then
            replaceVersion = BomVersionNumberFromLabel(saveVersionLabel)
        ElseIf versionChoice = BOX_VERSION_SAVE_NEW Then
            forceNewVersion = True
        End If
    End If

    Set components = CollectBomComponents(loBom, invLo, syncNotes, saveVersionLabel)
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
    If Not SaveShippingBomToRuntime(ws.Parent, boxRowValue, boxName, boxUOM, boxLoc, boxDesc, components, bomReport, replaceVersion, forceNewVersion) Then
        If bomReport = "" Then bomReport = "Unable to save Shipping BOM to the selected warehouse runtime."
        MsgBox bomReport, vbCritical
        Exit Sub
    End If
    RefreshShippingBomViewForWorkbook ws.Parent, bomReport

    Dim finalMsg As String
    If InStr(1, bomReport, "Shipping BOM unchanged:", vbTextCompare) > 0 Then
        finalMsg = bomReport
    Else
        finalMsg = "Saved BOM '" & boxName & "' to warehouse runtime (invSys ROW " & boxRowValue & ", " & components.count & " components)."
    End If
    If Len(syncNotes) > 0 Then
        finalMsg = finalMsg & vbCrLf & syncNotes
    End If
    If Len(bomReport) > 0 And InStr(1, finalMsg, bomReport, vbTextCompare) = 0 Then
        finalMsg = finalMsg & vbCrLf & bomReport
    End If
    MsgBox finalMsg, vbInformation

    EnsureTableHasRow loMeta
    EnsureTableHasRow loBom
    RefreshBoxBomVersionList ws, boxRowValue
    RebuildBoxBomVersionListFromDisplayedBom ws, boxRowValue
    InvalidateAggregates True
    Exit Sub

ErrHandler:
    MsgBox "BTN_SAVE_BOX failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnDeleteBoxVersion()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim loVersions As ListObject
    Dim versionLabel As String
    Dim boxName As String
    Dim packageRow As Long
    Dim runtimeMax As Long
    Dim report As String
    Dim deleteReport As String
    Dim loBuilder As ListObject

    If Not modRoleUiAccess.RequireCurrentUserCapability("ADMIN_MAINT") Then Exit Sub

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then
        MsgBox "BoxBuilder table was not found.", vbExclamation
        Exit Sub
    End If
    Set loVersions = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    If loVersions Is Nothing Then
        MsgBox "No BoxBOMVersions table is available.", vbExclamation
        Exit Sub
    End If

    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    If boxName = "" Then
        MsgBox "BoxBuilder Box Name is required before deleting a version.", vbExclamation
        Exit Sub
    End If
    packageRow = FindShippingBomPackageRowByName(ws.Parent, boxName, runtimeMax)
    If packageRow <= 0 Then
        MsgBox "Saved box '" & boxName & "' was not found in ShippingBOM runtime.", vbExclamation
        Exit Sub
    End If

    versionLabel = SelectedBoxBomVersionLabel(ws, loVersions, packageRow)
    If versionLabel = "" Then
        MsgBox "Select a version row in BoxBOMVersions before deleting.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Delete " & boxName & " " & versionLabel & " from the warehouse Shipping BOM?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    If Not DeleteShippingBomVersionFromRuntime(ws.Parent, packageRow, BomVersionNumberFromLabel(versionLabel), report) Then
        If report = "" Then report = "Could not delete selected Shipping BOM version."
        MsgBox report, vbCritical
        Exit Sub
    End If

    deleteReport = report
    DeleteLocalBoxBomRowsForVersion ws, versionLabel
    DeleteLocalBoxBomVersionSummaryRow loVersions, versionLabel
    DeleteLocalShippingBomViewRowsForVersion ws, packageRow, versionLabel
    ClearSelectedBoxBomVersionIfMatches ws, packageRow, versionLabel
    RefreshShippingBomViewForWorkbook ws.Parent, report
    DeleteLocalShippingBomViewRowsForVersion ws, packageRow, versionLabel
    RefreshBoxBomVersionList ws, packageRow
    MsgBox deleteReport, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BTN_DELETE_BOX_VERSION failed: " & Err.Description, vbCritical
End Sub

Private Function PromptBoxVersionSaveChoiceShipping(ByVal boxName As String, ByVal versionLabel As String) As Long
    On Error GoTo FallbackPrompt

    frmBoxVersionSaveChoice.InitializeChoice boxName, versionLabel
    frmBoxVersionSaveChoice.Show vbModal
    PromptBoxVersionSaveChoiceShipping = frmBoxVersionSaveChoice.Choice
    Unload frmBoxVersionSaveChoice
    Exit Function

FallbackPrompt:
    Dim response As VbMsgBoxResult
    On Error Resume Next
    Unload frmBoxVersionSaveChoice
    On Error GoTo 0
    response = MsgBox("Save edits to " & boxName & " " & versionLabel & "?" & vbCrLf & vbCrLf & _
                      "Yes updates the selected version." & vbCrLf & _
                      "No saves these rows as a new version.", _
                      vbQuestion + vbYesNoCancel, _
                      "Save Box Version")
    Select Case response
        Case vbYes
            PromptBoxVersionSaveChoiceShipping = BOX_VERSION_SAVE_UPDATE
        Case vbNo
            PromptBoxVersionSaveChoiceShipping = BOX_VERSION_SAVE_NEW
        Case Else
            PromptBoxVersionSaveChoiceShipping = BOX_VERSION_SAVE_CANCEL
    End Select
End Function

Public Sub BtnDeleteBox()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim boxName As String
    Dim packageRow As Long
    Dim runtimeMax As Long
    Dim report As String
    Dim deleteReport As String
    Dim loBuilder As ListObject
    Dim loVersions As ListObject

    If Not modRoleUiAccess.RequireCurrentUserCapability("ADMIN_MAINT") Then Exit Sub

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then
        MsgBox "BoxBuilder table was not found.", vbExclamation
        Exit Sub
    End If
    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    If boxName = "" Then
        MsgBox "BoxBuilder Box Name is required before deleting a box.", vbExclamation
        Exit Sub
    End If

    packageRow = FindShippingBomPackageRowByName(ws.Parent, boxName, runtimeMax)
    If packageRow <= 0 Then
        MsgBox "Saved box '" & boxName & "' was not found in ShippingBOM runtime.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Delete all saved BOM versions for '" & boxName & "' from the warehouse Shipping BOM?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    If Not DeleteShippingBomPackageFromRuntime(ws.Parent, packageRow, report) Then
        If report = "" Then report = "Could not delete selected Shipping BOM box."
        MsgBox report, vbCritical
        Exit Sub
    End If

    deleteReport = report
    RefreshShippingBomViewForWorkbook ws.Parent, report
    Set loVersions = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    If Not loVersions Is Nothing Then ClearListObjectData loVersions
    MsgBox deleteReport, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BTN_DELETE_BOX failed: " & Err.Description, vbCritical
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
    PersistActiveShipmentRowsLocal loShip
    PersistHoldRowsLocal loHold
    InvalidateAggregates

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
    If aggPack Is Nothing Then
        MsgBox "AggregatePackages table not found.", vbInformation
        Exit Sub
    End If
    If aggPack.DataBodyRange Is Nothing Then
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
    If Not RunShippingRuntimeQueueRefresh(ws.Parent, ResolveCurrentShippingWarehouseId(), runtimeReport) Then
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

Public Function ValidateShipmentsSentStagingFromCurrentWorkbook() As String
    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim deltas As Collection
    Dim errNotes As String

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        ValidateShipmentsSentStagingFromCurrentWorkbook = "ShipmentsTally sheet not found."
        Exit Function
    End If

    Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        ValidateShipmentsSentStagingFromCurrentWorkbook = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    If BuildQueueableShipmentsSentDeltas(invLo, ws, deltas, errNotes) Then
        ValidateShipmentsSentStagingFromCurrentWorkbook = "OK"
    Else
        If errNotes = "" Then errNotes = "No staged shipments found in invSys.SHIPMENTS."
        ValidateShipmentsSentStagingFromCurrentWorkbook = errNotes
    End If
End Function

Public Function ValidateToShipmentsFromCurrentWorkbook() As String
    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim aggPack As ListObject
    Dim deltas As Collection
    Dim errNotes As String

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        ValidateToShipmentsFromCurrentWorkbook = "ShipmentsTally sheet not found."
        Exit Function
    End If

    Set invLo = GetInvSysTable()
    Set aggPack = GetListObject(ws, TABLE_AGG_PACK)
    If invLo Is Nothing Then
        ValidateToShipmentsFromCurrentWorkbook = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If aggPack Is Nothing Then
        ValidateToShipmentsFromCurrentWorkbook = "AggregatePackages table not found."
        Exit Function
    End If
    If aggPack.DataBodyRange Is Nothing Then
        ValidateToShipmentsFromCurrentWorkbook = "AggregatePackages has no rows to stage."
        Exit Function
    End If

    Set deltas = BuildShipmentDeltaPacket(invLo, aggPack, errNotes)
    If deltas Is Nothing Then
        If errNotes = "" Then errNotes = "No additional shipments required; Shipments column already meets demand."
        ValidateToShipmentsFromCurrentWorkbook = errNotes
    ElseIf deltas.Count = 0 Then
        If errNotes = "" Then errNotes = "No additional shipments required; Shipments column already meets demand."
        ValidateToShipmentsFromCurrentWorkbook = errNotes
    Else
        ValidateToShipmentsFromCurrentWorkbook = "OK"
    End If
End Function

Public Function ValidateBoxesMadeFromCurrentWorkbook() As String
    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim aggBom As ListObject
    Dim shortage As String

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        ValidateBoxesMadeFromCurrentWorkbook = "ShipmentsTally sheet not found."
        Exit Function
    End If

    Set invLo = GetInvSysTable()
    Set aggBom = GetListObject(ws, TABLE_AGG_BOM)
    If invLo Is Nothing Then
        ValidateBoxesMadeFromCurrentWorkbook = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    If ValidateComponentInventory(invLo, aggBom, shortage) Then
        ValidateBoxesMadeFromCurrentWorkbook = "OK"
    Else
        ValidateBoxesMadeFromCurrentWorkbook = shortage
    End If
End Function

Public Function ValidateConfirmInventoryFromCurrentWorkbook() As String
    Dim ws As Worksheet

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        ValidateConfirmInventoryFromCurrentWorkbook = "ShipmentsTally sheet not found."
        Exit Function
    End If

    If UseExistingInventoryEnabled(ws) Then
        ValidateConfirmInventoryFromCurrentWorkbook = "Use existing inventory is enabled. Skip Confirm inventory and go to 'To Shipments'."
    Else
        ValidateConfirmInventoryFromCurrentWorkbook = "OK"
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

Private Function QueueShipmentsReserveEvent(ByVal deltas As Collection, ByRef errNotes As String, ByRef eventIdOut As String) As Boolean
    Dim payloadJson As String

    payloadJson = BuildPayloadJsonFromDeltas(deltas, "RESERVED")
    If payloadJson = "" Then
        If errNotes = "" Then errNotes = "No shipment reserve payload rows were generated."
        Exit Function
    End If

    QueueShipmentsReserveEvent = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_SHIP_RESERVE, _
        "", _
        payloadJson, _
        "BTN_TO_SHIPMENTS_RESERVE", _
        eventIdOut, _
        errNotes)
End Function

Private Function QueueShipmentsReleaseEvent(ByVal deltas As Collection, ByRef errNotes As String, ByRef eventIdOut As String) As Boolean
    Dim payloadJson As String

    payloadJson = BuildPayloadJsonFromDeltas(deltas, "RELEASED")
    If payloadJson = "" Then
        If errNotes = "" Then errNotes = "No shipment release payload rows were generated."
        Exit Function
    End If

    QueueShipmentsReleaseEvent = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_SHIP_RELEASE, _
        "", _
        payloadJson, _
        "BTN_NOT_SHIPPED_RELEASE", _
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

Private Function QueueBoxUnboxEventFromBuilder(ByVal loBuilder As ListObject, _
                                               ByVal loBom As ListObject, _
                                               ByVal invLo As ListObject, _
                                               ByRef componentsReturned As Double, _
                                               ByRef packageRemoved As Double, _
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
    Dim currentInv As Variant
    Dim foundCurrent As Boolean
    Dim snapshotCache As Object

    errNotes = ""
    componentsReturned = 0
    packageRemoved = 0
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

    currentInv = ResolveCurrentInventoryValue(loBuilder.Parent, invLo, packageRow, boxName, foundCurrent, snapshotCache)
    If foundCurrent Then
        If boxQty > NzDbl(currentInv) + 0.0000001 Then
            errNotes = "Box '" & boxName & "' only has " & Format$(NzDbl(currentInv), "0.###") & " in inventory but needs " & Format$(boxQty, "0.###") & "."
            Exit Function
        End If
    End If

    If itemCode = "" Then itemCode = boxName
    uomVal = Trim$(NzStr(ValueFromTable(loBuilder, "UOM")))
    locationVal = Trim$(NzStr(ValueFromTable(loBuilder, "LOCATION")))
    descrVal = Trim$(NzStr(ValueFromTable(loBuilder, "DESCRIPTION")))

    Set payloadItems = New Collection
    If Not AddBoxUnboxComponentPayloadItems(loBom, invLo, payloadItems, componentsReturned, errNotes) Then Exit Function
    AddBoxBuildPayloadItem payloadItems, packageRow, itemCode, boxName, boxQty, uomVal, locationVal, descrVal, "UNMADE"
    packageRemoved = boxQty

    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        errNotes = "No BoxMaker unbox payload rows were generated."
        Exit Function
    End If

    QueueBoxUnboxEventFromBuilder = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_BOX_UNBOX, _
        "", _
        payloadJson, _
        "BTN_BOX_UNBOXED", _
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
    Dim cArea As Long
    Dim cCarrier As Long
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

Private Function AddBoxUnboxComponentPayloadItems(ByVal loBom As ListObject, _
                                                  ByVal invLo As ListObject, _
                                                  ByVal payloadItems As Collection, _
                                                  ByRef returnedTotal As Double, _
                                                  ByRef errNotes As String) As Boolean
    Dim cItem As Long
    Dim cCode As Long
    Dim cRow As Long
    Dim cQty As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long
    Dim cTotalInv As Long
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

        AddBoxBuildPayloadItem payloadItems, rowVal, itemCode, itemName, qtyVal, uomVal, locVal, descVal, "RETURNED"
        returnedTotal = returnedTotal + qtyVal
NextRow:
    Next r

    If returnedTotal <= 0 Then
        errNotes = "No component quantities were found in BoxBOM."
        Exit Function
    End If
    AddBoxUnboxComponentPayloadItems = True
End Function

Private Sub AddBoxBuildPayloadItem(ByVal payloadItems As Collection, _
                                   ByVal rowVal As Long, _
                                   ByVal itemCode As String, _
                                   ByVal itemName As String, _
                                   ByVal qtyVal As Double, _
                                   ByVal uomVal As String, _
                                   ByVal locationVal As String, _
                                   ByVal descriptionVal As String, _
                                   ByVal ioType As String, _
                                   Optional ByVal bomVersionLabel As String = "")
    Dim payloadItem As Object

    If payloadItems Is Nothing Then Exit Sub
    Set payloadItem = modRoleEventWriter.CreatePayloadItem(rowVal, itemCode, qtyVal, locationVal, itemName, ioType)
    payloadItem("ROW") = rowVal
    payloadItem("ITEM_CODE") = itemCode
    payloadItem("ITEM") = itemName
    payloadItem("UOM") = uomVal
    payloadItem("DESCRIPTION") = descriptionVal
    payloadItem("LOCATION") = locationVal
    If Trim$(bomVersionLabel) <> "" Then payloadItem("BomVersionLabel") = NormalizeBoxBomVersionLabelShipping(bomVersionLabel)
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
    If ShippingPickerTargetTableName(targetCell) = "boxbom" Then
        mDynSearch.ShowForShippingComponentCell targetCell
    Else
        mDynSearch.ShowForCell targetCell
    End If
    Exit Sub
ErrHandler:
    MsgBox "Shipping item picker is unavailable: " & Err.Description, vbExclamation
End Sub

Private Function ShippingPickerTargetTableName(ByVal targetCell As Range) As String
    Dim lo As ListObject

    If targetCell Is Nothing Then Exit Function
    On Error Resume Next
    Set lo = targetCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    ShippingPickerTargetTableName = LCase$(Trim$(lo.Name))
End Function

Private Function ShouldRefreshShippingBomBeforePicker(ByVal targetCell As Range) As Boolean
    Dim lo As ListObject
    Dim tableName As String

    If targetCell Is Nothing Then Exit Function
    On Error Resume Next
    Set lo = targetCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function

    tableName = LCase$(Trim$(lo.Name))
    ShouldRefreshShippingBomBeforePicker = False
End Function

Public Sub HandleShippingSelectionChange(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    If target.Cells.CountLarge > 1 Then Exit Sub
    If target.Worksheet Is Nothing Then Exit Sub
    If target.Worksheet.Parent Is Nothing Then Exit Sub
    If target.Worksheet.Parent.IsAddin Then Exit Sub
    If StrComp(target.Worksheet.Name, SHEET_SHIPMENTS, vbTextCompare) <> 0 Then Exit Sub
    TrackGeneratedIdentityEditSelection target

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
            If Not IsBoxMakerMode(target.Worksheet) Then
                On Error Resume Next
                Set targetCol = lo.ListColumns("Box Name")
                On Error GoTo 0
                If targetCol Is Nothing Then Exit Sub
                If target.Column <> targetCol.Range.Column Then Exit Sub
                If target.Row <> lo.HeaderRowRange.Row Then Exit Sub
                Set gSelectedCell = target
                ShowDynamicItemSearch target
                SelectBoxBuilderDataCellForRepeatHeaderPicker lo
                Exit Sub
            End If
            On Error Resume Next
            Set targetCol = lo.ListColumns("Box Name")
            On Error GoTo 0
        Case "boxbomversions"
            HandleBoxBomVersionSelection target
            Exit Sub
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
    If mHandlingShippingSheetChange Then Exit Sub
    If target Is Nothing Then Exit Sub
    If target.Cells.CountLarge > 50 Then Exit Sub
    If target.Worksheet Is Nothing Then Exit Sub
    If target.Worksheet.Parent Is Nothing Then Exit Sub
    If target.Worksheet.Parent.IsAddin Then Exit Sub
    If StrComp(target.Worksheet.Name, SHEET_SHIPMENTS, vbTextCompare) <> 0 Then Exit Sub
    mHandlingShippingSheetChange = True

    If BoxBomGeneratedIdentityColumnWasEdited(target.Worksheet, target) Then
        Application.EnableEvents = False
        On Error Resume Next
        Application.Undo
        On Error GoTo ExitHandler
        MsgBox "Version, ITEM_CODE, and ROW are managed by invSys. Use the picker or box/version controls instead of editing those fields directly.", vbInformation
        GoTo ExitHandler
    End If

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

    If AutoFillBoxBomVersionForChange(target) Then GoTo ExitHandler

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
    If lo Is Nothing Then GoTo ExitHandler
    If lo.DataBodyRange Is Nothing Then GoTo ExitHandler

    On Error Resume Next
    Set qtyCol = lo.ListColumns("Box Name")
    On Error GoTo 0
    If Not qtyCol Is Nothing Then
        Set hit = Application.Intersect(target, qtyCol.DataBodyRange)
        If Not hit Is Nothing Then
            Application.EnableEvents = False
            RefreshBoxBomVersionListForCurrentBuilder target.Worksheet
            GoTo ExitHandler
        End If
    End If

    If Not IsBoxMakerMode(target.Worksheet) Then GoTo ExitHandler

    On Error Resume Next
    Set qtyCol = lo.ListColumns("Quantity")
    On Error GoTo 0
    If qtyCol Is Nothing Then GoTo ExitHandler

    Set hit = Application.Intersect(target, qtyCol.DataBodyRange)
    If hit Is Nothing Then GoTo ExitHandler

    Application.EnableEvents = False
    ReloadBoxMakerBomFromBuilder lo.Parent
ExitHandler:
    Application.EnableEvents = True
    mHandlingShippingSheetChange = False
End Sub

Private Function AutoFillBoxBomVersionForChange(ByVal target As Range) As Boolean
    On Error GoTo CleanFail

    Dim loBom As ListObject
    Dim hit As Range
    Dim packageRow As Long
    Dim runtimeMax As Long
    Dim boxName As String

    If target Is Nothing Then Exit Function
    Set loBom = GetListObject(target.Worksheet, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function

    Set hit = Application.Intersect(target, loBom.DataBodyRange)
    If hit Is Nothing Then Exit Function

    Application.EnableEvents = False
    modUiQuiet.BeginQuietUi target.Worksheet.Parent
    EnsureBoxBomEntryColumns loBom
    FillBlankBoxBomVersionShipping loBom

    RebuildBoxBomVersionsForCurrentBoxShipping target.Worksheet
    AutoFillBoxBomVersionForChange = True
    modUiQuiet.EndQuietUi
    Exit Function

CleanFail:
    On Error Resume Next
    modUiQuiet.EndQuietUi
    On Error GoTo 0
    AutoFillBoxBomVersionForChange = False
End Function

Private Sub TrackGeneratedIdentityEditSelection(ByVal target As Range)
    If BoxBomGeneratedIdentityColumnSelection(target) Then
        mGeneratedIdentityEditWorkbookName = target.Worksheet.Parent.Name
        mGeneratedIdentityEditWorksheetName = target.Worksheet.Name
        mGeneratedIdentityEditAddress = target.Address(False, False)
    Else
        ClearGeneratedIdentityEditSelection
    End If
End Sub

Private Sub ClearGeneratedIdentityEditSelection()
    mGeneratedIdentityEditWorkbookName = ""
    mGeneratedIdentityEditWorksheetName = ""
    mGeneratedIdentityEditAddress = ""
End Sub

Private Function BoxBomGeneratedIdentityColumnSelection(ByVal target As Range) As Boolean
    Dim loBom As ListObject
    Dim colName As Variant
    Dim colIdx As Long

    If target Is Nothing Then Exit Function
    If target.Cells.CountLarge <> 1 Then Exit Function
    Set loBom = GetListObject(target.Worksheet, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function
    If Application.Intersect(target, loBom.DataBodyRange) Is Nothing Then Exit Function

    For Each colName In Array("Version", "ITEM_CODE", "ROW")
        colIdx = ColumnIndex(loBom, CStr(colName))
        If colIdx > 0 Then
            If Not Application.Intersect(target, loBom.DataBodyRange.Columns(colIdx)) Is Nothing Then
                BoxBomGeneratedIdentityColumnSelection = True
                Exit Function
            End If
        End If
    Next colName
End Function

Private Function BoxBomGeneratedIdentityColumnWasEdited(ByVal ws As Worksheet, ByVal target As Range) As Boolean
    Dim loBom As ListObject
    Dim colName As Variant
    Dim colIdx As Long
    Dim hit As Range

    If ws Is Nothing Or target Is Nothing Then Exit Function
    If mSuppressGeneratedIdentityEditGuard Then Exit Function
    If target.Cells.CountLarge <> 1 Then Exit Function
    If StrComp(mGeneratedIdentityEditWorkbookName, ws.Parent.Name, vbTextCompare) <> 0 Then Exit Function
    If StrComp(mGeneratedIdentityEditWorksheetName, ws.Name, vbTextCompare) <> 0 Then Exit Function
    If StrComp(mGeneratedIdentityEditAddress, target.Address(False, False), vbTextCompare) <> 0 Then Exit Function

    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function

    For Each colName In Array("Version", "ITEM_CODE", "ROW")
        colIdx = ColumnIndex(loBom, CStr(colName))
        If colIdx > 0 Then
            Set hit = Application.Intersect(target, loBom.DataBodyRange.Columns(colIdx))
            If Not hit Is Nothing Then
                BoxBomGeneratedIdentityColumnWasEdited = True
                Exit Function
            End If
        End If
    Next colName
End Function

Private Sub RebuildBoxBomVersionsForCurrentBoxShipping(ByVal ws As Worksheet)
    On Error GoTo CleanExit

    Dim packageRow As Long
    Dim runtimeMax As Long
    Dim boxName As String

    If ws Is Nothing Then Exit Sub
    boxName = CurrentBoxBuilderName(ws)
    If boxName <> "" Then packageRow = FindShippingBomPackageRowByName(ws.Parent, boxName, runtimeMax)
    RebuildBoxBomVersionListFromDisplayedBom ws, packageRow

CleanExit:
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
    Dim lo3 As ListObject: Set lo3 = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    arrTables = Array(lo1, lo2, lo3)
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
        FillBlankBoxBomVersionShipping loBom
        SortBoxBomByVersionShipping loBom
    End If
    ArrangeBoxBuilderBandShipping loBuilder, loBom
End Sub

Private Sub NormalizeBoxBuilderTable(ByVal loBuilder As ListObject)
    If loBuilder Is Nothing Then Exit Sub
    UnhideListObjectWorksheetColumnsShipping loBuilder
    RepairBoxBuilderBoxNameHeader loBuilder
    EnsureColumnExists loBuilder, "Box Name"
    EnsureColumnExists loBuilder, "UOM"
    EnsureColumnExists loBuilder, "LOCATION"
    EnsureColumnExists loBuilder, "DESCRIPTION"
    RemoveDuplicateBoxBuilderColumn loBuilder, "Box Name"
    RemoveColumnIfExistsShipping loBuilder, "ROW"
    EnsureTableHasRow loBuilder
End Sub

Private Sub RepairBoxBuilderBoxNameHeader(ByVal loBuilder As ListObject)
    Dim boxNameIdx As Long

    If loBuilder Is Nothing Then Exit Sub
    If loBuilder.ListColumns.Count = 0 Then Exit Sub
    boxNameIdx = ColumnIndex(loBuilder, "Box Name")
    If boxNameIdx > 1 Then
        If Not loBuilder.DataBodyRange Is Nothing Then
            If Trim$(NzStr(loBuilder.DataBodyRange.Cells(1, 1).Value)) = "" Then
                loBuilder.DataBodyRange.Cells(1, 1).Value = loBuilder.DataBodyRange.Cells(1, boxNameIdx).Value
            End If
        End If
        loBuilder.HeaderRowRange.Cells(1, 1).Value = "Box Name"
    ElseIf boxNameIdx = 0 Then
        loBuilder.HeaderRowRange.Cells(1, 1).Value = "Box Name"
    End If
End Sub

Private Sub RemoveDuplicateBoxBuilderColumn(ByVal loBuilder As ListObject, ByVal columnName As String)
    Dim i As Long

    If loBuilder Is Nothing Then Exit Sub
    For i = loBuilder.ListColumns.Count To 1 Step -1
        If StrComp(Trim$(NzStr(loBuilder.HeaderRowRange.Cells(1, i).Value)), columnName, vbTextCompare) = 0 Then
            If i > 1 Then loBuilder.ListColumns(i).Delete
        End If
    Next i
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
        EnsureColumnExists loBuilder, COL_BOM_VERSION, "Quantity"
        FormatBoxMakerReadOnlyColumn loBuilder, COL_BOM_VERSION
    Else
        RemoveColumnIfExistsShipping loBuilder, "Quantity"
        RemoveColumnIfExistsShipping loBuilder, COL_BOM_VERSION
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
    FormatBoxMakerReadOnlyColumn lo, COL_CURRENT_INV
End Sub

Private Sub FormatBoxMakerReadOnlyColumn(ByVal lo As ListObject, ByVal colName As String)
    Dim colIdx As Long

    If lo Is Nothing Then Exit Sub
    colIdx = ColumnIndex(lo, colName)
    If colIdx = 0 Then Exit Sub

    EnsureShippingWorksheetEditable lo.Parent
    On Error Resume Next
    lo.ListColumns(colIdx).Range.Locked = True
    lo.ListColumns(colIdx).Range.Interior.Color = RGB(242, 242, 242)
    lo.ListColumns(colIdx).Range.Font.Color = RGB(96, 96, 96)
    On Error GoTo 0
    ProtectShippingWorksheetForReadOnlyColumns lo.Parent
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

Private Function ResolveCanonicalComponentInfoShipping(ByVal itemName As String, _
                                                       ByVal itemCode As String, _
                                                       ByRef rowValue As Long, _
                                                       ByRef resolvedItem As String, _
                                                       ByRef resolvedItemCode As String, _
                                                       ByRef resolvedUom As String, _
                                                       ByRef resolvedLocation As String, _
                                                       ByRef resolvedDescription As String) As Boolean
    On Error GoTo CleanExit

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim workbookPath As String
    Dim wb As Workbook
    Dim lo As ListObject
    Dim openedTransient As Boolean
    Dim cRow As Long
    Dim cCode As Long
    Dim cItem As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long
    Dim cTotalInv As Long
    Dim r As Long
    Dim candidate As String
    Dim candidateCode As String
    Dim candidateUom As String
    Dim candidateLocation As String
    Dim candidateDescription As String
    Dim normalizedItemName As String
    Dim normalizedItemCode As String
    Dim preferredUom As String
    Dim preferredLocation As String
    Dim bestRowValue As Long
    Dim bestScore As Long
    Dim score As Long

    itemName = Trim$(itemName)
    itemCode = Trim$(itemCode)
    If itemName = "" And itemCode = "" Then Exit Function
    normalizedItemName = NormalizeInventoryLookupTextShipping(itemName)
    normalizedItemCode = NormalizeInventoryLookupTextShipping(itemCode)
    preferredUom = NormalizeInventoryLookupTextShipping(resolvedUom)
    preferredLocation = NormalizeInventoryLookupTextShipping(resolvedLocation)

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then Exit Function

    workbookPath = rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb"
    Set wb = FindOpenWorkbookByFullNameShipping(workbookPath)
    If Not wb Is Nothing Then HideWorkbookWindowsShipping wb
    If wb Is Nothing Then Set wb = OpenWorkbookHiddenShipping(workbookPath, True, openedTransient, False)
    If wb Is Nothing Then GoTo CleanExit

    Set lo = FindListObjectByNameShipping(wb, "invSys")
    If lo Is Nothing Then Set lo = FindListObjectByNameShipping(wb, "tblItemSearchIndex")
    If lo Is Nothing Then Set lo = FindListObjectByNameShipping(wb, "tblSkuCatalog")
    If lo Is Nothing Then GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit

    cRow = ColumnIndex(lo, "ROW")
    cCode = ColumnIndex(lo, "ITEM_CODE")
    cItem = ColumnIndex(lo, "ITEM")
    cUom = ColumnIndex(lo, "UOM")
    cLoc = ColumnIndex(lo, "LOCATION")
    cDesc = ColumnIndex(lo, "DESCRIPTION")
    cTotalInv = ColumnIndex(lo, "TOTAL INV")
    If cRow = 0 Or cItem = 0 Then GoTo CleanExit

    For r = 1 To lo.ListRows.Count
        candidate = Trim$(NzStr(lo.DataBodyRange.Cells(r, cItem).Value))
        candidateCode = ""
        candidateUom = ""
        candidateLocation = ""
        candidateDescription = ""
        If cCode > 0 Then candidateCode = Trim$(NzStr(lo.DataBodyRange.Cells(r, cCode).Value))
        If cUom > 0 Then candidateUom = Trim$(NzStr(lo.DataBodyRange.Cells(r, cUom).Value))
        If cLoc > 0 Then candidateLocation = Trim$(NzStr(lo.DataBodyRange.Cells(r, cLoc).Value))
        If cDesc > 0 Then candidateDescription = Trim$(NzStr(lo.DataBodyRange.Cells(r, cDesc).Value))
        If normalizedItemCode <> "" Then
            If NormalizeInventoryLookupTextShipping(candidateCode) <> normalizedItemCode Then GoTo NextRow
        ElseIf NormalizeInventoryLookupTextShipping(candidate) <> normalizedItemName Then
            GoTo NextRow
        End If

        rowValue = NzLng(lo.DataBodyRange.Cells(r, cRow).Value)
        If rowValue <= 0 Then GoTo NextRow

        score = 1
        If normalizedItemCode <> "" Then score = score + 100
        If preferredLocation <> "" And NormalizeInventoryLookupTextShipping(candidateLocation) = preferredLocation Then score = score + 20
        If preferredUom <> "" And NormalizeInventoryLookupTextShipping(candidateUom) = preferredUom Then score = score + 10
        If score > bestScore Or (score = bestScore And rowValue > bestRowValue) Then
            bestScore = score
            bestRowValue = rowValue
            resolvedItem = candidate
            resolvedItemCode = candidateCode
            resolvedUom = candidateUom
            resolvedLocation = candidateLocation
            resolvedDescription = candidateDescription
        End If
NextRow:
    Next r

    If bestRowValue > 0 Then
        rowValue = bestRowValue
        ResolveCanonicalComponentInfoShipping = True
    End If

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wb
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
    If Not wbSnap Is Nothing Then HideWorkbookWindowsShipping wbSnap
    If wbSnap Is Nothing Then
        If Len(Dir$(snapshotPath)) = 0 Then Exit Function
        Set wbSnap = OpenWorkbookHiddenShipping(snapshotPath, True, openedTransient)
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
    If Not wb Is Nothing Then HideWorkbookWindowsShipping wb
    If wb Is Nothing Then Set wb = OpenWorkbookHiddenShipping(workbookPath, True, openedTransient, False)
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
    If Not wb Is Nothing Then HideWorkbookWindowsShipping wb
    If wb Is Nothing Then Set wb = OpenWorkbookHiddenShipping(workbookPath, True, openedTransient, False)
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

Public Function ShippingIsBoxMakerMode(ByVal ws As Worksheet) As Boolean
    ShippingIsBoxMakerMode = IsBoxMakerMode(ws)
End Function

Private Sub InvalidateShippingRibbonLabels()
    On Error Resume Next
    modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
    On Error GoTo 0
End Sub

Private Sub ArrangeBoxBuilderBandShipping(ByVal loBuilder As ListObject, ByVal loBom As ListObject)
    Dim versionsRow As Long
    Dim bomRow As Long
    Dim targetCol As Long
    Dim loVersions As ListObject
    Dim ws As Worksheet

    On Error GoTo CleanExit
    If loBuilder Is Nothing Or loBom Is Nothing Then Exit Sub
    targetCol = loBuilder.Range.Column
    If targetCol < 1 Then Exit Sub

    Set ws = WorksheetFromListObjectShipping(loBuilder)
    If ws Is Nothing Then GoTo CleanExit
    Set loVersions = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    If Not loVersions Is Nothing Then
        MoveListObjectToRowColShipping loBom, TemporaryBoxBomLayoutRowShipping(loBuilder, loBom, loVersions), targetCol
        versionsRow = loBuilder.Range.Row + loBuilder.Range.Rows.Count + SHIP_LAYOUT_GAP_ROWS + 1
        MoveListObjectToRowColShipping loVersions, versionsRow, targetCol
        bomRow = loVersions.Range.Row + loVersions.Range.Rows.Count + SHIP_LAYOUT_GAP_ROWS + 1
    Else
        bomRow = loBuilder.Range.Row + loBuilder.Range.Rows.Count + SHIP_LAYOUT_GAP_ROWS + 1
    End If
    If bomRow < 1 Then Exit Sub
    MoveListObjectToRowColShipping loBom, bomRow, targetCol
CleanExit:
End Sub

Private Function TemporaryBoxBomLayoutRowShipping(ByVal loBuilder As ListObject, _
                                                  ByVal loBom As ListObject, _
                                                  ByVal loVersions As ListObject) As Long
    Dim bottomRow As Long

    bottomRow = 1
    If Not loBuilder Is Nothing Then bottomRow = Application.WorksheetFunction.Max(bottomRow, loBuilder.Range.Row + loBuilder.Range.Rows.Count)
    If Not loBom Is Nothing Then bottomRow = Application.WorksheetFunction.Max(bottomRow, loBom.Range.Row + loBom.Range.Rows.Count)
    If Not loVersions Is Nothing Then bottomRow = Application.WorksheetFunction.Max(bottomRow, loVersions.Range.Row + loVersions.Range.Rows.Count)
    TemporaryBoxBomLayoutRowShipping = bottomRow + SHIP_LAYOUT_GAP_ROWS + 5
End Function

Private Function WorksheetFromListObjectShipping(ByVal lo As ListObject) As Worksheet
    On Error Resume Next
    If lo Is Nothing Then Exit Function
    Set WorksheetFromListObjectShipping = lo.Parent
    On Error GoTo 0
End Function

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

    headers = Array("Version", COL_BOXBOM_ITEM, "ITEM_CODE", "ROW", "QUANTITY", "UOM", "LOCATION", "DESCRIPTION")
    Set startCell = ws.Cells(startRow, startCol)
    For i = LBound(headers) To UBound(headers)
        startCell.Offset(0, i - LBound(headers)).Value = headers(i)
    Next i

    Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
    Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    lo.Name = TABLE_BOX_BOM
    Set CreateBoxBomTable = lo
End Function

Private Function EnsureBoxBomVersionsTable(ByVal ws As Worksheet, ByVal loBom As ListObject) As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim dataRange As Range
    Dim lo As ListObject
    Dim i As Long
    Dim startRow As Long
    Dim startCol As Long

    If ws Is Nothing Then Exit Function
    Set lo = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    If Not lo Is Nothing Then
        headers = BoxBomVersionHeadersShipping()
        For i = LBound(headers) To UBound(headers)
            EnsureColumnExists lo, CStr(headers(i))
        Next i
        ApplyBoxBomVersionStatusValidation lo
        HideBoxBomVersionIdentityColumns lo
        Set EnsureBoxBomVersionsTable = lo
        Exit Function
    End If

    If loBom Is Nothing Then Exit Function
    headers = BoxBomVersionHeadersShipping()
    startRow = loBom.Range.Row + loBom.Range.Rows.Count + SHIP_LAYOUT_GAP_ROWS + 1
    startCol = loBom.Range.Column
    Set startCell = ws.Cells(startRow, startCol)
    For i = LBound(headers) To UBound(headers)
        startCell.Offset(0, i - LBound(headers)).Value = headers(i)
    Next i

    Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
    Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    lo.Name = TABLE_BOX_BOM_VERSIONS
    If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    ApplyBoxBomVersionStatusValidation lo
    HideBoxBomVersionIdentityColumns lo
    Set EnsureBoxBomVersionsTable = lo
End Function

Private Function BoxBomVersionHeadersShipping() As Variant
    BoxBomVersionHeadersShipping = Array("Version", "Status", "Effective From", "Effective To", "Retired At", "Updated At", "Updated By", "Box Name")
End Function

Private Sub HideBoxBomVersionIdentityColumns(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    RemoveListColumnIfExistsShipping lo, "PackageRow"
End Sub

Private Sub ApplyBoxBomVersionStatusValidation(ByVal lo As ListObject)
    On Error GoTo CleanExit

    Dim statusCol As ListColumn

    If lo Is Nothing Then Exit Sub
    Set statusCol = lo.ListColumns("Status")
    If statusCol Is Nothing Then Exit Sub
    If statusCol.DataBodyRange Is Nothing Then Exit Sub

    With statusCol.DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="Active,Retired"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

CleanExit:
End Sub

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

Private Function FirstBlankListRowShipping(ByVal lo As ListObject) As ListRow
    Dim rowIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To lo.ListRows.Count
        If TableRowIsBlankShipping(lo, rowIndex) Then
            Set FirstBlankListRowShipping = lo.ListRows(rowIndex)
            Exit Function
        End If
    Next rowIndex
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
    Dim ws As Worksheet
    Dim quietStarted As Boolean
    If Not targetCell Is Nothing Then
        Set ws = targetCell.Worksheet
    End If
    If ws Is Nothing Then Set ws = WorkbookSheetExistsShipping(ResolveShippingWorkbook(, SHEET_SHIPMENTS), SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Sub
    EnsureBoxBomEntryColumns loBom
    mHandlingShippingSheetChange = True
    modUiQuiet.BeginQuietUi ws.Parent
    quietStarted = True

    Dim invIdx As Long
    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If Not invLo Is Nothing Then
        If itemRow > 0 Then invIdx = FindInvRowIndexByRow(invLo, itemRow)
        If invIdx = 0 And Len(Trim$(itemName)) > 0 Then
            invIdx = FindInvRowIndexByItem(invLo, itemName)
        End If
    End If

    Dim actualRow As Long
    Dim actualItem As String
    Dim actualUom As String, actualLoc As String, actualDesc As String

    If invIdx > 0 Then
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
    End If

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
    FillBlankBoxBomVersionShipping loBom
    SortBoxBomByVersionShipping loBom
    RebuildBoxBomVersionsForCurrentBoxShipping ws
    GoTo CleanExit

ErrHandler:
    MsgBox "ApplyItemToBoxBOM error: " & Err.Description, vbCritical
CleanExit:
    If quietStarted Then
        On Error Resume Next
        modUiQuiet.EndQuietUi
        On Error GoTo 0
    End If
    mHandlingShippingSheetChange = False
End Sub

Public Function LoadShippingBomPackagePickerItems() As Variant
    On Error GoTo FailSoft

    Dim loBom As ListObject
    Dim wb As Workbook

    Set wb = ResolveShippingWorkbook(, SHEET_SHIPMENTS)
    If wb Is Nothing Then Exit Function

    Set loBom = FindListObjectByNameShipping(wb, TABLE_SHIPPING_BOM_VIEW)
    If loBom Is Nothing Then Exit Function
    LoadShippingBomPackagePickerItems = BuildPackagePickerItemsFromShippingBom(loBom)
    Exit Function

FailSoft:
    LoadShippingBomPackagePickerItems = Empty
End Function

Private Function BuildPackagePickerItemsFromShippingBom(ByVal loBom As ListObject) As Variant
    Dim cPackageRow As Long
    Dim cPackageItem As Long
    Dim cPackageUom As Long
    Dim cPackageLocation As Long
    Dim cPackageDescription As Long
    Dim cActive As Long
    Dim dict As Object
    Dim src As Variant
    Dim rows As Collection
    Dim rowData As Variant
    Dim result() As Variant
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
    cActive = ColumnIndex(loBom, "IsActive")
    If cPackageRow = 0 Or cPackageItem = 0 Then Exit Function

    src = loBom.DataBodyRange.Value
    Set dict = CreateObject("Scripting.Dictionary")
    ReDim result(1 To UBound(src, 1), 1 To 7)

    For r = 1 To UBound(src, 1)
        If cActive > 0 Then
            If Not ShippingBomActiveValue(src(r, cActive)) Then GoTo NextPackage
        End If
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

Public Function LoadShippingComponentPickerItems() As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim result As Variant

    mLastComponentPickerStatus = ""
    result = LoadRuntimeInventoryPickerItems()
    If Not IsEmpty(result) Then
        LoadShippingComponentPickerItems = result
        Exit Function
    End If

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If Not ws Is Nothing Then

        Set lo = GetListObject(ws, TABLE_CHECK_INV)
        If ShippingInventoryPickerTableHasRows(lo) Then
            result = BuildShippingInventoryPickerItems(lo)
            If Not IsEmpty(result) Then
                mLastComponentPickerStatus = "Loaded inventory from Check_invSys."
                LoadShippingComponentPickerItems = result
                Exit Function
            End If
        End If

        Set lo = GetListObject(ws, "invSysData_Shipping")
        If ShippingInventoryPickerTableHasRows(lo) Then
            result = BuildShippingInventoryPickerItems(lo)
            If Not IsEmpty(result) Then
                mLastComponentPickerStatus = "Loaded inventory from ShipmentsTally read model."
                LoadShippingComponentPickerItems = result
                Exit Function
            End If
        End If

        Set lo = GetInvSysTableFromWorkbook(ws.Parent)
        If ShippingInventoryPickerTableHasRows(lo) Then
            result = BuildShippingInventoryPickerItems(lo)
            If Not IsEmpty(result) Then
                mLastComponentPickerStatus = "Loaded inventory from this workbook's invSys table."
                LoadShippingComponentPickerItems = result
                Exit Function
            End If
        End If
    End If

    If mLastComponentPickerStatus = "" Then mLastComponentPickerStatus = "No inventory source had rows. Connect/select a warehouse target or refresh inventory."
    Exit Function

FailSoft:
    mLastComponentPickerStatus = "Inventory load failed: " & Err.Description
    Exit Function
End Function

Private Function LoadRuntimeInventoryPickerItems() As Variant
    On Error GoTo FailSoft

    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim warehouseId As String
    Dim rootPath As String
    Dim workbookPath As String
    Dim wbInv As Workbook
    Dim openedTransient As Boolean
    Dim lo As ListObject
    Dim result As Variant

    If Not modNasConnection.ResolveWarehouseTarget(target, statusCode) Then
        mLastComponentPickerStatus = "No warehouse target is selected or remembered."
        Exit Function
    End If
    If target Is Nothing Then
        mLastComponentPickerStatus = "No warehouse target is selected or remembered."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        mLastComponentPickerStatus = "Selected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    workbookPath = rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb"
    If Len(Dir$(workbookPath)) = 0 Then
        mLastComponentPickerStatus = "Inventory runtime workbook was not found: " & workbookPath
        Exit Function
    End If

    Set wbInv = OpenWorkbookHiddenShipping(workbookPath, True, openedTransient, False)
    If wbInv Is Nothing Then
        mLastComponentPickerStatus = "Inventory runtime workbook could not be opened: " & workbookPath
        Exit Function
    End If

    Set lo = GetInvSysTableFromWorkbook(wbInv)
    If ShippingInventoryPickerTableHasRows(lo) Then
        result = BuildShippingInventoryPickerItems(lo)
        If Not IsEmpty(result) Then
            mLastComponentPickerStatus = "Loaded inventory from warehouse runtime invSys table."
            LoadRuntimeInventoryPickerItems = result
            GoTo CleanExit
        End If
    End If

    result = BuildCanonicalRuntimeInventoryPickerItems(wbInv)
    If Not IsEmpty(result) Then
        mLastComponentPickerStatus = "Loaded inventory from warehouse runtime SKU catalog."
        LoadRuntimeInventoryPickerItems = result
        GoTo CleanExit
    End If

    mLastComponentPickerStatus = "Inventory runtime workbook has no usable invSys or SKU catalog rows."

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbInv
    Exit Function

FailSoft:
    mLastComponentPickerStatus = "Runtime inventory lookup failed: " & Err.Description
    LoadRuntimeInventoryPickerItems = Empty
    Resume CleanExit
End Function

Public Function ShippingComponentPickerLastStatus() As String
    ShippingComponentPickerLastStatus = mLastComponentPickerStatus
End Function

Private Function BuildCanonicalRuntimeInventoryPickerItems(ByVal wbInv As Workbook) As Variant
    On Error GoTo FailSoft

    Dim loCatalog As ListObject
    Dim loBalance As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim balances As Object
    Dim sku As String
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim cSku As Long
    Dim cRow As Long
    Dim cCode As Long
    Dim cItem As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long

    If wbInv Is Nothing Then Exit Function
    Set loCatalog = FindListObjectByNameShipping(wbInv, "tblSkuCatalog")
    If loCatalog Is Nothing Then Exit Function
    If loCatalog.DataBodyRange Is Nothing Then Exit Function

    Set loBalance = FindListObjectByNameShipping(wbInv, "tblSkuBalance")
    Set balances = BuildSkuBalanceDictionaryShipping(loBalance)

    cSku = ColumnIndex(loCatalog, "SKU")
    cRow = ColumnIndex(loCatalog, "ROW")
    cCode = ColumnIndex(loCatalog, "ITEM_CODE")
    cItem = ColumnIndex(loCatalog, "ITEM")
    cUom = ColumnIndex(loCatalog, "UOM")
    cLoc = ColumnIndex(loCatalog, "LOCATION")
    cDesc = ColumnIndex(loCatalog, "DESCRIPTION")
    If cItem = 0 Then Exit Function

    src = loCatalog.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 7)
    For r = 1 To UBound(src, 1)
        If Trim$(NzStr(src(r, cItem))) = "" Then GoTo NextRow

        outRow = outRow + 1
        If cRow > 0 Then result(outRow, 1) = NzStr(src(r, cRow))
        If cCode > 0 Then result(outRow, 2) = NzStr(src(r, cCode))
        result(outRow, 3) = NzStr(src(r, cItem))
        If cUom > 0 Then result(outRow, 4) = NzStr(src(r, cUom))
        If cLoc > 0 Then result(outRow, 5) = NzStr(src(r, cLoc))
        If cDesc > 0 Then result(outRow, 6) = NzStr(src(r, cDesc))

        sku = ""
        If cSku > 0 Then sku = NzStr(src(r, cSku))
        If sku = "" And cCode > 0 Then sku = NzStr(src(r, cCode))
        If Not balances Is Nothing Then
            If sku <> "" And balances.Exists(UCase$(sku)) Then result(outRow, 7) = balances(UCase$(sku))
        End If
NextRow:
    Next r

    If outRow = 0 Then Exit Function
    If outRow = UBound(src, 1) Then
        BuildCanonicalRuntimeInventoryPickerItems = result
        Exit Function
    End If

    ReDim trimmed(1 To outRow, 1 To 7)
    For r = 1 To outRow
        For c = 1 To 7
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    BuildCanonicalRuntimeInventoryPickerItems = trimmed
    Exit Function

FailSoft:
    BuildCanonicalRuntimeInventoryPickerItems = Empty
End Function

Private Function BuildSkuBalanceDictionaryShipping(ByVal loBalance As ListObject) As Object
    On Error GoTo FailSoft

    Dim dict As Object
    Dim src As Variant
    Dim r As Long
    Dim cSku As Long
    Dim cQty As Long
    Dim sku As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set BuildSkuBalanceDictionaryShipping = dict
    If loBalance Is Nothing Then Exit Function
    If loBalance.DataBodyRange Is Nothing Then Exit Function

    cSku = ColumnIndex(loBalance, "SKU")
    cQty = ColumnIndex(loBalance, "QtyOnHand")
    If cSku = 0 Or cQty = 0 Then Exit Function

    src = loBalance.DataBodyRange.Value
    For r = 1 To UBound(src, 1)
        sku = UCase$(Trim$(NzStr(src(r, cSku))))
        If sku <> "" Then dict(sku) = NzDbl(src(r, cQty))
    Next r
    Exit Function

FailSoft:
    Set BuildSkuBalanceDictionaryShipping = CreateObject("Scripting.Dictionary")
End Function

Public Function BoxBuilderFormCurrentMeta() As Variant
    Dim result(1 To 4) As Variant
    Dim ws As Worksheet
    Dim loBuilder As ListObject

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        BoxBuilderFormCurrentMeta = result
        Exit Function
    End If

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then
        BoxBuilderFormCurrentMeta = result
        Exit Function
    End If

    EnsureTableHasRow loBuilder
    result(1) = NzStr(ValueFromTable(loBuilder, "Box Name"))
    result(2) = NzStr(ValueFromTable(loBuilder, "UOM"))
    result(3) = NzStr(ValueFromTable(loBuilder, "LOCATION"))
    result(4) = NzStr(ValueFromTable(loBuilder, "DESCRIPTION"))
    BoxBuilderFormCurrentMeta = result
End Function

Public Function BoxBuilderFormCurrentComponents() As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim loBom As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim cVersion As Long
    Dim cItem As Long
    Dim cCode As Long
    Dim cRow As Long
    Dim cQty As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function

    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Function
    EnsureBoxBomEntryColumns loBom
    If loBom.DataBodyRange Is Nothing Then Exit Function

    cVersion = ColumnIndex(loBom, "Version")
    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    cCode = ColumnIndex(loBom, "ITEM_CODE")
    cRow = ColumnIndex(loBom, "ROW")
    cQty = ColumnIndex(loBom, "QUANTITY")
    cUom = ColumnIndex(loBom, "UOM")
    cLoc = ColumnIndex(loBom, "LOCATION")
    cDesc = ColumnIndex(loBom, "DESCRIPTION")
    If cItem = 0 Or cRow = 0 Or cQty = 0 Then Exit Function

    src = loBom.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 8)
    For r = 1 To UBound(src, 1)
        If Trim$(NzStr(src(r, cItem))) <> "" Or NzLng(src(r, cRow)) > 0 Or NzDbl(src(r, cQty)) > 0 Then
            outRow = outRow + 1
            If cVersion > 0 Then result(outRow, 1) = NzStr(src(r, cVersion)) Else result(outRow, 1) = "v1"
            result(outRow, 2) = NzStr(src(r, cItem))
            If cCode > 0 Then result(outRow, 3) = NzStr(src(r, cCode))
            result(outRow, 4) = NzLng(src(r, cRow))
            result(outRow, 5) = NzDbl(src(r, cQty))
            If cUom > 0 Then result(outRow, 6) = NzStr(src(r, cUom))
            If cLoc > 0 Then result(outRow, 7) = NzStr(src(r, cLoc))
            If cDesc > 0 Then result(outRow, 8) = NzStr(src(r, cDesc))
        End If
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 8)
    For r = 1 To outRow
        For c = 1 To 8
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    BoxBuilderFormCurrentComponents = trimmed
    Exit Function

FailSoft:
    BoxBuilderFormCurrentComponents = Empty
End Function

Public Function BoxBuilderFormLoadSavedBoxes(Optional ByVal includeActive As Boolean = True, _
                                             Optional ByVal includeArchived As Boolean = False, _
                                             Optional ByVal skipRuntimeRefreshForTest As Boolean = False) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim loView As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim rowData As Variant
    Dim keys As Variant
    Dim dict As Object
    Dim packageRow As Long
    Dim r As Long
    Dim c As Long
    Dim cPackageRow As Long
    Dim cPackageItem As Long
    Dim cPackageUom As Long
    Dim cPackageLocation As Long
    Dim cPackageDescription As Long
    Dim cActive As Long
    Dim refreshReport As String
    Dim activePackages As Object
    Dim outputCount As Long
    Dim key As Variant

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function

    If Not skipRuntimeRefreshForTest Then RefreshShippingBomViewForWorkbook ws.Parent, refreshReport
    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If loView Is Nothing Then Exit Function
    If loView.DataBodyRange Is Nothing Then Exit Function

    cPackageRow = ColumnIndex(loView, "PackageRow")
    cPackageItem = ColumnIndex(loView, "PackageItem")
    cPackageUom = ColumnIndex(loView, "PackageUOM")
    cPackageLocation = ColumnIndex(loView, "PackageLocation")
    cPackageDescription = ColumnIndex(loView, "PackageDescription")
    cActive = ColumnIndex(loView, "IsActive")
    If cPackageRow = 0 Or cPackageItem = 0 Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    Set activePackages = CreateObject("Scripting.Dictionary")
    src = loView.DataBodyRange.Value
    For r = 1 To UBound(src, 1)
        packageRow = NzLng(src(r, cPackageRow))
        If packageRow <= 0 Then GoTo NextRow
        If cActive = 0 Or ShippingBomActiveValue(src(r, cActive)) Then activePackages(CStr(packageRow)) = True
        If dict.Exists(CStr(packageRow)) Then GoTo NextRow

        ReDim rowData(1 To 5)
        rowData(1) = packageRow
        rowData(2) = NzStr(src(r, cPackageItem))
        If cPackageUom > 0 Then rowData(3) = NzStr(src(r, cPackageUom))
        If cPackageLocation > 0 Then rowData(4) = NzStr(src(r, cPackageLocation))
        If cPackageDescription > 0 Then rowData(5) = NzStr(src(r, cPackageDescription))
        dict.Add CStr(packageRow), rowData
NextRow:
    Next r

    If dict.Count = 0 Then Exit Function
    keys = dict.Keys
    For Each key In keys
        If ShouldIncludeBoxBuilderPackage(CStr(key), activePackages, includeActive, includeArchived) Then outputCount = outputCount + 1
    Next key
    If outputCount = 0 Then Exit Function

    ReDim result(1 To outputCount, 1 To 5)
    r = 0
    For Each key In keys
        If Not ShouldIncludeBoxBuilderPackage(CStr(key), activePackages, includeActive, includeArchived) Then GoTo NextOutput
        r = r + 1
        rowData = dict(key)
        For c = 1 To 5
            result(r, c) = rowData(c)
        Next c
NextOutput:
    Next key
    BoxBuilderFormLoadSavedBoxes = result
    Exit Function

FailSoft:
    BoxBuilderFormLoadSavedBoxes = Empty
End Function

Private Function ShouldIncludeBoxBuilderPackage(ByVal packageKey As String, _
                                                ByVal activePackages As Object, _
                                                ByVal includeActive As Boolean, _
                                                ByVal includeArchived As Boolean) As Boolean
    Dim isActive As Boolean

    If activePackages Is Nothing Then Exit Function
    isActive = activePackages.Exists(packageKey)
    If isActive Then
        ShouldIncludeBoxBuilderPackage = includeActive
    Else
        ShouldIncludeBoxBuilderPackage = includeArchived
    End If
End Function

Public Function BoxBuilderFormLoadSavedBoxesReportForTest(Optional ByVal includeActive As Boolean = True, _
                                                          Optional ByVal includeArchived As Boolean = False) As String
    On Error GoTo FailSoft

    Dim rowsData As Variant
    Dim r As Long

    rowsData = BoxBuilderFormLoadSavedBoxes(includeActive, includeArchived, True)
    If IsEmpty(rowsData) Then
        BoxBuilderFormLoadSavedBoxesReportForTest = "COUNT=0"
        Exit Function
    End If

    BoxBuilderFormLoadSavedBoxesReportForTest = "COUNT=" & CStr(UBound(rowsData, 1))
    For r = 1 To UBound(rowsData, 1)
        BoxBuilderFormLoadSavedBoxesReportForTest = BoxBuilderFormLoadSavedBoxesReportForTest & _
            ";ROW=" & CStr(rowsData(r, 1)) & "|BOX=" & CStr(rowsData(r, 2))
    Next r
    Exit Function

FailSoft:
    BoxBuilderFormLoadSavedBoxesReportForTest = "ERROR=" & Err.Description
End Function

Public Function BoxBuilderFormInitializeSmokeForTest(ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim frm As frmShippingBoxBuilder

    Set frm = New frmShippingBoxBuilder
    report = "OK"
    BoxBuilderFormInitializeSmokeForTest = True
    Set frm = Nothing
    Exit Function

FailSoft:
    report = Err.Description
End Function

Public Function BoxBuilderFormLoadVersions(ByVal packageRow As Long) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim loView As ListObject
    Dim versionRows As Variant
    Dim versionCount As Long

    If packageRow <= 0 Then Exit Function
    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If loView Is Nothing Then Exit Function
    If loView.DataBodyRange Is Nothing Then Exit Function

    versionRows = BuildBoxBomVersionRows(loView, packageRow, versionCount)
    If versionCount > 0 Then BoxBuilderFormLoadVersions = versionRows
    Exit Function

FailSoft:
    BoxBuilderFormLoadVersions = Empty
End Function

Public Function BoxBuilderFormLoadVersionComponents(ByVal packageRow As Long, ByVal versionLabel As String) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim loView As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim versionNumber As Long
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim cLabel As Long
    Dim cComponentRow As Long
    Dim cComponentItemCode As Long
    Dim cComponentItem As Long
    Dim cComponentQty As Long
    Dim cComponentUom As Long
    Dim cComponentLocation As Long
    Dim cComponentDescription As Long
    Dim rowLabel As String

    If packageRow <= 0 Then Exit Function
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then versionLabel = "v1"
    versionNumber = BomVersionNumberFromLabel(versionLabel)

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If loView Is Nothing Then Exit Function
    If loView.DataBodyRange Is Nothing Then Exit Function

    cPackageRow = ColumnIndex(loView, "PackageRow")
    cVersion = ColumnIndex(loView, "BomVersion")
    cLabel = ColumnIndex(loView, "BomVersionLabel")
    cComponentRow = ColumnIndex(loView, "ComponentRow")
    cComponentItemCode = ColumnIndex(loView, "ComponentItemCode")
    cComponentItem = ColumnIndex(loView, "ComponentItem")
    cComponentQty = ColumnIndex(loView, "ComponentQty")
    cComponentUom = ColumnIndex(loView, "ComponentUOM")
    cComponentLocation = ColumnIndex(loView, "ComponentLocation")
    cComponentDescription = ColumnIndex(loView, "ComponentDescription")
    If cPackageRow = 0 Or cComponentRow = 0 Or cComponentQty = 0 Then Exit Function

    src = loView.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 8)
    For r = 1 To UBound(src, 1)
        If NzLng(src(r, cPackageRow)) <> packageRow Then GoTo NextRow
        If cVersion > 0 And NzLng(src(r, cVersion)) <> versionNumber Then GoTo NextRow
        If cLabel > 0 Then
            rowLabel = NormalizeBoxBomVersionLabelShipping(NzStr(src(r, cLabel)))
            If rowLabel <> "" And StrComp(rowLabel, versionLabel, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If Not ShippingBomSourceRowHasComponent(src, r, cComponentRow, cComponentItem, cComponentQty) Then GoTo NextRow

        outRow = outRow + 1
        result(outRow, 1) = versionLabel
        If cComponentItem > 0 Then result(outRow, 2) = NzStr(src(r, cComponentItem))
        If cComponentItemCode > 0 Then result(outRow, 3) = NzStr(src(r, cComponentItemCode))
        result(outRow, 4) = NzLng(src(r, cComponentRow))
        result(outRow, 5) = NzDbl(src(r, cComponentQty))
        If cComponentUom > 0 Then result(outRow, 6) = NzStr(src(r, cComponentUom))
        If cComponentLocation > 0 Then result(outRow, 7) = NzStr(src(r, cComponentLocation))
        If cComponentDescription > 0 Then result(outRow, 8) = NzStr(src(r, cComponentDescription))
NextRow:
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 8)
    For r = 1 To outRow
        For c = 1 To 8
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    BoxBuilderFormLoadVersionComponents = trimmed
    Exit Function

FailSoft:
    BoxBuilderFormLoadVersionComponents = Empty
End Function

Public Function BoxMakerFormLoadSavedBoxes() As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim loSource As ListObject
    Dim wbRuntime As Workbook
    Dim openedTransient As Boolean
    Dim report As String

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function

    Set loSource = BoxMakerShippingBomSourceTable(ws, wbRuntime, openedTransient, report)
    If loSource Is Nothing Then GoTo CleanExit
    BoxMakerFormLoadSavedBoxes = BuildPackagePickerItemsFromShippingBom(loSource)

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    Exit Function

FailSoft:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    BoxMakerFormLoadSavedBoxes = Empty
End Function

Public Function BoxMakerFormLoadShippableInventory(ByVal savedBoxes As Variant) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim snapshotCache As Object
    Dim result() As Variant
    Dim r As Long
    Dim rowCount As Long
    Dim rowValue As Long
    Dim itemName As String
    Dim foundCurrent As Boolean
    Dim currentInv As Variant
    Dim stagedDeltas As Object
    Dim stagedKey As String
    Dim stagedDelta As Double

    If IsEmpty(savedBoxes) Then Exit Function
    rowCount = UBound(savedBoxes, 1)
    If rowCount <= 0 Then Exit Function

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    Set invLo = GetInvSysTable()
    Set stagedDeltas = modRoleEventWriter.GetLocalStagedBoxInventoryDeltas()

    ReDim result(1 To rowCount, 1 To 5)
    For r = 1 To rowCount
        rowValue = NzLng(savedBoxes(r, 1))
        itemName = NzStr(savedBoxes(r, 2))
        foundCurrent = False
        currentInv = ResolveCurrentInventoryValue(ws, invLo, rowValue, itemName, foundCurrent, snapshotCache)
        If rowValue > 0 And Not stagedDeltas Is Nothing Then
            stagedKey = CStr(rowValue)
            If stagedDeltas.Exists(stagedKey) Then
                stagedDelta = CDbl(stagedDeltas(stagedKey))
                If foundCurrent And IsNumeric(currentInv) Then
                    currentInv = CDbl(currentInv) + stagedDelta
                ElseIf Abs(stagedDelta) > 0.0000001 Then
                    currentInv = stagedDelta
                    foundCurrent = True
                End If
                If IsNumeric(currentInv) Then
                    If CDbl(currentInv) < 0 Then currentInv = 0
                End If
            End If
        End If

        result(r, 1) = rowValue
        result(r, 2) = itemName
        If foundCurrent Then result(r, 3) = currentInv Else result(r, 3) = ""
        result(r, 4) = NzStr(savedBoxes(r, 4))
        result(r, 5) = NzStr(savedBoxes(r, 5))
    Next r
    BoxMakerFormLoadShippableInventory = result
    Exit Function

FailSoft:
    BoxMakerFormLoadShippableInventory = Empty
End Function

Public Function BoxMakerFormLoadShippableVersionInventory(ByVal savedBoxes As Variant) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim loSource As ListObject
    Dim wbRuntime As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim rows As Collection
    Dim rowData As Variant
    Dim result() As Variant
    Dim versions As Variant
    Dim versionInv As Object
    Dim invLo As ListObject
    Dim snapshotCache As Object
    Dim boxRow As Long
    Dim r As Long
    Dim v As Long
    Dim c As Long
    Dim activeCount As Long
    Dim versionLabel As String
    Dim foundCurrent As Boolean
    Dim currentInv As Variant

    If IsEmpty(savedBoxes) Then Exit Function
    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function

    Set loSource = BoxMakerShippingBomSourceTable(ws, wbRuntime, openedTransient, report)
    If loSource Is Nothing Then GoTo CleanExit
    Set invLo = GetInvSysTable()

    Set rows = New Collection
    For r = 1 To UBound(savedBoxes, 1)
        boxRow = NzLng(savedBoxes(r, 1))
        If boxRow <= 0 Then GoTo NextBox

        versions = BuildBoxBomVersionRows(loSource, boxRow, c)
        If IsEmpty(versions) Then GoTo NextBox
        Set versionInv = BoxMakerFormLoadBoxVersionInventory(boxRow, NzStr(savedBoxes(r, 2)))
        activeCount = CountActiveBoxBomVersionsShipping(versions)

        For v = 1 To UBound(versions, 1)
            If UCase$(NzStr(versions(v, 2))) <> "ACTIVE" Then GoTo NextVersion
            versionLabel = NormalizeBoxBomVersionLabelShipping(versions(v, 1))
            If versionLabel = "" Then GoTo NextVersion

            ReDim rowData(1 To 8)
            rowData(1) = boxRow
            rowData(2) = NzStr(savedBoxes(r, 2))
            rowData(3) = versionLabel
            If activeCount = 1 Then
                foundCurrent = False
                currentInv = ResolveCurrentInventoryValue(ws, invLo, boxRow, NzStr(savedBoxes(r, 2)), foundCurrent, snapshotCache)
                If foundCurrent Then rowData(4) = currentInv
            End If
            If Not versionInv Is Nothing Then
                If versionInv.Exists(versionLabel) And Trim$(NzStr(rowData(4))) = "" Then rowData(4) = versionInv(versionLabel)
            End If
            rowData(8) = PendingBoxVersionInventoryOverlayValue(boxRow, versionLabel, rowData(4))
            rowData(5) = NzStr(savedBoxes(r, 4))
            rowData(6) = NzStr(savedBoxes(r, 5))
            rowData(7) = NzStr(savedBoxes(r, 6))
            rows.Add rowData
NextVersion:
        Next v
NextBox:
    Next r

    If rows.Count = 0 Then GoTo CleanExit
    ReDim result(1 To rows.Count, 1 To 8)
    For r = 1 To rows.Count
        rowData = rows(r)
        For c = 1 To 8
            result(r, c) = rowData(c)
        Next c
    Next r
    BoxMakerFormLoadShippableVersionInventory = result

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    Exit Function

FailSoft:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    BoxMakerFormLoadShippableVersionInventory = Empty
End Function

Private Function CountActiveBoxBomVersionsShipping(ByVal versions As Variant) As Long
    Dim r As Long

    If IsEmpty(versions) Then Exit Function
    For r = 1 To UBound(versions, 1)
        If UCase$(NzStr(versions(r, 2))) = "ACTIVE" Then
            If NormalizeBoxBomVersionLabelShipping(versions(r, 1)) <> "" Then CountActiveBoxBomVersionsShipping = CountActiveBoxBomVersionsShipping + 1
        End If
    Next r
End Function

Public Sub RegisterPendingBoxVersionInventoryOverlay(ByVal packageRow As Long, _
                                                     ByVal versionLabel As String, _
                                                     ByVal projectedQty As Double, _
                                                     Optional ByVal baselineQty As Variant)
    Dim key As String
    Dim resolvedBaseline As Double

    EnsurePendingBoxVersionInventoryOverlayLoaded
    key = PendingBoxVersionInventoryKey(packageRow, versionLabel)
    If key = "" Then Exit Sub
    If mPendingBoxVersionInventoryOverlay Is Nothing Then
        Set mPendingBoxVersionInventoryOverlay = CreateObject("Scripting.Dictionary")
        mPendingBoxVersionInventoryOverlay.CompareMode = vbTextCompare
    End If
    If mPendingBoxVersionInventoryOverlayBaseline Is Nothing Then
        Set mPendingBoxVersionInventoryOverlayBaseline = CreateObject("Scripting.Dictionary")
        mPendingBoxVersionInventoryOverlayBaseline.CompareMode = vbTextCompare
    End If
    If projectedQty < 0 Then projectedQty = 0
    If IsMissing(baselineQty) Or Not IsNumeric(baselineQty) Then
        resolvedBaseline = projectedQty
    Else
        resolvedBaseline = CDbl(baselineQty)
        If resolvedBaseline < projectedQty Then resolvedBaseline = projectedQty
    End If
    mPendingBoxVersionInventoryOverlay(key) = projectedQty
    mPendingBoxVersionInventoryOverlayBaseline(key) = resolvedBaseline
    PersistPendingBoxVersionInventoryOverlay
End Sub

Public Sub ClearPendingBoxVersionInventoryOverlayForTest()
    Set mPendingBoxVersionInventoryOverlay = Nothing
    Set mPendingBoxVersionInventoryOverlayBaseline = Nothing
    mPendingBoxVersionInventoryOverlayPath = ""
End Sub

Public Function PendingBoxVersionInventoryOverlayPathForTest() As String
    PendingBoxVersionInventoryOverlayPathForTest = PersistentPendingBoxVersionInventoryOverlayPath()
End Function

Public Function PendingBoxVersionInventoryOverlayText(ByVal packageRow As Long, _
                                                      ByVal versionLabel As String, _
                                                      ByVal backendText As String) As String
    Dim overlayValue As Variant

    overlayValue = PendingBoxVersionInventoryOverlayValue(packageRow, versionLabel, backendText)
    PendingBoxVersionInventoryOverlayText = Trim$(NzStr(overlayValue))
End Function

Private Function PendingBoxVersionInventoryOverlayValue(ByVal packageRow As Long, _
                                                       ByVal versionLabel As String, _
                                                       ByVal backendValue As Variant) As Variant
    Dim key As String
    Dim pendingQty As Double
    Dim backendQty As Double
    Dim baselineQty As Double

    EnsurePendingBoxVersionInventoryOverlayLoaded
    PendingBoxVersionInventoryOverlayValue = backendValue
    If mPendingBoxVersionInventoryOverlay Is Nothing Then Exit Function

    key = PendingBoxVersionInventoryKey(packageRow, versionLabel)
    If key = "" Then Exit Function
    If Not mPendingBoxVersionInventoryOverlay.Exists(key) Then Exit Function

    pendingQty = CDbl(mPendingBoxVersionInventoryOverlay(key))
    If IsNumeric(backendValue) Then
        backendQty = CDbl(backendValue)
        baselineQty = pendingQty
        If Not mPendingBoxVersionInventoryOverlayBaseline Is Nothing Then
            If mPendingBoxVersionInventoryOverlayBaseline.Exists(key) Then baselineQty = CDbl(mPendingBoxVersionInventoryOverlayBaseline(key))
        End If
        If backendQty > baselineQty + 0.0000001 Then
            mPendingBoxVersionInventoryOverlay.Remove key
            If Not mPendingBoxVersionInventoryOverlayBaseline Is Nothing Then
                If mPendingBoxVersionInventoryOverlayBaseline.Exists(key) Then mPendingBoxVersionInventoryOverlayBaseline.Remove key
            End If
            PersistPendingBoxVersionInventoryOverlay
            Exit Function
        End If
    End If
    PendingBoxVersionInventoryOverlayValue = pendingQty
End Function

Private Sub EnsurePendingBoxVersionInventoryOverlayLoaded()
    On Error GoTo CleanExit

    Dim filePath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim parts As Variant
    Dim key As String
    Dim qtyText As String
    Dim baselineText As String

    filePath = PersistentPendingBoxVersionInventoryOverlayPath()
    If mPendingBoxVersionInventoryOverlayPath = filePath And Not mPendingBoxVersionInventoryOverlay Is Nothing Then Exit Sub

    Set mPendingBoxVersionInventoryOverlay = CreateObject("Scripting.Dictionary")
    mPendingBoxVersionInventoryOverlay.CompareMode = vbTextCompare
    Set mPendingBoxVersionInventoryOverlayBaseline = CreateObject("Scripting.Dictionary")
    mPendingBoxVersionInventoryOverlayBaseline.CompareMode = vbTextCompare
    mPendingBoxVersionInventoryOverlayPath = filePath
    If filePath = "" Then Exit Sub
    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Sub

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        parts = Split(lineText, vbTab)
        If UBound(parts) >= 1 Then
            key = Trim$(UnescapeHoldField(CStr(parts(0))))
            qtyText = Trim$(UnescapeHoldField(CStr(parts(1))))
            baselineText = qtyText
            If UBound(parts) >= 2 Then baselineText = Trim$(UnescapeHoldField(CStr(parts(2))))
            If key <> "" And IsNumeric(Replace$(qtyText, ",", "")) Then
                mPendingBoxVersionInventoryOverlay(key) = CDbl(Replace$(qtyText, ",", ""))
                If IsNumeric(Replace$(baselineText, ",", "")) Then
                    mPendingBoxVersionInventoryOverlayBaseline(key) = CDbl(Replace$(baselineText, ",", ""))
                Else
                    mPendingBoxVersionInventoryOverlayBaseline(key) = CDbl(Replace$(qtyText, ",", ""))
                End If
            End If
        End If
    Loop
    Close #fileNum

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
End Sub

Private Sub PersistPendingBoxVersionInventoryOverlay()
    On Error GoTo CleanExit

    Dim filePath As String
    Dim fileNum As Integer
    Dim key As Variant

    filePath = PersistentPendingBoxVersionInventoryOverlayPath()
    mPendingBoxVersionInventoryOverlayPath = filePath
    If filePath = "" Then Exit Sub
    EnsureLocalFolderExistsShipping ParentFolderPathShipping(filePath)

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    If Not mPendingBoxVersionInventoryOverlay Is Nothing Then
        For Each key In mPendingBoxVersionInventoryOverlay.Keys
            Print #fileNum, EscapeHoldField(CStr(key)) & vbTab & _
                            EscapeHoldField(CStr(mPendingBoxVersionInventoryOverlay(key))) & vbTab & _
                            EscapeHoldField(CStr(PendingOverlayBaselineForKey(CStr(key))))
        Next key
    End If
    Close #fileNum

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
End Sub

Private Function PendingOverlayBaselineForKey(ByVal key As String) As Double
    If Not mPendingBoxVersionInventoryOverlayBaseline Is Nothing Then
        If mPendingBoxVersionInventoryOverlayBaseline.Exists(key) Then
            PendingOverlayBaselineForKey = CDbl(mPendingBoxVersionInventoryOverlayBaseline(key))
            Exit Function
        End If
    End If
    If Not mPendingBoxVersionInventoryOverlay Is Nothing Then
        If mPendingBoxVersionInventoryOverlay.Exists(key) Then PendingOverlayBaselineForKey = CDbl(mPendingBoxVersionInventoryOverlay(key))
    End If
End Function

Private Function PendingBoxVersionInventoryKey(ByVal packageRow As Long, ByVal versionLabel As String) As String
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If packageRow <= 0 Or versionLabel = "" Then Exit Function
    PendingBoxVersionInventoryKey = CStr(packageRow) & "|" & versionLabel
End Function

Public Function BoxMakerFormLoadBoxVersionInventory(ByVal packageRow As Long, ByVal boxName As String) As Object
    On Error GoTo FailSoft

    Dim result As Object
    Dim invLo As ListObject
    Dim invIdx As Long
    Dim packageSku As String
    Dim packageItem As String
    Dim skuCandidates As Object
    Dim wb As Workbook
    Dim loLog As ListObject
    Dim src As Variant
    Dim r As Long
    Dim cEventType As Long
    Dim cSku As Long
    Dim cQtyDelta As Long
    Dim cNote As Long
    Dim eventType As String
    Dim skuValue As String
    Dim versionLabel As String
    Dim qtyDelta As Double
    Dim stagedDeltas As Object
    Dim key As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set BoxMakerFormLoadBoxVersionInventory = result
    If packageRow <= 0 And Trim$(boxName) = "" Then Exit Function

    Set invLo = GetInvSysTable()
    If Not invLo Is Nothing Then
        If packageRow > 0 Then invIdx = FindInvRowIndexByRow(invLo, packageRow)
        If invIdx <= 0 And Trim$(boxName) <> "" Then invIdx = FindInvRowIndexByItem(invLo, boxName)
        If invIdx > 0 Then packageSku = NzStr(GetInvSysValueByIndex(invLo, invIdx, "ITEM_CODE"))
        If invIdx > 0 Then packageItem = NzStr(GetInvSysValueByIndex(invLo, invIdx, "ITEM"))
        If packageSku = "" Then packageSku = packageItem
        Set wb = invLo.Parent.Parent
    End If
    If packageSku = "" Then packageSku = Trim$(boxName)

    Set skuCandidates = CreateObject("Scripting.Dictionary")
    skuCandidates.CompareMode = vbTextCompare
    If packageSku <> "" Then skuCandidates(packageSku) = True
    If packageItem <> "" Then skuCandidates(packageItem) = True
    If Trim$(boxName) <> "" Then skuCandidates(Trim$(boxName)) = True

    If Not wb Is Nothing And skuCandidates.Count > 0 Then
        Set loLog = FindListObjectByNameShipping(wb, "tblInventoryLog")
        If loLog Is Nothing Then Set loLog = FindListObjectByNameShipping(wb, "InventoryLog")
        If Not loLog Is Nothing Then
            If Not loLog.DataBodyRange Is Nothing Then
                cEventType = ColumnIndex(loLog, "EventType")
                cSku = ColumnIndex(loLog, "SKU")
                cQtyDelta = ColumnIndex(loLog, "QtyDelta")
                cNote = ColumnIndex(loLog, "Note")
                If cEventType > 0 And cSku > 0 And cQtyDelta > 0 And cNote > 0 Then
                    src = loLog.DataBodyRange.Value
                    For r = 1 To UBound(src, 1)
                        eventType = UCase$(Trim$(NzStr(src(r, cEventType))))
                        If eventType <> EVENT_TYPE_SHIP _
                           And eventType <> EVENT_TYPE_SHIP_RESERVE _
                           And eventType <> EVENT_TYPE_SHIP_RELEASE _
                           And eventType <> EVENT_TYPE_BOX_BUILD _
                           And eventType <> EVENT_TYPE_BOX_UNBOX Then GoTo NextLogRow

                        skuValue = Trim$(NzStr(src(r, cSku)))
                        If Not skuCandidates.Exists(skuValue) Then GoTo NextLogRow

                        versionLabel = ExtractBoxVersionLabelFromNoteShipping(NzStr(src(r, cNote)))
                        If versionLabel = "" Then GoTo NextLogRow
                        qtyDelta = NzDbl(src(r, cQtyDelta))
                        AddVersionInventoryDeltaShipping result, versionLabel, qtyDelta
NextLogRow:
                    Next r
                End If
            End If
        End If
    End If

    If packageRow > 0 Then
        Set stagedDeltas = modRoleEventWriter.GetLocalStagedBoxVersionInventoryDeltas(packageRow)
        If Not stagedDeltas Is Nothing Then
            For Each key In stagedDeltas.Keys
                AddVersionInventoryDeltaShipping result, CStr(key), CDbl(stagedDeltas(key))
            Next key
        End If
    End If

    For Each key In result.Keys
        If Abs(CDbl(result(key))) < 0.0000001 Then result(key) = 0#
        If CDbl(result(key)) < 0 Then result(key) = 0#
    Next key
    Exit Function

FailSoft:
    If BoxMakerFormLoadBoxVersionInventory Is Nothing Then
        Set BoxMakerFormLoadBoxVersionInventory = CreateObject("Scripting.Dictionary")
        BoxMakerFormLoadBoxVersionInventory.CompareMode = vbTextCompare
    End If
End Function

Private Sub AddVersionInventoryDeltaShipping(ByVal totals As Object, _
                                             ByVal versionLabel As String, _
                                             ByVal qtyDelta As Double)
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If totals Is Nothing Then Exit Sub
    If versionLabel = "" Then Exit Sub
    If Abs(qtyDelta) < 0.0000001 Then Exit Sub

    If totals.Exists(versionLabel) Then
        totals(versionLabel) = CDbl(totals(versionLabel)) + qtyDelta
    Else
        totals(versionLabel) = qtyDelta
    End If
End Sub

Private Function ExtractBoxVersionLabelFromNoteShipping(ByVal noteText As String) As String
    Dim pos As Long
    Dim tailText As String
    Dim i As Long
    Dim ch As String
    Dim token As String

    noteText = Trim$(noteText)
    If noteText = "" Then Exit Function
    pos = InStr(1, noteText, "VERSION=", vbTextCompare)
    If pos = 0 Then Exit Function

    tailText = Mid$(noteText, pos + Len("VERSION="))
    For i = 1 To Len(tailText)
        ch = Mid$(tailText, i, 1)
        If ch = ";" Or ch = "|" Or ch = "," Or ch = vbTab Or ch = " " Or ch = vbCr Or ch = vbLf Then Exit For
        token = token & ch
    Next i
    ExtractBoxVersionLabelFromNoteShipping = NormalizeBoxBomVersionLabelShipping(token)
End Function

Public Function BoxMakerFormLoadVersions(ByVal packageRow As Long) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim loSource As ListObject
    Dim wbRuntime As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim versionCount As Long

    If packageRow <= 0 Then Exit Function
    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function

    Set loSource = BoxMakerShippingBomSourceTable(ws, wbRuntime, openedTransient, report)
    If loSource Is Nothing Then GoTo CleanExit
    BoxMakerFormLoadVersions = BuildBoxBomVersionRows(loSource, packageRow, versionCount)

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    Exit Function

FailSoft:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    BoxMakerFormLoadVersions = Empty
End Function

Public Function BoxMakerFormLoadVersionComponents(ByVal packageRow As Long, ByVal versionLabel As String) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim loSource As ListObject
    Dim wbRuntime As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim invLo As ListObject
    Dim snapshotCache As Object
    Dim components As Variant
    Dim result() As Variant
    Dim r As Long
    Dim c As Long
    Dim foundCurrent As Boolean
    Dim currentInv As Variant
    Dim rowValue As Long
    Dim itemName As String

    If packageRow <= 0 Then Exit Function
    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function

    Set loSource = BoxMakerShippingBomSourceTable(ws, wbRuntime, openedTransient, report)
    If loSource Is Nothing Then GoTo CleanExit
    components = BoxMakerVersionComponentsFromTable(loSource, packageRow, versionLabel)
    If IsEmpty(components) Then GoTo CleanExit

    Set invLo = GetInvSysTable()

    ReDim result(1 To UBound(components, 1), 1 To 9)
    For r = 1 To UBound(components, 1)
        For c = 1 To 8
            result(r, c) = components(r, c)
        Next c
        rowValue = NzLng(components(r, 4))
        itemName = NzStr(components(r, 2))
        foundCurrent = False
        currentInv = ResolveCurrentInventoryValue(ws, invLo, rowValue, itemName, foundCurrent, snapshotCache)
        If foundCurrent Then result(r, 9) = currentInv Else result(r, 9) = ""
    Next r
    BoxMakerFormLoadVersionComponents = result

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    Exit Function

FailSoft:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    BoxMakerFormLoadVersionComponents = Empty
End Function

Private Function BoxMakerShippingBomSourceTable(ByVal ws As Worksheet, _
                                                ByRef wbRuntime As Workbook, _
                                                ByRef openedTransient As Boolean, _
                                                ByRef report As String) As ListObject
    On Error GoTo FailSoft

    Dim loView As ListObject
    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String

    If ws Is Nothing Then Exit Function

    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If Not loView Is Nothing Then
        If Not loView.DataBodyRange Is Nothing Then
            Set BoxMakerShippingBomSourceTable = loView
            Exit Function
        End If
    End If

    RefreshShippingBomViewForWorkbook ws.Parent, report
    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If Not loView Is Nothing Then
        If Not loView.DataBodyRange Is Nothing Then
            Set BoxMakerShippingBomSourceTable = loView
            Exit Function
        End If
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "A connected warehouse target is required to load BoxMaker designs."
        Exit Function
    End If
    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        report = "Selected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    Set wbRuntime = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbRuntime Is Nothing Then Exit Function
    Set BoxMakerShippingBomSourceTable = EnsureShippingBomSchema(wbRuntime, report)
    Exit Function

FailSoft:
    report = "BoxMakerShippingBomSourceTable failed: " & Err.Description
End Function

Private Function BoxMakerVersionComponentsFromTable(ByVal loSource As ListObject, _
                                                    ByVal packageRow As Long, _
                                                    ByVal versionLabel As String) As Variant
    On Error GoTo FailSoft

    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim versionNumber As Long
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim cLabel As Long
    Dim cComponentRow As Long
    Dim cComponentItemCode As Long
    Dim cComponentItem As Long
    Dim cComponentQty As Long
    Dim cComponentUom As Long
    Dim cComponentLocation As Long
    Dim cComponentDescription As Long
    Dim rowLabel As String

    If loSource Is Nothing Then Exit Function
    If loSource.DataBodyRange Is Nothing Then Exit Function
    If packageRow <= 0 Then Exit Function

    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then versionLabel = "v1"
    versionNumber = BomVersionNumberFromLabel(versionLabel)

    cPackageRow = ColumnIndex(loSource, "PackageRow")
    cVersion = ColumnIndex(loSource, "BomVersion")
    cLabel = ColumnIndex(loSource, "BomVersionLabel")
    cComponentRow = ColumnIndex(loSource, "ComponentRow")
    cComponentItemCode = ColumnIndex(loSource, "ComponentItemCode")
    cComponentItem = ColumnIndex(loSource, "ComponentItem")
    cComponentQty = ColumnIndex(loSource, "ComponentQty")
    cComponentUom = ColumnIndex(loSource, "ComponentUOM")
    cComponentLocation = ColumnIndex(loSource, "ComponentLocation")
    cComponentDescription = ColumnIndex(loSource, "ComponentDescription")
    If cPackageRow = 0 Or cComponentRow = 0 Or cComponentQty = 0 Then Exit Function

    src = loSource.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 8)
    For r = 1 To UBound(src, 1)
        If NzLng(src(r, cPackageRow)) <> packageRow Then GoTo NextRow
        If cVersion > 0 And NzLng(src(r, cVersion)) <> versionNumber Then GoTo NextRow
        If cLabel > 0 Then
            rowLabel = NormalizeBoxBomVersionLabelShipping(NzStr(src(r, cLabel)))
            If rowLabel <> "" And StrComp(rowLabel, versionLabel, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If Not ShippingBomSourceRowHasComponent(src, r, cComponentRow, cComponentItem, cComponentQty) Then GoTo NextRow

        outRow = outRow + 1
        result(outRow, 1) = versionLabel
        If cComponentItem > 0 Then result(outRow, 2) = NzStr(src(r, cComponentItem))
        If cComponentItemCode > 0 Then result(outRow, 3) = NzStr(src(r, cComponentItemCode))
        result(outRow, 4) = NzLng(src(r, cComponentRow))
        result(outRow, 5) = NzDbl(src(r, cComponentQty))
        If cComponentUom > 0 Then result(outRow, 6) = NzStr(src(r, cComponentUom))
        If cComponentLocation > 0 Then result(outRow, 7) = NzStr(src(r, cComponentLocation))
        If cComponentDescription > 0 Then result(outRow, 8) = NzStr(src(r, cComponentDescription))
NextRow:
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 8)
    For r = 1 To outRow
        For c = 1 To 8
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    BoxMakerVersionComponentsFromTable = trimmed
    Exit Function

FailSoft:
    BoxMakerVersionComponentsFromTable = Empty
End Function

Public Function BoxMakerFormCurrentInventory(ByVal rowValue As Long, ByVal itemName As String) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim snapshotCache As Object
    Dim foundCurrent As Boolean

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    Set invLo = GetInvSysTable()
    BoxMakerFormCurrentInventory = ResolveCurrentInventoryValue(ws, invLo, rowValue, itemName, foundCurrent, snapshotCache)
    If Not foundCurrent Then BoxMakerFormCurrentInventory = ""
    Exit Function

FailSoft:
    BoxMakerFormCurrentInventory = ""
End Function

Private Function QueueBoxMakerFormPayload(ByVal isMakeAction As Boolean, _
                                          ByVal packageRow As Long, _
                                          ByVal boxName As String, _
                                          ByVal boxUom As String, _
                                          ByVal boxLocation As String, _
                                          ByVal boxDescription As String, _
                                          ByVal versionLabel As String, _
                                          ByVal boxQty As Double, _
                                          ByVal componentRows As Variant, _
                                          ByRef componentTotalOut As Double, _
                                          ByRef packageTotalOut As Double, _
                                          ByRef eventIdOut As String, _
                                          ByRef errNotes As String) As Boolean
    On Error GoTo FailSoft

    Dim payloadItems As Collection
    Dim payloadJson As String
    Dim eventType As String
    Dim sourceName As String
    Dim componentIoType As String
    Dim packageIoType As String
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String
    Dim itemCode As String
    Dim qtyPerBox As Double
    Dim qtyTotal As Double
    Dim uomVal As String
    Dim locationVal As String
    Dim descriptionVal As String

    errNotes = ""
    eventIdOut = ""
    componentTotalOut = 0
    packageTotalOut = 0
    boxName = Trim$(boxName)
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then versionLabel = "v1"
    If packageRow <= 0 Then
        errNotes = "Saved box ROW was not resolved."
        Exit Function
    End If
    If boxName = "" Then
        errNotes = "Box name is required."
        Exit Function
    End If
    If boxQty <= 0 Then
        errNotes = "Box quantity must be greater than zero."
        Exit Function
    End If
    If IsEmpty(componentRows) Then
        errNotes = "Selected box version has no component rows."
        Exit Function
    End If

    If isMakeAction Then
        eventType = EVENT_TYPE_BOX_BUILD
        sourceName = "FORM_BOX_MAKER"
        componentIoType = "USED"
        packageIoType = "MADE"
    Else
        eventType = EVENT_TYPE_BOX_UNBOX
        sourceName = "FORM_BOX_UNMAKE"
        componentIoType = "RETURNED"
        packageIoType = "UNMADE"
    End If

    Set payloadItems = New Collection
    For r = 1 To UBound(componentRows, 1)
        itemName = BoxBuilderFormBomText(componentRows, r, 2, "")
        itemCode = BoxBuilderFormBomText(componentRows, r, 3, "")
        rowVal = BoxBuilderFormBomLong(componentRows, r, 4)
        qtyPerBox = BoxBuilderFormBomDouble(componentRows, r, 5)
        If BoxMakerComponentRowIsBlank(itemName, rowVal, qtyPerBox) Then GoTo NextComponent
        If qtyPerBox <= 0 Then
            errNotes = "Component row " & CStr(r) & " needs a quantity greater than zero."
            Exit Function
        End If
        qtyTotal = qtyPerBox * boxQty
        If qtyTotal <= 0 Then
            errNotes = "Component row " & CStr(r) & " produced a zero total quantity."
            Exit Function
        End If
        uomVal = BoxBuilderFormBomText(componentRows, r, 6, "")
        locationVal = BoxBuilderFormBomText(componentRows, r, 7, "")
        descriptionVal = BoxBuilderFormBomText(componentRows, r, 8, "")
        If itemCode = "" And itemName <> "" Then itemCode = itemName

        AddBoxBuildPayloadItem payloadItems, rowVal, itemCode, itemName, qtyTotal, uomVal, locationVal, descriptionVal, componentIoType, versionLabel
        componentTotalOut = componentTotalOut + qtyTotal
NextComponent:
    Next r

    If componentTotalOut <= 0 Then
        errNotes = "No component quantities were found for the selected box version."
        Exit Function
    End If

    AddBoxBuildPayloadItem payloadItems, _
                           packageRow, _
                           boxName, _
                           boxName, _
                           boxQty, _
                           boxUom, _
                           boxLocation, _
                           boxDescription, _
                           packageIoType, _
                           versionLabel
    packageTotalOut = boxQty

    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        errNotes = "No BoxMaker payload rows were generated."
        Exit Function
    End If

    QueueBoxMakerFormPayload = modRoleEventWriter.QueuePayloadEventCurrent( _
        eventType, _
        "", _
        payloadJson, _
        sourceName, _
        eventIdOut, _
        errNotes)
    Exit Function

FailSoft:
    errNotes = "QueueBoxMakerFormPayload failed: " & Err.Description
End Function

Public Function CommitBoxMakerFormAction(ByVal packageRow As Long, _
                                         ByVal boxName As String, _
                                         ByVal boxUom As String, _
                                         ByVal boxLocation As String, _
                                         ByVal boxDescription As String, _
                                         ByVal versionLabel As String, _
                                         ByVal boxQty As Double, _
                                         ByVal componentRows As Variant, _
                                         ByRef resultMessage As String, _
                                         Optional ByVal actionText As String = "MAKE", _
                                         Optional ByRef syncCompletedOut As Boolean = False, _
                                         Optional ByVal displayedAvailableQty As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim rowCount As Long
    Dim usedTotal As Double
    Dim madeTotal As Double
    Dim packageReturned As Double
    Dim componentsReturned As Double
    Dim errNotes As String
    Dim eventIdOut As String
    Dim runtimeReport As String
    Dim batchProcessed As Boolean
    Dim currentQty As Double
    Dim foundCurrentQty As Boolean

    resultMessage = ""
    syncCompletedOut = False
    actionText = UCase$(Trim$(actionText))
    boxName = Trim$(boxName)
    boxUom = Trim$(boxUom)
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then versionLabel = "v1"

    If boxName = "" Then
        resultMessage = "Select a saved box before posting BoxMaker inventory."
        Exit Function
    End If
    If boxQty <= 0 Then
        resultMessage = "Box quantity must be greater than zero."
        Exit Function
    End If
    If IsEmpty(componentRows) Then
        resultMessage = "Selected box version has no component rows."
        Exit Function
    End If
    rowCount = UBound(componentRows, 1)
    If rowCount <= 0 Then
        resultMessage = "Selected box version has no component rows."
        Exit Function
    End If

    If actionText = "UNMAKE" Or actionText = "UNBOX" Then
        If Not IsMissing(displayedAvailableQty) Then
            If IsNumeric(Replace$(Trim$(NzStr(displayedAvailableQty)), ",", "")) Then
                currentQty = CDbl(Replace$(Trim$(NzStr(displayedAvailableQty)), ",", ""))
                foundCurrentQty = True
            End If
        End If
        If Not foundCurrentQty Then currentQty = ResolveBoxMakerUnboxAvailableQty(packageRow, boxName, foundCurrentQty)
        If Not foundCurrentQty Then
            resultMessage = "Not allowed: current inventory was not resolved for " & boxName & " " & versionLabel & "."
            Exit Function
        End If
        If boxQty > currentQty + 0.0000001 Then
            resultMessage = "Not allowed: Qty exceeds inventory. " & boxName & " " & versionLabel & _
                            " has " & FormatBoxMakerQuantityText(currentQty) & _
                            " in inventory, but Qty is " & FormatBoxMakerQuantityText(boxQty) & "."
            Exit Function
        End If
        If Not QueueBoxMakerFormPayload(False, _
                                        packageRow, _
                                        boxName, _
                                        boxUom, _
                                        boxLocation, _
                                        boxDescription, _
                                        versionLabel, _
                                        boxQty, _
                                        componentRows, _
                                        componentsReturned, _
                                        packageReturned, _
                                        eventIdOut, _
                                        errNotes) Then
            If errNotes = "" Then errNotes = "Box could not be unboxed."
            resultMessage = errNotes
            Exit Function
        End If
        resultMessage = "Box unbox event queued for " & FormatBoxMakerQuantityText(boxQty) & " " & boxName & _
                        " " & versionLabel & _
                        ". Removes " & FormatBoxMakerQuantityText(packageReturned) & _
                        " shippable units and returns " & FormatBoxMakerQuantityText(componentsReturned) & _
                        " component units after processor sync."
    Else
        If Not QueueBoxMakerFormPayload(True, _
                                        packageRow, _
                                        boxName, _
                                        boxUom, _
                                        boxLocation, _
                                        boxDescription, _
                                        versionLabel, _
                                        boxQty, _
                                        componentRows, _
                                        usedTotal, _
                                        madeTotal, _
                                        eventIdOut, _
                                        errNotes) Then
            If errNotes = "" Then errNotes = "Box creation could not be posted."
            resultMessage = errNotes
            Exit Function
        End If
        resultMessage = "Box build event queued for " & FormatBoxMakerQuantityText(boxQty) & " " & boxName & _
                        " " & versionLabel & _
                        ". Uses " & FormatBoxMakerQuantityText(usedTotal) & _
                        " component units and adds " & FormatBoxMakerQuantityText(madeTotal) & _
                        " shippable units after processor sync."
    End If
    batchProcessed = RunShippingRuntimeQueueRefresh(ActiveWorkbook, ResolveCurrentShippingWarehouseId(), runtimeReport)
    If Not batchProcessed Then batchProcessed = BoxMakerRuntimeReportShowsProcessed(runtimeReport)
    syncCompletedOut = batchProcessed
    If batchProcessed Then
        AppendNote errNotes, "Sync complete."
    Else
        AppendNote errNotes, "Sync pending."
    End If
    If runtimeReport <> "" Then AppendNote errNotes, runtimeReport
    If eventIdOut <> "" Then AppendNote errNotes, "Inbox EventID: " & eventIdOut
    If errNotes <> "" Then resultMessage = resultMessage & vbCrLf & vbCrLf & errNotes
    ShowShippingStatus resultMessage
    CommitBoxMakerFormAction = True
    Exit Function

ErrHandler:
    resultMessage = "BOX_MAKER_FORM_COMMIT failed: " & Err.Description
End Function

Private Function ResolveBoxMakerUnboxAvailableQty(ByVal packageRow As Long, _
                                                  ByVal boxName As String, _
                                                  ByRef foundCurrentQty As Boolean) As Double
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim invIdx As Long
    Dim totalInv As Variant

    foundCurrentQty = False
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function

    Set invLo = GetInvSysTableFromWorkbook(wb)
    If invLo Is Nothing Then GoTo TryShippingReadModel
    If packageRow > 0 Then invIdx = FindInvRowIndexByRow(invLo, packageRow)
    If invIdx <= 0 And Trim$(boxName) <> "" Then invIdx = FindInvRowIndexByItem(invLo, boxName)
    If invIdx <= 0 Then GoTo TryShippingReadModel

    totalInv = GetInvSysValueByIndex(invLo, invIdx, "TOTAL INV")
    If Not IsNumeric(totalInv) Then GoTo TryShippingReadModel
    foundCurrentQty = True
    ResolveBoxMakerUnboxAvailableQty = CDbl(totalInv)
    Exit Function

TryShippingReadModel:
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function

    totalInv = ResolveCurrentInventoryFromTable(GetListObject(ws, "invSysData_Shipping"), packageRow, boxName, foundCurrentQty)
    If foundCurrentQty And IsNumeric(totalInv) Then
        ResolveBoxMakerUnboxAvailableQty = CDbl(totalInv)
        Exit Function
    End If

    totalInv = ResolveCurrentInventoryFromTable(GetListObject(ws, TABLE_CHECK_INV), packageRow, boxName, foundCurrentQty)
    If foundCurrentQty And IsNumeric(totalInv) Then ResolveBoxMakerUnboxAvailableQty = CDbl(totalInv)
End Function

Public Function CommitBoxMakerFormActionReportForTest(ByVal packageRow As Long, _
                                                      ByVal boxName As String, _
                                                      ByVal boxUom As String, _
                                                      ByVal boxLocation As String, _
                                                      ByVal boxDescription As String, _
                                                      ByVal versionLabel As String, _
                                                      ByVal boxQty As Double, _
                                                      ByVal componentRows As Variant, _
                                                      ByVal actionText As String, _
                                                      Optional ByVal displayedAvailableQty As Variant) As String
    Dim resultMessage As String
    Dim posted As Boolean
    Dim syncCompleted As Boolean

    posted = CommitBoxMakerFormAction(packageRow, _
                                      boxName, _
                                      boxUom, _
                                      boxLocation, _
                                      boxDescription, _
                                      versionLabel, _
                                      boxQty, _
                                      componentRows, _
                                      resultMessage, _
                                      actionText, _
                                      syncCompleted, _
                                      displayedAvailableQty)
    CommitBoxMakerFormActionReportForTest = "Posted=" & IIf(posted, "1", "0") & "; Report=" & resultMessage
End Function

Public Function ShipmentsFormLoadShippables() As Variant
    Dim savedBoxes As Variant

    savedBoxes = BoxMakerFormLoadSavedBoxes()
    If IsEmpty(savedBoxes) Then Exit Function
    ShipmentsFormLoadShippables = BoxMakerFormLoadShippableVersionInventory(savedBoxes)
End Function

Public Function ShipmentsProjectedDisplayQty(ByVal nasQty As Double, _
                                             ByVal lockedQty As Double, _
                                             ByVal unreservedLocalQty As Double, _
                                             Optional ByVal reservedLocalQty As Double = 0, _
                                             Optional ByVal pendingOverlayQty As Variant) As Double
    Dim baseQty As Double
    Dim projectedQty As Double
    Dim lockedToSubtract As Double
    Dim overlayQty As Double
    Dim overlayApplied As Boolean

    baseQty = nasQty
    lockedToSubtract = lockedQty
    If Not IsMissing(pendingOverlayQty) Then
        If IsNumeric(pendingOverlayQty) Then
            overlayQty = CDbl(pendingOverlayQty)
            If overlayQty < nasQty Then
                baseQty = overlayQty
                overlayApplied = True
            End If
        End If
    End If

    If overlayApplied And reservedLocalQty > 0 Then
        lockedToSubtract = lockedQty - reservedLocalQty
        If lockedToSubtract < 0 Then lockedToSubtract = 0
    End If

    projectedQty = baseQty - lockedToSubtract - unreservedLocalQty
    If projectedQty < 0 Then projectedQty = 0
    ShipmentsProjectedDisplayQty = projectedQty
End Function

Public Function ShipmentsProjectedDisplayQtyForTest(ByVal nasQty As Double, _
                                                    ByVal lockedQty As Double, _
                                                    ByVal unreservedLocalQty As Double, _
                                                    Optional ByVal reservedLocalQty As Double = 0, _
                                                    Optional ByVal pendingOverlayQty As Variant) As Double
    If IsMissing(pendingOverlayQty) Then
        ShipmentsProjectedDisplayQtyForTest = ShipmentsProjectedDisplayQty(nasQty, lockedQty, unreservedLocalQty, reservedLocalQty)
    Else
        ShipmentsProjectedDisplayQtyForTest = ShipmentsProjectedDisplayQty(nasQty, lockedQty, unreservedLocalQty, reservedLocalQty, pendingOverlayQty)
    End If
End Function

Public Function ShipmentsFormRefreshRuntimeInventory(ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim wb As Workbook
    Dim runtimeReport As String
    Dim bomReport As String
    Dim warehouseId As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        report = "No active operator workbook to refresh."
        Exit Function
    End If

    warehouseId = ResolveCurrentShippingWarehouseId()
    If Not RunShippingRuntimeQueueRefresh(wb, warehouseId, runtimeReport, False) Then
        report = runtimeReport
        Exit Function
    End If
    If Not RefreshShippingBomViewForWorkbook(wb, bomReport) Then
        report = bomReport
        Exit Function
    End If

    report = runtimeReport
    If Trim$(bomReport) <> "" Then AppendNote report, bomReport
    ShipmentsFormRefreshRuntimeInventory = True
    Exit Function

FailSoft:
    report = "Shipments refresh failed: " & Err.Description
End Function

Private Function RunShippingRuntimeQueueRefresh(ByVal wb As Workbook, _
                                                ByVal warehouseId As String, _
                                                ByRef report As String, _
                                                Optional ByVal requireQueuedWork As Boolean = True) As Boolean
    On Error GoTo FailSoft

    Dim resolvedWarehouseId As String
    Dim stationId As String
    Dim stagingReport As String
    Dim batchReport As String
    Dim publishReport As String
    Dim processedCount As Long
    Dim totalTimer As Single
    Dim batchTimer As Single
    Dim batchMs As Long

    If wb Is Nothing Then
        report = "Operator workbook not resolved."
        Exit Function
    End If

    resolvedWarehouseId = Trim$(warehouseId)
    If resolvedWarehouseId = "" Then resolvedWarehouseId = ResolveCurrentShippingWarehouseId()
    stationId = ResolveCurrentShippingStationId(resolvedWarehouseId)
    totalTimer = Timer
    batchTimer = Timer

    If Not modRoleEventWriter.SyncLocalStagedInboxRows(stagingReport, resolvedWarehouseId, stationId) Then
        batchMs = ElapsedMillisecondsShipping(batchTimer)
        report = "Local staged shipping rows could not be merged before runtime processing. " & _
                 "StagingReport=" & stagingReport & " BatchReport=Skipped; " & _
                 FormatShippingRuntimeTiming(ElapsedMillisecondsShipping(totalTimer), batchMs, 0)
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(resolvedWarehouseId, 0, batchReport)
    batchMs = ElapsedMillisecondsShipping(batchTimer)
    If Left$(batchReport, 15) = "RunBatch failed" Then
        report = "RunBatch failed after local shipping post/write. StagingReport=" & stagingReport & " " & _
                 batchReport & " RefreshReport=Skipped; " & _
                 FormatShippingRuntimeTiming(ElapsedMillisecondsShipping(totalTimer), batchMs, 0)
        Exit Function
    End If

    If Not ShippingRuntimeReportShowsProcessed(processedCount, batchReport) _
       And (requireQueuedWork Or ShippingRuntimeReportMetric(stagingReport, "LocalStagingMerged") > 0) Then
        report = "RunBatch did not handle the queued shipping event after local post/write. " & _
                 "StagingReport=" & stagingReport & " BatchReport=" & batchReport & " RefreshReport=Skipped; " & _
                 FormatShippingRuntimeTiming(ElapsedMillisecondsShipping(totalTimer), batchMs, 0)
        Exit Function
    End If

    If processedCount > 0 Then
        If Not modInventoryDomainBridge.PublishInventorySnapshotBridge(resolvedWarehouseId, Nothing, publishReport) Then
            If publishReport = "" Then publishReport = "Snapshot publish failed."
        End If
    End If

    report = "Processed=" & CStr(processedCount) & "; StagingReport=" & stagingReport & "; BatchReport=" & batchReport
    If publishReport <> "" Then report = report & "; PublishWarning=" & publishReport
    report = report & "; " & FormatShippingRuntimeTiming(ElapsedMillisecondsShipping(totalTimer), batchMs, 0)
    RunShippingRuntimeQueueRefresh = True
    Exit Function

FailSoft:
    report = "RunShippingRuntimeQueueRefresh failed: " & Err.Description
End Function

Private Function ResolveCurrentShippingStationId(ByVal warehouseId As String) As String
    On Error Resume Next

    Dim target As WarehouseTarget

    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then
        If Trim$(warehouseId) = "" Or StrComp(Trim$(target.WarehouseId), Trim$(warehouseId), vbTextCompare) = 0 Then
            ResolveCurrentShippingStationId = Trim$(target.StationId)
        End If
    End If
    If ResolveCurrentShippingStationId = "" Then ResolveCurrentShippingStationId = Trim$(modConfig.GetStationId())
End Function

Private Function ShippingRuntimeReportShowsProcessed(ByVal processedCount As Long, ByVal batchReport As String) As Boolean
    If processedCount > 0 Then
        ShippingRuntimeReportShowsProcessed = True
        Exit Function
    End If

    If ShippingRuntimeReportMetric(batchReport, "Applied") > 0 Then
        ShippingRuntimeReportShowsProcessed = True
        Exit Function
    End If

    If ShippingRuntimeReportMetric(batchReport, "SkipDup") > 0 Then
        ShippingRuntimeReportShowsProcessed = True
    End If
End Function

Private Function ShippingRuntimeReportMetric(ByVal runtimeReport As String, ByVal metricName As String) As Long
    Dim marker As String
    Dim pos As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim ch As String

    marker = metricName & "="
    pos = InStr(1, runtimeReport, marker, vbTextCompare)
    If pos <= 0 Then Exit Function

    valueStart = pos + Len(marker)
    valueEnd = valueStart
    Do While valueEnd <= Len(runtimeReport)
        ch = Mid$(runtimeReport, valueEnd, 1)
        If ch < "0" Or ch > "9" Then Exit Do
        valueEnd = valueEnd + 1
    Loop
    If valueEnd <= valueStart Then Exit Function
    ShippingRuntimeReportMetric = CLng(Mid$(runtimeReport, valueStart, valueEnd - valueStart))
End Function

Private Function FormatShippingRuntimeTiming(ByVal totalMs As Long, _
                                             ByVal batchMs As Long, _
                                             ByVal refreshMs As Long) As String
    FormatShippingRuntimeTiming = "TimingMs=Total:" & CStr(totalMs) & _
                                  ";Batch:" & CStr(batchMs) & _
                                  ";Refresh:" & CStr(refreshMs)
End Function

Private Function ElapsedMillisecondsShipping(ByVal startedAt As Single) As Long
    Dim deltaSeconds As Single

    deltaSeconds = Timer - startedAt
    If deltaSeconds < 0 Then deltaSeconds = deltaSeconds + 86400!
    ElapsedMillisecondsShipping = CLng(deltaSeconds * 1000)
End Function

Private Function ResolveCurrentShippingWarehouseId() As String
    On Error Resume Next

    Dim target As WarehouseTarget

    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then ResolveCurrentShippingWarehouseId = Trim$(target.WarehouseId)
    If ResolveCurrentShippingWarehouseId = "" Then ResolveCurrentShippingWarehouseId = Trim$(modConfig.GetWarehouseId())
    If ResolveCurrentShippingWarehouseId = "" Then ResolveCurrentShippingWarehouseId = Trim$(modConfig.GetString("WarehouseId", ""))
End Function

Public Function ShipmentsFormRecentHistoryText(Optional ByVal limitCount As Long = 20) As String
    On Error GoTo FailSoft

    Dim warehouseId As String
    Dim stationId As String
    Dim inventoryWb As Workbook
    Dim logText As String
    Dim pipelineText As String
    Dim target As WarehouseTarget

    If limitCount <= 0 Then limitCount = 20
    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then
        warehouseId = Trim$(target.WarehouseId)
        stationId = Trim$(target.StationId)
    End If
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetWarehouseId())
    If stationId = "" Then stationId = Trim$(modConfig.GetStationId())
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetString("WarehouseId", ""))
    If stationId = "" Then stationId = Trim$(modConfig.GetString("StationId", ""))

    Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(warehouseId)
    logText = RecentShipmentInventoryLogTextShipping(inventoryWb, limitCount)
    pipelineText = ShipmentPipelineStatusTextShipping(warehouseId, stationId)

    ShipmentsFormRecentHistoryText = logText
    If pipelineText <> "" Then ShipmentsFormRecentHistoryText = ShipmentsFormRecentHistoryText & vbCrLf & vbCrLf & pipelineText
    Exit Function

FailSoft:
    ShipmentsFormRecentHistoryText = "Shipments history failed: " & Err.Description
End Function

Private Function RecentShipmentInventoryLogTextShipping(ByVal inventoryWb As Workbook, ByVal limitCount As Long) As String
    On Error GoTo FailSoft

    Dim loLog As ListObject
    Dim cEventId As Long
    Dim cEventType As Long
    Dim cTime As Long
    Dim cSku As Long
    Dim cQtyDelta As Long
    Dim cLocation As Long
    Dim cNote As Long
    Dim rowIndex As Long
    Dim shown As Long
    Dim eventType As String
    Dim lineText As String

    If inventoryWb Is Nothing Then
        RecentShipmentInventoryLogTextShipping = "Processed server log: inventory workbook not open."
        Exit Function
    End If
    Set loLog = FindListObjectByNameShipping(inventoryWb, "tblInventoryLog")
    If loLog Is Nothing Or loLog.DataBodyRange Is Nothing Then
        RecentShipmentInventoryLogTextShipping = "Processed server log: no tblInventoryLog rows found."
        Exit Function
    End If

    cEventId = ColumnIndex(loLog, "EventID")
    cEventType = ColumnIndex(loLog, "EventType")
    cTime = ColumnIndex(loLog, "OccurredAtUTC")
    cSku = ColumnIndex(loLog, "SKU")
    cQtyDelta = ColumnIndex(loLog, "QtyDelta")
    cLocation = ColumnIndex(loLog, "Location")
    cNote = ColumnIndex(loLog, "Note")
    If cEventType = 0 Then
        RecentShipmentInventoryLogTextShipping = "Processed server log: EventType column missing."
        Exit Function
    End If

    RecentShipmentInventoryLogTextShipping = "Processed server shipment history:"
    For rowIndex = loLog.ListRows.Count To 1 Step -1
        eventType = UCase$(Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cEventType).Value)))
        If eventType = EVENT_TYPE_SHIP Or eventType = EVENT_TYPE_SHIP_RESERVE _
           Or eventType = EVENT_TYPE_SHIP_RELEASE Or eventType = EVENT_TYPE_ADMIN_SHIPMENT_RECONCILE Then
            shown = shown + 1
            lineText = CStr(shown) & ". "
            If cEventId > 0 Then lineText = lineText & NzStr(loLog.DataBodyRange.Cells(rowIndex, cEventId).Value) & " | "
            lineText = lineText & eventType
            If cTime > 0 Then lineText = lineText & " | " & FormatHistoryValueShipping(loLog.DataBodyRange.Cells(rowIndex, cTime).Value)
            If cSku > 0 Then lineText = lineText & " | " & NzStr(loLog.DataBodyRange.Cells(rowIndex, cSku).Value)
            If cQtyDelta > 0 Then lineText = lineText & " | Delta " & FormatBoxMakerQuantityText(NzDbl(loLog.DataBodyRange.Cells(rowIndex, cQtyDelta).Value))
            If cLocation > 0 And Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cLocation).Value)) <> "" Then lineText = lineText & " | " & NzStr(loLog.DataBodyRange.Cells(rowIndex, cLocation).Value)
            If cNote > 0 And Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cNote).Value)) <> "" Then lineText = lineText & " | " & Left$(Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cNote).Value)), 90)
            RecentShipmentInventoryLogTextShipping = RecentShipmentInventoryLogTextShipping & vbCrLf & lineText
            If shown >= limitCount Then Exit For
        End If
    Next rowIndex
    If shown = 0 Then RecentShipmentInventoryLogTextShipping = RecentShipmentInventoryLogTextShipping & vbCrLf & "None found."
    Exit Function

FailSoft:
    RecentShipmentInventoryLogTextShipping = "Processed server log failed: " & Err.Description
End Function

Private Function ShipmentPipelineStatusTextShipping(ByVal warehouseId As String, ByVal stationId As String) As String
    Dim pendingCount As Long
    Dim matchingPending As Long
    Dim stagedRows As Long
    Dim matchingStaged As Long
    Dim inboxReport As String
    Dim inboxError As String
    Dim stagedReport As String
    Dim stagedError As String

    inboxReport = modRoleEventWriter.DescribeInboxPendingRows(EVENT_TYPE_SHIP, warehouseId, stationId, "", pendingCount, matchingPending, inboxError)
    stagedReport = modRoleEventWriter.DescribeLocalStagedInboxRows(EVENT_TYPE_SHIP & "," & EVENT_TYPE_SHIP_RESERVE & "," & EVENT_TYPE_SHIP_RELEASE, _
                                                                    warehouseId, _
                                                                    stationId, _
                                                                    stagedRows, _
                                                                    matchingStaged, _
                                                                    stagedError)

    ShipmentPipelineStatusTextShipping = "Shipment pipeline status:" & vbCrLf & _
        "NAS shipping inbox pending rows: " & CStr(pendingCount)
    If pendingCount > 0 Then
        ShipmentPipelineStatusTextShipping = ShipmentPipelineStatusTextShipping & vbCrLf & _
            "Processor has not applied these rows yet; tblInventoryLog will update after processor catch-up."
    End If
    If inboxReport <> "" Then
        ShipmentPipelineStatusTextShipping = ShipmentPipelineStatusTextShipping & vbCrLf & inboxReport
    ElseIf inboxError <> "" Then
        ShipmentPipelineStatusTextShipping = ShipmentPipelineStatusTextShipping & vbCrLf & inboxError
    End If
    ShipmentPipelineStatusTextShipping = ShipmentPipelineStatusTextShipping & vbCrLf & _
        "Local staged shipment rows: " & CStr(matchingStaged)
    If stagedReport <> "" Then
        ShipmentPipelineStatusTextShipping = ShipmentPipelineStatusTextShipping & vbCrLf & stagedReport
    ElseIf stagedError <> "" Then
        ShipmentPipelineStatusTextShipping = ShipmentPipelineStatusTextShipping & vbCrLf & stagedError
    End If
End Function

Private Function FormatHistoryValueShipping(ByVal valueIn As Variant) As String
    If IsDate(valueIn) Then
        FormatHistoryValueShipping = Format$(CDate(valueIn), "yyyy-mm-dd hh:nn:ss")
    Else
        FormatHistoryValueShipping = NzStr(valueIn)
    End If
End Function

Public Function ShipmentsFormLoadLines(Optional ByVal holdRows As Boolean = False) As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tableName As String
    Dim cRef As Long
    Dim cItem As Long
    Dim cQty As Long
    Dim cRow As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long
    Dim cArea As Long
    Dim cCarrier As Long
    Dim cReserve As Long
    Dim src As Variant
    Dim rows As Variant
    Dim r As Long
    Dim outRow As Long
    Dim countRows As Long

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    If holdRows Then tableName = TABLE_NOTSHIPPED Else tableName = TABLE_SHIPMENTS
    Set lo = GetListObject(ws, tableName)
    If holdRows Then
        LoadPersistentHoldRowsLocal lo
    Else
        LoadPersistentActiveShipmentRowsLocal lo
    End If
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    cRef = ColumnIndex(lo, "REF_NUMBER")
    cItem = ColumnIndex(lo, "ITEMS")
    cQty = ColumnIndex(lo, "QUANTITY")
    cRow = ColumnIndex(lo, "ROW")
    cUom = ColumnIndex(lo, "UOM")
    cLoc = ColumnIndex(lo, "LOCATION")
    cDesc = ColumnIndex(lo, "DESCRIPTION")
    cArea = ColumnIndex(lo, "AREA")
    cCarrier = ColumnIndex(lo, "CARRIER")
    cReserve = ColumnIndex(lo, COL_SHIPMENT_RESERVE_EVENT_ID)
    If cItem = 0 Or cQty = 0 Then Exit Function

    src = lo.DataBodyRange.Value
    For r = 1 To UBound(src, 1)
        If Trim$(NzStr(src(r, cItem))) <> "" Or NzDbl(src(r, cQty)) <> 0 Then countRows = countRows + 1
    Next r
    If countRows = 0 Then Exit Function

    ReDim rows(1 To countRows, 1 To 11)
    outRow = 0
    For r = 1 To UBound(src, 1)
        If Trim$(NzStr(src(r, cItem))) = "" And NzDbl(src(r, cQty)) = 0 Then GoTo NextRow
        outRow = outRow + 1
        If cRef > 0 Then rows(outRow, 1) = NzStr(src(r, cRef))
        rows(outRow, 2) = NzStr(src(r, cItem))
        rows(outRow, 3) = NzDbl(src(r, cQty))
        If cUom > 0 Then rows(outRow, 4) = NzStr(src(r, cUom))
        If cLoc > 0 Then rows(outRow, 5) = NzStr(src(r, cLoc))
        If cRow > 0 Then rows(outRow, 6) = NzLng(src(r, cRow))
        If cDesc > 0 Then rows(outRow, 7) = NzStr(src(r, cDesc))
        rows(outRow, 8) = r
        If cArea > 0 Then rows(outRow, 9) = NormalizeShipmentArea(NzStr(src(r, cArea)), holdRows)
        If Trim$(NzStr(rows(outRow, 9))) = "" Then rows(outRow, 9) = NormalizeShipmentArea("", holdRows)
        If cCarrier > 0 Then rows(outRow, 10) = NzStr(src(r, cCarrier))
        If cReserve > 0 Then rows(outRow, 11) = NzStr(src(r, cReserve))
NextRow:
    Next r

    ShipmentsFormLoadLines = rows
    Exit Function

FailSoft:
End Function

Public Function ShipmentsFormCommitLine(ByVal targetName As String, _
                                        ByVal actionName As String, _
                                        ByVal tableRowIndex As Long, _
                                        ByVal refNumber As String, _
                                        ByVal itemName As String, _
                                        ByVal qtyValue As Double, _
                                        ByVal rowValue As Long, _
                                        ByVal uomValue As String, _
                                        ByVal locationValue As String, _
                                        ByVal descriptionValue As String, _
                                        ByVal carrierValue As String, _
                                        ByRef report As String) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim isHold As Boolean
    Dim previousVisibility As XlSheetVisibility
    Dim visibilityChanged As Boolean
    Dim previousEvents As Boolean
    Dim previousHandling As Boolean
    Dim finalQty As Double
    Dim mergedExisting As Boolean
    Dim invLo As ListObject
    Dim singleRow As Variant
    Dim releaseDeltas As Collection
    Dim reserveDeltas As Collection
    Dim errNotes As String
    Dim releaseEventId As String
    Dim reserveEventId As String
    Dim releasedTotal As Double
    Dim reservedTotal As Double
    Dim hadExistingReserve As Boolean

    actionName = UCase$(Trim$(actionName))
    isHold = (UCase$(Trim$(targetName)) = "HOLD")
    If Not ValidateShipmentCommitInputs(actionName, isHold, itemName, qtyValue, rowValue, carrierValue, report) Then Exit Function

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    If isHold Then
        Set lo = GetListObject(ws, TABLE_NOTSHIPPED)
    Else
        Set lo = GetListObject(ws, TABLE_SHIPMENTS)
    End If
    If lo Is Nothing Then
        report = "Shipment table not found."
        Exit Function
    End If
    BeginShippingTableMutation lo, previousVisibility, visibilityChanged, previousEvents, previousHandling

    If actionName = "DELETE" Then
        If tableRowIndex <= 0 Or tableRowIndex > lo.ListRows.Count Then
            report = "Select a row to remove."
            GoTo CleanExit
        End If
        If Not isHold Then
            singleRow = Array(tableRowIndex)
            If Trim$(ShipmentRowText(lo, tableRowIndex, COL_SHIPMENT_RESERVE_EVENT_ID)) <> "" Then
                Set invLo = GetShipmentReleaseInvSysTable(ws, report)
                If invLo Is Nothing Then
                    If report = "" Then report = "InventoryManagement!invSys table not found."
                    GoTo CleanExit
                End If
                Set releaseDeltas = BuildSelectedShipmentRowsDeltas(invLo, lo, singleRow, "Locked", errNotes)
                If releaseDeltas Is Nothing Then
                    If errNotes = "" Then errNotes = "Unable to build shipment release event."
                    report = errNotes
                    GoTo CleanExit
                End If
                releasedTotal = ApplyShipmentReleaseDeltasLocal(invLo, releaseDeltas, errNotes, True)
                If releasedTotal < 0 Then
                    If errNotes = "" Then errNotes = "Unable to release local shipment inventory."
                    report = errNotes
                    GoTo CleanExit
                End If
                If Not QueueShipmentsReleaseEvent(releaseDeltas, errNotes, releaseEventId) Then
                    If errNotes <> "" Then AppendNote report, errNotes
                    errNotes = ""
                End If
                If Not MarkShippingReservationRows(lo, singleRow, SHIP_RESERVATION_RELEASED, releaseEventId, report) Then GoTo CleanExit
            End If
        End If
        lo.ListRows(tableRowIndex).Delete
        InvalidateAggregates
        If isHold Then
            PersistHoldRowsLocal lo
        Else
            PersistActiveShipmentRowsLocal lo
        End If
        report = "Removed shipment row."
        ShipmentsFormCommitLine = True
        GoTo CleanExit
    End If

    itemName = Trim$(itemName)
    If itemName = "" Then
        report = "Select a shippable item."
        GoTo CleanExit
    End If
    If qtyValue <= 0 Then
        report = "Quantity must be greater than zero."
        GoTo CleanExit
    End If
    If Not isHold And actionName = "ADD" And Trim$(carrierValue) = "" Then
        report = "Select a Carrier."
        GoTo CleanExit
    End If
    If rowValue <= 0 Then
        report = "Selected shippable is missing ROW."
        GoTo CleanExit
    End If

    If actionName = "UPDATE" Then
        If tableRowIndex <= 0 Or tableRowIndex > lo.ListRows.Count Then
            report = "Select a row to update."
            GoTo CleanExit
        End If
        Set lr = lo.ListRows(tableRowIndex)
        hadExistingReserve = (Not isHold And Trim$(ShipmentRowText(lo, tableRowIndex, COL_SHIPMENT_RESERVE_EVENT_ID)) <> "")
    Else
        Set lr = FindShipmentLineByRefItemVersion(lo, Trim$(refNumber), itemName, Trim$(descriptionValue), Trim$(carrierValue))
        If Not lr Is Nothing Then
            finalQty = ExistingShipmentLineQuantity(lo, lr) + qtyValue
            mergedExisting = True
            hadExistingReserve = (Not isHold And Trim$(ShipmentRowText(lo, lr.Index, COL_SHIPMENT_RESERVE_EVENT_ID)) <> "")
        Else
            Set lr = FirstBlankListRowShipping(lo)
            If lr Is Nothing Then Set lr = lo.ListRows.Add
        End If
    End If
    If finalQty <= 0 Then finalQty = qtyValue

    If hadExistingReserve Then
        singleRow = Array(lr.Index)
        Set invLo = GetShipmentReleaseInvSysTable(ws, report)
        If invLo Is Nothing Then
            If report = "" Then report = "InventoryManagement!invSys table not found."
            GoTo CleanExit
        End If
        Set releaseDeltas = BuildSelectedShipmentRowsDeltas(invLo, lo, singleRow, "Locked", errNotes)
        If releaseDeltas Is Nothing Then
            If errNotes = "" Then errNotes = "Unable to build release event for the existing reservation."
            report = errNotes
            GoTo CleanExit
        End If
        If Not QueueShipmentsReleaseEvent(releaseDeltas, errNotes, releaseEventId) Then
            If errNotes = "" Then errNotes = "Unable to queue release event for the existing reservation."
            report = errNotes
            GoTo CleanExit
        End If
        releasedTotal = ApplyShipmentReleaseDeltasLocal(invLo, releaseDeltas, errNotes, True)
        If releasedTotal < 0 Then
            If errNotes = "" Then errNotes = "Unable to release the existing reservation locally."
            report = errNotes
            GoTo CleanExit
        End If
        If Not MarkShippingReservationRows(lo, singleRow, SHIP_RESERVATION_RELEASED, releaseEventId, report) Then GoTo CleanExit
        WriteValue lr, COL_SHIPMENT_RESERVE_EVENT_ID, vbNullString
        WriteValue lr, "AREA", "Warehouse"
    End If

    WriteValue lr, "REF_NUMBER", Trim$(refNumber)
    WriteValue lr, COL_SHIPMENT_LINE_ID, EnsureShipmentLineId(lo, lr.Index)
    If Trim$(ShipmentRowText(lo, lr.Index, COL_SHIPMENT_RESERVE_EVENT_ID)) = "" Or Not mergedExisting Then WriteValue lr, COL_SHIPMENT_RESERVE_EVENT_ID, vbNullString
    WriteValue lr, "ITEMS", itemName
    WriteValue lr, "QUANTITY", finalQty
    WriteValue lr, "ROW", rowValue
    WriteValue lr, "UOM", Trim$(uomValue)
    WriteValue lr, "LOCATION", Trim$(locationValue)
    WriteValue lr, "DESCRIPTION", Trim$(descriptionValue)
    If isHold Then
        If Trim$(ShipmentRowText(lo, lr.Index, "AREA")) = "" Then WriteValue lr, "AREA", "Warehouse"
    Else
        WriteValue lr, "AREA", "Warehouse"
    End If
    If Trim$(carrierValue) <> "" Or Not mergedExisting Then WriteValue lr, "CARRIER", Trim$(carrierValue)

    If Not isHold Then
        singleRow = Array(lr.Index)
        If invLo Is Nothing Then Set invLo = GetWritableShippingInvSysTable(ws, report)
        If invLo Is Nothing Then
            If report = "" Then report = "InventoryManagement!invSys table not found."
            GoTo CleanExit
        End If
        Set reserveDeltas = BuildSelectedShipmentRowsDeltas(invLo, lo, singleRow, "Warehouse", errNotes)
        If reserveDeltas Is Nothing Then
            If errNotes = "" Then errNotes = "Unable to build shipment reserve event."
            report = errNotes
            GoTo CleanExit
        End If
        If Not QueueShipmentsReserveEvent(reserveDeltas, errNotes, reserveEventId) Then
            If errNotes = "" Then errNotes = "Unable to queue shipment reserve event."
            report = errNotes
            GoTo CleanExit
        End If
        If Not UpsertShippingReservationForRow(lo, lr.Index, reserveEventId, report) Then GoTo CleanExit
        reservedTotal = ApplyShipmentDeltasLocal(invLo, reserveDeltas, errNotes)
        If reservedTotal < 0 Then
            If errNotes = "" Then errNotes = "Unable to lock selected shipment inventory locally."
            report = errNotes
            GoTo CleanExit
        End If
        ApplyStageVersionInventoryOverlayFromRows invLo, lo, singleRow
        WriteValue lr, COL_SHIPMENT_RESERVE_EVENT_ID, reserveEventId
        WriteValue lr, "AREA", "Warehouse"
    End If

    InvalidateAggregates
    If isHold Then
        PersistHoldRowsLocal lo
    Else
        PersistActiveShipmentRowsLocal lo
    End If
    If actionName = "UPDATE" Then
        report = "Updated shipment row."
    ElseIf mergedExisting Then
        report = "Added quantity to existing shipment row."
    Else
        report = "Added shipment row."
    End If
    If reserveEventId <> "" Then report = report & vbCrLf & "Locked " & Format$(reservedTotal, "0.###") & " package(s) for shipment." & vbCrLf & "Reserve EventID: " & reserveEventId
    ShipmentsFormCommitLine = True

CleanExit:
    EndShippingTableMutation lo, previousVisibility, visibilityChanged, previousEvents, previousHandling
    Exit Function

Fail:
    report = "Shipment row update failed: " & Err.Description
    On Error Resume Next
    EndShippingTableMutation lo, previousVisibility, visibilityChanged, previousEvents, previousHandling
    On Error GoTo 0
End Function

Public Function ValidateShipmentCommitInputsReportForTest(ByVal targetName As String, _
                                                          ByVal actionName As String, _
                                                          ByVal itemName As String, _
                                                          ByVal qtyValue As Double, _
                                                          ByVal rowValue As Long, _
                                                          ByVal carrierValue As String) As String
    Dim report As String

    If ValidateShipmentCommitInputs(UCase$(Trim$(actionName)), _
                                    (UCase$(Trim$(targetName)) = "HOLD"), _
                                    itemName, _
                                    qtyValue, _
                                    rowValue, _
                                    carrierValue, _
                                    report) Then
        ValidateShipmentCommitInputsReportForTest = "OK"
    Else
        ValidateShipmentCommitInputsReportForTest = report
    End If
End Function

Private Function ValidateShipmentCommitInputs(ByVal actionName As String, _
                                              ByVal isHold As Boolean, _
                                              ByVal itemName As String, _
                                              ByVal qtyValue As Double, _
                                              ByVal rowValue As Long, _
                                              ByVal carrierValue As String, _
                                              ByRef report As String) As Boolean
    ValidateShipmentCommitInputs = False
    actionName = UCase$(Trim$(actionName))
    If actionName = "DELETE" Then
        ValidateShipmentCommitInputs = True
        Exit Function
    End If
    itemName = Trim$(itemName)
    If itemName = "" Then
        report = "Select a shippable item."
        Exit Function
    End If
    If qtyValue <= 0 Then
        report = "Quantity must be greater than zero."
        Exit Function
    End If
    If Not isHold And actionName = "ADD" And Trim$(carrierValue) = "" Then
        report = "Select a Carrier."
        Exit Function
    End If
    If rowValue <= 0 Then
        report = "Selected shippable is missing ROW."
        Exit Function
    End If
    ValidateShipmentCommitInputs = True
End Function

Private Function FindShipmentLineByRefItemVersion(ByVal lo As ListObject, _
                                                  ByVal refNumber As String, _
                                                  ByVal itemName As String, _
                                                  ByVal versionText As String, _
                                                  Optional ByVal carrierText As String = "") As ListRow
    Dim cRef As Long
    Dim cItems As Long
    Dim cDesc As Long
    Dim cCarrier As Long
    Dim lr As ListRow

    refNumber = Trim$(refNumber)
    itemName = Trim$(itemName)
    versionText = Trim$(versionText)
    carrierText = Trim$(carrierText)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If refNumber = "" Or itemName = "" Then Exit Function

    cRef = ColumnIndex(lo, "REF_NUMBER")
    cItems = ColumnIndex(lo, "ITEMS")
    cDesc = ColumnIndex(lo, "DESCRIPTION")
    cCarrier = ColumnIndex(lo, "CARRIER")
    If cRef = 0 Or cItems = 0 Then Exit Function

    For Each lr In lo.ListRows
        If StrComp(Trim$(NzStr(lr.Range.Cells(1, cRef).Value)), refNumber, vbTextCompare) = 0 _
           And StrComp(Trim$(NzStr(lr.Range.Cells(1, cItems).Value)), itemName, vbTextCompare) = 0 Then
            If cDesc = 0 Or versionText = "" _
               Or StrComp(Trim$(NzStr(lr.Range.Cells(1, cDesc).Value)), versionText, vbTextCompare) = 0 Then
                If cCarrier = 0 Or carrierText = "" _
                   Or StrComp(Trim$(NzStr(lr.Range.Cells(1, cCarrier).Value)), carrierText, vbTextCompare) = 0 Then
                    Set FindShipmentLineByRefItemVersion = lr
                    Exit Function
                End If
            End If
        End If
NextLine:
    Next lr
End Function

Private Function ExistingShipmentLineQuantity(ByVal lo As ListObject, ByVal lr As ListRow) As Double
    Dim cQty As Long

    If lo Is Nothing Then Exit Function
    If lr Is Nothing Then Exit Function
    cQty = ColumnIndex(lo, "QUANTITY")
    If cQty = 0 Then Exit Function
    ExistingShipmentLineQuantity = NzDbl(lr.Range.Cells(1, cQty).Value)
End Function

Private Function ShipmentRowText(ByVal lo As ListObject, ByVal tableRowIndex As Long, ByVal columnName As String) As String
    Dim colIdx As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If tableRowIndex <= 0 Or tableRowIndex > lo.ListRows.Count Then Exit Function
    colIdx = ColumnIndex(lo, columnName)
    If colIdx = 0 Then Exit Function
    ShipmentRowText = NzStr(lo.DataBodyRange.Cells(tableRowIndex, colIdx).Value)
End Function

Private Function NormalizeShipmentArea(ByVal areaValue As String, Optional ByVal holdRows As Boolean = False) As String
    areaValue = Trim$(areaValue)
    If areaValue = "" Then
        NormalizeShipmentArea = "Warehouse"
    ElseIf StrComp(areaValue, "dock", vbTextCompare) = 0 _
        Or StrComp(areaValue, "shipping", vbTextCompare) = 0 _
        Or StrComp(areaValue, "shipments", vbTextCompare) = 0 _
        Or StrComp(areaValue, "shipment", vbTextCompare) = 0 Then
        NormalizeShipmentArea = "Shipments"
    ElseIf StrComp(areaValue, "hold", vbTextCompare) = 0 _
        Or StrComp(areaValue, "not shipped", vbTextCompare) = 0 Then
        NormalizeShipmentArea = "Hold"
    Else
        NormalizeShipmentArea = areaValue
    End If
End Function

Private Sub CopyShipmentLineRow(ByVal sourceTable As ListObject, ByVal targetTable As ListObject, ByVal sourceRowIndex As Long)
    Dim lr As ListRow
    Dim lc As ListColumn
    Dim targetIdx As Long

    If sourceTable Is Nothing Or targetTable Is Nothing Then Exit Sub
    If sourceTable.DataBodyRange Is Nothing Then Exit Sub
    If sourceRowIndex <= 0 Or sourceRowIndex > sourceTable.ListRows.Count Then Exit Sub

    Set lr = FirstBlankListRowShipping(targetTable)
    If lr Is Nothing Then Set lr = targetTable.ListRows.Add
    EnsureColumnExists targetTable, COL_SHIPMENT_LINE_ID
    EnsureColumnExists targetTable, COL_SHIPMENT_RESERVE_EVENT_ID
    For Each lc In sourceTable.ListColumns
        targetIdx = ColumnIndex(targetTable, CStr(lc.Name))
        If targetIdx > 0 Then
            lr.Range.Cells(1, targetIdx).Value = sourceTable.DataBodyRange.Cells(sourceRowIndex, lc.Index).Value
        End If
    Next lc
    WriteValue lr, COL_SHIPMENT_LINE_ID, EnsureShipmentLineId(targetTable, lr.Index)
End Sub

Private Sub SetShipmentRowsArea(ByVal loShip As ListObject, ByVal rowIndexes As Variant, ByVal areaText As String, ByVal carrierValue As String)
    Dim cArea As Long
    Dim cCarrier As Long
    Dim i As Long
    Dim rowIndex As Long

    If loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    cArea = ColumnIndex(loShip, "AREA")
    cCarrier = ColumnIndex(loShip, "CARRIER")
    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex >= 1 And rowIndex <= loShip.ListRows.Count Then
            If cArea > 0 Then loShip.DataBodyRange.Cells(rowIndex, cArea).Value = areaText
            If cCarrier > 0 And Trim$(carrierValue) <> "" Then loShip.DataBodyRange.Cells(rowIndex, cCarrier).Value = Trim$(carrierValue)
        End If
    Next i
End Sub

Private Sub SetShipmentRowsReserveEventId(ByVal loShip As ListObject, ByVal rowIndexes As Variant, ByVal reserveEventId As String)
    Dim i As Long
    Dim rowIndex As Long

    If loShip Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    EnsureColumnExists loShip, COL_SHIPMENT_RESERVE_EVENT_ID

    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex >= 1 And rowIndex <= loShip.ListRows.Count Then
            WriteValue loShip.ListRows(rowIndex), COL_SHIPMENT_RESERVE_EVENT_ID, Trim$(reserveEventId)
        End If
    Next i
End Sub

Private Function ShipmentRowIndexCount(ByVal rowIndexes As Variant) As Long
    If IsEmpty(rowIndexes) Then Exit Function
    ShipmentRowIndexCount = UBound(rowIndexes) - LBound(rowIndexes) + 1
End Function

Private Function ForEachShipmentReservationUpsert(ByVal loShip As ListObject, _
                                                  ByVal rowIndexes As Variant, _
                                                  ByVal reserveEventId As String, _
                                                  ByRef report As String) As Boolean
    Dim i As Long
    Dim rowIndex As Long

    ForEachShipmentReservationUpsert = True
    If loShip Is Nothing Then Exit Function
    If IsEmpty(rowIndexes) Then Exit Function
    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex >= 1 And rowIndex <= loShip.ListRows.Count Then
            If Not UpsertShippingReservationForRow(loShip, rowIndex, reserveEventId, report) Then
                ForEachShipmentReservationUpsert = False
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ShipmentRowsByReserveState(ByVal loShip As ListObject, ByVal rowIndexes As Variant, ByVal requireReserve As Boolean) As Variant
    Dim selected As Collection
    Dim cReserve As Long
    Dim i As Long
    Dim rowIndex As Long
    Dim hasReserve As Boolean
    Dim result() As Long

    If loShip Is Nothing Then Exit Function
    If IsEmpty(rowIndexes) Then Exit Function
    Set selected = New Collection
    cReserve = ColumnIndex(loShip, COL_SHIPMENT_RESERVE_EVENT_ID)

    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex < 1 Or rowIndex > loShip.ListRows.Count Then GoTo NextRow
        hasReserve = False
        If cReserve > 0 Then hasReserve = (Trim$(NzStr(loShip.DataBodyRange.Cells(rowIndex, cReserve).Value)) <> "")
        If hasReserve = requireReserve Then selected.Add rowIndex
NextRow:
    Next i

    If selected.Count = 0 Then Exit Function
    ReDim result(0 To selected.Count - 1)
    For i = 1 To selected.Count
        result(i - 1) = CLng(selected(i))
    Next i
    ShipmentRowsByReserveState = result
End Function

Private Sub SetShipmentRowsCarrier(ByVal loShip As ListObject, ByVal rowIndexes As Variant, ByVal carrierValue As String)
    Dim cCarrier As Long
    Dim i As Long
    Dim rowIndex As Long

    If loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    cCarrier = ColumnIndex(loShip, "CARRIER")
    If cCarrier = 0 Then Exit Sub
    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex >= 1 And rowIndex <= loShip.ListRows.Count Then
            loShip.DataBodyRange.Cells(rowIndex, cCarrier).Value = Trim$(carrierValue)
        End If
    Next i
End Sub

Private Sub DeleteShipmentRows(ByVal loShip As ListObject, ByVal rowIndexes As Variant)
    Dim i As Long
    Dim rowIndex As Long

    If loShip Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    For i = UBound(rowIndexes) To LBound(rowIndexes) Step -1
        rowIndex = CLng(rowIndexes(i))
        If rowIndex >= 1 And rowIndex <= loShip.ListRows.Count Then
            loShip.ListRows(rowIndex).Delete
        End If
    Next i
End Sub

Private Function BuildSelectedShipmentRowsDeltas(ByVal invLo As ListObject, _
                                                 ByVal loShip As ListObject, _
                                                 ByVal rowIndexes As Variant, _
                                                 ByVal requiredArea As String, _
                                                 ByRef errNotes As String) As Collection
    Dim cQtyShip As Long
    Dim cRowShip As Long
    Dim cItemShip As Long
    Dim cArea As Long
    Dim cDesc As Long
    Dim colItemCode As Long
    Dim colItemName As Long
    Dim colTotalInv As Long
    Dim colShipments As Long
    Dim requirements As Object
    Dim versionRequirements As Object
    Dim versionNames As Object
    Dim names As Object
    Dim versionLabel As String
    Dim i As Long
    Dim rowIndex As Long
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim currentArea As String
    Dim reqKey As String
    Dim result As New Collection
    Dim key As Variant
    Dim requireAreaMatch As Boolean
    Dim useShipmentStaging As Boolean
    Dim allowMissingShipmentStaging As Boolean

    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        Exit Function
    End If
    If invLo.DataBodyRange Is Nothing Then
        errNotes = "invSys has no inventory rows."
        Exit Function
    End If
    If loShip Is Nothing Then
        errNotes = "Shipment table not found."
        Exit Function
    End If
    If loShip.DataBodyRange Is Nothing Then
        errNotes = "No shipment rows are ready."
        Exit Function
    End If
    If IsEmpty(rowIndexes) Then
        errNotes = "Select shipment row(s) first."
        Exit Function
    End If

    cQtyShip = ColumnIndex(loShip, "QUANTITY")
    cRowShip = ColumnIndex(loShip, "ROW")
    cItemShip = ColumnIndex(loShip, "ITEMS")
    cArea = ColumnIndex(loShip, "AREA")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    If cQtyShip = 0 Or cRowShip = 0 Then
        errNotes = "Shipments table missing QUANTITY/ROW columns."
        Exit Function
    End If

    Set requirements = CreateObject("Scripting.Dictionary")
    requirements.CompareMode = vbTextCompare
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare
    Set versionRequirements = CreateObject("Scripting.Dictionary")
    versionRequirements.CompareMode = vbTextCompare
    Set versionNames = CreateObject("Scripting.Dictionary")
    versionNames.CompareMode = vbTextCompare
    requireAreaMatch = (StrComp(requiredArea, "Locked", vbTextCompare) <> 0)
    useShipmentStaging = (StrComp(requiredArea, "Shipments", vbTextCompare) = 0 Or StrComp(requiredArea, "Locked", vbTextCompare) = 0)
    allowMissingShipmentStaging = (StrComp(requiredArea, "Locked", vbTextCompare) = 0)
    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex <= 0 Or rowIndex > loShip.ListRows.Count Then
            AppendNote errNotes, "Selected shipment row " & CStr(rowIndex) & " is no longer valid."
            Exit Function
        End If
        If cArea > 0 Then currentArea = NormalizeShipmentArea(NzStr(loShip.DataBodyRange.Cells(rowIndex, cArea).Value)) Else currentArea = "Warehouse"
        If requireAreaMatch And StrComp(currentArea, requiredArea, vbTextCompare) <> 0 Then
            If StrComp(requiredArea, "Shipments", vbTextCompare) = 0 Then
                AppendNote errNotes, "Selected row " & CStr(rowIndex) & " is in " & currentArea & ". Use To Shipments before Shipments Sent."
            Else
                AppendNote errNotes, "Selected row " & CStr(rowIndex) & " is already in " & currentArea & "."
            End If
            Exit Function
        End If

        rowVal = NzLng(loShip.DataBodyRange.Cells(rowIndex, cRowShip).Value)
        qtyVal = NzDbl(loShip.DataBodyRange.Cells(rowIndex, cQtyShip).Value)
        If rowVal = 0 Or qtyVal <= 0 Then GoTo NextSelectedRow
        versionLabel = ""
        If cDesc > 0 Then versionLabel = NormalizeBoxBomVersionLabelShipping(NzStr(loShip.DataBodyRange.Cells(rowIndex, cDesc).Value))
        reqKey = CStr(rowVal)
        If useShipmentStaging And versionLabel <> "" Then reqKey = CStr(rowVal) & "|" & versionLabel
        If requirements.Exists(reqKey) Then
            requirements(reqKey) = NzDbl(requirements(reqKey)) + qtyVal
        Else
            requirements.Add reqKey, qtyVal
            If cItemShip > 0 Then names.Add reqKey, NzStr(loShip.DataBodyRange.Cells(rowIndex, cItemShip).Value)
        End If
        If StrComp(requiredArea, "Warehouse", vbTextCompare) = 0 And cDesc > 0 Then
            If versionLabel <> "" Then
                Dim versionKey As String: versionKey = CStr(rowVal) & "|" & versionLabel
                If versionRequirements.Exists(versionKey) Then
                    versionRequirements(versionKey) = NzDbl(versionRequirements(versionKey)) + qtyVal
                Else
                    versionRequirements.Add versionKey, qtyVal
                    If cItemShip > 0 Then versionNames.Add versionKey, NzStr(loShip.DataBodyRange.Cells(rowIndex, cItemShip).Value)
                End If
            End If
        End If
NextSelectedRow:
    Next i

    If requirements.Count = 0 Then
        errNotes = "No selected shipment quantities were found."
        Exit Function
    End If
    If StrComp(requiredArea, "Warehouse", vbTextCompare) = 0 Then
        If Not SelectedVersionInventoryAvailable(invLo, versionRequirements, versionNames, errNotes) Then Exit Function
    End If

    colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    colItemName = ColumnIndex(invLo, "ITEM")
    If StrComp(requiredArea, "Warehouse", vbTextCompare) = 0 Then
        colTotalInv = ColumnIndex(invLo, "TOTAL INV")
        If colTotalInv = 0 Then
            errNotes = "invSys table missing TOTAL INV column."
            Exit Function
        End If
    ElseIf useShipmentStaging Then
        colShipments = ColumnIndex(invLo, "SHIPMENTS")
        If colShipments = 0 Then
            errNotes = "invSys table missing SHIPMENTS column."
            Exit Function
        End If
    End If
    For Each key In requirements.Keys
        Dim rowKeyValue As Long: rowKeyValue = ShipmentRequirementRowValue(CStr(key))
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, rowKeyValue)
        If invRow Is Nothing Then
            AppendNote errNotes, "Package ROW " & CStr(rowKeyValue) & " not found in invSys."
            Exit Function
        End If
        If colTotalInv > 0 Then
            Dim availableQty As Double: availableQty = NzDbl(invRow.Range.Cells(1, colTotalInv).Value)
            If NzDbl(requirements(key)) > availableQty + 0.0000001 Then
                AppendNote errNotes, "ROW " & CStr(rowKeyValue) & " requires " & Format$(NzDbl(requirements(key)), "0.###") & " but only " & Format$(availableQty, "0.###") & " in TOTAL INV."
                Exit Function
            End If
        ElseIf colShipments > 0 Then
            Dim stagedQty As Double: stagedQty = NzDbl(invRow.Range.Cells(1, colShipments).Value)
            If Not allowMissingShipmentStaging And NzDbl(requirements(key)) > stagedQty + 0.0000001 Then
                AppendNote errNotes, "ROW " & CStr(rowKeyValue) & " only has " & Format$(stagedQty, "0.###") & " staged but needs " & Format$(NzDbl(requirements(key)), "0.###") & "."
                Exit Function
            End If
        End If

        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = rowKeyValue
        delta("QTY") = NzDbl(requirements(key))
        versionLabel = ShipmentRequirementVersionLabel(CStr(key))
        If versionLabel <> "" Then delta("VERSION") = versionLabel
        If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(invRow.Range.Cells(1, colItemCode).Value)
        If colItemName > 0 Then
            delta("ITEM_NAME") = NzStr(invRow.Range.Cells(1, colItemName).Value)
        ElseIf names.Exists(CStr(key)) Then
            delta("ITEM_NAME") = NzStr(names(CStr(key)))
        End If
        result.Add delta
    Next key

    Set BuildSelectedShipmentRowsDeltas = result
End Function

Private Function ShipmentRequirementRowValue(ByVal requirementKey As String) As Long
    Dim parts As Variant

    parts = Split(CStr(requirementKey), "|")
    ShipmentRequirementRowValue = CLng(Val(CStr(parts(0))))
End Function

Private Function ShipmentRequirementVersionLabel(ByVal requirementKey As String) As String
    Dim parts As Variant

    parts = Split(CStr(requirementKey), "|")
    If UBound(parts) >= 1 Then ShipmentRequirementVersionLabel = NormalizeBoxBomVersionLabelShipping(CStr(parts(1)))
End Function

Private Function SelectedVersionInventoryAvailable(ByVal invLo As ListObject, _
                                                   ByVal versionRequirements As Object, _
                                                   ByVal versionNames As Object, _
                                                   ByRef errNotes As String) As Boolean
    Dim key As Variant
    Dim parts As Variant
    Dim rowVal As Long
    Dim versionLabel As String
    Dim itemName As String
    Dim versionInv As Object
    Dim availableQty As Double
    Dim requiredQty As Double
    Dim hasVersionQty As Boolean
    Dim localStagedQty As Double
    Dim nasReservedQty As Double

    SelectedVersionInventoryAvailable = True
    If versionRequirements Is Nothing Then Exit Function
    If versionRequirements.Count = 0 Then Exit Function

    For Each key In versionRequirements.Keys
        parts = Split(CStr(key), "|")
        If UBound(parts) < 1 Then GoTo NextKey
        rowVal = CLng(Val(CStr(parts(0))))
        versionLabel = NormalizeBoxBomVersionLabelShipping(CStr(parts(1)))
        itemName = ""
        If Not versionNames Is Nothing Then
            If versionNames.Exists(CStr(key)) Then itemName = NzStr(versionNames(CStr(key)))
        End If

        Set versionInv = BoxMakerFormLoadBoxVersionInventory(rowVal, itemName)
        availableQty = 0#
        hasVersionQty = False
        If Not versionInv Is Nothing Then
            If versionInv.Exists(versionLabel) Then
                availableQty = NzDbl(versionInv(versionLabel))
                hasVersionQty = True
            End If
        End If
        If Not hasVersionQty Then
            availableQty = SingleVersionFallbackAvailableQty(invLo, rowVal, itemName, versionLabel)
        End If
        If availableQty <= 0.0000001 Then
            availableQty = PickerVersionAvailableQty(rowVal, itemName, versionLabel)
        End If
        localStagedQty = StagedShipmentVersionQty(rowVal, versionLabel)
        nasReservedQty = ActiveNasShippingReservationQty(rowVal, versionLabel)
        If nasReservedQty > localStagedQty Then
            availableQty = availableQty - nasReservedQty
        Else
            availableQty = availableQty - localStagedQty
        End If
        If availableQty < 0 Then availableQty = 0
        requiredQty = NzDbl(versionRequirements(CStr(key)))
        If requiredQty > availableQty + 0.0000001 Then
            If itemName = "" Then itemName = "ROW " & CStr(rowVal)
            AppendNote errNotes, itemName & " " & versionLabel & " requires " & Format$(requiredQty, "0.###") & " but only " & Format$(availableQty, "0.###") & " is available for that version."
            SelectedVersionInventoryAvailable = False
            Exit Function
        End If
NextKey:
    Next key
End Function

Private Function SingleVersionFallbackAvailableQty(ByVal invLo As ListObject, _
                                                   ByVal rowVal As Long, _
                                                   ByVal itemName As String, _
                                                   ByVal versionLabel As String) As Double
    Dim ws As Worksheet
    Dim invIdx As Long
    Dim totalVal As Variant
    Dim foundCurrent As Boolean
    Dim snapshotCache As Object

    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If rowVal <= 0 Or versionLabel = "" Then Exit Function
    If CountActiveVersionsForPackageShipping(rowVal) <> 1 Then Exit Function

    If Not invLo Is Nothing Then
        invIdx = FindInvRowIndexByRow(invLo, rowVal)
        If invIdx <= 0 And Trim$(itemName) <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
        If invIdx > 0 Then
            totalVal = GetInvSysValueByIndex(invLo, invIdx, "TOTAL INV")
            If Not IsBlankInventoryValue(totalVal) Then
                SingleVersionFallbackAvailableQty = NzDbl(totalVal)
                Exit Function
            End If
        End If
    End If

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    totalVal = ResolveCurrentInventoryValue(ws, invLo, rowVal, itemName, foundCurrent, snapshotCache)
    If foundCurrent Then SingleVersionFallbackAvailableQty = NzDbl(totalVal)
End Function

Private Function PickerVersionAvailableQty(ByVal rowVal As Long, _
                                           ByVal itemName As String, _
                                           ByVal versionLabel As String) As Double
    On Error GoTo CleanFail

    Dim savedBoxes As Variant
    Dim shippables As Variant
    Dim r As Long
    Dim rawQty As String

    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If rowVal <= 0 Or versionLabel = "" Then Exit Function

    savedBoxes = BoxMakerFormLoadSavedBoxes()
    If IsEmpty(savedBoxes) Then Exit Function
    shippables = BoxMakerFormLoadShippableVersionInventory(savedBoxes)
    If IsEmpty(shippables) Then Exit Function

    For r = 1 To UBound(shippables, 1)
        If NzLng(shippables(r, 1)) = rowVal _
           And StrComp(NormalizeBoxBomVersionLabelShipping(NzStr(shippables(r, 3))), versionLabel, vbTextCompare) = 0 Then
            rawQty = Trim$(NzStr(shippables(r, 4)))
            If rawQty <> "" And LCase$(rawQty) <> "unknown" Then PickerVersionAvailableQty = NzDbl(rawQty)
            Exit Function
        End If
    Next r

CleanFail:
End Function

Private Function CountActiveVersionsForPackageShipping(ByVal packageRow As Long) As Long
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim loSource As ListObject
    Dim wbRuntime As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim versions As Variant
    Dim ignoredCount As Long

    If packageRow <= 0 Then Exit Function
    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function

    Set loSource = BoxMakerShippingBomSourceTable(ws, wbRuntime, openedTransient, report)
    If loSource Is Nothing Then GoTo CleanExit

    versions = BuildBoxBomVersionRows(loSource, packageRow, ignoredCount)
    CountActiveVersionsForPackageShipping = CountActiveBoxBomVersionsShipping(versions)

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
    Exit Function

CleanFail:
    If openedTransient Then CloseWorkbookNoSaveShipping wbRuntime
End Function

Private Function StagedShipmentVersionQty(ByVal rowVal As Long, ByVal versionLabel As String) As Double
    Dim ws As Worksheet
    Dim loShip As ListObject
    Dim cRow As Long
    Dim cQty As Long
    Dim cArea As Long
    Dim cDesc As Long
    Dim r As Long

    If rowVal <= 0 Then Exit Function
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then Exit Function
    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    If loShip Is Nothing Then Exit Function
    If loShip.DataBodyRange Is Nothing Then Exit Function

    cRow = ColumnIndex(loShip, "ROW")
    cQty = ColumnIndex(loShip, "QUANTITY")
    cArea = ColumnIndex(loShip, "AREA")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    If cRow = 0 Or cQty = 0 Or cArea = 0 Or cDesc = 0 Then Exit Function

    For r = 1 To loShip.ListRows.Count
        If NzLng(loShip.DataBodyRange.Cells(r, cRow).Value) = rowVal _
           And StrComp(NormalizeShipmentArea(NzStr(loShip.DataBodyRange.Cells(r, cArea).Value)), "Shipments", vbTextCompare) = 0 _
           And StrComp(NormalizeBoxBomVersionLabelShipping(NzStr(loShip.DataBodyRange.Cells(r, cDesc).Value)), versionLabel, vbTextCompare) = 0 Then
            StagedShipmentVersionQty = StagedShipmentVersionQty + NzDbl(loShip.DataBodyRange.Cells(r, cQty).Value)
        End If
    Next r
End Function

Public Function ShipmentsFormMoveHold(ByVal refNumber As String, _
                                      ByVal itemName As String, _
                                      ByVal qtyValue As Double, _
                                      ByVal moveToHold As Boolean, _
                                      ByRef report As String) As Boolean
    report = MoveShipmentHoldForAutomation(refNumber, itemName, qtyValue, moveToHold)
    ShipmentsFormMoveHold = (Left$(report, 3) = "OK|")
End Function

Public Function ShipmentsFormMoveHoldRows(ByVal rowIndexes As Variant, _
                                          ByVal moveToHold As Boolean, _
                                          ByRef report As String) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Dim sourceTable As ListObject
    Dim targetTable As ListObject
    Dim previousVisibility As XlSheetVisibility
    Dim visibilityChanged As Boolean
    Dim previousEvents As Boolean
    Dim previousHandling As Boolean
    Dim mutationStarted As Boolean
    Dim movedRows As Long
    Dim i As Long
    Dim rowIndex As Long
    Dim invLo As ListObject
    Dim releaseRows As Variant
    Dim releaseDeltas As Collection
    Dim errNotes As String
    Dim releaseEventId As String
    Dim releasedTotal As Double

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    If moveToHold Then
        Set sourceTable = GetListObject(ws, TABLE_SHIPMENTS)
        Set targetTable = GetListObject(ws, TABLE_NOTSHIPPED)
    Else
        Set sourceTable = GetListObject(ws, TABLE_NOTSHIPPED)
        Set targetTable = GetListObject(ws, TABLE_SHIPMENTS)
    End If
    If sourceTable Is Nothing Or targetTable Is Nothing Then
        report = "Shipment hold tables not found."
        Exit Function
    End If
    If IsEmpty(rowIndexes) Then
        report = "Select shipment row(s) first."
        Exit Function
    End If
    If moveToHold Then
        releaseRows = ShipmentRowsByReserveState(sourceTable, rowIndexes, True)
        If Not IsEmpty(releaseRows) Then
            Set invLo = GetWritableShippingInvSysTable(ws, report)
            If invLo Is Nothing Then
                If report = "" Then report = "InventoryManagement!invSys table not found."
                Exit Function
            End If
            Set releaseDeltas = BuildSelectedShipmentRowsDeltas(invLo, sourceTable, releaseRows, "Shipments", errNotes)
            If releaseDeltas Is Nothing Then
                If errNotes = "" Then errNotes = "Unable to build shipment release event."
                report = errNotes
                Exit Function
            End If
            If Not QueueShipmentsReleaseEvent(releaseDeltas, errNotes, releaseEventId) Then
                If errNotes = "" Then errNotes = "Unable to queue shipment release event."
                report = errNotes
                Exit Function
            End If
            releasedTotal = ApplyShipmentReleaseDeltasLocal(invLo, releaseDeltas, errNotes)
            If releasedTotal < 0 Then
                If errNotes = "" Then errNotes = "Unable to release local shipment inventory."
                report = errNotes
                Exit Function
            End If
            If Not MarkShippingReservationRows(sourceTable, releaseRows, SHIP_RESERVATION_RELEASED, releaseEventId, report) Then Exit Function
            SetShipmentRowsReserveEventId sourceTable, releaseRows, vbNullString
        End If
    End If

    BeginShippingTableMutation sourceTable, previousVisibility, visibilityChanged, previousEvents, previousHandling
    mutationStarted = True
    For i = UBound(rowIndexes) To LBound(rowIndexes) Step -1
        rowIndex = CLng(rowIndexes(i))
        If rowIndex >= 1 And rowIndex <= sourceTable.ListRows.Count Then
            CopyShipmentLineRow sourceTable, targetTable, rowIndex
            sourceTable.ListRows(rowIndex).Delete
            movedRows = movedRows + 1
        End If
    Next i
    If moveToHold Then
        PersistActiveShipmentRowsLocal sourceTable
        PersistHoldRowsLocal targetTable
    Else
        PersistHoldRowsLocal sourceTable
        PersistActiveShipmentRowsLocal targetTable
    End If

    InvalidateAggregates True
    If movedRows = 0 Then
        report = "No selected shipment rows were moved."
    ElseIf moveToHold Then
        report = "Moved " & CStr(movedRows) & " row(s) to Not Shipped."
        If releasedTotal > 0 Then report = report & vbCrLf & "Released " & Format$(releasedTotal, "0.###") & " package(s) back to warehouse."
        If releaseEventId <> "" Then report = report & vbCrLf & "Release EventID: " & releaseEventId
    Else
        report = "Returned " & CStr(movedRows) & " row(s) to Shipments."
    End If
    ShipmentsFormMoveHoldRows = (movedRows > 0)

CleanExit:
    If mutationStarted Then EndShippingTableMutation sourceTable, previousVisibility, visibilityChanged, previousEvents, previousHandling
    Exit Function

Fail:
    report = "Hold action failed: " & Err.Description
    Resume CleanExit
End Function

Private Sub PersistHoldRowsLocal(ByVal loHold As ListObject)
    PersistShipmentRowsLocal loHold, PersistentHoldRowsPath()
End Sub

Private Sub PersistActiveShipmentRowsLocal(ByVal loShip As ListObject)
    PersistShipmentRowsLocal loShip, PersistentActiveShipmentRowsPath()
End Sub

Private Sub PersistShipmentRowsLocal(ByVal lo As ListObject, ByVal filePath As String)
    On Error GoTo CleanExit

    Dim fileNum As Integer
    Dim r As Long
    Dim lineText As String

    If lo Is Nothing Then Exit Sub
    If filePath = "" Then Exit Sub
    EnsureLocalFolderExistsShipping ParentFolderPathShipping(filePath)
    EnsureColumnExists lo, COL_SHIPMENT_LINE_ID
    EnsureColumnExists lo, COL_SHIPMENT_RESERVE_EVENT_ID
    EnsureShipmentLineIds lo

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.ListRows.Count
            lineText = HoldRowField(lo, r, "REF_NUMBER") & vbTab & _
                       HoldRowField(lo, r, "ITEMS") & vbTab & _
                       HoldRowField(lo, r, "QUANTITY") & vbTab & _
                       HoldRowField(lo, r, "ROW") & vbTab & _
                       HoldRowField(lo, r, "UOM") & vbTab & _
                       HoldRowField(lo, r, "LOCATION") & vbTab & _
                       HoldRowField(lo, r, "DESCRIPTION") & vbTab & _
                       HoldRowField(lo, r, "AREA") & vbTab & _
                       HoldRowField(lo, r, "CARRIER") & vbTab & _
                       HoldRowField(lo, r, COL_SHIPMENT_LINE_ID) & vbTab & _
                       HoldRowField(lo, r, COL_SHIPMENT_RESERVE_EVENT_ID)
            Print #fileNum, lineText
        Next r
    End If
    Close #fileNum

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
End Sub

Private Sub LoadPersistentHoldRowsLocal(ByVal loHold As ListObject)
    LoadPersistentShipmentRowsLocal loHold, PersistentHoldRowsPath(), "Hold"
End Sub

Private Sub LoadPersistentActiveShipmentRowsLocal(ByVal loShip As ListObject)
    LoadPersistentShipmentRowsLocal loShip, PersistentActiveShipmentRowsPath(), "Warehouse"
End Sub

Private Sub LoadPersistentShipmentRowsLocal(ByVal lo As ListObject, ByVal filePath As String, ByVal defaultArea As String)
    On Error GoTo CleanExit

    Static loading As Boolean
    Dim fileNum As Integer
    Dim lineText As String
    Dim parts As Variant
    Dim lr As ListRow

    If loading Then Exit Sub
    If lo Is Nothing Then Exit Sub
    If filePath = "" Then Exit Sub
    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Sub

    loading = True
    EnsureShippingWorksheetEditable lo.Parent
    EnsureColumnExists lo, COL_SHIPMENT_LINE_ID
    EnsureColumnExists lo, COL_SHIPMENT_RESERVE_EVENT_ID
    ClearListObjectData lo

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        If Trim$(lineText) <> "" Then
            parts = Split(lineText, vbTab)
            If StrComp(defaultArea, "Warehouse", vbTextCompare) = 0 Then
                If PersistentSentShipmentRowExists(parts) Then GoTo NextPersistedLine
            End If
            Set lr = FirstBlankListRowShipping(lo)
            If lr Is Nothing Then Set lr = lo.ListRows.Add
            WriteValue lr, "REF_NUMBER", HoldPart(parts, 0)
            WriteValue lr, "ITEMS", HoldPart(parts, 1)
            WriteValue lr, "QUANTITY", HoldPart(parts, 2)
            WriteValue lr, "ROW", HoldPart(parts, 3)
            WriteValue lr, "UOM", HoldPart(parts, 4)
            WriteValue lr, "LOCATION", HoldPart(parts, 5)
            WriteValue lr, "DESCRIPTION", HoldPart(parts, 6)
            WriteValue lr, "AREA", IIf(HoldPart(parts, 7) = "", defaultArea, HoldPart(parts, 7))
            WriteValue lr, "CARRIER", HoldPart(parts, 8)
            If Trim$(HoldPart(parts, 9)) <> "" Then
                WriteValue lr, COL_SHIPMENT_LINE_ID, HoldPart(parts, 9)
            Else
                WriteValue lr, COL_SHIPMENT_LINE_ID, NewShipmentLineId()
            End If
            WriteValue lr, COL_SHIPMENT_RESERVE_EVENT_ID, HoldPart(parts, 10)
        End If
NextPersistedLine:
    Loop
    Close #fileNum

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    loading = False
    On Error GoTo 0
End Sub

Private Function HoldRowField(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim c As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    c = ColumnIndex(lo, columnName)
    If c = 0 Then Exit Function
    HoldRowField = EscapeHoldField(NzStr(lo.DataBodyRange.Cells(rowIndex, c).Value))
End Function

Private Function HoldPart(ByVal parts As Variant, ByVal partIndex As Long) As String
    On Error GoTo CleanExit

    If IsArray(parts) Then
        If UBound(parts) >= partIndex Then HoldPart = UnescapeHoldField(CStr(parts(partIndex)))
    End If

CleanExit:
End Function

Private Sub AppendSentShipmentRowsLocal(ByVal loShip As ListObject, ByVal rowIndexes As Variant)
    On Error GoTo CleanExit

    Dim filePath As String
    Dim fileNum As Integer
    Dim i As Long
    Dim rowIndex As Long
    Dim keyText As String

    If loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    EnsureColumnExists loShip, COL_SHIPMENT_LINE_ID
    EnsureShipmentLineIds loShip
    filePath = PersistentSentShipmentRowsPath()
    If filePath = "" Then Exit Sub
    EnsureLocalFolderExistsShipping ParentFolderPathShipping(filePath)

    fileNum = FreeFile
    Open filePath For Append As #fileNum
    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex >= 1 And rowIndex <= loShip.ListRows.Count Then
            keyText = SentShipmentPersistTokenFromTableRow(loShip, rowIndex)
            If keyText <> "" Then Print #fileNum, EscapeHoldField(keyText)
        End If
    Next i
    Close #fileNum

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
End Sub

Private Function PersistentSentShipmentRowExists(ByVal parts As Variant) As Boolean
    On Error GoTo CleanExit

    Dim filePath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim wantedLineId As String
    Dim wantedKey As String

    wantedLineId = ShipmentLineIdFromParts(parts)
    wantedKey = ShipmentPersistKeyFromParts(parts)
    If wantedLineId = "" And wantedKey = "" Then Exit Function

    filePath = PersistentSentShipmentRowsPath()
    If filePath = "" Then Exit Function
    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Function

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = UnescapeHoldField(lineText)
        If ShipmentSentTokenMatches(lineText, wantedLineId, wantedKey) Then
            PersistentSentShipmentRowExists = True
            Exit Do
        End If
    Loop
    Close #fileNum

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
End Function

Private Function PersistentSentShipmentLineIdExists(ByVal lineId As String) As Boolean
    PersistentSentShipmentLineIdExists = PersistentSentShipmentLineIdExistsForWarehouse(lineId, CurrentShippingWarehouseIdForLocalState())
End Function

Private Function PersistentSentShipmentLineIdExistsForWarehouse(ByVal lineId As String, ByVal warehouseId As String) As Boolean
    On Error GoTo CleanExit

    Dim filePath As String
    Dim fileNum As Integer
    Dim lineText As String

    lineId = Trim$(lineId)
    If lineId = "" Then Exit Function

    filePath = PersistentSentShipmentRowsPathForWarehouse(warehouseId)
    If filePath = "" Then Exit Function
    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Function

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = UnescapeHoldField(lineText)
        If ShipmentSentTokenMatches(lineText, lineId, vbNullString) Then
            PersistentSentShipmentLineIdExistsForWarehouse = True
            Exit Do
        End If
    Loop
    Close #fileNum

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
End Function

Private Function SentShipmentPersistTokenFromTableRow(ByVal lo As ListObject, ByVal rowIndex As Long) As String
    Dim lineId As String

    lineId = EnsureShipmentLineId(lo, rowIndex)
    If lineId <> "" Then
        SentShipmentPersistTokenFromTableRow = "ID:" & lineId
    Else
        SentShipmentPersistTokenFromTableRow = ShipmentPersistKeyFromTableRow(lo, rowIndex)
    End If
End Function

Private Function ShipmentSentTokenMatches(ByVal sentToken As String, ByVal wantedLineId As String, ByVal wantedKey As String) As Boolean
    sentToken = Trim$(sentToken)
    If sentToken = "" Then Exit Function
    If Left$(sentToken, 3) = "ID:" Then
        If wantedLineId <> "" Then
            ShipmentSentTokenMatches = (StrComp(Mid$(sentToken, 4), wantedLineId, vbTextCompare) = 0)
        End If
    ElseIf wantedKey <> "" Then
        ShipmentSentTokenMatches = (StrComp(sentToken, wantedKey, vbTextCompare) = 0)
    End If
End Function

Private Function ShipmentLineIdFromParts(ByVal parts As Variant) As String
    ShipmentLineIdFromParts = Trim$(HoldPart(parts, 9))
End Function

Private Sub EnsureShipmentLineIds(ByVal lo As ListObject)
    Dim r As Long

    If lo Is Nothing Then Exit Sub
    EnsureColumnExists lo, COL_SHIPMENT_LINE_ID
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For r = 1 To lo.ListRows.Count
        EnsureShipmentLineId lo, r
    Next r
End Sub

Private Function EnsureShipmentLineId(ByVal lo As ListObject, ByVal rowIndex As Long) As String
    Dim cLineId As Long
    Dim existingId As String

    If lo Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    EnsureColumnExists lo, COL_SHIPMENT_LINE_ID
    cLineId = ColumnIndex(lo, COL_SHIPMENT_LINE_ID)
    If cLineId = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    existingId = Trim$(NzStr(lo.DataBodyRange.Cells(rowIndex, cLineId).Value))
    If existingId = "" Then
        existingId = NewShipmentLineId()
        lo.DataBodyRange.Cells(rowIndex, cLineId).Value = existingId
    End If
    EnsureShipmentLineId = existingId
End Function

Private Function NewShipmentLineId() As String
    NewShipmentLineId = Trim$(modUR_Snapshot.GenerateGUID())
    If NewShipmentLineId = "" Then NewShipmentLineId = "SHIPLINE-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(CLng(Timer * 1000), "00000000")
End Function

Private Function ShipmentPersistKeyFromTableRow(ByVal lo As ListObject, ByVal rowIndex As Long) As String
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function

    ShipmentPersistKeyFromTableRow = ShipmentPersistKey( _
        HoldRowFieldRaw(lo, rowIndex, "REF_NUMBER"), _
        HoldRowFieldRaw(lo, rowIndex, "ITEMS"), _
        HoldRowFieldRaw(lo, rowIndex, "QUANTITY"), _
        HoldRowFieldRaw(lo, rowIndex, "ROW"), _
        HoldRowFieldRaw(lo, rowIndex, "UOM"), _
        HoldRowFieldRaw(lo, rowIndex, "LOCATION"), _
        HoldRowFieldRaw(lo, rowIndex, "DESCRIPTION"))
End Function

Private Function ShipmentPersistKeyFromParts(ByVal parts As Variant) As String
    ShipmentPersistKeyFromParts = ShipmentPersistKey( _
        HoldPart(parts, 0), _
        HoldPart(parts, 1), _
        HoldPart(parts, 2), _
        HoldPart(parts, 3), _
        HoldPart(parts, 4), _
        HoldPart(parts, 5), _
        HoldPart(parts, 6))
End Function

Private Function ShipmentPersistKey(ByVal refNumber As String, _
                                    ByVal itemName As String, _
                                    ByVal qtyText As String, _
                                    ByVal rowText As String, _
                                    ByVal uomText As String, _
                                    ByVal locationText As String, _
                                    ByVal versionText As String) As String
    Dim rowValue As Long
    Dim qtyValue As Double

    refNumber = ShipmentPersistKeyPart(refNumber)
    itemName = ShipmentPersistKeyPart(itemName)
    rowValue = CLng(Val(rowText))
    qtyValue = NzDbl(qtyText)
    uomText = ShipmentPersistKeyPart(uomText)
    locationText = ShipmentPersistKeyPart(locationText)
    versionText = ShipmentPersistKeyPart(NormalizeBoxBomVersionLabelShipping(versionText))
    If refNumber = "" Or itemName = "" Or rowValue <= 0 Or qtyValue <= 0 Then Exit Function

    ShipmentPersistKey = refNumber & "|" & itemName & "|" & _
                         Format$(qtyValue, "0.############") & "|" & _
                         CStr(rowValue) & "|" & uomText & "|" & locationText & "|" & versionText
End Function

Private Function ShipmentPersistKeyPart(ByVal valueText As String) As String
    valueText = LCase$(Trim$(valueText))
    valueText = Replace$(valueText, "|", " ")
    Do While InStr(1, valueText, "  ", vbBinaryCompare) > 0
        valueText = Replace$(valueText, "  ", " ")
    Loop
    ShipmentPersistKeyPart = valueText
End Function

Private Function HoldRowFieldRaw(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim c As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    c = ColumnIndex(lo, columnName)
    If c = 0 Then Exit Function
    HoldRowFieldRaw = NzStr(lo.DataBodyRange.Cells(rowIndex, c).Value)
End Function

Private Function EscapeHoldField(ByVal valueText As String) As String
    valueText = Replace$(valueText, "\", "\\")
    valueText = Replace$(valueText, vbTab, "\t")
    valueText = Replace$(valueText, vbCr, "\r")
    valueText = Replace$(valueText, vbLf, "\n")
    EscapeHoldField = valueText
End Function

Private Function UnescapeHoldField(ByVal valueText As String) As String
    valueText = Replace$(valueText, "\n", vbLf)
    valueText = Replace$(valueText, "\r", vbCr)
    valueText = Replace$(valueText, "\t", vbTab)
    valueText = Replace$(valueText, "\\", "\")
    UnescapeHoldField = valueText
End Function

Private Function PersistentHoldRowsPath() As String
    Dim rootPath As String
    Dim warehouseId As String

    rootPath = Environ$("LOCALAPPDATA")
    If Trim$(rootPath) = "" Then rootPath = Environ$("TEMP")
    If Trim$(rootPath) = "" Then Exit Function

    warehouseId = Trim$(modConfig.GetWarehouseId())
    If warehouseId = "" Then warehouseId = "default"
    PersistentHoldRowsPath = NormalizeFolderPathShipping(rootPath) & "\invSys\shipping_hold_" & SafeFileTokenShipping(warehouseId) & ".tsv"
End Function

Private Function PersistentActiveShipmentRowsPath() As String
    Dim rootPath As String
    Dim warehouseId As String

    rootPath = Environ$("LOCALAPPDATA")
    If Trim$(rootPath) = "" Then rootPath = Environ$("TEMP")
    If Trim$(rootPath) = "" Then Exit Function

    warehouseId = Trim$(modConfig.GetWarehouseId())
    If warehouseId = "" Then warehouseId = "default"
    PersistentActiveShipmentRowsPath = NormalizeFolderPathShipping(rootPath) & "\invSys\shipping_active_" & SafeFileTokenShipping(warehouseId) & ".tsv"
End Function

Private Function PersistentSentShipmentRowsPath() As String
    PersistentSentShipmentRowsPath = PersistentSentShipmentRowsPathForWarehouse(Trim$(modConfig.GetWarehouseId()))
End Function

Private Function PersistentSentShipmentRowsPathForWarehouse(ByVal warehouseId As String) As String
    Dim rootPath As String

    rootPath = Environ$("LOCALAPPDATA")
    If Trim$(rootPath) = "" Then rootPath = Environ$("TEMP")
    If Trim$(rootPath) = "" Then Exit Function

    warehouseId = Trim$(warehouseId)
    If warehouseId = "" Then warehouseId = "default"
    PersistentSentShipmentRowsPathForWarehouse = NormalizeFolderPathShipping(rootPath) & "\invSys\shipping_sent_" & SafeFileTokenShipping(warehouseId) & ".tsv"
End Function

Private Function PersistentPendingBoxVersionInventoryOverlayPath() As String
    Dim rootPath As String
    Dim warehouseId As String

    rootPath = Environ$("LOCALAPPDATA")
    If Trim$(rootPath) = "" Then rootPath = Environ$("TEMP")
    If Trim$(rootPath) = "" Then Exit Function

    warehouseId = CurrentShippingWarehouseIdForLocalState()
    If Trim$(warehouseId) = "" Then warehouseId = "default"
    PersistentPendingBoxVersionInventoryOverlayPath = NormalizeFolderPathShipping(rootPath) & "\invSys\shipping_projected_" & SafeFileTokenShipping(warehouseId) & ".tsv"
End Function

Private Function ParentFolderPathShipping(ByVal filePath As String) As String
    Dim pos As Long

    pos = InStrRev(filePath, "\")
    If pos > 0 Then ParentFolderPathShipping = Left$(filePath, pos - 1)
End Function

Private Sub EnsureLocalFolderExistsShipping(ByVal folderPath As String)
    On Error Resume Next

    Dim fso As Object
    If Trim$(folderPath) = "" Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    End If
    On Error GoTo 0
End Sub

Private Function SafeFileTokenShipping(ByVal tokenText As String) As String
    Dim i As Long
    Dim ch As String

    tokenText = Trim$(tokenText)
    For i = 1 To Len(tokenText)
        ch = Mid$(tokenText, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then
            SafeFileTokenShipping = SafeFileTokenShipping & ch
        Else
            SafeFileTokenShipping = SafeFileTokenShipping & "_"
        End If
    Next i
    If SafeFileTokenShipping = "" Then SafeFileTokenShipping = "default"
End Function

Private Function ShipmentRowsActionSummary(ByVal loShip As ListObject, _
                                           ByVal rowIndexes As Variant, _
                                           ByVal actionText As String) As String
    Dim cItem As Long
    Dim cQty As Long
    Dim cDesc As Long
    Dim summary As Object
    Dim labels As Object
    Dim i As Long
    Dim rowIndex As Long
    Dim qtyVal As Double
    Dim itemName As String
    Dim versionLabel As String
    Dim key As Variant
    Dim textOut As String

    If loShip Is Nothing Then Exit Function
    If loShip.DataBodyRange Is Nothing Then Exit Function
    If IsEmpty(rowIndexes) Then Exit Function

    cItem = ColumnIndex(loShip, "ITEMS")
    cQty = ColumnIndex(loShip, "QUANTITY")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    If cItem = 0 Or cQty = 0 Then Exit Function

    Set summary = CreateObject("Scripting.Dictionary")
    summary.CompareMode = vbTextCompare
    Set labels = CreateObject("Scripting.Dictionary")
    labels.CompareMode = vbTextCompare

    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex < 1 Or rowIndex > loShip.ListRows.Count Then GoTo NextRow
        itemName = Trim$(NzStr(loShip.DataBodyRange.Cells(rowIndex, cItem).Value))
        qtyVal = NzDbl(loShip.DataBodyRange.Cells(rowIndex, cQty).Value)
        If itemName = "" Or qtyVal <= 0 Then GoTo NextRow
        If cDesc > 0 Then versionLabel = NormalizeBoxBomVersionLabelShipping(NzStr(loShip.DataBodyRange.Cells(rowIndex, cDesc).Value)) Else versionLabel = ""
        key = LCase$(itemName) & "|" & LCase$(versionLabel)
        If summary.Exists(CStr(key)) Then
            summary(CStr(key)) = NzDbl(summary(CStr(key))) + qtyVal
        Else
            summary.Add CStr(key), qtyVal
            labels.Add CStr(key), Trim$(itemName & " " & versionLabel)
        End If
NextRow:
    Next i

    If summary.Count = 0 Then Exit Function
    textOut = "Boxes " & LCase$(Trim$(actionText)) & ":"
    For Each key In summary.Keys
        textOut = textOut & vbCrLf & "- " & Format$(NzDbl(summary(CStr(key))), "0.###") & " " & NzStr(labels(CStr(key)))
    Next key
    ShipmentRowsActionSummary = textOut
End Function

Private Sub ApplyStageVersionInventoryOverlayFromRows(ByVal invLo As ListObject, _
                                                      ByVal loShip As ListObject, _
                                                      ByVal rowIndexes As Variant)
    Dim cRow As Long
    Dim cItem As Long
    Dim cQty As Long
    Dim cDesc As Long
    Dim i As Long
    Dim rowIndex As Long
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim itemName As String
    Dim versionLabel As String
    Dim currentQty As Double

    If loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    cRow = ColumnIndex(loShip, "ROW")
    cItem = ColumnIndex(loShip, "ITEMS")
    cQty = ColumnIndex(loShip, "QUANTITY")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    If cRow = 0 Or cQty = 0 Or cDesc = 0 Then Exit Sub

    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex < 1 Or rowIndex > loShip.ListRows.Count Then GoTo NextRow
        rowVal = NzLng(loShip.DataBodyRange.Cells(rowIndex, cRow).Value)
        qtyVal = NzDbl(loShip.DataBodyRange.Cells(rowIndex, cQty).Value)
        If cItem > 0 Then itemName = Trim$(NzStr(loShip.DataBodyRange.Cells(rowIndex, cItem).Value)) Else itemName = ""
        versionLabel = NormalizeBoxBomVersionLabelShipping(NzStr(loShip.DataBodyRange.Cells(rowIndex, cDesc).Value))
        If rowVal <= 0 Or qtyVal <= 0 Or versionLabel = "" Then GoTo NextRow

        currentQty = StageVersionInventoryCurrentQty(invLo, rowVal, itemName, versionLabel)
        RegisterPendingBoxVersionInventoryOverlay rowVal, versionLabel, currentQty, currentQty + qtyVal
NextRow:
    Next i
End Sub

Private Function StageVersionInventoryCurrentQty(ByVal invLo As ListObject, _
                                                 ByVal rowVal As Long, _
                                                 ByVal itemName As String, _
                                                 ByVal versionLabel As String) As Double
    Dim invIdx As Long
    Dim totalVal As Variant
    Dim versionInv As Object

    If rowVal <= 0 Or NormalizeBoxBomVersionLabelShipping(versionLabel) = "" Then Exit Function

    If Not invLo Is Nothing Then
        invIdx = FindInvRowIndexByRow(invLo, rowVal)
        If invIdx <= 0 And Trim$(itemName) <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
        If invIdx > 0 Then
            totalVal = GetInvSysValueByIndex(invLo, invIdx, "TOTAL INV")
            If Not IsBlankInventoryValue(totalVal) Then
                StageVersionInventoryCurrentQty = NzDbl(totalVal)
                Exit Function
            End If
        End If
    End If

    Set versionInv = BoxMakerFormLoadBoxVersionInventory(rowVal, itemName)
    If Not versionInv Is Nothing Then
        If versionInv.Exists(NormalizeBoxBomVersionLabelShipping(versionLabel)) Then
            StageVersionInventoryCurrentQty = NzDbl(versionInv(NormalizeBoxBomVersionLabelShipping(versionLabel)))
            Exit Function
        End If
    End If

    StageVersionInventoryCurrentQty = SingleVersionFallbackAvailableQty(invLo, rowVal, itemName, versionLabel)
End Function

Private Sub SyncSingleVersionInventoryOverlayFromInvSysRows(ByVal invLo As ListObject, _
                                                            ByVal loShip As ListObject, _
                                                            ByVal rowIndexes As Variant)
    Dim cRow As Long
    Dim cDesc As Long
    Dim cTotal As Long
    Dim i As Long
    Dim rowIndex As Long
    Dim rowVal As Long
    Dim invIdx As Long
    Dim versionLabel As String

    If invLo Is Nothing Or loShip Is Nothing Then Exit Sub
    If invLo.DataBodyRange Is Nothing Or loShip.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    cRow = ColumnIndex(loShip, "ROW")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    cTotal = ColumnIndex(invLo, "TOTAL INV")
    If cRow = 0 Or cDesc = 0 Or cTotal = 0 Then Exit Sub

    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex < 1 Or rowIndex > loShip.ListRows.Count Then GoTo NextRow
        rowVal = NzLng(loShip.DataBodyRange.Cells(rowIndex, cRow).Value)
        versionLabel = NormalizeBoxBomVersionLabelShipping(NzStr(loShip.DataBodyRange.Cells(rowIndex, cDesc).Value))
        If rowVal <= 0 Or versionLabel = "" Then GoTo NextRow
        If CountActiveVersionsForPackageShipping(rowVal) <> 1 Then GoTo NextRow
        invIdx = FindInvRowIndexByRow(invLo, rowVal)
        If invIdx > 0 Then RegisterPendingBoxVersionInventoryOverlay rowVal, versionLabel, NzDbl(invLo.DataBodyRange.Cells(invIdx, cTotal).Value), NzDbl(invLo.DataBodyRange.Cells(invIdx, cTotal).Value)
NextRow:
    Next i
End Sub

Private Sub ApplyShipmentsSentVersionInventoryOverlay(ByVal invLo As ListObject, _
                                                      ByVal loShip As ListObject, _
                                                      ByVal rowIndexes As Variant)
    Dim cRow As Long
    Dim cItem As Long
    Dim cQty As Long
    Dim cDesc As Long
    Dim i As Long
    Dim rowIndex As Long
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim itemName As String
    Dim versionLabel As String
    Dim backendText As String
    Dim projectedText As String
    Dim backendQty As Double
    Dim existingProjectedQty As Double
    Dim projectedQty As Double

    If loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    cRow = ColumnIndex(loShip, "ROW")
    cItem = ColumnIndex(loShip, "ITEMS")
    cQty = ColumnIndex(loShip, "QUANTITY")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    If cRow = 0 Or cQty = 0 Or cDesc = 0 Then Exit Sub

    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex < 1 Or rowIndex > loShip.ListRows.Count Then GoTo NextRow
        rowVal = NzLng(loShip.DataBodyRange.Cells(rowIndex, cRow).Value)
        qtyVal = NzDbl(loShip.DataBodyRange.Cells(rowIndex, cQty).Value)
        If cItem > 0 Then itemName = Trim$(NzStr(loShip.DataBodyRange.Cells(rowIndex, cItem).Value)) Else itemName = ""
        versionLabel = NormalizeBoxBomVersionLabelShipping(NzStr(loShip.DataBodyRange.Cells(rowIndex, cDesc).Value))
        If rowVal <= 0 Or qtyVal <= 0 Or versionLabel = "" Then GoTo NextRow

        backendText = ShipmentVersionInventoryBackendText(invLo, rowVal, itemName, versionLabel)
        projectedText = PendingBoxVersionInventoryOverlayText(rowVal, versionLabel, backendText)
        backendQty = NzDbl(backendText)
        existingProjectedQty = NzDbl(projectedText)
        projectedQty = backendQty - qtyVal
        If projectedQty < 0 Then projectedQty = 0
        If existingProjectedQty < backendQty Then projectedQty = existingProjectedQty
        RegisterPendingBoxVersionInventoryOverlay rowVal, versionLabel, projectedQty, backendQty
NextRow:
    Next i
End Sub

Private Function ShipmentVersionInventoryBackendText(ByVal invLo As ListObject, _
                                                     ByVal rowVal As Long, _
                                                     ByVal itemName As String, _
                                                     ByVal versionLabel As String) As String
    Dim qtyVal As Double
    Dim invIdx As Long
    Dim totalVal As Variant

    qtyVal = PickerVersionAvailableQty(rowVal, itemName, versionLabel)
    If qtyVal > 0.0000001 Then
        ShipmentVersionInventoryBackendText = CStr(qtyVal)
        Exit Function
    End If

    If CountActiveVersionsForPackageShipping(rowVal) <= 1 Then
        If Not invLo Is Nothing Then
            invIdx = FindInvRowIndexByRow(invLo, rowVal)
            If invIdx <= 0 And Trim$(itemName) <> "" Then invIdx = FindInvRowIndexByItem(invLo, itemName)
            If invIdx > 0 Then
                totalVal = GetInvSysValueByIndex(invLo, invIdx, "TOTAL INV")
                If Not IsBlankInventoryValue(totalVal) Then
                    ShipmentVersionInventoryBackendText = CStr(NzDbl(totalVal))
                    Exit Function
                End If
            End If
        End If
    End If

    qtyVal = SingleVersionFallbackAvailableQty(invLo, rowVal, itemName, versionLabel)
    ShipmentVersionInventoryBackendText = CStr(qtyVal)
End Function

Public Function ShipmentsFormRunToShipmentsRows(ByVal rowIndexes As Variant, _
                                                ByVal carrierValue As String, _
                                                ByRef report As String) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim loShip As ListObject
    Dim errNotes As String
    Dim deltas As Collection
    Dim shipLogs As New Collection
    Dim stagedTotal As Double
    Dim reserveEventId As String
    Dim previousVisibility As XlSheetVisibility
    Dim visibilityChanged As Boolean
    Dim previousEvents As Boolean
    Dim previousHandling As Boolean
    Dim mutationStarted As Boolean
    Dim unreservedRows As Variant
    Dim reservedRows As Variant
    Dim alreadyReservedCount As Long

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    Set invLo = GetWritableShippingInvSysTable(ws, report)
    Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    If invLo Is Nothing Then
        If report = "" Then report = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If loShip Is Nothing Then
        report = "Shipment table not found."
        Exit Function
    End If

    unreservedRows = ShipmentRowsByReserveState(loShip, rowIndexes, False)
    reservedRows = ShipmentRowsByReserveState(loShip, rowIndexes, True)
    alreadyReservedCount = ShipmentRowIndexCount(reservedRows)

    If IsEmpty(unreservedRows) Then
        If alreadyReservedCount > 0 Then
            BeginShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
            mutationStarted = True
            SetShipmentRowsArea loShip, rowIndexes, "Shipments", carrierValue
            PersistActiveShipmentRowsLocal loShip
            report = CStr(alreadyReservedCount) & " selected row(s) were already locked for shipment."
            If Trim$(carrierValue) <> "" Then report = report & vbCrLf & "Carrier: " & Trim$(carrierValue)
            ShipmentsFormRunToShipmentsRows = True
            GoTo CleanExit
        End If
        report = "No selected Warehouse rows are ready for To Shipments."
        Exit Function
    End If

    Set deltas = BuildSelectedShipmentRowsDeltas(invLo, loShip, unreservedRows, "Warehouse", errNotes)
    If deltas Is Nothing Then
        If errNotes = "" Then errNotes = "No selected Warehouse rows are ready for To Shipments."
        report = errNotes
        Exit Function
    End If
    If deltas.Count = 0 Then
        If errNotes = "" Then errNotes = "No selected Warehouse rows are ready for To Shipments."
        report = errNotes
        Exit Function
    End If

    If Not QueueShipmentsReserveEvent(deltas, errNotes, reserveEventId) Then
        If errNotes = "" Then errNotes = "Unable to queue shipment reserve event."
        report = errNotes
        Exit Function
    End If

    BeginShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
    mutationStarted = True
    PrepareShipmentStageLogEntries invLo, deltas, shipLogs
    stagedTotal = ApplyShipmentDeltasLocal(invLo, deltas, errNotes)
    If stagedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to stage selected shipment rows."
        report = errNotes
        GoTo CleanExit
    End If
    SetShipmentRowsArea loShip, unreservedRows, "Shipments", carrierValue
    If alreadyReservedCount > 0 Then SetShipmentRowsArea loShip, reservedRows, "Shipments", carrierValue
    If reserveEventId <> "" Then
        SetShipmentRowsReserveEventId loShip, unreservedRows, reserveEventId
        If Not ForEachShipmentReservationUpsert(loShip, unreservedRows, reserveEventId, report) Then GoTo CleanExit
    End If
    ApplyStageVersionInventoryOverlayFromRows invLo, loShip, unreservedRows
    InvalidateAggregates True
    PersistActiveShipmentRowsLocal loShip
    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs

    report = "Moved " & Format$(stagedTotal, "0.###") & " package(s) to Shipments."
    If alreadyReservedCount > 0 Then report = report & vbCrLf & CStr(alreadyReservedCount) & " selected row(s) were already locked."
    If Trim$(carrierValue) <> "" Then report = report & vbCrLf & "Carrier: " & Trim$(carrierValue)
    If reserveEventId <> "" Then report = report & vbCrLf & "Reserve EventID: " & reserveEventId
    ShipmentsFormRunToShipmentsRows = True

CleanExit:
    If mutationStarted Then EndShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
    Exit Function

Fail:
    report = "To Shipments failed: " & Err.Description
    Resume CleanExit
End Function

Public Function ShipmentsFormRunShipmentsSentRows(ByVal rowIndexes As Variant, _
                                                  ByVal carrierValue As String, _
                                                  ByRef report As String, _
                                                  Optional ByVal skipAuthForTest As Boolean = False) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim loShip As ListObject
    Dim deltas As Collection
    Dim errNotes As String
    Dim queuedEventId As String
    Dim shipLogs As New Collection
    Dim shippedTotal As Double
    Dim previousVisibility As XlSheetVisibility
    Dim visibilityChanged As Boolean
    Dim previousEvents As Boolean
    Dim previousHandling As Boolean
    Dim mutationStarted As Boolean
    Dim sentSummary As String
    Dim unreservedRows As Variant
    Dim unreservedDeltas As Collection

    If Not skipAuthForTest Then
        If Not modRoleUiAccess.CanCurrentUserPerformCapability("SHIP_POST", "", "", "", errNotes) Then
            report = errNotes
            Exit Function
        End If
    End If

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    Set invLo = GetWritableShippingInvSysTable(ws, report)
    Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    If invLo Is Nothing Then
        If report = "" Then report = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If loShip Is Nothing Then
        report = "Shipment table not found."
        Exit Function
    End If

    Set deltas = BuildSelectedShipmentRowsDeltas(invLo, loShip, rowIndexes, "Shipments", errNotes)
    If deltas Is Nothing Then
        If errNotes = "" Then errNotes = "Select row(s) in Shipments area. Use To Shipments first for Warehouse rows."
        report = errNotes
        Exit Function
    End If
    If deltas.Count = 0 Then
        If errNotes = "" Then errNotes = "Select row(s) in Shipments area. Use To Shipments first for Warehouse rows."
        report = errNotes
        Exit Function
    End If
    sentSummary = ShipmentRowsActionSummary(loShip, rowIndexes, "Sent")
    unreservedRows = ShipmentRowsByReserveState(loShip, rowIndexes, False)
    If Not IsEmpty(unreservedRows) Then
        Set unreservedDeltas = BuildSelectedShipmentRowsDeltas(invLo, loShip, unreservedRows, "Shipments", errNotes)
        If unreservedDeltas Is Nothing Then
            If errNotes = "" Then errNotes = "Unable to build unreserved shipment event."
            report = errNotes
            Exit Function
        End If
        If Not QueueShipmentsSentEvent(unreservedDeltas, errNotes, queuedEventId) Then
            If errNotes = "" Then errNotes = "Unable to queue shipment event."
            report = errNotes
            Exit Function
        End If
    End If

    BeginShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
    mutationStarted = True
    If Trim$(carrierValue) <> "" Then SetShipmentRowsCarrier loShip, rowIndexes, carrierValue
    shippedTotal = ApplyShipmentsSentRowsInventory(invLo, loShip, rowIndexes, shipLogs, errNotes)
    If shippedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to finalize selected shipment rows."
        report = errNotes
        GoTo CleanExit
    End If
    ApplyShipmentsSentVersionInventoryOverlay invLo, loShip, rowIndexes
    AppendSentShipmentRowsLocal loShip, rowIndexes
    If Not MarkShippingReservationRows(loShip, rowIndexes, SHIP_RESERVATION_COMPLETED, vbNullString, report) Then GoTo CleanExit
    DeleteShipmentRows loShip, rowIndexes
    InvalidateAggregates True
    PersistActiveShipmentRowsLocal loShip
    ClearInstructionStaging ws
    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs

    report = "Shipments sent: " & Format$(shippedTotal, "0.###") & " package(s)."
    If sentSummary <> "" Then report = report & vbCrLf & sentSummary
    If Trim$(carrierValue) <> "" Then report = report & vbCrLf & "Carrier: " & Trim$(carrierValue)
    If queuedEventId <> "" Then report = report & vbCrLf & "Inbox EventID: " & queuedEventId
    If queuedEventId = "" Then report = report & vbCrLf & _
        "Server inventory was reserved at To Shipments; Shipments Sent completed the reservation and is waiting for processor/log catch-up."
    If errNotes <> "" Then report = report & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
    ShipmentsFormRunShipmentsSentRows = True

CleanExit:
    If mutationStarted Then EndShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
    Exit Function

Fail:
    report = "Shipments Sent failed: " & Err.Description
    Resume CleanExit
End Function

Public Function ShipmentsFormRunShipmentsSentRowsReportForTest(ByVal rowIndexes As Variant, _
                                                               ByVal carrierValue As String) As String
    Dim report As String

    If ShipmentsFormRunShipmentsSentRows(rowIndexes, carrierValue, report, True) Then
        ShipmentsFormRunShipmentsSentRowsReportForTest = "OK|" & report
    Else
        ShipmentsFormRunShipmentsSentRowsReportForTest = "FAIL|" & report
    End If
End Function

Public Function ValidateApplyShipmentsSentRowsInventoryFromCurrentWorkbook(ByVal rowIndexes As Variant, ByRef report As String) As Boolean
    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim loShip As ListObject
    Dim shipLogs As New Collection
    Dim errNotes As String
    Dim shippedTotal As Double

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    Set invLo = GetInvSysTable()
    Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If loShip Is Nothing Then
        report = "Shipment table not found."
        Exit Function
    End If

    shippedTotal = ApplyShipmentsSentRowsInventory(invLo, loShip, rowIndexes, shipLogs, errNotes)
    If shippedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to finalize shipment rows."
        report = errNotes
        Exit Function
    End If
    report = "Validated Shipments Sent inventory application for " & Format$(shippedTotal, "0.###") & " package(s)."
    ValidateApplyShipmentsSentRowsInventoryFromCurrentWorkbook = True
End Function

Public Function ValidateCompleteShipmentsSentRowsFromCurrentWorkbook(ByVal rowIndexes As Variant, ByRef report As String) As Boolean
    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim loShip As ListObject
    Dim shipLogs As New Collection
    Dim errNotes As String
    Dim shippedTotal As Double
    Dim previousVisibility As XlSheetVisibility
    Dim visibilityChanged As Boolean
    Dim previousEvents As Boolean
    Dim previousHandling As Boolean
    Dim mutationStarted As Boolean

    On Error GoTo Fail
    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    Set invLo = GetInvSysTable()
    Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If loShip Is Nothing Then
        report = "Shipment table not found."
        Exit Function
    End If

    BeginShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
    mutationStarted = True
    shippedTotal = ApplyShipmentsSentRowsInventory(invLo, loShip, rowIndexes, shipLogs, errNotes)
    If shippedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to finalize shipment rows."
        report = errNotes
        GoTo CleanExit
    End If
    ApplyShipmentsSentVersionInventoryOverlay invLo, loShip, rowIndexes
    If Not MarkShippingReservationRows(loShip, rowIndexes, SHIP_RESERVATION_COMPLETED, vbNullString, report) Then GoTo CleanExit
    DeleteShipmentRows loShip, rowIndexes
    InvalidateAggregates True
    PersistActiveShipmentRowsLocal loShip

    report = "Validated Shipments Sent completion for " & Format$(shippedTotal, "0.###") & " package(s)."
    ValidateCompleteShipmentsSentRowsFromCurrentWorkbook = True

CleanExit:
    If mutationStarted Then EndShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
    Exit Function

Fail:
    report = "Validate Shipments Sent completion failed: " & Err.Description
    Resume CleanExit
End Function

Private Function ApplyShipmentsSentRowsInventory(ByVal invLo As ListObject, _
                                                 ByVal loShip As ListObject, _
                                                 ByVal rowIndexes As Variant, _
                                                 ByVal shipLogs As Collection, _
                                                 ByRef errNotes As String) As Double
    Dim reservedRows As Variant
    Dim unreservedRows As Variant
    Dim reservedDeltas As Collection
    Dim unreservedDeltas As Collection
    Dim appliedTotal As Double

    ApplyShipmentsSentRowsInventory = -1
    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If loShip Is Nothing Then
        errNotes = "Shipment table not found."
        Exit Function
    End If
    If IsEmpty(rowIndexes) Then
        errNotes = "Select shipment row(s) first."
        Exit Function
    End If

    reservedRows = ShipmentRowsByReserveState(loShip, rowIndexes, True)
    unreservedRows = ShipmentRowsByReserveState(loShip, rowIndexes, False)

    If Not IsEmpty(reservedRows) Then
        Set reservedDeltas = BuildSelectedShipmentRowsDeltas(invLo, loShip, reservedRows, "Shipments", errNotes)
        If reservedDeltas Is Nothing Then
            If errNotes = "" Then errNotes = "Unable to build reserved shipment finalization."
            Exit Function
        End If
    End If
    If Not IsEmpty(unreservedRows) Then
        Set unreservedDeltas = BuildSelectedShipmentRowsDeltas(invLo, loShip, unreservedRows, "Shipments", errNotes)
        If unreservedDeltas Is Nothing Then
            If errNotes = "" Then errNotes = "Unable to build unreserved shipment finalization."
            Exit Function
        End If
    End If
    If reservedDeltas Is Nothing And unreservedDeltas Is Nothing Then
        errNotes = "No selected Shipments rows were found."
        Exit Function
    End If

    ApplyShipmentsSentRowsInventory = 0
    If Not reservedDeltas Is Nothing Then
        PrepareShipmentsSentLogEntries invLo, reservedDeltas, shipLogs, False
        appliedTotal = ApplyShipmentsSentDeltas(invLo, reservedDeltas, errNotes, False)
        If appliedTotal < 0 Then
            If errNotes = "" Then errNotes = "Unable to finalize reserved shipment rows."
            ApplyShipmentsSentRowsInventory = -1
            Exit Function
        End If
        ApplyShipmentsSentRowsInventory = ApplyShipmentsSentRowsInventory + appliedTotal
    End If
    If Not unreservedDeltas Is Nothing Then
        PrepareShipmentsSentLogEntries invLo, unreservedDeltas, shipLogs, True
        appliedTotal = ApplyShipmentsSentDeltas(invLo, unreservedDeltas, errNotes, True)
        If appliedTotal < 0 Then
            If errNotes = "" Then errNotes = "Unable to finalize unreserved shipment rows."
            ApplyShipmentsSentRowsInventory = -1
            Exit Function
        End If
        ApplyShipmentsSentRowsInventory = ApplyShipmentsSentRowsInventory + appliedTotal
    End If
End Function

Private Sub BeginShippingTableMutation(ByVal lo As ListObject, _
                                       ByRef previousVisibility As XlSheetVisibility, _
                                       ByRef visibilityChanged As Boolean, _
                                       ByRef previousEvents As Boolean, _
                                       ByRef previousHandling As Boolean)
    previousEvents = Application.EnableEvents
    previousHandling = mHandlingShippingSheetChange
    Application.EnableEvents = False
    mHandlingShippingSheetChange = True
    If lo Is Nothing Then Exit Sub

    previousVisibility = lo.Parent.Visible
    If previousVisibility <> xlSheetVisible Then
        lo.Parent.Visible = xlSheetVisible
        visibilityChanged = True
    End If
    EnsureShippingWorksheetEditable lo.Parent
End Sub

Private Sub EndShippingTableMutation(ByVal lo As ListObject, _
                                     ByVal previousVisibility As XlSheetVisibility, _
                                     ByVal visibilityChanged As Boolean, _
                                     ByVal previousEvents As Boolean, _
                                     ByVal previousHandling As Boolean)
    On Error Resume Next
    If visibilityChanged Then
        If Not lo Is Nothing Then
            If Not Application.ActiveSheet Is Nothing Then
                If Application.ActiveSheet Is lo.Parent Then
                    lo.Parent.Parent.Worksheets(SHEET_SHIPMENTS).Activate
                End If
            End If
            lo.Parent.Visible = previousVisibility
        End If
    End If
    mHandlingShippingSheetChange = previousHandling
    Application.EnableEvents = previousEvents
    On Error GoTo 0
End Sub

Public Function ShipmentsFormUseExistingInventory() As Boolean
    Dim ws As Worksheet

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    ShipmentsFormUseExistingInventory = UseExistingInventoryEnabled(ws)
End Function

Public Sub ShipmentsFormSetUseExistingInventory(ByVal enabled As Boolean)
    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set shp = ws.Shapes(CHK_USE_EXISTING)
    If Not shp Is Nothing Then shp.ControlFormat.Value = IIf(enabled, 1, xlOff)
    On Error GoTo 0
    InvalidateAggregates
End Sub

Public Function ShipmentsFormLoadReadiness() As Variant
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim loPack As ListObject
    Dim loBom As ListObject
    Dim countRows As Long
    Dim rows As Variant
    Dim outRow As Long

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    RebuildShippingAggregates

    Set invLo = GetInvSysTable()
    Set loPack = GetListObject(ws, TABLE_AGG_PACK)
    Set loBom = GetListObject(ws, TABLE_AGG_BOM)

    countRows = CountReadableAggregateRows(loPack) + CountReadableAggregateRows(loBom)
    If countRows = 0 Then Exit Function
    ReDim rows(1 To countRows, 1 To 9)

    outRow = 0
    AppendReadinessRows rows, outRow, "Package", loPack, invLo, "SHIPMENTS"
    AppendReadinessRows rows, outRow, "Component", loBom, invLo, "USED"

    ShipmentsFormLoadReadiness = rows
    Exit Function

FailSoft:
End Function

Public Function ShipmentsFormRunToShipments(ByRef report As String) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim aggPack As ListObject
    Dim errNotes As String
    Dim deltas As Collection
    Dim shipLogs As New Collection
    Dim stagedTotal As Double

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    RebuildShippingAggregates
    Set invLo = GetInvSysTable()
    Set aggPack = GetListObject(ws, TABLE_AGG_PACK)
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If aggPack Is Nothing Then
        report = "AggregatePackages table not found."
        Exit Function
    End If
    If aggPack.DataBodyRange Is Nothing Then
        report = "AggregatePackages has no rows to stage."
        Exit Function
    End If

    Set deltas = BuildShipmentDeltaPacket(invLo, aggPack, errNotes)
    If deltas Is Nothing Then
        If errNotes = "" Then errNotes = "No additional shipments required; Shipments column already meets demand."
        report = errNotes
        Exit Function
    End If
    If deltas.Count = 0 Then
        If errNotes = "" Then errNotes = "No additional shipments required; Shipments column already meets demand."
        report = errNotes
        Exit Function
    End If

    PrepareShipmentStageLogEntries invLo, deltas, shipLogs
    stagedTotal = ApplyShipmentDeltasLocal(invLo, deltas, errNotes)
    If stagedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to stage shipments due to inventory shortage."
        report = errNotes
        Exit Function
    End If

    InvalidateAggregates True
    RestoreShipmentStageColumns invLo, deltas
    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs

    report = "Staged " & Format$(stagedTotal, "0.###") & " packages into invSys.SHIPMENTS."
    ShipmentsFormRunToShipments = True
    Exit Function

Fail:
    report = "To Shipments failed: " & Err.Description
End Function

Public Function ShipmentsFormRunShipmentsSent(ByRef report As String) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim deltas As Collection
    Dim errNotes As String
    Dim queuedEventId As String
    Dim runtimeReport As String
    Dim shipLogs As New Collection
    Dim shippedTotal As Double

    If Not modRoleUiAccess.CanCurrentUserPerformCapability("SHIP_POST", "", "", "", errNotes) Then
        report = errNotes
        Exit Function
    End If

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    If Not BuildQueueableShipmentsSentDeltas(invLo, ws, deltas, errNotes) Then
        If errNotes = "" Then errNotes = "Unable to build shipment event."
        report = errNotes
        Exit Function
    End If
    If Not QueueShipmentsSentEvent(deltas, errNotes, queuedEventId) Then
        If errNotes = "" Then errNotes = "Unable to queue shipment event."
        report = errNotes
        Exit Function
    End If

    PrepareShipmentsSentLogEntries invLo, deltas, shipLogs, SHIPMENTS_SENT_DEDUCTS_TOTALINV
    shippedTotal = ApplyShipmentsSentDeltas(invLo, deltas, errNotes, SHIPMENTS_SENT_DEDUCTS_TOTALINV)
    If shippedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to finalize shipments."
        report = errNotes
        Exit Function
    End If

    ClearShipmentEntryTables
    InvalidateAggregates True
    ClearInstructionStaging ws
    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs
    If Not RunShippingRuntimeQueueRefresh(ws.Parent, ResolveCurrentShippingWarehouseId(), runtimeReport) Then
        If runtimeReport = "" Then runtimeReport = "Local shipment post succeeded, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If
    ClearShipmentStageAfterRefresh ws.Parent, deltas

    report = "Finalized " & Format$(shippedTotal, "0.###") & " shipments."
    If queuedEventId <> "" Then report = report & vbCrLf & "Inbox EventID: " & queuedEventId
    If errNotes <> "" Then report = report & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
    ShipmentsFormRunShipmentsSent = True
    Exit Function

Fail:
    report = "Shipments Sent failed: " & Err.Description
End Function

Public Function ShipmentsFormRunDirectShipmentsSent(ByRef report As String, Optional ByVal displayedShipmentRows As Variant) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Dim invLo As ListObject
    Dim loShip As ListObject
    Dim deltas As Collection
    Dim errNotes As String
    Dim queuedEventId As String
    Dim runtimeReport As String
    Dim shipLogs As New Collection
    Dim shippedTotal As Double

    If Not modRoleUiAccess.CanCurrentUserPerformCapability("SHIP_POST", "", "", "", errNotes) Then
        report = errNotes
        Exit Function
    End If

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If
    Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    If loShip Is Nothing Then
        report = "Shipment table not found."
        Exit Function
    End If

    If Not IsMissing(displayedShipmentRows) Then
        If Not IsEmpty(displayedShipmentRows) Then Set deltas = BuildDisplayedShipmentRowsDeltas(invLo, displayedShipmentRows, errNotes)
    End If
    If deltas Is Nothing Then Set deltas = BuildShipmentLineDeltas(invLo, loShip, errNotes)
    If deltas Is Nothing Then
        If errNotes = "" Then errNotes = "No shipment rows are ready to send."
        report = errNotes
        Exit Function
    End If
    If deltas.Count = 0 Then
        If errNotes = "" Then errNotes = "No shipment rows are ready to send."
        report = errNotes
        Exit Function
    End If

    If Not QueueShipmentsSentEvent(deltas, errNotes, queuedEventId) Then
        If errNotes = "" Then errNotes = "Unable to queue shipment event."
        report = errNotes
        Exit Function
    End If

    PrepareShipmentsSentLogEntries invLo, deltas, shipLogs, True
    shippedTotal = ApplyDirectShipmentsSentDeltas(invLo, deltas, errNotes)
    If shippedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to finalize shipments."
        report = errNotes
        Exit Function
    End If

    ClearShipmentEntryTables
    InvalidateAggregates True
    ClearInstructionStaging ws
    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs
    If Not RunShippingRuntimeQueueRefresh(ws.Parent, ResolveCurrentShippingWarehouseId(), runtimeReport) Then
        If runtimeReport = "" Then runtimeReport = "Local shipment post succeeded, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If

    report = "Finalized " & Format$(shippedTotal, "0.###") & " shipments."
    If queuedEventId <> "" Then report = report & vbCrLf & "Inbox EventID: " & queuedEventId
    If errNotes <> "" Then report = report & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
    ShipmentsFormRunDirectShipmentsSent = True
    Exit Function

Fail:
    report = "Shipments Sent failed: " & Err.Description
End Function

Public Function ShipmentsFormRunStageAndShipmentsSent(ByRef report As String) As Boolean
    On Error GoTo Fail

    Dim stageReport As String
    Dim sendReport As String
    Dim stageOk As Boolean
    Dim sendOk As Boolean

    stageOk = ShipmentsFormRunToShipments(stageReport)
    If Not stageOk Then
        If InStr(1, stageReport, "No additional shipments required", vbTextCompare) = 0 Then
            report = stageReport
            Exit Function
        End If
    End If

    sendOk = ShipmentsFormRunShipmentsSent(sendReport)
    report = Trim$(stageReport)
    If Trim$(sendReport) <> "" Then
        If report <> "" Then report = report & vbCrLf
        report = report & sendReport
    End If
    ShipmentsFormRunStageAndShipmentsSent = sendOk
    Exit Function

Fail:
    report = "Shipments Sent failed: " & Err.Description
End Function

Private Function CountReadableAggregateRows(ByVal lo As ListObject) As Long
    Dim cQty As Long
    Dim cRow As Long
    Dim arr As Variant
    Dim r As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cQty = ColumnIndex(lo, "QUANTITY")
    cRow = ColumnIndex(lo, "ROW")
    If cQty = 0 Or cRow = 0 Then Exit Function
    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        If NzLng(arr(r, cRow)) > 0 And NzDbl(arr(r, cQty)) > 0 Then CountReadableAggregateRows = CountReadableAggregateRows + 1
    Next r
End Function

Private Sub AppendReadinessRows(ByRef rows As Variant, _
                                ByRef outRow As Long, _
                                ByVal kindText As String, _
                                ByVal lo As ListObject, _
                                ByVal invLo As ListObject, _
                                ByVal stagedColumnName As String)
    Dim cQty As Long
    Dim cRow As Long
    Dim cItem As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim arr As Variant
    Dim r As Long
    Dim rowValue As Long
    Dim requiredQty As Double
    Dim invRow As ListRow
    Dim currentInv As Double
    Dim stagedQty As Double
    Dim itemName As String
    Dim uomValue As String
    Dim locationValue As String

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    cQty = ColumnIndex(lo, "QUANTITY")
    cRow = ColumnIndex(lo, "ROW")
    cItem = ColumnIndex(lo, "ITEM")
    cUom = ColumnIndex(lo, "UOM")
    cLoc = ColumnIndex(lo, "LOCATION")
    If cQty = 0 Or cRow = 0 Then Exit Sub

    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        rowValue = NzLng(arr(r, cRow))
        requiredQty = NzDbl(arr(r, cQty))
        If rowValue <= 0 Or requiredQty <= 0 Then GoTo NextRow

        itemName = ""
        uomValue = ""
        locationValue = ""
        currentInv = 0
        stagedQty = 0
        If cItem > 0 Then itemName = NzStr(arr(r, cItem))
        If cUom > 0 Then uomValue = NzStr(arr(r, cUom))
        If cLoc > 0 Then locationValue = NzStr(arr(r, cLoc))

        If Not invLo Is Nothing Then
            Set invRow = FindInvListRowByRowValue(invLo, rowValue)
            If Not invRow Is Nothing Then
                If itemName = "" Then itemName = NzStr(GetInvSysValueFromRow(invRow, "ITEM"))
                If uomValue = "" Then uomValue = NzStr(GetInvSysValueFromRow(invRow, "UOM"))
                If locationValue = "" Then locationValue = NzStr(GetInvSysValueFromRow(invRow, "LOCATION"))
                currentInv = NzDbl(GetInvSysValueFromRow(invRow, "TOTAL INV"))
                stagedQty = NzDbl(GetInvSysValueFromRow(invRow, stagedColumnName))
            End If
        End If

        outRow = outRow + 1
        rows(outRow, 1) = kindText
        rows(outRow, 2) = itemName
        rows(outRow, 3) = requiredQty
        rows(outRow, 4) = currentInv
        rows(outRow, 5) = stagedQty
        rows(outRow, 6) = uomValue
        rows(outRow, 7) = locationValue
        rows(outRow, 8) = rowValue
        If currentInv + stagedQty + 0.0000001 >= requiredQty Then
            rows(outRow, 9) = "OK"
        Else
            rows(outRow, 9) = "Short " & Format$(requiredQty - currentInv - stagedQty, "0.###")
        End If
NextRow:
        Set invRow = Nothing
    Next r
End Sub

Private Function GetInvSysValueFromRow(ByVal invRow As ListRow, ByVal columnName As String) As Variant
    Dim idx As Long

    If invRow Is Nothing Then Exit Function
    idx = ColumnIndex(invRow.Parent, columnName)
    If idx = 0 Then Exit Function
    GetInvSysValueFromRow = invRow.Range.Cells(1, idx).Value
End Function

Private Function FormatBoxMakerQuantityText(ByVal qtyValue As Double) As String
    If Abs(qtyValue - Fix(qtyValue)) < 0.0000001 Then
        FormatBoxMakerQuantityText = Format$(qtyValue, "0")
    Else
        FormatBoxMakerQuantityText = Format$(qtyValue, "0.###")
    End If
End Function

Public Sub CommitBoxBuilderFormState(ByVal boxName As String, _
                                     ByVal boxUom As String, _
                                     ByVal boxLocation As String, _
                                     ByVal boxDescription As String, _
                                     ByVal bomRows As Variant, _
                                     Optional ByVal saveAction As String = "", _
                                     Optional ByVal versionLabel As String = "", _
                                     Optional ByVal statusText As String = "")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim loBuilder As ListObject
    Dim loBom As ListObject
    Dim prevEvents As Boolean
    Dim quietStarted As Boolean
    Dim rowCount As Long
    Dim r As Long
    Dim errText As String

    prevEvents = Application.EnableEvents
    boxName = Trim$(boxName)
    boxUom = Trim$(boxUom)
    If boxName = "" Then
        MsgBox "Enter a Box Name before saving.", vbExclamation
        Exit Sub
    End If
    If boxUom = "" Then
        MsgBox "Box UOM is required.", vbExclamation
        Exit Sub
    End If
    If IsEmpty(bomRows) Then
        MsgBox "Add at least one component before saving.", vbExclamation
        Exit Sub
    End If

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Set wb = ws.Parent

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBuilder Is Nothing Or loBom Is Nothing Then
        MsgBox "Box Builder tables not found on ShipmentsTally sheet.", vbExclamation
        Exit Sub
    End If

    rowCount = UBound(bomRows, 1)
    If rowCount <= 0 Then
        MsgBox "Add at least one component before saving.", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    mHandlingShippingSheetChange = True
    mSuppressGeneratedIdentityEditGuard = True
    modUiQuiet.BeginQuietUi wb
    quietStarted = True

    EnsureTableHasRow loBuilder
    RemoveColumnIfExistsShipping loBuilder, "ROW"
    ClearListObjectData loBuilder
    SetTableCellShipping loBuilder, 1, "Box Name", boxName
    SetTableCellShipping loBuilder, 1, "UOM", boxUom
    SetTableCellShipping loBuilder, 1, "LOCATION", boxLocation
    SetTableCellShipping loBuilder, 1, "DESCRIPTION", boxDescription

    EnsureBoxBomEntryColumns loBom
    ClearListObjectData loBom
    EnsureListObjectHasRowsShipping loBom, MaxLongShipping(rowCount, 10)

    For r = 1 To rowCount
        SetTableCellShipping loBom, r, "Version", BoxBuilderFormBomText(bomRows, r, 1, "v1")
        SetTableCellShipping loBom, r, COL_BOXBOM_ITEM, BoxBuilderFormBomText(bomRows, r, 2, "")
        SetTableCellShipping loBom, r, "ITEM_CODE", BoxBuilderFormBomText(bomRows, r, 3, "")
        SetTableCellShipping loBom, r, "ROW", BoxBuilderFormBomLong(bomRows, r, 4)
        SetTableCellShipping loBom, r, "QUANTITY", BoxBuilderFormBomDouble(bomRows, r, 5)
        SetTableCellShipping loBom, r, "UOM", BoxBuilderFormBomText(bomRows, r, 6, "")
        SetTableCellShipping loBom, r, "LOCATION", BoxBuilderFormBomText(bomRows, r, 7, "")
        SetTableCellShipping loBom, r, "DESCRIPTION", BoxBuilderFormBomText(bomRows, r, 8, "")
    Next r
    FillBlankBoxBomVersionShipping loBom
    SortBoxBomByVersionShipping loBom

CleanRestore:
    If quietStarted Then
        On Error Resume Next
        modUiQuiet.EndQuietUi
        On Error GoTo ErrHandler
    End If
    mSuppressGeneratedIdentityEditGuard = False
    mHandlingShippingSheetChange = False
    Application.EnableEvents = prevEvents

    If Err.Number = 0 Then
        If Trim$(saveAction) = "" Then
            BtnSaveBox
        Else
            SaveBoxBuilderFormTablesExplicit saveAction, versionLabel, statusText
        End If
    End If
    Exit Sub

ErrHandler:
    errText = Err.Description
    Resume CleanFail

CleanFail:
    If quietStarted Then
        On Error Resume Next
        modUiQuiet.EndQuietUi
        On Error GoTo 0
    End If
    mSuppressGeneratedIdentityEditGuard = False
    mHandlingShippingSheetChange = False
    Application.EnableEvents = prevEvents
    MsgBox "BOX_BUILDER_SAVE failed: " & errText, vbCritical
End Sub

Private Sub SaveBoxBuilderFormTablesExplicit(ByVal saveAction As String, ByVal versionLabel As String, ByVal statusText As String)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim loMeta As ListObject
    Dim loBom As ListObject
    Dim invLo As ListObject
    Dim components As Collection
    Dim syncNotes As String
    Dim boxName As String
    Dim boxUOM As String
    Dim boxLoc As String
    Dim boxDesc As String
    Dim boxRowValue As Long
    Dim replaceVersion As Long
    Dim forceNewVersion As Boolean
    Dim bomReport As String
    Dim statusReport As String
    Dim finalMsg As String
    Dim savedBomVersion As Long
    Dim desiredActive As Boolean

    saveAction = UCase$(Trim$(saveAction))
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then versionLabel = "v1"

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Set loMeta = GetListObject(ws, TABLE_BOX_BUILDER)
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loMeta Is Nothing Or loBom Is Nothing Then
        MsgBox "Box Builder tables not found on ShipmentsTally sheet.", vbExclamation
        Exit Sub
    End If

    EnsureTableHasRow loMeta
    RemoveColumnIfExistsShipping loMeta, "ROW"
    EnsureBoxBomEntryColumns loBom
    FillBlankBoxBomVersionShipping loBom

    boxName = Trim$(NzStr(ValueFromTable(loMeta, "Box Name")))
    boxUOM = Trim$(NzStr(ValueFromTable(loMeta, "UOM")))
    boxLoc = Trim$(NzStr(ValueFromTable(loMeta, "LOCATION")))
    boxDesc = Trim$(NzStr(ValueFromTable(loMeta, "DESCRIPTION")))
    If boxName = "" Then
        MsgBox "Enter a Box Name before saving.", vbExclamation
        Exit Sub
    End If
    If boxUOM = "" Then
        MsgBox "Box Builder UOM is required.", vbExclamation
        Exit Sub
    End If

    Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    If saveAction = "UPDATE" Then
        replaceVersion = BomVersionNumberFromLabel(versionLabel)
    ElseIf saveAction = "NEW" Then
        forceNewVersion = True
    Else
        MsgBox "Unknown BoxBuilder save action: " & saveAction, vbCritical
        Exit Sub
    End If

    Set components = CollectBomComponents(loBom, invLo, syncNotes, versionLabel)
    If components.Count = 0 Then
        MsgBox "Add at least one valid component row (ROW/QUANTITY) to the BoxBOM table.", vbExclamation
        Exit Sub
    End If
    If components.Count > SHIPPING_BOM_DATA_ROWS Then
        MsgBox "BOM exceeds the 50-row limit. Remove extra rows and try again.", vbExclamation
        Exit Sub
    End If

    boxRowValue = ResolveBoxPackageRowValue(ws.Parent, boxName, invLo)
    If boxRowValue = 0 Then Exit Sub
    boxRowValue = EnsureInvSysItem(boxName, boxUOM, boxLoc, boxDesc, invLo, boxRowValue)
    If boxRowValue = 0 Then Exit Sub

    If Not SaveShippingBomToRuntime(ws.Parent, boxRowValue, boxName, boxUOM, boxLoc, boxDesc, components, bomReport, replaceVersion, forceNewVersion) Then
        If bomReport = "" Then bomReport = "Unable to save Shipping BOM to the selected warehouse runtime."
        MsgBox bomReport, vbCritical
        Exit Sub
    End If

    If replaceVersion > 0 Then
        savedBomVersion = replaceVersion
    Else
        savedBomVersion = mLastSavedShippingBomVersion
    End If
    If savedBomVersion <= 0 Then savedBomVersion = BomVersionNumberFromLabel(versionLabel)
    If savedBomVersion <= 0 Then savedBomVersion = 1

    desiredActive = BoxBuilderStatusIsActive(statusText)
    If Not SetShippingBomVersionStatusInRuntime(ws.Parent, boxRowValue, savedBomVersion, desiredActive, statusReport) Then
        If statusReport = "" Then statusReport = "Unable to update Shipping BOM status."
        MsgBox statusReport, vbCritical
        Exit Sub
    End If

    RefreshShippingBomViewForWorkbook ws.Parent, bomReport
    RefreshBoxBomVersionList ws, boxRowValue
    RebuildBoxBomVersionListFromDisplayedBom ws, boxRowValue
    InvalidateAggregates True

    finalMsg = bomReport
    If Len(statusReport) > 0 Then finalMsg = finalMsg & vbCrLf & statusReport
    If Len(syncNotes) > 0 Then finalMsg = finalMsg & vbCrLf & syncNotes
    MsgBox finalMsg, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BOX_BUILDER_EXPLICIT_SAVE failed: " & Err.Description, vbCritical
End Sub

Public Sub BoxBuilderFormDeleteVersion(ByVal packageRow As Long, ByVal versionLabel As String)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim report As String

    If Not modRoleUiAccess.RequireCurrentUserCapability("ADMIN_MAINT") Then Exit Sub
    If packageRow <= 0 Then
        MsgBox "Select a saved box before deleting a version.", vbExclamation
        Exit Sub
    End If
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then
        MsgBox "Select a version before deleting.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Delete Shipping BOM ROW " & CStr(packageRow) & " " & versionLabel & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    If Not DeleteShippingBomVersionFromRuntime(ws.Parent, packageRow, BomVersionNumberFromLabel(versionLabel), report) Then
        If report = "" Then report = "Could not delete selected Shipping BOM version."
        MsgBox report, vbCritical
        Exit Sub
    End If

    RefreshShippingBomViewForWorkbook ws.Parent, report
    DeleteLocalShippingBomViewRowsForVersion ws, packageRow, versionLabel
    RefreshBoxBomVersionList ws, packageRow
    MsgBox report, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BOX_BUILDER_DELETE_VERSION failed: " & Err.Description, vbCritical
End Sub

Public Sub BoxBuilderFormDeleteBox(ByVal packageRow As Long)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim report As String

    If Not modRoleUiAccess.RequireCurrentUserCapability("ADMIN_MAINT") Then Exit Sub
    If packageRow <= 0 Then
        MsgBox "Select a saved box before deleting.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Delete all saved BOM versions for Shipping BOM ROW " & CStr(packageRow) & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    If Not DeleteShippingBomPackageFromRuntime(ws.Parent, packageRow, report) Then
        If report = "" Then report = "Could not delete selected Shipping BOM box."
        MsgBox report, vbCritical
        Exit Sub
    End If

    RefreshShippingBomViewForWorkbook ws.Parent, report
    MsgBox report, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BOX_BUILDER_DELETE_BOX failed: " & Err.Description, vbCritical
End Sub

Public Function BoxBuilderFormArchiveBox(ByVal packageRow As Long, Optional ByRef report As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim archiveReport As String
    Dim refreshReport As String

    If Not modRoleUiAccess.RequireCurrentUserCapability("ADMIN_MAINT") Then
        report = "ADMIN_MAINT is required to archive box designs."
        Exit Function
    End If
    If packageRow <= 0 Then
        report = "Select a saved box before archiving."
        Exit Function
    End If

    Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then
        report = "ShipmentsTally sheet was not found."
        Exit Function
    End If
    If Not ArchiveShippingBomPackageInRuntime(ws.Parent, packageRow, archiveReport) Then
        If archiveReport = "" Then archiveReport = "Could not archive selected Shipping BOM box."
        report = archiveReport
        Exit Function
    End If

    RefreshShippingBomViewForWorkbook ws.Parent, refreshReport
    RefreshBoxBomVersionList ws, packageRow
    InvalidateAggregates True
    report = archiveReport
    If report = "" Then report = "Archived Shipping BOM ROW " & CStr(packageRow) & "."
    If Trim$(refreshReport) <> "" Then AppendNote report, refreshReport
    BoxBuilderFormArchiveBox = True
    Exit Function

ErrHandler:
    report = "BOX_BUILDER_ARCHIVE_BOX failed: " & Err.Description
End Function

Private Function BoxBuilderFormBomText(ByVal rowsData As Variant, _
                                       ByVal rowIndex As Long, _
                                       ByVal colIndex As Long, _
                                       ByVal defaultValue As String) As String
    On Error GoTo UseDefault
    BoxBuilderFormBomText = NzStr(rowsData(rowIndex, colIndex))
    If BoxBuilderFormBomText = "" Then BoxBuilderFormBomText = defaultValue
    Exit Function
UseDefault:
    BoxBuilderFormBomText = defaultValue
End Function

Private Function BoxBuilderFormBomLong(ByVal rowsData As Variant, _
                                       ByVal rowIndex As Long, _
                                       ByVal colIndex As Long) As Long
    On Error GoTo UseZero
    BoxBuilderFormBomLong = NzLng(rowsData(rowIndex, colIndex))
    Exit Function
UseZero:
    BoxBuilderFormBomLong = 0
End Function

Private Function BoxBuilderFormBomDouble(ByVal rowsData As Variant, _
                                         ByVal rowIndex As Long, _
                                         ByVal colIndex As Long) As Double
    On Error GoTo UseZero
    BoxBuilderFormBomDouble = NzDbl(rowsData(rowIndex, colIndex))
    Exit Function
UseZero:
    BoxBuilderFormBomDouble = 0#
End Function

Private Function ShippingInventoryPickerTableHasRows(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If ColumnIndex(lo, "ROW") = 0 Then Exit Function
    If ColumnIndex(lo, "ITEM") = 0 Then Exit Function
    ShippingInventoryPickerTableHasRows = True
End Function

Private Function BuildShippingInventoryPickerItems(ByVal lo As ListObject) As Variant
    Dim cRow As Long
    Dim cCode As Long
    Dim cItem As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long
    Dim cTotalInv As Long
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim itemName As String

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    cRow = ColumnIndex(lo, "ROW")
    cCode = ColumnIndex(lo, "ITEM_CODE")
    cItem = ColumnIndex(lo, "ITEM")
    cUom = ColumnIndex(lo, "UOM")
    cLoc = ColumnIndex(lo, "LOCATION")
    cDesc = ColumnIndex(lo, "DESCRIPTION")
    cTotalInv = ColumnIndex(lo, "TOTAL INV")
    If cRow = 0 Or cItem = 0 Then Exit Function

    src = lo.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 7)
    For r = 1 To UBound(src, 1)
        itemName = Trim$(NzStr(src(r, cItem)))
        If itemName <> "" Then
            outRow = outRow + 1
            result(outRow, 1) = NzStr(src(r, cRow))
            If cCode > 0 Then
                result(outRow, 2) = NzStr(src(r, cCode))
            Else
                result(outRow, 2) = itemName
            End If
            result(outRow, 3) = itemName
            If cUom > 0 Then result(outRow, 4) = NzStr(src(r, cUom))
            If cLoc > 0 Then result(outRow, 5) = NzStr(src(r, cLoc))
            If cDesc > 0 Then result(outRow, 6) = NzStr(src(r, cDesc))
            If cTotalInv > 0 Then result(outRow, 7) = NzDbl(src(r, cTotalInv)) Else result(outRow, 7) = ""
        End If
    Next r

    If outRow = 0 Then Exit Function
    If outRow = UBound(src, 1) Then
        BuildShippingInventoryPickerItems = result
        Exit Function
    End If

    ReDim trimmed(1 To outRow, 1 To 7)
    For r = 1 To outRow
        For c = 1 To 7
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    BuildShippingInventoryPickerItems = trimmed
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
    Dim boxMakerMode As Boolean
    Dim packageQty As Double

    If targetCell Is Nothing Then
        Set ws = SheetExists(SHEET_SHIPMENTS)
    Else
        Set ws = targetCell.Worksheet
    End If
    If ws Is Nothing Then Exit Sub
    boxMakerMode = IsBoxMakerMode(ws)

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then Exit Sub
    NormalizeBoxBuilderTable loBuilder
    If boxMakerMode Then EnsureColumnExists loBuilder, "Quantity", "Box Name"
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

    ClearSelectedBoxBomVersionForWorksheet ws
    ClearBoxBomVersionListForWorksheet ws
    ClearDisplayedBoxBomForWorksheet ws

    Application.EnableEvents = False
    WriteValue loBuilder.ListRows(1), "Box Name", actualItem
    WriteValue loBuilder.ListRows(1), "UOM", actualUom
    WriteValue loBuilder.ListRows(1), "LOCATION", actualLoc
    WriteValue loBuilder.ListRows(1), "DESCRIPTION", actualDesc
    Application.EnableEvents = True

    If actualRow > 0 Then
        If boxMakerMode Then packageQty = NzDbl(ValueFromTable(loBuilder, "Quantity")) Else packageQty = 1#
        If LoadBoxMakerBomForPackage(ws, actualRow, packageQty, report) Then
            ShowShippingStatus report
        ElseIf report <> "" Then
            ShowShippingStatus report
        End If
    Else
        ClearBoxBomVersionListForWorksheet ws
    End If
    SelectBoxBuilderDataCellForRepeatHeaderPicker loBuilder
    RefreshBoxMakerCurrentInventory ws
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    MsgBox "ApplyItemToBoxBuilder error: " & Err.Description, vbCritical
End Sub

Private Sub SelectBoxBuilderDataCellForRepeatHeaderPicker(ByVal loBuilder As ListObject)
    On Error Resume Next
    If loBuilder Is Nothing Then Exit Sub
    If loBuilder.DataBodyRange Is Nothing Then Exit Sub
    loBuilder.DataBodyRange.Cells(1, 1).Select
    On Error GoTo 0
End Sub

Public Sub ShippingSelectBoxBuilderDataCellForRepeatHeaderPicker(ByVal loBuilder As ListObject)
    SelectBoxBuilderDataCellForRepeatHeaderPicker loBuilder
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

    ClearBoxBomVersionListForWorksheet ws
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

    ClearBoxBomVersionListForWorksheet ws
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
    Dim cPackageItem As Long
    Dim cVersion As Long
    Dim cVersionLabel As Long
    Dim cComponentRow As Long
    Dim cComponentItemCode As Long
    Dim cComponentItem As Long
    Dim cComponentQty As Long
    Dim cComponentUom As Long
    Dim cComponentLocation As Long
    Dim cComponentDescription As Long
    Dim cActive As Long
    Dim scaleQty As Double
    Dim arr As Variant
    Dim r As Long
    Dim outRow As Long
    Dim loView As ListObject
    Dim refreshReport As String
    Dim preservedCurrentInv As Object
    Dim componentName As String
    Dim componentCode As String
    Dim componentRow As Long
    Dim componentUom As String
    Dim componentLocation As String
    Dim componentDescription As String
    Dim versionLabel As String
    Dim repairedRows As Long

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
    ClearBoxBomVersionListForWorksheet ws
    ClearSelectedBoxBomVersionForWorksheet ws

    If Not TryLoadRuntimeShippingBomRows(arr, _
                                         cPackageRow, _
                                         cVersion, _
                                         cVersionLabel, _
                                         cComponentRow, _
                                         cComponentItemCode, _
                                         cComponentItem, _
                                         cComponentQty, _
                                         cComponentUom, _
                                         cComponentLocation, _
                                         cComponentDescription, _
                                         cActive, _
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
        cVersion = ColumnIndex(loView, "BomVersion")
        cVersionLabel = ColumnIndex(loView, "BomVersionLabel")
        cComponentRow = ColumnIndex(loView, "ComponentRow")
        cComponentItemCode = ColumnIndex(loView, "ComponentItemCode")
        cComponentItem = ColumnIndex(loView, "ComponentItem")
        cComponentQty = ColumnIndex(loView, "ComponentQty")
        cComponentUom = ColumnIndex(loView, "ComponentUOM")
        cComponentLocation = ColumnIndex(loView, "ComponentLocation")
        cComponentDescription = ColumnIndex(loView, "ComponentDescription")
        cActive = ColumnIndex(loView, "IsActive")
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
        If cActive > 0 Then
            If Not ShippingBomActiveValue(arr(r, cActive)) Then GoTo NextBomRow
        End If

        outRow = outRow + 1
        Do While loBom.ListRows.Count < outRow
            loBom.ListRows.Add
        Loop

        componentName = ValueFromArrayColumn(arr, r, cComponentItem)
        componentCode = ""
        If cComponentItemCode > 0 Then componentCode = ValueFromArrayColumn(arr, r, cComponentItemCode)
        componentRow = NzLng(arr(r, cComponentRow))
        componentUom = ValueFromArrayColumn(arr, r, cComponentUom)
        componentLocation = ValueFromArrayColumn(arr, r, cComponentLocation)
        componentDescription = ValueFromArrayColumn(arr, r, cComponentDescription)
        versionLabel = VersionLabelShipping(arr, r, cVersion, cVersionLabel)
        If componentRow <= 0 Then
            If ResolveCanonicalComponentInfoShipping(componentName, componentCode, componentRow, componentName, componentCode, componentUom, componentLocation, componentDescription) Then
                repairedRows = repairedRows + 1
            End If
        End If

        SetTableCellShipping loBom, outRow, "Version", versionLabel
        SetTableCellShipping loBom, outRow, COL_BOXBOM_ITEM, componentName
        SetTableCellShipping loBom, outRow, "ITEM_CODE", componentCode
        SetTableCellShipping loBom, outRow, "ROW", componentRow
        SetTableCellShipping loBom, outRow, "QUANTITY", NzDbl(arr(r, cComponentQty)) * scaleQty
        SetTableCellShipping loBom, outRow, "UOM", componentUom
        SetTableCellShipping loBom, outRow, "LOCATION", componentLocation
        SetTableCellShipping loBom, outRow, "DESCRIPTION", componentDescription
NextBomRow:
    Next r

    EnsureBoxBomStarterRows loBom
    FillBlankBoxBomVersionShipping loBom
    SortBoxBomByVersionShipping loBom
    If outRow = 0 Then
        report = "No saved BoxBOM components were found for invSys ROW " & CStr(packageRow) & "."
        Exit Function
    End If

    LoadBoxMakerBomForPackage = True
    report = "Loaded BoxBOM for invSys ROW " & CStr(packageRow) & " (" & CStr(outRow) & " component row(s))."
    If repairedRows > 0 Then report = report & " Repaired " & CStr(repairedRows) & " component ROW value(s) from inventory."
    RefreshBoxMakerCurrentInventory ws
    RestorePreservedBoxBomCurrentInventory loBom, preservedCurrentInv
    RefreshShippingBomViewForWorkbook ws.Parent, refreshReport
    RefreshBoxBomVersionList ws, packageRow
    If Not BoxBomVersionListHasRows(ws) Then RebuildBoxBomVersionListFromDisplayedBom ws, packageRow
    RefreshBoxMakerBomVersionDisplay ws, packageRow
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

Private Sub RefreshBoxBomVersionList(ByVal ws As Worksheet, ByVal packageRow As Long)
    On Error GoTo CleanExit

    Dim loView As ListObject
    Dim loBom As ListObject
    Dim loVersions As ListObject
    Dim versionRows As Variant
    Dim versionCount As Long

    If ws Is Nothing Then Exit Sub
    If packageRow <= 0 Then Exit Sub
    Set loVersions = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If Not loView Is Nothing Then
        If Not loView.DataBodyRange Is Nothing Then
            versionRows = BuildBoxBomVersionRows(loView, packageRow, versionCount)
        End If
    End If
    AugmentBoxBomVersionRowsFromLocalBom ws, versionRows, versionCount

    If versionCount = 0 Then
        If Not loVersions Is Nothing Then DeleteAllListObjectRowsShipping loVersions
        Exit Sub
    End If

    Set loVersions = EnsureBoxBomVersionsTable(ws, loBom)
    If loVersions Is Nothing Then Exit Sub
    WriteBoxBomVersionRowsToTable loVersions, versionRows
    ApplyBoxBomVersionStatusValidation loVersions
    ArrangeBoxBuilderBandShipping GetListObject(ws, TABLE_BOX_BUILDER), loBom

CleanExit:
End Sub

Private Sub RefreshBoxBomVersionListFromSourceTable(ByVal ws As Worksheet, _
                                                    ByVal loSource As ListObject, _
                                                    ByVal packageRow As Long)
    On Error GoTo CleanExit

    Dim loBom As ListObject
    Dim loVersions As ListObject
    Dim versionRows As Variant
    Dim versionCount As Long

    If ws Is Nothing Then Exit Sub
    If loSource Is Nothing Then Exit Sub
    If packageRow <= 0 Then Exit Sub

    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    Set loVersions = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    versionRows = BuildBoxBomVersionRows(loSource, packageRow, versionCount)
    AugmentBoxBomVersionRowsFromLocalBom ws, versionRows, versionCount

    If versionCount = 0 Then
        If Not loVersions Is Nothing Then DeleteAllListObjectRowsShipping loVersions
        Exit Sub
    End If

    Set loVersions = EnsureBoxBomVersionsTable(ws, loBom)
    If loVersions Is Nothing Then Exit Sub
    WriteBoxBomVersionRowsToTable loVersions, versionRows
    ApplyBoxBomVersionStatusValidation loVersions

CleanExit:
End Sub

Private Sub RefreshBoxBomVersionListForCurrentBuilder(ByVal ws As Worksheet)
    On Error GoTo CleanExit

    Dim packageRow As Long
    Dim runtimeMax As Long
    Dim boxName As String
    Dim loBuilder As ListObject
    Dim report As String

    If ws Is Nothing Then Exit Sub
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then
        ClearSelectedBoxBomVersionForWorksheet ws
        ClearBoxBomVersionListForWorksheet ws
        ClearDisplayedBoxBomForWorksheet ws
        Exit Sub
    End If
    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    If boxName = "" Then
        ClearSelectedBoxBomVersionForWorksheet ws
        ClearBoxBomVersionListForWorksheet ws
        ClearDisplayedBoxBomForWorksheet ws
        Exit Sub
    End If

    packageRow = FindShippingBomPackageRowByName(ws.Parent, boxName, runtimeMax)
    If packageRow <= 0 Then
        ClearSelectedBoxBomVersionForWorksheet ws
        ClearBoxBomVersionListForWorksheet ws
        ClearDisplayedBoxBomForWorksheet ws
        Exit Sub
    End If

    ClearSelectedBoxBomVersionForWorksheet ws
    ClearBoxBomVersionListForWorksheet ws
    ClearDisplayedBoxBomForWorksheet ws
    If LoadBoxMakerBomForPackage(ws, packageRow, 1#, report) Then
        ShowShippingStatus report
    ElseIf report <> "" Then
        ShowShippingStatus report
    End If

CleanExit:
End Sub

Private Sub ClearSelectedBoxBomVersionForWorksheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If StrComp(mSelectedBoxBomVersionWorkbookName, ws.Parent.Name, vbTextCompare) <> 0 Then Exit Sub
    If StrComp(mSelectedBoxBomVersionWorksheetName, ws.Name, vbTextCompare) <> 0 Then Exit Sub

    mSelectedBoxBomVersionLabel = ""
    mSelectedBoxBomVersionPackageRow = 0
    mSelectedBoxBomVersionWorkbookName = ""
    mSelectedBoxBomVersionWorksheetName = ""
End Sub

Private Sub ClearBoxBomVersionListForWorksheet(ByVal ws As Worksheet)
    Dim loVersions As ListObject

    If ws Is Nothing Then Exit Sub
    ClearBoxBomVersionSignatureCache ws
    Set loVersions = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    If Not loVersions Is Nothing Then DeleteAllListObjectRowsShipping loVersions
End Sub

Private Sub ClearDisplayedBoxBomForWorksheet(ByVal ws As Worksheet)
    Dim loBom As ListObject

    If ws Is Nothing Then Exit Sub
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Sub
    ClearListObjectData loBom
    EnsureBoxBomStarterRows loBom
End Sub

Private Sub RebuildBoxBomVersionListFromDisplayedBom(ByVal ws As Worksheet, Optional ByVal packageRow As Long = 0)
    On Error GoTo CleanExit

    Dim loBom As ListObject
    Dim loVersions As ListObject
    Dim cVersion As Long
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim versionLabel As String
    Dim seen As Object
    Dim rowsOut() As Variant
    Dim key As Variant
    Dim boxName As String

    If ws Is Nothing Then Exit Sub
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then
        ClearBoxBomVersionListForWorksheet ws
        Exit Sub
    End If
    If loBom.DataBodyRange Is Nothing Then
        ClearBoxBomVersionListForWorksheet ws
        Exit Sub
    End If
    cVersion = ColumnIndex(loBom, "Version")
    If cVersion = 0 Then
        ClearBoxBomVersionListForWorksheet ws
        Exit Sub
    End If

    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To loBom.ListRows.Count
        If Not BoxBomRowHasComponentDataShipping(loBom, r) Then GoTo NextBomRow
        versionLabel = NormalizeBoxBomVersionLabelShipping(loBom.DataBodyRange.Cells(r, cVersion).Value)
        If versionLabel = "" Then
            versionLabel = "v1"
            loBom.DataBodyRange.Cells(r, cVersion).Value = versionLabel
        End If
        If versionLabel <> "" Then seen(versionLabel) = True
NextBomRow:
    Next r

    Set loVersions = EnsureBoxBomVersionsTable(ws, loBom)
    If loVersions Is Nothing Then Exit Sub
    If seen.Count = 0 Then
        DeleteAllListObjectRowsShipping loVersions
        Exit Sub
    End If

    boxName = CurrentBoxBuilderName(ws)
    ReDim rowsOut(1 To seen.Count, 1 To 8)
    For Each key In seen.Keys
        outRow = outRow + 1
        rowsOut(outRow, 1) = CStr(key)
        rowsOut(outRow, 2) = "Active"
        For c = 3 To 7
            rowsOut(outRow, c) = vbNullString
        Next c
        rowsOut(outRow, 8) = boxName
    Next key
    If packageRow > 0 Then FillBoxBomVersionRowsMetadataFromView ws, packageRow, rowsOut

    If BoxBomVersionSignatureMatches(ws, boxName, rowsOut) Then GoTo CleanExit
    WriteDisplayedBoxBomVersionRowsDirect loVersions, rowsOut
    ApplyBoxBomVersionStatusValidation loVersions
    RememberBoxBomVersionSignature ws, boxName, rowsOut
    ArrangeBoxBuilderBandShipping GetListObject(ws, TABLE_BOX_BUILDER), loBom

CleanExit:
End Sub

Private Function BoxBomVersionSignatureMatches(ByVal ws As Worksheet, _
                                               ByVal boxName As String, _
                                               ByRef rowsOut As Variant) As Boolean
    If ws Is Nothing Then Exit Function
    If StrComp(mLastBoxBomVersionWorkbookName, ws.Parent.Name, vbTextCompare) <> 0 Then Exit Function
    If StrComp(mLastBoxBomVersionWorksheetName, ws.Name, vbTextCompare) <> 0 Then Exit Function
    BoxBomVersionSignatureMatches = (StrComp(mLastBoxBomVersionSignature, BoxBomVersionSignature(boxName, rowsOut), vbBinaryCompare) = 0)
End Function

Private Sub RememberBoxBomVersionSignature(ByVal ws As Worksheet, _
                                           ByVal boxName As String, _
                                           ByRef rowsOut As Variant)
    If ws Is Nothing Then Exit Sub
    mLastBoxBomVersionWorkbookName = ws.Parent.Name
    mLastBoxBomVersionWorksheetName = ws.Name
    mLastBoxBomVersionSignature = BoxBomVersionSignature(boxName, rowsOut)
End Sub

Private Sub ClearBoxBomVersionSignatureCache(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If StrComp(mLastBoxBomVersionWorkbookName, ws.Parent.Name, vbTextCompare) <> 0 Then Exit Sub
    If StrComp(mLastBoxBomVersionWorksheetName, ws.Name, vbTextCompare) <> 0 Then Exit Sub
    mLastBoxBomVersionSignature = ""
    mLastBoxBomVersionWorkbookName = ""
    mLastBoxBomVersionWorksheetName = ""
End Sub

Private Function BoxBomVersionSignature(ByVal boxName As String, ByRef rowsOut As Variant) As String
    Dim r As Long
    Dim c As Long
    Dim parts As String

    On Error GoTo CleanExit
    parts = LCase$(Trim$(boxName))
    For r = LBound(rowsOut, 1) To UBound(rowsOut, 1)
        parts = parts & "|"
        For c = LBound(rowsOut, 2) To UBound(rowsOut, 2)
            parts = parts & CStr(rowsOut(r, c)) & Chr$(30)
        Next c
    Next r
CleanExit:
    BoxBomVersionSignature = parts
End Function

Private Sub WriteDisplayedBoxBomVersionRowsDirect(ByVal loVersions As ListObject, ByRef rowsData As Variant)
    On Error GoTo CleanFail

    Dim rowsNeeded As Long
    Dim r As Long
    Dim c As Long
    Dim headers As Variant
    Dim colIndex As Long

    If loVersions Is Nothing Then Exit Sub
    headers = BoxBomVersionHeadersShipping()
    For c = LBound(headers) To UBound(headers)
        EnsureColumnExists loVersions, CStr(headers(c))
    Next c
    HideBoxBomVersionIdentityColumns loVersions
    ClearListObjectFiltersShipping loVersions

    rowsNeeded = UBound(rowsData, 1)
    Do While loVersions.ListRows.Count < rowsNeeded
        loVersions.ListRows.Add
    Loop
    Do While loVersions.ListRows.Count > rowsNeeded
        loVersions.ListRows(loVersions.ListRows.Count).Delete
    Loop
    If loVersions.DataBodyRange Is Nothing Then Exit Sub
    loVersions.DataBodyRange.ClearContents

    For r = 1 To rowsNeeded
        For c = LBound(headers) To UBound(headers)
            colIndex = ColumnIndex(loVersions, CStr(headers(c)))
            If colIndex > 0 Then loVersions.DataBodyRange.Cells(r, colIndex).Value = rowsData(r, c - LBound(headers) + 1)
        Next c
    Next r
    Exit Sub

CleanFail:
End Sub

Private Function BoxBomVersionListHasRows(ByVal ws As Worksheet) As Boolean
    On Error GoTo CleanFail

    Dim loVersions As ListObject
    Dim cVersion As Long
    Dim r As Long

    If ws Is Nothing Then Exit Function
    Set loVersions = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    If loVersions Is Nothing Then Exit Function
    If loVersions.DataBodyRange Is Nothing Then Exit Function
    cVersion = ColumnIndex(loVersions, "Version")
    If cVersion = 0 Then Exit Function

    For r = 1 To loVersions.ListRows.Count
        If Trim$(NzStr(loVersions.DataBodyRange.Cells(r, cVersion).Value)) <> "" Then
            BoxBomVersionListHasRows = True
            Exit Function
        End If
    Next r
    Exit Function

CleanFail:
    BoxBomVersionListHasRows = False
End Function

Private Sub WriteBoxBomVersionRowsToTable(ByVal loVersions As ListObject, ByRef rowsData As Variant)
    On Error GoTo CleanExit

    Dim rowsNeeded As Long
    Dim r As Long
    Dim headers As Variant
    Dim c As Long
    Dim colIndex As Long

    If loVersions Is Nothing Then Exit Sub
    ClearListObjectFiltersShipping loVersions

    headers = BoxBomVersionHeadersShipping()
    For c = LBound(headers) To UBound(headers)
        EnsureColumnExists loVersions, CStr(headers(c))
    Next c

    If IsEmpty(rowsData) Then
        ResizeListObjectDataRowsShipping loVersions, 0
        Exit Sub
    End If
    rowsNeeded = UBound(rowsData, 1)
    If rowsNeeded <= 0 Then
        ResizeListObjectDataRowsShipping loVersions, 0
        Exit Sub
    End If

    MoveBoxBomAwayFromVersionResize loVersions, rowsNeeded
    ResizeListObjectDataRowsShipping loVersions, rowsNeeded
    If loVersions.DataBodyRange Is Nothing Then Exit Sub
    loVersions.DataBodyRange.ClearContents

    For r = 1 To rowsNeeded
        For c = LBound(headers) To UBound(headers)
            colIndex = ColumnIndex(loVersions, CStr(headers(c)))
            If colIndex > 0 Then loVersions.DataBodyRange.Cells(r, colIndex).Value = rowsData(r, c - LBound(headers) + 1)
        Next c
    Next r
    HideBoxBomVersionIdentityColumns loVersions

CleanExit:
End Sub

Private Sub MoveBoxBomAwayFromVersionResize(ByVal loVersions As ListObject, ByVal rowsNeeded As Long)
    On Error GoTo CleanExit

    Dim ws As Worksheet
    Dim loBom As ListObject
    Dim targetRow As Long

    If loVersions Is Nothing Then Exit Sub
    Set ws = loVersions.Parent
    If ws Is Nothing Then Exit Sub
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Sub

    targetRow = loVersions.Range.Row + rowsNeeded + SHIP_LAYOUT_GAP_ROWS + 8
    If targetRow < 1 Then Exit Sub
    MoveListObjectToRowColShipping loBom, targetRow, loVersions.Range.Column

CleanExit:
End Sub

Private Sub ResizeListObjectDataRowsShipping(ByVal lo As ListObject, ByVal rowsNeeded As Long)
    On Error GoTo FallbackDelete

    Dim firstCell As Range
    Dim newRange As Range
    Dim oldBottomRow As Long
    Dim oldFirstCol As Long
    Dim oldRightCol As Long

    If lo Is Nothing Then Exit Sub
    If rowsNeeded < 0 Then rowsNeeded = 0
    ClearListObjectFiltersShipping lo

    Set firstCell = lo.HeaderRowRange.Cells(1, 1)
    oldBottomRow = lo.Range.Row + lo.Range.Rows.Count - 1
    oldFirstCol = lo.Range.Column
    oldRightCol = lo.Range.Column + lo.Range.Columns.Count - 1

    If rowsNeeded = 0 Then
        Set newRange = lo.Parent.Range(firstCell, firstCell.Offset(1, lo.Range.Columns.Count - 1))
        lo.Resize newRange
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
        DeleteAllListObjectRowsShipping lo
        ClearListObjectOldFootprintShipping lo, oldBottomRow, oldFirstCol, oldRightCol
        Exit Sub
    End If

    Set newRange = lo.Parent.Range(firstCell, firstCell.Offset(rowsNeeded, lo.Range.Columns.Count - 1))
    lo.Resize newRange
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    ClearListObjectOldFootprintShipping lo, oldBottomRow, oldFirstCol, oldRightCol
    Exit Sub

FallbackDelete:
    Err.Clear
    On Error Resume Next
    Do While lo.ListRows.Count < rowsNeeded
        lo.ListRows.Add
    Loop
    Do While lo.ListRows.Count > rowsNeeded
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    ClearListObjectOldFootprintShipping lo, oldBottomRow, oldFirstCol, oldRightCol
    On Error GoTo 0
End Sub

Private Sub ClearListObjectOldFootprintShipping(ByVal lo As ListObject, _
                                                ByVal oldBottomRow As Long, _
                                                ByVal oldFirstCol As Long, _
                                                ByVal oldRightCol As Long)
    On Error Resume Next

    Dim newBottomRow As Long
    Dim clearTopRow As Long
    Dim clearRange As Range

    If lo Is Nothing Then Exit Sub
    If oldBottomRow <= 0 Or oldFirstCol <= 0 Or oldRightCol < oldFirstCol Then Exit Sub

    newBottomRow = lo.Range.Row + lo.Range.Rows.Count - 1
    clearTopRow = newBottomRow + 1
    If clearTopRow > oldBottomRow Then Exit Sub

    Set clearRange = lo.Parent.Range(lo.Parent.Cells(clearTopRow, oldFirstCol), _
                                     lo.Parent.Cells(oldBottomRow, oldRightCol))
    clearRange.Clear
    On Error GoTo 0
End Sub

Private Sub FillBoxBomVersionRowsMetadataFromView(ByVal ws As Worksheet, ByVal packageRow As Long, ByRef rowsData As Variant)
    On Error GoTo CleanExit

    Dim loView As ListObject
    Dim cPackageRow As Long
    Dim cPackageItem As Long
    Dim cVersion As Long
    Dim cLabel As Long
    Dim cActive As Long
    Dim cEffectiveFrom As Long
    Dim cEffectiveTo As Long
    Dim cRetiredAt As Long
    Dim cUpdatedAt As Long
    Dim cUpdatedBy As Long
    Dim r As Long
    Dim outRow As Long
    Dim rowLabel As String
    Dim wantedLabel As String
    Dim rowVersion As Long

    If ws Is Nothing Then Exit Sub
    If packageRow <= 0 Then Exit Sub
    If IsEmpty(rowsData) Then Exit Sub
    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If loView Is Nothing Then Exit Sub
    If loView.DataBodyRange Is Nothing Then Exit Sub

    cPackageRow = ColumnIndex(loView, "PackageRow")
    cPackageItem = ColumnIndex(loView, "PackageItem")
    cVersion = ColumnIndex(loView, "BomVersion")
    cLabel = ColumnIndex(loView, "BomVersionLabel")
    cActive = ColumnIndex(loView, "IsActive")
    cEffectiveFrom = ColumnIndex(loView, "EffectiveFromUTC")
    cEffectiveTo = ColumnIndex(loView, "EffectiveToUTC")
    cRetiredAt = ColumnIndex(loView, "RetiredAtUTC")
    cUpdatedAt = ColumnIndex(loView, "UpdatedAtUTC")
    cUpdatedBy = ColumnIndex(loView, "UpdatedBy")
    If cPackageRow = 0 Then Exit Sub

    For outRow = 1 To UBound(rowsData, 1)
        wantedLabel = NormalizeBoxBomVersionLabelShipping(rowsData(outRow, 1))
        If wantedLabel = "" Then GoTo NextOutRow

        For r = 1 To loView.ListRows.Count
            If NzLng(loView.DataBodyRange.Cells(r, cPackageRow).Value) <> packageRow Then GoTo NextViewRow
            rowLabel = ""
            If cLabel > 0 Then rowLabel = NormalizeBoxBomVersionLabelShipping(loView.DataBodyRange.Cells(r, cLabel).Value)
            If rowLabel = "" And cVersion > 0 Then
                rowVersion = NzLng(loView.DataBodyRange.Cells(r, cVersion).Value)
                If rowVersion > 0 Then rowLabel = "v" & CStr(rowVersion)
            End If
            If StrComp(rowLabel, wantedLabel, vbTextCompare) <> 0 Then GoTo NextViewRow

            If cActive > 0 And Not ShippingBomActiveValue(loView.DataBodyRange.Cells(r, cActive).Value) Then
                rowsData(outRow, 2) = "Retired"
            Else
                rowsData(outRow, 2) = "Active"
            End If
            If cEffectiveFrom > 0 Then rowsData(outRow, 3) = loView.DataBodyRange.Cells(r, cEffectiveFrom).Value
            If cEffectiveTo > 0 Then rowsData(outRow, 4) = loView.DataBodyRange.Cells(r, cEffectiveTo).Value
            If cRetiredAt > 0 Then rowsData(outRow, 5) = loView.DataBodyRange.Cells(r, cRetiredAt).Value
            If cUpdatedAt > 0 Then rowsData(outRow, 6) = loView.DataBodyRange.Cells(r, cUpdatedAt).Value
            If cUpdatedBy > 0 Then rowsData(outRow, 7) = loView.DataBodyRange.Cells(r, cUpdatedBy).Value
            If cPackageItem > 0 Then
                rowsData(outRow, 8) = loView.DataBodyRange.Cells(r, cPackageItem).Value
            Else
                rowsData(outRow, 8) = CurrentBoxBuilderName(ws)
            End If
            Exit For
NextViewRow:
        Next r
NextOutRow:
    Next outRow

CleanExit:
End Sub

Private Sub AugmentBoxBomVersionRowsFromLocalBom(ByVal ws As Worksheet, _
                                                 ByRef versionRows As Variant, _
                                                 ByRef versionCount As Long)
    On Error GoTo CleanExit

    Dim loBom As ListObject
    Dim cVersion As Long
    Dim r As Long
    Dim c As Long
    Dim versionLabel As String
    Dim seen As Object
    Dim newRows() As Variant
    Dim oldCount As Long
    Dim copyRow As Long

    If ws Is Nothing Then Exit Sub
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Sub
    If loBom.DataBodyRange Is Nothing Then Exit Sub
    cVersion = ColumnIndex(loBom, "Version")
    If cVersion = 0 Then Exit Sub

    Set seen = CreateObject("Scripting.Dictionary")
    If versionCount > 0 Then
        For r = 1 To versionCount
            versionLabel = NormalizeBoxBomVersionLabelShipping(versionRows(r, 1))
            If versionLabel <> "" Then seen(versionLabel) = True
        Next r
    End If

    For r = 1 To loBom.ListRows.Count
        If Not BoxBomRowHasComponentDataShipping(loBom, r) Then GoTo NextBomRow
        versionLabel = NormalizeBoxBomVersionLabelShipping(loBom.DataBodyRange.Cells(r, cVersion).Value)
        If versionLabel = "" Then
            versionLabel = "v1"
            loBom.DataBodyRange.Cells(r, cVersion).Value = versionLabel
        End If
        If seen.Exists(versionLabel) Then GoTo NextBomRow

        oldCount = versionCount
        versionCount = versionCount + 1
        ReDim newRows(1 To versionCount, 1 To 8)
        If oldCount > 0 Then
            For c = 1 To 8
                For copyRow = 1 To oldCount
                    newRows(copyRow, c) = versionRows(copyRow, c)
                Next copyRow
            Next c
        End If
        newRows(versionCount, 1) = versionLabel
        newRows(versionCount, 2) = "Active"
        newRows(versionCount, 8) = CurrentBoxBuilderName(ws)
        versionRows = newRows
        seen(versionLabel) = True
NextBomRow:
    Next r

CleanExit:
End Sub

Public Sub HandleBoxBomVersionSelection(ByVal target As Range)
    On Error GoTo CleanExit

    Dim loVersions As ListObject
    Dim versionCol As ListColumn
    Dim versionLabel As String
    Dim report As String
    Dim packageRow As Long
    Dim rowIndex As Long

    If target Is Nothing Then Exit Sub
    If target.Cells.CountLarge > 1 Then Exit Sub
    On Error Resume Next
    Set loVersions = target.ListObject
    Set versionCol = loVersions.ListColumns("Version")
    On Error GoTo CleanExit
    If loVersions Is Nothing Or versionCol Is Nothing Then Exit Sub
    If StrComp(loVersions.Name, TABLE_BOX_BOM_VERSIONS, vbTextCompare) <> 0 Then Exit Sub
    If target.Row <= loVersions.HeaderRowRange.Row Then Exit Sub
    If target.Column <> versionCol.Range.Column Then Exit Sub

    versionLabel = Trim$(NzStr(target.Value))
    If versionLabel = "" Then Exit Sub
    packageRow = CurrentBoxBuilderPackageRow(target.Worksheet)
    rowIndex = target.Row - loVersions.DataBodyRange.Row + 1
    If Not BoxBomVersionRowMatchesPackage(loVersions, rowIndex, packageRow) Then
        MsgBox "Selected version does not belong to the current BoxBuilder box. Re-select the box and try again.", vbExclamation
        Exit Sub
    End If
    RememberSelectedBoxBomVersion target.Worksheet, versionLabel, packageRow
    If Not LoadSelectedBoxBomVersion(target.Worksheet, versionLabel, report) Then
        If report <> "" Then MsgBox report, vbExclamation
    End If

CleanExit:
End Sub

Private Function SelectedBoxBomVersionLabel(ByVal ws As Worksheet, _
                                            ByVal loVersions As ListObject, _
                                            Optional ByVal expectedPackageRow As Long = 0) As String
    On Error GoTo CleanExit

    Dim target As Range
    Dim versionCol As ListColumn
    Dim rowIndex As Long

    If ws Is Nothing Or loVersions Is Nothing Then Exit Function
    If Not ActiveCell Is Nothing Then
        If ActiveCell.Worksheet Is ws Then
            Set target = ActiveCell
            If Not target.ListObject Is Nothing Then
                If target.ListObject Is loVersions Then
                    If target.Row > loVersions.HeaderRowRange.Row Then
                        If Not loVersions.DataBodyRange Is Nothing Then
                            Set versionCol = loVersions.ListColumns("Version")
                            If Not versionCol Is Nothing Then
                                rowIndex = target.Row - loVersions.DataBodyRange.Row + 1
                                If rowIndex > 0 And rowIndex <= loVersions.ListRows.Count Then
                                    If BoxBomVersionRowMatchesPackage(loVersions, rowIndex, expectedPackageRow) Then
                                        SelectedBoxBomVersionLabel = Trim$(NzStr(loVersions.DataBodyRange.Cells(rowIndex, versionCol.Index).Value))
                                    End If
                                    If SelectedBoxBomVersionLabel <> "" Then
                                        RememberSelectedBoxBomVersion ws, SelectedBoxBomVersionLabel, expectedPackageRow
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If CachedBoxBomVersionSelectionMatches(ws, expectedPackageRow) Then
        SelectedBoxBomVersionLabel = mSelectedBoxBomVersionLabel
    End If

CleanExit:
End Function

Private Function ResolveBoxBomSaveVersionLabel(ByVal ws As Worksheet, ByVal loBom As ListObject) As String
    On Error GoTo CleanExit

    Dim loVersions As ListObject
    Dim packageRow As Long

    If ws Is Nothing Then Exit Function
    Set loVersions = GetListObject(ws, TABLE_BOX_BOM_VERSIONS)
    If loVersions Is Nothing Then Exit Function
    packageRow = CurrentBoxBuilderPackageRow(ws)
    If packageRow <= 0 Then Exit Function
    ResolveBoxBomSaveVersionLabel = SelectedBoxBomVersionLabel(ws, loVersions, packageRow)

CleanExit:
End Function

Private Function SingleVisibleBoxBomVersionLabel(ByVal loBom As ListObject) As String
    On Error GoTo CleanExit

    Dim cVersion As Long
    Dim r As Long
    Dim versionLabel As String
    Dim foundLabel As String

    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function
    cVersion = ColumnIndex(loBom, "Version")
    If cVersion = 0 Then Exit Function

    For r = 1 To loBom.ListRows.Count
        If BoxBomTableRowHiddenShipping(loBom, r) Then GoTo NextRow
        If Not BoxBomRowHasComponentDataShipping(loBom, r) Then GoTo NextRow
        versionLabel = NormalizeBoxBomVersionLabelShipping(loBom.DataBodyRange.Cells(r, cVersion).Value)
        If versionLabel = "" Then GoTo NextRow
        If foundLabel = "" Then
            foundLabel = versionLabel
        ElseIf StrComp(foundLabel, versionLabel, vbTextCompare) <> 0 Then
            SingleVisibleBoxBomVersionLabel = ""
            Exit Function
        End If
NextRow:
    Next r

    SingleVisibleBoxBomVersionLabel = foundLabel

CleanExit:
End Function

Private Function NormalizeBoxBomVersionLabelShipping(ByVal value As Variant) As String
    Dim textValue As String

    textValue = Trim$(NzStr(value))
    If textValue = "" Then Exit Function
    If LCase$(Left$(textValue, 1)) <> "v" Then textValue = "v" & textValue
    NormalizeBoxBomVersionLabelShipping = textValue
End Function

Private Function BoxBomTableRowHiddenShipping(ByVal loBom As ListObject, ByVal rowIndex As Long) As Boolean
    On Error GoTo CleanExit

    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > loBom.ListRows.Count Then Exit Function
    BoxBomTableRowHiddenShipping = loBom.DataBodyRange.Rows(rowIndex).EntireRow.Hidden

CleanExit:
End Function

Private Sub DeleteLocalBoxBomRowsForVersion(ByVal ws As Worksheet, ByVal versionLabel As String)
    On Error GoTo CleanExit

    Dim loBom As ListObject
    Dim cVersion As Long
    Dim i As Long

    If ws Is Nothing Then Exit Sub
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then Exit Sub
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBom Is Nothing Then Exit Sub
    If loBom.DataBodyRange Is Nothing Then Exit Sub
    cVersion = ColumnIndex(loBom, "Version")
    If cVersion = 0 Then Exit Sub

    For i = loBom.ListRows.Count To 1 Step -1
        If StrComp(NormalizeBoxBomVersionLabelShipping(loBom.DataBodyRange.Cells(i, cVersion).Value), versionLabel, vbTextCompare) = 0 Then
            loBom.ListRows(i).Delete
        End If
    Next i
    EnsureBoxBomStarterRows loBom

CleanExit:
End Sub

Private Sub DeleteLocalBoxBomVersionSummaryRow(ByVal loVersions As ListObject, ByVal versionLabel As String)
    On Error GoTo CleanExit

    Dim cVersion As Long
    Dim i As Long

    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If loVersions Is Nothing Or versionLabel = "" Then Exit Sub
    If loVersions.DataBodyRange Is Nothing Then Exit Sub
    cVersion = ColumnIndex(loVersions, "Version")
    If cVersion = 0 Then Exit Sub

    For i = loVersions.ListRows.Count To 1 Step -1
        If StrComp(NormalizeBoxBomVersionLabelShipping(loVersions.DataBodyRange.Cells(i, cVersion).Value), versionLabel, vbTextCompare) = 0 Then
            loVersions.ListRows(i).Delete
        End If
    Next i

CleanExit:
End Sub

Private Sub DeleteLocalShippingBomViewRowsForVersion(ByVal ws As Worksheet, ByVal packageRow As Long, ByVal versionLabel As String)
    On Error GoTo CleanExit

    Dim loView As ListObject
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim cLabel As Long
    Dim versionNumber As Long
    Dim rowVersion As Long
    Dim rowLabel As String
    Dim i As Long

    If ws Is Nothing Then Exit Sub
    If packageRow <= 0 Then Exit Sub
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then Exit Sub
    versionNumber = BomVersionNumberFromLabel(versionLabel)

    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If loView Is Nothing Then Exit Sub
    If loView.DataBodyRange Is Nothing Then Exit Sub
    cPackageRow = ColumnIndex(loView, "PackageRow")
    cVersion = ColumnIndex(loView, "BomVersion")
    cLabel = ColumnIndex(loView, "BomVersionLabel")
    If cPackageRow = 0 Then Exit Sub

    For i = loView.ListRows.Count To 1 Step -1
        If NzLng(loView.DataBodyRange.Cells(i, cPackageRow).Value) <> packageRow Then GoTo NextRow
        rowVersion = 0
        rowLabel = ""
        If cVersion > 0 Then rowVersion = NzLng(loView.DataBodyRange.Cells(i, cVersion).Value)
        If cLabel > 0 Then rowLabel = NormalizeBoxBomVersionLabelShipping(loView.DataBodyRange.Cells(i, cLabel).Value)
        If (versionNumber > 0 And rowVersion = versionNumber) _
           Or (rowLabel <> "" And StrComp(rowLabel, versionLabel, vbTextCompare) = 0) Then
            loView.ListRows(i).Delete
        End If
NextRow:
    Next i

CleanExit:
End Sub

Private Sub ClearSelectedBoxBomVersionIfMatches(ByVal ws As Worksheet, ByVal packageRow As Long, ByVal versionLabel As String)
    If ws Is Nothing Then Exit Sub
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then Exit Sub
    If Not CachedBoxBomVersionSelectionMatches(ws, packageRow) Then Exit Sub
    If StrComp(NormalizeBoxBomVersionLabelShipping(mSelectedBoxBomVersionLabel), versionLabel, vbTextCompare) <> 0 Then Exit Sub

    mSelectedBoxBomVersionLabel = ""
    mSelectedBoxBomVersionPackageRow = 0
    mSelectedBoxBomVersionWorkbookName = ""
    mSelectedBoxBomVersionWorksheetName = ""
End Sub

Private Sub DeleteAllListObjectRowsShipping(ByVal lo As ListObject)
    On Error GoTo CleanExit

    Dim i As Long

    If lo Is Nothing Then Exit Sub
    ClearListObjectFiltersShipping lo
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = lo.ListRows.Count To 1 Step -1
        lo.ListRows(i).Delete
    Next i

CleanExit:
End Sub

Private Sub ClearListObjectFiltersShipping(ByVal lo As ListObject)
    On Error Resume Next

    If lo Is Nothing Then Exit Sub
    If Not lo.AutoFilter Is Nothing Then lo.AutoFilter.ShowAllData
    If Not lo.Parent Is Nothing Then
        If lo.Parent.FilterMode Then lo.Parent.ShowAllData
    End If
    On Error GoTo 0
End Sub

Private Sub RememberSelectedBoxBomVersion(ByVal ws As Worksheet, _
                                          ByVal versionLabel As String, _
                                          ByVal packageRow As Long)
    If ws Is Nothing Then Exit Sub
    If packageRow <= 0 Then
        ClearSelectedBoxBomVersionForWorksheet ws
        Exit Sub
    End If
    mSelectedBoxBomVersionLabel = Trim$(versionLabel)
    mSelectedBoxBomVersionPackageRow = packageRow
    mSelectedBoxBomVersionWorkbookName = ws.Parent.Name
    mSelectedBoxBomVersionWorksheetName = ws.Name
End Sub

Private Function CachedBoxBomVersionSelectionMatches(ByVal ws As Worksheet, _
                                                     ByVal expectedPackageRow As Long) As Boolean
    If ws Is Nothing Then Exit Function
    If mSelectedBoxBomVersionLabel = "" Then Exit Function
    If StrComp(mSelectedBoxBomVersionWorkbookName, ws.Parent.Name, vbTextCompare) <> 0 Then Exit Function
    If StrComp(mSelectedBoxBomVersionWorksheetName, ws.Name, vbTextCompare) <> 0 Then Exit Function
    If expectedPackageRow > 0 Then
        If expectedPackageRow <> mSelectedBoxBomVersionPackageRow Then Exit Function
    End If
    CachedBoxBomVersionSelectionMatches = True
End Function

Private Function BoxBomVersionRowMatchesPackage(ByVal loVersions As ListObject, _
                                                ByVal rowIndex As Long, _
                                                ByVal expectedPackageRow As Long) As Boolean
    On Error GoTo CleanFail

    Dim cPackageRow As Long
    Dim cBoxName As Long
    Dim rowPackage As Long
    Dim currentPackageRow As Long
    Dim rowBoxName As String
    Dim currentBoxName As String

    If loVersions Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function
    If loVersions.DataBodyRange Is Nothing Then Exit Function
    If rowIndex > loVersions.ListRows.Count Then Exit Function
    If expectedPackageRow <= 0 Then Exit Function

    cBoxName = ColumnIndex(loVersions, "Box Name")
    If cBoxName > 0 Then
        rowBoxName = Trim$(NzStr(loVersions.DataBodyRange.Cells(rowIndex, cBoxName).Value))
        currentBoxName = CurrentBoxBuilderName(loVersions.Parent)
        If rowBoxName <> "" And currentBoxName <> "" Then
            If StrComp(rowBoxName, currentBoxName, vbTextCompare) <> 0 Then Exit Function
        End If
    End If

    cPackageRow = ColumnIndex(loVersions, "PackageRow")
    If cPackageRow = 0 Then
        currentPackageRow = CurrentBoxBuilderPackageRow(loVersions.Parent)
        BoxBomVersionRowMatchesPackage = (currentPackageRow = expectedPackageRow)
        Exit Function
    End If

    rowPackage = NzLng(loVersions.DataBodyRange.Cells(rowIndex, cPackageRow).Value)
    BoxBomVersionRowMatchesPackage = (rowPackage = expectedPackageRow)
    Exit Function

CleanFail:
    BoxBomVersionRowMatchesPackage = False
End Function

Private Function CurrentBoxBuilderPackageRow(ByVal ws As Worksheet) As Long
    On Error GoTo CleanExit

    Dim loBuilder As ListObject
    Dim boxName As String
    Dim runtimeMax As Long

    If ws Is Nothing Then Exit Function
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then Exit Function
    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    If boxName = "" Then Exit Function
    CurrentBoxBuilderPackageRow = FindShippingBomPackageRowByName(ws.Parent, boxName, runtimeMax)

CleanExit:
End Function

Private Function CurrentBoxBuilderName(ByVal ws As Worksheet) As String
    On Error GoTo CleanExit

    Dim loBuilder As ListObject

    If ws Is Nothing Then Exit Function
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then Exit Function
    CurrentBoxBuilderName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))

CleanExit:
End Function

Private Function LoadSelectedBoxBomVersion(ByVal ws As Worksheet, _
                                           ByVal versionLabel As String, _
                                           ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim loBuilder As ListObject
    Dim boxName As String
    Dim packageRow As Long
    Dim runtimeMax As Long
    Dim versionNumber As Long

    report = ""
    If ws Is Nothing Then Exit Function
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then
        report = "BoxBuilder table was not found."
        Exit Function
    End If

    boxName = Trim$(NzStr(ValueFromTable(loBuilder, "Box Name")))
    If boxName = "" Then
        report = "BoxBuilder Box Name is required before selecting a version."
        Exit Function
    End If

    packageRow = FindShippingBomPackageRowByName(ws.Parent, boxName, runtimeMax)
    If packageRow <= 0 Then
        report = "Saved box '" & boxName & "' was not found in ShippingBOM runtime."
        Exit Function
    End If

    versionNumber = BomVersionNumberFromLabel(versionLabel)
    If versionNumber <= 0 Then
        report = "Could not resolve selected version '" & versionLabel & "'."
        Exit Function
    End If

    LoadSelectedBoxBomVersion = LoadBoxBomForPackageVersion(ws, packageRow, versionNumber, report)
    Exit Function

FailSoft:
    report = "Load selected BoxBOM version failed: " & Err.Description
End Function

Private Function LoadBoxBomForPackageVersion(ByVal ws As Worksheet, _
                                             ByVal packageRow As Long, _
                                             ByVal versionNumber As Long, _
                                             ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim loBuilder As ListObject
    Dim loBom As ListObject
    Dim arr As Variant
    Dim cPackageRow As Long
    Dim cPackageItem As Long
    Dim cPackageUom As Long
    Dim cPackageLocation As Long
    Dim cPackageDescription As Long
    Dim cVersion As Long
    Dim cComponentRow As Long
    Dim cComponentItemCode As Long
    Dim cComponentItem As Long
    Dim cComponentQty As Long
    Dim cComponentUom As Long
    Dim cComponentLocation As Long
    Dim cComponentDescription As Long
    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loRuntime As ListObject
    Dim openedTransient As Boolean
    Dim r As Long
    Dim outRow As Long
    Dim componentName As String
    Dim componentCode As String
    Dim componentRow As Long
    Dim componentUom As String
    Dim componentLocation As String
    Dim componentDescription As String
    Dim repairedRows As Long
    Dim matchCount As Long
    Dim quietStarted As Boolean

    report = ""
    If ws Is Nothing Then Exit Function
    If packageRow <= 0 Or versionNumber <= 0 Then Exit Function

    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loBuilder Is Nothing Or loBom Is Nothing Then
        report = "BoxBuilder or BoxBOM table was not found."
        Exit Function
    End If
    EnsureTableHasRow loBuilder
    EnsureBoxBomEntryColumns loBom

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
    If wbBom Is Nothing Then Exit Function
    Set loRuntime = EnsureShippingBomSchema(wbBom, report)
    If loRuntime Is Nothing Then GoTo CleanExit
    If loRuntime.DataBodyRange Is Nothing Then
        report = "No saved Shipping BOM rows are available."
        GoTo CleanExit
    End If

    cPackageRow = ColumnIndex(loRuntime, "PackageRow")
    cPackageItem = ColumnIndex(loRuntime, "PackageItem")
    cPackageUom = ColumnIndex(loRuntime, "PackageUOM")
    cPackageLocation = ColumnIndex(loRuntime, "PackageLocation")
    cPackageDescription = ColumnIndex(loRuntime, "PackageDescription")
    cVersion = ColumnIndex(loRuntime, "BomVersion")
    cComponentRow = ColumnIndex(loRuntime, "ComponentRow")
    cComponentItemCode = ColumnIndex(loRuntime, "ComponentItemCode")
    cComponentItem = ColumnIndex(loRuntime, "ComponentItem")
    cComponentQty = ColumnIndex(loRuntime, "ComponentQty")
    cComponentUom = ColumnIndex(loRuntime, "ComponentUOM")
    cComponentLocation = ColumnIndex(loRuntime, "ComponentLocation")
    cComponentDescription = ColumnIndex(loRuntime, "ComponentDescription")
    If cPackageRow = 0 Or cVersion = 0 Or cComponentRow = 0 Or cComponentQty = 0 Then
        report = "ShippingBOM runtime is missing required PackageRow/BomVersion/Component columns."
        GoTo CleanExit
    End If

    arr = loRuntime.DataBodyRange.Value

    For r = 1 To UBound(arr, 1)
        If NzLng(arr(r, cPackageRow)) = packageRow _
           And NzLng(arr(r, cVersion)) = versionNumber Then
            If ShippingBomSourceRowHasComponent(arr, r, cComponentRow, cComponentItem, cComponentQty) Then
                matchCount = matchCount + 1
            End If
        End If
    Next r
    If matchCount = 0 Then
        If DisplayedBoxBomHasVersionRowsShipping(loBom, "v" & CStr(versionNumber)) Then
            RefreshBoxBomVersionList ws, packageRow
            RememberSelectedBoxBomVersion ws, "v" & CStr(versionNumber), packageRow
            report = "Displayed BoxBOM already shows v" & CStr(versionNumber) & "."
            LoadBoxBomForPackageVersion = True
            GoTo CleanExit
        End If
        report = "No saved BoxBOM rows were found for ROW " & CStr(packageRow) & " v" & CStr(versionNumber) & "."
        GoTo CleanExit
    End If

    mHandlingShippingSheetChange = True
    modUiQuiet.BeginQuietUi ws.Parent
    quietStarted = True
    ClearListObjectData loBom
    EnsureBoxBomStarterRows loBom

    For r = 1 To UBound(arr, 1)
        If NzLng(arr(r, cPackageRow)) <> packageRow Then GoTo NextBomRow
        If NzLng(arr(r, cVersion)) <> versionNumber Then GoTo NextBomRow
        If Not ShippingBomSourceRowHasComponent(arr, r, cComponentRow, cComponentItem, cComponentQty) Then GoTo NextBomRow
        If outRow = 0 Then
            If cPackageItem > 0 Then SetTableCellShipping loBuilder, 1, "Box Name", arr(r, cPackageItem)
            If cPackageUom > 0 Then SetTableCellShipping loBuilder, 1, "UOM", arr(r, cPackageUom)
            If cPackageLocation > 0 Then SetTableCellShipping loBuilder, 1, "LOCATION", arr(r, cPackageLocation)
            If cPackageDescription > 0 Then SetTableCellShipping loBuilder, 1, "DESCRIPTION", arr(r, cPackageDescription)
        End If

        outRow = outRow + 1
        Do While loBom.ListRows.Count < outRow
            loBom.ListRows.Add
        Loop
        componentName = ValueFromArrayColumn(arr, r, cComponentItem)
        componentCode = ""
        If cComponentItemCode > 0 Then componentCode = ValueFromArrayColumn(arr, r, cComponentItemCode)
        componentRow = NzLng(arr(r, cComponentRow))
        componentUom = ValueFromArrayColumn(arr, r, cComponentUom)
        componentLocation = ValueFromArrayColumn(arr, r, cComponentLocation)
        componentDescription = ValueFromArrayColumn(arr, r, cComponentDescription)
        If componentRow <= 0 Then
            If ResolveCanonicalComponentInfoShipping(componentName, componentCode, componentRow, componentName, componentCode, componentUom, componentLocation, componentDescription) Then
                repairedRows = repairedRows + 1
            End If
        End If

        SetTableCellShipping loBom, outRow, "Version", "v" & CStr(versionNumber)
        SetTableCellShipping loBom, outRow, COL_BOXBOM_ITEM, componentName
        SetTableCellShipping loBom, outRow, "ITEM_CODE", componentCode
        SetTableCellShipping loBom, outRow, "ROW", componentRow
        SetTableCellShipping loBom, outRow, "QUANTITY", NzDbl(arr(r, cComponentQty))
        SetTableCellShipping loBom, outRow, "UOM", componentUom
        SetTableCellShipping loBom, outRow, "LOCATION", componentLocation
        SetTableCellShipping loBom, outRow, "DESCRIPTION", componentDescription
NextBomRow:
    Next r

    EnsureBoxBomStarterRows loBom
    FillBlankBoxBomVersionShipping loBom
    SortBoxBomByVersionShipping loBom

    RefreshBoxMakerCurrentInventory ws
    RefreshBoxBomVersionListFromSourceTable ws, loRuntime, packageRow
    RememberSelectedBoxBomVersion ws, "v" & CStr(versionNumber), packageRow
    report = "Loaded " & CStr(outRow) & " component row(s) for v" & CStr(versionNumber) & "."
    If repairedRows > 0 Then report = report & " Repaired " & CStr(repairedRows) & " component ROW value(s) from inventory."
    LoadBoxBomForPackageVersion = True

CleanExit:
    If quietStarted Then
        On Error Resume Next
        modUiQuiet.EndQuietUi
        On Error GoTo 0
    End If
    If quietStarted Then mHandlingShippingSheetChange = False
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "LoadBoxBomForPackageVersion failed: " & Err.Description
    Resume CleanExit
End Function

Private Sub RefreshBoxMakerBomVersionDisplay(ByVal ws As Worksheet, ByVal packageRow As Long)
    On Error GoTo CleanExit

    Dim loBuilder As ListObject
    Dim versionLabel As String
    Dim cVersion As Long

    If ws Is Nothing Then Exit Sub
    If packageRow <= 0 Then Exit Sub
    Set loBuilder = GetListObject(ws, TABLE_BOX_BUILDER)
    If loBuilder Is Nothing Then Exit Sub
    If loBuilder.DataBodyRange Is Nothing Then Exit Sub
    cVersion = ColumnIndex(loBuilder, COL_BOM_VERSION)
    If cVersion = 0 Then Exit Sub

    versionLabel = ActiveBoxBomVersionLabel(ws, packageRow)
    If versionLabel = "" Then versionLabel = "v1"
    loBuilder.DataBodyRange.Cells(1, cVersion).Value = versionLabel
    FormatBoxMakerReadOnlyColumn loBuilder, COL_BOM_VERSION

CleanExit:
End Sub

Private Function ActiveBoxBomVersionLabel(ByVal ws As Worksheet, ByVal packageRow As Long) As String
    Dim loView As ListObject
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim cLabel As Long
    Dim cActive As Long
    Dim r As Long

    If ws Is Nothing Then Exit Function
    Set loView = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
    If loView Is Nothing Then Exit Function
    If loView.DataBodyRange Is Nothing Then Exit Function

    cPackageRow = ColumnIndex(loView, "PackageRow")
    cVersion = ColumnIndex(loView, "BomVersion")
    cLabel = ColumnIndex(loView, "BomVersionLabel")
    cActive = ColumnIndex(loView, "IsActive")
    If cPackageRow = 0 Then Exit Function

    For r = 1 To loView.ListRows.Count
        If NzLng(loView.DataBodyRange.Cells(r, cPackageRow).Value) = packageRow Then
            If cActive = 0 Or ShippingBomActiveValue(loView.DataBodyRange.Cells(r, cActive).Value) Then
                If cLabel > 0 Then ActiveBoxBomVersionLabel = Trim$(NzStr(loView.DataBodyRange.Cells(r, cLabel).Value))
                If ActiveBoxBomVersionLabel = "" And cVersion > 0 Then ActiveBoxBomVersionLabel = "v" & CStr(NzLng(loView.DataBodyRange.Cells(r, cVersion).Value))
                Exit Function
            End If
        End If
    Next r
End Function

Private Function BuildBoxBomVersionRows(ByVal loView As ListObject, _
                                        ByVal packageRow As Long, _
                                        ByRef versionCount As Long) As Variant
    Dim cPackageRow As Long
    Dim cPackageItem As Long
    Dim cVersion As Long
    Dim cLabel As Long
    Dim cActive As Long
    Dim cEffectiveFrom As Long
    Dim cEffectiveTo As Long
    Dim cRetiredAt As Long
    Dim cUpdatedAt As Long
    Dim cUpdatedBy As Long
    Dim cComponentRow As Long
    Dim cComponentItem As Long
    Dim cComponentQty As Long
    Dim dict As Object
    Dim src As Variant
    Dim rowData As Variant
    Dim result() As Variant
    Dim keys As Variant
    Dim r As Long
    Dim c As Long
    Dim key As String

    versionCount = 0
    If loView Is Nothing Then Exit Function
    If loView.DataBodyRange Is Nothing Then Exit Function

    cPackageRow = ColumnIndex(loView, "PackageRow")
    cPackageItem = ColumnIndex(loView, "PackageItem")
    cVersion = ColumnIndex(loView, "BomVersion")
    cLabel = ColumnIndex(loView, "BomVersionLabel")
    cActive = ColumnIndex(loView, "IsActive")
    cEffectiveFrom = ColumnIndex(loView, "EffectiveFromUTC")
    cEffectiveTo = ColumnIndex(loView, "EffectiveToUTC")
    cRetiredAt = ColumnIndex(loView, "RetiredAtUTC")
    cUpdatedAt = ColumnIndex(loView, "UpdatedAtUTC")
    cUpdatedBy = ColumnIndex(loView, "UpdatedBy")
    cComponentRow = ColumnIndex(loView, "ComponentRow")
    cComponentItem = ColumnIndex(loView, "ComponentItem")
    cComponentQty = ColumnIndex(loView, "ComponentQty")
    If cPackageRow = 0 Then Exit Function

    src = loView.DataBodyRange.Value
    Set dict = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(src, 1)
        If NzLng(src(r, cPackageRow)) <> packageRow Then GoTo NextRow
        If Not ShippingBomSourceRowHasComponent(src, r, cComponentRow, cComponentItem, cComponentQty) Then GoTo NextRow

        key = VersionKeyShipping(src, r, cVersion)
        If dict.Exists(key) Then GoTo NextRow

        ReDim rowData(1 To 8)
        rowData(1) = VersionLabelShipping(src, r, cVersion, cLabel)
        If cActive > 0 And Not ShippingBomActiveValue(src(r, cActive)) Then
            rowData(2) = "Retired"
        Else
            rowData(2) = "Active"
        End If
        If cEffectiveFrom > 0 Then rowData(3) = src(r, cEffectiveFrom)
        If cEffectiveTo > 0 Then rowData(4) = src(r, cEffectiveTo)
        If cRetiredAt > 0 Then rowData(5) = src(r, cRetiredAt)
        If cUpdatedAt > 0 Then rowData(6) = src(r, cUpdatedAt)
        If cUpdatedBy > 0 Then rowData(7) = src(r, cUpdatedBy)
        If cPackageItem > 0 Then rowData(8) = src(r, cPackageItem)
        dict.Add key, rowData
NextRow:
    Next r

    versionCount = dict.Count
    If versionCount = 0 Then Exit Function

    keys = dict.Keys
    ReDim result(1 To versionCount, 1 To 8)
    For r = 0 To UBound(keys)
        rowData = dict(keys(r))
        For c = 1 To 8
            result(r + 1, c) = rowData(c)
        Next c
    Next r
    BuildBoxBomVersionRows = result
End Function

Private Function ShippingBomSourceRowHasComponent(ByRef src As Variant, _
                                                  ByVal rowIndex As Long, _
                                                  ByVal cComponentRow As Long, _
                                                  ByVal cComponentItem As Long, _
                                                  ByVal cComponentQty As Long) As Boolean
    If cComponentRow > 0 Then
        If NzLng(src(rowIndex, cComponentRow)) > 0 Then
            ShippingBomSourceRowHasComponent = True
            Exit Function
        End If
    End If
    If cComponentItem > 0 Then
        If Trim$(NzStr(src(rowIndex, cComponentItem))) <> "" Then
            ShippingBomSourceRowHasComponent = True
            Exit Function
        End If
    End If
    If cComponentQty > 0 Then
        If NzDbl(src(rowIndex, cComponentQty)) <> 0 Then ShippingBomSourceRowHasComponent = True
    End If
End Function

Private Function VersionKeyShipping(ByRef src As Variant, ByVal rowIndex As Long, ByVal cVersion As Long) As String
    Dim versionValue As Long

    If cVersion > 0 Then versionValue = NzLng(src(rowIndex, cVersion))
    If versionValue <= 0 Then versionValue = 1
    VersionKeyShipping = Format$(versionValue, "0000000000")
End Function

Private Function VersionLabelShipping(ByRef src As Variant, _
                                      ByVal rowIndex As Long, _
                                      ByVal cVersion As Long, _
                                      ByVal cLabel As Long) As String
    Dim versionValue As Long

    If cLabel > 0 Then VersionLabelShipping = Trim$(NzStr(src(rowIndex, cLabel)))
    If VersionLabelShipping <> "" Then Exit Function

    If cVersion > 0 Then
        versionValue = NzLng(src(rowIndex, cVersion))
        If versionValue <= 0 Then versionValue = 1
        VersionLabelShipping = "v" & CStr(versionValue)
    Else
        VersionLabelShipping = "v1"
    End If
End Function

Private Function BomVersionNumberFromLabel(ByVal versionLabel As String) As Long
    Dim textValue As String

    textValue = LCase$(Trim$(versionLabel))
    If Left$(textValue, 1) = "v" Then textValue = Mid$(textValue, 2)
    BomVersionNumberFromLabel = NzLng(textValue)
    If BomVersionNumberFromLabel <= 0 Then BomVersionNumberFromLabel = 1
End Function

Private Function TryLoadRuntimeShippingBomRows(ByRef arr As Variant, _
                                               ByRef cPackageRow As Long, _
                                               ByRef cVersion As Long, _
                                               ByRef cVersionLabel As Long, _
                                               ByRef cComponentRow As Long, _
                                               ByRef cComponentItemCode As Long, _
                                               ByRef cComponentItem As Long, _
                                               ByRef cComponentQty As Long, _
                                               ByRef cComponentUom As Long, _
                                               ByRef cComponentLocation As Long, _
                                               ByRef cComponentDescription As Long, _
                                               ByRef cActive As Long, _
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
    cVersion = ColumnIndex(loBom, "BomVersion")
    cVersionLabel = ColumnIndex(loBom, "BomVersionLabel")
    cComponentRow = ColumnIndex(loBom, "ComponentRow")
    cComponentItemCode = ColumnIndex(loBom, "ComponentItemCode")
    cComponentItem = ColumnIndex(loBom, "ComponentItem")
    cComponentQty = ColumnIndex(loBom, "ComponentQty")
    cComponentUom = ColumnIndex(loBom, "ComponentUOM")
    cComponentLocation = ColumnIndex(loBom, "ComponentLocation")
    cComponentDescription = ColumnIndex(loBom, "ComponentDescription")
    cActive = ColumnIndex(loBom, "IsActive")
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
                                           ByRef errNotes As String, _
                                           Optional ByRef eventQueuedOut As Boolean = False, _
                                           Optional ByRef batchProcessedOut As Boolean = False, _
                                           Optional ByRef eventIdOut As String = "", _
                                           Optional ByRef runtimeReportOut As String = "") As Boolean
    Dim runtimeReport As String

    errNotes = ""
    eventQueuedOut = False
    batchProcessedOut = False
    eventIdOut = ""
    runtimeReportOut = ""
    If loBuilder Is Nothing Or loBom Is Nothing Then
        errNotes = "Box Created required tables are missing."
        Exit Function
    End If

    EnsureColumnExists loBuilder, "Quantity", "Box Name"
    EnsureBoxBomEntryColumns loBom

    If Not QueueBoxBuildEventFromBuilder(loBuilder, loBom, invLo, usedTotal, madeTotal, eventIdOut, errNotes) Then Exit Function
    eventQueuedOut = True

    batchProcessedOut = RunShippingRuntimeQueueRefresh(loBuilder.Parent.Parent, ResolveCurrentShippingWarehouseId(), runtimeReport)
    If Not batchProcessedOut Then batchProcessedOut = BoxMakerRuntimeReportShowsProcessed(runtimeReport)
    runtimeReportOut = runtimeReport
    If Not batchProcessedOut Then
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
                                           ByRef errNotes As String, _
                                           Optional ByRef eventQueuedOut As Boolean = False, _
                                           Optional ByRef batchProcessedOut As Boolean = False, _
                                           Optional ByRef eventIdOut As String = "", _
                                           Optional ByRef runtimeReportOut As String = "") As Boolean
    Dim runtimeReport As String

    errNotes = ""
    eventQueuedOut = False
    batchProcessedOut = False
    eventIdOut = ""
    runtimeReportOut = ""
    If loBuilder Is Nothing Or loBom Is Nothing Then
        errNotes = "Box Unboxed required tables are missing."
        Exit Function
    End If

    EnsureColumnExists loBuilder, "Quantity", "Box Name"
    EnsureBoxBomEntryColumns loBom

    If Not QueueBoxUnboxEventFromBuilder(loBuilder, loBom, invLo, componentsReturned, packageReturned, eventIdOut, errNotes) Then Exit Function
    eventQueuedOut = True

    batchProcessedOut = RunShippingRuntimeQueueRefresh(loBuilder.Parent.Parent, ResolveCurrentShippingWarehouseId(), runtimeReport)
    If Not batchProcessedOut Then batchProcessedOut = BoxMakerRuntimeReportShowsProcessed(runtimeReport)
    runtimeReportOut = runtimeReport
    If Not batchProcessedOut Then
        If runtimeReport = "" Then runtimeReport = "Box unbox event queued, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If
    If eventIdOut <> "" Then AppendNote errNotes, "Inbox EventID: " & eventIdOut

    ApplyBoxUnboxedFromBuilder = True
End Function

Private Function BoxMakerRuntimeReportShowsProcessed(ByVal runtimeReport As String) As Boolean
    Dim reportText As String

    reportText = Trim$(runtimeReport)
    If reportText = "" Then Exit Function

    If InStr(1, reportText, "RunBatch processed queued event", vbTextCompare) > 0 Then
        BoxMakerRuntimeReportShowsProcessed = True
        Exit Function
    End If
    If BoxMakerRuntimeReportMetric(reportText, "Processed") > 0 Then
        BoxMakerRuntimeReportShowsProcessed = True
        Exit Function
    End If
    If BoxMakerRuntimeReportMetric(reportText, "Applied") > 0 Then
        BoxMakerRuntimeReportShowsProcessed = True
        Exit Function
    End If
    If BoxMakerRuntimeReportMetric(reportText, "SkipDup") > 0 Then
        BoxMakerRuntimeReportShowsProcessed = True
    End If
End Function

Private Function BoxMakerRuntimeReportMetric(ByVal runtimeReport As String, ByVal metricName As String) As Long
    Dim marker As String
    Dim pos As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim ch As String

    marker = metricName & "="
    pos = InStr(1, runtimeReport, marker, vbTextCompare)
    If pos <= 0 Then Exit Function

    valueStart = pos + Len(marker)
    valueEnd = valueStart
    Do While valueEnd <= Len(runtimeReport)
        ch = Mid$(runtimeReport, valueEnd, 1)
        If ch < "0" Or ch > "9" Then Exit Do
        valueEnd = valueEnd + 1
    Loop
    If valueEnd <= valueStart Then Exit Function
    BoxMakerRuntimeReportMetric = CLng(Mid$(runtimeReport, valueStart, valueEnd - valueStart))
End Function

Private Function CaptureBoxMakerCurrentInventoryState(ByVal loBuilder As ListObject, _
                                                      ByVal loBom As ListObject) As Object
    Dim state As Object
    Dim cInv As Long
    Dim cQty As Long
    Dim r As Long
    Dim currentVal As Variant
    Dim qtyVal As Variant

    Set state = CreateObject("Scripting.Dictionary")
    state.Add "BuilderHasCurrent", False
    state.Add "BuilderCurrent", 0#
    state.Add "BuilderQty", 0#
    state.Add "BomRows", 0&

    If Not loBuilder Is Nothing Then
        If Not loBuilder.DataBodyRange Is Nothing Then
            cInv = ColumnIndex(loBuilder, COL_CURRENT_INV)
            cQty = ColumnIndex(loBuilder, "Quantity")
            If cQty > 0 Then state("BuilderQty") = NzDbl(loBuilder.DataBodyRange.Cells(1, cQty).Value)
            If cInv > 0 Then
                currentVal = loBuilder.DataBodyRange.Cells(1, cInv).Value
                If IsNumeric(currentVal) Then
                    state("BuilderHasCurrent") = True
                    state("BuilderCurrent") = CDbl(currentVal)
                End If
            End If
        End If
    End If

    If Not loBom Is Nothing Then
        If Not loBom.DataBodyRange Is Nothing Then
            cInv = ColumnIndex(loBom, COL_CURRENT_INV)
            cQty = ColumnIndex(loBom, "QUANTITY")
            If cInv > 0 And cQty > 0 Then
                state("BomRows") = CLng(loBom.ListRows.Count)
                For r = 1 To loBom.ListRows.Count
                    state.Add "BomHasCurrent_" & CStr(r), False
                    state.Add "BomCurrent_" & CStr(r), 0#
                    state.Add "BomQty_" & CStr(r), 0#

                    qtyVal = loBom.DataBodyRange.Cells(r, cQty).Value
                    If IsNumeric(qtyVal) Then state("BomQty_" & CStr(r)) = CDbl(qtyVal)

                    currentVal = loBom.DataBodyRange.Cells(r, cInv).Value
                    If IsNumeric(currentVal) Then
                        state("BomHasCurrent_" & CStr(r)) = True
                        state("BomCurrent_" & CStr(r)) = CDbl(currentVal)
                    End If
                Next r
            End If
        End If
    End If

    Set CaptureBoxMakerCurrentInventoryState = state
End Function

Private Sub ApplyBoxUnboxExpectedCurrentInventoryDisplay(ByVal loBuilder As ListObject, _
                                                         ByVal loBom As ListObject, _
                                                         ByVal state As Object)
    Dim cInv As Long
    Dim r As Long
    Dim maxRows As Long

    If state Is Nothing Then Exit Sub

    If Not loBuilder Is Nothing Then
        If Not loBuilder.DataBodyRange Is Nothing Then
            cInv = ColumnIndex(loBuilder, COL_CURRENT_INV)
            If cInv > 0 Then
                If CBool(state("BuilderHasCurrent")) Then
                    loBuilder.DataBodyRange.Cells(1, cInv).Value = CDbl(state("BuilderCurrent")) - CDbl(state("BuilderQty"))
                End If
                FormatCurrentInventoryColumn loBuilder
            End If
        End If
    End If

    If Not loBom Is Nothing Then
        If Not loBom.DataBodyRange Is Nothing Then
            cInv = ColumnIndex(loBom, COL_CURRENT_INV)
            If cInv > 0 Then
                maxRows = CLng(state("BomRows"))
                If maxRows > loBom.ListRows.Count Then maxRows = loBom.ListRows.Count
                For r = 1 To maxRows
                    If CBool(state("BomHasCurrent_" & CStr(r))) Then
                        loBom.DataBodyRange.Cells(r, cInv).Value = CDbl(state("BomCurrent_" & CStr(r))) + CDbl(state("BomQty_" & CStr(r)))
                    End If
                Next r
                FormatCurrentInventoryColumn loBom
            End If
        End If
    End If
End Sub

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

Private Function CollectBomComponents(loBom As ListObject, invLo As ListObject, ByRef syncNotes As String, Optional ByVal targetVersionLabel As String = "") As Collection
    Dim result As New Collection
    If loBom Is Nothing Then
        Set CollectBomComponents = result
        Exit Function
    End If

    targetVersionLabel = NormalizeBoxBomVersionLabelShipping(targetVersionLabel)
    Dim cName As Long: cName = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    Dim cVersion As Long: cVersion = ColumnIndex(loBom, "Version")
    Dim cCode As Long: cCode = ColumnIndex(loBom, "ITEM_CODE")
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

    Dim invRowCol As Long
    Dim invCodeCol As Long
    Dim invItemCol As Long
    Dim invUomCol As Long
    Dim invLocCol As Long
    Dim invDescCol As Long
    If Not invLo Is Nothing Then
        invRowCol = ColumnIndex(invLo, "ROW")
        invCodeCol = ColumnIndex(invLo, "ITEM_CODE")
        invItemCol = ColumnIndex(invLo, "ITEM")
        invUomCol = ColumnIndex(invLo, "UOM")
        invLocCol = ColumnIndex(invLo, "LOCATION")
        invDescCol = ColumnIndex(invLo, "DESCRIPTION")
    End If

    Dim arr As Variant: arr = loBom.DataBodyRange.Value
    Dim r As Long
    Dim seenVersions As Object
    Set seenVersions = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(arr, 1)
        If BoxBomTableRowHiddenShipping(loBom, r) Then GoTo NextComponent
        Dim rowVersionLabel As String
        If cVersion > 0 Then rowVersionLabel = NormalizeBoxBomVersionLabelShipping(NzStr(arr(r, cVersion)))
        If targetVersionLabel <> "" Then
            If rowVersionLabel <> "" And StrComp(rowVersionLabel, targetVersionLabel, vbTextCompare) <> 0 Then GoTo NextComponent
        ElseIf rowVersionLabel <> "" Then
            seenVersions(rowVersionLabel) = True
            If seenVersions.Count > 1 Then
                Err.Raise vbObjectError + 16, , "BoxBOM contains multiple versions. Filter or select one Version before saving edits."
            End If
        End If

        Dim partName As String: partName = Trim$(NzStr(arr(r, cName)))
        Dim partCode As String
        If cCode > 0 Then partCode = Trim$(NzStr(arr(r, cCode)))
        Dim partRow As Long: partRow = NzLng(arr(r, cRow))
        Dim qty As Double: qty = NzDbl(arr(r, cQty))
        Dim uomVal As String: uomVal = Trim$(NzStr(arr(r, cUom)))

        If partName = "" And partRow = 0 And qty = 0 Then GoTo NextComponent
        If qty <= 0 Then
            Err.Raise vbObjectError + 1, , "Component row " & r & " has no quantity."
        End If

        Dim invIdx As Long
        Dim partResolvedName As String
        Dim actualUom As String, actualLoc As String, actualDesc As String
        Dim actualItem As String
        actualUom = uomVal
        If cLoc > 0 Then actualLoc = Trim$(NzStr(arr(r, cLoc)))
        If cDesc > 0 Then actualDesc = Trim$(NzStr(arr(r, cDesc)))
        If Not invLo Is Nothing Then
            If partRow > 0 Then invIdx = FindInvRowIndexByRow(invLo, partRow)
            If invIdx = 0 And partCode <> "" Then invIdx = FindInvRowIndexByItemCode(invLo, partCode)
            If invIdx = 0 And partName <> "" Then invIdx = FindInvRowIndexByItem(invLo, partName)
        End If

        If partRow <= 0 And invIdx > 0 And invRowCol > 0 Then
            partRow = NzLng(invLo.DataBodyRange.Cells(invIdx, invRowCol).Value)
            If partRow > 0 And partName <> "" Then AppendSyncMessage syncNotes, "Updated ROW for '" & partName & "' to " & partRow & "."
        End If
        If partRow <= 0 And (partName <> "" Or partCode <> "") Then
            If ResolveCanonicalComponentInfoShipping(partName, partCode, partRow, actualItem, partCode, actualUom, actualLoc, actualDesc) Then
                If partRow > 0 Then AppendSyncMessage syncNotes, "Updated ROW for '" & partName & "' to " & partRow & "."
            End If
        End If
        If partRow <= 0 And partName = "" Then
            Err.Raise vbObjectError + 4, , "Component row " & r & " is missing both item name and ROW."
        End If
        If partRow <= 0 Then
            Err.Raise vbObjectError + 5, , "Component '" & partName & "' is missing a valid invSys ROW. Re-pick the component before saving."
        End If

        If invIdx > 0 Then
            If invCodeCol > 0 Then partCode = NzStr(invLo.DataBodyRange.Cells(invIdx, invCodeCol).Value)
            If invItemCol > 0 Then actualItem = NzStr(invLo.DataBodyRange.Cells(invIdx, invItemCol).Value)
            If invUomCol > 0 Then actualUom = NzStr(invLo.DataBodyRange.Cells(invIdx, invUomCol).Value)
            If invLocCol > 0 Then actualLoc = NzStr(invLo.DataBodyRange.Cells(invIdx, invLocCol).Value)
            If invDescCol > 0 Then actualDesc = NzStr(invLo.DataBodyRange.Cells(invIdx, invDescCol).Value)
        End If
        If actualItem <> "" Then partResolvedName = actualItem Else partResolvedName = partName
        If actualUom = "" Then actualUom = uomVal
        If uomVal <> "" And StrComp(uomVal, actualUom, vbTextCompare) <> 0 Then
            AppendSyncMessage syncNotes, "UOM for '" & partResolvedName & "' reset to " & actualUom & "."
        End If
        uomVal = actualUom
        If actualLoc = "" And cLoc > 0 Then actualLoc = Trim$(NzStr(arr(r, cLoc)))
        If actualDesc = "" And cDesc > 0 Then actualDesc = Trim$(NzStr(arr(r, cDesc)))

        If cName > 0 And partResolvedName <> "" Then
            loBom.DataBodyRange.Cells(r, cName).Value = partResolvedName
        End If
        If cCode > 0 Then loBom.DataBodyRange.Cells(r, cCode).Value = partCode
        loBom.DataBodyRange.Cells(r, cRow).Value = partRow
        loBom.DataBodyRange.Cells(r, cUom).Value = uomVal
        If cLoc > 0 Then loBom.DataBodyRange.Cells(r, cLoc).Value = actualLoc
        If cDesc > 0 Then loBom.DataBodyRange.Cells(r, cDesc).Value = actualDesc

        Dim entry(1 To 7) As Variant
        entry(1) = partRow
        entry(2) = qty
        entry(3) = uomVal
        entry(4) = partResolvedName
        entry(5) = actualLoc
        entry(6) = actualDesc
        entry(7) = partCode
        result.Add entry
NextComponent:
    Next

    Set CollectBomComponents = result
End Function

Private Sub EnsureBoxBomEntryColumns(loBom As ListObject)
    If loBom Is Nothing Then Exit Sub
    UnhideListObjectWorksheetColumnsShipping loBom
    RepairMisheadedBoxBomColumnsShipping loBom
    RemoveBlankUnexpectedBoxBomColumnsShipping loBom
    Dim idxItem As Long
    EnsureColumnExists loBom, "Version"
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
    ApplyBoxBomReadOnlyColumnFormatting loBom
End Sub

Private Sub ApplyBoxBomReadOnlyColumnFormatting(ByVal loBom As ListObject)
    If loBom Is Nothing Then Exit Sub
    EnsureShippingWorksheetEditable loBom.Parent
    On Error Resume Next
    loBom.Parent.Cells.Locked = False
    loBom.Range.Locked = False
    On Error GoTo 0
    FormatBoxMakerReadOnlyColumn loBom, "Version"
    FormatBoxMakerReadOnlyColumn loBom, "ITEM_CODE"
    FormatBoxMakerReadOnlyColumn loBom, "ROW"
    ProtectShippingWorksheetForReadOnlyColumns loBom.Parent
End Sub

Private Sub EnsureShippingWorksheetEditable(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
End Sub

Private Sub ProtectShippingWorksheetForReadOnlyColumns(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If ws.ProtectContents Then ws.Unprotect
    On Error GoTo 0
End Sub

Private Sub RepairMisheadedBoxBomColumnsShipping(ByVal loBom As ListObject)
    If loBom Is Nothing Then Exit Sub

    RenameBoxBomColumnIfMissingShipping loBom, "ITEM", "Status"
    RenameBoxBomColumnIfMissingShipping loBom, "ITEM_CODE", "Effective From"
    RenameBoxBomColumnIfMissingShipping loBom, "ROW", "Effective To"
    RenameBoxBomColumnIfMissingShipping loBom, "QUANTITY", "Retired At"
    RenameBoxBomColumnIfMissingShipping loBom, "UOM", "Updated At"
    RenameBoxBomColumnIfMissingShipping loBom, "LOCATION", "Updated By"
End Sub

Private Sub RenameBoxBomColumnIfMissingShipping(ByVal loBom As ListObject, _
                                                ByVal wantedName As String, _
                                                ByVal wrongName As String)
    Dim wrongIdx As Long

    If loBom Is Nothing Then Exit Sub
    If ColumnIndex(loBom, wantedName) > 0 Then Exit Sub
    wrongIdx = ColumnIndex(loBom, wrongName)
    If wrongIdx = 0 Then Exit Sub
    On Error Resume Next
    loBom.ListColumns(wrongIdx).Name = wantedName
    On Error GoTo 0
End Sub

Private Sub FillBlankBoxBomVersionShipping(ByVal loBom As ListObject)
    Dim cVersion As Long
    Dim r As Long
    Dim lastVersion As String

    If loBom Is Nothing Then Exit Sub
    If loBom.DataBodyRange Is Nothing Then Exit Sub
    cVersion = ColumnIndex(loBom, "Version")
    If cVersion = 0 Then Exit Sub

    For r = 1 To loBom.ListRows.Count
        If Trim$(NzStr(loBom.DataBodyRange.Cells(r, cVersion).Value)) <> "" Then
            lastVersion = Trim$(NzStr(loBom.DataBodyRange.Cells(r, cVersion).Value))
        ElseIf BoxBomRowHasComponentDataShipping(loBom, r) Then
            If lastVersion = "" Then lastVersion = "v1"
            loBom.DataBodyRange.Cells(r, cVersion).Value = lastVersion
        End If
    Next r
End Sub

Private Function BoxBomRowHasComponentDataShipping(ByVal loBom As ListObject, ByVal rowIndex As Long) As Boolean
    Dim cItem As Long
    Dim cRow As Long
    Dim cQty As Long

    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > loBom.ListRows.Count Then Exit Function

    cItem = ColumnIndex(loBom, COL_BOXBOM_ITEM)
    cRow = ColumnIndex(loBom, "ROW")
    cQty = ColumnIndex(loBom, "QUANTITY")
    If cItem > 0 Then
        If Trim$(NzStr(loBom.DataBodyRange.Cells(rowIndex, cItem).Value)) <> "" Then BoxBomRowHasComponentDataShipping = True
    End If
    If cRow > 0 Then
        If NzLng(loBom.DataBodyRange.Cells(rowIndex, cRow).Value) <> 0 Then BoxBomRowHasComponentDataShipping = True
    End If
    If cQty > 0 Then
        If NzDbl(loBom.DataBodyRange.Cells(rowIndex, cQty).Value) <> 0 Then BoxBomRowHasComponentDataShipping = True
    End If
End Function

Private Function DisplayedBoxBomHasVersionRowsShipping(ByVal loBom As ListObject, ByVal versionLabel As String) As Boolean
    Dim cVersion As Long
    Dim r As Long
    Dim rowVersion As String

    If loBom Is Nothing Then Exit Function
    If loBom.DataBodyRange Is Nothing Then Exit Function
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If versionLabel = "" Then Exit Function
    cVersion = ColumnIndex(loBom, "Version")
    If cVersion = 0 Then Exit Function

    For r = 1 To loBom.ListRows.Count
        If Not BoxBomRowHasComponentDataShipping(loBom, r) Then GoTo NextRow
        rowVersion = NormalizeBoxBomVersionLabelShipping(loBom.DataBodyRange.Cells(r, cVersion).Value)
        If StrComp(rowVersion, versionLabel, vbTextCompare) = 0 Then
            DisplayedBoxBomHasVersionRowsShipping = True
            Exit Function
        End If
NextRow:
    Next r
End Function

Private Sub SortBoxBomByVersionShipping(ByVal loBom As ListObject)
    Dim cVersion As Long

    If loBom Is Nothing Then Exit Sub
    If loBom.DataBodyRange Is Nothing Then Exit Sub
    cVersion = ColumnIndex(loBom, "Version")
    If cVersion = 0 Then Exit Sub

    On Error Resume Next
    With loBom.Sort
        .SortFields.Clear
        .SortFields.Add Key:=loBom.ListColumns(cVersion).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With
    On Error GoTo 0
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

Private Sub RemoveListColumnIfExistsShipping(ByVal lo As ListObject, ByVal columnName As String)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    idx = ColumnIndex(lo, columnName)
    If idx = 0 Then Exit Sub

    On Error Resume Next
    lo.ListColumns(idx).Delete
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
        Case "VERSION", "ITEM", "BOXBOM", "ITEM_CODE", "ROW", "QUANTITY", "CURRENT INV", "UOM", "LOCATION", "DESCRIPTION"
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
    EnsureShippingWorksheetEditable lo.Parent
    idx = ColumnIndex(lo, colName)
    If idx > 0 Then lo.ListColumns(idx).Delete
End Sub

Private Sub EnsureColumnExists(lo As ListObject, colName As String, Optional afterColumn As String = "")
    If lo Is Nothing Then Exit Sub
    If ColumnIndex(lo, colName) > 0 Then Exit Sub
    EnsureShippingWorksheetEditable lo.Parent
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
        "BomVersion", "BomVersionLabel", "IsActive", "EffectiveFromUTC", "EffectiveToUTC", "RetiredAtUTC", _
        "ComponentRow", "ComponentItemCode", "ComponentItem", "ComponentQty", "ComponentUOM", "ComponentLocation", "ComponentDescription", _
        "UpdatedAtUTC", "UpdatedBy")
End Function

Private Function ShippingBomPackageTableHeaders() As Variant
    ShippingBomPackageTableHeaders = Array( _
        "ComponentRow", "ComponentItemCode", "ComponentItem", "ComponentQty", "ComponentUOM", "ComponentLocation", "ComponentDescription", _
        "UpdatedAtUTC", "UpdatedBy")
End Function

Private Function SaveShippingBomToRuntime(ByVal operatorWb As Workbook, _
                                          ByVal packageRow As Long, _
                                          ByVal packageItem As String, _
                                          ByVal packageUom As String, _
                                          ByVal packageLocation As String, _
                                          ByVal packageDescription As String, _
                                          ByVal components As Collection, _
                                          ByRef report As String, _
                                          Optional ByVal replaceBomVersion As Long = 0, _
                                          Optional ByVal forceNewVersion As Boolean = False) As Boolean
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
    Dim bomVersion As Long
    Dim existingVersion As Long

    mLastSavedShippingBomVersion = 0
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

    updatedAt = Now
    updatedBy = modRoleEventWriter.ResolveCurrentUserId()

    If Not forceNewVersion Then
        existingVersion = MatchingShippingBomVersion(loBom, packageRow, components)
        If existingVersion > 0 And (replaceBomVersion <= 0 Or existingVersion = replaceBomVersion) Then
            mLastSavedShippingBomVersion = existingVersion
            SaveShippingBomToRuntime = True
            report = "Shipping BOM unchanged: " & packageItem & " v" & CStr(existingVersion) & " already matches the current BoxBOM."
            GoTo CleanExit
        End If
        If existingVersion > 0 And replaceBomVersion > 0 And existingVersion <> replaceBomVersion Then
            mLastSavedShippingBomVersion = existingVersion
            SaveShippingBomToRuntime = True
            report = "Shipping BOM not duplicated: current BoxBOM already matches " & packageItem & " v" & CStr(existingVersion) & "."
            GoTo CleanExit
        End If
    End If

    If replaceBomVersion > 0 Then
        bomVersion = replaceBomVersion
        DeleteShippingBomPackageVersionRows loBom, packageRow, bomVersion
        DeleteShippingBomPackageTable wbBom, packageRow, bomVersion
    Else
        bomVersion = NextShippingBomVersion(loBom, packageRow)
    End If
    mLastSavedShippingBomVersion = bomVersion

    For i = 1 To components.Count
        info = components(i)
        Set lr = loBom.ListRows.Add
        SetTableCellShipping loBom, lr.Index, "PackageRow", packageRow
        SetTableCellShipping loBom, lr.Index, "PackageItem", packageItem
        SetTableCellShipping loBom, lr.Index, "PackageUOM", packageUom
        SetTableCellShipping loBom, lr.Index, "PackageLocation", packageLocation
        SetTableCellShipping loBom, lr.Index, "PackageDescription", packageDescription
        SetTableCellShipping loBom, lr.Index, "BomVersion", bomVersion
        SetTableCellShipping loBom, lr.Index, "BomVersionLabel", "v" & CStr(bomVersion)
        SetTableCellShipping loBom, lr.Index, "IsActive", True
        SetTableCellShipping loBom, lr.Index, "EffectiveFromUTC", updatedAt
        SetTableCellShipping loBom, lr.Index, "EffectiveToUTC", vbNullString
        SetTableCellShipping loBom, lr.Index, "RetiredAtUTC", vbNullString
        SetTableCellShipping loBom, lr.Index, "ComponentRow", NzLng(info(1))
        SetTableCellShipping loBom, lr.Index, "ComponentQty", NzDbl(info(2))
        SetTableCellShipping loBom, lr.Index, "ComponentUOM", NzStr(info(3))
        If UBound(info) >= 4 Then SetTableCellShipping loBom, lr.Index, "ComponentItem", NzStr(info(4))
        If UBound(info) >= 5 Then SetTableCellShipping loBom, lr.Index, "ComponentLocation", NzStr(info(5))
        If UBound(info) >= 6 Then SetTableCellShipping loBom, lr.Index, "ComponentDescription", NzStr(info(6))
        If UBound(info) >= 7 Then SetTableCellShipping loBom, lr.Index, "ComponentItemCode", NzStr(info(7))
        SetTableCellShipping loBom, lr.Index, "UpdatedAtUTC", updatedAt
        SetTableCellShipping loBom, lr.Index, "UpdatedBy", updatedBy
    Next i

    WriteShippingBomPackageTable wbBom, packageRow, bomVersion, packageItem, components, updatedAt, updatedBy
    wbBom.Save
    SaveShippingBomToRuntime = True
    If replaceBomVersion > 0 Then
        report = "Shipping BOM runtime updated: " & wbBom.FullName & " (" & packageItem & " v" & CStr(bomVersion) & " edited)"
    Else
        report = "Shipping BOM runtime updated: " & wbBom.FullName & " (" & packageItem & " v" & CStr(bomVersion) & " active)"
    End If

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "SaveShippingBomToRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Private Function BoxBuilderStatusIsActive(ByVal statusText As String) As Boolean
    statusText = UCase$(Trim$(statusText))
    BoxBuilderStatusIsActive = (statusText <> "RETIRED")
End Function

Private Function SetShippingBomVersionStatusInRuntime(ByVal operatorWb As Workbook, _
                                                      ByVal packageRow As Long, _
                                                      ByVal bomVersion As Long, _
                                                      ByVal isActive As Boolean, _
                                                      ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim openedTransient As Boolean
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim cActive As Long
    Dim cEffectiveFrom As Long
    Dim i As Long
    Dim rowVersion As Long
    Dim updatedAt As Date
    Dim updatedBy As String
    Dim updatedRows As Long
    Dim changedRows As Long
    Dim currentActive As Boolean

    report = ""
    If packageRow <= 0 Or bomVersion <= 0 Then Exit Function

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "A connected warehouse target is required before updating Shipping BOM status."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        report = "Selected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbBom Is Nothing Then Exit Function
    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit

    cPackageRow = ColumnIndex(loBom, "PackageRow")
    cVersion = ColumnIndex(loBom, "BomVersion")
    cActive = ColumnIndex(loBom, "IsActive")
    cEffectiveFrom = ColumnIndex(loBom, "EffectiveFromUTC")
    If cPackageRow = 0 Then GoTo CleanExit
    If loBom.DataBodyRange Is Nothing Then GoTo CleanExit

    updatedAt = Now
    updatedBy = modRoleEventWriter.ResolveCurrentUserId()

    For i = 1 To loBom.ListRows.Count
        If NzLng(loBom.DataBodyRange.Cells(i, cPackageRow).Value) = packageRow Then
            If cVersion > 0 Then rowVersion = NzLng(loBom.DataBodyRange.Cells(i, cVersion).Value) Else rowVersion = 1
            If rowVersion <= 0 Then rowVersion = 1
            If rowVersion = bomVersion Then
                currentActive = True
                If cActive > 0 Then currentActive = ShippingBomActiveValue(loBom.DataBodyRange.Cells(i, cActive).Value)
                updatedRows = updatedRows + 1
                If currentActive <> isActive Then changedRows = changedRows + 1

                SetTableCellShipping loBom, i, "IsActive", isActive
                If isActive Then
                    If cEffectiveFrom = 0 Or Trim$(NzStr(loBom.DataBodyRange.Cells(i, cEffectiveFrom).Value)) = "" Then
                        SetTableCellShipping loBom, i, "EffectiveFromUTC", updatedAt
                    End If
                    SetTableCellShipping loBom, i, "EffectiveToUTC", vbNullString
                    SetTableCellShipping loBom, i, "RetiredAtUTC", vbNullString
                Else
                    SetTableCellShipping loBom, i, "EffectiveToUTC", updatedAt
                    SetTableCellShipping loBom, i, "RetiredAtUTC", updatedAt
                End If
                SetTableCellShipping loBom, i, "UpdatedAtUTC", updatedAt
                SetTableCellShipping loBom, i, "UpdatedBy", updatedBy
            End If
        End If
    Next i

    If updatedRows <= 0 Then
        report = "No Shipping BOM rows were found for ROW " & CStr(packageRow) & " v" & CStr(bomVersion) & "."
        GoTo CleanExit
    End If

    wbBom.Save
    SetShippingBomVersionStatusInRuntime = True
    If changedRows > 0 Then
        report = "Shipping BOM status updated: ROW " & CStr(packageRow) & " v" & CStr(bomVersion) & " is " & IIf(isActive, "Active", "Retired") & "."
    End If

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "SetShippingBomVersionStatusInRuntime failed: " & Err.Description
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

    Set GetShippingBomViewTable = FindListObjectByNameShipping(wb, TABLE_SHIPPING_BOM_VIEW)
    If Not GetShippingBomViewTable Is Nothing Then Exit Function

    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    Set GetShippingBomViewTable = GetListObject(ws, TABLE_SHIPPING_BOM_VIEW)
End Function

Public Function ShippingBomViewTableExistsForWorkbookForTest(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    On Error Resume Next
    Set wb = Application.Workbooks(workbookName)
    On Error GoTo 0
    If wb Is Nothing Then Exit Function
    ShippingBomViewTableExistsForWorkbookForTest = Not GetShippingBomViewTable(wb) Is Nothing
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
        HideWorkbookWindowsShipping wb
        Set OpenShippingBomWorkbook = wb
        Exit Function
    End If

    If Len(Dir$(targetPath)) > 0 Then
        Set wb = OpenWorkbookHiddenShipping(targetPath, False, openedTransient)
    ElseIf createIfMissing Then
        EnsureFolderRecursiveShipping GetParentFolderShipping(targetPath)
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        HideWorkbookWindowsShipping wb
        wb.Worksheets(1).Name = SHEET_BOM
        If EnsureShippingBomSchema(wb, report) Is Nothing Then
            CloseWorkbookNoSaveShipping wb
            Exit Function
        End If
        wb.SaveAs Filename:=targetPath, FileFormat:=50
        openedTransient = False
    Else
        report = "Shipping BOM runtime workbook was not found: " & targetPath
        Exit Function
    End If

    Set OpenShippingBomWorkbook = wb
    Exit Function

FailSoft:
    report = "OpenShippingBomWorkbook failed: " & Err.Description
End Function

Private Function OpenShippingReservationsWorkbook(ByVal warehouseId As String, _
                                                  ByVal rootPath As String, _
                                                  ByVal createIfMissing As Boolean, _
                                                  ByRef openedTransient As Boolean, _
                                                  ByRef report As String) As Workbook
    On Error GoTo FailSoft

    Dim targetPath As String
    Dim wb As Workbook

    targetPath = ShippingReservationsWorkbookPath(warehouseId, rootPath)
    If targetPath = "" Then
        report = "Shipping reservations workbook path could not be resolved."
        Exit Function
    End If

    Set wb = FindOpenWorkbookByFullNameShipping(targetPath)
    If Not wb Is Nothing Then
        HideWorkbookWindowsShipping wb
        Set OpenShippingReservationsWorkbook = wb
        Exit Function
    End If

    If Len(Dir$(targetPath)) > 0 Then
        Set wb = OpenWorkbookHiddenShipping(targetPath, False, openedTransient, False)
    ElseIf createIfMissing Then
        EnsureFolderRecursiveShipping GetParentFolderShipping(targetPath)
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        HideWorkbookWindowsShipping wb
        wb.Worksheets(1).Name = SHEET_SHIPPING_RESERVATIONS
        If EnsureShippingReservationsSchema(wb, report) Is Nothing Then
            CloseWorkbookNoSaveShipping wb
            Exit Function
        End If
        wb.SaveAs Filename:=targetPath, FileFormat:=50
        openedTransient = True
    Else
        report = "Shipping reservations runtime workbook was not found: " & targetPath
        Exit Function
    End If

    Set OpenShippingReservationsWorkbook = wb
    Exit Function

FailSoft:
    report = "OpenShippingReservationsWorkbook failed: " & Err.Description
End Function

Private Function OpenWorkbookHiddenShipping(ByVal workbookPath As String, _
                                            ByVal readOnly As Boolean, _
                                            ByRef openedTransient As Boolean, _
                                            Optional ByVal keepOpen As Boolean = True) As Workbook
    On Error GoTo FailSoft

    Dim wb As Workbook
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    openedTransient = False
    If Trim$(workbookPath) = "" Then Exit Function

    Set wb = FindOpenWorkbookByFullNameShipping(workbookPath)
    If Not wb Is Nothing Then
        HideWorkbookWindowsShipping wb
        If Not keepOpen Then openedTransient = (readOnly And wb.ReadOnly)
        Set OpenWorkbookHiddenShipping = wb
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Application.Workbooks.Open(Filename:=workbookPath, _
                                        UpdateLinks:=False, _
                                        ReadOnly:=readOnly, _
                                        AddToMru:=False, _
                                        IgnoreReadOnlyRecommended:=True, _
                                        Notify:=False)
    HideWorkbookWindowsShipping wb

    ' Keep stable runtime/read-model workbooks open for speed, but allow callers
    ' that read Data.Inventory to close it before processor write-lock attempts.
    openedTransient = Not keepOpen
    Set OpenWorkbookHiddenShipping = wb

CleanExit:
    On Error Resume Next
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    Exit Function

FailSoft:
    Resume CleanExit
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

Private Function EnsureShippingReservationsSchema(ByVal wb As Workbook, ByRef report As String) As ListObject
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim startCell As Range
    Dim dataRange As Range

    If wb Is Nothing Then Exit Function
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPPING_RESERVATIONS)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_SHIPPING_RESERVATIONS
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_SHIPPING_RESERVATIONS)
    On Error GoTo FailSoft

    headers = ShippingReservationHeaders()
    If lo Is Nothing Then
        Set startCell = ws.Range("A1")
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = TABLE_SHIPPING_RESERVATIONS
        If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureColumnExists lo, CStr(headers(i))
    Next i
    Set EnsureShippingReservationsSchema = lo
    Exit Function

FailSoft:
    report = "EnsureShippingReservationsSchema failed: " & Err.Description
End Function

Private Function ShippingReservationHeaders() As Variant
    ShippingReservationHeaders = Array( _
        "ReservationID", "Status", "WarehouseId", "StationId", "UserId", _
        "LineID", "EventID", "RefNumber", "ItemName", "PackageRow", _
        "Version", "Qty", "UOM", "Location", "SourceWorkbook", _
        "CreatedAtUTC", "UpdatedAtUTC", "ReleasedAtUTC", "CompletedAtUTC", "ReleaseEventID")
End Function

Public Function ShipmentsFormReservationKey(ByVal packageRow As Long, ByVal versionLabel As String) As String
    versionLabel = NormalizeBoxBomVersionLabelShipping(versionLabel)
    If packageRow <= 0 Or versionLabel = "" Then Exit Function
    ShipmentsFormReservationKey = CStr(packageRow) & "|" & LCase$(versionLabel)
End Function

Public Function ShipmentsFormLoadNasReservationTotals() As Object
    Dim totals As Object
    Dim report As String

    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare
    Set ShipmentsFormLoadNasReservationTotals = totals
    Set totals = LoadActiveShippingReservationTotals(report)
    If Not totals Is Nothing Then Set ShipmentsFormLoadNasReservationTotals = totals
End Function

Public Function ValidateShippingReservationTotalsFromTableForTest(ByVal lo As ListObject, ByVal warehouseId As String) As Object
    Dim report As String

    Set ValidateShippingReservationTotalsFromTableForTest = BuildActiveShippingReservationTotalsFromTable(lo, warehouseId, report)
End Function

Public Function ValidateShippingReservationTotalsFromTableWithLocalLinesForTest(ByVal lo As ListObject, _
                                                                                ByVal warehouseId As String, _
                                                                                ByVal localSourceWorkbook As String, _
                                                                                ByVal activeLineIdsCsv As String) As Object
    Dim report As String
    Dim activeLineIds As Object
    Dim parts As Variant
    Dim i As Long
    Dim lineId As String

    Set activeLineIds = CreateObject("Scripting.Dictionary")
    activeLineIds.CompareMode = vbTextCompare
    parts = Split(activeLineIdsCsv, ",")
    For i = LBound(parts) To UBound(parts)
        lineId = Trim$(NzStr(parts(i)))
        If lineId <> "" Then activeLineIds(lineId) = True
    Next i
    Set ValidateShippingReservationTotalsFromTableWithLocalLinesForTest = BuildActiveShippingReservationTotalsFromTable(lo, _
                                                                                                                        warehouseId, _
                                                                                                                        report, _
                                                                                                                        localSourceWorkbook, _
                                                                                                                        activeLineIds, _
                                                                                                                        "S31")
End Function

Private Function LoadActiveShippingReservationTotals(ByRef report As String) As Object
    On Error GoTo FailSoft

    Dim wb As Workbook
    Dim lo As ListObject
    Dim openedTransient As Boolean
    Dim totals As Object
    Dim warehouseId As String
    Dim localSourceWorkbook As String
    Dim activeLineIds As Object

    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare
    Set LoadActiveShippingReservationTotals = totals

    If Not OpenCurrentShippingReservationsWorkbook(False, wb, lo, openedTransient, report) Then GoTo CleanExit
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then GoTo CleanExit

    warehouseId = CurrentShippingWarehouseIdForLocalState()
    localSourceWorkbook = CurrentShippingOperatorWorkbookFullName()
    Set activeLineIds = ActiveShipmentLineIdsForCurrentOperatorWorkbook()
    Set totals = BuildActiveShippingReservationTotalsFromTable(lo, _
                                                               warehouseId, _
                                                               report, _
                                                               localSourceWorkbook, _
                                                               activeLineIds, _
                                                               CurrentShippingStationIdForLocalState())
    If Not totals Is Nothing Then Set LoadActiveShippingReservationTotals = totals

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wb
    Exit Function

FailSoft:
    report = "LoadActiveShippingReservationTotals failed: " & Err.Description
    Resume CleanExit
End Function

Private Function BuildActiveShippingReservationTotalsFromTable(ByVal lo As ListObject, _
                                                              ByVal warehouseId As String, _
                                                              ByRef report As String, _
                                                              Optional ByVal localSourceWorkbook As String = "", _
                                                              Optional ByVal activeLineIds As Object = Nothing, _
                                                              Optional ByVal localStationId As String = "") As Object
    On Error GoTo FailSoft

    Dim totals As Object
    Dim r As Long
    Dim cStatus As Long
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim cQty As Long
    Dim cLineId As Long
    Dim cSourceWorkbook As Long
    Dim cStationId As Long
    Dim key As String
    Dim qtyVal As Double
    Dim lineId As String
    Dim sourceWorkbook As String
    Dim reservationStationId As String
    Dim localReservation As Boolean

    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare
    Set BuildActiveShippingReservationTotalsFromTable = totals
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then GoTo CleanExit

    cStatus = ColumnIndex(lo, "Status")
    cPackageRow = ColumnIndex(lo, "PackageRow")
    cVersion = ColumnIndex(lo, "Version")
    cQty = ColumnIndex(lo, "Qty")
    cLineId = ColumnIndex(lo, "LineID")
    cSourceWorkbook = ColumnIndex(lo, "SourceWorkbook")
    cStationId = ColumnIndex(lo, "StationId")
    If cStatus = 0 Or cPackageRow = 0 Or cVersion = 0 Or cQty = 0 Then GoTo CleanExit

    For r = 1 To lo.ListRows.Count
        If StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cStatus).Value)), SHIP_RESERVATION_ACTIVE, vbTextCompare) = 0 Then
            If cLineId > 0 Then
                lineId = Trim$(NzStr(lo.DataBodyRange.Cells(r, cLineId).Value))
                If PersistentSentShipmentLineIdExistsForWarehouse(lineId, warehouseId) Then GoTo NextReservation
                If Not activeLineIds Is Nothing Then
                    sourceWorkbook = ""
                    If cSourceWorkbook > 0 Then sourceWorkbook = NormalizeFolderPathShipping(Trim$(NzStr(lo.DataBodyRange.Cells(r, cSourceWorkbook).Value)))
                    reservationStationId = ""
                    If cStationId > 0 Then reservationStationId = Trim$(NzStr(lo.DataBodyRange.Cells(r, cStationId).Value))
                    localReservation = (sourceWorkbook = "" _
                                        Or StrComp(sourceWorkbook, NormalizeFolderPathShipping(localSourceWorkbook), vbTextCompare) = 0 _
                                        Or reservationStationId = "" _
                                        Or (Trim$(localStationId) <> "" And StrComp(reservationStationId, Trim$(localStationId), vbTextCompare) = 0))
                    If localReservation Then
                        If lineId = "" Or Not activeLineIds.Exists(lineId) Then GoTo NextReservation
                    End If
                End If
            End If
            key = ShipmentsFormReservationKey(NzLng(lo.DataBodyRange.Cells(r, cPackageRow).Value), NzStr(lo.DataBodyRange.Cells(r, cVersion).Value))
            qtyVal = NzDbl(lo.DataBodyRange.Cells(r, cQty).Value)
            If key <> "" And qtyVal > 0 Then
                If totals.Exists(key) Then
                    totals(key) = NzDbl(totals(key)) + qtyVal
                Else
                    totals.Add key, qtyVal
                End If
            End If
        End If
NextReservation:
    Next r

CleanExit:
    Exit Function

FailSoft:
    report = "BuildActiveShippingReservationTotalsFromTable failed: " & Err.Description
    Resume CleanExit
End Function

Private Function CurrentShippingWarehouseIdForLocalState() As String
    Dim target As WarehouseTarget

    CurrentShippingWarehouseIdForLocalState = Trim$(modConfig.GetWarehouseId())
    If CurrentShippingWarehouseIdForLocalState <> "" Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then CurrentShippingWarehouseIdForLocalState = Trim$(target.WarehouseId)
    If CurrentShippingWarehouseIdForLocalState = "" Then CurrentShippingWarehouseIdForLocalState = "default"
End Function

Private Function CurrentShippingOperatorWorkbookFullName() As String
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then CurrentShippingOperatorWorkbookFullName = ActiveWorkbook.FullName
    On Error GoTo 0
End Function

Private Function CurrentShippingStationIdForLocalState() As String
    Dim target As WarehouseTarget
    Dim warehouseId As String

    warehouseId = CurrentShippingWarehouseIdForLocalState()
    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then
        If warehouseId = "" Or StrComp(Trim$(target.WarehouseId), warehouseId, vbTextCompare) = 0 Then
            CurrentShippingStationIdForLocalState = Trim$(target.StationId)
        End If
    End If
    If CurrentShippingStationIdForLocalState = "" Then CurrentShippingStationIdForLocalState = Trim$(modConfig.GetStationId())
End Function

Private Function ActiveShipmentLineIdsForCurrentOperatorWorkbook() As Object
    On Error GoTo CleanExit

    Dim result As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim cLine As Long
    Dim cItem As Long
    Dim cQty As Long
    Dim r As Long
    Dim lineId As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set ActiveShipmentLineIdsForCurrentOperatorWorkbook = result

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    Set ws = WorkbookSheetExistsShipping(wb, SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Function
    Set lo = GetListObject(ws, TABLE_SHIPMENTS)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    cLine = ColumnIndex(lo, COL_SHIPMENT_LINE_ID)
    cItem = ColumnIndex(lo, "ITEMS")
    cQty = ColumnIndex(lo, "QUANTITY")
    If cLine = 0 Then Exit Function

    For r = 1 To lo.ListRows.Count
        lineId = Trim$(NzStr(lo.DataBodyRange.Cells(r, cLine).Value))
        If lineId <> "" Then
            If cItem = 0 Or cQty = 0 Then
                result(lineId) = True
            ElseIf Trim$(NzStr(lo.DataBodyRange.Cells(r, cItem).Value)) <> "" Or NzDbl(lo.DataBodyRange.Cells(r, cQty).Value) <> 0 Then
                result(lineId) = True
            End If
        End If
    Next r

CleanExit:
End Function

Private Function ActiveNasShippingReservationQty(ByVal packageRow As Long, ByVal versionLabel As String) As Double
    Dim totals As Object
    Dim key As String
    Dim report As String

    Set totals = LoadActiveShippingReservationTotals(report)
    If totals Is Nothing Then Exit Function
    key = ShipmentsFormReservationKey(packageRow, versionLabel)
    If key <> "" Then
        If totals.Exists(key) Then ActiveNasShippingReservationQty = NzDbl(totals(key))
    End If
End Function

Private Function OpenCurrentShippingReservationsWorkbook(ByVal createIfMissing As Boolean, _
                                                        ByRef wb As Workbook, _
                                                        ByRef lo As ListObject, _
                                                        ByRef openedTransient As Boolean, _
                                                        ByRef report As String) As Boolean
    Dim target As WarehouseTarget
    Dim warehouseId As String
    Dim rootPath As String

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "A connected warehouse target is required before locking shipment inventory."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        report = "Selected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    Set wb = OpenShippingReservationsWorkbook(warehouseId, rootPath, createIfMissing, openedTransient, report)
    If wb Is Nothing Then Exit Function
    If wb.ReadOnly Then
        If createIfMissing Then
            report = "Shipping reservations workbook is read-only or locked by another Excel session."
            GoTo CleanFail
        End If
        Set lo = FindListObjectByNameShipping(wb, TABLE_SHIPPING_RESERVATIONS)
        OpenCurrentShippingReservationsWorkbook = Not lo Is Nothing
        Exit Function
    End If
    Set lo = EnsureShippingReservationsSchema(wb, report)
    If lo Is Nothing Then GoTo CleanFail

    OpenCurrentShippingReservationsWorkbook = True
    Exit Function

CleanFail:
    If openedTransient Then CloseWorkbookNoSaveShipping wb
End Function

Private Function UpsertShippingReservationForRow(ByVal loShip As ListObject, _
                                                 ByVal rowIndex As Long, _
                                                 ByVal reserveEventId As String, _
                                                 ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim wb As Workbook
    Dim lo As ListObject
    Dim openedTransient As Boolean
    Dim target As WarehouseTarget
    Dim lr As ListRow
    Dim lineId As String
    Dim sourceWorkbook As String
    Dim nowText As String

    If loShip Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > loShip.ListRows.Count Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "A connected warehouse target is required before locking shipment inventory."
        Exit Function
    End If

    If Not OpenCurrentShippingReservationsWorkbook(True, wb, lo, openedTransient, report) Then Exit Function
    lineId = Trim$(ShipmentRowText(loShip, rowIndex, COL_SHIPMENT_LINE_ID))
    If lineId = "" Then lineId = EnsureShipmentLineId(loShip, rowIndex)
    Set lr = FindShippingReservationRow(lo, reserveEventId, lineId)
    If lr Is Nothing Then Set lr = lo.ListRows.Add

    sourceWorkbook = ""
    On Error Resume Next
    sourceWorkbook = loShip.Parent.Parent.FullName
    On Error GoTo FailSoft
    nowText = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    WriteValue lr, "ReservationID", ShippingReservationId(lineId, reserveEventId)
    WriteValue lr, "Status", SHIP_RESERVATION_ACTIVE
    WriteValue lr, "WarehouseId", Trim$(target.WarehouseId)
    WriteValue lr, "StationId", Trim$(target.StationId)
    WriteValue lr, "UserId", modRoleEventWriter.ResolveCurrentUserId()
    WriteValue lr, "LineID", lineId
    WriteValue lr, "EventID", Trim$(reserveEventId)
    WriteValue lr, "RefNumber", ShipmentRowText(loShip, rowIndex, "REF_NUMBER")
    WriteValue lr, "ItemName", ShipmentRowText(loShip, rowIndex, "ITEMS")
    WriteValue lr, "PackageRow", NzLng(ShipmentRowText(loShip, rowIndex, "ROW"))
    WriteValue lr, "Version", NormalizeBoxBomVersionLabelShipping(ShipmentRowText(loShip, rowIndex, "DESCRIPTION"))
    WriteValue lr, "Qty", NzDbl(ShipmentRowText(loShip, rowIndex, "QUANTITY"))
    WriteValue lr, "UOM", ShipmentRowText(loShip, rowIndex, "UOM")
    WriteValue lr, "Location", ShipmentRowText(loShip, rowIndex, "LOCATION")
    WriteValue lr, "SourceWorkbook", sourceWorkbook
    If Trim$(NzStr(ReadListRowValue(lr, "CreatedAtUTC"))) = "" Then WriteValue lr, "CreatedAtUTC", nowText
    WriteValue lr, "UpdatedAtUTC", nowText
    WriteValue lr, "ReleasedAtUTC", vbNullString
    WriteValue lr, "CompletedAtUTC", vbNullString
    WriteValue lr, "ReleaseEventID", vbNullString

    wb.Save
    UpsertShippingReservationForRow = True

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wb
    Exit Function

FailSoft:
    report = "Unable to write shipping reservation lock: " & Err.Description
    Resume CleanExit
End Function

Private Function MarkShippingReservationRows(ByVal loShip As ListObject, _
                                             ByVal rowIndexes As Variant, _
                                             ByVal statusText As String, _
                                             ByVal releaseEventId As String, _
                                             ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim wb As Workbook
    Dim lo As ListObject
    Dim openedTransient As Boolean
    Dim i As Long
    Dim rowIndex As Long
    Dim reserveEventId As String
    Dim lineId As String
    Dim lr As ListRow
    Dim nowText As String

    MarkShippingReservationRows = True
    If loShip Is Nothing Then Exit Function
    If IsEmpty(rowIndexes) Then Exit Function
    If Not OpenCurrentShippingReservationsWorkbook(False, wb, lo, openedTransient, report) Then Exit Function
    If lo Is Nothing Then GoTo CleanExit

    nowText = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex < 1 Or rowIndex > loShip.ListRows.Count Then GoTo NextRow
        reserveEventId = Trim$(ShipmentRowText(loShip, rowIndex, COL_SHIPMENT_RESERVE_EVENT_ID))
        lineId = Trim$(ShipmentRowText(loShip, rowIndex, COL_SHIPMENT_LINE_ID))
        Set lr = FindShippingReservationRow(lo, reserveEventId, lineId)
        If lr Is Nothing Then GoTo NextRow
        WriteValue lr, "Status", statusText
        WriteValue lr, "UpdatedAtUTC", nowText
        If StrComp(statusText, SHIP_RESERVATION_RELEASED, vbTextCompare) = 0 Then
            WriteValue lr, "ReleasedAtUTC", nowText
            WriteValue lr, "ReleaseEventID", Trim$(releaseEventId)
        ElseIf StrComp(statusText, SHIP_RESERVATION_COMPLETED, vbTextCompare) = 0 Then
            WriteValue lr, "CompletedAtUTC", nowText
        End If
NextRow:
    Next i
    wb.Save

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wb
    Exit Function

FailSoft:
    MarkShippingReservationRows = False
    report = "Unable to update shipping reservation lock: " & Err.Description
    Resume CleanExit
End Function

Private Function FindShippingReservationRow(ByVal lo As ListObject, ByVal reserveEventId As String, ByVal lineId As String) As ListRow
    Dim cEvent As Long
    Dim cLine As Long
    Dim r As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    reserveEventId = Trim$(reserveEventId)
    lineId = Trim$(lineId)
    cEvent = ColumnIndex(lo, "EventID")
    cLine = ColumnIndex(lo, "LineID")

    If reserveEventId <> "" And cEvent > 0 Then
        For r = 1 To lo.ListRows.Count
            If StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cEvent).Value)), reserveEventId, vbTextCompare) = 0 Then
                Set FindShippingReservationRow = lo.ListRows(r)
                Exit Function
            End If
        Next r
    End If

    If lineId <> "" And cLine > 0 Then
        For r = 1 To lo.ListRows.Count
            If StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cLine).Value)), lineId, vbTextCompare) = 0 Then
                Set FindShippingReservationRow = lo.ListRows(r)
                Exit Function
            End If
        Next r
    End If
End Function

Private Function ShippingReservationId(ByVal lineId As String, ByVal reserveEventId As String) As String
    If Trim$(reserveEventId) <> "" Then
        ShippingReservationId = Trim$(reserveEventId)
    ElseIf Trim$(lineId) <> "" Then
        ShippingReservationId = Trim$(lineId)
    Else
        ShippingReservationId = "SHIPRES-" & Format$(Now, "yyyymmddhhnnss")
    End If
End Function

Private Function ReadListRowValue(ByVal lr As ListRow, ByVal columnName As String) As Variant
    Dim idx As Long

    If lr Is Nothing Then Exit Function
    idx = ColumnIndex(lr.Parent, columnName)
    If idx = 0 Then Exit Function
    ReadListRowValue = lr.Range.Cells(1, idx).Value
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

Private Function DeleteShippingBomVersionFromRuntime(ByVal operatorWb As Workbook, _
                                                     ByVal packageRow As Long, _
                                                     ByVal bomVersion As Long, _
                                                     ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim openedTransient As Boolean
    Dim deletedRows As Long

    report = ""
    If packageRow <= 0 Or bomVersion <= 0 Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "A connected warehouse target is required before deleting a Shipping BOM version."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        report = "Selected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbBom Is Nothing Then Exit Function
    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit

    deletedRows = DeleteShippingBomPackageVersionRows(loBom, packageRow, bomVersion)
    DeleteShippingBomPackageTable wbBom, packageRow, bomVersion
    If deletedRows <= 0 Then
        report = "No Shipping BOM rows were found for ROW " & CStr(packageRow) & " v" & CStr(bomVersion) & "."
        GoTo CleanExit
    End If

    wbBom.Save
    DeleteShippingBomVersionFromRuntime = True
    report = "Deleted Shipping BOM ROW " & CStr(packageRow) & " v" & CStr(bomVersion) & " (" & CStr(deletedRows) & " row(s))."

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "DeleteShippingBomVersionFromRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Private Function DeleteShippingBomPackageFromRuntime(ByVal operatorWb As Workbook, _
                                                     ByVal packageRow As Long, _
                                                     ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim openedTransient As Boolean
    Dim deletedRows As Long

    report = ""
    If packageRow <= 0 Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "A connected warehouse target is required before deleting a Shipping BOM box."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        report = "Selected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbBom Is Nothing Then Exit Function
    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit

    deletedRows = CountShippingBomPackageRows(loBom, packageRow)
    DeleteShippingBomPackageRows loBom, packageRow
    DeleteShippingBomPackageTables wbBom, packageRow
    If deletedRows <= 0 Then
        report = "No Shipping BOM rows were found for ROW " & CStr(packageRow) & "."
        GoTo CleanExit
    End If

    wbBom.Save
    DeleteShippingBomPackageFromRuntime = True
    report = "Deleted Shipping BOM ROW " & CStr(packageRow) & " (" & CStr(deletedRows) & " row(s))."

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "DeleteShippingBomPackageFromRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Private Function ArchiveShippingBomPackageInRuntime(ByVal operatorWb As Workbook, _
                                                    ByVal packageRow As Long, _
                                                    ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim target As Object
    Dim warehouseId As String
    Dim rootPath As String
    Dim wbBom As Workbook
    Dim loBom As ListObject
    Dim openedTransient As Boolean
    Dim retiredAt As Date
    Dim retiredBy As String
    Dim activeRows As Long

    report = ""
    If packageRow <= 0 Then Exit Function

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "A connected warehouse target is required before archiving Shipping BOM designs."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    rootPath = NormalizeFolderPathShipping(target.RuntimeRoot)
    If warehouseId = "" Or rootPath = "" Then
        report = "Selected warehouse target is missing WarehouseId or RuntimeRoot."
        Exit Function
    End If

    Set wbBom = OpenShippingBomWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbBom Is Nothing Then Exit Function
    Set loBom = EnsureShippingBomSchema(wbBom, report)
    If loBom Is Nothing Then GoTo CleanExit

    activeRows = CountActiveShippingBomPackageRows(loBom, packageRow)
    If activeRows <= 0 Then
        report = "No active Shipping BOM designs were found for ROW " & CStr(packageRow) & "."
        ArchiveShippingBomPackageInRuntime = True
        GoTo CleanExit
    End If

    retiredAt = Now
    retiredBy = modRoleEventWriter.ResolveCurrentUserId()
    RetireActiveShippingBomPackageRows loBom, packageRow, retiredAt, retiredBy

    wbBom.Save
    report = "Archived Shipping BOM ROW " & CStr(packageRow) & ": " & CStr(activeRows) & _
             " active design row(s) retired. Existing inventory remains available."
    ArchiveShippingBomPackageInRuntime = True

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveShipping wbBom
    Exit Function

FailSoft:
    report = "ArchiveShippingBomPackageInRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Private Function DeleteShippingBomPackageVersionRows(ByVal lo As ListObject, _
                                                     ByVal packageRow As Long, _
                                                     ByVal bomVersion As Long) As Long
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim i As Long
    Dim rowVersion As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cPackageRow = ColumnIndex(lo, "PackageRow")
    cVersion = ColumnIndex(lo, "BomVersion")
    If cPackageRow = 0 Then Exit Function

    For i = lo.ListRows.Count To 1 Step -1
        If NzLng(lo.DataBodyRange.Cells(i, cPackageRow).Value) = packageRow Then
            If cVersion > 0 Then rowVersion = NzLng(lo.DataBodyRange.Cells(i, cVersion).Value) Else rowVersion = 1
            If rowVersion <= 0 Then rowVersion = 1
            If rowVersion = bomVersion Then
                lo.ListRows(i).Delete
                DeleteShippingBomPackageVersionRows = DeleteShippingBomPackageVersionRows + 1
            End If
        End If
    Next i
End Function

Private Function CountShippingBomPackageRows(ByVal lo As ListObject, ByVal packageRow As Long) As Long
    Dim cPackageRow As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cPackageRow = ColumnIndex(lo, "PackageRow")
    If cPackageRow = 0 Then Exit Function
    For i = 1 To lo.ListRows.Count
        If NzLng(lo.DataBodyRange.Cells(i, cPackageRow).Value) = packageRow Then CountShippingBomPackageRows = CountShippingBomPackageRows + 1
    Next i
End Function

Private Function CountActiveShippingBomPackageRows(ByVal lo As ListObject, ByVal packageRow As Long) As Long
    Dim cPackageRow As Long
    Dim cActive As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cPackageRow = ColumnIndex(lo, "PackageRow")
    cActive = ColumnIndex(lo, "IsActive")
    If cPackageRow = 0 Then Exit Function
    For i = 1 To lo.ListRows.Count
        If NzLng(lo.DataBodyRange.Cells(i, cPackageRow).Value) = packageRow Then
            If cActive = 0 Or ShippingBomActiveValue(lo.DataBodyRange.Cells(i, cActive).Value) Then
                CountActiveShippingBomPackageRows = CountActiveShippingBomPackageRows + 1
            End If
        End If
    Next i
End Function

Private Sub DeleteShippingBomPackageTable(ByVal wbBom As Workbook, _
                                          ByVal packageRow As Long, _
                                          ByVal bomVersion As Long)
    Dim ws As Worksheet
    Dim tableName As String

    If wbBom Is Nothing Then Exit Sub
    Set ws = WorkbookSheetExistsShipping(wbBom, SHEET_BOM_TABLES)
    If ws Is Nothing Then Exit Sub
    tableName = BomTableNameFromRowVersion(packageRow, bomVersion)
    On Error Resume Next
    ws.ListObjects(tableName).Delete
    On Error GoTo 0
End Sub

Private Sub DeleteShippingBomPackageTables(ByVal wbBom As Workbook, ByVal packageRow As Long)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim prefix As String
    Dim i As Long

    If wbBom Is Nothing Then Exit Sub
    Set ws = WorkbookSheetExistsShipping(wbBom, SHEET_BOM_TABLES)
    If ws Is Nothing Then Exit Sub
    prefix = BomTableNameFromRow(packageRow) & "_V"

    For i = ws.ListObjects.Count To 1 Step -1
        Set lo = ws.ListObjects(i)
        If StrComp(Left$(lo.Name, Len(prefix)), prefix, vbTextCompare) = 0 Then
            lo.Delete
        End If
    Next i
End Sub

Private Function NextShippingBomVersion(ByVal lo As ListObject, ByVal packageRow As Long) As Long
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim i As Long
    Dim rowVersion As Long
    Dim maxVersion As Long

    If lo Is Nothing Then
        NextShippingBomVersion = 1
        Exit Function
    End If
    cPackageRow = ColumnIndex(lo, "PackageRow")
    cVersion = ColumnIndex(lo, "BomVersion")
    If cPackageRow = 0 Or lo.DataBodyRange Is Nothing Then
        NextShippingBomVersion = 1
        Exit Function
    End If

    For i = 1 To lo.ListRows.Count
        If NzLng(lo.DataBodyRange.Cells(i, cPackageRow).Value) = packageRow Then
            If cVersion > 0 Then rowVersion = NzLng(lo.DataBodyRange.Cells(i, cVersion).Value) Else rowVersion = 0
            If rowVersion <= 0 Then rowVersion = 1
            If rowVersion > maxVersion Then maxVersion = rowVersion
        End If
    Next i

    NextShippingBomVersion = maxVersion + 1
    If NextShippingBomVersion <= 1 Then NextShippingBomVersion = 1
End Function

Private Function MatchingShippingBomVersion(ByVal lo As ListObject, _
                                            ByVal packageRow As Long, _
                                            ByVal components As Collection) As Long
    On Error GoTo CleanExit

    Dim targetSignature As String
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim dict As Object
    Dim i As Long
    Dim rowVersion As Long
    Dim key As Variant

    If lo Is Nothing Then Exit Function
    If components Is Nothing Then Exit Function
    If components.Count = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    cPackageRow = ColumnIndex(lo, "PackageRow")
    cVersion = ColumnIndex(lo, "BomVersion")
    If cPackageRow = 0 Then Exit Function

    targetSignature = ShippingBomComponentSignatureFromCollection(components)
    If targetSignature = "" Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To lo.ListRows.Count
        If NzLng(lo.DataBodyRange.Cells(i, cPackageRow).Value) = packageRow Then
            If cVersion > 0 Then rowVersion = NzLng(lo.DataBodyRange.Cells(i, cVersion).Value) Else rowVersion = 1
            If rowVersion <= 0 Then rowVersion = 1
            dict(CStr(rowVersion)) = rowVersion
        End If
    Next i

    For Each key In dict.Keys
        rowVersion = CLng(key)
        If ShippingBomComponentSignatureFromTableVersion(lo, packageRow, rowVersion) = targetSignature Then
            MatchingShippingBomVersion = rowVersion
            Exit Function
        End If
    Next key

CleanExit:
End Function

Private Function ShippingBomComponentSignatureFromCollection(ByVal components As Collection) As String
    On Error GoTo CleanExit

    Dim dict As Object
    Dim info As Variant
    Dim i As Long

    If components Is Nothing Then Exit Function
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To components.Count
        info = components(i)
        AddShippingBomSignaturePart dict, _
                                    NzLng(info(1)), _
                                    NzDbl(info(2)), _
                                    NzStr(info(3)), _
                                    NzStr(info(4)), _
                                    NzStr(info(5)), _
                                    NzStr(info(6))
    Next i
    ShippingBomComponentSignatureFromCollection = ShippingBomSignatureFromDictionary(dict)

CleanExit:
End Function

Private Function ShippingBomComponentSignatureFromTableVersion(ByVal lo As ListObject, _
                                                               ByVal packageRow As Long, _
                                                               ByVal bomVersion As Long) As String
    On Error GoTo CleanExit

    Dim dict As Object
    Dim cPackageRow As Long
    Dim cVersion As Long
    Dim cComponentRow As Long
    Dim cComponentItem As Long
    Dim cComponentQty As Long
    Dim cComponentUom As Long
    Dim cComponentLocation As Long
    Dim cComponentDescription As Long
    Dim r As Long
    Dim rowVersion As Long
    Dim componentUom As String
    Dim componentItem As String
    Dim componentLocation As String
    Dim componentDescription As String

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    cPackageRow = ColumnIndex(lo, "PackageRow")
    cVersion = ColumnIndex(lo, "BomVersion")
    cComponentRow = ColumnIndex(lo, "ComponentRow")
    cComponentItem = ColumnIndex(lo, "ComponentItem")
    cComponentQty = ColumnIndex(lo, "ComponentQty")
    cComponentUom = ColumnIndex(lo, "ComponentUOM")
    cComponentLocation = ColumnIndex(lo, "ComponentLocation")
    cComponentDescription = ColumnIndex(lo, "ComponentDescription")
    If cPackageRow = 0 Or cComponentRow = 0 Or cComponentQty = 0 Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    For r = 1 To lo.ListRows.Count
        If NzLng(lo.DataBodyRange.Cells(r, cPackageRow).Value) <> packageRow Then GoTo NextRow
        If cVersion > 0 Then rowVersion = NzLng(lo.DataBodyRange.Cells(r, cVersion).Value) Else rowVersion = 1
        If rowVersion <= 0 Then rowVersion = 1
        If rowVersion <> bomVersion Then GoTo NextRow

        componentUom = ""
        componentItem = ""
        componentLocation = ""
        componentDescription = ""
        If cComponentUom > 0 Then componentUom = NzStr(lo.DataBodyRange.Cells(r, cComponentUom).Value)
        If cComponentItem > 0 Then componentItem = NzStr(lo.DataBodyRange.Cells(r, cComponentItem).Value)
        If cComponentLocation > 0 Then componentLocation = NzStr(lo.DataBodyRange.Cells(r, cComponentLocation).Value)
        If cComponentDescription > 0 Then componentDescription = NzStr(lo.DataBodyRange.Cells(r, cComponentDescription).Value)

        AddShippingBomSignaturePart dict, _
                                    NzLng(lo.DataBodyRange.Cells(r, cComponentRow).Value), _
                                    NzDbl(lo.DataBodyRange.Cells(r, cComponentQty).Value), _
                                    componentUom, _
                                    componentItem, _
                                    componentLocation, _
                                    componentDescription
NextRow:
    Next r
    ShippingBomComponentSignatureFromTableVersion = ShippingBomSignatureFromDictionary(dict)

CleanExit:
End Function

Private Sub AddShippingBomSignaturePart(ByVal dict As Object, _
                                        ByVal componentRow As Long, _
                                        ByVal componentQty As Double, _
                                        ByVal componentUom As String, _
                                        ByVal componentItem As String, _
                                        ByVal componentLocation As String, _
                                        ByVal componentDescription As String)
    Dim key As String

    If dict Is Nothing Then Exit Sub
    If componentRow <= 0 And Trim$(componentItem) = "" Then Exit Sub
    If componentQty = 0 Then Exit Sub

    key = Format$(componentRow, "0000000000") & "|" & _
          NormalizeShippingBomSignatureText(componentItem) & "|" & _
          NormalizeShippingBomSignatureText(componentUom) & "|" & _
          NormalizeShippingBomSignatureText(componentLocation) & "|" & _
          NormalizeShippingBomSignatureText(componentDescription)
    If dict.Exists(key) Then
        dict(key) = CDbl(dict(key)) + componentQty
    Else
        dict.Add key, componentQty
    End If
End Sub

Private Function ShippingBomSignatureFromDictionary(ByVal dict As Object) As String
    Dim keys As Variant
    Dim i As Long
    Dim key As Variant
    Dim result As String

    If dict Is Nothing Then Exit Function
    If dict.Count = 0 Then Exit Function

    keys = SortedTextKeysShipping(dict)
    For i = LBound(keys) To UBound(keys)
        key = keys(i)
        If result <> "" Then result = result & vbLf
        result = result & CStr(key) & "|" & Format$(CDbl(dict(key)), "0.############")
    Next i
    ShippingBomSignatureFromDictionary = result
End Function

Private Function SortedTextKeysShipping(ByVal dict As Object) As Variant
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant

    If dict Is Nothing Then Exit Function
    keys = dict.Keys
    If Not IsArray(keys) Then
        SortedTextKeysShipping = keys
        Exit Function
    End If

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(CStr(keys(j)), CStr(keys(i)), vbTextCompare) < 0 Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
    SortedTextKeysShipping = keys
End Function

Private Function NormalizeShippingBomSignatureText(ByVal valueIn As String) As String
    NormalizeShippingBomSignatureText = LCase$(Trim$(valueIn))
End Function

Private Sub RetireActiveShippingBomPackageRows(ByVal lo As ListObject, _
                                               ByVal packageRow As Long, _
                                               ByVal retiredAt As Date, _
                                               ByVal retiredBy As String)
    Dim cPackageRow As Long
    Dim cActive As Long
    Dim i As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    cPackageRow = ColumnIndex(lo, "PackageRow")
    cActive = ColumnIndex(lo, "IsActive")
    If cPackageRow = 0 Then Exit Sub

    For i = 1 To lo.ListRows.Count
        If NzLng(lo.DataBodyRange.Cells(i, cPackageRow).Value) = packageRow Then
            If cActive = 0 Or ShippingBomActiveValue(lo.DataBodyRange.Cells(i, cActive).Value) Then
                SetTableCellShipping lo, i, "IsActive", False
                SetTableCellShipping lo, i, "EffectiveToUTC", retiredAt
                SetTableCellShipping lo, i, "RetiredAtUTC", retiredAt
                SetTableCellShipping lo, i, "UpdatedAtUTC", retiredAt
                SetTableCellShipping lo, i, "UpdatedBy", retiredBy
            End If
        End If
    Next i
End Sub

Private Function ShippingBomActiveValue(ByVal rawValue As Variant) As Boolean
    Dim textValue As String

    If IsEmpty(rawValue) Or IsNull(rawValue) Then
        ShippingBomActiveValue = True
        Exit Function
    End If
    textValue = UCase$(Trim$(NzStr(rawValue)))
    If textValue = "" Then
        ShippingBomActiveValue = True
    ElseIf textValue = "TRUE" Or textValue = "YES" Or textValue = "ACTIVE" Or textValue = "1" Then
        ShippingBomActiveValue = True
    Else
        ShippingBomActiveValue = False
    End If
End Function

Private Sub CopyShippingBomTable(ByVal loSource As ListObject, ByVal loTarget As ListObject)
    Dim headers As Variant
    Dim r As Long
    Dim c As Long
    Dim sourceCol As Long
    Dim targetCol As Long
    Dim rowsNeeded As Long

    If loTarget Is Nothing Then Exit Sub
    ClearListObjectData loTarget
    If loSource Is Nothing Then Exit Sub
    If loSource.DataBodyRange Is Nothing Then Exit Sub

    headers = ShippingBomHeaders()
    For c = LBound(headers) To UBound(headers)
        EnsureColumnExists loTarget, CStr(headers(c))
    Next c

    rowsNeeded = loSource.DataBodyRange.Rows.Count
    Do While loTarget.ListRows.Count < rowsNeeded
        loTarget.ListRows.Add
    Loop
    Do While loTarget.ListRows.Count > rowsNeeded
        loTarget.ListRows(loTarget.ListRows.Count).Delete
    Loop
    If loTarget.DataBodyRange Is Nothing Then Exit Sub
    loTarget.DataBodyRange.ClearContents

    For r = 1 To rowsNeeded
        For c = LBound(headers) To UBound(headers)
            sourceCol = ColumnIndex(loSource, CStr(headers(c)))
            targetCol = ColumnIndex(loTarget, CStr(headers(c)))
            If sourceCol > 0 And targetCol > 0 Then
                loTarget.DataBodyRange.Cells(r, targetCol).Value = loSource.DataBodyRange.Cells(r, sourceCol).Value
            End If
        Next c
    Next r
End Sub

Private Sub WriteShippingBomPackageTable(ByVal wbBom As Workbook, _
                                         ByVal packageRow As Long, _
                                         ByVal bomVersion As Long, _
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
    If bomVersion <= 0 Then bomVersion = 1
    If components Is Nothing Then Exit Sub
    If components.Count = 0 Then Exit Sub

    Set ws = WorkbookSheetExistsShipping(wbBom, SHEET_BOM_TABLES)
    If ws Is Nothing Then
        Set ws = wbBom.Worksheets.Add(After:=wbBom.Worksheets(wbBom.Worksheets.Count))
        ws.Name = SHEET_BOM_TABLES
    End If

    Set lo = EnsureShippingBomPackageTable(ws, packageRow, bomVersion, packageItem)
    If lo Is Nothing Then Exit Sub

    headers = ShippingBomPackageTableHeaders()
    ReDim arr(1 To components.Count, 1 To UBound(headers) - LBound(headers) + 1)
    For i = 1 To components.Count
        info = components(i)
        arr(i, 1) = NzLng(info(1))
        If UBound(info) >= 7 Then arr(i, 2) = NzStr(info(7))
        If UBound(info) >= 4 Then arr(i, 3) = NzStr(info(4))
        arr(i, 4) = NzDbl(info(2))
        arr(i, 5) = NzStr(info(3))
        If UBound(info) >= 5 Then arr(i, 6) = NzStr(info(5))
        If UBound(info) >= 6 Then arr(i, 7) = NzStr(info(6))
        arr(i, 8) = updatedAt
        arr(i, 9) = updatedBy
    Next i

    For c = LBound(headers) To UBound(headers)
        EnsureColumnExists lo, CStr(headers(c))
    Next c
    WriteArrayToTable lo, arr

CleanExit:
End Sub

Private Function EnsureShippingBomPackageTable(ByVal ws As Worksheet, _
                                               ByVal packageRow As Long, _
                                               ByVal bomVersion As Long, _
                                               ByVal packageItem As String) As ListObject
    On Error GoTo FailSoft

    Dim tableName As String
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim dataRange As Range
    Dim i As Long

    If ws Is Nothing Then Exit Function
    If bomVersion <= 0 Then bomVersion = 1
    tableName = BomTableNameFromRowVersion(packageRow, bomVersion)

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo FailSoft

    headers = ShippingBomPackageTableHeaders()
    If lo Is Nothing Then
        Set startCell = NextShippingBomPackageTableStartCell(ws)
        startCell.Offset(-1, 0).Value = "PackageRow"
        startCell.Offset(-1, 1).Value = packageRow
        startCell.Offset(-1, 2).Value = packageItem
        startCell.Offset(-1, 3).Value = "BomVersion"
        startCell.Offset(-1, 4).Value = bomVersion
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
        lo.Range.Cells(1, 1).Offset(-1, 3).Value = "BomVersion"
        lo.Range.Cells(1, 1).Offset(-1, 4).Value = bomVersion
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

Private Function ShippingReservationsWorkbookPath(ByVal warehouseId As String, ByVal rootPath As String) As String
    rootPath = NormalizeFolderPathShipping(rootPath)
    warehouseId = Trim$(warehouseId)
    If rootPath = "" Or warehouseId = "" Then Exit Function
    ShippingReservationsWorkbookPath = rootPath & "\" & warehouseId & ".invSys.Data.ShippingReservations.xlsb"
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

Private Sub HideWorkbookWindowsShipping(ByVal wb As Workbook)
    Dim i As Long

    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To wb.Windows.Count
        wb.Windows(i).Visible = False
    Next i
    On Error GoTo 0
End Sub

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

Private Function BomTableNameFromRowVersion(ByVal rowValue As Long, ByVal bomVersion As Long) As String
    If bomVersion <= 0 Then bomVersion = 1
    BomTableNameFromRowVersion = SafeTableName(BomTableNameFromRow(rowValue) & "_V" & CStr(bomVersion))
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
    PersistActiveShipmentRowsLocal loShip
    PersistHoldRowsLocal loHold
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
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set GetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
    If GetListObject Is Nothing Then
        If ShippingBridgeTableName(tableName) Then
            Set GetListObject = FindListObjectByNameShipping(ws.Parent, tableName)
        End If
    End If
End Function

Private Function ShippingBridgeTableName(ByVal tableName As String) As Boolean
    Select Case LCase$(Trim$(tableName))
        Case LCase$(TABLE_SHIPMENTS), _
             LCase$(TABLE_NOTSHIPPED), _
             LCase$(TABLE_AGG_BOM), _
             LCase$(TABLE_AGG_PACK), _
             LCase$(TABLE_CHECK_INV), _
             "invsysdata_shipping", _
             LCase$(TABLE_SHIPPING_BOM_VIEW), _
             "aggregateboxbom_log", _
             "aggregatepackages_log"
            ShippingBridgeTableName = True
    End Select
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

Private Function GetWritableShippingInvSysTable(ByVal wsShip As Worksheet, ByRef report As String) As ListObject
    Dim invLo As ListObject
    Dim sourceLo As ListObject

    If wsShip Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If

    Set invLo = GetInvSysTableFromWorkbook(wsShip.Parent)
    If invLo Is Nothing Then
        modRoleWorkbookSurfaces.EnsureInventoryManagementSurface wsShip.Parent
        Set invLo = GetInvSysTableFromWorkbook(wsShip.Parent)
    End If
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    If ShippingInventoryPickerTableHasRows(invLo) Then
        EnsureMissingInvSysRowsFromShipmentLines invLo, GetListObject(wsShip, TABLE_SHIPMENTS)
        ReconcileShipmentStagingFromShipmentLines invLo, GetListObject(wsShip, TABLE_SHIPMENTS)
        ReconcileShippableTotalsFromVersionInventory invLo
        Set GetWritableShippingInvSysTable = invLo
        Exit Function
    End If

    Set sourceLo = GetListObject(wsShip, "invSysData_Shipping")
    If Not ShippingInventoryPickerTableHasRows(sourceLo) Then Set sourceLo = GetListObject(wsShip, TABLE_CHECK_INV)
    If Not ShippingInventoryPickerTableHasRows(sourceLo) Then
        HydrateInvSysFromShipmentLines invLo, GetListObject(wsShip, TABLE_SHIPMENTS)
        ReconcileShipmentStagingFromShipmentLines invLo, GetListObject(wsShip, TABLE_SHIPMENTS)
        ReconcileShippableTotalsFromVersionInventory invLo
        If ShippingInventoryPickerTableHasRows(invLo) Then
            Set GetWritableShippingInvSysTable = invLo
        Else
            report = "Inventory read model has no rows to initialize invSys."
        End If
        Exit Function
    End If

    HydrateInvSysFromShippingReadModel invLo, sourceLo
    EnsureMissingInvSysRowsFromShipmentLines invLo, GetListObject(wsShip, TABLE_SHIPMENTS)
    ReconcileShipmentStagingFromShipmentLines invLo, GetListObject(wsShip, TABLE_SHIPMENTS)
    ReconcileShippableTotalsFromVersionInventory invLo
    If ShippingInventoryPickerTableHasRows(invLo) Then
        Set GetWritableShippingInvSysTable = invLo
    Else
        report = "InventoryManagement!invSys could not be initialized from the shipping read model."
    End If
End Function

Private Function GetShipmentReleaseInvSysTable(ByVal wsShip As Worksheet, ByRef report As String) As ListObject
    Dim invLo As ListObject

    If wsShip Is Nothing Then
        report = "ShipmentsTally sheet not found."
        Exit Function
    End If

    Set invLo = GetInvSysTableFromWorkbook(wsShip.Parent)
    If invLo Is Nothing Then
        modRoleWorkbookSurfaces.EnsureInventoryManagementSurface wsShip.Parent
        Set invLo = GetInvSysTableFromWorkbook(wsShip.Parent)
    End If
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If
    Set GetShipmentReleaseInvSysTable = invLo
End Function

Private Sub HydrateInvSysFromShippingReadModel(ByVal invLo As ListObject, ByVal sourceLo As ListObject)
    Dim rowCount As Long
    Dim r As Long
    Dim lc As ListColumn
    Dim sourceCol As Long

    If invLo Is Nothing Or sourceLo Is Nothing Then Exit Sub
    If sourceLo.DataBodyRange Is Nothing Then Exit Sub
    rowCount = sourceLo.DataBodyRange.Rows.Count
    If rowCount <= 0 Then Exit Sub

    EnsureShippingWorksheetEditable invLo.Parent
    ClearListObjectData invLo
    ResizeListObjectRowsForWrite invLo, rowCount
    For r = 1 To rowCount
        For Each lc In invLo.ListColumns
            sourceCol = ColumnIndex(sourceLo, CStr(lc.Name))
            If sourceCol > 0 Then
                invLo.DataBodyRange.Cells(r, lc.Index).Value = sourceLo.DataBodyRange.Cells(r, sourceCol).Value
            End If
        Next lc
    Next r
End Sub

Private Sub HydrateInvSysFromShipmentLines(ByVal invLo As ListObject, ByVal loShip As ListObject)
    Dim cRow As Long
    Dim cItem As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long
    Dim seen As Object
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String
    Dim lr As ListRow

    If invLo Is Nothing Or loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    cRow = ColumnIndex(loShip, "ROW")
    cItem = ColumnIndex(loShip, "ITEMS")
    cUom = ColumnIndex(loShip, "UOM")
    cLoc = ColumnIndex(loShip, "LOCATION")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    If cRow = 0 Or cItem = 0 Then Exit Sub

    EnsureShippingWorksheetEditable invLo.Parent
    ClearListObjectData invLo
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To loShip.ListRows.Count
        rowVal = NzLng(loShip.DataBodyRange.Cells(r, cRow).Value)
        itemName = Trim$(NzStr(loShip.DataBodyRange.Cells(r, cItem).Value))
        If rowVal <= 0 Or itemName = "" Then GoTo NextLine
        If seen.Exists(CStr(rowVal)) Then GoTo NextLine
        seen.Add CStr(rowVal), True

        Set lr = FirstBlankListRowShipping(invLo)
        If lr Is Nothing Then Set lr = invLo.ListRows.Add
        WriteValue lr, "ROW", rowVal
        WriteValue lr, "ITEM_CODE", itemName
        WriteValue lr, "ITEM", itemName
        If cUom > 0 Then WriteValue lr, "UOM", NzStr(loShip.DataBodyRange.Cells(r, cUom).Value)
        If cLoc > 0 Then WriteValue lr, "LOCATION", NzStr(loShip.DataBodyRange.Cells(r, cLoc).Value)
        If cDesc > 0 Then WriteValue lr, "DESCRIPTION", NzStr(loShip.DataBodyRange.Cells(r, cDesc).Value)
        WriteValue lr, "TOTAL INV", 0
        WriteValue lr, "SHIPMENTS", 0
NextLine:
    Next r
End Sub

Private Sub EnsureMissingInvSysRowsFromShipmentLines(ByVal invLo As ListObject, ByVal loShip As ListObject)
    Dim cRow As Long
    Dim cItem As Long
    Dim cUom As Long
    Dim cLoc As Long
    Dim cDesc As Long
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String

    If invLo Is Nothing Or loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    cRow = ColumnIndex(loShip, "ROW")
    cItem = ColumnIndex(loShip, "ITEMS")
    cUom = ColumnIndex(loShip, "UOM")
    cLoc = ColumnIndex(loShip, "LOCATION")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    If cRow = 0 Or cItem = 0 Then Exit Sub

    EnsureShippingWorksheetEditable invLo.Parent
    For r = 1 To loShip.ListRows.Count
        rowVal = NzLng(loShip.DataBodyRange.Cells(r, cRow).Value)
        itemName = Trim$(NzStr(loShip.DataBodyRange.Cells(r, cItem).Value))
        If rowVal <= 0 Or itemName = "" Then GoTo NextLine
        If FindInvRowIndexByRow(invLo, rowVal) <= 0 Then
            AddInvSysRowFromShipmentLine invLo, loShip, r, cRow, cItem, cUom, cLoc, cDesc
        End If
NextLine:
    Next r
End Sub

Private Sub AddInvSysRowFromShipmentLine(ByVal invLo As ListObject, _
                                         ByVal loShip As ListObject, _
                                         ByVal sourceRow As Long, _
                                         ByVal cRow As Long, _
                                         ByVal cItem As Long, _
                                         ByVal cUom As Long, _
                                         ByVal cLoc As Long, _
                                         ByVal cDesc As Long)
    Dim lr As ListRow
    Dim rowVal As Long
    Dim itemName As String
    Dim versionInv As Object
    Dim key As Variant
    Dim totalQty As Double

    If invLo Is Nothing Or loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    rowVal = NzLng(loShip.DataBodyRange.Cells(sourceRow, cRow).Value)
    itemName = Trim$(NzStr(loShip.DataBodyRange.Cells(sourceRow, cItem).Value))
    If rowVal <= 0 Or itemName = "" Then Exit Sub

    Set lr = FirstBlankListRowShipping(invLo)
    If lr Is Nothing Then Set lr = invLo.ListRows.Add
    WriteValue lr, "ROW", rowVal
    WriteValue lr, "ITEM_CODE", itemName
    WriteValue lr, "ITEM", itemName
    If cUom > 0 Then WriteValue lr, "UOM", NzStr(loShip.DataBodyRange.Cells(sourceRow, cUom).Value)
    If cLoc > 0 Then WriteValue lr, "LOCATION", NzStr(loShip.DataBodyRange.Cells(sourceRow, cLoc).Value)
    If cDesc > 0 Then WriteValue lr, "DESCRIPTION", NzStr(loShip.DataBodyRange.Cells(sourceRow, cDesc).Value)

    Set versionInv = BoxMakerFormLoadBoxVersionInventory(rowVal, itemName)
    If Not versionInv Is Nothing Then
        For Each key In versionInv.Keys
            totalQty = totalQty + NzDbl(versionInv(key))
        Next key
    End If
    WriteValue lr, "TOTAL INV", totalQty
    WriteValue lr, "SHIPMENTS", 0
End Sub

Private Sub ReconcileShipmentStagingFromShipmentLines(ByVal invLo As ListObject, ByVal loShip As ListObject)
    On Error GoTo FailSoft

    Dim cInvShip As Long
    Dim cInvRow As Long
    Dim cShipRow As Long
    Dim cShipQty As Long
    Dim cShipArea As Long
    Dim stagedByRow As Object
    Dim r As Long
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim invIdx As Long
    Dim key As Variant

    If invLo Is Nothing Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub
    cInvShip = ColumnIndex(invLo, "SHIPMENTS")
    cInvRow = ColumnIndex(invLo, "ROW")
    If cInvShip = 0 Or cInvRow = 0 Then Exit Sub

    EnsureShippingWorksheetEditable invLo.Parent
    For r = 1 To invLo.ListRows.Count
        invLo.DataBodyRange.Cells(r, cInvShip).Value = 0
    Next r

    If loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    cShipRow = ColumnIndex(loShip, "ROW")
    cShipQty = ColumnIndex(loShip, "QUANTITY")
    cShipArea = ColumnIndex(loShip, "AREA")
    If cShipRow = 0 Or cShipQty = 0 Then Exit Sub

    Set stagedByRow = CreateObject("Scripting.Dictionary")
    stagedByRow.CompareMode = vbTextCompare
    For r = 1 To loShip.ListRows.Count
        If cShipArea > 0 Then
            If StrComp(NormalizeShipmentArea(NzStr(loShip.DataBodyRange.Cells(r, cShipArea).Value), False), "Shipments", vbTextCompare) <> 0 Then GoTo NextLine
        End If
        rowVal = NzLng(loShip.DataBodyRange.Cells(r, cShipRow).Value)
        qtyVal = NzDbl(loShip.DataBodyRange.Cells(r, cShipQty).Value)
        If rowVal <= 0 Or qtyVal <= 0 Then GoTo NextLine
        If stagedByRow.Exists(CStr(rowVal)) Then
            stagedByRow(CStr(rowVal)) = NzDbl(stagedByRow(CStr(rowVal))) + qtyVal
        Else
            stagedByRow.Add CStr(rowVal), qtyVal
        End If
NextLine:
    Next r

    For Each key In stagedByRow.Keys
        invIdx = FindInvRowIndexByRow(invLo, CLng(key))
        If invIdx > 0 Then invLo.DataBodyRange.Cells(invIdx, cInvShip).Value = NzDbl(stagedByRow(CStr(key)))
    Next key

FailSoft:
End Sub

Private Sub ReconcileShippableTotalsFromVersionInventory(ByVal invLo As ListObject)
    On Error GoTo FailSoft

    Dim savedBoxes As Variant
    Dim versionInv As Object
    Dim r As Long
    Dim invIdx As Long
    Dim totalVersionQty As Double
    Dim availableVersionQty As Double
    Dim key As Variant
    Dim cTotal As Long
    Dim cShip As Long

    If invLo Is Nothing Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub
    cTotal = ColumnIndex(invLo, "TOTAL INV")
    If cTotal = 0 Then Exit Sub
    cShip = ColumnIndex(invLo, "SHIPMENTS")

    savedBoxes = BoxMakerFormLoadSavedBoxes()
    If IsEmpty(savedBoxes) Then Exit Sub

    For r = 1 To UBound(savedBoxes, 1)
        invIdx = FindInvRowIndexByRow(invLo, NzLng(savedBoxes(r, 1)))
        If invIdx <= 0 Then invIdx = FindInvRowIndexByItem(invLo, NzStr(savedBoxes(r, 2)))
        If invIdx <= 0 Then GoTo NextBox

        Set versionInv = BoxMakerFormLoadBoxVersionInventory(NzLng(savedBoxes(r, 1)), NzStr(savedBoxes(r, 2)))
        totalVersionQty = 0#
        If Not versionInv Is Nothing Then
            For Each key In versionInv.Keys
                totalVersionQty = totalVersionQty + NzDbl(versionInv(key))
            Next key
        End If
        availableVersionQty = totalVersionQty
        If cShip > 0 Then availableVersionQty = availableVersionQty - NzDbl(invLo.DataBodyRange.Cells(invIdx, cShip).Value)
        If availableVersionQty < 0 Then availableVersionQty = 0
        If availableVersionQty > NzDbl(invLo.DataBodyRange.Cells(invIdx, cTotal).Value) + 0.0000001 Then
            invLo.DataBodyRange.Cells(invIdx, cTotal).Value = availableVersionQty
        End If
NextBox:
    Next r
    Exit Sub

FailSoft:
End Sub

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
    Dim wanted As String
    wanted = NormalizeInventoryLookupTextShipping(itemName)
    If wanted = "" Then Exit Function
    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If NormalizeInventoryLookupTextShipping(invLo.DataBodyRange.Cells(r, cItem).Value) = wanted Then
            FindInvRowIndexByItem = r
            Exit Function
        End If
    Next r
End Function

Private Function FindInvRowIndexByItemCode(invLo As ListObject, ByVal itemCode As String) As Long
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    Dim cCode As Long: cCode = ColumnIndex(invLo, "ITEM_CODE")
    If cCode = 0 Then Exit Function
    Dim wanted As String
    wanted = NormalizeInventoryLookupTextShipping(itemCode)
    If wanted = "" Then Exit Function
    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If NormalizeInventoryLookupTextShipping(invLo.DataBodyRange.Cells(r, cCode).Value) = wanted Then
            FindInvRowIndexByItemCode = r
            Exit Function
        End If
    Next r
End Function

Private Function NormalizeInventoryLookupTextShipping(ByVal rawValue As Variant) As String
    Dim textValue As String

    textValue = NzStr(rawValue)
    If textValue = "" Then Exit Function
    textValue = Replace(textValue, Chr$(160), " ")
    textValue = Replace(textValue, vbTab, " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    Do While InStr(1, textValue, "  ", vbBinaryCompare) > 0
        textValue = Replace(textValue, "  ", " ")
    Loop
    NormalizeInventoryLookupTextShipping = LCase$(Trim$(textValue))
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
    Dim cActive As Long
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
    cActive = ColumnIndex(loBom, "IsActive")
    If cPackageRow = 0 Then GoTo CleanExit
    If loBom.DataBodyRange Is Nothing Then GoTo CleanExit

    For i = 1 To loBom.ListRows.Count
        rowValue = NzLng(loBom.DataBodyRange.Cells(i, cPackageRow).Value)
        If rowValue > maxPackageRow Then maxPackageRow = rowValue
        If cPackageItem > 0 Then
            If StrComp(Trim$(NzStr(loBom.DataBodyRange.Cells(i, cPackageItem).Value)), Trim$(boxName), vbTextCompare) = 0 Then
                If cActive = 0 Or ShippingBomActiveValue(loBom.DataBodyRange.Cells(i, cActive).Value) Then
                    FindShippingBomPackageRowByName = rowValue
                End If
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

Private Sub EnsureListObjectHasRowsShipping(ByVal lo As ListObject, ByVal wantedRows As Long)
    Dim rowCount As Long

    If lo Is Nothing Then Exit Sub
    If wantedRows < 1 Then wantedRows = 1
    If lo.DataBodyRange Is Nothing Then
        rowCount = 0
    Else
        rowCount = lo.ListRows.Count
    End If
    Do While rowCount < wantedRows
        lo.ListRows.Add
        rowCount = rowCount + 1
    Loop
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
    Dim cRow As Long: cRow = ColumnIndex(loShip, "ROW")
    Dim cArea As Long: cArea = ColumnIndex(loShip, "AREA")
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
        If cArea > 0 Then
            If StrComp(NormalizeShipmentArea(NzStr(data(r, cArea))), "Shipments", vbTextCompare) = 0 Then GoTo NextRow
        End If
        If qty <= 0 Or itemName = "" Then GoTo NextRow
        Dim rowVal As Long
        If cRow > 0 Then rowVal = NzLng(data(r, cRow))
        If rowVal = 0 Then rowVal = ResolveRowFromCaches(itemName, nameCache)
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
                info("ITEM_CODE") = NzStr(infovalue(invInfo, "ITEM_CODE"))
                info("UOM") = NzStr(infovalue(invInfo, "UOM"))
                info("LOCATION") = NzStr(infovalue(invInfo, "LOCATION"))
            Else
                info("ITEM") = itemName
                info("ITEM_CODE") = ""
                info("UOM") = ""
                info("LOCATION") = ""
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
                    info("ITEM_CODE") = NzStr(infovalue(invInfo, "ITEM_CODE"))
                    info("UOM") = NzStr(infovalue(invInfo, "UOM"))
                    info("LOCATION") = NzStr(infovalue(invInfo, "LOCATION"))
                Else
                    info("ITEM") = ""
                    info("ITEM_CODE") = ""
                    info("UOM") = ""
                    info("LOCATION") = ""
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
                info("ITEM_CODE") = NzStr(infovalue(invInfo, "ITEM_CODE"))
                info("UOM") = NzStr(infovalue(invInfo, "UOM"))
                info("LOCATION") = NzStr(infovalue(invInfo, "LOCATION"))
            Else
                info("ITEM") = ""
                info("ITEM_CODE") = ""
                info("UOM") = ""
                info("LOCATION") = ""
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
    ResizeListObjectRowsForWrite lo, count
    Dim i As Long
    For i = 1 To count
        Dim key As Variant: key = keys(LBound(keys) + i - 1)
        Dim info As Object: Set info = pkgDict(key)
        Dim itemText As String: itemText = NzStr(infovalue(info, "ITEM"))
        If itemText = "" Then itemText = NzStr(key)
        WriteTableCellByName lo, i, "ROW", CLng(key)
        WriteTableCellByName lo, i, "ITEM_CODE", NzStr(infovalue(info, "ITEM_CODE"))
        WriteTableCellByName lo, i, "ITEM", itemText
        WriteTableCellByName lo, i, "QUANTITY", NzDbl(infovalue(info, "QTY"))
        WriteTableCellByName lo, i, "UOM", NzStr(infovalue(info, "UOM"))
        WriteTableCellByName lo, i, "LOCATION", NzStr(infovalue(info, "LOCATION"))
    Next i
End Sub

Private Sub WriteAggregateBOM(lo As ListObject, bomDict As Object)
    If lo Is Nothing Then Exit Sub
    ClearListObjectData lo
    If bomDict Is Nothing Then Exit Sub
    If bomDict.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = SortedKeys(bomDict)
    Dim count As Long: count = UBound(keys) - LBound(keys) + 1
    ResizeListObjectRowsForWrite lo, count
    Dim i As Long
    For i = 1 To count
        Dim key As Variant: key = keys(LBound(keys) + i - 1)
        Dim info As Object: Set info = bomDict(key)
        WriteTableCellByName lo, i, "ROW", CLng(key)
        WriteTableCellByName lo, i, "ITEM_CODE", NzStr(infovalue(info, "ITEM_CODE"))
        WriteTableCellByName lo, i, "ITEM", NzStr(infovalue(info, "ITEM"))
        WriteTableCellByName lo, i, "QUANTITY", NzDbl(infovalue(info, "QTY"))
        WriteTableCellByName lo, i, "UOM", NzStr(infovalue(info, "UOM"))
        WriteTableCellByName lo, i, "LOCATION", NzStr(infovalue(info, "LOCATION"))
    Next i
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
    ResizeListObjectRowsForWrite lo, count
    Dim i As Long
    For i = 1 To count
        rowKey = keys(LBound(keys) + i - 1)
        Dim info As Object
        If Not rowCache Is Nothing Then
            If rowCache.Exists(CStr(rowKey)) Then Set info = rowCache(CStr(rowKey))
        End If
        WriteTableCellByName lo, i, "ROW", CLng(rowKey)
        WriteTableCellByName lo, i, "ITEM_CODE", NzStr(infovalue(info, "ITEM_CODE"))
        WriteTableCellByName lo, i, "ITEM", NzStr(infovalue(info, "ITEM"))
        WriteTableCellByName lo, i, "UOM", NzStr(infovalue(info, "UOM"))
        WriteTableCellByName lo, i, "LOCATION", NzStr(infovalue(info, "LOCATION"))
        WriteTableCellByName lo, i, "USED", NzDbl(infovalue(info, "USED"))
        WriteTableCellByName lo, i, "MADE", NzDbl(infovalue(info, "MADE"))
        WriteTableCellByName lo, i, "SHIPMENTS", NzDbl(infovalue(info, "SHIPMENTS"))
        WriteTableCellByName lo, i, "TOTAL INV", NzDbl(infovalue(info, "TOTAL_INV"))
    Next i
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
    If aggPack Is Nothing Then Exit Function
    If aggPack.DataBodyRange Is Nothing Then Exit Function

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
    If invLo.DataBodyRange Is Nothing Then Exit Function
    If aggPack Is Nothing Then Exit Function
    If aggPack.DataBodyRange Is Nothing Then Exit Function

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

Private Function BuildShipmentLineDeltas(ByVal invLo As ListObject, ByVal loShip As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function
    If loShip Is Nothing Then Exit Function
    If loShip.DataBodyRange Is Nothing Then Exit Function

    Dim cQtyShip As Long: cQtyShip = ColumnIndex(loShip, "QUANTITY")
    Dim cRowShip As Long: cRowShip = ColumnIndex(loShip, "ROW")
    Dim cItemShip As Long: cItemShip = ColumnIndex(loShip, "ITEMS")
    If cQtyShip = 0 Or cRowShip = 0 Then
        errNotes = "Shipments table missing QUANTITY/ROW columns."
        Exit Function
    End If

    Dim colTotalInv As Long: colTotalInv = ColumnIndex(invLo, "TOTAL INV")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    If colTotalInv = 0 Then
        errNotes = "invSys table missing TOTAL INV column."
        Exit Function
    End If

    Dim requirements As Object: Set requirements = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loShip.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As Long: rowVal = NzLng(arr(r, cRowShip))
        Dim qtyVal As Double: qtyVal = NzDbl(arr(r, cQtyShip))
        If rowVal = 0 Or qtyVal <= 0 Then GoTo NextShipRow
        Dim reqKey As String: reqKey = CStr(rowVal)
        If requirements.Exists(reqKey) Then
            requirements(reqKey) = NzDbl(requirements(reqKey)) + qtyVal
        Else
            requirements.Add reqKey, qtyVal
        End If
NextShipRow:
    Next r
    If requirements.Count = 0 Then
        errNotes = "No shipment rows are ready to send."
        Exit Function
    End If

    Dim result As New Collection
    Dim key As Variant
    For Each key In requirements.Keys
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, CLng(key))
        If invRow Is Nothing Then
            AppendNote errNotes, "Package ROW " & CStr(key) & " not found in invSys."
            Exit Function
        End If

        Dim requiredQty As Double: requiredQty = NzDbl(requirements(key))
        Dim available As Double: available = NzDbl(invRow.Range.Cells(1, colTotalInv).Value)
        If requiredQty > available + 0.0000001 Then
            AppendNote errNotes, "ROW " & CStr(key) & " requires " & Format$(requiredQty, "0.###") & " but only " & Format$(available, "0.###") & " in TOTAL INV."
            Exit Function
        End If

        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = CLng(key)
        delta("QTY") = requiredQty
        If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(invRow.Range.Cells(1, colItemCode).Value)
        If colItemName > 0 Then
            delta("ITEM_NAME") = NzStr(invRow.Range.Cells(1, colItemName).Value)
        ElseIf cItemShip > 0 Then
            delta("ITEM_NAME") = ShipmentItemNameForRow(loShip, cRowShip, cItemShip, CLng(key))
        End If
        result.Add delta
    Next key

    If result.Count > 0 Then Set BuildShipmentLineDeltas = result
End Function

Private Function BuildDisplayedShipmentRowsDeltas(ByVal invLo As ListObject, ByVal rowsData As Variant, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function
    If IsEmpty(rowsData) Then Exit Function

    Dim lb1 As Long
    Dim ub1 As Long
    On Error GoTo BadRows
    lb1 = LBound(rowsData, 1)
    ub1 = UBound(rowsData, 1)
    On Error GoTo 0
    If ub1 < lb1 Then Exit Function

    Dim colTotalInv As Long: colTotalInv = ColumnIndex(invLo, "TOTAL INV")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    If colTotalInv = 0 Then
        errNotes = "invSys table missing TOTAL INV column."
        Exit Function
    End If

    Dim requirements As Object: Set requirements = CreateObject("Scripting.Dictionary")
    Dim names As Object: Set names = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = lb1 To ub1
        Dim rowVal As Long: rowVal = NzLng(rowsData(r, 6))
        Dim qtyVal As Double: qtyVal = NzDbl(rowsData(r, 3))
        If rowVal = 0 Or qtyVal <= 0 Then GoTo NextDisplayedRow
        Dim rowKey As String: rowKey = CStr(rowVal)
        If requirements.Exists(rowKey) Then
            requirements(rowKey) = NzDbl(requirements(rowKey)) + qtyVal
        Else
            requirements.Add rowKey, qtyVal
            names.Add rowKey, NzStr(rowsData(r, 2))
        End If
NextDisplayedRow:
    Next r
    If requirements.Count = 0 Then
        errNotes = "No shipment rows are ready to send."
        Exit Function
    End If

    Dim result As New Collection
    Dim key As Variant
    For Each key In requirements.Keys
        Dim invRow As ListRow: Set invRow = FindInvListRowByRowValue(invLo, CLng(key))
        If invRow Is Nothing Then
            AppendNote errNotes, "Package ROW " & CStr(key) & " not found in invSys."
            Exit Function
        End If

        Dim requiredQty As Double: requiredQty = NzDbl(requirements(key))
        Dim available As Double: available = NzDbl(invRow.Range.Cells(1, colTotalInv).Value)
        If requiredQty > available + 0.0000001 Then
            AppendNote errNotes, "ROW " & CStr(key) & " requires " & Format$(requiredQty, "0.###") & " but only " & Format$(available, "0.###") & " in TOTAL INV."
            Exit Function
        End If

        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("ROW") = CLng(key)
        delta("QTY") = requiredQty
        If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(invRow.Range.Cells(1, colItemCode).Value)
        If colItemName > 0 Then
            delta("ITEM_NAME") = NzStr(invRow.Range.Cells(1, colItemName).Value)
        ElseIf names.Exists(CStr(key)) Then
            delta("ITEM_NAME") = NzStr(names(CStr(key)))
        End If
        result.Add delta
    Next key

    If result.Count > 0 Then Set BuildDisplayedShipmentRowsDeltas = result
    Exit Function

BadRows:
    errNotes = "Shipment rows could not be read from the form."
End Function

Private Function ShipmentItemNameForRow(ByVal loShip As ListObject, _
                                        ByVal rowColumn As Long, _
                                        ByVal itemColumn As Long, _
                                        ByVal rowValue As Long) As String
    Dim r As Long

    If loShip Is Nothing Then Exit Function
    If loShip.DataBodyRange Is Nothing Then Exit Function
    If rowColumn <= 0 Or itemColumn <= 0 Then Exit Function
    For r = 1 To loShip.DataBodyRange.Rows.Count
        If NzLng(loShip.DataBodyRange.Cells(r, rowColumn).Value) = rowValue Then
            ShipmentItemNameForRow = NzStr(loShip.DataBodyRange.Cells(r, itemColumn).Value)
            Exit Function
        End If
    Next r
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
        If delta.Exists("VERSION") Then
            If Trim$(NzStr(delta("VERSION"))) <> "" Then
                Dim payloadVersion As String: payloadVersion = NormalizeBoxBomVersionLabelShipping(NzStr(delta("VERSION")))
                payloadItem("Version") = payloadVersion
                payloadItem("BomVersionLabel") = payloadVersion
                payloadItem("Note") = Trim$(NzStr(delta("ITEM_NAME")) & " VERSION=" & payloadVersion)
            End If
        End If
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
    Dim validationByRow As Object
    Dim key As Variant
    Set validationByRow = CreateObject("Scripting.Dictionary")
    validationByRow.CompareMode = vbTextCompare
    For Each delta In deltas
        Dim rowVal As Long: rowVal = CLng(delta("ROW"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextValidate
        If validationByRow.Exists(CStr(rowVal)) Then
            validationByRow(CStr(rowVal)) = NzDbl(validationByRow(CStr(rowVal))) + qtyVal
        Else
            validationByRow.Add CStr(rowVal), qtyVal
        End If
NextValidate:
    Next delta

    For Each key In validationByRow.Keys
        rowVal = CLng(key)
        qtyVal = NzDbl(validationByRow(CStr(key)))
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
    Next key

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

Private Function ApplyDirectShipmentsSentDeltas(ByVal invLo As ListObject, ByVal deltas As Collection, ByRef errNotes As String) As Double
    ApplyDirectShipmentsSentDeltas = 0
    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        ApplyDirectShipmentsSentDeltas = -1
        Exit Function
    End If
    If deltas Is Nothing Then Exit Function
    If deltas.Count = 0 Then Exit Function

    Dim colTotal As Long: colTotal = ColumnIndex(invLo, "TOTAL INV")
    Dim colLastEdited As Long: colLastEdited = ColumnIndex(invLo, "LAST EDITED")
    Dim colTotalLastEdit As Long: colTotalLastEdit = ColumnIndex(invLo, "TOTAL INV LAST EDIT")
    If colTotal = 0 Then
        errNotes = "invSys table missing TOTAL INV column."
        ApplyDirectShipmentsSentDeltas = -1
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
            ApplyDirectShipmentsSentDeltas = -1
            Exit Function
        End If

        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotal)
        Dim currentTotal As Double: currentTotal = NzDbl(totalCell.Value)
        If qtyVal > currentTotal + 0.0000001 Then
            AppendNote errNotes, "ROW " & rowVal & " only has " & Format$(currentTotal, "0.###") & " in TOTAL INV but needs " & Format$(qtyVal, "0.###") & "."
            ApplyDirectShipmentsSentDeltas = -1
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
        totalCell.Value = NzDbl(totalCell.Value) - qtyVal
        If colLastEdited > 0 Then invRow.Range.Cells(1, colLastEdited).Value = Now
        If colTotalLastEdit > 0 Then invRow.Range.Cells(1, colTotalLastEdit).Value = Now
        ApplyDirectShipmentsSentDeltas = ApplyDirectShipmentsSentDeltas + qtyVal
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

Private Function ApplyShipmentReleaseDeltasLocal(invLo As ListObject, deltas As Collection, ByRef errNotes As String, Optional ByVal allowMissingLocalStage As Boolean = False) As Double
    ApplyShipmentReleaseDeltasLocal = 0
    errNotes = ""
    If invLo Is Nothing Then
        errNotes = "invSys table not found."
        ApplyShipmentReleaseDeltasLocal = -1
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
        ApplyShipmentReleaseDeltasLocal = -1
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
            ApplyShipmentReleaseDeltasLocal = -1
            Exit Function
        End If

        Dim shipCell As Range: Set shipCell = invRow.Range.Cells(1, colShip)
        Dim currentShip As Double: currentShip = NzDbl(shipCell.Value)
        If qtyVal > currentShip + 0.0000001 Then
            If Not allowMissingLocalStage Then
                AppendNote errNotes, "ROW " & rowVal & " only has " & Format$(currentShip, "0.###") & " staged but needs " & Format$(qtyVal, "0.###") & " to release."
                ApplyShipmentReleaseDeltasLocal = -1
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
        Dim totalCell As Range: Set totalCell = invRow.Range.Cells(1, colTotal)
        Set shipCell = invRow.Range.Cells(1, colShip)
        currentShip = NzDbl(shipCell.Value)
        Dim localReleaseQty As Double: localReleaseQty = qtyVal
        If allowMissingLocalStage And localReleaseQty > currentShip Then localReleaseQty = currentShip
        If localReleaseQty <= 0 Then GoTo NextApply
        totalCell.Value = NzDbl(totalCell.Value) + localReleaseQty
        shipCell.Value = currentShip - localReleaseQty
        If NzDbl(shipCell.Value) < 0 Then shipCell.Value = 0
        If colLastEdited > 0 Then invRow.Range.Cells(1, colLastEdited).Value = Now
        If colTotalLastEdit > 0 Then invRow.Range.Cells(1, colTotalLastEdit).Value = Now
        ApplyShipmentReleaseDeltasLocal = ApplyShipmentReleaseDeltasLocal + localReleaseQty
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

Private Sub ResizeListObjectRowsForWrite(ByVal lo As ListObject, ByVal rowsNeeded As Long)
    Dim currentRows As Long
    Dim diff As Long

    If lo Is Nothing Then Exit Sub
    If rowsNeeded < 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = lo.DataBodyRange.Rows.Count
    End If

    If currentRows < rowsNeeded Then
        For diff = 1 To rowsNeeded - currentRows
            lo.ListRows.Add
        Next diff
    ElseIf currentRows > rowsNeeded Then
        For diff = currentRows To rowsNeeded + 1 Step -1
            lo.ListRows(diff).Delete
        Next diff
    End If
End Sub

Private Sub WriteTableCellByName(ByVal lo As ListObject, _
                                 ByVal rowIndex As Long, _
                                 ByVal columnName As String, _
                                 ByVal value As Variant)
    Dim colIndex As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If rowIndex <= 0 Or rowIndex > lo.DataBodyRange.Rows.Count Then Exit Sub
    colIndex = ColumnIndex(lo, columnName)
    If colIndex = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = value
End Sub

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
    Dim wb As Workbook
    Dim tbl As ListObject

    Set wb = ResolveShippingWorkbook(, SHEET_SHIPMENTS)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Set tbl = FindListObjectByNameShipping(wb, logTableName)
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
