Attribute VB_Name = "modTS_Shipments"
Option Explicit

' =============================================================
' Module: modTS_Shipments
' Purpose: All logic for the ShippingTally system (box builder,
'          holding subsystem, confirm/build/ship macros, logging).
' Notes:
'   - Buttons are generated dynamically (similar to modTS_Received).
'   - ShippingBOM sheet stores one ListObject per BOM (Box Name).
'   - BOM entries store ROW/QUANTITY/UOM only; item metadata is
'     resolved from invSys (InventoryManagement!invSys).
'   - Hold subsystem keeps packages on NotShipped until released.
'   - Additional confirm/build/ship routines will be implemented in
'     subsequent iterations (placeholders provided below).
' =============================================================

' ===== constants =====
Private Const SHEET_SHIPMENTS As String = "ShipmentsTally"
Private Const SHEET_INV As String = "InventoryManagement"
Private Const SHEET_BOM As String = "ShippingBOM"

Private Const TABLE_SHIPMENTS As String = "ShipmentsTally"
Private Const TABLE_NOTSHIPPED As String = "NotShipped"
Private Const TABLE_AGG_BOM As String = "AggregateBoxBOM"
Private Const TABLE_AGG_PACK As String = "AggregatePackages"
Private Const TABLE_BOX_BUILDER As String = "BoxBuilder"
Private Const TABLE_BOX_BOM As String = "BoxBOM"
Private Const TABLE_CHECK_INV As String = "Check_invSys"
Private Const COL_BOXBOM_ITEM As String = "ITEM"

Private Const BTN_TOGGLE_BUILDER As String = "BTN_TOGGLE_BUILDER"
Private Const BTN_SAVE_BOX As String = "BTN_SAVE_BOX"
Private Const BTN_UNSHIP As String = "BTN_UNSHIP"
Private Const BTN_SEND_HOLD As String = "BTN_SEND_HOLD"
Private Const BTN_RETURN_HOLD As String = "BTN_RETURN_HOLD"
Private Const BTN_CONFIRM_INV As String = "BTN_CONFIRM_INV"
Private Const BTN_BOXES_MADE As String = "BTN_BOXES_MADE"
Private Const BTN_TO_TOTALINV As String = "BTN_TO_TOTALINV"
Private Const BTN_TO_SHIPMENTS As String = "BTN_TO_SHIPMENTS"
Private Const BTN_SHIPMENTS_SENT As String = "BTN_SHIPMENTS_SENT"
Private Const CHK_USE_EXISTING As String = "CHK_USE_EXISTING"

Private Const SHIPPING_BOM_BLOCK_ROWS As Long = 52
Private Const SHIPPING_BOM_DATA_ROWS As Long = 50
Private Const SHIPPING_BOM_COLS As Long = 3 ' ROW, QUANTITY, UOM
Private Const SHIPMENTS_SENT_DEDUCTS_TOTALINV As Boolean = False

Private mDynSearch As cDynItemSearch
Private mNextInvSysRow As Long
Private mAggDirty As Boolean

' ===== public entry points =====
Public Sub InitializeShipmentsUI()
    EnsureShipmentsButtons
    EnsureBuilderTablesReady
    If mAggDirty Then RebuildShippingAggregates
End Sub

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
    EnsureColumnExists loMeta, "ROW"
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
    boxRowValue = EnsureInvSysItem(boxName, boxUOM, boxLoc, boxDesc, invLo)
    If boxRowValue = 0 Then Exit Sub
    Dim cBoxRowField As Long: cBoxRowField = ColumnIndex(loMeta, "ROW")
    If cBoxRowField > 0 Then
        loMeta.DataBodyRange.Cells(1, cBoxRowField).Value = boxRowValue
    End If

    Dim wsBOM As Worksheet: Set wsBOM = SheetExists(SHEET_BOM)
    If wsBOM Is Nothing Then
        MsgBox "ShippingBOM sheet not found.", vbCritical
        Exit Sub
    End If
    Dim bomTable As ListObject, blockRange As Range
    Set bomTable = EnsureBomTable(wsBOM, boxName, boxRowValue, blockRange)
    If bomTable Is Nothing Then Exit Sub

    WriteBomData bomTable, blockRange, components
    PropagateBomMetadata wsBOM, components

    Dim finalMsg As String
    finalMsg = "Saved BOM '" & boxName & "' (invSys ROW " & boxRowValue & ", " & components.count & " components)."
    If Len(syncNotes) > 0 Then
        finalMsg = finalMsg & vbCrLf & syncNotes
    End If
    MsgBox finalMsg, vbInformation

    ClearListObjectData loMeta
    ClearListObjectData loBom
    EnsureTableHasRow loMeta
    EnsureTableHasRow loBom
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
        MsgBox msg, vbInformation
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

    usedTotal = modInvMan.ApplyUsedDeltas(usedDeltas, errNotes, "BTN_BOXES_MADE - Components Used")
    If usedTotal < 0 Then
        MsgBox "Boxes made cancelled: insufficient inventory to cover all BOM components." & vbCrLf & vbCrLf & errNotes, vbExclamation
        Exit Sub
    End If

    madeTotal = modInvMan.ApplyMadeDeltas(madeDeltas, errNotes, "BTN_BOXES_MADE - Packages Staged")
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
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
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
    If deltas Is Nothing Or deltas.Count = 0 Then
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
    movedTotal = modInvMan.ApplyMadeToInventoryDeltas(deltas, errNotes, "BTN_TO_TOTALINV - Added To Total Inv")
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
    MsgBox msg, vbInformation
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
    If deltas Is Nothing Or deltas.Count = 0 Then
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
    stagedTotal = modInvMan.ApplyShipmentDeltas(deltas, errNotes, "BTN_TO_SHIPMENTS - Inventory Staged")
    If stagedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to stage shipments due to inventory shortage."
        MsgBox errNotes, vbCritical
        Exit Sub
    End If

    InvalidateAggregates True
    If shipLogs.Count > 0 Then LogShippingChanges "AggregatePackages_Log", shipLogs

    Dim msg As String
    msg = "Staged " & Format$(stagedTotal, "0.###") & " packages into invSys.SHIPMENTS."
    MsgBox msg, vbInformation
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
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim errNotes As String
    Dim deltas As Collection
    Set deltas = BuildShipmentsSentDeltaPacket(invLo, errNotes)
    If deltas Is Nothing Or deltas.Count = 0 Then
        If errNotes <> "" Then
            MsgBox errNotes, vbInformation
        Else
            MsgBox "No staged shipments found in invSys.SHIPMENTS.", vbInformation
        End If
        Exit Sub
    End If

    Dim aggPack As ListObject: Set aggPack = GetListObject(ws, TABLE_AGG_PACK)
    Dim rowFilter As Object
    If Not aggPack Is Nothing Then
        If Not aggPack.DataBodyRange Is Nothing Then
            Dim cRowAgg As Long: cRowAgg = ColumnIndex(aggPack, "ROW")
            If cRowAgg > 0 Then
                Set rowFilter = CreateObject("Scripting.Dictionary")
                Dim arrAgg As Variant: arrAgg = aggPack.DataBodyRange.Value
                Dim r As Long
                For r = 1 To UBound(arrAgg, 1)
                    Dim rowVal As Long: rowVal = NzLng(arrAgg(r, cRowAgg))
                    If rowVal > 0 Then rowFilter(CStr(rowVal)) = True
                Next r
            End If
        End If
    End If

    If Not rowFilter Is Nothing Then
        If rowFilter.Count > 0 Then
            Dim filtered As New Collection
            Dim delta As Variant
            For Each delta In deltas
                If rowFilter.Exists(CStr(delta("ROW"))) Then filtered.Add delta
            Next delta
            Set deltas = filtered
            If deltas.Count = 0 Then
                MsgBox "No staged shipments match the current AggregatePackages rows.", vbInformation
                Exit Sub
            End If
        End If
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

    Dim msg As String
    msg = "Finalized " & Format$(shippedTotal, "0.###") & " shipments."
    If SHIPMENTS_SENT_DEDUCTS_TOTALINV Then
        msg = msg & vbCrLf & "TOTAL INV reduced; SHIPMENTS cleared."
    Else
        msg = msg & vbCrLf & "SHIPMENTS cleared."
    End If
    MsgBox msg, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "BTN_SHIPMENTS_SENT failed: " & Err.Description, vbCritical
End Sub

Public Sub ShowDynamicItemSearch(ByVal targetCell As Range)
    On Error GoTo ErrHandler
    If targetCell Is Nothing Then Exit Sub
    If mDynSearch Is Nothing Then Set mDynSearch = New cDynItemSearch
    mDynSearch.ShowForCell targetCell
    Exit Sub
ErrHandler:
    On Error Resume Next
    frmItemSearch.Show vbModeless
End Sub

' ===== button scaffolding =====
Private Sub EnsureShipmentsButtons()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    DeleteShapeIfExists ws, "BTN_SHOW_BUILDER"
    DeleteShapeIfExists ws, "BTN_HIDE_BUILDER"
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
    Dim nextTop As Double: nextTop = chkTop + CHK_STACK_SPACING

    EnsureButtonCustom ws, BTN_TOGGLE_BUILDER, "Toggle builder", "modTS_Shipments.BtnToggleBuilder", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_SAVE_BOX, "Save box", "modTS_Shipments.BtnSaveBox", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_CONFIRM_INV, "Confirm inventory", "modTS_Shipments.BtnConfirmInventory", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_BOXES_MADE, "Boxes made", "modTS_Shipments.BtnBoxesMade", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_UNSHIP, "Toggle NotShipped", "modTS_Shipments.BtnUnship", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_SEND_HOLD, "Send to hold", "modTS_Shipments.BtnSendHold", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_RETURN_HOLD, "Return from hold", "modTS_Shipments.BtnReturnHold", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_TO_TOTALINV, "To TotalInv", "modTS_Shipments.BtnToTotalInv", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_TO_SHIPMENTS, "To Shipments", "modTS_Shipments.BtnToShipments", leftA, nextTop, colAWidth
    nextTop = nextTop + BTN_STACK_SPACING
    EnsureButtonCustom ws, BTN_SHIPMENTS_SENT, "Shipments sent", "modTS_Shipments.BtnShipmentsSent", leftA, nextTop, colAWidth
End Sub

Public Sub ToggleUseExistingInventory()
    InvalidateAggregates True
End Sub

Private Sub EnsureButtonCustom(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String, leftPos As Double, topPos As Double, Optional widthPts As Double = 118)
    Const BTN_HEIGHT As Double = 20
    If widthPts < 20 Then widthPts = 118
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, widthPts, BTN_HEIGHT)
        shp.Name = shapeName
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = onActionMacro
    Else
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = BTN_HEIGHT
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = onActionMacro
    End If
End Sub

Private Sub EnsureCheckbox(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String, leftPos As Double, topPos As Double, Optional widthPts As Double = 118)
    Const CHK_HEIGHT As Double = 26
    If widthPts < 20 Then widthPts = 118
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        On Error Resume Next
        If shp.Type <> msoFormControl Or shp.FormControlType <> xlCheckBox Then
            Set shp = Nothing
        End If
        On Error GoTo 0
    End If
    If shp Is Nothing Then
        Dim candidate As Shape
        Dim bestMatch As Shape
        Dim bestTop As Double: bestTop = 1E+30
        For Each candidate In ws.Shapes
            If candidate.Type = msoFormControl Then
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
        shp.OnAction = onActionMacro
    Else
        shp.Name = shapeName
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = CHK_HEIGHT
        shp.OnAction = onActionMacro
    End If
    ForceCheckboxCaption shp, caption
End Sub

Private Sub DeleteLegacyCheckBoxes(ws As Worksheet)
    Dim shp As Shape
    Dim toDelete As Collection: Set toDelete = New Collection
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
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

Private Sub EnsureBuilderTablesReady()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If Not loBom Is Nothing Then EnsureBoxBomEntryColumns loBom
End Sub

Public Sub ApplyItemSelection(targetCell As Range, lo As ListObject, rowIndex As Long, _
    ByVal itemName As String, ByVal itemCode As String, ByVal itemRow As Long, _
    ByVal uom As String, ByVal location As String, ByVal vendor As String, _
    Optional ByVal description As String = "")

    If lo Is Nothing Then Exit Sub
    
    Dim tableName As String
    tableName = LCase$(lo.Name)

    Select Case tableName
        Case "shipmentstally"
            targetCell.Value = itemName
            InvalidateAggregates True
            
        Case Else
            ' no-op
    End Select
End Sub

Public Sub ApplyItemToBoxBOM(targetCell As Range, ByVal itemName As String, ByVal itemRow As Long, _
    ByVal uom As String, ByVal location As String, ByVal description As String)

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
    Dim colItemInv As Long: colItemInv = ColumnIndex(invLo, "ITEM")
    Dim colUomInv As Long: colUomInv = ColumnIndex(invLo, "UOM")
    Dim colLocInv As Long: colLocInv = ColumnIndex(invLo, "LOCATION")
    Dim colDescInv As Long: colDescInv = ColumnIndex(invLo, "DESCRIPTION")

    If colRowInv > 0 Then actualRow = NzLng(invLo.DataBodyRange.Cells(invIdx, colRowInv).Value)
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
    WriteValue lr, "ROW", actualRow
    WriteValue lr, "UOM", actualUom
    WriteValue lr, "LOCATION", actualLoc
    WriteValue lr, "DESCRIPTION", actualDesc
    Exit Sub

ErrHandler:
    MsgBox "ApplyItemToBoxBOM error: " & Err.Description, vbCritical
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

        Dim entry(1 To 3) As Variant
        entry(1) = partRow
        entry(2) = qty
        entry(3) = uomVal
        result.Add entry
NextComponent:
    Next

    Set CollectBomComponents = result
End Function

Private Sub EnsureBoxBomEntryColumns(loBom As ListObject)
    If loBom Is Nothing Then Exit Sub
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
    EnsureColumnExists loBom, "QUANTITY", COL_BOXBOM_ITEM
    EnsureColumnExists loBom, "ROW"
    EnsureColumnExists loBom, "UOM"
    EnsureColumnExists loBom, "LOCATION"
    EnsureColumnExists loBom, "DESCRIPTION"
End Sub

Private Sub EnsureColumnExists(lo As ListObject, colName As String, Optional afterColumn As String = "")
    If lo Is Nothing Then Exit Sub
    If ColumnIndex(lo, colName) > 0 Then Exit Sub
    Dim insertPos As Long
    If afterColumn <> "" Then insertPos = ColumnIndex(lo, afterColumn)
    Dim newCol As ListColumn
    If insertPos > 0 Then
        Set newCol = lo.ListColumns.Add(insertPos + 1)
    Else
        Set newCol = lo.ListColumns.Add
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

    AddOrMergeHoldRow targetTable, refVal, itemVal, qtyMove

    Dim newQty As Double
    If moveToHold Then
        newQty = qtyVal - qtyMove
    Else
        newQty = qtyVal - qtyMove
    End If
    If newQty <= 0 Then
        sourceTable.ListRows(rowIndex).Range.ClearContents
    Else
        sourceTable.DataBodyRange.Cells(rowIndex, cQty).Value = newQty
    End If
End Sub

Private Sub AddOrMergeHoldRow(targetTable As ListObject, refVal As String, itemVal As String, qtyMove As Double)
    If targetTable Is Nothing Then Exit Sub
    If qtyMove <= 0 Then Exit Sub
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
    lr.Range.Cells(1, cRef).Value = refVal
    lr.Range.Cells(1, cItems).Value = itemVal
    lr.Range.Cells(1, cQty).Value = qtyMove
End Sub

' ===== helpers reused from modTS_Received =====
Private Function SheetExists(nameOrCode As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
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
    Dim wsInv As Worksheet: Set wsInv = SheetExists(SHEET_INV)
    If wsInv Is Nothing Then Exit Function
    On Error Resume Next
    Set GetInvSysTable = wsInv.ListObjects("invSys")
    On Error GoTo 0
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

Private Function EnsureInvSysItem(boxName As String, uom As String, location As String, descr As String, invLo As ListObject) As Long
    If invLo Is Nothing Then Exit Function
    EnsureInvSysRowSeed invLo
    Dim existingIdx As Long
    existingIdx = FindInvRowIndexByItem(invLo, boxName)
    Dim cRow As Long: cRow = ColumnIndex(invLo, "ROW")
    If existingIdx > 0 Then
        EnsureInvSysItem = NzLng(invLo.DataBodyRange.Cells(existingIdx, cRow).Value)
        If EnsureInvSysItem >= mNextInvSysRow Then
            mNextInvSysRow = EnsureInvSysItem + 1
        End If
        UpdateInvSysRow invLo.ListRows(existingIdx), boxName, uom, location, descr
        Exit Function
    End If

    Dim lr As ListRow: Set lr = invLo.ListRows.Add
    Dim newRowVal As Long: newRowVal = NextInvSysRowValue(invLo)
    EnsureInvSysItem = newRowVal
    UpdateInvSysRow lr, boxName, uom, location, descr, newRowVal
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
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
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

    Dim wsBOM As Worksheet: Set wsBOM = SheetExists(SHEET_BOM)
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
    If deltas Is Nothing Or deltas.Count = 0 Then Exit Function

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
            .Cells(1, 2).Value = Environ$("USERNAME")
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
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Public Function NzDbl(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzDbl = 0#
    Else
        NzDbl = CDbl(v)
    End If
End Function

Public Function NzLng(v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzLng = 0
    Else
        NzLng = CLng(v)
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
    modTS_Received.EnsureGeneratedButtons
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

    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
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
    On Error Resume Next
    If Not ThisWorkbook.Worksheets("ShipmentsTally") Is Nothing Then
        If Not ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("ShipmentsTally") Is Nothing Then
            ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("ShipmentsTally").AutoFilter.ShowAllData
        End If
        If Not ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("invSysData_Shipping") Is Nothing Then
            ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("invSysData_Shipping").AutoFilter.ShowAllData
        End If
    End If
    If Not ThisWorkbook.Worksheets("ReceivedTally") Is Nothing Then
        If Not ThisWorkbook.Worksheets("ReceivedTally").ListObjects("ReceivedTally") Is Nothing Then
            ThisWorkbook.Worksheets("ReceivedTally").ListObjects("ReceivedTally").AutoFilter.ShowAllData
        End If
        If Not ThisWorkbook.Worksheets("ReceivedTally").ListObjects("invSysData_Receiving") Is Nothing Then
            ThisWorkbook.Worksheets("ReceivedTally").ListObjects("invSysData_Receiving").AutoFilter.ShowAllData
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

    Set ws = ThisWorkbook.Worksheets("INVENTORY MANAGEMENT")
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
