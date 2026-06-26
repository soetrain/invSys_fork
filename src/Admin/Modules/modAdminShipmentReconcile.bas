Attribute VB_Name = "modAdminShipmentReconcile"
Option Explicit

Private Const ADMIN_RECON_EVENT_TYPE As String = "ADMIN_SHIPMENT_RECONCILE"
Private Const SERVER_SHIP_EVENT_TYPE As String = "SHIP"
Private Const SERVER_SHIP_RESERVE_EVENT_TYPE As String = "SHIP_RESERVE"
Private Const SERVER_SHIP_RELEASE_EVENT_TYPE As String = "SHIP_RELEASE"
Private Const FLAG_PROJECTED_BELOW_NAS_NO_ACTIVE_ROW As String = "PROJECTED_BELOW_NAS_NO_ACTIVE_ROW"
Private Const FLAG_LOCKED_WITH_NO_ACTIVE_ROW As String = "LOCKED_WITH_NO_ACTIVE_ROW"
Private Const FLAG_COMPLETED_RESERVATION_STILL_ACTIVE As String = "COMPLETED_RESERVATION_STILL_ACTIVE"
Private Const FLAG_NAS_INCREASED_AFTER_SHIP As String = "NAS_INCREASED_AFTER_SHIP"

Public Sub OpenShipmentReconcileTool()
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim sku As String
    Dim selectionText As String
    Dim selectedIndex As Long
    Dim recentRows As Collection
    Dim selectedRow As Object
    Dim report As String
    Dim eventIdOut As String
    Dim correctedShipEventId As String
    Dim narrative As String
    Dim inventoryWb As Workbook
    Dim currentNasQty As Double
    Dim qtyAfterShip As Double
    Dim adjustmentDelta As Double
    Dim rowVal As Long
    Dim locationVal As String
    Dim target As WarehouseTarget

    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then
        warehouseId = Trim$(target.WarehouseId)
        stationId = Trim$(target.StationId)
    End If
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetWarehouseId())
    If stationId = "" Then stationId = Trim$(modConfig.GetStationId())
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetString("WarehouseId", ""))
    If stationId = "" Then stationId = Trim$(modConfig.GetString("StationId", ""))
    userId = Trim$(modRoleEventWriter.ResolveCurrentUserId())

    Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(warehouseId)
    If inventoryWb Is Nothing Then
        MsgBox "Canonical inventory workbook could not be opened.", vbExclamation, "invSys Admin"
        Exit Sub
    End If

    Set recentRows = BuildRecentShipmentSentRows(inventoryWb, 20)
    If recentRows Is Nothing Or recentRows.Count = 0 Then
        MsgBox "No server shipment deduction log entries (SHIP) were found in tblInventoryLog." & vbCrLf & vbCrLf & _
               ShipmentLogDiagnosticsText(inventoryWb), vbInformation, "invSys Admin - Shipment Reconcile"
        Exit Sub
    End If

    selectionText = Trim$(InputBox( _
        "Enter the number of the server shipment deduction log line to inspect." & vbCrLf & vbCrLf & _
        BuildRecentShipmentSentLogText(inventoryWb, 20), _
        "invSys Admin - Shipment Reconcile"))
    If selectionText = "" Then Exit Sub
    If Not IsNumeric(selectionText) Then
        MsgBox "Enter a number from the recent server shipment deduction list.", vbExclamation, "invSys Admin - Shipment Reconcile"
        Exit Sub
    End If

    selectedIndex = CLng(Val(selectionText))
    If selectedIndex < 1 Or selectedIndex > recentRows.Count Then
        MsgBox "Selection must be between 1 and " & CStr(recentRows.Count) & ".", vbExclamation, "invSys Admin - Shipment Reconcile"
        Exit Sub
    End If

    Set selectedRow = recentRows(selectedIndex)
    sku = CStr(selectedRow("SKU"))
    correctedShipEventId = CStr(selectedRow("EventID"))

    If Not DetectNasIncreaseAfterShipEvent(inventoryWb, correctedShipEventId, sku, currentNasQty, qtyAfterShip, report) Then
        If report = "" Then report = "No NAS increase after Shipments Sent event " & correctedShipEventId & " was detected for " & sku & "."
        MsgBox report, vbInformation, "invSys Admin - Shipment Reconcile"
        Exit Sub
    End If

    adjustmentDelta = qtyAfterShip - currentNasQty
    rowVal = ResolveRowForSku(inventoryWb, sku)
    locationVal = ResolveLatestLocationForSku(inventoryWb, sku)

    narrative = Trim$(InputBox( _
        "Type one line explaining why this admin correction is being made." & vbCrLf & vbCrLf & _
        "Corrects server shipment deduction event: " & correctedShipEventId & vbCrLf & _
        "Current NAS Qty: " & Format$(currentNasQty, "0.########") & vbCrLf & _
        "Expected Qty after that event: " & Format$(qtyAfterShip, "0.########") & vbCrLf & _
        "Queued correction delta: " & Format$(adjustmentDelta, "0.########"), _
        "invSys Admin - Confirm Shipment Reconcile"))
    If narrative = "" Then
        MsgBox "Shipment reconcile cancelled. A human repair narrative is required.", vbExclamation, "invSys Admin"
        Exit Sub
    End If

    If QueueAdminShipmentReconcileEvent(warehouseId, stationId, userId, sku, rowVal, locationVal, adjustmentDelta, _
                                        correctedShipEventId, narrative, FLAG_NAS_INCREASED_AFTER_SHIP, _
                                        Nothing, eventIdOut, report) Then
        MsgBox "Shipment reconcile event queued." & vbCrLf & _
               "EventID: " & eventIdOut & vbCrLf & _
               "Corrects: " & correctedShipEventId & vbCrLf & _
               "Delta: " & Format$(adjustmentDelta, "0.########"), _
               vbInformation, "invSys Admin"
    Else
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Public Function QueueAdminShipmentReconcileEvent(ByVal warehouseId As String, _
                                                 ByVal stationId As String, _
                                                 ByVal userId As String, _
                                                 ByVal sku As String, _
                                                 ByVal rowVal As Long, _
                                                 ByVal locationVal As String, _
                                                 ByVal signedQtyDelta As Double, _
                                                 ByVal correctedShipEventId As String, _
                                                 ByVal repairNarrative As String, _
                                                 Optional ByVal mismatchFlag As String = "", _
                                                 Optional ByVal targetInboxWb As Workbook = Nothing, _
                                                 Optional ByRef eventIdOut As String = "", _
                                                 Optional ByRef report As String = "") As Boolean
    Dim item As Object
    Dim payloadJson As String

    If Not ValidateShipmentReconcileRequest(sku, signedQtyDelta, correctedShipEventId, repairNarrative, report) Then Exit Function
    sku = Trim$(sku)
    correctedShipEventId = Trim$(correctedShipEventId)
    repairNarrative = Trim$(repairNarrative)

    Set item = modRoleEventWriter.CreatePayloadItem(rowVal, sku, signedQtyDelta, locationVal, repairNarrative, "RECONCILE")
    item("CorrectedShipEventId") = correctedShipEventId
    item("RepairNarrative") = repairNarrative
    If Trim$(mismatchFlag) <> "" Then item("MismatchFlag") = Trim$(mismatchFlag)
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)

    QueueAdminShipmentReconcileEvent = modRoleEventWriter.QueuePayloadEvent( _
        ADMIN_RECON_EVENT_TYPE, warehouseId, stationId, userId, payloadJson, repairNarrative, _
        correctedShipEventId, "", Now, targetInboxWb, eventIdOut, report)
End Function

Public Function ValidateShipmentReconcileRequest(ByVal sku As String, _
                                                 ByVal signedQtyDelta As Double, _
                                                 ByVal correctedShipEventId As String, _
                                                 ByVal repairNarrative As String, _
                                                 Optional ByRef report As String = "") As Boolean
    If Trim$(sku) = "" Then
        report = "SKU is required."
        Exit Function
    End If
    If signedQtyDelta = 0 Then
        report = "Signed adjustment delta cannot be zero."
        Exit Function
    End If
    If Trim$(correctedShipEventId) = "" Then
        report = "CorrectedShipEventId is required."
        Exit Function
    End If
    If Trim$(repairNarrative) = "" Then
        report = "A one-line human repair narrative is required."
        Exit Function
    End If

    ValidateShipmentReconcileRequest = True
End Function

Public Function BuildShipmentMismatchFlags(ByVal projectedQty As Double, _
                                           ByVal nasQty As Double, _
                                           ByVal lockedQty As Double, _
                                           ByVal hasActiveShipmentRow As Boolean, _
                                           ByVal completedReservationStillActive As Boolean, _
                                           ByVal nasIncreasedAfterShip As Boolean) As String
    Dim flags As String

    If projectedQty < nasQty And Not hasActiveShipmentRow Then AddFlag flags, FLAG_PROJECTED_BELOW_NAS_NO_ACTIVE_ROW
    If lockedQty > 0 And Not hasActiveShipmentRow Then AddFlag flags, FLAG_LOCKED_WITH_NO_ACTIVE_ROW
    If completedReservationStillActive Then AddFlag flags, FLAG_COMPLETED_RESERVATION_STILL_ACTIVE
    If nasIncreasedAfterShip Then AddFlag flags, FLAG_NAS_INCREASED_AFTER_SHIP
    BuildShipmentMismatchFlags = flags
End Function

Public Function BuildRecentShipmentSentLogText(ByVal inventoryWb As Workbook, _
                                               Optional ByVal limitCount As Long = 20) As String
    Dim rows As Collection
    Dim rowData As Object
    Dim i As Long
    Dim lineText As String

    Set rows = BuildRecentShipmentSentRows(inventoryWb, limitCount)
    If rows Is Nothing Or rows.Count = 0 Then
        BuildRecentShipmentSentLogText = "No server shipment deduction log entries found."
        Exit Function
    End If

    For i = 1 To rows.Count
        Set rowData = rows(i)
        lineText = CStr(i) & ". " & CStr(rowData("EventID")) & _
                   " | " & CStr(rowData("EventType")) & _
                   " | " & CStr(rowData("OccurredAtUTC")) & _
                   " | " & CStr(rowData("SKU")) & _
                   " | Qty " & Format$(Abs(CDbl(rowData("QtyDelta"))), "0.########")
        If CStr(rowData("Location")) <> "" Then lineText = lineText & " | " & CStr(rowData("Location"))
        If BuildRecentShipmentSentLogText <> "" Then BuildRecentShipmentSentLogText = BuildRecentShipmentSentLogText & vbCrLf
        BuildRecentShipmentSentLogText = BuildRecentShipmentSentLogText & lineText
    Next i
End Function

Public Function DetectNasIncreaseAfterLastShip(ByVal inventoryWb As Workbook, _
                                               ByVal sku As String, _
                                               ByRef correctedShipEventId As String, _
                                               ByRef currentNasQty As Double, _
                                               ByRef qtyAfterShip As Double, _
                                               Optional ByRef report As String = "") As Boolean
    Dim loLog As ListObject
    Dim loSku As ListObject
    Dim rowIndex As Long
    Dim latestRow As Long
    Dim latestSeq As Long
    Dim seqVal As Long
    Dim latestStamp As Date
    Dim rowStamp As Date

    sku = Trim$(sku)
    If inventoryWb Is Nothing Or sku = "" Then
        report = "Inventory workbook and SKU are required."
        Exit Function
    End If

    Set loLog = FindListObjectByNameReconcile(inventoryWb, "tblInventoryLog")
    Set loSku = FindListObjectByNameReconcile(inventoryWb, "tblSkuBalance")
    If loLog Is Nothing Or loSku Is Nothing Then
        report = "Inventory log or SKU balance table was not found."
        Exit Function
    End If

    currentNasQty = GetSkuBalanceQtyReconcile(loSku, sku)
    If Not loLog.DataBodyRange Is Nothing Then
        For rowIndex = 1 To loLog.ListRows.Count
            If StrComp(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "SKU")), sku, vbTextCompare) = 0 _
               And IsShipmentDeductionEventType(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "EventType"))) Then
                seqVal = CLng(NzDblReconcile(GetCellByColumnReconcile(loLog, rowIndex, "AppliedSeq")))
                If IsDate(GetCellByColumnReconcile(loLog, rowIndex, "OccurredAtUTC")) Then
                    rowStamp = CDate(GetCellByColumnReconcile(loLog, rowIndex, "OccurredAtUTC"))
                Else
                    rowStamp = CDate(0)
                End If
                If latestRow = 0 Or rowStamp > latestStamp Or (rowStamp = latestStamp And seqVal >= latestSeq) Then
                    latestRow = rowIndex
                    latestSeq = seqVal
                    latestStamp = rowStamp
                End If
            End If
        Next rowIndex
    End If

    If latestRow = 0 Then
        report = "No SHIP event was found for " & sku & "."
        Exit Function
    End If

    correctedShipEventId = SafeTrimReconcile(GetCellByColumnReconcile(loLog, latestRow, "EventID"))
    If latestSeq <= 0 Then latestSeq = latestRow
    qtyAfterShip = SumQtyThroughAppliedSeqReconcile(loLog, sku, latestSeq, latestRow)
    DetectNasIncreaseAfterLastShip = (currentNasQty > qtyAfterShip)
    If Not DetectNasIncreaseAfterLastShip Then
        report = "Current NAS Qty is not higher than the log balance after " & correctedShipEventId & "."
    End If
End Function

Public Function DetectNasIncreaseAfterShipEvent(ByVal inventoryWb As Workbook, _
                                                ByVal shipEventId As String, _
                                                ByVal sku As String, _
                                                ByRef currentNasQty As Double, _
                                                ByRef qtyAfterShip As Double, _
                                                Optional ByRef report As String = "") As Boolean
    Dim loLog As ListObject
    Dim loSku As ListObject
    Dim shipRow As Long
    Dim appliedSeq As Long

    shipEventId = Trim$(shipEventId)
    sku = Trim$(sku)
    If inventoryWb Is Nothing Or shipEventId = "" Then
        report = "Inventory workbook and shipment deduction EventID are required."
        Exit Function
    End If

    Set loLog = FindListObjectByNameReconcile(inventoryWb, "tblInventoryLog")
    Set loSku = FindListObjectByNameReconcile(inventoryWb, "tblSkuBalance")
    If loLog Is Nothing Or loSku Is Nothing Then
        report = "Inventory log or SKU balance table was not found."
        Exit Function
    End If

    shipRow = FindShipmentLogRowByEventReconcile(loLog, shipEventId, sku)
    If shipRow = 0 Then
        report = "Shipment deduction EventID was not found in tblInventoryLog: " & shipEventId
        Exit Function
    End If

    If sku = "" Then sku = SafeTrimReconcile(GetCellByColumnReconcile(loLog, shipRow, "SKU"))
    currentNasQty = GetSkuBalanceQtyReconcile(loSku, sku)
    appliedSeq = CLng(NzDblReconcile(GetCellByColumnReconcile(loLog, shipRow, "AppliedSeq")))
    If appliedSeq <= 0 Then appliedSeq = shipRow
    qtyAfterShip = SumQtyThroughAppliedSeqReconcile(loLog, sku, appliedSeq, shipRow)

    DetectNasIncreaseAfterShipEvent = (currentNasQty > qtyAfterShip)
    If Not DetectNasIncreaseAfterShipEvent Then
        report = "Current NAS Qty is not higher than the log balance after " & shipEventId & "."
    End If
End Function

Private Sub AddFlag(ByRef flags As String, ByVal flagText As String)
    If flags <> "" Then flags = flags & ";"
    flags = flags & flagText
End Sub

Private Function SumQtyThroughAppliedSeqReconcile(ByVal loLog As ListObject, _
                                                  ByVal sku As String, _
                                                  ByVal appliedSeqLimit As Long, _
                                                  ByVal rowLimit As Long) As Double
    Dim rowIndex As Long
    Dim seqVal As Long

    If loLog Is Nothing Or loLog.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To loLog.ListRows.Count
        If StrComp(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "SKU")), sku, vbTextCompare) = 0 Then
            seqVal = CLng(NzDblReconcile(GetCellByColumnReconcile(loLog, rowIndex, "AppliedSeq")))
            If (seqVal > 0 And seqVal <= appliedSeqLimit) Or (seqVal = 0 And rowIndex <= rowLimit) Then
                SumQtyThroughAppliedSeqReconcile = SumQtyThroughAppliedSeqReconcile + NzDblReconcile(GetCellByColumnReconcile(loLog, rowIndex, "QtyDelta"))
            End If
        End If
    Next rowIndex
End Function

Private Function BuildRecentShipmentSentRows(ByVal inventoryWb As Workbook, _
                                             ByVal limitCount As Long) As Collection
    Dim loLog As ListObject
    Dim rows As Collection
    Dim rowIndex As Long
    Dim rowData As Object
    Dim eventType As String
    Dim qtyDelta As Double

    Set rows = New Collection
    Set BuildRecentShipmentSentRows = rows
    If inventoryWb Is Nothing Then Exit Function
    If limitCount <= 0 Then limitCount = 20

    Set loLog = FindListObjectByNameReconcile(inventoryWb, "tblInventoryLog")
    If loLog Is Nothing Or loLog.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = loLog.ListRows.Count To 1 Step -1
        eventType = UCase$(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "EventType")))
        qtyDelta = NzDblReconcile(GetCellByColumnReconcile(loLog, rowIndex, "QtyDelta"))
        If IsShipmentDeductionEventType(eventType) And qtyDelta < 0 Then
            Set rowData = CreateObject("Scripting.Dictionary")
            rowData.CompareMode = vbTextCompare
            rowData("EventID") = SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "EventID"))
            rowData("EventType") = eventType
            rowData("SKU") = SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "SKU"))
            rowData("QtyDelta") = qtyDelta
            rowData("Location") = SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "Location"))
            rowData("OccurredAtUTC") = FormatLogDateReconcile(GetCellByColumnReconcile(loLog, rowIndex, "OccurredAtUTC"))
            rows.Add rowData
            If rows.Count >= limitCount Then Exit For
        End If
    Next rowIndex
End Function

Private Function FindShipmentLogRowByEventReconcile(ByVal loLog As ListObject, _
                                                    ByVal shipEventId As String, _
                                                    ByVal sku As String) As Long
    Dim rowIndex As Long

    If loLog Is Nothing Or loLog.DataBodyRange Is Nothing Then Exit Function
    shipEventId = Trim$(shipEventId)
    sku = Trim$(sku)
    For rowIndex = 1 To loLog.ListRows.Count
        If StrComp(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "EventID")), shipEventId, vbTextCompare) = 0 _
           And IsShipmentDeductionEventType(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "EventType"))) Then
            If sku = "" Or StrComp(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "SKU")), sku, vbTextCompare) = 0 Then
                FindShipmentLogRowByEventReconcile = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Public Function ShipmentLogDiagnosticsText(ByVal inventoryWb As Workbook) As String
    Dim loLog As ListObject
    Dim rowIndex As Long
    Dim eventType As String
    Dim shipCount As Long
    Dim reserveCount As Long
    Dim deductionCount As Long
    Dim pathText As String
    Dim warehouseId As String
    Dim stationId As String
    Dim pendingCount As Long
    Dim matchingPending As Long
    Dim inboxReport As String
    Dim inboxError As String
    Dim stagedCount As Long
    Dim matchingStaged As Long
    Dim stagedReport As String
    Dim stagedError As String
    Dim target As WarehouseTarget

    If inventoryWb Is Nothing Then
        ShipmentLogDiagnosticsText = "Inventory workbook: not open."
        Exit Function
    End If

    pathText = inventoryWb.FullName
    If pathText = "" Then pathText = inventoryWb.Name
    Set loLog = FindListObjectByNameReconcile(inventoryWb, "tblInventoryLog")
    If loLog Is Nothing Then
        ShipmentLogDiagnosticsText = "Inventory workbook: " & pathText & vbCrLf & _
                                     "tblInventoryLog: missing."
        Exit Function
    End If

    If Not loLog.DataBodyRange Is Nothing Then
        For rowIndex = 1 To loLog.ListRows.Count
            eventType = UCase$(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "EventType")))
            If eventType = SERVER_SHIP_EVENT_TYPE Then shipCount = shipCount + 1
            If eventType = SERVER_SHIP_RESERVE_EVENT_TYPE Then reserveCount = reserveCount + 1
            If IsShipmentDeductionEventType(eventType) And NzDblReconcile(GetCellByColumnReconcile(loLog, rowIndex, "QtyDelta")) < 0 Then
                deductionCount = deductionCount + 1
            End If
        Next rowIndex
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then
        warehouseId = Trim$(target.WarehouseId)
        stationId = Trim$(target.StationId)
    End If
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetWarehouseId())
    If stationId = "" Then stationId = Trim$(modConfig.GetStationId())
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetString("WarehouseId", ""))
    If stationId = "" Then stationId = Trim$(modConfig.GetString("StationId", ""))
    inboxReport = modRoleEventWriter.DescribeInboxPendingRows(SERVER_SHIP_EVENT_TYPE, warehouseId, stationId, "", pendingCount, matchingPending, inboxError)
    stagedReport = modRoleEventWriter.DescribeLocalStagedInboxRows(SERVER_SHIP_EVENT_TYPE & "," & SERVER_SHIP_RESERVE_EVENT_TYPE & "," & SERVER_SHIP_RELEASE_EVENT_TYPE, _
                                                                    warehouseId, _
                                                                    stationId, _
                                                                    stagedCount, _
                                                                    matchingStaged, _
                                                                    stagedError)

    ShipmentLogDiagnosticsText = "Inventory workbook: " & pathText & vbCrLf & _
                                 "tblInventoryLog rows: " & CStr(loLog.ListRows.Count) & vbCrLf & _
                                 "SHIP rows: " & CStr(shipCount) & vbCrLf & _
                                 "SHIP_RESERVE rows: " & CStr(reserveCount) & vbCrLf & _
                                 "shipment deduction rows: " & CStr(deductionCount) & vbCrLf & _
                                 "NAS shipping inbox pending: " & CStr(pendingCount) & vbCrLf & _
                                 "NAS shipping inbox: " & IIf(inboxReport <> "", inboxReport, inboxError) & vbCrLf & _
                                 "local staged shipment rows: " & CStr(matchingStaged) & vbCrLf & _
                                 "local staging: " & IIf(stagedReport <> "", stagedReport, stagedError)
End Function

Private Function IsShipmentDeductionEventType(ByVal eventType As String) As Boolean
    eventType = UCase$(Trim$(eventType))
    IsShipmentDeductionEventType = (eventType = SERVER_SHIP_EVENT_TYPE)
End Function

Private Function FormatLogDateReconcile(ByVal valueIn As Variant) As String
    If IsDate(valueIn) Then
        FormatLogDateReconcile = Format$(CDate(valueIn), "yyyy-mm-dd hh:nn:ss")
    Else
        FormatLogDateReconcile = SafeTrimReconcile(valueIn)
    End If
End Function

Private Function GetSkuBalanceQtyReconcile(ByVal loSku As ListObject, ByVal sku As String) As Double
    Dim rowIndex As Long

    If loSku Is Nothing Or loSku.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To loSku.ListRows.Count
        If StrComp(SafeTrimReconcile(GetCellByColumnReconcile(loSku, rowIndex, "SKU")), sku, vbTextCompare) = 0 Then
            GetSkuBalanceQtyReconcile = NzDblReconcile(GetCellByColumnReconcile(loSku, rowIndex, "QtyOnHand"))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function ResolveRowForSku(ByVal inventoryWb As Workbook, ByVal sku As String) As Long
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = FindListObjectByNameReconcile(inventoryWb, "tblSkuCatalog")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(SafeTrimReconcile(GetCellByColumnReconcile(lo, rowIndex, "SKU")), sku, vbTextCompare) = 0 _
           Or StrComp(SafeTrimReconcile(GetCellByColumnReconcile(lo, rowIndex, "ITEM_CODE")), sku, vbTextCompare) = 0 Then
            ResolveRowForSku = CLng(NzDblReconcile(GetCellByColumnReconcile(lo, rowIndex, "ROW")))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function ResolveLatestLocationForSku(ByVal inventoryWb As Workbook, ByVal sku As String) As String
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = FindListObjectByNameReconcile(inventoryWb, "tblInventoryLog")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = lo.ListRows.Count To 1 Step -1
        If StrComp(SafeTrimReconcile(GetCellByColumnReconcile(lo, rowIndex, "SKU")), sku, vbTextCompare) = 0 Then
            ResolveLatestLocationForSku = SafeTrimReconcile(GetCellByColumnReconcile(lo, rowIndex, "Location"))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FindListObjectByNameReconcile(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameReconcile = ws.ListObjects(tableName)
        If Not FindListObjectByNameReconcile Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function GetCellByColumnReconcile(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long

    idx = GetColumnIndexReconcile(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnReconcile = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Function GetColumnIndexReconcile(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexReconcile = i
            Exit Function
        End If
    Next i
End Function

Private Function SafeTrimReconcile(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimReconcile = Trim$(CStr(valueIn))
End Function

Private Function NzDblReconcile(ByVal valueIn As Variant) As Double
    On Error Resume Next
    If IsNumeric(valueIn) Then NzDblReconcile = CDbl(valueIn)
End Function
