VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cEventsPublication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mInventory As Collection, mGroups As Object, mSources As Object, mState As Collection
Private mInventoryReason As String
Private Const INVENTORY_FIELDS As String = "EventID|EventType|OccurredAtUTC|AppliedAtUTC|StationId|UserId|System_Key|SKU|QtyDelta|Location|Condition|Note"

Public Sub CaptureInventory(ByVal source As ListObject)
    Dim columns As Object, required As Object, names As Variant, name As Variant, column As ListColumn, values As Variant
    Dim rows As Collection, row As Object, r As Long, text As String, note As String
    On Error GoTo Failed
    mInventoryReason = "InvalidSchema": Set mInventory = Nothing
    If source Is Nothing Then mInventoryReason = "MissingSource": Exit Sub
    names = Split(INVENTORY_FIELDS, "|")
    Set columns = CreateObject("Scripting.Dictionary"): columns.CompareMode = vbTextCompare
    Set required = CreateObject("Scripting.Dictionary"): required.CompareMode = vbTextCompare
    For Each name In names: required.Add CStr(name), True: Next name
    For Each name In Array("UndoOfEventId", "AppliedSeq", "WarehouseId", "UOM"): required.Add CStr(name), True: Next name
    For Each column In source.ListColumns
        text = Trim$(column.Name)
        If required.Exists(text) Then
            If columns.Exists(text) Then Exit Sub
            columns.Add text, column.Index
        End If
    Next column
    For Each name In names
        If Not columns.Exists(name) Then Exit Sub
    Next name
    Set rows = New Collection
    If Not source.DataBodyRange Is Nothing Then
        values = source.DataBodyRange.Value2
        For r = 1 To UBound(values, 1)
            Set row = CreateObject("Scripting.Dictionary")
            For Each name In names
                text = FieldText(values(r, CLng(columns(name))), CStr(name))
                If name = "Note" Then note = text Else row.Add CStr(name), text
            Next name
            If row("EventID") = "" Or row("System_Key") = "" Then Exit Sub
            For Each name In Array("UndoOfEventId", "AppliedSeq", "WarehouseId")
                If columns.Exists(name) Then row.Add CStr(name), FieldText(values(r, CLng(columns(name))), CStr(name))
            Next name
            row.Add "Reference", FirstToken(note, Array("Reference", "Ref", "PO", "BOL", "ReceiptId", "DispositionRef"))
            row.Add "Item", FirstToken(note, Array("Item", "Box", "Package"))
            text = FirstToken(note, Array("UOM"))
            If columns.Exists("UOM") Then text = FieldText(values(r, CLng(columns("UOM"))), "Uom")
            row.Add "Uom", text
            rows.Add row
        Next r
    End If
    Set mInventory = rows: mInventoryReason = "OK"
    Exit Sub
Failed:
    mInventoryReason = "ReadFailed"
End Sub

Public Function Publish(ByVal warehouseId As String, ByVal runtimeRoot As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget, record As Object, records As Collection, reason As String, encoded As String
    Dim model As Object, coverage As Object, sources As Collection, groups As Collection, source As Variant
    Dim version As Long, collect As Boolean, visible As Boolean, notice As String, session As Long
    On Error GoTo Failed
    report = "Events publication deferred: a matching allowed warehouse target is unavailable."
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    If Not MatchesTarget(warehouseId, runtimeRoot) Then Exit Function
    session = modAuthSession.Version()
    report = "Events publication unavailable: tracking policy could not be validated."
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice) Then Exit Function
    Set mGroups = CreateObject("Scripting.Dictionary"): mGroups.CompareMode = vbBinaryCompare
    Set mSources = CreateObject("Scripting.Dictionary"): Set mState = New Collection
    If Not mInventory Is Nothing Then
        For Each record In mInventory
            If record.Exists("WarehouseId") Then
                If StrComp(CStr(record("WarehouseId")), warehouseId, vbBinaryCompare) <> 0 Then
                    Set mInventory = Nothing: mInventoryReason = "WarehouseMismatch": Exit For
                End If
            End If
        Next record
    End If
    AddSource "Inventory", "Warehouse", mInventory, mInventoryReason
    encoded = modDesignsDomainBridge.ReadDesignsEventsForPublication(warehouseId, runtimeRoot)
    Set records = DecodeSource(encoded, "Designs", warehouseId, reason)
    AddSource "Designs", "Warehouse", records, reason
    Set records = modActivityStore.ReadPublicationRecords(target, reason)
    AddSource "Activity", "Warehouse", records, reason
    For Each source In Array("ShippingBOM", "ShippingHolds")
        encoded = ReadShipping(CStr(source), warehouseId, runtimeRoot)
        Set records = DecodeSource(encoded, CStr(source), warehouseId, reason)
        AddSource CStr(source), IIf(source = "ShippingHolds", "Station profile", "Warehouse"), records, reason
    Next source
    Set groups = SelectedGroups()
    Set sources = New Collection
    For Each source In Array("Inventory", "Designs", "Activity", "ShippingBOM", "ShippingHolds")
        Set record = mSources(source)
        If record("Availability") = "Available" Then
            record("OmittedGroups") = record("AvailableGroups") - record("IncludedGroups")
            record("OmittedLines") = record("AvailableLines") - record("IncludedLines")
        End If
        sources.Add record
    Next source
    Set coverage = CreateObject("Scripting.Dictionary"): coverage.Add "Sources", sources
    Set model = CreateObject("Scripting.Dictionary")
    model.Add "SchemaVersion", 1
    model.Add "PackageSetVersion", CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value)
    model.Add "BuildIdentity", CStr(ThisWorkbook.CustomDocumentProperties("invSysBuildIdentity").Value)
    model.Add "WarehouseId", warehouseId: model.Add "PublicationId", modTrainingWire.NewId()
    model.Add "PublishedAtUTC", modTrainingWire.UtcTimestamp(): model.Add "PolicyVersion", version
    model.Add "Groups", groups: model.Add "CurrentState", mState: model.Add "Coverage", coverage
    If Not MatchesTarget(warehouseId, runtimeRoot) Or session <> modAuthSession.Version() Then GoTo Failed
    Publish = modEventsPublicationStore.WriteModel(model, runtimeRoot, report)
    Exit Function
Failed:
    report = "Events publication failed; the prior complete Events artifact was retained."
End Function

Private Function MatchesTarget(ByVal warehouseId As String, ByVal root As String) As Boolean
    If Not modTrainingWire.ValidSegment(warehouseId) Then Exit Function
    If Not modNasConnection.IsCurrentTargetAllowed(True) Then Exit Function
    MatchesTarget = (StrComp(warehouseId, modNasConnection.GetCurrentTargetWarehouseId(), vbBinaryCompare) = 0 And _
        StrComp(modConfig.NormalizeFolderPathForRuntime(root), modConfig.NormalizeFolderPathForRuntime(modNasConnection.GetCurrentTargetRuntimeRoot()), vbTextCompare) = 0)
End Function

Private Sub AddSource(ByVal source As String, ByVal scope As String, ByVal rows As Collection, ByVal reason As String)
    Dim info As Object, state As Object, row As Object, before As Long, name As Variant, stamp As String
    Set info = CreateObject("Scripting.Dictionary")
    info.Add "Source", source: info.Add "Scope", scope
    info.Add "Availability", IIf(rows Is Nothing, "Unavailable", "Available")
    info.Add "Explanation", ReasonText(reason, scope)
    For Each name In Array("AvailableGroups", "IncludedGroups", "OmittedGroups", "AvailableLines", "IncludedLines", "OmittedLines")
        If rows Is Nothing Then info.Add CStr(name), "Unavailable" Else info.Add CStr(name), 0&
    Next name
    info.Add "IncludedFrom", "": info.Add "IncludedTo", ""
    mSources.Add source, info
    If rows Is Nothing Then Exit Sub
    info("AvailableLines") = rows.Count
    before = mGroups.Count
    If source = "ShippingBOM" Or source = "ShippingHolds" Then
        If rows.Count > 0 Then
            Set state = CreateObject("Scripting.Dictionary")
            state.Add "Source", source: state.Add "SourceKind", "Current state"
            state.Add "TimeProvenance", "Recorded time (zone unavailable)": state.Add "Lines", rows
            mState.Add state
            info("AvailableGroups") = 1&: info("IncludedGroups") = 1&: info("IncludedLines") = rows.Count
        End If
    Else
        For Each row In rows
            If source = "Activity" Then
                AddGroup source, "User activity", row("ActivityId"), row("OccurredAtUTC"), "Verified UTC", row, (row("OutcomeCode") <> "REQUESTED")
            Else
                stamp = CStr(row("OccurredAtUTC"))
                If stamp = "" Then stamp = CStr(row("AppliedAtUTC"))
                If CStr(row("AppliedAtUTC")) <> "" And StrComp(CStr(row("AppliedAtUTC")), stamp, vbBinaryCompare) < 0 Then stamp = CStr(row("AppliedAtUTC"))
                AddGroup source, "Business event", row("EventID"), stamp, "Recorded time (zone unavailable)", row, (CStr(row("AppliedAtUTC")) <> "")
            End If
        Next row
        info("AvailableGroups") = mGroups.Count - before
    End If
End Sub

Private Sub AddGroup(ByVal source As String, ByVal kind As String, ByVal id As String, ByVal stamp As String, ByVal zone As String, ByVal row As Object, ByVal outcome As Boolean)
    Dim key As String, group As Object, lines As Collection, outcomes As Collection
    If id = "" Then Err.Raise 5
    key = source & Chr$(30) & id
    If mGroups.Exists(key) Then
        Set group = mGroups(key)
        If stamp <> "" And (group("RecordedAt") = "" Or StrComp(stamp, group("RecordedAt"), vbBinaryCompare) < 0) Then group("RecordedAt") = stamp
    Else
        Set group = CreateObject("Scripting.Dictionary")
        Set lines = New Collection: Set outcomes = New Collection
        group.Add "Source", source: group.Add "SourceId", id: group.Add "SourceKind", kind
        group.Add "RecordedAt", stamp: group.Add "TimeProvenance", zone
        group.Add "Lines", lines: group.Add "Outcomes", outcomes
        mGroups.Add key, group
    End If
    group("Lines").Add row
    If outcome Then group("Outcomes").Add row
End Sub

Private Function SelectedGroups() As Collection
    Dim keys As Variant, result As Collection, group As Object, info As Object, i As Long, stamp As String
    Set result = New Collection
    If mGroups.Count > 0 Then
        keys = mGroups.Keys: SortKeys keys, LBound(keys), UBound(keys)
        For i = LBound(keys) To UBound(keys)
            If result.Count = 5000 Then Exit For
            Set group = mGroups(keys(i)): result.Add group
            Set info = mSources(group("Source"))
            info("IncludedGroups") = info("IncludedGroups") + 1&
            info("IncludedLines") = info("IncludedLines") + group("Lines").Count
            stamp = group("RecordedAt")
            If stamp <> "" Then
                If info("IncludedFrom") = "" Or StrComp(stamp, info("IncludedFrom"), vbBinaryCompare) < 0 Then info("IncludedFrom") = stamp
                If info("IncludedTo") = "" Or StrComp(stamp, info("IncludedTo"), vbBinaryCompare) > 0 Then info("IncludedTo") = stamp
            End If
        Next i
    End If
    Set SelectedGroups = result
End Function

Private Sub SortKeys(ByRef keys As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As String, swap As Variant
    i = low: j = high: pivot = keys((low + high) \ 2)
    Do While i <= j
        Do While Compare(keys(i), pivot) < 0: i = i + 1: Loop
        Do While Compare(keys(j), pivot) > 0: j = j - 1: Loop
        If i <= j Then
            swap = keys(i): keys(i) = keys(j): keys(j) = swap
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then SortKeys keys, low, j
    If i < high Then SortKeys keys, i, high
End Sub

Private Function Compare(ByVal left As String, ByVal right As String) As Long
    Dim a As Object, b As Object
    Set a = mGroups(left): Set b = mGroups(right)
    Compare = -StrComp(a("RecordedAt"), b("RecordedAt"), vbBinaryCompare)
    If Compare = 0 Then Compare = StrComp(a("SourceKind"), b("SourceKind"), vbBinaryCompare)
    If Compare = 0 Then Compare = StrComp(a("SourceId"), b("SourceId"), vbBinaryCompare)
    If Compare = 0 Then Compare = StrComp(a("Source"), b("Source"), vbBinaryCompare)
End Function

Private Function ReadShipping(ByVal source As String, ByVal warehouse As String, ByVal root As String) As String
    On Error GoTo Failed
    ReadShipping = CStr(Application.Run("'invSys.Operations.xlam'!modShippingPublicationSource.ReadForPublication", source, warehouse, root))
Failed:
End Function

Private Function DecodeSource(ByVal encoded As String, ByVal source As String, ByVal warehouse As String, ByRef reason As String) As Collection
    Dim lines As Variant, header As Variant, names As Variant, values As Variant, rows As Collection, row As Object
    Dim r As Long, c As Long, count As Long, name As String
    On Error GoTo Failed
    reason = "ReadFailed"
    lines = Split(encoded, vbCrLf)
    If UBound(lines) < 1 Then Exit Function
    header = Split(lines(0), vbTab)
    If UBound(header) <> 6 Then Exit Function
    If header(0) <> "EVTSRC1" Or header(1) <> source Or Unescape(CStr(header(3))) <> warehouse Then Exit Function
    If header(2) <> "Available" Then reason = CStr(header(6)): Exit Function
    If header(5) <> "ReadOnly" And header(5) <> "Borrowed" Then Exit Function
    If header(6) <> "OK" Then Exit Function
    If CStr(lines(1)) <> Replace$(SourceFields(source), "|", vbTab) Then reason = "InvalidSchema": Exit Function
    If Not IsNumeric(header(4)) Then Exit Function
    count = CLng(header(4))
    If count < 0 Or CStr(count) <> header(4) Or UBound(lines) <> count + 1 Then Exit Function
    names = Split(lines(1), vbTab): Set rows = New Collection
    For r = 2 To UBound(lines)
        values = Split(lines(r), vbTab)
        If UBound(values) <> UBound(names) Then Exit Function
        Set row = CreateObject("Scripting.Dictionary")
        For c = 0 To UBound(names)
            name = CStr(names(c))
            If name = "" Or row.Exists(name) Then Exit Function
            row.Add name, Unescape(CStr(values(c)))
        Next c
        rows.Add row
    Next r
    reason = "OK": Set DecodeSource = rows
    Exit Function
Failed:
    reason = "InvalidSchema"
End Function

Private Function SourceFields(ByVal source As String) As String
    Select Case source
        Case "Designs"
            SourceFields = "EventID|UndoOfEventId|AppliedSeq|EventType|OccurredAtUTC|AppliedAtUTC|WarehouseId|StationId|UserId|DefinitionType|DefinitionId|DefinitionVersion|Note"
        Case "ShippingBOM"
            SourceFields = "PackageSystemKey|PackageItem|PackageUOM|PackageLocation|PackageDescription|BomVersion|BomVersionLabel|IsActive|EffectiveFromUTC|EffectiveToUTC|RetiredAtUTC|ComponentSystemKey|ComponentItemCode|ComponentItem|ComponentQty|ComponentUOM|ComponentLocation|ComponentDescription|UpdatedAtUTC|UpdatedBy"
        Case "ShippingHolds"
            SourceFields = "Ref|Item|Qty|UOM|Location|Description|Area|Carrier|ShipmentLineId|ReserveEventId|System_Key"
    End Select
End Function

Private Function Unescape(ByVal value As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch = "\" Then
            i = i + 1: ch = Mid$(value, i, 1)
            Select Case ch
                Case "\"
                Case "t": ch = vbTab
                Case "r": ch = vbCr
                Case "n": ch = vbLf
                Case Else: Err.Raise 5
            End Select
        End If
        Unescape = Unescape & ch
    Next i
End Function

Private Function FieldText(ByVal value As Variant, ByVal field As String) As String
    If IsError(value) Or IsNull(value) Then Err.Raise 5
    If Right$(field, 3) = "UTC" And Not IsEmpty(value) Then
        If IsNumeric(value) Or IsDate(value) Then
            FieldText = Format$(CDate(value), "yyyy-mm-dd\Thh:nn:ss"): Exit Function
        End If
    End If
    FieldText = CStr(value)
End Function

Private Function FirstToken(ByVal note As String, ByVal names As Variant) As String
    Dim name As Variant, part As Variant, prefix As String, text As String
    For Each name In names
        prefix = CStr(name) & "="
        For Each part In Split(note, ";")
            text = Trim$(CStr(part))
            If StrComp(Left$(text, Len(prefix)), prefix, vbTextCompare) = 0 Then
                FirstToken = Trim$(Mid$(text, Len(prefix) + 1)): Exit Function
            End If
        Next part
    Next name
End Function


Private Function ReasonText(ByVal reason As String, ByVal scope As String) As String
    Select Case reason
        Case "OK": ReasonText = "Source available; " & scope & " scope."
        Case "MissingSource": ReasonText = "Source unavailable: no readable source has been published by its owner."
        Case "UnsavedSource": ReasonText = "Source unavailable: its workbook contains unsaved changes."
        Case "ContextMismatch", "WarehouseMismatch": ReasonText = "Source unavailable: warehouse context does not match."
        Case "InvalidSchema", "InvalidRecord": ReasonText = "Source unavailable: saved evidence failed validation."
        Case Else: ReasonText = "Source unavailable: the owning read did not complete."
    End Select
End Function
