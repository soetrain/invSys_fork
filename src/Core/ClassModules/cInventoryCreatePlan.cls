VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cInventoryCreatePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Exact transient expectations for one owning INVENTORY_CREATE payload.
Private mKey As String, mSku As String, mQty As Double
Private mLocation As String, mCondition As String, mAttributes As String

Public Function Plan(ByVal payload As String, ByVal catalog As cInventoryTableInput, _
        ByVal entities As cInventoryTableInput, ByRef selectedSku As String, _
        Optional ByVal existingResult As Boolean = False) As Boolean
    Dim item As Object, report As String, field As Variant, name As String, value As String
    On Error GoTo Invalid
    payload = Trim$(payload)
    If Left$(payload, 1) <> "[" Or Right$(payload, 1) <> "]" Then Exit Function
    Set item = modRoleEventJson.ParseJsonObjectRole(Mid$(payload, 2, Len(payload) - 2), report)
    If item Is Nothing Then Exit Function
    mKey = Trim$(CStr(item("System_Key"))): mSku = Trim$(CStr(item("SKU")))
    mQty = CDbl(item("Qty")): mLocation = Trim$(CStr(item("Location")))
    mCondition = UCase$(Trim$(CStr(item("Condition"))))
    mAttributes = Trim$(CStr(item("AttributesJson")))
    If mKey = "" Or mSku = "" Or mQty < 0 Then Exit Function
    If mCondition = "" Then mCondition = "GOOD"
    If Not entities.PlanNewRow(mKey) Then
        If Not existingResult Then Exit Function
        If CStr(entities.Scalar(mKey, "SKU")) <> mSku Then Exit Function
    End If
    entities.PlanValue mKey, "System_Key", mKey
    entities.PlanValue mKey, "SKU", mSku
    entities.PlanValue mKey, "QtyOnHand", mQty
    entities.PlanValue mKey, "Location", mLocation
    entities.PlanValue mKey, "Condition", mCondition
    entities.PlanValue mKey, "AttributesJson", mAttributes
    entities.PlanValue mKey, "InventoryState", "ACTIVE"
    If IsEmpty(catalog.Scalar(mSku, "SKU")) Then
        If Not catalog.PlanNewRow(mSku) Then Exit Function
        catalog.PlanValue mSku, "SKU", mSku
        catalog.PlanValue mSku, "ITEM_CODE", mSku
        catalog.PlanValue mSku, "ITEM", mSku
    End If
    For Each field In item.Keys
        name = NormalizeHeader(CStr(field)): value = Trim$(CStr(item(field)))
        If name <> "" And value <> "" Then
            If IsCatalogField(name) Then catalog.PlanValue mSku, name, value
        End If
    Next field
    selectedSku = mSku: Plan = True
Invalid:
End Function

Public Function VerifySaved(ByVal book As Workbook, ByVal eventId As String, ByVal payload As String) As Boolean
    Dim catalog As New cInventoryTableInput, entities As New cInventoryTableInput, selected As String
    On Error GoTo Invalid
    If Not catalog.Capture(modExecutionTarget.Table(book, "tblSkuCatalog"), "SKU", False) Then Exit Function
    If Not entities.Capture(modExecutionTarget.Table(book, "tblInventoryEntities"), "System_Key", False) Then Exit Function
    If Not Plan(payload, catalog, entities, selected, True) Then Exit Function
    If Not PlanSaved(book, eventId, entities) Then Exit Function
    If Not catalog.MatchesPersistedCopy(modExecutionTarget.Table(book, "tblSkuCatalog")) Then Exit Function
    VerifySaved = entities.MatchesPersistedCopy(modExecutionTarget.Table(book, "tblInventoryEntities"))
Invalid:
End Function

Private Function NormalizeHeader(ByVal name As String) As String
    name = Replace(Replace(Replace(name, vbCr, " "), vbLf, " "), vbTab, " ")
    name = Replace(Replace(name, "[", "("), "]", ")")
    Do While InStr(name, "  ") > 0: name = Replace(name, "  ", " "): Loop
    NormalizeHeader = UCase$(Left$(Trim$(name), 48))
End Function

Private Function IsCatalogField(ByVal name As String) As Boolean
    Select Case name
        Case "SYSTEM_KEY", "SYSTEMKEY", "CONDITION", "ATTRIBUTESJSON", "SKU", "ROW", _
             "QTY", "QTYDELTA", "QTYAVAILABLE", "TOTAL INV", "NOTE", "IOTYPE", _
             "VERSION", "BOMVERSIONLABEL", "CORRECTEDSHIPEVENTID", "REPAIRNARRATIVE", _
             "MISMATCHFLAG", "LOCATIONFROM", "LOCATIONTO", "PROCESSKEY"
            Exit Function
    End Select
    IsCatalogField = True
End Function

Public Function PlanSaved(ByVal book As Workbook, ByVal eventId As String, _
        ByVal entities As cInventoryTableInput) As Boolean
    Dim table As ListObject, fields As Object, row As Long, receipts As Long, lines As Long
    Dim appliedAt As Variant, delta As Variant
    On Error GoTo Invalid
    Set table = modExecutionTarget.Table(book, "tblAppliedEvents")
    Set fields = modAggregationSources.RequiredColumns(table, Array("EventID", "Status"))
    If fields Is Nothing Then Exit Function
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = eventId Then
            receipts = receipts + 1
            If CStr(table.DataBodyRange.Cells(row, fields("Status")).Value2) <> "APPLIED" Then Exit Function
        End If
    Next row
    If receipts <> 1 Then Exit Function
    Set table = modExecutionTarget.Table(book, "tblInventoryLog")
    Set fields = modAggregationSources.RequiredColumns(table, Array("EventID", "EventType", _
        "System_Key", "SKU", "QtyDelta", "Location", "Condition", "AttributesJson", "AppliedAtUTC"))
    If fields Is Nothing Then Exit Function
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = eventId Then
            lines = lines + 1
            If CStr(table.DataBodyRange.Cells(row, fields("EventType")).Value2) <> "INVENTORY_CREATE" Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("System_Key")).Value2) <> mKey Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2) <> mSku Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("Location")).Value2) <> mLocation Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("Condition")).Value2) <> mCondition Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("AttributesJson")).Value2) <> mAttributes Then Exit Function
            delta = table.DataBodyRange.Cells(row, fields("QtyDelta")).Value2
            If Not IsNumeric(delta) Then Exit Function
            If CDbl(delta) <> mQty Then Exit Function
            appliedAt = table.DataBodyRange.Cells(row, fields("AppliedAtUTC")).Value2
            If IsNumeric(appliedAt) Then appliedAt = CDate(appliedAt)
            If Not IsDate(appliedAt) Then Exit Function
        End If
    Next row
    If lines <> 1 Then Exit Function
    entities.PlanValue mKey, "LastAppliedUTC", CDbl(CDate(appliedAt))
    PlanSaved = True
Invalid:
End Function
