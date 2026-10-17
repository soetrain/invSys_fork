VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cInventoryRetirementPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Transient exact all-active-entity retirement expectations, never a receipt store.
Private mSku As String, mExpected As Object

Public Function Plan(ByVal entities As cInventoryTableInput, ByVal sku As String, _
        ByVal reason As String, ByRef payload As String, ByRef count As Long) As Boolean
    Dim key As Variant, state As String, qty As Variant, location As String, condition As String
    Dim attributes As String, items As New Collection, item As Object
    On Error GoTo Invalid
    If sku = "" Or Trim$(reason) = "" Then Exit Function
    mSku = sku: Set mExpected = CreateObject("Scripting.Dictionary")
    For Each key In entities.KeysForSku(sku)
        state = UCase$(Trim$(CStr(entities.Scalar(CStr(key), "InventoryState"))))
        If state = "ACTIVE" Or state = "" Then
            qty = entities.Scalar(CStr(key), "QtyOnHand")
            If Not IsNumeric(qty) Then Exit Function
            If CDbl(qty) < 0 Then Exit Function
            location = CStr(entities.Scalar(CStr(key), "Location"))
            condition = UCase$(CStr(entities.Scalar(CStr(key), "Condition")))
            If condition = "" Then condition = "GOOD"
            attributes = CStr(entities.Scalar(CStr(key), "AttributesJson"))
            mExpected.Add CStr(key), Array(-CDbl(qty), location, condition, attributes)
            Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(CStr(key), sku, _
                -CDbl(qty), location, condition, attributes, "Admin delete inventory item. Reason: " & reason)
            item("ITEM_CODE") = sku: item("IoType") = "RETIRE"
            item("InventoryState") = "RETIRED": item("Reason") = reason
            items.Add item
            entities.PlanValue CStr(key), "QtyOnHand", 0#
            entities.PlanValue CStr(key), "InventoryState", "RETIRED"
        End If
    Next key
    payload = modRoleEventWriter.BuildPayloadJsonFromCollection(items)
    count = items.Count: Plan = True
Invalid:
End Function

Public Function PlanSaved(ByVal book As Workbook, ByVal eventId As String, _
        ByVal entities As cInventoryTableInput) As Boolean
    Dim table As ListObject, fields As Object, row As Long, receipts As Long
    Dim key As String, expected As Variant, seen As Object, appliedAt As Variant, delta As Variant
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
        "System_Key", "SKU", "QtyDelta", "Location", "Condition", "AttributesJson", "InventoryState", "AppliedAtUTC"))
    If fields Is Nothing Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = eventId Then
            key = CStr(table.DataBodyRange.Cells(row, fields("System_Key")).Value2)
            If Not mExpected.Exists(key) Or seen.Exists(key) Then Exit Function
            seen.Add key, True: expected = mExpected(key)
            If CStr(table.DataBodyRange.Cells(row, fields("EventType")).Value2) <> "ADMIN_INVENTORY_ADJUST" Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2) <> mSku Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("InventoryState")).Value2) <> "RETIRED" Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("Location")).Value2) <> expected(1) Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("Condition")).Value2) <> expected(2) Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("AttributesJson")).Value2) <> expected(3) Then Exit Function
            delta = table.DataBodyRange.Cells(row, fields("QtyDelta")).Value2
            If Not IsNumeric(delta) Then Exit Function
            If CDbl(delta) <> expected(0) Then Exit Function
            appliedAt = table.DataBodyRange.Cells(row, fields("AppliedAtUTC")).Value2
            If IsNumeric(appliedAt) Then appliedAt = CDate(appliedAt)
            If Not IsDate(appliedAt) Then Exit Function
            entities.PlanValue key, "LastAppliedUTC", CDbl(CDate(appliedAt))
        End If
    Next row
    PlanSaved = (seen.Count = mExpected.Count And seen.Count > 0)
Invalid:
End Function
