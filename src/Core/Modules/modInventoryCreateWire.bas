Attribute VB_Name = "modInventoryCreateWire"
Option Explicit

' Cross-XLAM creation contracts contain scalars and parallel one-dimensional
' arrays of scalar fields. The existing event JSON format remains unchanged.
Public Function BuildPayload(ByVal systemKey As String, ByVal sku As String, _
        ByVal qty As Double, ByVal location As String, ByVal condition As String, _
        ByVal attributesJson As String, ByVal note As String, _
        ByVal fieldNames As Variant, ByVal fieldValues As Variant) As String
    Dim item As Object, items As New Collection, i As Long
    If Not IsArray(fieldNames) Or Not IsArray(fieldValues) Then GoTo Invalid
    If LBound(fieldNames) <> LBound(fieldValues) Or UBound(fieldNames) <> UBound(fieldValues) Then GoTo Invalid
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        systemKey, sku, qty, location, condition, attributesJson, note)
    For i = LBound(fieldNames) To UBound(fieldNames)
        If VarType(fieldNames(i)) <> vbString Then GoTo Invalid
        If IsObject(fieldValues(i)) Or IsArray(fieldValues(i)) Then GoTo Invalid
        item(CStr(fieldNames(i))) = fieldValues(i)
    Next i
    items.Add item
    BuildPayload = modRoleEventWriter.BuildPayloadJsonFromCollection(items)
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2191, "modInventoryCreateWire.BuildPayload", _
        "Inventory payload fields must be matching scalar arrays."
End Function

' Admin never supplies a Workbook object. Existing provisioned authorization,
' exact event identity and queue behavior remain at the owning Core boundary.
Public Function QueueCreate(ByVal warehouse As String, ByVal station As String, _
        ByVal user As String, ByVal payload As String, ByVal note As String, _
        ByRef eventId As String, ByRef report As String, Optional ByRef observedEvent As String = "") As Boolean
    Dim attempted As Boolean
    observedEvent = ""
    QueueCreate = modRoleEventWriter.QueueInventoryCreateEvent(warehouse, station, _
        user, payload, note, 0, Nothing, eventId, report, "", attempted)
    If QueueCreate Then
        observedEvent = eventId & vbTab & "Submitted"
    ElseIf attempted Then
        observedEvent = eventId & vbTab & "Unknown"
    End If
End Function
