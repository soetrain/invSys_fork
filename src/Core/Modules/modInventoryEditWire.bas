Attribute VB_Name = "modInventoryEditWire"
Option Explicit

' EDIT references are not credentials. Exact saved receipt and event proof are
' still required; a fingerprint alone never establishes application.
Public Function HashFields(ByVal names As Variant, ByVal values As Variant) As String
    Dim fields As Object, keys As Variant, key As Variant, i As Long, j As Long
    Dim text As String, swap As String
    On Error GoTo Invalid
    If Not IsArray(names) Or Not IsArray(values) Then Exit Function
    If LBound(names) <> LBound(values) Or UBound(names) <> UBound(values) Then Exit Function
    Set fields = CreateObject("Scripting.Dictionary"): fields.CompareMode = vbBinaryCompare
    For i = LBound(names) To UBound(names)
        If VarType(names(i)) <> vbString Or IsObject(values(i)) Or IsArray(values(i)) Then Exit Function
        If IsError(values(i)) Or IsNull(values(i)) Then Exit Function
        fields.Add CStr(names(i)), CStr(values(i))
    Next i
    keys = fields.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(CStr(keys(i)), CStr(keys(j)), vbBinaryCompare) > 0 Then
                swap = keys(i): keys(i) = keys(j): keys(j) = swap
            End If
        Next j
    Next i
    text = "{"
    For Each key In keys
        If Len(text) > 1 Then text = text & ","
        text = text & modTrainingJson.Quote(CStr(key)) & ":" & modTrainingJson.Quote(CStr(fields(key)))
    Next key
    HashFields = modTrainingWire.Sha256(text & "}")
Invalid:
End Function

Public Function ReceiptHeaders() As Variant
    ReceiptHeaders = Split("CommandId|WarehouseId|StationId|UserId|SKU|InputHash|OriginalStateHash|PlannedEventId|System_Key|PreviousQty|TargetQty|PayloadJson|QuantityIntent|CatalogPhase|BindingHash", "|")
End Function

Public Function ReferenceField(ByVal text As String, ByVal prefix As String) As String
    Dim line As Variant, found As String, value As String, count As Long
    For Each line In Split(Replace$(text, vbCr, ""), vbLf)
        line = Trim$(CStr(line))
        If StrComp(Left$(line, Len(prefix)), prefix, vbTextCompare) = 0 Then
            count = count + 1
            value = Trim$(Mid$(line, Len(prefix) + 1))
            If value = "" Or InStr(value, " ") > 0 Or Len(value) > 200 Then Exit Function
            found = value
        End If
    Next line
    If count = 1 Then ReferenceField = found
End Function

Public Function Reference(ByVal command As String, ByVal binding As String, ByVal eventId As String) As String
    Reference = "Command: " & command & vbCrLf & "Edit binding: " & binding
    If eventId <> "" Then Reference = Reference & vbCrLf & "Event: " & eventId
End Function

Public Function WorksheetHeader(ByVal name As String) As String
    Dim separator As Variant
    name = UCase$(Trim$(name))
    For Each separator In Array(" ", "_", "-", "/", "(", ")")
        name = Replace$(name, CStr(separator), "")
    Next separator
    WorksheetHeader = name
End Function
