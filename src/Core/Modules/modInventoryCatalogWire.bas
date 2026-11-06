Attribute VB_Name = "modInventoryCatalogWire"
Option Explicit

' Pure planning only: no workbook access, authorization or canonical writes.
' Inputs and result use declared scalar arrays. The result contains two parallel
' arrays, field names and values; no Dictionary crosses the XLAM boundary.
Public Function PlanFields(ByVal values As Variant, ByVal customNames As Variant, _
        ByVal customValues As Variant, ByVal priorHistory As String, _
        ByVal existingHeaders As Variant) As Variant
    Dim fields As Object, names As Variant, refs As Variant, i As Long, count As Long
    Dim text As String, existing As String, nameBase As Long, valueBase As Long
    If ScalarCount(values) <> 10 Then GoTo Invalid
    If LBound(values) <> 0 Then GoTo Invalid
    count = ScalarCount(customNames)
    If count < 0 Or count <> ScalarCount(customValues) Then GoTo Invalid
    If ScalarCount(existingHeaders) < 0 Then GoTo Invalid
    Set fields = CreateObject("Scripting.Dictionary"): fields.CompareMode = vbTextCompare
    names = Array("ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION", _
                  "VENDOR(s)", "VENDOR_CODE", "CATEGORY", "EXTERNAL_CODE")
    For i = 0 To UBound(names)
        PlanValue fields, CStr(names(i)), values(i)
    Next i
    refs = PictureValues(CStr(values(9)))
    If ScalarCount(refs) > 0 Then fields("IMAGE_PATH") = refs(0) Else fields("IMAGE_PATH") = ""
    For i = 2 To 12
        text = "IMAGE_PATH_" & CStr(i)
        If i <= ScalarCount(refs) Then
            fields(text) = refs(i - 1)
        ElseIf HasHeader(existingHeaders, text) Then
            fields(text) = ""
        End If
    Next i
    If count > 0 Then
        nameBase = LBound(customNames): valueBase = LBound(customValues)
        For i = 0 To count - 1
            text = Trim$(CStr(customNames(nameBase + i)))
            If StrComp(text, "EDIT_HISTORY_APPEND", vbTextCompare) = 0 Then
                text = Trim$(CStr(customValues(valueBase + i)))
                If text <> "" Then
                    If fields.Exists("EDIT_HISTORY") Then existing = CStr(fields("EDIT_HISTORY")) Else existing = priorHistory
                    If existing <> "" Then text = existing & vbLf & text
                    fields("EDIT_HISTORY") = text
                End If
            Else
                PlanValue fields, text, customValues(valueBase + i)
            End If
        Next i
    End If
    PlanFields = Array(fields.Keys, fields.Items)
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2221, "modInventoryCatalogWire.PlanFields", _
              "Catalog fields must use scalar arrays with matching custom fields."
End Function

Public Function NormalizeHeader(ByVal name As String) As String
    name = Replace(name, vbCr, " ")
    name = Replace(name, vbLf, " ")
    name = Replace(name, vbTab, " ")
    name = Replace(name, "[", "(")
    name = Replace(name, "]", ")")
    Do While InStr(1, name, "  ", vbBinaryCompare) > 0
        name = Replace(name, "  ", " ")
    Loop
    name = Trim$(name)
    If Len(name) > 48 Then name = Left$(name, 48)
    NormalizeHeader = name
End Function

Public Function PictureValues(ByVal text As String) As Variant
    Dim parts As Variant, part As Variant, values As New Collection
    Dim result() As String, value As String, i As Long
    text = Replace(text, "|", ";")
    text = Replace(text, vbCr, ";")
    text = Replace(text, vbLf, ";")
    parts = Split(text, ";")
    For Each part In parts
        value = Trim$(CStr(part))
        If value <> "" Then values.Add value
    Next part
    If values.Count = 0 Then PictureValues = Array(): Exit Function
    ReDim result(0 To values.Count - 1)
    For i = 1 To values.Count: result(i - 1) = CStr(values(i)): Next i
    PictureValues = result
End Function

Private Sub PlanValue(ByVal fields As Object, ByVal name As String, ByVal value As Variant)
    name = NormalizeHeader(name)
    If name <> "" Then fields(name) = value
End Sub

Private Function HasHeader(ByVal headers As Variant, ByVal name As String) As Boolean
    Dim header As Variant
    For Each header In headers
        If StrComp(CStr(header), name, vbTextCompare) = 0 Then HasHeader = True: Exit Function
    Next header
End Function

Private Function ScalarCount(ByVal values As Variant) As Long
    Dim i As Long, count As Long
    ScalarCount = -1
    On Error GoTo Invalid
    If Not IsArray(values) Then Exit Function
    count = UBound(values) - LBound(values) + 1
    For i = LBound(values) To UBound(values)
        If IsObject(values(i)) Or IsArray(values(i)) Then Exit Function
    Next i
    ScalarCount = count
Invalid:
End Function
