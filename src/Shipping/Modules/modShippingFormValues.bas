Attribute VB_Name = "modShippingFormValues"
Option Explicit
Option Private Module

Public Function ShippableNasChangeSummary(ByVal beforeMap As Object, ByVal afterMap As Object) As String
    On Error GoTo CleanExit

    Dim key As Variant
    Dim beforeParts As Variant
    Dim afterParts As Variant
    Dim changes As String
    Dim labelText As String

    If afterMap Is Nothing Or afterMap.Count = 0 Then
        ShippableNasChangeSummary = "NAS Inv checked: no visible shippable rows loaded."
        Exit Function
    End If

    For Each key In afterMap.Keys
        afterParts = Split(CStr(afterMap(key)), vbTab)
        If beforeMap Is Nothing Or Not beforeMap.Exists(CStr(key)) Then GoTo NextKey
        beforeParts = Split(CStr(beforeMap(key)), vbTab)
        If UBound(beforeParts) < 2 Or UBound(afterParts) < 2 Then GoTo NextKey
        If Trim$(CStr(beforeParts(2))) <> Trim$(CStr(afterParts(2))) Then
            labelText = Trim$(CStr(afterParts(0)) & " " & CStr(afterParts(1)))
            If labelText = "" Then labelText = CStr(key)
            If changes <> "" Then changes = changes & "; "
            changes = changes & labelText & " " & _
                      IIf(Trim$(CStr(beforeParts(2))) = "", "blank", Trim$(CStr(beforeParts(2)))) & _
                      " -> " & IIf(Trim$(CStr(afterParts(2))) = "", "blank", Trim$(CStr(afterParts(2))))
        End If
NextKey:
    Next key

    If changes = "" Then
        ShippableNasChangeSummary = "NAS Inv checked: no visible changes."
    Else
        ShippableNasChangeSummary = "NAS Inv updated: " & changes & "."
    End If
    Exit Function

CleanExit:
    ShippableNasChangeSummary = "NAS Inv checked."
End Function

Public Function DisplayVersionOrNA(ByVal versionText As String) As String
    versionText = Trim$(versionText)
    If versionText = "" Then
        DisplayVersionOrNA = "NA"
    Else
        DisplayVersionOrNA = versionText
    End If
End Function

Public Function SelectedListTableRowCount(ByVal lst As MSForms.ListBox) As Long
    Dim rows As Variant

    rows = SelectedListTableRows(lst)
    If IsEmpty(rows) Then Exit Function
    SelectedListTableRowCount = UBound(rows) - LBound(rows) + 1
End Function

Public Function SelectedListTableRows(ByVal lst As MSForms.ListBox) As Variant
    Dim rowIndexes() As Long
    Dim i As Long
    Dim countRows As Long
    Dim tableRow As Long

    If lst Is Nothing Then Exit Function
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            tableRow = CLng(Val(NzText(lst.List(i, 10))))
            If tableRow > 0 Then
                countRows = countRows + 1
                ReDim Preserve rowIndexes(1 To countRows)
                rowIndexes(countRows) = tableRow
            End If
        End If
    Next i
    If countRows > 0 Then SelectedListTableRows = rowIndexes
End Function

Public Function AppendTiming(ByVal report As String, ByVal elapsedMs As Long) As String
    If Trim$(report) <> "" Then
        AppendTiming = report & vbCrLf & vbCrLf
    End If
    AppendTiming = AppendTiming & "Completed in " & Format$(elapsedMs, "#,##0") & " ms."
End Function

Public Function NzText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        NzText = ""
    Else
        NzText = CStr(value)
    End If
End Function

Public Function ParseNumber(ByVal textValue As String) As Double
    On Error GoTo UseZero
    textValue = Trim$(textValue)
    If textValue = "" Then Exit Function
    ParseNumber = CDbl(textValue)
    Exit Function
UseZero:
    ParseNumber = 0
End Function

Public Function FormatQuantity(ByVal qtyValue As Double) As String
    If Abs(qtyValue - Fix(qtyValue)) < 0.0000001 Then
        FormatQuantity = Format$(qtyValue, "0")
    Else
        FormatQuantity = Format$(qtyValue, "0.###")
    End If
End Function

Public Function DisplayQtyText(ByVal rawText As String) As String
    Dim qty As Double

    rawText = Trim$(rawText)
    If rawText = "" Then Exit Function
    If LCase$(rawText) = "unknown" Then
        DisplayQtyText = "unknown"
        Exit Function
    End If
    qty = ParseNumber(rawText)
    DisplayQtyText = FormatQuantity(qty)
End Function
