Attribute VB_Name = "modShippingReportText"
Option Explicit
Option Private Module

Public Sub AppendNote(ByRef target As String, ByVal text As String)
    If Len(text) = 0 Then Exit Sub
    If Len(target) > 0 Then
        target = target & vbCrLf & text
    Else
        target = text
    End If
End Sub

Public Sub AppendShippingPersistenceSummary(ByRef report As String, _
                                             ByVal inboxSaved As Boolean, _
                                             ByVal reservationLedgerSaved As Boolean, _
                                             Optional ByVal processorDurabilitySaves As Long = 0)
    Dim detail As String

    If inboxSaved Then detail = "warehouse inbox saved"
    If reservationLedgerSaved Then
        If detail <> "" Then detail = detail & "; "
        detail = detail & "reservation ledger saved"
    End If
    If processorDurabilitySaves > 0 Then
        If detail <> "" Then detail = detail & "; "
        detail = detail & "processor durability saves=" & CStr(processorDurabilitySaves)
    End If
    If detail <> "" Then AppendNote report, "Persistence summary: " & detail & "."
End Sub

Public Function AppendHistoryTokenShipping(ByVal baseText As String, ByVal tokenText As String) As String
    baseText = Trim$(baseText)
    tokenText = Trim$(tokenText)
    If tokenText = "" Then
        AppendHistoryTokenShipping = baseText
    ElseIf baseText = "" Then
        AppendHistoryTokenShipping = tokenText
    Else
        AppendHistoryTokenShipping = baseText & "; " & tokenText
    End If
End Function

Public Function ShippingRuntimeReportMetric(ByVal runtimeReport As String, ByVal metricName As String) As Long
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

Public Function FormatShippingRuntimeTiming(ByVal totalMs As Long, _
                                             ByVal batchMs As Long, _
                                             ByVal refreshMs As Long) As String
    FormatShippingRuntimeTiming = "TimingMs=Total:" & CStr(totalMs) & _
                                  ";Batch:" & CStr(batchMs) & _
                                  ";Refresh:" & CStr(refreshMs)
End Function

Public Function SortedTextKeysShipping(ByVal dict As Object) As Variant
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

Public Function NormalizeShippingBomSignatureText(ByVal valueIn As String) As String
    NormalizeShippingBomSignatureText = LCase$(Trim$(valueIn))
End Function

Public Function ShippingRuntimeReportShowsProcessed(ByVal processedCount As Long, ByVal batchReport As String) As Boolean
    If processedCount > 0 Then
        ShippingRuntimeReportShowsProcessed = True
        Exit Function
    End If

    If modShippingReportText.ShippingRuntimeReportMetric(batchReport, "Applied") > 0 Then
        ShippingRuntimeReportShowsProcessed = True
        Exit Function
    End If

    If modShippingReportText.ShippingRuntimeReportMetric(batchReport, "SkipDup") > 0 Then
        ShippingRuntimeReportShowsProcessed = True
    End If
End Function

Public Function ElapsedMillisecondsShipping(ByVal startedAt As Single) As Long
    Dim deltaSeconds As Single

    deltaSeconds = Timer - startedAt
    If deltaSeconds < 0 Then deltaSeconds = deltaSeconds + 86400!
    ElapsedMillisecondsShipping = CLng(deltaSeconds * 1000)
End Function
