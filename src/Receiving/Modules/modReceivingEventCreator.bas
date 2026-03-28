Attribute VB_Name = "modReceivingEventCreator"
Option Explicit

Public Function QueueReceiveEventsFromWorkbook(ByVal wb As Workbook, Optional ByRef errorMessage As String = "") As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim cols As Object
    Dim arr As Variant
    Dim r As Long
    Dim currentUserId As String
    Dim rowError As String
    Dim eventIdOut As String

    If wb Is Nothing Then
        errorMessage = "Receiving workbook not provided."
        Exit Function
    End If
    If Not CanCurrentUserPerformCapability("RECEIVE_POST", "", "", "", errorMessage) Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets("ReceivedTally")
    If Not ws Is Nothing Then Set lo = ws.ListObjects("AggregateReceived")
    On Error GoTo 0
    If lo Is Nothing Then
        errorMessage = "AggregateReceived table not found."
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        errorMessage = "AggregateReceived has no rows to confirm."
        Exit Function
    End If

    Set cols = AggColMapReceive(lo)
    If cols Is Nothing Then
        errorMessage = "AggregateReceived is missing required columns."
        Exit Function
    End If

    currentUserId = modRoleEventWriter.ResolveCurrentUserId()
    If currentUserId = "" Then
        errorMessage = "Unable to resolve current user identity."
        Exit Function
    End If

    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        eventIdOut = ""
        rowError = ""
        If Not modRoleEventWriter.QueueReceiveEventCurrent( _
            currentUserId, _
            NzStrReceive(arr(r, cols("ITEM_CODE"))), _
            NzDblReceive(arr(r, cols("QUANTITY"))), _
            NzStrReceive(arr(r, cols("LOCATION"))), _
            BuildReceiveEventNote(arr, cols, r), _
            eventIdOut, _
            rowError) Then
            errorMessage = "Inbox queue failed for row " & r & ": " & rowError
            Exit Function
        End If
    Next r

    QueueReceiveEventsFromWorkbook = True
End Function

Private Function AggColMapReceive(ByVal lo As ListObject) As Object
    Dim d As Object
    Dim names As Variant
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    names = Array("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW")
    For i = LBound(names) To UBound(names)
        d(CStr(names(i))) = ColumnIndexReceive(lo, CStr(names(i)))
        If d(CStr(names(i))) = 0 Then Exit Function
    Next i
    Set AggColMapReceive = d
End Function

Private Function ColumnIndexReceive(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(colName), vbTextCompare) = 0 Then
            ColumnIndexReceive = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function BuildReceiveEventNote(ByVal arr As Variant, ByVal cols As Object, ByVal rowIndex As Long) As String
    BuildReceiveEventNote = "REF_NUMBER=" & NzStrReceive(arr(rowIndex, cols("REF_NUMBER")))
    If NzStrReceive(arr(rowIndex, cols("ITEM"))) <> "" Then
        BuildReceiveEventNote = BuildReceiveEventNote & "; ITEM=" & NzStrReceive(arr(rowIndex, cols("ITEM")))
    End If
    If NzStrReceive(arr(rowIndex, cols("VENDORS"))) <> "" Then
        BuildReceiveEventNote = BuildReceiveEventNote & "; VENDORS=" & NzStrReceive(arr(rowIndex, cols("VENDORS")))
    End If
End Function

Private Function NzStrReceive(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    NzStrReceive = CStr(valueIn)
End Function

Private Function NzDblReceive(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or valueIn = "" Then Exit Function
    NzDblReceive = CDbl(valueIn)
End Function
