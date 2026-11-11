Attribute VB_Name = "modInventoryWorksheetActivity"
Option Explicit

' Observation wire: newline-separated exact event ID, TAB, Submitted/Unknown.
' This carries owner facts only. It never submits or confirms business work.
Public Function References(ByVal action As String, ByVal events As String) As String
    Dim seen As Object, line As Variant, fields As Variant, id As Variant, json As String, part As String
    On Error GoTo Invalid
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare
    If events <> "" Then
        For Each line In Split(events, vbLf)
            fields = Split(CStr(line), vbTab)
            If UBound(fields) <> 1 Then Exit Function
            If fields(1) <> "Submitted" And fields(1) <> "Unknown" Then Exit Function
            If Not seen.Exists(fields(0)) Then
                seen.Add fields(0), fields(1)
            ElseIf fields(1) = "Submitted" Then
                seen(fields(0)) = "Submitted"
            End If
        Next line
    End If
    For Each id In seen.Keys
        part = modActivity.InventorySourceReferences(action, CStr(id), CStr(seen(id)))
        If Len(part) < 3 Then Exit Function
        If json <> "" Then json = json & ","
        json = json & Mid$(part, 2, Len(part) - 2)
    Next id
    References = "[" & json & "]"
Invalid:
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim message As String, severity As String, effect As String, record As Object, create As Boolean
    create = (id = "ADMIN_INVENTORY_CREATE_TABLE")
    If Not create And id <> "ADMIN_INVENTORY_UPLOAD_TABLE" Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = "Inventory worksheet action requested."
        Case "REFUSED"
            severity = "Blocked": effect = "Unchanged"
            message = "The worksheet action was refused before writes. Reopen Add/Edit Inventory."
        Case "SAVED"
            If Not create Then Exit Function
            effect = "Unchanged": message = "The local Inventory table was saved and presented. No inventory application is asserted."
        Case "EMPTY", "REJECTED"
            If create Then Exit Function
            effect = "Unchanged"
            If code = "EMPTY" Then
                message = "The selected Inventory table has no pending rows to upload."
            Else
                severity = "Blocked": message = "The selected Inventory table failed whole-selection preflight. No inventory application was attempted."
            End If
        Case "CONFIRMED"
            If create Then Exit Function
            message = "The owning Inventory commands and saved worksheet results were confirmed. Previously applied work may have been reconciled."
        Case "PARTIAL"
            severity = "Warning": message = "Worksheet completion is unconfirmed; contributing changes may remain. Inspect the owning result before retrying."
        Case "FAILED"
            severity = "Error": message = "The Inventory worksheet action did not finish. Inspect the owning result; no rollback is asserted."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
