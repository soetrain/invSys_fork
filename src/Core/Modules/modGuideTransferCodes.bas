Attribute VB_Name = "modGuideTransferCodes"
Option Explicit
Option Private Module

Public Function Control(ByVal id As String) As Object
    Dim record As Object, caption As String
    Select Case id
        Case "VIEWER_GUIDE_EXPORT": caption = "Export"
        Case "VIEWER_GUIDE_IMPORT": caption = "Import"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "CORE_GUIDE_TRANSFER"
    record.Add "Class", "Command": record.Add "Role", "Viewer"
    record.Add "Caption", caption: record.Add "Surface", "Operations > Viewer > Published guides"
    record.Add "Capability", "ACTION_PATH_MAINT": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim record As Object, severity As String, effect As String, message As String
    If id <> "VIEWER_GUIDE_EXPORT" And id <> "VIEWER_GUIDE_IMPORT" Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED": effect = "Unknown": message = "Guide transfer requested."
        Case "COMPLETED"
            effect = "Changed"
            If id = "VIEWER_GUIDE_EXPORT" Then
                message = "A new guide package was exported."
            Else
                message = "A new local guide was imported with its original evidence retained."
            End If
        Case "CANCELLED": severity = "Notice": message = "Guide file selection was cancelled."
        Case "DENIED": severity = "Blocked": message = "Guide transfer requires ACTION_PATH_MAINT."
        Case "REJECTED": severity = "Warning": message = "Validation prevented guide transfer before writing."
        Case "FAILED": severity = "Error": effect = "Unknown": message = "Guide transfer failed; verify the destination before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
