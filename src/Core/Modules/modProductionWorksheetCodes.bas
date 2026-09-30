Attribute VB_Name = "modProductionWorksheetCodes"
Option Explicit
Option Private Module

' D18 catalog 22. Local saves and owning import completion are distinct facts.
Public Function ControlIds() As Variant
    ControlIds = Array("PRODUCTION_PROCESS_WORKSHEET_SEND", "PRODUCTION_PROCESS_WORKSHEET_ADD_ITEM", _
                       "PRODUCTION_PROCESS_WORKSHEET_RETRIEVE")
End Function

Public Function Control(ByVal id As String) As Object
    Dim ids As Variant, captions As Variant, index As Long
    ids = ControlIds()
    captions = Array("Send Process to Sheet", "Add Acceptable Item", "Retrieve Selected Process")
    For index = LBound(ids) To UBound(ids)
        If id = ids(index) Then
            Set Control = modProductionControlCatalog.Command(id, "PRODUCTION_PROCESS_WORKSHEET", CStr(captions(index)), _
                                                             "Operations > Production > Process Designer")
            Exit Function
        End If
    Next index
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim record As Object, definition As Object, severity As String, effect As String
    Dim message As String, nextStep As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED"
            effect = "Unknown": message = "Process worksheet action requested."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; no worksheet edit or Designs submission was made."
        Case "REJECTED"
            severity = "Warning": message = "Validation or selection prevented submission; saved Designs definitions were not changed."
            nextStep = "Review the owning form's validation results; local worksheet normalization may remain."
        Case "STAGED"
            If id = "PRODUCTION_PROCESS_WORKSHEET_RETRIEVE" Then Exit Function
            message = "The Process worksheet edit was saved locally; saved Designs definitions were not changed."
        Case "CONFIRMED"
            If id <> "PRODUCTION_PROCESS_WORKSHEET_RETRIEVE" Then Exit Function
            effect = "Unknown"
            message = "Every selected Process import and worksheet removal/save finished; application requires exact Designs evidence."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The Process worksheet action failed or partly completed; its final data state requires verification."
            nextStep = "Inspect the captured worksheet and any exact referenced Designs events before retrying; prior work may remain."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
