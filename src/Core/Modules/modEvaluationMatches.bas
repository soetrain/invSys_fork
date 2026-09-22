Attribute VB_Name = "modEvaluationMatches"
Option Explicit
Option Private Module

Public Function PolicyControls(ByVal policy As Object, ByVal capture As Boolean) As Object
    Dim result As Object, row As Object, definition As Object, allowed As Boolean
    Set result = CreateObject("Scripting.Dictionary")
    If Not policy Is Nothing Then
        For Each row In policy("Controls")
            Set definition = modActivityCatalog.Control(CStr(row("ControlId")), CLng(policy("SavedCatalogVersion")))
            allowed = False
            If Not definition Is Nothing Then
                If capture Then
                    allowed = CBool(policy("ViewerActionPathCaptureEnabled")) And CBool(row("SequenceEligible")) And _
                        (CBool(row("Collect")) Or definition("Class") = "Navigation")
                Else
                    allowed = CBool(row("Visible")) And _
                        (definition("Role") <> "Admin" Or CBool(policy("AdminViewerEventLoggingEnabled")))
                End If
            End If
            result.Add CStr(row("ControlId")), allowed
        Next row
    End If
    Set PolicyControls = result
End Function

Public Function Permitted(ByVal controls As Object, ByVal controlId As String) As Boolean
    If controls.Exists(controlId) Then Permitted = CBool(controls(controlId))
End Function

' One action occurrence can match at most one step, regardless of its record count.
Public Function Match(ByVal result As Object, ByVal observations As Collection, ByVal historical As Object, _
                      ByVal visible As Object) As Object
    Dim actions As New Collection, outcomes As Object, selected As Object, used As Object
    Dim record As Object, step As Object, attempt As Object, found As Object, matchRow As Object
    Dim cursor As Long, index As Long, first As Long, id As String, terminal As Object
    Set outcomes = CreateObject("Scripting.Dictionary"): Set used = CreateObject("Scripting.Dictionary")
    For Each record In observations
        If record("OutcomeCode") = "REQUESTED" Then
            actions.Add record
        Else
            outcomes.Add CStr(record("ActivityId")), record
        End If
    Next record
    For Each step In result("ExpectedConclusion")("Steps")
        Set found = Nothing: first = 0
        id = CStr(step("ControlId"))
        If Not Permitted(historical, id) Or Not Permitted(visible, id) Then
            result("UnavailableSteps").Add CStr(step("StepId"))
            If Not Permitted(historical, id) Then modEvaluationModel.Reason result, "REQUIRED_CAPTURE_UNAVAILABLE"
            If Not Permitted(visible, id) Then modEvaluationModel.Reason result, "REQUIRED_OBSERVATION_RESTRICTED"
        Else
            For index = cursor + 1 To actions.Count
                Set attempt = actions(index)
                If attempt("ControlId") = id Then
                    If first = 0 Then first = index
                    Set selected = attempt
                    If step("RequiredOutcome") <> "REQUESTED" And outcomes.Exists(attempt("ActivityId")) Then
                        Set selected = outcomes(attempt("ActivityId"))
                    End If
                    If selected("OutcomeCode") = step("RequiredOutcome") Then Set found = selected: Exit For
                    If Not CBool(step("RetryAllowed")) Then Exit For
                End If
            Next index
            If Not found Is Nothing Then
                cursor = CLng(found("Ordinal"))
                Set matchRow = CreateObject("Scripting.Dictionary")
                matchRow.Add "StepId", step("StepId"): matchRow.Add "ActivityId", found("ActivityId")
                matchRow.Add "Ordinal", found("Ordinal"): matchRow.Add "ControlId", found("ControlId")
                matchRow.Add "OutcomeCode", found("OutcomeCode")
                result("Matches").Add matchRow: used.Add CStr(found("ActivityId")), True
                If step("StepId") = result("ExpectedConclusion")("TerminalStepId") Then Set terminal = found
            ElseIf first = 0 Then
                result("MissingSteps").Add CStr(step("StepId"))
                modEvaluationModel.Reason result, "REQUIRED_STEP_MISSING"
            Else
                cursor = first
                result("FailedSteps").Add CStr(step("StepId"))
                modEvaluationModel.Reason result, "REQUIRED_OUTCOME_MISMATCH"
            End If
        End If
    Next step
    For Each attempt In actions
        If Not used.Exists(attempt("ActivityId")) And Permitted(visible, CStr(attempt("ControlId"))) Then
            result("ExtraActivityIds").Add CStr(attempt("ActivityId"))
        End If
    Next attempt
    Set Match = terminal
End Function

Public Function CommandCompleted(ByVal record As Object) As Boolean
    Dim definition As Object, outcome As Object, code As String
    Set definition = modActivityCatalog.Control(CStr(record("ControlId")), CLng(record("CatalogVersion")))
    If definition Is Nothing Then Exit Function
    Set outcome = modActivityCatalog.Outcome(CStr(record("ControlId")), CStr(record("OutcomeCode")))
    If outcome Is Nothing Then Exit Function
    code = CStr(record("OutcomeCode"))
    ' Explicit owner facts; severity and data effect are deliberately not classifiers.
    Select Case CStr(record("ControlId"))
        Case "ADMIN_SETTINGS_SAVE_VALUE", "PRODUCTION_UOM_RETRIEVE", _
             "ADMIN_UOM_ADD", "ADMIN_UOM_REMOVE", "ADMIN_UOM_RESET"
            CommandCompleted = (code = "COMPLETED" Or code = "UNCHANGED")
        Case "RECEIVING_CONFIRM_WRITES", "RECEIVING_WORKSHEET_CONFIRM", "DISPOSITION_CONFIRM", "SHIPPING_SEND", _
             "BOXING_MAKE", "BOXING_UNBOX"
            CommandCompleted = (code = "CONFIRMED")
        Case "RECEIVING_ADD_SELECTED", "DISPOSITION_ADD_SELECTED", "SHIPPING_ADD", "SHIPPING_UPDATE", "SHIPPING_REMOVE", "SHIPPING_HOLD", "SHIPPING_RETURN", "SHIPPING_STAGE": CommandCompleted = (code = "STAGED")
        Case "RECEIVING_REFRESH": CommandCompleted = (code = "REFRESHED")
        Case "RECEIVING_CLEAR": CommandCompleted = (code = "CLEARED" Or code = "EMPTY")
        Case "RECEIVING_OPEN": CommandCompleted = (code = "OPENED" Or code = "REUSED")
        Case "RECEIVING_CLOSE": CommandCompleted = (code = "CLOSED")
        Case Else
            If definition("Class") = "Navigation" And definition("OwnerId") = "RECEIVING_NAVIGATION" Then CommandCompleted = (code = "SELECTED")
    End Select
End Function
