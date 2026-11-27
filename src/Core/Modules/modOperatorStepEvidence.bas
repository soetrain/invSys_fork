Attribute VB_Name = "modOperatorStepEvidence"
Option Explicit
Option Private Module

Public Const CONTROL_ID As String = "PRODUCTION_RUN_COMPLETE"

Public Function IncludesCompletion(ByVal steps As Collection) As Boolean
    Dim step As Object
    For Each step In steps
        If step("ControlId") = CONTROL_ID Then IncludesCompletion = True: Exit Function
    Next step
End Function

Public Function NewClaim(ByVal stepId As String) As Object
    Dim row As Object
    Set row = CreateObject("Scripting.Dictionary")
    row.Add "StepId", stepId: row.Add "BindingId", modTrainingWire.NewId()
    row.Add "Kind", "": row.Add "SelectionSha256", "": row.Add "ActivityId", ""
    row.Add "OwnerValidated", False
    row.Add "AfterObservation", 0&
    Set NewClaim = row
End Function

Public Function ValidClaims(ByVal model As Object) As Boolean
    Dim claim As Object, step As Object, field As Variant, ids As Object, steps As Object
    Set ids = CreateObject("Scripting.Dictionary"): Set steps = CreateObject("Scripting.Dictionary")
    If TypeName(model("OperatorClaims")) <> "Collection" Then Exit Function
    If model("OperatorClaims").Count > 256 Then Exit Function
    For Each claim In model("OperatorClaims")
        If Not modEvaluationModel.HasFields(claim, "StepId|BindingId|Kind|SelectionSha256|ActivityId|OwnerValidated|AfterObservation") Then Exit Function
        For Each field In Array("StepId", "BindingId", "Kind", "SelectionSha256", "ActivityId")
            If VarType(claim(field)) <> vbString Then Exit Function
        Next field
        If VarType(claim("OwnerValidated")) <> vbBoolean Then Exit Function
        If Not modEvaluationModel.IsInteger(claim("AfterObservation")) Then Exit Function
        If claim("AfterObservation") < 0 Or claim("AfterObservation") > 512 Then Exit Function
        If Not modTrainingWire.ValidId(CStr(claim("StepId"))) Or Not modTrainingWire.ValidId(CStr(claim("BindingId"))) Then Exit Function
        If ids.Exists(claim("BindingId")) Or steps.Exists(claim("StepId")) Then Exit Function
        ids.Add claim("BindingId"), True: steps.Add claim("StepId"), claim
        If claim("Kind") = "" Then
            If claim("SelectionSha256") <> "" Or claim("ActivityId") <> "" Or claim("AfterObservation") <> 0 Then Exit Function
        Else
            If claim("Kind") <> "ReusableProcess" And claim("Kind") <> "WorksheetOutput" Then Exit Function
            If Not modEvaluationModel.IsHash(claim("SelectionSha256")) Then Exit Function
        End If
        If claim("ActivityId") <> "" Then
            If Not modTrainingWire.ValidId(CStr(claim("ActivityId"))) Or ids.Exists(claim("ActivityId")) Then Exit Function
            ids.Add claim("ActivityId"), True
        ElseIf claim("OwnerValidated") Then
            Exit Function
        End If
    Next claim
    For Each step In model("Steps")
        If step("ControlId") = CONTROL_ID Then
            If Not steps.Exists(step("StepId")) Then Exit Function
            Set claim = steps(step("StepId"))
            If step("ActivityId") <> claim("ActivityId") Then Exit Function
            If step("State") = "Completed" Then Exit Function
            If step("State") = "OperatorCompleted" Then
                If Not claim("OwnerValidated") Or step("ReasonCode") <> "CONFIRMED" Or step("SourceEventRefs").Count = 0 Then Exit Function
            End If
        ElseIf step("State") = "OperatorCompleted" Then
            Exit Function
        End If
    Next step
    ValidClaims = True
End Function

Public Function Retains(ByVal previous As Collection, ByVal current As Collection) As Boolean
    Dim i As Long, field As Variant, old As Object, item As Object
    If current.Count < previous.Count Or current.Count > previous.Count + 1 Then Exit Function
    For i = 1 To previous.Count
        Set old = previous(i): Set item = current(i)
        If i < previous.Count Or current.Count > previous.Count Then
            If modTrainingJson.EncodeObject(old) <> modTrainingJson.EncodeObject(item) Then Exit Function
        End If
        For Each field In Array("StepId", "BindingId", "Kind", "SelectionSha256", "ActivityId")
            If old(field) <> "" And old(field) <> item(field) Then Exit Function
        Next field
        If old("OwnerValidated") And Not item("OwnerValidated") Then Exit Function
        If old("Kind") <> "" And old("AfterObservation") <> item("AfterObservation") Then Exit Function
        If old("Kind") = "" And item("ActivityId") <> "" Then Exit Function
        If old("ActivityId") = "" And item("OwnerValidated") Then Exit Function
    Next i
    If current.Count > previous.Count Then
        Set item = current(current.Count)
        If item("Kind") <> "" Then Exit Function
    End If
    Retains = True
End Function

' Only observations from the recording that Core opened for this run are read.
Public Function ReadPair(ByVal context As String, ByVal sequence As String, ByVal activity As String, _
                         ByRef terminal As Object, ByVal afterObservation As Long) As Boolean
    Dim header As Object, rows As Collection, record As Object, requested As Boolean, i As Long
    Set terminal = Nothing
    If activity = "" Then Exit Function
    If Not modRecordingSession.ExecutionEvidence(context, header, rows) Then Exit Function
    If header("SequenceId") <> sequence Then Exit Function
    For i = afterObservation + 1 To rows.Count
        Set record = rows(i)
        If record("ActivityId") = activity Then
            If record("SequenceId") <> sequence Or record("ControlId") <> CONTROL_ID Or record("OwnerId") <> "PRODUCTION_RUN_COMPLETION" Then Exit Function
            If record("OutcomeCode") = "REQUESTED" Then
                If requested Then Exit Function
                requested = True
            Else
                If Not requested Or Not terminal Is Nothing Then Exit Function
                Set terminal = record
            End If
        End If
    Next i
    ReadPair = requested
End Function

Public Function Applied(ByVal context As String, ByVal warehouse As String, ByVal terminal As Object) As Boolean
    Dim pub As Object, provenance As Object, result As Object, refs As New Collection, ref As Object
    If terminal("OutcomeCode") <> "CONFIRMED" Or terminal("SourceEventRefs").Count = 0 Then Exit Function
    If Not modLoadedEvents.Read(context, pub, provenance) Then Exit Function
    If pub("WarehouseId") <> warehouse Then Exit Function
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "WarehouseId", warehouse: result.Add "TerminalSources", refs
    modEvaluationSources.Retain result, terminal, pub
    If refs.Count <> terminal("SourceEventRefs").Count Then Exit Function
    For Each ref In refs
        If ref("OwnerStatus") <> "Applied" Or ref("SubmissionState") <> "Submitted" Then Exit Function
    Next ref
    Applied = True
End Function

' Expectation StepIds are distinct from guide StepIds. Use the profile's validated order.
Public Function CreditMap(ByVal profile As Object, ByVal run As Object) As Object
    Dim result As Object, expected As Object, step As Object, credit As Object, position As Long
    Set result = CreateObject("Scripting.Dictionary")
    For Each expected In profile("ExpectedConclusion")("Steps")
        Do
            position = position + 1
            Set step = profile("Steps")(position)
        Loop Until expected("ControlId") = step("ControlId")
        If step("ControlId") = CONTROL_ID Then
            result.Add CStr(expected("StepId")), ""
            For Each credit In run("Steps")
                If credit("StepId") = step("StepId") And credit("State") = "OperatorCompleted" Then result(expected("StepId")) = credit("ActivityId")
            Next credit
        End If
    Next expected
    Set CreditMap = result
End Function
