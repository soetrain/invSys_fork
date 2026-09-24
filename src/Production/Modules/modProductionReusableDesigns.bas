Attribute VB_Name = "modProductionReusableDesigns"
Option Explicit

Public Function SubmitReusableDesignEvent(ByVal eventType As String, _
                                          ByVal definitionId As String, _
                                          ByVal definitionVersion As String, _
                                          Optional ByVal payloadJson As String = "", _
                                          Optional ByVal noteText As String = "", _
                                          Optional ByRef report As String = "", _
                                          Optional ByVal facts As cProductionLifecycleFacts = Nothing) As Boolean
    On Error GoTo Failed

    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim appliedCount As Long
    Dim warehouseId As String
    Dim expectedStatus As String
    Dim actualStatus As String
    Dim submitted As Boolean, writeAttempted As Boolean

    eventType = UCase$(Trim$(eventType))
    definitionId = Trim$(definitionId)
    definitionVersion = Trim$(definitionVersion)
    If definitionId = "" Or definitionVersion = "" Then
        report = "Definition ID and version are required."
        If Not facts Is Nothing Then facts.OutcomeCode = "REJECTED"
        Exit Function
    End If
    If Not IsReusableLifecycleEvent(eventType) Then
        report = "Unsupported reusable design event: " & eventType
        Exit Function
    End If
    submitted = modRoleEventWriter.QueueDesignEventCurrent( _
            eventType, definitionId, definitionVersion, payloadJson, noteText, _
            "", eventId, queueError, writeAttemptedOut:=writeAttempted)
    If Not facts Is Nothing Then facts.ObserveSubmission eventId, submitted, writeAttempted
    If Not submitted Then
        report = "Event was not queued: " & queueError
        Exit Function
    End If

    warehouseId = Trim$(modConfig.GetWarehouseId())
    appliedCount = modProcessor.RunBatch(warehouseId, 0, processorReport)
    expectedStatus = ExpectedReusableStatus(eventType)
    actualStatus = ReusableDefinitionStatus(eventType, definitionId, definitionVersion)
    If StrComp(actualStatus, expectedStatus, vbTextCompare) <> 0 Then
        report = "Event " & eventId & " was queued but the expected " & expectedStatus & _
                 " projection is not visible. Processor applied=" & CStr(appliedCount) & _
                 ". " & processorReport & "; ProjectionIdentities=" & _
                 ReusableDefinitionIdentityEvidence(eventType)
        Exit Function
    End If

    report = ReusableDefinitionKind(eventType) & " " & definitionId & _
             " version " & definitionVersion & " is " & expectedStatus & _
             ". EventID=" & eventId & "; Processor applied=" & CStr(appliedCount) & "."
    SubmitReusableDesignEvent = True
    Exit Function

Failed:
    If Not facts Is Nothing Then facts.ObserveSubmission eventId, submitted, writeAttempted
    report = "Reusable design action failed: " & Err.Description
End Function

Private Function ReusableDefinitionIdentityEvidence(ByVal eventType As String) As String
    Dim rows As Variant
    Dim r As Long
    Dim shown As Long

    If Left$(eventType, 7) = "PROCESS" Then
        rows = modOperationsPrimitiveBridge.ListProcesses("")
    Else
        rows = modOperationsPrimitiveBridge.ListRecipes("")
    End If
    If Not IsArray(rows) Then
        ReusableDefinitionIdentityEvidence = "NoRows"
        Exit Function
    End If
    On Error GoTo Done
    For r = LBound(rows, 1) To UBound(rows, 1)
        If ReusableDefinitionIdentityEvidence <> "" Then _
            ReusableDefinitionIdentityEvidence = ReusableDefinitionIdentityEvidence & ","
        ReusableDefinitionIdentityEvidence = ReusableDefinitionIdentityEvidence & _
            Trim$(CStr(rows(r, 1))) & "/" & Trim$(CStr(rows(r, 2))) & "/" & _
            Trim$(CStr(rows(r, 5)))
        shown = shown + 1
        If shown >= 12 Then Exit For
    Next r
Done:
    If ReusableDefinitionIdentityEvidence = "" Then _
        ReusableDefinitionIdentityEvidence = "Empty"
End Function

Public Function NextReusableDefinitionVersion(ByVal definitionId As String, _
                                               ByVal processDefinition As Boolean) As String
    Dim rows As Variant
    Dim r As Long
    Dim best As Long
    Dim candidate As Long

    definitionId = Trim$(definitionId)
    If processDefinition Then
        rows = modOperationsPrimitiveBridge.ListProcesses("")
    Else
        rows = modOperationsPrimitiveBridge.ListRecipes("")
    End If
    If IsArray(rows) Then
        On Error GoTo NoRows
        For r = LBound(rows, 1) To UBound(rows, 1)
            If StrComp(Trim$(CStr(rows(r, 1))), definitionId, vbTextCompare) = 0 Then
                If IsNumeric(rows(r, 2)) Then
                    candidate = CLng(rows(r, 2))
                    If candidate > best Then best = candidate
                End If
            End If
        Next r
    End If
NoRows:
    NextReusableDefinitionVersion = CStr(best + 1)
End Function

Public Function ParseReusableDefinitionRecords(ByVal jsonText As String, _
                                                Optional ByRef report As String = "") As Collection
    On Error GoTo Failed

    Dim records As New Collection
    Dim position As Long
    Dim record As Object

    position = 1
    SkipJsonWhitespace jsonText, position
    If JsonCharacter(jsonText, position) <> "[" Then
        report = "Definition JSON must be an array."
        Exit Function
    End If
    position = position + 1
    Do
        SkipJsonWhitespace jsonText, position
        If JsonCharacter(jsonText, position) = "]" Then
            position = position + 1
            Exit Do
        End If
        Set record = ParseFlatJsonObject(jsonText, position, report)
        If record Is Nothing Then Exit Function
        records.Add record
        SkipJsonWhitespace jsonText, position
        Select Case JsonCharacter(jsonText, position)
            Case ",": position = position + 1
            Case "]": position = position + 1: Exit Do
            Case Else
                report = "Expected ',' or ']' in definition JSON."
                Exit Function
        End Select
    Loop
    Set ParseReusableDefinitionRecords = records
    Exit Function

Failed:
    report = "Definition JSON parse failed: " & Err.Description
End Function

Public Function ReusableRecordText(ByVal record As Object, ByVal keyName As String) As String
    If record Is Nothing Then Exit Function
    If record.Exists(keyName) Then ReusableRecordText = Trim$(CStr(record(keyName)))
End Function

Public Function ReusableRecordValue(ByVal record As Object, ByVal keyName As String) As Variant
    If record Is Nothing Then Exit Function
    If record.Exists(keyName) Then ReusableRecordValue = record(keyName)
End Function

Private Function ParseFlatJsonObject(ByVal jsonText As String, ByRef position As Long, _
                                     ByRef report As String) As Object
    Dim result As Object
    Dim keyName As String
    Dim value As Variant

    SkipJsonWhitespace jsonText, position
    If JsonCharacter(jsonText, position) <> "{" Then
        report = "Expected an object in definition JSON."
        Exit Function
    End If
    position = position + 1
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Do
        SkipJsonWhitespace jsonText, position
        If JsonCharacter(jsonText, position) = "}" Then
            position = position + 1
            Exit Do
        End If
        keyName = ParseJsonString(jsonText, position, report)
        If report <> "" Then Exit Function
        SkipJsonWhitespace jsonText, position
        If JsonCharacter(jsonText, position) <> ":" Then
            report = "Expected ':' after a definition JSON key."
            Exit Function
        End If
        position = position + 1
        value = ParseJsonScalar(jsonText, position, report)
        If report <> "" Then Exit Function
        result(keyName) = value
        SkipJsonWhitespace jsonText, position
        Select Case JsonCharacter(jsonText, position)
            Case ",": position = position + 1
            Case "}": position = position + 1: Exit Do
            Case Else
                report = "Expected ',' or '}' in definition JSON object."
                Exit Function
        End Select
    Loop
    Set ParseFlatJsonObject = result
End Function

Private Function ParseJsonScalar(ByVal jsonText As String, ByRef position As Long, _
                                 ByRef report As String) As Variant
    Dim startAt As Long
    Dim token As String
    Dim ch As String

    SkipJsonWhitespace jsonText, position
    If JsonCharacter(jsonText, position) = Chr$(34) Then
        ParseJsonScalar = ParseJsonString(jsonText, position, report)
        Exit Function
    End If
    startAt = position
    Do While position <= Len(jsonText)
        ch = JsonCharacter(jsonText, position)
        If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit Do
        position = position + 1
    Loop
    token = Mid$(jsonText, startAt, position - startAt)
    Select Case LCase$(token)
        Case "null": ParseJsonScalar = Empty
        Case "true": ParseJsonScalar = True
        Case "false": ParseJsonScalar = False
        Case Else
            If IsNumeric(token) Then
                ParseJsonScalar = CDbl(token)
            Else
                report = "Unsupported definition JSON value: " & token
            End If
    End Select
End Function

Private Function ParseJsonString(ByVal jsonText As String, ByRef position As Long, _
                                 ByRef report As String) As String
    Dim ch As String
    Dim escaped As Boolean

    If JsonCharacter(jsonText, position) <> Chr$(34) Then
        report = "Expected a quoted definition JSON string."
        Exit Function
    End If
    position = position + 1
    Do While position <= Len(jsonText)
        ch = Mid$(jsonText, position, 1)
        position = position + 1
        If escaped Then
            Select Case ch
                Case Chr$(34), "\", "/": ParseJsonString = ParseJsonString & ch
                Case "n": ParseJsonString = ParseJsonString & vbLf
                Case "r": ParseJsonString = ParseJsonString & vbCr
                Case "t": ParseJsonString = ParseJsonString & vbTab
                Case Else: ParseJsonString = ParseJsonString & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = Chr$(34) Then
            Exit Function
        Else
            ParseJsonString = ParseJsonString & ch
        End If
    Loop
    report = "Unterminated definition JSON string."
End Function

Private Sub SkipJsonWhitespace(ByVal jsonText As String, ByRef position As Long)
    Dim ch As String
    Do While position <= Len(jsonText)
        ch = Mid$(jsonText, position, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        position = position + 1
    Loop
End Sub

Private Function JsonCharacter(ByVal jsonText As String, ByVal position As Long) As String
    If position >= 1 And position <= Len(jsonText) Then JsonCharacter = Mid$(jsonText, position, 1)
End Function

Private Function IsReusableLifecycleEvent(ByVal eventType As String) As Boolean
    Select Case eventType
        Case "PROCESS_SAVE", "PROCESS_RELEASE", "PROCESS_OBSOLETE", _
             "RECIPE_SAVE", "RECIPE_RELEASE", "RECIPE_OBSOLETE"
            IsReusableLifecycleEvent = True
    End Select
End Function

Private Function ExpectedReusableStatus(ByVal eventType As String) As String
    Select Case eventType
        Case "PROCESS_SAVE", "RECIPE_SAVE": ExpectedReusableStatus = "DRAFT"
        Case "PROCESS_RELEASE", "RECIPE_RELEASE": ExpectedReusableStatus = "RELEASED"
        Case "PROCESS_OBSOLETE", "RECIPE_OBSOLETE": ExpectedReusableStatus = "OBSOLETE"
    End Select
End Function

Private Function ReusableDefinitionKind(ByVal eventType As String) As String
    If Left$(eventType, 7) = "PROCESS" Then
        ReusableDefinitionKind = "Process"
    Else
        ReusableDefinitionKind = "Recipe"
    End If
End Function

Private Function ReusableDefinitionStatus(ByVal eventType As String, _
                                          ByVal definitionId As String, _
                                          ByVal definitionVersion As String) As String
    Dim rows As Variant
    Dim r As Long

    If Left$(eventType, 7) = "PROCESS" Then
        rows = modOperationsPrimitiveBridge.ListProcesses("")
    Else
        rows = modOperationsPrimitiveBridge.ListRecipes("")
    End If
    If Not IsArray(rows) Then Exit Function
    On Error GoTo NoRows
    For r = LBound(rows, 1) To UBound(rows, 1)
        If StrComp(Trim$(CStr(rows(r, 1))), definitionId, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(rows(r, 2))), definitionVersion, vbTextCompare) = 0 Then
            ReusableDefinitionStatus = Trim$(CStr(rows(r, 5)))
            Exit Function
        End If
    Next r
NoRows:
End Function
