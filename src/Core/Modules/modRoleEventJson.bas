Attribute VB_Name = "modRoleEventJson"
Option Explicit
Option Private Module

Public Function DictionaryToJson(ByVal d As Object) As String
    Dim keys As Variant
    Dim i As Long
    Dim key As String

    DictionaryToJson = "{"
    keys = d.Keys
    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i))
        If i > LBound(keys) Then DictionaryToJson = DictionaryToJson & ","
        DictionaryToJson = DictionaryToJson & """" & EscapeJsonRole(key) & """:" & JsonValueRole(d(key))
    Next i
    DictionaryToJson = DictionaryToJson & "}"
End Function

Private Function JsonValueRole(ByVal valueIn As Variant) As String
    Select Case True
        Case IsObject(valueIn)
            JsonValueRole = "null"
        Case IsNull(valueIn), IsEmpty(valueIn)
            JsonValueRole = "null"
        Case VarType(valueIn) = vbBoolean
            JsonValueRole = IIf(CBool(valueIn), "true", "false")
        Case IsNumeric(valueIn)
            JsonValueRole = Replace$(CStr(valueIn), ",", "")
        Case Else
            JsonValueRole = """" & EscapeJsonRole(CStr(valueIn)) & """"
    End Select
End Function

Private Function EscapeJsonRole(ByVal textIn As String) As String
    EscapeJsonRole = textIn
    EscapeJsonRole = Replace$(EscapeJsonRole, "\", "\\")
    EscapeJsonRole = Replace$(EscapeJsonRole, Chr$(34), "\" & Chr$(34))
    EscapeJsonRole = Replace$(EscapeJsonRole, vbCrLf, "\n")
    EscapeJsonRole = Replace$(EscapeJsonRole, vbCr, "\n")
    EscapeJsonRole = Replace$(EscapeJsonRole, vbLf, "\n")
    EscapeJsonRole = Replace$(EscapeJsonRole, vbTab, "\t")
End Function

Public Function ReadStagedInboxRowsRole(ByVal filePath As String, ByRef report As String) As Collection
    On Error GoTo FailRead

    Dim rows As Collection
    Dim fileNum As Integer
    Dim lineText As String
    Dim rowValues As Object
    Dim parseReport As String

    Set rows = New Collection
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim$(lineText)
        If lineText <> "" Then
            parseReport = vbNullString
            Set rowValues = ParseJsonObjectRole(lineText, parseReport)
            If rowValues Is Nothing Then
                report = "Unable to parse local staged inbox row in " & filePath & ": " & parseReport
                Close #fileNum
                Exit Function
            End If
            rows.Add rowValues
        End If
    Loop
    Close #fileNum
    Set ReadStagedInboxRowsRole = rows
    Exit Function

FailRead:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    report = "Unable to read local staged inbox file " & filePath & ": " & Err.Description
End Function

Public Function ParseJsonObjectRole(ByVal jsonText As String, ByRef report As String) As Object
    On Error GoTo FailParse

    Dim pos As Long
    Dim keyText As String
    Dim valueText As Variant
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    pos = 1
    SkipJsonWhitespaceRole jsonText, pos
    If Mid$(jsonText, pos, 1) <> "{" Then
        report = "Expected object start."
        Exit Function
    End If
    pos = pos + 1

    Do
        SkipJsonWhitespaceRole jsonText, pos
        If Mid$(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Exit Do
        End If
        keyText = ParseJsonStringRole(jsonText, pos, report)
        If report <> "" Then Exit Function
        SkipJsonWhitespaceRole jsonText, pos
        If Mid$(jsonText, pos, 1) <> ":" Then
            report = "Expected ':' after key."
            Exit Function
        End If
        pos = pos + 1
        valueText = ParseJsonValueRole(jsonText, pos, report)
        If report <> "" Then Exit Function
        d(keyText) = valueText
        SkipJsonWhitespaceRole jsonText, pos
        Select Case Mid$(jsonText, pos, 1)
            Case ","
                pos = pos + 1
            Case "}"
                pos = pos + 1
                Exit Do
            Case Else
                report = "Expected ',' or object end."
                Exit Function
        End Select
    Loop While pos <= Len(jsonText)

    Set ParseJsonObjectRole = d
    Exit Function

FailParse:
    report = Err.Description
End Function

Private Function ParseJsonValueRole(ByVal jsonText As String, ByRef pos As Long, ByRef report As String) As Variant
    Dim startPos As Long
    Dim token As String

    SkipJsonWhitespaceRole jsonText, pos
    Select Case Mid$(jsonText, pos, 1)
        Case Chr$(34)
            ParseJsonValueRole = ParseJsonStringRole(jsonText, pos, report)
        Case Else
            startPos = pos
            Do While pos <= Len(jsonText)
                Select Case Mid$(jsonText, pos, 1)
                    Case ",", "}", " ", vbTab, vbCr, vbLf
                        Exit Do
                End Select
                pos = pos + 1
            Loop
            token = Trim$(Mid$(jsonText, startPos, pos - startPos))
            Select Case LCase$(token)
                Case "null"
                    ParseJsonValueRole = ""
                Case "true"
                    ParseJsonValueRole = True
                Case "false"
                    ParseJsonValueRole = False
                Case Else
                    If IsNumeric(token) Then
                        ParseJsonValueRole = CDbl(token)
                    Else
                        ParseJsonValueRole = token
                    End If
            End Select
    End Select
End Function

Private Function ParseJsonStringRole(ByVal jsonText As String, ByRef pos As Long, ByRef report As String) As String
    Dim ch As String
    Dim escaped As Boolean

    If Mid$(jsonText, pos, 1) <> Chr$(34) Then
        report = "Expected string."
        Exit Function
    End If
    pos = pos + 1
    Do While pos <= Len(jsonText)
        ch = Mid$(jsonText, pos, 1)
        pos = pos + 1
        If escaped Then
            Select Case ch
                Case Chr$(34), "\", "/"
                    ParseJsonStringRole = ParseJsonStringRole & ch
                Case "n"
                    ParseJsonStringRole = ParseJsonStringRole & vbLf
                Case "r"
                    ParseJsonStringRole = ParseJsonStringRole & vbCr
                Case "t"
                    ParseJsonStringRole = ParseJsonStringRole & vbTab
                Case Else
                    ParseJsonStringRole = ParseJsonStringRole & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = Chr$(34) Then
            Exit Function
        Else
            ParseJsonStringRole = ParseJsonStringRole & ch
        End If
    Loop
    report = "Unterminated string."
End Function

Private Sub SkipJsonWhitespaceRole(ByVal jsonText As String, ByRef pos As Long)
    Do While pos <= Len(jsonText)
        Select Case Mid$(jsonText, pos, 1)
            Case " ", vbTab, vbCr, vbLf
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub
