Attribute VB_Name = "modTrainingJson"
Option Explicit
Option Private Module

Public Function Quote(ByVal value As String) As String
    Dim i As Long, code As Long, ch As String, result As String
    result = """"
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case code
            Case 34: result = result & "\"""
            Case 92: result = result & "\\"
            Case 32 To 126: result = result & ch
            Case Else: result = result & "\u" & Right$("0000" & Hex$(code), 4)
        End Select
    Next i
    Quote = result & """"
End Function

Public Function EncodeObject(ByVal values As Object) As String
    Dim key As Variant, result As String
    result = "{"
    For Each key In values.Keys
        If Len(result) > 1 Then result = result & ","
        result = result & Quote(CStr(key)) & ":" & EncodeValue(values(key))
    Next key
    EncodeObject = result & "}"
End Function

Private Function EncodeValue(ByVal value As Variant) As String
    Dim i As Long, result As String
    If IsObject(value) Then
        If TypeName(value) = "Collection" Then
            result = "["
            For i = 1 To value.Count
                If i > 1 Then result = result & ","
                result = result & EncodeValue(value(i))
            Next i
            EncodeValue = result & "]"
        Else
            EncodeValue = EncodeObject(value)
        End If
    Else
        Select Case VarType(value)
            Case vbString: EncodeValue = Quote(CStr(value))
            Case vbBoolean: EncodeValue = IIf(CBool(value), "true", "false")
            Case vbInteger, vbLong: EncodeValue = CStr(value)
            Case Else: Err.Raise 5, , "Unsupported training value."
        End Select
    End If
End Function

Public Function DecodeObject(ByVal text As String) As Object
    Dim pos As Long, result As Object
    On Error GoTo Invalid
    If Len(text) = 0 Or Len(text) > 1048576 Then Exit Function
    pos = 1
    Set result = ReadObject(text, pos, 0)
    SkipSpace text, pos
    If pos <> Len(text) + 1 Then Exit Function
    Set DecodeObject = result
Invalid:
End Function

Private Function ReadObject(ByVal text As String, ByRef pos As Long, ByVal depth As Long) As Object
    Dim result As Object, key As String, value As Variant
    If depth > 16 Then Err.Raise 5
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbBinaryCompare
    Require text, pos, "{"
    SkipSpace text, pos
    If Mid$(text, pos, 1) <> "}" Then
        Do
            key = ReadString(text, pos)
            If result.Exists(key) Then Err.Raise 5
            Require text, pos, ":"
            ReadValue text, pos, depth + 1, value
            result.Add key, value
            SkipSpace text, pos
            If Mid$(text, pos, 1) = "}" Then Exit Do
            Require text, pos, ","
        Loop
    End If
    Require text, pos, "}"
    Set ReadObject = result
End Function

Private Sub ReadValue(ByVal text As String, ByRef pos As Long, ByVal depth As Long, ByRef value As Variant)
    Dim list As Collection, item As Variant, start As Long, token As String
    If depth > 16 Then Err.Raise 5
    If IsObject(value) Then Set value = Nothing
    value = Empty
    SkipSpace text, pos
    Select Case Mid$(text, pos, 1)
        Case """": value = ReadString(text, pos)
        Case "{": Set value = ReadObject(text, pos, depth)
        Case "["
            Set list = New Collection
            pos = pos + 1
            SkipSpace text, pos
            If Mid$(text, pos, 1) <> "]" Then
                Do
                    ReadValue text, pos, depth + 1, item
                    list.Add item
                    SkipSpace text, pos
                    If Mid$(text, pos, 1) = "]" Then Exit Do
                    Require text, pos, ","
                Loop
            End If
            Require text, pos, "]"
            Set value = list
        Case Else
            start = pos
            Do While pos <= Len(text)
                If InStr(1, ",}] " & vbCr & vbLf & vbTab, Mid$(text, pos, 1), vbBinaryCompare) > 0 Then Exit Do
                pos = pos + 1
            Loop
            token = Mid$(text, start, pos - start)
            Select Case token
                Case "true": value = True
                Case "false": value = False
                Case Else
                    If token = "" Or token Like "*[!0-9]*" Then Err.Raise 5
                    If Len(token) > 1 And Left$(token, 1) = "0" Then Err.Raise 5
                    value = CLng(token)
            End Select
    End Select
End Sub

Private Function ReadString(ByVal text As String, ByRef pos As Long) As String
    Dim ch As String, digits As String, result As String
    Require text, pos, """"
    Do While pos <= Len(text)
        ch = Mid$(text, pos, 1)
        pos = pos + 1
        If ch = """" Then ReadString = result: Exit Function
        If ch = "\" Then
            ch = Mid$(text, pos, 1)
            pos = pos + 1
            Select Case ch
                Case """", "\", "/": result = result & ch
                Case "b": result = result & Chr$(8)
                Case "f": result = result & Chr$(12)
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    digits = Mid$(text, pos, 4)
                    If Len(digits) <> 4 Or digits Like "*[!0-9A-Fa-f]*" Then Err.Raise 5
                    result = result & ChrW$(CLng("&H" & digits))
                    pos = pos + 4
                Case Else: Err.Raise 5
            End Select
        Else
            If AscW(ch) < 32 Then Err.Raise 5
            result = result & ch
        End If
    Loop
    Err.Raise 5
End Function

Private Sub Require(ByVal text As String, ByRef pos As Long, ByVal expected As String)
    SkipSpace text, pos
    If Mid$(text, pos, 1) <> expected Then Err.Raise 5
    pos = pos + 1
End Sub

Private Sub SkipSpace(ByVal text As String, ByRef pos As Long)
    Do While pos <= Len(text)
        If InStr(1, " " & vbCr & vbLf & vbTab, Mid$(text, pos, 1), vbBinaryCompare) = 0 Then Exit Do
        pos = pos + 1
    Loop
End Sub
