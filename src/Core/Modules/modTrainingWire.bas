Attribute VB_Name = "modTrainingWire"
Option Explicit

Private Type SystemTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef value As SystemTime)
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As Any) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As Any, ByVal buffer As LongPtr, ByVal length As Long) As Long
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef algorithm As LongPtr, ByVal name As LongPtr, ByVal provider As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt" (ByVal algorithm As LongPtr, ByRef hash As LongPtr, ByVal buffer As LongPtr, ByVal bufferLength As Long, ByVal secret As LongPtr, ByVal secretLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptHashData Lib "bcrypt" (ByVal hash As LongPtr, ByRef data As Any, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt" (ByVal hash As LongPtr, ByRef digest As Any, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal algorithm As LongPtr, ByVal flags As Long) As Long

Public Function NewId() As String
    Dim bytes(0 To 15) As Byte, buffer As String
    buffer = String$(39, vbNullChar)
    If CoCreateGuid(bytes(0)) <> 0 Then Err.Raise 5, , "Training identity unavailable."
    If StringFromGUID2(bytes(0), StrPtr(buffer), 39) <> 39 Then Err.Raise 5, , "Training identity unavailable."
    NewId = LCase$(Mid$(buffer, 2, 36))
End Function

Public Function UtcTimestamp() As String
    Dim value As SystemTime
    GetSystemTime value
    UtcTimestamp = Format$(DateSerial(value.Year, value.Month, value.Day), "yyyy-mm-dd") & "T" & _
        Format$(TimeSerial(value.Hour, value.Minute, value.Second), "hh:nn:ss") & "." & _
        Right$("000" & CStr(value.Milliseconds), 3) & "Z"
End Function

' JSON is ASCII-escaped UTF-8: hash exactly the persisted bytes, without a BOM.
' CNG owns the hash buffer when BCryptCreateHash receives a null buffer.
Public Function Sha256(ByVal asciiText As String) As String
    Dim algorithm As LongPtr, hash As LongPtr, bytes() As Byte
    Dim digest(0 To 31) As Byte, i As Long, name As String
    On Error GoTo Failed
    For i = 1 To Len(asciiText)
        If AscW(Mid$(asciiText, i, 1)) < 0 Or AscW(Mid$(asciiText, i, 1)) > 127 Then GoTo Failed
    Next i
    name = "SHA256"
    If BCryptOpenAlgorithmProvider(algorithm, StrPtr(name), 0, 0) <> 0 Then GoTo Failed
    If BCryptCreateHash(algorithm, hash, 0, 0, 0, 0, 0) <> 0 Then GoTo Failed
    If Len(asciiText) > 0 Then
        bytes = StrConv(asciiText, vbFromUnicode)
        If BCryptHashData(hash, bytes(0), UBound(bytes) + 1, 0) <> 0 Then GoTo Failed
    End If
    If BCryptFinishHash(hash, digest(0), 32, 0) <> 0 Then GoTo Failed
    For i = 0 To 31
        Sha256 = Sha256 & LCase$(Right$("0" & Hex$(digest(i)), 2))
    Next i
CleanExit:
    If hash <> 0 Then BCryptDestroyHash hash
    If algorithm <> 0 Then BCryptCloseAlgorithmProvider algorithm, 0
    Exit Function
Failed:
    Sha256 = ""
    Resume CleanExit
End Function

Public Function ValidId(ByVal value As String) As Boolean
    Dim i As Long, ch As String
    If Len(value) <> 36 Then Exit Function
    For i = 1 To 36
        ch = Mid$(value, i, 1)
        If i = 9 Or i = 14 Or i = 19 Or i = 24 Then
            If ch <> "-" Then Exit Function
        ElseIf InStr(1, "0123456789abcdef", ch, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next i
    ValidId = True
End Function

Public Function ValidUtcTimestamp(ByVal value As String) As Boolean
    Dim day As Date
    On Error GoTo Invalid
    If Not value Like "####-##-##T##:##:##.###Z" Then Exit Function
    day = DateSerial(CLng(Left$(value, 4)), CLng(Mid$(value, 6, 2)), CLng(Mid$(value, 9, 2)))
    If Format$(day, "yyyy-mm-dd") <> Left$(value, 10) Then Exit Function
    If CLng(Mid$(value, 12, 2)) > 23 Or CLng(Mid$(value, 15, 2)) > 59 Or CLng(Mid$(value, 18, 2)) > 59 Then Exit Function
    ValidUtcTimestamp = True
Invalid:
End Function

Public Function ValidSegment(ByVal value As String) As Boolean
    Dim i As Long, ch As String
    If Len(value) = 0 Or Len(value) > 128 Or value = "." Or value = ".." Then Exit Function
    If Right$(value, 1) = "." Or Right$(value, 1) = " " Then Exit Function
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If AscW(ch) >= 0 And AscW(ch) < 32 Then Exit Function
        If InStr(1, "\/:*?""<>|", ch, vbBinaryCompare) > 0 Then Exit Function
    Next i
    ValidSegment = True
End Function
