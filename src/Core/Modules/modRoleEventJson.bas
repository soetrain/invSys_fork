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
