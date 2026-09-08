Attribute VB_Name = "modSystemIdentity"
Option Explicit
Option Private Module

Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As Any) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As Any, ByVal buffer As LongPtr, ByVal length As Long) As Long

' D14 identities cannot depend on the host's shared, reseedable VBA Rnd state.
Public Function NewId() As String
    Dim bytes(0 To 15) As Byte, buffer As String
    buffer = String$(39, vbNullChar)
    If CoCreateGuid(bytes(0)) <> 0 Then Err.Raise 5, , "System identity unavailable."
    If StringFromGUID2(bytes(0), StrPtr(buffer), 39) <> 39 Then Err.Raise 5, , "System identity unavailable."
    NewId = LCase$(Mid$(buffer, 2, 36))
End Function
