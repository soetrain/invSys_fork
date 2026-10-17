VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCapturedFormStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Each authority owns a separate instance. This stores references only; callers
' retain all authorization, session and source checks at their owning boundary.
Private mEntries As Object

Private Sub Class_Initialize()
    Set mEntries = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Add(ByVal key As String, ByVal entry As Object)
    mEntries.Add key, entry
End Sub

Public Function Item(ByVal key As String) As Object
    If mEntries.Exists(key) Then Set Item = mEntries(key)
End Function

Public Sub Remove(ByVal key As String)
    If mEntries.Exists(key) Then mEntries.Remove key
End Sub
