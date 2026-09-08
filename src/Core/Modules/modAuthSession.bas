Attribute VB_Name = "modAuthSession"
Option Explicit
Option Private Module

Private mVersion As Long

' Called only by the owning authentication boundary, never by observers.
Public Sub Invalidate()
    mVersion = mVersion + 1
End Sub

Public Function Version() As Long
    Version = mVersion
End Function
