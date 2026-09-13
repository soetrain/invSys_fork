VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShippingActionTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLog() As String
Private mCount As Long
Private mStart As Single

Public Sub Start()
    mCount = 0
    Erase mLog
    mStart = Timer
End Sub

Public Sub Mark(ByVal label As String)
    Dim elapsedMs As Long
    If mStart <= 0 Then mStart = Timer
    elapsedMs = ElapsedMilliseconds(mStart)
    mCount = mCount + 1
    ReDim Preserve mLog(1 To mCount)
    mLog(mCount) = Format$(elapsedMs, "00000") & " ms  " & label
End Sub

Public Function Summary() As String
    Dim i As Long, lines As String
    For i = 1 To mCount
        If lines <> "" Then lines = lines & vbCrLf
        lines = lines & mLog(i)
    Next i
    Summary = lines
End Function

Public Function ElapsedMilliseconds(ByVal startedAt As Single) As Long
    Dim deltaSeconds As Single
    deltaSeconds = Timer - startedAt
    If deltaSeconds < 0 Then deltaSeconds = deltaSeconds + 86400!
    ElapsedMilliseconds = CLng(deltaSeconds * 1000)
End Function
