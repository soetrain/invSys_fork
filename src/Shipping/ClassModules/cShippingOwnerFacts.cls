VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShippingOwnerFacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Per-action Shipping facts. No authority reads, activity writes or UI work.
Private mSources As Object
Private mOutcome As String
Private mUnavailable As Boolean
Private mFailed As Boolean
Private mCompletedOwners As Long

Public Sub BeginOwner()
    mOutcome = "FAILED"
End Sub

Public Sub RejectBeforeMutation(Optional ByVal denied As Boolean = False)
    If mCompletedOwners > 0 Then
        mFailed = True: mOutcome = "FAILED"
        Exit Sub
    End If
    mOutcome = "REJECTED"
    If denied Then mOutcome = "DENIED"
End Sub

' Called only after entering the real submission boundary, including uncertain exits.
Public Sub ObserveSource(ByVal eventId As String, ByVal accepted As Boolean)
    Dim i As Long, code As Long
    On Error GoTo Unavailable
    If eventId = "" Then
        If accepted Then mUnavailable = True
        Exit Sub
    End If
    If Len(eventId) > 128 Then GoTo Unavailable
    For i = 1 To Len(eventId)
        code = AscW(Mid$(eventId, i, 1))
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or _
                (code >= 97 And code <= 122) Or code = 45 Or code = 95) Then GoTo Unavailable
    Next i
    If mSources Is Nothing Then
        Set mSources = CreateObject("Scripting.Dictionary")
        mSources.CompareMode = vbBinaryCompare
    End If
    If mSources.Exists(eventId) Then
        mSources(eventId) = CBool(mSources(eventId)) Or accepted
    Else
        mSources.Add eventId, accepted
    End If
    Exit Sub
Unavailable:
    mUnavailable = True
End Sub

Public Sub CompleteOwner(ByVal completed As Boolean, _
                         Optional ByVal requiredStepFailed As Boolean = False, _
                         Optional ByVal processingFinished As Boolean = False)
    mOutcome = "FAILED"
    If Not completed Or requiredStepFailed Then mFailed = True
    If mFailed Then Exit Sub
    mCompletedOwners = mCompletedOwners + 1
    If SourceIds(False) <> "" Then Exit Sub
    mOutcome = "STAGED"
    If SourceIds(True) <> "" Then
        mOutcome = "PENDING"
        If processingFinished Then mOutcome = "CONFIRMED"
    End If
End Sub

Public Property Get OutcomeCode() As String
    OutcomeCode = mOutcome
    If OutcomeCode = "" Then OutcomeCode = "FAILED"
End Property

Public Property Get EvidenceAvailable() As Boolean
    EvidenceAvailable = Not mUnavailable
End Property

Public Function SourceIds(ByVal accepted As Boolean) As String
    Dim key As Variant
    On Error GoTo Unavailable
    If mSources Is Nothing Then Exit Function
    For Each key In mSources.Keys
        If CBool(mSources(key)) = accepted Then
            If SourceIds <> "" Then SourceIds = SourceIds & vbLf
            SourceIds = SourceIds & CStr(key)
        End If
    Next key
    Exit Function
Unavailable:
    mUnavailable = True
    SourceIds = ""
End Function
