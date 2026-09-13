VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShippingActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFacts As cShippingOwnerFacts
Private mActivityId As String
Private mNotice As String
Private mFinished As Boolean

Public Property Get Facts() As cShippingOwnerFacts
    If mFacts Is Nothing Then Set mFacts = New cShippingOwnerFacts
    Set Facts = mFacts
End Property

' Only real handlers call Start; owning services receive facts, not this observer.
Public Sub Start(ByVal controlId As String, ByVal context As String)
    On Error GoTo Unavailable
    mActivityId = modActivity.BeginAction(controlId, context, mNotice)
    Exit Sub
Unavailable:
    mNotice = "Tracking unavailable: the action could not be recorded."
End Sub

Public Function Finish(Optional ByVal outcomeOverride As String = "") As String
    Dim outcome As String, submitted As String, uncertain As String, references As String
    On Error GoTo Unavailable
    If mFinished Then Exit Function
    mFinished = True
    If mActivityId = "" Then GoTo Done
    outcome = outcomeOverride
    If outcome = "" Then outcome = Facts.OutcomeCode
    submitted = modActivity.InventorySourceReferences(mActivityId, Facts.SourceIds(True), "Submitted")
    uncertain = modActivity.InventorySourceReferences(mActivityId, Facts.SourceIds(False), "Unknown")
    If Not Facts.EvidenceAvailable Or submitted = "" Or uncertain = "" Then GoTo Unavailable
    references = submitted
    If submitted = "[]" Then
        references = uncertain
    ElseIf uncertain <> "[]" Then
        references = Left$(submitted, Len(submitted) - 1) & "," & Mid$(uncertain, 2)
    End If
    modActivity.FinishAction mActivityId, outcome, mNotice, references
Done:
    Finish = mNotice
    Exit Function
Unavailable:
    mNotice = "Tracking unavailable: the result could not be recorded."
    Resume Done
End Function
