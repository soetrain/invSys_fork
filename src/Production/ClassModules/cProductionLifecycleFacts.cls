VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProductionLifecycleFacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Operations-local observations; never a claim of Domain application.
Public EventId As String
Public SubmissionState As String
Public OutcomeCode As String

Public Sub ObserveSubmission(ByVal id As String, ByVal submitted As Boolean, ByVal writeAttempted As Boolean)
    EventId = "": SubmissionState = "": OutcomeCode = "FAILED"
    If Not submitted And Not writeAttempted Then Exit Sub
    EventId = id: SubmissionState = IIf(submitted, "Submitted", "Unknown")
    If submitted Then OutcomeCode = "PENDING"
End Sub
