VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProductionWorksheetAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWorkbook As Workbook
Private mContext As String
Private mActivityId As String
Private mNotice As String
Private mReferences As String
Private mInvalidReferences As Boolean
Public OutcomeCode As String

Public Property Get OperatorWorkbook() As Workbook
    Set OperatorWorkbook = mWorkbook
End Property

' Check In uses the same current-action guard before its observation contract exists.
Public Function BindForValidation(ByVal capturedContext As String, ByVal operatorBook As Workbook, _
                                  ByRef report As String) As Boolean
    Set mWorkbook = operatorBook: mContext = capturedContext
    OutcomeCode = "FAILED"
    If Not CanContinue(report) Then Exit Function
    BindForValidation = CurrentBinding(report)
End Function

Public Function Begin(ByVal controlId As String, ByVal capturedContext As String, _
                      ByVal operatorBook As Workbook, ByRef report As String) As Boolean
    Set mWorkbook = operatorBook: mContext = capturedContext
    OutcomeCode = "FAILED": mReferences = ""
    If Not CurrentBinding(report) Then Exit Function
    StartTracking controlId
    If Not CurrentBinding(report) Then Exit Function
    If Not Permitted() Then
        OutcomeCode = "DENIED"
        report = "Production permission is required; the draft was not changed."
        Exit Function
    End If
    Begin = True
End Function

Private Sub StartTracking(ByVal controlId As String)
    On Error GoTo Unavailable
    mActivityId = modActivity.BeginAction(controlId, mContext, mNotice)
    Exit Sub
Unavailable:
    mNotice = "Tracking unavailable: the action could not be recorded."
End Sub

Public Function CanContinue(ByRef report As String) As Boolean
    If Not CurrentBinding(report) Then OutcomeCode = "FAILED": Exit Function
    If Not Permitted() Then
        OutcomeCode = "FAILED"
        report = "Production permission changed. Reopen Production before continuing."
        Exit Function
    End If
    CanContinue = True
End Function

Private Function CurrentBinding(ByRef report As String) As Boolean
    CurrentBinding = modProductionDesignerActions.ContextIsCurrent(mContext, mWorkbook)
    If Not CurrentBinding Then report = "Session, warehouse, or captured workbook changed. Reopen Production before editing the draft."
End Function

Private Function Permitted() As Boolean
    Permitted = modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST")
    If Not Permitted Then Permitted = modRoleUiAccess.CanCurrentUserPerformCapability("ADMIN_MAINT")
End Function

' Only the owning submission supplies identities and acknowledgment states.
' Core serializes each reference using this action's captured warehouse.
Public Sub ObserveSubmission(ByVal facts As cProductionLifecycleFacts)
    Dim wire As String
    On Error GoTo Unavailable
    If facts Is Nothing Or mActivityId = "" Then Exit Sub
    If facts.EventId = "" Then Exit Sub
    wire = modActivity.DesignsSourceReferences(mActivityId, facts.EventId, facts.SubmissionState)
    If Len(wire) <= 2 Then mInvalidReferences = True: Exit Sub
    If mReferences <> "" Then mReferences = mReferences & ","
    mReferences = mReferences & Mid$(wire, 2, Len(wire) - 2)
    Exit Sub
Unavailable:
    mInvalidReferences = True
End Sub

Public Sub Finish(ByRef report As String)
    Dim references As String
    On Error GoTo Unavailable
    references = "[" & mReferences & "]"
    If mInvalidReferences Then references = ""
    If mActivityId <> "" Then modActivity.FinishAction mActivityId, OutcomeCode, mNotice, references
Done:
    If mNotice <> "" Then report = report & " " & mNotice
    Exit Sub
Unavailable:
    mNotice = "Tracking unavailable: the result could not be recorded."
    Resume Done
End Sub
