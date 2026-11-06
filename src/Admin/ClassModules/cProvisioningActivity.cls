VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProvisioningActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mContext As String, mNotice As String, mOpening As String
Private mTarget As Workbook, mPresented As Boolean
Private mPrefix As String, mLabel As String

Public Sub Initialize(ByVal target As Workbook, Optional ByVal prefix As String = "ADMIN_CREATE_", Optional ByVal label As String = "Create Warehouse")
    On Error GoTo Unavailable
    mPrefix = prefix: mLabel = label
    Set mTarget = target
    mContext = modActivity.CaptureContext()
    mOpening = BeginAction("OPEN")
    Exit Sub
Unavailable:
    UnavailableNotice
End Sub

Public Function BeginAction(ByVal suffix As String) As String
    Dim surface As Object
    On Error GoTo Unavailable
    mNotice = ""
    Set surface = modAdminActionObservation.CurrentSurface()
    If Not TargetLive() Then GoTo Unavailable
    BeginAction = modActivity.BeginAction(mPrefix & suffix, mContext, mNotice)
    If mContext <> modActivity.CaptureContext() Then UnavailableNotice
Finished:
    modAdminActionObservation.RestoreSurface surface
    Exit Function
Unavailable:
    UnavailableNotice
    Resume Finished
End Function

Public Sub Finish(ByVal action As String, ByVal outcome As String, Optional ByVal sourceReferencesJson As String = "[]")
    Dim surface As Object, notice As String
    On Error GoTo Unavailable
    Set surface = modAdminActionObservation.CurrentSurface()
    If action <> "" Then
        If Not TargetLive() Then GoTo Unavailable
        If Not modActivity.FinishAction(action, outcome, notice, sourceReferencesJson) And notice = "" Then _
            notice = "Tracking unavailable: the " & mLabel & " result could not be recorded."
        If notice <> "" Then mNotice = notice
    End If
Finished:
    modAdminActionObservation.RestoreSurface surface
    ShowNotice
    Exit Sub
Unavailable:
    UnavailableNotice
    Resume Finished
End Sub

Public Sub Presented(ByVal visible As Boolean)
    If mPresented Or Not visible Then Exit Sub
    mPresented = True
    Finish mOpening, "OPENED"
End Sub

Public Sub OpeningFailed()
    If mPresented Then Exit Sub
    mPresented = True
    Finish mOpening, "FAILED"
End Sub

Private Function TargetLive() As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb Is mTarget Then TargetLive = True: Exit Function
    Next wb
End Function

Private Sub UnavailableNotice()
    mNotice = "Tracking unavailable: " & mLabel & " could not be recorded."
    On Error Resume Next
    modActionRecording.ContextClosed mContext, True
End Sub

Private Sub ShowNotice()
    Dim notice As String
    notice = mNotice: mNotice = ""
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
End Sub
