VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cRecordingControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mButtons As Collection
Private mBindings As Collection
Private mStatus As MSForms.Label
Private mContext As String
Private mRendering As Boolean

Public Sub Initialize(ByVal form As Object)
    Dim command As Variant, button As MSForms.CommandButton, binding As cRecordingButton
    Set mButtons = New Collection: Set mBindings = New Collection
    For Each command In Array("Start", "Stop", "Cancel")
        Set button = form.Controls.Add("Forms.CommandButton.1", "btn" & CStr(command) & "Recording", True)
        button.Caption = CStr(command) & " Recording"
        mButtons.Add button, CStr(command)
        Set binding = New cRecordingButton
        binding.Connect button, Me, CStr(command)
        mBindings.Add binding
    Next command
    Set mStatus = form.Controls.Add("Forms.Label.1", "lblRecordingStatus", True)
    mStatus.WordWrap = True
End Sub

Public Sub BindContext(ByVal context As String)
    If mContext <> "" And mContext <> context Then modActionRecording.ContextClosed mContext
    mContext = context
    Render
End Sub

Public Sub Execute(ByVal command As String)
    Dim report As String, succeeded As Boolean
    succeeded = modActionRecording.Control(mContext, command, report)
    Render
    If Not succeeded Then mStatus.Caption = report: mStatus.ControlTipText = report
End Sub

Public Sub Render()
    Dim active As Boolean, canStart As Boolean
    If mStatus Is Nothing Or mRendering Then Exit Sub
    On Error GoTo Done
    mRendering = True
    mStatus.Caption = modActionRecording.Status(mContext, active, canStart)
    mStatus.ControlTipText = mStatus.Caption
    mButtons("Start").Enabled = canStart
    mButtons("Stop").Enabled = active
    mButtons("Cancel").Enabled = active
Done:
    mRendering = False
End Sub

Public Sub Arrange(ByVal visible As Boolean, ByVal width As Single)
    Dim button As MSForms.CommandButton, index As Long
    If mButtons Is Nothing Then Exit Sub
    For Each button In mButtons
        button.Visible = visible
        button.Move 12 + index * 124, 184, 116, 26
        index = index + 1
    Next button
    mStatus.Visible = visible
    mStatus.Move 392, 184, width - 380, 40
End Sub

Public Sub Disconnect()
    Dim binding As cRecordingButton
    If mBindings Is Nothing Then Exit Sub
    modActionRecording.ContextClosed mContext
    For Each binding In mBindings: binding.Disconnect: Next binding
    Set mBindings = Nothing: Set mButtons = Nothing: Set mStatus = Nothing
    mContext = ""
End Sub
