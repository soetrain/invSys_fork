VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cRecordingButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mButton As MSForms.CommandButton
Private mOwner As cRecordingControls
Private mCommand As String

Public Sub Connect(ByVal button As MSForms.CommandButton, ByVal owner As cRecordingControls, ByVal command As String)
    Set mButton = button: Set mOwner = owner: mCommand = command
End Sub

Public Sub Disconnect()
    Set mButton = Nothing: Set mOwner = Nothing
End Sub

Private Sub mButton_Click()
    mOwner.Execute mCommand
End Sub
