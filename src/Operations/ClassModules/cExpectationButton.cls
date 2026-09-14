VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cExpectationButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mButton As MSForms.CommandButton
Private mEditor As frmActionPathExpectation
Private mCommand As String

Public Sub Connect(ByVal button As MSForms.CommandButton, ByVal editor As frmActionPathExpectation, ByVal command As String)
    Set mButton = button: Set mEditor = editor: mCommand = command
End Sub

Public Sub Disconnect()
    Set mButton = Nothing: Set mEditor = Nothing
End Sub

Private Sub mButton_Click()
    mEditor.ChangeDraft mCommand
End Sub
