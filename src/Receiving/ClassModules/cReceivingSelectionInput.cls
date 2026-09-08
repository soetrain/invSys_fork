VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cReceivingSelectionInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Form-local input provenance only: no key, button or selected value is retained.
Private WithEvents mList As MSForms.ListBox
Private WithEvents mChoice As MSForms.ComboBox
Private WithEvents mPage As MSForms.TabStrip
Private mArmed As Boolean
Private mDropdownOpen As Boolean

Public Sub Attach(ByVal control As Object)
    If TypeOf control Is MSForms.ListBox Then Set mList = control
    If TypeOf control Is MSForms.ComboBox Then Set mChoice = control
    If TypeOf control Is MSForms.TabStrip Then Set mPage = control
End Sub

Public Sub Detach()
    mArmed = False
    mDropdownOpen = False
    Set mList = Nothing: Set mChoice = Nothing: Set mPage = Nothing
End Sub

Public Function Consume() As Boolean
    Consume = mArmed
    mArmed = False
End Function

Private Function SelectionKey(ByVal key As Long) As Boolean
    Select Case key
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyHome, vbKeyEnd, vbKeyPageUp, vbKeyPageDown
            SelectionKey = True
    End Select
End Function

Private Sub mList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mArmed = (Button = 1)
End Sub
Private Sub mList_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mArmed = False
End Sub
Private Sub mList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mArmed = SelectionKey(KeyCode)
End Sub
Private Sub mList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mArmed = False
End Sub

Private Sub mChoice_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mArmed = (Button = 1)
End Sub
Private Sub mChoice_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' A dropdown commits after MouseUp, then signals DropButtonClick as it closes.
    ' Opening/cancelling the dropdown clears the token even when no Click follows.
    If Not mDropdownOpen Then mArmed = False
End Sub
Private Sub mChoice_DropButtonClick()
    mDropdownOpen = Not mDropdownOpen
    mArmed = False
End Sub
Private Sub mChoice_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mArmed = SelectionKey(KeyCode)
    If mDropdownOpen And (KeyCode = vbKeyReturn Or KeyCode = vbKeySpace) Then mArmed = True
End Sub
Private Sub mChoice_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mArmed = False
End Sub

Private Sub mPage_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mArmed = (Button = 1 And Index >= 0)
End Sub
Private Sub mPage_MouseUp(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mArmed = False
End Sub
Private Sub mPage_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mArmed = SelectionKey(KeyCode)
    If KeyCode = vbKeyTab And (Shift And 2) <> 0 Then mArmed = True
End Sub
Private Sub mPage_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mArmed = False
End Sub
