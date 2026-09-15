VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEventDetail
   Caption         =   "Event Detail"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   StartUpPosition =   1
End
Attribute VB_Name = "frmEventDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mOwner As cEventDetailController
Private WithEvents mLines As MSForms.ListBox
Private WithEvents mClose As MSForms.CommandButton
Private mFields As MSForms.ListBox
Private mProfile As MSForms.Label
Private mStatus As MSForms.Label
Private mLayout As cOperationsAnchorManager
Private mLoading As Boolean
Private mResizeReady As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Event Detail": Me.Width = 820: Me.Height = 640
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 820, 640
    Set mProfile = MakeControl("Label", "lblDetailProfile", "", 12, 10, 788, 32, 7)
    MakeControl "Label", "lblDetailLines", "Contributing lines - select a line to inspect its fields", 12, 48, 788, 20, 7
    Set mLines = MakeControl("ListBox", "lstEventLines", "", 12, 72, 788, 104, 7)
    mLines.ColumnCount = 1
    MakeControl "Label", "lblDetailFields", "Permitted event and selected-line fields", 12, 188, 788, 20, 7
    Set mFields = MakeControl("ListBox", "lstEventFields", "", 12, 212, 788, 306, 15)
    mFields.ColumnCount = 2: mFields.ColumnWidths = "195 pt;565 pt": mFields.Locked = True
    Set mStatus = MakeControl("Label", "lblDetailStatus", "Read-only published evidence. No workflow action is executed.", 12, 530, 680, 42, 13)
    Set mClose = MakeControl("CommandButton", "btnClose", "Close", 714, 570, 86, 28, 12)
End Sub

Public Sub Bind(ByVal owner As cEventDetailController, ByVal profileStatus As String)
    Dim key As Variant, keys As Collection
    Set mOwner = owner: mLoading = True
    Set keys = owner.LineLabels()
    mLines.Clear: mFields.Clear
    For Each key In keys
        mLines.AddItem CStr(key)
    Next key
    mProfile.Caption = profileStatus
    mStatus.Caption = CStr(mLines.ListCount) & " contributing line(s). Read-only published evidence; quantities are not combined."
    mLoading = False
    If mLines.ListCount > 0 Then mLines.ListIndex = 0
    RefreshFields
End Sub

Public Sub RefreshFields()
    Dim fields As Variant, index As Long
    If mLoading Or mOwner Is Nothing Then Exit Sub
    fields = mOwner.Fields(mLines.ListIndex + 1)
    mFields.Clear
    If IsEmpty(fields) Then Exit Sub
    For index = LBound(fields, 1) To UBound(fields, 1)
        If Not IsEmpty(fields(index, 0)) Then
            mFields.AddItem CStr(fields(index, 0))
            mFields.List(mFields.ListCount - 1, 1) = CStr(fields(index, 1))
        End If
    Next index
End Sub

Public Sub ClearContent(ByVal notice As String)
    mLoading = True
    mLines.Clear: mFields.Clear: mProfile.Caption = ""
    mStatus.Caption = notice
    mLoading = False
End Sub

Private Sub mLines_Click()
    RefreshFields
End Sub

Private Sub mClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not mOwner Is Nothing Then mOwner.ReleaseForm Me
    Set mOwner = Nothing
End Sub

Private Sub UserForm_Activate()
    If Not mOwner Is Nothing Then
        If Not mOwner.ContextValid() Then Exit Sub
    End If
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeReady = True
    End If
    mLayout.ApplyAnchoredLayout
End Sub

Private Sub UserForm_Layout()
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
    If Not mOwner Is Nothing Then mOwner.ContextValid
End Sub

Private Function MakeControl(ByVal kind As String, ByVal name As String, ByVal caption As String, ByVal x As Single, _
                             ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal anchors As Long) As Object
    Dim control As Object
    Set control = Me.Controls.Add("Forms." & kind & ".1", name, True)
    control.Move x, y, width, height
    If kind = "Label" Or kind = "CommandButton" Then control.Caption = caption
    mLayout.RegisterControl control, anchors
    Set MakeControl = control
End Function
