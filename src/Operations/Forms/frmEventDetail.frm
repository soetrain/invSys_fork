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
Private mMultiline As MSForms.Frame
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
    Set mFields = MakeControl("ListBox", "lstEventFields", "", 12, 212, 788, 176, 15)
    mFields.ColumnCount = 2: mFields.ColumnWidths = "195 pt;565 pt": mFields.Locked = False
    MakeControl "Label", "lblDetailMultiline", "Multiline fields for selected line", 12, 398, 788, 20, 13
    Set mMultiline = MakeControl("Frame", "fraDetailMultiline", "", 12, 420, 788, 98, 13)
    mMultiline.Caption = "": mMultiline.ScrollBars = fmScrollBarsBoth
    mMultiline.KeepScrollBarsVisible = fmScrollBarsNone
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
    ClearMultiline
    If IsEmpty(fields) Then Exit Sub
    For index = LBound(fields, 1) To UBound(fields, 1)
        If Not IsEmpty(fields(index, 0)) Then
            mFields.AddItem CStr(fields(index, 0))
            mFields.List(mFields.ListCount - 1, 1) = CStr(fields(index, 1))
        End If
    Next index
    FitFieldColumns
    RenderMultiline
End Sub

Private Sub ClearMultiline()
    Dim index As Long
    If mMultiline Is Nothing Then Exit Sub
    With mMultiline
        .ScrollTop = 0: .ScrollLeft = 0
        For index = .Controls.Count - 1 To 0 Step -1
            .Controls.Remove .Controls(index).Name
        Next index
        .ScrollWidth = .InsideWidth: .ScrollHeight = .InsideHeight
    End With
End Sub

Private Sub RenderMultiline()
    Dim row As Long, count As Long, top As Single, value As String, label As MSForms.Label
    top = 8
    For row = 0 To mFields.ListCount - 1
        value = CStr(mFields.List(row, 1))
        If InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then
            count = count + 1
            Set label = AddMultilineLabel("lblMultilineCaption" & CStr(count), CStr(mFields.List(row, 0)), top, True)
            top = label.Top + label.Height + 4
            Set label = AddMultilineLabel("lblMultilineValue" & CStr(count), value, top, False)
            top = label.Top + label.Height + 12
        End If
    Next row
    If count = 0 Then Set label = AddMultilineLabel("lblMultilineEmpty", "No multiline fields for this line.", top, False)
End Sub

Private Function AddMultilineLabel(ByVal name As String, ByVal text As String, ByVal top As Single, ByVal bold As Boolean) As MSForms.Label
    Dim label As MSForms.Label
    Set label = mMultiline.Controls.Add("Forms.Label.1", name, True)
    label.Font.Name = mFields.Font.Name: label.Font.Size = mFields.Font.Size: label.Font.Bold = bold
    label.Left = 8: label.Top = top
    label.Accelerator = "": label.WordWrap = False: label.AutoSize = True: label.Caption = text
    If label.Left + label.Width + 24 > mMultiline.ScrollWidth Then mMultiline.ScrollWidth = label.Left + label.Width + 24
    If label.Top + label.Height + 24 > mMultiline.ScrollHeight Then mMultiline.ScrollHeight = label.Top + label.Height + 24
    Set AddMultilineLabel = label
End Function

Private Sub FitFieldColumns()
    Dim measure As MSForms.Label, widths As Variant, row As Long, column As Long
    widths = Array(195!, 565!)
    Set measure = Me.Controls.Add("Forms.Label.1", "lblMeasureFieldText", False)
    measure.Font.Name = mFields.Font.Name: measure.Font.Size = mFields.Font.Size
    measure.Font.Bold = mFields.Font.Bold: measure.Font.Italic = mFields.Font.Italic
    measure.WordWrap = False: measure.AutoSize = True
    For row = 0 To mFields.ListCount - 1
        For column = 0 To 1
            measure.Caption = CStr(mFields.List(row, column))
            If measure.Width + 8 > widths(column) Then widths(column) = measure.Width + 8
        Next column
    Next row
    Me.Controls.Remove "lblMeasureFieldText"
    mFields.ColumnWidths = CStr(widths(0)) & " pt;" & CStr(widths(1)) & " pt"
End Sub

Public Sub ClearContent(ByVal notice As String)
    mLoading = True
    mLines.Clear: mFields.Clear: mProfile.Caption = ""
    ClearMultiline
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
