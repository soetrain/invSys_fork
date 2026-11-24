Attribute VB_Name = "modGuideReaderClose"
Option Explicit

' Native QueryClose must finish before the owner unloads the same MSForms window.
Private mPending As Boolean

Public Function RequestFlush() As Boolean
    On Error GoTo Failed
    If Not mPending Then
        Application.OnTime EarliestTime:=DateAdd("s", 1, Now), _
            Procedure:="'" & ThisWorkbook.Name & "'!modGuideReaderClose.Flush", Schedule:=True
        mPending = True
    End If
    RequestFlush = True
Failed:
End Function

Public Sub Flush()
    Dim reader As frmActionPathLibrary
    Dim index As Long, instance As Object, picker As frmGuideActionPicker
    If Not mPending Then Exit Sub
    mPending = False
    ' Reopening can already have unloaded the old window. Each surviving form
    ' carries its own close request. Unloading can also remove its child forms.
    For index = VBA.UserForms.Count - 1 To 0 Step -1
        If index < VBA.UserForms.Count Then
            Set instance = VBA.UserForms(index)
            If TypeOf instance Is frmGuideActionPicker Then
                Set picker = instance
                picker.FinishNativeClose
            ElseIf TypeOf instance Is frmActionPathLibrary Then
                Set reader = instance
                reader.FinishNativeClose
            End If
        End If
    Next index
End Sub
