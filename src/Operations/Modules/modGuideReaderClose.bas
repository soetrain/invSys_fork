Attribute VB_Name = "modGuideReaderClose"
Option Explicit

' Native QueryClose must finish before the owner unloads the same MSForms window.
Private mOwner As frmActionPaths
Private mReader As frmActionPathLibrary
Private mPending As Boolean

Public Function Schedule(ByVal owner As frmActionPaths, ByVal reader As frmActionPathLibrary) As Boolean
    If owner Is Nothing Or reader Is Nothing Then Exit Function
    If Not mReader Is Nothing Then
        Schedule = (mOwner Is owner And mReader Is reader)
        Exit Function
    End If
    If Not RequestFlush() Then Exit Function
    Set mOwner = owner: Set mReader = reader
    Schedule = True
End Function

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
    Dim owner As frmActionPaths, reader As frmActionPathLibrary
    Dim index As Long, instance As Object, picker As frmGuideActionPicker
    If Not mPending Then Exit Sub
    Set owner = mOwner: Set reader = mReader
    Set mOwner = Nothing: Set mReader = Nothing
    mPending = False
    If modOperationsFormLifetime.IsLoaded(owner) Then owner.ClosePublishedGuidesIfCurrent reader
    ' Reopening can already have unloaded the old picker. Only its live instance
    ' can carry the native-close flag. Unloading can also remove its child forms.
    For index = VBA.UserForms.Count - 1 To 0 Step -1
        If index < VBA.UserForms.Count Then
            Set instance = VBA.UserForms(index)
            If TypeOf instance Is frmGuideActionPicker Then
                Set picker = instance
                picker.FinishNativeClose
            End If
        End If
    Next index
End Sub
