Attribute VB_Name = "modGuideReaderClose"
Option Explicit

' Native QueryClose must finish before the owner unloads the same MSForms window.
Private mOwner As frmActionPaths
Private mReader As frmActionPathLibrary
Private mPickers As Collection
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

Public Function SchedulePicker(ByVal picker As frmGuideActionPicker) As Boolean
    If Not RequestFlush() Then Exit Function
    If mPickers Is Nothing Then Set mPickers = New Collection
    mPickers.Add picker
    SchedulePicker = True
End Function

Private Function RequestFlush() As Boolean
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
    Dim pickers As Collection, picker As frmGuideActionPicker
    If Not mPending Then Exit Sub
    Set owner = mOwner: Set reader = mReader
    Set pickers = mPickers: Set mPickers = Nothing
    Set mOwner = Nothing: Set mReader = Nothing
    mPending = False
    If modOperationsFormLifetime.IsLoaded(owner) Then owner.ClosePublishedGuidesIfCurrent reader
    If Not pickers Is Nothing Then
        For Each picker In pickers
            If modOperationsFormLifetime.IsLoaded(picker) Then Unload picker
        Next picker
    End If
End Sub
