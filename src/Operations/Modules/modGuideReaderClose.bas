Attribute VB_Name = "modGuideReaderClose"
Option Explicit

' Native QueryClose must finish before the owner unloads the same MSForms window.
Private mOwner As frmActionPaths
Private mReader As frmActionPathLibrary
Private mPending As Boolean
Private mNextRun As Date

Public Function Schedule(ByVal owner As frmActionPaths, ByVal reader As frmActionPathLibrary) As Boolean
    If owner Is Nothing Or reader Is Nothing Then Exit Function
    If mPending Then
        Schedule = (mOwner Is owner And mReader Is reader)
        Exit Function
    End If
    On Error GoTo Failed
    Set mOwner = owner: Set mReader = reader
    mNextRun = DateAdd("s", 1, Now)
    Application.OnTime EarliestTime:=mNextRun, _
        Procedure:="'" & ThisWorkbook.Name & "'!modGuideReaderClose.Flush", Schedule:=True
    mPending = True
    Schedule = True
    Exit Function
Failed:
    Set mOwner = Nothing: Set mReader = Nothing
    mNextRun = 0: mPending = False
End Function

Public Sub Flush()
    Dim owner As frmActionPaths, reader As frmActionPathLibrary
    If Not mPending Then Exit Sub
    Set owner = mOwner: Set reader = mReader
    Set mOwner = Nothing: Set mReader = Nothing
    mNextRun = 0: mPending = False
    If Not owner Is Nothing And Not reader Is Nothing Then owner.ClosePublishedGuidesIfCurrent reader
End Sub
