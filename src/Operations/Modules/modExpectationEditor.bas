Attribute VB_Name = "modExpectationEditor"
Option Explicit

' One Operations-owned editor shared by recording and selected-run analysis.
Private mEditor As frmActionPathExpectation
Private mLibrary As frmActionPaths
Private mContext As String
Private mPathId As String

Public Sub OpenForRecording(ByVal context As String)
    Present context, "", Nothing
End Sub

Public Sub OpenForRun(ByVal context As String, ByVal pathId As String, ByVal library As frmActionPaths)
    If pathId <> "" Then Present context, pathId, library
End Sub

Private Sub Present(ByVal context As String, ByVal pathId As String, ByVal library As frmActionPaths)
    If Not mEditor Is Nothing Then
        If mContext = context And mPathId = pathId And mEditor.Visible Then
            If mEditor.ValidateBinding() Then Exit Sub
        End If
        CloseEditor
    End If
    Set mEditor = New frmActionPathExpectation
    mContext = context: mPathId = pathId: Set mLibrary = library
    If mEditor.BindContext(context, pathId) Then mEditor.Show vbModeless Else CloseEditor
End Sub

Public Sub Staged()
    If Not mLibrary Is Nothing Then mLibrary.RefreshExpectation
End Sub

Public Sub CloseRecording(ByVal context As String)
    If mContext = context And mPathId = "" Then CloseEditor
End Sub

Public Sub CloseLibrary(ByVal library As frmActionPaths)
    If mLibrary Is Nothing Then Exit Sub
    If mLibrary Is library Then CloseEditor
End Sub

Private Sub CloseEditor()
    Set mLibrary = Nothing
    If Not mEditor Is Nothing Then
        mEditor.ReleaseDraft
        Unload mEditor: Set mEditor = Nothing
    End If
    mContext = "": mPathId = ""
End Sub
