Attribute VB_Name = "modExpectationEditor"
Option Explicit

' One Operations-owned editor with explicit recording, run and guide scopes.
Private mEditor As frmActionPathExpectation
Private mLibrary As frmActionPaths
Private mContext As String
Private mPathId As String
Private mGuideDraftId As String
Private mGuide As frmActionPathGuide

Public Sub OpenForRecording(ByVal context As String)
    Present context, "", Nothing, "", Nothing
End Sub

Public Sub OpenForRun(ByVal context As String, ByVal pathId As String, ByVal library As frmActionPaths)
    If pathId <> "" Then Present context, pathId, library, "", Nothing
End Sub

Public Sub OpenForGuide(ByVal context As String, ByVal draftId As String, ByVal guide As frmActionPathGuide)
    If draftId <> "" Then Present context, "", Nothing, draftId, guide
End Sub

Private Sub Present(ByVal context As String, ByVal pathId As String, ByVal library As frmActionPaths, _
                    ByVal guideDraftId As String, ByVal guide As frmActionPathGuide)
    If Not mEditor Is Nothing Then
        If mContext = context And mPathId = pathId And mGuideDraftId = guideDraftId And mEditor.Visible Then
            If mEditor.ValidateBinding() Then Exit Sub
        End If
        CloseEditor
    End If
    Set mEditor = New frmActionPathExpectation
    mContext = context: mPathId = pathId: Set mLibrary = library
    mGuideDraftId = guideDraftId: Set mGuide = guide
    If mEditor.BindContext(context, pathId, guideDraftId) Then mEditor.Show vbModeless Else CloseEditor
End Sub

Public Sub Staged()
    If Not mLibrary Is Nothing Then mLibrary.RefreshExpectation
    If Not mGuide Is Nothing Then mGuide.RefreshExpectation
End Sub

Public Sub CloseRecording(ByVal context As String)
    If mContext = context And mPathId = "" And mGuideDraftId = "" Then CloseEditor
End Sub

Public Sub CloseLibrary(ByVal library As frmActionPaths)
    If mLibrary Is Nothing Then Exit Sub
    If mLibrary Is library Then CloseEditor
End Sub

Public Sub CloseGuide(ByVal guide As frmActionPathGuide)
    If mGuide Is Nothing Then Exit Sub
    If mGuide Is guide Then CloseEditor
End Sub

Private Sub CloseEditor()
    Set mLibrary = Nothing: Set mGuide = Nothing
    If Not mEditor Is Nothing Then
        mEditor.ReleaseDraft
        Unload mEditor: Set mEditor = Nothing
    End If
    mContext = "": mPathId = "": mGuideDraftId = ""
End Sub
