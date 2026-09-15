Attribute VB_Name = "modGuideEditor"
Option Explicit

Private mEditor As frmActionPathGuide
Private mLibrary As frmActionPaths
Private mContext As String
Private mPathId As String

Public Sub OpenForRun(ByVal context As String, ByVal pathId As String, ByVal library As frmActionPaths)
    If Not mEditor Is Nothing Then
        If mContext = context And mPathId = pathId And mEditor.Visible Then
            If mEditor.ValidateBinding() Then Exit Sub
        End If
        CloseEditor
    End If
    If Not modActionGuideDraft.CanCreate(context, pathId) Then Exit Sub
    Set mEditor = New frmActionPathGuide
    Set mLibrary = library: mContext = context: mPathId = pathId
    If mEditor.BindContext(context, pathId) Then mEditor.Show vbModeless Else CloseEditor
End Sub

Public Sub CancelEditor(ByVal editor As frmActionPathGuide)
    If mEditor Is Nothing Then Exit Sub
    If mEditor Is editor Then CloseEditor
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
