Attribute VB_Name = "modGuideEditor"
Option Explicit

Private mEditor As frmActionPathGuide
Private mLibrary As frmActionPaths
Private mContext As String
Private mPathId As String
Private mPublishedReader As frmActionPathLibrary
Private mPublishedKey As String
Private mActionPicker As frmGuideActionPicker
Private mCurationId As String

Public Sub OpenForRun(ByVal context As String, ByVal pathId As String, ByVal library As frmActionPaths)
    If Not mEditor Is Nothing Then
        If mPublishedReader Is Nothing And mActionPicker Is Nothing And mContext = context And mPathId = pathId And mEditor.Visible Then
            If mEditor.ValidateBinding() Then Exit Sub
        End If
        CloseEditor
    End If
    If Not modActionGuideDraft.CanCreate(context, pathId) Then Exit Sub
    Set mEditor = New frmActionPathGuide
    Set mLibrary = library: mContext = context: mPathId = pathId
    If mEditor.BindContext(context, pathId) Then mEditor.Show vbModeless Else CloseEditor
End Sub

Public Sub OpenForSelectedActions(ByVal context As String, ByVal token As String, ByVal picker As frmGuideActionPicker)
    If Not mEditor Is Nothing Then
        If mActionPicker Is picker Then
            If mContext = context And mCurationId = token And mEditor.Visible Then
                If mEditor.ValidateBinding() Then Exit Sub
            End If
        End If
        CloseEditor
    End If
    Set mEditor = New frmActionPathGuide
    Set mActionPicker = picker: mContext = context: mCurationId = token
    If mEditor.BindSelectedActions(context, token) Then mEditor.Show vbModeless Else CloseEditor
End Sub

Public Sub CloseActionPicker(ByVal picker As frmGuideActionPicker)
    If mActionPicker Is Nothing Then Exit Sub
    If mActionPicker Is picker Then CloseEditor
End Sub

Public Sub OpenForPublishedGuide(ByVal context As String, ByVal key As String, ByVal reader As frmActionPathLibrary)
    If Not mEditor Is Nothing Then
        If mPublishedReader Is reader Then
            If mContext = context And mPublishedKey = key And mEditor.Visible Then
                If mEditor.ValidateBinding() Then Exit Sub
            End If
        End If
        CloseEditor
    End If
    If Not modActionGuideDraft.CanEditPublished(context, key) Then Exit Sub
    Set mEditor = New frmActionPathGuide
    Set mPublishedReader = reader: mContext = context: mPublishedKey = key
    If mEditor.BindPublishedContext(context, key) Then mEditor.Show vbModeless Else CloseEditor
End Sub

Public Sub ValidatePublishedSelection(ByVal reader As frmActionPathLibrary, ByVal key As String)
    If mPublishedReader Is Nothing Then Exit Sub
    If mPublishedReader Is reader Then
        If key <> mPublishedKey Then CloseEditor
    End If
End Sub

Public Sub ClosePublishedReader(ByVal reader As frmActionPathLibrary)
    If mPublishedReader Is Nothing Then Exit Sub
    If mPublishedReader Is reader Then CloseEditor
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
    Set mLibrary = Nothing: Set mPublishedReader = Nothing: Set mActionPicker = Nothing
    If Not mEditor Is Nothing Then
        mEditor.ReleaseDraft
        Unload mEditor: Set mEditor = Nothing
    End If
    mContext = "": mPathId = "": mPublishedKey = "": mCurationId = ""
End Sub
