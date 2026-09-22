Attribute VB_Name = "modGuideActionPicker"
Option Explicit

' Declared primitive/serialized boundary; original bodies and selection stay in Core.
Public Function CanOpen(ByVal context As String, ByRef notice As String) As Boolean
    CanOpen = modCuratedGuideSource.CanOpen(context, notice)
End Function

Public Function OpenSelection(ByVal context As String, ByRef token As String, ByRef notice As String) As Boolean
    OpenSelection = modCuratedGuideSource.OpenSelection(context, token, notice)
End Function

Public Function ReadRows(ByVal context As String, ByVal token As String, ByVal query As String, ByRef rows As String, _
                         ByRef source As String, ByRef selectedCount As Long, ByRef notice As String) As Boolean
    ReadRows = modCuratedGuideSource.ReadRows(context, token, query, rows, source, selectedCount, notice)
End Function

Public Function Choose(ByVal context As String, ByVal token As String, ByVal id As String, ByVal selected As Boolean, ByRef notice As String) As Boolean
    Choose = modCuratedGuideSource.Choose(context, token, id, selected, notice)
End Function

Public Sub CloseSelection(ByVal context As String, ByVal token As String)
    modCuratedGuideSource.CloseSelection context, token
End Sub
