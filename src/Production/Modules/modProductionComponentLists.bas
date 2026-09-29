Attribute VB_Name = "modProductionComponentLists"
Option Explicit
Option Private Module

Public Function FindEditorRow(ByVal listControl As MSForms.ListBox, ByVal componentId As String) As Long
    Dim rowIndex As Long, value As Variant
    FindEditorRow = -1
    If listControl Is Nothing Then Exit Function
    componentId = UCase$(Trim$(componentId))
    If componentId = "" Then Exit Function
    For rowIndex = 0 To listControl.ListCount - 1
        value = listControl.List(rowIndex, 0)
        If Not IsError(value) And Not IsNull(value) And Not IsEmpty(value) Then
            If StrComp(UCase$(Trim$(CStr(value))), componentId, vbTextCompare) = 0 Then
                FindEditorRow = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

' Component callers pass the seven/ten owned fields. Recipe callers retain their
' existing declared-column behavior and instruction-ordinal normalization.
Public Function MoveRow(ByVal listControl As MSForms.ListBox, ByVal direction As Long, _
                        ByVal instructions As MSForms.ListBox, Optional ByVal fields As Long = 0) As Boolean
    Dim sourceIndex As Long, targetIndex As Long, columnIndex As Long, index As Long
    Dim tempValue As Variant
    If listControl Is Nothing Then Exit Function
    sourceIndex = listControl.ListIndex
    targetIndex = sourceIndex + direction
    If sourceIndex < 0 Or targetIndex < 0 Or targetIndex >= listControl.ListCount Then Exit Function
    If fields = 0 Then fields = listControl.ColumnCount
    For columnIndex = 0 To fields - 1
        tempValue = listControl.List(sourceIndex, columnIndex)
        listControl.List(sourceIndex, columnIndex) = listControl.List(targetIndex, columnIndex)
        listControl.List(targetIndex, columnIndex) = tempValue
    Next columnIndex
    listControl.ListIndex = targetIndex
    If Not instructions Is Nothing Then
        For index = 0 To instructions.ListCount - 1
            instructions.List(index, 0) = CStr(index + 1)
        Next index
    End If
    MoveRow = True
End Function
