Attribute VB_Name = "modProductionRecipeLists"
Option Explicit
Option Private Module

' Recipe node identity comparisons preserve the existing case-insensitive,
' non-trimming local-list rules. These lists never own saved definitions.
Public Function NodeIndex(ByVal nodes As MSForms.ListBox, ByVal nodeId As String) As Long
    Dim i As Long
    NodeIndex = -1
    For i = 0 To nodes.ListCount - 1
        If StrComp(ListText(nodes.List(i, 0)), nodeId, vbTextCompare) = 0 Then NodeIndex = i: Exit Function
    Next i
End Function

Public Function ListText(ByVal value As Variant) As String
    If Not IsError(value) Then
        If Not IsNull(value) And Not IsEmpty(value) Then ListText = CStr(value)
    End If
End Function

Public Function AddNode(ByVal nodes As MSForms.ListBox, ByVal released As MSForms.ListBox) As String
    Dim selected As Long, candidate As Long, row As Long, nodeId As String
    selected = released.ListIndex
    If selected < 0 Then Exit Function
    candidate = nodes.ListCount + 1
    Do While NodeIndex(nodes, "N" & CStr(candidate)) >= 0
        candidate = candidate + 1
    Loop
    nodeId = "N" & CStr(candidate)
    nodes.AddItem nodeId
    row = nodes.ListCount - 1
    nodes.List(row, 1) = ListText(released.List(selected, 0))
    nodes.List(row, 2) = ListText(released.List(selected, 1))
    nodes.List(row, 3) = ListText(released.List(selected, 2))
    nodes.List(row, 4) = CStr(row + 1)
    nodes.ListIndex = row
    AddNode = nodeId
End Function

Public Function RemoveNode(ByVal nodes As MSForms.ListBox, ByVal connections As MSForms.ListBox) As Boolean
    Dim nodeId As String, i As Long
    If nodes.ListIndex < 0 Then Exit Function
    nodeId = ListText(nodes.List(nodes.ListIndex, 0))
    For i = connections.ListCount - 1 To 0 Step -1
        If StrComp(ListText(connections.List(i, 0)), nodeId, vbTextCompare) = 0 _
           Or StrComp(ListText(connections.List(i, 2)), nodeId, vbTextCompare) = 0 Then
            connections.RemoveItem i
        End If
    Next i
    nodes.RemoveItem nodes.ListIndex
    RemoveNode = True
End Function

Public Function ConnectionIndex(ByVal display As MSForms.ListBox, ByVal connections As MSForms.ListBox) As Long
    Dim idx As Long, selected As Long
    selected = display.ListIndex
    If selected >= 0 Then idx = CLng(Val(ListText(display.List(selected, 7)))) Else idx = -1
    If idx < 0 Then idx = connections.ListIndex
    ConnectionIndex = -1
    If idx >= 0 And idx < connections.ListCount Then ConnectionIndex = idx
End Function
