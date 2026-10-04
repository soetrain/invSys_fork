Attribute VB_Name = "modExecutionInputs"
Option Explicit
Option Private Module

' Registered B0 adapter inputs. This registry never dispatches a workflow.
Public Function Supported(ByVal controlId As String) As Boolean
    Select Case controlId
        Case "RECEIVING_OPEN", "RECEIVING_REFRESH", "RECEIVING_CLEAR", "RECEIVING_SELECT_ITEM", "RECEIVING_ADD_SELECTED", "RECEIVING_CONFIRM_WRITES"
            Supported = True
    End Select
End Function

Public Function Names(ByVal controlId As String) As String
    Select Case controlId
        Case "RECEIVING_SELECT_ITEM": Names = "SourceEntity"
        Case "RECEIVING_ADD_SELECTED": Names = "Reference|Quantity|Location|LotNumber|Condition"
    End Select
End Function

Public Function InputType(ByVal name As String) As String
    Select Case name
        Case "SourceEntity": InputType = "InventoryEntity"
        Case "Quantity": InputType = "PositiveDecimalText"
        Case "Reference", "Location": InputType = "RequiredText"
        Case "LotNumber": InputType = "OptionalText"
        Case "Condition": InputType = "ReceivingCondition"
    End Select
End Function

Public Function Create(ByVal controlId As String) As Collection
    Dim result As New Collection, name As Variant, item As Object, binding As Object
    If Names(controlId) <> "" Then
        For Each name In Split(Names(controlId), "|")
            Set item = CreateObject("Scripting.Dictionary"): Set binding = CreateObject("Scripting.Dictionary")
            item.Add "Name", CStr(name): item.Add "Type", InputType(CStr(name))
            If name = "SourceEntity" Then
                binding.Add "Kind", "Prompt": binding.Add "PromptId", "RECEIVING_SOURCE_ENTITY"
            Else
                binding.Add "Kind", "Literal": binding.Add "Value", ""
            End If
            item.Add "Binding", binding: result.Add item
        Next name
    End If
    Set Create = result
End Function

Public Function ValidBinding(ByVal name As String, ByVal binding As Object) As Boolean
    Dim value As String, i As Long, code As Long, parts As Variant
    On Error GoTo Invalid
    If name = "SourceEntity" Then
        If Not modEvaluationModel.HasFields(binding, "Kind|PromptId") Then Exit Function
        If VarType(binding("Kind")) <> vbString Or VarType(binding("PromptId")) <> vbString Then Exit Function
        ValidBinding = (binding("Kind") = "Prompt" And binding("PromptId") = "RECEIVING_SOURCE_ENTITY")
        Exit Function
    End If
    If Not modEvaluationModel.HasFields(binding, "Kind|Value") Then Exit Function
    If VarType(binding("Kind")) <> vbString Or VarType(binding("Value")) <> vbString Then Exit Function
    If binding("Kind") <> "Literal" Then Exit Function
    value = CStr(binding("Value"))
    If Len(value) > 128 Then Exit Function
    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1)) And &HFFFF&
        If code < 32 Or code = 127 Then Exit Function
    Next i
    Select Case name
        Case "Reference", "Location": ValidBinding = (Trim$(value) <> "")
        Case "LotNumber": ValidBinding = True
        Case "Condition"
            Select Case value
                Case "GOOD", "BAD", "DAMAGED", "EXPIRED", "REJECTED": ValidBinding = True
            End Select
        Case "Quantity"
            If Len(value) = 0 Or Len(value) > 32 Or value Like "*[!0-9.]*" Then Exit Function
            parts = Split(value, ".")
            If UBound(parts) > 1 Or parts(0) = "" Then Exit Function
            If UBound(parts) = 1 Then
                If parts(1) = "" Then Exit Function
            End If
            ValidBinding = (Val(value) > 0 And Val(value) <= 1000000000000#)
    End Select
Invalid:
End Function

Public Function Validate(ByVal controlId As String, ByVal inputs As Collection) As Boolean
    Dim namesList As Variant, item As Object, index As Long, expectedCount As Long
    On Error GoTo Invalid
    If Not Supported(controlId) Then Exit Function
    If Names(controlId) <> "" Then
        namesList = Split(Names(controlId), "|"): expectedCount = UBound(namesList) + 1
    End If
    If inputs.Count <> expectedCount Then Exit Function
    For Each item In inputs
        If Not modEvaluationModel.HasFields(item, "Name|Type|Binding") Then Exit Function
        If VarType(item("Name")) <> vbString Or VarType(item("Type")) <> vbString Then Exit Function
        If item("Name") <> CStr(namesList(index)) Or item("Type") <> InputType(CStr(item("Name"))) Then Exit Function
        If Not ValidBinding(CStr(item("Name")), item("Binding")) Then Exit Function
        index = index + 1
    Next item
    Validate = True
Invalid:
End Function
