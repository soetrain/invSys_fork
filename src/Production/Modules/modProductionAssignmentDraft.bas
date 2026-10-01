Attribute VB_Name = "modProductionAssignmentDraft"
Option Explicit
Option Private Module

' Existing shared local alternatives, identified by requirement and ITEM_CODE.
Public Function AddAlternative(ByVal requirements As MSForms.ListBox, ByVal inventory As MSForms.ListBox, _
                               ByVal alternatives As Collection, ByRef report As String) As String
    Dim inventoryIndex As Long, requirementId As String, itemCode As String
    Dim alternative As Object, existing As Variant
    AddAlternative = "REJECTED"
    If requirements.ListIndex < 0 Then report = "Select an ingredient requirement first.": Exit Function
    inventoryIndex = inventory.ListIndex
    If inventoryIndex < 0 Then report = "Select a managed item first.": Exit Function
    requirementId = modProductionRecipeLists.ListText(requirements.List(requirements.ListIndex, 0))
    itemCode = modProductionRecipeLists.ListText(inventory.List(inventoryIndex, 6))
    If itemCode = "" Then report = "The selected inventory row has no managed item code.": Exit Function
    For Each existing In alternatives
        If StrComp(modProductionReusableDesigns.ReusableRecordText(existing, "RequirementId"), requirementId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(existing, "ITEM_CODE"), itemCode, vbTextCompare) = 0 Then
            report = "That acceptable item is already assigned."
            Exit Function
        End If
    Next existing
    Set alternative = CreateObject("Scripting.Dictionary")
    alternative.CompareMode = vbTextCompare: alternative("RecordType") = "ALTERNATIVE"
    alternative("RequirementId") = requirementId: alternative("ITEM_CODE") = itemCode
    alternatives.Add alternative
    report = "Added acceptable managed item " & itemCode & "."
    AddAlternative = "STAGED"
End Function

Public Function RemoveAlternative(ByVal allowed As MSForms.ListBox, ByVal alternatives As Collection, _
                                  ByRef refresh As Boolean) As String
    Dim index As Long, requirementId As String, itemCode As String, i As Long
    RemoveAlternative = "REJECTED": refresh = False
    index = allowed.ListIndex
    If index < 0 Then Exit Function
    requirementId = modProductionRecipeLists.ListText(allowed.List(index, 0))
    itemCode = modProductionRecipeLists.ListText(allowed.List(index, 6))
    For i = alternatives.Count To 1 Step -1
        If StrComp(modProductionReusableDesigns.ReusableRecordText(alternatives(i), "RequirementId"), requirementId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(alternatives(i), "ITEM_CODE"), itemCode, vbTextCompare) = 0 Then
            alternatives.Remove i
            RemoveAlternative = "STAGED"
            Exit For
        End If
    Next i
    refresh = True
End Function

Public Sub RefreshAllowed(ByVal allowed As MSForms.ListBox, ByVal requirements As MSForms.ListBox, _
                          ByVal alternatives As Collection, ByVal inventory As MSForms.ListBox, ByVal inventoryRows As Variant)
    Dim alternative As Variant, requirementId As String, itemCode As String
    Dim itemName As String, itemUom As String, rowIndex As Long
    allowed.Clear
    If requirements.ListIndex >= 0 Then requirementId = modProductionRecipeLists.ListText(requirements.List(requirements.ListIndex, 0))
    If alternatives Is Nothing Then Exit Sub
    For Each alternative In alternatives
        If requirementId = "" Or StrComp(modProductionReusableDesigns.ReusableRecordText( _
                alternative, "RequirementId"), requirementId, vbTextCompare) = 0 Then
            allowed.AddItem modProductionReusableDesigns.ReusableRecordText(alternative, "RequirementId")
            rowIndex = allowed.ListCount - 1
            itemCode = modProductionReusableDesigns.ReusableRecordText(alternative, "ITEM_CODE")
            itemName = DisplayForCode(itemCode, itemUom, inventory, inventoryRows)
            If itemName = "" Then itemName = itemCode
            allowed.List(rowIndex, 1) = itemName: allowed.List(rowIndex, 2) = itemUom
            allowed.List(rowIndex, 3) = itemCode: allowed.List(rowIndex, 6) = itemCode
        End If
    Next alternative
End Sub

Private Function DisplayForCode(ByVal itemCode As String, ByRef itemUom As String, _
                                ByVal inventory As MSForms.ListBox, ByVal inventoryRows As Variant) As String
    Dim i As Long, r As Long
    itemCode = Trim$(itemCode): itemUom = ""
    If itemCode = "" Then Exit Function
    If Not inventory Is Nothing Then
        For i = 0 To inventory.ListCount - 1
            If StrComp(modProductionRecipeLists.ListText(inventory.List(i, 6)), itemCode, vbTextCompare) = 0 Then
                DisplayForCode = modProductionRecipeLists.ListText(inventory.List(i, 1))
                itemUom = modProductionRecipeLists.ListText(inventory.List(i, 2))
                Exit Function
            End If
        Next i
    End If
    If IsArray(inventoryRows) Then
        For r = LBound(inventoryRows, 1) To UBound(inventoryRows, 1)
            If StrComp(modProductionRecipeLists.ListText(inventoryRows(r, 7)), itemCode, vbTextCompare) = 0 Then
                DisplayForCode = modProductionRecipeLists.ListText(inventoryRows(r, 2))
                itemUom = modProductionRecipeLists.ListText(inventoryRows(r, 3))
                Exit Function
            End If
        Next r
    End If
End Function
