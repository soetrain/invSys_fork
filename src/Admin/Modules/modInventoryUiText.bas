Attribute VB_Name = "modInventoryUiText"
Option Explicit
Option Private Module

Public Function SafeFormText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Then Exit Function
    If IsNull(valueIn) Then Exit Function
    If IsEmpty(valueIn) Then Exit Function
    SafeFormText = Trim$(CStr(valueIn))
End Function

Public Function MaxLongAdminForm(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLongAdminForm = leftValue
    Else
        MaxLongAdminForm = rightValue
    End If
End Function

Public Function MinLongAdminForm(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue <= rightValue Then
        MinLongAdminForm = leftValue
    Else
        MinLongAdminForm = rightValue
    End If
End Function

Public Function AdminNormalizeSearchText(ByVal valueIn As String) As String
    Dim textOut As String

    textOut = Trim$(CStr(valueIn))
    textOut = Replace(textOut, vbTab, " ")
    Do While InStr(1, textOut, "  ", vbBinaryCompare) > 0
        textOut = Replace(textOut, "  ", " ")
    Loop
    AdminNormalizeSearchText = LCase$(textOut)
End Function

Public Function AdminTextMatchesSearch(ByVal candidate As String, ByVal searchTerm As String) As Boolean
    If searchTerm = "" Then
        AdminTextMatchesSearch = True
    Else
        AdminTextMatchesSearch = (InStr(1, AdminNormalizeSearchText(candidate), searchTerm, vbTextCompare) > 0)
    End If
End Function

Public Function AdminAnyTextMatchesSearch(ByVal searchTerm As String, ParamArray candidates() As Variant) As Boolean
    Dim i As Long

    If searchTerm = "" Then
        AdminAnyTextMatchesSearch = True
        Exit Function
    End If
    For i = LBound(candidates) To UBound(candidates)
        If AdminTextMatchesSearch(CStr(candidates(i)), searchTerm) Then
            AdminAnyTextMatchesSearch = True
            Exit Function
        End If
    Next i
End Function


Public Function NormalizeCustomFieldName(ByVal fieldName As String) As String
    fieldName = Replace(fieldName, vbCr, " ")
    fieldName = Replace(fieldName, vbLf, " ")
    fieldName = Replace(fieldName, vbTab, " ")
    fieldName = Replace(fieldName, "[", "(")
    fieldName = Replace(fieldName, "]", ")")
    Do While InStr(1, fieldName, "  ", vbBinaryCompare) > 0
        fieldName = Replace(fieldName, "  ", " ")
    Loop
    fieldName = Trim$(fieldName)
    If Len(fieldName) > 48 Then fieldName = Left$(fieldName, 48)
    NormalizeCustomFieldName = fieldName
End Function

Public Function IsReservedCustomField(ByVal fieldName As String) As Boolean
    Select Case UCase$(Trim$(fieldName))
        Case "SYSTEM_KEY", "SYSTEMKEY", "SYSTEM KEY", "SYSTEM-KEY", "CONDITION", "SKU", "ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "QTY", "TOTAL INV", _
             "QTYAVAILABLE", "DESCRIPTION", "VENDOR(S)", "VENDOR_CODE", "CATEGORY", _
             "EXTERNAL_CODE", "IMAGE_PATH", "TRACK_QTY", "ITEM_KIND", "NOTE", "IOTYPE"
            IsReservedCustomField = True
    End Select
End Function

Public Function CatalogItemTextAdmin(ByVal item As Variant, ByVal fieldName As String) As String
    On Error Resume Next
    CatalogItemTextAdmin = Trim$(CStr(item(fieldName)))
    On Error GoTo 0
End Function

Public Function CatalogItemIsNonCounted(ByVal item As Variant) As Boolean
    Dim trackQty As String
    Dim itemKind As String

    trackQty = UCase$(CatalogField(item, "TRACK_QTY"))
    itemKind = UCase$(CatalogField(item, "ITEM_KIND"))
    CatalogItemIsNonCounted = (trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
                               Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED")
End Function

Public Function FindCatalogItemBySku(ByVal catalogItems As Object, ByVal sku As String) As Object
    Dim item As Variant

    If catalogItems Is Nothing Then Exit Function
    For Each item In catalogItems
        If StrComp(CatalogField(item, "SKU"), sku, vbTextCompare) = 0 Then
            Set FindCatalogItemBySku = item
            Exit Function
        End If
    Next item
End Function

Public Function CatalogField(ByVal item As Variant, ByVal fieldName As String) As String
    On Error Resume Next
    CatalogField = Trim$(CStr(item(fieldName)))
    On Error GoTo 0
End Function

Public Function QuantityModeIsNonCounted(ByVal valueText As String) As Boolean
    Select Case UCase$(Trim$(valueText))
        Case "UTILITY", "SERVICE", "NOT COUNTED", "NON-COUNTED", "NONCOUNTED", "UNTRACKED"
            QuantityModeIsNonCounted = True
    End Select
End Function

Public Function NonCountedItemKind(ByVal valueText As String) As String
    Select Case UCase$(valueText)
        Case "UTILITY"
            NonCountedItemKind = "UTILITY"
        Case "SERVICE"
            NonCountedItemKind = "SERVICE"
        Case Else
            NonCountedItemKind = "NON_COUNTED"
    End Select
End Function

Public Sub AddDimensionOption(ByVal targetCombo As MSForms.ComboBox, _
                               ByVal seen As Object, _
                               ByVal valueText As String)
    valueText = Trim$(valueText)
    If valueText = "" Then Exit Sub
    If seen.Exists(valueText) Then Exit Sub
    seen.Add valueText, True
    targetCombo.AddItem valueText
End Sub

Public Function ComboContainsValue(ByVal targetCombo As MSForms.ComboBox, _
                                    ByVal valueText As String) As Boolean
    Dim i As Long

    If targetCombo Is Nothing Then Exit Function
    For i = 0 To targetCombo.ListCount - 1
        If StrComp(Trim$(CStr(targetCombo.List(i))), Trim$(valueText), vbTextCompare) = 0 Then
            ComboContainsValue = True
            Exit Function
        End If
    Next i
End Function

Public Function FindCustomFieldRow(ByVal fields As MSForms.ListBox, ByVal fieldName As String) As Long
    Dim i As Long

    FindCustomFieldRow = -1
    For i = 0 To fields.ListCount - 1
        If StrComp(CStr(fields.List(i, 0)), fieldName, vbTextCompare) = 0 Then
            FindCustomFieldRow = i
            Exit Function
        End If
    Next i
End Function

Public Function AddLabel(ByVal controls As MSForms.Controls, ByVal controlName As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal captionText As String) As MSForms.Label
    Set AddLabel = controls.Add("Forms.Label.1", controlName, True)
    AddLabel.Left = leftPos
    AddLabel.Top = topPos
    AddLabel.Width = widthVal
    AddLabel.Height = heightVal
    AddLabel.Caption = captionText
End Function

Public Function AddTextBox(ByVal controls As MSForms.Controls, ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = controls.Add("Forms.TextBox.1", controlName, True)
    AddTextBox.Left = leftPos
    AddTextBox.Top = topPos
    AddTextBox.Width = widthVal
    AddTextBox.Height = heightVal
End Function

Public Function AddCombo(ByVal controls As MSForms.Controls, ByVal controlName As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single) As MSForms.ComboBox
    Set AddCombo = controls.Add("Forms.ComboBox.1", controlName, True)
    AddCombo.Left = leftPos
    AddCombo.Top = topPos
    AddCombo.Width = widthVal
    AddCombo.Height = heightVal
End Function

Public Function AddListBox(ByVal controls As MSForms.Controls, ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = controls.Add("Forms.ListBox.1", controlName, True)
    AddListBox.Left = leftPos
    AddListBox.Top = topPos
    AddListBox.Width = widthVal
    AddListBox.Height = heightVal
End Function

Public Function AddButton(ByVal controls As MSForms.Controls, ByVal controlName As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single, _
                           ByVal captionText As String) As MSForms.CommandButton
    Set AddButton = controls.Add("Forms.CommandButton.1", controlName, True)
    AddButton.Left = leftPos
    AddButton.Top = topPos
    AddButton.Width = widthVal
    AddButton.Height = heightVal
    AddButton.Caption = captionText
End Function
