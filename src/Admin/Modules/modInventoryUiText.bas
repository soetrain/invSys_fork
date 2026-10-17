Attribute VB_Name = "modInventoryUiText"
Option Explicit
Option Private Module

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

