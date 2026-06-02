Attribute VB_Name = "modItemSearch"
Option Explicit

Public Function NormalizeSearchText(ByVal valueIn As String) As String
    Dim textOut As String

    textOut = Trim$(valueIn)
    If textOut = "" Then Exit Function

    textOut = Replace$(textOut, vbCr, " ")
    textOut = Replace$(textOut, vbLf, " ")
    textOut = Replace$(textOut, vbTab, " ")
    Do While InStr(textOut, "  ") > 0
        textOut = Replace$(textOut, "  ", " ")
    Loop

    NormalizeSearchText = LCase$(textOut)
End Function

Public Function TextMatchesSearch(ByVal candidate As String, ByVal searchTerm As String) As Boolean
    Dim normalizedTerm As String

    normalizedTerm = NormalizeSearchText(searchTerm)
    If normalizedTerm = "" Then
        TextMatchesSearch = True
        Exit Function
    End If

    TextMatchesSearch = (InStr(1, NormalizeSearchText(candidate), normalizedTerm, vbTextCompare) > 0)
End Function

Public Function AnyTextMatchesSearch(ByVal searchTerm As String, ParamArray candidates() As Variant) As Boolean
    Dim i As Long

    If NormalizeSearchText(searchTerm) = "" Then
        AnyTextMatchesSearch = True
        Exit Function
    End If

    For i = LBound(candidates) To UBound(candidates)
        If TextMatchesSearch(CStr(candidates(i)), searchTerm) Then
            AnyTextMatchesSearch = True
            Exit Function
        End If
    Next i
End Function

Public Function IdentifierTokens(ByVal valueIn As String) As Variant
    Dim normalized As String
    Dim parts As Variant
    Dim cleaned() As String
    Dim i As Long
    Dim n As Long

    normalized = NormalizeSearchText(valueIn)
    If normalized = "" Then Exit Function

    parts = Split(normalized, " ")
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then
            n = n + 1
            ReDim Preserve cleaned(0 To n - 1)
            cleaned(n - 1) = Trim$(parts(i))
        End If
    Next i

    If n = 0 Then Exit Function
    IdentifierTokens = cleaned
End Function

Public Function IdentifiersMatch(ByVal leftValue As String, ByVal rightValue As String) As Boolean
    Dim leftTokens As Variant
    Dim rightTokens As Variant
    Dim i As Long
    Dim j As Long

    leftTokens = IdentifierTokens(leftValue)
    rightTokens = IdentifierTokens(rightValue)
    If IsEmpty(leftTokens) Or IsEmpty(rightTokens) Then Exit Function

    For i = LBound(leftTokens) To UBound(leftTokens)
        For j = LBound(rightTokens) To UBound(rightTokens)
            If StrComp(leftTokens(i), rightTokens(j), vbTextCompare) = 0 Then
                IdentifiersMatch = True
                Exit Function
            End If
        Next j
    Next i
End Function

Public Function ResolveSearchRole(ByVal templateFormName As String) As String
    Select Case LCase$(Trim$(templateFormName))
        Case "ufreceivingitemsearch"
            ResolveSearchRole = "receiving"
        Case "ufshippingitemsearch"
            ResolveSearchRole = "shipping"
        Case "ufproductionitemsearch"
            ResolveSearchRole = "production"
        Case "ufadminitemsearch"
            ResolveSearchRole = "admin"
    End Select
End Function

Public Function ResolveSearchCaption(ByVal roleKey As String, ByVal pickerMode As String) As String
    Dim resolvedRole As String
    Dim resolvedMode As String

    resolvedRole = LCase$(Trim$(roleKey))
    resolvedMode = LCase$(Trim$(pickerMode))

    Select Case resolvedMode
        Case "ingredient"
            Select Case resolvedRole
                Case "production"
                    ResolveSearchCaption = "Production Ingredient Search"
                Case Else
                    ResolveSearchCaption = "Ingredient Search"
            End Select
        Case "palette_item"
            Select Case resolvedRole
                Case "production"
                    ResolveSearchCaption = "Production Palette Item Search"
                Case Else
                    ResolveSearchCaption = "Palette Item Search"
            End Select
        Case "recipe", "palette_recipe", "recipe_chooser"
            Select Case resolvedRole
                Case "production"
                    ResolveSearchCaption = "Production Recipe Search"
                Case Else
                    ResolveSearchCaption = "Recipe Search"
            End Select
        Case Else
            Select Case resolvedRole
                Case "receiving"
                    ResolveSearchCaption = "Receiving Item Search"
                Case "shipping"
                    ResolveSearchCaption = "Shipping Item Search"
                Case "production"
                    ResolveSearchCaption = "Production Item Search"
                Case "admin"
                    ResolveSearchCaption = "Admin Item Search"
                Case Else
                    ResolveSearchCaption = "Item Search"
            End Select
    End Select
End Function

Public Function ShouldDefaultShippableForRole(ByVal roleKey As String, _
                                              ByVal pickerMode As String, _
                                              Optional ByVal sourceTableName As String = "") As Boolean
    Dim resolvedRole As String
    Dim resolvedMode As String

    resolvedRole = LCase$(Trim$(roleKey))
    resolvedMode = LCase$(Trim$(pickerMode))

    If resolvedMode = "recipe" Or resolvedMode = "palette_recipe" _
        Or resolvedMode = "recipe_chooser" Or resolvedMode = "ingredient" _
        Or resolvedMode = "palette_item" Then Exit Function

    If resolvedRole = "shipping" Then
        ShouldDefaultShippableForRole = True
        Exit Function
    End If

    If LCase$(Trim$(sourceTableName)) = "shipmentstally" Then
        ShouldDefaultShippableForRole = True
    End If
End Function

Public Function IsShippingRelevantCategory(ByVal categoryText As String) As Boolean
    Dim normalized As String

    normalized = NormalizeSearchText(categoryText)
    If normalized = "" Then
        IsShippingRelevantCategory = True
        Exit Function
    End If

    If normalized = "shippable" Or normalized = "sell" Then
        IsShippingRelevantCategory = True
        Exit Function
    End If

    If InStr(1, normalized, "packaging.ship", vbTextCompare) > 0 _
       Or InStr(1, normalized, " ship", vbTextCompare) > 0 _
       Or InStr(1, normalized, ".ship", vbTextCompare) > 0 _
       Or InStr(1, normalized, "ship.", vbTextCompare) > 0 Then
        IsShippingRelevantCategory = True
    End If
End Function
