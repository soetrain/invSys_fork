Attribute VB_Name = "TestCoreItemSearch"
Option Explicit

Public Sub RunCoreItemSearchTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestNormalizeSearchText_CollapsesWhitespace(), passed, failed
    Tally TestAnyTextMatchesSearch_MatchesAcrossFields(), passed, failed
    Tally TestIdentifiersMatch_UsesTokenOverlap(), passed, failed
    Tally TestResolveSearchCaption_ReturnsRoleSpecificText(), passed, failed
    Tally TestShouldDefaultShippableForRole_UsesRoleDefaults(), passed, failed
    Tally TestIsShippingRelevantCategory_AcceptsShippingSeedCategories(), passed, failed

    Debug.Print "Core.ItemSearch tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestNormalizeSearchText_CollapsesWhitespace() As Long
    Dim normalized As String

    normalized = modItemSearch.NormalizeSearchText("  SKU" & vbTab & "  123 " & vbCrLf & " Blue ")
    If normalized <> "sku 123 blue" Then Exit Function

    TestNormalizeSearchText_CollapsesWhitespace = 1
End Function

Public Function TestAnyTextMatchesSearch_MatchesAcrossFields() As Long
    If Not modItemSearch.AnyTextMatchesSearch("sku-2", "Widget Blue", "SKU-200", "Dock A") Then Exit Function
    If modItemSearch.AnyTextMatchesSearch("missing", "Widget Blue", "SKU-200", "Dock A") Then Exit Function

    TestAnyTextMatchesSearch_MatchesAcrossFields = 1
End Function

Public Function TestIdentifiersMatch_UsesTokenOverlap() As Long
    If Not modItemSearch.IdentifiersMatch("recipe-123 alpha", "alpha batch-7") Then Exit Function
    If modItemSearch.IdentifiersMatch("recipe-123", "batch-7") Then Exit Function

    TestIdentifiersMatch_UsesTokenOverlap = 1
End Function

Public Function TestResolveSearchCaption_ReturnsRoleSpecificText() As Long
    If modItemSearch.ResolveSearchCaption("receiving", "item") <> "Receiving Item Search" Then Exit Function
    If modItemSearch.ResolveSearchCaption("shipping", "item") <> "Shipping Item Search" Then Exit Function
    If modItemSearch.ResolveSearchCaption("production", "recipe") <> "Production Recipe Search" Then Exit Function

    TestResolveSearchCaption_ReturnsRoleSpecificText = 1
End Function

Public Function TestShouldDefaultShippableForRole_UsesRoleDefaults() As Long
    If Not modItemSearch.ShouldDefaultShippableForRole("shipping", "item") Then Exit Function
    If modItemSearch.ShouldDefaultShippableForRole("receiving", "recipe") Then Exit Function
    If modItemSearch.ShouldDefaultShippableForRole("receiving", "item") Then Exit Function
    If Not modItemSearch.ShouldDefaultShippableForRole("receiving", "item", "ShipmentsTally") Then Exit Function

    TestShouldDefaultShippableForRole_UsesRoleDefaults = 1
End Function

Public Function TestIsShippingRelevantCategory_AcceptsShippingSeedCategories() As Long
    If Not modItemSearch.IsShippingRelevantCategory("shippable") Then Exit Function
    If Not modItemSearch.IsShippingRelevantCategory("sell") Then Exit Function
    If Not modItemSearch.IsShippingRelevantCategory("packaging.ship") Then Exit Function
    If Not modItemSearch.IsShippingRelevantCategory("Ship.used") Then Exit Function
    If Not modItemSearch.IsShippingRelevantCategory("") Then Exit Function
    If modItemSearch.IsShippingRelevantCategory("packaging.cook") Then Exit Function
    If modItemSearch.IsShippingRelevantCategory("raw") Then Exit Function

    TestIsShippingRelevantCategory_AcceptsShippingSeedCategories = 1
End Function

Private Sub Tally(ByVal resultIn As Long, ByRef passed As Long, ByRef failed As Long)
    If resultIn = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
