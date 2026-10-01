Attribute VB_Name = "modProductionRunEntityReads"
Option Explicit
Option Private Module

' Existing exact-entity readers with optional captured-action continuation.
Public Function AvailableQuantity(ByVal systemKey As String, ByRef locationOut As String, _
                                  ByVal action As cProductionWorksheetAction) As Double
    Dim entities As Variant, r As Long
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(r, 1))), Trim$(systemKey), vbTextCompare) = 0 Then
            If IsNumeric(entities(r, 6)) Then AvailableQuantity = CDbl(entities(r, 6))
            locationOut = Trim$(CStr(entities(r, 7)))
            Exit Function
        End If
    Next r
End Function

Public Function IsNonCounted(ByVal systemKey As String, ByVal action As cProductionWorksheetAction) As Boolean
    Dim entities As Variant, r As Long
    Dim trackQty As String, itemKind As String, categoryValue As String
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(r, 1))), Trim$(systemKey), vbTextCompare) = 0 Then
            If UBound(entities, 2) >= 11 Then trackQty = UCase$(Trim$(CStr(entities(r, 11))))
            If UBound(entities, 2) >= 12 Then itemKind = UCase$(Trim$(CStr(entities(r, 12))))
            If UBound(entities, 2) >= 13 Then categoryValue = UCase$(Trim$(CStr(entities(r, 13))))
            IsNonCounted = (trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
                Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED" _
                Or categoryValue = "UTILITY" Or categoryValue = "SERVICE")
            Exit Function
        End If
    Next r
End Function
