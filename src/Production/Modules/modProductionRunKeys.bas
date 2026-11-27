Attribute VB_Name = "modProductionRunKeys"
Option Explicit
Option Private Module

Public Function AllocationKey(ByVal nodeId As String, ByVal requirementId As String, _
                               ByVal systemKey As String) As String
    AllocationKey = nodeId & vbTab & requirementId & vbTab & systemKey
End Function

Public Function AllocationSystemKey(ByVal allocationKeyValue As String) As String
    Dim parts() As String
    parts = Split(allocationKeyValue, vbTab)
    If UBound(parts) >= 2 Then AllocationSystemKey = parts(2)
End Function

Public Function AllocationBelongsToNode(ByVal allocationId As String, _
                                         ByVal nodeId As String) As Boolean
    Dim parts() As String
    parts = Split(allocationId, vbTab)
    If UBound(parts) = 2 Then _
        AllocationBelongsToNode = _
            (StrComp(parts(0), nodeId, vbTextCompare) = 0)
End Function

Public Function OutputIdentityKey(ByVal nodeId As String, ByVal outputId As String) As String
    OutputIdentityKey = nodeId & vbTab & outputId
End Function

Public Function StockBucketKey(ByVal sku As String, ByVal uom As String, _
                                ByVal locationValue As String, _
                                ByVal conditionValue As String) As String
    StockBucketKey = UCase$(Trim$(sku)) & vbTab & UCase$(Trim$(uom)) & vbTab & _
                     UCase$(Trim$(locationValue)) & vbTab & UCase$(Trim$(conditionValue))
End Function

Public Function StockBucketKeyForEntity(ByVal entities As Variant, _
                                         ByVal entityRow As Long) As String
    Dim conditionValue As String

    If UBound(entities, 2) >= 8 Then conditionValue = CStr(entities(entityRow, 8))
    StockBucketKeyForEntity = StockBucketKey(CStr(entities(entityRow, 2)), _
        CStr(entities(entityRow, 5)), CStr(entities(entityRow, 7)), conditionValue)
End Function
