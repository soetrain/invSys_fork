Attribute VB_Name = "modCarrierSettings"
Option Explicit

Private Const SETTINGS_APP As String = "invSys"
Private Const SETTINGS_SECTION_SHIPPING As String = "Shipping"
Private Const SETTINGS_CARRIERS As String = "Carriers"
Private Const CARRIER_DELIMITER As String = "|"

Public Function GetConfiguredCarriers() As Variant
    Dim carriers As Collection
    Dim result() As Variant
    Dim idx As Long

    Set carriers = ConfiguredCarrierCollection()
    If carriers Is Nothing Then Exit Function
    If carriers.Count = 0 Then Exit Function

    ReDim result(1 To carriers.Count)
    For idx = 1 To carriers.Count
        result(idx) = carriers(idx)
    Next idx
    GetConfiguredCarriers = result
End Function

Public Function GetConfiguredCarriersText() As String
    Dim carriers As Collection
    Dim idx As Long

    Set carriers = ConfiguredCarrierCollection()
    If carriers Is Nothing Then Exit Function

    For idx = 1 To carriers.Count
        If GetConfiguredCarriersText <> "" Then GetConfiguredCarriersText = GetConfiguredCarriersText & vbCrLf
        GetConfiguredCarriersText = GetConfiguredCarriersText & CStr(carriers(idx))
    Next idx
End Function

Public Sub SaveConfiguredCarriersText(ByVal carrierText As String)
    Dim carriers As Collection

    Set carriers = ParseCarrierLines(carrierText)
    SaveCarrierCollection carriers
End Sub

Public Function AddConfiguredCarrier(ByVal carrierName As String) As Boolean
    Dim carriers As Collection

    carrierName = NormalizeCarrierName(carrierName)
    If carrierName = "" Then Exit Function

    Set carriers = ConfiguredCarrierCollection()
    If CarrierCollectionContains(carriers, carrierName) Then
        AddConfiguredCarrier = True
        Exit Function
    End If

    carriers.Add carrierName
    SaveCarrierCollection carriers
    AddConfiguredCarrier = True
End Function

Public Function RemoveConfiguredCarrier(ByVal carrierName As String) As Boolean
    Dim carriers As Collection
    Dim idx As Long

    carrierName = NormalizeCarrierName(carrierName)
    If carrierName = "" Then Exit Function

    Set carriers = ConfiguredCarrierCollection()
    For idx = carriers.Count To 1 Step -1
        If StrComp(CStr(carriers(idx)), carrierName, vbTextCompare) = 0 Then
            carriers.Remove idx
            RemoveConfiguredCarrier = True
        End If
    Next idx

    If RemoveConfiguredCarrier Then SaveCarrierCollection carriers
End Function

Public Sub ResetConfiguredCarriers()
    SaveCarrierCollection DefaultCarrierCollection()
End Sub

Private Function ConfiguredCarrierCollection() As Collection
    Dim packed As String

    packed = Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION_SHIPPING, SETTINGS_CARRIERS, ""))
    If packed = "" Then
        Set ConfiguredCarrierCollection = DefaultCarrierCollection()
    Else
        Set ConfiguredCarrierCollection = ParseCarrierPackedText(packed)
        If ConfiguredCarrierCollection.Count = 0 Then Set ConfiguredCarrierCollection = DefaultCarrierCollection()
    End If
End Function

Private Function DefaultCarrierCollection() As Collection
    Dim carriers As New Collection

    carriers.Add "UPS"
    carriers.Add "USPS"
    carriers.Add "FedEx"
    carriers.Add "DHL"
    Set DefaultCarrierCollection = carriers
End Function

Private Function ParseCarrierLines(ByVal carrierText As String) As Collection
    Dim normalizedText As String
    Dim parts As Variant
    Dim idx As Long
    Dim carriers As New Collection
    Dim carrierName As String

    normalizedText = Replace$(carrierText, vbCrLf, vbLf)
    normalizedText = Replace$(normalizedText, vbCr, vbLf)
    parts = Split(normalizedText, vbLf)
    For idx = LBound(parts) To UBound(parts)
        carrierName = NormalizeCarrierName(CStr(parts(idx)))
        If carrierName <> "" Then AddCarrierUnique carriers, carrierName
    Next idx
    Set ParseCarrierLines = carriers
End Function

Private Function ParseCarrierPackedText(ByVal packedText As String) As Collection
    Dim parts As Variant
    Dim idx As Long
    Dim carriers As New Collection
    Dim carrierName As String

    parts = Split(packedText, CARRIER_DELIMITER)
    For idx = LBound(parts) To UBound(parts)
        carrierName = NormalizeCarrierName(CStr(parts(idx)))
        If carrierName <> "" Then AddCarrierUnique carriers, carrierName
    Next idx
    Set ParseCarrierPackedText = carriers
End Function

Private Sub SaveCarrierCollection(ByVal carriers As Collection)
    Dim idx As Long
    Dim packed As String
    Dim carrierName As String

    If carriers Is Nothing Then Set carriers = New Collection
    For idx = 1 To carriers.Count
        carrierName = NormalizeCarrierName(CStr(carriers(idx)))
        If carrierName <> "" Then
            If packed <> "" Then packed = packed & CARRIER_DELIMITER
            packed = packed & carrierName
        End If
    Next idx
    SaveSetting SETTINGS_APP, SETTINGS_SECTION_SHIPPING, SETTINGS_CARRIERS, packed
End Sub

Private Sub AddCarrierUnique(ByVal carriers As Collection, ByVal carrierName As String)
    If CarrierCollectionContains(carriers, carrierName) Then Exit Sub
    carriers.Add carrierName
End Sub

Private Function CarrierCollectionContains(ByVal carriers As Collection, ByVal carrierName As String) As Boolean
    Dim idx As Long

    If carriers Is Nothing Then Exit Function
    For idx = 1 To carriers.Count
        If StrComp(CStr(carriers(idx)), carrierName, vbTextCompare) = 0 Then
            CarrierCollectionContains = True
            Exit Function
        End If
    Next idx
End Function

Private Function NormalizeCarrierName(ByVal carrierName As String) As String
    carrierName = Trim$(carrierName)
    carrierName = Replace$(carrierName, CARRIER_DELIMITER, " ")
    Do While InStr(1, carrierName, "  ", vbBinaryCompare) > 0
        carrierName = Replace$(carrierName, "  ", " ")
    Loop
    NormalizeCarrierName = carrierName
End Function
