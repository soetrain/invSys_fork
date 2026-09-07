Attribute VB_Name = "modUomSettings"
Option Explicit

Private Const CONFIG_KEY_UOM_CATALOG As String = "UomCatalog"
Private Const CONFIG_KEY_UOM_CONVERSION_CATALOG As String = "UomConversionCatalog"
Private Const CONFIG_KEY_UOM_CONVERSION_CATALOG_VERSION As String = "UomConversionCatalogVersion"
Private Const UOM_DELIMITER As String = "|"
Private Const DEFAULT_UOMS As String = "EA|LB|LBS|OZ|KG|G|GAL|QT|PT|L|ML|CS"
Private Const UOM_RECORD_DELIMITER As String = "~"
Private Const DEFAULT_UOM_CONVERSION_CATALOG As String = _
    "EA~DISCRETE~EA~1~FALSE~TRUE|CS~PACKAGING~CS~1~FALSE~TRUE|" & _
    "LB~MASS~LB~1~TRUE~TRUE|LBS~MASS~LB~1~TRUE~TRUE|OZ~MASS~LB~16~TRUE~TRUE|" & _
    "KG~MASS~LB~2.2046226218~TRUE~TRUE|G~MASS~LB~453.59237~TRUE~TRUE|" & _
    "GAL~VOLUME~GAL~1~TRUE~TRUE|QT~VOLUME~GAL~4~TRUE~TRUE|PT~VOLUME~GAL~8~TRUE~TRUE|" & _
    "L~VOLUME~GAL~3.785411784~TRUE~TRUE|ML~VOLUME~GAL~3785.411784~TRUE~TRUE"

Public Function GetConfiguredUoms() As Variant
    Dim uoms As Collection
    Dim result() As Variant
    Dim idx As Long

    Set uoms = ConfiguredUomCollection()
    If uoms Is Nothing Or uoms.Count = 0 Then Exit Function

    ReDim result(1 To uoms.Count)
    For idx = 1 To uoms.Count
        result(idx) = CStr(uoms(idx))
    Next idx
    GetConfiguredUoms = result
End Function

Public Function AddConfiguredUom(ByVal uomName As String, _
                                 Optional ByRef report As String = "") As Boolean
    Dim uoms As Collection

    uomName = NormalizeUomName(uomName)
    If uomName = "" Then
        report = "Enter a UOM containing letters or numbers."
        Exit Function
    End If

    Set uoms = ConfiguredUomCollection()
    If UomCollectionContains(uoms, uomName) Then
        report = uomName & " is already in the warehouse UOM catalog."
        AddConfiguredUom = True
        Exit Function
    End If

    uoms.Add uomName
    AddConfiguredUom = SaveUomCollection(uoms, report)
End Function

Public Function RemoveConfiguredUom(ByVal uomName As String, _
                                    Optional ByRef report As String = "") As Boolean
    Dim uoms As Collection
    Dim idx As Long

    uomName = NormalizeUomName(uomName)
    If uomName = "" Then
        report = "Select a UOM to remove."
        Exit Function
    End If

    Set uoms = ConfiguredUomCollection()
    For idx = uoms.Count To 1 Step -1
        If StrComp(CStr(uoms(idx)), uomName, vbTextCompare) = 0 Then uoms.Remove idx
    Next idx
    If uoms.Count = 0 Then
        report = "The warehouse UOM catalog must contain at least one value."
        Exit Function
    End If
    RemoveConfiguredUom = SaveUomCollection(uoms, report)
End Function

Public Function ResetConfiguredUoms(Optional ByRef report As String = "") As Boolean
    ResetConfiguredUoms = SaveUomCollection(ParseUomPackedText(DEFAULT_UOMS), report)
End Function

Public Function GetConfiguredUomsPackedText() As String
    GetConfiguredUomsPackedText = PackUomCollection(ConfiguredUomCollection())
End Function

Public Function NormalizeConfiguredUomName(ByVal uomName As String) As String
    NormalizeConfiguredUomName = NormalizeUomName(uomName)
End Function

Public Function UomRequiresWholeQuantity(ByVal uomName As String) As Boolean
    UomRequiresWholeQuantity = _
        (StrComp(NormalizeUomName(uomName), "EA", vbTextCompare) = 0)
End Function

Public Function ValidateQuantityForUom(ByVal quantity As Double, _
                                       ByVal uomName As String, _
                                       Optional ByRef report As String = "") As Boolean
    Dim normalizedUom As String

    normalizedUom = NormalizeUomName(uomName)
    If UomRequiresWholeQuantity(normalizedUom) Then
        If Abs(quantity - Fix(quantity)) > 0.0000001 Then
            report = "UOM EA requires a whole quantity; fractional EA is not allowed."
            Exit Function
        End If
    End If
    ValidateQuantityForUom = True
End Function

Public Function GetUomConversion(ByVal fromUom As String, ByVal toUom As String, _
                                 ByRef factorOut As Double, _
                                 Optional ByRef catalogVersionOut As String = "", _
                                 Optional ByRef report As String = "") As Boolean
    Dim records As Collection
    Dim source As Object
    Dim target As Object

    fromUom = NormalizeUomName(fromUom)
    toUom = NormalizeUomName(toUom)
    If fromUom = "" Or toUom = "" Then
        report = "Both source and requirement UOM are required."
        Exit Function
    End If
    Set records = ConfiguredUomConversionRecords()
    Set source = FindUomConversionRecord(records, fromUom)
    Set target = FindUomConversionRecord(records, toUom)
    If source Is Nothing Or target Is Nothing Then
        report = "No UOM Catalog row exists for " & IIf(source Is Nothing, fromUom, toUom) & "."
        Exit Function
    End If
    If Not CBool(source("Enabled")) Or Not CBool(target("Enabled")) Then
        report = "The required UOM Catalog row is disabled."
        Exit Function
    End If
    If StrComp(fromUom, toUom, vbTextCompare) = 0 Then
        factorOut = 1#
    Else
        If Not CBool(source("Convertible")) Or Not CBool(target("Convertible")) Then
            report = "UOM conversion is not enabled for " & fromUom & " or " & toUom & "."
            Exit Function
        End If
        If StrComp(CStr(source("Dimension")), CStr(target("Dimension")), vbTextCompare) <> 0 _
           Or StrComp(CStr(source("BaseUom")), CStr(target("BaseUom")), vbTextCompare) <> 0 Then
            report = "UOM conversion requires matching Dimension and Base UOM."
            Exit Function
        End If
        factorOut = CDbl(target("UnitsPerBase")) / CDbl(source("UnitsPerBase"))
    End If
    If factorOut <= 0# Then
        report = "The UOM Catalog conversion factor is invalid."
        Exit Function
    End If
    catalogVersionOut = CStr(modConfig.GetLong(CONFIG_KEY_UOM_CONVERSION_CATALOG_VERSION, 1))
    GetUomConversion = True
End Function

Public Function TestUomCatalogConversionContract() As String
    Dim factor As Double
    Dim versionText As String
    Dim report As String
    Dim lbToOz As Boolean
    Dim sameUom As Boolean
    Dim eaRejected As Boolean

    lbToOz = GetUomConversion("LB", "OZ", factor, versionText, report) _
        And Abs(factor - 16#) < 0.0000001
    sameUom = GetUomConversion("LBS", "LB", factor, versionText, report) _
        And Abs(factor - 1#) < 0.0000001
    eaRejected = Not GetUomConversion("EA", "OZ", factor, versionText, report)
    TestUomCatalogConversionContract = IIf(lbToOz And sameUom And eaRejected, "OK", "FAIL") & _
        "|LbToOz=" & CStr(lbToOz) & "|LbsToLb=" & CStr(sameUom) & _
        "|EaRejected=" & CStr(eaRejected)
End Function

Public Function GetUomCatalogRows() As Variant
    Dim records As Collection
    Dim result() As Variant
    Dim index As Long
    Dim record As Object

    Set records = ConfiguredUomConversionRecords()
    If records Is Nothing Or records.Count = 0 Then Exit Function
    ReDim result(1 To records.Count, 1 To 7)
    For index = 1 To records.Count
        Set record = records(index)
        result(index, 1) = record("UOM")
        result(index, 2) = record("Dimension")
        result(index, 3) = record("BaseUom")
        result(index, 4) = record("UnitsPerBase")
        result(index, 5) = record("Convertible")
        result(index, 6) = record("Enabled")
        result(index, 7) = ""
    Next index
    GetUomCatalogRows = result
End Function

Public Function PublishUomCatalogRows(ByVal rows As Variant, _
                                      Optional ByRef report As String = "") As Boolean
    PublishUomCatalogRows = modConfigCommands.PublishUomCatalogRows(rows, report)
End Function

Public Function UomCatalogMatches(ByVal packedUoms As String, ByVal packedConversions As String) As Boolean
    Dim currentConversions As String
    currentConversions = modConfig.GetString(CONFIG_KEY_UOM_CONVERSION_CATALOG, "")
    If currentConversions = "" Then currentConversions = DEFAULT_UOM_CONVERSION_CATALOG
    UomCatalogMatches = (StrComp(modConfig.GetString(CONFIG_KEY_UOM_CATALOG, DEFAULT_UOMS), packedUoms, vbTextCompare) = 0 _
        And StrComp(currentConversions, packedConversions, vbTextCompare) = 0)
End Function

Public Function PrepareUomCatalogRows(ByVal rows As Variant, ByRef packedUoms As String, ByRef packedConversions As String, _
                                      Optional ByRef report As String = "") As Boolean
    Dim records As Object
    Dim rowIndex As Long
    Dim uomName As String
    Dim dimension As String
    Dim baseUom As String
    Dim unitsPerBase As Double
    Dim convertible As Boolean
    Dim enabled As Boolean
    Dim key As Variant
    Dim record As Variant
    Dim baseRecord As Variant
    packedUoms = ""
    packedConversions = ""
    If Not IsArray(rows) Then
        report = "Select a populated UOM Catalog table."
        Exit Function
    End If
    Set records = CreateObject("Scripting.Dictionary")
    records.CompareMode = vbTextCompare
    For rowIndex = LBound(rows, 1) To UBound(rows, 1)
        uomName = NormalizeUomName(CStr(rows(rowIndex, 1)))
        dimension = UCase$(Trim$(CStr(rows(rowIndex, 2))))
        baseUom = NormalizeUomName(CStr(rows(rowIndex, 3)))
        If uomName = "" And dimension = "" And baseUom = "" Then GoTo NextRow
        If uomName = "" Or dimension = "" Or baseUom = "" Or Not IsNumeric(rows(rowIndex, 4)) Then
            report = "Every UOM Catalog row needs UOM, Dimension, Base UOM, and positive Units Per Base UOM."
            Exit Function
        End If
        unitsPerBase = CDbl(rows(rowIndex, 4))
        If unitsPerBase <= 0# Then
            report = "Units Per Base UOM must be positive for " & uomName & "."
            Exit Function
        End If
        If records.Exists(uomName) Then
            report = "Duplicate UOM Catalog row " & uomName & "."
            Exit Function
        End If
        convertible = (StrComp(Trim$(CStr(rows(rowIndex, 5))), "TRUE", vbTextCompare) = 0)
        enabled = (StrComp(Trim$(CStr(rows(rowIndex, 6))), "TRUE", vbTextCompare) = 0)
        If (uomName = "EA" Or uomName = "CS") And convertible Then
            report = uomName & " cannot be a globally convertible UOM."
            Exit Function
        End If
        records.Add uomName, Array(dimension, baseUom, unitsPerBase, convertible, enabled)
NextRow:
    Next rowIndex
    If records.Count = 0 Then
        report = "The UOM Catalog requires at least one row."
        Exit Function
    End If
    For Each key In records.Keys
        record = records(key)
        If Not records.Exists(CStr(record(1))) Then
            report = "Base UOM " & CStr(record(1)) & " is missing for " & CStr(key) & "."
            Exit Function
        End If
        baseRecord = records(CStr(record(1)))
        If StrComp(CStr(baseRecord(0)), CStr(record(0)), vbTextCompare) <> 0 Then
            report = "Base UOM dimension must match for " & CStr(key) & "."
            Exit Function
        End If
        If StrComp(CStr(key), CStr(record(1)), vbTextCompare) = 0 And Abs(CDbl(record(2)) - 1#) > 0.0000001 Then
            report = "Base UOM " & CStr(key) & " must use Units Per Base UOM of 1."
            Exit Function
        End If
    Next key
    For Each key In records.Keys
        record = records(key)
        If packedConversions <> "" Then packedConversions = packedConversions & UOM_DELIMITER
        packedConversions = packedConversions & CStr(key) & UOM_RECORD_DELIMITER & CStr(record(0)) & _
            UOM_RECORD_DELIMITER & CStr(record(1)) & UOM_RECORD_DELIMITER & _
            Replace$(CStr(record(2)), ",", ".") & UOM_RECORD_DELIMITER & _
            CStr(CBool(record(3))) & UOM_RECORD_DELIMITER & CStr(CBool(record(4)))
        If packedUoms <> "" Then packedUoms = packedUoms & UOM_DELIMITER
        packedUoms = packedUoms & CStr(key)
    Next key
    PrepareUomCatalogRows = True
End Function

Private Function ConfiguredUomConversionRecords() As Collection
    Dim packed As String
    Dim rows As Variant
    Dim rowValue As Variant
    Dim fields As Variant
    Dim record As Object

    packed = Trim$(modConfig.GetString(CONFIG_KEY_UOM_CONVERSION_CATALOG, ""))
    If packed = "" Then packed = DEFAULT_UOM_CONVERSION_CATALOG
    Set ConfiguredUomConversionRecords = New Collection
    rows = Split(packed, UOM_DELIMITER)
    For Each rowValue In rows
        fields = Split(CStr(rowValue), UOM_RECORD_DELIMITER)
        If UBound(fields) >= 5 Then
            Set record = CreateObject("Scripting.Dictionary")
            record.CompareMode = vbTextCompare
            record("UOM") = NormalizeUomName(CStr(fields(0)))
            record("Dimension") = UCase$(Trim$(CStr(fields(1))))
            record("BaseUom") = NormalizeUomName(CStr(fields(2)))
            If IsNumeric(fields(3)) Then record("UnitsPerBase") = CDbl(fields(3)) Else record("UnitsPerBase") = 0#
            record("Convertible") = (StrComp(Trim$(CStr(fields(4))), "TRUE", vbTextCompare) = 0)
            record("Enabled") = (StrComp(Trim$(CStr(fields(5))), "TRUE", vbTextCompare) = 0)
            If record("UOM") <> "" Then ConfiguredUomConversionRecords.Add record
        End If
    Next rowValue
End Function

Private Function FindUomConversionRecord(ByVal records As Collection, ByVal uomName As String) As Object
    Dim rawRecord As Variant
    Dim record As Object

    If records Is Nothing Then Exit Function
    For Each rawRecord In records
        Set record = rawRecord
        If StrComp(CStr(record("UOM")), uomName, vbTextCompare) = 0 Then
            Set FindUomConversionRecord = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function ConfiguredUomCollection() As Collection
    Dim packed As String

    packed = Trim$(modConfig.GetString(CONFIG_KEY_UOM_CATALOG, DEFAULT_UOMS))
    If packed = "" Then packed = DEFAULT_UOMS
    Set ConfiguredUomCollection = ParseUomPackedText(packed)
    If ConfiguredUomCollection.Count = 0 Then _
        Set ConfiguredUomCollection = ParseUomPackedText(DEFAULT_UOMS)
End Function

Private Function ParseUomPackedText(ByVal packedText As String) As Collection
    Dim normalized As String
    Dim parts As Variant
    Dim idx As Long
    Dim uoms As New Collection
    Dim uomName As String

    normalized = Replace$(packedText, vbCrLf, UOM_DELIMITER)
    normalized = Replace$(normalized, vbCr, UOM_DELIMITER)
    normalized = Replace$(normalized, vbLf, UOM_DELIMITER)
    normalized = Replace$(normalized, ",", UOM_DELIMITER)
    normalized = Replace$(normalized, ";", UOM_DELIMITER)
    parts = Split(normalized, UOM_DELIMITER)
    For idx = LBound(parts) To UBound(parts)
        uomName = NormalizeUomName(CStr(parts(idx)))
        If uomName <> "" Then
            If Not UomCollectionContains(uoms, uomName) Then uoms.Add uomName
        End If
    Next idx
    Set ParseUomPackedText = uoms
End Function

Private Function SaveUomCollection(ByVal uoms As Collection, ByRef report As String) As Boolean
    Dim packed As String

    packed = PackUomCollection(uoms)
    If packed = "" Then
        report = "The warehouse UOM catalog must contain at least one value."
        Exit Function
    End If
    SaveUomCollection = modConfigCommands.UpdateConfigValue(CONFIG_KEY_UOM_CATALOG, packed, report)
End Function

Private Function PackUomCollection(ByVal uoms As Collection) As String
    Dim idx As Long
    Dim uomName As String

    If uoms Is Nothing Then Exit Function
    For idx = 1 To uoms.Count
        uomName = NormalizeUomName(CStr(uoms(idx)))
        If uomName <> "" Then
            If PackUomCollection <> "" Then PackUomCollection = PackUomCollection & UOM_DELIMITER
            PackUomCollection = PackUomCollection & uomName
        End If
    Next idx
End Function

Private Function UomCollectionContains(ByVal uoms As Collection, ByVal uomName As String) As Boolean
    Dim idx As Long

    If uoms Is Nothing Then Exit Function
    For idx = 1 To uoms.Count
        If StrComp(CStr(uoms(idx)), uomName, vbTextCompare) = 0 Then
            UomCollectionContains = True
            Exit Function
        End If
    Next idx
End Function

Private Function NormalizeUomName(ByVal uomName As String) As String
    Dim idx As Long
    Dim ch As String
    Dim normalized As String

    uomName = UCase$(Trim$(uomName))
    For idx = 1 To Len(uomName)
        ch = Mid$(uomName, idx, 1)
        If ch Like "[A-Z0-9]" Or ch = "/" Or ch = "-" Then
            normalized = normalized & ch
        ElseIf ch = " " Then
            If normalized <> "" Then
                If Right$(normalized, 1) <> " " Then normalized = normalized & " "
            End If
        End If
    Next idx
    NormalizeUomName = Trim$(normalized)
End Function
