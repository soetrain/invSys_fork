Attribute VB_Name = "modEvaluationSources"
Option Explicit
Option Private Module

Public Sub Retain(ByVal result As Object, ByVal terminal As Object, ByVal publication As Object)
    Dim reference As Variant, evidence As Object, group As Object, source As Object, keys As Collection
    Dim groups As Object, coverage As Object
    Set groups = CreateObject("Scripting.Dictionary"): Set coverage = CreateObject("Scripting.Dictionary")
    If Not publication Is Nothing Then
        For Each group In publication("Groups"): groups.Add CStr(group("Source")) & vbTab & CStr(group("SourceId")), group: Next group
        For Each source In publication("Coverage")("Sources"): coverage.Add CStr(source("Source")), source: Next source
    End If
    For Each reference In terminal("SourceEventRefs")
        Set evidence = CreateObject("Scripting.Dictionary")
        evidence.Add "WarehouseId", reference("WarehouseId"): evidence.Add "SourceKind", reference("SourceKind")
        evidence.Add "EventId", reference("EventId"): evidence.Add "SubmissionState", reference("SubmissionState")
        evidence.Add "OwnerStatus", "Unavailable": evidence.Add "LineCount", 0&: evidence.Add "LinesSha256", ""
        Set keys = New Collection: evidence.Add "SystemKeys", keys
        If reference("WarehouseId") = result("WarehouseId") And reference("SourceKind") = "Inventory" Then
            If groups.Exists("Inventory" & vbTab & CStr(reference("EventId"))) Then
                Set group = groups("Inventory" & vbTab & CStr(reference("EventId")))
                If AppliedInventory(group, CStr(result("WarehouseId")), evidence) Then evidence("OwnerStatus") = "Applied"
            ElseIf coverage.Exists("Inventory") And reference("SubmissionState") = "Submitted" Then
                Set source = coverage("Inventory")
                If source("Availability") = "Available" Then
                    If source("OmittedGroups") = 0 And source("OmittedLines") = 0 Then evidence("OwnerStatus") = "Awaiting"
                End If
            End If
        End If
        result("TerminalSources").Add evidence
    Next reference
End Sub

Public Sub Assess(ByVal result As Object)
    Dim evidence As Object, pending As Boolean, unavailable As Boolean
    unavailable = (result("TerminalSources").Count = 0)
    For Each evidence In result("TerminalSources")
        If evidence("OwnerStatus") = "Unavailable" Then unavailable = True
        If evidence("OwnerStatus") = "Awaiting" Then pending = True
    Next evidence
    If unavailable Then
        result("ResultState") = "Incomplete": modEvaluationModel.Reason result, "SOURCE_UNAVAILABLE"
    ElseIf pending Then
        result("ResultState") = "Awaiting": modEvaluationModel.Reason result, "SOURCE_PENDING"
    Else
        result("ResultState") = "Concluded": modEvaluationModel.Reason result, "SOURCE_APPLIED"
    End If
End Sub

Private Function AppliedInventory(ByVal group As Object, ByVal warehouseId As String, ByVal evidence As Object) As Boolean
    Dim line As Object, index As Long, wrapper As Object, keys As New Collection
    On Error GoTo Unavailable
    If group("Source") <> "Inventory" Or group("SourceKind") <> "Business event" Then Exit Function
    If group("Lines").Count = 0 Or group("Outcomes").Count <> group("Lines").Count Then Exit Function
    For Each line In group("Lines")
        index = index + 1
        If VarType(line("EventID")) <> vbString Or line("EventID") <> group("SourceId") Then Exit Function
        If VarType(line("System_Key")) <> vbString Or line("System_Key") = "" Then Exit Function
        If VarType(line("AppliedAtUTC")) <> vbString Or line("AppliedAtUTC") = "" Then Exit Function
        If line.Exists("WarehouseId") Then
            If line("WarehouseId") <> warehouseId Then Exit Function
        End If
        If modTrainingJson.EncodeObject(line) <> modTrainingJson.EncodeObject(group("Outcomes")(index)) Then Exit Function
        keys.Add CStr(line("System_Key"))
    Next line
    Set wrapper = CreateObject("Scripting.Dictionary"): wrapper.Add "Lines", group("Lines")
    evidence("LinesSha256") = modTrainingWire.Sha256(modTrainingJson.EncodeObject(wrapper))
    If Not modEvaluationModel.IsHash(CStr(evidence("LinesSha256"))) Then Exit Function
    evidence("LineCount") = group("Lines").Count: Set evidence("SystemKeys") = keys
    AppliedInventory = True
Unavailable:
End Function

Public Function ValidSavedSources(ByVal sources As Collection, ByVal warehouseId As String) As Boolean
    Dim source As Object, value As Variant, field As Variant, seen As Object, key As String
    On Error GoTo Invalid
    Set seen = CreateObject("Scripting.Dictionary")
    For Each source In sources
        If Not modEvaluationModel.HasFields(source, "WarehouseId|SourceKind|EventId|SubmissionState|OwnerStatus|LineCount|LinesSha256|SystemKeys") Then Exit Function
        For Each field In Array("WarehouseId", "SourceKind", "EventId", "SubmissionState", "OwnerStatus", "LinesSha256")
            If VarType(source(field)) <> vbString Then Exit Function
        Next field
        If source("WarehouseId") <> warehouseId Or source("EventId") = "" Then Exit Function
        If source("SourceKind") <> "Inventory" And source("SourceKind") <> "Designs" Then Exit Function
        If source("SubmissionState") <> "Submitted" And source("SubmissionState") <> "Unknown" Then Exit Function
        key = source("SourceKind") & vbTab & source("EventId")
        If seen.Exists(key) Then Exit Function
        seen.Add key, True
        If Not modEvaluationModel.IsInteger(source("LineCount")) Or source("LineCount") < 0 Then Exit Function
        If TypeName(source("SystemKeys")) <> "Collection" Then Exit Function
        If source("OwnerStatus") = "Applied" Then
            If source("SourceKind") <> "Inventory" Or source("LineCount") = 0 Or source("SystemKeys").Count <> source("LineCount") Then Exit Function
            If Not modEvaluationModel.IsHash(CStr(source("LinesSha256"))) Then Exit Function
            For Each value In source("SystemKeys")
                If VarType(value) <> vbString Or value = "" Then Exit Function
            Next value
        Else
            If source("OwnerStatus") <> "Awaiting" And source("OwnerStatus") <> "Unavailable" Then Exit Function
            If source("OwnerStatus") = "Awaiting" And source("SubmissionState") <> "Submitted" Then Exit Function
            If source("LineCount") <> 0 Or source("LinesSha256") <> "" Or source("SystemKeys").Count <> 0 Then Exit Function
        End If
    Next source
    ValidSavedSources = True
Invalid:
End Function

Public Function ValidPublication(ByVal publication As Object, ByVal warehouseId As String) As Boolean
    Dim field As Variant, source As Object, seen As Object, kind As Variant
    On Error GoTo Invalid
    If Not modEvaluationModel.HasFields(publication, "Availability|WarehouseId|PublicationId|ContentSha256|SchemaVersion|PackageSetVersion|BuildIdentity|PublishedAtUTC|LoadedAtUTC|PolicyVersion|Coverage") Then Exit Function
    For Each field In Array("Availability", "WarehouseId", "PublicationId", "ContentSha256", "PackageSetVersion", "BuildIdentity", "PublishedAtUTC", "LoadedAtUTC")
        If VarType(publication(field)) <> vbString Then Exit Function
    Next field
    If Not modEvaluationModel.IsInteger(publication("SchemaVersion")) Or Not modEvaluationModel.IsInteger(publication("PolicyVersion")) Then Exit Function
    If publication("Availability") = "Unavailable" Then
        If publication("SchemaVersion") <> 0 Or publication("PolicyVersion") <> 0 Or TypeName(publication("Coverage")) <> "Dictionary" Then Exit Function
        For Each field In Array("WarehouseId", "PublicationId", "ContentSha256", "PackageSetVersion", "BuildIdentity", "PublishedAtUTC", "LoadedAtUTC")
            If publication(field) <> "" Then Exit Function
        Next field
        ValidPublication = (publication("Coverage").Count = 0): Exit Function
    End If
    If publication("Availability") <> "Loaded" And publication("Availability") <> "Stale" Then Exit Function
    If publication("WarehouseId") <> warehouseId Or publication("SchemaVersion") <> 1 Or publication("PolicyVersion") < 0 Then Exit Function
    If Not modTrainingWire.ValidId(CStr(publication("PublicationId"))) Or Not modEvaluationModel.IsHash(CStr(publication("ContentSha256"))) Then Exit Function
    If Not modTrainingWire.ValidUtcTimestamp(CStr(publication("PublishedAtUTC"))) Or Not modTrainingWire.ValidUtcTimestamp(CStr(publication("LoadedAtUTC"))) Then Exit Function
    If publication("BuildIdentity") = "" Or publication("PackageSetVersion") = "" Then Exit Function
    If Not modEvaluationModel.HasFields(publication("Coverage"), "Sources") Then Exit Function
    If TypeName(publication("Coverage")("Sources")) <> "Collection" Or publication("Coverage")("Sources").Count <> 5 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each source In publication("Coverage")("Sources")
        If Not modEvaluationModel.HasFields(source, "Source|Scope|Availability|Explanation|AvailableGroups|IncludedGroups|OmittedGroups|AvailableLines|IncludedLines|OmittedLines|IncludedFrom|IncludedTo") Then Exit Function
        For Each field In Array("Source", "Scope", "Availability", "Explanation", "IncludedFrom", "IncludedTo")
            If VarType(source(field)) <> vbString Then Exit Function
        Next field
        If InStr("|Inventory|Designs|Activity|ShippingBOM|ShippingHolds|", "|" & CStr(source("Source")) & "|") = 0 Or source("Source") = "" Then Exit Function
        If seen.Exists(source("Source")) Then Exit Function
        seen.Add source("Source"), True
        If source("Scope") <> IIf(source("Source") = "ShippingHolds", "Station profile", "Warehouse") Then Exit Function
        For Each field In Array("AvailableGroups", "IncludedGroups", "OmittedGroups", "AvailableLines", "IncludedLines", "OmittedLines")
            If source("Availability") = "Available" Then
                If Not modEvaluationModel.IsInteger(source(field)) Or source(field) < 0 Then Exit Function
            ElseIf source("Availability") = "Unavailable" Then
                If VarType(source(field)) <> vbString Or source(field) <> "Unavailable" Then Exit Function
            Else: Exit Function
            End If
        Next field
        If source("Availability") = "Available" Then
            For Each kind In Array("Groups", "Lines")
                If source("Available" & kind) <> source("Included" & kind) + source("Omitted" & kind) Then Exit Function
            Next kind
        End If
    Next source
    ValidPublication = True
Invalid:
End Function
