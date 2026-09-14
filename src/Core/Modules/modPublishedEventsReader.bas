Attribute VB_Name = "modPublishedEventsReader"
Option Explicit
Option Private Module

' D18: projection/config reads only; no owner workbook or publication command.
Public Function ReadCurrent() As String
    Dim context As String, target As WarehouseTarget, model As Object, policy As Object
    Dim version As Long, collect As Boolean, visible As Boolean, notice As String
    Dim allowed As Object, entry As Object, definition As Object, group As Object, row As Object
    Dim lines As New Collection, fields As Object, ids As Variant, header As Variant
    Dim loaded As String, coverage As String, hidden As Long, values() As String, i As Long
    On Error GoTo Unavailable
    ReadCurrent = "FAIL" & vbTab & "Published Events unavailable. Select a valid signed-in warehouse and compatible publication."
    context = modActivity.CaptureContext()
    If context = "" Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    Set policy = CreateObject("Scripting.Dictionary")
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice, policy) Then Exit Function
    Set model = modEventsPublicationStore.Read(target.RuntimeRoot & "\" & target.WarehouseId & ".invSys.Snapshot.Events.json", target.WarehouseId)
    If model Is Nothing Then Exit Function
    Set allowed = CreateObject("Scripting.Dictionary"): allowed.CompareMode = vbBinaryCompare
    For Each entry In policy("Controls")
        Set definition = modActivityCatalog.Control(CStr(entry("ControlId")), CLng(policy("SavedCatalogVersion")))
        visible = CBool(entry("Visible"))
        If definition Is Nothing Then
            visible = False
        ElseIf definition("Role") = "Admin" Then
            visible = visible And CBool(policy("AdminViewerEventLoggingEnabled"))
        End If
        allowed.Add CStr(entry("ControlId")), visible
    Next entry
    For Each group In model("Groups")
        visible = True
        If group("Source") = "Activity" Then
            For Each row In group("Lines")
                If Not allowed.Exists(CStr(row("ControlId"))) Then
                    visible = False
                ElseIf Not CBool(allowed(CStr(row("ControlId")))) Then
                    visible = False
                End If
            Next row
        End If
        If Not visible Then hidden = hidden + 1
    Next group
    coverage = CoverageText(model) & " Hidden by policy: " & CStr(hidden) & " activity group(s)."
    loaded = modTrainingWire.UtcTimestamp()
    Set fields = modEventDetailCatalog.Fields(): ids = fields.Keys
    For Each group In model("Groups")
        visible = True
        If group("Source") = "Activity" Then
            For Each row In group("Lines")
                If Not allowed.Exists(CStr(row("ControlId"))) Then
                    visible = False
                ElseIf Not CBool(allowed(CStr(row("ControlId")))) Then
                    visible = False
                End If
            Next row
        End If
        If visible Then
            For Each row In group("Lines")
                lines.Add RenderLine(group, row, model, ids, loaded, coverage)
            Next row
        End If
    Next group
    For Each group In model("CurrentState")
        For Each row In group("Lines")
            lines.Add RenderLine(group, row, model, ids, loaded, coverage)
        Next row
    Next group
    If context <> modActivity.CaptureContext() Then Exit Function
    header = Array("OK", target.WarehouseId, CStr(model("PublishedAtUTC")), CStr(lines.Count), "EVENTS1", loaded, _
                   CStr(model("PublicationId")), CStr(version), coverage, Join(ids, ","), CStr(model("PackageSetVersion")), CStr(model("BuildIdentity")))
    For i = 0 To UBound(header): header(i) = Escape(CStr(header(i))): Next i
    ReDim values(0 To lines.Count): values(0) = Join(header, vbTab)
    For i = 1 To lines.Count: values(i) = CStr(lines(i)): Next i
    ReadCurrent = Join(values, vbCrLf)
Unavailable:
End Function

Private Function RenderLine(ByVal group As Object, ByVal row As Object, ByVal model As Object, _
                            ByVal ids As Variant, ByVal loaded As String, ByVal coverage As String) As String
    Dim detail As Object, source As String, kind As String, code As String, label As String, family As String
    Dim values() As String, i As Long
    Set detail = CreateObject("Scripting.Dictionary"): detail.CompareMode = vbBinaryCompare
    For i = LBound(ids) To UBound(ids): detail.Add CStr(ids(i)), "": Next i
    source = Text(group, "Source"): kind = Text(group, "SourceKind")
    code = Text(row, "EventType"): label = FriendlyType(code): family = EventFamily(code)
    detail("SourceId") = Text(group, "SourceId")
    detail("SourceKind") = kind: detail("WarehouseId") = CStr(model("WarehouseId"))
    detail("OccurredAt") = DisplayTime(Text(row, "OccurredAtUTC"))
    detail("AppliedAt") = DisplayTime(Text(row, "AppliedAtUTC"))
    detail("TimeProvenance") = Text(group, "TimeProvenance")
    detail("Coverage") = coverage: detail("PublishedAt") = DisplayTime(CStr(model("PublishedAtUTC")))
    detail("LoadedAt") = DisplayTime(loaded): detail("Freshness") = "Loaded published projection. See Published and Loaded times."
    detail("System_Key") = Text(row, "System_Key"): detail("Reference") = Text(row, "Reference")
    detail("ItemCode") = Text(row, "SKU"): detail("ItemName") = Text(row, "Item")
    detail("Quantity") = Text(row, "QtyDelta"): detail("Uom") = Text(row, "Uom")
    detail("Location") = Text(row, "Location"): detail("Condition") = Text(row, "Condition")
    detail("ActorId") = Text(row, "UserId"): detail("StationId") = Text(row, "StationId")
    detail("UndoEventId") = Text(row, "UndoOfEventId"): detail("DataEffect") = "Unknown"
    detail("Explanation") = "Published owner evidence. A workflow conclusion requires its complete owning outcome evidence."
    detail("NextStep") = "Inspect the owning workflow for completion evidence."
    Select Case source
        Case "Inventory": detail("OwnerId") = "Inventory Domain"
        Case "Designs"
            family = "Designs": detail("OwnerId") = "Designs Domain"
            detail("Reference") = Text(row, "DefinitionId")
            detail("ItemName") = Text(row, "DefinitionType")
            If Text(row, "DefinitionType") = "PROCESS" Then
                detail("ProcessId") = Text(row, "DefinitionId"): detail("ProcessVersion") = Text(row, "DefinitionVersion")
            ElseIf Text(row, "DefinitionType") = "RECIPE" Then
                detail("RecipeId") = Text(row, "DefinitionId"): detail("RecipeVersion") = Text(row, "DefinitionVersion")
            End If
        Case "Activity"
            code = Text(row, "EventCode"): label = Text(row, "Caption"): family = Text(row, "SourceRole")
            detail("OwnerId") = Text(row, "OwnerId"): detail("Severity") = Text(row, "Severity")
            detail("DataEffect") = Text(row, "DataEffect"): detail("Outcome") = Text(row, "OutcomeCode")
            detail("SourceRole") = family: detail("Explanation") = Text(row, "UserMessage")
            detail("NextStep") = Text(row, "NextStep"): detail("ItemName") = Text(row, "Surface")
        Case "ShippingBOM"
            family = "Boxing": code = "BOX_DESIGNED": label = code
            detail("System_Key") = Text(row, "ComponentSystemKey")
            detail("Reference") = Text(row, "BomVersionLabel")
            detail("BomId") = Text(row, "PackageSystemKey")
            detail("Quantity") = Text(row, "ComponentQty"): detail("Uom") = Text(row, "ComponentUOM")
            detail("ItemCode") = Text(row, "ComponentItemCode"): detail("ItemName") = Text(row, "ComponentItem")
            detail("Location") = Text(row, "ComponentLocation")
            detail("BomVersion") = Text(row, "BomVersion")
            detail("OwnerId") = "Shipping workflow"
        Case "ShippingHolds"
            family = "Shipping": code = "SHIP_HELD": label = code
            detail("Reference") = Text(row, "Ref"): detail("ItemName") = Text(row, "Item")
            detail("Quantity") = Text(row, "Qty"): detail("Uom") = Text(row, "UOM")
            detail("OwnerId") = "Shipping workflow"
    End Select
    If kind = "Current state" Then
        detail("Explanation") = "Current-state supplement. Historical control usage and completed workflow outcomes are unavailable."
        detail("OccurredAt") = "": detail("AppliedAt") = "": detail("SourceId") = ""
    End If
    detail("EventType") = label: detail("EventCode") = code: detail("EventFamily") = family
    ReDim values(0 To 18 + UBound(ids))
    values(0) = DisplayTime(Text(group, "RecordedAt")): values(1) = label
    values(2) = detail("Reference"): values(3) = detail("ItemName"): values(4) = detail("Quantity")
    values(5) = detail("Uom"): values(6) = detail("Location"): values(7) = detail("Condition")
    values(8) = detail("ActorId"): values(9) = detail("Explanation")
    values(10) = detail("SourceId"): values(11) = detail("System_Key"): values(12) = code
    values(13) = detail("ItemCode"): values(14) = detail("StationId")
    values(15) = detail("OccurredAt"): values(16) = detail("AppliedAt"): values(17) = source
    If source = "ShippingBOM" Then
        ' Package summary and component detail are distinct published views.
        values(3) = Text(row, "PackageItem"): values(4) = ""
        values(5) = Text(row, "PackageUOM"): values(6) = Text(row, "PackageLocation")
    End If
    For i = 0 To UBound(ids): values(18 + i) = CStr(detail(CStr(ids(i)))): Next i
    For i = 0 To UBound(values): values(i) = Escape(values(i)): Next i
    RenderLine = Join(values, vbTab)
End Function

Private Function Text(ByVal row As Object, ByVal field As String) As String
    If row.Exists(field) Then
        If VarType(row(field)) = vbString Then Text = CStr(row(field))
    End If
End Function

Private Function DisplayTime(ByVal value As String) As String
    DisplayTime = Replace$(Replace$(value, "T", " "), "Z", " UTC")
End Function

Private Function Escape(ByVal value As String) As String
    Escape = Replace$(Replace$(Replace$(Replace$(value, "\", "\\"), vbTab, "\t"), vbCr, "\r"), vbLf, "\n")
End Function

Private Function CoverageText(ByVal model As Object) As String
    Dim source As Object, result As String
    For Each source In model("Coverage")("Sources")
        If result <> "" Then result = result & "; "
        result = result & CStr(source("Source")) & " (" & CStr(source("Scope")) & "): " & CStr(source("Availability"))
        If source("Availability") = "Available" Then
            result = result & ", " & CStr(source("IncludedGroups")) & "/" & CStr(source("AvailableGroups")) & " groups, " & _
                     CStr(source("IncludedLines")) & "/" & CStr(source("AvailableLines")) & " lines; omitted " & _
                     CStr(source("OmittedGroups")) & " groups / " & CStr(source("OmittedLines")) & " lines"
        End If
    Next source
    CoverageText = result
End Function

Private Function EventFamily(ByVal code As String) As String
    Select Case UCase$(code)
        Case "RECEIVE", "RETURN", "DUMP": EventFamily = "Receiving"
        Case "SHIP", "SHIP_RESERVE", "SHIP_RELEASE": EventFamily = "Shipping"
        Case "BOX_BUILD", "BOX_UNBOX": EventFamily = "Boxing"
        Case "PROD_CONSUME", "PROD_COMPLETE": EventFamily = "Production"
        Case Else: EventFamily = "Inventory"
    End Select
End Function

Private Function FriendlyType(ByVal code As String) As String
    Select Case UCase$(code)
        Case "RECEIVE": FriendlyType = "Receipt"
        Case "RETURN": FriendlyType = "Return"
        Case "DUMP": FriendlyType = "Dump"
        Case "BOX_BUILD": FriendlyType = "Box Made"
        Case "BOX_UNBOX": FriendlyType = "Box Unboxed"
        Case "SHIP": FriendlyType = "Shipped"
        Case "SHIP_RESERVE": FriendlyType = "Inventory Reserved"
        Case "SHIP_RELEASE": FriendlyType = "Remove"
        Case "PROD_CONSUME": FriendlyType = "Production Input Consumed"
        Case "PROD_COMPLETE": FriendlyType = "Production Output Created"
        Case "ADMIN_INVENTORY_ADJUST": FriendlyType = "Inventory Adjustment"
        Case Else: FriendlyType = Replace$(code, "_", " ")
    End Select
End Function
