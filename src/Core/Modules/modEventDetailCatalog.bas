Attribute VB_Name = "modEventDetailCatalog"
Option Explicit
Option Private Module

' D18 SchemaVersion 1 display vocabulary; these are not canonical event headers.
Public Function Families() As Variant
    Families = Array("Receiving", "Shipping", "Boxing", "Production", "Inventory", "Designs", "Admin", "Session", "Viewer")
End Function

Public Function Fields() As Object
    Dim catalog As Object
    Set catalog = CreateObject("Scripting.Dictionary")
    AddField catalog, "SourceId", "Source event / activity ID", "EXAMPLE-EVENT", True
    AddField catalog, "EventCode", "Event code", "EXAMPLE_CONFIRMED", True
    AddField catalog, "Severity", "Severity", "Notice", True
    AddField catalog, "DataEffect", "Data effect", "Unchanged", True
    AddField catalog, "OwnerId", "Owning operation", "Example workflow", True
    AddField catalog, "EventFamily", "Event family", "Example family", True
    AddField catalog, "EventType", "Event type", "Example action", True
    AddField catalog, "Outcome", "Outcome", "Example confirmed outcome", True
    AddField catalog, "SourceKind", "Source classification", "User activity", True
    AddField catalog, "WarehouseId", "Warehouse", "EXAMPLE-WAREHOUSE", True
    AddField catalog, "OccurredAt", "Occurred / recorded time", "Example timestamp", True
    AddField catalog, "TimeProvenance", "Time provenance", "Verified UTC (synthetic example)", True
    AddField catalog, "Coverage", "Coverage", "Example coverage; no live records", True
    AddField catalog, "PublishedAt", "Published", "Example publication time", True
    AddField catalog, "LoadedAt", "Loaded", "Example load time", True
    AddField catalog, "Freshness", "Freshness", "Synthetic preview only", True
    AddField catalog, "Explanation", "Explanation", "Example observation; no business action was performed.", True
    AddField catalog, "NextStep", "Advisory next step", "Example guidance only; no action is executed.", True
    AddField catalog, "System_Key", "Inventory identity (System_Key)", "EXAMPLE-INVENTORY-IDENTITY", True
    AddField catalog, "Reference", "Reference", "EXAMPLE-REFERENCE", False, True
    AddField catalog, "ParentEventId", "Parent event", "EXAMPLE-PARENT"
    AddField catalog, "UndoEventId", "Undo event", "EXAMPLE-UNDO"
    AddField catalog, "ItemCode", "Item code", "EXAMPLE-ITEM", False, True
    AddField catalog, "ItemName", "Item name", "Example item", False, True
    AddField catalog, "Quantity", "Quantity", "2", False, True
    AddField catalog, "Uom", "Unit of measure", "EA", False, True
    AddField catalog, "Location", "Location", "EXAMPLE-LOCATION", False, True
    AddField catalog, "Condition", "Condition", "GOOD", False, True
    AddField catalog, "SourceRole", "Source role", "Example role", False, True
    AddField catalog, "ActorId", "invSys actor", "EXAMPLE-ACTOR", False, True
    AddField catalog, "StationId", "Station", "EXAMPLE-STATION"
    AddField catalog, "AppliedAt", "Applied time", "Example applied time (UTC)"
    AddField catalog, "RecipeId", "Recipe identity", "EXAMPLE-RECIPE"
    AddField catalog, "RecipeVersion", "Recipe version", "1"
    AddField catalog, "ProcessId", "Process identity", "EXAMPLE-PROCESS"
    AddField catalog, "ProcessVersion", "Process version", "1"
    AddField catalog, "RunId", "Production run", "EXAMPLE-RUN"
    AddField catalog, "ShipmentId", "Shipment reference", "EXAMPLE-SHIPMENT"
    AddField catalog, "BomId", "BOM reference", "EXAMPLE-BOM"
    AddField catalog, "BomVersion", "BOM version", "1"
    AddField catalog, "BusinessReason", "Business reason", "Example sanitized business reason", False, True
    AddField catalog, "BatchNote", "Batch note", "Example permitted business note"
    Set Fields = catalog
End Function

Private Sub AddField(ByVal catalog As Object, ByVal id As String, ByVal caption As String, ByVal sample As String, _
                     Optional ByVal required As Boolean = False, Optional ByVal enabled As Boolean = False)
    Dim definition As Object
    Set definition = CreateObject("Scripting.Dictionary")
    definition.Add "Caption", caption
    definition.Add "Sample", sample
    definition.Add "Required", required
    definition.Add "DefaultEnabled", required Or enabled
    catalog.Add id, definition
End Sub
