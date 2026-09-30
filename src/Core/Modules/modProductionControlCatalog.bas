Attribute VB_Name = "modProductionControlCatalog"
Option Explicit
Option Private Module

' Shared record shape; each owning catalog keeps its exact ID/caption mapping.
Public Function Command(ByVal id As String, ByVal owner As String, ByVal caption As String, _
                        ByVal surface As String) As Object
    Dim record As Object
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", owner
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", caption: record.Add "Surface", surface
    record.Add "Capability", "PROD_POST": record.Add "CodePrefix", id & "_"
    Set Command = record
End Function
