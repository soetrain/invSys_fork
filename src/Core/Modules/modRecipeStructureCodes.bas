Attribute VB_Name = "modRecipeStructureCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("PRODUCTION_RECIPE_ADD_PROCESS", "PRODUCTION_RECIPE_REMOVE_PROCESS", _
                       "PRODUCTION_RECIPE_CONNECT", "PRODUCTION_RECIPE_UPDATE_CONNECTION", "PRODUCTION_RECIPE_DISCONNECT")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, record As Object
    Select Case id
        Case "PRODUCTION_RECIPE_ADD_PROCESS": caption = "Add Process"
        Case "PRODUCTION_RECIPE_REMOVE_PROCESS": caption = "Remove Process"
        Case "PRODUCTION_RECIPE_CONNECT": caption = "Connect"
        Case "PRODUCTION_RECIPE_UPDATE_CONNECTION": caption = "Update"
        Case "PRODUCTION_RECIPE_DISCONNECT": caption = "Disconnect"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_DESIGNER"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", caption: record.Add "Surface", "Operations > Production > Recipe Designer"
    record.Add "Capability", "PROD_POST": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function
