Attribute VB_Name = "modProductionRecipeOrderCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("PRODUCTION_RECIPE_MOVE_UP", "PRODUCTION_RECIPE_MOVE_DOWN", "PRODUCTION_RECIPE_AUTO_ORDER")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, record As Object
    Select Case id
        Case "PRODUCTION_RECIPE_MOVE_UP": caption = "Move Up"
        Case "PRODUCTION_RECIPE_MOVE_DOWN": caption = "Move Down"
        Case "PRODUCTION_RECIPE_AUTO_ORDER": caption = "Auto Order"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_DESIGNER"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", caption: record.Add "Surface", "Operations > Production > Recipe Designer"
    record.Add "Capability", "PROD_POST": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object
    Set definition = modProductionRecipeOrderCodes.Control(id)
    If Not definition Is Nothing Then Set Outcome = modProductionLocalEditCodes.Outcome(id, code)
End Function
