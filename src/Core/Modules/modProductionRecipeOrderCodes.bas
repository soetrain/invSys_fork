Attribute VB_Name = "modProductionRecipeOrderCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("PRODUCTION_RECIPE_MOVE_UP", "PRODUCTION_RECIPE_MOVE_DOWN", "PRODUCTION_RECIPE_AUTO_ORDER")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String
    Select Case id
        Case "PRODUCTION_RECIPE_MOVE_UP": caption = "Move Up"
        Case "PRODUCTION_RECIPE_MOVE_DOWN": caption = "Move Down"
        Case "PRODUCTION_RECIPE_AUTO_ORDER": caption = "Auto Order"
        Case Else: Exit Function
    End Select
    Set Control = modProductionControlCatalog.Command(id, "PRODUCTION_DESIGNER", caption, _
                                                     "Operations > Production > Recipe Designer")
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object
    Set definition = modProductionRecipeOrderCodes.Control(id)
    If Not definition Is Nothing Then Set Outcome = modProductionLocalEditCodes.Outcome(id, code)
End Function
