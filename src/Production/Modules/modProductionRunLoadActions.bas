Attribute VB_Name = "modProductionRunLoadActions"
Option Explicit
Option Private Module

Public Function Execute(ByVal owner As frmProduction, ByVal context As String, _
                        ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean, _
                        ByVal recipes As MSForms.ListBox, ByVal scaleText As String) As String
    Dim action As cProductionWorksheetAction, report As String, priorLoading As Boolean
    Dim number As Long, source As String, description As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    priorLoading = loading: busy = True
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_RUN_LOAD", context, operatorBook, report) Then GoTo Done
    If Not action.CanContinue(report) Then GoTo Done
    If recipes.ListIndex < 0 Then
        report = "Select a released Recipe version first."
        action.OutcomeCode = "REJECTED"
        GoTo Done
    End If
    If Stage(owner, owner.SelectedRunRecipeText(0), owner.SelectedRunRecipeText(1), scaleText, report, action) Then _
        action.OutcomeCode = "STAGED"
Done:
    loading = priorLoading
    If Not action Is Nothing Then action.Finish report
    busy = False
    Execute = report
    If number <> 0 Then
        On Error GoTo 0
        Err.Raise number, source, description
    End If
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    Resume Done
End Function

' Direct preparation callers retain the original behavior without an observation.
Public Function Stage(ByVal owner As frmProduction, ByVal recipeId As String, _
                      ByVal recipeVersion As String, ByVal scaleText As String, ByRef report As String, _
                      Optional ByVal action As cProductionWorksheetAction = Nothing) As Boolean
    Dim scalePercent As Double
    If Not CanContinue(action, report) Then Exit Function
    If Not owner.TryParseBatchScalePercent(scaleText, scalePercent, report) Then
        If Not action Is Nothing Then action.OutcomeCode = "REJECTED"
        Exit Function
    End If
    If Not modProductionReusableRun.LoadReleasedReusableRecipe(recipeId, recipeVersion, scalePercent, report, action) Then Exit Function
    If Not CanContinue(action, report) Then Exit Function
    owner.RefreshReusableRunControls False, action
    If Not CanContinue(action, report) Then Exit Function
    Stage = True
End Function

Public Function CanContinue(ByVal action As cProductionWorksheetAction, Optional ByRef report As String = "") As Boolean
    CanContinue = True
    If Not action Is Nothing Then CanContinue = action.CanContinue(report)
End Function
