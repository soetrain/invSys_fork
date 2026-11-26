Attribute VB_Name = "modAdminSeedObservation"
Option Explicit
Option Private Module

' Invocation-local optional observations; the existing owner determines effects.
Public Function Seed(ByVal warehouseId As String, ByVal stationId As String, _
                     ByVal userId As String, ByVal runtimeRoot As String, _
                     ByVal csvPath As String, ByRef report As String, Optional ByVal replayContext As String = "") As Boolean
    Dim action As String, notice As String, outcome As String, eventId As String
    Dim surface As Object, number As Long, source As String, description As String
    On Error GoTo Failed
    Set surface = modAdminActionObservation.CurrentSurface()
    action = BeginSeed(warehouseId, stationId, userId, runtimeRoot, notice)
    modAdminActionObservation.RestoreSurface surface
    If replayContext <> "" Then
        If action = "" Or modActivity.CaptureContext() <> replayContext Then
            FinishSeed action, "FAILED", "", notice
            Exit Function
        End If
    End If
    If csvPath = "" Then
        Seed = modAdminInventorySeed.SeedDemoInventoryForWarehouse(warehouseId, stationId, userId, report, outcome, eventId)
    Else
        Seed = modAdminInventorySeed.UploadDemoInventoryForWarehouse(warehouseId, stationId, userId, csvPath, report, outcome, eventId)
    End If
    Set surface = modAdminActionObservation.CurrentSurface()
    FinishSeed action, outcome, eventId, notice
    modAdminActionObservation.RestoreSurface surface
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    Set surface = modAdminActionObservation.CurrentSurface()
    FinishSeed action, "FAILED", eventId, notice
    modAdminActionObservation.RestoreSurface surface
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Err.Raise number, source, description
End Function

Private Function BeginSeed(ByVal warehouseId As String, ByVal stationId As String, _
                           ByVal userId As String, ByVal runtimeRoot As String, ByRef notice As String) As String
    Dim context As String
    On Error GoTo Unavailable
    context = modActivity.CaptureContext()
    If context = "" Then Exit Function
    If Not modActivity.MatchesSelection(warehouseId, stationId, userId, runtimeRoot) Then GoTo Unavailable
    BeginSeed = modActivity.BeginAction("ADMIN_DEMO_SEED", context, notice)
    If context <> modActivity.CaptureContext() Then GoTo Unavailable
    Exit Function
Unavailable:
    notice = "Tracking unavailable: Seed could not be recorded for the captured warehouse and user."
End Function

Private Sub FinishSeed(ByVal action As String, ByVal outcome As String, ByVal eventId As String, ByRef notice As String)
    Dim references As String, resultNotice As String
    On Error GoTo Unavailable
    If action = "" Then Exit Sub
    references = "[]"
    If eventId <> "" Then references = modActivity.InventorySourceReferences(action, eventId, "Submitted")
    modActivity.FinishAction action, outcome, resultNotice, references
    If resultNotice <> "" Then notice = resultNotice
    Exit Sub
Unavailable:
    notice = "Tracking unavailable: the Seed result could not be recorded."
End Sub
