Attribute VB_Name = "modAdminSeedReplay"
Option Explicit

' Fixed Operations bridge. All target and actor values come from the bound session.
Public Function ApplyBuiltIn(ByVal token As String, ByVal stepId As String, ByVal kit As String) As Boolean
    Dim target As WarehouseTarget, choices As Collection, picker As frmSeedInventory
    Dim actor As String, context As String, report As String
    On Error GoTo Done
    If kit <> modAdminInventorySeed.R1_DATA_SET_ID Then Exit Function
    If Not modExecutionRun.PendingSeed(token, stepId, kit, False) Then Exit Function
    context = modActivity.CaptureContext(): actor = modAuth.GetCurrentUserId()
    Set target = modNasConnection.GetCurrentTarget()
    If Not modAuth.CanPerform("ADMIN_MAINT", actor, target.WarehouseId, target.StationId) Then Exit Function
    Set choices = modAdminConsole.GetWarehouseDirectoryOptions(Nothing, report, True)
    If choices Is Nothing Then Exit Function
    Set picker = New frmSeedInventory
    picker.Configure choices, target.WarehouseId, target.StationId, actor
    If Not picker.SelectBoundKit(target.StationId, kit) Then GoTo Done
    If Not modActivity.MatchesSelection(picker.SelectedWarehouseId, picker.SelectedStationId, picker.SelectedUserId, picker.SelectedRuntimeRoot) Then GoTo Done
    If context <> modActivity.CaptureContext() Then GoTo Done
    modRuntimeWorkbooks.SetCoreDataRootOverride target.RuntimeRoot
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then GoTo Done
    If Not modExecutionRun.PendingSeed(token, stepId, kit, True) Then GoTo Done
    ApplyBuiltIn = modAdminSeedObservation.Seed(target.WarehouseId, target.StationId, actor, target.RuntimeRoot, "", report, context)
Done:
    If Not picker Is Nothing Then Unload picker
End Function
