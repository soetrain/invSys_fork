# Test-only typed probes. The callback, writer and processor are not replaced.
# Generated fixtures only; observations expose counts/status/booleans, never keys.
. (Join-Path $PSScriptRoot 'Slice4beAdminSeedFeedback.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminSeedWorkflow.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminSeedRecording.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminSeedBoundaries.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminSeedGuide.ps1')
function Install-AdminSeedOutcomeProbe {
    $probe=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Add(1)
    $probe.Name='TestSeedOutcome'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mInbox As Workbook
Private mOtherEvent As String
Private mSeedEvent As String
Private mWarehouse As String
Private mPriorEvents As Object

Public Function Prepare(ByVal warehouse As String, ByVal pendingOther As Boolean) As Boolean
    Dim report As String, item As Object, payload As String, lo As ListObject, row As Long, id As String
    mWarehouse = warehouse: mOtherEvent = "": mSeedEvent = ""
    Set mInbox = modRoleEventWriter.OpenInboxWorkbook("INVENTORY_CREATE", warehouse, "S1", report)
    If mInbox Is Nothing Then Exit Function
    Set mPriorEvents = CreateObject("Scripting.Dictionary"): mPriorEvents.CompareMode = vbBinaryCompare
    Set lo = FindTable(mInbox, "tblInboxProd")
    For row = 1 To lo.ListRows.Count
        id = CStr(Value(lo, row, "EventID"))
        If id <> "" Then mPriorEvents(id) = True
    Next row
    If pendingOther Then
        Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
            modRoleEventWriter.CreateSystemKey(), "OUTCOME-OTHER", 7, "OUTCOME", "GOOD")
        item("ITEM_CODE") = "OUTCOME-OTHER": item("ITEM") = "Independent event"
        item("UOM") = "each"
        payload = modRoleEventWriter.BuildPayloadJson(item)
        If Not modRoleEventWriter.QueueInventoryCreateEvent(warehouse, "S1", "config-admin", _
            payload, "Independent test event", 0, mInbox, mOtherEvent, report) Then Exit Function
    End If
    Prepare = True
End Function

Public Function Invoke(Optional ByVal action As String = "SEED") As Boolean
    Dim value As String
    value = modAdmin.RunDemoInventoryActionCallbackForAutomation(mWarehouse, "S1", "config-admin", action, "", True)
    Invoke = (Left$(value, 3) = "OK|")
End Function

Private Function FindTable(ByVal wb As Workbook, ByVal name As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = name Then Set FindTable = lo: Exit Function
        Next lo
    Next ws
    Err.Raise 5, , "Seed outcome fixture table unavailable"
End Function

Private Function Value(ByVal lo As ListObject, ByVal row As Long, ByVal name As String) As Variant
    Value = lo.ListColumns(name).DataBodyRange.Cells(row, 1).Value2
End Function

Private Function Applied(ByVal lo As ListObject, ByVal id As String) As Boolean
    Dim row As Long
    If id = "" Then Exit Function
    For row = 1 To lo.ListRows.Count
        If StrComp(CStr(Value(lo, row, "EventID")), id, vbBinaryCompare) = 0 Then Applied = True: Exit Function
    Next row
End Function

Public Function State() As String
    Dim inventory As Workbook, lo As ListObject, history As ListObject
    Dim row As Long, eventId As String, status As String, seedRows As Long
    Dim seedCount As Long, activeDemo As Long, otherQty As Double, keys As Object, key As String, validKeys As Boolean
    Set lo = FindTable(mInbox, "tblInboxProd")
    For row = 1 To lo.ListRows.Count
        eventId = CStr(Value(lo, row, "EventID"))
        If eventId <> "" And eventId <> mOtherEvent And Not mPriorEvents.Exists(eventId) Then
            mSeedEvent = eventId: status = CStr(Value(lo, row, "Status")): seedRows = seedRows + 1
        End If
    Next row
    Set inventory = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(mWarehouse)
    If inventory Is Nothing Then Err.Raise 5, , "Seed outcome inventory unavailable"
    Set history = FindTable(inventory, "tblAppliedEvents")
    Set lo = FindTable(inventory, "tblInventoryEntities")
    Set keys = CreateObject("Scripting.Dictionary"): keys.CompareMode = vbBinaryCompare
    validKeys = True
    For row = 1 To lo.ListRows.Count
        If CStr(Value(lo, row, "SKU")) <> "" Then
            key = CStr(Value(lo, row, "System_Key"))
            If key = "" Or keys.Exists(key) Then validKeys = False Else keys.Add key, True
            If Left$(CStr(Value(lo, row, "SKU")), 5) = "DEMO-" Then
                seedCount = seedCount + 1
                If CDbl(Value(lo, row, "QtyOnHand")) > 0 Then activeDemo = activeDemo + 1
            End If
            If CStr(Value(lo, row, "SKU")) = "OUTCOME-OTHER" Then otherQty = CDbl(Value(lo, row, "QtyOnHand"))
        End If
    Next row
    State = CStr(seedRows) & "|" & status & "|" & CStr(Applied(history, mSeedEvent)) & "|" & _
        CStr(Applied(history, mOtherEvent)) & "|" & CStr(seedCount) & "|" & CStr(otherQty) & "|" & CStr(validKeys) & "|" & CStr(activeDemo)
End Function

Public Function ApplyPending() As Long
    Dim report As String
    ApplyPending = modProcessor.RunBatch(mWarehouse, 1, report)
End Function



Public Sub Release()
    If Not mInbox Is Nothing Then mInbox.Close SaveChanges:=True
    Set mInbox = Nothing
End Sub
'@)
    if($Phase -ceq 'GREEN'){
        $probe.CodeModule.AddFromString(@'
Public Function ReceiptReadChecks() As String
    Dim wb As Workbook, ws As Worksheet, confirmed As Boolean, wrong As Boolean, missing As Boolean
    Dim dirtyDenied As Boolean, preserved As Boolean, report As String, stale As Boolean, processed As Long
    Set wb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(mWarehouse)
    Set ws = wb.Worksheets.Add: ws.Name = "ReceiptReadCanary"
    ws.Cells(1, 1).Value2 = "Saved custom value": wb.Save
    confirmed = modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    wrong = modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent & " ")
    missing = modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, "missing-event")
    ws.Cells(1, 1).Value2 = "Unsaved custom value"
    dirtyDenied = Not modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    preserved = (Not wb.Saved And ws.Cells(1, 1).Value2 = "Unsaved custom value")
    ws.Cells(1, 1).Value2 = "Saved custom value": wb.Save
    stale = True
    processed = modProcessor.RunBatch(mWarehouse, 1, report, "missing-event", stale)
    ReceiptReadChecks = CStr(confirmed) & "|" & CStr(Not wrong) & "|" & CStr(Not missing) & "|" & CStr(dirtyDenied) & "|" & CStr(preserved) & "|" & CStr(Not stale)
End Function
'@)
    }
    if($AdminSeedNativeFeedback){
        Install-AdminSeedFeedbackProbe $probe.CodeModule
    }
    if($AdminSeedWorkflowEdges -or $AdminSeedRecording){Install-AdminSeedWorkflowProbe $probe.CodeModule}
    if($AdminSeedRecording){Install-AdminSeedRecordingProbe $probe.CodeModule}
    if($AdminSeedRecordingBoundaries){
        Install-AdminSeedBoundariesProbe $probe.CodeModule
    }
    if($AdminSeedGuidePaths){Install-AdminSeedGuideProbe}
    if($AdminSeedReplayProfile){
        . (Join-Path $PSScriptRoot 'Slice4beGuideTransfer.ps1')
        Install-GuideTransferDialogProbe
        if($AdminSeedReplayRun){
            . (Join-Path $PSScriptRoot 'Slice4beAdminSeedReplay.ps1')
            Install-SeedReplayGuardProbe
        }
    }
}

function Test-AdminSeedOutcome($Fixture,$Other) {
    if($AdminSeedReplayProfile){
        . (Join-Path $PSScriptRoot 'Slice4beAdminSeedReplay.ps1')
        Test-AdminSeedReplayProfile $Fixture $Other
        return
    }
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestSeedOutcome.'+$Name) $Values}
    $observations=[Collections.Generic.List[object]]::new()
    $nativeObservations=[Collections.Generic.List[object]]::new()
    foreach($case in @('EmptyBatch','EarlierEvent')){
        $target=if($case -ceq 'EmptyBatch'){$Fixture}else{$Other}
        SelectTarget $target
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','1',$target.Warehouse,'S1'))){throw 'Batch size fixture unavailable; not product RED.'}
        Check ('SeedOutcome.'+$case+'.BatchSizeOne') ([long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('BatchSize',0)) -eq 1)
        if(-not [bool](Probe 'Prepare' @($target.Warehouse,($case -ceq 'EarlierEvent')))){throw 'Seed queue fixture unavailable; not product RED.'}
        try {
            $before=([string](Probe 'State')).Split('|')
            $prerequisite=($before[0] -ceq '0' -and $before[4] -ceq '3')
            Check ('SeedOutcome.'+$case+'.NoSeedBeforeCallback') $prerequisite
            if(-not $prerequisite){throw 'Expected three bootstrap entities and no new Seed event; not product RED.'}
            if($AdminSeedNativeFeedback -and $case -ceq 'EarlierEvent'){
                $native=Invoke-AdminSeedFeedback 'SEED';$nativeObservations.Add($native);$ack=$native.Acknowledged
            }else{$ack=[bool](Probe 'Invoke')}
            $state=([string](Probe 'State')).Split('|')
            Check ('SeedOutcome.'+$case+'.ExactlyOneSeedQueued') ($state[0] -ceq '1')
            Check ('SeedOutcome.'+$case+'.ExactUniqueKeys') ($state[6] -ceq 'True')
            if($case -ceq 'EmptyBatch'){
                Check 'SeedOutcome.EmptyBatch.SuccessOnlyAfterOwnApplication' ($ack -and $state[2] -ceq 'True' -and $state[4] -ceq '24')
            }else{
                Check 'SeedOutcome.EarlierEvent.AppliedIndependently' ($state[3] -ceq 'True' -and $state[5] -ceq '7')
                Check 'SeedOutcome.EarlierEvent.OwnSeedStillPending' ($state[1] -ceq 'NEW' -and $state[2] -ceq 'False' -and $state[4] -ceq $before[4])
                Check 'SeedOutcome.EarlierEvent.NoFalseSeedSuccess' (-not $ack)
                $applied=[int](Probe 'ApplyPending')
                $after=([string](Probe 'State')).Split('|')
                Check 'SeedOutcome.EarlierEvent.PendingSeedIsValid' ($applied -eq 1 -and $after[2] -ceq 'True' -and $after[4] -ceq '24' -and $after[5] -ceq '7' -and $after[6] -ceq 'True')
            }
            $observations.Add([pscustomobject]@{Case=$case;CallbackAcknowledged=$ack;OwnEventApplied=($state[2] -ceq 'True');OtherEventApplied=($state[3] -ceq 'True');DemoEntities=[int]$state[4]})
        }finally{[void](Probe 'Release')}
        if(-not [bool](Probe 'Prepare' @($target.Warehouse,($case -ceq 'EarlierEvent')))){throw 'Delete queue fixture unavailable; not product RED.'}
        try {
            $before=([string](Probe 'State')).Split('|')
            if($before[0] -cne '0' -or $before[7] -cne '24'){throw 'Delete requires all 24 active demo entities; not product RED.'}
            Check ('DeleteOutcome.'+$case+'.ActiveDemoPrecondition') $true
            if($AdminSeedNativeFeedback -and $case -ceq 'EarlierEvent'){
                $native=Invoke-AdminSeedFeedback 'DELETE';$nativeObservations.Add($native);$ack=$native.Acknowledged
            }else{$ack=[bool](Probe 'Invoke' @('DELETE'))}
            $state=([string](Probe 'State')).Split('|')
            Check ('DeleteOutcome.'+$case+'.ExactlyOneDeleteQueued') ($state[0] -ceq '1')
            Check ('DeleteOutcome.'+$case+'.HistoryAndUniqueKeysRetained') ($state[4] -ceq '24' -and $state[6] -ceq 'True')
            if($case -ceq 'EmptyBatch'){
                Check 'DeleteOutcome.EmptyBatch.SuccessOnlyAfterOwnApplication' ($ack -and $state[2] -ceq 'True' -and $state[7] -ceq '0')
            }else{
                Check 'DeleteOutcome.EarlierEvent.AppliedIndependently' ($state[3] -ceq 'True' -and $state[5] -ceq '7')
                Check 'DeleteOutcome.EarlierEvent.OwnDeleteStillPending' ($state[1] -ceq 'NEW' -and $state[2] -ceq 'False' -and $state[7] -ceq '24')
                Check 'DeleteOutcome.EarlierEvent.NoFalseDeleteSuccess' (-not $ack)
                $applied=[int](Probe 'ApplyPending')
                $after=([string](Probe 'State')).Split('|')
                Check 'DeleteOutcome.EarlierEvent.PendingDeleteIsValid' ($applied -eq 1 -and $after[2] -ceq 'True' -and $after[4] -ceq '24' -and $after[7] -ceq '0' -and $after[5] -ceq '7')
            }
            $observations.Add([pscustomobject]@{Case=('Delete'+$case);CallbackAcknowledged=$ack;OwnEventApplied=($state[2] -ceq 'True');OtherEventApplied=($state[3] -ceq 'True');DemoEntities=[int]$state[4]})
        }finally{[void](Probe 'Release')}
    }
    if($Phase -ceq 'GREEN'){
        $reads=([string](Probe 'ReceiptReadChecks')).Split('|')
        foreach($entry in @(@('ExactSavedEvent',0),@('NoTrimmedIdentityMatch',1),@('AbsentEventUnconfirmed',2),@('DirtyAuthorityUnconfirmed',3),@('BorrowedEditsPreserved',4),@('ProcessorResetsStaleReceipt',5))){
            Check ('SeedOutcome.ReceiptRead.'+$entry[0]) ($reads[$entry[1]] -ceq 'True')
        }
    }
    $observations|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'seed-outcome-observations.json')
    foreach($native in $nativeObservations){
        $prefix='SeedOutcome.Native.'+$native.Action
        Check ($prefix+'.ExactUnconfirmedMessage') $native.Observed.ExpectedMessage
        Check ($prefix+'.TextFits') ($native.Observed.FontAvailable -and $native.Observed.TextFits -and $native.Observed.WindowWithinDisplay)
        Check ($prefix+'.VisibleCapture') ($native.Observed.Captured -and $native.Observed.ForegroundBefore -and $native.Observed.ForegroundAfter)
        Check ($prefix+'.DismissedWithoutSuccess') ($native.Observed.OkDispatched -and -not $native.Acknowledged)
    }
    if($AdminSeedWorkflowEdges){Test-AdminSeedWorkflow $Other}
    if($AdminSeedRecording){Test-AdminSeedRecording $Fixture}
    if($AdminSeedRecordingBoundaries){Test-AdminSeedBoundaries $Fixture $Other}
    if($AdminSeedGuidePaths){Test-AdminSeedGuide $Fixture $Other}
}
