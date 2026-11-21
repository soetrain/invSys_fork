# Approved D18-LOCAL-01: actual recording/library handlers in each target kind.
# The legacy switch/name is retained; four old local-absence assertions are superseded.
# Only generated fixtures and disposable, unsaved test modules are touched.
function Install-ReconcileLocalDiagnostic {
    Install-AdminReconcileRecordingProbe
    Install-ReconcileGuideProbe
    $project=$packages['invSys.Core.xlam'].VBProject
    $project.VBComponents.Item('modActivityStore').CodeModule.AddFromString(@'
Public Function LocalDiagnosticRootForTest(ByVal target As WarehouseTarget, ByVal create As Boolean) As Boolean
    LocalDiagnosticRootForTest = (StoreRoot(target, create) <> "")
End Function
'@)
    $module=$project.VBComponents.Add(1);$module.Name='TestReconcileLocal'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function PreparePolicyAndRoots() As Boolean
    Dim context As String, version As Long, wire As String, report As String, model As Object
    Dim target As WarehouseTarget
    context = modActivity.CaptureContext()
    If context = "" Then Exit Function
    If Not modTrackingPolicySettings.ReadEditor(context, version, wire, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(True)
    model("ViewerActionPathCaptureEnabled") = True
    wire = modTrainingJson.EncodeObject(model)
    If Not modTrackingPolicySettings.SavePolicy(context, version, wire, report) Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    PreparePolicyAndRoots = modActivityStore.LocalDiagnosticRootForTest(target, True) And _
        (modRecordingJournal.JournalRoot(target, True) <> "")
End Function
Public Function ResolveExplicitLocal(ByVal root As String, ByVal secret As String) As Boolean
    Dim target As WarehouseTarget, status As NasStatusCode, signed As String
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.SetCoreDataRootOverride root
    If Not modNasConnection.ResolveWarehouseTarget(target, status) Then Exit Function
    If status <> NAS_OK Or target.SourceType <> WH_SOURCE_LOCAL Then Exit Function
    If StrComp(target.HubRoot, root, vbTextCompare) <> 0 Or StrComp(target.RuntimeRoot, root, vbTextCompare) <> 0 Then Exit Function
    signed = modAuth.SignInCurrentTargetForAutomation("config-admin", secret, "")
    secret = ""
    ResolveExplicitLocal = (Left$(signed, 3) = "OK|" And modAuth.IsSignedIn())
End Function
Public Function ObserveBoundary() As String
    Dim target As WarehouseTarget, version As Long, collect As Boolean, visible As Boolean
    Dim notice As String, capture As Boolean, eligible As Boolean, policy As Boolean
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    policy = modActivityPolicy.ReadPolicy(target, "ADMIN_RECONCILE_INSPECT", version, collect, visible, notice, Nothing, capture, eligible)
    ObserveBoundary = CStr(target.SourceType) & "|" & CStr(modAuth.IsSignedIn()) & "|" & _
        CStr(modNasConnection.IsCurrentTargetAllowed(False)) & "|" & CStr(modNasConnection.IsCurrentTargetAllowed(True)) & "|" & _
        CStr(modActivity.CaptureContext() <> "") & "|" & CStr(modActivityStore.LocalDiagnosticRootForTest(target, False)) & "|" & _
        CStr(modRecordingJournal.JournalRoot(target, False) <> "") & "|" & CStr(policy And collect And capture And eligible)
End Function
Public Function ReplayTargetEligible() As Boolean
    Dim target As WarehouseTarget, snapshot As String, notice As String
    ReplayTargetEligible = modExecutionTarget.Read(modActivity.CaptureContext(), target, snapshot, notice)
End Function
'@)
}

function Test-ReconcileLocalDiagnostic($Fixture,$Events,$State,$OtherFixture) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    function TrainingPins {
        $result=@{};$root=Join-Path $Fixture.Root 'Training'
        if(Test-Path -LiteralPath $root){Get-ChildItem -LiteralPath $root -Recurse -File|ForEach-Object {$result[$_.FullName]=Hash $_.FullName}}
        return $result
    }
    function Preserved($Before,$After){foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $Before[$path] -cne $After[$path]){return $false}};return $true}
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestReconcileLocal.PreparePolicyAndRoots')){throw 'Local diagnostic policy/root fixture unavailable; not product RED.'}
    $policyHash=Hash $Fixture.Config
    $observations=@();$hostBook=$null
    try{
        $hostBook=$excel.Workbooks.Add();$hostBook.Windows.Item(1).Visible=$true;$excel.Visible=$true
        $hostBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Local recording host canary'
        foreach($case in @('NASFixture','ExplicitLocal','FallbackFixture')){
            CloseRecordingViewer
            if($case -ceq 'ExplicitLocal'){
                if(-not [bool](Run 'invSys.Core.xlam' 'TestReconcileLocal.ResolveExplicitLocal' @($Fixture.Root,$Fixture.Secret))){throw 'Actual explicit local resolution failed; not product RED.'}
            }elseif($case -ceq 'FallbackFixture'){
                if(-not [bool](Run 'invSys.Core.xlam' 'modNasConnection.SetCurrentTargetSourceTypeForTest' @(3))){throw 'Fallback fixture unavailable.'}
            }
            $values=([string](Run 'invSys.Core.xlam' 'TestReconcileLocal.ObserveBoundary')) -split '\|'
            if($values.Count -ne 8){throw 'Local diagnostic envelope unavailable.'}
            $expectedType=if($case -ceq 'NASFixture'){0}elseif($case -ceq 'ExplicitLocal'){2}else{3}
            $nas=$case -ceq 'NASFixture'
            $prefix='Reconcile.LocalDiagnostic.'+$case+'.'
            Check ($prefix+'ExactTargetKindAndLiveBusinessAccess') ($values[0] -ceq [string]$expectedType -and $values[1] -ceq 'True' -and $values[2] -ceq 'True')
            Check ($prefix+'EnabledPolicyStillReadable') ($values[7] -ceq 'True')
            $boundaryName=if($nas){'ObservedNASContextAndStoreBoundary'}else{'ApprovedOwnRuntimeContextAndStores'}
            Check ($prefix+$boundaryName) (($values[3..6] -join '|') -ceq ((@([string]$nas,'True','True','True')) -join '|'))
            if(-not $nas){Check ($prefix+'ReplayTargetRemainsIneligible') (-not [bool](Run 'invSys.Core.xlam' 'TestReconcileLocal.ReplayTargetEligible'))}
            $before=TrainingPins;$inventoryHash=Hash $State.InventoryPath;$rows=$Events.ListRows.Count
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.PrepareAuthorityCase' @('LocalSignedIn'))){throw 'Reconcile callback fixture unavailable.'}
            try{
                $hostBook.Activate();OpenRecordingViewer
                Check ($prefix+'ActualStartHandler') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
                $starts=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
                $sequence='';if($starts.Count -eq 1){$sequence=[string]$starts[0].SequenceId}
                Check ($prefix+'ActualStartCreatesOwnJournal') ($starts.Count -eq 1 -and $starts[0].WarehouseId -ceq $Fixture.Warehouse)
                $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.InvokeCallback')
                $notices=[int](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.TrackingNoticeCount')
                Check ($prefix+'ActualStopHandler') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
                $after=TrainingPins
                Check ($prefix+'ActualCallbackQueuesExactlyOnce') ($called -and $Events.ListRows.Count -eq $rows+1)
                $row=$Events.ListRows.Item($Events.ListRows.Count)
                $items=ConvertFrom-Json -InputObject ([string]$row.Range.Cells.Item(1,$Events.ListColumns.Item('PayloadJson').Index).Value2)
                Check ($prefix+'ExactEntityCorrection') ($items.Count -eq 1 -and $items[0].System_Key -ceq $State.KeyA)
                Check ($prefix+'InventoryAndCustomValuesPreserved') ((Hash $State.InventoryPath) -ceq $inventoryHash -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.CanaryPreserved'))
                Check ($prefix+'PriorTrainingEvidencePreserved') (Preserved $before $after)
                $availabilityName=if($nas){'ObservedRecordingAvailability'}else{'ApprovedLocalActivityCreated'}
                $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {-not $before.ContainsKey($_)}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
                # The deterministic presentation seam invokes the real selection/
                # inspection handlers without showing the modal confirmation form.
                # Active capture adds the Open REQUESTED navigation observation.
                # Inspect emits REQUESTED/CONFIRMED; no unseen OPENED is invented.
                @($records|Group-Object ControlId,OutcomeCode,Ordinal|ForEach-Object {[pscustomobject]@{ControlId=$_.Group[0].ControlId;OutcomeCode=$_.Group[0].OutcomeCode;Ordinal=$_.Group[0].Ordinal;Count=$_.Count}})|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('local-observation-counts-'+$case+'.json'))
                Check ($prefix+$availabilityName) ($records.Count -eq 3)
                $exact=$records.Count -eq 3 -and $sequence -ne ''
                foreach($record in $records){$exact=$exact -and $record.SequenceId -ceq $sequence -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin'}
                foreach($pair in @(@('ADMIN_RECONCILE_OPEN','REQUESTED',1),@('ADMIN_RECONCILE_INSPECT','REQUESTED',2),@('ADMIN_RECONCILE_INSPECT','CONFIRMED',2))){$exact=$exact -and @($records|Where-Object {$_.ControlId -ceq $pair[0] -and $_.OutcomeCode -ceq $pair[1] -and $_.Ordinal -eq $pair[2]}).Count -eq 1}
                $confirmed=@($records|Where-Object OutcomeCode -CEQ 'CONFIRMED')
                if($confirmed.Count -eq 1){$refs=@($confirmed[0].SourceEventRefs);$exact=$exact -and $refs.Count -eq 1;if($refs.Count -eq 1){$exact=$exact -and $refs[0].EventId -ceq [string]$row.Range.Cells.Item(1,$Events.ListColumns.Item('EventID').Index).Value2 -and $refs[0].SubmissionState -ceq 'Submitted'}}else{$exact=$false}
                Check ($prefix+'ExactRecordedActorContextAndControls') $exact
                $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
                # Stop retains the unmatched Open request honestly as incomplete.
                Check ($prefix+'ExactIncompleteJournal') ($closed.Count -eq 1 -and $closed[0].ActionCount -eq 2 -and $closed[0].Lifecycle -ceq 'Incomplete' -and $closed[0].ReasonCode -ceq 'UNFINISHED_ACTIONS' -and @($closed[0].Observations).Count -eq 3 -and $closed[0].ExpectedConclusion.TerminalKind -ceq 'None')
                Check ($prefix+'ImmutableJournalChain') (JournalChain $sequence 5)
                Test-LocalRecordedGuide $Fixture $case $closed $prefix
                Check ($prefix+'PolicyUnchanged') ((Hash $Fixture.Config) -ceq $policyHash)
                Check ($prefix+'GuideReadsPreservePriorEvidenceAndHost') ((Preserved $after (TrainingPins)) -and [string]$hostBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Local recording host canary')
                $observations+=[pscustomobject]@{Case=$case;SourceType=[int]$values[0];SignedIn=[bool]::Parse($values[1]);BusinessAllowed=[bool]::Parse($values[2]);NASAllowed=[bool]::Parse($values[3]);ContextAvailable=[bool]::Parse($values[4]);ActivityStoreAvailable=[bool]::Parse($values[5]);JournalStoreAvailable=[bool]::Parse($values[6]);PolicyEnabled=[bool]::Parse($values[7]);Queued=$Events.ListRows.Count-$rows;NewTrainingFiles=$after.Count-$before.Count;TrackingNotices=$notices;ContractApproved=$true;FullAcceptance=$false}
            }finally{CloseRecordingViewer;[void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RestoreAuthorityFixture')}
        }
        $observations|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'local-recording-diagnostic.json')
        if($Phase -ceq 'GREEN'){Test-ReconcileLocalBoundaries $Fixture $OtherFixture $Events $State $hostBook}
    }finally{CloseRecordingViewer;if($null -ne $hostBook){$hostBook.Close($false)};SelectTarget $Fixture}
}

function Test-LocalRecordedGuide($Fixture,[string]$Case,$Closed,[string]$Prefix) {
    $opened=(BoundLibrary 'Open') -ceq 'DELIVERED'
    $selected=$false;$shown=$false;$saved=$false;$exact=$false;$read=$false
    if($opened -and $Closed.Count -eq 1){
        $selected=(BoundLibrary 'Select' $Closed[0].ActionPathId) -ceq 'SELECTED'
        if($selected){
            $delivered=BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths'
            $shown=$delivered -ceq 'DELIVERED' -and (BoundControl '' 'Count' '' 'frmActionPathGuide') -ceq '1'
            if($shown){
                $name='Local recording '+$Case
                [void](BoundControl 'txtGuideSubject' 'Write' 'Admin reconciliation' 'frmActionPathGuide')
                [void](BoundControl 'txtGuideName' 'Write' $name 'frmActionPathGuide')
                [void](BoundControl 'txtGuideInstructions' 'Write' 'Inspect the selected entity and submit its correction.' 'frmActionPathGuide')
                $saved=(BoundControl 'btnSaveGuide' 'Click' '' 'frmActionPathGuide') -ceq 'DELIVERED'
                $guideRoot=Join-Path $journalRoot 'Guides'
                $guides=@();if(Test-Path -LiteralPath $guideRoot){$guides=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|ForEach-Object {ReadGuideExpectationRecord $_.FullName}|Where-Object {$null -ne $_ -and $_.Name -ceq $name})}
                $exact=$guides.Count -eq 1
                if($exact){
                    $guide=$guides[0]
                    $exact=$guide.WarehouseId -ceq $Fixture.Warehouse -and $guide.SourceRun.RecordId -ceq $Closed[0].RecordId -and $guide.SourceRun.ContentSha256 -ceq $Closed[0].ContentSha256 -and @($guide.Observations).Count -eq 3 -and $guide.ExpectedConclusion.TerminalKind -ceq 'None' -and ($guide.Observations|ConvertTo-Json -Depth 16 -Compress) -ceq ($Closed[0].Observations|ConvertTo-Json -Depth 16 -Compress)
                    if($saved -and $Phase -ceq 'GREEN'){
                        if((BoundControl '' 'Fit' 'Current' 'frmActionPathGuide') -cne 'True'){throw 'Saved guide layout is unavailable for capture.'}
                        $captured=$false
                        try{CaptureOwnedFormByCaptionEvidence 'Action Path guide' ('local-guide-'+$Case.ToLowerInvariant()+'.png');$captured=$true}
                        catch{if($_.Exception.GetBaseException().Message -cne 'No uncovered owned form caption point is available.'){throw}}
                        Check ($Prefix+'SavedGuideVisibleCapture') $captured
                    }
                    $saved=$saved -and (BoundControl 'btnCancelGuide' 'Click' '' 'frmActionPathGuide') -ceq 'DELIVERED'
                    BoundOpen;BoundSelect $guide
                    $read=(BoundControl 'btnEditPublishedGuide' 'Click') -ceq 'DELIVERED' -and (BoundControl 'txtGuideName' 'Text' '' 'frmActionPathGuide') -ceq $name
                }
            }
        }
    }
    Check ($Prefix+'ActualLibrarySelectsOwnRecording') ($opened -and $selected)
    Check ($Prefix+'ActualCreateGuideOpensEditor') $shown
    Check ($Prefix+'ActualSaveCreatesExactOwnGuide') ($saved -and $exact)
    Check ($Prefix+'ActualLibraryReopensSavedGuide') $read
}

# Supplemental retained-boundary checks; actual callback and Viewer handlers.
function Test-ReconcileLocalBoundaries($Fixture,$OtherFixture,$Events,$State,$HostBook) {
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestReconcileRecording.'+$Name) $Values}
    function ResolveLocal {
        CloseRecordingViewer
        if(-not [bool](Run 'invSys.Core.xlam' 'TestReconcileLocal.ResolveExplicitLocal' @($Fixture.Root,$Fixture.Secret))){throw 'Boundary local resolver fixture failed.'}
    }
    function BeginCapture {
        $script:boundaryPrior=BoundPins (Join-Path $Fixture.Root 'Training')
        $HostBook.Activate();OpenRecordingViewer
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Boundary capture start handler unavailable.'}
        $starts=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|Where-Object {-not $script:boundaryPrior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Boundary capture start journal unavailable.'}
        return [string]$starts[0].SequenceId
    }
    $bytes=Hash $State.InventoryPath
    $otherTraining=Join-Path $OtherFixture.Root 'Training'
    ResolveLocal
    foreach($mode in @('Disabled','ControlsOff','UserOff')){
        SelectTarget $Fixture
        if(-not [bool](Probe 'EnablePolicy' @($mode))){throw 'Local suppression policy fixture unavailable.'}
        ResolveLocal
        $before=BoundPins (Join-Path $Fixture.Root 'Training');$other=BoundPins $otherTraining;$rows=$Events.ListRows.Count
        $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RecordingSubmission')
        $prefix='Reconcile.LocalBoundary.'+$mode+'.'
        Check ($prefix+'CommandQueuesOnce') ($called -and $Events.ListRows.Count -eq $rows+1)
        Check ($prefix+'TrainingUnchanged') (BoundSame $before (BoundPins (Join-Path $Fixture.Root 'Training')))
        Check ($prefix+'NoFalseWarning') ([int](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.TrackingNoticeCount') -eq 0)
        Check ($prefix+'InventoryAndOtherRootPreserved') ((Hash $State.InventoryPath) -ceq $bytes -and (BoundSame $other (BoundPins $otherTraining)))
    }
    foreach($target in @('Local','NASFixture')){
        CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Probe 'EnablePolicy' @('Capture'))){throw 'Storage-failure capture policy unavailable.'}
        if($target -ceq 'Local'){ResolveLocal}
        foreach($mode in @('REQUESTED','CONFIRMED')){
            CloseRecordingViewer
            $sequence=BeginCapture;$before=$script:boundaryPrior;$other=BoundPins $otherTraining;$rows=$Events.ListRows.Count
            $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
            Probe 'ArmStoreFailure' @($mode)
            try{
                $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RecordingSubmission')
                $faults=[int](Probe 'StoreFaultCalls')
            }finally{Probe 'ArmStoreFailure' @('')}
            if($faults -eq 0){throw 'Real storage write boundary was not reached.'}
            $prefix='Reconcile.LocalBoundary.'+$target+'.StoreFailure.'+$mode+'.'
            Check ($prefix+'CommandQueuesOnce') ($called -and $Events.ListRows.Count -eq $rows+1 -and $faults -eq $(if($mode -ceq 'REQUESTED'){2}else{1}))
            Check ($prefix+'ExactWarningOnce') ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ExactStoreWarning'))
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
            Check ($prefix+'HonestIncompleteJournal') ($closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Incomplete' -and $closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE' -and $closed[0].ExpectedConclusion.TerminalKind -ceq 'None' -and @($closed[0].Observations|Where-Object OutcomeCode -CNE 'REQUESTED').Count -eq 0)
            Check ($prefix+'OriginalTargetAndOtherRootPreserved') ($context -cne '' -and $context -ceq [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -and (BoundSame $other (BoundPins $otherTraining)))
            Check ($prefix+'PriorHistoryImmutable') (Preserved $before (BoundPins (Join-Path $Fixture.Root 'Training')))
            Check ($prefix+'InventoryPreserved') ((Hash $State.InventoryPath) -ceq $bytes)
        }
    }
    foreach($case in @('SignOutDuringNarrative','SwitchTargetDuringNarrative')){
        CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Probe 'EnablePolicy' @('Capture'))){throw 'Context-failure capture policy unavailable.'}
        ResolveLocal
        [void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ConfigureAuthorityMutations' @($Fixture.Secret,$OtherFixture.Root,$OtherFixture.Secret))
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.PrepareAuthorityCase' @($case))){throw 'Context mutation fixture unavailable.'}
        $sequence=BeginCapture;$before=$script:boundaryPrior;$other=BoundPins $otherTraining;$rows=$Events.ListRows.Count
        try{
            $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.InvokeCallback')
            $prefix='Reconcile.LocalBoundary.'+$case+'.'
            Check ($prefix+'LiveMutationObserved') ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.AuthorityMutationObserved'))
            Check ($prefix+'NoBusinessQueue') (-not $called -and $Events.ListRows.Count -eq $rows)
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
            Check ($prefix+'OriginalJournalInterrupted') ($closed.Count -eq 1 -and $closed[0].WarehouseId -ceq $Fixture.Warehouse -and $closed[0].Lifecycle -ceq 'Incomplete' -and $closed[0].ReasonCode -ceq 'SESSION_CHANGED' -and @($closed[0].Observations|Where-Object OutcomeCode -CNE 'REQUESTED').Count -eq 0)
            [void](BoundLibrary 'Open')
            Check ($prefix+'StaleViewerCannotOpenLibrary') ((BoundLibrary 'Count') -ceq '0')
            Check ($prefix+'PriorHistoryAndOtherRootPreserved') ((Preserved $before (BoundPins (Join-Path $Fixture.Root 'Training'))) -and (BoundSame $other (BoundPins $otherTraining)))
            Check ($prefix+'InventoryPreserved') ((Hash $State.InventoryPath) -ceq $bytes)
        }finally{CloseRecordingViewer;[void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RestoreAuthorityFixture');[void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ConfigureAuthorityMutations' @('','',''))}
    }
    ResolveLocal
}
