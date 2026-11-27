# Guide input is transferred through the ordinary UI. Each gate uses a new destination.
function Test-ProductionOperatorStep($Source,$SourceGuide,[string]$Mode){
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beSeedReplayProof.ps1')
    if(-not (Get-Variable guideTransferDialogProbeInstalled -ValueOnly -ErrorAction SilentlyContinue)){throw 'Destination guide transfer file picker fixture is missing.'}
    $runner='frmActionPathRun'
    $temporaryGuide=Join-Path $runRoot 'operator-completion-guide.json'
    $sourceInventory=Join-Path $Source.Root ($Source.Warehouse+'.invSys.Data.Inventory.xlsb')
    $sourceHash=Hash $sourceInventory
    $work=$null
    $case=[pscustomobject]@{RunId=''}
    function ClickRun([string]$name){Delivered (BoundControl $name 'Click' '' $runner)}
    function StartCase {
        $dir=Join-Path $destRoot 'Runs'
        $before=BoundPins $dir
        ClickRun 'btnStartRun'
        $ids=@(Get-ChildItem $dir -Filter '*.json' -File|Where-Object {!$before.ContainsKey($_.FullName)}|ForEach-Object {($_.BaseName -split '\.')[0]}|Sort-Object -Unique)
        if($ids.Count -ne 1 -or $ids[0] -ceq $case.RunId){throw 'Start must create exactly one new execution identity.'}
        $case.RunId=$ids[0]
    }
    function LastRun {
        if(!$case.RunId){throw 'Missing captured run identity.'}
        $files=@(Get-ChildItem -LiteralPath (Join-Path $destRoot 'Runs') -Filter ($case.RunId+'.*.json') -File)
        $latest=@($files|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision -Descending)
        if(!$latest.Count){throw 'Expected a persisted execution run.'}
        $latest[0]
    }
    try{
        TransferSetFile $temporaryGuide
        Delivered (BoundControl 'btnExportGuide' 'Click')
        if(!(Test-Path -LiteralPath $temporaryGuide)){throw 'Source guide export unavailable.'}
        CloseRecordingViewer
        $dest=NewFixture 'operator-destination'
        if($dest.Warehouse -ceq $Source.Warehouse){throw 'Distinct Training destination required.'}
        SelectTarget $dest
        if(-not ([string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($dest.Warehouse,'S1','config-admin'))).StartsWith('OK|')){throw 'Destination seed prerequisite failed.'}
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Destination read policy unavailable.'}
        SetRecordingPolicy $true
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Destination publication unavailable.'}
        OpenRecordingViewer;Delivered (BoundLibrary 'Open');BoundOpen
        Delivered (BoundControl 'btnImportGuide' 'Click')
        $destRoot=Join-Path $dest.Root ('Training/ActionPaths/'+$dest.Warehouse)
        $files=@(Get-ChildItem (Join-Path $destRoot 'Guides') -Filter '*.json' -File)
        if($files.Count -ne 1){throw 'One imported destination guide required.'}
        $guide=ReadGuideExpectationRecord $files[0].FullName
        if($guide.TransferOrigin.ContentSha256 -cne $SourceGuide.ContentSha256){throw 'Incorrect transfer origin.'}
        BoundSelect $guide;Delivered (BoundControl 'btnConfigureExecution' 'Click')
        $available=(BoundControl '' 'Count' '' 'frmActionPathExecution') -ceq '1'
        Pass 'CanConfigureOperatorRequiredStep' $available
        if(!$available){return}
        Delivered (BoundControl 'btnSaveExecutionProfile' 'Click' '' 'frmActionPathExecution')
        $files=@(Get-ChildItem (Join-Path $destRoot 'ExecutionProfiles') -Filter '*.json' -File)
        if($files.Count -ne 1){throw 'Expected one saved execution profile.'}
        $profile=Get-Content $files[0].FullName -Raw|ConvertFrom-Json
        Pass 'ExactOperatorProfile' ($profile.Guide.ContentSha256 -ceq $guide.ContentSha256 -and $profile.Steps.Count -eq 1 -and $profile.Steps[0].StepId -ceq $guide.Steps[0].StepId -and $profile.Steps[0].AdapterId -ceq $control -and $profile.Steps[0].Inputs.Count -eq 0)
        CloseRecordingViewer;SelectTarget $dest 'config-producer'
        $work=$excel.Workbooks.Add()
        [void](Probe 'OpenDesigner' @($work.Name))
        $contextName='BOUND'+[guid]::NewGuid().ToString('N')
        if(-not [bool](Probe 'ReadPrepare' @($contextName))){throw 'Released destination definitions unavailable.'}
        [void](Probe 'RunLocalRememberFixture')
        if($Mode -ceq 'Reusable'){$ready=[bool](Probe 'CompleteBaselinePrepare' @($true))}
        else{
            $ready=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CompleteWorksheetInventoryForTest' @($work.Name))
            if($ready){$ready=[bool](Probe 'CompleteWorksheetPrepare' @($contextName))}
        }
        if(!$ready){throw 'Destination Check In / output preparation unavailable.'}
        [void](Probe 'RunLocalShowAndCapture' @($work.Name,'CHECK_IN'))
        OpenRecordingViewer;Delivered (BoundLibrary 'Open');BoundOpen;BoundSelect $guide
        Delivered (BoundControl 'btnRunHowTo' 'Click')
        StartCase
        $waiting=LastRun
        Pass 'EligibleStepWaitsWithoutOwnerCall' ($waiting.State -ceq 'Running' -and $waiting.Steps.Count -eq 0 -and $waiting.OperatorClaims.Count -eq 1 -and $waiting.OperatorClaims[0].StepId -ceq $profile.Steps[0].StepId -and !$waiting.OperatorClaims[0].Kind -and (Run 'invSys.Operations.xlam' 'TestCompleteReplay.Counts') -ceq '2|2')
        Pass 'NextStepDisabledWhileOwnerPending' ((BoundControl 'btnNextRunStep' 'Click' '' $runner) -ceq 'DISABLED')
        ClickRun 'btnCheckProduction'
        Pass 'EarlierCompletionsCannotAdvance' ((LastRun).ContentSha256 -ceq $waiting.ContentSha256)
        ClickRun 'btnBindProduction'
        $bound=LastRun
        $kind=if($Mode -ceq 'Reusable'){'ReusableProcess'}else{'WorksheetOutput'}
        $confirmed=$bound.OperatorClaims[0].Kind -ceq $kind -and $bound.OperatorClaims[0].BindingId -ceq $waiting.OperatorClaims[0].BindingId -and $bound.OperatorClaims[0].SelectionSha256 -cmatch '^[a-f0-9]{64}$' -and !$bound.OperatorClaims[0].ActivityId
        Pass 'BoundExactDestinationBeforeOwnerClick' $confirmed
        if(!$confirmed){return}
        $inventoryPath=Join-Path $dest.Root ($dest.Warehouse+'.invSys.Data.Inventory.xlsb')
        $initialAuthority=Read-SeedReplayInventory $inventoryPath
        $before=@(Get-Slice4beActivityFiles $dest)
        [void](Probe 'CompleteSubmissionArm' @('','','ObserveOnly'))
        [void](Run 'invSys.Operations.xlam' 'TestCompleteReplay.LoseNextNotification')
        if(-not [bool](Probe 'CompleteEntryAct' @(''))){throw 'Ordinary destination completion unavailable.'}
        $records=@(Get-Slice4beActivityFiles $dest|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
        $terminal=@($records|Where-Object OutcomeCode -CEQ 'CONFIRMED')
        if($terminal.Count -ne 1){throw 'Expected an ordinary CONFIRMED outcome.'}
        $terminal=$terminal[0]
        $claimed=LastRun
        Pass 'OrdinaryOwnerClaimedOnce' ((Run 'invSys.Operations.xlam' 'TestCompleteReplay.Counts') -ceq '3|3' -and $claimed.Steps.Count -eq 0 -and $claimed.OperatorClaims[0].ActivityId -ceq $terminal.ActivityId -and $claimed.OperatorClaims[0].OwnerValidated -and $terminal.SequenceId -ceq $waiting.Recording.SequenceId)
        if($Mode -ceq 'Reusable'){
            $allocation=Owner 'CompleteSubmissionInputForTest';$output=Owner 'CompleteSubmissionOutputSelectionForTest'
            $keys=@($allocation[0],$output[0]);$delta=@(-[double]$allocation[1],[double]$output[1])
        }else{$keys=Probe 'CompletePathWorksheetKeys';$delta=@(-10.0,1.0);Pass 'OrdinaryWorksheetLog' ([bool](Probe 'CompleteWorksheetFact' @('Log',$contextName)))}
        Test-CompletePathAuthority $dest @($terminal.SourceEventRefs.EventId) $keys $delta 'OperatorDestination'|Out-Host
        $afterAuthority=Read-SeedReplayInventory $inventoryPath
        $priorRows=@($initialAuthority.tblInventoryEntities|Where-Object System_Key -CNE $keys[0]|Sort-Object System_Key)
        $otherRows=@($afterAuthority.tblInventoryEntities|Where-Object {$_.System_Key -cnotin $keys}|Sort-Object System_Key)
        $inputBefore=@($initialAuthority.tblInventoryEntities|Where-Object System_Key -CEQ $keys[0])
        $inputAfter=@($afterAuthority.tblInventoryEntities|Where-Object System_Key -CEQ $keys[0])
        $outputAfter=@($afterAuthority.tblInventoryEntities|Where-Object System_Key -CEQ $keys[1])
        Pass 'SavedEffectsAndAllUnrelatedInventory' ($terminal.SourceEventRefs.Count -eq 2 -and $keys[0] -cne $keys[1] -and $inputBefore.Count -eq 1 -and $inputAfter.Count -eq 1 -and $outputAfter.Count -eq 1 -and $keys[1] -cnotin @($initialAuthority.tblInventoryEntities.System_Key) -and ($inputAfter[0].QtyOnHand-$inputBefore[0].QtyOnHand) -eq $delta[0] -and $outputAfter[0].QtyOnHand -eq $delta[1] -and (ConvertTo-Json -Compress -InputObject $priorRows) -ceq (ConvertTo-Json -Compress -InputObject $otherRows))
        $afterEffects=Hash $inventoryPath
        ClickRun 'btnCheckProduction'
        Pass 'OlderPublicationKeepsStepPending' ((LastRun).ContentSha256 -ceq $claimed.ContentSha256)
        Delivered (BoundControl 'btnRefresh' 'Click' '' 'frmInventoryViewer')
        ClickRun 'btnCheckProduction'
        $done=LastRun
        $credited=$done.State -ceq 'Completed' -and $done.Steps.Count -eq 1 -and $done.Steps[0].State -ceq 'OperatorCompleted' -and $done.Steps[0].ActivityId -ceq $terminal.ActivityId -and ($done.Steps[0].SourceEventRefs.EventId -join '|') -ceq ($terminal.SourceEventRefs.EventId -join '|')
        Pass 'OnlyFreshAppliedClaimAdvances' $credited
        if(!$credited){return}
        $repeat=BoundControl 'btnCheckProduction' 'Click' '' $runner
        Pass 'DuplicateEvidenceCheckCannotConsumeAgain' ($repeat -ceq 'DISABLED' -and (LastRun).ContentSha256 -ceq $done.ContentSha256 -and (Hash $inventoryPath) -ceq $afterEffects -and (Run 'invSys.Operations.xlam' 'TestCompleteReplay.Counts') -ceq '3|3')
        [ordered]@{RunId=$done.RunId;SequenceId=$done.Recording.SequenceId;StepId=$done.Steps[0].StepId;BindingId=$done.OperatorClaims[0].BindingId;ActivityId=$terminal.ActivityId;EventIds=@($terminal.SourceEventRefs.EventId)}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'operator-completion-identities.json')
        foreach($i in 1..2){
            $evaluations=Join-Path $destRoot 'Evaluations'
            $pins=BoundPins $evaluations
            ClickRun 'btnVerifyRun'
            $result=BoundControl 'txtRunVerification' 'Text' '' $runner
            $added=@(Get-ChildItem $evaluations -File -Filter '*.json'|Where-Object {!$pins.ContainsKey($_.FullName)})
            if($added.Count -ne 1){throw 'Verify must create one saved diagnostic.'}
            $saved=Get-Content $added[0].FullName -Raw|ConvertFrom-Json
            Pass ("VerifyFreshCompletion"+$i) ($result.StartsWith('Conclusion observed') -and $saved.ResultState -ceq 'Concluded' -and $saved.Matches.Count -eq 1 -and $saved.Matches[0].ActivityId -ceq $terminal.ActivityId -and (LastRun).ContentSha256 -ceq $done.ContentSha256 -and (Hash $inventoryPath) -ceq $afterEffects -and (Run 'invSys.Operations.xlam' 'TestCompleteReplay.Counts') -ceq '3|3')
        }
        . (Join-Path $PSScriptRoot 'Slice4beProductionOperatorControls.ps1')
        Test-OperatorBindingControls
    }finally{
        [void](Probe 'CompleteSubmissionReset')
        CloseRecordingViewer;[void](Probe 'CloseDesigner')
        if($null -ne $work){$work.Close($false)}
        TransferSetFile '';SelectTarget $Source
        Pass 'SourceInventoryRemainsIndependent' ((Hash $sourceInventory) -ceq $sourceHash)
    }
}
