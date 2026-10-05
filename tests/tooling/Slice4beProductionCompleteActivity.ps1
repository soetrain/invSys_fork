# Observe actual Complete Run clicks and verify their independent owner effects.
# No activity records or business results are manufactured by this fixture.
function Test-ProductionCompleteActivity($Fixture,$Other,$Book,$Decoy,[string]$Canary) {
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    $activities=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    function Pair([string[]]$Before,[string]$Outcome,[string]$Label,[string[]]$EventIds=@(),[string]$Actor='config-producer') {
        $raw=@(Files|Where-Object {$_ -cnotin $Before}|ForEach-Object {[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object {$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
        $last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $incomplete=$Outcome -ceq 'REQUESTED'
        $paired=$records.Count -eq $(if($incomplete){1}else{2}) -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$safe=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false;$sources=$false
        foreach($record in $records){
            $context=$context -and $record.ControlId -ceq 'PRODUCTION_RUN_COMPLETE' -and $record.OwnerId -ceq 'PRODUCTION_RUN_COMPLETION' -and $record.UserId -ceq $Actor -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.StationId -ceq 'S1' -and $record.CatalogVersion -eq 28
        }
        foreach($value in $raw){
            foreach($private in @($Canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','DEMO-RAW-BLACK-TEA')){
                $encoded=ConvertTo-Json -InputObject $private -Compress
                if($value.Contains($private) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$safe=$false}
            }
            $match=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
        }
        if($paired){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and ($incomplete -or $first[0].RecordId -cne $last[0].RecordId) -and $activities.Add([string]$first[0].ActivityId)
            $severity=switch($Outcome){'DENIED'{'Blocked'} 'FAILED'{'Error'} 'PENDING'{'Warning'} 'REJECTED'{'Warning'} default{'Info'}}
            $effect=if($Outcome -cin @('REJECTED','DENIED')){'Unchanged'}else{'Unknown'}
            $facts=$first[0].Severity -ceq 'Info' -and $first[0].DataEffect -ceq 'Unknown' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and $last[0].EventCode -ceq ('PRODUCTION_RUN_COMPLETE_'+$Outcome)
            $sources=@($first[0].SourceEventRefs).Count -eq 0 -and @($last[0].SourceEventRefs).Count -eq $EventIds.Count
            $actual=@($last[0].SourceEventRefs|ForEach-Object EventId)
            $sources=$sources -and ($actual -join '|') -ceq ($EventIds -join '|')
            foreach($reference in $last[0].SourceEventRefs){$sources=$sources -and $reference.WarehouseId -ceq $Fixture.Warehouse -and $reference.SourceKind -ceq 'Inventory' -and $reference.SubmissionState -ceq 'Submitted'}
            $terminal=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress))) -and [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress))) -eq ($Outcome -ceq 'CONFIRMED')
        }
        Check ($Label+'.AttemptAndOutcome') $paired
        Check ($Label+'.ExactContext') $context
        Check ($Label+'.RedactedPayload') $safe
        Check ($Label+'.Integrity') $integrity
        Check ($Label+'.DistinctLinkedAttempt') $linked
        Check ($Label+'.OwnerOutcomeFacts') $facts
        Check ($Label+'.ExactOwnerSourceReferences') $sources
        Check ($Label+'.ExactCommandTerminal') $terminal
    }
    Test-ProductionCompleteCatalog
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Tracking policy prerequisite unavailable; not product RED.'}
    SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
    foreach($guard in @('Loading','Busy','Normal','NoProcess')){
        [void](Probe 'CompleteSubmissionReset')
        if(-not [bool](Probe 'CompleteBaselinePrepare' @($guard -cne 'NoProcess'))){throw 'Actual Check In prerequisite unavailable; not product RED.'}
        [void](Probe 'RunLocalShowAndCapture' @($Book.Name,'CHECK_IN'));$Decoy.Activate()
        $before=Files;$state=[string](Probe 'RunLocalOwnerState');$label='CompleteActivity.Reusable.'+$guard
        # Capture the real writer-return IDs without interrupting execution.
        [void](Probe 'CompleteSubmissionArm' @('','','ObserveOnly'))
        try {
            $entryGuard=if($guard -cin @('Loading','Busy')){$guard}else{''}
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CompleteEntryAct' @($entryGuard)))
            Check ($label+'.GuardsRestored') ([bool](Probe 'CompleteEntryFact' @('GuardsRestored')))
            if($guard -cin @('Loading','Busy')){
                Check ($label+'.NoObservation') ((@(Files) -join '|') -ceq ($before -join '|'))
                Check ($label+'.OwnerPreserved') ([string](Probe 'RunLocalOwnerState') -ceq $state)
                Check ($label+'.ExactInputsPreserved') ([bool](Owner 'CompleteBaselineBalancesForTest' @($false)))
            }elseif($guard -ceq 'Normal'){
                $eventIds=@([string](Probe 'CompleteSubmissionEvent'),[string](Probe 'CompleteSubmissionOutputEvent'))
                if($eventIds.Count -ne 2 -or $eventIds[0] -ceq '' -or $eventIds[1] -ceq '' -or $eventIds[0] -ceq $eventIds[1]){throw 'Actual submission IDs unavailable; not product RED.'}
                Check ($label+'.CompletedByOwner') ([bool](Probe 'CompleteBaselineCompleted'))
                Check ($label+'.ExactInputsConsumedOnce') ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)))
                Check ($label+'.ExactFreshOutput') ([bool](Owner 'CompleteBaselineOutputForTest'))
                Pair $before 'CONFIRMED' $label $eventIds
                CaptureOwnedFormByCaptionEvidence 'Production' 'complete-activity-reusable-confirmed.png'
                [void](Probe 'CompleteSubmissionReset')
                $before=Files;$label='CompleteActivity.Reusable.RefusalAfterSuccess'
                Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CompleteBaselineAct'))
                Check ($label+'.InputsNotConsumedAgain') ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)))
                Check ($label+'.OutputNotDuplicated') ([bool](Owner 'CompleteBaselineOutputForTest'))
                Pair $before 'REJECTED' $label
            }else{
                Check ($label+'.OwnerNotEntered') ([string](Probe 'CompleteBaselineOwners') -ceq '0|0')
                Check ($label+'.ExactInputsPreserved') ([bool](Owner 'CompleteBaselineBalancesForTest' @($false)))
                Pair $before 'REJECTED' $label
            }
            Check ($label+'.CustomValueAndFormulaPreserved') ($Book.Worksheets.Item(1).Cells.Item(2,1).Value2 -ceq $Canary -and $Book.Worksheets.Item(1).Cells.Item(2,2).Formula -ceq '=1+2')
            Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
        } finally {[void](Probe 'CompleteSubmissionReset')}
    }
    Test-ProductionCompleteWorksheet $Fixture $Book $Decoy $Canary
    SelectTarget $Fixture 'config-reader';[void](Probe 'RunLocalReopen' @($Book.Name))
    if(-not [bool](Probe 'CheckBaselineReusableStage' @('Selected'))){throw 'Denied actor staging unavailable; not product RED.'}
    $before=Files;$state=[string](Probe 'RunLocalOwnerState');$entries=[string](Probe 'CompleteBaselineOwners')
    Check 'CompleteActivity.Permission.ActualHandlerReturned' ([bool](Probe 'CompleteEntryAct' @('')))
    Check 'CompleteActivity.Permission.OwnerNotEntered' ([string](Probe 'CompleteBaselineOwners') -ceq $entries)
    Check 'CompleteActivity.Permission.OwnerPreserved' ([string](Probe 'RunLocalOwnerState') -ceq $state)
    Check 'CompleteActivity.Permission.ExistingRefusalWording' ([bool](Probe 'CheckBaselinePermissionRefused'))
    Pair $before 'DENIED' 'CompleteActivity.Permission' -Actor 'config-reader'
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($false))){throw 'Disabled recording policy unavailable; not product RED.'}
    SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
    if(-not [bool](Probe 'CompleteBaselinePrepare' @($true))){throw 'Disabled-policy completion staging unavailable; not product RED.'}
    $before=Files
    Check 'CompleteActivity.Disabled.ActualHandlerReturned' ([bool](Probe 'CompleteEntryAct' @('')))
    Check 'CompleteActivity.Disabled.OwnerCompleted' ([bool](Probe 'CompleteBaselineCompleted'))
    Check 'CompleteActivity.Disabled.ExactInputsConsumedOnce' ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)))
    Check 'CompleteActivity.Disabled.ExactFreshOutput' ([bool](Owner 'CompleteBaselineOutputForTest'))
    Check 'CompleteActivity.Disabled.NoRecords' (@(Files|Where-Object {$_ -cnotin $before}).Count -eq 0)
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Partial-submission recording policy unavailable; not product RED.'}
    Test-ProductionCompleteSubmission $Fixture $Book $Decoy $Canary -ObserveActivity -Other $Other
    Test-ProductionCompleteSubmission $Fixture $Book $Decoy $Canary -AfterOutput -ObserveActivity -Other $Other
    Test-ProductionCompletePolicy $Fixture $Other $Book $Decoy $Canary
}

function Test-ProductionCompleteCatalog {
    $id='PRODUCTION_RUN_COMPLETE'
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(27))).Split("`n")|Where-Object {$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(28))).Split("`n")|Where-Object {$_})
    Check 'CompleteCatalog.Extends27Exactly' ($old.Count -eq 131 -and $new.Count -eq 132 -and @($new|Select-Object -Unique).Count -eq 132 -and ($new[0..130] -join '|') -ceq ($old -join '|') -and $new[-1] -ceq $id)
    $same=$true
    foreach($prior in $old){$before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($prior,27));$after=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($prior,28));$same=$same -and $before -cne '' -and $after -ceq $before}
    Check 'CompleteCatalog.All131PriorDefinitionsPreserved' $same
    $excluded=$true
    foreach($version in 1..27){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
    Check 'CompleteCatalog.ExcludedFromHistoricalVersions' $excluded
    $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,28));$record=if($wire){$wire|ConvertFrom-Json}else{$null}
    Check 'CompleteCatalog.ExactControlContract' ($null -ne $record -and $record.ControlId -ceq $id -and $record.OwnerId -ceq 'PRODUCTION_RUN_COMPLETION' -and $record.Class -ceq 'Command' -and $record.Role -ceq 'Production' -and $record.Caption -ceq 'Complete Run' -and $record.Surface -ceq 'Operations > Production > Production Run - List' -and $record.Capability -ceq 'PROD_POST')
    foreach($code in @('REQUESTED','DENIED','REJECTED','FAILED','PENDING','CONFIRMED','STAGED','COMPLETED','APPLIED')){
        $json=@{ControlId=$id;OwnerId='PRODUCTION_RUN_COMPLETION';CatalogVersion=28;OutcomeCode=$code}|ConvertTo-Json -Compress
        Check ('CompleteCatalog.Terminal.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @($json)) -eq ($code -ceq 'CONFIRMED'))
        $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$code))
        $record=if($wire){$wire|ConvertFrom-Json}else{$null}
        $supported=$code -cin @('REQUESTED','DENIED','REJECTED','FAILED','PENDING','CONFIRMED')
        $valid=$null -eq $record
        if($supported){
            $severity=switch($code){'DENIED'{'Blocked'} 'FAILED'{'Error'} 'PENDING'{'Warning'} 'REJECTED'{'Warning'} default{'Info'}}
            $effect=if($code -cin @('DENIED','REJECTED')){'Unchanged'}else{'Unknown'}
            $valid=$null -ne $record -and $record.EventCode -ceq ($id+'_'+$code) -and $record.Severity -ceq $severity -and $record.DataEffect -ceq $effect
        }
        Check ('CompleteCatalog.Outcome.'+$code) $valid
        Check ('CompleteCatalog.EmptyReferences.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,'[]')) -eq ($code -cin @('REQUESTED','DENIED','REJECTED','FAILED')))
    }
    foreach($kind in @('Inventory','Designs')){foreach($state in @('Submitted','Unknown')){foreach($code in @('REQUESTED','FAILED','PENDING','CONFIRMED')){
        $json=ConvertTo-Json -InputObject @(@{WarehouseId='CATALOG_TEST';SourceKind=$kind;EventId='Source_A';SubmissionState=$state}) -Compress
        $allowed=$kind -ceq 'Inventory' -and ($code -ceq 'FAILED' -or ($state -ceq 'Submitted' -and $code -cin @('PENDING','CONFIRMED')))
        Check ('CompleteCatalog.Source.'+$kind+'.'+$state+'.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,$json)) -eq $allowed)
    }}}
}
