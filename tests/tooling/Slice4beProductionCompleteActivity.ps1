# Observe actual Complete Run clicks and verify their independent owner effects.
# No activity records or business results are manufactured by this fixture.
function Test-ProductionCompleteActivity($Fixture,$Other,$Book,$Decoy,[string]$Canary) {
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    $activities=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    function Pair([string[]]$Before,[string]$Outcome,[string]$Label,[string[]]$EventIds=@()) {
        $raw=@(Files|Where-Object {$_ -cnotin $Before}|ForEach-Object {[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object {$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
        $last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$safe=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false;$sources=$false
        foreach($record in $records){
            $context=$context -and $record.ControlId -ceq 'PRODUCTION_RUN_COMPLETE' -and $record.OwnerId -ceq 'PRODUCTION_RUN_COMPLETION' -and $record.UserId -ceq 'config-producer' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.StationId -ceq 'S1' -and $record.CatalogVersion -eq 28
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
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId -and $activities.Add([string]$first[0].ActivityId)
            $severity=if($Outcome -ceq 'REJECTED'){'Warning'}else{'Info'}
            $effect=if($Outcome -ceq 'REJECTED'){'Unchanged'}else{'Unknown'}
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
}
