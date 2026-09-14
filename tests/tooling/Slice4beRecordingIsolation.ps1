# D18: replay only a real owning handler's recorded result. No synthetic
# command completion or business mutation substitutes for operator actions.
function Test-Slice4beRecordingIsolation($Fixture) {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActivity').CodeModule
    $source=$core.Lines(1,$core.CountOfLines)
    $start=$source.IndexOf('Public Function FinishAction(',[StringComparison]::OrdinalIgnoreCase)
    $finish=$source.IndexOf('End Function',$start,[StringComparison]::OrdinalIgnoreCase)
    if($start -lt 0 -or $finish -lt $start){throw 'FinishAction exception fixture boundary is unavailable.'}
    $body=$source.Substring($start,$finish-$start)
    $markers=[regex]::Matches($body,'(?im)^[ \t]*Set[ \t]+target[ \t]*=[ \t]*action\("Target"\)[ \t]*\r?$')
    if($markers.Count -ne 1){throw 'Exception fixture requires one resolved owning action.'}
    $marker=$markers[0]
    $replacement=$marker.Value.TrimEnd("`r")+"`r`n"+'    If activityId = RecordingExceptionForTest() Then Err.Raise 5'
    $changed=$body.Substring(0,$marker.Index)+$replacement+$body.Substring($marker.Index+$marker.Length)
    $core.DeleteLines(1,$core.CountOfLines)
    $core.AddFromString($source.Substring(0,$start)+$changed+$source.Substring($finish))
    $core.AddFromString(@'
Public Function RecordingReplayForTest(ByVal activityId As String, ByVal outcome As String, Optional ByVal raiseResultError As Boolean = False) As String
    Dim notice As String, accepted As Boolean
    If raiseResultError Then RecordingExceptionForTest activityId
    accepted = FinishAction(activityId, outcome, notice)
    RecordingExceptionForTest "", True
    RecordingReplayForTest = CStr(accepted) & "|" & notice
End Function
Private Function RecordingExceptionForTest(Optional ByVal activityId As String = "", Optional ByVal clear As Boolean = False) As String
    Static selected As String
    If clear Then selected = ""
    If activityId <> "" Then selected = activityId
    RecordingExceptionForTest = selected
End Function
'@)
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    SetRecordingPolicy $true
    OpenRecordingViewer
    [void](RecordingControl 'Start Recording' 'Click')
    $old=SaveRecordedSetting '621'
    [void](RecordingControl 'Stop Recording' 'Click')
    if(-not (HasSequence $old 1) -or -not (JournalChain $old.Attempt.SequenceId 4)) {
        throw 'Isolation fixture requires an actual saved owning action and stopped journal.'
    }
    Check 'RecordingIsolation.RealOldOutcomePrepared' (JournalFact $old.Attempt.SequenceId 'Close' 1 'Stopped')
    $settingValue=621
    foreach($suffix in @('SamePolicy','StorageConflict','ResultException','CurrentStorageConflict','CurrentResultException','ChangedPolicy')) {
        $settingValue++
        $changedPolicy=$suffix -ceq 'ChangedPolicy'
        $ownsCurrent=$suffix -like 'Current*'
        CloseRecordingViewer
        if($changedPolicy){SetRecordingPolicy $false; SetRecordingPolicy $true}
        OpenRecordingViewer
        [void](RecordingControl 'Start Recording' 'Click')
        $new=SaveRecordedSetting ([string]$settingValue)
        if(-not (HasSequence $new 1) -or -not (JournalChain $new.Attempt.SequenceId 3) -or
            $new.Attempt.SequenceId -ceq $old.Attempt.SequenceId) {
            throw 'Isolation fixture requires a distinct active sequence with a real owning action.'
        }
        Check ('RecordingIsolation.'+$suffix+'.DistinctCurrentRun') (
            ($changedPolicy -and $new.Attempt.PolicyVersion -ne $old.Attempt.PolicyVersion) -or
            (-not $changedPolicy -and $new.Attempt.PolicyVersion -eq $old.Attempt.PolicyVersion))
        $activityBefore=ActivityPins
        $journalBefore=@{}
        foreach($file in Get-ChildItem -LiteralPath $journalRoot -File){$journalBefore[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
        $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $replayed=if($ownsCurrent){$new}else{$old}
        $outcomePath=Join-Path $activityRoot ($replayed.Outcome.RecordId+'.json')
        $originalBytes=[IO.File]::ReadAllBytes($outcomePath)
        try {
            # Only this generated fixture's existing outcome file is obstructed;
            # restore its exact bytes before preservation checks and later reads.
            if($suffix -like '*StorageConflict'){[IO.File]::WriteAllText($outcomePath,'Conflicting isolated activity record.')}
            $reply=[string](Run 'invSys.Core.xlam' 'modActivity.RecordingReplayForTest' @($replayed.Attempt.ActivityId,$replayed.Outcome.OutcomeCode,($suffix -like '*ResultException')))
        } finally {
            if($suffix -like '*StorageConflict'){[IO.File]::WriteAllBytes($outcomePath,$originalBytes)}
        }
        $expectedReply=switch -Wildcard ($suffix){
            'SamePolicy'{'^True\|'}
            'ChangedPolicy'{'^False\|.*policy changed'}
            '*StorageConflict'{'^False\|.*could not be saved'}
            '*ResultException'{'^False\|.*result could not be recorded'}
        }
        $name=if($ownsCurrent){'OwningResultDisposition'}else{'OldResultDisposition'}
        Check ('RecordingIsolation.'+$suffix+'.'+$name) (
            $reply -match $expectedReply)
        $expectedJournalCount=$journalBefore.Count+[int]$ownsCurrent
        $unchanged=$expectedJournalCount -eq @(Get-ChildItem -LiteralPath $journalRoot -File).Count
        foreach($file in $journalBefore.Keys){if((Get-FileHash -LiteralPath $file).Hash -cne $journalBefore[$file]){$unchanged=$false}}
        $name=if($ownsCurrent){'FailurePreservesOriginalEntries'}else{'ReplayPreservesBothJournals'}
        Check ('RecordingIsolation.'+$suffix+'.'+$name) $unchanged
        Check ('RecordingIsolation.'+$suffix+'.ReplayPreservesActivityAndConfig') (
            (PinsRetained $activityBefore) -and $activityBefore.Count -eq (ActivityPins).Count -and
            $configBefore -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
        if($ownsCurrent){
            Check ('RecordingIsolation.'+$suffix+'.CurrentRunExplainsIncomplete') (
                (RecordingControl 'Stop Recording') -ceq 'True|False' -and
                (RecordingStatus) -match '(?i)incomplete evidence.*tracking unavailable')
            Check ('RecordingIsolation.'+$suffix+'.FailureClosesOwningSequence') (
                (JournalFact $new.Attempt.SequenceId 'Close' 1 'Incomplete') -and
                (JournalChain $new.Attempt.SequenceId 4))
        } else {
            Check ('RecordingIsolation.'+$suffix+'.CurrentRunRemainsActive') (
                (RecordingControl 'Stop Recording') -ceq 'True|True' -and
                (RecordingStatus) -match '(?i)recording.*1\s*/\s*256')
            [void](RecordingControl 'Stop Recording' 'Click')
            Check ('RecordingIsolation.'+$suffix+'.CurrentRunStopsNormally') (
                (JournalFact $new.Attempt.SequenceId 'Close' 1 'Stopped') -and
                (JournalChain $new.Attempt.SequenceId 4))
        }
    }
    CloseRecordingViewer
}
