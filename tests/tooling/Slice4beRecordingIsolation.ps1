# D18: replay only a real owning handler's recorded result. No synthetic
# command completion or business mutation substitutes for operator actions.
function Test-Slice4beRecordingIsolation($Fixture) {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActivity').CodeModule
    $core.AddFromString(@'
Public Function RecordingReplayForTest(ByVal activityId As String, ByVal outcome As String) As String
    Dim notice As String, accepted As Boolean
    accepted = FinishAction(activityId, outcome, notice)
    RecordingReplayForTest = CStr(accepted) & "|" & notice
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
    foreach($changedPolicy in @($false,$true)) {
        CloseRecordingViewer
        if($changedPolicy){SetRecordingPolicy $false; SetRecordingPolicy $true}
        OpenRecordingViewer
        [void](RecordingControl 'Start Recording' 'Click')
        $new=SaveRecordedSetting $(if($changedPolicy){'623'}else{'622'})
        $suffix=if($changedPolicy){'ChangedPolicy'}else{'SamePolicy'}
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
        $reply=[string](Run 'invSys.Core.xlam' 'modActivity.RecordingReplayForTest' @($old.Attempt.ActivityId,$old.Outcome.OutcomeCode))
        Check ('RecordingIsolation.'+$suffix+'.OldResultDisposition') (
            ($changedPolicy -and $reply -match '^False\|.*policy changed') -or
            (-not $changedPolicy -and $reply -match '^True\|'))
        $unchanged=$journalBefore.Count -eq @(Get-ChildItem -LiteralPath $journalRoot -File).Count
        foreach($file in $journalBefore.Keys){if((Get-FileHash -LiteralPath $file).Hash -cne $journalBefore[$file]){$unchanged=$false}}
        Check ('RecordingIsolation.'+$suffix+'.ReplayPreservesBothJournals') $unchanged
        Check ('RecordingIsolation.'+$suffix+'.ReplayPreservesActivityAndConfig') (
            (PinsRetained $activityBefore) -and $activityBefore.Count -eq (ActivityPins).Count -and
            $configBefore -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
        Check ('RecordingIsolation.'+$suffix+'.CurrentRunRemainsActive') (
            (RecordingControl 'Stop Recording') -ceq 'True|True' -and
            (RecordingStatus) -match '(?i)recording.*1\s*/\s*256')
        [void](RecordingControl 'Stop Recording' 'Click')
        Check ('RecordingIsolation.'+$suffix+'.CurrentRunStopsNormally') (
            (JournalFact $new.Attempt.SequenceId 'Close' 1 'Stopped') -and
            (JournalChain $new.Attempt.SequenceId 4))
    }
    CloseRecordingViewer
}
