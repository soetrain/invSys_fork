# Mutate only the disposable saved result; invoke the actual library Refresh for
# every read. Recompute integrity so schema/journal validation is independently tested.
function Test-EvaluationIntegrity($File) {
    $cases=@('UnknownField','StringVersion','WrongWarehouse','JournalHash','ReversedMatches','ForgedOrdinal','TerminalIdentity','DuplicateReason','Oversize','IncompleteTerminalRefsOmitted')
    foreach($case in $cases){
        $rejected=$false;$restored=$false
        if($null -ne $File){
            $original=[IO.File]::ReadAllBytes($File.FullName)
            $count=@(EvaluationFiles).Count
            try {
                $model=[Text.Encoding]::UTF8.GetString($original)|ConvertFrom-Json
                $model.PSObject.Properties.Remove('ContentSha256')
                switch($case){
                    'UnknownField' {$model|Add-Member NoteProperty Unexpected $true}
                    'StringVersion' {$model.Version='1'}
                    'WrongWarehouse' {$model.WarehouseId='different-fixture-warehouse'}
                    'JournalHash' {$model.JournalSha256=('0'*64)}
                    'ReversedMatches' {$model.Matches=@($model.Matches[3],$model.Matches[2],$model.Matches[1],$model.Matches[0])}
                    'ForgedOrdinal' {$model.Matches[0].Ordinal=2}
                    'TerminalIdentity' {$model.TerminalSources[0].EventId='unobserved-fixture-event'}
                    'DuplicateReason' {$model.ReasonCodes=@($model.ReasonCodes[0],$model.ReasonCodes[0])}
                    'IncompleteTerminalRefsOmitted' {$model.ResultState='Incomplete';$model.ReasonCodes=@('SOURCE_UNAVAILABLE');$model.TerminalSources=@()}
                }
                $body=$model|ConvertTo-Json -Depth 24 -Compress
                if($case -eq 'Oversize'){$body=$body.Substring(0,$body.Length-1)+(' '*1048576)+'}'}
                $changed=$body.Substring(0,$body.Length-1)+',"ContentSha256":"'+(EvaluationSha $body)+'"}'
                [IO.File]::WriteAllText($File.FullName,$changed,[Text.UTF8Encoding]::new($false))
                [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
                $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
                $text=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
                $rejected=$status.StartsWith('Incomplete evidence',[StringComparison]::Ordinal) -and $text -ceq '' -and @(EvaluationFiles).Count -eq $count
            } finally {
                [IO.File]::WriteAllBytes($File.FullName,$original)
                [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
            }
            $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
            $restored=$status.StartsWith('Conclusion observed',[StringComparison]::Ordinal) -and @(EvaluationFiles).Count -eq $count -and
                [Convert]::ToBase64String([IO.File]::ReadAllBytes($File.FullName)) -ceq [Convert]::ToBase64String($original)
        }
        Check ('EvaluationIntegrity.'+$case+'.RejectedByActualRead') $rejected
        Check ('EvaluationIntegrity.'+$case+'.OriginalRestoredWithoutAppend') $restored
    }
}
