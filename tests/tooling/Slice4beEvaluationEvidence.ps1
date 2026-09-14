# Independent checks of saved product results, reached through the real Evaluate button.
function EvaluationSha([string]$Body) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Body))).Replace('-','').ToLowerInvariant()}
    finally {$sha.Dispose()}
}

function Test-EvaluationEvidence([string]$Stage,$File) {
    $prefix='EvaluationEvidence.'+$Stage+'.'
    $integrity=$false;$provenance=$false;$intent=$false;$publication=$false;$matches=$false;$sources=$false;$schema=$false
    if($null -ne $File){
        try {
            $text=[IO.File]::ReadAllText($File.FullName)
            $record=$text|ConvertFrom-Json
            $tail=[regex]::Match($text,',"ContentSha256":"([0-9a-f]{64})"}$')
            $integrity=$tail.Success -and (EvaluationSha ($text.Substring(0,$tail.Index)+'}')) -ceq $tail.Groups[1].Value
            $fields='SchemaVersion|RecordKind|EvaluationId|Version|ActionPathId|SequenceId|JournalVersion|JournalRecordId|JournalSha256|RecordedByUserId|EvaluatedByUserId|EvaluatedAtUTC|WarehouseId|OriginWarehouseId|CatalogVersion|PackageSetVersion|BuildIdentity|CapturePolicyVersion|EvaluationPolicyVersion|ExpectationSource|ExpectedConclusion|ExpectationSha256|Guide|PreviousEvaluationId|ResultState|ReasonCodes|Matches|MissingSteps|FailedSteps|UnavailableSteps|ExtraActivityIds|TerminalSources|Publication|ContentSha256' -split '\|'
            $names=@($record.PSObject.Properties.Name)
            $schema=$names.Count -eq $fields.Count -and @($names|Where-Object {$_ -cnotin $fields}).Count -eq 0
            $journal=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')[0]
            $provenance=$record.ActionPathId -ceq $journal.ActionPathId -and $record.SequenceId -ceq $journal.SequenceId -and
                $record.JournalVersion -eq $journal.Version -and $record.JournalRecordId -ceq $journal.RecordId -and
                $record.JournalSha256 -ceq $journal.ContentSha256 -and $record.RecordedByUserId -ceq $journal.CreatedByUserId -and
                $record.EvaluatedByUserId -ceq 'config-admin' -and $record.WarehouseId -ceq $Fixture.Warehouse -and
                $record.OriginWarehouseId -ceq $journal.OriginWarehouseId -and $record.CapturePolicyVersion -eq $journal.PolicyVersion -and
                $record.EvaluationPolicyVersion -ge $journal.PolicyVersion -and $record.CatalogVersion -ge $journal.CatalogVersion -and
                $record.EvaluatedAtUTC -cmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$' -and
                $record.PackageSetVersion.Length -gt 0 -and $record.BuildIdentity.Length -gt 0
            $definition=$record.ExpectedConclusion
            $intent=$record.ExpectationSource -ceq 'This evaluation' -and $definition.TerminalKind -ceq 'SourceEventsApplied' -and
                @($definition.Steps).Count -eq 4 -and $definition.TerminalStepId -ceq $definition.Steps[2].StepId -and
                $record.ExpectationSha256 -ceq (EvaluationSha ($definition|ConvertTo-Json -Depth 12 -Compress)) -and
                @($record.Guide.PSObject.Properties).Count -eq 0
            $publication=$record.Publication.Availability -ceq 'Loaded' -and $record.Publication.PublicationId -ceq $published.PublicationId -and
                $record.Publication.ContentSha256 -ceq $published.ContentSha256 -and $record.Publication.SchemaVersion -eq $published.SchemaVersion -and
                $record.Publication.WarehouseId -ceq $Fixture.Warehouse -and $record.Publication.PackageSetVersion -ceq $published.PackageSetVersion -and
                $record.Publication.BuildIdentity -ceq $published.BuildIdentity -and $record.Publication.PublishedAtUTC -ceq $published.PublishedAtUTC -and
                $record.Publication.LoadedAtUTC -cmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$' -and
                ($record.Publication.Coverage|ConvertTo-Json -Depth 12 -Compress) -ceq ($published.Coverage|ConvertTo-Json -Depth 12 -Compress)
            $matches=@($record.Matches).Count -eq 4 -and @($record.MissingSteps).Count -eq 0 -and @($record.FailedSteps).Count -eq 0 -and
                @($record.UnavailableSteps).Count -eq 0 -and @($record.ExtraActivityIds).Count -eq 5
            $ordinal=0;$seen=@{}
            for($i=0;$i -lt @($record.Matches).Count;$i++){
                $match=$record.Matches[$i];$step=$definition.Steps[$i]
                $original=@($journal.Observations|Where-Object {$_.ActivityId -ceq $match.ActivityId -and $_.OutcomeCode -ceq $step.RequiredOutcome})
                $matches=$matches -and $match.StepId -ceq $step.StepId -and $match.ControlId -ceq $step.ControlId -and
                    $match.OutcomeCode -ceq $step.RequiredOutcome -and $match.Ordinal -gt $ordinal -and $original.Count -eq 1 -and
                    $original[0].Ordinal -eq $match.Ordinal -and -not $seen.ContainsKey($match.ActivityId)
                $ordinal=$match.Ordinal;$seen[$match.ActivityId]=$true
            }
            $state=if($Stage -ceq 'Applied'){'Concluded'}else{'Awaiting'}
            $sources=$record.ResultState -ceq $state -and @($record.TerminalSources).Count -eq 4
            foreach($id in $submissionIds){
                $source=@($record.TerminalSources|Where-Object EventId -CEQ $id)
                $group=@($published.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $id})
                $sources=$sources -and $source.Count -eq 1 -and $source[0].WarehouseId -ceq $Fixture.Warehouse -and
                    $source[0].SourceKind -ceq 'Inventory' -and $source[0].SubmissionState -ceq 'Submitted'
                if($group.Count -eq 1){
                    $keys=@($group[0].Lines|ForEach-Object System_Key)
                    $lineBody=[ordered]@{Lines=@($group[0].Lines)}|ConvertTo-Json -Depth 16 -Compress
                    $sources=$sources -and $source[0].OwnerStatus -ceq 'Applied' -and $source[0].LineCount -eq $keys.Count -and
                        ($source[0].SystemKeys -join '|') -ceq ($keys -join '|') -and $source[0].LinesSha256 -ceq (EvaluationSha $lineBody)
                }else{
                    $sources=$sources -and $source[0].OwnerStatus -ceq 'Awaiting' -and $source[0].LineCount -eq 0 -and
                        $source[0].LinesSha256 -ceq '' -and @($source[0].SystemKeys).Count -eq 0
                }
            }
        } catch {
            # Malformed/missing product fields are behavioral failure. Never print
            # the record, actors, source values, paths or exception payload.
        }
    }
    Check ($prefix+'ExactSchemaFields') $schema
    Check ($prefix+'FinalContentHash') $integrity
    Check ($prefix+'ExactJournalActorsAndPolicies') $provenance
    Check ($prefix+'ExplicitExpectationHashAndProvenance') $intent
    Check ($prefix+'LoadedPublicationIdentityHashAndCoverage') $publication
    Check ($prefix+'OrderedDistinctMatchesAndExtras') $matches
    Check ($prefix+'EveryTerminalSourceRetainsExactAppliedEvidence') $sources
    $display=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
    $evaluationTime=$false;$loadedTime=$false
    if($provenance){$evaluationTime=$display.Contains('Evaluated at: '+$record.EvaluatedAtUTC.Substring(0,19).Replace('T',' ')+' UTC')}
    if($publication){$loadedTime=$display.Contains('Publication loaded at: '+$record.Publication.LoadedAtUTC.Substring(0,19).Replace('T',' ')+' UTC')}
    Check ($prefix+'EvaluationTimeUsesVerifiedUtcDisplay') $evaluationTime
    Check ($prefix+'LoadedTimeUsesVerifiedUtcDisplay') $loadedTime
    $rejected=$false;$restored=$false
    if($null -ne $File -and $integrity){
        $originalBytes=[IO.File]::ReadAllBytes($File.FullName)
        $savedCount=@(EvaluationFiles).Count
        try {
            [IO.File]::WriteAllText($File.FullName,$text.Replace(',"ContentSha256":"',',"ContentSha256":"0'),[Text.UTF8Encoding]::new($false))
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
            $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
            $display=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
            $rejected=$status.StartsWith('Incomplete evidence',[StringComparison]::Ordinal) -and $display -ceq '' -and
                @(EvaluationFiles).Count -eq $savedCount
        } finally {
            [IO.File]::WriteAllBytes($File.FullName,$originalBytes)
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
        }
        $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
        $expected=if($Stage -ceq 'Applied'){'Conclusion observed'}else{'Awaiting published result'}
        $restored=$status.StartsWith($expected,[StringComparison]::Ordinal) -and @(EvaluationFiles).Count -eq $savedCount -and
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($File.FullName)) -ceq [Convert]::ToBase64String($originalBytes)
    }
    Check ($prefix+'CorruptSavedResultClearsConclusionWithoutWriting') $rejected
    Check ($prefix+'RestoredSavedResultReadsWithoutAppending') $restored
    if($Stage -ceq 'Applied'){
        . (Join-Path $PSScriptRoot 'Slice4beEvaluationIntegrity.ps1')
        Test-EvaluationIntegrity $File
    }
}
