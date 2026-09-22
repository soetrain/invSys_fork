# D18: synthetic format variants of an actual run, preserving owning observations.
# All mutations stay inside the generated fixture and exact bytes are restored.
function Test-ExpectationCompatibility {
    $entries=@(RecordingJournal $sequence|Sort-Object Version)
    if($entries.Count -ne 20){throw 'Actual Operations compatibility fixture is incomplete.'}
    # Inspect the original owner-written journal before any compatibility variant.
    # Reader support for schema 2 alone does not satisfy the approved writer.
    Check 'ExpectationCapture.NewJournalUsesSchema2' (@($entries|Where-Object SchemaVersion -NE 2).Count -eq 0)
    $nonClosingNone=$true
    foreach($entry in @($entries|Where-Object RecordType -CNE 'Close')){
        if($null -eq $entry.PSObject.Properties['ExpectedConclusion']){$nonClosingNone=$false;continue}
        $definition=$entry.ExpectedConclusion
        if($definition.SchemaVersion -ne 1 -or $definition.TerminalKind -cne 'None' -or
            $definition.TerminalStepId -cne '' -or @($definition.Steps).Count -ne 0){$nonClosingNone=$false}
    }
    Check 'ExpectationCapture.NonClosingEntriesHaveNoExpectation' $nonClosingNone
    $defaultClose=$entries[-1]
    $defaultNone=$false
    if($null -ne $defaultClose.PSObject.Properties['ExpectedConclusion']){
        $definition=$defaultClose.ExpectedConclusion
        $defaultNone=$defaultClose.RecordType -ceq 'Close' -and $definition.SchemaVersion -eq 1 -and
            $definition.TerminalKind -ceq 'None' -and $definition.TerminalStepId -ceq '' -and @($definition.Steps).Count -eq 0
    }
    Check 'ExpectationCapture.DefaultCloseHasNoExpectation' $defaultNone
    $bytes=@{}
    foreach($entry in $entries){$file=Join-Path $journalRoot ($pathId+'.'+$entry.Version+'.json');$bytes[$file]=[IO.File]::ReadAllBytes($file)}
    $identity=[string]$entries[1].Observations[0].ActivityId
    $utf8=[Text.UTF8Encoding]::new($false)
    foreach($case in @('Schema1','Schema2None','Schema2Expectation','DuplicateStepId','UnsupportedControl','UnsupportedOutcome','RetryWrongType','UnexpectedField','MissingTerminal','TooManySteps','MixedSchema','EarlyExpectation','ExpectationSchema','DuplicateField')){
        try {
            $previous=''
            foreach($entry in $entries){
                $file=Join-Path $journalRoot ($pathId+'.'+$entry.Version+'.json')
                $model=$utf8.GetString($bytes[$file])|ConvertFrom-Json
                $model.PSObject.Properties.Remove('ContentSha256')
                $model.PSObject.Properties.Remove('ExpectedConclusion')
                $model.PreviousSha256=$previous
                $model.SchemaVersion=2
                $expectation=[ordered]@{SchemaVersion=1;Steps=@();TerminalStepId='';TerminalKind='None'}
                if($case -ne 'Schema2None' -and ($model.RecordType -ceq 'Close' -or $case -eq 'EarlyExpectation')){
                    $step=[ordered]@{StepId='expected-step-1';ControlId='ADMIN_SETTINGS_SAVE_VALUE';RequiredOutcome='COMPLETED';RetryAllowed=$true}
                    $expectation.Steps=@($step);$expectation.TerminalStepId=$step.StepId;$expectation.TerminalKind='CommandCompleted'
                    switch($case){
                        'DuplicateStepId' {$expectation.Steps=@($step,$step)}
                        'UnsupportedControl' {$step.ControlId='UNREGISTERED_EXPECTATION_CONTROL'}
                        'UnsupportedOutcome' {$step.RequiredOutcome='UNREGISTERED_EXPECTATION_OUTCOME'}
                        'RetryWrongType' {$step.RetryAllowed='True'}
                        'UnexpectedField' {$expectation.Add('Executable','NOT_EXECUTED')}
                        'MissingTerminal' {$expectation.TerminalStepId='missing-step'}
                        'TooManySteps' {$expectation.Steps=@(foreach($i in 1..257){[ordered]@{StepId=('expected-step-'+$i);ControlId='ADMIN_SETTINGS_SAVE_VALUE';RequiredOutcome='COMPLETED';RetryAllowed=$true}})}
                        'ExpectationSchema' {$expectation.SchemaVersion=2}
                    }
                }
                if($case -eq 'Schema1' -or ($case -eq 'MixedSchema' -and $model.Version -eq 1)){$model.SchemaVersion=1}
                else{$model|Add-Member -MemberType NoteProperty -Name ExpectedConclusion -Value $expectation}
                $body=$model|ConvertTo-Json -Depth 40 -Compress
                if($case -eq 'DuplicateField' -and $model.RecordType -ceq 'Close'){$body=$body.Replace('"TerminalKind":"CommandCompleted"','"TerminalKind":"CommandCompleted","TerminalKind":"CommandCompleted"')}
                $sha=[Security.Cryptography.SHA256]::Create()
                try{$previous=([BitConverter]::ToString($sha.ComputeHash($utf8.GetBytes($body)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
                $content=$body.Substring(0,$body.Length-1)+',"ContentSha256":"'+$previous+'"}'
                [IO.File]::WriteAllText($file,$content,$utf8)
            }
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
            $selected=Select-EvaluationRun $pathId
            $status=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Status',''))
            $evidence=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Evidence',''))
            $expectedValid=$case -in @('Schema1','Schema2None','Schema2Expectation')
            $valid=$selected -ceq 'SELECTED' -and $status.StartsWith('Stopped',[StringComparison]::Ordinal) -and $evidence.Contains($identity)
            $invalid=$selected -ceq 'SELECTED' -and $status.StartsWith('Incomplete evidence',[StringComparison]::Ordinal) -and -not $evidence.Contains($identity)
            Check ('ExpectationCompatibility.'+$case) $(if($expectedValid){$valid}else{$invalid})
        } finally {foreach($file in $bytes.Keys){[IO.File]::WriteAllBytes($file,$bytes[$file])}}
    }
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
    [void](Select-EvaluationRun $pathId)
    $preserved=$true
    foreach($file in $bytes.Keys){if([Convert]::ToBase64String([IO.File]::ReadAllBytes($file)) -cne [Convert]::ToBase64String($bytes[$file])){$preserved=$false}}
    Check 'ExpectationCompatibility.OriginalJournalRestoredExactly' $preserved
}

function Capture-ExpectationEditorEvidence([string]$Name) {
    try {
        Initialize-SettingsCapture
        $editorHandle=[InvSysSettingsCapture]::OwnedVisibleForm('Expected steps and conclusion',[IntPtr]$excel.Hwnd).ToInt64()
        if($editorHandle -eq 0){throw 'Requested form window unavailable.'}
        $editorActivation=New-Object -ComObject WScript.Shell
        try {[void]$editorActivation.AppActivate('Expected steps and conclusion')}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($editorActivation)}
        CaptureFormEvidence 'Expected steps and conclusion' ('expectation-editor-'+$Name.ToLowerInvariant()+'.png') $editorHandle
        Check ('Harness.ExpectationVisibleCapture.'+$Name) $true
    } catch {
        $message=$_.Exception.GetBaseException().Message
        if($message -cnotin @('Requested form is not in the foreground.','Requested form window unavailable.')){throw}
        # A visible-evidence failure is not product RED. Continue the actual
        # editor actions so one unavailable capture does not mask binding results.
        Check ('Harness.ExpectationVisibleCapture.'+$Name) $false
        Write-Output ('Visible editor evidence unavailable: '+$Name+'; '+$message)
    }
}

function Test-ExpectationEditorBinding {
    # Show only this owned disposable Excel instance for operator-view evidence.
    # A modeless form's Visible property alone does not establish a native window.
    $previousVisibility=[bool]$excel.Visible
    $excel.Visible=$true
    if($CheckAdminUomExpectationChoice){Test-AdminUomExpectedCancellation}
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual recorder baseline unavailable for editor binding.'}
    $action=SaveRecordedSetting '681'
    $activeSequence=[string]$action.Attempt.SequenceId
    $opened=ExpectationControl 'btnRecordingExpectation' 'Click' '' 'frmInventoryViewer'
    $caption=ExpectationControl '' 'FormCaption'
    Check 'ExpectationEditor.ApprovedCaption' ($caption -ceq 'Expected steps and conclusion')
    Check 'ExpectationEditor.RetryDefaultsToAllowed' ((ExpectationControl 'chkExpectedRetry' 'Value') -ceq 'True')
    if($caption -cne 'Expected steps and conclusion'){Write-Output ('Expectation editor caption observed: '+$caption)}
    foreach($size in @(@('Minimum',760,550),@('Large',1040,750),@('Default',900,630))){
        $layout=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingExpectationLayoutForTest' @('frmActionPathExpectation',[single]$size[1],[single]$size[2]))
        Check ('ExpectationEditor.Layout.'+$size[0]) ($layout -ceq 'FITS')
        if($layout -cne 'FITS'){Write-Output ('Expectation geometry '+$size[0]+': '+$layout)}
        if($opened -ceq 'DELIVERED'){
            Capture-ExpectationEditorEvidence $size[0]
        }
    }
    foreach($step in @(@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED'),@('RECEIVING_CONFIRM_WRITES','PENDING'))){
        [void](ExpectationControl 'cboExpectedControl' 'Select' $step[0])
        [void](ExpectationControl 'cboExpectedOutcome' 'Select' $step[1])
        if($step[0] -ceq 'RECEIVING_CONFIRM_WRITES'){
            [void](ExpectationControl 'chkExpectedRetry' 'Boolean' 'False')
        }
        [void](ExpectationControl 'btnAddExpectedStep' 'Click')
    }
    [void](ExpectationControl 'lstExpectedSteps' 'Index' '0')
    $firstId=ExpectationControl 'lstExpectedSteps' 'Value'
    [void](ExpectationControl 'lstExpectedSteps' 'Index' '1')
    $secondId=ExpectationControl 'lstExpectedSteps' 'Value'
    [void](ExpectationControl 'btnExpectedStepUp' 'Click')
    [void](ExpectationControl 'cboTerminalStep' 'Index' '1')
    [void](ExpectationControl 'cboTerminalKind' 'Select' 'CommandCompleted')
    if($opened -ceq 'DELIVERED'){Capture-ExpectationEditorEvidence 'Ordered'}
    $used=ExpectationControl 'btnUseExpectation' 'Click'
    $reopened=ExpectationControl 'btnRecordingExpectation' 'Click' '' 'frmInventoryViewer'
    # Leave the modeless editor open across the end of this exact sequence.
    [void](RecordingControl 'Stop Recording' 'Click')
    $closed=@(RecordingJournal $activeSequence|Where-Object RecordType -CEQ 'Close')
    if($closed.Count -ne 1){throw 'Actual editor-binding run did not close.'}
    $preserved=$false
    if($null -ne $closed[0].PSObject.Properties['ExpectedConclusion']){
        $steps=@($closed[0].ExpectedConclusion.Steps)
        $preserved=$steps.Count -eq 2 -and $firstId -ne '' -and $firstId -cne $secondId -and
            $steps[0].StepId -ceq $secondId -and $steps[1].StepId -ceq $firstId -and
            $steps[0].ControlId -ceq 'RECEIVING_CONFIRM_WRITES' -and $steps[1].ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and
            $closed[0].ExpectedConclusion.TerminalStepId -ceq $firstId
    }
    Check 'ExpectationEditor.ReorderPreservesExactStepIdsAndTerminal' ($opened -ceq 'DELIVERED' -and $used -ceq 'DELIVERED' -and $preserved)
    $retryPreserved=$false
    if($preserved){$retryPreserved=$steps[0].RetryAllowed -ceq $false -and $steps[1].RetryAllowed -ceq $true}
    Check 'ExpectationEditor.DefaultAndExplicitRetryPersistAfterReorder' $retryPreserved
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual second recording unavailable.'}
    $next=SaveRecordedSetting '682'
    [void](ExpectationControl 'btnUseExpectation' 'Click')
    [void](RecordingControl 'Stop Recording' 'Click')
    $nextClose=@(RecordingJournal $next.Attempt.SequenceId|Where-Object RecordType -CEQ 'Close')
    if($nextClose.Count -ne 1){throw 'Actual second run did not close.'}
    $none=$false
    if($null -ne $nextClose[0].PSObject.Properties['ExpectedConclusion']){
        $definition=$nextClose[0].ExpectedConclusion
        $none=$definition.TerminalKind -ceq 'None' -and $definition.TerminalStepId -ceq '' -and @($definition.Steps).Count -eq 0
    }
    Check 'ExpectationEditor.StaleEditorCannotRetargetAnotherSequence' ($reopened -ceq 'DELIVERED' -and $activeSequence -cne $next.Attempt.SequenceId -and $none)
    # Context loss must clear or close an already open expectation draft too.
    [void](RecordingControl 'Start Recording' 'Click')
    $opened=ExpectationControl 'btnRecordingExpectation' 'Click' '' 'frmInventoryViewer'
    [void](ExpectationControl 'cboExpectedControl' 'Select' 'ADMIN_SETTINGS_SAVE_VALUE')
    [void](ExpectationControl 'cboExpectedOutcome' 'Select' 'COMPLETED')
    [void](ExpectationControl 'btnAddExpectedStep' 'Click')
    $beforeCount=ExpectationControl 'lstExpectedSteps' 'Count'
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    [void](ExpectationControl 'btnUseExpectation' 'Click')
    $count=ExpectationControl 'lstExpectedSteps' 'CountAnyVisibility'
    Check 'ExpectationEditor.SignOutInvalidatesDraft' ($opened -ceq 'DELIVERED' -and $beforeCount -ceq '1' -and $count -cin @('MISSING','0'))
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    OpenRecordingViewer
    $excel.Visible=$previousVisibility
}
