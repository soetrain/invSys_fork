[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2

# Exercise the actual nested restart-fixture validator without Excel or authority files.
$source=Join-Path $PSScriptRoot 'Slice4beGuideRestartFixture.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if(@($errors).Count){throw 'Restart fixture does not parse.'}
$definitions=@($ast.FindAll({param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'FixtureRecording'
},$true))
if($definitions.Count -ne 1){throw 'Expected one real FixtureRecording function.'}
. ([scriptblock]::Create($definitions[0].Extent.Text))
$reportRoot=''

$actionCall=0
$journal=@()
function RecordingControl([string]$Caption,[string]$Operation) {'DELIVERED'}
function SaveRecordedSetting([string]$Value) {
    $script:actionCall++
    $ordinal=if($script:actionCall -eq 1){3}else{6}
    $one=[pscustomobject]@{SequenceId='test-sequence';Ordinal=$ordinal}
    [pscustomobject]@{Attempt=$one;Outcome=$one}
}
function RecordingJournal([string]$Sequence) {
    if($Sequence -cne 'test-sequence'){throw 'Wrong source sequence.'}
    $script:journal
}
function JournalChain([string]$Sequence,[int]$ExpectedCount) {
    $Sequence -ceq 'test-sequence' -and $ExpectedCount -eq $script:journal.Count
}
function HasSequence($Action,[int]$Ordinal) {
    $Action.Attempt.SequenceId -ceq 'test-sequence' -and
        $Action.Outcome.SequenceId -ceq 'test-sequence' -and
        $Action.Attempt.Ordinal -eq $Ordinal -and $Action.Outcome.Ordinal -eq $Ordinal
}

$cases=@(
    @{Name='SelectionInclusiveJournal';Entries=14;Observations=12;Actions=6;Lifecycle='Stopped';ExpectSuccess=$true},
    @{Name='MissingObservation';Entries=14;Observations=11;Actions=6;Lifecycle='Stopped';ExpectSuccess=$false},
    @{Name='MissingJournalVersion';Entries=13;Observations=12;Actions=6;Lifecycle='Stopped';ExpectSuccess=$false},
    @{Name='WrongActionCount';Entries=14;Observations=12;Actions=5;Lifecycle='Stopped';ExpectSuccess=$false},
    @{Name='IncompleteLifecycle';Entries=14;Observations=12;Actions=6;Lifecycle='Incomplete';ExpectSuccess=$false}
)
$failed=0
foreach($case in $cases) {
    $actionCall=0
    $journal=@(for($n=1;$n -le $case.Entries;$n++) {
        [pscustomobject]@{
            Version=$n
            Lifecycle=if($n -eq $case.Entries){$case.Lifecycle}else{'Recording'}
            ActionCount=$case.Actions
            Observations=if($n -eq $case.Entries){@(1..$case.Observations)}else{@()}
        }
    })
    $record=$null;$errorText=''
    try {$record=FixtureRecording '701' '702'} catch {$errorText=$_.Exception.Message}
    $passed=if($case.ExpectSuccess){$errorText -ceq '' -and $null -ne $record -and $record.Lifecycle -ceq 'Stopped'}
        else{$errorText -ceq 'Actual restart journal fixture is invalid; not behavioral RED.'}
    if(-not $passed){$failed++}
    Write-Output ('GuideRestartFixtureShape.'+$case.Name+': '+$(if($passed){'PASS'}else{'FAIL'}))
}
Write-Output ('Guide restart fixture shape: '+($cases.Count-$failed)+' PASS / '+$failed+' FAIL')
if($failed){exit 1}
