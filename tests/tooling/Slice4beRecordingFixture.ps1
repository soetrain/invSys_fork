# Shared recording fixture helpers; same handlers and journal assertions in both controllers.
function RecordingControl([string]$Caption,[string]$Operation='State') {
    [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingControlForTest' @($Caption,$Operation))
}
function RecordingStatus { [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingStatusForTest') }
function OpenRecordingViewer {
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Events',''))
}
function CloseRecordingViewer { [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest') }
function SetRecordingPolicy([bool]$Enabled) {
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.RecordingCapturePolicyForTest' @($Enabled))) {
            throw 'Actual recording policy fixture save failed; not behavioral RED.'
        }
    } finally { [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings') }
}
$activityRoot=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
function ActivityPins {
    $pins=@{}
    if(Test-Path -LiteralPath $activityRoot) {
        foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File) {
            $pins[$file.Name]=(Get-FileHash -LiteralPath $file.FullName).Hash
        }
    }
    return $pins
}
function SaveRecordedSetting([string]$Value) {
    $before=ActivityPins
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize',$Value))) {
            throw 'Actual Admin Save Value fixture failed; not behavioral RED.'
        }
    } finally { [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings') }
    $records=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File) {
        if(-not $before.ContainsKey($file.Name)) { [IO.File]::ReadAllText($file.FullName)|ConvertFrom-Json }
    })
    $attempts=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
    $outcomes=@($records|Where-Object OutcomeCode -CEQ 'COMPLETED')
    if($records.Count -ne 2 -or $attempts.Count -ne 1 -or $outcomes.Count -ne 1 -or
        $attempts[0].ControlId -cne 'ADMIN_SETTINGS_SAVE_VALUE' -or
        $attempts[0].ActivityId -cne $outcomes[0].ActivityId) {
        throw 'Existing Admin activity fixture is invalid; not recording RED.'
    }
    [pscustomobject]@{Attempt=$attempts[0];Outcome=$outcomes[0]}
}
function HasSequence($Action,[int]$Ordinal) {
    $id=[guid]::Empty
    return ([guid]::TryParse([string]$Action.Attempt.SequenceId,[ref]$id) -and
        $id -ne [guid]::Empty -and $Action.Attempt.SequenceId -ceq $Action.Outcome.SequenceId -and
        $Action.Attempt.Ordinal -eq $Ordinal -and $Action.Outcome.Ordinal -eq $Ordinal)
}
function PinsRetained($Pins) {
    foreach($name in $Pins.Keys) {
        if((Get-FileHash -LiteralPath (Join-Path $activityRoot $name)).Hash -cne $Pins[$name]) { return $false }
    }
    return $true
}
$journalRoot=Join-Path $Fixture.Root ('Training/ActionPaths/'+$Fixture.Warehouse)
function RecordingJournal([string]$Sequence) {
    if(-not $Sequence -or -not (Test-Path -LiteralPath $journalRoot)){return}
    foreach($file in Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File) {
        if($file.Length -gt 1048576){throw 'Recording product wrote an oversized record.'}
        $text=[IO.File]::ReadAllText($file.FullName)
        $record=$text|ConvertFrom-Json
        if($record.SequenceId -ceq $Sequence){$record}
    }
}
function JournalFact([string]$Sequence,[string]$Type,[int]$Count,[string]$Lifecycle='Recording') {
    $entries=@(RecordingJournal $Sequence|Where-Object {$_.RecordType -ceq $Type}|Sort-Object Version)
    if(-not $entries.Count){return $false}
    $record=$entries[-1]
    # D18 retains schema-1 reads while requiring new schema-2 writes separately.
    # Keep these lifecycle/identity facts applicable to both supported envelopes.
    $supportedSchema=$record.SchemaVersion -eq 1
    if($record.SchemaVersion -eq 2 -and $null -ne $record.PSObject.Properties['ExpectedConclusion']){
        $definition=$record.ExpectedConclusion
        $supportedSchema=$definition.SchemaVersion -eq 1 -and
            $definition.TerminalKind -cin @('None','CommandCompleted','SourceEventsApplied')
        if($Type -cne 'Close' -or $definition.TerminalKind -ceq 'None'){
            $supportedSchema=$supportedSchema -and $definition.TerminalKind -ceq 'None' -and
                $definition.TerminalStepId -ceq '' -and @($definition.Steps).Count -eq 0
        }
    }
    return ($supportedSchema -and $record.RecordKind -ceq 'Recording' -and
        $record.Lifecycle -ceq $Lifecycle -and $record.ActionCount -eq $Count -and
        $record.WarehouseId -ceq $Fixture.Warehouse -and $record.CreatedByUserId -ceq 'config-admin' -and
        $record.ActionPathId -cne $record.SequenceId -and $record.RecordId -cne $record.ActionPathId)
}
function JournalChain([string]$Sequence,[int]$ExpectedCount) {
    $entries=@(RecordingJournal $Sequence|Sort-Object Version)
    if($entries.Count -ne $ExpectedCount){return $false}
    $previousId='';$previousHash='';$version=0
    foreach($entry in $entries){
        $version++
        $path=Join-Path $journalRoot ($entry.ActionPathId+'.'+$version+'.json')
        if(-not (Test-Path -LiteralPath $path)){return $false}
        $text=[IO.File]::ReadAllText($path);$marker=$text.LastIndexOf(',"ContentSha256":"',[StringComparison]::Ordinal)
        if($marker -lt 0){return $false}
        $body=$text.Substring(0,$marker)+'}'
        $sha=[Security.Cryptography.SHA256]::Create()
        try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
        if($entry.Version -ne $version -or $entry.PreviousRecordId -cne $previousId -or $entry.PreviousSha256 -cne $previousHash -or $entry.ContentSha256 -cne $hash){return $false}
        $previousId=$entry.RecordId;$previousHash=$hash
    }
    return $true
}
function RestartPins([string]$Root) {
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $Root -Recurse -File) {
        $pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash
    }
    return $pins
}
function RestartPinsEqual($Before,[string]$Root) {
    $after=RestartPins $Root
    if($after.Count -ne $Before.Count){return $false}
    foreach($path in $Before.Keys){if(-not $after.ContainsKey($path) -or $after[$path] -cne $Before[$path]){return $false}}
    return $true
}
