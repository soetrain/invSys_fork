# Calibrate the narrow test allowance; these are synthetic records, not product RED.
[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Slice4bePreferenceSavePins.ps1')
. (Join-Path $PSScriptRoot 'Slice4beVisibilitySavePins.ps1')
$root=Join-Path (Resolve-Path $RepoRoot).Path ('reports/runtime/preference-save-pins/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$results=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})}
function Pins([string]$Path){$p=@{};foreach($f in Get-ChildItem -LiteralPath $Path -File -Recurse){$p[$f.FullName]=(Get-FileHash -LiteralPath $f.FullName).Hash};return $p}
foreach($fault in @('ExactPair','ChangedConfig','ChangedOriginal','MissingOriginal','ExtraFile','WrongDirectory','WrongControl','WrongContext','WrongSequence','WrongOutcome','WrongEffect','DifferentActivity','DuplicateOutcome','DamagedHash','MissingRecord')){
    $case=Join-Path $root $fault;$area=Join-Path $case 'Training/Activity/PIN-FIXTURE'
    New-Item -ItemType Directory -Path $area -Force|Out-Null
    $fixture=[pscustomobject]@{Root=$case;Warehouse='PIN-FIXTURE';Config=(Join-Path $case 'Config.fixture')}
    [IO.File]::WriteAllText($fixture.Config,'immutable config')
    $original=Join-Path $case 'Training/prior.fixture';[IO.File]::WriteAllText($original,'immutable training')
    $before=Pins $case;$saved=$before.Clone();$activity=[guid]::NewGuid().ToString()
    foreach($outcome in @('REQUESTED','COMPLETED')){
        $r=[ordered]@{SchemaVersion=1;CatalogVersion=11;PolicyVersion=1;ControlId='VIEWER_PATH_PREFERENCE_SAVE';OwnerId='CORE_PERSONAL_PREFERENCE';SourceKind='User activity';SourceRole='Viewer';WarehouseId='PIN-FIXTURE';StationId='S1';UserId='config-reader';SequenceId='';Ordinal=0;SourceEventRefs=@();RecordId=[guid]::NewGuid().ToString();ActivityId=$activity;OutcomeCode=$outcome;EventCode=('VIEWER_PATH_PREFERENCE_SAVE_'+$outcome);DataEffect=$(if($outcome -ceq 'REQUESTED'){'Unknown'}else{'Changed'})}
        if($outcome -ceq 'COMPLETED'){
            switch($fault){
                WrongControl {$r.ControlId='VIEWER_PATH_PREFERENCE_RESET'}
                WrongContext {$r.WarehouseId='OTHER'}
                WrongSequence {$r.SequenceId=[guid]::NewGuid().ToString()}
                WrongOutcome {$r.OutcomeCode='FAILED';$r.EventCode='VIEWER_PATH_PREFERENCE_SAVE_FAILED'}
                WrongEffect {$r.DataEffect='Unknown'}
                DifferentActivity {$r.ActivityId=[guid]::NewGuid().ToString()}
                DuplicateOutcome {$r.OutcomeCode='REQUESTED';$r.EventCode='VIEWER_PATH_PREFERENCE_SAVE_REQUESTED';$r.DataEffect='Unknown'}
            }
        }
        $body=ConvertTo-Json -InputObject $r -Compress;$sha=[Security.Cryptography.SHA256]::Create()
        try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
        if($fault -ceq 'DamagedHash'){$hash='0'*64}
        $destination=if($fault -ceq 'WrongDirectory'){$case}else{$area}
        if($fault -cne 'MissingRecord' -or $outcome -ceq 'REQUESTED'){
            [IO.File]::WriteAllText((Join-Path $destination ($r.RecordId+'.json')),($body.Substring(0,$body.Length-1)+',"ContentSha256":"'+$hash+'"}'))
        }
    }
    switch($fault){
        ChangedConfig {[IO.File]::WriteAllText($fixture.Config,'changed')}
        ChangedOriginal {[IO.File]::WriteAllText($original,'changed')}
        MissingOriginal {Remove-Item -LiteralPath $original}
        ExtraFile {[IO.File]::WriteAllText((Join-Path $area 'extra.fixture'),'extra')}
    }
    $after=Pins $case;$accepted=Update-Slice4bePreferenceSavePins $before $after $fixture
    Check ($fault+'.ExactAllowance') ($accepted -eq ($fault -ceq 'ExactPair'))
    $expected=if($accepted){$after}else{$saved}
    $same=$before.Count -eq $expected.Count;foreach($path in $expected.Keys){$same=$same -and $before[$path] -ceq $expected[$path]}
    Check ($fault+'.PinsAdvanceOnlyAfterExactPair') $same
    if($fault -ceq 'ExactPair'){
        Check 'FollowingReadInterval.Strict' (Test-Slice4beReadPins $before $case)
        $newPath=@($after.Keys|Where-Object {-not $saved.ContainsKey($_)})[0]
        [IO.File]::WriteAllText($newPath,'later observation mutation')
        Check 'FollowingReadInterval.NewObservationMutationRejected' (-not (Test-Slice4beReadPins $before $case))
    }
}
$results|ConvertTo-Json|Set-Content (Join-Path $root results.json)
Write-Output ('Calibration: '+@($results|Where-Object Passed).Count+'/'+$results.Count+'; '+$root)
if(@($results|Where-Object {-not $_.Passed}).Count){exit 1}
