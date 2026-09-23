# Calibrate the reader fixture's allowance; no product behavior is simulated.
[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
. (Join-Path $PSScriptRoot 'Slice4beVisibilitySavePins.ps1')
$root=Join-Path $repo ('reports/runtime/visibility-save-pins/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$results=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){
    $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))
}
foreach($scope in @('Warehouse','Training')){
    foreach($fault in @('ExactSave','UnexpectedFile','ChangedOriginal','MissingOriginal','WrongControl','WrongContext','WrongSequence','DamagedHash','MissingObservation')){
        $case=Join-Path $root ($scope+'-'+$fault)
        $activity=Join-Path $case 'Training/Activity/PIN-FIXTURE'
        New-Item -ItemType Directory -Path $activity -Force|Out-Null
        $fixture=[pscustomobject]@{Root=$case;Config=(Join-Path $case 'Config.fixture');Warehouse='PIN-FIXTURE'}
        [IO.File]::WriteAllText($fixture.Config,'original config')
        $original=Join-Path $activity 'existing.fixture'
        [IO.File]::WriteAllText($original,'original immutable bytes')
        $guardRoot=if($scope -ceq 'Warehouse'){$case}else{Join-Path $case 'Training'}
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $guardRoot -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
        $before=$pins.Clone()
        Check ($scope+'.'+$fault+'.InitialReadInterval') (Test-Slice4beReadPins $pins $guardRoot)
        [IO.File]::WriteAllText($fixture.Config,'deliberately saved config')
        $record=[ordered]@{SchemaVersion=1;ControlId='ADMIN_TRACKING_SAVE';OutcomeCode='REQUESTED';EventCode='ADMIN_TRACKING_SAVE_REQUESTED';OwnerId='CORE_CONFIGURATION';SourceRole='Admin';WarehouseId=$fixture.Warehouse;StationId='S1';UserId='config-admin';SequenceId='';Ordinal=0;DataEffect='Unknown';SourceEventRefs=@();RecordId=[guid]::NewGuid().ToString();ActivityId=[guid]::NewGuid().ToString()}
        switch($fault){
            'WrongControl' {$record.ControlId='ADMIN_DETAIL_SAVE'}
            'WrongContext' {$record.WarehouseId='OTHER-FIXTURE'}
            'WrongSequence' {$record.SequenceId=[guid]::NewGuid().ToString();$record.Ordinal=1}
        }
        $body=ConvertTo-Json -InputObject $record -Compress
        $sha=[Security.Cryptography.SHA256]::Create()
        try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body)))).Replace('-','').ToLowerInvariant()}
        finally{$sha.Dispose()}
        if($fault -ceq 'DamagedHash'){$hash='0'*64}
        if($fault -cne 'MissingObservation'){
            [IO.File]::WriteAllText((Join-Path $activity ($record.RecordId+'.json')),($body.Substring(0,$body.Length-1)+',"ContentSha256":"'+$hash+'"}'))
        }
        switch($fault){
            'UnexpectedFile' {[IO.File]::WriteAllText((Join-Path $activity 'unexpected.fixture'),'unexpected')}
            'ChangedOriginal' {[IO.File]::WriteAllText($original,'changed')}
            'MissingOriginal' {Remove-Item -LiteralPath $original}
        }
        $readUnchanged=$scope -ceq 'Training' -and $fault -ceq 'MissingObservation'
        Check ($scope+'.'+$fault+'.ReadIntervalMatchesScope') ((Test-Slice4beReadPins $pins $guardRoot) -eq $readUnchanged)
        $accepted=Update-Slice4beVisibilitySavePins $pins $guardRoot $fixture
        Check ($scope+'.'+$fault+'.ExactAllowance') ($accepted -eq ($fault -ceq 'ExactSave'))
        if($fault -ceq 'ExactSave'){
            Check ($scope+'.'+$fault+'.FollowingReadInterval') (Test-Slice4beReadPins $pins $guardRoot)
            [IO.File]::WriteAllText($original,'later reader mutation')
            Check ($scope+'.'+$fault+'.FollowingMutationRejected') (-not (Test-Slice4beReadPins $pins $guardRoot))
        }else{
            $unchanged=$pins.Count -eq $before.Count
            foreach($path in $before.Keys){$unchanged=$unchanged -and $pins[$path] -ceq $before[$path]}
            Check ($scope+'.'+$fault+'.RejectedDifferenceNeverAdvancesPins') $unchanged
        }
    }
}
ConvertTo-Json -InputObject @($results.ToArray())|Set-Content (Join-Path $root results.json)
Write-Output ('Report: '+$root)
if(@($results|Where-Object {-not $_.Passed}).Count){exit 1}
