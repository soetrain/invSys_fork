[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2

# File-only contract check for the shared recording fixture. No Excel or XLAM is opened.
$fixtureScript=Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1'
$tempParent=[IO.Path]::GetTempPath().TrimEnd([IO.Path]::DirectorySeparatorChar)
$tempRoot=Join-Path $tempParent ('InvSysRecordingFixtureFiltering-'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $tempRoot)
$caseMode=''
$activityIndex=0
$caseRoot=''

function Write-TestActivity([string]$ControlId,[string]$OutcomeCode,[string]$ActivityId) {
    $script:activityIndex++
    $path=Join-Path $script:caseRoot ('Training/Activity/test-warehouse/'+('{0:D3}.json' -f $script:activityIndex))
    $record=[ordered]@{ControlId=$ControlId;OutcomeCode=$OutcomeCode;ActivityId=$ActivityId}
    [IO.File]::WriteAllText($path,($record|ConvertTo-Json -Compress),[Text.UTF8Encoding]::new($false))
}
function Write-SelectionPair {
    $id=[guid]::NewGuid().ToString('N')
    Write-TestActivity 'ADMIN_SETTINGS_SELECT_CONFIG' 'REQUESTED' $id
    Write-TestActivity 'ADMIN_SETTINGS_SELECT_CONFIG' 'SELECTED' $id
}
function Run([string]$Package,[string]$Procedure,[object[]]$Arguments) {
    if($Package -cne 'invSys.Admin.xlam'){throw 'Unexpected package in fixture test.'}
    switch -CaseSensitive ($Procedure) {
        'TestD5Commands.OpenSettings' {
            if($script:caseMode -cne 'BaselineOnly'){Write-SelectionPair}
            return $true
        }
        'TestD5Commands.SaveSettings' {
            $id=[guid]::NewGuid().ToString('N')
            Write-TestActivity 'ADMIN_SETTINGS_SAVE_VALUE' 'REQUESTED' $id
            $resultId=if($script:caseMode -ceq 'MismatchedSaveIds'){[guid]::NewGuid().ToString('N')}else{$id}
            Write-TestActivity 'ADMIN_SETTINGS_SAVE_VALUE' 'COMPLETED' $resultId
            if($script:caseMode -ceq 'DuplicateSavePair'){
                $other=[guid]::NewGuid().ToString('N')
                Write-TestActivity 'ADMIN_SETTINGS_SAVE_VALUE' 'REQUESTED' $other
                Write-TestActivity 'ADMIN_SETTINGS_SAVE_VALUE' 'COMPLETED' $other
            }
            return $true
        }
        'TestD5Commands.CloseSettings' {
            if($script:caseMode -cne 'BaselineOnly'){Write-SelectionPair}
            return $true
        }
        default {throw ('Unexpected procedure in fixture test: '+$Procedure)}
    }
}

$results=@()
try {
    foreach($mode in @('BaselineOnly','ExtraSelections','MismatchedSaveIds','DuplicateSavePair')) {
        $caseMode=$mode
        $activityIndex=0
        $caseRoot=Join-Path $tempRoot $mode
        $activityPath=Join-Path $caseRoot 'Training/Activity/test-warehouse'
        $reportRoot=Join-Path $caseRoot 'report'
        [void](New-Item -ItemType Directory -Path $activityPath -Force)
        [void](New-Item -ItemType Directory -Path $reportRoot -Force)
        $Fixture=[pscustomobject]@{Root=$caseRoot;Warehouse='test-warehouse'}
        . $fixtureScript
        $errorText=''
        $pair=$null
        try {$pair=SaveRecordedSetting '7'} catch {$errorText=$_.Exception.Message}
        $expectFailure=$mode -in @('MismatchedSaveIds','DuplicateSavePair')
        $files=@(Get-ChildItem -LiteralPath $activityPath -Filter '*.json' -File)
        $expectedFiles=switch($mode){'BaselineOnly'{2};'ExtraSelections'{6};'MismatchedSaveIds'{6};'DuplicateSavePair'{8}}
        $pairValid=$null -ne $pair -and $pair.Attempt.ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and
            $pair.Outcome.ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and
            $pair.Attempt.ActivityId -ceq $pair.Outcome.ActivityId
        $passed=if($expectFailure){$errorText -ceq 'Existing Admin activity fixture is invalid; not recording RED.'}else{$errorText -ceq '' -and $pairValid}
        $passed=$passed -and $files.Count -eq $expectedFiles
        $results+=[pscustomobject]@{Case=$mode;Passed=$passed;ActivityFiles=$files.Count;Error=$errorText}
        Write-Output ('RecordingFixtureFiltering.'+$mode+': '+$(if($passed){'PASS'}else{'FAIL'}))
    }
} finally {
    $resolved=[IO.Path]::GetFullPath($tempRoot)
    $parent=[IO.Path]::GetFullPath($tempParent).TrimEnd([IO.Path]::DirectorySeparatorChar)+[IO.Path]::DirectorySeparatorChar
    if(-not $resolved.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe temporary cleanup path.'}
    if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
}
$fail=@($results|Where-Object {-not $_.Passed})
Write-Output ('Recording fixture filtering: '+($results.Count-$fail.Count)+' PASS / '+$fail.Count+' FAIL')
if($fail.Count){exit 1}
