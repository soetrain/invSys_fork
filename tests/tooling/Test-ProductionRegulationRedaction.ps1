[CmdletBinding()]
param([ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Slice4beProductionRegulationActivity.ps1')
$cases=@(
    @{Name='FixedFacts';Record=@{UserMessage='Local staging finished';SourceEventRefs=@()};Expected=$true},
    @{Name='IdentitySubstringInRecordId';Record=@{RecordId='11110012-1111-4111-8111-111111111111'};Expected=$true},
    @{Name='IdentitySubstringInTimestamp';Record=@{OccurredAtUTC='2026-01-01T00:00:00.001Z'};Expected=$true},
    @{Name='IdentitySubstringInHash';Record=@{ContentSha256=('a'*29+'001'+'b'*32)};Expected=$true},
    @{Name='WholeIdentityValue';Record=@{Unexpected='001'};Expected=$false},
    @{Name='IdentityInMessage';Record=@{UserMessage='Changed output A01'};Expected=$false},
    @{Name='IdentityInGuidance';Record=@{NextStep='Inspect NODE1'};Expected=$false},
    @{Name='IdentityInNestedValue';Record=@{Unexpected=@{Selected='001'}};Expected=$false},
    @{Name='EscapedSensitiveText';Raw='{"UserMessage":"Do not emit TEST\u005fCANARY"}';Expected=$false},
    @{Name='SensitiveAnywhere';Record=@{Unexpected='contains TEST_CANARY here'};Expected=$false}
)
$results=@(foreach($case in $cases){
    $raw=if($case.ContainsKey('Raw')){$case.Raw}else{$case.Record|ConvertTo-Json -Depth 5 -Compress}
    $actual=Test-RegulationRecordRedacted $raw @('TEST_CANARY') @('001','A01','NODE1')
    [pscustomobject]@{Name=$case.Name;Passed=($actual -eq $case.Expected)}
})
$root=Join-Path 'reports/runtime/production-regulation-redaction' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$results|ConvertTo-Json|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
$failed=@($results|Where-Object {-not $_.Passed}).Count
Write-Output ($root+': '+($results.Count-$failed)+' PASS / '+$failed+' FAIL')
if($failed){exit 1}
