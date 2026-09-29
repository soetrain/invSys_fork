[CmdletBinding()]
param([string]$RepoRoot='.',[string]$OutputDirectory='reports/runtime/launcher-cleanup-evidence')
$ErrorActionPreference='Stop'; Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$validator=Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'
$source=[IO.File]::ReadAllText($validator)
$final=[regex]::Matches($source,'(?ms)^    \$cleanupTerminationRequested=\$false\r?\n.*?(?=^    \$tempRoot =)')
$restart=[regex]::Matches($source,'(?ms)^        \$excel = \$null\r?\n(?<body>.*?)(?=^        \$opened = New-Object)')
if($final.Count -ne 1 -or $restart.Count -ne 1){throw 'Cleanup source boundaries are ambiguous.'}
$blocks=@{Final=[scriptblock]::Create($final[0].Value);Restart=[scriptblock]::Create($restart[0].Groups['body'].Value)}
$root=Join-Path (Join-Path $repo $OutputDirectory) ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){
    $checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))
}
# Execute the validator's actual decision blocks. Replace only external process
# operations; no Excel, wait, process termination or business action occurs.
function Start-Sleep([int]$Milliseconds){}
function Get-Process([int]$Id,[string]$ErrorAction){$script:fixtureProcess}
function Stop-Process([int]$Id,[switch]$Force){
    if($Id -ne 123456789 -or -not $Force){throw 'Unexpected synthetic termination.'}
    $script:stopCalls++
}
foreach($stage in @('Restart','Final')){
    $cases=@('AlreadyGone','Termination','DifferentProcess','UnknownProcess')
    if($stage -eq 'Final'){$cases+='PaletteWaitExit'}
    foreach($case in $cases){
        $outputPath=Join-Path $root ($stage+'-'+$case)
        New-Item -ItemType Directory -Path $outputPath|Out-Null
        $excelProcessId=123456789
        if($case -eq 'UnknownProcess'){$excelProcessId=0}
        $script:fixtureProcess=$null;$script:stopCalls=0
        if($case -in @('Termination','PaletteWaitExit','DifferentProcess')){
            $name=if($case -eq 'DifferentProcess'){'OTHER'}else{'EXCEL'}
            $script:fixtureProcess=[pscustomobject]@{ProcessName=$name}
            $script:fixtureProcess|Add-Member -MemberType ScriptMethod -Name WaitForExit -Value {param($Milliseconds) $true}
        }
        $ProductionPaletteProbe=($case -eq 'PaletteWaitExit');$paletteBooksClosed=$true
        $packages=@{};$opened=[Collections.Generic.List[object]]::new()
        $excel=$null;$configWb=$null;$authWb=$null;$packageWb=$null;$wb=$null
        . $blocks[$stage]
        $expected=($case -eq 'Termination')
        Check ($stage+'.'+$case+'.ExistingTerminationDecision') ($script:stopCalls -eq [int]$expected)
        $file=Join-Path $outputPath ($stage.ToLowerInvariant()+'-cleanup-observation.json')
        $exists=Test-Path -LiteralPath $file
        Check ($stage+'.'+$case+'.ReceiptExists') $exists
        $valid=$false
        if($exists){
            $record=Get-Content -LiteralPath $file -Raw|ConvertFrom-Json
            $keys=@($record.PSObject.Properties.Name|Sort-Object)
            $valid=($keys -join '|') -ceq 'ProcessIdAvailable|Stage|TerminationRequested'
            if($valid){$valid=$record.Stage -ceq $stage -and $record.ProcessIdAvailable -eq ($excelProcessId -gt 0) -and $record.TerminationRequested -eq $expected}
        }
        Check ($stage+'.'+$case+'.ExactBoundedEvidence') $valid
    }
}
$checks|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'checks.json')
[pscustomobject]@{ValidatorHash=(Get-FileHash -LiteralPath $validator).Hash;Passed=@($checks|Where-Object Passed).Count;Failed=@($checks|Where-Object {-not $_.Passed}).Count;NoLiveExcelOrTermination=$true}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'summary.json')
Write-Output ('Evidence: '+$root)
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
