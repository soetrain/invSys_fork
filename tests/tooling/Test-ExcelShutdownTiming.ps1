[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$WorkerScript)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'ExcelShutdownTiming.ps1')
$results=[Collections.Generic.List[object]]::new()
function Assert([string]$Name,[bool]$Passed){$results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})}
$root=Join-Path ([IO.Path]::GetTempPath()) ('invsys-shutdown-trace-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$script:ExcelShutdownTimingRoot=$root
Write-ExcelShutdownTiming BeforeGC
Write-ExcelShutdownTiming AfterGC
Write-ExcelShutdownTiming BeforeFinalizerWait
Write-ExcelShutdownTiming AfterFinalizerWait
$rows=@(Get-Content (Join-Path $root 'phases.jsonl')|ForEach-Object {$_|ConvertFrom-Json})
Assert 'SeparateCollectAndFinalizerIntervals' (($rows.Stage -join '|') -ceq 'BeforeGC|AfterGC|BeforeFinalizerWait|AfterFinalizerWait')
Assert 'MonotonicTiming' ($rows[3].Ticks -ge $rows[0].Ticks -and $rows[0].Frequency -gt 0)
Assert 'OnlyFixedMetadataFields' ((@($rows[0].PSObject.Properties.Name|Sort-Object) -join '|') -ceq 'Frequency|Kind|Ordinal|Stage|Ticks|UTC')
$rejected=$false
try{Write-ExcelShutdownTiming 'unsafe-fixture-content'}catch{$rejected=$true}
Assert 'RejectUnapprovedStageContent' ($rejected -and @(Get-Content (Join-Path $root 'phases.jsonl')).Count -eq 4)
$source=[IO.File]::ReadAllText((Resolve-Path $WorkerScript))
$generated=ConvertTo-ExcelShutdownTimedWorker $source 'C:\synthetic tooling' (Join-Path $PSScriptRoot 'ExcelShutdownTiming.ps1') $root
$tokens=$null;$errors=$null
[void][Management.Automation.Language.Parser]::ParseInput($generated,[ref]$tokens,[ref]$errors)
Assert 'GeneratedWorkerParses' ($errors.Count -eq 0)
foreach($call in @('[GC]::Collect()','[GC]::WaitForPendingFinalizers()','$excel.Quit()','[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)')){
    Assert ('OriginalCallCountPreserved.'+$call) ([regex]::Matches($source,[regex]::Escape($call)).Count -eq [regex]::Matches($generated,[regex]::Escape($call)).Count)
}
Assert 'AllHelperPathsStayAtOriginalSource' (-not $generated.Contains('$PSScriptRoot') -and $generated.Contains("'C:\synthetic tooling'"))
$rejected=$false
try{[void](ConvertTo-ExcelShutdownTimedWorker ($source.Replace('[GC]::WaitForPendingFinalizers()','[GC]::KeepAlive($null)')) 'synthetic' 'synthetic' $root)}catch{$rejected=$true}
Assert 'AnchorDriftRefusesBeforeDispatch' $rejected
Initialize-ExcelExitObserver
$child=Start-Process powershell.exe -WindowStyle Hidden -PassThru -ArgumentList @('-NoProfile','-NonInteractive','-Command','Start-Sleep -Seconds 2')
$created=$child.StartTime.ToUniversalTime().ToFileTimeUtc()
$rejected=$false
try{[InvSysExcelExitObserver]::Watch($child.Id,($created+1),(Join-Path $root 'wrong-owner.json'))}catch{$rejected=$true}
Assert 'RejectDifferentCreationIdentity' $rejected
[InvSysExcelExitObserver]::Watch($child.Id,$created,(Join-Path $root 'native-exit.json'))
if(-not $child.WaitForExit(10000)){throw 'Synthetic observer child did not exit.'}
for($attempt=0;$attempt -lt 30 -and -not(Test-Path (Join-Path $root 'native-exit.json'));$attempt++){Start-Sleep -Milliseconds 100}
$native=Get-Content -Raw (Join-Path $root 'native-exit.json')|ConvertFrom-Json
Assert 'IndependentNativeExitProof' ($native.WaitResult -eq 0 -and $native.TimesAvailable -and $native.CreatedFileTime -eq $created -and $native.ProcessId -eq $child.Id)
Assert 'NativeExitPrecedesObservation' ([DateTimeOffset]::Parse($native.NativeExitUTC) -le [DateTimeOffset]::Parse($native.ObservedUTC))
$results|Format-Table -AutoSize
if(@($results|Where-Object {-not $_.Passed}).Count){throw 'Shutdown timing checks failed.'}
