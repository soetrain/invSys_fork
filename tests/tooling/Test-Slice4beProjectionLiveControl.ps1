# Same-session diagnostic cut of the actual ordered live validator; not full R1.
[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-settings-diagnostic',
      [ValidateSet('BeforeProjection','AfterProjection')][string]$Cut='AfterProjection')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before diagnostic.'}
$root=Join-Path $repo ('reports/runtime/projection-live-control/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
Write-Output ('Report: '+$root)
$tempRoot=$root
$tokens=$null;$errors=$null
$chain=Join-Path $repo 'tools/validate_release1_full_chain.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($chain,[ref]$tokens,[ref]$errors)
$generator=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'New-OrderedLiveValidator'},$false)
if($errors.Count -or $null -eq $generator){throw 'Ordered generator unavailable.'}
# Transform in memory before the existing generator writes its diagnostic copy.
# Use an ephemeral synthetic credential, never a credential literal in that copy.
$definition=$generator.Extent.Text
$read='$source = Get-Content -LiteralPath $sourcePath -Raw'
if([regex]::Matches($definition,[regex]::Escape($read)).Count -ne 1){throw 'Generator source anchor differs.'}
$credentialTransform=@'
$credentialPattern='(?m)^\$testPin = "[0-9]+"\r?$'
if([regex]::Matches($source,$credentialPattern).Count -ne 1){throw 'Fixture credential anchor differs.'}
$source=[regex]::Replace($source,$credentialPattern,[Text.RegularExpressions.MatchEvaluator]{param($m) '$testPin = [guid]::NewGuid().ToString(''N'')'})
'@
$definition=$definition.Replace($read,$read+"`r`n"+$credentialTransform)
. ([scriptblock]::Create($definition))
$generated=New-OrderedLiveValidator
$source=Get-Content -LiteralPath $generated -Raw
$cutAnchor=if($Cut -ceq 'BeforeProjection'){'$currentStep = "Delete and rebuild canonical inventory projections"'}else{'$currentStep = "Run two consecutive Production batches through form actions"'}
$cutAt=$source.IndexOf($cutAnchor,[StringComparison]::Ordinal)
$tail=[regex]::Match($source,'(?m)^}\r?\ncatch \{')
if($cutAt -lt 0 -or -not $tail.Success -or $tail.Index -le $cutAt){throw 'Phase cut anchors differ.'}
$source=$source.Substring(0,$cutAt)+'Write-ControlMark ''CutReached'''+"`r`n"+$source.Substring($tail.Index)
$outputAnchor='$resultPath = Join-Path $repo "tests/unit/phase6_live_role_workflow_results.md"'
if([regex]::Matches($source,[regex]::Escape($outputAnchor)).Count -ne 1){throw 'Report anchor differs.'}
$source=$source.Replace($outputAnchor,'$resultPath = Join-Path $PSScriptRoot ''results.md''')
$detailAnchor='$lines += "| $($row.Check) | $result | $detail |"'
if([regex]::Matches($source,[regex]::Escape($detailAnchor)).Count -ne 1){throw 'Report redaction anchor differs.'}
$source=$source.Replace($detailAnchor,'$lines += "| $($row.Check) | $result | Details omitted |"')
$helpers=@'
function Write-ControlMark([string]$Stage){
    [pscustomobject]@{Stage=$Stage;UTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json -Compress|Add-Content (Join-Path $PSScriptRoot 'lifecycle.jsonl')
}
function Write-ControlOwner($Application){
    [uint32]$owner=0
    [void][InvSysLiveValidationWindow]::GetWindowThreadProcessId([IntPtr]$Application.Hwnd,[ref]$owner)
    $process=Get-Process -Id $owner
    [pscustomobject]@{ProcessId=$owner;StartUTC=$process.StartTime.ToUniversalTime().ToString('o')}|ConvertTo-Json|Set-Content (Join-Path $PSScriptRoot 'owner.json')
}
'@
$init='$repo = (Resolve-Path $RepoRoot).Path'
$source=$source.Replace($init,$helpers+"`r`n"+$init)
foreach($pair in @(
    @('$excel = New-Object -ComObject Excel.Application',"Write-ControlOwner `$excel"),
    @('Write-Output "PHASE6_LIVE_PROJECTION_DELETE_BEGIN"',"Write-ControlMark 'DeleteBegin'"),
    @('Write-Output "PHASE6_LIVE_PROJECTION_DELETE_END"',"Write-ControlMark 'DeleteEnd'"),
    @('Write-Output "PHASE6_LIVE_PROJECTION_RUN_BEGIN"',"Write-ControlMark 'ProcessorBegin'"),
    @('Write-Output "PHASE6_LIVE_PROJECTION_RUN_END"',"Write-ControlMark 'ProcessorEnd'"),
    @('try { $excel.Quit() } catch {}',"Write-ControlMark 'OriginalQuitReturned'"))) {
    if($source.Contains($pair[0])){$source=$source.Replace($pair[0],$pair[0]+"`r`n"+$pair[1])}
}
[IO.File]::WriteAllText($generated,$source,[Text.UTF8Encoding]::new($false))
$tokens=$null;$errors=$null
[void][Management.Automation.Language.Parser]::ParseFile($generated,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Generated diagnostic does not parse.'}
$deploy=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$pins=Get-Content (Join-Path $repo 'reports/runtime/settings-diagnostic-package-pins.json') -Raw|ConvertFrom-Json
foreach($pin in $pins){if((Get-FileHash -LiteralPath (Join-Path $deploy $pin.Package)).Hash -cne $pin.Hash){throw 'Candidate package differs.'}}
$tracked=Join-Path $repo 'tests/unit/phase6_live_role_workflow_results.md'
$trackedHash=(Get-FileHash -LiteralPath $tracked).Hash
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$child=$null;$restored=$false;$start=[DateTimeOffset]::UtcNow
[pscustomobject]@{Cut=$Cut;StartUTC=$start.ToString('o');GeneratedHash=(Get-FileHash -LiteralPath $generated).Hash;ChainHash=(Get-FileHash -LiteralPath $chain).Hash;LiveHash=(Get-FileHash -LiteralPath (Join-Path $repo 'tools/validate_phase6_live_role_workflows.ps1')).Hash;FullChainAccepted=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'start.json')
try {
    $arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$generated+'"'),'-RepoRoot',('"'+$repo+'"'),'-DeployRoot',('"'+$DeployRoot+'"'))
    $child=Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $root 'worker.stdout.log') -RedirectStandardError (Join-Path $root 'worker.stderr.log')
    $null=$child.Handle
    while(-not $child.HasExited){Start-Sleep -Milliseconds 500}
    $child.Refresh()
    [pscustomobject]@{ExitCode=$child.ExitCode;UTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content (Join-Path $root 'worker-exit.json')
    for($i=0;$i -lt 15 -and (Get-Process EXCEL -ErrorAction SilentlyContinue);$i++){Start-Sleep -Seconds 2}
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){
        [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');ClosurePending=$true}|ConvertTo-Json|Set-Content (Join-Path $root 'cleanup-wait.json')
        Wait-RecordingCleanup -Creator $null -Worker $child
    }
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    if(-not $restored){throw 'Settings restoration differs.'}
    Start-Sleep -Seconds 12
    $audit=[DateTimeOffset]::UtcNow;$queryErrors=@()
    $events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=$audit.LocalDateTime} -ErrorAction SilentlyContinue -ErrorVariable queryErrors)
    if(@($queryErrors|Where-Object FullyQualifiedErrorId -NotLike 'NoMatchingEventsFound*').Count){throw 'Application audit unavailable.'}
    foreach($pin in $pins){if((Get-FileHash -LiteralPath (Join-Path $deploy $pin.Package)).Hash -cne $pin.Hash){throw 'Package changed.'}}
    if((Get-FileHash -LiteralPath $tracked).Hash -cne $trackedHash){throw 'Tracked report changed.'}
    $checks=@();$resultFile=Join-Path $root 'results.md'
    if(Test-Path -LiteralPath $resultFile){$checks=@([regex]::Matches((Get-Content -LiteralPath $resultFile -Raw),'(?m)^\| ([^|]+) \| (PASS|FAIL) \|')|ForEach-Object {[pscustomobject]@{Check=$_.Groups[1].Value.Trim();Passed=$_.Groups[2].Value -ceq 'PASS'}})}
    $marks=@(Get-Content (Join-Path $root 'lifecycle.jsonl')|ForEach-Object {ConvertFrom-Json $_})
    $complete=@($marks|Where-Object Stage -CEQ 'CutReached').Count -eq 1
    $failed=@($checks|Where-Object {-not $_.Passed}).Count
    $result=[pscustomobject]@{Cut=$Cut;WorkerExit=$child.ExitCode;Checks=$checks;PassedChecks=$checks.Count-$failed;FailedChecks=$failed;CutReached=$complete;StartUTC=$start.ToString('o');AuditUTC=$audit.ToString('o');ApplicationEvents=$events.Count;SettingsRestored=$restored;PackagePins=5;TrackedReportUnchanged=$true;ExcelClosed=@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0;DiagnosticPassed=($child.ExitCode -eq 0 -and $complete -and $failed -eq 0 -and $events.Count -eq 0);FullChainAccepted=$false}
    $result|ConvertTo-Json -Depth 5|Set-Content (Join-Path $root 'result.json')
    $result|Select-Object Cut,WorkerExit,PassedChecks,FailedChecks,CutReached,ApplicationEvents,SettingsRestored,ExcelClosed,DiagnosticPassed|ConvertTo-Json
    if(-not $result.DiagnosticPassed){exit 1}
} finally {
    if(-not $restored){Wait-RecordingCleanup -Creator $null -Worker $child;$restored=Restore-InvSysTestSettingsSnapshot $settings;[pscustomobject]@{SettingsRestored=$restored}|ConvertTo-Json|Set-Content (Join-Path $root 'failure-restoration.json')}
}
