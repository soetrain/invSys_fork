[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$DeployRoot,
    [Parameter(Mandatory=$true)][string]$RunRoot,
    [switch]$RunApproved
)

$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$worker=Join-Path $PSScriptRoot 'Run-D14PackagedCreate.ps1'
$lifecycle=Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1'
$deploy=(Resolve-Path -LiteralPath $DeployRoot).Path
$runPath=[IO.Path]::GetFullPath($RunRoot)
$names=@('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')
if(Test-Path -LiteralPath $runPath){throw 'D14 run root already exists.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel is open.'}
foreach($path in @($worker,$lifecycle)){
    $tokens=$null;$errors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
    if(@($errors).Count){throw "PowerShell parse failed: $path"}
}
$before=@{}
foreach($name in $names){
    $path=Join-Path $deploy $name
    if(-not(Test-Path -LiteralPath $path)){throw "Package missing: $name"}
    $before[$name]=(Get-FileHash -LiteralPath $path).Hash
}
if(-not $RunApproved){
    Write-Output 'CHECKS ONLY: D14 packaged worker parsed, five packages pinned in memory, Excel closed, output absent. Nothing launched.'
    return
}

. $lifecycle
$settings=Get-InvSysTestSettingsSnapshot
$restored=$false
try{
New-Item -ItemType Directory -Path $runPath | Out-Null
$out=Join-Path $runPath 'artifacts'
([pscustomobject]@{Utc=[datetime]::UtcNow.ToString('o');WorkerSHA256=(Get-FileHash $worker).Hash;PackageHashes=$before}|ConvertTo-Json -Depth 3)|Set-Content -LiteralPath (Join-Path $runPath 'dispatch.json')
$psi=[Diagnostics.ProcessStartInfo]::new()
$psi.FileName='powershell.exe'
$psi.Arguments='-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$worker+'" -DeployRoot "'+$deploy+'" -OutputRoot "'+$out+'"'
$psi.WorkingDirectory=$PSScriptRoot
$psi.UseShellExecute=$false
$psi.CreateNoWindow=$true
$psi.RedirectStandardOutput=$true
$psi.RedirectStandardError=$true
$child=[Diagnostics.Process]::Start($psi)
$stdoutTask=$child.StandardOutput.ReadToEndAsync()
$stderrTask=$child.StandardError.ReadToEndAsync()
$workerExited=$child.WaitForExit(1200000)
if(-not $workerExited){throw 'D14 worker exceeded 20 minutes; no process was terminated.'}
$exitCode=$child.ExitCode
[IO.File]::WriteAllText((Join-Path $runPath 'worker.stdout.txt'),$stdoutTask.Result,[Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $runPath 'worker.stderr.txt'),$stderrTask.Result,[Text.UTF8Encoding]::new($false))
$workerResultPath=Join-Path $out 'result.json'
if(-not(Test-Path -LiteralPath $workerResultPath)){throw 'D14 worker exited without a result; no retry.'}
$result=Get-Content -LiteralPath $workerResultPath -Raw|ConvertFrom-Json
$pidExact=[int]$result.ExcelPid
$nativeExit=$false
$until=[datetime]::UtcNow.AddMinutes(10)
do{
    if(-not(Get-Process -Id $pidExact -ErrorAction SilentlyContinue)){$nativeExit=$true;break}
    Start-Sleep -Milliseconds 500
}while([datetime]::UtcNow -lt $until)
$restored=$false
if($nativeExit){$restored=Restore-InvSysTestSettingsSnapshot $settings}
$preserved=$true
foreach($name in $names){if((Get-FileHash -LiteralPath (Join-Path $deploy $name)).Hash -cne $before[$name]){$preserved=$false}}
$expected=@('D14.GeneratedSchemas.ManagedHeadersNoROW','D14.Seed.UniqueKeysConditionGood','D14.RoundTrip.PreservesSystemKey','D14.AdminSeed.RepeatedCreatesUniqueKeysNoMigration','D14.OperatorRefresh.PreservesCustomColumn','D14.OperatorReopen.PreservesCustomColumn')
$observed=@($result.Checks|ForEach-Object {$_.Name})
$exact=[string]::Join('|',$observed) -ceq [string]::Join('|',$expected)
$qualified=$exitCode -eq 0 -and [bool]$result.QuitIssued -and $nativeExit -and $restored -and $preserved -and $exact -and [int]$result.Passed -eq 6 -and [int]$result.Failed -eq 0
$closure=[pscustomobject]@{Utc=[datetime]::UtcNow.ToString('o');WorkerExitCode=$exitCode;ExcelPid=$pidExact;QuitIssued=$result.QuitIssued;ExcelNaturalExit=$nativeExit;SettingsRestored=$restored;PackagesPreserved=$preserved;ExactSixChecks=$exact;Passed=$result.Passed;Failed=$result.Failed;Qualified=$qualified}
$closure|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $runPath 'closure.json')
$closure|ConvertTo-Json -Depth 4
if(-not $qualified){throw 'D14 packaged gate did not qualify; no retry.'}
Write-Output 'D14 PACKAGED GATE: six exact checks, natural exit, settings and packages preserved.'
}finally{
    if(-not $restored){
        $until=[datetime]::UtcNow.AddMinutes(10)
        do{
            $remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
            if($remaining.Count -eq 0){break}
            Start-Sleep -Milliseconds 500
        }while([datetime]::UtcNow -lt $until)
        if($remaining.Count -eq 0){$restored=Restore-InvSysTestSettingsSnapshot $settings}
        else{Write-Warning 'D14 worker ended with Excel still running; settings restoration remains pending.'}
    }
    if(Test-Path -LiteralPath $runPath){
        [pscustomobject]@{Utc=[datetime]::UtcNow.ToString('o');ExcelClosed=(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0);SettingsRestored=$restored} |
            ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $runPath 'finally.json')
    }
}
