# Calibrate transport, ordinary exception handling, detach and output redaction.
[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$root=Join-Path $repo ('reports/runtime/native-exception-calibration/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$binary=Join-Path $root 'NativeExceptionObserver.exe'
$compiler=New-Object CodeDom.Compiler.CompilerParameters -Property @{CompilerOptions='/platform:x64';GenerateExecutable=$true;OutputAssembly=$binary}
[void]$compiler.ReferencedAssemblies.Add('System.dll')
Add-Type -Path (Join-Path $PSScriptRoot 'NativeExceptionObserver.cs') -CompilerParameters $compiler
$checks=[Collections.Generic.List[object]]::new()
foreach($mode in @('Exit','Detach','FaultsOnly')){
    $case=Join-Path $root $mode
    New-Item -ItemType Directory -Path $case|Out-Null
    if($mode -eq 'Detach'){[IO.File]::WriteAllText((Join-Path $case 'test-detach'),'Test')}
    $target=Start-Process -FilePath $binary -ArgumentList @('calibrate',('"'+$case+'"')) -WindowStyle Hidden -PassThru
    $observerArgs=@([string]$target.Id,('"'+$case+'"'),'30')
    if($mode -eq 'FaultsOnly'){$observerArgs+='faults-only'}
    $observer=Start-Process -FilePath $binary -ArgumentList $observerArgs -WindowStyle Hidden -PassThru
    try {
        $deadline=[DateTime]::UtcNow.AddSeconds(20)
        while(-not (Test-Path (Join-Path $case 'continued')) -and [DateTime]::UtcNow -lt $deadline -and -not $observer.HasExited){Start-Sleep -Milliseconds 100}
        if($mode -eq 'Detach'){
            [IO.File]::WriteAllText((Join-Path $case 'stop'),'Stop')
            if(-not $observer.WaitForExit(5000)){throw 'Observer failed to detach.'}
            $checks.Add([pscustomobject]@{Check='Detach.TargetRemainsAlive';Passed=(-not $target.HasExited)})
            [IO.File]::WriteAllText((Join-Path $case 'target-exit'),'Exit')
        }
        if(-not $target.WaitForExit(5000) -or -not $observer.WaitForExit(5000)){throw 'Calibration child did not exit.'}
        $raw=[IO.File]::ReadAllText((Join-Path $case 'native-events.jsonl'))
        $events=@($raw -split '\r?\n'|Where-Object {$_}|ForEach-Object {$_|ConvertFrom-Json})
        $exceptions=@($events|Where-Object {$_.PSObject.Properties['Code']})
        $custom=@($exceptions|Where-Object Code -CEQ 'E042BEEF')
        $states=@($events|Where-Object {$_.PSObject.Properties['State']}|ForEach-Object State)
        $checks.Add([pscustomobject]@{Check="$mode.ProcessExitCodes";Passed=($target.ExitCode -eq 0 -and $observer.ExitCode -eq 0)})
        $checks.Add([pscustomobject]@{Check="$mode.KnownFirstChance";Passed=($custom.Count -eq 1 -and $custom[0].FirstChance)})
        $checks.Add([pscustomobject]@{Check="$mode.ModuleOffsetResolved";Passed=($custom.Count -eq 1 -and $custom[0].Module -eq 'kernelbase.dll' -and $custom[0].Offset -match '^[0-9A-F]+$')})
        $checks.Add([pscustomobject]@{Check="$mode.TargetHandledException";Passed=(Test-Path (Join-Path $case 'handled'))})
        $checks.Add([pscustomobject]@{Check="$mode.Ready";Passed=('Ready' -cin $states)})
        $checks.Add([pscustomobject]@{Check="$mode.Completed";Passed=($(if($mode -eq 'Detach'){'Detached'}else{'Exited'}) -cin $states)})
        $checks.Add([pscustomobject]@{Check="$mode.DeclaredFilter";Passed=($(if($mode -eq 'FaultsOnly'){'FaultsOnly'}else{'AllExceptions'}) -cin $states)})
        $noticeCount=@($exceptions|Where-Object Code -CEQ '00000005').Count
        $checks.Add([pscustomobject]@{Check="$mode.FilterPreservesHandling";Passed=((Test-Path (Join-Path $case 'filtered-handled')) -and $noticeCount -eq $(if($mode -eq 'FaultsOnly'){0}else{1}))})
        $checks.Add([pscustomobject]@{Check="$mode.NoObserverErrors";Passed=(@($events|Where-Object {$_.PSObject.Properties['Win32Error'] -and $_.Win32Error -ne 0}).Count -eq 0)})
        $checks.Add([pscustomobject]@{Check="$mode.NoPayloadOrPaths";Passed=($raw -notmatch 'SENTINEL|[A-Za-z]:\\|ExceptionAddress|Parameters|ProcessId|ThreadId')})
        $allowed=@('UTC','State','Win32Error','Code','FirstChance','Module','Offset')
        $checks.Add([pscustomobject]@{Check="$mode.ExactMetadataSchema";Passed=(@($events|ForEach-Object {$_.PSObject.Properties.Name}|Where-Object {$_ -cnotin $allowed}).Count -eq 0)})
    } finally {
        [IO.File]::WriteAllText((Join-Path $case 'stop'),'Stop')
        [IO.File]::WriteAllText((Join-Path $case 'target-exit'),'Exit')
        if(-not $observer.HasExited){[void]$observer.WaitForExit(5000)}
        if(-not $target.HasExited){[void]$target.WaitForExit(5000)}
        if(-not $observer.HasExited){Stop-Process -Id $observer.Id}
        if(-not $target.HasExited){Stop-Process -Id $target.Id}
    }
}
$checks|ConvertTo-Json|Set-Content (Join-Path $root 'checks.json')
$checks|Format-Table -AutoSize
Write-Output ('Evidence: '+$root)
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
