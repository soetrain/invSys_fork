[CmdletBinding()]
param([string]$RepoRoot='', [string]$DeployRoot='deploy/validation-recording-isolation',
    [ValidateSet('RED','GREEN')][string]$Phase='RED', [switch]$CaptureEvidence)
# Neutral coordinator: never creates Excel, never persists the fixture transfer.
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
if(-not $RepoRoot){$RepoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)}
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$deploy=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before recording controller validation.'}
$state=$null; $creator=$null; $worker=$null; $stream=$null; $failed=$false
$results=[Collections.Generic.List[object]]::new()
$settingsRoot='HKCU:\Software\VB and VBA Program Settings\invSys'; $registryBefore=@{}
if(Test-Path -LiteralPath $settingsRoot){
    foreach($key in @(Get-Item -LiteralPath $settingsRoot)+@(Get-ChildItem -LiteralPath $settingsRoot -Recurse)){
        $values=@{}
        foreach($name in $key.GetValueNames()){$values[$name]=@($key.GetValue($name),$key.GetValueKind($name))}
        $registryBefore[$key.Name]=$values
    }
}
function StartController([string]$Arguments,[bool]$InputPipe=$false){
    $info=[Diagnostics.ProcessStartInfo]::new()
    $info.FileName=Join-Path $PSHOME 'powershell.exe';$info.Arguments=$Arguments
    $info.UseShellExecute=$false;$info.CreateNoWindow=$true
    $info.RedirectStandardInput=$InputPipe;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
    $process=[Diagnostics.Process]::new();$process.StartInfo=$info
    if(-not $process.Start()){throw 'Controller could not start.'}
    return $process
}
function RecordCheck([string]$Name,[bool]$Passed){
    $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))
}
$name='invsys-recording-'+[guid]::NewGuid().ToString('N')
. (Join-Path $PSScriptRoot 'Slice4beRecordingTransfer.ps1')
$pipe=New-RecordingHandoffServer $name
try {
    $connect=$pipe.WaitForConnectionAsync()
    $argsText='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1')+'" -RepoRoot "'+$repo+'" -DeployRoot "'+$DeployRoot+'" -Phase '+$Phase+' -CheckRecordingRestart -RecordingContinuationPipeName '+$name
    if($CaptureEvidence){$argsText+=' -CaptureEvidence'}
    $creator=StartController $argsText
    $creatorErrors=$creator.StandardError.ReadToEndAsync()
    if(-not $connect.Wait(10000)){throw 'Creator continuation connection failed.'}
    $stream=[IO.StreamReader]::new($pipe,[Text.Encoding]::UTF8)
    $transfer=$stream.ReadToEndAsync()
    while($null -ne ($line=$creator.StandardOutput.ReadLine())){
        if($line -match '^([A-Za-z][A-Za-z0-9. ]+): (PASS|FAIL)$'){
            RecordCheck $Matches[1] ($Matches[2] -ceq 'PASS')
        } elseif($line -in @('Admin-generated fixtures','load packages')){Write-Output $line}
        # Other creator diagnostics are suppressed here, never mistaken for a check.
    }
    $creator.WaitForExit()
    $packet=$transfer.GetAwaiter().GetResult()
    if($packet){$state=$packet|ConvertFrom-Json}
    $packet=$null
    if($creator.ExitCode -ne 0 -or $creatorErrors.Result.Length -gt 0 -or $null -eq $state){throw 'Creator did not complete the private fixture transfer.'}
    if(-not $creator.HasExited -or [int]$state.CreatorControllerId -ne $creator.Id -or -not $state.RequireCreatorExit){throw 'Creator exit identity is unverified.'}
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Creator left Excel open.'}
    $state.ParentControllerId=$PID
    $worker=StartController ('-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Slice4beRecordingRestartWorker.ps1')+'"') $true
    $workerErrors=$worker.StandardError.ReadToEndAsync()
    Write-RecordingHandoff $worker.StandardInput.BaseStream ($state|ConvertTo-Json -Depth 40 -Compress)
    $worker.StandardInput.Close()
    $badProtocol=$false
    while($null -ne ($line=$worker.StandardOutput.ReadLine())){
        try{
            $message=$line|ConvertFrom-Json
            if($message.Type -ceq 'Check' -and $message.Name -match '^RecordingRestart\.[A-Za-z]+$' -and $message.Passed -is [bool]){
                RecordCheck $message.Name $message.Passed
            } elseif($message.Type -ceq 'Stage' -and $message.Name -in @('Pristine Viewer','Install drivers','Read journal')){
                Write-Output ('Fresh recording reader: '+$message.Name)
            } else {$badProtocol=$true}
        }catch{$badProtocol=$true}
    }
    $worker.WaitForExit()
    if($badProtocol -or $worker.ExitCode -ne 0 -or $workerErrors.Result.Length -gt 0){throw 'Fresh reader failed; see sanitized stage/check evidence.'}
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Fresh reader left Excel open.'}
} catch {
    $failed=$true
    RecordCheck 'RecordingRestart.CoordinatorHarnessFailure' $false
} finally {
    # Never remove a fixture or restore shared settings beneath a live controller/Excel.
    $cleanupSafe=($null -eq $creator -or $creator.HasExited) -and
        ($null -eq $worker -or $worker.HasExited) -and
        -not (Get-Process EXCEL -ErrorAction SilentlyContinue)
    if(-not $cleanupSafe){
        $failed=$true
        RecordCheck 'RecordingRestart.CleanupDeferredForLiveProcess' $false
    }
    if($cleanupSafe -and $null -ne $state){
        $resolved=[IO.Path]::GetFullPath([string]$state.RunRoot)
        $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
        if($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'invsys-config-command-*'){
            if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
        }
        $results|ConvertTo-Json|Set-Content -LiteralPath (Join-Path ([string]$state.ReportRoot) 'controller-exit-checks.json')
    }
    if($cleanupSafe -and (Test-Path -LiteralPath $settingsRoot)){
        foreach($key in @(Get-Item -LiteralPath $settingsRoot)+@(Get-ChildItem -LiteralPath $settingsRoot -Recurse)){
            foreach($keyName in $key.GetValueNames()){
                if(-not $registryBefore.ContainsKey($key.Name) -or -not $registryBefore[$key.Name].ContainsKey($keyName)){
                    Remove-ItemProperty -LiteralPath ('Registry::'+$key.Name) -Name $keyName
                }
            }
        }
    }
    if($cleanupSafe){foreach($keyPath in $registryBefore.Keys){foreach($keyName in $registryBefore[$keyPath].Keys){
        $saved=$registryBefore[$keyPath][$keyName]
        New-ItemProperty -LiteralPath ('Registry::'+$keyPath) -Name $keyName -Value $saved[0] -PropertyType $saved[1] -Force|Out-Null
    }}}
    if($null -ne $stream){$stream.Dispose()};$pipe.Dispose()
    if($null -ne $creator){$creator.Dispose()};if($null -ne $worker){$worker.Dispose()}
    $state=$null
}
$bad=@($results|Where-Object {-not $_.Passed}).Count
Write-Output ('Controller exit: '+($results.Count-$bad)+' passed, '+$bad+' failed')
if($failed -or $bad){exit 1}
