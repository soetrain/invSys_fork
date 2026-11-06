$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$wrapper=Join-Path $PSScriptRoot 'Invoke-InventoryEditCreationDeferred.ps1'
$root=Join-Path $env:TEMP ('invsys-deferred-test-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$fake=Join-Path $root 'fake-controller.ps1'
@'
param([string]$Phase,[string]$Case,[string]$Deploy)
$controllerRoot=Join-Path $env:TEMP ('invsys-deferred-controller-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $controllerRoot|Out-Null
Write-Output ('Inventory recording controller: '+$controllerRoot)
Start-Sleep -Milliseconds 200
if($env:INVSYS_DEFERRED_INVENTORY_SIGNAL){
    [pscustomobject]@{Stage='AfterQuit';Checks=3;Failed=0;QuitReturned=$true;ExcelProcessId=0;UTC=[DateTimeOffset]::UtcNow.ToString('o')}|
        ConvertTo-Json -Compress|Set-Content -LiteralPath $env:INVSYS_DEFERRED_INVENTORY_SIGNAL
}
Start-Sleep -Seconds 4
[pscustomobject]@{ExitCode=0;ExcelClosed=$true;SettingsRestored=$true;PackagesPreserved=$true}|
    ConvertTo-Json|Set-Content -LiteralPath (Join-Path $controllerRoot 'closure.json')
'@|Set-Content -LiteralPath $fake
$launchArgs=@('-NoProfile','-NonInteractive','-File',$wrapper,'-ControllerScript',$fake,'-WorkingDirectory',$root,'-RootDirectory',$root)
function Run-Wrapper([string]$Label){
    $out=Join-Path $root ($Label+'.out.txt');$err=Join-Path $root ($Label+'.err.txt')
    $watch=[Diagnostics.Stopwatch]::StartNew()
    $quotedArgs=@($launchArgs|ForEach-Object {if($_ -match '[\s]'){'"'+$_+'"'}else{$_}})
    $p=Start-Process powershell.exe -ArgumentList $quotedArgs -WindowStyle Hidden -PassThru -RedirectStandardOutput $out -RedirectStandardError $err
    if(-not $p.WaitForExit(15000)){throw 'The wrapper did not return within 15 seconds.'}
    $p.Refresh()
    $code=$p.ExitCode
    $watch.Stop()
    return [pscustomobject]@{ExitCode=$code;Seconds=$watch.Elapsed.TotalSeconds;Output=(Get-Content $out -Raw -ErrorAction SilentlyContinue);Error=(Get-Content $err -Raw -ErrorAction SilentlyContinue)}
}
$completed=$false
try {
    $first=Run-Wrapper 'first'
    $early=$first.Output|ConvertFrom-Json
    if($early.State -cne 'PROVISIONAL' -or $early.FullAcceptance -or $early.Checks -ne 3 -or $early.Failed -ne 0 -or (Test-Path -LiteralPath (Join-Path $early.RunRoot 'final.json'))){throw 'RED: launcher did not return a provisional result before final cleanup.'}
    $second=Run-Wrapper 'second'
    if($second.Output -match 'PROVISIONAL' -or $second.Error -notmatch 'cleanup is pending'){throw 'RED: a second Excel test was not refused while cleanup is pending.'}
    $deadline=[DateTimeOffset]::UtcNow.AddSeconds(10)
    do {
        $final=@(Get-ChildItem -LiteralPath $root -Recurse -Filter final.json -File -ErrorAction SilentlyContinue)
        if($final.Count -gt 0){break}
        Start-Sleep -Milliseconds 200
    }while([DateTimeOffset]::UtcNow -lt $deadline)
    if($final.Count -ne 1 -or -not ((Get-Content $final[0].FullName -Raw|ConvertFrom-Json).CleanupComplete)){throw ('GREEN: final cleanup qualification missing. Root='+$root)}
    $completed=$true
    Write-Output 'Deferred cleanup: 3/3 PASS (early result, pending gate, final closure)'
} finally {
    # The fake controller is test-owned and must finish before removing its files.
    Start-Sleep -Seconds 1
    if($completed){Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue}
}
