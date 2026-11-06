# Development-only early result for one Inventory edit-creation controller.
# The detached supervisor retains natural Excel closure, settings restoration,
# and final qualification after the foreground command returns.
# From the main code checkout: powershell.exe -NoProfile -File
# ../invSys_edit_receipts/tests/tooling/Invoke-InventoryEditCreationDeferred.ps1
# -Phase GREEN -Case ApplyFailure -Deploy deploy/validation-inventory-config-03
# PROVISIONAL is not release GREEN. Inspect RunRoot/final.json before another
# Excel test; this launcher refuses a second run while cleanup is pending.
# The existing controller requires a fresh desktop-monitor report before dispatch.
param(
    [string]$ControllerScript = '',
    [string]$WorkingDirectory = '',
    [string]$RootDirectory = '',
    [ValidateSet('RED','GREEN')][string]$Phase = 'GREEN',
    [string]$Case = 'ApplyFailure',
    [string]$Deploy = 'deploy/validation-inventory-config-03',
    [switch]$Supervisor,
    [string]$RunRoot = ''
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest

function Write-AtomicJson([string]$Path,[object]$Value){
    $temporary=$Path+'.tmp'
    $Value|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $temporary -Encoding UTF8
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

if($Supervisor){
    $dispatch=Get-Content -LiteralPath (Join-Path $RunRoot 'dispatch.json') -Raw|ConvertFrom-Json
    $env:INVSYS_DEFERRED_INVENTORY_SIGNAL=Join-Path $RunRoot 'after-quit.json'
    $controller=$null
    try {
        $quotedPath=(Join-Path $RunRoot 'dispatch.json').Replace("'","''")
        $command="`$d=Get-Content -LiteralPath '$quotedPath' -Raw|ConvertFrom-Json; & `$d.ControllerScript -Phase `$d.Phase -Case `$d.Case -Deploy `$d.Deploy; if(-not `$?){exit 1}"
        $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
        $controller=Start-Process powershell.exe -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) -WorkingDirectory $dispatch.WorkingDirectory -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput (Join-Path $RunRoot 'controller.stdout.log') -RedirectStandardError (Join-Path $RunRoot 'controller.stderr.log')
        Write-AtomicJson (Join-Path $RunRoot 'controller.json') ([pscustomobject]@{Id=$controller.Id;ExitCode=$controller.ExitCode})
        $line=@(Get-Content -LiteralPath (Join-Path $RunRoot 'controller.stdout.log') -ErrorAction SilentlyContinue | Where-Object {$_ -clike 'Inventory recording controller: *'} | Select-Object -First 1)
        $closure=$null
        if($line.Count -eq 1){
            $resultRoot=$line[0].Substring('Inventory recording controller: '.Length)
            $closurePath=Join-Path $resultRoot 'closure.json'
            if(Test-Path -LiteralPath $closurePath){$closure=Get-Content -LiteralPath $closurePath -Raw|ConvertFrom-Json}
        }
        $excelClosed=(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0)
        $cleanupComplete=($null -ne $closure -and $closure.ExcelClosed -eq $true -and $closure.SettingsRestored -eq $true -and $closure.PackagesPreserved -eq $true -and $excelClosed)
        $testPassed=($cleanupComplete -and $controller.ExitCode -eq 0 -and $closure.ExitCode -eq 0)
        $state=if(-not $cleanupComplete){'CLEANUP_FAILED'}elseif($testPassed){'TEST_PASSED_CLEANUP_COMPLETE'}else{'TEST_FAILED_CLEANUP_COMPLETE'}
        Write-AtomicJson (Join-Path $RunRoot 'final.json') ([pscustomobject]@{State=$state;CleanupComplete=$cleanupComplete;TestPassed=$testPassed;FullAcceptance=$false;ControllerExitCode=$controller.ExitCode;ExcelClosed=$excelClosed;SettingsRestored=($null -ne $closure -and $closure.SettingsRestored -eq $true);PackagesPreserved=($null -ne $closure -and $closure.PackagesPreserved -eq $true);UTC=[DateTimeOffset]::UtcNow.ToString('o')})
    } catch {
        Write-AtomicJson (Join-Path $RunRoot 'final.json') ([pscustomobject]@{State='FAILED';CleanupComplete=$false;FullAcceptance=$false;FailureType=$_.Exception.GetType().FullName;ControllerId=if($controller){$controller.Id}else{$null};UTC=[DateTimeOffset]::UtcNow.ToString('o')})
    }
    exit 0
}

$sourceRoot=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$codeRoot=if((Split-Path $sourceRoot -Leaf) -ceq 'invSys_fork'){$sourceRoot}else{Join-Path (Split-Path $sourceRoot -Parent) 'invSys_fork'}
if(-not $WorkingDirectory){$WorkingDirectory=$codeRoot}
if(-not $ControllerScript){$ControllerScript=Join-Path $PSScriptRoot 'Run-InventoryEditCreationController.ps1'}
if(-not $RootDirectory){$RootDirectory=Join-Path $codeRoot 'reports/runtime/deferred-inventory-edit-creation'}
if(-not(Test-Path -LiteralPath $ControllerScript -PathType Leaf)){throw 'Inventory controller script is unavailable.'}
if(-not(Test-Path -LiteralPath $WorkingDirectory -PathType Container)){throw 'Controller working directory is unavailable.'}
New-Item -ItemType Directory -Path $RootDirectory -Force|Out-Null
foreach($prior in @(Get-ChildItem -LiteralPath $RootDirectory -Directory -Filter 'run-*')){
    $finalPath=Join-Path $prior.FullName 'final.json'
    if(-not(Test-Path -LiteralPath $finalPath)){throw ('Previous Inventory cleanup is pending: '+$prior.FullName)}
    $priorFinal=Get-Content -LiteralPath $finalPath -Raw|ConvertFrom-Json
    $priorCleanupSafe=($priorFinal.CleanupComplete -eq $true -or ($priorFinal.ExcelClosed -eq $true -and $priorFinal.SettingsRestored -eq $true -and $priorFinal.PackagesPreserved -eq $true))
    if(-not $priorCleanupSafe){throw ('Previous Inventory cleanup failed; inspect before another test: '+$prior.FullName)}
}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed before an Inventory test.'}
$RunRoot=Join-Path $RootDirectory ('run-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $RunRoot|Out-Null
Write-AtomicJson (Join-Path $RunRoot 'dispatch.json') ([pscustomobject]@{ControllerScript=(Resolve-Path -LiteralPath $ControllerScript).Path;WorkingDirectory=(Resolve-Path -LiteralPath $WorkingDirectory).Path;Phase=$Phase;Case=$Case;Deploy=$Deploy;UTC=[DateTimeOffset]::UtcNow.ToString('o')})
$supervisorProcess=Start-Process powershell.exe -ArgumentList @('-NoProfile','-NonInteractive','-File',('"'+$PSCommandPath+'"'),'-Supervisor','-RunRoot',('"'+$RunRoot+'"')) -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $RunRoot 'supervisor.stdout.log') -RedirectStandardError (Join-Path $RunRoot 'supervisor.stderr.log')
$signal=Join-Path $RunRoot 'after-quit.json'
$deadline=[DateTimeOffset]::UtcNow.AddMinutes(30)
while(-not(Test-Path -LiteralPath $signal) -and -not $supervisorProcess.HasExited -and [DateTimeOffset]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
if(-not(Test-Path -LiteralPath $signal)){throw ('No early Quit result; controller remains under supervisor. Inspect '+$RunRoot)}
$early=Get-Content -LiteralPath $signal -Raw|ConvertFrom-Json
if($early.Stage -cne 'AfterQuit' -or $early.QuitReturned -ne $true){throw ('Quit not confirmed. Inspect '+$RunRoot)}
[pscustomobject]@{State='PROVISIONAL';FullAcceptance=$false;Checks=$early.Checks;Failed=$early.Failed;QuitReturned=$true;ExcelProcessId=$early.ExcelProcessId;RunRoot=$RunRoot;Message='Controller still owns natural Excel closure and settings restoration.'}|ConvertTo-Json -Compress|Write-Output
if($early.Failed -gt 0){exit 1}
