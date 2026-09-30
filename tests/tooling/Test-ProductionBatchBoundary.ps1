# Calibrated batch-boundary diagnostic; optionally retains the standard run flow.
[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-production-paths',
      [string]$PackagePinsPath='reports/runtime/production-lifecycle-native-controller/65e8e27585ca47289cf27e87e147e83e/package-pins.json',
      [ValidateRange(0,30000)][int]$ReleaseObservationMilliseconds=30000,
      [switch]$FullProductionFlow,
      [switch]$TraceBoundaries,[switch]$StandardRunFlow,[switch]$CompileOnly,[switch]$NativeExceptions,[switch]$NativeFaultsOnly,[switch]$NativeBeforeRun,[switch]$ReleaseAutomationForTest,[switch]$ClearErrorReferencesForTest,[switch]$CloseOperatorFirstForTest,[switch]$ObserveShutdownForTest,[switch]$CloseOwnedWorkbooksForTest,[switch]$ReverseOpenedCloseForTest,[switch]$GenerateOnly)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if($StandardRunFlow -and -not ($TraceBoundaries -or $CompileOnly -or $NativeExceptions -or $ReleaseAutomationForTest -or $CloseOperatorFirstForTest -or $ObserveShutdownForTest -or $CloseOwnedWorkbooksForTest -or $ReverseOpenedCloseForTest)){throw 'Standard-flow diagnosis requires an explicit observer/control.'}
if($CloseOperatorFirstForTest -and -not $StandardRunFlow){throw 'Operator-first cleanup requires the completed standard flow.'}
if($CompileOnly -and ($TraceBoundaries -or -not $StandardRunFlow)){throw 'VBE preparation control requires standard flow without tracing.'}
if($NativeExceptions -and ($TraceBoundaries -or $CompileOnly -or -not $StandardRunFlow)){throw 'Native observation requires standard flow without VBE preparation.'}
if($NativeBeforeRun -and -not $NativeExceptions){throw 'Late attach requires the native observer.'}
if($FullProductionFlow -and (-not $StandardRunFlow -or -not $NativeExceptions -or $NativeBeforeRun)){throw 'Full Production diagnosis requires early native observation of the standard flow.'}
if($NativeFaultsOnly -and -not $NativeExceptions){throw 'Native filtering requires the native observer.'}
if($ClearErrorReferencesForTest -and -not $ReleaseAutomationForTest){throw 'Post-report error-reference control requires explicit automation cleanup diagnosis.'}
if($ReleaseObservationMilliseconds -ne 30000 -and -not $ReleaseAutomationForTest){throw 'Release wait control requires automation release diagnosis.'}
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated diagnostic.'}
$deploy=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$rawPins=Get-Content -LiteralPath (Join-Path $repo $PackagePinsPath) -Raw|ConvertFrom-Json
$pins=@($rawPins)
if($pins.Count -ne 5){throw 'Five package pins required.'}
foreach($pin in $pins){
    $name=if($pin.PSObject.Properties['Package']){$pin.Package}else{Split-Path -Leaf $pin.File}
    if((Get-FileHash -LiteralPath (Join-Path $deploy $name)).Hash -cne $pin.Hash){throw 'Candidate package differs.'}
}
$root=Join-Path $repo ('reports/runtime/production-batch-boundary/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
Write-Output ('Diagnostic: '+$root)
$validator=Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'
$validatorHash=(Get-FileHash -LiteralPath $validator).Hash
$source=[IO.File]::ReadAllText($validator)
$originalSource=$source
function Replace-Once([string]$Text,[string]$Anchor,[string]$Replacement){
    $pattern='(?m)^'+[regex]::Escape($Anchor)+'\r?$'
    if([regex]::Matches($Text,$pattern).Count -ne 1){throw 'Diagnostic harness anchor missing/ambiguous.'}
    [regex]::Replace($Text,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($match) $Replacement})
}
if($NativeExceptions){
    $nativeAnchor='    $currentStep = "create isolated config and auth"'
    if($NativeBeforeRun){$nativeAnchor='                    [IO.File]::WriteAllText($progressPath, "invoke Production reusable run actions")'}
    $nativeInstall=@'
    $nativeBinary=Join-Path $outputPath 'NativeExceptionObserver.exe'
    $nativeArgs=@([string]$excelProcessId,('"'+$outputPath+'"'),'600')
    $nativeObserver=Start-Process -FilePath $nativeBinary -ArgumentList $nativeArgs -WindowStyle Hidden -PassThru
    [IO.File]::WriteAllText((Join-Path $outputPath 'observer-pid.txt'),[string]$nativeObserver.Id)
    $nativeDeadline=[DateTime]::UtcNow.AddSeconds(20)
    while(-not (Test-Path (Join-Path $outputPath 'ready'))){
        if($nativeObserver.HasExited -or [DateTime]::UtcNow -gt $nativeDeadline){throw 'Native observer did not become ready; no workflow RED.'}
        Start-Sleep -Milliseconds 100
    }
'@
    if($NativeFaultsOnly){$nativeInstall=$nativeInstall.Replace("    `$nativeObserver=Start-Process", "    `$nativeArgs+='faults-only'`r`n    `$nativeObserver=Start-Process")}
    $source=Replace-Once $source $nativeAnchor ($nativeInstall+"`r`n"+$nativeAnchor)
}
if($TraceBoundaries -or $CompileOnly){
    $anchor='    $coreName = [string]$packages["invSys.Core.xlam"].Name'
    $install=@'
    . (Join-Path $repo 'tests/tooling/ProductionBatchBoundaryTrace.ps1')
    Install-ProductionBatchTrace -Excel $excel -Packages $packages -PackageRoot $deployPath
'@
    if($CompileOnly){$install+=' -CompileOnly'}
    $source=Replace-Once $source $anchor ($install+"`r`n"+$anchor)
}
if($TraceBoundaries){
    $anchor='    $currentStep = "invoke packaged callbacks for state $WorkbookState"'
    $arm=@'
    $tracePath=Join-Path $outputPath 'stages.txt'
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName -MacroName 'TestProductionBatchTrace.Arm' -Arguments @($tracePath))
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName -MacroName 'TestProductionBatchTrace.Mark' -Arguments @('UNREGISTERED_CALIBRATION_VALUE'))
    if(-not (Test-Path -LiteralPath $tracePath) -or [IO.File]::ReadAllText($tracePath).Trim() -cne 'Arm'){throw 'Trace transport/redaction calibration failed; not product RED.'}
'@
    $source=Replace-Once $source $anchor ($arm+"`r`n"+$anchor)
}
if($CloseOwnedWorkbooksForTest){
    $ownedCloseInstall=@'
    $ownedCloseFailure=$null;$ownedCloseCount=0;$ownedCloseBefore=-1;$ownedCloseAfter=-1
    $ownedCloseBooks=@()
    try {
        $ownedCloseRuntime=[IO.Path]::GetFullPath($runtimeRoot).TrimEnd('\')+'\'
        $ownedClosePackages=[IO.Path]::GetFullPath($deployPath).TrimEnd('\')+'\'
        if(-not $ownedCloseRuntime.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase) -or
            (Split-Path -Leaf $runtimeRoot) -notlike 'invsys-plan022-launcher-red-*'){throw 'Disposable runtime ownership failed.'}
        $ownedCloseBooks=@($excel.Workbooks)
        $ownedCloseBefore=$ownedCloseBooks.Count
        foreach($ownedCloseBook in $ownedCloseBooks){
            $ownedClosePath=[IO.Path]::GetFullPath([string]$ownedCloseBook.FullName)
            $ownedCloseFixture=$ownedClosePath.StartsWith($ownedCloseRuntime,[StringComparison]::OrdinalIgnoreCase)
            $ownedClosePackage=[bool]$ownedCloseBook.IsAddin -and $ownedClosePath.StartsWith($ownedClosePackages,[StringComparison]::OrdinalIgnoreCase)
            if(-not ($ownedCloseFixture -or $ownedClosePackage)){throw 'Workbook outside isolated roots.'}
        }
        foreach($ownedCloseBook in @($ownedCloseBooks|Sort-Object {[bool]$_.IsAddin})){
            $ownedCloseBook.Close($false);$ownedCloseCount++
        }
        $ownedCloseAfter=[int]$excel.Workbooks.Count
    }catch{
        $ownedCloseFailure=[pscustomobject]@{ExceptionType=$_.Exception.GetType().Name;HResult=$_.Exception.GetBaseException().HResult;Line=$_.InvocationInfo.ScriptLineNumber}
    }finally{
        foreach($ownedCloseBook in $ownedCloseBooks){Release-ComObject $ownedCloseBook}
        $ownedCloseBook=$null;$ownedCloseBooks=@()
    }
    [pscustomobject]@{WorkbooksBefore=$ownedCloseBefore;CloseReturned=$ownedCloseCount;WorkbooksAfter=$ownedCloseAfter;Failure=$ownedCloseFailure}|
        ConvertTo-Json|Set-Content (Join-Path $outputPath 'owned-workbook-cleanup.json')
'@
    $source=Replace-Once $source '    foreach ($wb in $opened) {' ($ownedCloseInstall+"`r`n"+'    foreach ($wb in $opened) {')
}
if($ObserveShutdownForTest){
    $shutdownInit=@'
    $shutdownState=[ordered]@{WorkbooksBeforeClose=-1;WorkbooksBeforeQuit=-1;EnableEventsBeforeQuit=$null;CloseReturned=0;CloseFailures=@();QuitReturned=$false;QuitFailure=$null;ReadFailures=@()}
    try{$shutdownState.WorkbooksBeforeClose=[int]$excel.Workbooks.Count}catch{$shutdownState.ReadFailures+= $_.Exception.GetBaseException().HResult}
'@
    $shutdownBeforeQuit=@'
    try{$shutdownState.WorkbooksBeforeQuit=[int]$excel.Workbooks.Count;$shutdownState.EnableEventsBeforeQuit=[bool]$excel.EnableEvents}catch{$shutdownState.ReadFailures+= $_.Exception.GetBaseException().HResult}
'@
    $shutdownWrite="    `$shutdownState|ConvertTo-Json|Set-Content (Join-Path `$outputPath 'shutdown-observation.json')"
    $shutdownCloseOriginal='        try { $wb.Close($false) } catch {}'
    $shutdownCloseObserved='        $shutdownCloseKind="FixtureOrUnavailable";try{$shutdownCloseName=[string]$wb.Name;if($shutdownCloseName -cin $packageNames){$shutdownCloseKind=$shutdownCloseName}}catch{};try { $wb.Close($false);$shutdownState.CloseReturned++ } catch { $shutdownState.CloseFailures+= [pscustomobject]@{Package=$shutdownCloseKind;HResult=$_.Exception.GetBaseException().HResult} }'
    $shutdownQuitOriginal='        try { $excel.Quit() } catch {}'
    $shutdownQuitObserved='        try { $excel.Quit();$shutdownState.QuitReturned=$true } catch { $shutdownState.QuitFailure=$_.Exception.GetBaseException().HResult }'
    $source=Replace-Once $source '    $paletteBooksClosed=$false' ($shutdownInit+"`r`n"+'    $paletteBooksClosed=$false')
    $source=Replace-Once $source $shutdownCloseOriginal $shutdownCloseObserved
    $shutdownQuitPattern='(?m)(^    if \(\$null -ne \$excel\) \{\r?\n)'+[regex]::Escape($shutdownQuitOriginal)+'\r?$'
    if([regex]::Matches($source,$shutdownQuitPattern).Count -ne 1){throw 'Final Quit observation anchor missing/ambiguous.'}
    $source=[regex]::Replace($source,$shutdownQuitPattern,[Text.RegularExpressions.MatchEvaluator]{param($match) $match.Groups[1].Value+$shutdownQuitObserved})
    $source=Replace-Once $source '    if ($null -ne $excel) {' ($shutdownBeforeQuit+"`r`n"+'    if ($null -ne $excel) {')
    $source=Replace-Once $source '    $cleanupTerminationRequested=$false' ($shutdownWrite+"`r`n"+'    $cleanupTerminationRequested=$false')
}
if($CloseOperatorFirstForTest){
    $operatorCloseAnchor='    foreach ($wb in $opened) {'
    $operatorCloseInstall=@'
    $operatorCloseFailure=$null;$operatorCloseMatches=0;$operatorClosed=$false;$operatorCountReduced=$false
    $operatorEventsEnabled=$false
    try {
        $ownedRuntime=[IO.Path]::GetFullPath($runtimeRoot).TrimEnd('\')+'\'
        $ownedOperators=[IO.Path]::GetFullPath($operatorRoot).TrimEnd('\')+'\'
        $ownedOperator=[IO.Path]::GetFullPath($productionOperatorPath)
        if(-not $ownedRuntime.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase) -or
            (Split-Path -Leaf $runtimeRoot) -notlike 'invsys-plan022-launcher-red-*' -or
            -not $ownedOperators.StartsWith($ownedRuntime,[StringComparison]::OrdinalIgnoreCase) -or
            -not $ownedOperator.StartsWith($ownedOperators,[StringComparison]::OrdinalIgnoreCase)){
            throw 'Disposable operator ownership failed.'
        }
        $operatorEventsEnabled=[bool]$excel.EnableEvents
        $operatorCountBefore=[int]$excel.Workbooks.Count
        for($operatorIndex=[int]$excel.Workbooks.Count;$operatorIndex -ge 1;$operatorIndex--){
            $operatorCloseBook=$excel.Workbooks.Item($operatorIndex)
            try {
                if([IO.Path]::GetFullPath([string]$operatorCloseBook.FullName).Equals($ownedOperator,[StringComparison]::OrdinalIgnoreCase)){
                    $operatorCloseMatches++
                    $operatorCloseBook.Close($false)
                    $operatorClosed=$true
                }
            }finally{Release-ComObject $operatorCloseBook;$operatorCloseBook=$null}
        }
        $operatorCountReduced=([int]$excel.Workbooks.Count -eq ($operatorCountBefore-1))
    }catch{
        $operatorCloseFailure=[pscustomobject]@{ExceptionType=$_.Exception.GetType().Name;HResult=$_.Exception.HResult;Line=$_.InvocationInfo.ScriptLineNumber}
    }
    [pscustomobject]@{MatchingOwnedWorkbooks=$operatorCloseMatches;EnableEvents=$operatorEventsEnabled;CloseReturned=$operatorClosed;WorkbookCountReducedByOne=$operatorCountReduced;Failure=$operatorCloseFailure}|
        ConvertTo-Json|Set-Content (Join-Path $outputPath 'operator-first-cleanup.json')
'@
    $source=Replace-Once $source $operatorCloseAnchor ($operatorCloseInstall+"`r`n"+$operatorCloseAnchor)
}
if($ReleaseAutomationForTest){
    $releaseAnchor='    $cleanupTerminationRequested=$false'
    $releaseInstall=@'
    . (Join-Path $repo 'tests/tooling/IsolatedAutomationCleanup.ps1')
    $releasedReferences=$null;$releaseFailure=$null
    $clearedErrorRecords=0
    $normalExit=$false
    $releaseWaitElapsed=0;$releaseWaitStartUTC=$null;$releaseWaitEndUTC=$null
    try {
        $releasedReferences=Release-IsolatedAutomationVariables -Variables (Get-Variable -Scope Script)
        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
        if($excelProcessId -gt 0){
            $releaseOwner=Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
            $normalExit=($null -eq $releaseOwner)
            if($null -ne $releaseOwner){
                $releaseWaitStartUTC=[DateTimeOffset]::UtcNow.ToString('o')
                $releaseWaitClock=[Diagnostics.Stopwatch]::StartNew()
                $normalExit=$releaseOwner.WaitForExit(30000)
                $releaseWaitClock.Stop();$releaseWaitElapsed=$releaseWaitClock.ElapsedMilliseconds
                $releaseWaitEndUTC=[DateTimeOffset]::UtcNow.ToString('o')
            }
        }
    } catch {
        $failureId=[string]$_.FullyQualifiedErrorId
        if($failureId -cnotmatch '^[A-Za-z0-9.,_]+$'){$failureId='Unclassified'}
        $releaseFailure=[pscustomobject]@{ErrorId=$failureId;ExceptionType=$_.Exception.GetType().Name;HResult=$_.Exception.HResult;Line=$_.InvocationInfo.ScriptLineNumber}
    }
    [pscustomobject]@{References=$releasedReferences;Failure=$releaseFailure;ExitedAfterRelease=$normalExit;ObservationSeconds=30;WaitElapsedMilliseconds=$releaseWaitElapsed;WaitStartUTC=$releaseWaitStartUTC;WaitEndUTC=$releaseWaitEndUTC;ClearedErrorRecords=$clearedErrorRecords}|ConvertTo-Json|Set-Content (Join-Path $outputPath 'automation-release.json')
'@
    if($ClearErrorReferencesForTest){
        $releaseInstall=$releaseInstall.Replace('        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()',
            '        $clearedErrorRecords=$Error.Count;$Error.Clear()'+"`r`n"+'        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()')
    }
    $releaseInstall=$releaseInstall.Replace('WaitForExit(30000)',('WaitForExit('+$ReleaseObservationMilliseconds+')')).Replace('ObservationSeconds=30;',('ObservationSeconds='+([string]($ReleaseObservationMilliseconds/1000))+';'))
    $source=Replace-Once $source $releaseAnchor ($releaseInstall+"`r`n"+$releaseAnchor)
}
if($ReverseOpenedCloseForTest){
    $reverseCloseOriginal='    foreach ($wb in $opened) {'
    $reverseCloseReplacement=@'
    $reverseCloseBooks=$opened.ToArray()
    [array]::Reverse($reverseCloseBooks)
    foreach ($wb in $reverseCloseBooks) {
'@
    $source=Replace-Once $source $reverseCloseOriginal $reverseCloseReplacement
}
$anchor='                $observedText += " || PRODUCTION_BATCH_SCALE=" + $workflowControlReport'
$cutAnchor=$anchor
$cut=@'
                $localPaths=@($newWorkbookPaths|Where-Object {[IO.Path]::GetFullPath($_).StartsWith([IO.Path]::GetFullPath($operatorRoot),[StringComparison]::OrdinalIgnoreCase)})
                [pscustomobject]@{NewWorkbookCount=$newWorkbooks.Count;NewPathCount=$newWorkbookPaths.Count;StationLocalCount=$localPaths.Count;SecondNewWorkbookCount=$secondNewWorkbooks.Count}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $outputPath 'workbook-counts.json')
                $focused=[ordered]@{
                    Setup=(@($setupEvidence|Where-Object {$_ -notmatch '=True$'}).Count -eq 0)
                    FirstLauncher=([string]::IsNullOrWhiteSpace($capture.Error))
                    FirstSavedWorkbook=($localPaths.Count -eq 1 -and (Test-Path -LiteralPath $localPaths[0]))
                    SecondLauncher=([string]::IsNullOrWhiteSpace($secondCapture.Error))
                    SecondReusesWorkbook=($secondNewWorkbooks.Count -eq 0)
                    OriginalScaleContract=$workflowControlPassed
                    ExactScaleResult=($workflowControlReport -ceq 'OK|Min=.001%|Default=100%|Max=1000%|BoundsRejected=True|ListOnly=True')
                }
                foreach($name in $focused.Keys){Add-Evidence -Rows $evidence -Callback ('BatchBoundary.'+$name) -Expected 'Existing packaged launcher/scale contract holds.' -Passed ([bool]$focused[$name]) -Observed 'Details omitted'}
                if(@($focused.Values|Where-Object {-not $_}).Count){throw 'Focused launcher/scale contract failed.'}
                [IO.File]::WriteAllText((Join-Path $outputPath 'cut-reached.txt'),'AfterBatchScale')
                return
'@
if(-not $StandardRunFlow){$source=Replace-Once $source $anchor $cut}
$anchor='        $observed = ([string]$row.Observed).Replace("|", "/")'
$redactionAnchor=$anchor
$source=Replace-Once $source $anchor '        $observed = "Details omitted"'
$heading='# Production batch-boundary diagnostic; not full reusable acceptance'
$source=$source.Replace('# Plan 022 Slice 4x Packaged Reusable Production Evidence',$heading)
# Reverse only the declared diagnostic insertions/redaction. Every original
# statement, including all standard workflow assertions and cleanup, must remain.
$restoredSource=$source.Replace('        $observed = "Details omitted"',$redactionAnchor).Replace($heading,'# Plan 022 Slice 4x Packaged Reusable Production Evidence')
if(-not $StandardRunFlow){$restoredSource=$restoredSource.Replace($cut,$cutAnchor)}
if($TraceBoundaries -or $CompileOnly){$restoredSource=$restoredSource.Replace($install+"`r`n",'')}
if($TraceBoundaries){$restoredSource=$restoredSource.Replace($arm+"`r`n",'')}
if($NativeExceptions){$restoredSource=$restoredSource.Replace($nativeInstall+"`r`n",'')}
if($ReleaseAutomationForTest){$restoredSource=$restoredSource.Replace($releaseInstall+"`r`n",'')}
if($CloseOperatorFirstForTest){$restoredSource=$restoredSource.Replace($operatorCloseInstall+"`r`n",'')}
if($CloseOwnedWorkbooksForTest){$restoredSource=$restoredSource.Replace($ownedCloseInstall+"`r`n",'')}
if($ReverseOpenedCloseForTest){$restoredSource=$restoredSource.Replace($reverseCloseReplacement,$reverseCloseOriginal)}
if($ObserveShutdownForTest){
    $restoredSource=$restoredSource.Replace($shutdownInit+"`r`n",'').Replace($shutdownBeforeQuit+"`r`n",'').Replace($shutdownWrite+"`r`n",'').Replace($shutdownCloseObserved,$shutdownCloseOriginal).Replace($shutdownQuitObserved,$shutdownQuitOriginal)
}
$statementsPreserved=($restoredSource -replace "`r`n","`n") -ceq ($originalSource -replace "`r`n","`n")
if(-not $statementsPreserved){throw 'Diagnostic changed undeclared standard-validator statements.'}
$generated=Join-Path $root $(if($StandardRunFlow){'standard-run-validator.ps1'}else{'scoped-validator.ps1'})
[IO.File]::WriteAllText($generated,$source,[Text.UTF8Encoding]::new($false))
$tokens=$null;$errors=$null
[void][Management.Automation.Language.Parser]::ParseFile($generated,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Diagnostic does not parse.'}
$flowArguments=@();if($StandardRunFlow -and -not $FullProductionFlow){$flowArguments+='-ProductionRunOnly'}
$expectedChecks=if($FullProductionFlow){2}elseif($StandardRunFlow){1}else{7}
[pscustomobject]@{StandardRunFlow=[bool]$StandardRunFlow;FullProductionFlow=[bool]$FullProductionFlow;ExpectedChecks=$expectedChecks;TraceBoundaries=[bool]$TraceBoundaries;CompileOnly=[bool]$CompileOnly;NativeExceptions=[bool]$NativeExceptions;NativeFaultsOnly=[bool]$NativeFaultsOnly;NativeBeforeRun=[bool]$NativeBeforeRun;ReleaseAutomationForTest=[bool]$ReleaseAutomationForTest;ReleaseObservationMilliseconds=$ReleaseObservationMilliseconds;ClearErrorReferencesForTest=[bool]$ClearErrorReferencesForTest;CloseOperatorFirstForTest=[bool]$CloseOperatorFirstForTest;ObserveShutdownForTest=[bool]$ObserveShutdownForTest;CloseOwnedWorkbooksForTest=[bool]$CloseOwnedWorkbooksForTest;ReverseOpenedCloseForTest=[bool]$ReverseOpenedCloseForTest;OriginalStatementsPreserved=$statementsPreserved;ParseErrors=$errors.Count;GenerationOpenedExcel=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'generation.json')
if($GenerateOnly){Write-Output 'Diagnostic generation calibrated; no runtime invoked.';return}
if($NativeExceptions){
    @('NativeExceptionObserver.cs','Test-ProductionBatchBoundary.ps1')|ForEach-Object {
        [pscustomobject]@{File=$_;Hash=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot $_)).Hash}
    }|ConvertTo-Json|Set-Content (Join-Path $root 'observer-source-pins.json')
    $compiler=New-Object CodeDom.Compiler.CompilerParameters -Property @{CompilerOptions='/platform:x64';GenerateExecutable=$true;OutputAssembly=(Join-Path $root 'NativeExceptionObserver.exe')}
    [void]$compiler.ReferencedAssemblies.Add('System.dll')
    Add-Type -Path (Join-Path $PSScriptRoot 'NativeExceptionObserver.cs') -CompilerParameters $compiler
}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$restored=$false;$code=1;$nativeStopped=$true;$start=[DateTimeOffset]::UtcNow
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $generated -RepoRoot $repo -DeployRoot $DeployRoot -OutputDirectory ($root.Substring($repo.Length+1)) -CallbackFilter Production -WorkbookState ProductionReusable @flowArguments *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
} finally {
    if($NativeExceptions){
        [IO.File]::WriteAllText((Join-Path $root 'stop'),'Stop')
        $observerPidPath=Join-Path $root 'observer-pid.txt'
        if(Test-Path $observerPidPath){
            $observerProcess=Get-Process -Id ([int][IO.File]::ReadAllText($observerPidPath)) -ErrorAction SilentlyContinue
            if($null -ne $observerProcess){$nativeStopped=$observerProcess.WaitForExit(5000)}
        }
    }
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $preserved=$true
    foreach($pin in $pins){
        $name=if($pin.PSObject.Properties['Package']){$pin.Package}else{Split-Path -Leaf $pin.File}
        $preserved=$preserved -and (Get-FileHash -LiteralPath (Join-Path $deploy $name)).Hash -ceq $pin.Hash
    }
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$preserved;ExcelClosed=$true;ValidatorPreserved=((Get-FileHash -LiteralPath $validator).Hash -ceq $validatorHash);TraceBoundaries=[bool]$TraceBoundaries;FullProductionAccepted=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'closure.json')
    if(-not $restored -or -not $preserved){throw 'Diagnostic preservation failed.'}
    if(-not $nativeStopped){throw 'Native observer did not stop; preserve diagnostic before continuing.'}
}
Start-Sleep -Seconds 12
$queryErrors=@();$audit=[DateTimeOffset]::UtcNow
$events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=$audit.LocalDateTime} -ErrorAction SilentlyContinue -ErrorVariable queryErrors|Where-Object {$_.Message -match '(?i)excel\.exe'})
if(@($queryErrors|Where-Object FullyQualifiedErrorId -NotLike 'NoMatchingEventsFound*').Count){throw 'Application audit unavailable.'}
$checks=@();$report=Join-Path $root 'production-reusable-production.md'
if(Test-Path -LiteralPath $report){$checks=@([regex]::Matches([IO.File]::ReadAllText($report),'(?m)^\| ([^|]+) \| (PASS|RED) \|')|ForEach-Object {[pscustomobject]@{Check=$_.Groups[1].Value;Passed=$_.Groups[2].Value -ceq 'PASS'}})}
$cutReached=Test-Path -LiteralPath (Join-Path $root 'cut-reached.txt')
$cleanup=$null;$cleanupPath=Join-Path $root 'final-cleanup-observation.json'
if(Test-Path -LiteralPath $cleanupPath){$cleanup=Get-Content -LiteralPath $cleanupPath -Raw|ConvertFrom-Json}
$traceValid=$true;$stages=@()
if($TraceBoundaries){
    . (Join-Path $PSScriptRoot 'ProductionBatchBoundaryTrace.ps1')
    $allow=@('Arm')+@(Get-ProductionBatchTracePlan|ForEach-Object Stage)
    $stagePath=Join-Path $root 'stages.txt'
    if(Test-Path -LiteralPath $stagePath){$stages=@(Get-Content -LiteralPath $stagePath)}
    $traceValid=$stages.Count -gt 1 -and @($stages|Where-Object {$_ -cnotin $allow}).Count -eq 0
}
$failed=@($checks|Where-Object {-not $_.Passed}).Count
$scopeSatisfied=($checks.Count -eq $expectedChecks -and ($cutReached -eq (-not $StandardRunFlow)))
$nativeValid=$true
if($NativeExceptions){
    $nativeEvents=@(Get-Content (Join-Path $root 'native-events.jsonl')|ForEach-Object {$_|ConvertFrom-Json})
    $nativeStates=@($nativeEvents|Where-Object {$_.PSObject.Properties['State']}|ForEach-Object State)
    $nativeValid=('Ready' -cin $nativeStates -and ('Exited' -cin $nativeStates -or 'Detached' -cin $nativeStates) -and @($nativeEvents|Where-Object {$_.PSObject.Properties['Win32Error'] -and $_.Win32Error -ne 0}).Count -eq 0)
}
$releaseValid=$true
if($ReleaseAutomationForTest){
    $released=Get-Content (Join-Path $root 'automation-release.json') -Raw|ConvertFrom-Json
    $releaseValid=(($released.ExitedAfterRelease -or $ReleaseObservationMilliseconds -eq 0) -and $null -eq $released.Failure -and $null -ne $released.References -and $released.References.ReleaseFailures -eq 0 -and $null -ne $cleanup -and $cleanup.ProcessIdAvailable -and -not $cleanup.TerminationRequested)
}
$operatorCloseValid=$true
if($CloseOperatorFirstForTest){
    $operatorReceipt=Get-Content (Join-Path $root 'operator-first-cleanup.json') -Raw|ConvertFrom-Json
    $operatorCloseValid=($operatorReceipt.MatchingOwnedWorkbooks -eq 1 -and $operatorReceipt.EnableEvents -and $operatorReceipt.CloseReturned -and $operatorReceipt.WorkbookCountReducedByOne -and $null -eq $operatorReceipt.Failure -and $null -ne $cleanup -and -not $cleanup.TerminationRequested)
}
$ownedCloseValid=$true
if($CloseOwnedWorkbooksForTest){
    $ownedReceipt=Get-Content (Join-Path $root 'owned-workbook-cleanup.json') -Raw|ConvertFrom-Json
    $ownedCloseValid=($ownedReceipt.WorkbooksBefore -gt 0 -and $ownedReceipt.CloseReturned -eq $ownedReceipt.WorkbooksBefore -and $ownedReceipt.WorkbooksAfter -eq 0 -and $null -eq $ownedReceipt.Failure -and $null -ne $cleanup -and $cleanup.ProcessIdAvailable -and -not $cleanup.TerminationRequested)
}
$reverseCloseValid=(-not $ReverseOpenedCloseForTest -or ($null -ne $cleanup -and $cleanup.ProcessIdAvailable -and -not $cleanup.TerminationRequested))
$result=[pscustomobject]@{TraceBoundaries=[bool]$TraceBoundaries;StandardRunFlow=[bool]$StandardRunFlow;FullProductionFlow=[bool]$FullProductionFlow;ExpectedChecks=$expectedChecks;CompileOnly=[bool]$CompileOnly;NativeExceptions=[bool]$NativeExceptions;NativeFaultsOnly=[bool]$NativeFaultsOnly;NativeBeforeRun=[bool]$NativeBeforeRun;NativeObserverValid=$nativeValid;ReleaseAutomationForTest=[bool]$ReleaseAutomationForTest;ReleaseObservationMilliseconds=$ReleaseObservationMilliseconds;ClearErrorReferencesForTest=[bool]$ClearErrorReferencesForTest;AutomationReleaseValid=$releaseValid;CloseOperatorFirstForTest=[bool]$CloseOperatorFirstForTest;ObserveShutdownForTest=[bool]$ObserveShutdownForTest;CloseOwnedWorkbooksForTest=[bool]$CloseOwnedWorkbooksForTest;ReverseOpenedCloseForTest=[bool]$ReverseOpenedCloseForTest;OperatorFirstCleanupValid=$operatorCloseValid;OwnedWorkbooksCleanupValid=$ownedCloseValid;ReverseOpenedCleanupValid=$reverseCloseValid;OriginalStatementsPreserved=$statementsPreserved;PassedChecks=$checks.Count-$failed;FailedChecks=$failed;Checks=$checks;CutReached=$cutReached;TraceAllowlistValid=$traceValid;TraceEntries=$stages.Count;ExcelApplicationEvents=$events.Count;FinalCleanup=$cleanup;DiagnosticPassed=($code -eq 0 -and $scopeSatisfied -and $failed -eq 0 -and $traceValid -and $nativeValid -and $releaseValid -and $operatorCloseValid -and $ownedCloseValid -and $reverseCloseValid -and $events.Count -eq 0);FullProductionAccepted=$false}
$result|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $root 'result.json')
$result|Select-Object TraceBoundaries,StandardRunFlow,CompileOnly,OriginalStatementsPreserved,PassedChecks,FailedChecks,CutReached,TraceAllowlistValid,TraceEntries,ExcelApplicationEvents,FinalCleanup,DiagnosticPassed|ConvertTo-Json -Depth 4
if(-not $result.DiagnosticPassed){exit 1}
