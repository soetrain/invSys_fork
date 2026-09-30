# Test-only fresh-process restart trace; preserved fixtures are never release evidence.
[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-process-worksheet-picker',
      [Parameter(Mandatory=$true)][string]$PackagePinsPath,[switch]$GenerateOnly,[switch]$UnobservedReplay)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before isolated restart diagnosis.'}
$deploy=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$diagnosticPins=Get-Content (Join-Path $repo $PackagePinsPath) -Raw|ConvertFrom-Json
if($diagnosticPins.Count -ne 5){throw 'Five frozen package pins required.'}
foreach($pin in $diagnosticPins){if((Get-FileHash (Join-Path $deploy (Split-Path -Leaf $pin.File))).Hash -cne $pin.Hash){throw 'Frozen package differs.'}}
$diagnosticRoot=Join-Path $repo ('reports/runtime/production-restart-diagnostic/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $diagnosticRoot|Out-Null
Write-Output ('Restart diagnostic: '+$diagnosticRoot)
$validator=Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'
$validatorHash=(Get-FileHash $validator).Hash
$source=[IO.File]::ReadAllText($validator);$original=$source
$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Original validator parse failed.'}
$pinAssignment=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -ceq '$testPin'},$true))
if($pinAssignment.Count -ne 1){throw 'Credential declaration ambiguous.'}
# Reuse the existing validator's fixture credential only in memory; never copy it into generated files.
$RestartDiagnosticPin=& ([scriptblock]::Create($pinAssignment[0].Right.Extent.Text))
$credentialOriginal=$pinAssignment[0].Extent.Text
$credentialReplacement='$testPin = $RestartDiagnosticPin'
$source=$source.Replace($credentialOriginal,$credentialReplacement)
$patches=[Collections.Generic.List[object]]::new()
function Replace-RestartDiagnosticAnchor([string]$Anchor,[string]$Replacement) {
    $pattern='(?m)^'+[regex]::Escape($Anchor)+'\r?$'
    if([regex]::Matches($script:source,$pattern).Count -ne 1){throw 'Restart diagnostic anchor missing/ambiguous.'}
    $script:source=[regex]::Replace($script:source,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m)$Replacement})
    $patches.Add([pscustomobject]@{Original=$Anchor;Replacement=$Replacement})
}
$init=@'
$diagnosticFixtureReady=$false
$restartObserver=$null;$restartObservedProcess=$null
$restartNativeRoot=Join-Path $outputPath 'restart-native'
New-Item -ItemType Directory -Path $restartNativeRoot|Out-Null
'@
Replace-RestartDiagnosticAnchor '$excel = $null' ($init+"`r`n"+'$excel = $null')
$retain=@'
        if($restartTerminationRequested -or -not $restartRelease.UnassistedExitObserved){throw 'Initial process did not close normally; restart withheld.'}
        $runtimePrefix=[IO.Path]::GetFullPath($runtimeRoot).TrimEnd('\')+'\'
        $operatorPath=[IO.Path]::GetFullPath($productionOperatorPath)
        if(-not $operatorPath.StartsWith($runtimePrefix,[StringComparison]::OrdinalIgnoreCase)){throw 'Operator fixture outside owned runtime.'}
        [pscustomobject]@{RuntimeLeaf=(Split-Path -Leaf $runtimeRoot);WarehouseId=$warehouseId;StationId=$stationId;RecipeId=$reusableRecipeId;RecipeVersion=$reusableRecipeVersion;OperatorRelativePath=$operatorPath.Substring($runtimePrefix.Length);RetainedStateMayIncludeRestartMutations=$true}|
            ConvertTo-Json|Set-Content (Join-Path $outputPath 'fixture-context.json')
        $diagnosticFixtureReady=$true
'@
$anchor="        `$opened = New-Object 'System.Collections.Generic.List[object]'"
Replace-RestartDiagnosticAnchor $anchor ($retain+"`r`n"+$anchor)
$observe=@'
        $restartObservedProcess=Get-Process -Id $excelProcessId
        if($restartObservedProcess.ProcessName -cne 'EXCEL'){throw 'Fresh owned Excel unavailable.'}
        $null=$restartObservedProcess.Handle
        $nativeArguments=@([string]$excelProcessId,('"'+$restartNativeRoot+'"'),'600','faults-only')
        $restartObserver=Start-Process -FilePath (Join-Path $outputPath 'NativeExceptionObserver.exe') -ArgumentList $nativeArguments -WindowStyle Hidden -PassThru
        $deadline=[DateTime]::UtcNow.AddSeconds(20)
        while(-not (Test-Path (Join-Path $restartNativeRoot 'ready'))){
            if($restartObserver.HasExited -or [DateTime]::UtcNow -gt $deadline){throw 'Restart observer not ready; workflow withheld.'}
            Start-Sleep -Milliseconds 100
        }
'@
if(-not $UnobservedReplay){Replace-RestartDiagnosticAnchor '        $excel.Visible = $true' ($observe+"`r`n"+'        $excel.Visible = $true')}
$trace=@'
        . (Join-Path $repo 'tests/tooling/ProductionRestartTrace.ps1')
        Install-ProductionRestartTrace -Excel $excel -Packages $packages -PackageRoot $deployPath
        $tracePath=Join-Path $outputPath 'restart-phases.txt'
        [void](Run-WorkbookMacro -Excel $excel -WorkbookName 'invSys.Operations.xlam' -MacroName 'TestProductionRestartTrace.Arm' -Arguments @($tracePath))
        [void](Run-WorkbookMacro -Excel $excel -WorkbookName 'invSys.Operations.xlam' -MacroName 'TestProductionRestartTrace.Mark' -Arguments @('UNREGISTERED_RESTART_SENTINEL'))
        if(-not (Test-Path $tracePath) -or [IO.File]::ReadAllText($tracePath).Trim() -cne 'Arm'){throw 'Restart trace redaction/transport failed; no product RED.'}
        [pscustomobject]@{UnknownLabelRejected=$true;ArmTransportPassed=$true;LoadedProjectsCompiled=4}|ConvertTo-Json|Set-Content (Join-Path $outputPath 'trace-calibration.json')
'@
$anchor='        $coreName = [string]$packages["invSys.Core.xlam"].Name'
if(-not $UnobservedReplay){Replace-RestartDiagnosticAnchor $anchor ($trace+"`r`n"+$anchor)}
$stop=@'
    if($null -ne $restartObserver){
        [IO.File]::WriteAllText((Join-Path $restartNativeRoot 'stop'),'Stop')
        if(-not $restartObserver.WaitForExit(5000)){throw 'Restart observer remains live; preserve it.'}
        $exitCode=$null;$exited=$false
        if($null -ne $restartObservedProcess){$exited=$restartObservedProcess.HasExited;if($exited){$exitCode=$restartObservedProcess.ExitCode.ToString('X8')}}
        [pscustomobject]@{ObserverExitCode=$restartObserver.ExitCode;ObservedExcelExited=$exited;ObservedExcelExitCode=$exitCode;ForcedTermination=$cleanupTerminationRequested}|
            ConvertTo-Json|Set-Content (Join-Path $outputPath 'restart-native-exit.json')
    }
'@
$anchor='    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())'
Replace-RestartDiagnosticAnchor $anchor ($stop+"`r`n"+$anchor)
$anchor='        Remove-Item -LiteralPath $resolvedRuntime -Recurse -Force -ErrorAction SilentlyContinue'
Replace-RestartDiagnosticAnchor $anchor ('        if(-not $diagnosticFixtureReady){'+"`r`n"+$anchor+"`r`n"+'        }')
$restored=$source
for($i=$patches.Count-1;$i -ge 0;$i--){$restored=$restored.Replace($patches[$i].Replacement,$patches[$i].Original)}
$restored=$restored.Replace($credentialReplacement,$credentialOriginal)
$same=($restored.Replace("`r`n","`n") -ceq $original.Replace("`r`n","`n"))
if(-not $same){throw 'Diagnostic changed original workflow statements.'}
$tokens=$null;$errors=$null;$generatedAst=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Generated restart diagnostic parse failed.'}
$credentialLiterals=@($generatedAst.FindAll({param($n)$n -is [Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -ceq $RestartDiagnosticPin},$true))
if($credentialLiterals.Count){throw 'Generated diagnostic must not persist fixture credentials.'}
$generated=Join-Path $diagnosticRoot 'restart-validator.ps1'
[IO.File]::WriteAllText($generated,$source,[Text.UTF8Encoding]::new($false))
[pscustomobject]@{OriginalStatementsPreserved=$same;ObserverOnlyAtFreshRestart=(-not $UnobservedReplay);TraceOnlyAtFreshRestart=(-not $UnobservedReplay);UnobservedReplay=[bool]$UnobservedReplay;CredentialInheritedOnlyInMemory=$true;RetainsExistingDisposableRuntime=$true;GenerationOpenedExcel=$false;ParseErrors=0;FullAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $diagnosticRoot 'generation.json')
if($GenerateOnly){return}
$nativeSource=Join-Path $PSScriptRoot 'NativeExceptionObserver.cs'
$compiler=New-Object CodeDom.Compiler.CompilerParameters -Property @{CompilerOptions='/platform:x64';GenerateExecutable=$true;OutputAssembly=(Join-Path $diagnosticRoot 'NativeExceptionObserver.exe')}
[void]$compiler.ReferencedAssemblies.Add('System.dll')
if(-not $UnobservedReplay){Add-Type -Path $nativeSource -CompilerParameters $compiler}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot;$start=[DateTimeOffset]::UtcNow;$code=1
try {
    & $generated -RepoRoot $repo -DeployRoot $DeployRoot -OutputDirectory ($diagnosticRoot.Substring($repo.Length+1)) -CallbackFilter Production -WorkbookState ProductionReusable *> (Join-Path $diagnosticRoot 'worker.log')
    $code=$LASTEXITCODE
    if($UnobservedReplay -and $code -eq 0){
        $code=1
        $replayCredential=ConvertTo-SecureString $RestartDiagnosticPin -AsPlainText -Force
        & (Join-Path $PSScriptRoot 'Test-ProductionRestartReplay.ps1') -FixtureContextPath (Join-Path ($diagnosticRoot.Substring($repo.Length+1)) 'fixture-context.json') -PackagePinsPath $PackagePinsPath -DeployRoot $DeployRoot -FixtureCredential $replayCredential
        $code=0
    }
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $settingsRestored=Restore-InvSysTestSettingsSnapshot $settings
    $preserved=@($diagnosticPins|Where-Object{(Get-FileHash (Join-Path $deploy (Split-Path -Leaf $_.File))).Hash -cne $_.Hash}).Count -eq 0
    $validatorPreserved=(Get-FileHash $validator).Hash -ceq $validatorHash
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$settingsRestored;PackagesPreserved=$preserved;ValidatorPreserved=$validatorPreserved;ExcelClosed=$true;FullAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $diagnosticRoot 'closure.json')
    if(-not ($settingsRestored -and $preserved -and $validatorPreserved)){throw 'Restart diagnostic preservation failed.'}
}
Get-Content (Join-Path $diagnosticRoot 'worker.log') -Tail 7
exit $code
