# Targeted diagnosis on a retained disposable fixture, never full release evidence.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$FixtureContextPath,
      [Parameter(Mandatory=$true)][string]$PackagePinsPath,
      [Security.SecureString]$FixtureCredential,
      [string]$DeployRoot='deploy/validation-process-worksheet-picker',[switch]$ValidateOnly)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before isolated replay.'}
$deploy=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$context=Get-Content -LiteralPath (Join-Path $repo $FixtureContextPath) -Raw|ConvertFrom-Json
if($context.RuntimeLeaf -cnotmatch '^invsys-plan022-launcher-red-[a-f0-9]{32}$'){throw 'Not a retained disposable fixture.'}
$runtimeRoot=[IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) $context.RuntimeLeaf))
$operatorPath=[IO.Path]::GetFullPath((Join-Path $runtimeRoot $context.OperatorRelativePath))
if(-not $operatorPath.StartsWith($runtimeRoot.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase) -or
    -not (Test-Path -LiteralPath $operatorPath -PathType Leaf)){throw 'Retained operator ownership failed.'}
$pins=Get-Content -LiteralPath (Join-Path $repo $PackagePinsPath) -Raw|ConvertFrom-Json
if($pins.Count -ne 5){throw 'Five package pins required.'}
foreach($pin in $pins){if((Get-FileHash (Join-Path $deploy (Split-Path -Leaf $pin.File))).Hash -cne $pin.Hash){throw 'Frozen candidate changed.'}}
$output=Join-Path $repo ('reports/runtime/production-restart-replay/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output|Out-Null
Write-Output ('Restart replay: '+$output)
# Import only the existing sanitized cross-package macro dispatcher.
$tokens=$null;$errors=$null
$helperAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_phase6_live_role_workflows.ps1'),[ref]$tokens,[ref]$errors)
$dispatcher=@($helperAst.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Run-WorkbookMacro'},$true))
if($errors.Count -or $dispatcher.Count -ne 1){throw 'Dispatcher declaration ambiguous.'}
. ([scriptblock]::Create($dispatcher[0].Extent.Text))
# The creator must pass the SAME random fixture credential in this process.
# Re-evaluating its random expression cannot authenticate an old fixture.
# Never serialize this parameter or pass a credential through a command line.
$testPin=$null
$testUser=if([string]::IsNullOrWhiteSpace($env:USERNAME)){'user1'}else{$env:USERNAME}
$checks=[Collections.Generic.List[object]]::new()
function Assert-Replay([string]$Name,[bool]$Passed){
    $checks.Add([pscustomobject]@{Name=$Name;Passed=$Passed})
    if(-not $Passed){throw 'Replay assertion failed; see sanitized checks.'}
}
function Assert-ReplayEnvelope([string]$Prefix,[string]$Value,[string[]]$Names){
    Assert-Replay ($Prefix+'.OK') ($Value.StartsWith('OK|'))
    foreach($name in $Names){Assert-Replay ($Prefix+'.'+$name) ($Value -cmatch ('(?:^|\|)'+[regex]::Escape($name)+'=True(?:\||$)'))}
}
if($ValidateOnly){
    Assert-ReplayEnvelope 'Calibration' 'OK|Allowed=True|Private=DO_NOT_COPY' @('Allowed')
    $serialized=$checks|ConvertTo-Json
    if($serialized.Contains('DO_NOT_COPY')){throw 'Envelope redaction failed.'}
    [pscustomobject]@{FixtureOwnership=$true;FivePackagePins=$true;DispatcherUnique=$true;CredentialInMemoryOnly=$true;EnvelopeRedacted=$true;ExcelOpened=$false;FullAcceptance=$false}|
        ConvertTo-Json|Set-Content (Join-Path $output 'calibration.json')
    return
}
if($null -eq $FixtureCredential -or $FixtureCredential.Length -eq 0){throw 'Live replay requires the creator-held credential in memory.'}
$testPin=[Net.NetworkCredential]::new('', $FixtureCredential).Password
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
. (Join-Path $PSScriptRoot 'ProductionReusableCleanup.ps1')
Add-Type 'using System; using System.Runtime.InteropServices; public static class RestartReplayWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint id); }'
function Release-ComObject($Value){if($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)}}
function Invoke-ReplaySession([bool]$Prepare){
    $stage=if($Prepare){'Prepare'}else{'Restart'}
    $excel=$null;$packages=@{};$ownedId=0;$closure=$null;$release=$null
    $packageNames=@('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')
    try{
        $excel=New-Object -ComObject Excel.Application
        [uint32]$windowId=0;[void][RestartReplayWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$windowId);$ownedId=[int]$windowId
        $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$true;$excel.AutomationSecurity=1
        foreach($name in $packageNames){$packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name))}
        function Core([string]$Macro,[object[]]$Values){Run-WorkbookMacro -Excel $excel -WorkbookName 'invSys.Core.xlam' -MacroName $Macro -Arguments $Values}
        [void](Core 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($runtimeRoot))
        Assert-Replay ($stage+'.Config') ([bool](Core 'modConfig.LoadConfig' @($context.WarehouseId,$context.StationId)))
        Assert-Replay ($stage+'.Auth') ([bool](Core 'modAuth.LoadAuth' @($context.WarehouseId)))
        Assert-Replay ($stage+'.Target') ([string](Core 'modNasConnection.SelectWarehouseTargetForAutomation' @($runtimeRoot,$runtimeRoot,$context.StationId,$true))).StartsWith('OK|')
        Assert-Replay ($stage+'.Paths') ([bool](Core 'modNasConnection.SetCurrentTargetPathsForTest' @('\\plan022-test\warehouse',$runtimeRoot)))
        Assert-Replay ($stage+'.OperatorRoot') ([bool](Core 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runtimeRoot 'operator-workbooks'))))
        $signIn=[string](Core 'modAuth.SignInCurrentTargetForAutomation' @($testUser,$testPin,'PROD_POST'))
        if($signIn -cmatch '^FAIL\|([0-9])$'){
            [pscustomobject]@{StatusCode=[int]$Matches[1]}|ConvertTo-Json|Set-Content (Join-Path $output ($stage+'-auth-status.json'))
        }
        Assert-Replay ($stage+'.SignIn') ($signIn.StartsWith('OK|'))
        [void](Run-WorkbookMacro $excel 'invSys.Operations.xlam' 'modOperationsInit.Auto_Open')
        [void](Run-WorkbookMacro $excel 'invSys.Operations.xlam' 'mProduction.BtnOpenProductionForm')
        $books=@($excel.Workbooks|Where-Object{-not $_.IsAddin -and [string]::Equals([string]$_.FullName,$operatorPath,[StringComparison]::OrdinalIgnoreCase)})
        Assert-Replay ($stage+'.SameOperator') ($books.Count -eq 1)
        if($Prepare){
            $tables=0
            foreach($sheet in $books[0].Worksheets){foreach($table in $sheet.ListObjects){if($table.Name -like 'invSys_Process_*'){$tables++}}}
            Assert-Replay 'Prepare.NoOutstandingTables' ($tables -eq 0)
            $value=[string](Run-WorkbookMacro $excel 'invSys.Operations.xlam' 'mProduction.RunProcessWorksheetWorkbenchContractTest')
            Assert-ReplayEnvelope 'Workbench' $value @('SeparateActions','MultipleTables','SelectedOnly','RecordTypeDropdown','CalculatedPercent','GeneratedDesign','ItemCodeRemoved','Assignments','ItemSearch')
            $books[0].Save()
        }else{
            $before=[int]$excel.Workbooks.Count
            $value=[string](Run-WorkbookMacro $excel 'invSys.Operations.xlam' 'mProduction.RunReusableProductionRestartActionContractTest' @($context.RecipeId,$context.RecipeVersion,$operatorPath))
            Assert-ReplayEnvelope 'Restart' $value @('RecipeFound','Loaded','SameWorkbook','WorksheetRediscovered','WorksheetRetrieved','MultipleTablesRediscovered','SelectedOnly','AllRetrieved')
            Assert-Replay 'Restart.NoNewBooks' ([int]$excel.Workbooks.Count -eq $before)
        }
    }finally{
        if($null -ne $excel){
            $closure=Invoke-ReusableWorkbookClosure $excel $runtimeRoot $packages $packageNames $deploy
            try{$excel.Quit()}catch{}
            $release=Wait-ReusableAutomationExit -ProcessId $ownedId -Variables (Get-Variable -Scope Local)
        }
        [pscustomobject]@{Stage=$stage;Closure=$closure;Release=$release;EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ForcedTermination=$false}|
            ConvertTo-Json -Depth 6|Set-Content (Join-Path $output ($stage+'-closure.json'))
    }
    Assert-Replay ($stage+'.NormalExit') ($closure.Completed -and $release.UnassistedExitObserved -and $null -eq $release.Failure -and $release.References.ReleaseFailures -eq 0)
}
$settings=Get-InvSysTestSettingsSnapshot;$started=[DateTimeOffset]::UtcNow;$failure=$null
try{Invoke-ReplaySession $true;Invoke-ReplaySession $false}
catch{$failure=Get-RecordingFailureFacts $_}
finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $preserved=@($pins|Where-Object{(Get-FileHash (Join-Path $deploy (Split-Path -Leaf $_.File))).Hash -cne $_.Hash}).Count -eq 0
    $checks|ConvertTo-Json|Set-Content (Join-Path $output 'checks.json')
    [pscustomobject]@{StartUTC=$started.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');Failure=$failure;SettingsRestored=$restored;PackagesPreserved=$preserved;ExcelClosed=$true;Debugger=$false;VbeInstrumentation=$false;FullAcceptance=$false}|
        ConvertTo-Json -Depth 4|Set-Content (Join-Path $output 'closure.json')
}
if($null -ne $failure -or -not ($restored -and $preserved)){throw 'Targeted replay failed; see sanitized evidence.'}
Write-Output ('Targeted replay passed: '+$checks.Count+' checks; not full acceptance.')
