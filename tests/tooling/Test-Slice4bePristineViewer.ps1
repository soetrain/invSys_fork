[CmdletBinding()]
param([string]$RepoRoot='', [string]$DeployRoot='deploy/validation-recording-isolation')
# Isolate cold packaged Viewer launch: no prior interruption and no VBA edits.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(-not $RepoRoot){$RepoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)}
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$deploy=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before pristine Viewer validation.'}
$runRoot=Join-Path ([IO.Path]::GetTempPath()) ('invsys-pristine-viewer-'+[guid]::NewGuid().ToString('N'))
$reportRoot=Join-Path $repo ('reports/runtime/pristine-viewer/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $reportRoot -Force|Out-Null
$results=[Collections.Generic.List[object]]::new()
$excel=$null; $packages=@{}; $preparedShippingFixtures=@{}; $preparedShippingBoundaries=@{}
$step='startup'; $initialExcelProcessIds=@()
function Check([string]$Name,[bool]$Passed){
    $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))
}
# Import only the established fixture/bootstrap and invocation functions, not
# the harness body or its VBA instrumentation. Their sources stay authoritative.
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Shared fixture source did not parse.'}
foreach($name in @('Run','Table','CredentialHash','SelectTarget','NewFixture')){
    $definitions=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq $name},$true))
    if($definitions.Count -ne 1){throw 'Shared fixture function is not unique.'}
    . ([scriptblock]::Create($definitions[0].Extent.Text))
}
$settingsRoot='HKCU:\Software\VB and VBA Program Settings\invSys'
$registryBefore=@{}
if(Test-Path -LiteralPath $settingsRoot){
    foreach($key in @(Get-Item -LiteralPath $settingsRoot)+@(Get-ChildItem -LiteralPath $settingsRoot -Recurse)){
        $values=@{}
        foreach($name in $key.GetValueNames()){$values[$name]=@($key.GetValue($name),$key.GetValueKind($name))}
        $registryBefore[$key.Name]=$values
    }
}
$packageNames=@('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')
$packagePins=@{}
foreach($name in $packageNames){$path=Join-Path $deploy $name;$packagePins[$path]=(Get-FileHash -LiteralPath $path).Hash}
try {
    $excel=New-Object -ComObject Excel.Application
    $initialExcelProcessIds=@(Get-Process EXCEL -ErrorAction Stop|Select-Object -ExpandProperty Id)
    if($initialExcelProcessIds.Count -ne 1){throw 'Pristine test requires one isolated Excel process.'}
    [pscustomobject]@{ExcelProcessId=$initialExcelProcessIds[0];ControllerId=$PID;NoPriorInterruption=$true}|
        ConvertTo-Json|Set-Content -LiteralPath (Join-Path $reportRoot 'processes.json')
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    foreach($name in $packageNames){
        $packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name),0,$true)
        if([string]$packages[$name].FullName -ine (Join-Path $deploy $name)){throw 'Pristine package path mismatch.'}
    }
    Check 'PristineViewer.FiveCandidatePackagesLoaded' ($packages.Count -eq 5)
    $step='Admin-generated fixture'
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $repo 'deploy/current/templates')))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    $fixture=NewFixture 'a'
    SelectTarget $fixture 'config-admin'
    Check 'PristineViewer.GeneratedFixtureSignedIn' $true
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $fixture.Root -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    $step='pristine Viewer callback'
    Write-Output 'Pristine Viewer: enter the existing packaged action wrapper.'
    $first=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RunInventoryViewerActionForTest')
    $match=[regex]::Match($first,'\|Generation=([1-9][0-9]*)\|')
    $opened=$first.StartsWith('OK|') -and $match.Success
    Check 'PristineViewer.FirstLaunch' $opened
    if($opened){
        $generation=$match.Groups[1].Value
        $second=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RunInventoryViewerActionForTest')
        Check 'PristineViewer.RepeatedLaunchReusesForm' ($second.StartsWith('OK|') -and $second.Contains('|Generation='+$generation+'|'))
    }
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
    $unchanged=$pins.Count -eq @(Get-ChildItem -LiteralPath $fixture.Root -Recurse -File).Count
    foreach($path in $pins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $pins[$path]){$unchanged=$false}}
    Check 'PristineViewer.FixtureFilesPreserved' $unchanged
} catch {
    Check ('PristineViewer.HarnessFailure.'+$step.Replace(' ','')) $false
} finally {
    if($null -ne $excel){
        try{[void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')}catch{}
        try{foreach($book in @($excel.Workbooks)){$book.Close($false)};$excel.Quit()}catch{}
        try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}catch{}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if(Test-Path -LiteralPath $settingsRoot){
        foreach($key in @(Get-Item -LiteralPath $settingsRoot)+@(Get-ChildItem -LiteralPath $settingsRoot -Recurse)){
            foreach($name in $key.GetValueNames()){
                if(-not $registryBefore.ContainsKey($key.Name) -or -not $registryBefore[$key.Name].ContainsKey($name)){
                    Remove-ItemProperty -LiteralPath ('Registry::'+$key.Name) -Name $name
                }
            }
        }
    }
    foreach($path in $registryBefore.Keys){
        foreach($name in $registryBefore[$path].Keys){
            $saved=$registryBefore[$path][$name]
            New-ItemProperty -LiteralPath ('Registry::'+$path) -Name $name -Value $saved[0] -PropertyType $saved[1] -Force|Out-Null
        }
    }
    $resolved=[IO.Path]::GetFullPath($runRoot)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'invsys-pristine-viewer-*'){
        if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
    }
    $unchanged=$true
    foreach($path in $packagePins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $packagePins[$path]){$unchanged=$false}}
    Check 'PristineViewer.FivePackageFilesUnchanged' $unchanged
    $results|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $reportRoot 'checks.json')
}
$failed=@($results|Where-Object {-not $_.Passed}).Count
Write-Output ('Pristine Viewer: '+($results.Count-$failed)+' passed, '+$failed+' failed')
if($failed){exit 1}
