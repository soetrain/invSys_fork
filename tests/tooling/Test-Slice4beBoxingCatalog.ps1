# Supplemental catalog/reference regression; real-handler RED/GREEN remains required.
[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/current')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before catalog validation.'}
$path=Join-Path (Join-Path $repo $DeployRoot) 'invSys.Core.xlam'
$pin=(Get-FileHash -LiteralPath $path).Hash
$checks=New-Object 'System.Collections.Generic.List[object]'
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))}
function Probe([string]$Name,$First,$Second,$Third){
    if($PSBoundParameters.ContainsKey('Third')){return $excel.Run("'invSys.Core.xlam'!TestShippingCatalog."+$Name,$First,$Second,$Third)}
    if($PSBoundParameters.ContainsKey('Second')){return $excel.Run("'invSys.Core.xlam'!TestShippingCatalog."+$Name,$First,$Second)}
    return $excel.Run("'invSys.Core.xlam'!TestShippingCatalog."+$Name,$First)
}
$excel=$null;$book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    $book=$excel.Workbooks.Open($path,0,$true)
    if($book.ReadOnly -isnot [bool] -or -not $book.ReadOnly){throw 'Catalog package did not open read-only.'}
    $packages=@{'invSys.Core.xlam'=$book}
    . (Join-Path $PSScriptRoot 'Slice4beShippingCatalog.ps1')
    Install-Slice4beShippingCatalogProbe
    $old=@(([string](Probe 'Ids' 8)).Split("`n")|Where-Object {$_ -ne ''})
    $current=@(([string](Probe 'Ids' 9)).Split("`n")|Where-Object {$_ -ne ''})
    Check 'Boxing.Catalog.NineExtendsEight' ($old.Count -eq 31 -and $current.Count -eq 33 -and @($old|Where-Object {$_ -cnotin $current}).Count -eq 0 -and @($current|Select-Object -Unique).Count -eq 33)
    foreach($id in $old){Check ('Boxing.Catalog.Preserved.'+$id) ([string](Probe 'Definition' $id 8) -cne '' -and [string](Probe 'Definition' $id 8) -ceq [string](Probe 'Definition' $id 9))}
    $submitted='[{"WarehouseId":"CATALOG_TEST","SourceKind":"Inventory","EventId":"Source_A","SubmissionState":"Submitted"}]'
    $unknown=$submitted.Replace('Submitted','Unknown')
    $invalid=@{
        Duplicate=$submitted.Substring(0,$submitted.Length-1)+','+$submitted.Substring(1)
        CrossWarehouse=$submitted.Replace('CATALOG_TEST','ANOTHER_TEST')
        InvalidIdentity=$submitted.Replace('Source_A','Source A')
        UnknownField=$submitted.Replace('"EventId":','"Extra":"forbidden","EventId":')
        UnknownSource=$submitted.Replace('Inventory','Other')
        UnknownState=$submitted.Replace('Submitted','Applied')
    }
    foreach($action in @('MAKE','UNBOX')){
        $id='BOXING_'+$action;$prefix='Boxing.Catalog.'+$action
        $value=[string](Probe 'Definition' $id 9)|ConvertFrom-Json
        $caption=if($action -ceq 'MAKE'){'Make Boxes'}else{'Unbox'}
        Check ($prefix+'.Definition') ($null -ne $value -and $value.ControlId -ceq $id -and $value.OwnerId -ceq 'BOXING_WORKFLOW' -and $value.Role -ceq 'Boxing' -and $value.Class -ceq 'Command' -and $value.Caption -ceq $caption -and $value.Surface -ceq 'Operations > Shipping > Box Maker' -and $value.Capability -ceq 'SHIP_POST' -and $value.CodePrefix -ceq ($id+'_'))
        foreach($version in 1..8){Check ($prefix+'.ExcludedFrom.'+$version) ([string](Probe 'Definition' $id $version) -ceq '')}
        Check ($prefix+'.UnsupportedVersionRejected') ([string](Probe 'Definition' $id 10) -ceq '')
        $outcomes=[ordered]@{REQUESTED=@('Info','Unknown');DENIED=@('Blocked','Unchanged');REJECTED=@('Warning','Unchanged');PENDING=@('Notice','Unknown');CONFIRMED=@('Info','Unknown');FAILED=@('Error','Unknown')}
        foreach($code in $outcomes.Keys){
            $value=[string](Probe 'Outcome' $id $code)|ConvertFrom-Json
            Check ($prefix+'.Outcome.'+$code) ($null -ne $value -and $value.OutcomeCode -ceq $code -and $value.EventCode -ceq ($id+'_'+$code) -and $value.Severity -ceq $outcomes[$code][0] -and $value.DataEffect -ceq $outcomes[$code][1] -and $value.UserMessage -ne '' -and $null -ne $value.PSObject.Properties['NextStep'])
        }
        foreach($code in @('STAGED','COMPLETED')){
            Check ($prefix+'.InvalidOutcome.'+$code) ([string](Probe 'Outcome' $id $code) -ceq '')
            $result=Probe 'References' $id $code '[]'
            Check ($prefix+'.InvalidReferencesOutcome.'+$code) ($result -is [bool] -and -not $result)
        }
        foreach($code in @('REQUESTED','DENIED','REJECTED')){
            $empty=Probe 'References' $id $code '[]';$source=Probe 'References' $id $code $submitted
            Check ($prefix+'.NoSource.'+$code) ($empty -is [bool] -and $empty -and $source -is [bool] -and -not $source)
        }
        foreach($code in @('PENDING','CONFIRMED')){
            $accepted=Probe 'References' $id $code $submitted;$empty=Probe 'References' $id $code '[]';$uncertain=Probe 'References' $id $code $unknown
            Check ($prefix+'.AcceptedSourcesOnly.'+$code) ($accepted -is [bool] -and $accepted -and $empty -is [bool] -and -not $empty -and $uncertain -is [bool] -and -not $uncertain)
        }
        foreach($refs in @('[]',$submitted,$unknown)){
            $kind=if($refs -ceq '[]'){'Empty'}elseif($refs -ceq $submitted){'Submitted'}else{'Unknown'}
            $result=Probe 'References' $id 'FAILED' $refs
            Check ($prefix+'.Failed.'+$kind) ($result -is [bool] -and $result)
        }
        foreach($kind in $invalid.Keys){$result=Probe 'References' $id 'FAILED' $invalid[$kind];Check ($prefix+'.InvalidSource.'+$kind) ($result -is [bool] -and -not $result)}
    }
}finally{
    if($null -ne $book){$book.Close($false)}
    if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    Check 'Boxing.Catalog.PackageBytesPreserved' ($pin -ceq (Get-FileHash -LiteralPath $path).Hash)
    $out=Join-Path $repo 'reports/runtime/boxing-catalog-regression.json'
    $checks|ConvertTo-Json|Set-Content -LiteralPath $out
}
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
