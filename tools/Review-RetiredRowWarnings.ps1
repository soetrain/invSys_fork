# These exact source revisions use ROW only as a prohibited-name guard or a
# transient, rechecked staging-row locator. Keep the scanner warnings visible.
# Any edit to a reviewed file requires a fresh review; new paths always fail.
function Get-ReviewedSourceHash {
    param([Parameter(Mandatory=$true)][string]$Path)
    $bytes=[IO.File]::ReadAllBytes($Path)
    $normalized=[IO.MemoryStream]::new()
    try{
        for($i=0;$i -lt $bytes.Length;$i++){
            if($bytes[$i] -eq 10 -and ($i -eq 0 -or $bytes[$i-1] -ne 13)){
                $normalized.WriteByte(13)
            }
            $normalized.WriteByte($bytes[$i])
        }
        $sha=[Security.Cryptography.SHA256]::Create()
        try{return [BitConverter]::ToString($sha.ComputeHash($normalized.ToArray())).Replace('-','')}
        finally{$sha.Dispose()}
    }finally{$normalized.Dispose()}
}

function Classify-StaticWarningDelta {
    [CmdletBinding()]
    param(
        [object[]]$ActualWarnings=@(),
        [object[]]$BaselineWarnings=@(),
        [Parameter(Mandatory=$true)][string]$RepoRoot
    )

    $reviewedHashes=@{
        'src/Admin/Modules/modInventoryUiText.bas'='89DDFDBE42BF020A021436D6548DEDFEC767FABDCC2088E134A6E170E22DA87A'
        'src/Core/ClassModules/cInventoryCreatePlan.cls'='FF1BE2980484C6488DA62962272E5BD51A6E12087DD3C07C5591A151B2097707'
        'src/Core/ClassModules/cInventoryEditRow.cls'='FC2CCF269832E7E538C5F8C92B2D8080031F16BFCBD30E0ACA681527DF8546A2'
        'src/Core/ClassModules/cInventoryTableInput.cls'='C22B3306CEA93E2244419BB8396C9CED6A85A59BA731282577DBD101A657F868'
        'src/Core/Modules/modAggregationSources.bas'='B4E604BAA94CDCFB388D98D1C695A5AA39D68B2A02D259C6695A1DFEEDC956CB'
        'src/Core/Modules/modInventoryEditCommands.bas'='835DBF8B6C98526116B9692DF6FC84F18CB6452B1BD19C43EA92FE603D0ED178'
    }
    $baseline=[System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($warning in $BaselineWarnings){[void]$baseline.Add(([string]$warning.code+'|'+[string]$warning.sourcePath))}
    $reviewed=[System.Collections.Generic.List[object]]::new()
    $unreviewed=[System.Collections.Generic.List[object]]::new()
    foreach($warning in $ActualWarnings){
        $code=[string]$warning.code
        $path=[string]$warning.sourcePath
        if($baseline.Contains($code+'|'+$path)){continue}
        if($code -ceq 'RETIRED_ROW_HEADER' -and $reviewedHashes.ContainsKey($path)){
            $source=Join-Path $RepoRoot ($path -replace '/','\')
            if(Test-Path -LiteralPath $source){
                if((Get-ReviewedSourceHash -Path $source) -ceq $reviewedHashes[$path]){
                    $reviewed.Add($warning)
                    continue
                }
            }
        }
        $unreviewed.Add($warning)
    }
    return [pscustomobject]@{Reviewed=@($reviewed.ToArray());Unreviewed=@($unreviewed.ToArray())}
}
