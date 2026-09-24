[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ReportRoot)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=(Resolve-Path -LiteralPath $ReportRoot).Path
$durations=New-Object 'System.Collections.Generic.List[object]'
foreach($role in @('Operations','Core')){
    $stack=New-Object 'System.Collections.Generic.Stack[object]'
    foreach($line in Get-Content -LiteralPath (Join-Path $root ('view-calls-'+$role+'.tsv'))){
        if([string]::IsNullOrWhiteSpace($line)){continue}
        $fields=$line.Trim() -split "`t"
        if($fields.Count -ne 3 -or $fields[2] -notmatch '^([A-Za-z0-9_.]+)\|(Enter|Exit|End)$'){throw 'Unexpected trace record.'}
        $label=$matches[1];$action=$matches[2]
        $tick=[long]$fields[1];if($tick -lt 0){$tick+=4294967296}
        $utc=[DateTime]::SpecifyKind([DateTime]::ParseExact($fields[0],'yyyy-MM-dd HH:mm:ss',[Globalization.CultureInfo]::InvariantCulture),[DateTimeKind]::Local).ToUniversalTime()
        if($action -ceq 'Enter'){
            $stack.Push([pscustomobject]@{Label=$label;Tick=$tick;UTC=$utc})
        } else {
            if($stack.Count -eq 0 -or $stack.Peek().Label -cne $label){throw 'Unbalanced procedure trace; do not infer durations.'}
            $entry=$stack.Pop();$elapsed=$tick-$entry.Tick;if($elapsed -lt 0){$elapsed+=4294967296}
            $durations.Add([pscustomobject]@{Role=$role;Procedure=$label;StartUTC=$entry.UTC.ToString('o');ElapsedMs=$elapsed;Depth=$stack.Count})
        }
    }
    if($stack.Count -ne 0){throw 'Incomplete procedure trace; wait for normal completion.'}
}
$summary=@($durations|Group-Object Role,Procedure|ForEach-Object {
    $measure=$_.Group|Measure-Object ElapsedMs -Sum -Maximum
    [pscustomobject]@{Role=$_.Group[0].Role;Procedure=$_.Group[0].Procedure;Count=$_.Count;InclusiveTotalMs=$measure.Sum;MaximumMs=$measure.Maximum}
})
$summary|Sort-Object InclusiveTotalMs -Descending|Format-Table -AutoSize
$durations|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $root 'view-call-durations.json')
$summary|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $root 'view-call-summary.json')
