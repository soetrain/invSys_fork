# Optional harness timing only: no Excel queries, arguments, return values,
# exception text, paths or warehouse/user identifiers are inspected or emitted.
function Write-InventoryCallTiming([string]$CallId,[string]$Macro,
        [ValidateSet('Started','Ready','Finished')][string]$Stage,
        [long]$ElapsedMilliseconds,[bool]$Returned) {
    if($Macro -cnotmatch '^[A-Za-z_][A-Za-z0-9_]*\.[A-Za-z_][A-Za-z0-9_]*$'){$Macro='Unclassified'}
    [pscustomobject]@{
        UTC=[DateTimeOffset]::UtcNow.ToString('o')
        CallId=$CallId
        Macro=$Macro
        Stage=$Stage
        ElapsedMilliseconds=$ElapsedMilliseconds
        Returned=$Returned
    }|ConvertTo-Json -Compress|Add-Content -LiteralPath (Join-Path $reportRoot 'inventory-call-timing.jsonl')
}
