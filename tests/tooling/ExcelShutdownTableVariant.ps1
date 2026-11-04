# Diagnostic-only transformation of a temporary worker. Never edits the source
# worker or any package, and never changes its final Quit/release/GC sequence.
function ConvertTo-ExcelShutdownTableVariant {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [ValidateSet('Baseline','OwnedTableTraversal')][string]$Mode='Baseline'
    )
    if($Mode -ceq 'Baseline'){return $Source}
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Table variant requires a parseable worker.'}
    $functions=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $n.Name -in @('Table','Release-ExcelShutdownTableReference')},$true))
    if($functions.Count -ne 1 -or $functions[0].Name -cne 'Table'){throw 'Unique original Table helper required.'}
    $expected=@'
function Table($Workbook,[string]$Name) {
    foreach ($sheet in $Workbook.Worksheets) {
        foreach ($candidate in $sheet.ListObjects) { if($candidate.Name -eq $Name){return $candidate} }
    }
    throw "Fixture table missing: $Name"
}
'@
    $original=$functions[0].Extent.Text
    if($original.Replace("`r`n","`n") -cne $expected.Replace("`r`n","`n")){
        throw 'Table helper anchor drift; no workload dispatched.'
    }
    $replacement=@'
function Release-ExcelShutdownTableReference($Value) {
    if($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)){
        # One release per owned acquisition, never FinalReleaseComObject.
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}
function Table($Workbook,[string]$Name) {
    # Workbook is borrowed. The matched table is transferred to the caller.
    $sheets=$Workbook.Worksheets
    try {
        for($sheetIndex=1;$sheetIndex -le $sheets.Count;$sheetIndex++){
            $sheet=$sheets.Item($sheetIndex)
            try {
                $tables=$sheet.ListObjects
                try {
                    for($tableIndex=1;$tableIndex -le $tables.Count;$tableIndex++){
                        $candidate=$tables.Item($tableIndex)
                        try {
                            if($candidate.Name -eq $Name){
                                $result=$candidate
                                $candidate=$null
                                return $result
                            }
                        } finally {Release-ExcelShutdownTableReference $candidate}
                    }
                } finally {Release-ExcelShutdownTableReference $tables}
            } finally {Release-ExcelShutdownTableReference $sheet}
        }
    } finally {Release-ExcelShutdownTableReference $sheets}
    throw "Fixture table missing: $Name"
}
'@
    $result=$Source.Remove($functions[0].Extent.StartOffset,$original.Length).Insert($functions[0].Extent.StartOffset,$replacement)
    [void][Management.Automation.Language.Parser]::ParseInput($result,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Table variant does not parse.'}
    return $result
}
