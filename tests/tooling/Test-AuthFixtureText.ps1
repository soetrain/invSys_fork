# Storage-only fixture test. Values are generated in memory, never reported;
# its private workbook files are removed after each save/reopen observation.
[CmdletBinding()]
param([string]$RepoRoot='.',[ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the fixture gate.'}
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_phase6_live_role_workflows.ps1'),[ref]$tokens,[ref]$errors)
if($errors){throw 'Fixture source does not parse; not behavioral RED.'}
foreach($definition in $ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$true)){
    . ([scriptblock]::Create($definition.Extent.Text))
}
$root=Join-Path $repo ('reports/runtime/auth-fixture-text/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$rows=New-Object 'System.Collections.Generic.List[object]'
$excel=New-Object -ComObject Excel.Application
$paths=@()
try {
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    foreach($kind in @('LeadingZero','ExponentLike','OrdinaryHex')){
        $value=switch($kind){
            'LeadingZero' {'0'+(Get-Random -Minimum 1000000 -Maximum 10000000).ToString()}
            'ExponentLike' {(Get-Random -Minimum 1 -Maximum 8).ToString()+'E'+(Get-Random -Minimum 10 -Maximum 30).ToString('D6')}
            'OrdinaryHex' {'A'+[guid]::NewGuid().ToString('N').Substring(0,7).ToUpperInvariant()}
        }
        $path=Join-Path $root ($kind+'.xlsb');$paths+=$path
        $book=New-AuthWorkbook -Excel $excel -Path $path -WarehouseId 'FIXTURE' -StationId 'S1' -CurrentUserIds @('fixture') -CredentialHash $value
        foreach($stage in @('Saved','Reopened')){
            if($stage -eq 'Reopened'){$book.Close($false);$book=$excel.Workbooks.Open($path,0,$true)}
            $actual=$book.Worksheets.Item('Users').ListObjects.Item('tblUsers').ListColumns.Item('PinHash').DataBodyRange.Cells.Item(1,1).Value2
            $rows.Add([pscustomobject]@{Check=$kind+'.'+$stage+'.TextType';Passed=($actual -is [string])})
            $rows.Add([pscustomobject]@{Check=$kind+'.'+$stage+'.Exact';Passed=([string]$actual -ceq $value)})
        }
        $book.Close($false);Release-ComObject $book;$book=$null
        $actual=$null;$value=$null
    }
} finally {
    foreach($open in @($excel.Workbooks)){try{$open.Close($false)}catch{}}
    $excel.Quit();Release-ComObject $excel;$excel=$null
    foreach($path in $paths){
        $resolved=[IO.Path]::GetFullPath($path)
        if(-not $resolved.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Fixture cleanup path escaped its private root.'}
        if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Force}
    }
}
$rows.ToArray()|ConvertTo-Json|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
$failed=@($rows|Where-Object {-not $_.Passed}).Count
Write-Output ($Phase+': '+($rows.Count-$failed)+' passed, '+$failed+' failed; '+$root)
if($failed){exit 1}
