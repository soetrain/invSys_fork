[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

# Tooling regression only: JSON strings must remain strings in both supported
# execution hosts (Windows PowerShell 5.1 and PowerShell 7.5+).
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$root = Join-Path $repo ('reports/runtime/json-literal-types/' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $root -Force)
$json = Join-Path $root 'literal.json'
$schema = Join-Path $root 'schema.json'
$literal = '{"utc":"2026-11-03T19:52:08Z","offset":"2026-11-03T12:52:08.1234567-07:00","nested":["2026-11-03T19:52:08Z"],"flag":true,"count":4,"empty":null}'
[IO.File]::WriteAllText($json, $literal)
[IO.File]::WriteAllText($schema, @'
{"type":"object","required":["utc"],"properties":{"utc":{"type":"string","format":"date-time","const":"2026-11-03T19:52:08Z"}}}
'@)
$script:passed = 0
$script:failed = 0
function Check([string]$Name, [bool]$Passed) {
    if ($Passed) { $script:passed++; Write-Output ('PASS ' + $Name) }
    else { $script:failed++; Write-Output ('FAIL ' + $Name) }
}

foreach ($tool in @('inventory-vba-surface.ps1', 'validate-json-contract.ps1')) {
    # Exercise the real reader without executing the full scanner or duplicating
    # its implementation in the test.
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile(
        (Join-Path $repo ('tools/' + $tool)), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw 'Tool could not be parsed; not a behavioral test result.' }
    $readers = @($ast.FindAll({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Read-Json'
    }, $true))
    if ($readers.Count -ne 1) { throw 'Expected exactly one production JSON reader.' }
    $body = $readers[0].Body.Extent.Text
    $reader = [scriptblock]::Create($body.Substring(1, $body.Length - 2))
    $value = & $reader $json
    Check ($tool + '.UtcLiteral') ($value.utc -is [string] -and $value.utc -ceq '2026-11-03T19:52:08Z')
    Check ($tool + '.OffsetLiteral') ($value.offset -is [string] -and $value.offset -ceq '2026-11-03T12:52:08.1234567-07:00')
    Check ($tool + '.NestedLiteral') ($value.nested[0] -is [string] -and $value.nested[0] -ceq '2026-11-03T19:52:08Z')
    Check ($tool + '.OtherTypes') ($value.flag -is [bool] -and $value.flag -and $value.count -eq 4 -and $null -eq $value.empty)
}

$validator = Join-Path $repo 'tools/validate-json-contract.ps1'
$accepted = $false
try { & $validator -JsonPath $json -SchemaPath $schema; $accepted = $true } catch {}
Check 'Validator.ValidTimestampAccepted' $accepted

foreach ($invalid in @('{"utc":4}', '{"utc":"not-a-date"}', '{"utc":"2026-11-03T19:52:09Z"}')) {
    [IO.File]::WriteAllText($json, $invalid)
    $rejected = $false
    try { & $validator -JsonPath $json -SchemaPath $schema } catch { $rejected = $true }
    Check ('Validator.Rejects.' + $invalid) $rejected
}
Write-Output ('JSON_LITERAL_TYPES passed={0} failed={1} total={2}' -f $passed, $failed, ($passed + $failed))
if ($failed) { exit 1 }
