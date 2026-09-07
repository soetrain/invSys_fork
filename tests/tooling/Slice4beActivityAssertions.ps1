# D18 assertions shared with the proven packaged D5 fixture/real form handlers.
# Emit check names and booleans only; activity payloads and fixture credentials
# must never enter console output or the persisted results report.
function Get-Slice4beActivityFiles($Fixture) {
    $root = Join-Path $Fixture.Root ('Training\Activity\' + $Fixture.Warehouse)
    if (Test-Path -LiteralPath $root) {
        Get-ChildItem -LiteralPath $root -Filter '*.json' -File -Recurse |
            Select-Object -ExpandProperty FullName
    }
}

function Get-Slice4beField($Record, [string]$Name) {
    $property = $Record.PSObject.Properties[$Name]
    if ($null -ne $property) { return [string]$property.Value }
    return ''
}

function Test-Slice4beObservedAction {
    param($Fixture, [string[]]$BeforeFiles, [string]$ControlId,
          [string]$RequestedCode, [string]$OutcomeCode, [string]$ExpectedEffect,
          [string]$ExpectedSeverity, [string]$Actor, [string]$CheckPrefix)
    $observations = @()
    $payloads = @()
    $readable = $true
    foreach ($path in @(Get-Slice4beActivityFiles $Fixture)) {
        if ($path -in $BeforeFiles) { continue }
        try {
            $raw = [IO.File]::ReadAllText($path)
            $record = $raw | ConvertFrom-Json -ErrorAction Stop
            if ((Get-Slice4beField $record 'ControlId') -eq $ControlId) {
                $observations += $record
                $payloads += $raw
            }
        } catch { $readable = $false }
    }
    $attempts = @($observations | Where-Object { (Get-Slice4beField $_ 'EventCode') -eq $RequestedCode })
    $outcomes = @($observations | Where-Object { (Get-Slice4beField $_ 'EventCode') -eq $OutcomeCode })
    $pair = $readable -and $attempts.Count -eq 1 -and $outcomes.Count -eq 1
    Check ($CheckPrefix + '.AttemptAndOutcome') $pair

    $correlated = $false
    $owned = $false
    if ($pair) {
        $attemptId = Get-Slice4beField $attempts[0] 'ActivityId'
        $firstId = Get-Slice4beField $attempts[0] 'RecordId'
        $lastId = Get-Slice4beField $outcomes[0] 'RecordId'
        $correlated = $attemptId -ne '' -and $firstId -ne '' -and $lastId -ne '' -and
            $firstId -ne $lastId -and
            $attemptId -eq (Get-Slice4beField $outcomes[0] 'ActivityId')
        $owned = (Get-Slice4beField $attempts[0] 'DataEffect') -eq 'Unknown' -and
            (Get-Slice4beField $outcomes[0] 'DataEffect') -eq $ExpectedEffect -and
            (Get-Slice4beField $outcomes[0] 'Severity') -eq $ExpectedSeverity
        foreach ($record in @($attempts[0], $outcomes[0])) {
            $owned = $owned -and (Get-Slice4beField $record 'OwnerId') -eq 'CORE_CONFIGURATION' -and
                (Get-Slice4beField $record 'WarehouseId') -eq $Fixture.Warehouse -and
                (Get-Slice4beField $record 'StationId') -eq 'S1' -and
                (Get-Slice4beField $record 'UserId') -eq $Actor
        }
    }
    Check ($CheckPrefix + '.StableCorrelation') $correlated
    Check ($CheckPrefix + '.OwnerContextAndEffect') $owned

    $redacted = $pair
    foreach ($raw in $payloads) {
        foreach ($forbidden in @($Fixture.Secret, (CredentialHash $Fixture.Secret), $Fixture.Root,
                                 'PinHash', 'Err.Description', 'mBtnSaveConfig_Click',
                                 'mBtnUomCatalogRetrieve_Click', '"BatchSize"', '"601"')) {
            if ($raw.IndexOf($forbidden, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $redacted = $false }
        }
        if ($raw -match '(?i)[A-Z]:\\|\\\\fixture-host') { $redacted = $false }
    }
    Check ($CheckPrefix + '.RedactedPayload') $redacted
}
