# D18 assertions shared with the proven packaged D5 fixture/real form handlers.
# Emit check names and booleans only; activity payloads and fixture credentials
# must never enter console output or the persisted results report.
function Test-Slice4bePackageIdentity([string]$Deploy) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    foreach ($name in @('Core','Inventory.Domain','Designs.Domain','Operations','Admin')) {
        $archive = $null
        try {
            $archive = [IO.Compression.ZipFile]::OpenRead((Join-Path $Deploy ('invSys.'+$name+'.xlam')))
            $entry = $archive.GetEntry('docProps/custom.xml')
            if ($null -eq $entry) { return $false }
            $reader = [IO.StreamReader]::new($entry.Open())
            try { [xml]$metadata = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $version = @($metadata.Properties.property | Where-Object { $_.name -eq 'invSysPackageSetVersion' })
            $identity = @($metadata.Properties.property | Where-Object { $_.name -eq 'invSysBuildIdentity' })
            if ($version.Count -ne 1 -or $identity.Count -ne 1) { return $false }
            if ($version[0].InnerText -ne 'R1-5' -or $identity[0].InnerText -notmatch '^[0-9a-f]{32}$') { return $false }
        } catch { return $false }
        finally { if ($null -ne $archive) { $archive.Dispose() } }
    }
    return $true
}

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
    $integrity = $pair
    foreach ($raw in $payloads) {
        $match = [regex]::Match($raw, '^(?<body>\{.*),"ContentSha256":"(?<hash>[a-f0-9]{64})"\}$')
        if (-not $match.Success) { $integrity = $false; continue }
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            $digest = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($match.Groups['body'].Value + '}'))).Replace('-','').ToLowerInvariant()
            if ($digest -cne $match.Groups['hash'].Value) { $integrity = $false }
        } finally { $sha.Dispose() }
    }
    Check ($CheckPrefix + '.ContentIntegrity') $integrity
}

function Test-Slice4beUnavailableStore($Fixture) {
    $parent = Join-Path $Fixture.Root 'Training\Activity'
    $leaf = Join-Path $parent $Fixture.Warehouse
    $held = $leaf + '-test-held'
    $expectedRoot = [IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\') + '\'
    foreach ($path in @($leaf,$held)) {
        if (-not [IO.Path]::GetFullPath($path).StartsWith($expectedRoot,[StringComparison]::OrdinalIgnoreCase)) {
            throw 'Activity fixture path escaped its disposable root.'
        }
    }
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    $moved = $false
    try {
        if (Test-Path -LiteralPath $leaf) {
            Move-Item -LiteralPath $leaf -Destination $held
            $moved = $true
        }
        [IO.File]::WriteAllText($leaf, 'blocked fixture path')
        $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','611'))
        Check 'Activity.StoreFailureDoesNotBlockCommand' ($ok -and [long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('BatchSize',0)) -eq 611)
        $status = [string](Run 'invSys.Admin.xlam' 'TestD5Commands.LastStatus')
        Check 'Activity.StoreFailureIsVisible' ($status.Contains('Tracking unavailable'))
    } finally {
        if (Test-Path -LiteralPath $leaf -PathType Leaf) { Remove-Item -LiteralPath $leaf -Force }
        if ($moved) { Move-Item -LiteralPath $held -Destination $leaf }
    }
}
