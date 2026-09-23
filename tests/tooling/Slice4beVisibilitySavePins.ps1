# Reader tests deliberately save visibility policy between read-only intervals.
# Permit that exact save's Config update and one request observation only.
function Test-Slice4beReadPins([hashtable]$Pins,[string]$Root) {
    $files=@(Get-ChildItem -LiteralPath $Root -Recurse -File)
    if($files.Count -ne $Pins.Count){return $false}
    foreach($file in $files){
        if(-not $Pins.ContainsKey($file.FullName) -or
            (Get-FileHash -LiteralPath $file.FullName).Hash -cne $Pins[$file.FullName]){return $false}
    }
    return $true
}

function Update-Slice4beVisibilitySavePins([hashtable]$Pins,[string]$Root,$Fixture) {
    $files=@(Get-ChildItem -LiteralPath $Root -Recurse -File)
    $added=@($files | Where-Object {-not $Pins.ContainsKey($_.FullName)})
    if($files.Count -ne $Pins.Count+1 -or $added.Count -ne 1){return $false}
    foreach($path in $Pins.Keys){
        if(-not (Test-Path -LiteralPath $path -PathType Leaf)){return $false}
        if($path -cne $Fixture.Config -and
            (Get-FileHash -LiteralPath $path).Hash -cne $Pins[$path]){return $false}
    }
    $activityRoot=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
    if($added[0].DirectoryName -cne [IO.Path]::GetFullPath($activityRoot)){return $false}
    try {
        $raw=[IO.File]::ReadAllText($added[0].FullName)
        $record=$raw|ConvertFrom-Json -ErrorAction Stop
        if($record.SchemaVersion -ne 1 -or $record.ControlId -cne 'ADMIN_TRACKING_SAVE' -or
            $record.OutcomeCode -cne 'REQUESTED' -or $record.EventCode -cne 'ADMIN_TRACKING_SAVE_REQUESTED' -or
            $record.OwnerId -cne 'CORE_CONFIGURATION' -or $record.SourceRole -cne 'Admin' -or
            $record.WarehouseId -cne $Fixture.Warehouse -or $record.StationId -cne 'S1' -or
            $record.UserId -cne 'config-admin' -or $record.SequenceId -cne '' -or $record.Ordinal -ne 0 -or
            $record.DataEffect -cne 'Unknown' -or @($record.SourceEventRefs).Count -ne 0){return $false}
        $recordId=[guid]::Empty;$activityId=[guid]::Empty
        if(-not [guid]::TryParse([string]$record.RecordId,[ref]$recordId) -or $recordId -eq [guid]::Empty -or
            -not [guid]::TryParse([string]$record.ActivityId,[ref]$activityId) -or $activityId -eq [guid]::Empty -or
            $added[0].BaseName -cne [string]$record.RecordId){return $false}
        $match=[regex]::Match($raw,'^(?<body>\{.*),"ContentSha256":"(?<hash>[a-f0-9]{64})"\}$')
        if(-not $match.Success){return $false}
        $sha=[Security.Cryptography.SHA256]::Create()
        try {$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($match.Groups['body'].Value+'}')))).Replace('-','').ToLowerInvariant()}
        finally {$sha.Dispose()}
        if($hash -cne $match.Groups['hash'].Value){return $false}
    } catch {return $false}
    # Advance only these named writes after validating the entire difference.
    if($Pins.ContainsKey($Fixture.Config)){$Pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash}
    $Pins[$added[0].FullName]=(Get-FileHash -LiteralPath $added[0].FullName).Hash
    return $true
}
