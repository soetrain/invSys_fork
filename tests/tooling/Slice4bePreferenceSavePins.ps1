# The explicit preference save may append exactly its correlated activity pair.
# Config and every preceding file remain immutable; subsequent reads stay strict.
function Update-Slice4bePreferenceSavePins([hashtable]$Before,[hashtable]$After,$Fixture) {
    if($After.Count -ne $Before.Count+2){return $false}
    foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $Before[$path] -cne $After[$path]){return $false}}
    $added=@($After.Keys|Where-Object {-not $Before.ContainsKey($_)})
    if($added.Count -ne 2){return $false}
    $root=[IO.Path]::GetFullPath((Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)))
    $records=@()
    try {
        foreach($path in $added){
            if([IO.Path]::GetDirectoryName($path) -cne $root -or [IO.Path]::GetExtension($path) -cne '.json'){return $false}
            if((Get-FileHash -LiteralPath $path).Hash -cne $After[$path]){return $false}
            $raw=[IO.File]::ReadAllText($path);$record=$raw|ConvertFrom-Json -ErrorAction Stop
            if($record.SchemaVersion -ne 1 -or $record.CatalogVersion -lt 11 -or $record.PolicyVersion -ne 1 -or
                $record.ControlId -cne 'VIEWER_PATH_PREFERENCE_SAVE' -or $record.OwnerId -cne 'CORE_PERSONAL_PREFERENCE' -or
                $record.SourceKind -cne 'User activity' -or $record.SourceRole -cne 'Viewer' -or
                $record.WarehouseId -cne $Fixture.Warehouse -or $record.StationId -cne 'S1' -or $record.UserId -cne 'config-reader' -or
                $record.SequenceId -cne '' -or $record.Ordinal -ne 0 -or @($record.SourceEventRefs).Count -ne 0){return $false}
            $id=[guid]::Empty;$activity=[guid]::Empty
            if(-not [guid]::TryParse([string]$record.RecordId,[ref]$id) -or $id -eq [guid]::Empty -or
                -not [guid]::TryParse([string]$record.ActivityId,[ref]$activity) -or $activity -eq [guid]::Empty -or
                [IO.Path]::GetFileNameWithoutExtension($path) -cne $record.RecordId){return $false}
            if($record.OutcomeCode -cnotin @('REQUESTED','COMPLETED') -or
                $record.EventCode -cne ('VIEWER_PATH_PREFERENCE_SAVE_'+$record.OutcomeCode)){return $false}
            $effect=if($record.OutcomeCode -ceq 'REQUESTED'){'Unknown'}else{'Changed'}
            if($record.DataEffect -cne $effect){return $false}
            $match=[regex]::Match($raw,'^(?<body>\{.*),"ContentSha256":"(?<hash>[a-f0-9]{64})"\}$')
            if(-not $match.Success){return $false}
            $sha=[Security.Cryptography.SHA256]::Create()
            try {$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($match.Groups['body'].Value+'}')))).Replace('-','').ToLowerInvariant()}
            finally {$sha.Dispose()}
            if($hash -cne $match.Groups['hash'].Value){return $false}
            $records+=$record
        }
        if($records[0].ActivityId -cne $records[1].ActivityId -or $records[0].RecordId -ceq $records[1].RecordId -or
            $records[0].CatalogVersion -ne $records[1].CatalogVersion -or $records[0].OutcomeCode -ceq $records[1].OutcomeCode){return $false}
    } catch {return $false}
    foreach($path in $added){$Before[$path]=$After[$path]}
    return $true
}
