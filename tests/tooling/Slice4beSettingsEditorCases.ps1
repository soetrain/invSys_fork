# D18 specified Settings controls exercised through actual packaged callbacks.
# Each row invokes one real private callback through the separate disposable probe.
function Test-SettingsEditorCatalog {
    $old=@(([string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.CatalogIds' @(10))).Split("`n")|Where-Object {$_ -ne ''})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.CatalogIds' @(11))).Split("`n")|Where-Object {$_ -ne ''})
    $expected=@((SettingsActivityCases).Id)+@('ADMIN_TRACKING_SAVE')
    Check 'SettingsActivity.Catalog.ElevenExtendsTenExactly' ($old.Count -eq 36 -and $new.Count -eq 62 -and
        @($new|Select-Object -Unique).Count -eq 62 -and @($old|Where-Object {$_ -cnotin $new}).Count -eq 0 -and
        @($expected|Where-Object {$_ -cnotin $new}).Count -eq 0)
    foreach($id in $old){
        $prior=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,10))
        Check ('SettingsActivity.Catalog.Preserve.'+$id) ($prior -ne '' -and $prior -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,11)))
    }
    foreach($id in $expected){
        $excluded=$true
        foreach($version in 1..10){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
        Check ('SettingsActivity.Catalog.OlderVersionsExclude.'+$id) $excluded
    }
}

function SettingsActivityCases {
    @(
        @{Section='Tracking';Action='SelectControl';Value='ADMIN_UOM_ADD';Id='ADMIN_TRACKING_SELECT_CONTROL';Outcome='SELECTED';Owner='Selection'},
        @{Section='Tracking';Action='Capture';Value='False';Id='ADMIN_TRACKING_CAPTURE';Outcome='STAGED';Owner='PolicyStaging'},
        @{Section='Tracking';Action='AdminVisible';Value='False';Id='ADMIN_TRACKING_ADMIN_VISIBLE';Outcome='STAGED';Owner='PolicyStaging'},
        @{Section='Tracking';Action='DefaultView';Value='Compare both';Id='ADMIN_TRACKING_DEFAULT_VIEW';Outcome='STAGED';Owner='PolicyStaging'},
        @{Section='Tracking';Action='Collect';Value='False';Id='ADMIN_TRACKING_COLLECT';Outcome='STAGED';Owner='PolicyStaging'},
        @{Section='Tracking';Action='Visible';Value='False';Id='ADMIN_TRACKING_VISIBLE';Outcome='STAGED';Owner='PolicyStaging'},
        @{Section='Tracking';Action='Sequence';Value='False';Id='ADMIN_TRACKING_SEQUENCE';Outcome='STAGED';Owner='PolicyStaging'},
        @{Section='Tracking';Action='Reset';Value='';Id='ADMIN_TRACKING_RESET';Outcome='STAGED';Owner='PolicyDefaults'},
        @{Section='Tracking';Action='Reload';Value='';Id='ADMIN_TRACKING_RELOAD';Outcome='REFRESHED';Owner='PolicyRead'},
        @{Section='Detail';Action='SelectFamily';Value='Admin';Id='ADMIN_DETAIL_SELECT_FAMILY';Outcome='SELECTED';Owner='Selection'},
        @{Section='Detail';Action='SelectField';Value='Reference';Id='ADMIN_DETAIL_SELECT_FIELD';Outcome='SELECTED';Owner='Selection'},
        @{Section='Detail';Action='ShowField';Value='False';Id='ADMIN_DETAIL_SHOW_FIELD';Outcome='STAGED';Owner='ProfileStaging'},
        @{Section='Detail';Action='MoveUp';Value='';Id='ADMIN_DETAIL_MOVE_UP';Outcome='STAGED';Owner='ProfileStaging'},
        @{Section='Detail';Action='MoveDown';Value='';Id='ADMIN_DETAIL_MOVE_DOWN';Outcome='STAGED';Owner='ProfileStaging'},
        @{Section='Detail';Action='Save';Value='';Id='ADMIN_DETAIL_SAVE';Outcome='COMPLETED';Owner='ProfileAppend'},
        @{Section='Detail';Action='Reset';Value='';Id='ADMIN_DETAIL_RESET';Outcome='STAGED';Owner='ProfileDefaults'},
        @{Section='Detail';Action='Reload';Value='';Id='ADMIN_DETAIL_RELOAD';Outcome='REFRESHED';Owner='ProfileRead'},
        @{Section='Preference';Action='Select';Value='Diagnostic';Id='ADMIN_PATH_PREFERENCE_SELECT';Outcome='STAGED';Owner='PreferenceStaging'},
        @{Section='Preference';Action='Save';Value='';Id='ADMIN_PATH_PREFERENCE_SAVE';Outcome='COMPLETED';Owner='PreferenceSave'},
        @{Section='Preference';Action='Reset';Value='';Id='ADMIN_PATH_PREFERENCE_RESET';Outcome='STAGED';Owner='PreferenceDefaults'},
        @{Section='Preference';Action='Reload';Value='';Id='ADMIN_PATH_PREFERENCE_RELOAD';Outcome='REFRESHED';Owner='PreferenceRead'},
        @{Section='Operations';Action='Select';Value='Compare both';Id='VIEWER_PATH_PREFERENCE_SELECT';Outcome='STAGED';Owner='PreferenceStaging'},
        @{Section='Operations';Action='Save';Value='';Id='VIEWER_PATH_PREFERENCE_SAVE';Outcome='COMPLETED';Owner='PreferenceSave'},
        @{Section='Operations';Action='Reset';Value='';Id='VIEWER_PATH_PREFERENCE_RESET';Outcome='STAGED';Owner='PreferenceDefaults'},
        @{Section='Operations';Action='Reload';Value='';Id='VIEWER_PATH_PREFERENCE_RELOAD';Outcome='REFRESHED';Owner='PreferenceRead'}
    )
    # ADMIN_TRACKING_SAVE is exercised last in a separate recording: the owner
    # appends a policy version and interrupts capture. Only REQUESTED survives.
}

function AssertSettingsPair($Fixture,[string[]]$Before,$Case,[string]$Sequence,[int]$Ordinal,[string]$Label='') {
    if($Label -eq ''){$Label=$Case.Id}
    $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before}|ForEach-Object {
        [IO.File]::ReadAllText($_)|ConvertFrom-Json
    })
    $attempts=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
    $outcomes=@($records|Where-Object OutcomeCode -CEQ $Case.Outcome)
    $paired=$records.Count -eq 2 -and $attempts.Count -eq 1 -and $outcomes.Count -eq 1
    if($paired){
        $paired=$attempts[0].ControlId -ceq $Case.Id -and $outcomes[0].ControlId -ceq $Case.Id -and
            $attempts[0].ActivityId -ceq $outcomes[0].ActivityId -and $attempts[0].RecordId -cne $outcomes[0].RecordId -and
            $attempts[0].SequenceId -ceq $Sequence -and $outcomes[0].SequenceId -ceq $Sequence -and
            $attempts[0].Ordinal -eq $Ordinal -and $outcomes[0].Ordinal -eq $Ordinal
    }
    Check ('SettingsActivity.'+$Label+'.ExactlyOneActualHandlerPair') $paired
    $effect=if($Case.Outcome -ceq 'COMPLETED'){'Changed'}elseif($Case.Outcome -ceq 'FAILED' -and $Case.Owner -cin @('ProfileAppend','PreferenceSave','PolicySave')){'Unknown'}else{'Unchanged'}
    $facts=$paired
    if($facts){
        $facts=$attempts[0].DataEffect -ceq 'Unknown' -and $outcomes[0].DataEffect -ceq $effect -and
            @($attempts[0].SourceEventRefs).Count -eq 0 -and @($outcomes[0].SourceEventRefs).Count -eq 0 -and
            $outcomes[0].CatalogVersion -eq 12 -and $outcomes[0].EventCode -ceq ($Case.Id+'_'+$Case.Outcome)
    }
    Check ('SettingsActivity.'+$Label+'.ExplicitOwnerFacts') $facts
    $metadata=$paired
    $integrity=$paired
    $owner=if($Case.Id.Contains('_PATH_PREFERENCE_')){'CORE_PERSONAL_PREFERENCE'}elseif($Case.Id.EndsWith('_SAVE')){'CORE_CONFIGURATION'}else{'ADMIN_SETTINGS_UI'}
    $role=if($Case.Id.StartsWith('VIEWER_')){'Viewer'}else{'Admin'}
    $surface=if($role -ceq 'Viewer'){'Operations > Viewer > Settings > Event Tracking'}elseif($Case.Id.StartsWith('ADMIN_TRACKING_')){'Admin > Settings > Event Tracking > Tracking'}elseif($Case.Id.StartsWith('ADMIN_DETAIL_')){'Admin > Settings > Event Tracking > Event Detail'}else{'Admin > Settings > Event Tracking > Action Paths'}
    $actor=[string](Run 'invSys.Core.xlam' 'modAuth.GetCurrentUserId')
    foreach($record in $records){
        $vocabulary=([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($Case.Id,$record.OutcomeCode)))|ConvertFrom-Json
        $metadata=$metadata -and $record.OwnerId -ceq $owner -and $record.SourceRole -ceq $role -and
            $record.Surface -ceq $surface -and $record.UserId -ceq $actor -and $record.WarehouseId -ceq $Fixture.Warehouse -and
            $record.UserMessage -ceq $vocabulary.UserMessage -and $record.NextStep -ceq $vocabulary.NextStep
    }
    foreach($file in @(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before})){
        $raw=[IO.File]::ReadAllText($file)
        $record=$raw|ConvertFrom-Json
        $integrity=$integrity -and @($record.PSObject.Properties).Count -eq 27
        foreach($forbidden in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'PinHash','Err.Description','mRequest')){
            if($raw.IndexOf($forbidden,[StringComparison]::OrdinalIgnoreCase) -ge 0){$integrity=$false}
        }
        $match=[regex]::Match($raw,'^(?<body>\{.*),"ContentSha256":"(?<hash>[a-f0-9]{64})"\}$')
        if(-not $match.Success){$integrity=$false;continue}
        $sha=[Security.Cryptography.SHA256]::Create()
        try{
            $digest=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($match.Groups['body'].Value+'}')))).Replace('-','').ToLowerInvariant()
            $integrity=$integrity -and $digest -ceq $match.Groups['hash'].Value
        }finally{$sha.Dispose()}
    }
    Check ('SettingsActivity.'+$Label+'.ExactOriginAndFixedVocabulary') $metadata
    Check ('SettingsActivity.'+$Label+'.RedactedImmutablePayload') $integrity
}
