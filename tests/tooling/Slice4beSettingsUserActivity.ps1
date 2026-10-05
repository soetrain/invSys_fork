# A1 additions retain the original 25-action recording unchanged.
function Test-SettingsUserActivity($Fixture) {
    $cases=@(
        @{Action='SelectUser';Value='config-reader';Id='ADMIN_TRACKING_SELECT_USER';Outcome='SELECTED';Owner='Selection'},
        @{Action='UserRecord';Value='False';Id='ADMIN_TRACKING_USER_RECORD';Outcome='STAGED';Owner='PolicyStaging'}
    )
    foreach($case in $cases){
        $definition=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($case.Id,27))
        $excluded=$true
        foreach($version in 1..26){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($case.Id,$version)) -ceq ''}
        Check ('SettingsActivity.Catalog27.'+$case.Id) ($definition -ne '' -and $excluded)
    }
    OpenSettingsEditor 'Tracking'
    $sequence=StartSettingsRecording
    $ordinal=0
    $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    foreach($case in $cases){
        $ordinal++
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $request=SettingsAct 'Tracking' 'Request'
        [void](SettingsAct 'Tracking' $case.Action $case.Value)
        $after=SettingsAct 'Tracking' 'Request'
        $ownerCorrect=if($case.Action -ceq 'SelectUser'){$after -ceq $request}else{
            $model=$after|ConvertFrom-Json
            @($model.Users|Where-Object {$_.UserId -ceq 'config-reader' -and $_.Record -is [bool] -and -not $_.Record}).Count -eq 1
        }
        Check ('SettingsActivity.'+$case.Id+'.IndependentOwnerResult') ($ownerCorrect -and $config -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        AssertSettingsPair $Fixture $before $case $sequence $ordinal
        $redacted=$true
        foreach($file in @(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before})){
            $raw=[IO.File]::ReadAllText($file)
            $redacted=$redacted -and -not $raw.Contains('config-reader') -and -not $raw.Contains('"Record":')
        }
        Check ('SettingsActivity.'+$case.Id+'.SelectedValuesNotLogged') $redacted
    }
    CloseSettingsEditors
    Check 'SettingsActivity.UserControls.ActualStop' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
    $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
    Check 'SettingsActivity.UserControls.CompleteTwoActionRun' ($closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq 4 -and (JournalChain $sequence 6))
}
