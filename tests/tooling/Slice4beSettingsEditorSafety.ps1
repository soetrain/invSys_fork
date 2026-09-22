# Supplemental actual-handler cases; changes affect disposable generated fixtures only.
function Test-SettingsEditorSafety($Fixture,$Evidence=$null) {
    $initial=ActivityPins
    Test-SettingsFailedProfileRead $Fixture $Evidence
    Test-SettingsSaveDenial $Fixture $Evidence
    Test-SettingsUnavailableTracking $Fixture
    Test-SettingsOlderPolicy $Fixture
    Check 'SettingsActivity.Safety.OriginalActivityBytesPreserved' (PinsRetained $initial)
}

function Test-SettingsFailedProfileRead($Fixture,$Evidence=$null) {
    $bytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $book=$null
    try {
        CloseSettingsEditors
        CloseRecordingViewer
        OpenRecordingViewer
        OpenSettingsEditor 'Detail'
        $sequence=StartSettingsRecording
        $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $table=Table $book 'tblEventDetailProfiles'
        $table.ListRows.Item($table.ListRows.Count).Range.Cells.Item(1,$table.ListColumns.Item('SchemaVersion').Index).Value2=99.0
        $book.Save();$book.Close($false);$book=$null
        $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $unreadable=-not [bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.ProfileReadable')
        Check 'SettingsActivity.FailedProfileRead.OwnerRejectsUnsupportedSchema' $unreadable
        if(-not $unreadable){throw 'Invalid profile fixture remained readable; not callback RED.'}
        $before=@(Get-Slice4beActivityFiles $Fixture)
        [void](SettingsAct 'Detail' 'Reload')
        Check 'SettingsActivity.FailedProfileRead.DoesNotRepairConfig' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config)
        AssertSettingsPair $Fixture $before @{Id='ADMIN_DETAIL_RELOAD';Outcome='FAILED';Owner='ProfileRead'} $sequence 1 'FailedProfileRead'
        $capturePins=ActivityPins
        CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'settings-failed-profile-read.png'
        Check 'SettingsActivity.FailedProfileRead.Visible' $true
        Check 'SettingsActivity.FailedProfileRead.CaptureDoesNotInventAction' ((ActivityPins).Count -eq $capturePins.Count -and (PinsRetained $capturePins))
        Check 'SettingsActivity.FailedProfileRead.ActualStop' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        Check 'SettingsActivity.FailedProfileRead.RenderingDoesNotInventSelection' ($closed.Count -eq 1 -and
            $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 1 -and @($closed[0].Observations).Count -eq 2)
        if($null -ne $Evidence){$Evidence.FailedProfileRead=$closed[0]}
    }finally{
        if($null -ne $book){$book.Close($false)}
        CloseSettingsEditors
        CloseRecordingViewer
        [IO.File]::WriteAllBytes($Fixture.Config,$bytes)
    }
    Check 'SettingsActivity.FailedProfileRead.ConfigRestored' ([bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.ProfileReadable'))
}

function Test-SettingsSaveDenial($Fixture,$Evidence=$null) {
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $bytes=[IO.File]::ReadAllBytes($authPath)
    $book=$null
    try {
        OpenRecordingViewer
        OpenSettingsEditor 'Tracking'
        $sequence=StartSettingsRecording
        $ordinal=0
        $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
        $book=$excel.Workbooks.Open($authPath,0,$false)
        $caps=Table $book 'tblCapabilities'
        $changed=0
        foreach($row in $caps.ListRows){
            if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and
               $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){
                $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive'
                $changed++
            }
        }
        $book.Save();$book.Close($false);$book=$null
        [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
        $denied=-not [bool](Run 'invSys.Core.xlam' 'modRoleUiAccess.CanCurrentUserPerformCapabilityCached' @('ADMIN_MAINT'))
        $sameContext=$context -cne '' -and [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context
        Check 'SettingsActivity.SaveDenial.SameSessionLostOnlyCapability' ($changed -gt 0 -and $denied -and $sameContext)
        if($changed -eq 0 -or -not $denied -or -not $sameContext){throw 'Capability fixture did not retain captured context; not outcome RED.'}
        foreach($section in @('Tracking','Detail')){
            $ordinal++
            OpenSettingsEditor $section
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            $staged=SettingsAct $section 'Request'
            [void](SettingsAct $section 'Save')
            Check ('SettingsActivity.SaveDenial.'+$section+'.OwnerPreservesSavedAndStagedState') (
                (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config -and (SettingsAct $section 'Request') -ceq $staged)
            $id=if($section -ceq 'Tracking'){'ADMIN_TRACKING_SAVE'}else{'ADMIN_DETAIL_SAVE'}
            AssertSettingsPair $Fixture $before @{Id=$id;Outcome='DENIED';Owner='DeniedSave'} $sequence $ordinal ('SaveDenial.'+$section)
        }
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](SettingsAct 'Detail' 'Reload')
        Check 'SettingsActivity.SaveDenial.Reload.DoesNotWriteConfig' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config)
        AssertSettingsPair $Fixture $before @{Id='ADMIN_DETAIL_RELOAD';Outcome='FAILED';Owner='ProfileRead'} $sequence 3 'SaveDenial.Reload'
        $capturePins=ActivityPins
        CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'settings-denied-reload.png'
        Check 'SettingsActivity.SaveDenial.Reload.Visible' $true
        Check 'SettingsActivity.SaveDenial.Reload.CaptureDoesNotInventAction' ((ActivityPins).Count -eq $capturePins.Count -and (PinsRetained $capturePins))
        Check 'SettingsActivity.SaveDenial.ActualStop' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        Check 'SettingsActivity.SaveDenial.ReloadDoesNotInventSelection' ($closed.Count -eq 1 -and
            $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 3 -and @($closed[0].Observations).Count -eq 6)
        if($null -ne $Evidence){$Evidence.SaveDenial=$closed[0]}
    }finally{
        if($null -ne $book){$book.Close($false)}
        CloseSettingsEditors
        CloseRecordingViewer
        [IO.File]::WriteAllBytes($authPath,$bytes)
        [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
    }
    Check 'SettingsActivity.SaveDenial.CapabilityRestored' ([bool](Run 'invSys.Core.xlam' 'modRoleUiAccess.CanCurrentUserPerformCapabilityCached' @('ADMIN_MAINT')))
}

function Test-SettingsUnavailableTracking($Fixture) {
    $leaf=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
    $held=$leaf+'-settings-test-held'
    $expected=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    foreach($path in @($leaf,$held)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($expected,[StringComparison]::OrdinalIgnoreCase)){throw 'Fixture activity path escaped generated root.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Preserve earlier activity backup.'}
    $moved=$false
    try {
        OpenRecordingViewer
        OpenSettingsEditor 'Operations'
        $current=[string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.PersonalChoice')
        $choice=if($current -ceq 'How-To'){'Diagnostic'}else{'How-To'}
        [void](SettingsAct 'Operations' 'Select' $choice)
        $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        Move-Item -LiteralPath $leaf -Destination $held
        $moved=$true
        [IO.File]::WriteAllText($leaf,'blocked fixture path')
        $status=SettingsAct 'Operations' 'Save'
        Check 'SettingsActivity.StoreFailure.OwnerSaveStillSucceeds' (
            [string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.PersonalChoice') -ceq $choice -and
            (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config)
        Check 'SettingsActivity.StoreFailure.SeparateOwnerAndTrackingStatus' ($status.Contains('saved') -and $status.Contains('Tracking unavailable'))
        CaptureOwnedFormByCaptionEvidence 'Event Tracking Settings' 'settings-tracking-unavailable.png'
        Check 'SettingsActivity.StoreFailure.Visible' $true
    }finally{
        if(Test-Path -LiteralPath $leaf -PathType Leaf){Remove-Item -LiteralPath $leaf -Force}
        if($moved){Move-Item -LiteralPath $held -Destination $leaf}
        CloseSettingsEditors
        CloseRecordingViewer
    }
}

function Test-SettingsOlderPolicy($Fixture) {
    $bytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $book=$null
    try {
        $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $headers=Table $book 'tblEventTrackingPolicies'
        $controls=Table $book 'tblEventTrackingControls'
        $header=$headers.ListRows.Item($headers.ListRows.Count)
        $version=[int]$header.Range.Cells.Item(1,$headers.ListColumns.Item('PolicyVersion').Index).Value2
        $header.Range.Cells.Item(1,$headers.ListColumns.Item('CatalogVersion').Index).Value2=10.0
        $newIds=@((SettingsActivityCases).Id)+@('ADMIN_TRACKING_SAVE')
        $removed=0
        for($index=$controls.ListRows.Count;$index -ge 1;$index--){
            $row=$controls.ListRows.Item($index)
            if([int]$row.Range.Cells.Item(1,$controls.ListColumns.Item('PolicyVersion').Index).Value2 -eq $version -and
                [string]$row.Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -cin $newIds){$row.Delete();$removed++}
        }
        $book.Save();$book.Close($false);$book=$null
        $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        Check 'SettingsActivity.OlderPolicy.ExactCatalogTenFixture' ($removed -eq 26 -and
            [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('ADMIN_SETTINGS_SAVE_VALUE')) -ceq ('True|True|True|'+$version))
        foreach($id in $newIds){Check ('SettingsActivity.OlderPolicy.Excludes.'+$id) (
            [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @($id)) -ceq 'False|False|False|0')}
        OpenRecordingViewer
        OpenSettingsEditor 'Operations'
        $before=ActivityPins
        $current=[string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.PersonalChoice')
        $choice=if($current -ceq 'Diagnostic'){'How-To'}else{'Diagnostic'}
        [void](SettingsAct 'Operations' 'Select' $choice)
        [void](SettingsAct 'Operations' 'Save')
        Check 'SettingsActivity.OlderPolicy.ActualOwnerSaveWithoutRetroactiveCollection' (
            [string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.PersonalChoice') -ceq $choice -and
            (ActivityPins).Count -eq $before.Count -and (PinsRetained $before) -and
            (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config)
    }finally{
        if($null -ne $book){$book.Close($false)}
        CloseSettingsEditors
        CloseRecordingViewer
        [IO.File]::WriteAllBytes($Fixture.Config,$bytes)
    }
}
