# Real Save callbacks; failed writes are observed through Excel BeforeSave.
# No owner outcome is derived from status text or a Boolean return.
function Test-SettingsEditorSaveCases($Fixture,$Evidence=$null) {
    $sequence='';$ordinal=0
    if($null -ne $Evidence){
        CloseSettingsEditors
        CloseRecordingViewer
        OpenRecordingViewer
        $sequence=StartSettingsRecording
    }
    foreach($section in @('Preference','Operations')){
        CloseSettingsEditors
        OpenSettingsEditor $section
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $choice=[string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.PersonalChoice')
        [void](SettingsAct $section 'Save')
        $id=if($section -ceq 'Preference'){'ADMIN_PATH_PREFERENCE_SAVE'}else{'VIEWER_PATH_PREFERENCE_SAVE'}
        Check ('SettingsActivity.Unchanged.'+$section+'.IndependentOwnerResult') (
            [string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.PersonalChoice') -ceq $choice -and
            (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config)
        if($null -ne $Evidence){$ordinal++}
        AssertSettingsPair $Fixture $before @{Id=$id;Outcome='UNCHANGED';Owner='PreferenceSave'} $sequence $ordinal ('Unchanged.'+$section)
    }
    if($null -ne $Evidence){Save-SettingsDiagnosticRun $Evidence 'Unchanged' $sequence 2}
    foreach($section in @('Tracking','Detail')){
        foreach($failure in @('StaleVersion','SaveCancelled')){
            CloseSettingsEditors
            if($null -ne $Evidence){CloseRecordingViewer;OpenRecordingViewer}
            OpenSettingsEditor $section
            if($failure -ceq 'StaleVersion'){[void](SettingsAct $section 'StaleVersionFixture')}
            $sequence='';$ordinal=0
            if($null -ne $Evidence){$sequence=StartSettingsRecording;$ordinal=1}
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $ownerBefore=SettingsOwnerState
            $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            $staged=SettingsAct $section 'Request'
            $cancelCount=0
            $eventsEnabled=$excel.EnableEvents
            if($failure -ceq 'SaveCancelled'){[void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCancelSave' @($Fixture.Config))}
            try {
                if($failure -ceq 'SaveCancelled'){$excel.EnableEvents=$true}
                [void](SettingsAct $section 'Save')
                if($failure -ceq 'SaveCancelled'){$cancelCount=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCancelledCount')}
            }finally{
                $excel.EnableEvents=$eventsEnabled
                if($failure -ceq 'SaveCancelled'){[void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyStopCancelling')}
            }
            $ownerAfter=SettingsOwnerState
            $label=$failure+'.'+$section
            if($failure -ceq 'SaveCancelled'){Check ('SettingsActivity.'+$label+'.ExcelCancelledExactlyOnce') ($cancelCount -eq 1)}
            $valid=(Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config -and
                $ownerBefore.ProfileVersion -eq $ownerAfter.ProfileVersion -and $ownerBefore.PolicyVersion -eq $ownerAfter.PolicyVersion -and
                (SettingsAct $section 'Request') -ceq $staged -and ($failure -cne 'SaveCancelled' -or $cancelCount -eq 1)
            Check ('SettingsActivity.'+$label+'.IndependentOwnerResult') $valid
            if(-not $valid){throw 'Save rejection/cancellation fixture did not preserve owner state; not observation RED.'}
            $id=if($section -ceq 'Tracking'){'ADMIN_TRACKING_SAVE'}else{'ADMIN_DETAIL_SAVE'}
            $outcome=if($failure -ceq 'SaveCancelled'){'FAILED'}else{'REJECTED'}
            $owner=if($section -ceq 'Tracking'){'PolicySave'}else{'ProfileAppend'}
            AssertSettingsPair $Fixture $before @{Id=$id;Outcome=$outcome;Owner=$owner} $sequence $ordinal $label
            if($null -ne $Evidence){Save-SettingsDiagnosticRun $Evidence $label $sequence 1}
        }
    }
    # Personal preference is available to a signed-in Operations user without
    # ADMIN_MAINT. Exercise both changed and unchanged actual Save callbacks.
    CloseSettingsEditors
    CloseRecordingViewer
    SelectTarget $Fixture 'config-reader'
    OpenRecordingViewer
    OpenSettingsEditor 'Operations'
    $current=[string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.PersonalChoice')
    $choice=if($current -ceq 'Diagnostic'){'How-To'}else{'Diagnostic'}
    [void](SettingsAct 'Operations' 'Select' $choice)
    foreach($outcome in @('COMPLETED','UNCHANGED')){
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](SettingsAct 'Operations' 'Save')
        $valid=[string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.PersonalChoice') -ceq $choice -and
            (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config
        Check ('SettingsActivity.PersonalNonAdmin.'+$outcome+'.IndependentOwnerResult') $valid
        if(-not $valid){throw 'Personal preference owner failed for existing signed-in Operations role.'}
        AssertSettingsPair $Fixture $before @{Id='VIEWER_PATH_PREFERENCE_SAVE';Outcome=$outcome;Owner='PreferenceSave'} '' 0 ('PersonalNonAdmin.'+$outcome)
    }
    CloseSettingsEditors
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
}
