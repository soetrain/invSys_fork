# D18 direct curation: original unrecorded Admin actions, real publication,
# actual selection/Create/Save handlers. No synthetic guide or recording source.
function Install-GuideDeferredCloseProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $probe=$project.VBComponents.Add(1)
    $probe.Name='TestGuideDeferredTiming'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mHolding As Boolean, mRequests As Long
Public Stage As String
Public Sub Arm()
    mRequests = 0: Stage = "entry": mHolding = True
End Sub
Public Function HoldCallback() As Boolean
    If mHolding Then mRequests = mRequests + 1
    HoldCallback = mHolding
End Function
Public Function Deliver() As String
    On Error GoTo Failed
    mHolding = False
    modGuideReaderClose.Flush
    Deliver = "OK|" & CStr(mRequests)
    Exit Function
Failed:
    Deliver = CStr(Err.Number) & "|" & Stage & "|" & CStr(mRequests)
End Function
'@)
    $module=$project.VBComponents.Item('modGuideReaderClose').CodeModule
    $body=$module.Lines(1,$module.CountOfLines)
    $onTime='        Application.OnTime EarliestTime:='
    if([regex]::Matches($body,[regex]::Escape($onTime)).Count -ne 1){throw 'Native close timer fixture unavailable.'}
    $body=$body.Replace($onTime,'        If Not TestGuideDeferredTiming.HoldCallback() Then Application.OnTime EarliestTime:=')
    # Attribute a former implementation's failure without requiring its loop in a repaired build.
    foreach($point in @(@('        For Each picker In pickers','enumerating queued pickers'),@('            If modOperationsFormLifetime.IsLoaded(picker) Then Unload picker','checking or unloading queued picker'))){
        $body=$body.Replace($point[0],('        TestGuideDeferredTiming.Stage = "'+$point[1]+'"'+[Environment]::NewLine+$point[0]))
    }
    $module.DeleteLines(1,$module.CountOfLines)
    $module.AddFromString($body)
}

function Test-GuideActionCuration($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    function Pick([string]$Name,[string]$Action,[string]$Value='') {BoundControl $Name $Action $Value 'frmGuideActionPicker'}
    function Draft([string]$Name,[string]$Action,[string]$Value='') {BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function OpenPicker {[void](BoundControl 'btnChooseGuideActions' 'Click' '' 'frmActionPaths')}
    function Test-ReopenedPickerNativeClose {
        . (Join-Path $PSScriptRoot 'Slice4beGuideInputNative.ps1')
        OpenPicker
        $pickerWindow=Get-OwnedGuideInputWindow 'Choose tracked actions'
        $elapsed=[Diagnostics.Stopwatch]::StartNew()
        if(-not [GuideInputNative]::PostMessage($pickerWindow,0x0010,[IntPtr]::Zero,[IntPtr]::Zero)){
            throw 'Native Close message could not be delivered to the reopened picker.'
        }
        do{
            Start-Sleep -Milliseconds 100
            $windowAlive=[GuideInputNative]::IsWindow($pickerWindow)
            $pickerGone=-not $windowAlive -and (Pick '' 'Count') -ceq '0'
        }until($pickerGone -or $elapsed.Elapsed.TotalSeconds -ge 10)
        $elapsed.Stop()
        Check 'GuideCuration.ReopenedPickerNativeXDisposesWindowAndForm' $pickerGone
        Write-Host ('Native picker close observation: '+$elapsed.ElapsedMilliseconds+' ms')
        if(-not $pickerGone){
            $visible=[InvSysSettingsCapture]::OwnedVisibleForm('Choose tracked actions',[IntPtr]$excel.Hwnd)
            $loaded=Pick '' 'Count'
            throw ("Native picker close failed after {0} ms. Original window exists={1}, exact caption visible={2}, loaded form count={3}." -f $elapsed.ElapsedMilliseconds,$windowAlive,($visible -ne [IntPtr]::Zero),$loaded)
        }
    }
    function Test-PickerReentryAfterNativeClose([switch]$HoldCallback) {
        . (Join-Path $PSScriptRoot 'Slice4beGuideInputNative.ps1')
        OpenPicker
        $previous=Get-OwnedGuideInputWindow 'Choose tracked actions'
        if($HoldCallback){[void](Run 'invSys.Operations.xlam' 'TestGuideDeferredTiming.Arm')}
        if(-not [GuideInputNative]::PostMessage($previous,0x0010,[IntPtr]::Zero,[IntPtr]::Zero)){
            throw 'Unable to deliver the native picker close before reentry.'
        }
        $deadline=[datetime]::UtcNow.AddSeconds(10)
        do{
            Start-Sleep -Milliseconds 30
            $visible=[InvSysSettingsCapture]::OwnedVisibleForm('Choose tracked actions',[IntPtr]$excel.Hwnd)
        }until($visible -eq [IntPtr]::Zero -or [datetime]::UtcNow -ge $deadline)
        if($visible -ne [IntPtr]::Zero){throw 'The picker remained visible after native close.'}
        OpenPicker
        $replacement=Get-OwnedGuideInputWindow 'Choose tracked actions'
        if($HoldCallback){
            Check 'GuideCuration.OldPickerWindowGoneBeforeDeferredDelivery' ($previous -ne $replacement -and -not [GuideInputNative]::IsWindow($previous))
            $delivery=[string](Run 'invSys.Operations.xlam' 'TestGuideDeferredTiming.Deliver')
            Write-Host ('Deferred picker delivery: '+$delivery)
            Check 'GuideCuration.DeferredCallbackCompletesAfterReplacement' ($delivery -ceq 'OK|1')
        }else{Start-Sleep -Seconds 2}
        $current=[InvSysSettingsCapture]::OwnedVisibleForm('Choose tracked actions',[IntPtr]$excel.Hwnd)
        Check 'GuideCuration.PendingClosePreservesReopenedPicker' ($replacement -eq $current -and (Pick '' 'Count') -ceq '1')
        [void](Pick 'btnCancelGuideActions' 'Click')
        Check 'GuideCuration.ReenteredPickerCancelDisposesForm' ((Pick '' 'Count') -ceq '0')
    }
    function ToggleAction([string]$Id) {
        $ids=@((Pick 'lstGuideActions' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        $index=[Array]::IndexOf($ids,$Id)
        if($index -lt 0){return $false}
        return (Pick 'lstGuideActions' 'Toggle' ([string]$index)) -ceq 'SELECTED'
    }
    function OpenChosenDraft {
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        [void](BoundLibrary 'Open');OpenPicker
        [void](ToggleAction $first.Attempt.ActivityId);[void](ToggleAction $second.Attempt.ActivityId)
        [void](Pick 'btnCreateSelectedGuide' 'Click')
    }
    function CurateCapture([string]$Stage,[string]$Caption='Choose tracked actions') {
        if($CaptureGuideEvidence){CaptureOwnedFormByCaptionEvidence $Caption ('guide-curation-'+$Stage+'.png')}
    }
    function CurateNativeLayout([string]$Form,[string]$Caption,[string]$Stage) {
        if(-not $CaptureGuideEvidence){return}
        if(-not ('CurationNativeLayout' -as [type])){
            Add-Type @"
using System; using System.Runtime.InteropServices;
public static class CurationNativeLayout {
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h,int command);
    [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr h);
}
"@
        }
        Initialize-SettingsCapture
        $handle=[InvSysSettingsCapture]::OwnedVisibleForm($Caption,[IntPtr]$excel.Hwnd)
        if($handle -eq [IntPtr]::Zero){throw 'Native curation layout needs the exact owned visible form.'}
        foreach($command in @(3,9)){
            [void][CurationNativeLayout]::ShowWindow($handle,$command)
            Start-Sleep -Milliseconds 200
            $size=if($command -eq 3){'Maximized'}else{'NativeRestored'}
            Check ('GuideCuration.'+$Stage+'NativeLayout.'+$size) ([CurationNativeLayout]::IsZoomed($handle) -eq ($command -eq 3) -and (BoundControl '' 'Fit' 'Current' $Form) -ceq 'True')
            CurateCapture ($Stage.ToLowerInvariant()+'-'+$size.ToLowerInvariant()) $Caption
        }
    }
    function CurationDigest([string]$Body) {
        if($Body -match '[^\x00-\x7f]'){throw 'Integrity fixture must preserve the ASCII wire.'}
        $sha=[Security.Cryptography.SHA256]::Create()
        try{return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
    }
    CloseRecordingViewer;SelectTarget $Fixture 'config-admin';SetRecordingPolicy $false
    $first=SaveRecordedSetting '801';$second=SaveRecordedSetting '802'
    if($GuidePickerScaleOnly){
        foreach($value in 803..830){[void](SaveRecordedSetting ([string]$value))}
    }
    if($first.Attempt.SequenceId -cne '' -or $second.Attempt.SequenceId -cne '' -or $first.Attempt.Ordinal -ne 0 -or $second.Attempt.Ordinal -ne 0){throw 'Actual unrecorded activity fixture failed; not curation RED.'}
    $published=Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest'
    $publication=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $valid=Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($publication,$Fixture.Warehouse)
    if($published -isnot [bool] -or -not $published -or $valid -isnot [bool] -or -not $valid){throw 'Actual curation publication fixture unavailable.'}
    $model=Get-Content -LiteralPath $publication -Raw|ConvertFrom-Json
    $selectedIds=@($first.Attempt.ActivityId,$second.Attempt.ActivityId)
    $groups=@($model.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -cin $selectedIds})
    if($groups.Count -ne 2 -or @($groups|ForEach-Object Lines).Count -ne 4){throw 'Actual publication omitted the selected source bodies.'}
    $observations=@(foreach($line in ($groups|ForEach-Object Lines)){
        $raw=[IO.File]::ReadAllText((Join-Path $activityRoot ($line.RecordId+'.json')))
        $marker=[regex]::Match($raw,',"ContentSha256":"([0-9a-f]{64})"\}$')
        if(-not $marker.Success){throw 'Original activity envelope lacks its existing digest.'}
        $body=$raw.Substring(0,$marker.Index)+'}';$sha=[Security.Cryptography.SHA256]::Create()
        try{$digest=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
        $original=$raw|ConvertFrom-Json;$record=$body|ConvertFrom-Json
        if($digest -cne $marker.Groups[1].Value -or @($record.PSObject.Properties).Count -ne 26 -or
           (ConvertTo-Json -InputObject $original -Depth 25 -Compress) -cne (ConvertTo-Json -InputObject $line -Depth 25 -Compress)){
            throw 'Published activity differs from its verified original body/envelope.'
        }
        $record
    })
    Check 'GuideCuration.Fixture.ActualUnrecordedActionsAndPublication' $true
    $sourcePins=BoundPins $Fixture.Root
    $activityBefore=ActivityPins;$trainingBefore=BoundPins $journalRoot
    $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $publishBefore=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
    $authorityBefore=Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest'
    $guideRoot=Join-Path $journalRoot 'Guides'
    try {
        OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Accepted Action Paths entry unavailable.'}
        Check 'GuideCuration.EntryIndependentOfRecording' ((BoundControl 'btnChooseGuideActions' 'State' '' 'frmActionPaths') -ceq 'True|True' -and (BoundControl 'lstActionPaths' 'Rows' '' 'frmActionPaths') -ceq '0')
        Check 'GuideCuration.EntryCaption' ((BoundControl 'btnChooseGuideActions' 'Label' '' 'frmActionPaths') -ceq 'Choose tracked actions')
        if($GuideDeferredReentryOnly){Test-PickerReentryAfterNativeClose -HoldCallback;return}
        foreach($size in @('Minimum','Default','Larger','Restored')){
            Check ('GuideCuration.LibraryLayout.'+$size) ((BoundControl '' 'Fit' $size 'frmActionPaths') -ceq 'True')
            CurateCapture ('library-'+$size.ToLowerInvariant()) 'Action Paths'
        }
        CurateNativeLayout 'frmActionPaths' 'Action Paths' 'Library'
        if($CheckGuideSearchNative){
            . (Join-Path $PSScriptRoot 'Slice4beGuideInputNative.ps1')
            foreach($child in @(
                @{Form='frmActionPathLibrary';Caption='Published guides';Open='btnPublishedGuides';Close='btnCloseGuides'},
                @{Form='frmGuideActionPicker';Caption='Choose tracked actions';Open='btnChooseGuideActions';Close='btnCancelGuideActions'}
            )){
                foreach($cycle in 1..2){
                    $owner=[IntPtr]$excel.Hwnd
                    $elapsed=[Diagnostics.Stopwatch]::new()
                    Click-GuideControlNative 'frmActionPaths' 'Action Paths' $child.Open $elapsed
                    do {
                        $handle=[InvSysSettingsCapture]::OwnedVisibleForm($child.Caption,$owner)
                        if($handle -ne [IntPtr]::Zero){break}
                        Start-Sleep -Milliseconds 50
                    } while($elapsed.Elapsed.TotalSeconds -lt 10)
                    $visible=$handle -ne [IntPtr]::Zero -and (Wait-GuideFormCountNative $child.Form '1' ([datetime]::UtcNow.AddMilliseconds(10000-$elapsed.ElapsedMilliseconds)))
                    Check ("GuideCuration.$($child.Form).Cycle$cycle.SingleClickOpens") $visible
                    Write-Host ("Cycle {0}, {1} open: {2} ms." -f $cycle,$child.Caption,$elapsed.ElapsedMilliseconds)
                    if(-not $visible){throw 'Guide window did not open after one delivered click.'}
                    Click-GuideControlNative $child.Form $child.Caption $child.Close $elapsed
                    $closed=(Wait-GuideFormCountNative $child.Form '0' ([datetime]::UtcNow.AddMilliseconds(10000-$elapsed.ElapsedMilliseconds))) -and -not [GuideInputNative]::IsWindow($handle)
                    Check ("GuideCuration.$($child.Form).Cycle$cycle.SingleClickCloses") $closed
                    Write-Host ("Cycle {0}, {1} close: {2} ms." -f $cycle,$child.Caption,$elapsed.ElapsedMilliseconds)
                    if(-not $closed){throw 'Guide window survived the delivered Close click.'}
                }
            }
        }
        if($VisibleExcelForTest -and -not $GuidePickerScaleOnly){
            OpenPicker
            [void](Pick 'btnCancelGuideActions' 'Click')
            Test-ReopenedPickerNativeClose
            Test-PickerReentryAfterNativeClose
        }
        $openTime=[Diagnostics.Stopwatch]::StartNew()
        OpenPicker
        $openTime.Stop()
        $opened=(Pick '' 'Count') -ceq '1'
        Check 'GuideCuration.ActualEntryOpensOnePicker' $opened
        Check 'GuideCuration.EmptySelectionDisablesCreate' ($opened -and (Pick 'btnCreateSelectedGuide' 'State') -ceq 'True|False')
        $source=Pick 'lblGuideActionSource' 'Label'
        Check 'GuideCuration.SourceNamesExactLoadedPublication' ($opened -and $source.Contains([string]$model.PublicationId) -and $source.Contains([string]$model.ContentSha256))
        $ids=@((Pick 'lstGuideActions' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        $permittedIds=@($model.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object SourceId)
        Check 'GuideCuration.CandidatesArePublishedActionIdentities' ($opened -and $ids.Count -gt 1 -and @($ids|Where-Object {$_ -cnotin $permittedIds}).Count -eq 0 -and $first.Attempt.ActivityId -cin $ids -and $second.Attempt.ActivityId -cin $ids)
        if($GuidePickerScaleOnly){
            Write-Host ('Picker scale: '+$ids.Count+' authorized rows; open '+$openTime.ElapsedMilliseconds+' ms')
            if($ids.Count -lt 30){throw 'The required thirty published actions are unavailable; scale timing cannot be qualified.'}
        }
        Check 'GuideCuration.SelectFirstActualAction' (ToggleAction $first.Attempt.ActivityId)
        if($CheckGuideSearchNative -or $GuidePickerScaleOnly){
            . (Join-Path $PSScriptRoot 'Slice4beGuideSearchNative.ps1')
            $beforeKeys=[string](Run 'invSys.Core.xlam' 'TestGuideSearchReadCount.Facts')
            Test-GuideSearchNative 'frmGuideActionPicker' 'Choose tracked actions' 'txtGuideActionSearch' 'lstGuideActions' 'GuideCuration' 'A' 'D' 'M' $false
            Check 'GuideCuration.NativeKeysUseNoCoreLibraryOrPolicyReads' (([string](Run 'invSys.Core.xlam' 'TestGuideSearchReadCount.Facts')) -ceq $beforeKeys)
            Check 'GuideCuration.NativeSearchRetainsExactSelectedAction' ((Pick 'lstGuideActions' 'SelectedValues') -ceq ($first.Attempt.ActivityId+"`n"))
        }
        # A caption copied from the visible list must find exactly its published actions.
        $caption=[string]$first.Attempt.Caption
        $expectedIds=@(foreach($group in $model.Groups){
            if($group.Source -ceq 'Activity' -and $group.SourceId -cin $ids -and
                @($group.Lines | Where-Object {$_.OutcomeCode -ceq 'REQUESTED' -and $_.Caption -ceq $caption}).Count){$group.SourceId}
        })
        [void](Pick 'txtGuideActionSearch' 'Write' $caption.ToLowerInvariant())
        $actualIds=@((Pick 'lstGuideActions' 'Values') -split "`n" | Where-Object {$_})
        Check 'GuideCuration.SearchDisplayedCaptionReturnsExactActions' ($caption -cne '' -and $first.Attempt.ActivityId -cin $expectedIds -and
            (($actualIds | Sort-Object) -join '|') -ceq (($expectedIds | Sort-Object) -join '|'))
        [void](Pick 'txtGuideActionSearch' 'Write' 'no-match-curation-filter')
        Check 'GuideCuration.SearchFiltersWithoutLosingSelection' ($opened -and (Pick 'lstGuideActions' 'Rows') -ceq '0' -and (Pick 'lblGuideActionStatus' 'Label') -match '(?i)1 selected' -and (Pick 'btnCreateSelectedGuide' 'State') -ceq 'True|True')
        [void](Pick 'txtGuideActionSearch' 'Write' 'ADMIN_SETTINGS_SAVE_VALUE')
        Check 'GuideCuration.SearchByControlRestoresSelectedIdentity' ($opened -and (Pick 'lstGuideActions' 'SelectedValues') -ceq ($first.Attempt.ActivityId+"`n"))
        if($GuidePickerScaleOnly){
            [void](Pick 'btnCancelGuideActions' 'Click')
            Check 'GuideCuration.ScaleCloseDisposesPickerWithSourcesPreserved' ((Pick '' 'Count') -ceq '0' -and (BoundSame $sourcePins (BoundPins $Fixture.Root)))
            Test-ReopenedPickerNativeClose
            Test-PickerReentryAfterNativeClose
            return
        }
        Check 'GuideCuration.SelectSecondActualAction' (ToggleAction $second.Attempt.ActivityId)
        [void](Pick 'txtGuideActionSearch' 'Write' '')
        $retainedIds=@((Pick 'lstGuideActions' 'SelectedValues') -split "`n" | Where-Object {$_ -cne ''})
        Check 'GuideCuration.HiddenAndNewSelectionsKeepExactActivityIds' (
            $retainedIds.Count -eq 2 -and $first.Attempt.ActivityId -cin $retainedIds -and
            $second.Attempt.ActivityId -cin $retainedIds)
        # D18-SEARCH-INPUT-02: a loaded header is presentation data until the
        # operator explicitly reloads. A changed publication cannot authorize
        # Create, but it cannot make each Search key perform a guarded read.
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
        $searchReadsBefore=[string](Run 'invSys.Core.xlam' 'TestGuideSearchReadCount.Facts')
        [void](Pick 'txtGuideActionSearch' 'Write' 'no-match-after-refresh')
        Check 'GuideCuration.SearchUsesNoCoreLibraryOrPolicyReads' (
            ([string](Run 'invSys.Core.xlam' 'TestGuideSearchReadCount.Facts')) -ceq $searchReadsBefore)
        Check 'GuideCuration.SearchUsesLoadedSnapshotAfterPublicationChange' (
            (Pick 'lstGuideActions' 'Rows') -ceq '0' -and
            (Pick 'lblGuideActionStatus' 'Label') -match '(?i)2 selected' -and
            (Pick 'btnCreateSelectedGuide' 'State') -ceq 'True|True')
        [void](Pick 'btnCreateSelectedGuide' 'Click')
        Check 'GuideCuration.StaleHiddenSelectionCannotCreate' ((Draft '' 'Count') -ceq '0')
        Check 'GuideCuration.CreateRevalidatesCoreAfterSearch' (
            ([string](Run 'invSys.Core.xlam' 'TestGuideSearchReadCount.Facts')) -cne $searchReadsBefore)
        [void](Pick 'btnCancelGuideActions' 'Click')
        OpenPicker
        [void](ToggleAction $first.Attempt.ActivityId);[void](ToggleAction $second.Attempt.ActivityId)
        OpenPicker
        Check 'GuideCuration.RepeatedEntryReusesSelection' ($opened -and (Pick '' 'Count') -ceq '1' -and (Pick 'lblGuideActionStatus' 'Label') -match '(?i)2 selected')
        foreach($size in @('Minimum','Default','Larger','Restored')){
            Check ('GuideCuration.PickerLayout.'+$size) ($opened -and (Pick '' 'Fit' $size) -ceq 'True')
            if($opened){CurateCapture ('picker-'+$size.ToLowerInvariant())}
        }
        if($opened){CurateNativeLayout 'frmGuideActionPicker' 'Choose tracked actions' 'Picker'}
        if($CheckGuideSearchNative){
            Open-GuideEditorNative 'frmGuideActionPicker' 'Choose tracked actions' 'btnCreateSelectedGuide' 'GuideCuration.OneClickShowsGuideEditor'
        }else{[void](Pick 'btnCreateSelectedGuide' 'Click')}
        $drafted=(Draft '' 'Count') -ceq '1'
        Check 'GuideCuration.CreateUsesExistingGuideEditor' ($opened -and $drafted)
        Check 'GuideCuration.NoRecordedSequenceClaim' ($drafted -and (Draft 'lblGuideSource' 'Label') -match 'No recorded sequence; authored order is not execution evidence')
        $stepIds=@((Draft 'lstGuideSteps' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        Check 'GuideCuration.DistinctAuthoredStepIdentities' ($drafted -and $stepIds.Count -eq 2 -and @($stepIds|Select-Object -Unique).Count -eq 2 -and @($stepIds|Where-Object {$_ -cin $selectedIds}).Count -eq 0)
        Check 'GuideCuration.NoAutomaticExpectation' ($drafted -and (Draft 'lblGuideExpectationSummary' 'Label') -ceq 'Guide expectation: None')
        $evidence=Draft 'txtGuideEvidence' 'Text'
        Check 'GuideCuration.OriginalActionEvidenceDisplayed' ($drafted -and $evidence.Contains([string]$first.Attempt.ActivityId) -and $evidence.Contains([string]$second.Attempt.ActivityId) -and $evidence -match 'REQUESTED' -and $evidence -match 'COMPLETED')
        [void](Draft 'txtGuideSubject' 'Write' 'Direct action training')
        [void](Draft 'txtGuideName' 'Write' 'Direct tracked-action guide')
        [void](Draft 'txtGuideTags' 'Write' 'direct, unrecorded')
        [void](Draft 'txtGuideInstructions' 'Write' 'Authored order is not execution evidence.')
        [void](Pick 'btnCreateSelectedGuide' 'Click')
        Check 'GuideCuration.RepeatedCreateRetainsStagedText' ($drafted -and (Draft '' 'Count') -ceq '1' -and (Draft 'txtGuideName' 'Text') -ceq 'Direct tracked-action guide')
        [void](Draft 'lstGuideSteps' 'Select' '1');[void](Draft 'btnGuideStepUp' 'Click')
        [void](Draft 'txtGuideStepInstruction' 'Write' 'First authored instruction.')
        Check 'GuideCuration.ReorderRetainsStableStepIdsAndOriginalEvidence' ($drafted -and (Draft 'lstGuideSteps' 'Values') -ceq ($stepIds[1]+"`n"+$stepIds[0]+"`n") -and (Draft 'txtGuideEvidence' 'Text') -ceq $evidence)
        Check 'GuideCuration.NonSaveActionsPreserveEverySourceByte' (BoundSame $sourcePins (BoundPins $Fixture.Root))
        foreach($size in @('Minimum','Default','Larger','Restored')){
            Check ('GuideCuration.EditorLayout.'+$size) ($drafted -and (Draft '' 'Fit' $size) -ceq 'True')
            if($drafted){CurateCapture ('editor-'+$size.ToLowerInvariant()) 'Action Path guide'}
        }
        if($drafted){CurateNativeLayout 'frmActionPathGuide' 'Action Path guide' 'Editor'}
        [void](Draft 'btnSaveGuide' 'Click')
        $files=@();if(Test-Path -LiteralPath $guideRoot){$files=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.1.json')}
        $guide=$null;if($files.Count -eq 1){$guide=ReadGuideExpectationRecord $files[0].FullName}
        $saved=$null -ne $guide
        Check 'GuideCuration.SavePublishesValidatedSchemaThree' ($drafted -and $saved -and $guide.SchemaVersion -eq 3 -and $guide.Version -eq 1 -and $guide.Subject -ceq 'Direct action training' -and $guide.Name -ceq 'Direct tracked-action guide')
        Check 'GuideCuration.SaveHasNoFabricatedRunOrExpectation' ($saved -and @($guide.SourceRun.PSObject.Properties).Count -eq 0 -and $guide.ExpectedConclusion.TerminalKind -ceq 'None' -and @($guide.ExpectedConclusion.Steps).Count -eq 0)
        Check 'GuideCuration.SavePreservesEverySelectedOriginalBody' ($saved -and (ConvertTo-Json -InputObject @($guide.Observations) -Depth 25 -Compress) -ceq (ConvertTo-Json -InputObject $observations -Depth 25 -Compress))
        Check 'GuideCuration.SavePreservesAuthoredOrderAndSourceIdentity' ($saved -and $guide.Steps[0].StepId -ceq $stepIds[1] -and $guide.Steps[1].StepId -ceq $stepIds[0] -and $guide.Steps[0].SourceActivityId -ceq $groups[1].SourceId -and $guide.Steps[1].SourceActivityId -ceq $groups[0].SourceId -and $guide.Steps[0].Instruction -ceq 'First authored instruction.')
        [void](Draft 'btnSaveGuide' 'Click')
        $revision=$null;if($saved){$revision=ReadGuideExpectationRecord (Join-Path $guideRoot ($guide.ActionPathId+'.2.json'))}
        Check 'GuideCuration.RepeatedSaveAppendsImmutableRevision' ($saved -and $null -ne $revision -and $revision.PreviousRecordId -ceq $guide.RecordId -and $revision.PreviousSha256 -ceq $guide.ContentSha256 -and @($revision.SourceRun.PSObject.Properties).Count -eq 0)
        $savedPins=BoundPins $journalRoot
        [void](Draft 'btnGuideExpectedConclusion' 'Click');[void](ToggleAction $first.Attempt.ActivityId)
        Check 'GuideCuration.SelectionChangeClosesDraftAndExpectation' ($drafted -and (Draft '' 'Count') -ceq '0' -and (BoundControl '' 'Count' '' 'frmActionPathExpectation') -ceq '0')
        [void](Pick 'btnCreateSelectedGuide' 'Click');[void](Pick 'btnCancelGuideActions' 'Click')
        Check 'GuideCuration.CancelClosesOwnedStagingPreservesSavedGuides' ($opened -and (Pick '' 'Count') -ceq '0' -and (Draft '' 'Count') -ceq '0' -and (BoundSame $savedPins (BoundPins $journalRoot)))
        Check 'GuideCuration.DoesNotCreateRecordingOrEvaluation' (@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json' -ErrorAction SilentlyContinue).Count -eq 0 -and (BoundPins (Join-Path $journalRoot 'Evaluations')).Count -eq 0)
        Check 'GuideCuration.SaveCancelPreserveActivityAndConfig' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore -and (PinsRetained $activityBefore) -and (ActivityPins).Count -eq $activityBefore.Count)
        OpenChosenDraft
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
        [void](Draft 'btnSaveGuide' 'Click')
        Check 'GuideCuration.ReloadedPublicationInvalidatesCapturedDraft' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (Draft 'btnSaveGuide' 'State') -cne 'True|True' -and (BoundSame $savedPins (BoundPins $journalRoot)))
        OpenChosenDraft;[void](Draft 'btnGuideExpectedConclusion' 'Click')
        [void](BoundControl 'btnClose' 'Click' '' 'frmActionPaths')
        Check 'GuideCuration.LibraryCloseDisposesPickerDraftAndExpectation' ($opened -and (Pick '' 'Count') -ceq '0' -and (Draft '' 'Count') -ceq '0' -and (BoundControl '' 'Count' '' 'frmActionPathExpectation') -ceq '0')
        OpenChosenDraft;CloseRecordingViewer
        Check 'GuideCuration.ViewerCloseDisposesOwnedStaging' ($drafted -and (Pick '' 'Count') -ceq '0' -and (Draft '' 'Count') -ceq '0' -and (BoundSame $savedPins (BoundPins $journalRoot)))
        OpenChosenDraft
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut');[void](Draft 'btnSaveGuide' 'Click')
        Check 'GuideCuration.SignOutClearsEvidenceAndCannotSave' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $savedPins (BoundPins $journalRoot)))
        OpenChosenDraft
        $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
        $authBytes=[IO.File]::ReadAllBytes($authPath)
        try {
            if($WaitForExcelReadyForTest){Wait-ExcelReadyForTest 'CurationFixture.OpenPermissionWorkbook'}
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            try {
                $caps=Table $auth 'tblCapabilities';$revoked=0
                foreach($row in $caps.ListRows){
                    if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and
                       $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ACTION_PATH_MAINT'){
                        $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$revoked++
                    }
                }
                if($revoked -ne 1){throw 'Guide maintenance fixture is not unique.'};$auth.Save()
            } finally {$auth.Close($false)}
            [void](Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Fixture.Warehouse))
            $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
            $signed=Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'
            if($allowed -isnot [bool] -or $allowed -or $signed -isnot [bool] -or -not $signed){throw 'Permission loss fixture unavailable.'}
            [void](Draft 'btnSaveGuide' 'Click')
            Check 'GuideCuration.MaintenanceRevocationInvalidatesStagedDraft' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $savedPins (BoundPins $journalRoot)))
        } finally {CloseRecordingViewer;[IO.File]::WriteAllBytes($authPath,$authBytes);SelectTarget $Fixture 'config-admin'}
        OpenChosenDraft
        try {
            SaveGuideExpectationVisibility $false
            [void](Draft 'btnSaveGuide' 'Click')
            Check 'GuideCuration.PolicyLossInvalidatesWithoutDroppingSteps' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $savedPins (BoundPins $journalRoot)))
        } finally {CloseRecordingViewer;SaveGuideExpectationVisibility $true}
        OpenChosenDraft;SelectTarget $Other 'config-admin';[void](Draft 'btnSaveGuide' 'Click')
        Check 'GuideCuration.ContextLossCannotRetargetSave' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $savedPins (BoundPins $journalRoot)))
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer;[void](BoundLibrary 'Open')
        Check 'GuideCuration.ReaderCannotOpenPicker' ((BoundControl 'btnChooseGuideActions' 'State' '' 'frmActionPaths') -ceq 'True|False' -and (Pick '' 'Count') -ceq '0')
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
        $publicationBytes=[IO.File]::ReadAllBytes($publication)
        try {
            $altered=[Text.Encoding]::UTF8.GetString($publicationBytes)|ConvertFrom-Json
            $badGroup=@($altered.Groups|Where-Object SourceId -CEQ $first.Attempt.ActivityId)
            $badAttempt=@($badGroup[0].Lines|Where-Object OutcomeCode -CEQ 'REQUESTED')
            if($badGroup.Count -ne 1 -or $badAttempt.Count -ne 1){throw 'Invalid-digest fixture source is ambiguous.'}
            $badAttempt[0].ContentSha256='0'*64
            $altered.PSObject.Properties.Remove('ContentSha256')
            $body=ConvertTo-Json -InputObject $altered -Depth 30 -Compress
            if($body -match '[^\x00-\x7f]'){throw 'Diagnostic publication fixture must retain ASCII wire.'}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$digest=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            [IO.File]::WriteAllText($publication,$body.Substring(0,$body.Length-1)+',"ContentSha256":"'+$digest+'"}',[Text.UTF8Encoding]::new($false))
            $valid=Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($publication,$Fixture.Warehouse)
            if($valid -isnot [bool] -or -not $valid){throw 'Outer publication fixture must remain valid.'}
            OpenRecordingViewer;[void](BoundLibrary 'Open');OpenPicker
            $ids=@((Pick 'lstGuideActions' 'Values') -split "`n"|Where-Object {$_ -cne ''})
            Check 'GuideCuration.InvalidInnerDigestRejectedWithoutSourceMutation' ($opened -and $first.Attempt.ActivityId -cnotin $ids -and (Pick 'lblGuideActionStatus' 'Label') -match '1 unavailable' -and (PinsRetained $activityBefore))
            Check 'GuideCuration.InvalidInnerDigestDoesNotSuppressOtherValidAction' ($opened -and $second.Attempt.ActivityId -cin $ids -and (BoundSame $savedPins (BoundPins $journalRoot)))
        } finally {CloseRecordingViewer;[IO.File]::WriteAllBytes($publication,$publicationBytes)}
        # Preserve exact original JSON bytes except the single diagnosed member;
        # rehash both layers so these are semantic-integrity cases, not bad hashes.
        $originalPublication=[Text.Encoding]::UTF8.GetString($publicationBytes)
        $originalOutcome=[IO.File]::ReadAllText((Join-Path $activityRoot ($first.Outcome.RecordId+'.json')))
        # The publisher indexes the same completion in Lines and Outcomes.
        # Replace both representations, without inventing a second observation.
        if([regex]::Matches($originalPublication,[regex]::Escape($originalOutcome)).Count -ne 2){throw 'Exact original outcome must occur in both Lines and Outcomes.'}
        $innerMarker=[regex]::Match($originalOutcome,',"ContentSha256":"([0-9a-f]{64})"\}$')
        if(-not $innerMarker.Success){throw 'Original outcome digest missing.'}
        $originalBody=$originalOutcome.Substring(0,$innerMarker.Index)+'}'
        if((CurationDigest $originalBody) -cne $innerMarker.Groups[1].Value){throw 'Original outcome digest invalid.'}
        foreach($case in @(
            @{Field='RecordId';Value='not-a-guid';Wire='"not-a-guid"'},
            @{Field='CatalogVersion';Value=1;Wire='1'},
            @{Field='PackageSetVersion';Value='different-package-fixture';Wire='"different-package-fixture"'},
            @{Field='BuildIdentity';Value='different-build-fixture';Wire='"different-build-fixture"'}
        )){
            try {
                $member=[regex]::Matches($originalBody,'"'+$case.Field+'":("(?:[^"\\]|\\.)*"|[0-9]+)')
                if($member.Count -ne 1){throw 'Source-integrity fixture member is ambiguous.'}
                $changedBody=$originalBody.Replace($member[0].Value,'"'+$case.Field+'":'+$case.Wire)
                $changed=$changedBody|ConvertFrom-Json
                if($changedBody -ceq $originalBody -or $changed.($case.Field) -cne $case.Value){throw 'Source-integrity fixture did not change only its intended member.'}
                $changedEnvelope=$changedBody.Substring(0,$changedBody.Length-1)+',"ContentSha256":"'+(CurationDigest $changedBody)+'"}'
                $changedPublication=$originalPublication.Replace($originalOutcome,$changedEnvelope)
                $outerMarker=[regex]::Match($changedPublication,',"ContentSha256":"[0-9a-f]{64}"\}$')
                if(-not $outerMarker.Success){throw 'Original publication digest missing.'}
                $outerBody=$changedPublication.Substring(0,$outerMarker.Index)+'}'
                [IO.File]::WriteAllText($publication,$outerBody.Substring(0,$outerBody.Length-1)+',"ContentSha256":"'+(CurationDigest $outerBody)+'"}',[Text.UTF8Encoding]::new($false))
                $valid=Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($publication,$Fixture.Warehouse)
                if($valid -isnot [bool] -or -not $valid){throw 'Semantic-integrity fixture needs a valid outer publication.'}
                OpenRecordingViewer;[void](BoundLibrary 'Open');OpenPicker
                $ids=@((Pick 'lstGuideActions' 'Values') -split "`n"|Where-Object {$_ -cne ''})
                Check ('GuideCuration.SourceIntegrity.'+$case.Field+'.InvalidActionUnavailable') ($first.Attempt.ActivityId -cnotin $ids -and (Pick 'lblGuideActionStatus' 'Label') -match '1 unavailable')
                Check ('GuideCuration.SourceIntegrity.'+$case.Field+'.ValidNeighborAndSourcesRetained') ($second.Attempt.ActivityId -cin $ids -and (PinsRetained $activityBefore) -and (BoundSame $savedPins (BoundPins $journalRoot)))
            } finally {CloseRecordingViewer;[IO.File]::WriteAllBytes($publication,$publicationBytes)}
        }
        if($saved -and $null -ne $revision){
            OpenRecordingViewer;[void](BoundLibrary 'Open');BoundOpen;BoundSelect $revision
        }
        Check 'GuideCuration.PublishedReaderLoadsCuratedGuideWithoutRun' ($saved -and $null -ne $revision -and (BoundControl 'txtPublishedInstructions' 'Text').Contains([string]$guide.Name) -and (BoundControl 'btnUseGuideForRun' 'State') -ceq 'True|False')
        if($saved -and $null -ne $revision){[void](BoundControl 'btnEditPublishedGuide' 'Click')}
        Check 'GuideCuration.PublishedEditorRestoresCuratedStepIds' ($saved -and $null -ne $revision -and (Draft 'lstGuideSteps' 'Values') -ceq ((@($revision.Steps.StepId) -join "`n")+"`n"))
        if($saved -and $null -ne $revision){
            [void](Draft 'txtGuideName' 'Write' 'Reopened directly curated guide');[void](Draft 'btnSaveGuide' 'Click')
            $third=ReadGuideExpectationRecord (Join-Path $guideRoot ($guide.ActionPathId+'.3.json'))
        } else {$third=$null}
        Check 'GuideCuration.PublishedRevisionRetainsEmptyRunAndOriginalBodies' ($saved -and $null -ne $third -and $third.PreviousRecordId -ceq $revision.RecordId -and $third.PreviousSha256 -ceq $revision.ContentSha256 -and @($third.SourceRun.PSObject.Properties).Count -eq 0 -and (ConvertTo-Json -InputObject @($third.Observations) -Depth 25 -Compress) -ceq (ConvertTo-Json -InputObject $observations -Depth 25 -Compress))
        Check 'GuideCuration.NoPublicationOrAuthorityRead' ((Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishBefore -and (Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest') -eq $authorityBefore)
        foreach($path in $sourcePins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $sourcePins[$path] -and $path -cne $Fixture.Config){throw 'Curation changed a pre-existing source file.'}}
    } finally {CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
