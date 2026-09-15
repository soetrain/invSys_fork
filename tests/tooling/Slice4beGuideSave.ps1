# D18 immutable guide publication through the actual packaged editor controls.
# The fixture journal comes from real Admin actions; no guide is fabricated.
function Test-GuideSave($Fixture,$Other) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $deploy 'invSys.Core.xlam'))
    try {
        $entry=$archive.GetEntry('docProps/custom.xml')
        if($null -eq $entry){throw 'Guide publisher package metadata is missing.'}
        $reader=[IO.StreamReader]::new($entry.Open())
        try {[xml]$metadata=$reader.ReadToEnd()}finally{$reader.Dispose()}
        $versions=@($metadata.Properties.property|Where-Object name -CEQ 'invSysPackageSetVersion')
        $builds=@($metadata.Properties.property|Where-Object name -CEQ 'invSysBuildIdentity')
        if($versions.Count -ne 1 -or $builds.Count -ne 1){throw 'Guide publisher package identity is ambiguous.'}
        $publisherVersion=[string]$versions[0].InnerText;$publisherBuild=[string]$builds[0].InnerText
        if($publisherVersion -ceq '' -or $publisherBuild -notmatch '^[0-9a-f]{32}$'){throw 'Guide publisher package identity is invalid.'}
    } finally {$archive.Dispose()}
    function GuideSaveControl([string]$Name,[string]$Action,[string]$Value='', [string]$Form='frmActionPathGuide') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($Form,$Name,$Action,$Value))
    }
    function GuideSaveLibrary([string]$Action,[string]$Value='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @($Action,$Value))
    }
    function CaptureSavedGuide([string]$Version) {
        if(-not $CaptureEvidence){return}
        Initialize-SettingsCapture
        $handle=[InvSysSettingsCapture]::OwnedVisibleForm('Action Path guide',[IntPtr]$excel.Hwnd).ToInt64()
        CaptureOwnedFormEvidence 'Action Path guide' ('guide-save-version-'+$Version+'.png') $handle
        Check ('GuideSave.VisibleCapture.Version'+$Version) $true
    }
    function GuideFiles {
        if(Test-Path -LiteralPath $guideRoot){Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'}
    }
    function ReadSavedGuide($File) {
        if($null -eq $File){return $null}
        try {
            $text=[IO.File]::ReadAllText($File.FullName)
            if($File.Length -ne $text.Length -or $text -match '[^\x00-\x7f]'){return $null}
            $value=$text|ConvertFrom-Json
            $fields='SchemaVersion|RecordKind|ActionPathId|Version|RecordId|PreviousRecordId|PreviousSha256|WarehouseId|OriginWarehouseId|CreatedByUserId|CreatedAtUTC|Lifecycle|Name|Tags|Instructions|CatalogVersion|PackageSetVersion|BuildIdentity|PolicyVersion|Steps|Observations|SourceRun|ExpectedConclusion|ContentSha256' -split '\|'
            $actual=@($value.PSObject.Properties.Name)
            if($actual.Count -ne $fields.Count -or @($actual|Where-Object {$_ -cnotin $fields}).Count){return $null}
            $sourceFields='ActionPathId|SequenceId|JournalVersion|RecordId|ContentSha256|RecordedByUserId|EntryCreatedAtUTC|Lifecycle|ReasonCode|ActionCount|CapturePolicyVersion|CatalogVersion|PackageSetVersion|BuildIdentity|RestrictedObservationCount' -split '\|'
            if(@($value.SourceRun.PSObject.Properties).Count -ne $sourceFields.Count){return $null}
            foreach($field in $sourceFields){if($null -eq $value.SourceRun.PSObject.Properties[$field]){return $null}}
            if($null -eq $value.ExpectedConclusion.PSObject.Properties['TerminalKind'] -or $null -eq $value.ExpectedConclusion.PSObject.Properties['Steps']){return $null}
            foreach($step in $value.Steps){
                $stepFields='StepId|Method|ControlId|Caption|Instruction|SourceActivityId' -split '\|'
                if(@($step.PSObject.Properties).Count -ne $stepFields.Count){return $null}
                foreach($field in $stepFields){if($null -eq $step.PSObject.Properties[$field]){return $null}}
            }
            $match=[regex]::Match($text,',"ContentSha256":"([0-9a-f]{64})"\}$')
            if(-not $match.Success -or $File.Length -gt 1048576){return $null}
            $body=$text.Substring(0,$match.Index)+'}'
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            if($hash -cne $value.ContentSha256){return $null}
            return $value
        } catch {return $null}
    }
    function SourcePins {
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $journalRoot -Recurse -File){
            if(-not $file.FullName.StartsWith($guideRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){
                $pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash
            }
        }
        return $pins
    }
    function SamePins($Before,$After) {
        if($Before.Count -ne $After.Count){return $false}
        foreach($key in $Before.Keys){if(-not $After.ContainsKey($key) -or $Before[$key] -cne $After[$key]){return $false}}
        return $true
    }
    function GuidePins {
        $pins=@{}
        foreach($file in @(GuideFiles)){$pins[$file.Name]=(Get-FileHash -LiteralPath $file.FullName).Hash}
        return $pins
    }
    function OpenSaveDraft {
        if((GuideSaveLibrary 'Open') -cne 'DELIVERED'){throw 'Save guard fixture cannot open the current Viewer library.'}
        if((GuideSaveLibrary 'Select' $source.ActionPathId) -cne 'SELECTED'){throw 'Save guard fixture cannot select its source.'}
        if((GuideSaveControl 'btnCreateGuide' 'Click' '' 'frmActionPaths') -cne 'DELIVERED'){throw 'Save guard fixture cannot deliver Create guide.'}
        if((GuideSaveControl '' 'Count') -cne '1'){throw 'Save guard fixture has no unique guide editor.'}
        [void](GuideSaveControl 'txtGuideName' 'Write' 'Guarded draft must not be published')
    }
    function SaveGuideVisibility([bool]$Visible) {
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            $saved=Run 'invSys.Admin.xlam' 'TestD5Commands.PublishedReadVisibilityForTest' @($Visible)
            if($saved -isnot [bool] -or -not $saved){throw 'Actual policy save fixture failed; not guide-save RED.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    SetRecordingPolicy $true
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Existing recorder unavailable; not guide-save RED.'}
    $firstAction=SaveRecordedSetting '640'
    $secondAction=SaveRecordedSetting '641'
    if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Existing recorder cannot stop; not guide-save RED.'}
    $sequence=[string]$firstAction.Attempt.SequenceId
    if(-not (HasSequence $firstAction 1) -or -not (HasSequence $secondAction 2) -or -not (JournalChain $sequence 6)){
        throw 'Actual source sequence invalid; not guide-save RED.'
    }
    $entries=@(RecordingJournal $sequence|Sort-Object Version)
    $source=$entries[-1]
    if((GuideSaveLibrary 'Open') -cne 'DELIVERED' -or (GuideSaveLibrary 'Select' $source.ActionPathId) -cne 'SELECTED'){
        throw 'Existing recorded-run selection unavailable; not guide-save RED.'
    }
    if((GuideSaveControl 'btnCreateGuide' 'Click' '' 'frmActionPaths') -cne 'DELIVERED' -or
       (GuideSaveControl '' 'Count') -cne '1' -or (GuideSaveControl 'lstGuideSteps' 'Rows') -cne '2'){
        throw 'Accepted guide draft editor unavailable; not guide-save RED.'
    }
    $guideRoot=Join-Path $journalRoot 'Guides'
    if(@(GuideFiles).Count){throw 'Guide-save fixture is not empty.'}
    $sourcePins=SourcePins; $activity=ActivityPins; $configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $publishedBefore=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
    $observed=GuideSaveControl 'txtGuideEvidence' 'Text'
    try {
        [void](GuideSaveControl 'txtGuideName' 'Write' 'Review and verify a Settings change')
        [void](GuideSaveControl 'txtGuideTags' 'Write' 'training, settings')
        [void](GuideSaveControl 'txtGuideInstructions' 'Write' 'Read the setting, save it, and inspect the observed outcome.')
        [void](GuideSaveControl 'lstGuideSteps' 'Select' '0')
        $firstStep=GuideSaveControl 'lstGuideSteps' 'Selected'
        $firstInstruction='First reviewed instruction '+[char]0x03A9+"`r`n"+'Check "quotes".'
        [void](GuideSaveControl 'txtGuideStepInstruction' 'Write' $firstInstruction)
        [void](GuideSaveControl 'lstGuideSteps' 'Select' '1')
        $secondStep=GuideSaveControl 'lstGuideSteps' 'Selected'
        [void](GuideSaveControl 'txtGuideStepInstruction' 'Write' 'Second reviewed instruction')
        $savePresent=(GuideSaveControl 'btnSaveGuide' 'State') -ceq 'True|True'
        Check 'GuideSave.ActualSaveControlAvailable' $savePresent
        Check 'GuideSave.PublicationMeaningVisibleBeforeSave' ((GuideSaveControl '' 'Labels') -match '(?i)save.*publish.*version')
        [void](GuideSaveControl 'txtGuideName' 'Write' '   ')
        $blank=GuideSaveControl 'btnSaveGuide' 'Click'
        Check 'GuideSave.BlankNameRejectedWithoutPublication' ($savePresent -and $blank -ceq 'DELIVERED' -and @(GuideFiles).Count -eq 0 -and (GuideSaveControl 'lblGuideStatus' 'Label') -match '(?i)name.*required|enter.*name')
        [void](GuideSaveControl 'txtGuideName' 'Write' 'Review and verify a Settings change')
        $save=GuideSaveControl 'btnSaveGuide' 'Click'
        $files=@(GuideFiles)
        $firstFile=if($files.Count -eq 1){$files[0]}else{$null}
        $firstRecord=ReadSavedGuide $firstFile
        $validFirst=($null -ne $firstRecord)
        Check 'GuideSave.FirstSavePublishesOneValidatedRecord' ($save -ceq 'DELIVERED' -and $validFirst -and $firstRecord.RecordKind -ceq 'Guide' -and $firstRecord.SchemaVersion -eq 1 -and $firstRecord.Version -eq 1 -and $firstRecord.Lifecycle -ceq 'Published')
        $guideId=if($validFirst){[string]$firstRecord.ActionPathId}else{''}
        $guideGuid=[guid]::Empty
        Check 'GuideSave.NewGuideIdentityIsDistinctFromSourceAndSteps' ($validFirst -and [guid]::TryParse($guideId,[ref]$guideGuid) -and $guideGuid -ne [guid]::Empty -and $guideId -cnotin @($source.ActionPathId,$sequence,$firstStep,$secondStep) -and $firstFile.Name -ceq ($guideId+'.1.json'))
        Check 'GuideSave.FirstVersionHasNoPreviousLink' ($validFirst -and $firstRecord.PreviousRecordId -ceq '' -and $firstRecord.PreviousSha256 -ceq '')
        $recordGuid=[guid]::Empty
        Check 'GuideSave.CreatorWarehouseAndReleaseProvenance' ($validFirst -and $firstRecord.WarehouseId -ceq $Fixture.Warehouse -and $firstRecord.OriginWarehouseId -ceq $Fixture.Warehouse -and
            $firstRecord.CreatedByUserId -ceq 'config-admin' -and $firstRecord.CreatedAtUTC -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$' -and
            $firstRecord.CatalogVersion -eq $source.CatalogVersion -and $firstRecord.PolicyVersion -eq $source.PolicyVersion -and
            $firstRecord.PackageSetVersion -ceq $publisherVersion -and
            $firstRecord.BuildIdentity -ceq $publisherBuild -and
            [guid]::TryParse([string]$firstRecord.RecordId,[ref]$recordGuid) -and $recordGuid -ne [guid]::Empty -and $firstRecord.RecordId -cnotin @($guideId,$source.RecordId,$sequence,$firstStep,$secondStep))
        Check 'GuideSave.AuthoredFieldsAndTagsPreserved' ($validFirst -and $firstRecord.Name -ceq 'Review and verify a Settings change' -and @($firstRecord.Tags).Count -eq 2 -and $firstRecord.Tags[0] -ceq 'training' -and $firstRecord.Tags[1] -ceq 'settings' -and $firstRecord.Instructions -ceq 'Read the setting, save it, and inspect the observed outcome.')
        $steps=@(if($validFirst){$firstRecord.Steps})
        Check 'GuideSave.StableAuthoredStepsRetainExactSourceActions' ($steps.Count -eq 2 -and $steps[0].StepId -ceq $firstStep -and $steps[1].StepId -ceq $secondStep -and
            $steps[0].SourceActivityId -ceq $firstAction.Attempt.ActivityId -and $steps[1].SourceActivityId -ceq $secondAction.Attempt.ActivityId -and
            $steps[0].ControlId -ceq $firstAction.Attempt.ControlId -and $steps[1].ControlId -ceq $secondAction.Attempt.ControlId -and
            $steps[0].Caption -ceq $firstAction.Attempt.Caption -and $steps[1].Caption -ceq $secondAction.Attempt.Caption -and
            $steps[0].Method -ceq 'How-To' -and $steps[1].Method -ceq 'How-To' -and
            $steps[0].Instruction -ceq $firstInstruction -and $steps[1].Instruction -ceq 'Second reviewed instruction')
        Check 'GuideSave.SourceRunVersionAndHashAreExact' ($validFirst -and $firstRecord.SourceRun.ActionPathId -ceq $source.ActionPathId -and $firstRecord.SourceRun.SequenceId -ceq $sequence -and $firstRecord.SourceRun.JournalVersion -eq 6 -and $firstRecord.SourceRun.RecordId -ceq $source.RecordId -and $firstRecord.SourceRun.ContentSha256 -ceq $source.ContentSha256)
        Check 'GuideSave.SourceLifecycleActorAndCaptureProvenancePreserved' ($validFirst -and $firstRecord.SourceRun.RecordedByUserId -ceq $source.CreatedByUserId -and
            $firstRecord.SourceRun.EntryCreatedAtUTC -ceq $source.CreatedAtUTC -and $firstRecord.SourceRun.Lifecycle -ceq $source.Lifecycle -and
            $firstRecord.SourceRun.ReasonCode -ceq $source.ReasonCode -and $firstRecord.SourceRun.ActionCount -eq $source.ActionCount -and
            $firstRecord.SourceRun.CapturePolicyVersion -eq $source.PolicyVersion -and $firstRecord.SourceRun.CatalogVersion -eq $source.CatalogVersion -and
            $firstRecord.SourceRun.PackageSetVersion -ceq $source.PackageSetVersion -and $firstRecord.SourceRun.BuildIdentity -ceq $source.BuildIdentity -and
            $firstRecord.SourceRun.RestrictedObservationCount -eq 0)
        Check 'GuideSave.OriginalObservationsRetainedInOrder' ($validFirst -and ($firstRecord.Observations|ConvertTo-Json -Depth 12 -Compress) -ceq ($source.Observations|ConvertTo-Json -Depth 12 -Compress))
        Check 'GuideSave.NoInferredExpectationOrConclusion' ($validFirst -and $firstRecord.ExpectedConclusion.TerminalKind -ceq 'None' -and @($firstRecord.ExpectedConclusion.Steps).Count -eq 0 -and (GuideSaveControl 'lblGuideStatus' 'Label') -notmatch '(?i)conclusion observed')
        Check 'GuideSave.SuccessNamesPublishedGuideAndVersion' ($validFirst -and (GuideSaveControl 'lblGuideStatus' 'Label').Contains($guideId) -and (GuideSaveControl 'lblGuideStatus' 'Label') -match '(?i)version\s*:?\s*1\b')
        if($validFirst){CaptureSavedGuide '1'}
        $firstHash=if($validFirst){(Get-FileHash -LiteralPath $firstFile.FullName).Hash}else{''}
        [void](GuideSaveControl 'txtGuideName' 'Write' 'Revised Settings training guide')
        [void](GuideSaveControl 'btnGuideStepUp' 'Click')
        [void](GuideSaveControl 'btnSaveGuide' 'Click')
        $files=@(GuideFiles)
        $secondFile=@($files|Where-Object Name -CEQ ($guideId+'.2.json'))
        $secondRecord=if($secondFile.Count -eq 1){ReadSavedGuide $secondFile[0]}else{$null}
        $validSecond=($null -ne $secondRecord)
        Check 'GuideSave.SecondSaveAppendsSameGuideNextVersion' ($validFirst -and $validSecond -and $files.Count -eq 2 -and $secondRecord.ActionPathId -ceq $guideId -and $secondRecord.Version -eq 2 -and $secondRecord.RecordId -cne $firstRecord.RecordId)
        Check 'GuideSave.SecondVersionLinksExactPriorRecord' ($validFirst -and $validSecond -and $secondRecord.PreviousRecordId -ceq $firstRecord.RecordId -and $secondRecord.PreviousSha256 -ceq $firstRecord.ContentSha256 -and (Get-FileHash -LiteralPath $firstFile.FullName).Hash -ceq $firstHash)
        Check 'GuideSave.RevisionChangesAuthoredOrderOnly' ($validFirst -and $validSecond -and $secondRecord.Name -ceq 'Revised Settings training guide' -and $secondRecord.Steps[0].StepId -ceq $secondStep -and $secondRecord.Steps[1].StepId -ceq $firstStep -and ($secondRecord.Observations|ConvertTo-Json -Depth 12 -Compress) -ceq ($firstRecord.Observations|ConvertTo-Json -Depth 12 -Compress) -and (GuideSaveControl 'txtGuideEvidence' 'Text') -ceq $observed)
        if($validSecond){CaptureSavedGuide '2'}
        $saved=GuidePins
        $oversize='x'*1048576
        [void](GuideSaveControl 'txtGuideInstructions' 'Write' $oversize)
        [void](GuideSaveControl 'btnSaveGuide' 'Click')
        Check 'GuideSave.OversizeRejectedWithoutTruncatingDraftOrVersions' ($savePresent -and $validSecond -and (GuideSaveControl 'txtGuideInstructions' 'Text').Length -eq $oversize.Length -and (GuideSaveControl 'lblGuideStatus' 'Label') -match '(?i)1\s*MiB' -and (SamePins $saved (GuidePins)))
        [void](GuideSaveControl 'txtGuideInstructions' 'Write' 'Unsaved later edit')
        $conflictRejected=$false
        if($validSecond){
            $conflict=[IO.Path]::GetFullPath((Join-Path $guideRoot ($guideId+'.3.json')))
            if(-not $conflict.StartsWith([IO.Path]::GetFullPath($guideRoot)+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $conflict)){throw 'Invalid guide conflict fixture path.'}
            [void](New-Item -ItemType Directory -Path $conflict)
            try {
                [void](GuideSaveControl 'btnSaveGuide' 'Click')
                $conflictRejected=(Test-Path -LiteralPath $conflict -PathType Container) -and @(Get-ChildItem -LiteralPath $conflict -Force).Count -eq 0 -and
                    (SamePins $saved (GuidePins)) -and (GuideSaveControl 'txtGuideInstructions' 'Text') -ceq 'Unsaved later edit' -and
                    (GuideSaveControl 'lblGuideStatus' 'Label') -match '(?i)conflict|unavailable|could not'
            } finally {
                if((Test-Path -LiteralPath $conflict -PathType Container) -and @(Get-ChildItem -LiteralPath $conflict -Force).Count -eq 0){Remove-Item -LiteralPath $conflict -Force}
            }
        }
        Check 'GuideSave.FilenameConflictPreservesVersionsAndDraft' $conflictRejected
        [void](GuideSaveControl 'btnCancelGuide' 'Click')
        Check 'GuideSave.CancelKeepsPublishedVersionsAndDiscardsUnsavedText' ($validSecond -and (GuideSaveControl '' 'Count') -ceq '0' -and (SamePins $saved (GuidePins)))
        Check 'GuideSave.PublicationPreservesOriginalTrainingActivityAndConfig' ((SamePins $sourcePins (SourcePins)) -and (PinsRetained $activity) -and (ActivityPins).Count -eq $activity.Count -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configHash)
        Check 'GuideSave.DoesNotPublishBusinessEvents' ([long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishedBefore)
        Check 'GuideSave.NoPendingArtifactAfterAttemptedSaves' ($savePresent -and @(Get-ChildItem -LiteralPath $journalRoot -Recurse -File -Filter '*.pending').Count -eq 0)
        OpenSaveDraft
        [void](GuideSaveControl 'lstGuideSteps' 'Select' '0')
        [void](GuideSaveControl 'btnRemoveGuideStep' 'Click')
        [void](GuideSaveControl 'lstGuideSteps' 'Select' '0')
        [void](GuideSaveControl 'btnRemoveGuideStep' 'Click')
        [void](GuideSaveControl 'btnSaveGuide' 'Click')
        Check 'GuideSave.EmptyStepListRejectedWithoutLosingDraft' ($savePresent -and (GuideSaveControl '' 'Count') -ceq '1' -and (GuideSaveControl 'lstGuideSteps' 'Rows') -ceq '0' -and
            (GuideSaveControl 'lblGuideStatus' 'Label') -match '(?i)step.*required|at least one' -and (SamePins $saved (GuidePins)))
        [void](GuideSaveControl 'btnCancelGuide' 'Click')
        OpenSaveDraft
        try {
            SaveGuideVisibility $false
            [void](GuideSaveControl 'btnSaveGuide' 'Click')
            Check 'GuideSave.CurrentPolicyRestrictionPreventsRetainedDraftSave' ($savePresent -and (SamePins $saved (GuidePins)) -and (GuideSaveControl 'txtGuideEvidence' 'Text') -cin @('','MISSING'))
        } finally {
            [void](GuideSaveControl 'btnCancelGuide' 'Click')
            SaveGuideVisibility $true
        }
        OpenSaveDraft
        $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
        $authBytes=[IO.File]::ReadAllBytes($authPath)
        try {
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            try {
                $caps=Table $auth 'tblCapabilities';$revoked=0
                foreach($row in $caps.ListRows){
                    if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and
                       $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ACTION_PATH_MAINT'){
                        $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$revoked++
                    }
                }
                if($revoked -ne 1){throw 'Guide-author fixture grant is not unique.'}
                $auth.Save()
            } finally {$auth.Close($false)}
            [void](Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Fixture.Warehouse))
            $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
            $signedIn=Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'
            if($allowed -isnot [bool] -or $allowed -or $signedIn -isnot [bool] -or -not $signedIn){throw 'Capability loss fixture is not isolated from sign-out.'}
            [void](GuideSaveControl 'btnSaveGuide' 'Click')
            Check 'GuideSave.RevokedMaintenanceCapabilityPreventsSave' ($savePresent -and (SamePins $saved (GuidePins)) -and ((GuideSaveControl '' 'Count') -ceq '0' -or (GuideSaveControl 'lblGuideStatus' 'Label') -match '(?i)permission|unavailable|changed'))
        } finally {
            [void](GuideSaveControl 'btnCancelGuide' 'Click')
            [IO.File]::WriteAllBytes($authPath,$authBytes)
            # Re-sign-in creates a new session; reopen Viewer under that binding.
            CloseRecordingViewer
            SelectTarget $Fixture 'config-admin'
        }
        $restored=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
        if($restored -isnot [bool] -or -not $restored){throw 'Guide-author fixture capability was not restored.'}
        if([Convert]::ToBase64String([IO.File]::ReadAllBytes($authPath)) -cne [Convert]::ToBase64String($authBytes)){throw 'Guide-author fixture bytes were not restored.'}
        OpenRecordingViewer
        OpenSaveDraft
        $otherRoot=Join-Path $Other.Root ('Training/ActionPaths/'+$Other.Warehouse+'/Guides')
        if(Test-Path -LiteralPath $otherRoot){throw 'Other-target guide fixture must be empty.'}
        SelectTarget $Other 'config-admin'
        [void](GuideSaveControl 'btnSaveGuide' 'Click')
        Check 'GuideSave.TargetChangeCannotRedirectSave' ($savePresent -and (SamePins $saved (GuidePins)) -and -not (Test-Path -LiteralPath $otherRoot))
        Check 'GuideSave.GuardRejectionsPreserveOriginalSourceRecords' (SamePins $sourcePins (SourcePins))
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
        $closing=[IO.Path]::GetFullPath((Join-Path $journalRoot ($source.ActionPathId+'.6.json')))
        $held=$closing+'.withheld'
        $safeRoot=[IO.Path]::GetFullPath($journalRoot)+[IO.Path]::DirectorySeparatorChar
        foreach($path in @($closing,$held)){if(-not $path.StartsWith($safeRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Interrupted-source fixture path escapes its generated journal.'}}
        if(-not (Test-Path -LiteralPath $closing -PathType Leaf) -or (Test-Path -LiteralPath $held)){throw 'Interrupted-source fixture unavailable.'}
        Move-Item -LiteralPath $closing -Destination $held
        try {
            OpenRecordingViewer
            OpenSaveDraft
            Check 'GuideSave.UnclosedSourceIsLabelledInterrupted' ((GuideSaveControl 'lblGuideSource' 'Label') -match '(?i)interrupted' -and
                (GuideSaveControl 'lblGuideSource' 'Label') -notmatch '(?i);\s*recording\b')
        } finally {
            CloseRecordingViewer
            if(Test-Path -LiteralPath $closing){throw 'Interrupted-source read recreated the missing close.'}
            Move-Item -LiteralPath $held -Destination $closing
        }
        Check 'GuideSave.InterruptedFixtureRestoresExactSourceBytes' (SamePins $sourcePins (SourcePins))
    } finally {
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
    }
}
