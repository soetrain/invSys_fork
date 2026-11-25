# Record, author, edit and transfer under the new destination Admin's session.
function New-OnboardingGuide($case,[ref]$result) {
    $fixture=[pscustomobject]@{Warehouse=$case.Id;Root=$case.Root;Config=$case.Config}
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    SetRecordingPolicy $true -SaveValueOnly
    Write-Host 'Onboarding guide fixture: recording policy saved'
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Training recording could not start.'}
    Write-Host 'Onboarding guide fixture: Start Recording delivered'
    $first=SaveRecordedSetting '630'
    $second=SaveRecordedSetting '631'
    if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED' -or -not (HasSequence $second 2)){throw 'Training recording did not capture two ordered actions.'}
    $record=@(RecordingJournal $first.Attempt.SequenceId | Sort-Object Version)
    $pathId=$record[0].ActionPathId
    Write-Host 'Onboarding guide fixture: stopped recording located'
    if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $pathId) -cne 'SELECTED'){throw 'Saved Training recording could not be selected.'}
    $editor='frmActionPathGuide'
    [void](BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
    if((BoundControl '' 'Count' '' $editor) -cne '1'){throw 'Authorized guide editor unavailable.'}
    Check 'Onboarding.Author.SelectedRecordingOpensEditor' $true
    [void](BoundControl 'txtGuideSubject' 'Write' 'Training basics' $editor)
    [void](BoundControl 'txtGuideName' 'Write' 'Training guide' $editor)
    [void](BoundControl 'btnSaveGuide' 'Click' '' $editor)
    $root=Join-Path $journalRoot 'Guides'
    $files=@(if(Test-Path -LiteralPath $root){Get-ChildItem -LiteralPath $root -File -Filter '*.json'})
    $firstGuide=if($files.Count -eq 1){ReadGuideExpectationRecord $files[0].FullName}
    Check 'Onboarding.Author.PublishesRecordedTrainingGuide' ($null -ne $firstGuide -and $firstGuide.SourceRun.ActionPathId -ceq $pathId)
    if(-not $firstGuide){throw 'Initial publication missing.'}
    [void](BoundControl 'btnCancelGuide' 'Click' '' $editor)
    BoundOpen
    BoundSelect $firstGuide
    [void](BoundControl 'btnEditPublishedGuide' 'Click')
    [void](BoundControl 'txtGuideInstructions' 'Write' 'Revised Training instruction' $editor)
    [void](BoundControl 'btnSaveGuide' 'Click' '' $editor)
    [void](BoundControl 'btnCancelGuide' 'Click' '' $editor)
    $revision=@(Get-ChildItem -LiteralPath $root -Filter '*.json' -File | ForEach-Object {ReadGuideExpectationRecord $_.FullName} | Where-Object Version -eq 2)
    Check 'Onboarding.Edit.PublishesSecondVersion' ($revision.Count -eq 1)
    if($revision.Count -ne 1){throw 'Edited publication missing.'}
    [void](BoundControl 'btnRefreshGuides' 'Click')
    BoundSelect $revision[0]
    $export=Join-Path $runRoot 'onboarding-guide-export.json'
    TransferSetFile $export
    [void](BoundControl 'btnExportGuide' 'Click')
    $exported=Test-Path -LiteralPath $export -PathType Leaf
    Check 'Onboarding.Transfer.ExportsExactEditedVersion' ($exported -and ([IO.File]::ReadAllText($export)|ConvertFrom-Json).Guide.ContentSha256 -ceq $revision[0].ContentSha256)
    if(-not $exported){throw 'Guide export unavailable.'}
    [void](BoundControl 'btnImportGuide' 'Click')
    Check 'Onboarding.Transfer.ImportPublishesNewGuide' (@(Get-ChildItem -LiteralPath $root -File -Filter '*.json').Count -eq 3)
    TransferSetFile ''
    $result.Value=[pscustomobject]@{Root=$root;Revision=$revision[0];Export=$export}
}

function Test-OpenGuideRevocation($guide,[scriptblock]$setPermission) {
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideInputNative.ps1')
    function Editor($control,$action,$value=''){BoundControl $control $action $value 'frmActionPathGuide'}
    $original=BoundPins $guide.Root
    foreach($operation in @('Remove','Save','Cancel','WindowX')){
        & $setPermission $true
        [void](BoundControl 'btnRefreshGuides' 'Click')
        BoundSelect $guide.Revision
        [void](BoundControl 'btnEditPublishedGuide' 'Click')
        if((Editor '' 'Count') -cne '1' -or (Editor '' 'Place') -cne 'PLACED' -or (Editor 'txtGuideName' 'Focus') -cne 'FOCUSED'){throw 'Live editor focus unavailable.'}
        [void](Run 'invSys.Core.xlam' 'TestGuideInputWriteCount.ResetCount')
        if((Editor 'txtGuideName' 'Type' 'Unpublished text') -cne 'DELIVERED'){throw 'Dirty guide entry failed.'}
        & $setPermission $false
        switch($operation){
            'Cancel' {[void](Click-GuideCancelNative)}
            'WindowX' {[void](Close-GuideWindowXNative)}
            'Save' {[void](Editor 'btnSaveGuide' 'Click')}
            'Remove' {[void](Editor 'btnRemoveGuideStep' 'Click')}
        }
        $attempts=[long](Run 'invSys.Core.xlam' 'TestGuideInputWriteCount.ReadCount')
        if($operation -in @('Cancel','WindowX')){
            Check "Onboarding.OpenEditor.Dirty${operation}DiscardsWithoutWrite" ($attempts -eq 0 -and (Editor '' 'Count') -ceq '0' -and (BoundSame $original (BoundPins $guide.Root)))
        }else{
            Check "Onboarding.OpenEditor.RevokedFlushRefuses$operation" ($attempts -eq 1 -and (Editor 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $original (BoundPins $guide.Root)))
            [void](Editor 'btnCancelGuide' 'Click')
        }
    }
    $refusedExport=Join-Path $runRoot 'revoked-guide-export.json'
    TransferSetFile $refusedExport
    [void](BoundControl 'btnExportGuide' 'Click')
    Check 'Onboarding.RevokedExport.NoFile' (-not (Test-Path -LiteralPath $refusedExport))
    TransferSetFile $guide.Export
    [void](BoundControl 'btnImportGuide' 'Click')
    Check 'Onboarding.RevokedImport.NoPublication' (BoundSame $original (BoundPins $guide.Root))
    TransferSetFile ''
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
}
