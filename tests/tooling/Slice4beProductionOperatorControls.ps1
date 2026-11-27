# Each Start captures its own exact RunId. Only the positive case may gain credit.
function Test-OperatorBindingControls {
    $books=[Collections.Generic.List[object]]::new()
    function PrepareBook {
        $book=$excel.Workbooks.Add()
        $books.Add($book)
        [void](Probe 'RunLocalReopen' @($book.Name))
        if($Mode -ceq 'Reusable'){$ready=[bool](Probe 'CompleteBaselinePrepare' @($true))}
        else{
            $ready=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CompleteWorksheetInventoryForTest' @($book.Name))
            if($ready){$ready=[bool](Probe 'CompleteWorksheetPrepare' @($contextName))}
        }
        if(!$ready){throw 'Negative control needs actual prepared and checked-in inventory.'}
        [void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))
    }
    function OpenAndBind {
        BoundSelect $guide
        Delivered (BoundControl 'btnRunHowTo' 'Click')
        StartCase
        ClickRun 'btnBindProduction'
        $history=LastRun
        if($history.OperatorClaims.Count -ne 1 -or $history.OperatorClaims[0].Kind -cne $kind -or $history.OperatorClaims[0].ActivityId){throw 'Fresh pending binding missing.'}
    }
    function OrdinaryClick {
        [void](Probe 'CompleteSubmissionArm' @('','','ObserveOnly'))
        if(-not [bool](Probe 'CompleteEntryAct' @(''))){throw 'Ordinary button unavailable.'}
    }
    function VerifySaved {
        $evalRoot=Join-Path $destRoot 'Evaluations'
        $before=BoundPins $evalRoot
        ClickRun 'btnVerifyRun'
        $new=@(Get-ChildItem $evalRoot -File -Filter '*.json'|Where-Object {!$before.ContainsKey($_.FullName)})
        if($new.Count -ne 1){throw 'Exactly one saved Verify result required.'}
        Get-Content $new[0].FullName -Raw|ConvertFrom-Json
    }
    try{
        ClickRun 'btnCloseRun'
        PrepareBook;OpenAndBind
        $stopping=LastRun
        ClickRun 'btnStopRun'
        $stopped=LastRun
        $recordingPins=BoundPins (Join-Path $destRoot 'Runs')
        $before=@(Get-Slice4beActivityFiles $dest)
        OrdinaryClick
        $later=@(Get-Slice4beActivityFiles $dest|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object {$_.ControlId -ceq $control -and $_.OutcomeCode -ceq 'CONFIRMED'})
        Pass 'StopExcludesLaterOrdinaryCompletion' ($later.Count -eq 1 -and $later[0].SourceEventRefs.Count -eq 2 -and $stopped.State -ceq 'Stopped' -and $stopped.RunId -ceq $stopping.RunId -and !$stopped.OperatorClaims[0].ActivityId -and $stopped.Steps.Count -eq 0 -and (BoundSame $recordingPins (BoundPins (Join-Path $destRoot 'Runs'))))
        ClickRun 'btnCloseRun'
        PrepareBook;OpenAndBind
        $original=LastRun
        # Replacing the Production form and workbook is a legitimate ordinary
        # action. The new owner must succeed without adopting the old binding.
        PrepareBook
        $before=@(Get-Slice4beActivityFiles $dest)
        OrdinaryClick
        $rows=@(Get-Slice4beActivityFiles $dest|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
        $other=@($rows|Where-Object {$_.ControlId -ceq $control -and $_.OutcomeCode -ceq 'CONFIRMED'})
        if($other.Count -ne 1 -or $other[0].SequenceId -cne $original.Recording.SequenceId){throw 'A successful unrelated completion in this recording is required.'}
        $pin=Hash $inventoryPath
        ClickRun 'btnCheckProduction'
        $blocked=LastRun
        Pass 'DifferentArtifactCompletedWithoutCredit' ($blocked.State -ceq 'Blocked' -and $blocked.Steps.Count -eq 0 -and !$blocked.OperatorClaims[0].ActivityId -and $blocked.OperatorClaims[0].BindingId -ceq $original.OperatorClaims[0].BindingId)
        Delivered (BoundControl 'btnRefresh' 'Click' '' 'frmInventoryViewer')
        $verification=VerifySaved
        Pass 'VerifyExcludesUnclaimedSuccessInSameRecording' ($verification.ResultState -ceq 'Failed' -and 'REQUIRED_STEP_MISSING' -cin $verification.ReasonCodes -and $other[0].ActivityId -cin $verification.ExtraActivityIds -and (Hash $inventoryPath) -ceq $pin)
        ClickRun 'btnCloseRun'
        PrepareBook;OpenAndBind
        $pending=LastRun
        ClickRun 'btnCloseRun'
        $closed=LastRun
        BoundSelect $guide;Delivered (BoundControl 'btnRunHowTo' 'Click')
        $opened=LastRun
        Pass 'CloseReopenPreservesStoppedBinding' ($closed.State -ceq 'Stopped' -and $closed.OperatorClaims[0].BindingId -ceq $pending.OperatorClaims[0].BindingId -and $opened.ContentSha256 -ceq $closed.ContentSha256 -and $opened.Steps.Count -eq 0)
        ClickRun 'btnCloseRun'
        PrepareBook;OpenAndBind
        [void](Run 'invSys.Operations.xlam' 'TestCompleteReplay.LoseNextNotification')
        OrdinaryClick
        $unobserved=LastRun
        ClickRun 'btnStopRun'
        $afterStop=LastRun
        Pass 'StopAfterLostNotificationReleasesReturnedOwner' ($unobserved.OperatorClaims[0].OwnerValidated -and $afterStop.State -ceq 'Stopped' -and $afterStop.Steps.Count -eq 0 -and $afterStop.OperatorClaims[0].ActivityId -ceq $unobserved.OperatorClaims[0].ActivityId)
        if($afterStop.State -ceq 'Running'){
            # Deliver the skipped notification only for cleanup of a failing candidate.
            [void](Run 'invSys.Operations.xlam' 'TestCompleteReplay.ReleaseForTest' @($unobserved.OperatorClaims[0].ActivityId))
        }
        ClickRun 'btnCloseRun'
        PrepareBook;OpenAndBind
        $pin=Hash $inventoryPath
        $before=@(Get-Slice4beActivityFiles $dest)
        [void](Probe 'ClearEnteredActual')
        OrdinaryClick
        $rows=@(Get-Slice4beActivityFiles $dest|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
        $refusal=@($rows|Where-Object {$_.ControlId -ceq $control -and $_.OutcomeCode -ceq 'REJECTED'})
        ClickRun 'btnCheckProduction'
        $refused=LastRun
        Pass 'OwnerRefusalPreservedWithoutInventoryEffectOrCredit' ($refusal.Count -eq 1 -and $refusal[0].SequenceId -ceq $refused.Recording.SequenceId -and $refused.Steps.Count -eq 0 -and !$refused.OperatorClaims[0].ActivityId -and $refused.State -ceq 'Blocked' -and (Hash $inventoryPath) -ceq $pin)
        ClickRun 'btnCloseRun'
        PrepareBook;OpenAndBind
        $old=LastRun
        SelectTarget $Source
        $reply=BoundControl 'btnCheckProduction' 'Click' '' $runner
        if($reply -cnotin @('DELIVERED','MISSING','HIDDEN','DISABLED')){throw 'Unexpected change of session control result.'}
        CloseRecordingViewer
        $revoked=LastRun
        Pass 'DifferentSessionAndWarehouseCannotRetarget' ($revoked.WarehouseId -ceq $dest.Warehouse -and $old.RunId -ceq $revoked.RunId -and !$revoked.OperatorClaims[0].ActivityId -and $revoked.Steps.Count -eq 0 -and $old.ContextSha256 -ceq $revoked.ContextSha256 -and (Hash $inventoryPath) -ceq $pin -and (Hash $sourceInventory) -ceq $sourceHash)
        SelectTarget $dest 'config-producer'
    }finally{
        [void](Probe 'CompleteSubmissionReset')
        [void](Probe 'CloseDesigner')
        foreach($book in $books){$book.Close($false)}
    }
}
