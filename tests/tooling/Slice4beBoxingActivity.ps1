# D18 Boxing observations through actual form handlers, with independent owner
# submission/refresh evidence. Probes are installed before fixtures and compiled.
function Install-Slice4beBoxingActivityProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmShipmentsTally').CodeModule.AddFromString(@'
Public Function BoxingPrepareForTest() As String
    mPages.Value = 2: mPages_Change
    mBtnBoxMakerRefresh_Click
    If mLstBoxMakerDesigns.ListCount <> 1 Then Exit Function
    mLstBoxMakerDesigns.ListIndex = 0: mLstBoxMakerDesigns_Click
    If mLstBoxMakerComponents.ListCount <> 1 Then Exit Function
    CancelAutoSync
    BoxingPrepareForTest = mSelectedBoxMakerPackageSystemKey
End Function
Public Sub BoxingClickForTest(ByVal action As String, ByVal quantity As String)
    mTxtBoxMakerQty.Value = quantity
    Select Case action
        Case "MAKE": mBtnBoxMakerMake_Click
        Case "UNBOX": mBtnBoxMakerUnmake_Click
        Case Else: Err.Raise 5, , "Unknown Boxing test action."
    End Select
    CancelAutoSync
End Sub
Public Sub BoxingReturnToShippingForTest()
    mPages.Value = 0: mPages_Change
    mBtnRefresh_Click
    CancelAutoSync
End Sub
'@)
    $module=$project.VBComponents.Item('modTS_Shipments').CodeModule
    $module.InsertLines($module.CountOfDeclarationLines+1,'Private BoxingRefreshCompletedForTest As Boolean')
    $start=$module.ProcStartLine('CommitBoxMakerFormAction',0)
    $count=$module.ProcCountLines('CommitBoxMakerFormAction',0)
    $source=$module.Lines($start,$count)
    $anchor='    If Not batchProcessed Then batchProcessed = BoxMakerRuntimeReportShowsProcessed(runtimeReport)'
    $extracted=$anchor.Replace('BoxMakerRuntimeReportShowsProcessed','modShippingReportText.BoxMakerRuntimeReportShowsProcessed')
    if($source.Contains($extracted)){$anchor=$extracted}
    if(-not $source.Contains($anchor)){throw 'Boxing explicit refresh observer anchor unavailable.'}
    $source=$source.Replace('    syncCompletedOut = False',"    syncCompletedOut = False`r`n    BoxingRefreshCompletedForTest = False")
    $source=$source.Replace($anchor,"    BoxingRefreshCompletedForTest = batchProcessed`r`n"+$anchor)
    $module.DeleteLines($start,$count);$module.InsertLines($start,$source)
    $module.AddFromString(@'
Public Function BoxingActionForTest(ByVal action As String, Optional ByVal quantity As String = "1") As String
    If mShipmentsLauncherForm Is Nothing Then Err.Raise 5, , "Boxing fixture form missing."
    If action = "Prepare" Then
        BoxingActionForTest = mShipmentsLauncherForm.BoxingPrepareForTest()
    ElseIf action = "Shipping" Then
        mShipmentsLauncherForm.BoxingReturnToShippingForTest
    Else
        mShipmentsLauncherForm.BoxingClickForTest action, quantity
        BoxingActionForTest = "DELIVERED"
    End If
End Function
Public Function BoxingRefreshForTest() As Boolean
    BoxingRefreshForTest = BoxingRefreshCompletedForTest
End Function
'@)
}

function Test-BoxingObservation($Fixture,$Before,[string]$Label,[string]$ControlId,[string]$Caption,[string]$Outcome,$Ids,[string]$Sequence,[int]$Ordinal,[string]$SubmissionState='Submitted') {
    $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before}|ForEach-Object {[IO.File]::ReadAllText($_)|ConvertFrom-Json})
    $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
    $result=@($records|Where-Object OutcomeCode -CEQ $Outcome)
    $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $result.Count -eq 1
    Check ($Label+'.Activity.AttemptAndOwnerOutcome') $pair
    $metadata=$false;$facts=$false;$references=$false;$ordered=$false;$safe=$false
    if($pair){
        $metadata=@($records|Where-Object {$_.ControlId -cne $ControlId -or $_.Caption -cne $Caption -or
            $_.OwnerId -cne 'BOXING_WORKFLOW' -or $_.SourceRole -cne 'Boxing' -or
            $_.Surface -cne 'Operations > Shipping > Box Maker' -or $_.UserId -cne 'config-reader' -or
            $_.WarehouseId -cne $Fixture.Warehouse -or $_.StationId -cne 'S1' -or
            $_.EventCode -cne ($ControlId+'_'+$_.OutcomeCode)}).Count -eq 0
        $severity=if($Outcome -ceq 'REJECTED'){'Warning'}elseif($Outcome -ceq 'PENDING'){'Notice'}elseif($Outcome -ceq 'FAILED'){'Error'}else{'Info'}
        $effect=if($Outcome -ceq 'REJECTED'){'Unchanged'}else{'Unknown'}
        $facts=$attempt[0].Severity -ceq 'Info' -and $attempt[0].DataEffect -ceq 'Unknown' -and
            $result[0].Severity -ceq $severity -and $result[0].DataEffect -ceq $effect
        $refs=@($result[0].SourceEventRefs)
        $references=@($attempt[0].SourceEventRefs).Count -eq 0 -and $refs.Count -eq @($Ids).Count
        foreach($id in $Ids){$references=$references -and @($refs|Where-Object {$_.EventId -ceq $id -and
            $_.WarehouseId -ceq $Fixture.Warehouse -and $_.SourceKind -ceq 'Inventory' -and $_.SubmissionState -ceq $SubmissionState}).Count -eq 1}
        $ordered=$attempt[0].ActivityId -ne '' -and $attempt[0].ActivityId -ceq $result[0].ActivityId -and $attempt[0].RecordId -cne $result[0].RecordId
        if($Ordinal -gt 0){
            $ordered=$ordered -and (HasSequence ([pscustomobject]@{Attempt=$attempt[0];Outcome=$result[0]}) $Ordinal) -and $attempt[0].SequenceId -ceq $Sequence
        }else{
            foreach($record in $records){
                $sequenceProperty=$record.PSObject.Properties['SequenceId']
                if($null -ne $sequenceProperty -and -not [string]::IsNullOrEmpty([string]$sequenceProperty.Value)){$ordered=$false}
            }
        }
        $raw=$records|ConvertTo-Json -Depth 12 -Compress;$safe=$true
        foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'SHIPPING-ACTIVITY-BOX','SHIPPING-PRIVATE-DESCRIPTION','SHIPPING-FIXTURE-BIN','mBtnBoxMaker','PinHash')){
            if($raw.IndexOf($value,[StringComparison]::OrdinalIgnoreCase) -ge 0){$safe=$false}
        }
    }
    Check ($Label+'.Activity.RegisteredOwnerAndContext') $metadata
    Check ($Label+'.Activity.ExplicitOutcomeFacts') $facts
    Check ($Label+'.Activity.EverySubmittedSource') $references
    $correlation=if($Ordinal -gt 0){'OrderedRecordingOccurrence'}else{'CorrelatedUnrecordedPair'}
    Check ($Label+'.Activity.'+$correlation) $ordered
    Check ($Label+'.Activity.PrivateInputsExcluded') $safe
}

function Test-Slice4beBoxingActivity($Fixture,$Operator,$Other,$Ship,$Hold) {
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $prior=@{}
    if(Test-Path -LiteralPath $journalRoot){$prior=RestartPins $journalRoot}
    $applied=@();$ordinal=0;$displayLines=@()
    OpenRecordingViewer
    try {
        $started=(RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED'
        Check 'Boxing.Recording.ActualStart' $started
        if(-not $started){throw 'Boxing Start control was not delivered.'}
        $starts=@(Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {
            [IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json
        }|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Boxing recording start fixture is unavailable.'}
        $sequence=[string]$starts[0].SequenceId
        foreach($case in @(@('MAKE','1'),@('UNBOX','1'),@('MAKE','0'),@('UNBOX','0'))){
            $ordinal++;$action=$case[0];$valid=$case[1] -ceq '1'
            $label='Boxing.'+$action+$(if($valid){'.Accepted'}else{'.ZeroQuantity'})
            $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @('Prepare'))
            if(-not $key){throw 'Boxing fixture selection unavailable.'}
            $beforeLog=@(Get-ShippingActivityLog $Fixture)
            $before=@(Get-Slice4beActivityFiles $Fixture)
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingResetSources')
            $queue=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue'))
            $Other.Activate()
            $delivered=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @($action,$case[1]))
            Check ($label+'.ActualHandlerReturned') ($delivered -ceq 'DELIVERED')
            $ids=@(([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingReadSources')).Split("`n")|Where-Object {$_ -ne ''})
            $expectedCount=if($valid){1}else{0}
            $submissions=$ids.Count -eq $expectedCount -and
                ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue')))-$queue -eq $expectedCount -and
                [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSourceFailures') -eq 0
            Check ($label+'.IndependentSubmissionCount') $submissions
            $pin=Get-ShippingActivityHash $inventory
            $afterLog=@(Get-ShippingActivityLog $Fixture)
            Check ($label+'.EvidenceReadPreservesAuthorityBytes') ($pin -ceq (Get-ShippingActivityHash $inventory))
            $newRows=@($afterLog|Where-Object {$_.EventID -cnotin @($beforeLog.EventID)})
            $owner=$false
            if($valid -and $submissions){
                $type=if($action -ceq 'MAKE'){'BOX_BUILD'}else{'BOX_UNBOX'}
                $qty=if($action -ceq 'MAKE'){1}else{-1}
                $owner=$newRows.Count -eq 2 -and @($newRows|Where-Object {$_.EventID -cne $ids[0] -or $_.EventType -cne $type -or $_.WarehouseId -cne $Fixture.Warehouse}).Count -eq 0 -and
                    @($newRows|Where-Object {$_.System_Key -ceq $key -and [double]$_.QtyDelta -eq $qty}).Count -eq 1 -and
                    @($newRows|Where-Object {$_.System_Key -cne $key -and $_.System_Key -ne '' -and [double]$_.QtyDelta -eq -$qty}).Count -eq 1
                $applied+=@($newRows)
            }elseif(-not $valid){
                $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
                $owner=$newRows.Count -eq 0 -and $status.Contains('Box quantity must be greater than zero.')
            }
            Check ($label+'.IndependentOwnerResult') $owner
            if(-not $owner -or -not $submissions){throw 'Boxing owning fixture result not established; activity assertions are not meaningful.'}
            Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)))
            Check ($label+'.UnknownHeadersPreserved') ($null -ne $Ship.ListColumns.Item('Shipping Extra') -and $null -ne $Hold.ListColumns.Item('Shipping Extra'))
            Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel')
            $outcome='REJECTED'
            if($valid){$outcome=if([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingRefreshForTest')){'CONFIRMED'}else{'PENDING'}}
            $caption=if($action -ceq 'MAKE'){'Make Boxes'}else{'Unbox'}
            Test-BoxingObservation $Fixture $before $label ('BOXING_'+$action) $caption $outcome $ids $sequence $ordinal
            $displayLines+=@(($ordinal.ToString()+'. '+$caption+' - REQUESTED'),($ordinal.ToString()+'. '+$caption+' - '+$outcome))
            if($CaptureEvidence){Capture-BoxingFormEvidence 'Shipping Shipments' ($label.ToLowerInvariant()+'.png') ($label+'.VisibleCapture')}
        }
        $groups=@($applied|Group-Object System_Key)
        Check 'Boxing.MakeUnbox.RestoresEveryExactEntityBalance' ($applied.Count -eq 4 -and $groups.Count -eq 2 -and
            @($groups|Where-Object {($_.Group|Measure-Object QtyDelta -Sum).Sum -ne 0}).Count -eq 0)
        Check 'Boxing.Recording.ActualStop' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        Check 'Boxing.Recording.FourActionsEightObservations' ($closed.Count -eq 1 -and $closed[0].ActionCount -eq 4 -and @($closed[0].Observations).Count -eq 8)
        Check 'Boxing.Recording.CompleteIntegrityChain' (JournalChain $sequence 10)
        $script:BoxingPublicationEvidence=[pscustomobject]@{Rows=@($applied);Observations=@($closed[0].Observations);ActionPathId=[string]$closed[0].ActionPathId}
        $preserved=$true
        foreach($path in $prior.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $prior[$path]){$preserved=$false}}
        Check 'Boxing.Recording.PriorShippingJournalPreserved' $preserved
        if($CaptureEvidence){Test-BoxingVisibleRecording $Fixture $closed[0].ActionPathId $displayLines @($applied.EventID|Select-Object -Unique)}
    }finally{
        CloseRecordingViewer
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @('Shipping'))
    }
}

# Visible acceptance evidence uses only the owned generated-fixture windows.
function Capture-BoxingFormEvidence([string]$Title,[string]$File,[string]$Label) {
    Initialize-SettingsCapture
    for($attempt=1;$attempt -le 3;$attempt++){
        try{
            $handle=[InvSysSettingsCapture]::OwnedVisibleForm($Title,[IntPtr]$excel.Hwnd).ToInt64()
            if($handle -eq 0){throw 'Requested form window unavailable.'}
            $activation=New-Object -ComObject WScript.Shell
            try{[void]$activation.AppActivate($Title)}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($activation)}
            CaptureFormEvidence $Title $File $handle
            Check $Label $true
            return
        }catch{
            if($_.Exception.GetBaseException().Message -cnotin @('Requested form is not in the foreground.','Requested form window unavailable.')){throw}
            if($attempt -lt 3){Start-Sleep -Milliseconds 300}
        }
    }
    Check $Label $false
}

function Test-BoxingVisibleRecording($Fixture,[string]$PathId,$ExpectedLines,$SourceIds) {
    $journal=RestartPins $journalRoot;$activity=RestartPins $activityRoot
    function BoxingLibrary([string]$Action,[string]$Value=''){
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @($Action,$Value))
    }
    try{
        Check 'Boxing.VisualLibrary.ActualOpen' ((BoxingLibrary 'Open') -ceq 'DELIVERED')
        Check 'Boxing.VisualLibrary.ActualSelection' ((BoxingLibrary 'Select' $PathId) -ceq 'SELECTED')
        $evidence=BoxingLibrary 'Evidence';$previous=-1;$ordered=$ExpectedLines.Count -eq 8
        foreach($line in $ExpectedLines){$index=$evidence.IndexOf($line,[StringComparison]::Ordinal);$ordered=$ordered -and $index -gt $previous;$previous=$index}
        Check 'Boxing.VisualLibrary.EightOrderedObservations' $ordered
        $sources=$SourceIds.Count -eq 2
        foreach($id in $SourceIds){$sources=$sources -and $evidence.Contains('Source event: '+$id+' (Submitted; application not asserted)')}
        Check 'Boxing.VisualLibrary.ExactSourcesWithoutApplicationClaim' $sources
        Check 'Boxing.VisualLibrary.CaptureOnlyStatus' ((BoxingLibrary 'Status') -ceq 'Stopped. Capture frozen.')
        Check 'Boxing.VisualLibrary.ReadOnlyEvidence' ((BoxingLibrary 'ReadOnly') -ceq 'True')
        foreach($size in @('Default','Minimum','Larger')){
            Check ('Boxing.VisualLibrary.Layout.'+$size) ((BoxingLibrary ('Fit'+$size)) -ceq 'True')
            Capture-BoxingFormEvidence 'Action Paths' ('boxing-path-'+$size.ToLowerInvariant()+'.png') ('Boxing.VisualLibrary.Capture.'+$size)
        }
    }finally{[void](BoxingLibrary 'Close')}
    Check 'Boxing.VisualLibrary.JournalBytesPreserved' (RestartPinsEqual $journal $journalRoot)
    Check 'Boxing.VisualLibrary.ActivityBytesPreserved' (RestartPinsEqual $activity $activityRoot)
}
