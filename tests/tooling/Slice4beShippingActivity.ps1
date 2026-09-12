# D18 test-first discovery. Only unsaved handler facades; no auth bypass or
# activity fabrication. Generated fixture values remain in memory only.
function Get-ShippingActivityRows($List) {
    foreach($row in $List.ListRows) {
        $values=@{}
        foreach($column in $List.ListColumns){$values[$column.Name]=$row.Range.Cells.Item(1,$column.Index).Value2}
        [pscustomobject]$values
    }
}
function Get-ShippingActivityHash([string]$Path) {
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $sha=[Security.Cryptography.SHA256]::Create()
    try{[BitConverter]::ToString($sha.ComputeHash($stream))}finally{$sha.Dispose();$stream.Dispose()}
}
function Get-ShippingActivityLog($Fixture) {
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$null;$opened=$false
    foreach($candidate in $excel.Workbooks){if($candidate.FullName -ieq $path){$book=$candidate;break}}
    if($null -eq $book){$book=$excel.Workbooks.Open($path,0,$true);$opened=$true}
    try{Get-ShippingActivityRows (Table $book 'tblInventoryLog')}finally{if($opened){$book.Close($false)}}
}
function Test-ShippingActivityPair($Fixture,$Before,[string]$Caption,[string]$Label,$AppliedIds) {
    $records=@();$raws=@()
    foreach($path in @(Get-Slice4beActivityFiles $Fixture)) {
        if($path -in $Before){continue}
        $raw=[IO.File]::ReadAllText($path);$r=$raw|ConvertFrom-Json
        if((Get-Slice4beField $r 'SourceRole') -ceq 'Shipping' -and (Get-Slice4beField $r 'Caption') -ceq $Caption){$records+=$r;$raws+=$raw}
    }
    $attempt=@($records|Where-Object {(Get-Slice4beField $_ 'OutcomeCode') -ceq 'REQUESTED'})
    $result=@($records|Where-Object {(Get-Slice4beField $_ 'OutcomeCode') -cne 'REQUESTED'})
    $pair=$attempt.Count -eq 1 -and $result.Count -eq 1
    Check ($Label+'.Activity.AttemptAndOutcome') $pair
    $correlated=$false;$owned=$false
    if($pair){
        $correlated=(Get-Slice4beField $attempt[0] 'ActivityId') -ne '' -and $attempt[0].ActivityId -ceq $result[0].ActivityId -and $attempt[0].RecordId -cne $result[0].RecordId -and $attempt[0].ControlId -ne '' -and $attempt[0].ControlId -ceq $result[0].ControlId
        $owned=$true
        foreach($r in $records){$owned=$owned -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.UserId -ceq 'config-reader' -and $r.StationId -ceq 'S1' -and $r.OwnerId -ne '' -and $r.EventCode -ne ''}
    }
    Check ($Label+'.Activity.StableCorrelation') $correlated
    Check ($Label+'.Activity.TrustedOwnerContext') $owned
    $references=$false
    if($pair){
        $property=$result[0].PSObject.Properties['SourceEventRefs']
        if($null -ne $property){
            $refs=@($property.Value);$ids=@($refs|ForEach-Object EventId)
            $references=$ids.Count -eq $AppliedIds.Count -and @($ids|Select-Object -Unique).Count -eq $ids.Count
            foreach($id in $AppliedIds){$references=$references -and $id -cin $ids}
            foreach($ref in $refs){$references=$references -and $ref.WarehouseId -ceq $Fixture.Warehouse -and $ref.SourceKind -ceq 'Inventory' -and $ref.SubmissionState -ceq 'Submitted'}
        }
    }
    Check ($Label+'.Activity.ExactAppliedSourceReferences') $references
    $safe=$pair
    foreach($raw in $raws){foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'SHIPPING-PRIVATE-REFERENCE','SHIPPING-PRIVATE-CARRIER','PinHash','Err.Description','mBtnAdd_Click')){if($raw.IndexOf($value,[StringComparison]::OrdinalIgnoreCase) -ge 0){$safe=$false}}}
    Check ($Label+'.Activity.InputValuesExcluded') $safe
}
function Test-Slice4beShippingActivity {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmShipmentsTally').CodeModule
    # Intercept only existing report presentation, retaining the real handlers.
    $source=$form.Lines(1,$form.CountOfLines)
    $source=$source.Replace('MsgBox report,','ActivityShippingNotice report,')
    $form.DeleteLines(1,$form.CountOfLines);$form.AddFromString($source)
    $form.AddFromString(@'
Private Sub ActivityShippingNotice(ByVal report As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "")
End Sub
Public Function ActivityShippingWorkbook() As String
    ActivityShippingWorkbook = mOperatorWorkbook.Name
End Function
Public Sub ActivityShippingSetQuantity(ByVal value As String)
    mTxtQty.Value = value
End Sub
Public Function ActivityShippingStatus() As String
    ActivityShippingStatus = NzText(mTxtStatus.Value)
End Function
Public Function ActivityShippingCreateBox() As Boolean
    Dim i As Long, chosen As Boolean
    mBtnRefresh_Click
    mPages.Value = 1: mPages_Change
    mBtnBoxBuilderRefresh_Click
    mBtnBoxBuilderNew_Click
    For i = 0 To mLstBoxBuilderInventory.ListCount - 1
        If ParseNumber(NzText(mLstBoxBuilderInventory.List(i, 6))) >= 10 Then
            mLstBoxBuilderInventory.ListIndex = i: chosen = True: Exit For
        End If
    Next i
    If Not chosen Then Exit Function
    mTxtBoxBuilderName.Value = "SHIPPING-ACTIVITY-BOX"
    mTxtBoxBuilderUom.Value = "EA"
    mTxtBoxBuilderLocation.Value = "SHIPPING-FIXTURE-BIN"
    mTxtBoxBuilderDescription.Value = "SHIPPING-PRIVATE-DESCRIPTION"
    mTxtBoxBuilderComponentQty.Value = "1"
    mBtnBoxBuilderAddComponent_Click
    If mLstBoxBuilderComponents.ListCount <> 1 Then Exit Function
    mBtnBoxBuilderSave_Click
    mPages.Value = 2: mPages_Change
    mBtnBoxMakerRefresh_Click
    If mLstBoxMakerDesigns.ListCount <> 1 Then Exit Function
    mLstBoxMakerDesigns.ListIndex = 0: mLstBoxMakerDesigns_Click
    If mLstBoxMakerComponents.ListCount <> 1 Then Exit Function
    mTxtBoxMakerQty.Value = "10"
    mBtnBoxMakerMake_Click
    mPages.Value = 0: mPages_Change
    mOperatorWorkbook.Activate
    mChkUseExisting.Value = True: mChkUseExisting_Click
    mBtnRefresh_Click
    CancelAutoSync
    For i = 0 To mLstShippables.ListCount - 1
        If ParseNumber(NzText(mLstShippables.List(i, 3))) >= 10 Then ActivityShippingCreateBox = True
    Next i
End Function
Public Function ActivityShippingPrepare(ByVal action As String) As String
    Dim i As Long, chosen As Boolean
    CancelAutoSync
    If action = "Add" Or action = "AddAgain" Then
        For i = 0 To mLstShipments.ListCount - 1: mLstShipments.Selected(i) = False: Next i
        For i = 0 To mLstShippables.ListCount - 1
            If ParseNumber(NzText(mLstShippables.List(i, 3))) >= 5 Then
                mLstShippables.ListIndex = i: mLstShippables_Click: chosen = True: Exit For
            End If
        Next i
        If Not chosen Then Exit Function
        mTxtRef.Value = "SHIPPING-PRIVATE-REFERENCE"
        mTxtCarrier.Value = "SHIPPING-PRIVATE-CARRIER"
        mTxtQty.Value = "2"
    ElseIf action = "Return" Then
        If mLstHold.ListCount <> 1 Then Exit Function
        mLstHold.ListIndex = 0: mLstHold.Selected(0) = True: mLstHold_Click
    Else
        If mLstShipments.ListCount <> 1 Then Exit Function
        mLstShipments.ListIndex = 0: mLstShipments.Selected(0) = True: mLstShipments_Click
        If action = "Update" Then mTxtQty.Value = "3"
    End If
    ActivityShippingPrepare = NzText(mTxtSystemKey.Value)
End Function
Public Sub ActivityShippingClick(ByVal action As String)
    Select Case action
        Case "Add", "AddAgain": mBtnAdd_Click
        Case "Update": mBtnUpdate_Click
        Case "Remove": mBtnRemove_Click
        Case "Hold": mBtnHold_Click
        Case "Return": mBtnReturn_Click
        Case "Stage": mBtnStage_Click
        Case "Send": mBtnSend_Click
        Case Else: Err.Raise 5, , "Unknown fixture action."
    End Select
    CancelAutoSync
End Sub
Public Function ActivityShippingBound(ByVal expected As Workbook) As Boolean
    ActivityShippingBound = (mOperatorWorkbook Is expected)
End Function
'@)
    $module=$project.VBComponents.Item('modTS_Shipments').CodeModule
    # Count entry to the existing owner/submission boundaries without replacing
    # their logic. No payload, event identity or credential enters the counters.
    $source=$module.Lines(1,$module.CountOfLines)
    $source=$source.Replace('Option Explicit',"Option Explicit`r`nPublic ActivityShippingOwnerEntries As Long`r`nPublic ActivityShippingQueueEntries As Long")
    foreach($entry in @(
        @{Name='ShipmentsFormCommitLine';Counter='ActivityShippingOwnerEntries'},
        @{Name='QueueShippingPayloadEventServerFirst';Counter='ActivityShippingQueueEntries'}
    )) {
        $pattern='(?s)((?:Public|Private) Function '+$entry.Name+'\(.*?\) As Boolean\s*\r?\n)(\s*On Error)'
        if([regex]::Matches($source,$pattern).Count -ne 1){throw 'Shipping boundary observer anchor unavailable.'}
        $replacement='$1'+'    '+$entry.Counter+' = '+$entry.Counter+" + 1`r`n"+'$2'
        $source=[regex]::Replace($source,$pattern,$replacement)
    }
    $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($source)
    $module.AddFromString(@'
Public Function ActivityShippingBoundaryCount(ByVal boundary As String) As Long
    If boundary = "Owner" Then
        ActivityShippingBoundaryCount = ActivityShippingOwnerEntries
    ElseIf boundary = "Queue" Then
        ActivityShippingBoundaryCount = ActivityShippingQueueEntries
    Else
        Err.Raise 5, , "Unknown fixture boundary."
    End If
End Function
Public Function ActivityShippingOpen() As String
    BtnOpenShipmentsForm
    If mShipmentsLauncherForm Is Nothing Then Exit Function
    mShipmentsLauncherForm.CancelAutoSync
    ActivityShippingOpen = mShipmentsLauncherForm.ActivityShippingWorkbook()
End Function
Public Function ActivityShippingCreateBox() As Boolean
    ActivityShippingCreateBox = mShipmentsLauncherForm.ActivityShippingCreateBox()
End Function
Public Function ActivityShippingPrepare(ByVal action As String) As String
    ActivityShippingPrepare = mShipmentsLauncherForm.ActivityShippingPrepare(action)
End Function
Public Sub ActivityShippingSetQuantity(ByVal value As String)
    mShipmentsLauncherForm.ActivityShippingSetQuantity value
End Sub
Public Function ActivityShippingStatus() As String
    ActivityShippingStatus = mShipmentsLauncherForm.ActivityShippingStatus()
End Function
Public Sub ActivityShippingClick(ByVal action As String)
    mShipmentsLauncherForm.ActivityShippingClick action
End Sub
Public Function ActivityShippingBound(ByVal workbookName As String) As Boolean
    ActivityShippingBound = mShipmentsLauncherForm.ActivityShippingBound(Application.Workbooks(workbookName))
End Function
Public Sub ActivityShippingClose()
    If Not mShipmentsLauncherForm Is Nothing Then Unload mShipmentsLauncherForm
    Set mShipmentsLauncherForm = Nothing
    Set mShipmentsAutoSyncForm = Nothing
    mShipmentsLauncherWorkbookName = ""
End Sub
'@)
    $fixture=NewFixture 'shipping-activity'
    $auth=$excel.Workbooks.Open((Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Auth.xlsb')),0,$false)
    $caps=Table $auth 'tblCapabilities';$row=$caps.ListRows.Add()
    foreach($pair in @{UserId='config-reader';Capability='SHIP_POST';WarehouseId=$fixture.Warehouse;StationId='S1';Status='Active'}.GetEnumerator()){$row.Range.Cells.Item(1,$caps.ListColumns.Item($pair.Key).Index).Value2=$pair.Value}
    $auth.Save();$auth.Close($false)
    SelectTarget $fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Shipping generated fixture seed failed.'}
    SelectTarget $fixture 'config-reader'
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $repo 'deploy/current/templates')))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'shipping-operators')))
    $other=$excel.Workbooks.Add();$other.Worksheets.Item(1).Cells.Item(1,1).Value2='shipping unrelated sentinel'
    $other.SaveAs((Join-Path $runRoot 'shipping-unrelated.xlsm'),52)
    $otherHash=Get-ShippingActivityHash $other.FullName
    $configHash=Get-ShippingActivityHash $fixture.Config
    $operator=$null;$dialogJob=$null
    $dialogStop=Join-Path $runRoot 'shipping-dialog-observer.stop'
    try {
        $other.Activate()
        $name=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingOpen')
        if(-not $name){throw 'Public Shipping launcher did not establish its form.'}
        $operator=$excel.Workbooks.Item($name)
        Check 'Shipping.PublicLauncher.Captured' ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($name)))
        $observerPath=Join-Path $repo 'tools/plan022-dialog-observer.ps1'
        . $observerPath
        if(-not ('Plan022NativeDialogs' -as [type])){Invoke-Plan022NativeDialogObservation -ProcessId 0 -TimeoutSeconds 0}
        [uint32]$ownedProcess=0
        [void][Plan022NativeDialogs]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$ownedProcess)
        if($ownedProcess -eq 0){throw 'Owned Shipping fixture process unavailable.'}
        $dialogJob=Start-Job -ArgumentList $observerPath,$ownedProcess,$dialogStop -ScriptBlock {
            param($observer,$owned,$stop)
            . $observer
            # Existing observer accepts only one owned informational OK. Do not
            # persist or return the native dialog's arbitrary text or values.
            Invoke-Plan022NativeDialogObservation -ProcessId $owned -TimeoutSeconds 180 -StopPath $stop | Out-Null
        }
        $created=[bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingCreateBox')
        Check 'Shipping.Setup.BoxDesignerAndMakerHandlers' $created
        if(-not $created){throw 'Shipping fixture box creation was not established through the form handlers.'}
        $ship=Table $operator 'ShipmentsTally';$hold=Table $operator 'NotShipped'
        foreach($table in @($ship,$hold)){$column=$table.ListColumns.Add();$column.Name='Shipping Extra'}
        $operator.Save()
        $inventoryPath=Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        $lastKey=''
        $captions=@{Add='Add';AddAgain='Add';Update='Update Row';Remove='Remove';Hold='Send Hold';Return='Return';Stage='To Shipments';Send='Shipments Sent'}
        foreach($action in @('Add','Update','Hold','Return','Remove','AddAgain','Stage','Send')) {
            $label='Shipping.'+$action
            $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @($action))
            if(-not $key){throw ('Shipping fixture preparation failed: '+$action)}
            $lastKey=$key
            if($action -in @('Update','Stage')){$ship.ListColumns.Item('Shipping Extra').DataBodyRange.Value2='preserve shipping extension'}
            $beforeLog=@(Get-ShippingActivityLog $fixture)
            $before=@(Get-Slice4beActivityFiles $fixture)
            $other.Activate()
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @($action))
            $rows=@(Get-ShippingActivityRows $ship);$held=@(Get-ShippingActivityRows $hold)
            $live=@($rows|Where-Object {[double]$_.QUANTITY -gt 0});$heldLive=@($held|Where-Object {[double]$_.QUANTITY -gt 0})
            $expectedQty=if($action -in @('Update','Hold','Return')){3}else{2}
            $effect=if($action -eq 'Hold'){$live.Count -eq 0 -and $heldLive.Count -eq 1 -and $heldLive[0].System_Key -ceq $key -and [double]$heldLive[0].QUANTITY -eq $expectedQty}elseif($action -in @('Remove','Send')){$live.Count -eq 0 -and $heldLive.Count -eq 0}else{$live.Count -eq 1 -and $heldLive.Count -eq 0 -and $live[0].System_Key -ceq $key -and [double]$live[0].QUANTITY -eq $expectedQty}
            Check ($label+'.OwnerStagingResult') $effect
            if(-not $effect){throw ('Shipping owner staging result not established: '+$action)}
            Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($name)))
            Check ($label+'.UnknownHeadersPreserved') ($null -ne $ship.ListColumns.Item('Shipping Extra') -and $null -ne $hold.ListColumns.Item('Shipping Extra'))
            if($action -in @('Update','Stage')){Check ($label+'.UnknownValuePreserved') ($live[0].'Shipping Extra' -ceq 'preserve shipping extension')}
            Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel' -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
            $beforeReadHash=Get-ShippingActivityHash $inventoryPath
            $afterLog=@(Get-ShippingActivityLog $fixture)
            Check ($label+'.EvidenceReadPreservedAuthorityBytes') ($beforeReadHash -ceq (Get-ShippingActivityHash $inventoryPath))
            $newLog=@($afterLog|Where-Object {$_.EventID -cnotin @($beforeLog.EventID)})
            $appliedIds=@($newLog|ForEach-Object EventID|Select-Object -Unique)
            Check ($label+'.SourceEvidenceHasExactWarehouseAndKey') (@($newLog|Where-Object {$_.WarehouseId -cne $fixture.Warehouse -or $_.System_Key -cne $key -or [string]::IsNullOrWhiteSpace($_.EventID)}).Count -eq 0)
            Test-ShippingActivityPair $fixture $before $captions[$action] $label $appliedIds
        }
        $terminal=@(Get-ShippingActivityLog $fixture)
        $shipEvents=@($terminal|Where-Object {$_.EventType -ceq 'SHIP' -and $_.System_Key -ceq $lastKey})
        Check 'Shipping.Domain.ExactShipmentAppliedOnce' ($shipEvents.Count -eq 1 -and [double]$shipEvents[0].QtyDelta -eq -2 -and $shipEvents[0].WarehouseId -ceq $fixture.Warehouse)
        $keyRows=@($terminal|Where-Object {$_.System_Key -ceq $lastKey})
        $balance=($keyRows|Measure-Object -Property QtyDelta -Sum).Sum
        Check 'Shipping.Domain.BoxQuantityReconciled' ($balance -eq 8 -and @($keyRows|Where-Object {$_.EventType -ceq 'BOX_BUILD' -and [double]$_.QtyDelta -eq 10}).Count -eq 1)
        Check 'Shipping.BoundaryObservers.CalibratedByNormalActions' ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Owner')) -gt 0 -and [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue')) -gt 0)
        # Negative cases follow the preserved normal sequence. Read values only
        # in memory; reports contain fixed assertion names and booleans.
        $sessionModule=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
        $sessionModule.Name='TestShippingSession'
        $sessionModule.CodeModule.AddFromString(@'
Public Function SessionVersion() As Long
    SessionVersion = modAuthSession.Version()
End Function
'@)
        foreach($case in @('InvalidQuantity','ReauthenticatedSession')) {
            $label='Shipping.'+$case
            $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @('Add'))
            if(-not $key){throw 'Shipping negative fixture selection unavailable.'}
            if($case -eq 'InvalidQuantity') {
                [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSetQuantity' @('0'))
            } else {
                $sessionBefore=[long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion')
                [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
                $signed=[string](Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @('config-reader',$fixture.Secret,''))
                $changed=$signed.StartsWith('OK|') -and [long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion') -ne $sessionBefore -and [string](Run 'invSys.Core.xlam' 'modAuth.GetCurrentUserId') -ceq 'config-reader'
                Check ($label+'.FixtureSameActorNewSession') $changed
                if(-not $changed){throw 'Shipping reauthentication fixture was not established.'}
            }
            $beforeRows=@(Get-ShippingActivityRows $ship)|ConvertTo-Json -Depth 5 -Compress
            $beforeHeld=@(Get-ShippingActivityRows $hold)|ConvertTo-Json -Depth 5 -Compress
            $beforeLog=@(Get-ShippingActivityLog $fixture)
            $before=@(Get-Slice4beActivityFiles $fixture)
            $ownerBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Owner'))
            $queueBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue'))
            $other.Activate()
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @('Add'))
            $ownerAfter=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Owner'))
            $queueAfter=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue'))
            Check ($label+'.NoSubmissionBoundaryEntered') ($queueBefore -eq $queueAfter)
            $afterRows=@(Get-ShippingActivityRows $ship)|ConvertTo-Json -Depth 5 -Compress
            $afterHeld=@(Get-ShippingActivityRows $hold)|ConvertTo-Json -Depth 5 -Compress
            Check ($label+'.StagingValuesPreserved') ($beforeRows -ceq $afterRows -and $beforeHeld -ceq $afterHeld)
            $beforeReadHash=Get-ShippingActivityHash $inventoryPath
            $afterLog=@(Get-ShippingActivityLog $fixture)
            Check ($label+'.EvidenceReadPreservedAuthorityBytes') ($beforeReadHash -ceq (Get-ShippingActivityHash $inventoryPath))
            Check ($label+'.NoCanonicalEventApplied') (($beforeLog|ConvertTo-Json -Depth 5 -Compress) -ceq ($afterLog|ConvertTo-Json -Depth 5 -Compress))
            Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($name)))
            Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel' -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
            if($case -eq 'InvalidQuantity') {
                $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
                Check ($label+'.OwnerValidationReached') ($status.Contains('Quantity must be greater than zero.'))
                Test-ShippingActivityPair $fixture $before 'Add' $label @()
                $outcomes=@(Get-Slice4beActivityFiles $fixture|Where-Object {$_ -notin $before}|ForEach-Object {[IO.File]::ReadAllText($_)|ConvertFrom-Json}|Where-Object {$_.SourceRole -ceq 'Shipping' -and $_.Caption -ceq 'Add' -and $_.OutcomeCode -ceq 'REJECTED'})
                Check ($label+'.Activity.RejectedWithNoSourceReferences') ($outcomes.Count -eq 1 -and @($outcomes[0].SourceEventRefs).Count -eq 0)
            } else {
                # A new sign-in must not revive the old form or attribute a
                # stale action to that new session, even for the same user.
                Check ($label+'.StopsBeforeStagingOwner') ($ownerBefore -eq $ownerAfter)
                Check ($label+'.NoActivityAttributedToNewSession') (@(Get-Slice4beActivityFiles $fixture|Where-Object {$_ -notin $before}).Count -eq 0)
            }
        }
        Check 'Shipping.ConfigBytesPreserved' ($configHash -ceq (Get-ShippingActivityHash $fixture.Config))
    } finally {
        if($null -ne $dialogJob){
            [IO.File]::WriteAllText($dialogStop,'stop')
            [void](Wait-Job $dialogJob -Timeout 5)
            $observerFailed=$dialogJob.State -eq 'Failed'
            if($dialogJob.State -eq 'Running'){Stop-Job $dialogJob}
            Remove-Job $dialogJob
            if($observerFailed){Check 'Shipping.DialogObserver.CompletedWithoutError' $false}
        }
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClose')
        if($null -ne $operator){$operator.Close($false)}
        $other.Close($false)
    }
}
