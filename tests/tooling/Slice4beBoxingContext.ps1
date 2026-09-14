# D18 context and permission checks use the real Make/Unbox handlers. Only the
# mutation-entry probe stops work, after counting entry; service authorization
# and form binding remain real. Normal Boxing tests run with this probe disabled.
function Install-Slice4beBoxingContextProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $module=$project.VBComponents.Item('modTS_Shipments').CodeModule
    $module.InsertLines($module.CountOfDeclarationLines+1,@'
Public BoxingServiceEntriesForTest As Long
Private BoxingOwnerEntriesForTest As Long
Private BoxingStopOwnerForTest As Boolean
'@)
    $start=$module.ProcStartLine('CommitBoxMakerFormAction',0)
    $count=$module.ProcCountLines('CommitBoxMakerFormAction',0)
    $source=$module.Lines($start,$count)
    $anchor='    syncCompletedOut = False'
    if([regex]::Matches($source,[regex]::Escape($anchor)).Count -ne 1){throw 'Boxing owner probe anchor unavailable.'}
    $source=$source.Replace($anchor,$anchor+@'

    BoxingOwnerEntriesForTest = BoxingOwnerEntriesForTest + 1
    If BoxingStopOwnerForTest Then
        resultMessage = "Boxing mutation boundary probe."
        Exit Function
    End If
'@)
    $module.DeleteLines($start,$count);$module.InsertLines($start,$source)
    $module.AddFromString(@'
Public Sub BoxingStopOwnerProbeForTest(ByVal enabled As Boolean)
    BoxingStopOwnerForTest = enabled
End Sub
Public Function BoxingEntryCountForTest(ByVal boundary As String) As Long
    If boundary = "Service" Then
        BoxingEntryCountForTest = BoxingServiceEntriesForTest
    ElseIf boundary = "Owner" Then
        BoxingEntryCountForTest = BoxingOwnerEntriesForTest
    Else
        Err.Raise 5, , "Unknown Boxing probe boundary."
    End If
End Function
'@)
    $service=$project.VBComponents.Item('modBoxingService').CodeModule
    $start=$service.ProcStartLine('PostBoxMakerAction',0)
    $count=$service.ProcCountLines('PostBoxMakerAction',0)
    $source=$service.Lines($start,$count)
    $anchor='    On Error GoTo Fail'
    if([regex]::Matches($source,[regex]::Escape($anchor)).Count -ne 1){throw 'Boxing service probe anchor unavailable.'}
    $source=$source.Replace($anchor,"    modTS_Shipments.BoxingServiceEntriesForTest = modTS_Shipments.BoxingServiceEntriesForTest + 1`r`n"+$anchor)
    $service.DeleteLines($start,$count);$service.InsertLines($start,$source)
}

function Test-BoxingDeniedObservation($Fixture,$Before,[string]$Action) {
    $prefix='Boxing.Capability.Revoked.'+$Action+'.Activity.'
    $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before}|ForEach-Object {[IO.File]::ReadAllText($_)|ConvertFrom-Json})
    $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
    $result=@($records|Where-Object OutcomeCode -CEQ 'DENIED')
    $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $result.Count -eq 1
    Check ($prefix+'RequestedThenDenied') $pair
    $facts=$false;$metadata=$false;$refs=$false
    if($pair){
        $facts=$attempt[0].Severity -ceq 'Info' -and $attempt[0].DataEffect -ceq 'Unknown' -and
            $result[0].Severity -ceq 'Blocked' -and $result[0].DataEffect -ceq 'Unchanged' -and
            $attempt[0].ActivityId -ceq $result[0].ActivityId -and $attempt[0].RecordId -cne $result[0].RecordId
        $metadata=@($records|Where-Object {$_.ControlId -cne ('BOXING_'+$Action) -or $_.OwnerId -cne 'BOXING_WORKFLOW' -or
            $_.SourceRole -cne 'Boxing' -or $_.WarehouseId -cne $Fixture.Warehouse -or $_.UserId -cne 'config-reader' -or
            $_.EventCode -cne ('BOXING_'+$Action+'_'+$_.OutcomeCode)}).Count -eq 0
        $refs=@($attempt[0].SourceEventRefs).Count -eq 0 -and @($result[0].SourceEventRefs).Count -eq 0
    }
    Check ($prefix+'ExplicitDenialFacts') $facts
    Check ($prefix+'TrustedContextAndOwner') $metadata
    Check ($prefix+'NoSubmittedSources') $refs
}

function Test-Slice4beBoxingContext($Fixture,$Operator,$Other,$Ship,$Hold) {
    $second=NewFixture 'boxing-context-other'
    # The alternate target also grants SHIP_POST so its stale-form case cannot
    # pass the mutation guard merely because this actor lacks that permission.
    $secondAuth=$excel.Workbooks.Open((Join-Path $second.Root ($second.Warehouse+'.invSys.Auth.xlsb')),0,$false)
    try{
        $caps=Table $secondAuth 'tblCapabilities';$row=$caps.ListRows.Add()
        foreach($pair in @{UserId='config-reader';Capability='SHIP_POST';WarehouseId=$second.Warehouse;StationId='S1';Status='Active'}.GetEnumerator()){
            $row.Range.Cells.Item(1,$caps.ListColumns.Item($pair.Key).Index).Value2=$pair.Value
        }
        $secondAuth.Save()
    }finally{$secondAuth.Close($false)}
    SelectTarget $Fixture 'config-reader'
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
    $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @('Prepare'))
    if(-not $key){throw 'Boxing context fixture selection unavailable.'}
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $secondInventory=Join-Path $second.Root ($second.Warehouse+'.invSys.Data.Inventory.xlsb')
    $authorityHash=Get-ShippingActivityHash $inventory
    $secondHash=Get-ShippingActivityHash $secondInventory
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath)
    $authHash=Get-ShippingActivityHash $authPath
    $revoked=$false
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingStopOwnerProbeForTest' @($true))
    try {
        foreach($context in @('Healthy','SignedOut','Reauthenticated','OtherTarget','Revoked')){
            switch($context){
                'SignedOut' {
                    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
                    $ready=-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn')
                }
                'Reauthenticated' {
                    $version=[long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion')
                    $signed=[string](Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @('config-reader',$Fixture.Secret,''))
                    $ready=$signed.StartsWith('OK|') -and [long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion') -ne $version -and
                        [string](Run 'invSys.Core.xlam' 'modAuth.GetCurrentUserId') -ceq 'config-reader'
                }
                'OtherTarget' {
                    SelectTarget $second 'config-reader'
                    $ready=$second.Warehouse -cne $Fixture.Warehouse -and [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn') -and
                        [string](Run 'invSys.Core.xlam' 'modConfig.GetWarehouseId') -ceq $second.Warehouse -and
                        [bool](Run 'invSys.Core.xlam' 'TestShippingSession.CanShip')
                }
                'Revoked' {
                    SelectTarget $Fixture 'config-reader'
                    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
                    $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @('Prepare'))
                    if(-not $key){throw 'Boxing permission fixture selection unavailable.'}
                    $version=[long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion')
                    $auth=$excel.Workbooks.Open($authPath,0,$false)
                    try {
                        $caps=Table $auth 'tblCapabilities';$count=0
                        foreach($row in $caps.ListRows){
                            if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader' -and
                               $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'SHIP_POST'){
                                $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$count++
                            }
                        }
                        if($count -ne 1){throw 'Boxing fixture permission is not unique.'}
                        $revoked=$true;$auth.Save()
                    }finally{$auth.Close($false)}
                    $ready=-not [bool](Run 'invSys.Core.xlam' 'TestShippingSession.CanShip') -and
                        [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn') -and $version -eq [long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion')
                }
                default {$ready=[bool](Run 'invSys.Core.xlam' 'TestShippingSession.CanShip')}
            }
            Check ('Boxing.Context.'+$context+'.Established') $ready
            if(-not $ready){throw 'Boxing context/permission case was not established.'}
            foreach($action in @('MAKE','UNBOX')){
                $prefix='Boxing.Context.'+$context+'.'+$action
                $beforeService=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingEntryCountForTest' @('Service'))
                $beforeOwner=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingEntryCountForTest' @('Owner'))
                $beforeQueue=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue'))
                $before=@(Get-Slice4beActivityFiles $Fixture);$beforeOther=@(Get-Slice4beActivityFiles $second)
                $rows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
                $held=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
                $Other.Activate()
                Check ($prefix+'.ActualHandlerReturned') ([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @($action,'1')) -ceq 'DELIVERED')
                $service=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingEntryCountForTest' @('Service'))
                $owner=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingEntryCountForTest' @('Owner'))
                if($context -ceq 'Healthy'){
                    $calibrated=$service -eq $beforeService+1 -and $owner -eq $beforeOwner+1
                    Check ($prefix+'.BothEntryProbesCalibrated') $calibrated
                    if(-not $calibrated){throw 'Boxing healthy entry probes were not calibrated.'}
                }else{
                    Check ($prefix+'.StopsBeforeMutationOwner') ($owner -eq $beforeOwner)
                    $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
                    if($context -ceq 'Revoked'){
                        Check ($prefix+'.VisiblePermissionDenial') ($status -match '(?i)permission|capability|not authorized')
                        Test-BoxingDeniedObservation $Fixture $before $action
                    }else{
                        Check ($prefix+'.StopsBeforeServiceDispatch') ($service -eq $beforeService)
                        Check ($prefix+'.VisibleContextRejection') ($status -match '(?i)session|warehouse' -and $status -match '(?i)reopen')
                        Check ($prefix+'.NoCrossContextActivity') (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}).Count -eq 0 -and
                            @(Get-Slice4beActivityFiles $second|Where-Object {$_ -cnotin $beforeOther}).Count -eq 0)
                    }
                }
                Check ($prefix+'.NoSubmissionEntry') ($beforeQueue -eq [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue')))
                Check ($prefix+'.StagingAndUnknownValuesPreserved') ($rows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $held -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
                Check ($prefix+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)))
                Check ($prefix+'.BothAuthorityBytesPreserved') ($authorityHash -ceq (Get-ShippingActivityHash $inventory) -and $secondHash -ceq (Get-ShippingActivityHash $secondInventory))
                Check ($prefix+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel')
            }
        }
    }finally{
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingStopOwnerProbeForTest' @($false))
        if($revoked){[IO.File]::WriteAllBytes($authPath,$authBytes)}
        SelectTarget $Fixture 'config-reader'
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
    }
    Check 'Boxing.Context.AuthBytesRestored' ($authHash -ceq (Get-ShippingActivityHash $authPath))
    Check 'Boxing.Context.ReopenRetainsWorkbook' ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)))
    Check 'Boxing.Context.RepeatedLauncherReusesForm' ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses'))
}
