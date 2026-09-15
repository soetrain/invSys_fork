# Unsaved fault/return observers. Real Make/Unbox handlers, authorization, event
# construction and writes remain active; injected failures are explicit and counted.
function Edit-BoxingProbeProcedure($Module,[string]$Name,[string]$Anchor,[string]$Replacement) {
    $start=$Module.ProcStartLine($Name,0);$count=$Module.ProcCountLines($Name,0)
    $source=$Module.Lines($start,$count)
    if([regex]::Matches($source,[regex]::Escape($Anchor)).Count -ne 1){throw ('Boxing outcome probe anchor unavailable: '+$Name)}
    $source=$source.Replace($Anchor,$Replacement)
    $Module.DeleteLines($start,$count);$Module.InsertLines($start,$source)
}

function Install-Slice4beBoxingOutcomeProbes {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modRoleEventWriter').CodeModule
    $core.InsertLines($core.CountOfDeclarationLines+1,"Private BoxingStageFaultArmed As Boolean`r`nPrivate BoxingStageFaultCount As Long`r`nPrivate BoxingSubmissionWriteState As String")
    # Capture direct submission counters before a later staging merge can enter
    # the same low-level inbox writer. Catch-up writes are not this submission.
    foreach($name in @('QueuePayloadEventServer','QueuePayloadEventCurrent')){
        $start=$core.ProcStartLine($name,0);$count=$core.ProcCountLines($name,0)
        $source=$core.Lines($start,$count)
        $source=$source.Replace('Exit Function','Call BoxingCaptureSubmissionWrites: Exit Function')
        $source=$source.Replace('End Function',"    BoxingCaptureSubmissionWrites`r`nEnd Function")
        $core.DeleteLines($start,$count);$core.InsertLines($start,$source)
    }
    Edit-BoxingProbeProcedure $core 'SyncLocalStagedInboxRows' '    On Error GoTo FailSync' @'
    On Error GoTo FailSync
    If BoxingStageFaultArmed Then
        BoxingStageFaultArmed = False
        BoxingStageFaultCount = BoxingStageFaultCount + 1
        report = "Fixture local staging unavailable."
        Exit Function
    End If
'@
    $core.AddFromString(@'
Public Sub BoxingStageFaultForTest(ByVal enabled As Boolean)
    BoxingStageFaultArmed = enabled: BoxingStageFaultCount = 0
    BoxingSubmissionWriteState = ""
End Sub
Private Sub BoxingCaptureSubmissionWrites()
    BoxingSubmissionWriteState = ActivityShippingWriteEntryState()
End Sub
Public Function BoxingSubmissionWritesForTest() As String
    BoxingSubmissionWritesForTest = BoxingSubmissionWriteState
End Function
Public Function BoxingStageFaultCountForTest() As Long
    BoxingStageFaultCountForTest = BoxingStageFaultCount
End Function
Public Function BoxingSubmissionIdForTest() As String
    BoxingSubmissionIdForTest = ActivityCurrentOutput
    If BoxingSubmissionIdForTest = "" Then BoxingSubmissionIdForTest = ActivityServerId
End Function
'@)
    $processor=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modProcessor').CodeModule
    $processor.InsertLines($processor.CountOfDeclarationLines+1,"Private BoxingBatchMode As Long`r`nPrivate BoxingBatchFaultCount As Long")
    Edit-BoxingProbeProcedure $processor 'RunBatch' '    On Error GoTo FailRun' @'
    On Error GoTo FailRun
    If BoxingBatchMode <> 0 Then
        BoxingBatchFaultCount = BoxingBatchFaultCount + 1
        report = "Fixture processing deferred."
        If BoxingBatchMode = 2 Then report = "RunBatch failed: fixture processing unavailable."
        BoxingBatchMode = 0
        Exit Function
    End If
'@
    $processor.AddFromString(@'
Public Sub BoxingBatchFaultForTest(ByVal mode As Long)
    BoxingBatchMode = mode: BoxingBatchFaultCount = 0
End Sub
Public Function BoxingBatchFaultCountForTest() As Long
    BoxingBatchFaultCountForTest = BoxingBatchFaultCount
End Function
'@)
    $module=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modTS_Shipments').CodeModule
    $module.InsertLines($module.CountOfDeclarationLines+1,@'
Private BoxingRefreshFaultArmed As Boolean
Private BoxingRefreshFaultCount As Long
Private BoxingOwnerReturned As Boolean, BoxingOwnerReturnValue As Boolean
Private BoxingLegacySyncValue As Boolean, BoxingReturnedSource As String
'@)
    $start=$module.ProcStartLine('ShipmentsFormRefreshReadModelForWorkbook',0)
    $count=$module.ProcCountLines('ShipmentsFormRefreshReadModelForWorkbook',0)
    $source=$module.Lines($start,$count)
    $pattern='(?s)(Public Function ShipmentsFormRefreshReadModelForWorkbook\(.*?\) As Boolean\s*\r?\n)'
    if([regex]::Matches($source,$pattern).Count -ne 1){throw 'Boxing refresh failure anchor unavailable.'}
    $source=[regex]::Replace($source,$pattern,'$1'+@'
    If BoxingRefreshFaultArmed Then
        BoxingRefreshFaultArmed = False
        BoxingRefreshFaultCount = BoxingRefreshFaultCount + 1
        report = "Fixture read-model refresh unavailable."
        Exit Function
    End If

'@)
    $module.DeleteLines($start,$count);$module.InsertLines($start,$source)
    $start=$module.ProcStartLine('CommitBoxMakerFormAction',0)
    $count=$module.ProcCountLines('CommitBoxMakerFormAction',0)
    $source=$module.Lines($start,$count)
    $observe='BoxingObserveOwnerReturn CommitBoxMakerFormAction, eventIdOut, syncCompletedOut'
    $source=$source.Replace('Exit Function',$observe+': Exit Function')
    $source=$source.Replace('End Function',"    $observe`r`nEnd Function")
    $module.DeleteLines($start,$count);$module.InsertLines($start,$source)
    $module.AddFromString(@'
Public Sub BoxingResetOutcomeForTest(ByVal failRefresh As Boolean)
    BoxingRefreshFaultArmed = failRefresh: BoxingRefreshFaultCount = 0
    BoxingOwnerReturned = False: BoxingOwnerReturnValue = False
    BoxingLegacySyncValue = False: BoxingReturnedSource = ""
End Sub
Private Sub BoxingObserveOwnerReturn(ByVal succeeded As Boolean, ByVal sourceId As String, ByVal synced As Boolean)
    BoxingOwnerReturned = True: BoxingOwnerReturnValue = succeeded
    BoxingReturnedSource = sourceId: BoxingLegacySyncValue = synced
End Sub
Public Function BoxingOwnerResultForTest(ByVal expectedId As String) As String
    BoxingOwnerResultForTest = CStr(BoxingOwnerReturned) & "|" & CStr(BoxingOwnerReturnValue) & "|" & _
        CStr(BoxingLegacySyncValue) & "|" & CStr(expectedId <> "" And expectedId = BoxingReturnedSource)
End Function
Public Function BoxingRefreshFaultCountForTest() As Long
    BoxingRefreshFaultCountForTest = BoxingRefreshFaultCount
End Function
'@)
}

function Reset-BoxingOutcomeFaults {
    [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionMode' @(0))
    [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.BoxingStageFaultForTest' @($false))
    [void](Run 'invSys.Core.xlam' 'modProcessor.BoxingBatchFaultForTest' @(0))
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingResetOutcomeForTest' @($false))
}

function Test-Slice4beBoxingOutcomes($Fixture,$Operator,$Other,$Ship,$Hold) {
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $observedFacts=New-Object 'System.Collections.Generic.List[object]'
    # Deferred/uncertain submissions precede real processing cases. Their catch-up
    # must never be attributed to a later click's SourceEventRefs.
    $cases=@(
        @{Name='UncertainAcceptance';Submit=3;Outcome='FAILED';Applied=$false;State='Unknown'},
        @{Name='ProcessingDeferred';Submit=7;Batch=1;Outcome='PENDING';Applied=$false},
        @{Name='ProcessingFailed';Submit=7;Batch=2;Outcome='FAILED';Applied=$false},
        @{Name='StagingFailed';Submit=7;Stage=$true;Outcome='FAILED';Applied=$true},
        @{Name='RefreshFailed';Submit=7;Refresh=$true;Outcome='FAILED';Applied=$true},
        @{Name='ServerUnavailable';Submit=1;Outcome='CONFIRMED';Applied=$true},
        @{Name='LostAcknowledgment';Submit=2;Outcome='CONFIRMED';Applied=$true},
        @{Name='ExceptionalAcknowledgment';Submit=4;Outcome='CONFIRMED';Applied=$true},
        @{Name='PrewriteRefusal';Submit=5;Outcome='FAILED';Applied=$false}
    )
    try {
        foreach($case in $cases){foreach($action in @('MAKE','UNBOX')){
            Reset-BoxingOutcomeFaults
            $label='Boxing.Outcomes.'+$case.Name+'.'+$action
            $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @('Prepare'))
            if(-not $key){throw 'Boxing outcome fixture selection unavailable.'}
            $rows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
            $held=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
            $beforeLog=@(Get-ShippingActivityLog $Fixture)
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $beforeAuthority=Get-ShippingActivityHash $inventory
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionMode' @($case.Submit))
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.BoxingStageFaultForTest' @([bool]$case['Stage']))
            [void](Run 'invSys.Core.xlam' 'modProcessor.BoxingBatchFaultForTest' @([int]$case['Batch']))
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingResetOutcomeForTest' @([bool]$case['Refresh']))
            $Other.Activate()
            Check ($label+'.ActualHandlerReturned') ([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @($action,'1')) -ceq 'DELIVERED')
            $id=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.BoxingSubmissionIdForTest')
            $state=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionState')).Split('|')
            $entries=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.BoxingSubmissionWritesForTest')
            $owner=([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingOwnerResultForTest' @($id))).Split('|')
            $expectedEntries=switch($case.Submit){1{'0|0|1'} 3{'0|1|0'} 5{'2|0|0'} 7{'0|1|0'} default{'0|1|1'}}
            $serverAccepted=$case.Submit -notin @(1,5)
            $localAccepted=$case.Submit -in @(1,2,4)
            $calibrated=$state.Count -eq 6 -and $state[0] -ceq '1' -and $state[1] -ceq $(if($case.Submit -eq 7){'0'}else{'1'}) -and
                $state[2] -ceq [string]$serverAccepted -and $state[3] -ceq [string]$localAccepted -and $entries -ceq $expectedEntries -and $id -ne ''
            Check ($label+'.SubmissionAndWriteFaultCalibrated') $calibrated
            if(-not $calibrated){throw 'Boxing submission fault was not established; not product RED.'}
            Check ($label+'.ExactIdAcrossEnteredSubmissionRoutes') ($case.Submit -eq 7 -or $(if($case.Submit -eq 1){$state[5] -ceq 'True'}else{$state[4] -ceq 'True'}))
            $returned=$owner.Count -eq 4 -and $owner[0] -ceq 'True' -and $owner[1] -cin @('True','False') -and $owner[2] -cin @('True','False') -and $owner[3] -ceq 'True'
            Check ($label+'.ActualOwnerReturnAndSourceObserved') $returned
            if(-not $returned){throw 'Boxing owner return/source observation unavailable.'}
            $stageCount=Run 'invSys.Core.xlam' 'modRoleEventWriter.BoxingStageFaultCountForTest'
            $batchCount=Run 'invSys.Core.xlam' 'modProcessor.BoxingBatchFaultCountForTest'
            $refreshCount=Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingRefreshFaultCountForTest'
            $faults=$null -ne $stageCount -and $null -ne $batchCount -and $null -ne $refreshCount -and
                [int]$stageCount -eq [int][bool]$case['Stage'] -and [int]$batchCount -eq [int]([int]$case['Batch'] -gt 0) -and [int]$refreshCount -eq [int][bool]$case['Refresh']
            Check ($label+'.RequiredStepFaultCalibrated') $faults
            if(-not $faults){throw 'Boxing required-step fault was not entered exactly as configured.'}
            $explicitRefresh=Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingRefreshForTest'
            if($explicitRefresh -isnot [bool]){throw 'Boxing explicit refresh result is unavailable.'}
            $observedFacts.Add([pscustomobject]@{Case=$case.Name;Action=$action;OwnerReturned=$true;OwnerSuccess=($owner[1] -ceq 'True');LegacySync=($owner[2] -ceq 'True');ExplicitRefresh=$explicitRefresh;StageFault=[int]$stageCount;BatchFault=[int]$batchCount;RefreshFault=[int]$refreshCount})
            $pin=Get-ShippingActivityHash $inventory
            $afterLog=@(Get-ShippingActivityLog $Fixture)
            Check ($label+'.EvidenceReadPreservesAuthorityBytes') ($pin -ceq (Get-ShippingActivityHash $inventory))
            $applied=@($afterLog|Where-Object EventID -CEQ $id)
            $expectedType=if($action -ceq 'MAKE'){'BOX_BUILD'}else{'BOX_UNBOX'}
            $qty=if($action -ceq 'MAKE'){1}else{-1}
            $domain=$applied.Count -eq 0 -and $beforeAuthority -ceq $pin
            if($case.Applied){
                $domain=$applied.Count -eq 2 -and @($beforeLog|Where-Object EventID -CEQ $id).Count -eq 0 -and
                    @($applied|Where-Object {$_.EventType -cne $expectedType -or $_.WarehouseId -cne $Fixture.Warehouse}).Count -eq 0 -and
                    @($applied|Where-Object {$_.System_Key -ceq $key -and [double]$_.QtyDelta -eq $qty}).Count -eq 1 -and
                    @($applied|Where-Object {$_.System_Key -cne $key -and $_.System_Key -ne '' -and [double]$_.QtyDelta -eq -$qty}).Count -eq 1
            }
            Check ($label+'.IndependentDomainApplicationAndExactLines') $domain
            if(-not $domain){throw 'Boxing independent application fixture did not match the configured branch.'}
            $serverPath=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionStoragePath' @($true))
            $serverRows=@()
            if($serverPath){
                if(-not [IO.Path]::GetFullPath($serverPath).StartsWith([IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Boxing submission escaped fixture root.'}
                $hash=Get-ShippingActivityHash $serverPath
                $book=$excel.Workbooks.Open($serverPath,0,$true)
                try{$serverRows=@(Get-ShippingActivityRows (Table $book 'tblInboxShip')|Where-Object EventID -CEQ $id)}finally{$book.Close($false)}
                if($hash -cne (Get-ShippingActivityHash $serverPath)){throw 'Read-only Boxing submission inspection changed bytes.'}
            }
            Check ($label+'.IndependentServerSubmissionIdentity') ($serverRows.Count -eq [int]$serverAccepted -and
                @($serverRows|Where-Object {$_.WarehouseId -cne $Fixture.Warehouse -or $_.EventType -cne $expectedType}).Count -eq 0)
            $ids=@(if($case.Submit -ne 5){$id})
            $submissionState=if($case.ContainsKey('State')){$case.State}else{'Submitted'}
            $caption=if($action -ceq 'MAKE'){'Make Boxes'}else{'Unbox'}
            Test-BoxingObservation $Fixture $before $label ('BOXING_'+$action) $caption $case.Outcome $ids '' 0 $submissionState
            Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)))
            Check ($label+'.StagingKeysAndUnknownValuesPreserved') ($rows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $held -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
            Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping submission sentinel')
        }}
    }finally{
        $observedFacts|ConvertTo-Json -Depth 4|Set-Content (Join-Path $reportRoot 'boxing-owner-return-facts.json')
        Reset-BoxingOutcomeFaults
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @('Shipping'))
    }
}
