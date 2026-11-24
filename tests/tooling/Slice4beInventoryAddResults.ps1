# Draft next D13 gate: actual Add handler, owning generated keys and saved outcomes.
function Install-InventoryAddResultProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $project.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function AddResultShowsForTest() As Long
    AddResultShowsForTest = mShows
End Function
Public Function AddResultPartialForTest() As Boolean
    AddResultPartialForTest = (InStr(1, mLastNotice, "partial", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "Event:", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "saved", vbTextCompare) > 0)
End Function
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryAddResult'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mCase As String, mReturned As Boolean, mKey As String, mEvent As String
Private mTrackingNotices As Long, mUnexpectedNotice As Boolean
Public Sub Arm(ByVal scenario As String)
    mCase = scenario: mReturned = False: mKey = "": mEvent = ""
    mTrackingNotices = 0: mUnexpectedNotice = False
End Sub
Public Sub TrackingNotice(ByVal text As String, ByVal style As VbMsgBoxStyle, ByVal title As String)
    If Left$(text, 21) <> "Tracking unavailable:" Or style <> vbExclamation Or title <> "invSys Admin" Then mUnexpectedNotice = True
    mTrackingNotices = mTrackingNotices + 1
End Sub
Public Function ExpectedTrackingNotices(ByVal expected As Long) As Boolean
    ExpectedTrackingNotices = (Not mUnexpectedNotice And mTrackingNotices = expected)
End Function
Public Sub RecordKey(ByVal key As String)
    mKey = key
End Sub
Public Sub AfterQueue(ByVal context As String, ByVal eventId As String)
    mEvent = eventId
    If mCase = "Pending" Then modInventoryCommands.Release context
End Sub
Public Sub AfterProcessor(ByVal context As String, ByVal eventId As String)
    mReturned = True: mEvent = eventId
    If mCase = "AuthorityLost" Then modAuth.SignOut
    If mCase = "Partial" Then modInventoryCommands.Release context
End Sub
Public Function Returned() As Boolean
    Returned = mReturned
End Function
Public Function KeyMatches(ByVal key As String) As Boolean
    KeyMatches = (mKey <> "" And mKey = key)
End Function
Public Function EventMatches(ByVal eventId As String) As Boolean
    EventMatches = (mEvent <> "" And mEvent = eventId)
End Function
'@)
    $owner=$project.VBComponents.Item('modAdmin').CodeModule
    $start=$owner.ProcStartLine('AddInventoryItemForWarehouse',0);$count=$owner.ProcCountLines('AddInventoryItemForWarehouse',0)
    $body=$owner.Lines($start,$count)
    foreach($point in @(
        @('    systemKey = modRoleEventWriter.CreateSystemKey()','    TestInventoryAddResult.RecordKey systemKey'),
        @('    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport, eventIdOut, applied, True)','    TestInventoryAddResult.AfterProcessor interactiveContext, eventIdOut')
    )){
        $pattern='(?im)^[\t ]*'+[regex]::Escape($point[0].Trim())+'[\t ]*\r?$'
        if([regex]::Matches($body,$pattern).Count -ne 1){throw 'Add result observation boundary unavailable.'}
        $mark=$point[1]
        $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+$mark})
    }
    if($CheckInventoryAddActivity){
        $queueGuard='If Not modInventoryCommands.BeginCreateSave(interactiveContext, report) Then GoTo Unconfirmed'
        if([regex]::Matches($body,[regex]::Escape($queueGuard)).Count -ne 1){throw 'Add queued-event owner boundary unavailable.'}
        $body=$body.Replace($queueGuard,('TestInventoryAddResult.AfterQueue interactiveContext, eventIdOut' + "`r`n        " + $queueGuard))
    }
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    if($CheckInventoryAddActivity){
        Install-InventoryTrackingNoticeProbe $project 'TestInventoryAddResult'
        . (Join-Path $PSScriptRoot 'Slice4beAdminDirectoryBoundaries.ps1')
        $policy=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestInventoryRecordingPolicy').CodeModule
        $policy.AddFromString(@'
Public Function AddReferencesValid(ByVal warehouse As String, ByVal code As String, ByVal wire As String) As Boolean
    Dim refs As Collection
    Set refs = modActivityReferences.Decode(warehouse, "ADMIN_INVENTORY_ADD", code, wire)
    AddReferencesValid = Not refs Is Nothing
End Function
'@)
        Install-AdminDirectoryBoundaries $policy $null -ControlId 'ADMIN_INVENTORY_ADD' -Outcome 'APPLIED' -ProbeName 'TestInventoryRecordingPolicy' -RootAction
    }
}

function Test-InventoryAddResults($Fixture,$Other) {
    $cases=@('Authorized','AuthorityLost')
    if($CheckInventoryAddActivity){
        $cases+=@('TrackingOff','Refused','AttemptWriteFailure','OutcomeWriteFailure','Partial','Pending')
        $definition=[string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @('ADMIN_INVENTORY_ADD',48))
        $control=if($definition){$definition|ConvertFrom-Json}
        Check 'InventoryAddResult.Catalog.AddItem' ($null -ne $control -and $control.Caption -ceq 'Add Item' -and $control.Class -ceq 'Command' -and $control.OwnerId -ceq 'ADMIN_INVENTORY' -and $control.Capability -ceq 'ADMIN_MAINT')
        Check 'InventoryAddResult.Catalog.PreviousDefinitionsPreserved' ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.PriorDefinitionsPreserved' @(47)))
        $effects=@{REQUESTED='Unknown';REFUSED='Unchanged';FAILED='Unknown';PENDING='Unknown';PARTIAL='Changed';APPLIED='Changed'}
        foreach($code in $effects.Keys){
            Check ('InventoryAddResult.Conclusion.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.InventoryOutcome' @('ADMIN_INVENTORY_ADD',48,$code,$effects[$code],($code -ceq 'APPLIED'))))
        }
        $first=@{WarehouseId=$Fixture.Warehouse;EventId='TEST-EVENT-A';SourceKind='Inventory';SubmissionState='Submitted'}
        $second=$first.Clone();$second.EventId='TEST-EVENT-B'
        $unknown=$first.Clone();$unknown.SubmissionState='Unknown'
        $otherKind=$first.Clone();$otherKind.SourceKind='Designs'
        $otherWarehouse=$first.Clone();$otherWarehouse.WarehouseId=$Other.Warehouse
        $sets=@{Empty=@();Submitted=@($first);Unknown=@($unknown);Multiple=@($first,$second);WrongKind=@($otherKind);WrongWarehouse=@($otherWarehouse)}
        foreach($code in $effects.Keys){
            $matrix=$true
            foreach($kind in $sets.Keys){
                $v=ConvertTo-Json -InputObject $sets[$kind] -Depth 3 -Compress
                $expected=switch($code){
                    {$_ -in 'REQUESTED','REFUSED'} {$kind -ceq 'Empty'}
                    'FAILED' {$kind -in 'Empty','Submitted','Unknown'}
                    default {$kind -ceq 'Submitted'}
                }
                $actual=[bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.AddReferencesValid' @($Fixture.Warehouse,$code,$v))
                if($expected -ne $actual){$matrix=$false}
            }
            Check ('InventoryAddResult.References.'+$code) $matrix
        }
    }
    foreach($case in $cases){
        SelectTarget $Fixture
        if($CheckInventoryAddActivity){
            $policy=if($case -ceq 'TrackingOff'){'Disabled'}else{'Enabled'}
            if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.SetPolicy' @($policy))){throw 'Inventory tracking policy fixture failed.'}
            [void](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.ArmBoundary' @($case,'','',''))
            $beforeActivity=@(Get-Slice4beActivityFiles $Fixture)
        }
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('add-result-'+$case+'.xlsx')),51)
        try{
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            [void](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.Arm' @($case))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            $gesture='Authorized'
            if($case -ceq 'Refused'){
                $gesture='SetQuantity'
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetQuantityTarget' @(-1.0))
            }
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Add',$gesture,'ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $id='InventoryAddResult.'+$case
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualAddHandlerReached') $ready
            if(-not $ready){throw 'Actual Add action unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            if($CheckInventoryAddActivity){
                $newActivity=@(Get-Slice4beActivityFiles $Fixture | Where-Object {$_ -cnotin $beforeActivity} | ForEach-Object {Get-Content -LiteralPath $_ -Raw | ConvertFrom-Json})
                $add=@($newActivity | Where-Object ControlId -CEQ 'ADMIN_INVENTORY_ADD')
                $wrongSource=@($add | Where-Object {$_.WarehouseId -cne $Fixture.Warehouse -or $_.UserId -cne 'config-admin' -or $_.OwnerId -cne 'ADMIN_INVENTORY'})
                Check ($id+'.OriginalActivitySource') ($wrongSource.Count -eq 0)
                $wire=ConvertTo-Json -InputObject $add -Depth 10 -Compress
                Check ($id+'.RedactedItemData') ($wire -notmatch 'Boundary fixture item|CLEARVIEW|System_Key')
            }
            if($case -ceq 'Refused'){
                $attempt=@($add | Where-Object OutcomeCode -CEQ 'REQUESTED')
                $refusal=@($add | Where-Object OutcomeCode -CEQ 'REFUSED')
                Check ($id+'.RefusalBeforeProcessing') (-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.Returned') -and $after.EntityQuantities.Count -eq $before.EntityQuantities.Count -and $after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows)
                Check ($id+'.RecordedPrewriteRefusal') ($add.Count -eq 2 -and $attempt.Count -eq 1 -and $refusal.Count -eq 1 -and $refusal[0].ActivityId -ceq $attempt[0].ActivityId -and $refusal[0].DataEffect -ceq 'Unchanged' -and @($refusal[0].SourceEventRefs).Count -eq 0)
                continue
            }
            if($case -in 'Partial','Pending'){
                $applied=$case -ceq 'Partial'
                $delta=[int]$applied
                Check ($id+'.IndependentSavedApplication') ($after.AppliedRows -eq $before.AppliedRows+$delta -and $after.LogRows -eq $before.LogRows+$delta -and $after.EntityQuantities.Count -eq $before.EntityQuantities.Count+$delta)
                Check ($id+'.ExpectedProcessorBoundary') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.Returned') -eq $applied)
                $attempt=@($add | Where-Object OutcomeCode -CEQ 'REQUESTED')
                $terminal=@($add | Where-Object OutcomeCode -CEQ $case.ToUpperInvariant())
                Check ($id+'.OwnerDerivedTerminal') ($add.Count -eq 2 -and $attempt.Count -eq 1 -and $terminal.Count -eq 1 -and $attempt[0].ActivityId -ceq $terminal[0].ActivityId -and $terminal[0].DataEffect -ceq $effects[$case])
                $refs=@($terminal|ForEach-Object {$_.SourceEventRefs})
                Check ($id+'.ExactSubmittedEvent') ($refs.Count -eq 1 -and $refs[0].SubmissionState -ceq 'Submitted' -and $refs[0].SourceKind -ceq 'Inventory' -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and [bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.EventMatches' @([string]$refs[0].EventId)))
                continue
            }
            Check ($id+'.ProcessorReturned') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.Returned'))
            $shows=[long](Run 'invSys.Admin.xlam' 'TestInventoryCommands.AddResultShowsForTest')
            Check ($id+'.ExpectedContinuation') ($shows -eq $(if($case -ceq 'AuthorityLost'){1}else{2}))
            $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
            $exact=$newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq 0 -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1
            Check ($id+'.ExactNewZeroEntitySaved') $exact
            if($newKeys.Count -ne 1){throw 'Expected owning creation output unavailable.'}
            Check ($id+'.OwningGeneratedKeyPreserved') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.KeyMatches' @($newKeys[0])))
            if($CheckInventoryAddActivity){
                $writeFailure=$case -in 'AttemptWriteFailure','OutcomeWriteFailure'
                if($writeFailure -and [int](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.WriteFailures') -ne 1){throw 'Expected Add store failure did not reach its actual write.'}
                $expectedNotices=[int]($writeFailure -or $case -ceq 'AuthorityLost')
                Check ($id+'.ExpectedTrackingNotices') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.ExpectedTrackingNotices' @($expectedNotices)))
                if($case -ceq 'TrackingOff'){
                    Check ($id+'.SavedWithoutOptionalRecording') ($exact -and $newActivity.Count -eq 0)
                    continue
                }
                $opening=@($newActivity | Where-Object {$_.ControlId -ceq 'ADMIN_INVENTORY_OPEN' -and $_.OutcomeCode -ceq 'REQUESTED'})
                if($opening.Count -ne 1){throw 'The enabled tracking fixture did not capture the actual Inventory entry.'}
                $attempt=@($add | Where-Object OutcomeCode -CEQ 'REQUESTED')
                $terminal=@($add | Where-Object OutcomeCode -CNE 'REQUESTED')
                if($writeFailure){
                    Check ($id+'.OptionalStoreFailurePreservesBusinessResult') ($exact -and $terminal.Count -eq 0 -and $attempt.Count -eq [int]($case -ceq 'OutcomeWriteFailure'))
                    continue
                }
                Check ($id+'.RecordedAddRequest') ($attempt.Count -eq 1 -and $attempt[0].DataEffect -ceq 'Unknown' -and @($attempt[0].SourceEventRefs).Count -eq 0)
                if($case -ceq 'AuthorityLost'){
                    Check ($id+'.OriginalSessionLeftIncomplete') ($attempt.Count -eq 1 -and $terminal.Count -eq 0)
                }else{
                    $applied=$attempt.Count -eq 1 -and $terminal.Count -eq 1 -and $terminal[0].ActivityId -ceq $attempt[0].ActivityId -and $terminal[0].OutcomeCode -ceq 'APPLIED' -and $terminal[0].DataEffect -ceq 'Changed'
                    Check ($id+'.RecordedSavedApplication') $applied
                    $refs=@($terminal | ForEach-Object {$_.SourceEventRefs})
                    $exactSource=$refs.Count -eq 1 -and $refs[0].SourceKind -ceq 'Inventory' -and $refs[0].SubmissionState -ceq 'Submitted' -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and [bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.EventMatches' @([string]$refs[0].EventId))
                    Check ($id+'.RecordedExactOwnerEvent') $exactSource
                }
            }
            if($case -ceq 'AuthorityLost'){
                Check ($id+'.HonestSavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.AddResultPartialForTest'))
            }
        }finally{$hostBook.Close($false)}
    }
}
