# Actual Delete handler; all inventory starts at the owning Admin CSV boundary.
function Install-InventoryDeleteResultProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    if($CheckInventorySubmissionActivity){Install-InventorySubmissionNoticeProbe $project}
    $project.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function DeleteResultShowsForTest() As Long
    DeleteResultShowsForTest = mShows
End Function
Public Function DeleteResultPartialForTest() As Boolean
    DeleteResultPartialForTest = (InStr(1, mLastNotice, "partial", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "Event:", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "saved", vbTextCompare) > 0)
End Function
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryDeleteResult'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mLose As Boolean, mReturned As Boolean
Public Sub Arm(ByVal lose As Boolean)
    mLose = lose: mReturned = False
End Sub
Public Sub AfterProcessor()
    mReturned = True
    If mLose Then modAuth.SignOut
End Sub
Public Function Returned() As Boolean
    Returned = mReturned
End Function
Public Function PrepareExtensions(ByVal warehouse As String, ByVal sku As String) As String
    Dim book As Workbook, sheet As Worksheet, table As ListObject, extra As ListColumn
    Dim row As Long, keys As String
    Set book = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(warehouse)
    For Each sheet In book.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = "tblInventoryEntities" Or table.Name = "tblSkuCatalog" Then
                modAdmin.InventoryBoundaryProtectionForTest sheet, False
                Set extra = Nothing
                On Error Resume Next
                Set extra = table.ListColumns("DELETE UNKNOWN")
                On Error GoTo 0
                If extra Is Nothing Then Set extra = table.ListColumns.Add: extra.Name = "DELETE UNKNOWN"
                extra.DataBodyRange.Value2 = "Preserved extension"
                If table.Name = "tblInventoryEntities" Then
                    For row = 1 To table.ListRows.Count
                        If CStr(table.DataBodyRange.Cells(row, table.ListColumns("SKU").Index).Value2) = sku Then
                            If keys <> "" Then keys = keys & "|"
                            keys = keys & CStr(table.DataBodyRange.Cells(row, table.ListColumns("System_Key").Index).Value2)
                        End If
                    Next row
                End If
                modAdmin.InventoryBoundaryProtectionForTest sheet, True
            End If
        Next table
    Next sheet
    book.Save
    PrepareExtensions = keys
End Function
'@)
    $core=$packages['invSys.Core.xlam'].VBProject
    $processor=$core.VBComponents.Add(1);$processor.Name='TestInventoryRetirementResult'
    $processor.CodeModule.AddFromString(@'
Option Explicit
Private mLose As Boolean, mReturned As Boolean
Private mEvent As String
Private mStopBeforeApply As Boolean
Public Sub StopBeforeApply()
    mStopBeforeApply = True
End Sub
Public Sub BeforeApply(ByVal context As String)
    If mStopBeforeApply Then modInventoryCommands.Release context
    mStopBeforeApply = False
End Sub
Public Sub Arm(ByVal lose As Boolean)
    mLose = lose: mReturned = False: mEvent = ""
End Sub
Public Sub AfterProcessor(ByVal eventId As String)
    mReturned = True: mEvent = eventId
    If mLose Then modAuth.SignOut
End Sub
Public Function EventMatches(ByVal observed As String) As Boolean
    EventMatches = (mEvent <> "" And mEvent = observed)
End Function
Public Function Returned() As Boolean
    Returned = mReturned
End Function
'@)
    $owner=$core.VBComponents.Item('modInventoryRetirementCommands').CodeModule
    $text=$owner.Lines(1,$owner.CountOfLines)
    $pattern='(?im)^[\t ]*processed = modProcessor\.RunBatch\(warehouse, 0, batchReport, eventId, applied, True\)[\t ]*\r?$'
    if([regex]::Matches($text,$pattern).Count -ne 1){throw 'Delete processor-return observation boundary unavailable.'}
    $text=[regex]::Replace($text,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+'    TestInventoryRetirementResult.AfterProcessor eventId'})
    if($CheckInventorySubmissionActivity){
        $anchor='(?im)^([\t ]*If Not modInventoryCommands\.BeginRetirementSave\(context, report\) Then GoTo Unconfirmed)[\t ]*\r?$'
        if([regex]::Matches($text,$anchor).Count -ne 1){throw 'Retirement pre-application test boundary unavailable.'}
        $text=[regex]::Replace($text,$anchor,[Text.RegularExpressions.MatchEvaluator]{param($m) '        TestInventoryRetirementResult.BeforeApply context'+"`r`n"+$m.Value})
    }
    $owner.DeleteLines(1,$owner.CountOfLines);$owner.AddFromString($text)
}

function Test-InventoryDeleteResults($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    $cases=@('Authorized','AuthorityLost')
    if($CheckInventorySubmissionActivity){
        $cases+=@('TrackingOff','PriorPending')
        $wire=[string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @('ADMIN_INVENTORY_DELETE',49))
        $definition=if($wire){$wire|ConvertFrom-Json}
        Check 'InventoryDeleteResult.Catalog.Definition' ($null -ne $definition -and $definition.Caption -ceq 'Delete Item' -and $definition.Class -ceq 'Command' -and $definition.OwnerId -ceq 'ADMIN_INVENTORY' -and $definition.Capability -ceq 'ADMIN_MAINT')
        Check 'InventoryDeleteResult.Catalog.Preserves48' ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.PriorDefinitionsPreserved' @(48)))
        $effects=@{REQUESTED='Unknown';CANCELLED='Unchanged';REFUSED='Unchanged';SAVED='Changed';APPLIED='Changed';PENDING='Unknown';PARTIAL='Changed';FAILED='Unknown'}
        foreach($code in $effects.Keys){
            Check ('InventoryDeleteResult.Conclusion.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.InventoryOutcome' @('ADMIN_INVENTORY_DELETE',49,$code,$effects[$code],($code -in 'APPLIED','SAVED'))))
        }
    }
    foreach($case in $cases){
        SelectTarget $Fixture
        if($CheckInventorySubmissionActivity){
            $policy=if($case -ceq 'TrackingOff'){'Disabled'}else{'Enabled'}
            if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.SetPolicy' @($policy))){throw 'Retirement tracking fixture unavailable.'}
            [void](Run 'invSys.Admin.xlam' 'TestSubmissionNotice.ClearNotices')
        }
        $sku='DEMO-DELETE-'+$case.ToUpperInvariant()
        $csv=Join-Path $runRoot ('delete-result-'+$case+'.csv')
        @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION',($sku+',Delete fixture,LB,CLEARVIEW,10,GOOD'),($sku+',Delete fixture,LB,CLEARVIEW,7,HOLD'))|Set-Content $csv -Encoding ASCII
        $created=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))
        $id='InventoryDeleteResult.'+$case
        Check ($id+'.OwningGeneratedFixtureApplied') $created
        if(-not $created){throw 'Generated retirement fixture unavailable.'}
        $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        $keys=([string](Run 'invSys.Admin.xlam' 'TestInventoryDeleteResult.PrepareExtensions' @($Fixture.Warehouse,$sku))).Split('|')
        Check ($id+'.TwoDistinctGeneratedEntities') ($keys.Count -eq 2 -and $keys[0] -cne $keys[1])
        if($keys.Count -ne 2){throw 'Exact all-active fixture unavailable.'}
        [pscustomobject]@{Case=$case;Stage='BeforeColdBaseline';FileExists=(Test-Path $path);Ready=[bool]$excel.Ready}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'delete-fixture-stages.jsonl')
        $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
        [pscustomobject]@{Case=$case;Stage='ColdBaselineRead'}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'delete-fixture-stages.jsonl')
        $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
        $authBefore=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
        [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('delete-result-'+$case+'.xlsx')),51)
        try{
            [void](Run 'invSys.Core.xlam' 'TestInventoryRetirementResult.Arm' @(($case -ceq 'AuthorityLost')))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
            if($CheckInventorySubmissionActivity -and $case -ceq 'Authorized'){
                foreach($cancel in 'DeclineDelete','EmptyDeleteReason'){
                    $priorFiles=@(Get-Slice4beActivityFiles $Fixture)
                    $priorBytes=Get-AdminInventoryEntryHash $path
                    $phase=if($cancel -ceq 'DeclineDelete'){'DeleteConfirmation'}else{'DeleteReason'}
                    [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Delete',$cancel,$phase,$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
                    $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $priorFiles}|ForEach-Object {Get-Content -LiteralPath $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_INVENTORY_DELETE')
                    $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
                    $end=@($records|Where-Object OutcomeCode -CEQ 'CANCELLED')
                    Check ('InventoryDeleteResult.'+$cancel+'.NoInventoryWrite') ((Get-AdminInventoryEntryHash $path) -ceq $priorBytes -and -not [bool](Run 'invSys.Core.xlam' 'TestInventoryRetirementResult.Returned'))
                    Check ('InventoryDeleteResult.'+$cancel+'.CorrelatedCancellation') ($records.Count -eq 2 -and $attempt.Count -eq 1 -and $end.Count -eq 1 -and $end[0].ActivityId -ceq $attempt[0].ActivityId -and $end[0].DataEffect -ceq 'Unchanged' -and @($end[0].SourceEventRefs).Count -eq 0)
                }
            }
            if($case -ceq 'PriorPending'){
                $priorFiles=@(Get-Slice4beActivityFiles $Fixture)
                [void](Run 'invSys.Core.xlam' 'TestInventoryRetirementResult.StopBeforeApply')
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Delete','Authorized','DeleteConfirmation',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
                $pending=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $priorFiles}|ForEach-Object {Get-Content -LiteralPath $_ -Raw|ConvertFrom-Json}|Where-Object {$_.ControlId -ceq 'ADMIN_INVENTORY_DELETE' -and $_.OutcomeCode -ceq 'PENDING'})
                $priorRefs=@($pending|ForEach-Object {$_.SourceEventRefs})
                Check ($id+'.EarlierClickOwnsPendingReference') ($pending.Count -eq 1 -and $priorRefs.Count -eq 1 -and $priorRefs[0].SubmissionState -ceq 'Submitted' -and -not [bool](Run 'invSys.Core.xlam' 'TestInventoryRetirementResult.Returned'))
                if($priorRefs.Count -ne 1){throw 'Earlier click did not leave its own pending event for retry.'}
            }
            if($CheckInventorySubmissionActivity){$beforeActivity=@(Get-Slice4beActivityFiles $Fixture)}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Delete','Authorized','DeleteConfirmation',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualDeleteAndConfirmationReached') $ready
            if(-not $ready){throw 'Actual Delete/confirmation unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            Check ($id+'.ProcessorReturned') ([bool](Run 'invSys.Core.xlam' 'TestInventoryRetirementResult.Returned'))
            Check ($id+'.ExpectedContinuation') ([long](Run 'invSys.Admin.xlam' 'TestInventoryCommands.DeleteResultShowsForTest') -eq $(if($case -ceq 'AuthorityLost'){1}else{2}))
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $exact=$after.AppliedRows -eq $before.AppliedRows+1 -and $after.LogRows -eq $before.LogRows+2 -and $after.EntityQuantities.Count -eq $before.EntityQuantities.Count
            foreach($key in $before.EntityQuantities.Keys){$expected=if($keys -ccontains $key){0}else{$before.EntityQuantities[$key]};if(-not $after.EntityQuantities.ContainsKey($key) -or $after.EntityQuantities[$key] -ne $expected){$exact=$false}}
            Check ($id+'.AllAndOnlyExactEntitiesRetired') $exact
            $hash=Get-AdminInventoryEntryHash $path;$book=$excel.Workbooks.Open($path,0,$true)
            try{
                $entities=Table $book 'tblInventoryEntities';$catalog=Table $book 'tblSkuCatalog';$states=0;$catalogState='';$extras=$true
                foreach($table in @($entities,$catalog)){
                    foreach($row in $table.ListRows){
                        if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('DELETE UNKNOWN').Index).Value2 -cne 'Preserved extension'){$extras=$false}
                        if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index).Value2 -ceq $sku){
                            if($table.Name -ceq 'tblInventoryEntities'){if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('InventoryState').Index).Value2 -ceq 'RETIRED'){$states++}}
                            else{$catalogState=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('CATALOG_STATE').Index).Value2}
                        }
                    }
                }
                Check ($id+'.ExactRetiredStatesSaved') ($states -eq 2)
                Check ($id+'.CatalogWriteFollowsAuthority') ($(if($case -ceq 'AuthorityLost'){$catalogState -cne 'RETIRED'}else{$catalogState -ceq 'RETIRED'}))
                Check ($id+'.UnknownColumnsPreserved') $extras
            }finally{$book.Close($false)}
            Check ($id+'.ColdInspectionPreservedBytes') ((Get-AdminInventoryEntryHash $path) -ceq $hash)
            if($CheckInventorySubmissionActivity){
                $rows=@(Get-Slice4beActivityFiles $Fixture | Where-Object {$_ -cnotin $beforeActivity} | ForEach-Object {Get-Content -LiteralPath $_ -Raw|ConvertFrom-Json})
                if($case -ceq 'TrackingOff'){
                    Check ($id+'.RetiredWithoutOptionalTracking') ($exact -and $catalogState -ceq 'RETIRED' -and $rows.Count -eq 0 -and [bool](Run 'invSys.Admin.xlam' 'TestSubmissionNotice.Expected' @(0)))
                    continue
                }
                $open=@($rows|Where-Object {$_.ControlId -ceq 'ADMIN_INVENTORY_OPEN' -and $_.OutcomeCode -ceq 'OPENED'})
                if($open.Count -ne 1){throw 'Expected ordinary enabled Inventory tracking during the delete gesture.'}
                $deletes=@($rows|Where-Object ControlId -CEQ 'ADMIN_INVENTORY_DELETE')
                $wire=ConvertTo-Json -InputObject $deletes -Compress -Depth 8
                $redacted=$wire -notmatch 'DEMO-DELETE-|Delete fixture|CLEARVIEW'
                foreach($key in $keys){if($wire.Contains($key)){$redacted=$false}}
                Check ($id+'.RedactedRetirement') $redacted
                Check ($id+'.ExpectedTrackingNotices') ([bool](Run 'invSys.Admin.xlam' 'TestSubmissionNotice.Expected' @([int]($case -ceq 'AuthorityLost'))))
                $attempt=@($deletes|Where-Object OutcomeCode -CEQ 'REQUESTED')
                $end=@($deletes|Where-Object OutcomeCode -CNE 'REQUESTED')
                Check ($id+'.RetirementRequestedByOriginalUser') ($attempt.Count -eq 1 -and $attempt[0].OwnerId -ceq 'ADMIN_INVENTORY' -and $attempt[0].UserId -ceq 'config-admin' -and $attempt[0].WarehouseId -ceq $Fixture.Warehouse -and $attempt[0].DataEffect -ceq 'Unknown' -and @($attempt[0].SourceEventRefs).Count -eq 0)
                if($case -ceq 'Authorized'){
                    Check ($id+'.RecordedCompletedRetirement') ($end.Count -eq 1 -and $attempt.Count -eq 1 -and $end[0].ActivityId -ceq $attempt[0].ActivityId -and $end[0].OutcomeCode -ceq 'APPLIED' -and $end[0].DataEffect -ceq 'Changed')
                    $refs=@($end|ForEach-Object {$_.SourceEventRefs})
                    Check ($id+'.ExactRetirementEvent') ($refs.Count -eq 1 -and $refs[0].SourceKind -ceq 'Inventory' -and $refs[0].SubmissionState -ceq 'Submitted' -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and [bool](Run 'invSys.Core.xlam' 'TestInventoryRetirementResult.EventMatches' @([string]$refs[0].EventId)))
                }elseif($case -ceq 'PriorPending'){
                    Check ($id+'.CatalogSavedWithoutClaimingOldSubmission') ($end.Count -eq 1 -and $attempt.Count -eq 1 -and $attempt[0].ActivityId -ceq $end[0].ActivityId -and $end[0].OutcomeCode -ceq 'SAVED' -and $end[0].DataEffect -ceq 'Changed' -and @($end[0].SourceEventRefs).Count -eq 0 -and [bool](Run 'invSys.Core.xlam' 'TestInventoryRetirementResult.EventMatches' @([string]$priorRefs[0].EventId)))
                }else{
                    Check ($id+'.LostAuthorityLeavesOriginalAttempt') ($attempt.Count -eq 1 -and $end.Count -eq 0)
                }
            }
            if($case -ceq 'AuthorityLost'){Check ($id+'.HonestSavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.DeleteResultPartialForTest'))}
            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Config save facts unavailable.'}
            $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $phase
            [pscustomobject]@{Case=$case;AuthBytesPreserved=($authBefore -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))}|
                ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'delete-authority-observations.jsonl')
        }finally{$hostBook.Close($false)}
    }
}
