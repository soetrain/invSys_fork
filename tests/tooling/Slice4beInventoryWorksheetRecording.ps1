# Actual packaged worksheet handlers; existing owner fault seams and cold reads.
. (Join-Path $PSScriptRoot 'Slice4beInventoryCreatePhases.ps1')
. (Join-Path $PSScriptRoot 'Slice4beInventoryWorksheetResults.ps1')
function Install-InventoryWorksheetRecordingProbe {
    Install-InventoryCreatePhaseProbe
    Install-InventoryWorksheetResultProbe
    if($AdminInventoryWorksheetRecordingGroup -ceq 'UploadNoFormat'){
        $config=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modConfig').CodeModule
        $start=$config.ProcStartLine('EnsureStationInbox',0);$count=$config.ProcCountLines('EnsureStationInbox',0)
        $body=$config.Lines($start,$count)
        $call='EnsureConfigSchema(wbCfg, resolvedWh, resolvedSt, report)'
        if([regex]::Matches($body,[regex]::Escape($call)).Count -ne 1){throw 'Config formatting diagnostic boundary unavailable.'}
        $body=$body.Replace($call,'EnsureConfigSchema(wbCfg, resolvedWh, resolvedSt, report, False)')
        $config.DeleteLines($start,$count);$config.InsertLines($start,$body)
        Write-Output 'DIAGNOSTIC ONLY: inbox Config formatting suppressed in memory; no product GREEN claim.'
    }
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestInventoryRecordingPolicy').CodeModule
    $core.AddFromString(@'
Public Function Prior199Preserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(46)
        Set before = modActivityCatalog.Control(CStr(id), 46)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    Prior199Preserved = True
End Function
Public Function WorksheetOutcome(ByVal id As String, ByVal code As String, ByVal effect As String, ByVal completed As Boolean) As Boolean
    Dim outcome As Object, record As Object
    Set outcome = modActivityCatalog.Outcome(id, code)
    If outcome Is Nothing Then Exit Function
    If outcome("DataEffect") <> effect Then Exit Function
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "CatalogVersion", 47
    record.Add "OwnerId", "ADMIN_INVENTORY": record.Add "OutcomeCode", code
    WorksheetOutcome = (modEvaluationMatches.CommandCompleted(record) = completed)
End Function
'@)
    $queueProbe=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1);$queueProbe.Name='TestWorksheetQueue'
    $queueProbe.CodeModule.AddFromString(@'
Option Explicit
Private mThrow As Boolean, mEvent As String
Public Sub Arm(ByVal shouldThrow As Boolean)
    mThrow = shouldThrow: mEvent = ""
End Sub
Public Sub BeforeWrite(ByVal eventId As String)
    If Not mThrow Then Exit Sub
    mEvent = eventId: mThrow = False
    Err.Raise vbObjectError + 2186, , "Owned queue write refusal."
End Sub
Public Function EventIdentity() As String
    EventIdentity = mEvent
End Function
'@)
    $writer=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modRoleEventWriter').CodeModule
    $pattern='(?im)^[\t ]*writeAttemptedOut = True[\t ]*\r?$'
    foreach($procedure in @('QueueEventCore','AppendInboxRowToLocalStagingRole')){
        $start=$writer.ProcStartLine($procedure,0);$count=$writer.ProcCountLines($procedure,0)
        $body=$writer.Lines($start,$count)
        if([regex]::Matches($body,$pattern).Count -ne 1){throw 'Exact queue write boundary unavailable; not product RED.'}
        $injection=if($procedure -ceq 'QueueEventCore'){'    TestWorksheetQueue.BeforeWrite eventIdOut'}else{'    TestWorksheetQueue.BeforeWrite CStr(rowValues("EventID"))'}
        $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+$injection})
        $writer.DeleteLines($start,$count);$writer.InsertLines($start,$body)
    }
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $form=$admin.VBComponents.Item('frmAddInventoryItem').CodeModule
    $form.AddFromString(@'
Public Sub ExpireWorksheetOwnerForTest()
    modInventoryCommands.Release mInteractiveContext
End Sub
'@)
    $probe=$admin.VBComponents.Add(1);$probe.Name='TestWorksheetRecording'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mExpire As Boolean
Public Sub Arm(ByVal expire As Boolean)
    mExpire = expire
End Sub
Public Sub BeforeInvoke(ByVal form As frmAddInventoryItem)
    If mExpire Then form.ExpireWorksheetOwnerForTest
End Sub
'@)
    $driver=$admin.VBComponents.Item('TestInventoryCommands').CodeModule
    $start=$driver.ProcStartLine('Drive',0);$count=$driver.ProcCountLines('Drive',0)
    $body=$driver.Lines($start,$count)
    $pattern='(?im)^[\t ]*mHandler = mHandler \+ 1[\t ]*\r?$'
    if([regex]::Matches($body,$pattern).Count -ne 1){throw 'Worksheet handler entry adapter unavailable; not product RED.'}
    $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) '    TestWorksheetRecording.BeforeInvoke form'+"`r`n"+$m.Value.TrimEnd("`r")})
    $driver.DeleteLines($start,$count);$driver.InsertLines($start,$body)
}

function Test-InventoryWorksheetRecording($Fixture,$Other,[string]$Group='All') {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    SelectTarget $Fixture
    Check 'WorksheetRecording.Prior199DefinitionsPreserved' ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Prior199Preserved'))
    foreach($suffix in @('CREATE_TABLE','UPLOAD_TABLE')){
        $id='ADMIN_INVENTORY_'+$suffix
        $wire=[string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @($id,47))
        $definition=if($wire){$wire|ConvertFrom-Json}else{$null}
        Check ('WorksheetRecording.Catalog47.'+$suffix) ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_INVENTORY' -and $definition.Capability -ceq 'ADMIN_MAINT' -and $definition.Class -ceq 'Command')
        Check ('WorksheetRecording.Version46Excludes.'+$suffix) ([string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @($id,46)) -ceq '')
        $codes=if($suffix -ceq 'CREATE_TABLE'){@('REQUESTED','SAVED','REFUSED','PARTIAL','FAILED')}else{@('REQUESTED','CONFIRMED','EMPTY','REJECTED','REFUSED','PARTIAL','FAILED')}
        foreach($code in $codes){
            $effect=if($code -cin @('SAVED','REFUSED','EMPTY','REJECTED')){'Unchanged'}else{'Unknown'}
            $positive=$code -cin @('SAVED','CONFIRMED')
            Check ('WorksheetRecording.Conclusion.'+$suffix+'.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.WorksheetOutcome' @($id,$code,$effect,$positive)))
        }
    }
    $cases=@('CreateEnabled','CreateDisabled','CreateControlOff','CreateUserOff','CreateExpired','CreateSignedOut','CreateSaveThrows',
        'UploadEnabled','UploadEmpty','UploadInvalid','UploadControlOff','UploadExpired','UploadSignedOut','UploadSaveThrows','UploadQueueThrows')
    if($Group -ceq 'UploadNoFormat'){$cases=@('UploadEnabled')}
    elseif($Group -ceq 'Evidence'){$cases=@('UploadEnabled','UploadSaveThrows','UploadQueueThrows')}
    elseif($Group -ne 'All'){$cases=@($cases|Where-Object {$_.StartsWith($Group)})}
    foreach($case in $cases){
        SelectTarget $Fixture
        $create=$case.StartsWith('Create');$action=if($create){'Create'}else{'Upload'}
        $suffix=if($create){'CREATE_TABLE'}else{'UPLOAD_TABLE'};$id='WorksheetRecording.'+$case
        $policy=if($case.EndsWith('Disabled')){'Disabled'}elseif($case.EndsWith('ControlOff')){'ControlOff'}elseif($case.EndsWith('UserOff')){'UserOff'}else{'Enabled'}
        if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.SetPolicy' @($policy))){throw 'Worksheet recording policy fixture unavailable.'}
        $book=$excel.Workbooks.Add();$path=Join-Path $runRoot ($case+'.xlsx')
        $book.Worksheets.Item(1).Range('A1').Value2='Unrelated operator content';$book.SaveAs($path,51)
        try{
            # A previous action's injected failure must not affect fixture setup.
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCreatePhase.Arm' @('Authorized',$book.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('Authorized',$book.Name))
            [void](Run 'invSys.Core.xlam' 'TestWorksheetQueue.Arm' @($false))
            $table=$null
            if(-not $create){
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($book.Name))){throw 'Owned worksheet fixture unavailable; not product RED.'}
                $table=@($book.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                $extra=$table.ListColumns.Add();$extra.Name='USER NOTE';$extra.DataBodyRange.Cells.Item(1,1).Value2='WorksheetRecordingSentinel'
                if($case -ceq 'UploadEmpty'){$table.DataBodyRange.ClearContents()}
                if($case -ceq 'UploadInvalid'){$table.ListColumns.Item('Quantity').DataBodyRange.Cells.Item(1,1).Value2=-1}
                $book.Save();$table.DataBodyRange.Cells.Item(1,1).Select()
            }
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            $pins=@{};foreach($store in @('Config','Auth')){$storePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.'+$store+'.xlsb');$pins[$storePath]=Get-AdminInventoryEntryHash $storePath}
            $activityBefore=@(Get-Slice4beActivityFiles $Fixture)
            $fault=if($case.EndsWith('SaveThrows')){'SaveThrows'}else{'Authorized'}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCreatePhase.Arm' @($(if($create){$fault}else{'Authorized'}),$book.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @($(if($create){'Authorized'}else{$fault}),$book.Name))
            [void](Run 'invSys.Admin.xlam' 'TestWorksheetRecording.Arm' @($case.EndsWith('Expired')))
            [void](Run 'invSys.Core.xlam' 'TestWorksheetQueue.Arm' @($case -ceq 'UploadQueueThrows'))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            $authority=if($case.EndsWith('SignedOut')){'SignedOut'}else{'Authorized'}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($book.Name,$action,$authority,'ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq 'True'
            Check ($id+'.ActualBoundHandler') $ready
            if(-not $ready){throw 'Actual worksheet handler not reached; not product RED.'}
            $eventId=''
            if($case -ceq 'UploadQueueThrows'){
                $eventId=[string](Run 'invSys.Core.xlam' 'TestWorksheetQueue.EventIdentity')
                Check ($id+'.ActualWriteBoundaryReached') ($eventId -cne '')
                if($eventId -ceq ''){throw 'Queue write boundary not reached; not product RED.'}
            }
            $refused=$case.EndsWith('Expired') -or $case.EndsWith('SignedOut')
            $applied=(-not $create -and -not $refused -and $case -cnotin @('UploadEmpty','UploadInvalid','UploadQueueThrows'))
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $keys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
            $owner=if($applied){$keys.Count -eq 1 -and $after.EntityQuantities[$keys[0]] -eq 0 -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1}else{$after.FileHash -ceq $before.FileHash}
            foreach($key in $before.EntityQuantities.Keys){$owner=$owner -and $after.EntityQuantities.ContainsKey($key) -and $after.EntityQuantities[$key] -eq $before.EntityQuantities[$key]}
            Check ($id+'.IndependentCanonicalOutcome') $owner
            if(-not $owner){throw 'Unexpected owning business result; recording RED cannot mask it.'}
            if($applied){
                $sourcePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$hash=Get-AdminInventoryEntryHash $sourcePath
                $source=$excel.Workbooks.Open($sourcePath,0,$true)
                try{
                    $entities=Table $source 'tblInventoryEntities';$matches=@()
                    foreach($row in $entities.ListRows){if([string]$row.Range.Cells.Item(1,$entities.ListColumns.Item('System_Key').Index).Value2 -ceq $keys[0]){$matches+=$row}}
                    Check ($id+'.ColdActiveEntity') ($matches.Count -eq 1 -and [double]$matches[0].Range.Cells.Item(1,$entities.ListColumns.Item('QtyOnHand').Index).Value2 -eq 0 -and [string]$matches[0].Range.Cells.Item(1,$entities.ListColumns.Item('InventoryState').Index).Value2 -ceq 'ACTIVE')
                    $log=Table $source 'tblInventoryLog';$events=@()
                    foreach($row in $log.ListRows){if([string]$row.Range.Cells.Item(1,$log.ListColumns.Item('System_Key').Index).Value2 -ceq $keys[0]){$events+=[string]$row.Range.Cells.Item(1,$log.ListColumns.Item('EventID').Index).Value2}}
                    if($events.Count -ne 1){throw 'Exact generated event missing.'};$eventId=$events[0]
                }finally{$source.Close($false)}
                Check ($id+'.ColdCanonicalReadPreservedBytes') ((Get-AdminInventoryEntryHash $sourcePath) -ceq $hash)
            }
            $hostHash=Get-AdminInventoryEntryHash $path;$copy=Join-Path $runRoot ('cold-'+$case+'.xlsx');Copy-Item -LiteralPath $path -Destination $copy
            $cold=$excel.Workbooks.Open($copy,0,$true)
            try{
                $tables=@($cold.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})
                if($create){
                    $saved=-not $refused -and -not $case.EndsWith('SaveThrows')
                    Check ($id+'.IndependentLocalSave') ($tables.Count -eq $(if($saved){1}else{0}))
                }else{
                    $status=[string]$tables[0].ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2
                    Check ($id+'.IndependentResultSave') ($(if($applied -and -not $case.EndsWith('SaveThrows')){$status -ceq 'IMPORTED'}else{$status -cne 'IMPORTED'}))
                    if($case -cne 'UploadEmpty'){Check ($id+'.UnknownColumnPreserved') ([string]$tables[0].ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'WorksheetRecordingSentinel')}
                }
                Check ($id+'.UnrelatedContentPreserved') ([string]$cold.Worksheets.Item(1).Range('A1').Value2 -ceq 'Unrelated operator content')
            }finally{$cold.Close($false)}
            Check ($id+'.ColdOperatorReadPreservedBytes') ((Get-AdminInventoryEntryHash $path) -ceq $hostHash)
            $allRows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $activityBefore}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
            Check ($id+'.NoInventedClose') (@($allRows|Where-Object ControlId -CEQ 'ADMIN_INVENTORY_CLOSE').Count -eq 0)
            $rows=@($allRows|Where-Object ControlId -CEQ ('ADMIN_INVENTORY_'+$suffix))
            $suppressed=$policy -cne 'Enabled' -or $case.EndsWith('SignedOut')
            Check ($id+'.ExactRecordCount') ($rows.Count -eq $(if($suppressed){0}else{2}))
            if(-not $suppressed){
                $outcome=if($case.EndsWith('Expired')){'REFUSED'}elseif($case.EndsWith('SaveThrows')){'PARTIAL'}elseif($case -ceq 'UploadQueueThrows'){'FAILED'}elseif($case -ceq 'UploadEmpty'){'EMPTY'}elseif($case -ceq 'UploadInvalid'){'REJECTED'}elseif($create){'SAVED'}else{'CONFIRMED'}
                $effect=if($outcome -cin @('CONFIRMED','PARTIAL','FAILED')){'Unknown'}else{'Unchanged'}
                $attempts=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$ends=@($rows|Where-Object OutcomeCode -CEQ $outcome)
                Check ($id+'.ExactCorrelatedOwnerOutcome') ($attempts.Count -eq 1 -and $ends.Count -eq 1 -and $attempts[0].ActivityId -ceq $ends[0].ActivityId -and $ends[0].DataEffect -ceq $effect)
                $refs=@($ends|ForEach-Object {$_.SourceEventRefs})
                $state=if($case -ceq 'UploadQueueThrows'){'Unknown'}else{'Submitted'}
                $exact=if($eventId -cne ''){$refs.Count -eq 1 -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SourceKind -ceq 'Inventory' -and $refs[0].EventId -ceq $eventId -and $refs[0].SubmissionState -ceq $state}else{$refs.Count -eq 0}
                Check ($id+'.ExactOwnerEventReferences') $exact
            }
            $json=[string]($rows|ConvertTo-Json -Depth 12 -Compress)
            Check ($id+'.Redacted') ($json -notmatch 'WorksheetRecordingSentinel|Unrelated operator content|CLEARVIEW|Boundary fixture')
            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Config save observations unavailable.'}
            $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $phase
            $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
            $authSame=(Get-AdminInventoryEntryHash $authPath) -ceq $pins[$authPath]
            if($create){
                Check ($id+'.AuthorityStoresPreserved') ($authSame -and $configBefore.FileHash -ceq $configAfter.FileHash)
            }else{
                Check ($id+'.AuthBytesPreserved') $authSame
                Check ($id+'.ConfigBytesOrExplainedSaveMetadata') ($configBefore.FileHash -ceq $configAfter.FileHash -or $script:InventoryConfigLastComparison.CauseEstablished)
            }
            if($case -ceq 'UploadSaveThrows'){
                $book.Close($false);$book=$excel.Workbooks.Open($path,0,$false)
                $table=@($book.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                $book.Activate();$table.DataBodyRange.Cells.Item(1,1).Select()
                SelectTarget $Fixture
                [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('Authorized',$book.Name))
                $prior=@(Get-Slice4beActivityFiles $Fixture)
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($book.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
                $retryFacts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
                Check ($id+'.Retry.ActualBoundHandler') ($retryFacts.Count -eq 5 -and $retryFacts[0] -ceq 'True' -and $retryFacts[2] -ceq '1' -and $retryFacts[3] -ceq 'True')
                $retry=Get-AdminInventoryEntryFacts $Fixture -Persisted
                Check ($id+'.Retry.CanonicalBytesUnchanged') ($retry.FileHash -ceq $after.FileHash)
                $retryRows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $prior}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_INVENTORY_UPLOAD_TABLE')
                $attempt=@($retryRows|Where-Object OutcomeCode -CEQ 'REQUESTED');$end=@($retryRows|Where-Object OutcomeCode -CEQ 'CONFIRMED')
                Check ($id+'.Retry.CorrelatedConfirmation') ($retryRows.Count -eq 2 -and $attempt.Count -eq 1 -and $end.Count -eq 1 -and $attempt[0].ActivityId -ceq $end[0].ActivityId -and $end[0].DataEffect -ceq 'Unknown')
                $refs=@($end|ForEach-Object {$_.SourceEventRefs})
                Check ($id+'.Retry.ExactPriorEvent') ($refs.Count -eq 1 -and $refs[0].EventId -ceq $eventId -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SourceKind -ceq 'Inventory' -and $refs[0].SubmissionState -ceq 'Submitted')
                $hash=Get-AdminInventoryEntryHash $path;$copy=Join-Path $runRoot 'cold-upload-retry.xlsx';Copy-Item -LiteralPath $path -Destination $copy
                $cold=$excel.Workbooks.Open($copy,0,$true)
                try{
                    $savedTable=@($cold.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                    Check ($id+'.Retry.SavedImportedStatus') ([string]$savedTable.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'IMPORTED')
                    Check ($id+'.Retry.SavedCustomColumn') ([string]$savedTable.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'WorksheetRecordingSentinel')
                }finally{$cold.Close($false)}
                Check ($id+'.Retry.ColdReadPreservedBytes') ((Get-AdminInventoryEntryHash $path) -ceq $hash)
                $configPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Config.xlsb')
                Check ($id+'.Retry.AuthorityBytesPreserved') ((Get-AdminInventoryEntryHash $authPath) -ceq $pins[$authPath] -and (Get-AdminInventoryEntryHash $configPath) -ceq $configAfter.FileHash)
            }
        }finally{$book.Close($false)}
    }
    Test-InventoryConfigDiagnosticPreservation
}
