# Actual worksheet Upload handler and real source application; owned fixtures only.
function Install-InventoryWorksheetResultProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $project.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function WorksheetResultPartialForTest() As Boolean
    WorksheetResultPartialForTest = (InStr(1, mLastNotice, "partial", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "Event:", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "saved", vbTextCompare) > 0)
End Function
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryWorksheetResult'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mHost As String, mApplied As Boolean, mReached As Boolean
Private mYieldHash As String, mFailed As Boolean, mSavesAfterLoss As Long
Public Sub Arm(ByVal mode As String, ByVal host As String)
    mMode = mode: mHost = host: mApplied = False: mReached = False
    mYieldHash = "": mFailed = False: mSavesAfterLoss = 0
End Sub
Private Sub ObserveYield()
    mReached = True
    mYieldHash = modTrainingWire.FileSha256(Application.Workbooks(mHost).FullName)
End Sub
Public Sub AfterPreflight()
    Dim sheet As Worksheet, table As ListObject
    If mMode = "PreflightAuthorityLost" Then
        ObserveYield
        modAuth.SignOut
    ElseIf mMode = "PendingRowChanged" Then
        For Each sheet In Application.Workbooks(mHost).Worksheets
            For Each table In sheet.ListObjects
                If Left$(table.Name, 17) = "invSys_Inventory_" Then
                    table.ListColumns("Quantity").DataBodyRange.Cells(1, 1).Value2 = 1
                    ObserveYield
                    Exit Sub
                End If
            Next table
        Next sheet
    End If
End Sub
Public Sub AfterApply()
    mApplied = True
    If mMode = "AfterApplyAuthorityLost" Then
        ObserveYield
        modAuth.SignOut
    End If
End Sub
Public Sub SaveReturn(ByVal book As Workbook)
    If mApplied And mMode = "SaveThrows" And Not mFailed Then
        mFailed = True: ObserveYield
        Err.Raise vbObjectError + 2140, , "Owned result-save failure."
    End If
    If mReached And Not modAuth.IsSignedIn() Then mSavesAfterLoss = mSavesAfterLoss + 1
    book.Save
    If mApplied And mMode = "SavedAuthorityLost" And Not mReached Then
        ObserveYield
        modAuth.SignOut
    End If
End Sub
Public Function Facts() As String
    Dim unchanged As Boolean
    If mYieldHash <> "" Then unchanged = (mYieldHash = modTrainingWire.FileSha256(Application.Workbooks(mHost).FullName))
    Facts = CStr(mReached) & "|" & CStr(mApplied) & "|" & CStr(unchanged) & "|" & CStr(mSavesAfterLoss)
End Function
'@)
    # Observe the exact worksheet form message shown to the operator.
    $form=$project.VBComponents.Item('frmAddInventoryItem').CodeModule
    $start=$form.ProcStartLine('MsgBox',0);$count=$form.ProcCountLines('MsgBox',0)
    $body=$form.Lines($start,$count)
    $notice='(?im)^[\t ]*MsgBox = TestInventoryCommands\.DialogChoice\(Prompt, Buttons\)[\t ]*\r?$'
    if([regex]::Matches($body,$notice).Count -ne 1){throw 'Actual worksheet notice observation unavailable.'}
    $body=[regex]::Replace($body,$notice,[Text.RegularExpressions.MatchEvaluator]{param($m) '    TestInventoryCommands.RecordNotice Prompt'+"`r`n"+$m.Value.TrimEnd("`r")})
    $form.DeleteLines($start,$count);$form.InsertLines($start,$body)
    $owner=$project.VBComponents.Item('modAdminInventoryWorksheet').CodeModule
    $start=$owner.ProcStartLine('UploadSelectedInventoryWorksheetTable',0);$count=$owner.ProcCountLines('UploadSelectedInventoryWorksheetTable',0)
    $body=$owner.Lines($start,$count)
    $before='(?im)^[\t ]*If records Is Nothing Or records.Count = 0 Then[\t ]*\r?$'
    $after='(?im)^[\t ]*If ApplyInventoryWorksheetRecord\(record, warehouseId, stationId, userId, applyReport(?:, interactiveContext)?\) Then[\t ]*\r?$'
    if([regex]::Matches($body,$before).Count -ne 1 -or [regex]::Matches($body,$after).Count -ne 1){throw 'Worksheet action observation boundary unavailable.'}
    $body=[regex]::Replace($body,$before,[Text.RegularExpressions.MatchEvaluator]{param($m) '    TestInventoryWorksheetResult.AfterPreflight'+"`r`n"+$m.Value.TrimEnd("`r")})
    $body=[regex]::Replace($body,$after,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+'            TestInventoryWorksheetResult.AfterApply'})
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $start=$owner.ProcStartLine('SaveInventoryWorksheetWorkbook',0);$count=$owner.ProcCountLines('SaveInventoryWorksheetWorkbook',0)
    $body=$owner.Lines($start,$count)
    $pattern='(?i)Then wb.Save'
    if([regex]::Matches($body,$pattern).Count -ne 1){throw 'Worksheet save-return seam unavailable.'}
    $body=[regex]::Replace($body,$pattern,'Then TestInventoryWorksheetResult.SaveReturn wb')
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
}

function Test-InventoryWorksheetResults($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    foreach($case in @('Authorized','PreflightAuthorityLost','AfterApplyAuthorityLost','PendingRowChanged','SaveThrows','SavedAuthorityLost')){
        SelectTarget $Fixture
        $hostBook=$excel.Workbooks.Add();$hostPath=Join-Path $runRoot ('worksheet-result-'+$case+'.xlsx');$hostBook.SaveAs($hostPath,51)
        try{
            $prepared=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
            $id='InventoryWorksheetResult.'+$case
            Check ($id+'.OperatorTablePrepared') $prepared
            if(-not $prepared){throw 'Owned staging table unavailable.'}
            $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
            $extra=$table.ListColumns.Add();$extra.Name='USER NOTE';$extra.DataBodyRange.Cells.Item(1,1).Value2='Preserved worksheet extension';$hostBook.Save()
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $authBefore=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @($case,$hostBook.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualUploadHandlerReached') $ready
            if(-not $ready){throw 'Actual Upload handler unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            $boundary=([string](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Facts')).Split('|')
            if($boundary.Count -ne 4){throw 'Worksheet boundary facts unavailable.'}
            if($case -cne 'Authorized'){Check ($id+'.RequiredBoundaryReached') ($boundary[0] -ceq 'True')}
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
            if($case -cin @('PreflightAuthorityLost','PendingRowChanged')){
                Check ($id+'.NoAuthoritativeOutputAfterYield') ($newKeys.Count -eq 0 -and $after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows)
            }else{
                Check ($id+'.ExactZeroEntitySaved') ($newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq 0 -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1)
            }
            $eventId='';$createdKey=''
            if($case -cnotin @('PreflightAuthorityLost','PendingRowChanged') -and $newKeys.Count -eq 1){
                $createdKey=[string]$newKeys[0];$sourcePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
                $sourceHash=Get-AdminInventoryEntryHash $sourcePath;$source=$excel.Workbooks.Open($sourcePath,0,$true)
                try{
                    $entities=Table $source 'tblInventoryEntities';$matches=@()
                    foreach($row in $entities.ListRows){if([string]$row.Range.Cells.Item(1,$entities.ListColumns.Item('System_Key').Index).Value2 -ceq $createdKey){$matches+=$row}}
                    $active=$matches.Count -eq 1 -and [double]$matches[0].Range.Cells.Item(1,$entities.ListColumns.Item('QtyOnHand').Index).Value2 -eq 0 -and [string]$matches[0].Range.Cells.Item(1,$entities.ListColumns.Item('InventoryState').Index).Value2 -ceq 'ACTIVE'
                    Check ($id+'.IndependentActiveZeroEntitySaved') $active
                    $log=Table $source 'tblInventoryLog';$events=@()
                    foreach($row in $log.ListRows){if([string]$row.Range.Cells.Item(1,$log.ListColumns.Item('System_Key').Index).Value2 -ceq $createdKey){$events+=[string]$row.Range.Cells.Item(1,$log.ListColumns.Item('EventID').Index).Value2}}
                    if($events.Count -ne 1){throw 'Exact create event identity unavailable.'};$eventId=$events[0]
                }finally{$source.Close($false)}
                Check ($id+'.SourceColdInspectionPreservedBytes') ((Get-AdminInventoryEntryHash $sourcePath) -ceq $sourceHash)
            }
            if($case -cne 'Authorized'){
                Check ($id+'.NoOperatorSaveAfterYield') ($boundary[2] -ceq 'True' -and $boundary[3] -ceq '0')
            }
            if($case -cin @('AfterApplyAuthorityLost','SaveThrows','SavedAuthorityLost')){
                Check ($id+'.HonestSavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.WorksheetResultPartialForTest'))
            }
            Check ($id+'.UnknownStagingColumnPreserved') ([string]$table.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Preserved worksheet extension')
            if($case -ceq 'Authorized'){
                $hash=Get-AdminInventoryEntryHash $hostPath
                $copy=Join-Path $runRoot ('cold-worksheet-'+[guid]::NewGuid().ToString('N')+'.xlsx');Copy-Item $hostPath $copy
                $cold=$excel.Workbooks.Open($copy,0,$true)
                try{
                    $savedTable=@($cold.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                    Check ($id+'.IndependentImportedStatusSaved') ([string]$savedTable.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'IMPORTED')
                    Check ($id+'.IndependentUnknownColumnSaved') ([string]$savedTable.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Preserved worksheet extension')
                }finally{$cold.Close($false)}
                Check ($id+'.ColdInspectionPreservedBytes') ((Get-AdminInventoryEntryHash $hostPath) -ceq $hash)
            }

            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Config save facts unavailable.'}
            $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $phase
            [pscustomobject]@{Case=$case;AuthBytesPreserved=($authBefore -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))}|
                ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'worksheet-authority-observations.jsonl')
            if($case -ceq 'SaveThrows'){
                $hostBook.Close($false);$hostBook=$excel.Workbooks.Open($hostPath,0,$false)
                $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                $hostBook.Activate();$table.DataBodyRange.Cells.Item(1,1).Select()
                SelectTarget $Fixture
                [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('Retry',$hostBook.Name))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
                $retried=Get-AdminInventoryEntryFacts $Fixture -Persisted
                $same=$retried.LogRows -eq $after.LogRows -and $retried.AppliedRows -eq $after.AppliedRows -and $retried.EntityQuantities.Count -eq $after.EntityQuantities.Count
                foreach($key in $after.EntityQuantities.Keys){if(-not $retried.EntityQuantities.ContainsKey($key) -or $retried.EntityQuantities[$key] -ne $after.EntityQuantities[$key]){$same=$false}}
                Check ($id+'.FreshRetryDoesNotDuplicateSavedCreate') $same
                $retryFacts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
                Check ($id+'.FreshRetryActualHandlerReached') ($retryFacts.Count -eq 5 -and $retryFacts[0] -ceq 'True' -and $retryFacts[2] -ceq '1' -and $retryFacts[3] -ceq 'True')
                $retryCopy=Join-Path $runRoot ('retry-worksheet-'+[guid]::NewGuid().ToString('N')+'.xlsx');Copy-Item $hostPath $retryCopy
                $cold=$excel.Workbooks.Open($retryCopy,0,$true)
                try{
                    $savedTable=@($cold.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                    Check ($id+'.FreshRetryImportedStatusSaved') ([string]$savedTable.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'IMPORTED')
                    $receipt=[string]$savedTable.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2
                    Check ($id+'.FreshRetryRetainsExactEventAndEntity') ($eventId -ne '' -and $createdKey -ne '' -and $receipt.Contains($eventId) -and $receipt.Contains($createdKey))
                }finally{$cold.Close($false)}

            }
        }finally{$hostBook.Close($false)}
    }
}
