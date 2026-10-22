# Shared read-only observation for actual EDIT setup, failure and retry actions.
# Values and hashes remain memory-only; output contains fixed labels and facts.
. (Join-Path $PSScriptRoot 'Slice4beInventoryReceiptScenarios.ps1')
function Start-InventoryEditConfigDiagnosis($Fixture) {
    $before=Get-InventoryConfigDiagnosticSnapshot $Fixture
    $auth=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
    [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
    return [pscustomobject]@{Config=$before;Auth=$auth}
}
function Invoke-InventoryEditObservedUpload($Fixture,$HostBook,[string]$Label) {
    $diagnosis=Start-InventoryEditConfigDiagnosis $Fixture
    [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($HostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
    Complete-InventoryEditConfigDiagnosis $Fixture $diagnosis $Label
}
function Complete-InventoryEditConfigDiagnosis($Fixture,$Before,[string]$Label) {
    $after=Get-InventoryConfigDiagnosticSnapshot $Fixture
    $facts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
    if($facts.Count -ne 3 -or @($facts|Where-Object {$_ -cnotmatch '^\d+$'}).Count){throw 'Config save observer unavailable.'}
    $phase=[pscustomobject]@{Calls=[int]$facts[0];ByteChangesDuringSave=[int]$facts[1];UnavailableHashComparisons=[int]$facts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
    Write-InventoryConfigDiagnosticComparison $Before.Config $after $Label $phase
    $comparison=$script:InventoryConfigLastComparison
    $authAfter=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
    $authPreserved=$Before.Auth -ne '' -and $Before.Auth -ceq $authAfter
    $bytesPreserved=$Before.Config.FileHash -ne '' -and $Before.Config.FileHash -ceq $after.FileHash
    Check ($Label+'.ConfigCellsAndShapesPreserved') ($comparison.ChangedTableCells -eq 0 -and $comparison.TableShapesPreserved)
    Check ($Label+'.AuthBytesPreserved') $authPreserved
    [pscustomobject]@{Case=$Label;ConfigBytesPreserved=$bytesPreserved;AuthBytesPreserved=$authPreserved;ConfigChangeExplained=($bytesPreserved -or $comparison.CauseEstablished)}|
        ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'inventory-edit-authority-observations.jsonl')
}
# Owned, service-created test warehouses only. Business values and exact command /
# event identities remain in memory or ignored disposable workbook copies. Reports
# contain fixed check labels and booleans; no row values or raw errors are emitted.
function Install-InventoryEditReceiptProbe {
    $project=$packages['invSys.Core.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryEditReceipt'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mSource As String, mHostPath As String, mCopy As String
Private mReached As Boolean, mCopied As Boolean, mPlanned As String, mPreserved As Boolean
Private mQueued As String, mPayload As String, mLookup As String
Public Sub Arm(ByVal mode As String, ByVal source As String, ByVal hostPath As String, ByVal copyPrefix As String)
    mMode = mode: mSource = source: mHostPath = hostPath: mCopy = copyPrefix
    mReached = False: mCopied = False: mPlanned = "": mPreserved = False
    mQueued = "": mPayload = "": mLookup = ""
End Sub
Public Sub LookupObserved(ByVal payload As String)
    If mMode <> "" Then mPayload = payload
End Sub
Public Sub LookupReturned(ByVal eventId As String)
    If mMode <> "" Then mLookup = eventId
End Sub
Public Function LookupEventForOwnedTest() As String
    LookupEventForOwnedTest = mLookup
End Function
Public Function BeforeQueue(ByVal plannedEvent As String, ByVal payload As String) As Boolean
    Dim fso As Object, sourceHash As String, hostHash As String
    If mMode = "" Then Exit Function
    mReached = True: mPlanned = plannedEvent
    mPayload = payload
    BeforeQueue = (mMode = "QueueFailure")
    On Error GoTo CopyFailed
    sourceHash = modTrainingWire.FileSha256(mSource)
    hostHash = modTrainingWire.FileSha256(mHostPath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile mSource, mCopy & ".xlsb", False
    fso.CopyFile mHostPath, mCopy & ".xlsx", False
    mCopied = True
    mPreserved = (sourceHash <> "" And hostHash <> "" And _
        sourceHash = modTrainingWire.FileSha256(mSource) And _
        hostHash = modTrainingWire.FileSha256(mHostPath))
CopyFailed:
End Function
Public Function Facts() As String
    Facts = CStr(mReached) & "|" & CStr(mCopied) & "|" & CStr(mPreserved)
End Function
Public Function PlannedEventForOwnedTest() As String
    PlannedEventForOwnedTest = mPlanned
End Function
Public Sub AfterQueue(ByVal eventId As String)
    If mMode <> "" Then mQueued = eventId
End Sub
Public Function QueuedEventForOwnedTest() As String
    QueuedEventForOwnedTest = mQueued
End Function
Public Function PayloadForOwnedTest() As String
    PayloadForOwnedTest = mPayload
End Function
Public Function ResetQuantityForOwnedTest(ByVal warehouse As String, ByVal sku As String) As Boolean
    Dim report As String
    ResetQuantityForOwnedTest = modInventoryQuantityCommands.SetSingleEntityQuantity( _
        warehouse, "S1", "config-admin", sku, 0, "Owned intervening reset", report)
End Function
Public Function SetLaterQuantityForOwnedTest(ByVal warehouse As String, ByVal sku As String) As Boolean
    Dim report As String
    SetLaterQuantityForOwnedTest = modInventoryQuantityCommands.SetSingleEntityQuantity( _
        warehouse, "S1", "config-admin", sku, 7, "Owned later quantity change", report)
End Function
'@)
    $owner=$project.VBComponents.Item('modRoleEventWriter').CodeModule
    $start=$owner.ProcStartLine('QueueAdminInventoryAdjustEvent',0);$count=$owner.ProcCountLines('QueueAdminInventoryAdjustEvent',0)
    $body=$owner.Lines($start,$count)
    $boundary='(?im)^[\t ]*QueueAdminInventoryAdjustEvent = QueueEventCore\('
    if([regex]::Matches($body,$boundary).Count -ne 1){throw 'Owned queue observation boundary unavailable; not RED.'}
    $hook="    If TestInventoryEditReceipt.BeforeQueue(eventIdOut, payloadJson) Then`r`n        errorMessage = ""Owned queue failure.""`r`n        Exit Function`r`n    End If`r`n"
    $body=[regex]::Replace($body,$boundary,[Text.RegularExpressions.MatchEvaluator]{param($m) $hook+$m.Value})
    $body=[regex]::Replace($body,'(?im)^End Function[\t ]*\r?$',"    TestInventoryEditReceipt.AfterQueue eventIdOut`r`nEnd Function")
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $owner=$project.VBComponents.Item('modInventoryCommandReceipts').CodeModule
    $start=$owner.ProcStartLine('FindExact',0);$count=$owner.ProcCountLines('FindExact',0)
    $body=$owner.Lines($start,$count)
    $entry='    eventId = ""';$returned='    report = "": FindExact = True'
    if([regex]::Matches($body,[regex]::Escape($entry)).Count -ne 1 -or
       [regex]::Matches($body,[regex]::Escape($returned)).Count -ne 1){throw 'Exact event lookup observation unavailable; not RED.'}
    $body=$body.Replace($entry,($entry+"`r`n"+'    TestInventoryEditReceipt.LookupObserved payload'))
    $body=$body.Replace($returned,('    TestInventoryEditReceipt.LookupReturned eventId'+"`r`n"+$returned))
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
}
function Get-InventoryEditCell($Row,$Table,[string]$Header) {
    if($Row.ContainsKey($Header)){return [string]$Row[$Header]}
    return ''
}
function Get-InventoryEditRows($Table) {
    # Read each header and the data block once. Do not retain per-cell COM objects
    # across closed books; business values remain private to this owned test.
    if($Table.ListRows.Count -eq 0){return}
    $names=@($Table.ListColumns|ForEach-Object {([string]$_.Name).Trim()})
    $values=$Table.DataBodyRange.Value2
    if($values -isnot [Array] -or $values.Rank -ne 2){throw 'Owned snapshot block unavailable; not RED.'}
    for($r=$values.GetLowerBound(0);$r -le $values.GetUpperBound(0);$r++){
        $row=@{}
        for($c=$values.GetLowerBound(1);$c -le $values.GetUpperBound(1);$c++){
            $name=$names[$c-$values.GetLowerBound(1)]
            if($row.ContainsKey($name)){throw 'Ambiguous owned snapshot headers; not RED.'}
            $row[$name]=$values.GetValue($r,$c)
        }
        $row
    }
}
function Read-InventoryEditReceiptCopy([string]$Path,[string]$Sku,[string]$Command='') {
    $hash=Get-AdminInventoryEntryHash $Path
    $book=$excel.Workbooks.Open($Path,0,$true)
    $result=[pscustomobject]@{CatalogMatches=0;Item='';Reason='';History='';Custom='';Receipts=@();Applied=@();AppliedRecords=@();Log=@();Preserved=$false}
    try{
        foreach($sheet in $book.Worksheets){foreach($table in $sheet.ListObjects){
            if($table.Name -notin @('tblSkuCatalog','tblAppliedEvents','tblInventoryLog','tblInventoryEditReceipts')){continue}
            $rows=@(Get-InventoryEditRows $table)
            if($table.Name -ieq 'tblSkuCatalog'){
                foreach($row in $rows){if((Get-InventoryEditCell $row $table 'ITEM_CODE') -ceq $Sku){
                    $result.CatalogMatches++
                    $result.Item=Get-InventoryEditCell $row $table 'ITEM'
                    $result.Reason=Get-InventoryEditCell $row $table 'LAST_EDIT_REASON'
                    $result.History=Get-InventoryEditCell $row $table 'EDIT_HISTORY'
                    $result.Custom=Get-InventoryEditCell $row $table 'USER NOTE'
                }}
            }elseif($table.Name -ieq 'tblAppliedEvents'){
                foreach($row in $rows){
                    $event=Get-InventoryEditCell $row $table 'EventID'
                    $result.Applied+=,$event
                    $result.AppliedRecords+=,[pscustomobject]@{Event=$event;Status=(Get-InventoryEditCell $row $table 'Status')}
                }
            }elseif($table.Name -ieq 'tblInventoryLog'){
                foreach($row in $rows){
                    if((Get-InventoryEditCell $row $table 'SKU') -ceq $Sku){
                        $result.Log+=,[pscustomobject]@{Event=(Get-InventoryEditCell $row $table 'EventID');Key=(Get-InventoryEditCell $row $table 'System_Key');Delta=(Get-InventoryEditCell $row $table 'QtyDelta')}
                    }
                }
            }elseif($table.Name -ieq 'tblInventoryEditReceipts'){
                foreach($row in $rows){
                    $rowSku=Get-InventoryEditCell $row $table 'SKU'
                    $rowCommand=Get-InventoryEditCell $row $table 'CommandId'
                    if($rowSku -ceq $Sku -or ($Command -ne '' -and $rowCommand -ceq $Command)){
                        $result.Receipts+=,[pscustomobject]@{
                            Command=$rowCommand;Sku=$rowSku
                            Warehouse=(Get-InventoryEditCell $row $table 'WarehouseId')
                            Station=(Get-InventoryEditCell $row $table 'StationId')
                            User=(Get-InventoryEditCell $row $table 'UserId')
                            Event=(Get-InventoryEditCell $row $table 'PlannedEventId')
                        }
                    }
                }
            }
        }}
    }finally{$book.Close($false)}
    $result.Preserved=((Get-AdminInventoryEntryHash $Path) -ceq $hash)
    return $result
}
function Get-InventoryEditSavedResult([string]$Path,[int]$Row=1) {
    $hash=Get-AdminInventoryEntryHash $Path;$book=$excel.Workbooks.Open($Path,0,$true)
    try{
        $table=@($book.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
        $text=[string]$table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item($Row,1).Value2
        $match=[regex]::Match($text,'(?im)^Command:\s*([^\r\n]+)')
        $command='';if($match.Success){$command=$match.Groups[1].Value.Trim()}
        $status=[string]$table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item($Row,1).Value2
    }finally{$book.Close($false)}
    return [pscustomobject]@{Command=$command;Status=$status;Text=$text;Preserved=((Get-AdminInventoryEntryHash $Path) -ceq $hash)}
}
function Test-InventoryEditBinding($Snapshot,$Saved,$Fixture,[string]$Sku) {
    if($Saved.Command -eq ''){return $false}
    $boundReceipts=@($Snapshot.Receipts|Where-Object {$_.Command -ceq $Saved.Command})
    if($boundReceipts.Count -ne 1){return $false}
    $receipt=$boundReceipts[0]
    return ($receipt.Command -ceq $Saved.Command -and $receipt.Sku -ceq $Sku -and $receipt.Warehouse -ceq $Fixture.Warehouse -and $receipt.Station -ceq 'S1' -and $receipt.User -ceq 'config-admin')
}
function Test-InventoryEditReceiptPhases($Fixture,$Other,[ValidateSet('','Phase','Result','Canonical','Ownership')][string]$Cohort='') {
    if($Cohort -ne ''){
        $AdminInventoryEditReceiptResultsOnly=($Cohort -eq 'Result')
        $AdminInventoryEditCanonicalSaveOnly=($Cohort -eq 'Canonical')
    }
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    $caseCohort='Phase'
    if($Cohort -eq 'Ownership'){$caseCohort='Ownership'}
    if($AdminInventoryEditReceiptResultsOnly){$caseCohort='Result'}
    if($AdminInventoryEditCanonicalSaveOnly){$caseCohort='Canonical'}
    $cases=@(Get-InventoryReceiptScenarios -Cohort $caseCohort -Scenario $AdminInventoryReceiptScenarioForTest)
    [pscustomobject]@{Cohort=$caseCohort;SelectedCases=$cases;FullCohort=($AdminInventoryReceiptScenarioForTest -eq '')}|
        ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'inventory-receipt-scope.jsonl')
    foreach($case in $cases){
        SelectTarget $Fixture
        $id='InventoryEditReceipt.'+$case
        $hostBook=$excel.Workbooks.Add();$hostPath=Join-Path $runRoot ('edit-receipt-'+$case+'.xlsx');$hostBook.SaveAs($hostPath,51)
        $sourcePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        try{
            $prepared=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
            if(-not $prepared){throw 'Owned operator table unavailable; not RED.'}
            $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
            $extra=$table.ListColumns.Add();$extra.Name='USER NOTE';$extra.DataBodyRange.Cells.Item(1,1).Value2='Retained receipt extension'
            $hostBook.Save();$initial=Get-AdminInventoryEntryFacts $Fixture -Persisted
            [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('','','',''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('ReceiptSetup',$hostBook.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            Invoke-InventoryEditObservedUpload $Fixture $hostBook ($id+'.SetupAdd')
            $created=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $newKeys=@($created.EntityQuantities.Keys|Where-Object {-not $initial.EntityQuantities.ContainsKey($_)})
            $ready=$newKeys.Count -eq 1 -and $created.EntityQuantities[$newKeys[0]] -eq 0 -and $created.LogRows -eq $initial.LogRows+1 -and $created.AppliedRows -eq $initial.AppliedRows+1
            Check ($id+'.OwningAddFixtureSaved') $ready
            if(-not $ready){throw 'Owning ADD fixture unavailable; not RED.'}
            $key=[string]$newKeys[0];$sku=[string]$table.ListColumns.Item('Item Code').DataBodyRange.Cells.Item(1,1).Value2
            $item='Edited receipt '+$case;$reason='Owned receipt reason '+$case
            $table.ListColumns.Item('Action').DataBodyRange.Cells.Item(1,1).Value2='EDIT'
            $table.ListColumns.Item('Item Name').DataBodyRange.Cells.Item(1,1).Value2=$item
            $table.ListColumns.Item('Edit Reason').DataBodyRange.Cells.Item(1,1).Value2=$reason
            $quantity='';if($case -eq 'NoopQuantity'){$quantity='0'}elseif($case -notin @('CatalogOnly','NoopQuantity')){$quantity='5'}
            $table.ListColumns.Item('Quantity').DataBodyRange.Cells.Item(1,1).Value2=$quantity
            $table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2=''
            $table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2=''
            $hostBook.Save();$table.DataBodyRange.Cells.Item(1,1).Select()
            $prefix=Join-Path $runRoot ('before-queue-'+$case)
            if($AdminInventoryEditCanonicalSaveOnly){
                [void](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Arm' @($case,$sourcePath,$hostPath,$prefix))
            }
            [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @($case,$sourcePath,$hostPath,$prefix))
            $resultMode='ReceiptEdit'
            if($case -like 'OperatorSaveThrows*'){$resultMode='SaveThrows'}elseif($case -eq 'AfterApplyAuthorityLost'){$resultMode=$case}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @($resultMode,$hostBook.Name))
            Invoke-InventoryEditObservedUpload $Fixture $hostBook ($id+'.Edit')
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $reached=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq 'True'
            Check ($id+'.ActualCapturedUploadReached') $reached
            if(-not $reached){throw 'Actual EDIT action unavailable; not RED.'}
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $savedCopy=Join-Path $runRoot ('saved-edit-'+$case+'.xlsx');Copy-Item -LiteralPath $hostPath -Destination $savedCopy
            $saved=Get-InventoryEditSavedResult $savedCopy
            $snapshot=Read-InventoryEditReceiptCopy $sourcePath $sku $saved.Command
            if($AdminInventoryEditCanonicalSaveOnly){
                Test-InventoryEditCanonicalSaveOutcome $Fixture ([ref]$hostBook) $hostPath $table $created $after $snapshot $saved $sku $key $item $reason $case $prefix
                continue
            }
            Check ($id+'.SavedCatalogAndReason') ($snapshot.CatalogMatches -eq 1 -and $snapshot.Item -ceq $item -and $snapshot.Reason -ceq $reason)
            Check ($id+'.SingleCatalogApplication') ([regex]::Matches($snapshot.History,[regex]::Escape($reason)).Count -eq 1)
            Check ($id+'.ReciprocalSavedCommandAndSource') (Test-InventoryEditBinding $snapshot $saved $Fixture $sku)
            Check ($id+'.CustomValuesPreserved') ($snapshot.Custom -ceq 'Retained receipt extension' -and [string]$extra.DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Retained receipt extension')
            Check ($id+'.ReadOnlyInspectionPreservedBytes') ($snapshot.Preserved -and $saved.Preserved)
            if($case -in @('CatalogOnly','NoopQuantity','QuantityChange')){Check ($id+'.CompletionSavedInOperator') ($saved.Status -ceq 'IMPORTED')}
            $delta=0;if($case -eq 'QuantityChange' -or $AdminInventoryEditReceiptResultsOnly){$delta=1}
            $expectedQty=0;if($delta){$expectedQty=5}
            Check ($id+'.OnlyIntendedQuantityEffect') ($after.LogRows -eq $created.LogRows+$delta -and $after.AppliedRows -eq $created.AppliedRows+$delta -and $after.EntityQuantities.Count -eq $created.EntityQuantities.Count -and $after.EntityQuantities[$key] -eq $expectedQty)
            $queue=([string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Facts')).Split('|')
            $planned=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.PlannedEventForOwnedTest')
            if($case -notin @('CatalogOnly','NoopQuantity')){
                $observed=$queue.Count -eq 3 -and $queue[0] -ceq 'True' -and $queue[1] -ceq 'True' -and $queue[2] -ceq 'True'
                Check ($id+'.QueueBoundaryObservedWithoutSourceMutation') $observed
                if(-not $observed){throw 'Queue boundary snapshot unavailable; not RED.'}
                $beforeQueue=Read-InventoryEditReceiptCopy ($prefix+'.xlsb') $sku
                $beforeResult=Get-InventoryEditSavedResult ($prefix+'.xlsx')
                $joint=Test-InventoryEditBinding $beforeQueue $beforeResult $Fixture $sku
                Check ($id+'.CatalogReceiptAndPlannedEventSavedBeforeQueue') ($joint -and $planned -ne '' -and $beforeQueue.Receipts[0].Event -ceq $planned -and $beforeQueue.Item -ceq $item -and $beforeQueue.Reason -ceq $reason -and $beforeQueue.Preserved -and $beforeResult.Preserved)
                if($case -eq 'QuantityChange' -or $AdminInventoryEditReceiptResultsOnly){
                    $exact=@($snapshot.Log|Where-Object {$_.Event -ceq $planned -and $_.Key -ceq $key -and [double]$_.Delta -eq 5})
                    Check ($id+'.PlannedEventHasExactAppliedAndLogProof') ($planned -ne '' -and @($snapshot.Applied|Where-Object {$_ -ceq $planned}).Count -eq 1 -and $exact.Count -eq 1)
                }
            }else{
                # This requires an explicit saved no-event receipt, not merely
                # an absent queue call. A missing receipt fails this assertion.
                Check ($id+'.NoInventedQuantityEvent') ($queue[0] -ceq 'False' -and $snapshot.Receipts.Count -eq 1 -and $snapshot.Receipts[0].Event -eq '')
            }
            if($case -eq 'QueueFailure' -or $AdminInventoryEditReceiptResultsOnly){
                if($case -eq 'QueueFailure'){Check ($id+'.IncompleteQuantityNotMarkedImported') ($saved.Status -cne 'IMPORTED')}
                if($AdminInventoryEditReceiptResultsOnly){
                    $boundary=([string](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Facts')).Split('|')
                    $boundaryReached=$boundary.Count -eq 4 -and $boundary[0] -ceq 'True' -and $boundary[1] -ceq 'True'
                    Check ($id+'.ActualResultBoundaryReached') $boundaryReached
                    if(-not $boundaryReached){throw 'Owned result failure or sign-out boundary unavailable; not RED.'}
                    Check ($id+'.OperatorBytesPreservedAfterBoundary') ($boundary[2] -ceq 'True' -and $boundary[3] -ceq '0')
                    Check ($id+'.HonestSavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.WorksheetResultPartialForTest'))
                    Check ($id+'.UnconfirmedResultNotDurablyImported') ($saved.Status -cne 'IMPORTED')
                }
                $hostBook.Close($false);$hostBook=$excel.Workbooks.Open($hostPath,0,$false)
                $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                SelectTarget $Fixture;$hostBook.Activate();$table.DataBodyRange.Cells.Item(1,1).Select()
                $retryBefore=$after;$retryTarget=5;$retryDelta=1
                if($case -eq 'QueueFailure' -and ($AdminInventoryEditReceiptAllOnly -or $AdminInventoryEditOwnershipOnly)){
                    Test-InventoryEditPendingOwnership $Fixture $hostBook $sourcePath $after
                }
                if($AdminInventoryEditReceiptResultsOnly){$retryDelta=0}
                if($case -eq 'OperatorSaveThrowsLaterQuantity'){
                    [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('','','',''))
                    $laterDiagnosis=Start-InventoryEditConfigDiagnosis $Fixture
                    $moved=[bool](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.SetLaterQuantityForOwnedTest' @($Fixture.Warehouse,$sku))
                    Complete-InventoryEditConfigDiagnosis $Fixture $laterDiagnosis ($id+'.LaterQuantity')
                    $retryBefore=Get-AdminInventoryEntryFacts $Fixture -Persisted;$retryTarget=7
                    $moved=$moved -and $retryBefore.EntityQuantities[$key] -eq 7 -and $retryBefore.LogRows -eq $after.LogRows+1 -and $retryBefore.AppliedRows -eq $after.AppliedRows+1
                    Check ($id+'.LaterOwningQuantityChangeSaved') $moved
                    if(-not $moved){throw 'Later owning quantity fixture unavailable; not RED.'}
                    $hostBook.Activate();$table.DataBodyRange.Cells.Item(1,1).Select()
                }
                [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('Retry',$sourcePath,$hostPath,($prefix+'-retry')))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('ReceiptRetry',$hostBook.Name))
                Invoke-InventoryEditObservedUpload $Fixture $hostBook ($id+'.Retry')
                $retry=Get-AdminInventoryEntryFacts $Fixture -Persisted
                $retrySnapshot=Read-InventoryEditReceiptCopy $sourcePath $sku
                $retryEvent=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.PlannedEventForOwnedTest')
                $retryFacts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
                Check ($id+'.FreshRetryActualCapturedHandlerReached') ($retryFacts.Count -eq 5 -and $retryFacts[0] -ceq 'True' -and $retryFacts[2] -ceq '1' -and $retryFacts[3] -ceq 'True')
                if($case -eq 'QueueFailure'){
                    Check ($id+'.FreshRetryKeepsPlannedEvent') ($planned -ne '' -and $retryEvent -ceq $planned)
                }else{
                    $retryQueue=([string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Facts')).Split('|')
                    Check ($id+'.AppliedRetryNeverQueuesAgain') ($retryQueue[0] -ceq 'False')
                    Check ($id+'.FreshRetryKeepsPlannedEvent') ($planned -ne '' -and $retrySnapshot.Receipts.Count -eq 1 -and $retrySnapshot.Receipts[0].Event -ceq $planned)
                }
                Check ($id+'.FreshRetryDoesNotRepeatCatalog') ($retrySnapshot.History -ceq $snapshot.History -and $retrySnapshot.Item -ceq $item -and $retrySnapshot.Preserved)
                Check ($id+'.FreshRetryAppliesQuantityOnce') ($retry.EntityQuantities.Count -eq $retryBefore.EntityQuantities.Count -and $retry.EntityQuantities[$key] -eq $retryTarget -and $retry.LogRows -eq $retryBefore.LogRows+$retryDelta -and $retry.AppliedRows -eq $retryBefore.AppliedRows+$retryDelta)
                $exact=@($retrySnapshot.Log|Where-Object {$_.Event -ceq $planned -and $_.Key -ceq $key -and [double]$_.Delta -eq 5})
                Check ($id+'.RetryPlannedEventHasExactAppliedAndLogProof') ($planned -ne '' -and @($retrySnapshot.Applied|Where-Object {$_ -ceq $planned}).Count -eq 1 -and $exact.Count -eq 1)
            }
        }finally{$hostBook.Close($false)}
    }
    Test-InventoryConfigDiagnosticPreservation
}
