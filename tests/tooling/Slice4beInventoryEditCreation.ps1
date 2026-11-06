# D8-INV-EDIT-02. Only Admin-generated disposable warehouses and real Upload handlers.
# Identity/payload values stay in memory or ignored owned copies; reports are fixed checks.
. (Join-Path $PSScriptRoot 'Slice4beInventoryConfigPreservation.ps1')
. (Join-Path $PSScriptRoot 'Slice4beInventoryWorksheetResults.ps1')
. (Join-Path $PSScriptRoot 'Slice4beInventoryEditReceipts.ps1')
. (Join-Path $PSScriptRoot 'Slice4beInventoryEditBindings.ps1')
. (Join-Path $PSScriptRoot 'Slice4beInventoryEditCanonicalSave.ps1')
function Close-ReceiptFixtureInventory($Fixture) {
    $source=[IO.Path]::GetFullPath((Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')))
    $ownedRoot=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    if(-not $source.StartsWith($ownedRoot,[StringComparison]::OrdinalIgnoreCase)){
        throw 'Receipt fixture Inventory path is outside its owned root.'
    }
    $books=$excel.Workbooks
    try{
        for($index=$books.Count;$index -ge 1;$index--){
            $book=$books.Item($index);$match=$false
            try{
                $match=[string]::Equals([IO.Path]::GetFullPath([string]$book.FullName),$source,[StringComparison]::OrdinalIgnoreCase)
                if($match){
                    if($book.ReadOnly){throw 'Receipt fixture Inventory is open read-only.'}
                    if(-not $book.Saved){$book.Save();if(-not $book.Saved){throw 'Receipt fixture Inventory save was not confirmed.'}}
                    $book.Close($false)
                }
            }finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
            if($match){break}
        }
    }finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($books)}
    if(-not (Test-Path -LiteralPath $source -PathType Leaf) -or (Get-AdminInventoryEntryHash $source) -eq ''){
        throw 'Receipt fixture Inventory persistence was not confirmed.'
    }
}
function Install-InventoryEditCreationProbe {
    Install-InventoryConfigPreservationProbe
    Install-InventoryWorksheetResultProbe
    Install-InventoryEditBindingsProbe
    Install-InventoryEditCanonicalSaveProbe
    Install-InventoryEditReceiptProbe
    Install-InventoryCreationNoticeProbe
    $project=$packages['invSys.Core.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryEditCreation'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mSource As String, mHost As String, mCopy As String
Private mQueueReached As Boolean, mBatchBlocked As Boolean, mCopied As Boolean
Private mPlanned As String, mQueued As String
Public Sub Arm(ByVal mode As String, ByVal source As String, ByVal host As String, ByVal copy As String)
    mMode = mode: mSource = source: mHost = host: mCopy = copy
    mQueueReached = False: mBatchBlocked = False: mCopied = False
    mPlanned = "": mQueued = ""
End Sub
Public Function BeforeQueue(ByVal planned As String) As Boolean
    Dim fso As Object, a As String, b As String
    If mMode = "" Then Exit Function
    mQueueReached = True: mPlanned = planned
    BeforeQueue = (mMode = "QueueFailure" Or mMode Like "Intervening*")
    a = modTrainingWire.FileSha256(mSource): b = modTrainingWire.FileSha256(mHost)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile mSource, mCopy & ".xlsb", False
    fso.CopyFile mHost, mCopy & ".xlsx", False
    mCopied = (a <> "" And b <> "" And a = modTrainingWire.FileSha256(mSource) And b = modTrainingWire.FileSha256(mHost))
End Function
Public Sub AfterQueue(ByVal eventId As String, ByVal succeeded As Boolean)
    If mMode <> "" And succeeded Then mQueued = eventId
End Sub
Public Function BlockBatch(ByVal eventId As String) As Boolean
    If (mMode = "ApplyFailure" Or mMode Like "Queued*") And mQueued <> "" And mQueued = eventId Then
        mBatchBlocked = True: BlockBatch = True
    End If
End Function
Public Function Facts() As String
    Facts = CStr(mQueueReached) & "|" & CStr(mCopied) & "|" & CStr(mBatchBlocked)
End Function
Public Function PlannedForOwnedTest() As String
    PlannedForOwnedTest = mPlanned
End Function
Public Function QueuedForOwnedTest() As String
    QueuedForOwnedTest = mQueued
End Function
Public Function LaterQuantity(ByVal warehouse As String, ByVal sku As String) As Boolean
    Dim report As String
    LaterQuantity = modInventoryQuantityCommands.SetSingleEntityQuantity( _
        warehouse, "S1", "config-admin", sku, 9, "Owned later quantity", report)
End Function
Public Function ApplyQueued(ByVal warehouse As String, ByVal eventId As String) As Boolean
    Dim report As String, processed As Long, applied As Boolean
    processed = modProcessor.RunBatch(warehouse, 0, report, eventId, applied, True)
    ApplyQueued = applied
End Function
'@)
    $owner=$project.VBComponents.Item('modRoleEventWriter').CodeModule
    $start=$owner.ProcStartLine('QueueInventoryCreateEvent',0);$count=$owner.ProcCountLines('QueueInventoryCreateEvent',0)
    $body=$owner.Lines($start,$count)
    $boundary='(?im)^[\t ]*QueueInventoryCreateEvent = QueueEventCore\('
    if([regex]::Matches($body,$boundary).Count -ne 1){throw 'Creation queue seam unavailable; not RED.'}
    $hook="    If TestInventoryEditCreation.BeforeQueue(eventIdOut) Then`r`n        errorMessage = ""Owned CREATE queue failure."": Exit Function`r`n    End If`r`n"
    $body=[regex]::Replace($body,$boundary,[Text.RegularExpressions.MatchEvaluator]{param($m) $hook+$m.Value})
    $body=[regex]::Replace($body,'(?im)^End Function[\t ]*\r?$',"    TestInventoryEditCreation.AfterQueue eventIdOut, QueueInventoryCreateEvent`r`nEnd Function")
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $owner=$project.VBComponents.Item('modProcessor').CodeModule
    $start=$owner.ProcStartLine('RunBatch',0);$count=$owner.ProcCountLines('RunBatch',0)
    $body=$owner.Lines($start,$count);$anchor='    receiptApplied = False'
    if([regex]::Matches($body,[regex]::Escape($anchor)).Count -ne 1){throw 'Exact processor seam unavailable; not RED.'}
    $body=$body.Replace($anchor,($anchor+"`r`n"+'    If TestInventoryEditCreation.BlockBatch(receiptEventId) Then report = "Owned application failure.": Exit Function'))
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestInventoryCatalogCompletion').CodeModule.AddFromString(@'
Public Function ConfigureCreationCase(ByVal host As String, ByVal tableName As String, ByVal mode As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, col As ListColumn
    For Each ws In Application.Workbooks(host).Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                If mode = "ZeroQuantity" Then lo.ListColumns("Quantity").DataBodyRange.Cells(1, 1).Value2 = 0#
                If mode = "DuplicateRows" Then
                    For Each col In lo.ListColumns
                        col.DataBodyRange.Cells(2, 1).Value2 = col.DataBodyRange.Cells(1, 1).Value2
                    Next col
                End If
                ConfigureCreationCase = True: Exit Function
            End If
        Next lo
    Next ws
End Function
Public Function CreateOtherEntity(ByVal warehouse As String, ByVal sku As String, ByVal retire As Boolean) As Boolean
    Dim report As String, count As Long, eventId As String
    If Not modAdminInventoryCreate.CompleteCatalogItem(warehouse, "S1", "config-admin", sku, _
        "Synthetic catalog-only fixture", "LB", "CLEARVIEW", 2, "Owned intervening creation", report) Then Exit Function
    If retire Then
        If Not modInventoryRetirementCommands.RetireActiveEntities(warehouse, "S1", "config-admin", sku, _
            "Owned retired history", count, eventId, report) Then Exit Function
        If count <> 1 Then Exit Function
    End If
    CreateOtherEntity = True
End Function
Public Function ChangeLaterCatalog(ByVal warehouse As String, ByVal sku As String, Optional ByVal mode As String = "InterveningCatalog") As Boolean
    Dim report As String, custom As Object, location As String
    Set custom = CreateObject("Scripting.Dictionary")
    custom.Add "USER NOTE", "Owned later custom field"
    location = "CLEARVIEW"
    If mode = "InterveningLocation" Then location = "OWNED-OTHER"
    If mode = "InterveningQuantityMode" Then
        custom.Add "TRACK_QTY", "FALSE": custom.Add "ITEM_KIND", "UTILITY"
    End If
    ChangeLaterCatalog = modAdmin.UpdateInventoryItemCatalogForWarehouse(warehouse, sku, _
        "Owned later catalog item", "LB", location, , , , , , , custom, report, "", "config-admin", "S1")
End Function
'@)
    $packages['invSys.Inventory.Domain.xlam'].VBProject.VBComponents.Item('modInventoryApply').CodeModule.AddFromString(@'
Public Function TestRetiredCatalogFixture(ByVal bookName As String, ByVal sku As String) As Boolean
    Dim lo As ListObject, row As Long
    Set lo = FindListObjectByNameApply(Application.Workbooks(bookName), "tblSkuCatalog")
    row = FindRowByColumnValueApply(lo, "SKU", sku)
    If row = 0 Or Left$(sku, 24) <> "DEMO-CATALOG-COMPLETION-" Then Exit Function
    lo.ListColumns("CATALOG_STATE").DataBodyRange.Cells(row, 1).Value2 = "RETIRED"
    TestRetiredCatalogFixture = True
End Function
'@)
}
function Invoke-InventoryCreationUpload($Fixture,$HostBook,[string]$Id,[string]$AuthorityCase='Authorized') {
    $diagnosis=Start-InventoryEditConfigDiagnosis $Fixture
    [void](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.ResetTrackingNotices')
    [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($HostBook.Name,'Upload',$AuthorityCase,'ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
    $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
    $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq 'True'
    Check ($Id+'.ActualCapturedUploadReached') $ready
    if(-not $ready){throw 'Actual Upload unavailable; not RED.'}
    $notices=([string](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.TrackingNoticeFacts')).Split('|')
    if($notices.Count -ne 3){throw 'Tracking notice observer unavailable; not RED.'}
    if($Id -ceq 'InventoryEditCreation.AuthorityLost'){
        Check ($Id+'.StaleActionTrackingNoticeObserved') ([int]$notices[0] -ge 1 -and $notices[1] -ceq 'True' -and $notices[2] -ceq 'False')
    }elseif($AuthorityCase -in @('ReplacementSession','OtherWarehouseAdmin')){
        Check ($Id+'.ChangedAuthorityTrackingNoticeObserved') ([int]$notices[0] -ge 1 -and $notices[2] -ceq 'False')
    }elseif([int]$notices[0] -gt 0){
        Check ($Id+'.UnexpectedTrackingWarning') $false
    }
    Complete-InventoryEditConfigDiagnosis $Fixture $diagnosis $Id
    Check ($Id+'.ConfigBytesPreserved') ($diagnosis.Config.FileHash -ceq (Get-AdminInventoryEntryHash $Fixture.Config))
    $saves=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
    Check ($Id+'.NoConfigSave') ([int]$saves[0] -eq 0)
}
function Install-InventoryCreationNoticeProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $probe=$admin.VBComponents.Item('TestInventoryCatalogCompletion').CodeModule
    $probe.InsertLines($probe.CountOfDeclarationLines+1,'Private mTrackingNotices As Long, mStaleActionNotice As Boolean, mUnexpectedNotice As Boolean')
    $probe.AddFromString(@'
Public Sub ResetTrackingNotices()
    mTrackingNotices = 0: mStaleActionNotice = False: mUnexpectedNotice = False
End Sub
Public Sub ObserveTrackingNotice(ByVal text As String, ByVal buttons As VbMsgBoxStyle, ByVal title As String)
    mTrackingNotices = mTrackingNotices + 1
    If Trim$(text) = "Tracking unavailable: the action context is no longer current." Then mStaleActionNotice = True
    If Left$(text, Len("Tracking unavailable: ")) <> "Tracking unavailable: " Or buttons <> vbExclamation Or title <> "invSys Admin" Then mUnexpectedNotice = True
End Sub
Public Function TrackingNoticeFacts() As String
    TrackingNoticeFacts = CStr(mTrackingNotices) & "|" & CStr(mStaleActionNotice) & "|" & CStr(mUnexpectedNotice)
End Function
'@)
    $admin.VBComponents.Item('cProvisioningActivity').CodeModule.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
        Optional ByVal Title As String = "") As VbMsgBoxResult
    TestInventoryCatalogCompletion.ObserveTrackingNotice Prompt, Buttons, Title
    MsgBox = vbOK
End Function
'@)
}
function Get-InventoryCreationSavedResult($Book,[string]$Label) {
    $copy=Join-Path $runRoot ($Label+'.xlsx');Copy-Item -LiteralPath $Book.FullName -Destination $copy
    Get-InventoryEditSavedResult $copy
}
function Test-InventoryCreationIdentityLoss($Fixture,$HostBook,[string]$Sku,$Saved) {
    $source=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($source).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Identity mutations require an owned fixture.'}
    $backup=Join-Path $runRoot 'creation-receipt-before-mutations.xlsb';Copy-Item -LiteralPath $source -Destination $backup
    $table=@($HostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
    $cases=@('LostEventReference','LostReceiptEvent','LostReceiptKey','ChangedReceiptKey','MissingReceipt','DuplicateReceipt')
    foreach($case in $cases){
        SelectTarget $Fixture
        $table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2=$Saved.Text
        $table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2='IMPORTED'
        $HostBook.Save()
        try{
            if($case -eq 'LostEventReference'){
                $text=[regex]::Replace($Saved.Text,'(?im)^Event:[^\r\n]*\r?\n?','')
                if($text -ceq $Saved.Text){throw 'Produced event reference unavailable; not RED.'}
                $table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2=$text
                $HostBook.Save()
            }else{
                $book=$excel.Workbooks.Open($source,0,$false)
                try{
                    $receipts=Table $book 'tblInventoryEditReceipts';$rows=@()
                    for($row=1;$row -le $receipts.ListRows.Count;$row++){
                        if([string]$receipts.ListColumns.Item('CommandId').DataBodyRange.Cells.Item($row,1).Value2 -ceq $Saved.Command){$rows+=,$row}
                    }
                    if($rows.Count -ne 1){throw 'Actual command receipt unavailable; not RED.'}
                    $row=$rows[0]
                    switch($case){
                        'MissingReceipt' {$receipts.ListRows.Item($row).Delete()}
                        'DuplicateReceipt' {$extra=$receipts.ListRows.Add();[void]$receipts.ListRows.Item($row).Range.Copy($extra.Range)}
                        default {
                            $field=if($case -eq 'LostReceiptEvent'){'PlannedEventId'}else{'System_Key'}
                            $value=if($case -eq 'ChangedReceiptKey'){[guid]::NewGuid().ToString().ToUpperInvariant()}else{''}
                            $receipts.ListColumns.Item($field).DataBodyRange.Cells.Item($row,1).Value2=$value
                        }
                    }
                    $book.Save()
                }finally{$book.Close($false)}
            }
            $before=Get-AdminInventoryEntryHash $source
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('IdentityLoss',$HostBook.Name))
            Invoke-InventoryCreationUpload $Fixture $HostBook ('InventoryEditCreation.IdentityLoss.'+$case)
            Check ('InventoryEditCreation.IdentityLoss.'+$case+'.NoCanonicalWriteOrReplacementIdentity') ($before -ceq (Get-AdminInventoryEntryHash $source))
            Check ('InventoryEditCreation.IdentityLoss.'+$case+'.HonestRefusalNotice') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.EditBindingRefusalForTest'))
            $result=Get-InventoryCreationSavedResult $HostBook ('identity-'+$case)
            Check ('InventoryEditCreation.IdentityLoss.'+$case+'.OriginalCommandRetained') ($result.Command -ceq $Saved.Command)
        }finally{
            Copy-Item -LiteralPath $backup -Destination $source -Force
        }
    }
    $table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2=$Saved.Text
    $table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2=$Saved.Status
    $HostBook.Save()
}
function Test-InventoryEditCreation($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    SelectTarget $Fixture
    $csv=Join-Path $runRoot 'edit-creation-baseline.csv'
    @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION','DEMO-CREATE-BASE,Creation baseline,LB,CLEARVIEW,3,GOOD')|Set-Content $csv -Encoding ASCII
    if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))){throw 'Owning seed unavailable; not RED.'}
    if($AdminInventoryEditCreationCase -eq 'RetainedReceipts'){
        Close-ReceiptFixtureInventory $Fixture
        Test-InventoryEditReceiptPhases $Fixture $Other 'Phase'
        return
    }
    $source=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $safety=@('BeforeCatalogSaveThrows','AfterCatalogSaveThrows','ReplacementSession','OtherWarehouseAdmin','RetiredHistory','InterveningEntity','InterveningCatalog','InterveningLocation','InterveningQuantityMode')
    $queuedCases=@('QueuedLocation','QueuedQuantityMode')
    $cases=@('Success','QueueFailure','ApplyFailure','ResultSaveFailure','AuthorityLost','ZeroQuantity','DuplicateRows','RetiredCatalog')+$safety+$queuedCases
    if($AdminInventoryEditCreationCase -eq 'Safety'){$cases=$safety}
    elseif($AdminInventoryEditCreationCase -eq 'QuantityPhase'){$cases=@('InterveningCatalog','InterveningLocation','InterveningQuantityMode')+$queuedCases}
    elseif($AdminInventoryEditCreationCase -eq 'InterveningInputs'){$cases=@('InterveningLocation','InterveningQuantityMode')}
    elseif($AdminInventoryEditCreationCase -eq 'QueuedInputs'){$cases=$queuedCases}
    elseif($AdminInventoryEditCreationCase -ne 'All'){$cases=@($AdminInventoryEditCreationCase)}
    foreach($case in $cases){
        SelectTarget $Fixture
        [void](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.Arm' @('','','',''))
        $id='InventoryEditCreation.'+$case;$sku='DEMO-CATALOG-COMPLETION-'+$case.ToUpperInvariant()
        $book=$excel.Workbooks.Open($source,0,$false)
        try{
            if(-not [bool](Run 'invSys.Inventory.Domain.xlam' 'modInventoryApply.TestCatalogOnlyFixture' @($book.Name,$sku))){throw 'Catalog fixture unavailable; not RED.'}
            if($case -eq 'RetiredCatalog' -and -not [bool](Run 'invSys.Inventory.Domain.xlam' 'modInventoryApply.TestRetiredCatalogFixture' @($book.Name,$sku))){throw 'Retired catalog fixture unavailable; not RED.'}
            $book.Save()
        }finally{$book.Close($false)}
        if($case -eq 'RetiredHistory'){
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.CreateOtherEntity' @($Fixture.Warehouse,$sku,$true))){throw 'Owning retired-history fixture unavailable; not RED.'}
        }
        $hostBook=$excel.Workbooks.Add();$hostPath=Join-Path $runRoot ('edit-creation-'+$case+'.xlsx');$hostBook.SaveAs($hostPath,51)
        try{
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))){throw 'Operator fixture unavailable; not RED.'}
            $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.PrepareCatalogEditTable' @($hostBook.Name,$table.Name,$sku))){throw 'Typed EDIT fixture unavailable; not RED.'}
            $extra=$table.ListColumns.Add();$extra.Name='USER NOTE';$extra.DataBodyRange.Cells.Item(1,1).Value2='Retained creation extension'
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.ConfigureCreationCase' @($hostBook.Name,$table.Name,$case))){throw 'Creation case unavailable; not RED.'}
            $hostBook.Save();Set-InventoryConfigPresentation $Fixture
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted;$beforeHash=Get-AdminInventoryEntryHash $source
            $prefix=Join-Path $runRoot ('creation-before-queue-'+$case)
            [void](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.Arm' @($case,$source,$hostPath,$prefix))
            $saveMode=if($case -like '*CatalogSaveThrows'){$case}else{''}
            [void](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Arm' @($saveMode,$source,$hostPath,($prefix+'-canonical')))
            $resultMode=switch($case){'ResultSaveFailure'{'SaveThrows'} 'AuthorityLost'{'PreflightAuthorityLost'} default {'Creation'}}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @($resultMode,$hostBook.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
            $authorityCase=if($case -in @('ReplacementSession','OtherWarehouseAdmin')){$case}else{'Authorized'}
            Invoke-InventoryCreationUpload $Fixture $hostBook $id $authorityCase
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $saved=Get-InventoryCreationSavedResult $hostBook ($case+'-saved')
            $snapshot=Read-InventoryEditReceiptCopy $source $sku $saved.Command
            Test-InventoryConfigPresentation $Fixture $id
            Check ($id+'.ReadOnlyInspection') ($saved.Preserved -and $snapshot.Preserved)
            Check ($id+'.UnknownOperatorColumnPreserved') ([string]$extra.DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Retained creation extension')
            if($case -in @('AuthorityLost','ZeroQuantity','DuplicateRows','RetiredCatalog','ReplacementSession','OtherWarehouseAdmin','RetiredHistory')){
                Check ($id+'.WholeSelectionRefusedWithoutCanonicalWrite') ($beforeHash -ceq (Get-AdminInventoryEntryHash $source) -and $saved.Status -cne 'IMPORTED')
                if($case -eq 'AuthorityLost'){
                    $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Facts')).Split('|')
                    Check ($id+'.AuthorityYieldReached') ($facts[0] -ceq 'True')
                }
                continue
            }
            if($saveMode){
                $saveFacts=([string](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Facts')).Split('|')
                $observed=$saveFacts.Count -eq 6 -and $saveFacts[0] -ceq 'True' -and $saveFacts[1] -ceq 'True'
                Check ($id+'.ActualCanonicalSaveBoundaryReached') $observed
                $queueFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.Facts')).Split('|')
                Check ($id+'.NoCreationAfterUnconfirmedSave') ($queueFacts[0] -ceq 'False' -and $after.LogRows -eq $before.LogRows -and $after.EntityQuantities.Count -eq $before.EntityQuantities.Count -and $saved.Status -cne 'IMPORTED')
                if(-not $observed){continue}
                Check ($id+'.SaveUncertaintyDisclosed') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.CanonicalSaveUnconfirmedForTest'))
                if($case -eq 'BeforeCatalogSaveThrows'){
                    Check ($id+'.UnsavedCatalogAndReceiptPreserved') ($beforeHash -ceq (Get-AdminInventoryEntryHash $source) -and $snapshot.Receipts.Count -eq 0)
                    $originalCommand=$saved.Command
                    $hostBook.Close($false);$hostBook=$excel.Workbooks.Open($hostPath,0,$false)
                    SelectTarget $Fixture
                    [void](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Arm' @('',$source,$hostPath,''))
                    Invoke-InventoryCreationUpload $Fixture $hostBook ($id+'.ReopenedWithoutReceipt')
                    $retryResult=Get-InventoryCreationSavedResult $hostBook ($case+'-retry')
                    Check ($id+'.MissingBoundIdentityRefusedWithoutRegeneration') ($originalCommand -ne '' -and $retryResult.Command -ceq $originalCommand -and $beforeHash -ceq (Get-AdminInventoryEntryHash $source) -and $retryResult.Status -cne 'IMPORTED')
                    continue
                }
            }
            $bound=Test-InventoryEditBinding $snapshot $saved $Fixture $sku
            Check ($id+'.BoundCreateReceiptSaved') ($bound -and $snapshot.Receipts[0].Intent -ceq 'CREATE')
            Check ($id+'.OneCatalogApplication') ([regex]::Matches($snapshot.History,'Config preservation completion').Count -eq 1)
            Check ($id+'.CanonicalCustomFieldPreserved') ($snapshot.Custom -ceq 'Retained creation extension')
            $facts=([string](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.Facts')).Split('|')
            if(-not $saveMode){Check ($id+'.QueueBoundaryReached') ($facts[0] -ceq 'True')}
            if(-not $bound){
                # The required owning output is absent. Never fabricate receipts or
                # claim dependent corruption/retry checks were executed.
                [pscustomobject]@{Case=$case;DependentChecksExecuted=$false;Reason='NoActualBoundReceipt'}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'creation-dependent-scope.jsonl')
                continue
            }
            $receipt=$snapshot.Receipts[0];$planned=[string](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.PlannedForOwnedTest')
            $queued=[string](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.QueuedForOwnedTest')
            Check ($id+'.ImmutableDistinctPlannedIdentities') ($receipt.Event -ne '' -and $receipt.Key -ne '' -and $receipt.Event -cne $receipt.Key -and $receipt.Command -cne $receipt.Event -and $receipt.Command -cne $receipt.Key -and ($saveMode -or $planned -ceq $receipt.Event))
            if(-not $saveMode){
            $pre=Read-InventoryEditReceiptCopy ($prefix+'.xlsb') $sku $saved.Command
            $preResult=Get-InventoryEditSavedResult ($prefix+'.xlsx')
            Check ($id+'.CatalogReceiptAndOperatorSavedBeforeQueue') ($facts[1] -ceq 'True' -and (Test-InventoryEditBinding $pre $preResult $Fixture $sku) -and $pre.Receipts[0].Event -ceq $receipt.Event -and $pre.Receipts[0].Key -ceq $receipt.Key -and $pre.Log.Count -eq 0)
            }
            if($case -in @('QueueFailure','ApplyFailure') -or $case -like 'Intervening*' -or $case -like 'Queued*'){
                Check ($id+'.CreationNotAppliedOrReportedImported') ($after.EntityQuantities.Count -eq $before.EntityQuantities.Count -and $after.LogRows -eq $before.LogRows -and $saved.Status -cne 'IMPORTED')
                Check ($id+'.ExpectedQueueAndApplicationBoundary') ($(if($case -eq 'ApplyFailure' -or $case -like 'Queued*'){$queued -ceq $receipt.Event -and $facts[2] -ceq 'True'}else{$queued -eq ''}))
            }elseif(-not $saveMode){
                Check ($id+'.ExactFreshCreationApplied') ((Test-InventoryExactAppliedEffect $snapshot $receipt.Event $receipt.Key 7) -and $after.EntityQuantities.Count -eq $before.EntityQuantities.Count+1 -and $after.EntityQuantities[$receipt.Key] -eq 7)
                Check ($id+'.ExpectedOperatorSaveOutcome') ($(if($case -eq 'Success'){$saved.Status -ceq 'IMPORTED'}else{$saved.Status -cne 'IMPORTED'}))
            }
            $hostBook.Close($false);$hostBook=$null
            SelectTarget $Fixture
            [void](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.Arm' @('','','',''))
            [void](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Arm' @('',$source,$hostPath,''))
            if($case -like 'Queued*'){
                $change=if($case -eq 'QueuedLocation'){'InterveningLocation'}else{'InterveningQuantityMode'}
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.ChangeLaterCatalog' @($Fixture.Warehouse,$sku,$change))){throw 'Owning post-queue catalog update unavailable; not RED.'}
                $priorApply=Read-InventoryEditReceiptCopy $source $sku $saved.Command
                $applied=[bool](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.ApplyQueued' @($Fixture.Warehouse,$receipt.Event))
                $proof=Read-InventoryEditReceiptCopy $source $sku $saved.Command
                Check ($id+'.HeadlessApplicationPreservesLaterCatalog') ($proof.Item -ceq $priorApply.Item -and $proof.Custom -ceq $priorApply.Custom -and $proof.Location -ceq $priorApply.Location -and $proof.TrackQty -ceq $priorApply.TrackQty -and $proof.ItemKind -ceq $priorApply.ItemKind)
                Check ($id+'.HeadlessApplicationRetainsReceiptIdentities') ($proof.Receipts.Count -eq 1 -and $proof.Receipts[0].Command -ceq $receipt.Command -and $proof.Receipts[0].Event -ceq $receipt.Event -and $proof.Receipts[0].Key -ceq $receipt.Key)
                if($case -eq 'QueuedQuantityMode'){
                    Check ($id+'.ChangedCountingModeCannotApplyDifferentEffect') (-not $applied -and $proof.Log.Count -eq 0 -and @($proof.AppliedRecords|Where-Object Event -CEQ $receipt.Event).Count -eq 0)
                }else{
                    Check ($id+'.HeadlessExactCreationApplied') ($applied -and (Test-InventoryExactAppliedEffect $proof $receipt.Event $receipt.Key 7))
                    $hostBook=$excel.Workbooks.Open($hostPath,0,$false)
                    Invoke-InventoryCreationUpload $Fixture $hostBook ($id+'.ReopenedAfterHeadlessApply')
                    $savedRetry=Get-InventoryCreationSavedResult $hostBook ($case+'-retry')
                    $retried=Read-InventoryEditReceiptCopy $source $sku $savedRetry.Command
                    Check ($id+'.ImportedRetryDoesNotReapplyEitherPhase') ($savedRetry.Status -ceq 'IMPORTED' -and $savedRetry.Command -ceq $saved.Command -and $retried.Log.Count -eq 1 -and $retried.Item -ceq $priorApply.Item -and $retried.Location -ceq $priorApply.Location -and $retried.Custom -ceq $priorApply.Custom)
                }
                continue
            }
            if($case -in @('Success','ResultSaveFailure')){
                if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryEditCreation.LaterQuantity' @($Fixture.Warehouse,$sku))){throw 'Owning later quantity unavailable; not RED.'}
            }
            if($case -eq 'InterveningEntity'){
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.CreateOtherEntity' @($Fixture.Warehouse,$sku,$false))){throw 'Owning intervening creation unavailable; not RED.'}
                Set-InventoryConfigPresentation $Fixture
            }
            if($case -in @('InterveningCatalog','InterveningLocation','InterveningQuantityMode')){
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.ChangeLaterCatalog' @($Fixture.Warehouse,$sku,$case))){throw 'Owning later catalog update unavailable; not RED.'}
            }
            $priorRetry=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $hostBook=$excel.Workbooks.Open($hostPath,0,$false)
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('CreationRetry',$hostBook.Name))
            Invoke-InventoryCreationUpload $Fixture $hostBook ($id+'.ReopenedRetry')
            $retried=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $savedRetry=Get-InventoryCreationSavedResult $hostBook ($case+'-retry')
            $proof=Read-InventoryEditReceiptCopy $source $sku $savedRetry.Command
            if($case -in @('InterveningEntity','InterveningLocation','InterveningQuantityMode')){
                Check ($id+'.InterveningEntityRefusedWithoutReplacement') ($retried.FileHash -ceq $priorRetry.FileHash -and $savedRetry.Status -cne 'IMPORTED' -and $savedRetry.Command -ceq $saved.Command -and $proof.Receipts.Count -eq 1 -and $proof.Receipts[0].Event -ceq $receipt.Event -and $proof.Receipts[0].Key -ceq $receipt.Key)
                Check ($id+'.SavedCatalogPartialDisclosed') ($savedRetry.Text -match '(?i)catalog.*receipt.*saved' -and $savedRetry.Text -match '(?i)unconfirmed')
                continue
            }
            $expectedEvents=if($case -in @('QueueFailure','ApplyFailure','AfterCatalogSaveThrows','InterveningCatalog')){1}else{0}
            $expectedQty=if($expectedEvents){7}else{9}
            Check ($id+'.RetryRetainsCommandEventAndKey') ($savedRetry.Command -ceq $saved.Command -and $proof.Receipts.Count -eq 1 -and $proof.Receipts[0].Event -ceq $receipt.Event -and $proof.Receipts[0].Key -ceq $receipt.Key)
            Check ($id+'.RetryExactApplicationWithoutDuplication') ((Test-InventoryExactAppliedEffect $proof $receipt.Event $receipt.Key 7) -and $retried.LogRows -eq $priorRetry.LogRows+$expectedEvents -and $retried.AppliedRows -eq $priorRetry.AppliedRows+$expectedEvents -and $retried.EntityQuantities[$receipt.Key] -eq $expectedQty)
            Check ($id+'.RetryImportedAndOneCatalogApplication') ($savedRetry.Status -ceq 'IMPORTED' -and [regex]::Matches($proof.History,'Config preservation completion').Count -eq 1)
            if($case -eq 'InterveningCatalog'){
                Check ($id+'.QuantityRetryPreservesLaterCatalogAndCustomEdits') ($proof.Item -ceq 'Owned later catalog item' -and $proof.Custom -ceq 'Owned later custom field')
            }
            if($case -eq 'Success'){Test-InventoryCreationIdentityLoss $Fixture $hostBook $sku $savedRetry}
        }finally{if($null -ne $hostBook){$hostBook.Close($false)}}
    }
    Test-InventoryConfigDiagnosticPreservation
    if($AdminInventoryEditCreationCase -eq 'All'){
        Test-InventoryEditReceiptPhases $Fixture $Other 'Phase'
    }
}
