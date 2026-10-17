# Operator Edit with independent saved results; first23 use owning event fixtures.
# Separate completion cases use the explicitly scoped synthetic catalog-only fixture.
. (Join-Path $PSScriptRoot 'Slice4beInventoryCatalogCompletion.ps1')
function Install-InventoryOwnResultProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $admin.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function OwnResultShowsForTest() As Long
    OwnResultShowsForTest = mShows
End Function
Public Function OwnResultPartialForTest() As Boolean
    OwnResultPartialForTest = (InStr(1, mLastNotice, "Exact event application was saved:", vbTextCompare) > 0 _
        And InStr(1, mLastNotice, "Inventory quantity set.", vbTextCompare) = 0)
End Function
Public Function OwnResultReportFactsForTest() As String
    OwnResultReportFactsForTest = CStr(InStr(1, mLastNotice, "Partial result:", vbTextCompare) > 0) & "|" & _
        CStr(InStr(1, mLastNotice, "unsaved changes", vbTextCompare) > 0) & "|" & _
        CStr(InStr(1, mLastNotice, "read-only", vbTextCompare) > 0) & "|" & _
        CStr(InStr(1, mLastNotice, "source", vbTextCompare) > 0)
End Function
'@)
    $trace=$admin.VBComponents.Add(1);$trace.Name='TestInventoryCatalogTrace'
    $trace.CodeModule.AddFromString(@'
Option Explicit
Private mPhase As String
Public Sub Mark(ByVal phase As String)
    mPhase = phase
End Sub
Public Function Phase() As String
    Phase = mPhase
End Function
'@)
    $owner=$admin.VBComponents.Item('modAdmin').CodeModule
    $start=$owner.ProcStartLine('UpdateInventoryItemCatalogForWarehouse',0)
    $count=$owner.ProcCountLines('UpdateInventoryItemCatalogForWarehouse',0)
    $body=$owner.Lines($start,$count)
    foreach($point in @(
        @('    On Error GoTo FailUpdate','Entered'),
        @('        If Not modInventoryCommands.BeginCatalog(interactiveContext, sku, report) Then GoTo CleanExit','CatalogBegun'),
        @('        If Not modInventoryCommands.BindCatalogWriter(interactiveContext, wb.Name, report) Then GoTo CleanExit','WriterBound'),
        @('        Next field','FieldsPlanned'),
        @('        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then GoTo CleanExit','BeforeWritePassed'),
        @('    WriteInventoryCatalogFieldsAdmin lo, rowIndex, fields','FieldsWritten'),
        @('        If Not modInventoryCommands.BeginCatalogSave(interactiveContext, report) Then GoTo CleanExit','BeforeSavePassed'),
        @('    wb.Save','SaveReturned'),
        @('    UpdateInventoryItemCatalogForWarehouse = True','Confirmed')
    )){
        # VBE may recase identifiers when it inserts source into a project.
        $pattern='(?im)^'+[regex]::Escape($point[0])+'[\t ]*\r?$'
        if([regex]::Matches($body,$pattern).Count -ne 1){throw ('Catalog fixed-stage anchor unavailable: '+$point[1])}
        $mark='    TestInventoryCatalogTrace.Mark "'+$point[1]+'"'
        $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+$mark})
    }
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $core=$packages['invSys.Core.xlam'].VBProject
    $valueTrace=$core.VBComponents.Add(1);$valueTrace.Name='TestInventoryValueTrace'
    $valueTrace.CodeModule.AddFromString(@'
Option Explicit
Private mChecks As Long, mTextMismatch As Boolean, mRawMatches As Boolean
Public Sub Reset()
    mChecks = 0: mTextMismatch = False: mRawMatches = False
End Sub
Public Sub Observe(ByVal actual As Variant, ByVal expected As String)
    If Not IsNumeric(actual(0)) Or Not IsDate(expected) Then Exit Sub
    mChecks = mChecks + 1
    mTextMismatch = (CDbl(actual(0)) <> CDbl(CDate(expected)))
    mRawMatches = (CDbl(actual(3)) = CDbl(CDate(expected)))
End Sub
Public Function Facts() As String
    Facts = CStr(mChecks) & "|" & CStr(mTextMismatch) & "|" & CStr(mRawMatches)
End Function
'@)
    $tableInput=$core.VBComponents.Item('cInventoryTableInput').CodeModule
    $inputText=$tableInput.Lines(1,$tableInput.CountOfLines)
    $rawAnchor=' & text & ":" & formula)'
    $rawPattern='(?i)'+[regex]::Escape($rawAnchor)
    # The extra observation slot is unused by the unchanged production guard.
    if([regex]::Matches($inputText,$rawPattern).Count -eq 1){
        $inputText=[regex]::Replace($inputText,$rawPattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.Substring(0,$m.Value.Length-1)+', cell.Value2)'})
    }elseif([regex]::Matches($inputText,('(?i)'+[regex]::Escape(' & text & ":" & formula, cell.Value2)'))).Count -ne 1){
        throw 'Read-only raw-value observation anchor unavailable.'
    }
    $matchAnchor='(?im)^[\t ]*If PlannedValueMatches\(actual\(header\), plannedValue\) Then GoTo NextField[\t ]*\r?$'
    if([regex]::Matches($inputText,$matchAnchor).Count -ne 1){throw 'Planned-value observation anchor unavailable.'}
    $inputText=[regex]::Replace($inputText,$matchAnchor,[Text.RegularExpressions.MatchEvaluator]{param($m) '                            If header = "LAST_EDIT_AT" Then TestInventoryValueTrace.Observe actual(header), plannedValue'+"`r`n"+$m.Value.TrimEnd("`r")})
    $tableInput.DeleteLines(1,$tableInput.CountOfLines);$tableInput.AddFromString($inputText)
    $probe=$core.VBComponents.Add(1);$probe.Name='TestInventoryOwnResult'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mSignOut As Boolean, mReturned As Boolean, mConfirmation As String
Public Sub Arm(ByVal signOut As Boolean)
    mSignOut = signOut: mReturned = False
    mConfirmation = ""
    TestInventoryValueTrace.Reset
End Sub
Public Sub AfterProcessor()
    mReturned = True
    If mSignOut Then modAuth.SignOut
End Sub
Public Function Returned() As Boolean
    Returned = mReturned
End Function
Public Sub ConfirmationStage(ByVal stage As String)
    mConfirmation = stage
End Sub
Public Function Confirmation() As String
    Confirmation = mConfirmation
End Function
'@)
    $inputs=$core.VBComponents.Item('cInventoryCommandInputs').CodeModule
    $inputCode=$inputs.Lines(1,$inputs.CountOfLines)
    if($inputCode -match '(?im)^Private Function ConfirmSavedInputs'){
        $start=$inputs.ProcStartLine('ConfirmSavedInputs',0);$count=$inputs.ProcCountLines('ConfirmSavedInputs',0)
        $body=$inputs.Lines($start,$count)
        foreach($point in @(
            @('    On Error GoTo Invalid','Entered'),
            @('    If Not mSaving Then Exit Function','Saving'),
            @('    Set book = ReadSource(owner, borrowed)','SourceRead'),
            @('    If Not book.Saved Then Exit Function','SourceSaved'),
            @('    Set persisted = modAggregationFiles.OpenCopy(mPath, copyPath, True)','CopyOpenReturned'),
            @('    If modTrainingWire.FileSha256(copyPath) <> savedHash Then GoTo CleanExit','CopyHashVerified'),
            @('        If Not PlanSavedQuantity(persisted, eventId) Then GoTo CleanExit','ReceiptVerified'),
            @('    If Not mCatalog.MatchesPersistedCopy(modExecutionTarget.Table(persisted, "tblSkuCatalog")) Then GoTo CleanExit','SavedCatalogVerified'),
            @('    If Not mEntities.MatchesPersistedCopy(modExecutionTarget.Table(persisted, "tblInventoryEntities")) Then GoTo CleanExit','SavedEntitiesVerified'),
            @('    If modAggregationFiles.Identity(mPath) <> savedIdentity Or modTrainingWire.FileSha256(mPath) <> savedHash Then GoTo CleanExit','StableSourceVerified'),
            @('    If Not Unchanged() Then GoTo CleanExit','LiveInputsVerified'),
            @('    If Not mCatalog.Confirm(modExecutionTarget.Table(book, "tblSkuCatalog")) Then GoTo CleanExit','CatalogAdopted'),
            @('    If Not mEntities.Confirm(modExecutionTarget.Table(book, "tblInventoryEntities")) Then GoTo CleanExit','EntitiesAdopted'),
            @('    mIdentity = savedIdentity: mHash = savedHash','BaselineBound')
        )){
            $pattern='(?im)^[\t ]*'+[regex]::Escape($point[0].Trim())+'[\t ]*\r?$'
            if([regex]::Matches($body,$pattern).Count -ne 1){throw ('Confirmation diagnostic anchor unavailable: '+$point[1])}
            $mark='    If mPhase = "Quantity" Then TestInventoryOwnResult.ConfirmationStage "'+$point[1]+'"'
            $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+$mark})
        }
        $inputs.DeleteLines($start,$count);$inputs.InsertLines($start,$body)
    }
    $module=$core.VBComponents.Item('modInventoryQuantityCommands').CodeModule
    $text=$module.Lines(1,$module.CountOfLines)
    $boundary='    processed = modProcessor.RunBatch(warehouse, 0, batchReport, eventId, applied, True)'
    if([regex]::Matches($text,[regex]::Escape($boundary)).Count -ne 1){throw 'Quantity processor-return test boundary unavailable.'}
    $module.DeleteLines(1,$module.CountOfLines)
    $module.AddFromString($text.Replace($boundary,($boundary+"`r`n"+'    TestInventoryOwnResult.AfterProcessor')))
    Install-InventoryCatalogCompletionProbe
}

function Test-InventoryOwnResults($Fixture,$Other) {
    SelectTarget $Fixture
    $sku='DEMO-OWN-RESULT'
    $csv=Join-Path $runRoot 'own-result-fixture.csv'
    @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION',($sku+',Own result fixture,LB,CLEARVIEW,10,GOOD'))|
        Set-Content $csv -Encoding ASCII
    $created=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))
    Check 'InventoryOwnResult.GeneratedSingleEntityFixtureApplied' $created
    if(-not $created){throw 'Owning own-result fixture unavailable; not product RED.'}
    $original=Get-AdminInventoryEntryFacts $Fixture -Persisted
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$excel.Workbooks.Open($path,0,$true)
    try{
        $table=Table $book 'tblInventoryEntities';$entityKeys=@()
        foreach($row in $table.ListRows){
            if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index).Value2 -ceq $sku -and [string]$row.Range.Cells.Item(1,$table.ListColumns.Item('InventoryState').Index).Value2 -ceq 'ACTIVE'){
                $entityKeys+=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2
            }
        }
    }finally{$book.Close($false)}
    if($entityKeys.Count -ne 1){throw 'Own-result SKU requires exactly one generated active entity.'}
    $key=$entityKeys[0]
    foreach($spec in @(@('NoOp',10),@('Applied',7),@('AuthorityLost',4))){
        $case=[string]$spec[0];$target=[double]$spec[1];$id='InventoryOwnResult.'+$case
        SelectTarget $Fixture
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('own-result-'+$case+'.xlsx')),51)
        try{
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            [void](Run 'invSys.Core.xlam' 'TestInventoryOwnResult.Arm' @(($case -ceq 'AuthorityLost')))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetQuantityTarget' @($target))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Edit','SetQuantity','EditReason',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualEditHandlerAndReasonReached') $ready
            if(-not $ready){throw 'Own-result actual Edit handler unavailable; not product RED.'}
            $shows=[long](Run 'invSys.Admin.xlam' 'TestInventoryCommands.OwnResultShowsForTest')
            $catalogPhase=[string](Run 'invSys.Admin.xlam' 'TestInventoryCatalogTrace.Phase')
            if($catalogPhase -cnotin @('Entered','CatalogBegun','WriterBound','FieldsPlanned','BeforeWritePassed','FieldsWritten','BeforeSavePassed','SaveReturned','Confirmed')){throw 'Unknown catalog diagnostic stage.'}
            $noticeFacts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.OwnResultReportFactsForTest')).Split('|')
            $dateFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryValueTrace.Facts')).Split('|')
            $confirmation=[string](Run 'invSys.Core.xlam' 'TestInventoryOwnResult.Confirmation')
            if($confirmation -and $confirmation -cnotin @('Entered','Saving','SourceRead','SourceSaved','CopyOpenReturned','CopyHashVerified','ReceiptVerified','SavedCatalogVerified','SavedEntitiesVerified','StableSourceVerified','LiveInputsVerified','CatalogAdopted','EntitiesAdopted','BaselineBound')){throw 'Unknown confirmation stage.'}
            [pscustomobject]@{Case=$case;ConfirmationStage=$confirmation}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'own-result-confirmation-stages.jsonl')
            [pscustomobject]@{Case=$case;CatalogPhase=$catalogPhase;FormShows=$shows;PartialNotice=($noticeFacts[0] -ceq 'True');DirtyNotice=($noticeFacts[1] -ceq 'True');ReadOnlyNotice=($noticeFacts[2] -ceq 'True');SourceNotice=($noticeFacts[3] -ceq 'True');DateComparisons=[int]$dateFacts[0];DateTextRoundTripMismatch=($dateFacts[1] -ceq 'True');RawDateMatches=($dateFacts[2] -ceq 'True')}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'own-result-fixed-stages.jsonl')
            Check ($id+'.ExpectedFormContinuation') ($shows -eq $(if($case -ceq 'AuthorityLost'){1}else{2}))
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $othersSame=@($original.EntityQuantities.Keys|Where-Object {$_ -cne $key -and (-not $after.EntityQuantities.ContainsKey($_) -or $after.EntityQuantities[$_] -ne $original.EntityQuantities[$_])}).Count -eq 0
            Check ($id+'.ExactEntityAndTargetSaved') ($after.EntityQuantities.Count -eq $original.EntityQuantities.Count -and $after.EntityQuantities.ContainsKey($key) -and $after.EntityQuantities[$key] -eq $target -and $othersSame)
            $events=if($case -ceq 'NoOp'){0}else{1}
            Check ($id+'.ExactEventCountSaved') ($after.LogRows -eq $before.LogRows+$events -and $after.AppliedRows -eq $before.AppliedRows+$events)
            $returned=[bool](Run 'invSys.Core.xlam' 'TestInventoryOwnResult.Returned')
            Check ($id+'.ExpectedProcessorReturn') ($returned -eq ($case -cne 'NoOp'))
            if($case -ceq 'AuthorityLost'){
                Check ($id+'.SavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.OwnResultPartialForTest'))
            }
        }finally{$hostBook.Close($false)}
    }
    Test-InventoryCatalogCompletionResults $Fixture $Other
}
