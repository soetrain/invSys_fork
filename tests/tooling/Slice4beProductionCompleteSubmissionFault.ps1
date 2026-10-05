# Disposable writer faults preserve real appends, sync/processor results and handlers.
function Install-ProductionCompleteSubmissionFaultProbe {
    $writer=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modRoleEventWriter').CodeModule
    function InsertFault([string]$Procedure,[string]$Anchor,[string]$Code,[switch]$After){
        $start=$writer.ProcStartLine($Procedure,0);$end=$start+$writer.ProcCountLines($Procedure,0)
        $hits=@(for($line=$start;$line -lt $end;$line++){if($writer.Lines($line,1).Trim() -ceq $Anchor){$line}})
        if($hits.Count -ne 1){throw ('Completion writer seam changed: '+$Procedure+'; not product RED.')}
        $writer.InsertLines($hits[0]+[int][bool]$After,$Code)
    }
    $writer.InsertLines(1,@'
Private mCompleteFaultRoot As String, mCompleteFaultMode As String, mCompleteFaultWarehouse As String
Private mCompleteFaultTarget As Long, mCompleteFaultCalls As Long, mCompleteFaultReached As Boolean
Private mCompleteFaultAllocated As String, mCompleteFaultPrinted As String, mCompleteFaultHold As Integer
'@)
    InsertFault 'LocalStagingRootRole' 'rootPath = Trim$(Environ$("LOCALAPPDATA"))' '    If mCompleteFaultRoot <> "" Then LocalStagingRootRole = mCompleteFaultRoot: Exit Function'
    InsertFault 'AppendInboxRowToLocalStagingRole' 'writeAttemptedOut = True' '    CompleteFaultBeforeForTest rowValues'
    InsertFault 'AppendInboxRowToLocalStagingRole' 'Print #fileNum, serializedRow' '    CompleteFaultAfterForTest rowValues, stagingPath' -After
    $writer.AddFromString(@'
Public Sub CompleteFaultArmForTest(ByVal root As String, ByVal mode As String, ByVal ordinal As Long, ByVal warehouse As String)
    CompleteFaultResetForTest
    mCompleteFaultRoot = root: mCompleteFaultMode = mode: mCompleteFaultTarget = ordinal
    mCompleteFaultWarehouse = warehouse
    mCompleteFaultCalls = 0: mCompleteFaultReached = False
    mCompleteFaultAllocated = "": mCompleteFaultPrinted = ""
End Sub
Private Sub CompleteFaultBeforeForTest(ByVal row As Object)
    If mCompleteFaultMode = "" Then Exit Sub
    If CStr(row("WarehouseId")) <> mCompleteFaultWarehouse Then Err.Raise vbObjectError + 7860, , "Completion fixture warehouse mismatch."
    If CStr(row("EventType")) <> "PROD_CONSUME" And CStr(row("EventType")) <> "PROD_COMPLETE" Then _
        Err.Raise vbObjectError + 7860, , "Unexpected completion fixture writer."
    mCompleteFaultCalls = mCompleteFaultCalls + 1
    If mCompleteFaultAllocated <> "" Then mCompleteFaultAllocated = mCompleteFaultAllocated & vbLf
    mCompleteFaultAllocated = mCompleteFaultAllocated & CStr(row("EventID"))
    If mCompleteFaultCalls = mCompleteFaultTarget And mCompleteFaultMode = "BeforeAppend" Then
        mCompleteFaultReached = True
        Err.Raise vbObjectError + 7861, , "Injected completion pre-append failure."
    End If
End Sub
Private Sub CompleteFaultAfterForTest(ByVal row As Object, ByVal path As String)
    Dim marker As String
    If mCompleteFaultMode = "" Then Exit Sub
    If mCompleteFaultPrinted <> "" Then mCompleteFaultPrinted = mCompleteFaultPrinted & vbLf
    mCompleteFaultPrinted = mCompleteFaultPrinted & CStr(row("EventID"))
    If mCompleteFaultCalls <> mCompleteFaultTarget Then Exit Sub
    If mCompleteFaultMode = "AfterAppend" Then
        mCompleteFaultReached = True
        Err.Raise vbObjectError + 7862, , "Injected completion acknowledgment failure."
    ElseIf mCompleteFaultMode = "Pending" Then
        ' A held recovery file prevents the real synchronizer from merging rows.
        ' It contains no business row and never substitutes a processor result.
        mCompleteFaultHold = FreeFile
        Open path & ".syncing" For Binary Access Read Write Lock Read Write As #mCompleteFaultHold
        marker = "Disposable completion synchronization hold"
        Put #mCompleteFaultHold, , marker
        mCompleteFaultReached = True
    End If
End Sub
Public Function CompleteFaultEvidenceForTest() As Variant
    CompleteFaultEvidenceForTest = Array(mCompleteFaultAllocated, mCompleteFaultPrinted, _
        mCompleteFaultCalls, mCompleteFaultReached, (mCompleteFaultHold <> 0))
End Function
Public Sub CompleteFaultResetForTest()
    If mCompleteFaultHold <> 0 Then Close #mCompleteFaultHold
    mCompleteFaultHold = 0: mCompleteFaultMode = "": mCompleteFaultRoot = ""
End Sub
'@)
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function CompleteFaultWorksheetSelectionForTest() As Variant
    Dim output As ListObject, key As String
    Set output = ProductionTable(TABLE_MANAGER_OUTPUT)
    key = CellByHeader(output, 1, "System_Key")
    CompleteFaultWorksheetSelectionForTest = Array(mRunSheetKey, 10#, key)
End Function
Public Function CompleteFaultWorksheetOutputAbsentForTest() As Boolean
    Dim output As ListObject, key As String
    Set output = ProductionTable(TABLE_MANAGER_OUTPUT)
    key = CellByHeader(output, 1, "System_Key")
    If key = "" Or key = mRunSheetKey Then Exit Function
    CompleteFaultWorksheetOutputAbsentForTest = _
        Abs(CompleteWorksheetExactQtyForTest(key)) < 0.0000001 And HasProductionCheckRows()
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function CompleteFaultWorksheetSelection() As Variant
    CompleteFaultWorksheetSelection = mForm.CompleteFaultWorksheetSelectionForTest()
End Function
Public Function CompleteFaultWorksheetOutputAbsent() As Boolean
    CompleteFaultWorksheetOutputAbsent = mForm.CompleteFaultWorksheetOutputAbsentForTest()
End Function
'@)
}

function Test-ProductionCompleteSubmissionFault($Fixture,$Other,$Book,$Decoy,[string]$Canary){
    $queueBase=[IO.Path]::GetFullPath((Join-Path $runRoot 'complete-fault-queues'))
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not $queueBase.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Completion queues escaped their disposable fixture.'}
    $otherBefore=@(Get-Slice4beActivityFiles $Other)
    foreach($branch in @('Reusable','Worksheet')){foreach($ordinal in 1,2){foreach($fault in @('BeforeAppend','AfterAppend','Pending')){
        $work=$null;$label='CompleteWriterFault.'+$branch+'.'+$ordinal+'.'+$fault
        $queue=Join-Path $queueBase ([guid]::NewGuid().ToString('N'))
        try{
            SelectTarget $Fixture 'config-producer'
            $work=$excel.Workbooks.Add();$sheet=$work.Worksheets.Item(1)
            $sheet.Cells.Item(2,1).Value2=$Canary;$sheet.Cells.Item(2,2).Formula='=1+2'
            [void](Probe 'RunLocalReopen' @($work.Name))
            if($branch -ceq 'Reusable'){
                $ready=[bool](Probe 'CompleteBaselinePrepare' @($true))
                $inputSelection=Owner 'CompleteSubmissionInputForTest'
            }else{
                $ready=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CompleteWorksheetInventoryForTest' @($work.Name))
                if($ready){$ready=[bool](Probe 'CompleteWorksheetPrepare' @($Canary))}
                $inputSelection=Probe 'CompleteFaultWorksheetSelection'
            }
            if(-not $ready -or $inputSelection -isnot [array] -or [string]$inputSelection[0] -ceq '' -or [double]$inputSelection[1] -le 0){throw 'Completion writer fixture unavailable; not product RED.'}
            [void](Probe 'RunLocalShowAndCapture' @($work.Name,'CHECK_IN'));$Decoy.Activate()
            if($branch -ceq 'Worksheet' -and $ordinal -eq 1 -and $fault -ceq 'BeforeAppend'){
                Test-ProductionCompleteWorksheetPermission $Fixture $work $Decoy $Canary
            }
            $before=@(Get-Slice4beActivityFiles $Fixture);$pins=@{}
            foreach($file in $before){$pins[$file]=Hash $file}
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.CompleteFaultArmForTest' @($queue,$fault,$ordinal,$Fixture.Warehouse))
            if($branch -ceq 'Worksheet'){
                $notice=Invoke-CompleteWorksheetPermissionNotice -Kind Submission
                $returned=$notice.Returned
            }else{$returned=[bool](Probe 'CompleteEntryAct' @(''))}
            $facts=Run 'invSys.Core.xlam' 'modRoleEventWriter.CompleteFaultEvidenceForTest'
            $calls=if($branch -ceq 'Worksheet' -and $fault -ceq 'Pending'){2}else{$ordinal}
            if($facts -isnot [array] -or $facts.Count -ne 5 -or -not [bool]$facts[3] -or [int]$facts[2] -ne $calls){throw ('Actual writer fault boundary unavailable: '+$branch+'/'+$ordinal+'/'+$fault+'; not product RED.')}
            $allocated=@(([string]$facts[0]).Split("`n")|Where-Object{$_})
            $printed=@(([string]$facts[1]).Split("`n")|Where-Object{$_})
            $printedCount=if($fault -ceq 'BeforeAppend'){$ordinal-1}else{$calls}
            if($allocated.Count -ne $calls -or $printed.Count -ne $printedCount -or @($allocated|Select-Object -Unique).Count -ne $calls){throw 'Writer identity evidence is incomplete; not product RED.'}
            Check ($label+'.ActualHandlerReturned') $returned
            if($branch -ceq 'Worksheet'){Check ($label+'.NativeFailureAcknowledged') $notice.CompletionNotice}
            Check ($label+'.RealFaultAndAppendBoundary') ($printed.Count -eq $printedCount -and [bool]$facts[4] -eq ($fault -ceq 'Pending'))
            Check ($label+'.GuardsRestored') ([bool](Probe 'CompleteEntryFact' @('GuardsRestored')))
            $refs=@();$states=@()
            for($index=0;$index -lt $printed.Count;$index++){
                $refs+=$printed[$index]
                $states+=if($fault -ceq 'AfterAppend' -and $index -eq ($ordinal-1)){'Unknown'}else{'Submitted'}
            }
            $outcome=if($fault -ceq 'Pending'){'PENDING'}else{'FAILED'}
            Pair $before $outcome $label $refs -ReferenceStates $states
            $raw=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
            $unusedHidden=$true
            foreach($id in @($allocated|Where-Object{$_ -cnotin $printed})){foreach($wire in $raw){$unusedHidden=$unusedHidden -and -not $wire.Contains($id)}}
            Check ($label+'.UnattemptedIdentitiesNotExposed') $unusedHidden
            # Drop the override before any later ordinary work can synchronize these pending rows.
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.CompleteFaultResetForTest')
            $consumed=$branch -ceq 'Reusable' -and $ordinal -eq 2
            if($branch -ceq 'Reusable'){
                Check ($label+'.ExactInputEffect') ([bool](Owner 'CompleteBaselineBalancesForTest' @($consumed)))
                Check ($label+'.NoCompletedOutputOrSuccessState') ([bool](Owner 'CompleteSubmissionOutputsAbsentForTest'))
                $output=Owner 'CompleteSubmissionOutputSelectionForTest';$outputKey=[string]$output[0]
            }else{
                Check ($label+'.ExactInputEffect') ([bool](Probe 'CompleteWorksheetFact' @('InputUnchanged',$Canary)))
                Check ($label+'.NoCompletedOutputOrSuccessState') ([bool](Probe 'CompleteFaultWorksheetOutputAbsent'))
                $output=Probe 'CompleteFaultWorksheetSelection';$outputKey=[string]$output[2]
                foreach($fact in @('OutputCustom','CheckCustom')){Check ($label+'.'+$fact) ([bool](Probe 'CompleteWorksheetFact' @($fact,$Canary)))}
            }
            $safe=$outputKey -cne ''
            foreach($wire in $raw){foreach($private in @([string]$inputSelection[0],$outputKey,$queue,'Injected completion')){if($private -cne '' -and $wire.Contains($private)){$safe=$false}}}
            Check ($label+'.NoBusinessIdentityOrRawFaultInObservation') $safe
            $activeRoot=Join-Path (Join-Path $queue $Fixture.Warehouse) 'S1'
            $staged=@(Get-ChildItem -LiteralPath $activeRoot -File -Filter '*.staging.jsonl'|ForEach-Object{[IO.File]::ReadAllLines($_.FullName)}|Where-Object{$_}|ForEach-Object{$_|ConvertFrom-Json})
            $expectedStaged=@($printed|Where-Object{-not($consumed -and $_ -ceq $allocated[0])})
            $exact=@($staged|ForEach-Object EventID)
            $staging=$staged.Count -eq $expectedStaged.Count -and ($exact -join '|') -ceq ($expectedStaged -join '|')
            foreach($row in $staged){$staging=$staging -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.EventType -cin @('PROD_CONSUME','PROD_COMPLETE')}
            Check ($label+'.ExactRetainedQueueRows') $staging
            Test-CompleteFaultAuthority $Fixture $allocated $inputSelection $consumed $label
            $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
            Check ($label+'.PriorRecordsImmutable') $same
            Check ($label+'.NoRedirectedActivity') ((@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
            Check ($label+'.CustomValueAndFormula') ($sheet.Cells.Item(2,1).Value2 -ceq $Canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
            Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
            if($ordinal -eq 2 -and $fault -cin @('AfterAppend','Pending')){
                CaptureOwnedFormByCaptionEvidence 'Production' ('complete-writer-'+$branch.ToLowerInvariant()+'-'+$fault.ToLowerInvariant()+'.png')
                if($fault -ceq 'Pending'){Test-CompleteStatusScrolling $label ('complete-writer-'+$branch.ToLowerInvariant()+'-pending-scroll')}
            }
        }finally{
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.CompleteFaultResetForTest')
            [void](Probe 'CloseDesigner');if($null -ne $work){$work.Close($false)}
        }
    }}}
    SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
}

function Test-CompleteFaultAuthority($Fixture,[string[]]$Allocated,$InputSelection,[bool]$Consumed,[string]$Label){
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $before=@($excel.Workbooks|Where-Object{$_.FullName -ieq $path});$pin=Hash $path
    if($before.Count -gt 1){throw 'Duplicate authority during writer audit.'}
    $opened=$before.Count -eq 0
    if($opened){$authority=$excel.Workbooks.Open($path,0,$true)}else{$authority=$before[0]}
    try{
        $applied=Table $authority 'tblAppliedEvents';$audit=Table $authority 'tblInventoryLog';$exact=$true
        for($i=0;$i -lt $Allocated.Count;$i++){
            $id=$Allocated[$i];$expected=if($Consumed -and $i -eq 0){1}else{0}
            $events=@($applied.ListRows|Where-Object{$_.Range.Cells.Item(1,$applied.ListColumns.Item('EventID').Index).Value2 -ceq $id})
            $rows=@($audit.ListRows|Where-Object{$_.Range.Cells.Item(1,$audit.ListColumns.Item('EventID').Index).Value2 -ceq $id})
            $exact=$exact -and $events.Count -eq $expected -and $rows.Count -eq $expected
            if($expected -eq 1 -and $rows.Count -eq 1){$exact=$exact -and $rows[0].Range.Cells.Item(1,$audit.ListColumns.Item('System_Key').Index).Value2 -ceq $InputSelection[0] -and $rows[0].Range.Cells.Item(1,$audit.ListColumns.Item('QtyDelta').Index).Value2 -eq (-[double]$InputSelection[1])}
        }
        Check ($Label+'.ExactAppliedEventsAndAuditEffects') $exact
    }finally{if($opened){$authority.Close($false)}}
    Check ($Label+'.AuditWorkbookBytesPreserved') ((Hash $path) -ceq $pin)
    Check ($Label+'.AuditWorkbookLifetimePreserved') (@($excel.Workbooks|Where-Object{$_.FullName -ieq $path}).Count -eq $before.Count)
}
