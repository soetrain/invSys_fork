# Actual Print Recall handler; missing output naturally refuses before print preview.
. (Join-Path $PSScriptRoot 'Slice4beProductionPrintRebuild.ps1')
. (Join-Path $PSScriptRoot 'Slice4beProductionPrintInventory.ps1')
function Install-ProductionPrintBaselineProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $module=$project.VBComponents.Item('mProduction').CodeModule
    $line=$module.ProcBodyLine('BtnPrintRecallCodes',0)
    $module.InsertLines($line+1,'    TestProductionDesigner.PrintOwnerHit')
    $start=$module.ProcStartLine('BuildRecallCodesReportFromCurrentWorkbook',0)
    $end=$start+$module.ProcCountLines('BuildRecallCodesReportFromCurrentWorkbook',0)
    $anchor=@(for($i=$start;$i -lt $end;$i++){if($module.Lines($i,1).Trim() -ceq 'Set wsProd = SheetExists(SHEET_PRODUCTION)'){$i+1}})
    if($anchor.Count -ne 1){throw 'Print report read anchor changed; not product RED.'}
    $module.InsertLines($anchor[0],'    TestProductionDesigner.PrintSheetHit wsProd')
    Install-ProductionPrintRebuildProbe $module
    Install-ProductionPrintInventoryProbe $module
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function PrintActForTest() As Boolean
    On Error GoTo Failed
    mBtnManagerPrint_Click
    PrintActForTest = True
Failed:
End Function
Public Function PrintStatusForTest() As String
    PrintStatusForTest = mTxtStatus.Text
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mPrintOwners As Long, mPrintReads As Long, mPrintBook As Workbook')
    $adapter.AddFromString(@'
Public Sub PrintOwnerHit()
    mPrintOwners = mPrintOwners + 1
End Sub
Public Sub PrintSheetHit(ByVal sheet As Worksheet)
    mPrintReads = mPrintReads + 1
    If Not sheet Is Nothing Then Set mPrintBook = sheet.Parent
End Sub
Public Function PrintAct() As Boolean
    mPrintOwners = 0: mPrintReads = 0: Set mPrintBook = Nothing
    PrintAct = mForm.PrintActForTest()
End Function
Public Function PrintOwnerEntries() As Long
    PrintOwnerEntries = mPrintOwners
End Function
Public Function PrintReportReads() As Long
    PrintReportReads = mPrintReads
End Function
Public Function PrintReadCapturedBook(ByVal name As String) As Boolean
    If mPrintBook Is Nothing Then Exit Function
    PrintReadCapturedBook = (mPrintBook Is Application.Workbooks(name))
End Function
Public Function PrintStatus() As String
    PrintStatus = mForm.PrintStatusForTest()
End Function
'@)
}

function Test-ProductionPrintBaseline($Fixture,$Other){
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){(Get-FileHash -LiteralPath $Path).Hash}
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'),[ref]$tokens,[ref]$errors)
    $definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Start-DialogCaptureAndDismiss'},$false)
    if($errors.Count -or $null -eq $definition){throw 'Native refusal observer unavailable; not product RED.'}
    . ([scriptblock]::Create($definition.Extent.Text))
    $canary='PRINT-FIXTURE';$book=$null;$decoy=$null;$observer=$null
    $otherPins=RestartPins $Other.Root
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$sheet.Name='Production'
        $sheet.Cells.Item(1,1).Value2=$canary;$sheet.Cells.Item(2,1).Formula='=1+2'
        $path=Join-Path $runRoot 'print-binding.xlsb'
        if(Test-Path -LiteralPath $path){throw 'Preserve existing saved Print fixture.'}
        $book.SaveAs($path,50);$book.Close($false);$pin=Hash $path
        $book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$otherSheet=$decoy.Worksheets.Item(1);$otherSheet.Name='Production'
        $otherSheet.Cells.Item(1,1).Value2=$canary;$otherSheet.Cells.Item(2,1).Formula='=4+5'
        foreach($case in @('Current','Target','Session','SignedOut','MissingSheet')){
            SelectTarget $Fixture 'config-producer'
            [void](Probe 'OpenDesigner' @($book.Name))
            [void](Probe 'RunLocalShowAndCapture' @($book.Name,'PRINT'))
            $decoy.Activate()
            if($case -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($case -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($case -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($case -ceq 'MissingSheet'){$sheet.Name='PrintMissing'}
            $processes=@(Get-Process EXCEL);if($processes.Count -ne 1){throw 'Isolated Excel required.'}
            $stop=Join-Path $runRoot ('print-notice-'+$case)
            $observer=Start-DialogCaptureAndDismiss -ExcelProcessId $processes[0].Id -TimeoutSeconds 30 -StopPath $stop
            $label='PrintBaseline.'+$case
            try{Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintAct'))}finally{
                [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
                $notice=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
                if($observer.State -ne 'Completed'){Stop-Job $observer;throw 'Native observer did not close normally.'}
                Remove-Job $observer;$observer=$null
            }
            if($case -ceq 'Current'){
                Check ($label+'.OriginalOwnerEnteredOnce') ([int](Probe 'PrintOwnerEntries') -eq 1)
                Check ($label+'.BothExistingReportReadsRetained') ([int](Probe 'PrintReportReads') -eq 2)
                Check ($label+'.ReadsCapturedBookWithDecoyActive') ([bool](Probe 'PrintReadCapturedBook' @($book.Name)))
                Check ($label+'.ExistingNativeRefusal') ($notice.Contains('ProductionOutput table not found on Production sheet.'))
            }else{
                Check ($label+'.NoOwnerEntry') ([int](Probe 'PrintOwnerEntries') -eq 0)
                Check ($label+'.NoReportReadOrFallback') ([int](Probe 'PrintReportReads') -eq 0)
                $expected=if($case -ceq 'MissingSheet'){'Production sheet not found.'}else{'Session, warehouse, or captured workbook changed. Reopen Production before editing the draft.'}
                Check ($label+'.VisibleRefusal') ([string](Probe 'PrintStatus') -ceq $expected)
            }
            Check ($label+'.CapturedCustomValuesPreserved') ($sheet.Cells.Item(1,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,1).Formula -ceq '=1+2')
            Check ($label+'.DecoyPreserved') ($decoy.Worksheets.Count -eq 1 -and $otherSheet.Cells.Item(1,1).Value2 -ceq $canary -and $otherSheet.Cells.Item(2,1).Formula -ceq '=4+5')
            Check ($label+'.NoReportCreated') ($book.Worksheets.Count -eq 1)
            CaptureOwnedFormByCaptionEvidence 'Production' ('print-binding-'+$case.ToLowerInvariant()+'.png')
            [void](Probe 'CloseDesigner');$sheet.Name='Production'
        }
        Test-ProductionPrintRefusalPreservation $Fixture $book $sheet
        Test-ProductionPrintRebuild $Fixture $book $sheet
        Test-ProductionPrintInventory $Fixture $book $sheet $decoy
        $book.Close($false);$book=$null
        Check 'PrintBaseline.SavedOperatorBytesPreserved' ((Hash $path) -ceq $pin)
        Check 'PrintBaseline.OtherWarehousePreserved' (RestartPinsEqual $otherPins $Other.Root)
    }finally{
        if($null -ne $observer){Stop-Job $observer;Remove-Job $observer}
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture 'config-producer'
    }
}

function Test-ProductionPrintRefusalPreservation($Fixture,$Book,$Sheet){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Fingerprint($Worksheet){
        $tables=@(foreach($table in $Worksheet.ListObjects){
            [pscustomobject]@{Name=$table.Name;Address=$table.Range.Address();Formula=@($table.Range.Formula)}
        })
        ConvertTo-Json -Depth 12 -Compress -InputObject ([pscustomobject]@{Tables=$tables;Formula=@($Worksheet.UsedRange.Formula)})
    }
    SelectTarget $Fixture 'config-producer'
    $Sheet.Cells.Item(5,1).Value2='PROCESS';$Sheet.Cells.Item(5,2).Value2='OUTPUT'
    $Sheet.Cells.Item(5,3).Value2='RECALL CODE';$Sheet.Cells.Item(5,4).Value2='LOCAL_NOTE'
    $Sheet.Cells.Item(6,1).Value2='PRINT-DRAFT';$Sheet.Cells.Item(6,2).Value2='PRINT-OUTPUT'
    $Sheet.Cells.Item(6,4).Formula='=2+3'
    $output=$Sheet.ListObjects.Add(1,$Sheet.Range('A5:D6'),$null,1);$output.Name='ProductionOutput'
    $report=$null
    foreach($case in @('NoReport','NoRecall','MissingRecallHeader','EmptyOutput')){
        if($case -cne 'NoReport'){
            if($null -eq $report){
                foreach($candidate in $Book.Worksheets){if($candidate.Name -ceq 'RecallCodesPrint'){$report=$candidate}}
                if($null -eq $report){$report=$Book.Worksheets.Add();$report.Name='RecallCodesPrint'}
            }
            # Reset only this owned fixture so RED in one case cannot contaminate another.
            while($report.ListObjects.Count){$report.ListObjects.Item(1).Delete()};$report.Cells.Clear()
            $headers=@('LOCAL_NOTE','RECALL CODE','RECIPE','RECIPE_ID','PROCESS','OUTPUT','REAL OUTPUT','UOM','BATCH','LOCATION')
            for($i=0;$i -lt $headers.Count;$i++){$report.Cells.Item(5,$i+1).Value2=$headers[$i]}
            $report.Cells.Item(6,1).Formula='=6+7';$report.Cells.Item(6,2).Value2='RETAINED-REPORT'
            $table=$report.ListObjects.Add(1,$report.Range('A5:J6'),$null,1);$table.Name='RecallCodesReport'
            $report.Range('N5').Value2='USER_DATA';$report.Range('N6').Value2='RETAINED-TABLE'
            $other=$report.ListObjects.Add(1,$report.Range('N5:N6'),$null,1);$other.Name='PrintUserTable'
            $report.Range('Z1').Value2='RETAINED-CELL';$report.Range('Z2').Formula='=8+9'
            $before=Fingerprint $report
        }
        if($case -ceq 'MissingRecallHeader'){$output.ListColumns.Item('RECALL CODE').Name='USER_RECALL'}
        if($case -ceq 'EmptyOutput'){$output.ListRows.Item(1).Delete()}
        $source=Fingerprint $Sheet
        [void](Probe 'OpenDesigner' @($Book.Name));[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'PRINT'))
        $processes=@(Get-Process EXCEL);if($processes.Count -ne 1){throw 'Isolated Excel required.'}
        $stop=Join-Path $runRoot ('print-preserve-'+$case)
        $observer=Start-DialogCaptureAndDismiss -ExcelProcessId $processes[0].Id -TimeoutSeconds 30 -StopPath $stop
        $label='PrintPreservation.'+$case
        try{Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintAct'))}finally{
            [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
            $notice=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
            if($observer.State -ne 'Completed'){Stop-Job $observer;Remove-Job $observer;throw 'Native observer did not close normally.'}
            Remove-Job $observer
        }
        $expected=if($case -ceq 'EmptyOutput'){'ProductionOutput has no rows to print.'}else{'No recall-coded ProductionOutput rows found. Generate recall codes from checked output rows before printing.'}
        Check ($label+'.ExistingNativeRefusal') ($notice.Contains($expected))
        Check ($label+'.BothExistingReportReads') ([int](Probe 'PrintReportReads') -eq 2)
        Check ($label+'.SourcePreserved') ((Fingerprint $Sheet) -ceq $source)
        if($case -ceq 'NoReport'){
            Check ($label+'.NoEmptyReportCreated') (@($Book.Worksheets|Where-Object Name -CEQ 'RecallCodesPrint').Count -eq 0)
        }else{
            Check ($label+'.WholeReportPreserved') ((Fingerprint $report) -ceq $before)
            Check ($label+'.UserTablesPreserved') ($report.ListObjects.Count -eq 2)
            Check ($label+'.UnrelatedCellsPreserved') ($report.Range('Z1').Value2 -ceq 'RETAINED-CELL' -and $report.Range('Z2').Formula -ceq '=8+9')
        }
        CaptureOwnedFormByCaptionEvidence 'Production' ('print-preserve-'+$case.ToLowerInvariant()+'.png')
        [void](Probe 'CloseDesigner')
    }
}
