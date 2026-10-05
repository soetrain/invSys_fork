# Actual Print Recall handler; missing output naturally refuses before print preview.
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
