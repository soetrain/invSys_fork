# Fixed owner feedback through the real form handler; native preview remains a declared seam.
function Install-ProductionPrintOutcomeProbe {
    $adapter=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mPrintPreviewFailure As Boolean')
    $start=$adapter.ProcStartLine('PrintPreviewForTest',0);$end=$start+$adapter.ProcCountLines('PrintPreviewForTest',0)
    $lines=@(for($i=$start;$i -lt $end;$i++){if($adapter.Lines($i,1).Trim() -ceq 'mPrintPreviews = mPrintPreviews + 1'){$i}})
    if($lines.Count -ne 1){throw 'Preview observer anchor changed; not product RED.'}
    $adapter.InsertLines($lines[0]+1,'    If mPrintPreviewFailure Then Err.Raise 1004, "Print preview", "Print preview unavailable (fixture)."')
    $adapter.AddFromString(@'
Public Sub PrintPreviewFailureForTest(ByVal fail As Boolean)
    mPrintPreviewFailure = fail
End Sub
'@)
}

function Test-ProductionPrintOutcome($Fixture,$Book,$Sheet){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Fingerprint($Worksheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Worksheet.UsedRange.Formula)}
    SelectTarget $Fixture 'config-producer'
    $output=$Sheet.ListObjects.Item('ProductionOutput')
    $recall=$output.ListColumns.Item('RECALL CODE').DataBodyRange.Cells.Item(1,1)
    $report=$Book.Worksheets.Item('RecallCodesPrint')
    [void](Probe 'OpenDesigner' @($Book.Name));[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'PRINT'))
    foreach($case in @('PreviewReturn','PreviewFailure','RefusalAfterPreview','Recovery')){
        $recall.Value2=if($case -ceq 'RefusalAfterPreview'){''}else{'PRINT-RECALL-1'}
        $source=Fingerprint $Sheet;$reportBefore=Fingerprint $report
        [void](Probe 'ResetPrintPreviewForTest');[void](Probe 'PrintPreviewFailureForTest' @($case -ceq 'PreviewFailure'))
        $stop=Join-Path $runRoot ('print-outcome-'+$case)
        $observer=Start-DialogCaptureAndDismiss -ExcelProcessId (@(Get-Process EXCEL)[0].Id) -TimeoutSeconds 30 -StopPath $stop
        $label='PrintOutcome.'+$case
        try{Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintAct'))}finally{
            [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
            $notice=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
            if($observer.State -ne 'Completed'){Stop-Job $observer;Remove-Job $observer;throw 'Native observer did not close normally.'}
            Remove-Job $observer
        }
        $expected=switch($case){
            'PreviewFailure'{'BTN_PRINT_CODES failed: Print preview unavailable (fixture).'}
            'RefusalAfterPreview'{'No recall-coded ProductionOutput rows found. Generate recall codes from checked output rows before printing.'}
            default{'Print preview closed.'}
        }
        Check ($label+'.SingleOwnerEntry') ([int](Probe 'PrintOwnerEntries') -eq 1)
        Check ($label+'.SingleReportBuild') ([int](Probe 'PrintReportReads') -eq 1)
        Check ($label+'.TruthfulStatus') ([string](Probe 'PrintStatus') -ceq $expected)
        $previews=if($case -ceq 'RefusalAfterPreview'){0}else{1}
        Check ($label+'.PreviewBoundaryCount') ([int](Probe 'PrintPreviewCountForTest') -eq $previews)
        if($case -in @('PreviewFailure','RefusalAfterPreview')){Check ($label+'.ExistingNativeNotice') ($notice.Contains($expected))}
        Check ($label+'.SourceAndExactKeyPreserved') ((Fingerprint $Sheet) -ceq $source)
        if($case -ceq 'RefusalAfterPreview'){Check ($label+'.PriorReportPreserved') ((Fingerprint $report) -ceq $reportBefore)}
        CaptureOwnedFormByCaptionEvidence 'Production' ('print-outcome-'+$case.ToLowerInvariant()+'.png')
    }
    [void](Probe 'PrintPreviewFailureForTest' @($false));[void](Probe 'CloseDesigner')
}
