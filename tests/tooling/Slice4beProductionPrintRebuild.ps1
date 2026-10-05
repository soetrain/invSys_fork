# Real Print Recall handler/report builder; native preview is an explicitly observed seam.
function Install-ProductionPrintRebuildProbe($Module){
    $start=$Module.ProcStartLine('BtnPrintRecallCodes',0);$end=$start+$Module.ProcCountLines('BtnPrintRecallCodes',0)
    $lines=@(for($i=$start;$i -lt $end;$i++){if($Module.Lines($i,1).Trim() -ceq 'wsReport.PrintOut Preview:=True'){$i}})
    if($lines.Count -ne 1){throw 'Native preview boundary changed; not product RED.'}
    $Module.ReplaceLine($lines[0],'    TestProductionDesigner.PrintPreviewForTest wsReport')
    $adapter=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mPrintPreviews As Long, mPrintPreviewRows As Long, mPrintPreviewMetadata As Boolean')
    $adapter.AddFromString(@'
Public Sub PrintPreviewForTest(ByVal report As Worksheet)
    mPrintPreviews = mPrintPreviews + 1
    mPrintPreviewRows = report.ListObjects("RecallCodesReport").ListRows.Count
    mPrintPreviewMetadata = (Len(CStr(report.Range("D2").Text)) = 19 And InStr(CStr(report.Range("D2").Text), "#") = 0)
End Sub
Public Sub ResetPrintPreviewForTest()
    mPrintPreviews = 0: mPrintPreviewRows = 0: mPrintPreviewMetadata = False
End Sub
Public Function PrintPreviewCountForTest() As Long
    PrintPreviewCountForTest = mPrintPreviews
End Function
Public Function PrintPreviewRowsForTest() As Long
    PrintPreviewRowsForTest = mPrintPreviewRows
End Function
Public Function PrintPreviewMetadataForTest() As Boolean
    PrintPreviewMetadataForTest = mPrintPreviewMetadata
End Function
'@)
}

function Test-ProductionPrintRebuild($Fixture,$Book,$Sheet){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Fingerprint($Worksheet){
        ConvertTo-Json -Depth 12 -Compress -InputObject ([pscustomobject]@{Cells=@($Worksheet.UsedRange.Formula);Tables=@(foreach($t in $Worksheet.ListObjects){$t.Name+'|'+$t.Range.Address()})})
    }
    SelectTarget $Fixture 'config-producer'
    $output=$Sheet.ListObjects.Item('ProductionOutput')
    $output.ListColumns.Item(3).Name='RECALL CODE'
    $report=$Book.Worksheets.Item('RecallCodesPrint')
    $managed=@('RECIPE','RECIPE_ID','PROCESS','OUTPUT','REAL OUTPUT','UOM','BATCH','RECALL CODE','LOCATION')
    $headers=@('LOCAL_NOTE','RECALL CODE','LOCATION','BATCH','UOM','REAL OUTPUT','OUTPUT','PROCESS','RECIPE_ID','RECIPE','LOCAL_FORMULA')
    foreach($case in @('Reuse','Grow','UnknownGrowth','Shrink','MissingHeader','AmbiguousHeader','OccupiedGrowth','NewTable')){
        # Reset only owned fixture state; each case independently demonstrates its contract.
        while($report.ListObjects.Count){$report.ListObjects.Item(1).Delete()};$report.Cells.Clear()
        $oldRows=if($case -ceq 'Shrink'){3}else{1}
        $newRows=if($case -in @('Grow','UnknownGrowth','OccupiedGrowth')){3}else{1}
        $customRows=if($case -ceq 'UnknownGrowth'){3}else{$oldRows}
        while($output.ListRows.Count){$output.ListRows.Item(1).Delete()}
        for($r=1;$r -le $newRows;$r++){
            $row=$output.ListRows.Add()
            $row.Range.Cells.Item(1,1).Value2='PRINT-PROCESS-'+$r
            $row.Range.Cells.Item(1,2).Value2='PRINT-OUTPUT-'+$r
            $row.Range.Cells.Item(1,3).Value2='PRINT-RECALL-'+$r
            $row.Range.Cells.Item(1,4).Formula='=2+3'
        }
        $source=Fingerprint $Sheet
        $oldTable=$null
        if($case -cne 'NewTable'){
            for($i=0;$i -lt $headers.Count;$i++){$report.Cells.Item(5,$i+1).Value2=$headers[$i]}
            for($r=1;$r -le $oldRows;$r++){
                $report.Cells.Item(5+$r,1).Value2='LOCAL-'+$r
                $report.Cells.Item(5+$r,11).Formula='='+$r+'+7'
                $report.Cells.Item(5+$r,2).Value2='OLD-RECALL-'+$r
            }
            if($case -ceq 'UnknownGrowth'){
                for($r=2;$r -le $customRows;$r++){$report.Cells.Item(5+$r,1).Value2='LOCAL-'+$r;$report.Cells.Item(5+$r,11).Formula='='+$r+'+7'}
            }
            if($case -ceq 'OccupiedGrowth'){$report.Range('B7').Value2='BLOCKED-USER-DATA'}
            $oldTable=$report.ListObjects.Add(1,$report.Range('A5').Resize($oldRows+1,11),$null,1);$oldTable.Name='RecallCodesReport'
            if($oldTable.ListRows.Count -ne $oldRows){throw 'Report row-count fixture changed; not product RED.'}
            if($case -ceq 'MissingHeader'){$oldTable.ListColumns.Item('PROCESS').Name='USER_PROCESS'}
            if($case -ceq 'AmbiguousHeader'){$oldTable.ListColumns.Item('LOCAL_FORMULA').Name=' PROCESS '}
        }
        $report.Range('N5').Value2='USER_DATA';$report.Range('N6').Value2='RETAINED-TABLE'
        if($case -cne 'NewTable'){$report.Range('N7').Formula='=COUNTA(RecallCodesReport[RECALL CODE])'}
        $other=$report.ListObjects.Add(1,$report.Range('N5:N7'),$null,1);$other.Name='PrintUserTable'
        $report.Range('Z1').Value2='RETAINED-CELL';$report.Range('Z2').Formula='=8+9'
        $before=Fingerprint $report
        [void](Probe 'OpenDesigner' @($Book.Name));[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'PRINT'))
        [void](Probe 'ResetPrintPreviewForTest')
        $stop=Join-Path $runRoot ('print-rebuild-'+$case)
        $observer=Start-DialogCaptureAndDismiss -ExcelProcessId (@(Get-Process EXCEL)[0].Id) -TimeoutSeconds 30 -StopPath $stop
        $label='PrintRebuild.'+$case
        try{Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintAct'))}finally{
            [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
            $notice=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
            if($observer.State -ne 'Completed'){Stop-Job $observer;Remove-Job $observer;throw 'Native observer did not close normally.'}
            Remove-Job $observer
        }
        $refusal=$case -in @('MissingHeader','AmbiguousHeader','OccupiedGrowth')
        Check ($label+'.SourcePreserved') ((Fingerprint $Sheet) -ceq $source)
        if($refusal){
            Check ($label+'.NoPreviewBoundary') ([int](Probe 'PrintPreviewCountForTest') -eq 0)
            Check ($label+'.ReportPreserved') ((Fingerprint $report) -ceq $before)
            Check ($label+'.VisibleRefusal') ($notice.Contains('Recall report cannot be refreshed safely. Check its managed headers and available space.'))
        }else{
            Check ($label+'.PreviewBoundaryOnce') ([int](Probe 'PrintPreviewCountForTest') -eq 1)
            Check ($label+'.PreviewRows') ([int](Probe 'PrintPreviewRowsForTest') -eq $newRows)
            Check ($label+'.PreviewMetadataReadable') ([bool](Probe 'PrintPreviewMetadataForTest'))
            $current=@($report.ListObjects|Where-Object Name -CEQ 'RecallCodesReport')
            $correct=$current.Count -eq 1
            if($correct){
                $correct=$current[0].ListRows.Count -eq $newRows
                for($r=1;$r -le $newRows;$r++){
                    foreach($field in @('PROCESS','OUTPUT','RECALL CODE')){
                        $prefix=if($field -ceq 'RECALL CODE'){'RECALL'}else{$field}
                        $correct=$correct -and $current[0].ListColumns.Item($field).DataBodyRange.Cells.Item($r,1).Value2 -ceq ('PRINT-'+$prefix+'-'+$r)
                    }
                }
            }
            Check ($label+'.ManagedValuesByHeader') $correct
            if($case -cne 'NewTable'){
                $custom=$true
                for($r=1;$r -le $customRows;$r++){$custom=$custom -and $report.Cells.Item(5+$r,1).Value2 -ceq ('LOCAL-'+$r) -and $report.Cells.Item(5+$r,11).Formula -ceq ('='+$r+'+7')}
                Check ($label+'.UnknownCellsAndPositionsPreserved') $custom
                Check ($label+'.OriginalHeadersAndTableReused') ($current.Count -eq 1 -and (@($current[0].HeaderRowRange.Value2) -join '|') -ceq ($headers -join '|') -and $current[0].Range.Row -eq 5)
                if($case -ceq 'Shrink'){Check ($label+'.OldManagedTailCleared') ($report.Range('B7').Value2 -eq $null -and $report.Range('B8').Value2 -eq $null)}
            }
        }
        Check ($label+'.OtherTablePreserved') (@($report.ListObjects|Where-Object Name -CEQ 'PrintUserTable').Count -eq 1 -and $report.Range('N6').Value2 -ceq 'RETAINED-TABLE')
        if($case -cne 'NewTable'){Check ($label+'.StructuredReferencePreserved') ($report.Range('N7').Formula -ceq '=COUNTA(RecallCodesReport[RECALL CODE])')}
        Check ($label+'.UnrelatedCellsPreserved') ($report.Range('Z1').Value2 -ceq 'RETAINED-CELL' -and $report.Range('Z2').Formula -ceq '=8+9')
        CaptureOwnedFormByCaptionEvidence 'Production' ('print-rebuild-'+$case.ToLowerInvariant()+'.png')
        [void](Probe 'CloseDesigner')
        if($Phase -ceq 'GREEN' -and $case -in @('Reuse','Grow','UnknownGrowth','Shrink')){
            $visible=$excel.Visible;$updating=$excel.ScreenUpdating;$windowState=$excel.WindowState
            try{
                $excel.Visible=$true;$excel.ScreenUpdating=$true;$excel.WindowState=-4137
                $report.Activate();$excel.ActiveWindow.ScrollRow=1;$excel.ActiveWindow.ScrollColumn=1
                CaptureFormEvidence '' ('print-report-'+$case.ToLowerInvariant()+'.png') ([long]$excel.Hwnd)
            }finally{$excel.WindowState=$windowState;$excel.ScreenUpdating=$updating;$excel.Visible=$visible}
        }
    }
}
