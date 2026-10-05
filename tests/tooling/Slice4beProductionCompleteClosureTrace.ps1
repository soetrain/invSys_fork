# Disposable fixed-label diagnostics; owner results are never replaced.
# A second cleanup failure escapes the loop into the actual-handler assertion.
function Install-ProductionCompleteClosureTrace {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,"Private mCompleteClosureTraceCount As Long`r`nPrivate mCompleteClosureRecoveryCount As Long")
    $path=(Join-Path $reportRoot 'complete-owner-closure-trace.log').Replace('"','""')
    $adapter.AddFromString(@'
Public Sub CompleteOwnerClosureTrace(ByVal stage As String)
    Dim channel As Integer
    If stage = "Before.ExceptionRecovery" Then
        mCompleteClosureRecoveryCount = mCompleteClosureRecoveryCount + 1
        If mCompleteClosureRecoveryCount > 1 Then Err.Raise vbObjectError + 2811, "CompleteClosureProbe", "Repeated completion cleanup failure."
    End If
    If mCompleteClosureTraceCount >= 128 Then Exit Sub
    mCompleteClosureTraceCount = mCompleteClosureTraceCount + 1
    channel = FreeFile
    Open "TRACE_PATH" For Append As #channel
    Print #channel, stage
    Close #channel
End Sub
'@.Replace('TRACE_PATH',$path))
    function TraceLines($Module,[string]$Procedure,$Anchors) {
        $start=$Module.ProcStartLine($Procedure,0);$end=$start+$Module.ProcCountLines($Procedure,0)
        $edits=@();$found=@{}
        for($line=$start;$line -lt $end;$line++){
            $text=$Module.Lines($line,1).Trim()
            foreach($anchor in $Anchors.Keys){
                if($text -ceq $anchor){
                    $label=$Anchors[$anchor];$found[$label]=$true
                    if($text.EndsWith(':')){
                        $edits+=@{Line=$line+1;Code='    TestProductionDesigner.CompleteOwnerClosureTrace "'+$label+'"'}
                    }else{
                        $edits+=@{Line=$line;Code='    TestProductionDesigner.CompleteOwnerClosureTrace "Before.'+$label+'"'}
                        $edits+=@{Line=$line+1;Code='    TestProductionDesigner.CompleteOwnerClosureTrace "After.'+$label+'"'}
                    }
                }
            }
        }
        if($edits.Count -eq 0){throw 'Closure trace anchors unavailable; not product RED.'}
        foreach($edit in @($edits|Sort-Object {[int]$_.Line} -Descending)){$Module.InsertLines($edit.Line,$edit.Code)}
        return $found
    }
    $wrapper=$project.VBComponents.Item('modProductionCompleteActions').CodeModule
    $hits=TraceLines $wrapper 'Execute' @{
        'owner.CompleteProductionRun'='OwnerCall'
        'owner.CompleteProductionRun action'='OwnerCall'
        'report = owner.RunAllocationStatus()'='StatusRead'
        'Set owner.RunActionContinuation = action'='AttachContinuation'
        'Set owner.RunActionContinuation = Nothing'='DetachContinuation'
        'If Not action Is Nothing Then action.Finish report'='FinishObservation'
        'If report <> "" Then owner.ShowStatus report'='FinalStatus'
        'Done:'='Cleanup'
        'Resume Done'='ExceptionRecovery'
    }
    if(-not $hits.ContainsKey('OwnerCall') -or -not $hits.ContainsKey('Cleanup')){throw 'Actual completion wrapper trace incomplete.'}
    [void](TraceLines $adapter 'CheckClosedYieldClose' @{'mRunLocalCapturedBookForTest.Close False'='NativeWorkbookClose'})
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    [void](TraceLines $form 'CompleteProductionRun' @{
        'TestProductionDesigner.CheckYieldReturned "CompletePending", True'='PendingInterruption'
        'If Not action.CanContinue(reusableReport) Then ShowStatus reusableReport: Exit Sub'='OwnerContinuation'
    })
}
