# Request closure while the ordinary owner has returned but Core still owns the
# pending step. The hook changes neither dispatch results nor recording evidence.
function Install-ReceivingRunCloseProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmActionPathRun').CodeModule
    $start=$form.ProcStartLine('DispatchSteps',0);$count=$form.ProcCountLines('DispatchSteps',0)
    $body=$form.Lines($start,$count)
    $anchor='        completed = modExecutionRun.FinishStep(mContext, mToken, delivered, notice)'
    if(-not $body.Contains($anchor)){throw 'Run close probe requires the actual owner-return boundary.'}
    $body=$body.Replace($anchor,('        TestRunClose.DeliverForTest Me'+"`r`n"+$anchor))
    $form.DeleteLines($start,$count);$form.InsertLines($start,$body)
    $form.AddFromString(@'
Public Function RequestBusyCloseForTest(ByVal kind As String) As Boolean
    Dim cancelled As Integer
    If Not mExecuting Then Err.Raise 5, , "Close probe requires an active dispatch frame."
    If kind = "Button" Then
        Me.Controls("btnCloseRun").Value = True
        cancelled = 1
    ElseIf kind = "QueryClose" Then
        UserForm_QueryClose cancelled, 5
    Else
        Err.Raise 5, , "Unknown close probe."
    End If
    RequestBusyCloseForTest = (cancelled <> 0 And mCloseRequested And Not mClosing And modOperationsFormLifetime.IsLoaded(Me))
End Function
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestRunClose'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mKind As String
Private mCalls As Long
Private mDeferred As Boolean
Public Sub ArmForTest(ByVal kind As String)
    mKind = kind: mCalls = 0: mDeferred = False
End Sub
Public Sub DeliverForTest(ByVal form As frmActionPathRun)
    Dim kind As String
    If mKind = "" Then Exit Sub
    kind = mKind: mKind = "": mCalls = mCalls + 1
    mDeferred = form.RequestBusyCloseForTest(kind)
End Sub
Public Function ResultForTest() As String
    ResultForTest = CStr(mCalls) & "|" & CStr(mDeferred)
End Function
'@)
}

function Test-ReceivingRunBusyClose {
    foreach($case in @('Button','QueryClose')) {
        $before=BoundPins $runRoot;$business=BusinessPins;$activity=ActivityPins
        if(-not (OpenRunSetup)){throw 'Busy-close setup fixture is unavailable.'}
        if((RunnerControl 'lstRunSourceEntities' 'Select' '0') -cne 'SELECTED'){throw 'Busy-close source entity unavailable.'}
        [void](RunnerControl 'cboRunMode' 'Write' 'Run all')
        [void](Run 'invSys.Operations.xlam' 'TestRunClose.ArmForTest' @($case))
        Write-Host ('BusyClose.'+$case+'.Start.Before')
        [void](RunnerControl 'btnStartRun' 'Click')
        Write-Host ('BusyClose.'+$case+'.Start.Returned')
        $rows=@(RunFiles|Where-Object {-not $before.ContainsKey($_.FullName)}|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision)
        if(-not $rows.Count){throw 'Busy-close attempt was not created.'}
        $latest=$rows[-1];$steps=@($latest.Steps)
        $stopped=$latest.State -ceq 'Stopped' -and $steps.Count -eq 1 -and $steps[0].State -ceq 'Completed' -and $steps[0].ControlId -ceq 'RECEIVING_OPEN'
        Check ('ReceivingRun.BusyClose.'+$case+'.DefersWhileOwnerReturns') ((Run 'invSys.Operations.xlam' 'TestRunClose.ResultForTest') -ceq '1|True')
        Check ('ReceivingRun.BusyClose.'+$case+'.StopsAfterOneRealOwner') ($stopped -and $steps[0].ActivityId -cne '' -and (BoundSame $business (BusinessPins)))
        $fresh=@(Get-ChildItem -LiteralPath $activityRoot -File -Filter '*.json'|Where-Object {-not $activity.ContainsKey($_.Name)}|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json})
        $actual=@($fresh|Where-Object {$_.SequenceId -ceq $latest.Recording.SequenceId})
        Check ('ReceivingRun.BusyClose.'+$case+'.FreshEvidenceOnlyForCurrentOwner') ($actual.Count -eq 2 -and @($actual|Where-Object ControlId -CNE 'RECEIVING_OPEN').Count -eq 0 -and @($actual|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($actual|Where-Object OutcomeCode -CIN @('OPENED','REUSED')).Count -eq 1)
        Check ('ReceivingRun.BusyClose.'+$case+'.UnloadsAfterOuterHandler') ((RunnerControl '' 'Count') -ceq '0' -and (RunnerControl 'btnNextRunStep' 'Click') -ceq 'MISSING')
    }
}
