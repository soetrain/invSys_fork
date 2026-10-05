# Static callback names only. Disposable instrumentation preserves error handling.
function Install-ReceivingRunEventTrace {
    $targets=@{
        'invSys.Operations.xlam'=@('frmActionPathRun','frmActionPathLibrary','frmActionPaths','frmInventoryViewer','frmReceiving','cRecordingControls')
        'invSys.Admin.xlam'=@('frmAdminSettings','cAdminTrackingPolicy')
    }
    foreach($package in $targets.Keys){
        $project=$packages[$package].VBProject
        $logger=$project.VBComponents.Add(1);$logger.Name='TestRunGuardTrace'
        $logger.CodeModule.AddFromString(@'
Option Explicit
Private mPath As String
Public Sub InitializeForTest(ByVal path As String)
    mPath = path
    Emit "Trace.initialized"
End Sub
Public Sub Emit(ByVal marker As String)
    Dim channel As Integer
    If mPath = "" Then Exit Sub
    On Error Resume Next
    channel = FreeFile
    Open mPath For Append As #channel
    Print #channel, CStr(Timer) & "|" & marker
    Close #channel
End Sub
'@)
        foreach($name in $targets[$package]){
            $module=$project.VBComponents.Item($name).CodeModule
            foreach($procedure in @('UserForm_Activate','UserForm_QueryClose','UserForm_Terminate','mClose_Click','mSave_Click','mCapture_Click','Render','Execute','Disconnect','ReleaseRunner','CloseRunner','ClearContent','CloseGuideActionPicker','CloseActionPathView','ClosePublishedGuides','ClearEvaluation','ReleaseReader','ClearSelection','CloseExecution')){
                $start=0;$count=0
                try{$start=$module.ProcStartLine($procedure,0);$count=$module.ProcCountLines($procedure,0)}catch{continue}
                $body=$module.Lines($start,$count)
                $rows=$body -split '\r?\n';$signature=-1
                for($i=0;$i -lt $rows.Count;$i++){if($rows[$i] -match '^\s*(Public |Private )?(Sub|Function) '){$signature=$i;break}}
                if($signature -lt 0){throw 'Callback trace signature unavailable.'}
                while($rows[$signature].TrimEnd().EndsWith('_')){$signature++}
                $kind=if($body -match '\bEnd Function\b'){'Function'}else{'Sub'}
                $prefix=$name+'.'+$procedure
                $body=(@($rows[0..$signature])+@('    TestRunGuardTrace.Emit "'+$prefix+'.enter"')+@($rows[($signature+1)..($rows.Count-1)])) -join "`r`n"
                if($procedure -ceq 'UserForm_QueryClose'){
                    $body=$body.Replace(('TestRunGuardTrace.Emit "'+$prefix+'.enter"'),('TestRunGuardTrace.Emit "'+$prefix+'.mode=" & CStr(CloseMode)'))
                }
                $body=$body.Replace(('Exit '+$kind),('TestRunGuardTrace.Emit "'+$prefix+'.exit": Exit '+$kind))
                $body=$body.Replace(('End '+$kind),('    TestRunGuardTrace.Emit "'+$prefix+'.return"'+"`r`n"+'End '+$kind))
                $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
            }
        }
    }
}
