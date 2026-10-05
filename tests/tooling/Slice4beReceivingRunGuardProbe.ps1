# Disposable probes attempt re-entry while the actual Start/Next handler owns the run.
# They call public boundaries, never fabricate owner outcomes or bypass the guard.
function Install-ReceivingRunGuardProbe {
    $component=$null
    foreach($item in $packages['invSys.Core.xlam'].VBProject.VBComponents){if($item.Name -ceq 'modExecutionRun'){$component=$item;break}}
    if($null -eq $component){return}
    $module=$component.CodeModule
    $module.InsertLines($module.CountOfDeclarationLines+1,@'
Private mProbeRunEntry As Boolean
Private mProbeStartEntries As Long
Private mProbeStepEntries As Long
Private mProbeEntriesRefused As Boolean
'@)
    foreach($procedure in @('Start','BeginStep')) {
        $start=$module.ProcStartLine($procedure,0);$count=$module.ProcCountLines($procedure,0)
        $body=$module.Lines($start,$count)
        if($procedure -ceq 'Start'){
            $body=$body.Replace('    mStarting = True',('    mStarting = True'+"`r`n"+'    ProbeRunEntryForTest context, token, entity, mode, True'))
        }else{
            $body=$body.Replace('    BeginStep = True: Exit Function',('    ProbeRunEntryForTest context, token, mEntity, CStr(mRun("Mode")), False'+"`r`n"+'    BeginStep = True: Exit Function'))
        }
        $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    }
    $module.AddFromString(@'
Public Sub ArmRunEntryProbeForTest()
    mProbeStartEntries = 0: mProbeStepEntries = 0: mProbeEntriesRefused = True: mProbeRunEntry = True
End Sub
Public Function RunEntryProbeResultForTest() As String
    RunEntryProbeResultForTest = CStr(mProbeStartEntries) & "|" & CStr(mProbeStepEntries) & "|" & CStr(mProbeEntriesRefused)
End Function
Private Sub ProbeRunEntryForTest(ByVal context As String, ByVal token As String, ByVal entity As String, ByVal mode As String, ByVal starting As Boolean)
    Dim accepted As Boolean, notice As String, nestedToken As String, snapshot As String, guide As String, profile As String, target As String, control As String, inputs As String
    If Not mProbeRunEntry Then Exit Sub
    mProbeRunEntry = False
    If starting Then mProbeStartEntries = mProbeStartEntries + 1 Else mProbeStepEntries = mProbeStepEntries + 1
    accepted = Start(context, token, entity, mode, notice)
    mProbeEntriesRefused = mProbeEntriesRefused And Not accepted
    accepted = OpenSetup(context, mKey, nestedToken, snapshot, guide, profile, target, notice)
    mProbeEntriesRefused = mProbeEntriesRefused And Not accepted
    accepted = BeginStep(context, token, control, inputs, notice)
    mProbeEntriesRefused = mProbeEntriesRefused And Not accepted
    mProbeRunEntry = starting
End Sub
'@)
}
