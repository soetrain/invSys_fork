# Disposable package instrumentation only. Observe the existing completion
# bracket, never replace the handler, service or quiet-state implementation.
function Install-ProductionQuietBoundaryProbe {
    param($Excel,[hashtable]$Packages,[string]$PackageRoot)
    $core=$Packages['invSys.Core.xlam']
    $operations=$Packages['invSys.Operations.xlam']
    $probe=$core.VBProject.VBComponents.Add(1)
    $probe.Name='TestProductionQuietProbe'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mExpected As String
Private mArmed As Boolean
Private mCalls As Long
Private mNamesMatch As Boolean
Private mCompleted As Long
Private mPrimitive As Long
Private mBound As Long
Private mRestored As Long
Private mScreen As Boolean, mEvents As Boolean, mAlerts As Boolean
Private mStatus As Boolean, mCalc As Long, mQuiet As Boolean

Public Sub Configure(ByVal expectedName As String)
    mExpected = expectedName
    mArmed = False
    mCompleted = 0: mPrimitive = 0: mBound = 0: mRestored = 0
End Sub

Public Sub Arm(ByVal capturedName As String)
    If mExpected = vbNullString Then Exit Sub
    mArmed = True
    mCalls = 0
    mNamesMatch = (StrComp(capturedName, mExpected, vbBinaryCompare) = 0)
    mScreen = Application.ScreenUpdating
    mEvents = Application.EnableEvents
    mAlerts = Application.DisplayAlerts
    mStatus = Application.DisplayStatusBar
    mCalc = Application.Calculation
    mQuiet = modUiQuiet.QuietUiIsActive()
End Sub

Public Sub Observe(ByVal workbookName As String)
    If Not mArmed Then Exit Sub
    mCalls = mCalls + 1
    mNamesMatch = mNamesMatch And (StrComp(workbookName, mExpected, vbBinaryCompare) = 0)
End Sub

Public Sub Finish()
    If Not mArmed Then Exit Sub
    mArmed = False
    mCompleted = mCompleted + 1
    If mCalls = 1 Then mPrimitive = mPrimitive + 1
    If mCalls = 1 And mNamesMatch Then mBound = mBound + 1
    If Application.ScreenUpdating = mScreen And Application.EnableEvents = mEvents And _
       Application.DisplayAlerts = mAlerts And Application.DisplayStatusBar = mStatus And _
       Application.Calculation = mCalc And modUiQuiet.QuietUiIsActive() = mQuiet Then
        mRestored = mRestored + 1
    End If
End Sub

Public Function ResultCount(ByVal fact As String) As Long
    Select Case fact
        Case "Completions": ResultCount = mCompleted
        Case "PrimitiveEntries": ResultCount = mPrimitive
        Case "CapturedBinding": ResultCount = mBound
        Case "QuietStateRestored": ResultCount = mRestored
        Case Else: Err.Raise 5
    End Select
End Function
'@)
    $bridge=$core.VBProject.VBComponents.Item('modOperationsPrimitiveBridge').CodeModule
    $start=$bridge.ProcBodyLine('BeginQuietUiForWorkbook',0)
    $bridge.InsertLines($start+1,'    TestProductionQuietProbe.Observe workbookName')
    $production=$operations.VBProject.VBComponents.Item('mProduction').CodeModule
    $start=$production.ProcStartLine('CompleteProductionRunAfterCheckInForOutput',0)
    $count=$production.ProcCountLines('CompleteProductionRunAfterCheckInForOutput',0)
    $text=$production.Lines($start,$count)
    $before='    Dim completionResult As cProductionCompletionResult'
    $after='    quietStarted = False'
    if(([regex]::Matches($text,[regex]::Escape($before))).Count -ne 1 -or
       ([regex]::Matches($text,[regex]::Escape($after))).Count -ne 1){throw 'Completion probe anchors ambiguous; not behavioral RED.'}
    $text=$text.Replace($before,$before+"`r`n    TestProductionQuietProbe.Arm wsProd.Parent.Name")
    $text=$text.Replace($after,$after+"`r`n    TestProductionQuietProbe.Finish")
    $production.DeleteLines($start,$count)
    $production.InsertLines($start,$text)
    # Explicitly compile every instrumented dependency before executing a form.
    $visible=$Excel.VBE.MainWindow.Visible
    try {
        foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')){
            $project=$Packages[$name].VBProject
            foreach($ref in $project.References){
                if($ref.IsBroken){throw 'Broken probe reference; not behavioral RED.'}
                if($ref.Name -like 'invSys_*' -and -not [string]::Equals((Split-Path -Parent $ref.FullPath),$PackageRoot,[StringComparison]::OrdinalIgnoreCase)){
                    throw 'Probe dependency outside disposable package directory.'
                }
            }
            $Excel.VBE.ActiveVBProject=$project
            foreach($component in $project.VBComponents){if($component.Type -eq 1){$component.CodeModule.CodePane.Show();break}}
            $command=$Excel.VBE.CommandBars.FindControl(1,578)
            if($null -eq $command){throw 'Compile command unavailable.'}
            if($command.Enabled){$command.Execute()}
            if($Excel.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Probe compile incomplete; not behavioral RED.'}
            Write-Output ('PRODUCTION_QUIET_PROBE_COMPILE_PASS '+$name)
        }
    } finally {$Excel.VBE.MainWindow.Visible=$visible}
}
