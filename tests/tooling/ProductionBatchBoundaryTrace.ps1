# Fixed-stage, unsaved instrumentation; no runtime values or business changes.
function Get-ProductionBatchTracePlan {
    $groups=@(
        @('mProduction','RunProductionBatchScaleContractTest',@(
            'BtnOpenProductionForm','If Not frmProduction.Visible Then',
            'RunProductionBatchScaleContractTest = frmProduction.TestBatchScaleContract()',
            'Exit Function','RunProductionBatchScaleContractTest = _','End Function')),
        @('mProduction','BtnOpenProductionForm',@(
            'launcherStage = "capability"','launcherStage = "capture active workbook"',
            'launcherStage = "resolve or provision Production workbook"',
            'launcherStage = "capture resolved Production workbook"',
            'launcherStage = "show production form"','ShowProductionLauncherError launcherStage,')),
        @('mProduction','ShowProductionForm',@(
            'launcherStage = "validate Production workbook"','launcherStage = "begin quiet UI"',
            'launcherStage = "repair Production surface"','launcherStage = "initialize Production UI"',
            'launcherStage = "bind Production form"','launcherStage = "initialize Production form"',
            'launcherStage = "end quiet UI"','launcherStage = "show Production form"',
            'On Error Resume Next','ShowProductionLauncherError launcherStage,')),
        @('frmProduction','TestBatchScaleContract',@(
            'minimumOk = TryParseBatchScalePercent(', 'normalOk = TryParseBatchScalePercent(',
            'maximumOk = TryParseBatchScalePercent(', 'lowRejected = Not TryParseBatchScalePercent(',
            'highRejected = Not TryParseBatchScalePercent(',
            'If minimumOk And normalOk And maximumOk And lowRejected And highRejected Then',
            'TestBatchScaleContract = _','TestBatchScaleContract = "FAIL|Batch scale contract did not hold."',
            'End Function'))
    )
    foreach($group in $groups){
        $index=0
        foreach($anchor in $group[2]){
            $index++
            [pscustomobject]@{Module=$group[0];Procedure=$group[1];Anchor=$anchor;Stage=($group[1]+'.'+$index)}
        }
    }
}
function Get-ProductionBatchTraceEdits([hashtable]$Sources) {
    foreach($step in Get-ProductionBatchTracePlan){
        if(-not $Sources.ContainsKey($step.Module)){throw 'Trace source module missing.'}
        $source=[string]$Sources[$step.Module]
        $pattern='(?ims)^\s*(?:Public|Private) (?:Sub|Function) '+[regex]::Escape($step.Procedure)+'\b.*?^End (?:Sub|Function)\b'
        $procedures=[regex]::Matches($source,$pattern)
        if($procedures.Count -ne 1){throw ('Trace procedure missing/ambiguous: '+$step.Procedure)}
        $procedure=$procedures[0]
        $offset=([regex]::Matches($source.Substring(0,$procedure.Index),"`n")).Count
        $lines=$procedure.Value -split '\r?\n'
        $found=@(for($i=0;$i -lt $lines.Count;$i++){
            if($lines[$i].Trim().StartsWith($step.Anchor,[StringComparison]::OrdinalIgnoreCase)){$offset+$i+1}
        })
        if($found.Count -ne 1){throw ('Trace anchor missing/ambiguous: '+$step.Stage)}
        [pscustomobject]@{Module=$step.Module;Line=$found[0];Stage=$step.Stage;Code=('    TestProductionBatchTrace.Mark "'+$step.Stage+'"')}
    }
}
function Get-ProductionBatchTraceLogger {
    $cases=@('Arm')+@(Get-ProductionBatchTracePlan|ForEach-Object Stage)
    $allow=@($cases|ForEach-Object {'        Case "'+$_+'"'}) -join "`r`n"
    @'
Option Explicit
Private mPath As String
Public Sub Arm(ByVal path As String)
    mPath = path
    Mark "Arm"
End Sub
Public Sub Mark(ByVal stage As String)
    Dim channel As Integer
    Select Case stage
'@ + "`r`n"+$allow+"`r`n"+@'
        Case Else: Exit Sub
    End Select
    If mPath = "" Then Exit Sub
    On Error GoTo Failed
    channel = FreeFile
    Open mPath For Append As #channel
    Print #channel, stage
    Close #channel
    Exit Sub
Failed:
    On Error Resume Next
    If channel > 0 Then Close #channel
End Sub
'@
}
function Install-ProductionBatchTrace {
    param($Excel,[hashtable]$Packages,[string]$PackageRoot)
    $project=$Packages['invSys.Operations.xlam'].VBProject
    $sources=@{}
    foreach($name in @('mProduction','frmProduction')){
        $module=$project.VBComponents.Item($name).CodeModule
        $sources[$name]=[string]$module.Lines(1,$module.CountOfLines)
    }
    # Resolve every anchor before the first mutation. VBA normalizes identifier case.
    $edits=@(Get-ProductionBatchTraceEdits $sources)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestProductionBatchTrace'
    $probe.CodeModule.AddFromString((Get-ProductionBatchTraceLogger))
    foreach($group in $edits|Group-Object Module){
        $module=$project.VBComponents.Item($group.Name).CodeModule
        foreach($edit in $group.Group|Sort-Object Line -Descending){$module.InsertLines($edit.Line,$edit.Code)}
    }
    $visible=$Excel.VBE.MainWindow.Visible
    try {
        foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')){
            $target=$Packages[$name].VBProject
            foreach($ref in $target.References){
                if($ref.IsBroken){throw 'Trace dependency is broken; not product RED.'}
                if($ref.Name -like 'invSys_*' -and -not [string]::Equals((Split-Path -Parent $ref.FullPath),$PackageRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Trace dependency outside pinned directory.'}
            }
            $Excel.VBE.ActiveVBProject=$target
            foreach($component in $target.VBComponents){if($component.Type -eq 1){$component.CodeModule.CodePane.Show();break}}
            $command=$Excel.VBE.CommandBars.FindControl(1,578)
            if($null -eq $command){throw 'Trace compile command unavailable.'}
            if($command.Enabled){$command.Execute()}
            if($Excel.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Trace compile incomplete; not product RED.'}
            Write-Output ('BATCH_TRACE_COMPILE_PASS '+$name)
        }
    } finally {$Excel.VBE.MainWindow.Visible=$visible}
}
