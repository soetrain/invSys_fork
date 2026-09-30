# Diagnostic only: unsaved instrumentation, fixed labels, no business values.
function Get-ProductionRestartTracePlan {
    $groups=@(
        @('mProduction','RunReusableProductionRestartActionContractTest',@(
            @('BtnOpenProductionForm','Launcher.Before',1,1),
            @('If Not frmProduction.Visible Then','Launcher.After',1,1),
            @('RunReusableProductionRestartActionContractTest = _','Contract.Before',1,2))),
        @('frmProduction','TestReusableProductionRestartActionContract',@(
            @('If Not mBuilt Then BuildLayout','Form.Enter',1,1),
            @('RefreshReusableDesignLists','Refresh.Before',1,1),
            @('rowIndex = FindIdentityListRow','Refresh.After',1,1),
            @('mBtnLoaderLoad_Click','Load.Before',1,1),
            @('If Not modProductionReusableRun.ReusableRunIsLoaded() Then','Load.After',1,1),
            @('multipleTablesRediscovered = _','Tables.Discover',1,1),
            @('mBtnProcessWorksheetRetrieve_Click','Retrieve.First.Before',1,2),
            @('selectedOnly = _','Retrieve.First.After',1,1),
            @('mBtnProcessWorksheetRetrieve_Click','Retrieve.Second.Before',2,2),
            @('allRetrieved = _','Retrieve.Second.After',1,1),
            @('If Not multipleTablesRediscovered Or Not worksheetRediscovered Or _','Tables.Check',1,1),
            @('TestReusableProductionRestartActionContract = _','Result.RecipeMissing',1,5),
            @('TestReusableProductionRestartActionContract = _','Result.LoadFailed',2,5),
            @('TestReusableProductionRestartActionContract = _','Result.TablesFailed',3,5),
            @('TestReusableProductionRestartActionContract = _','Result.Success',4,5)))
    )
    foreach($group in $groups){foreach($step in $group[2]){
        [pscustomobject]@{Module=$group[0];Procedure=$group[1];Anchor=$step[0];Stage=$step[1];Occurrence=[int]$step[2];Matches=[int]$step[3]}
    }}
}
function Get-ProductionRestartTraceEdits([hashtable]$Sources) {
    foreach($step in Get-ProductionRestartTracePlan){
        if(-not $Sources.ContainsKey($step.Module)){throw 'Restart trace source absent.'}
        $source=[string]$Sources[$step.Module]
        $pattern='(?ims)^[ \t]*(?:Public|Private) (?:Sub|Function) '+[regex]::Escape($step.Procedure)+'\b.*?^End (?:Sub|Function)\b'
        $matches=[regex]::Matches($source,$pattern)
        if($matches.Count -ne 1){throw 'Restart trace procedure ambiguous.'}
        $body=$matches[0];$offset=[regex]::Matches($source.Substring(0,$body.Index),"`n").Count
        $lines=$body.Value -split '\r?\n'
        $found=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Trim().StartsWith($step.Anchor,[StringComparison]::OrdinalIgnoreCase)){$offset+$i+1}})
        if($found.Count -ne $step.Matches){throw ('Restart trace anchor changed: '+$step.Stage)}
        [pscustomobject]@{Module=$step.Module;Line=$found[$step.Occurrence-1];Stage=$step.Stage;Code=('    TestProductionRestartTrace.Mark "'+$step.Stage+'"')}
    }
}
function Get-ProductionRestartTraceLogger {
    $allow=(@('Arm')+@(Get-ProductionRestartTracePlan|ForEach-Object Stage)|ForEach-Object{'        Case "'+$_+'"'}) -join "`r`n"
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
'@+"`r`n"+$allow+"`r`n"+@'
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
function Install-ProductionRestartTrace {
    param($Excel,[hashtable]$Packages,[string]$PackageRoot)
    $project=$Packages['invSys.Operations.xlam'].VBProject;$sources=@{}
    foreach($name in @('mProduction','frmProduction')){$module=$project.VBComponents.Item($name).CodeModule;$sources[$name]=[string]$module.Lines(1,$module.CountOfLines)}
    $edits=@(Get-ProductionRestartTraceEdits $sources)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestProductionRestartTrace'
    $probe.CodeModule.AddFromString((Get-ProductionRestartTraceLogger))
    foreach($group in $edits|Group-Object Module){
        $module=$project.VBComponents.Item($group.Name).CodeModule
        foreach($edit in $group.Group|Sort-Object Line -Descending){$module.InsertLines($edit.Line,$edit.Code)}
    }
    . (Join-Path $PSScriptRoot 'ProductionBatchBoundaryTrace.ps1')
    Install-ProductionBatchTrace -Excel $Excel -Packages $Packages -PackageRoot $PackageRoot -CompileOnly
}
