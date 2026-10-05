# Disposable, compiled stage markers. Only static stage names/error numbers leave Core.
function Install-ExecutionSetupDiagnostics {
    $project=$packages['invSys.Core.xlam'].VBProject
    foreach($name in @('modExecutionTarget','modExecutionBinding')) {
        $component=$null
        foreach($item in $project.VBComponents){if($item.Name -ceq $name){$component=$item;break}}
        if($null -eq $component){continue}
        $module=$component.CodeModule
        $module.InsertLines($module.CountOfDeclarationLines+1,'Public ExecutionStageForTest As String')
        $start=$module.ProcStartLine('Read',0);$count=$module.ProcCountLines('Read',0)
        $body=$module.Lines($start,$count)
        $markers=if($name -ceq 'modExecutionTarget'){
            @(@('    snapshot = ""','context'),@('    Set fso =','root'),@('    For Each suffix','authority-files'),@('    Set wb = OpenRead','config-read'),@('    Set warehouse =','tables'),@('            If CStr(Value(warehouse, row, "WarehousePurpose"))','purpose'),@('            If StrComp(fso.GetAbsolutePathName','data-root'),@('    For row = 1 To station','station'),@('            path = fso.GetAbsolutePathName','inbox-root'),@('    Read = True','complete'))
        }else{
            @(@('    If Not modExecutionTarget.Read','target'),@('    If Not modTrainingReadContext.Read','policy'),@('    parts = Split','guide'),@('    Set capture =','capture'),@('    For Each step In guide("Steps")','permissions'),@('    For Each step In guide("Observations")','visibility'),@('    If Not modExecutionProfileStore.Latest','profile'),@('    Set builds =','packages'),@('    notice = "": Read = True','complete'))
        }
        foreach($marker in $markers){$needle=$marker[0];$body=$body.Replace($needle,('    ExecutionStageForTest = "'+$marker[1]+'"'+"`r`n"+$needle))}
        $body=$body.Replace('Failed:',('Failed:'+"`r`n"+'    ExecutionStageForTest = ExecutionStageForTest & ":" & CStr(Err.Number)'))
        $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    }
    $facade=$null
    foreach($item in $project.VBComponents){if($item.Name -ceq 'modExecutionRun'){$facade=$item;break}}
    if($null -ne $facade){$facade.CodeModule.AddFromString(@'
Public Function ExecutionSetupStageForTest() As String
    ExecutionSetupStageForTest = modExecutionTarget.ExecutionStageForTest & "|" & modExecutionBinding.ExecutionStageForTest
End Function
'@)}
}
