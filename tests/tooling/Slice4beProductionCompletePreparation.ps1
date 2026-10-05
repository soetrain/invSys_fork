# Explicit native diagnosis in disposable probes; only fixed labels and counts.
# This route never replaces the full completion acceptance gate.
function Install-ProductionCompletePreparationTrace {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
#If VBA7 Then
Private Declare PtrSafe Function GetCurrentProcessPreparation Lib "kernel32" Alias "GetCurrentProcess" () As LongPtr
Private Declare PtrSafe Function GetGuiResourcesPreparation Lib "user32" Alias "GetGuiResources" (ByVal process As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function GetCurrentProcessPreparation Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function GetGuiResourcesPreparation Lib "user32" Alias "GetGuiResources" (ByVal process As Long, ByVal flags As Long) As Long
#End If
'@)
    $path=(Join-Path $reportRoot 'preparation-native-counts.log').Replace('"','""')
    $adapter.AddFromString((@'
Public Sub CompletePreparationMark(ByVal phase As String)
    Dim handle As Integer
    handle = FreeFile
    Open "__TRACE_FILE__" For Append As #handle
    Print #handle, phase & "|" & CStr(GetGuiResourcesPreparation(GetCurrentProcessPreparation(), 0)) & "|" & _
        CStr(GetGuiResourcesPreparation(GetCurrentProcessPreparation(), 2)) & "|" & _
        CStr(GetGuiResourcesPreparation(GetCurrentProcessPreparation(), 1))
    Close #handle
End Sub
'@).Replace('__TRACE_FILE__',$path))
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $groups=@(
        @{Procedure='CompleteBaselinePrepareForTest';Points=@(
            @('If Not CheckBaselineReusableStageForTest(', 'Staging'),
            @('mBtnManagerCheckIn_Click', 'CheckIn'),
            @('modProductionReusableRun.CompleteBaselineRememberBalancesForTest', 'Balances'))},
        @{Procedure='CheckBaselineReusableStageForTest';Points=@(
            @('If RunLocalStageForTest(', 'LoadAndStage'),
            @('RefreshReusableRunControls False', 'AllocatedRefresh'))},
        @{Procedure='RunLocalStageForTest';Points=@(
            @('RefreshRecipeLists', 'RecipeList'),
            @('If Not LoadReusableRecipeIntoRun(', 'RecipeLoad'),
            @('RefreshReusableRunControls False', 'StagedRefresh'))}
    )
    foreach($group in $groups){
        $start=$form.ProcStartLine($group.Procedure,0);$count=$form.ProcCountLines($group.Procedure,0)
        $original=$form.Lines($start,$count) -split '\r?\n'
        $markers=@{}
        foreach($point in $group.Points){
            $hits=@(for($line=0;$line -lt $original.Count;$line++){if($original[$line].Trim().StartsWith($point[0],[StringComparison]::OrdinalIgnoreCase)){$line}})
            if($hits.Count -ne 1){throw ('Preparation trace anchor unavailable: '+$group.Procedure+'/'+$point[1])}
            $markers[[int]$hits[0]]=$point[1]
        }
        $replacement=@(for($line=0;$line -lt $original.Count;$line++){
            if($markers.ContainsKey($line)){'    TestProductionDesigner.CompletePreparationMark "Before'+$markers[$line]+'"'}
            $original[$line]
            if($markers.ContainsKey($line)){'    TestProductionDesigner.CompletePreparationMark "After'+$markers[$line]+'"'}
        }) -join "`r`n"
        $form.DeleteLines($start,$count);$form.InsertLines($start,$replacement)
        $readback=$form.Lines($form.ProcStartLine($group.Procedure,0),$form.ProcCountLines($group.Procedure,0))
        foreach($line in $markers.Keys){
            $bracket='    TestProductionDesigner.CompletePreparationMark "Before'+$markers[$line]+'"'+"`r`n"+$original[$line]+"`r`n"+'    TestProductionDesigner.CompletePreparationMark "After'+$markers[$line]+'"'
            if($readback.IndexOf($bracket,[StringComparison]::OrdinalIgnoreCase) -lt 0){throw 'Preparation trace statement bracket was not preserved.'}
        }
    }
}
