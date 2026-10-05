# Explicit native diagnosis in disposable probes; only fixed labels and counts.
# This route never replaces the full completion acceptance gate.
function Install-ProductionCompletePreparationTrace {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $declarations=@'
#If VBA7 Then
Private Declare PtrSafe Function GetCurrentProcessPreparation Lib "kernel32" Alias "GetCurrentProcess" () As LongPtr
Private Declare PtrSafe Function GetCurrentProcessIdPreparation Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare PtrSafe Function GetGuiResourcesPreparation Lib "user32" Alias "GetGuiResources" (ByVal process As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function GetCurrentProcessPreparation Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function GetCurrentProcessIdPreparation Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function GetGuiResourcesPreparation Lib "user32" Alias "GetGuiResources" (ByVal process As Long, ByVal flags As Long) As Long
#End If
'@
    $adapter.InsertLines(1,$declarations)
    $path=(Join-Path $reportRoot 'preparation-native-counts.log').Replace('"','""')
    $trace=(@'
Public Sub CompletePreparationMark(ByVal phase As String)
    Dim handle As Integer
    handle = FreeFile
    Open "__TRACE_FILE__" For Append As #handle
    Print #handle, phase & "|" & CStr(GetGuiResourcesPreparation(GetCurrentProcessPreparation(), 0)) & "|" & _
        CStr(GetGuiResourcesPreparation(GetCurrentProcessPreparation(), 2)) & "|" & _
        CStr(GetGuiResourcesPreparation(GetCurrentProcessPreparation(), 1)) & "|" & CStr(GetCurrentProcessIdPreparation())
    Close #handle
End Sub
'@).Replace('__TRACE_FILE__',$path)
    $adapter.AddFromString($trace)
    $core=$packages['invSys.Core.xlam'].VBProject
    $coreComponent=$core.VBComponents.Add(1)
    $coreComponent.Name='TestCompletePreparation'
    $coreAdapter=$coreComponent.CodeModule
    $coreAdapter.InsertLines(1,$declarations)
    $coreAdapter.AddFromString($trace)
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
    # Mark boundaries before statements so continued calls and block Ifs retain
    # their exact semantics. Fixed labels/counts only; no operational values.
    $boundaries=@(
        @{Module='modProductionCheckInActions';Procedure='Execute';Points=@(
            @('If Not action.Begin(', 'ActionBegin'),
            @('Set owner.RunActionContinuation = action', 'ActionAuthorized'),
            @('action.OutcomeCode = owner.CheckInProductionRun()', 'CheckInOwner'),
            @('report = owner.RunAllocationStatus()', 'CheckInOwnerReturned'),
            @('If Not action Is Nothing Then action.Finish report', 'ActionFinish'),
            @('busy = False', 'ActionFinishReturned'))},
        @{Module='frmProduction';Procedure='CheckInProductionRun';Points=@(
            @('If Not SyncReusableRunBatchNote(', 'BatchNote'),
            @('reusableChecked = modProductionReusableRun.CheckInReusableProcess(', 'CheckInService'),
            @('If reusableChecked Then RefreshReusableRunControls', 'CheckInServiceReturned'),
            @('If reusableChecked Then CheckInProductionRun =', 'CheckInRefreshReturned'),
            @('ShowStatus reusableReport', 'CheckInStatus'))},
        @{Module='modProductionReusableRun';Procedure='CheckInReusableProcess';Points=@(
            @('If Not ValidateProcessRequirementsReady(', 'RequirementsRead'),
            @('If Not ValidateProcessAllocationsLive(', 'AllocationsRead'),
            @('mCheckedIn = True', 'AllocationsReadReturned'))},
        @{Module='frmProduction';Procedure='RefreshReusableRunControls';Points=@(
            @('loaderRows =', 'RefreshLoaderRead'),
            @('paletteRows =', 'RefreshPaletteRead'),
            @('checkRows =', 'RefreshCheckRead'),
            @('outputRows =', 'RefreshOutputRead'),
            @('FillListFromArray mLstLoaderLines,', 'RefreshFillLists'),
            @('activeOutputRow =', 'RefreshListsFilled'))},
        @{Core=$true;Module='modRoleUiAccess';Procedure='CanCurrentUserPerformCapability';Points=@(
            @('If Not modConfig.LoadConfig(', 'CapabilityConfigLoad'),
            @('If resolvedWh = "" Then resolvedWh = modConfig.GetWarehouseId()', 'CapabilityConfigReturned'),
            @('If Not modAuth.LoadAuth(', 'CapabilityAuthLoad'),
            @('If Not modAuth.CanPerform(', 'CapabilityAuthReturned'),
            @('CanCurrentUserPerformCapability = True', 'CapabilityAllowed'))},
        @{Core=$true;Module='modConfig';Procedure='LoadConfig';Points=@(
            @('Set wb = ResolveExistingConfigForRead(', 'ConfigResolve'),
            @('openedTransient =', 'ConfigResolved'),
            @('Set loWh =', 'ConfigHidden'),
            @('CloseTransientConfigAfterLoad wb,', 'ConfigClose'),
            @('End Function', 'ConfigClosed'))},
        @{Core=$true;Module='modAuth';Procedure='LoadAuth';Points=@(
            @('Set wb = ResolveAuthWorkbook(', 'AuthResolve'),
            @('openedTransient =', 'AuthResolved'),
            @('mAuthWorkbook = wb.Name', 'AuthHidden'),
            @('CloseTransientAuthAfterLoad wb,', 'AuthClose'),
            @('End Function', 'AuthClosed'))}
    )
    foreach($group in $boundaries){
        $traceOwner='TestProductionDesigner';$targetProject=$project
        if($group.ContainsKey('Core')){$traceOwner='TestCompletePreparation';$targetProject=$core}
        $module=$targetProject.VBComponents.Item($group.Module).CodeModule
        $start=$module.ProcStartLine($group.Procedure,0);$count=$module.ProcCountLines($group.Procedure,0)
        $original=$module.Lines($start,$count) -split '\r?\n';$markers=@{}
        foreach($point in $group.Points){
            $hits=@(for($line=0;$line -lt $original.Count;$line++){if($original[$line].Trim().StartsWith($point[0],[StringComparison]::OrdinalIgnoreCase)){$line}})
            # ShowStatus has a refusal branch as well as the successful branch.
            if($point[1] -eq 'CheckInStatus'){$hits=@($hits|Select-Object -Last 1)}
            if($hits.Count -ne 1){throw ('Check In trace anchor unavailable: '+$group.Procedure+'/'+$point[1])}
            $markers[[int]$hits[0]]=$point[1]
        }
        $replacement=@(for($line=0;$line -lt $original.Count;$line++){
            if($markers.ContainsKey($line)){'    '+$traceOwner+'.CompletePreparationMark "'+$markers[$line]+'"'}
            $original[$line]
        }) -join "`r`n"
        $module.DeleteLines($start,$count);$module.InsertLines($start,$replacement)
        $readback=$module.Lines($module.ProcStartLine($group.Procedure,0),$module.ProcCountLines($group.Procedure,0))
        foreach($line in $markers.Keys){
            $bracket='    '+$traceOwner+'.CompletePreparationMark "'+$markers[$line]+'"'+"`r`n"+$original[$line]
            if($readback.IndexOf($bracket,[StringComparison]::OrdinalIgnoreCase) -lt 0){throw 'Check In trace statement was not preserved.'}
        }
    }
}
