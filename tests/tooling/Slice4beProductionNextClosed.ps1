# Reuse native workbook-closure receipts; never query a dismissed form's controls.
function Install-ProductionNextClosedProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $line=$form.ProcBodyLine('mBtnManagerNext_Click',0)
    $form.InsertLines($line+1,'    TestProductionDesigner.NextClosedHandlerHit')
    $form.AddFromString(@'
Public Function NextClosedActForTest() As Boolean
    On Error GoTo Failed
    mBtnManagerNext_Click
    NextClosedActForTest = True
Failed:
End Function
Public Function NextClosedGuardsForTest() As Boolean
    NextClosedGuardsForTest = Not mLoading And Not mDesignerActionInProgress
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mNextClosedEntries As Long, mNextClosedRecoveries As Long')
    $adapter.AddFromString(@'
Public Sub NextClosedHandlerHit()
    mNextClosedEntries = mNextClosedEntries + 1
End Sub
Public Function NextClosedAct() As Boolean
    mNextClosedEntries = 0: mNextClosedRecoveries = 0
    NextClosedAct = mForm.NextClosedActForTest()
End Function
Public Function NextClosedEntries() As Long
    NextClosedEntries = mNextClosedEntries
End Function
Public Function NextClosedRecoveries() As Long
    NextClosedRecoveries = mNextClosedRecoveries
End Function
Public Function NextClosedGuards() As Boolean
    NextClosedGuards = mForm.NextClosedGuardsForTest()
End Function
Public Sub NextClosedRecovery()
    mNextClosedRecoveries = mNextClosedRecoveries + 1
    If mNextClosedRecoveries > 1 Then Err.Raise vbObjectError + 2812, "NextClosureProbe", "Repeated Next Batch cleanup failure."
End Sub
'@)
    if($RunNextClosedEscapeForTest){
        $module=$project.VBComponents.Item('modProductionNextActions').CodeModule
        $start=$module.ProcStartLine('Execute',0);$end=$start+$module.ProcCountLines('Execute',0)
        $hits=@(for($i=$start;$i -lt $end;$i++){if($module.Lines($i,1).Trim() -ceq 'Resume Done'){$i}})
        if($hits.Count -ne 1){throw 'Next Batch cleanup diagnostic anchor changed; not product RED.'}
        $module.InsertLines($hits[0],'    TestProductionDesigner.NextClosedRecovery')
    }
}

function Test-ProductionNextClosed($Fixture,$Other,$Book,$Decoy,[string]$Canary){
    . (Join-Path $PSScriptRoot 'Slice4beProductionNextPaths.ps1')
    function SavedHash([string]$Path){Hash $Path}
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'),[ref]$tokens,[ref]$errors)
    $definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Start-DialogCaptureAndDismiss'},$false)
    if($errors.Count -or $null -eq $definition){throw 'Native notice observer unavailable; not product RED.'}
    . ([scriptblock]::Create($definition.Extent.Text))
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Next closure policy unavailable; not product RED.'}
    $catalog=[int](Run 'invSys.Core.xlam' 'TestShippingCatalog.NextPolicyCatalogVersion')
    $cases=@(@{Mode='Reusable';Boundary='NextOwner'},@{Mode='Reusable';Boundary='RunPalette'},
        @{Mode='Worksheet';Boundary='NextOwner'},@{Mode='Worksheet';Boundary='InventoryPicker'})
    foreach($case in $cases){
        $work=$null;$observer=$null
        try{
            [void](Probe 'CheckYieldReset');SelectTarget $Fixture 'config-producer'
            $work=$excel.Workbooks.Add();$sheet=$work.Worksheets.Item(1)
            $sheet.Cells.Item(2,1).Value2=$Canary;$sheet.Cells.Item(2,2).Formula='=1+2'
            $label='NextClosed.'+$case.Mode+'.'+$case.Boundary
            $path=Join-Path $runRoot ($label+'.xlsb')
            if(Test-Path -LiteralPath $path){throw 'Preserve existing native closure fixture.'}
            $work.SaveAs($path,50)
            [void](Probe 'RunLocalReopen' @($work.Name));Prepare-NextPath $work $Canary $case.Mode
            $work.Save();$saved=Hash $path
            Initialize-SettingsCapture
            $visibleBefore=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
            if(-not $visibleBefore){throw 'Native Next Batch surface unavailable; not product RED.'}
            $capturedName=$work.Name;$before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $pins=Get-NextPathAuthorityPins $Fixture;$history=@{};foreach($file in $before){$history[$file]=Hash $file}
            [void](Probe 'CheckClosedYieldArm' @($case.Boundary,1,$Decoy.Name));$Decoy.Activate()
            if($case.Mode -ceq 'Worksheet'){
                $processes=@(Get-Process EXCEL);if($processes.Count -ne 1){throw 'Isolated Excel required.'}
                $stop=Join-Path $runRoot ($label+'-notice-stop')
                $observer=Start-DialogCaptureAndDismiss -ExcelProcessId $processes[0].Id -TimeoutSeconds 30 -StopPath $stop
            }
            $started=[DateTimeOffset]::UtcNow.ToString('o')
            try{$returned=[bool](Probe 'NextClosedAct')}finally{
                if($null -ne $observer){
                    [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
                    $notice=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
                    Check ($label+'.NativeOwnerNotice') ($observer.State -eq 'Completed' -and $notice.Contains('Next Batch ready. Inventory selections cleared for unchecked processes.'))
                    if($observer.State -ne 'Completed'){Stop-Job $observer};Remove-Job $observer;$observer=$null
                }
            }
            $facts=([string](Probe 'CheckYieldEvidence')).Split('|');$native=([string](Probe 'CheckClosedYieldReceipt')).Split('|')
            if($facts.Count -ne 4 -or $facts[0] -cne 'True' -or $facts[1] -cne 'True' -or $facts[3] -cne 'True' -or $native.Count -ne 7 -or $native[0] -cne 'True' -or $native[1] -cne 'True' -or $native[4] -cne 'True' -or $native[5] -cne 'True'){
                throw ('Actual Next Batch boundary/closure unavailable: '+$case.Mode+'/'+$case.Boundary+'; not product RED.')
            }
            $work=$null
            $visibleAfter=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
            $projection=$true;$guards=$true;$refusal=$true
            if($visibleAfter){
                $projection=[bool](Probe 'CheckClosedYieldFormProjectionPreserved')
                $guards=[bool](Probe 'NextClosedGuards');$refusal=[bool](Probe 'CheckBaselineContextRefused')
                CaptureOwnedFormByCaptionEvidence 'Production' ($label.ToLowerInvariant()+'.png')
            }
            $recoveries=[int](Probe 'NextClosedRecoveries');$entries=[int](Probe 'NextClosedEntries')
            Check ($label+'.ActualHandlerReturnsCleanly') ($returned -and $recoveries -eq 0 -and $entries -eq 1)
            Check ($label+'.NativeWorkbookRemoved') ([int]$native[2] -eq [int]$native[3]+1 -and $capturedName -cnotin @($excel.Workbooks|ForEach-Object Name))
            Check ($label+'.NoLaterReads') ([int]$facts[2] -eq 0)
            Check ($label+'.OwnerAtBoundaryPreserved') ([bool](Probe 'CheckYieldOwnerPreserved'))
            Check ($label+'.CapturedBindingRejected') (-not [bool](Probe 'RunLocalClosedBindingCurrent'))
            Check ($label+'.SurvivingProjectionPreserved') $projection
            Check ($label+'.SurvivingGuardsRestored') $guards
            Check ($label+'.DismissedOrVisibleRefusal') $refusal
            Check ($label+'.NoFormReinitialization') ([int]$native[6] -eq 0)
            Check ($label+'.SavedOperatorBytesPreserved') ((Hash $path) -ceq $saved)
            Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
            Test-ProductionCheckInJournal $Fixture $Other $before $otherBefore 'FAILED' $label $Canary -ControlId 'PRODUCTION_RUN_NEXT_BATCH' -CatalogVersion $catalog
            $same=$true;foreach($file in $history.Keys){$same=$same -and (Test-Path -LiteralPath $file) -and (Hash $file) -ceq $history[$file]}
            Check ($label+'.PriorHistoryImmutable') $same
            $after=Get-NextPathAuthorityPins $Fixture;$same=$pins.Count -eq $after.Count
            foreach($file in $pins.Keys){$same=$same -and $after.ContainsKey($file) -and $after[$file] -ceq $pins[$file]}
            Check ($label+'.CanonicalBytesPreserved') $same
            [pscustomobject]@{Mode=$case.Mode;Boundary=$case.Boundary;StartUTC=$started;EndUTC=[DateTimeOffset]::UtcNow.ToString('o');VisibleBefore=$visibleBefore;VisibleAfter=$visibleAfter;HandlerEntries=$entries;HandlerReturned=$returned;CleanupExceptions=$recoveries;DiagnosticEscapeInstalled=[bool]$RunNextClosedEscapeForTest;NativeWorkbookRemoved=$true;DecoyOpen=$true;LaterReads=[int]$facts[2];FormInitializations=[int]$native[6];DismissedControlsQueried=$false}|ConvertTo-Json|Set-Content (Join-Path $reportRoot ($label.ToLowerInvariant()+'.json'))
        }finally{
            [void](Probe 'CheckYieldReset');[void](Probe 'RunLocalSafeClose')
            if($null -ne $work){try{$work.Close($false)}catch{}}
        }
    }
    SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
}
