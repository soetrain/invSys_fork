# Interrupt the actual Next Batch owner return or an existing inventory-read return.
function Install-ProductionNextYieldProbe {
    $module=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modProductionNextActions').CodeModule
    $start=$module.ProcStartLine('Execute',0);$end=$start+$module.ProcCountLines('Execute',0)
    $anchors=@(for($i=$start;$i -lt $end;$i++){
        if($module.Lines($i,1).Trim() -ceq 'action.OutcomeCode = "FAILED"' -or
           $module.Lines($i,1).Trim() -ceq 'mProduction.BtnNextBatch completed, report'){$i+1}
    })
    if($anchors.Count -ne 2){throw 'Next Batch owner-return anchors changed; not product RED.'}
    foreach($line in @($anchors|Sort-Object -Descending)){
        $module.InsertLines($line,'        TestProductionDesigner.CheckYieldReturned "NextOwner", True')
    }
}

function Test-ProductionNextYield($Fixture,$Other,$Book,$Decoy,[string]$Canary){
    . (Join-Path $PSScriptRoot 'Slice4beProductionNextPaths.ps1')
    function SavedHash([string]$Path){Hash $Path}
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($auth);$authPin=Hash $auth
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    if(-not $root.StartsWith([IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){
        throw 'Next interruption fixture escaped owned root.'
    }
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'),[ref]$tokens,[ref]$errors)
    $definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Start-DialogCaptureAndDismiss'},$false)
    if($errors.Count -or $null -eq $definition){throw 'Native notice observer unavailable; not product RED.'}
    . ([scriptblock]::Create($definition.Extent.Text))
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Next interruption policy unavailable; not product RED.'}
    $catalog=[int](Run 'invSys.Core.xlam' 'TestShippingCatalog.NextPolicyCatalogVersion')
    $cases=@(
        @{Mode='Reusable';Boundary='NextOwner'},@{Mode='Reusable';Boundary='RunPalette'},
        @{Mode='Worksheet';Boundary='NextOwner'},@{Mode='Worksheet';Boundary='InventoryPicker'}
    )
    foreach($interruption in @('SignedOut','Permission')){foreach($case in $cases){
        $work=$null;$observer=$null
        try{
            [void](Probe 'CheckYieldReset');SelectTarget $Fixture 'config-producer'
            $work=$excel.Workbooks.Add();$sheet=$work.Worksheets.Item(1)
            $sheet.Cells.Item(2,1).Value2=$Canary;$sheet.Cells.Item(2,2).Formula='=1+2'
            [void](Probe 'RunLocalReopen' @($work.Name));Prepare-NextPath $work $Canary $case.Mode
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $pins=Get-NextPathAuthorityPins $Fixture
            $history=@{};foreach($path in $before){$history[$path]=Hash $path}
            $label='NextYield.'+$interruption+'.'+$case.Mode+'.'+$case.Boundary
            [void](Probe 'CheckYieldArm' @($case.Boundary,1,$interruption,$auth));$Decoy.Activate()
            if($case.Mode -ceq 'Worksheet'){
                $processes=@(Get-Process EXCEL);if($processes.Count -ne 1){throw 'Isolated Excel required.'}
                $stop=Join-Path $runRoot ($label+'-notice-stop')
                $observer=Start-DialogCaptureAndDismiss -ExcelProcessId $processes[0].Id -TimeoutSeconds 30 -StopPath $stop
            }
            try{$returned=[bool](Probe 'NextActivityAct' @(''))}finally{
                if($null -ne $observer){
                    [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
                    $notice=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
                    Check ($label+'.NativeOwnerNotice') ($observer.State -eq 'Completed' -and $notice.Contains('Next Batch ready. Inventory selections cleared for unchecked processes.'))
                    if($observer.State -ne 'Completed'){Stop-Job $observer};Remove-Job $observer;$observer=$null
                }
            }
            $facts=([string](Probe 'CheckYieldEvidence')).Split('|')
            if($facts.Count -ne 4 -or $facts[0] -cne 'True' -or $facts[1] -cne 'True' -or $facts[3] -cne 'True'){
                throw ('Next Batch real boundary/interruption unavailable: '+$case.Mode+'/'+$case.Boundary+'; not product RED.')
            }
            Check ($label+'.ActualHandlerAndRestoredGuards') $returned
            Check ($label+'.NoLaterInventoryReads') ([int]$facts[2] -eq 0)
            Check ($label+'.OwnerAtBoundaryPreserved') ([bool](Probe 'CheckYieldOwnerPreserved'))
            Check ($label+'.ProjectionAtBoundaryPreserved') ([bool](Probe 'CheckYieldProjectionPreserved'))
            $refused=if($interruption -ceq 'Permission'){[bool](Probe 'CheckBaselinePermissionRefused')}else{[bool](Probe 'CheckBaselineContextRefused')}
            Check ($label+'.VisibleContextRefusal') $refused
            $outcome=if($interruption -ceq 'Permission'){'FAILED'}else{'REQUESTED'}
            Test-ProductionCheckInJournal $Fixture $Other $before $otherBefore $outcome $label $Canary -ControlId 'PRODUCTION_RUN_NEXT_BATCH' -CatalogVersion $catalog
            $unchanged=$true;foreach($path in $history.Keys){$unchanged=$unchanged -and (Test-Path -LiteralPath $path) -and (Hash $path) -ceq $history[$path]}
            Check ($label+'.PriorHistoryImmutable') $unchanged
            if($case.Mode -ceq 'Worksheet'){
                foreach($fact in @('Output','Identity','Custom','Palette')){Check ($label+'.Owner'+$fact) ([bool](Probe 'NextActivityWorksheetFact' @($fact,$Canary)))}
            }else{Check ($label+'.ExactlyOneOwnerEntry') ([int](Probe 'NextBaselineEntries') -eq 1)}
            Check ($label+'.CapturedCustomValueAndFormula') ($sheet.Cells.Item(2,1).Value2 -ceq $Canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
            Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
            if($case.Boundary -ceq 'NextOwner'){CaptureOwnedFormByCaptionEvidence 'Production' ($label.ToLowerInvariant()+'.png')}
        }finally{
            [void](Probe 'CheckYieldReset');[void](Probe 'CloseDesigner')
            if($null -ne $work){$work.Close($false)}
            [IO.File]::WriteAllBytes($auth,$authBytes)
            SelectTarget $Fixture 'config-producer'
        }
        $after=Get-NextPathAuthorityPins $Fixture;$same=$pins.Count -eq $after.Count
        foreach($path in $pins.Keys){$same=$same -and $after.ContainsKey($path) -and $after[$path] -ceq $pins[$path]}
        Check ($label+'.CanonicalBytesPreservedAfterAuthRestoration') $same
    }}
    Check 'NextYield.AuthBytesRestored' ((Hash $auth) -ceq $authPin)
    [void](Probe 'RunLocalReopen' @($Book.Name))
}
