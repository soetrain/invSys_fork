# D18 captured binding and D-NAS Core authorization through real completion.
# Reuse the existing disposable Check In interruption fixture; no owner is replaced.
function Install-ProductionCompleteInterruptionProbe {
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    $start=$form.ProcStartLine('CompleteProductionRun',0);$end=$start+$form.ProcCountLines('CompleteProductionRun',0)
    $anchor='ShowPersistencePending "Saving the reusable Process run to the warehouse server..."'
    $hits=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim() -ieq $anchor){$line}})
    if($hits.Count -ne 1){throw 'Complete Run pending-yield anchor changed; not product RED.'}
    $form.InsertLines($hits[0]+1,'    TestProductionDesigner.CheckYieldReturned "CompletePending", True')
}

function Test-ProductionCompleteInterruptions($Fixture,$Other,$Book,$Decoy,[string]$Canary){
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath);$authHash=(Get-FileHash -LiteralPath $authPath).Hash
    $sheet=$Book.Worksheets.Item(1);$decoySheet=$Decoy.Worksheets.Item(1)
    try{
        foreach($interruption in @('SignedOut','Permission')){
            foreach($boundary in @('CompletePending','AvailableQuantity','EntityKind')){
                try{
                    [void](Probe 'CheckYieldReset')
                    SelectTarget $Fixture 'config-producer'
                    [void](Probe 'RunLocalReopen' @($Book.Name))
                    if(-not [bool](Probe 'CompleteBaselinePrepare' @($true))){throw 'Completion interruption prerequisite unavailable; not product RED.'}
                    [void](Probe 'RunLocalShowAndCapture' @($Book.Name,'CHECK_IN'));$Decoy.Activate()
                    [void](Probe 'CheckYieldArm' @($boundary,1,$interruption,$authPath))
                    $returned=[bool](Probe 'CompleteBaselineAct')
                    $evidence=([string](Probe 'CheckYieldEvidence')).Split('|')
                    if($evidence.Count -ne 4 -or $evidence[0] -cne 'True' -or $evidence[1] -cne 'True' -or $evidence[3] -cne 'True'){
                        throw ('Completion real boundary/interruption fixture unavailable: '+$interruption+'/'+$boundary+'; not product RED.')
                    }
                    $label='CompleteYield.'+$interruption+'.'+$boundary
                    Check ($label+'.ActualHandlerReturned') $returned
                    Check ($label+'.RealBoundaryAndInterruption') $true
                    Check ($label+'.NoLaterReads') ([int]$evidence[2] -eq 0)
                    $expectedOwners=if($boundary -ceq 'CompletePending'){'0|0'}else{'1|0'}
                    Check ($label+'.OwnerEntryBoundary') ([string](Probe 'CompleteBaselineOwners') -ceq $expectedOwners)
                    Check ($label+'.OwnerAtBoundaryPreserved') ([bool](Probe 'CheckYieldOwnerPreserved'))
                    Check ($label+'.ProjectionAtBoundaryPreserved') ([bool](Probe 'CheckYieldProjectionPreserved'))
                    $refused=if($interruption -ceq 'Permission'){[bool](Probe 'CheckBaselinePermissionRefused')}else{[bool](Probe 'CheckBaselineContextRefused')}
                    Check ($label+'.VisibleInterruptionRefusal') $refused
                    CaptureOwnedFormByCaptionEvidence 'Production' ('complete-yield-'+$interruption.ToLowerInvariant()+'-'+$boundary.ToLowerInvariant()+'.png')
                }finally{
                    [void](Probe 'CheckYieldReset')
                    if($interruption -ceq 'Permission'){
                        [IO.File]::WriteAllBytes($authPath,$authBytes)
                        [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
                    }
                }
                SelectTarget $Fixture 'config-producer'
                Check ($label+'.ExactInputBalancesPreserved') ([bool](Owner 'CompleteBaselineBalancesForTest' @($false)))
                Check ($label+'.CapturedBookCustomValueAndFormula') ($sheet.Cells.Item(2,1).Value2 -ceq $Canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
                Check ($label+'.DecoyPreserved') ($decoySheet.Cells.Item(1,1).Value2 -ceq $Canary -and $Decoy.Worksheets.Count -eq 1)
            }
        }
        Check 'CompleteYield.FixtureAuthorizationRestored' ((Get-FileHash -LiteralPath $authPath).Hash -ceq $authHash)
    }finally{
        [void](Probe 'CheckYieldReset')
        SelectTarget $Fixture 'config-producer'
    }
}
