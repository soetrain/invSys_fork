# Optional tracking faults are confined to disposable Config and Training fixtures.
function Install-ProductionCheckInPolicyProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function CheckPolicyMessageForTest(ByVal policy As String, ByVal authorized As Boolean) As Boolean
    Dim notice As String, expected As String, text As String
    Select Case policy
        Case "Off": notice = ""
        Case "Older": notice = "Tracking unavailable: the saved policy does not include this control."
        Case "Invalid": notice = "Tracking unavailable: the saved tracking policy is invalid."
        Case "StoreUnavailable": notice = "Tracking unavailable: the training record could not be saved."
        Case Else: Exit Function
    End Select
    text = mTxtStatus.Text
    If authorized Then
        If InStr(1, text, "Checked in ", vbBinaryCompare) <> 1 Then Exit Function
        If notice = "" Then
            CheckPolicyMessageForTest = (InStr(1, text, "Tracking unavailable:", vbBinaryCompare) = 0)
        Else
            CheckPolicyMessageForTest = (Right$(text, Len(notice) + 1) = " " & notice)
        End If
    Else
        expected = "Production permission changed. Reopen Production before continuing."
        If notice <> "" Then expected = expected & " " & notice
        CheckPolicyMessageForTest = (StrComp(text, expected, vbBinaryCompare) = 0)
    End If
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function CheckPolicyMessage(ByVal policy As String, ByVal authorized As Boolean) As Boolean
    CheckPolicyMessage = mForm.CheckPolicyMessageForTest(policy, authorized)
End Function
'@)
}

function Test-ProductionCheckInPolicy($Fixture,$Other,$Book,$Decoy,[string]$SelectedKey,[string]$Canary,[string[]]$Keys){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function TrainingPins($Target){
        $result=@{};$training=Join-Path $Target.Root 'Training'
        if(Test-Path -LiteralPath $training){foreach($file in Get-ChildItem -LiteralPath $training -Recurse -File){$result[$file.FullName]=Hash $file.FullName}}
        $result
    }
    function SamePins($Before,$After){
        if($Before.Count -ne $After.Count){return $false}
        foreach($file in $Before.Keys){if(-not $After.ContainsKey($file) -or $After[$file] -cne $Before[$file]){return $false}}
        return $true
    }
    $fixtureRoot=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $ownedRoot=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not $fixtureRoot.StartsWith($ownedRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Policy fixture must remain within the owned test root.'}
    $blocked=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
    $held=$blocked+'-check-in-policy-held'
    foreach($candidate in @($blocked,$held,$Fixture.Config)){
        if(-not [IO.Path]::GetFullPath($candidate).StartsWith($fixtureRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Policy path escaped its disposable fixture.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Preserve existing held policy fixture.'}
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configPin=Hash $Fixture.Config
    $oldTraining=TrainingPins $Fixture;$otherTraining=TrainingPins $Other
    foreach($policy in @('Off','Older','Invalid','StoreUnavailable')){
        $moved=$false;$markerCreated=$false
        try{
            SelectTarget $Fixture
            if($policy -ceq 'Off'){
                if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($false))){throw 'Owning disabled-policy save unavailable; not product RED.'}
            }elseif($policy -cin @('Older','Invalid')){
                $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                try{
                    $headers=Table $cfg 'tblEventTrackingPolicies';$controls=Table $cfg 'tblEventTrackingControls'
                    if($policy -ceq 'Older'){
                        $headers.ListColumns.Item('CatalogVersion').DataBodyRange.Value2=24.0
                        for($i=$controls.ListRows.Count;$i -ge 1;$i--){
                            if([string]$controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -ceq 'PRODUCTION_RUN_CHECK_IN'){$controls.ListRows.Item($i).Delete()}
                        }
                    }else{$headers.ListColumns.Item('SchemaVersion').DataBodyRange.Value2=999.0}
                    $cfg.Save()
                }finally{$cfg.Close($false)}
            }
            $effective=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_CHECK_IN'))
            $valid=if($policy -ceq 'Off'){$effective.StartsWith('True|False|')}elseif($policy -ceq 'StoreUnavailable'){$effective.StartsWith('True|True|')}else{$effective.StartsWith('False|False|')}
            if(-not $valid){throw 'Requested policy fixture not effective; not product RED.'}
            if($policy -ceq 'Older' -and -not ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_CLOSE'))).StartsWith('True|True|')){throw 'Older policy must remain valid for an existing control; not product RED.'}
            Check ('CheckInPolicy.'+$policy+'.EffectiveFixture') $true
            $policyPin=Hash $Fixture.Config
            if($policy -ceq 'StoreUnavailable'){
                if(-not (Test-Path -LiteralPath $blocked -PathType Container)){throw 'Prior Activity store unavailable; not product RED.'}
                Move-Item -LiteralPath $blocked -Destination $held;$moved=$true
                [IO.File]::WriteAllText($blocked,'Disposable Check In activity store is intentionally unavailable.');$markerCreated=$true
            }
            foreach($actor in @('config-producer','config-reader')){
                SelectTarget $Fixture $actor
                [void](Probe 'CheckBaselineReopen' @($Book.Name))
                $authorized=$actor -ceq 'config-producer'
                foreach($mode in @('Reusable','Worksheet')){
                    $before=TrainingPins $Fixture
                    $ready=if($mode -ceq 'Reusable'){[bool](Probe 'CheckBaselineReusableStage' @('Selected'))}else{[bool](Probe 'CheckBaselineWorksheetStage' @($SelectedKey,$Canary))}
                    if(-not $ready){throw 'Policy owner staging unavailable; not product RED.'}
                    $label='CheckInPolicy.'+$policy+'.'+$(if($authorized){'Allowed'}else{'Denied'})+'.'+$mode
                    Check ($label+'.SetupNotUserAction') (SamePins $before (TrainingPins $Fixture))
                    $owner=[string](Probe 'RunLocalOwnerState');$projection=[string](Probe 'RunLocalState')+'|'+[string](Probe 'CheckBaselineWorksheetState')
                    [void](Probe 'CheckBaselineResetOwnerEntries');$decoy.Activate()
                    $visible=$CaptureEvidence -and (($authorized -and ($mode -ceq 'Reusable' -or $policy -ceq 'StoreUnavailable')) -or (-not $authorized -and $policy -ceq 'StoreUnavailable' -and $mode -ceq 'Reusable'))
                    if($visible){[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'CHECK_IN'));$decoy.Activate()}
                    Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
                    Check ($label+'.ExpectedOwnerEntries') ([int](Probe 'CheckBaselineOwnerEntries') -eq $(if($authorized){1}else{0}))
                    if($authorized){
                        if($mode -ceq 'Reusable'){
                            Check ($label+'.CheckedNoteFrozenWithoutCompletion') ([bool](Probe 'CheckBaselineReusableResult' @($true)))
                            Check ($label+'.IdentityUnderMatchingHeading') ([bool](Probe 'CheckBaselineReusableDisplay' @($Keys[0],$Keys[1])))
                        }else{
                            foreach($fact in @('ReachedCheckRows','ExactSelectedKey','CustomValue','CustomFormula','PalettePreserved','HeadersPreserved','DisplayColumns')){
                                Check ($label+'.'+$fact) ([bool](Probe 'CheckBaselineWorksheetFact' @($fact)))
                            }
                        }
                    }else{
                        Check ($label+'.OwnerPreserved') ([string](Probe 'RunLocalOwnerState') -ceq $owner)
                        Check ($label+'.ProjectionPreserved') (([string](Probe 'RunLocalState')+'|'+[string](Probe 'CheckBaselineWorksheetState')) -ceq $projection)
                    }
                    Check ($label+'.ExactMessageAndNoticePolicy') ([bool](Probe 'CheckPolicyMessage' @($policy,$authorized)))
                    Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
                    Check ($label+'.NoFalseFallbackOrChangedRecords') (SamePins $before (TrainingPins $Fixture))
                    Check ($label+'.NoRedirectedRecords') (SamePins $otherTraining (TrainingPins $Other))
                    Check ($label+'.NoPolicyRepairOrRewrite') ((Hash $Fixture.Config) -ceq $policyPin)
                    Check ($label+'.CanonicalEntitiesPreserved') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockSourcePreservedForTest'))
                    if($visible){CaptureOwnedFormByCaptionEvidence 'Production' ('check-in-policy-'+$policy.ToLowerInvariant()+'-'+$(if($authorized){'allowed'}else{'denied'})+'-'+$mode.ToLowerInvariant()+'.png')}
                }
            }
        }finally{
            if($markerCreated){Remove-Item -LiteralPath $blocked -Force}
            if($moved){Move-Item -LiteralPath $held -Destination $blocked}
            # Config readers close their handles; never overwrite an open fixture.
            if(@($excel.Workbooks|Where-Object{$_.FullName -ieq $Fixture.Config}).Count){throw 'Close the owned Config fixture before restoring its bytes.'}
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        }
        Check ('CheckInPolicy.'+$policy+'.ConfigBytesRestored') ((Hash $Fixture.Config) -ceq $configPin)
    }
    # Existing evidence must survive the temporary unavailable-store fixture.
    $after=TrainingPins $Fixture;$retained=$true
    foreach($file in $oldTraining.Keys){$retained=$retained -and $after.ContainsKey($file) -and $after[$file] -ceq $oldTraining[$file]}
    Check 'CheckInPolicy.OlderTrainingEvidenceImmutable' $retained
    SelectTarget $Fixture 'config-producer';[void](Probe 'CheckBaselineReopen' @($Book.Name))
}
