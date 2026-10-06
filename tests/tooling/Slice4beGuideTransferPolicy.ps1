# Optional observation loss never grants authority or changes a permitted transfer.
# All fault injection is confined to the generated fixture; owner handlers stay real.
function Test-GuideTransferPolicy($Fixture,$Other,$Guide) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $targetRoot=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    if(-not $targetRoot.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated transfer policy fixture required.'}
    $activityRoot=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
    $held=$activityRoot+'-transfer-policy-held'
    foreach($path in @($activityRoot,$held,$Fixture.Config)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($targetRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Transfer policy path escaped its fixture.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Preserve existing held activity fixture.'}
    $files=Join-Path $runRoot 'transfer-policy';New-Item -ItemType Directory -Path $files|Out-Null
    $packagePath=Join-Path $runRoot 'transfer-activity/export.json'
    if(-not(Test-Path -LiteralPath $packagePath -PathType Leaf)){throw 'Retained actual export prerequisite missing.'}
    $packageHash=(Get-FileHash -LiteralPath $packagePath).Hash
    $invalid=Join-Path $files 'invalid.json';[IO.File]::WriteAllText($invalid,'{}',[Text.Encoding]::ASCII)
    $guideRoot=Join-Path $Fixture.Root ('Training/ActionPaths/'+$Fixture.Warehouse+'/Guides')
    $otherPins=BoundPins $Other.Root;$originalActivity=BoundPins $activityRoot
    $initialVersion=Get-TrackingPolicyVersion $Fixture

    function ConfigurePolicy([string]$Mode) {
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
        $previous=Get-TrackingPolicyVersion $Fixture
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        try {
            if((UserPolicyControl 'btnResetTrackingPolicy' 'Click') -cne 'CLICKED'){throw 'Actual policy Reset control unavailable.'}
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($true))
            if($Mode -ceq 'ControlOff'){
                foreach($id in @('VIEWER_GUIDE_EXPORT','VIEWER_GUIDE_IMPORT')){
                    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyControl' @($id,$false))
                }
            }
            if($Mode -ceq 'UserOff'){
                if((UserPolicyControl 'lstTrackingUsers' 'Select' 'config-admin') -cne 'SELECTED'){throw 'Generated actor not selectable.'}
                if((UserPolicyControl 'chkUserRecord' 'Write' 'False') -cne 'SET'){throw 'Actual user recording control unavailable.'}
            }
            $entries=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyEntries')
            $saved=(UserPolicyControl 'btnSaveTrackingPolicy' 'Click') -ceq 'CLICKED'
            $rendered=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyShowsSaved')
            $dispatched=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyEntries') -eq $entries+1
            if(-not($saved -and $rendered -and $dispatched)){throw 'Actual policy save prerequisite failed; not product RED.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
        $version=Get-TrackingPolicyVersion $Fixture
        return [pscustomobject]@{Version=$version;Previous=$previous}
    }

    function PolicyAction([string]$Mode,[string]$Case) {
        $label='GuideTransferPolicy.'+$policyMode+'.'+$Mode+'.'+$Case
        $path='';$outcome='CANCELLED';$message=$Mode+' cancelled.'
        switch($Case){
            Success {
                $outcome='COMPLETED'
                if($Mode -ceq 'Export'){$path=Join-Path $files ($policyMode+'-export.json');$message='Guide exported.'}
                else{$path=$packagePath;$message='Guide imported. Imported origin evidence; not locally observed.'}
            }
            Rejected {
                $outcome='REJECTED'
                if($Mode -ceq 'Export'){$path=$packagePath;$message='Export destination already exists. Choose a new file.'}
                else{$path=$invalid;$message='Unavailable: the transfer file has invalid integrity, schema, bounds or provenance.'}
            }
            PickerFailure {
                $outcome='FAILED';$path='TRANSFER_PICKER_FAULT'
                $message='Unavailable: the guide file could not be selected for '+$Mode.ToLowerInvariant()+'.'
            }
        }
        $before=TransferPins $Fixture.Root;$beforeActivity=BoundPins $activityRoot
        $configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        TransferSetFile $path
        $delivered=(BoundControl ('btn'+$Mode+'Guide') 'Click') -ceq 'DELIVERED'
        $status=BoundControl 'lblPublishedGuideStatus' 'Label'
        $notice=switch($policyMode){
            Older {'Tracking unavailable: the saved policy does not include this control.'}
            StoreUnavailable {'Tracking unavailable: the training record could not be saved.'}
            default {''}
        }
        $expected=$message;if($notice){$expected+="`r`n"+$notice}
        $after=TransferPins $Fixture.Root;$afterActivity=BoundPins $activityRoot
        $added=@($after.Keys|Where-Object {-not $before.ContainsKey($_)})
        $newActivity=@($afterActivity.Keys|Where-Object {-not $beforeActivity.ContainsKey($_)})
        $records=@($newActivity|ForEach-Object {[IO.File]::ReadAllText($_)|ConvertFrom-Json})
        $owner=$delivered -and (TransferRetained $before $after)
        if($Case -ceq 'Success' -and $Mode -ceq 'Import'){
            $owner=$owner -and $added.Count -eq 1
            if($added.Count -eq 1){$owner=$owner -and (Split-Path $added[0] -Parent) -ceq $guideRoot}
        }else{$owner=$owner -and $added.Count -eq 0}
        if($Case -ceq 'Success' -and $Mode -ceq 'Export'){
            $wire=TransferRead $path 'GuideTransfer' 1 'SchemaVersion|RecordKind|TransferId|ExportedAtUTC|ExportedByUserId|SourceWarehouseId|Guide|ContentSha256'
            $owner=$owner -and $null -ne $wire -and $wire.Guide.ContentSha256 -ceq $Guide.ContentSha256
        }
        $evidence=$records.Count -eq 0
        if($policyMode -ceq 'Recovery'){
            $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $outcome)
            $evidence=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
            if($evidence){$evidence=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId -and @($records|Where-Object {$_.ControlId -cne ('VIEWER_GUIDE_'+$Mode.ToUpperInvariant()) -or $_.UserId -cne 'config-admin' -or $_.WarehouseId -cne $Fixture.Warehouse -or @($_.SourceEventRefs).Count}).Count -eq 0}
        }
        Check ($label+'.ActualHandlerReturned') $delivered
        Check ($label+'.OwnerResultAndOnlyExpectedWrites') $owner
        Check ($label+'.ExactOwnerMessageAndTrackingNotice') ($status -ceq $expected)
        Check ($label+'.ExpectedEvidenceOrSuppression') $evidence
        Check ($label+'.PriorActivityImmutable') (TransferRetained $beforeActivity $afterActivity)
        Check ($label+'.NoPolicyRepairOrRewrite') ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configHash)
        Check ($label+'.SourcePackagePreserved') ((Get-FileHash -LiteralPath $packagePath).Hash -ceq $packageHash)
        Check ($label+'.OtherWarehousePreserved') (BoundSame $otherPins (BoundPins $Other.Root))
        if($Case -ceq 'Success' -and $policyMode -cin @('StoreUnavailable','Recovery')){
            Check ($label+'.MinimumLayout') ((BoundControl '' 'Fit' 'Minimum') -ceq 'True')
            Check ($label+'.ActualActivationDelivered') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.TransferActivateForTest'))
            if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Published guides' ('transfer-policy-'+$policyMode.ToLowerInvariant()+'-'+$Mode.ToLowerInvariant()+'.png')}
            Check ($label+'.NoticeSurvivesLayoutAndCapture') ((BoundControl 'lblPublishedGuideStatus' 'Label') -ceq $expected)
            [void](BoundControl '' 'Fit' 'Restored')
        }
        if($Case -ceq 'Success' -and $policyMode -ceq 'Recovery' -and $Mode -ceq 'Import'){
            $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
            $authBytes=[IO.File]::ReadAllBytes($authPath)
            try {
                $auth=$excel.Workbooks.Open($authPath,0,$false)
                try {
                    $caps=Table $auth 'tblCapabilities';$revoked=0
                    foreach($row in $caps.ListRows){
                        if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ACTION_PATH_MAINT'){
                            $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$revoked++
                        }
                    }
                    if($revoked -ne 1){throw 'Unique maintenance capability fixture required.'}
                    $auth.Save()
                }finally{$auth.Close($false)}
                [void](Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Fixture.Warehouse))
                if([bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))){throw 'Capability revocation prerequisite failed.'}
                $guardPins=TransferPins $Fixture.Root
                $activated=[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.TransferActivateForTest')
                Check ($label+'.ActivationRevalidatesPermissionBeforeRetainingNotice') ($activated -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match 'ACTION_PATH_MAINT' -and (BoundControl 'btnExportGuide' 'Click') -ceq 'DISABLED' -and (BoundSame $guardPins (TransferPins $Fixture.Root)))
            }finally{
                [IO.File]::WriteAllBytes($authPath,$authBytes)
                [void](Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Fixture.Warehouse))
                [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.TransferActivateForTest')
            }
        }
    }

    try {
        foreach($policyMode in @('ControlOff','UserOff','Older','StoreUnavailable','Recovery')){
            $savedPolicy=ConfigurePolicy $policyMode;$version=[int]$savedPolicy.Version
            Check ('GuideTransferPolicy.'+$policyMode+'.ActualPolicyHandlerAppendsOneVersion') ($version -eq $savedPolicy.Previous+1)
            $pristine=$null;$moved=$false;$marker=$false
            try {
                if($policyMode -ceq 'Older'){
                    $pristine=[IO.File]::ReadAllBytes($Fixture.Config)
                    $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                    try {
                        $headers=Table $cfg 'tblEventTrackingPolicies';$controls=Table $cfg 'tblEventTrackingControls';$found=0
                        foreach($row in $headers.ListRows){if([int]$row.Range.Cells.Item(1,$headers.ListColumns.Item('PolicyVersion').Index).Value2 -eq $version){$row.Range.Cells.Item(1,$headers.ListColumns.Item('CatalogVersion').Index).Value2=29.0;$found++}}
                        if($found -ne 1){throw 'Unique current policy required.'}
                        for($index=$controls.ListRows.Count;$index -ge 1;$index--){
                            $row=$controls.ListRows.Item($index)
                            if([int]$row.Range.Cells.Item(1,$controls.ListColumns.Item('PolicyVersion').Index).Value2 -eq $version -and [string]$row.Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -cin @('VIEWER_GUIDE_EXPORT','VIEWER_GUIDE_IMPORT')){$row.Delete()}
                        }
                        $cfg.Save()
                    }finally{$cfg.Close($false)}
                }
                TransferOpen $Fixture;BoundSelect $Guide
                $expected=if($policyMode -cin @('ControlOff','UserOff')){'True|False'}elseif($policyMode -ceq 'Older'){'False|False'}else{'True|True'}
                foreach($id in @('VIEWER_GUIDE_EXPORT','VIEWER_GUIDE_IMPORT')){
                    if([string](Run 'invSys.Core.xlam' 'TestGuideTransferActivity.Policy' @($id)) -cne $expected){throw 'Effective transfer policy prerequisite mismatch; not product RED.'}
                }
                $ordinary=if($policyMode -ceq 'UserOff'){'True|False'}else{'True|True'}
                if([string](Run 'invSys.Core.xlam' 'TestGuideTransferActivity.Policy' @('ADMIN_SETTINGS_SAVE_VALUE')) -cne $ordinary){throw 'Independent valid policy prerequisite failed.'}
                Check ('GuideTransferPolicy.'+$policyMode+'.PermissionStillRequiredAndPresent') ([bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')))
                $heldPins=BoundPins $activityRoot
                if($policyMode -ceq 'StoreUnavailable'){
                    if(-not(Test-Path -LiteralPath $activityRoot -PathType Container)){throw 'Existing activity store required.'}
                    Move-Item -LiteralPath $activityRoot -Destination $held;$moved=$true
                    [IO.File]::WriteAllText($activityRoot,'Disposable transfer activity storage unavailable.');$marker=$true
                }
                foreach($mode in @('Export','Import')){foreach($case in @('Success','Cancel','Rejected','PickerFailure')){PolicyAction $mode $case}}
            }finally{
                TransferSetFile '';CloseRecordingViewer
                if($marker){Remove-Item -LiteralPath $activityRoot -Force}
                if($moved){Move-Item -LiteralPath $held -Destination $activityRoot}
                if($null -ne $pristine){[IO.File]::WriteAllBytes($Fixture.Config,$pristine)}
            }
            Check ('GuideTransferPolicy.'+$policyMode+'.ExistingActivityRestoredOrRetained') (TransferRetained $heldPins (BoundPins $activityRoot))
        }
        Check 'GuideTransferPolicy.PolicyHistoryAppendsFiveVersions' ((Get-TrackingPolicyVersion $Fixture) -eq $initialVersion+5)
        Check 'GuideTransferPolicy.OriginalActivityRemainsImmutable' (TransferRetained $originalActivity (BoundPins $activityRoot))
    } finally {TransferSetFile '';CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
