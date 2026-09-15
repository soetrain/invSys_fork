# D18 optional tracking must preserve actual Make/Unbox submission and application.
# Policies and unavailable storage exist only in the generated disposable fixture.
function Set-BoxingTrackingFixturePolicy($Fixture,[int]$Catalog,[bool]$Collect) {
    $cfg=$null
    try {
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $ids=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(8))).Split("`n")|Where-Object {$_ -ne ''})
        if($ids.Count -ne 31){throw 'Existing catalog-8 policy fixture is unavailable.'}
        if($Catalog -eq 9){$ids+=@('BOXING_MAKE','BOXING_UNBOX')}
        [void](Add-ActivityFixtureTable $cfg 'tblEventTrackingPolicies' @('PolicyVersion','SchemaVersion','CatalogVersion','CreatedAtUTC','CreatedByUserId','DefaultView','ViewerActionPathCaptureEnabled','AdminViewerEventLoggingEnabled','Operator Extra') @(,@(1.0,1.0,[double]$Catalog,'2026-09-07T12:00:00.000Z','config-admin','How-To',$false,$true,'preserve')))
        $rows=@(foreach($id in $ids){,@(1.0,$id,$Collect,$true,$false,'preserve')})
        [void](Add-ActivityFixtureTable $cfg 'tblEventTrackingControls' @('PolicyVersion','ControlId','Collect','Visible','SequenceEligible','Operator Extra') $rows)
        $cfg.Save()
    }finally{if($null -ne $cfg){$cfg.Close($false)}}
}

function Test-Slice4beBoxingTracking($Fixture,$Operator,$Other,$Ship,$Hold) {
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $leaf=Join-Path $Fixture.Root ('Training\Activity\'+$Fixture.Warehouse)
    $held=$leaf+'-boxing-tracking-test-held'
    foreach($path in @($leaf,$held,$Fixture.Config)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Boxing tracking fixture escaped its generated root.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Boxing tracking hold path already exists.'}
    $bytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $originalConfig=Get-ShippingActivityHash $Fixture.Config
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    try {
        foreach($mode in @('UnavailableStore','OlderPolicy','DisabledPolicy')){
            $history=@{};foreach($file in @(Get-Slice4beActivityFiles $Fixture)){$history[$file]=Get-ShippingActivityHash $file}
            $moved=$false;$blocked=$false
            try {
                if($mode -eq 'UnavailableStore'){
                    New-Item -ItemType Directory -Path (Split-Path $leaf -Parent) -Force|Out-Null
                    if(Test-Path -LiteralPath $leaf){Move-Item -LiteralPath $leaf -Destination $held;$moved=$true}
                    [IO.File]::WriteAllText($leaf,'blocked fixture activity path');$blocked=$true
                }else{
                    $catalog=if($mode -eq 'OlderPolicy'){8}else{9}
                    Set-BoxingTrackingFixturePolicy $Fixture $catalog ($mode -eq 'OlderPolicy')
                    $policyPin=Get-ShippingActivityHash $Fixture.Config
                    # Collection-off assertions require a valid catalog-9 policy;
                    # a rejected policy never proves disabled tracking behavior.
                    $known=if($mode -eq 'OlderPolicy'){'True|True|True|1'}else{'True|False|True|1'}
                    Check ('Boxing.Tracking.'+$mode+'.WholePolicyValid') ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('ADMIN_SETTINGS_SAVE_VALUE')) -ceq $known)
                    foreach($action in @('MAKE','UNBOX')){
                        $expected=if($mode -eq 'OlderPolicy'){'False|False|False|0'}else{'True|False|True|1'}
                        Check ('Boxing.Tracking.'+$mode+'.ExplicitControlPolicy.'+$action) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('BOXING_'+$action)) -ceq $expected)
                    }
                    Check ('Boxing.Tracking.'+$mode+'.PolicyReadsPreserveBytes') ($policyPin -ceq (Get-ShippingActivityHash $Fixture.Config))
                }
                $configHash=Get-ShippingActivityHash $Fixture.Config
                foreach($action in @('MAKE','UNBOX')){
                    Reset-BoxingOutcomeFaults
                    $label='Boxing.Tracking.'+$mode+'.'+$action
                    $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @('Prepare'))
                    if(-not $key){throw 'Boxing tracking fixture selection unavailable.'}
                    $rows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
                    $holdRows=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
                    $beforeLog=@(Get-ShippingActivityLog $Fixture)
                    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingResetSources')
                    $queue=Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue')
                    if($null -eq $queue){throw 'Boxing submission counter unavailable.'}
                    $Other.Activate()
                    Check ($label+'.ActualHandlerReturned') ([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @($action,'1')) -ceq 'DELIVERED')
                    $ids=@(([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingReadSources')).Split("`n")|Where-Object {$_ -ne ''})
                    $afterQueue=Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue')
                    $failures=Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSourceFailures'
                    $submitted=$null -ne $afterQueue -and $null -ne $failures -and $ids.Count -eq 1 -and [long]$afterQueue-[long]$queue -eq 1 -and [long]$failures -eq 0
                    Check ($label+'.ExactlyOneOwningSubmission') $submitted
                    if(-not $submitted){throw 'Boxing tracking owning submission not established.'}
                    $pin=Get-ShippingActivityHash $inventory
                    $afterLog=@(Get-ShippingActivityLog $Fixture)
                    Check ($label+'.EvidenceReadPreservesAuthorityBytes') ($pin -ceq (Get-ShippingActivityHash $inventory))
                    $applied=@($afterLog|Where-Object {$_.EventID -cnotin @($beforeLog.EventID)})
                    $type=if($action -ceq 'MAKE'){'BOX_BUILD'}else{'BOX_UNBOX'}
                    $qty=if($action -ceq 'MAKE'){1}else{-1}
                    $owner=$applied.Count -eq 2 -and @($applied|Where-Object {$_.EventID -cne $ids[0] -or $_.EventType -cne $type -or $_.WarehouseId -cne $Fixture.Warehouse}).Count -eq 0 -and
                        @($applied|Where-Object {$_.System_Key -ceq $key -and [double]$_.QtyDelta -eq $qty}).Count -eq 1 -and
                        @($applied|Where-Object {$_.System_Key -cne $key -and $_.System_Key -ne '' -and [double]$_.QtyDelta -eq -$qty}).Count -eq 1
                    Check ($label+'.ExactPackageAndComponentApplication') $owner
                    if(-not $owner){throw 'Boxing tracking independent application not established.'}
                    $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
                    $notice=$status.Contains('Tracking unavailable')
                    Check ($label+'.TrackingNoticeMatchesPolicy') ($status -ne '' -and $notice -eq ($mode -ne 'DisabledPolicy'))
                    if($CaptureEvidence){Capture-BoxingFormEvidence 'Shipping Shipments' ($label.ToLowerInvariant()+'.png') ($label+'.VisibleCapture')}
                    $bound=Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)
                    Check ($label+'.CapturedWorkbookRetained') ($bound -is [bool] -and $bound)
                    Check ($label+'.StagingKeysAndUnknownValuesPreserved') ($rows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $holdRows -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
                    Check ($label+'.PolicyBytesPreserved') ($configHash -ceq (Get-ShippingActivityHash $Fixture.Config))
                    Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping submission sentinel')
                    if($blocked){
                        Check ($label+'.BlockedStorePreserved') ((Test-Path -LiteralPath $leaf -PathType Leaf) -and [IO.File]::ReadAllText($leaf) -ceq 'blocked fixture activity path')
                    }else{
                        Check ($label+'.NoImplicitCollection') (@(Get-Slice4beActivityFiles $Fixture|Where-Object {-not $history.ContainsKey($_)}).Count -eq 0)
                    }
                }
            }finally{
                if($blocked -and (Test-Path -LiteralPath $leaf -PathType Leaf)){Remove-Item -LiteralPath $leaf -Force}
                if($moved){Move-Item -LiteralPath $held -Destination $leaf}
                [IO.File]::WriteAllBytes($Fixture.Config,$bytes)
                [void](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($Fixture.Warehouse,'S1'))
            }
            $after=@(Get-Slice4beActivityFiles $Fixture);$preserved=$after.Count -eq $history.Count
            foreach($file in $after){$preserved=$preserved -and $history.ContainsKey($file) -and $history[$file] -ceq (Get-ShippingActivityHash $file)}
            Check ('Boxing.Tracking.'+$mode+'.PriorActivityBytesPreserved') $preserved
            Check ('Boxing.Tracking.'+$mode+'.OriginalConfigRestored') ($originalConfig -ceq (Get-ShippingActivityHash $Fixture.Config))
        }
    }finally{
        Reset-BoxingOutcomeFaults
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.BoxingActionForTest' @('Shipping'))
    }
}
