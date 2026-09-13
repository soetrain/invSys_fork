# Tests real handler yield points and the callback registered with OnTime.
# Owner probes remain enabled by the context matrix; no business owner is run.
function Test-Slice4beShippingInterruptions($Fixture,$Operator,$Ship,$Hold) {
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $authorityHash=Get-ShippingActivityHash $inventory
    foreach($action in @('Add','Stage','Send')) {
        SelectTarget $Fixture 'config-reader'
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
        $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @($action))
        if(-not $key){throw 'Shipping UI-yield fixture selection unavailable.'}
        $beforeRows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
        $beforeHeld=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
        $beforeActivity=@(Get-Slice4beActivityFiles $Fixture)
        $before=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount')
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingInterruptAtPending')
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @($action))
        $label='Shipping.DuringPending.'+$action
        Check ($label+'.RealYieldInterruptedOnce') ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingInterruptionCount') -eq 1 -and -not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))
        Check ($label+'.StopsBeforeOwner') ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount') -eq $before)
        $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
        Check ($label+'.VisibleContextRejection') ($status -match '(?i)session' -and $status -match '(?i)reopen')
        Check ($label+'.StagingPreserved') ($beforeRows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $beforeHeld -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
        Check ($label+'.NoCrossContextActivity') (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -notin $beforeActivity}).Count -eq 0)
        Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)))
    }
    foreach($context in @('Healthy','SignedOut')) {
        SelectTarget $Fixture 'config-reader'
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
        $pending=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingArmTimer')
        Check ('Shipping.AutoSync.'+$context+'.PendingFixtureEstablished') ($pending -gt 0)
        if($pending -le 0){throw 'Shipping auto-sync fixture has no pending rows.'}
        if($context -eq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
        $beforeRows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
        $beforeHeld=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
        $beforeActivity=@(Get-Slice4beActivityFiles $Fixture)
        $before=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount')
        try {
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.TriggerShipmentsFormAutoSync')
            $after=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount')
            $scheduled=[bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingTimerScheduled')
            $label='Shipping.AutoSync.'+$context
            if($context -eq 'Healthy') {
                Check ($label+'.CalibratedOwnerEntry') ($after -eq $before+1)
                Check ($label+'.PendingWorkRescheduled') $scheduled
            } else {
                Check ($label+'.StopsBeforeOwner') ($after -eq $before)
                Check ($label+'.StaleTimerNotRescheduled') (-not $scheduled)
                $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
                Check ($label+'.VisibleContextRejection') ($status -match '(?i)session' -and $status -match '(?i)reopen')
            }
            Check ($label+'.NoUserActivity') (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -notin $beforeActivity}).Count -eq 0)
            Check ($label+'.StagingPreserved') ($beforeRows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $beforeHeld -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
        } finally {
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingCancelTimer')
        }
    }
    Check 'Shipping.Interruptions.AuthorityBytesPreserved' ($authorityHash -ceq (Get-ShippingActivityHash $inventory))
}
