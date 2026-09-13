# D18 handler guard matrix. The normal sequence in the parent helper exercises
# the real owners. This probe deliberately stops at owner entry, after counting
# it, so stale-context coverage cannot mutate or submit business data.
function Test-Slice4beShippingContextMatrix($Fixture,$Operator,$Other,$Ship,$Hold) {
    $second=NewFixture 'shipping-context-other'
    SelectTarget $Fixture 'config-reader'
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClose')
    $Operator.Activate()
    $name=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingOpen')
    if($name -cne $Operator.Name){throw 'Shipping context fixture rebound to another workbook.'}
    # Prepare genuine active and held lines through normal owners before the
    # probe is enabled. Do not inject business inventory or bypass authorization.
    foreach($action in @('Add','Hold','Add')) {
        $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @($action))
        if(-not $key){throw 'Shipping context staging selection unavailable.'}
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @($action))
    }
    $active=@(Get-ShippingActivityRows $Ship|Where-Object {[double]$_.QUANTITY -gt 0})
    $held=@(Get-ShippingActivityRows $Hold|Where-Object {[double]$_.QUANTITY -gt 0})
    $ready=$active.Count -eq 1 -and $held.Count -eq 1
    Check 'Shipping.ContextMatrix.GenuineActiveAndHeldFixture' $ready
    if(-not $ready){throw 'Shipping context fixture requires one actual active and held line.'}
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $authorityHash=Get-ShippingActivityHash $inventory
    $otherHash=Get-ShippingActivityHash $Other.FullName
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSetProbeMode' @($true))
    try {
        Check 'Shipping.ContextMatrix.Healthy.PublicLauncherReusesForm' ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses'))
        foreach($context in @('Healthy','SignedOut','Reauthenticated','OtherTarget')) {
            switch($context) {
                'SignedOut' {
                    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
                    Check 'Shipping.ContextMatrix.SignedOut.Established' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))
                }
                'Reauthenticated' {
                    $version=[long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion')
                    $signed=[string](Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @('config-reader',$Fixture.Secret,''))
                    $ready=$signed.StartsWith('OK|') -and [long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion') -ne $version
                    Check 'Shipping.ContextMatrix.Reauthenticated.Established' $ready
                    if(-not $ready){throw 'Shipping context reauthentication failed.'}
                }
                'OtherTarget' { SelectTarget $second 'config-reader' }
            }
            foreach($action in @('Add','Update','Remove','Hold','Return','Stage','Send')) {
                $label='Shipping.ContextMatrix.'+$context+'.'+$action
                $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @($action))
                if($context -eq 'Healthy' -and -not $key){throw 'Shipping healthy guard calibration selection unavailable.'}
                $beforeRows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
                $beforeHeld=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
                $beforeActivity=@(Get-Slice4beActivityFiles $Fixture)
                $beforeOtherActivity=@(Get-Slice4beActivityFiles $second)
                $before=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount')
                $Other.Activate()
                [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @($action))
                $after=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount')
                if($context -eq 'Healthy') {
                    Check ($label+'.CalibratedOwnerEntry') ($after -eq $before+1)
                } else {
                    Check ($label+'.StopsBeforeOwner') ($after -eq $before)
                    $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
                    Check ($label+'.VisibleContextRejection') ($status -match '(?i)session' -and $status -match '(?i)reopen')
                    if($CaptureEvidence -and $context -eq 'SignedOut' -and $action -eq 'Add') {
                        CaptureFormEvidence 'Shipping Shipments' 'shipping-stale-session.png'
                    }
                    Check ($label+'.NoCrossContextActivity') (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -notin $beforeActivity}).Count -eq 0 -and @(Get-Slice4beActivityFiles $second|Where-Object {$_ -notin $beforeOtherActivity}).Count -eq 0)
                }
                $afterRows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
                $afterHeld=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
                Check ($label+'.StagingAndUnknownValuesPreserved') ($beforeRows -ceq $afterRows -and $beforeHeld -ceq $afterHeld)
                Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($name)))
            }
            Check ('Shipping.ContextMatrix.'+$context+'.AuthorityBytesPreserved') ($authorityHash -ceq (Get-ShippingActivityHash $inventory))
            Check ('Shipping.ContextMatrix.'+$context+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel')
        }
        SelectTarget $Fixture 'config-reader'
        $beforeRows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
        $beforeHeld=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
        $reused=[bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
        Check 'Shipping.ContextMatrix.Reopen.ReplacesStaleForm' (-not $reused)
        Check 'Shipping.ContextMatrix.Reopen.RetainsWorkbook' ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($name)))
        Check 'Shipping.ContextMatrix.Reopen.PreservesStaging' ($beforeRows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $beforeHeld -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
        Check 'Shipping.ContextMatrix.Reopen.NextLaunchReusesForm' ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses'))
        Test-Slice4beShippingInterruptions $Fixture $Operator $Ship $Hold
    } finally {
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSetProbeMode' @($false))
        SelectTarget $Fixture 'config-reader'
    }
}
