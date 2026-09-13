# D18: existing role capability governs the action independently of collection.
# Probe cases stop at calibrated mutation-owner entry; Hold also runs its real owner.
function Test-Slice4beShippingCapability($Fixture,$Operator,$Other,$Ship,$Hold) {
    SelectTarget $Fixture 'config-reader'
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
    $name=[string]$Operator.Name
    $session=[long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion')
    $allowed=[bool](Run 'invSys.Core.xlam' 'TestShippingSession.CanShip')
    Check 'Shipping.Capability.Healthy.CorePermissionEstablished' $allowed
    if(-not $allowed){throw 'Shipping capability fixture lacks initial permission.'}
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $authorityHash=Get-ShippingActivityHash $inventory
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath)
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSetProbeMode' @($true))
    try {
        foreach($phase in @('Healthy','Revoked')) {
            if($phase -eq 'Revoked') {
                $auth=$excel.Workbooks.Open($authPath,0,$false)
                try {
                    $caps=Table $auth 'tblCapabilities';$revoked=0
                    foreach($row in $caps.ListRows) {
                        if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader' -and
                           $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'SHIP_POST') {
                            $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive'
                            $revoked++
                        }
                    }
                    if($revoked -ne 1){throw 'Shipping fixture permission is not unique.'}
                    $auth.Save()
                } finally { $auth.Close($false) }
                $denied=-not [bool](Run 'invSys.Core.xlam' 'TestShippingSession.CanShip')
                $sameSession=[bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn') -and $session -eq [long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion')
                Check 'Shipping.Capability.Revoked.CorePermissionDenied' $denied
                Check 'Shipping.Capability.Revoked.SameSignedInSession' $sameSession
                if(-not $denied -or -not $sameSession){throw 'Shipping capability loss was not isolated from session loss.'}
            }
            foreach($action in @('Add','Update','Remove','Hold','Return','Stage','Send')) {
                $prefix='Shipping.Capability.'+$phase+'.'+$action
                $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @($action))
                if(-not $key){throw 'Shipping capability case lacks genuine selected staging.'}
                $beforeRows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
                $beforeHeld=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
                $before=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount')
                $Other.Activate()
                [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @($action))
                $after=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount')
                if($phase -eq 'Healthy') {
                    Check ($prefix+'.CalibratedOwnerEntry') ($after -eq $before+1)
                } else {
                    Check ($prefix+'.StopsBeforeMutationOwner') ($after -eq $before)
                    $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
                    Check ($prefix+'.VisiblePermissionDenial') ($status -match '(?i)permission|not authorized|capability')
                    if($CaptureEvidence -and $action -eq 'Hold') {
                        CaptureFormEvidence 'Shipping Shipments' 'shipping-permission-denied.png'
                    }
                }
                Check ($prefix+'.StagingAndUnknownValuesPreserved') ($beforeRows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $beforeHeld -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
                Check ($prefix+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($name)))
            }
        }
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSetProbeMode' @($false))
        $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @('Hold'))
        if(-not $key){throw 'Shipping real denied-Hold fixture lacks active staging.'}
        $beforeRows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
        $beforeHeld=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
        $queueBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue'))
        $Other.Activate()
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @('Hold'))
        Check 'Shipping.Capability.RealHold.StagingAndUnknownValuesPreserved' ($beforeRows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $beforeHeld -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
        Check 'Shipping.Capability.RealHold.NoSubmissionEntry' ($queueBefore -eq [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue')))
        Check 'Shipping.Capability.AuthorityBytesPreserved' ($authorityHash -ceq (Get-ShippingActivityHash $inventory))
        Check 'Shipping.Capability.UnrelatedWorkbookPreserved' ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel')
    } finally {
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSetProbeMode' @($false))
        [IO.File]::WriteAllBytes($authPath,$authBytes)
        SelectTarget $Fixture 'config-reader'
    }
}
