# D18: allocation alone is not an Inventory submission. Real Add still owns the
# action; the Core probe refuses only after allocation and before either writer.
function Test-Slice4beShippingPrewriteRefusal($Fixture,$Operator,$Other,$Ship,$Hold){
    $label='Shipping.Submission.PrewriteRefusal'
    [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionMode' @(0))
    if(@(Get-ShippingActivityRows $Ship|Where-Object {[double]$_.QUANTITY -gt 0}).Count){
        $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @('Remove'))
        if(-not $key){throw 'Prewrite refusal cleanup selection unavailable.'}
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @('Remove'))
    }
    $empty=@(Get-ShippingActivityRows $Ship|Where-Object {[double]$_.QUANTITY -gt 0}).Count -eq 0
    Check ($label+'.EmptyStagingEstablishedThroughOwner') $empty
    if(-not $empty){throw 'Prewrite refusal requires empty actual staging.'}
    $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @('Add'))
    if(-not $key){throw 'Prewrite refusal Add selection unavailable.'}
    $paths=@(Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object Extension -EQ '.xlsb'|ForEach-Object FullName)
    $localRoot=Join-Path $runRoot 'shipping-submission-staging'
    if(Test-Path -LiteralPath $localRoot){$paths+=@(Get-ChildItem -LiteralPath $localRoot -Recurse -File|ForEach-Object FullName)}
    $hashes=@{};foreach($path in $paths){$hashes[$path]=Get-ShippingActivityHash $path}
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $before=@(Get-Slice4beActivityFiles $Fixture)
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSubmissionResetOwner')
    [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionMode' @(5))
    try{
        $Other.Activate()
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @('Add'))
        $entries=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingWriteEntryState')
        $state=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionState')).Split('|')
        $id=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionId')
        $owner=([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSubmissionOwnerState' @($id))).Split('|')
        $calibrated=$entries -ceq '2|0|0' -and $id -ne '' -and $state[4] -ceq 'True'
        Check ($label+'.BothRefusedAfterAllocationBeforeWrites') $calibrated
        if(-not $calibrated){throw 'Allocation/write boundary calibration failed; not a product RED.'}
        Check ($label+'.BothPublicRoutesEnteredOnce') ($state[0] -ceq '1' -and $state[1] -ceq '1')
        Check ($label+'.NeitherRouteAccepted') ($state[2] -ceq 'False' -and $state[3] -ceq 'False')
        Check ($label+'.OwnerReportsFailureAndRetainsAllocatedId') ($owner[0] -ceq 'True' -and $owner[1] -ceq 'False' -and $owner[2] -ceq 'True')
        Test-ShippingActivityPair $Fixture $before 'Add' $label @() 'SHIPPING_ADD' 'FAILED'
        $preserved=$true;foreach($path in $hashes.Keys){$preserved=$preserved -and (Test-Path -LiteralPath $path) -and $hashes[$path] -ceq (Get-ShippingActivityHash $path)}
        $afterPaths=@(Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object Extension -EQ '.xlsb'|ForEach-Object FullName)
        if(Test-Path -LiteralPath $localRoot){$afterPaths+=@(Get-ChildItem -LiteralPath $localRoot -Recurse -File|ForEach-Object FullName)}
        Check ($label+'.AuthorityAndSubmissionFilesUnchanged') ($preserved -and $afterPaths.Count -eq $paths.Count -and @($afterPaths|Where-Object {$_ -notin $paths}).Count -eq 0)
        Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)))
        Check ($label+'.UnknownHeadersPreserved') ($null -ne $Ship.ListColumns.Item('Shipping Extra') -and $null -ne $Hold.ListColumns.Item('Shipping Extra'))
        Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName))
    }finally{[void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionMode' @(0))}
}
