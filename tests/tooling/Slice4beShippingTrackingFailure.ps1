# Actual Add/Remove through the existing form facade with a physically unavailable
# optional activity directory. Only paths under this generated fixture may move.
function Test-Slice4beShippingTrackingFailure($Fixture,$Operator,$Other,$Ship,$Hold){
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $leaf=Join-Path $Fixture.Root ('Training\Activity\'+$Fixture.Warehouse)
    $held=$leaf+'-tracking-test-held'
    foreach($path in @($leaf,$held)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fixture escaped its generated root.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Tracking hold path already exists.'}
    $history=@{};foreach($file in @(Get-Slice4beActivityFiles $Fixture)){$history[[IO.Path]::GetFileName($file)]=Get-ShippingActivityHash $file}
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $moved=$false
    New-Item -ItemType Directory -Path (Split-Path $leaf -Parent) -Force|Out-Null
    try{
        if(Test-Path -LiteralPath $leaf){Move-Item -LiteralPath $leaf -Destination $held;$moved=$true}
        [IO.File]::WriteAllText($leaf,'blocked fixture activity path')
        foreach($action in @('Add','Remove')){
            $label='Shipping.TrackingFailure.'+$action
            $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @($action))
            if(-not $key){throw 'Tracking fixture selection unavailable.'}
            $ownerBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Owner'))
            $queueBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue'))
            $Other.Activate()
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @($action))
            $rows=@(Get-ShippingActivityRows $Ship|Where-Object {[double]$_.QUANTITY -gt 0})
            $effect=if($action -eq 'Add'){$rows.Count -eq 1 -and $rows[0].System_Key -ceq $key -and [double]$rows[0].QUANTITY -eq 2}else{$rows.Count -eq 0}
            Check ($label+'.OwnerEnteredOnce') (([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Owner')))-$ownerBefore -eq 1)
            Check ($label+'.SubmissionEnteredOnce') (([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue')))-$queueBefore -eq 1)
            Check ($label+'.OwnerStagingResult') $effect
            Check ($label+'.BlockedStoreUnchanged') ((Test-Path -LiteralPath $leaf -PathType Leaf) -and [IO.File]::ReadAllText($leaf) -ceq 'blocked fixture activity path')
            $status=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus')
            Check ($label+'.TrackingUnavailableVisible') ($status.Contains('Tracking unavailable'))
            Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)))
            Check ($label+'.UnknownHeadersPreserved') ($null -ne $Ship.ListColumns.Item('Shipping Extra') -and $null -ne $Hold.ListColumns.Item('Shipping Extra'))
            Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName))
        }
    }finally{
        if(Test-Path -LiteralPath $leaf -PathType Leaf){Remove-Item -LiteralPath $leaf -Force}
        if($moved){Move-Item -LiteralPath $held -Destination $leaf}
    }
    $after=@(Get-Slice4beActivityFiles $Fixture)
    $preserved=$after.Count -eq $history.Count
    foreach($file in $after){$name=[IO.Path]::GetFileName($file);$preserved=$preserved -and $history.ContainsKey($name) -and $history[$name] -ceq (Get-ShippingActivityHash $file)}
    Check 'Shipping.TrackingFailure.PriorActivityBytesPreserved' $preserved
}
