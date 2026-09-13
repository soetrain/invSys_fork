# D18 workbook lifetime proof through Excel's actual WorkbookBeforeClose event.
# No event suppression, direct close-handler invocation or authority mutation.
function Test-Slice4beShippingWorkbookClose($Fixture,$Operator,$Other,$Ship,$Hold) {
    # The shared COM fixture opens packages with events off and does not run
    # Auto_Open. Establish the ordinary Shipping startup hook for this case.
    [void](Run 'invSys.Operations.xlam' 'modShippingInit.ShippingPackageAutoOpen')
    $previousEvents=[bool]$excel.EnableEvents
    try {
        SelectTarget $Fixture 'config-reader'
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
        $name=[string]$Operator.Name
        $path=[string]$Operator.FullName
        $beforeRows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
        $beforeHeld=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
        $Operator.Save()
        $savedHash=Get-ShippingActivityHash $path
        $otherHash=Get-ShippingActivityHash $Other.FullName
        $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        $authorityHash=Get-ShippingActivityHash $inventory
        $beforeActivity=@(Get-Slice4beActivityFiles $Fixture)
        $prefix='Shipping.WorkbookClose.'
        $state=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingLifetimeState')
        Check ($prefix+'LiveFormAndTimerBindingEstablished') ($state -ceq 'True|True|1')
        if($state -cne 'True|True|1'){throw 'Shipping close fixture lacks its live form and callback binding.'}
        $excel.EnableEvents=$true
        Check ($prefix+'ExcelEventsEnabled') ([bool]$excel.EnableEvents)
        if(-not $excel.EnableEvents){throw 'Shipping close fixture requires normal Excel events.'}
        $beforeOwner=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Owner'))
        $beforeQueue=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue'))
        $Other.Activate()
        $Operator.Close($false)
        $openNames=@($excel.Workbooks|ForEach-Object Name)
        Check ($prefix+'CapturedWorkbookClosed') ($name -cnotin $openNames)
        Check ($prefix+'FormAndCallbackBindingReleased') ([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingLifetimeState') -ceq 'False|False|0')
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.TriggerShipmentsFormAutoSync')
        Check ($prefix+'LateTimerDoesNotReopenFormOrWorkbook') ([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingLifetimeState') -ceq 'False|False|0' -and $name -cnotin @($excel.Workbooks|ForEach-Object Name))
        Check ($prefix+'NoStagingOrSubmissionOwnerEntry') ($beforeOwner -eq [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Owner')) -and $beforeQueue -eq [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBoundaryCount' @('Queue')))
        Check ($prefix+'NoFabricatedControlActivity') (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -notin $beforeActivity}).Count -eq 0)
        Check ($prefix+'SavedWorkbookBytesPreserved') ($savedHash -ceq (Get-ShippingActivityHash $path))
        Check ($prefix+'AuthorityBytesPreserved') ($authorityHash -ceq (Get-ShippingActivityHash $inventory))
        Check ($prefix+'UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel')
        $reopened=$excel.Workbooks.Open($path,0,$true)
        try {
            $savedShip=Table $reopened 'ShipmentsTally'
            $savedHold=Table $reopened 'NotShipped'
            Check ($prefix+'SavedStagingKeysAndUnknownValuesPreserved') ($beforeRows -ceq (@(Get-ShippingActivityRows $savedShip)|ConvertTo-Json -Depth 5 -Compress) -and $beforeHeld -ceq (@(Get-ShippingActivityRows $savedHold)|ConvertTo-Json -Depth 5 -Compress))
        } finally { $reopened.Close($false) }
    } finally { $excel.EnableEvents=$previousEvents }
}
