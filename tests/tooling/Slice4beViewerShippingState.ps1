# D18: owner-created current state, real Admin publication and Viewer handlers.
# Expected owner values stay in memory; reports contain fixed checks only.
function Test-Slice4beViewerShippingState($Fixture) {
    function PublicationSourceHash([string]$Path) {
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        $sha=[Security.Cryptography.SHA256]::Create()
        try{([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','')}finally{$sha.Dispose();$stream.Dispose()}
    }
    . (Join-Path $PSScriptRoot 'Slice4beShippingPublicationFixture.ps1')
    $shipping=$null
    New-Slice4beShippingPublicationFixture $Fixture ([ref]$shipping)
    try {
        SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Shipping Viewer fixture publication failed.'}
        $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
        if(-not [bool](Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($path,$Fixture.Warehouse))){throw 'Shipping Viewer fixture publication is invalid.'}
        $artifact=[IO.File]::ReadAllText($path)|ConvertFrom-Json
        Test-Slice4beShippingPublication $artifact $shipping
        $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
        $form.AddFromString(@'
Public Function ShippingStateProbeForTest(ByVal action As String, ByVal source As String, Optional ByVal column As Long = 0) As String
    Dim index As Long, first As Long, count As Long, keys As Collection, key As Variant
    If IsEmpty(mRows) Then Exit Function
    For index = 1 To mVisibleIndexes.Count
        If CStr(mRows(CLng(mVisibleIndexes(index)), 18)) = source Then
            count = count + 1
            If first = 0 Then first = index
        End If
    Next index
    If action = "Count" Then ShippingStateProbeForTest = CStr(count): Exit Function
    If first = 0 Then Exit Function
    Select Case action
        Case "Value": ShippingStateProbeForTest = CStr(mLstInventory.List(first - 1, column))
        Case "Select": mLstInventory.ListIndex = first - 1: ShippingStateProbeForTest = "Selected"
        Case "Keys"
            If mDetail Is Nothing Then Exit Function
            Set keys = mDetail.Keys()
            For Each key In keys
                If ShippingStateProbeForTest <> "" Then ShippingStateProbeForTest = ShippingStateProbeForTest & vbLf
                ShippingStateProbeForTest = ShippingStateProbeForTest & CStr(key)
            Next key
    End Select
End Function
'@)
        $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
        $manager.AddFromString(@'
Public Function ShippingStateProbeForTest(ByVal action As String, ByVal source As String, Optional ByVal column As Long = 0) As String
    If Not mInventoryViewer Is Nothing Then ShippingStateProbeForTest = mInventoryViewer.ShippingStateProbeForTest(action, source, column)
End Function
'@)
        function StateProbe([string]$Action,[string]$Source,[int]$Column=0){[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.ShippingStateProbeForTest' @($Action,$Source,$Column))}
        function StateDetail([string]$Caption,[string]$Value){[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadDetailForTest' @($Caption,$Value))}
        function StateAction([string]$Action,[string]$Value=''){[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @($Action,$Value))}
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){$pins[$file.FullName]=PublicationSourceHash $file.FullName}
        foreach($file in $shipping.Pins.Keys){$pins[$file]=PublicationSourceHash $file}
        $published=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
        $reads=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest')
        SelectTarget $Fixture 'config-reader'
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
        [void](StateAction 'Events')
        $box=$shipping.BomRows[0];$hold=$shipping.HoldRows[0]
        Check 'ShippingViewer.OneSummaryPerPackageAlternative' ((StateProbe 'Count' 'ShippingBOM') -ceq '1')
        Check 'ShippingViewer.AcceptedBoxSummaryValues' ((StateProbe 'Value' 'ShippingBOM' 1) -ceq 'BOX_DESIGNED' -and
            (StateProbe 'Value' 'ShippingBOM' 2) -ceq $box.BomVersionLabel -and
            (StateProbe 'Value' 'ShippingBOM' 3) -ceq $box.PackageItem -and
            (StateProbe 'Value' 'ShippingBOM' 4) -ceq '' -and
            (StateProbe 'Value' 'ShippingBOM' 5) -ceq $box.PackageUOM -and
            (StateProbe 'Value' 'ShippingBOM' 6) -ceq $box.PackageLocation)
        Check 'ShippingViewer.BoxSelectedThroughActualListHandler' ((StateProbe 'Select' 'ShippingBOM') -ceq 'Selected')
        $keys=@((StateProbe 'Keys' 'ShippingBOM') -split "`n")
        $expected=@($shipping.BomRows|ForEach-Object {$_.ComponentSystemKey})
        Check 'ShippingViewer.BoxDetailRetainsBothExactComponentLines' ($keys.Count -eq 2 -and @(Compare-Object $expected $keys -CaseSensitive).Count -eq 0)
        Check 'ShippingViewer.BoxIsCurrentStateWithoutEventIdentity' ((StateDetail 'Source classification' 'Current state') -and (StateDetail 'Source event / activity ID' 'Unavailable'))
        [void](StateAction 'Search' ([string]$box.ComponentSystemKey))
        Check 'ShippingViewer.ComponentSearchRetainsPackageSummary' ((StateProbe 'Count' 'ShippingBOM') -ceq '1' -and (StateProbe 'Value' 'ShippingBOM' 3) -ceq $box.PackageItem)
        [void](StateProbe 'Select' 'ShippingBOM')
        $filteredKeys=@((StateProbe 'Keys' 'ShippingBOM') -split "`n")
        Check 'ShippingViewer.FilteredBoxRetainsBothComponentLines' ($filteredKeys.Count -eq 2 -and @(Compare-Object $expected $filteredKeys -CaseSensitive).Count -eq 0)
        [void](StateAction 'Search' '')
        Check 'ShippingViewer.HoldSummaryPreservesOwnerValues' ((StateProbe 'Count' 'ShippingHolds') -ceq '1' -and
            (StateProbe 'Value' 'ShippingHolds' 1) -ceq 'SHIP_HELD' -and
            (StateProbe 'Value' 'ShippingHolds' 2) -ceq $hold.Ref -and
            (StateProbe 'Value' 'ShippingHolds' 3) -ceq $hold.Item -and
            (StateProbe 'Value' 'ShippingHolds' 4) -ceq $hold.Qty -and
            (StateProbe 'Value' 'ShippingHolds' 5) -ceq $hold.UOM -and
            (StateProbe 'Value' 'ShippingHolds' 6) -ceq $hold.Location)
        Check 'ShippingViewer.HoldSelectedThroughActualListHandler' ((StateProbe 'Select' 'ShippingHolds') -ceq 'Selected')
        Check 'ShippingViewer.HoldRetainsExactInventoryKey' ((StateProbe 'Keys' 'ShippingHolds') -ceq $hold.System_Key)
        Check 'ShippingViewer.HoldIsStateNotCompletedShipment' ((StateDetail 'Source classification' 'Current state') -and
            (StateDetail 'Source event / activity ID' 'Unavailable') -and (StateDetail 'Outcome' 'Unavailable') -and
            (StateDetail 'Explanation' 'completed workflow outcomes are unavailable'))
        Check 'ShippingViewer.ReadActionsAvoidAuthorityAndPublication' ($published -eq [long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -and
            $reads -eq [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest'))
        $unchanged=$true
        foreach($file in $pins.Keys){$unchanged=$unchanged -and (Test-Path -LiteralPath $file) -and (PublicationSourceHash $file) -ceq $pins[$file]}
        Check 'ShippingViewer.OwnerProjectionAndLocalSourceBytesUnchanged' $unchanged
    } finally {
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
        Remove-Slice4beShippingPublicationFiles $shipping.LocalFiles
    }
}
