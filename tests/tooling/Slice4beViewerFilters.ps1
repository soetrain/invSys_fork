# D18: actual selector events over a synthetic declared wire; real reader checks
# precede this extension. No owner events or authority rows are fabricated.
function Test-Slice4beViewerFilters($Fixture) {
    SelectTarget $Fixture 'config-reader'
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modInventoryViewerData').CodeModule
    $core.InsertLines($core.CountOfDeclarationLines+1,'Private FilterReadsForTest As Long')
    $core.InsertLines($core.ProcBodyLine('LoadCurrentInventoryEventViewerData',0)+1,'    FilterReadsForTest = FilterReadsForTest + 1')
    $core.AddFromString(@'
Public Function FilterReadCountForTest() As Long
    FilterReadCountForTest = FilterReadsForTest
End Function
'@)
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    $form.AddFromString(@'
Public Function ChooseEventFilterForTest(ByVal name As String, ByVal caption As String) As Boolean
    Dim control As Object, index As Long
    On Error GoTo Missing
    Set control = Me.Controls(name)
    For index = 0 To control.ListCount - 1
        If CStr(control.List(index, 0)) = caption Then
            control.ListIndex = index
            ChooseEventFilterForTest = True: Exit Function
        End If
    Next index
Missing:
End Function
Public Function EventFilterFactForTest(ByVal action As String, Optional ByVal expected As String = "") As Boolean
    Dim index As Long, control As Object, name As Variant, lastRight As Single
    On Error GoTo Missing
    Select Case action
        Case "Rows": EventFilterFactForTest = (CStr(mLstInventory.ListCount) = expected)
        Case "StageDate": mCboEventRange.Value = expected: EventFilterFactForTest = True
        Case "CloseDetail": If Not mDetail Is Nothing Then mDetail.CloseDetail
        Case "Count": EventFilterFactForTest = (InStr(1, mLblEventPage.Caption, ". " & expected & " matching /", vbBinaryCompare) > 0)
        Case "FirstPage": EventFilterFactForTest = (Left$(mLblEventPage.Caption, 10) = "Page 1 of ")
        Case "Next": Me.Controls("btnEventsNext").Value = True: EventFilterFactForTest = True
        Case "DetailLines"
            If mLstInventory.ListCount = 0 Then Exit Function
            mLstInventory.ListIndex = 0
            EventFilterFactForTest = (CStr(mDetail.Keys().Count) = expected)
        Case "ReservationLabel"
            For index = 1 To mVisibleIndexes.Count
                If CStr(mRows(CLng(mVisibleIndexes(index)), 13)) = "SHIP_RESERVE" Then
                    EventFilterFactForTest = (CStr(mLstInventory.List(index - 1, 1)) = "Inventory Reserved"): Exit Function
                End If
            Next index
        Case "Fit"
            lastRight = 0
            For Each name In Array("cboEventsView", "cboEventsFamily", "cboEventsSource", "cboEventsOutcome")
                Set control = Me.Controls(CStr(name))
                If Not control.Visible Or control.Left < lastRight Or control.Width < 100 Or control.Top < 0 Then Exit Function
                If control.Left + control.Width > Me.InsideWidth + 1 Or control.Top + control.Height > mLstInventory.Top - 20 Then Exit Function
                lastRight = control.Left + control.Width
            Next name
            EventFilterFactForTest = True
        Case "InventoryHidden"
            mTabs.Value = 0
            For Each name In Array("cboEventsView", "cboEventsFamily", "cboEventsSource", "cboEventsOutcome")
                If Me.Controls(CStr(name)).Visible Then Exit Function
            Next name
            EventFilterFactForTest = True
        Case "InventoryListSpace"
            EventFilterFactForTest = (Abs(mLstInventory.Top + mLstInventory.Height - (mBtnClose.Top - 12)) <= 1)
    End Select
Missing:
End Function
'@)
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function ChooseEventFilterForTest(ByVal name As String, ByVal caption As String) As Boolean
    If Not mInventoryViewer Is Nothing Then ChooseEventFilterForTest = mInventoryViewer.ChooseEventFilterForTest(name, caption)
End Function
Public Function EventFilterFactForTest(ByVal action As String, Optional ByVal expected As String = "") As Boolean
    If Not mInventoryViewer Is Nothing Then EventFilterFactForTest = mInventoryViewer.EventFilterFactForTest(action, expected)
End Function
'@)
    function Choose([string]$Name,[string]$Caption){[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.ChooseEventFilterForTest' @(('cboEvents'+$Name),$Caption))}
    function Fact([string]$Action,[string]$Expected=''){[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.EventFilterFactForTest' @($Action,$Expected))}
    function Act([string]$Action,[string]$Expected=''){[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @($Action,$Expected))}
    $wire=[string](Run 'invSys.Core.xlam' 'modInventoryViewerData.LoadCurrentInventoryEventViewerData')
    $header=($wire -split "`r`n")[0] -split "`t"
    if($header.Count -ne 12 -or $header[4] -cne 'EVENTS1'){throw 'Filter fixture requires validated EVENTS1.'}
    $ids=$header[9] -split ','
    $rows=[Collections.Generic.List[string]]::new()
    $records=@(foreach($n in 1..105){,@('Inventory',('FILTER-RECEIVE-{0:D3}' -f $n),'Business event','Receiving','RECEIVE','')})
    $records+=,@('Inventory','FILTER-RESERVE','Business event','Shipping','SHIP_RESERVE','')
    $records+=,@('Designs','FILTER-DESIGN','Business event','Designs','PROCESS_SAVE','')
    $records+=,@('Activity','FILTER-ACTIVITY','User activity','Admin','CONFIG_SAVE_REQUESTED','REQUESTED')
    $records+=,@('Activity','FILTER-ACTIVITY','User activity','Admin','CONFIG_SAVE_COMPLETED','COMPLETED')
    $records+=,@('ShippingHolds','','Current state','Shipping','SHIP_HELD','')
    foreach($record in $records){
        $values=[string[]]::new(18+$ids.Count);for($i=0;$i -lt $values.Length;$i++){$values[$i]=''}
        $values[0]=$header[2].Replace('T',' ').Replace('Z',' UTC');$values[1]=if($record[4] -ceq 'SHIP_RESERVE'){'Inventory Reserved'}else{$record[4]}
        if($record[1] -ceq 'FILTER-RECEIVE-105'){$values[0]=[DateTime]::UtcNow.AddDays(-10).ToString('yyyy-MM-dd HH:mm:ss')+' UTC'}
        $values[2]=$record[1];$values[10]=$record[1];$values[12]=$record[4];$values[17]=$record[0]
        $fields=@{SourceId=$record[1];SourceKind=$record[2];EventFamily=$record[3];EventCode=$record[4];EventType=$values[1];Outcome=$record[5];WarehouseId=$Fixture.Warehouse;TimeProvenance='Verified UTC'}
        foreach($field in $fields.Keys){$column=[array]::IndexOf($ids,$field);if($column -lt 0){throw 'Filter detail field missing'};$values[18+$column]=$fields[$field]}
        $rows.Add(($values -join "`t"))
    }
    if($rows.Count -ne 110){throw 'Filter fixture row count is invalid'}
    $header[3]=[string]$rows.Count;$header[6]=[guid]::NewGuid().ToString();$header[8]='Synthetic filter fixture; no business effects.'
    $payload=($header -join "`t")+"`r`n"+($rows -join "`r`n")
    try {
        [void](Run 'invSys.Core.xlam' 'modInventoryViewerData.SetPublishedOrderingForTest' @($payload))
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
        [void](Act 'Events')
        Check 'EventFilters.LoadedFixtureHasCompleteDefaultGroups' ((Fact 'Count' '108') -and (Fact 'Rows' '100'))
        $readCount=[long](Run 'invSys.Core.xlam' 'modInventoryViewerData.FilterReadCountForTest')
        Check 'EventFilters.ViewSelectorInvokesActualChange' (Choose 'View' 'All published events')
        Check 'EventFilters.AllPublishedIncludesReservation' (Fact 'Count' '109')
        Check 'EventFilters.FamilySelectorFiltersWholeProjection' ((Choose 'Family' 'Shipping') -and (Fact 'Count' '2'))
        Check 'EventFilters.ReservationKeepsHonestLabel' (Fact 'ReservationLabel')
        Check 'EventFilters.SourceCombinesWithFamily' ((Choose 'Source' 'Inventory') -and (Fact 'Count' '1'))
        Check 'EventFilters.OperatorActionsHidesInternalReserve' ((Choose 'View' 'Operator actions') -and (Fact 'Rows' '0'))
        [void](Choose 'View' 'All published events');[void](Choose 'Family' 'All families');[void](Choose 'Source' 'All sources')
        Check 'EventFilters.RecordedOutcomeMatchesResultLine' ((Choose 'Outcome' 'Completed') -and (Fact 'Count' '1'))
        Check 'EventFilters.OutcomeSelectionRetainsAttemptAndResult' (Fact 'DetailLines' '2')
        Check 'EventFilters.ThreeCriteriaCombineWithoutInventedMatch' ((Choose 'Family' 'Receiving') -and (Choose 'Source' 'Inventory') -and (Fact 'Rows' '0'))
        [void](Choose 'Family' 'All families');[void](Choose 'Source' 'All sources')
        Check 'EventFilters.UnavailableDoesNotClaimOutcome' ((Choose 'Outcome' 'Unavailable') -and (Fact 'Count' '108'))
        [void](Choose 'Outcome' 'All outcomes')
        Check 'EventFilters.FamilyBeforePagingFindsAll105' ((Choose 'Family' 'Receiving') -and (Fact 'Count' '105') -and (Fact 'Rows' '100'))
        [void](Fact 'Next')
        Check 'EventFilters.SecondFilteredPageRetainsFive' (Fact 'Rows' '5')
        Check 'EventFilters.SourceChangeStartsFirstPage' ((Choose 'Source' 'Inventory') -and (Fact 'FirstPage') -and (Fact 'Rows' '100'))
        [void](Act 'Search' 'FILTER-RECEIVE-105')
        Check 'EventFilters.SearchCombinesWithSelectors' ((Fact 'Count' '1') -and (Fact 'Rows' '1'))
        [void](Act 'Search' '')
        Check 'EventFilters.SelectorsNeverReadProjection' ($readCount -eq [long](Run 'invSys.Core.xlam' 'modInventoryViewerData.FilterReadCountForTest'))
        [void](Fact 'StageDate' 'Day')
        Check 'EventFilters.PendingDateWaitsForRefresh' ((Choose 'Family' 'All families') -and (Fact 'Count' '106'))
        [void](Act 'Refresh')
        Check 'EventFilters.RefreshAppliesPendingDay' (Fact 'Count' '105')
        [void](Fact 'StageDate' 'All');[void](Act 'Refresh')
        Check 'EventFilters.RefreshRestoresAllDates' (Fact 'Count' '106')
        foreach($size in @('FitMinimum','FitDefault','FitLarger','FitDefault')){
            [void](Act $size)
            $suffix=if($size -eq 'FitDefault' -and @($results.Check) -contains 'EventFilters.Layout.FitDefault'){'.Restored'}else{''}
            Check ('EventFilters.Layout.'+$size+$suffix) (Fact 'Fit')
        }
        if($CaptureEvidence){
            [void](Fact 'CloseDetail')
            CaptureFormEvidence '' 'viewer-event-filters.png' ([long](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadWindowForTest'))
            Check 'EventFilters.VisibleCapture' $true
        }
        Check 'EventFilters.InventoryHidesSelectors' (Fact 'InventoryHidden')
        foreach($size in @('FitMinimum','FitDefault','FitLarger','FitDefault')){
            [void](Act $size)
            $suffix=if($size -eq 'FitDefault' -and @($results.Check) -contains 'EventFilters.InventoryListSpace.FitDefault'){'.Restored'}else{''}
            Check ('EventFilters.InventoryListSpace.'+$size+$suffix) (Fact 'InventoryListSpace')
        }
        [void](Act 'Events')
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        [void](Choose 'Family' 'Shipping')
        Check 'EventFilters.ContextChangeInvalidatesLoadedContent' (Act 'Empty')
    } finally {
        [void](Run 'invSys.Core.xlam' 'modInventoryViewerData.SetPublishedOrderingForTest' @(''))
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
    }
}
