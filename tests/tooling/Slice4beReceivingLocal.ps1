# D18 actual Refresh/Clear handlers. Faults and row inspection stay within
# generated disposable fixtures; reports contain check names and booleans only.
function Install-ReceivingLocalCoreSeam($Bridge) {
    $start=$Bridge.ProcStartLine('RefreshInventoryReadModel',0)
    $count=$Bridge.ProcCountLines('RefreshInventoryReadModel',0)
    $source=$Bridge.Lines($start,$count)
    $changed=$source -replace '(?m)^(\s*)Dim wb As Workbook', ('$1Dim wb As Workbook'+"`r`n"+
        '    TestReceivingActivityGate.RefreshCalls = TestReceivingActivityGate.RefreshCalls + 1'+"`r`n"+
        '    If TestReceivingActivityGate.RefreshFault Then report = "Fixture read-model refresh withheld.": Exit Function')
    if ($changed -eq $source) { throw 'Refresh owner fault seam is unavailable.' }
    $Bridge.DeleteLines($start,$count)
    $Bridge.InsertLines($start,$changed)
}

function Install-ReceivingLocalFormSeams($Form,$Helper) {
    $Form.AddFromString(@'
Public Function ActivityTestLocalAction(ByVal action As String) As String
    Select Case action
        Case "Refresh": mBtnRefresh_Click
        Case "Clear": mBtnClear_Click
        Case "Internal": RefreshAllViews
        Case Else: Err.Raise 5, , "Unsupported fixture action."
    End Select
    ActivityTestLocalAction = CStr(mTxtStatus.Value)
End Function
'@)
    $Helper.AddFromString(@'
Public Function LocalAction(ByVal action As String, ByVal otherName As String) As String
    Application.Workbooks(otherName).Activate
    LocalAction = mForm.ActivityTestLocalAction(action)
End Function
Public Sub DirectLocal(ByVal workbookName As String)
    Dim report As String, ignored As String
    modTS_Received.RefreshReceivingUiForWorkbook Application.Workbooks(workbookName), "LOCAL", report
    ignored = mForm.ActivityTestLocalAction("Internal")
    modTS_Received.ClearReceivingFormStagingForWorkbook Application.Workbooks(workbookName)
End Sub
'@)
}

function Get-ReceivingAuthorityHashes($Fixture) {
    $hashes=@{}
    foreach ($file in Get-ChildItem -LiteralPath $Fixture.Root -File -Recurse -Filter '*.xls*') {
        $hashes[$file.FullName]=Get-ReceivingFixtureHash $file.FullName
    }
    if ($hashes.Count -lt 3) { throw 'Local-action authority fixture is incomplete.' }
    return $hashes
}

function Test-ReceivingAuthorityHashes($Fixture,$Before) {
    $after=Get-ReceivingAuthorityHashes $Fixture
    if ($after.Count -ne $Before.Count) { return $false }
    foreach ($key in $Before.Keys) {
        if (-not $after.ContainsKey($key) -or $after[$key] -cne $Before[$key]) { return $false }
    }
    return $true
}

function Hold-ReceivingLocalActivityStore($Fixture) {
    $leaf=Join-Path $Fixture.Root ('Training\Activity\'+$Fixture.Warehouse)
    $held=$leaf+'-local-fixture-held'
    foreach($path in @($leaf,$held)) {
        if (-not [IO.Path]::GetFullPath($path).StartsWith($Fixture.Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Local-action store fault escaped its fixture.' }
    }
    if (-not (Test-Path -LiteralPath $leaf -PathType Container) -or (Test-Path -LiteralPath $held)) { throw 'Local-action store fault is not ready.' }
    Move-Item -LiteralPath $leaf -Destination $held
    [IO.File]::WriteAllText($leaf,'Blocked disposable local-action activity path')
    return @($leaf,$held)
}

function Test-ReceivingLocalActivity($Fixture) {
    foreach($disposition in @($false,$true)) {
        foreach($action in @('Refresh','Clear')) {
            foreach($mode in @('Normal','Failure','Stale','StoreFailure')) {
                $label='Local.'+$(if($disposition){'Returns'}else{'Receipts'})+'.'+$action+'.'+$mode
                SelectTarget $Fixture 'config-reader'
                $operator=$excel.Workbooks.Add()
                $operator.SaveAs((Join-Path $runRoot ($label+'.xlsm')),52)
                $other=$excel.Workbooks.Add()
                $other.Worksheets.Item(1).Cells.Item(1,1).Value2='unrelated local-action sentinel'
                $other.SaveAs((Join-Path $runRoot ($label+'-other.xlsm')),52)
                $otherHash=Get-ReceivingFixtureHash $other.FullName
                if (-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name,$disposition))) { throw 'Local-action fixture staging failed.' }
                $staging=Table $operator 'ReceivedTally'
                $aggregate=Table $operator 'AggregateReceived'
                $extra=$staging.ListColumns.Add(); $extra.Name='Local Extra'; $extra.DataBodyRange.Value2='preserved staging extension'
                $aggregateExtra=$aggregate.ListColumns.Add(); $aggregateExtra.Name='Local Aggregate Extra'
                $stagedBefore=@(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress
                $aggregateBefore=@(Get-ReceivingFixtureRows $aggregate) | ConvertTo-Json -Depth 5 -Compress
                $before=@(Get-Slice4beActivityFiles $Fixture)
                $held=@()
                [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetRefreshFault' @($action -eq 'Refresh' -and $mode -eq 'Failure'))
                if ($mode -eq 'Stale') {
                    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
                    SelectTarget $Fixture 'config-reader'
                }
                if ($mode -eq 'Failure' -and $action -eq 'Clear') {
                    # Protect only the second table, so the real first deletion
                    # completes before Excel rejects the second owner write.
                    $faultSheet=$operator.Worksheets.Add()
                    $faultSheet.Name='Protected aggregate fixture'
                    [void]$aggregate.Range.Cut($faultSheet.Range('A1'))
                    $aggregate=Table $operator 'AggregateReceived'
                    if ($aggregate.Parent.Name -cne $faultSheet.Name) { throw 'Second-table failure fixture did not move its table.' }
                    if (@($aggregate.ListColumns | Where-Object Name -ceq 'Local Aggregate Extra').Count -ne 1) { throw 'Second-table fixture lost its extension header before the action.' }
                    $aggregate.Parent.Protect()
                }
                if ($mode -eq 'StoreFailure') { $held=Hold-ReceivingLocalActivityStore $Fixture }
                $authorityBefore=Get-ReceivingAuthorityHashes $Fixture
                try {
                    $status=[string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.LocalAction' @($action,$other.Name))
                    $refreshCalls=[int](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.RefreshCallCount')
                }
                finally {
                    if ($mode -eq 'Failure' -and $action -eq 'Clear') { $aggregate.Parent.Unprotect() }
                    [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetRefreshFault' @($false))
                    if ($held.Count -eq 2) {
                        Remove-Item -LiteralPath $held[0]
                        Move-Item -LiteralPath $held[1] -Destination $held[0]
                    }
                }
                $stagedAfter=@(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress
                if ($mode -eq 'Stale') {
                    Check "$label.RejectedBeforeLocalCommand" ($status.Contains('Session or warehouse changed') -and $stagedAfter -ceq $stagedBefore)
                    if ($action -eq 'Refresh') { Check "$label.NoRefreshOwnerCall" ($refreshCalls -eq 0) }
                    Check "$label.NoNewSessionAttribution" (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
                } else {
                    if ($action -eq 'Refresh') {
                        Check "$label.RefreshOwnerCalledOnce" ($refreshCalls -eq 1)
                        Check "$label.StagingAndUnknownValuesPreserved" ($stagedAfter -ceq $stagedBefore)
                        if ($mode -eq 'Failure') {
                            Check "$label.OwnerFailureVisible" ($status.Contains('Fixture read-model refresh withheld.') -and -not $status.Contains('staging refreshed.'))
                        } else { Check "$label.VisibleCompletion" $status.Contains('staging refreshed.') }
                    } elseif ($mode -eq 'Failure') {
                        $aggregateAfter=@(Get-ReceivingFixtureRows $aggregate) | ConvertTo-Json -Depth 5 -Compress
                        Check "$label.PartialChangeAndCauseVisible" ($staging.ListRows.Count -eq 0 -and $aggregateAfter -ceq $aggregateBefore -and $status.Contains('Clear failed:'))
                    } else {
                        Check "$label.LocalStagingCleared" ($staging.ListRows.Count -eq 0 -and $aggregate.ListRows.Count -eq 0 -and $status.Contains('staging cleared.'))
                    }
                    if ($mode -eq 'StoreFailure') {
                        Check "$label.TrackingFailureVisibleWithoutInventedRecords" ($status.Contains('Tracking unavailable') -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
                    } else {
                        $owner=if($action -eq 'Clear'){'RECEIVING_STAGING'}else{'RECEIVING_WORKFLOW'}
                        $outcome=if($mode -eq 'Failure'){'FAILED'}elseif($action -eq 'Clear'){'CLEARED'}else{'REFRESHED'}
                        $effect=if($mode -eq 'Failure'){'Unknown'}else{'Changed'}
                        $severity=if($mode -eq 'Failure'){'Error'}else{'Info'}
                        Test-ReceivingControlRecords $Fixture $before ('RECEIVING_'+$action.ToUpperInvariant()) $owner ('RECEIVE_'+$action.ToUpperInvariant()+'_') $outcome 1 $effect @() $label $severity 4
                    }
                }
                Check "$label.AuthorityBytesPreserved" (Test-ReceivingAuthorityHashes $Fixture $authorityBefore)
                Check "$label.StagingHeaderPreserved" (@($staging.ListColumns | Where-Object Name -ceq 'Local Extra').Count -eq 1)
                Check "$label.AggregateHeaderPreserved" (@($aggregate.ListColumns | Where-Object Name -ceq 'Local Aggregate Extra').Count -eq 1)
                Check "$label.OtherWorkbookPreserved" ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'unrelated local-action sentinel')
                Show-ReceivingStagingEvidence $label
                if ($mode -eq 'Normal') {
                    if ($action -eq 'Clear') {
                        $beforeEmpty=@(Get-Slice4beActivityFiles $Fixture)
                        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.LocalAction' @('Clear',$other.Name))
                        Test-ReceivingControlRecords $Fixture $beforeEmpty 'RECEIVING_CLEAR' 'RECEIVING_STAGING' 'RECEIVE_CLEAR_' 'EMPTY' 1 'Unchanged' @() ($label+'.Empty') 'Info' 4
                    }
                    $beforeDirect=@(Get-Slice4beActivityFiles $Fixture)
                    [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.DirectLocal' @($operator.Name))
                    Check "$label.DirectAndInternalCallsAreNotClicks" ($staging.ListRows.Count -eq 0 -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $beforeDirect.Count)
                }
                [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
                $operator.Close($false); $other.Close($false)
            }
        }
    }
    Test-ReceivingLocalClosedWorkbook $Fixture
}

function Test-ReceivingLocalClosedWorkbook($Fixture) {
    foreach($action in @('Refresh','Clear')) {
        SelectTarget $Fixture 'config-reader'
        $operator=$excel.Workbooks.Add()
        $operator.SaveAs((Join-Path $runRoot ('local-closed-'+$action+'.xlsm')),52)
        $other=$excel.Workbooks.Add()
        $other.Worksheets.Item(1).Cells.Item(1,1).Value2='closed-binding sentinel'
        $other.SaveAs((Join-Path $runRoot ('local-closed-other-'+$action+'.xlsm')),52)
        if (-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name))) { throw 'Closed local-action fixture staging failed.' }
        $operator.Close($false)
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $otherHash=Get-ReceivingFixtureHash $other.FullName
        $authorityBefore=Get-ReceivingAuthorityHashes $Fixture
        [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetRefreshFault' @($false))
        $status=[string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.LocalAction' @($action,$other.Name))
        Check "Local.Closed.$action.VisibleBindingRejection" $status.Contains('Receiving workbook is no longer open')
        Check "Local.Closed.$action.NoActivityOrRefreshOwnerCall" (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count -and [int](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.RefreshCallCount') -eq 0)
        Check "Local.Closed.$action.NoOtherWorkbookRedirection" ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'closed-binding sentinel' -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
        Check "Local.Closed.$action.AuthorityBytesPreserved" (Test-ReceivingAuthorityHashes $Fixture $authorityBefore)
        Show-ReceivingStagingEvidence ('local-closed-'+$action)
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
        $other.Close($false)
    }
}

function Test-ReceivingLocalOlderPolicy($Fixture,[string]$RecordId,[string]$Context) {
    $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
    $meta=Table $cfg 'tblEventTrackingPolicies'; $rows=Table $cfg 'tblEventTrackingControls'
    $meta.ListColumns.Item('CatalogVersion').DataBodyRange.Cells.Item(1,1).Value2=3.0
    foreach($control in @('RECEIVING_CONFIRM_WRITES','RECEIVING_ADD_SELECTED','DISPOSITION_ADD_SELECTED','DISPOSITION_CONFIRM')) {
        $row=$rows.ListRows.Add()
        foreach($field in @('Collect','Visible','SequenceEligible')) { $row.Range.Cells.Item(1,$rows.ListColumns.Item($field).Index).Value2=$true }
        $row.Range.Cells.Item(1,$rows.ListColumns.Item('PolicyVersion').Index).Value2=1.0
        $row.Range.Cells.Item(1,$rows.ListColumns.Item('ControlId').Index).Value2=$control
    }
    $cfg.Save(); $cfg.Close($false)
    try {
        Check 'Local.Policy.CatalogThreeRemainsValid' ((Get-ActivityRead $RecordId).StartsWith('OK|'))
        foreach($control in @('RECEIVING_REFRESH','RECEIVING_CLEAR')) {
            $count=@(Get-Slice4beActivityFiles $Fixture).Count
            $id=[string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @($control,$Context))
            Check ('Local.Policy.OlderCatalogDoesNotEnable.'+$control) ($id -eq '' -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $count)
        }
    } finally {
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $rows=Table $cfg 'tblEventTrackingControls'
        while($rows.ListRows.Count -gt 2) { $rows.ListRows.Item($rows.ListRows.Count).Delete() }
        (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Cells.Item(1,1).Value2=1.0
        $cfg.Save(); $cfg.Close($false)
    }
}
