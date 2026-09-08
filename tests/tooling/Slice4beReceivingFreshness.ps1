# A True read-model result may describe cached/stale state. Exercise the real
# source resolver; never infer freshness from Boolean success or error text.
function Get-ReceivingFreshnessBusinessRows($Table) {
    @(Get-ReceivingFixtureRows $Table | Select-Object System_Key,QtyOnHand,QtyAvailable,Condition,'Freshness Extra') | ConvertTo-Json -Depth 5 -Compress
}

function Test-ReceivingRefreshFreshness($Fixture) {
    foreach($disposition in @($false,$true)) {
        foreach($mode in @('Cached','Stale')) {
            SelectTarget $Fixture 'config-reader'
            $label='Freshness.'+$(if($disposition){'Returns'}else{'Receipts'})+'.'+$mode
            $operator=$excel.Workbooks.Add()
            $operator.SaveAs((Join-Path $runRoot ($label+'.xlsm')),52)
            $other=$excel.Workbooks.Add()
            $other.Worksheets.Item(1).Cells.Item(1,1).Value2='freshness unrelated sentinel'
            $other.SaveAs((Join-Path $runRoot ($label+'-other.xlsm')),52)
            if (-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name,$disposition))) { throw 'Freshness fixture staging failed.' }
            $inventory=Table $operator 'invSys'; $staging=Table $operator 'ReceivedTally'
            if ($inventory.ListRows.Count -eq 0) { throw 'Freshness fixture has no managed inventory.' }
            $extra=$inventory.ListColumns.Add(); $extra.Name='Freshness Extra'; $extra.DataBodyRange.Value2='preserved local inventory extension'
            $inventoryBefore=Get-ReceivingFreshnessBusinessRows $inventory
            $stagedBefore=@(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress
            $otherHash=Get-ReceivingFixtureHash $other.FullName
            $authorityBefore=Get-ReceivingAuthorityHashes $Fixture
            $source=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Inventory.xlsb')
            $snapshots=@(Get-ChildItem -LiteralPath $Fixture.Root -File -Filter ($Fixture.Warehouse+'*.invSys.Snapshot.Inventory.xls*'))
            if ($snapshots.Count -ne 1 -or -not (Test-Path -LiteralPath $source -PathType Leaf)) { throw 'Freshness fixture requires exactly its canonical source snapshot.' }
            if (@($excel.Workbooks | Where-Object { $_.FullName -eq $source }).Count -ne 0) { throw 'Freshness fixture source is still open.' }
            $held=Join-Path $Fixture.Root ($Fixture.Warehouse+'-fallback.invSys.Snapshot.Inventory.xlsb')
            if ($mode -eq 'Cached') {
                $folder=Join-Path $Fixture.Root 'FreshnessHeld'
                New-Item -ItemType Directory -Path $folder -Force | Out-Null
                $held=Join-Path $folder ([IO.Path]::GetFileName($source))
            }
            foreach($path in @($source,$held)) {
                if (-not [IO.Path]::GetFullPath($path).StartsWith($Fixture.Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Freshness fixture move escaped its root.' }
            }
            if (Test-Path -LiteralPath $held) { throw 'Freshness fixture destination already exists.' }
            $before=@(Get-Slice4beActivityFiles $Fixture)
            Move-Item -LiteralPath $source -Destination $held
            try {
                [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetRefreshFault' @($false))
                $status=[string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.LocalAction' @('Refresh',$other.Name))
                $rows=@(Get-ReceivingFixtureRows $inventory)
                $expectedSource=if($mode -eq 'Cached'){'CACHED'}else{'LOCAL'}
                $staleRows=@($rows | Where-Object { $_.IsStale -eq $true -and $_.SourceType -ceq $expectedSource }).Count
                Check "$label.RealOwnerMarkedStaleState" ($rows.Count -gt 0 -and $staleRows -eq $rows.Count -and [int](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.RefreshCallCount') -eq 1)
                Check "$label.VisibleStaleCauseInsteadOfFreshSuccess" ($status.ToLowerInvariant().Contains($mode.ToLowerInvariant()) -and -not $status.Contains('staging refreshed.'))
                Check "$label.ExactKeysQuantitiesAndUnknownValues" ($inventoryBefore -ceq (Get-ReceivingFreshnessBusinessRows $inventory))
                Check "$label.StagingPreserved" ($stagedBefore -ceq (@(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress))
                Test-ReceivingControlRecords $Fixture $before 'RECEIVING_REFRESH' 'RECEIVING_WORKFLOW' 'RECEIVE_REFRESH_' 'STALE' 1 'Changed' @() $label 'Warning' 4
                Show-ReceivingStagingEvidence $label
                $beforeDirect=@(Get-Slice4beActivityFiles $Fixture)
                $accepted=[bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.DirectRefresh' @($operator.Name))
                Check "$label.BooleanSuccessCanRetainStaleEvidence" ($accepted -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $beforeDirect.Count)
            } finally {
                Move-Item -LiteralPath $held -Destination $source
            }
            Check "$label.AuthorityAndSourceBytesPreserved" (Test-ReceivingAuthorityHashes $Fixture $authorityBefore)
            Check "$label.NoOtherWorkbookRedirection" ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'freshness unrelated sentinel')
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
            $operator.Close($false); $other.Close($false)
        }
    }
}
