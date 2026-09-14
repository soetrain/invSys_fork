param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path $RepoRoot).Path

function Read-Source([string]$relativePath) {
    Get-Content -Raw -LiteralPath (Join-Path $repo $relativePath)
}

function Procedure-Text([string]$text, [string]$name) {
    [regex]::Match(
        $text,
        "(?ms)^(?:Public|Private) (?:Function|Sub) $([regex]::Escape($name))\b.*?^End (?:Function|Sub)"
    ).Value
}

$connectionForm = Read-Source "src/Core/Forms/frmWarehouseConnection.frm"
$roleWriter = Read-Source "src/Core/Modules/modRoleEventWriter.bas"
$warehouseSync = Read-Source "src/Core/Modules/modWarehouseSync.bas"
$viewerData = Read-Source "src/Core/Modules/modInventoryViewerData.bas"
$publishedReader = Read-Source "src/Core/Modules/modPublishedEventsReader.bas"
$eventsPublisher = Read-Source "src/Core/ClassModules/cEventsPublication.cls"
$viewerForm = Read-Source "src/Operations/Forms/frmInventoryViewer.frm"
$operationsAnchors = Read-Source "src/Operations/ClassModules/cOperationsAnchorManager.cls"
$viewerController = Read-Source "src/Operations/Modules/modInventoryViewer.bas"
$receivingForm = Read-Source "src/Receiving/Forms/frmReceiving.frm"
$receivingActivity = Read-Source "src/Receiving/Modules/modReceivingActivityAction.bas"
$shippingForm = Read-Source "src/Shipping/Forms/frmShipmentsTally.frm"
$shippingService = Read-Source "src/Shipping/Modules/modTS_Shipments.bas"

$connectClick = Procedure-Text $connectionForm "mBtnConnect_Click"
$connectAction = Procedure-Text $roleWriter "ConnectWarehouseStorageForCapability"
$receivingConfirm = Procedure-Text $receivingForm "mBtnConfirm_Click"
$receivingAggregateClick = Procedure-Text $receivingForm "mLstAggregate_Click"
$receivingNavigation = Procedure-Text $receivingForm "NavigationSelection"
$receivingConfirmEntry = Procedure-Text $receivingActivity "ConfirmWrites"
$receivingConfirmOwner = Procedure-Text $receivingActivity "ConfirmAction"
$viewerBuild = Procedure-Text $viewerForm "BuildLayout"
$viewerTab = Procedure-Text $viewerForm "ApplyViewerTab"
$viewerRefreshEvents = Procedure-Text $viewerForm "RefreshEvents"
$viewerEventsTest = Procedure-Text $viewerForm "TestEventsReport"
$viewerRenderRows = Procedure-Text $viewerForm "RenderRows"
$viewerEventsAction = Procedure-Text $viewerController "RunInventoryViewerEventsForTest"
$shippingCommit = Procedure-Text $shippingForm "CommitCurrentLine"
$shippingSend = Procedure-Text $shippingForm "mBtnSend_Click"
$snapshotAction = Procedure-Text $warehouseSync "GenerateWarehouseSnapshot"

$checks = @(
    [pscustomobject]@{
        Check = "ServerConnection.ProgressBeforeBlockingIO"
        Passed = ($connectClick -match 'Connecting to warehouse storage') -and
            ($connectClick -match 'Me\.Repaint') -and ($connectClick -match 'DoEvents') -and
            ($connectAction -match 'BeginServerConnectionProgressRole') -and
            ($connectAction -match 'EndServerConnectionProgressRole')
        Contract = "Manual and ribbon Server Sign In render progress before the synchronous Windows SMB call and restore Excel UI afterward."
    },
    [pscustomobject]@{
        Check = "Receiving.AggregateReferenceDetail"
        Passed = ($receivingForm -match 'txtAggregateReferences') -and
            ($receivingForm -match 'Selected references') -and
            ($receivingForm -match 'mTxtAggregateReferences\.MultiLine\s*=\s*True') -and
            ($receivingAggregateClick -match 'NavigationSelection\s+"lstAggregate"') -and
            ($receivingNavigation -match 'Case\s+"lstAggregate"\s*:\s*ShowSelectedAggregateReferences') -and
            ($receivingForm -match 'ClearAggregateReferenceDetail')
        Contract = "The fixed-height aggregate list retains one-line rows while a dedicated multiline detail surface shows every concatenated reference and clears with staging."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.ReadOnlyTab"
        Passed = ($viewerBuild -match 'Tabs\(0\)\.Caption\s*=\s*"Inventory"') -and
            ($viewerBuild -match 'Tabs\(1\)\.Caption\s*=\s*"Events"') -and
            ($viewerTab -match 'RefreshEvents') -and
            ($viewerRefreshEvents -match 'modInventoryViewer\.LoadInventoryViewerEvents') -and
            ($viewerController -match 'modInventoryViewerData\.LoadCurrentInventoryEventViewerData') -and
            ($viewerController -notmatch 'LoadShippingViewerSupplementEvents') -and
            ($publishedReader -match 'Case\s+"ShippingBOM"') -and
            ($publishedReader -match 'Case\s+"ShippingHolds"')
        Contract = "Viewer reads published event observations and labelled current Box Design/Held Shipment state; opening or refreshing it never reads Shipping authority supplements."
    },
    [pscustomobject]@{
        Check = "Viewer.Tabs.InventoryEventsAndListBoxTable"
        Passed = ($viewerBuild -notmatch 'Tabs\.Add\s+"tabInventory"') -and
            ($viewerBuild -notmatch 'Tabs\.Add\s+"tabEvents"') -and
            ($viewerBuild -match 'Tabs\.Add\s+"tabListBoxTable"') -and
            ($viewerBuild -match 'ListBox->Table') -and
            ($viewerEventsTest -match 'TabCount=') -and
            ($viewerEventsTest -match 'TabCaptions=') -and
            ($viewerEventsTest -match 'SelectedTab=')
        Contract = "The runtime Viewer retains one Inventory and one Events page, adds the approved ListBox->Table page, and its public Events action selects the operator-visible Events tab."
    },
    [pscustomobject]@{
        Check = "Viewer.Layout.GuardsNativeWindowState"
        Passed = ($operationsAnchors -match 'GetUserFormWindowHandle') -and
            ($operationsAnchors -match 'IsIconic') -and ($operationsAnchors -match 'IsZoomed') -and
            ($operationsAnchors -match 'ApplyMinimumFormSize') -and
            ($operationsAnchors -match 'Err\.Number\s*=\s*384')
        Contract = "Operations anchoring skips native form-size enforcement while minimized or maximized and contains residual run-time error 384 without disabling restored-state layout."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.ReadableTimestampRefresh"
        Passed = ($eventsPublisher -match 'Format\$\(CDate\(value\),\s*"yyyy-mm-dd\\Thh:nn:ss"\)') -and
            ($publishedReader -match 'values\(0\)\s*=\s*DisplayTime') -and
            ($publishedReader -match '"Z",\s*" UTC"') -and
            ($viewerEventsTest -match 'ReadableDates=') -and
            ($viewerEventsTest -match 'FirstReference=')
        Contract = "Events renders readable timestamps and the public Events refresh reports the newly published first event rather than retaining stale rows."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.RollingDateFilters"
        Passed = ($viewerForm -match 'mCboEventRange') -and
            ($viewerBuild -match '\.AddItem\s+"All"') -and
            ($viewerBuild -match '\.AddItem\s+"Day"') -and
            ($viewerBuild -match '\.AddItem\s+"Week"') -and
            ($viewerBuild -match '\.AddItem\s+"Month"') -and
            ($viewerRenderRows -match 'Case\s+"DAY"') -and
            ($viewerRenderRows -match 'Case\s+"WEEK"') -and
            ($viewerRenderRows -match 'Case\s+"MONTH"') -and
            ($viewerRenderRows -match 'IsNumeric') -and
            ($viewerRenderRows -match 'DateAdd\("d"') -and
            ($viewerEventsTest -match 'rangeText') -and
            ($viewerEventsTest -match 'mBtnRefresh_Click') -and
            ($viewerEventsAction -match 'rangeText')
        Contract = "The operator-visible Events Refresh action combines text search with All, rolling Day/Week/Month, or a typed positive whole-number-of-days filter; Inventory remains unfiltered."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.RemembersDateFilter"
        Passed = ($viewerForm -match 'SETTINGS_EVENT_RANGE') -and
            ($viewerBuild -match 'GetSetting\(') -and
            ($viewerForm -match 'InventoryViewerEventRange') -and
            ($viewerRenderRows -match 'SaveSetting\s+SETTINGS_APP') -and
            ($viewerEventsTest -match 'EventRange=') -and
            ($viewerController -match 'CloseInventoryViewerForTest')
        Contract = "A valid applied Event range is stored as a per-Windows-user Operations preference and restored when the public Viewer action creates a new form instance."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.PublishedProjection"
        Passed = ($snapshotAction -match 'WriteSnapshotEventRows') -and
            ($warehouseSync -match 'tblInventoryEvents') -and
            ($viewerData -match 'LoadCurrentInventoryEventViewerData\s*=\s*modPublishedEventsReader\.ReadCurrent') -and
            ($publishedReader -match 'modEventsPublicationStore\.Read') -and
            ($publishedReader -match '\.invSys\.Snapshot\.Events\.json')
        Contract = "Viewer event history is read from the published snapshot projection rather than making the form a canonical writer or authority."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.RemoveRelease"
        Passed = ($publishedReader -match 'Case\s+"SHIP_RELEASE"\s*:\s*FriendlyType\s*=\s*"Remove"') -and
            ($shippingService -match 'EVENT_TYPE_SHIP_RELEASE')
        Contract = "Shipping Remove releases locked inventory through SHIP_RELEASE and the operator-facing Events view labels that event Remove."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.ExcludesInternalReservation"
        Passed = ($publishedReader -match 'Case\s+"SHIP_RESERVE"\s*:\s*FriendlyType\s*=\s*"Inventory Reserved"') -and
            ($publishedReader -notmatch 'Case\s+"SHIP_RESERVE"\s*:\s*FriendlyType\s*=\s*"Shipment Held"') -and
            ($viewerEventsTest -match 'ShipmentHeldRows=')
        Contract = "An ordinary Shipping Add may write an internal SHIP_RESERVE row, but the operator-facing Events view does not misreport that zero-delta reservation as Shipment Held; actual held-shipment supplements remain visible."
    },
    [pscustomobject]@{
        Check = "OperatorPersistence.PendingStatus"
        Passed = ($receivingConfirm -match 'ShowPersistencePending') -and
            ($receivingConfirm -match 'modReceivingActivityAction\.ConfirmWrites') -and
            ($receivingConfirmEntry -match 'ConfirmAction\(') -and
            ($receivingConfirmOwner -match 'modReceivingPostingService\.ExecuteConfirmWrites') -and
            ($shippingCommit -match 'ShowPersistencePending') -and
            ($shippingSend -match 'ShowPersistencePending') -and
            ($shippingForm -match 'Me\.Repaint') -and ($shippingForm -match 'DoEvents')
        Contract = "Receiving/Returns and Shipping render their own saving-to-server status before required persistence begins; Office-native progress UI remains separate."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests/integration/plan022_slice4w_operator_responsiveness_and_events_results.md"
$lines = @(
    "# Plan 022 Slice 4w Operator Responsiveness and Events Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Contract |",
    "|---|---|---|"
)
foreach ($check in $checks) {
    $lines += "| $($check.Check) | $(if ($check.Passed) { 'PASS' } else { 'FAIL' }) | $($check.Contract) |"
}
Set-Content -LiteralPath $resultPath -Value $lines -Encoding utf8
$checks | Format-Table -AutoSize
Write-Host "Plan 022 Slice 4w operator responsiveness/events: $passed passed, $failed failed"
if ($failed -gt 0) { exit 1 }
