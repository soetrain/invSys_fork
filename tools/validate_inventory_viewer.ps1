[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$DeployRoot = "deploy/current",
    [switch]$RejectFixtureSignInForTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Import-FunctionDefinitions {
    param([string]$ScriptPath)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $ScriptPath, [ref]$tokens, [ref]$errors
    )
    if ($errors.Count -gt 0) { throw "Unable to parse helper source: $($errors[0].Message)" }
    foreach ($definition in $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true)) {
        $scriptDefinition = $definition.Extent.Text -replace (
            '^(?i)function\s+' + [regex]::Escape($definition.Name)
        ), ('function script:' + $definition.Name)
        . ([scriptblock]::Create($scriptDefinition))
    }
}

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deploy = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
Import-FunctionDefinitions -ScriptPath (Join-Path $repo "tools\validate_phase6_live_role_workflows.ps1")

$runtimeRoot = Join-Path ([IO.Path]::GetTempPath()) (
    "invsys-inventory-viewer-" + [guid]::NewGuid().ToString("N")
)
$warehouseId = "WHV" + [guid]::NewGuid().ToString("N").Substring(0, 6).ToUpperInvariant()
$stationId = "S1"
$testUser = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "user1" } else { $env:USERNAME }
$testPin = [guid]::NewGuid().ToString("N")
$testPinHash = Get-InvSysCredentialHash -Credential $testPin
$configPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Config.xlsb")
$authPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Auth.xlsb")
$inventoryPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Data.Inventory.xlsb")
$snapshotPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Snapshot.Inventory.xlsb")
$resultPath = Join-Path $repo "tests\integration\inventory_viewer_results.md"
$excel = $null
$opened = New-Object System.Collections.Generic.List[object]
$facts = [ordered]@{}
$passed = $false
$detail = ""
$step = "startup"
$preferenceRegistryPath = "HKCU:\Software\VB and VBA Program Settings\invSys\Operations"
$preferenceName = "InventoryViewerEventRange"
$preferenceExisted = $false
$preferenceBefore = ""
$preferenceItem = Get-ItemProperty -LiteralPath $preferenceRegistryPath `
    -Name $preferenceName -ErrorAction SilentlyContinue
if ($null -ne $preferenceItem) {
    $preferenceExisted = $true
    $preferenceBefore = [string]$preferenceItem.PSObject.Properties[$preferenceName].Value
}

try {
    New-Item -Path $preferenceRegistryPath -Force | Out-Null
    Set-ItemProperty -LiteralPath $preferenceRegistryPath -Name $preferenceName `
        -Value "invalid-test-range" -Type String
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
    $step = "start isolated Excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    $step = "create isolated runtime"
    $configWb = New-ConfigWorkbook -Excel $excel -Path $configPath `
        -WarehouseId $warehouseId -StationId $stationId -RuntimeRoot $runtimeRoot
    $authWb = New-AuthWorkbook -Excel $excel -Path $authPath `
        -WarehouseId $warehouseId -StationId $stationId `
        -CurrentUserIds @($testUser) -CredentialHash $testPinHash
    $inventoryWb = New-InventoryWorkbook -Excel $excel -Path $inventoryPath `
        -WarehouseId $warehouseId -SkuRows @("SKU-SHIP", "SKU-SUGAR", "SKU-COMP")
    $inventoryLog = $inventoryWb.Worksheets.Item("InventoryLog").ListObjects.Item("tblInventoryLog")
    $fixtureNow = [DateTime]::Now.AddMinutes(-2)
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-RECEIVE"
        "AppliedSeq" = 4
        "EventType" = "RECEIVE"
        "OccurredAtUTC" = $fixtureNow.AddDays(-40).ToOADate()
        "AppliedAtUTC" = $fixtureNow.AddDays(-40).ToOADate()
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-VIEWER-RECEIVE"
        "SKU" = "SKU-SHIP"
        "QtyDelta" = 2
        "Location" = "DOCK"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=BOL-VIEWER;Item=Viewer Shipment Item;UOM=each"
    }
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-RETURN"
        "AppliedSeq" = 5
        "EventType" = "RETURN"
        "OccurredAtUTC" = $fixtureNow.AddDays(-20).ToOADate()
        "AppliedAtUTC" = $fixtureNow.AddDays(-20).ToOADate()
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-VIEWER-RETURN"
        "SKU" = "SKU-SHIP"
        "QtyDelta" = -1
        "Location" = "DOCK"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=RETURN-VIEWER;Item=Viewer Shipment Item;UOM=each"
    }
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-SHIP-REMOVE"
        "AppliedSeq" = 6
        "EventType" = "SHIP_RELEASE"
        "OccurredAtUTC" = $fixtureNow.AddDays(-6).ToOADate()
        "AppliedAtUTC" = $fixtureNow.AddDays(-6).ToOADate()
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-LIVE-SHIP"
        "SKU" = "SKU-SHIP"
        "QtyDelta" = 0
        "Location" = "DOCK"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=SHIP-REMOVE-VIEWER;Item=Viewer Shipment Item;UOM=each"
    }
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-DUMP"
        "AppliedSeq" = 7
        "EventType" = "DUMP"
        "OccurredAtUTC" = $fixtureNow.AddHours(-12).ToOADate()
        "AppliedAtUTC" = $fixtureNow.AddHours(-12).ToOADate()
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-VIEWER-DUMP"
        "SKU" = "SKU-SHIP"
        "QtyDelta" = -1
        "Location" = "DOCK"
        "Condition" = "DAMAGED"
        "AttributesJson" = "{}"
        "Note" = "Reference=DUMP-VIEWER;Item=Viewer Shipment Item;UOM=each"
    }
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-INTERNAL-RESERVE"
        "AppliedSeq" = 8
        "EventType" = "SHIP_RESERVE"
        "OccurredAtUTC" = $fixtureNow.AddHours(-10).ToOADate()
        "AppliedAtUTC" = $fixtureNow.AddHours(-10).ToOADate()
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-LIVE-SHIP"
        "SKU" = "SKU-SHIP"
        "QtyDelta" = 0
        "Location" = "DOCK"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=SHIP-ADD-VIEWER;Item=Viewer Shipment Item;UOM=each;IO=RESERVED"
    }
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-PROD-CONSUME"
        "AppliedSeq" = 9
        "EventType" = "PROD_CONSUME"
        "OccurredAtUTC" = $fixtureNow.AddHours(-8).ToOADate()
        "AppliedAtUTC" = $fixtureNow.AddHours(-8).ToOADate()
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-VIEWER-PROD-INPUT"
        "SKU" = "SKU-SUGAR"
        "QtyDelta" = -2
        "Location" = "LINE"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=PROD-INPUT-VIEWER;Item=Viewer Production Input;UOM=LB;RunId=RUN-VIEWER"
    }
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-PROD-COMPLETE"
        "AppliedSeq" = 10
        "EventType" = "PROD_COMPLETE"
        "OccurredAtUTC" = $fixtureNow.AddHours(-6).ToOADate()
        "AppliedAtUTC" = $fixtureNow.AddHours(-6).ToOADate()
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-VIEWER-PROD-OUTPUT"
        "SKU" = "SKU-COMP"
        "QtyDelta" = 2
        "Location" = "LINE"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=PROD-OUTPUT-VIEWER;Item=Viewer Production Output;UOM=LB;RunId=RUN-VIEWER"
    }
    $inventoryWb.Save()
    $opened.Add($configWb) | Out-Null
    $opened.Add($authWb) | Out-Null
    $opened.Add($inventoryWb) | Out-Null

    $step = "open packaged add-ins"
    $packages = @{}
    foreach ($packageName in @(
        "invSys.Core.xlam", "invSys.Inventory.Domain.xlam",
        "invSys.Designs.Domain.xlam", "invSys.Operations.xlam"
    )) {
        $package = $excel.Workbooks.Open((Join-Path $deploy $packageName))
        $opened.Add($package) | Out-Null
        $packages[$packageName] = $package
    }
    $coreName = [string]$packages["invSys.Core.xlam"].Name
    $operationsName = [string]$packages["invSys.Operations.xlam"].Name

    $step = "configure signed-in target"
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($runtimeRoot))
    $configLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modConfig.LoadConfig" -Arguments @($warehouseId, $stationId))
    $authLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.LoadAuth" -Arguments @($warehouseId))
    $targetResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" `
        -Arguments @($runtimeRoot, $runtimeRoot, $stationId, $true))
    $targetPathsSet = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SetCurrentTargetPathsForTest" `
        -Arguments @("\\inventory-viewer-test\warehouse", $runtimeRoot))
    $signInCredential = $testPin
    if($RejectFixtureSignInForTest) { $signInCredential = $testPin + '-invalid-fixture-credential' }
    $signInResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.SignInCurrentTargetForAutomation" `
        -Arguments @($testUser, $signInCredential, "RECEIVE_POST"))
    $facts.ConfigLoaded = $configLoaded
    $facts.AuthLoaded = $authLoaded
    $facts.TargetSelected = $targetResult.StartsWith("OK|")
    $facts.TargetPathsSet = $targetPathsSet
    $facts.SignedIn = $signInResult.StartsWith("OK|")
    # Never write the raw response: exceptional responses may contain entered data.
    $facts.SignInStatus = if($facts.SignedIn) { 'OK' } elseif($signInResult -match '^FAIL\|([1-9])$') {
        'AUTH_STATUS_' + $Matches[1]
    } elseif($signInResult -match '^FAIL\|NO_TARGET\|[0-9]+$') { 'NO_TARGET' } else { 'UNAVAILABLE' }
    $fixtureUsers=$authWb.Worksheets.Item('Users').ListObjects.Item('tblUsers')
    $matchingCredential=$false
    foreach($row in $fixtureUsers.ListRows) {
        if([string]$row.Range.Cells.Item(1,$fixtureUsers.ListColumns.Item('UserId').Index).Value2 -ceq $testUser) {
            $matchingCredential=([string]$row.Range.Cells.Item(1,$fixtureUsers.ListColumns.Item('PinHash').Index).Value2 -ceq $testPinHash)
        }
    }
    $facts.FixtureCredentialMatches=$matchingCredential
    $facts.PublicViewerEntryReached=$false
    if(-not $facts.SignedIn) { throw 'Viewer fixture sign-in failed before any public Viewer action; see sanitized status.' }

    $step = "publish isolated snapshot"
    $snapshotCreated = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modWarehouseSync.GenerateWarehouseSnapshot" `
        -Arguments @($warehouseId, $inventoryWb, $snapshotPath))
    if (-not $snapshotCreated) { throw "The isolated inventory snapshot was not created." }
    $inventoryWb.Close($true)
    $authWb.Close($false)
    $configWb.Close($false)
    $snapshotHashBefore = (Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash

    $step = "invoke public Viewer action twice"
    $facts.PublicViewerEntryReached=$true
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modOperationsInit.Auto_Open")
    $firstReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerActionForTest")
    $secondReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerActionForTest")
    $filterReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerFilterForTest" `
        -Arguments @("SKU-SHIP"))
    $step = "export Viewer ListBox to a new worksheet table"
    $listBoxExportReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerListBoxTableActionForTest")
    $listBoxExportWb = $excel.ActiveWorkbook
    if ($null -ne $listBoxExportWb -and $listBoxExportWb.Name -ne $operationsName) {
        $listBoxExportWb.Close($false)
    }
    $eventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest")

    $step = "publish a newer event and refresh the open Events page"
    $refreshInventoryWb = $excel.Workbooks.Open($inventoryPath)
    $opened.Add($refreshInventoryWb) | Out-Null
    $refreshInventoryLog = $refreshInventoryWb.Worksheets.Item("InventoryLog").ListObjects.Item("tblInventoryLog")
    Add-ListObjectRow -ListObject $refreshInventoryLog -Values @{
        "EventID" = "EVT-VIEWER-NEW-RECEIVE"
        "AppliedSeq" = 8
        "EventType" = "RECEIVE"
        "OccurredAtUTC" = [DateTime]::Now.AddMinutes(-1).ToOADate()
        "AppliedAtUTC" = [DateTime]::Now.AddMinutes(-1).ToOADate()
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-VIEWER-NEW-RECEIVE"
        "SKU" = "SKU-COMP"
        "QtyDelta" = 7
        "Location" = "DOCK"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=BOL-VIEWER-NEW;Item=New Viewer Receipt;UOM=each"
    }
    $refreshInventoryWb.Save()
    $newSnapshotCreated = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modWarehouseSync.GenerateWarehouseSnapshot" `
        -Arguments @($warehouseId, $refreshInventoryWb, $snapshotPath))
    if (-not $newSnapshotCreated) { throw "The newer isolated event snapshot was not created." }
    $refreshInventoryWb.Close($true)
    $snapshotHashPublished = (Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash
    $refreshedEventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest")
    $dayEventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest" -Arguments @("Day"))
    $weekEventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest" -Arguments @("Week"))
    $monthEventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest" -Arguments @("Month"))
    $customEventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest" -Arguments @("14"))
    $allEventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest" -Arguments @("All"))
    $rememberedSetupReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest" -Arguments @("14"))
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.CloseInventoryViewerForTest")
    $reopenedRememberedReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.CloseInventoryViewerForTest")
    $snapshotHashAfter = (Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash

    $firstOk = $firstReport -match '^OK\|' -and $firstReport -match '(?:^|\|)VisibleRows=3(?:\||$)' -and
        $firstReport -match '(?:^|\|)Generation=1(?:\||$)'
    $secondReused = $secondReport -match '^OK\|' -and
        $secondReport -match '(?:^|\|)VisibleRows=3(?:\||$)' -and
        $secondReport -match '(?:^|\|)Generation=1(?:\||$)'
    $filterOk = $filterReport -match '^OK\|' -and
        $filterReport -match '(?:^|\|)VisibleRows=1(?:\||$)' -and
        $filterReport -match '(?:^|\|)Generation=1(?:\||$)'
    $listBoxExportOk = $listBoxExportReport -match '^OK\|' -and
        $listBoxExportReport -match '(?:^|\|)ListBox->Table=True(?:\||$)' -and
        $listBoxExportReport -match '(?:^|\|)Exported=True(?:\||$)' -and
        $listBoxExportReport -match 'new unsaved workbook'
    $snapshotAdvanced = $snapshotHashBefore -ne $snapshotHashPublished
    $snapshotUnchanged = $snapshotHashPublished -eq $snapshotHashAfter
    $eventsOk = $eventsReport -match '^OK\|' -and
        $eventsReport -match '(?:^|\|)TabCount=3(?:\||$)' -and
        $eventsReport -match '(?:^|\|)TabCaptions=Inventory,Events,ListBox->Table(?:\||$)' -and
        $eventsReport -match '(?:^|\|)SelectedTab=Events(?:\||$)' -and
        $eventsReport -match '(?:^|\|)Title=Inventory and shipping events(?:\||$)' -and
        $eventsReport -match '(?:^|\|)VisibleRows=6(?:\||$)' -and
        $eventsReport -match '(?:^|\|)ReadableDates=6(?:\||$)' -and
        $eventsReport -match '(?:^|\|)EventHeaderAligned=True(?:\||$)' -and
        $eventsReport -match '(?:^|\|)FirstReference=PROD-OUTPUT-VIEWER(?:\||$)' -and
        $eventsReport -match '(?:^|\|)RemoveRows=1(?:\||$)' -and
        $eventsReport -match '(?:^|\|)ShipmentHeldRows=0(?:\||$)' -and
        $eventsReport -match '(?:^|\|)ProductionInputRows=1(?:\||$)' -and
        $eventsReport -match '(?:^|\|)ProductionOutputRows=1(?:\||$)' -and
        $eventsReport -match '(?:^|\|)EventRange=All(?:\||$)' -and
        $eventsReport -match '(?:^|\|)RangeControlVisible=True(?:\||$)' -and
        $eventsReport -match '(?:^|\|)Columns=10(?:\||$)' -and
        $eventsReport -match '(?:^|\|)ReadOnly=True(?:\||$)'
    $invalidRememberedFallbackOk = $eventsReport -match '(?:^|\|)EventRange=All(?:\||$)'
    $refreshedEventsOk = $refreshedEventsReport -match '^OK\|' -and
        $refreshedEventsReport -match '(?:^|\|)VisibleRows=7(?:\||$)' -and
        $refreshedEventsReport -match '(?:^|\|)ReadableDates=7(?:\||$)' -and
        $refreshedEventsReport -match '(?:^|\|)FirstReference=BOL-VIEWER-NEW(?:\||$)'
    $dateFiltersOk = $dayEventsReport -match '(?:^|\|)EventRange=Day(?:\||$)' -and
        $dayEventsReport -match '(?:^|\|)VisibleRows=4(?:\||$)' -and
        $weekEventsReport -match '(?:^|\|)EventRange=Week(?:\||$)' -and
        $weekEventsReport -match '(?:^|\|)VisibleRows=5(?:\||$)' -and
        $monthEventsReport -match '(?:^|\|)EventRange=Month(?:\||$)' -and
        $monthEventsReport -match '(?:^|\|)VisibleRows=6(?:\||$)' -and
        $customEventsReport -match '(?:^|\|)EventRange=14(?:\||$)' -and
        $customEventsReport -match '(?:^|\|)VisibleRows=5(?:\||$)' -and
        $allEventsReport -match '(?:^|\|)EventRange=All(?:\||$)' -and
        $allEventsReport -match '(?:^|\|)VisibleRows=7(?:\||$)'
    $rememberedRangeOk = $rememberedSetupReport -match '(?:^|\|)EventRange=14(?:\||$)' -and
        $reopenedRememberedReport -match '(?:^|\|)EventRange=14(?:\||$)' -and
        $reopenedRememberedReport -match '(?:^|\|)VisibleRows=5(?:\||$)' -and
        $reopenedRememberedReport -match '(?:^|\|)Generation=2(?:\||$)'

    $facts.ConfigLoaded = $configLoaded
    $facts.AuthLoaded = $authLoaded
    $facts.TargetSelected = $targetResult.StartsWith("OK|")
    $facts.TargetPathsSet = $targetPathsSet
    $facts.SignedIn = $signInResult.StartsWith("OK|")
    $facts.SnapshotCreated = $snapshotCreated
    $facts.FirstActionRows = if ($firstOk) { 3 } else { 0 }
    $facts.RepeatedLaunchReusedGeneration = $secondReused
    $facts.FilterVisibleRows = if ($filterOk) { 1 } else { 0 }
    $facts.ListBoxTableExport = $listBoxExportOk
    $facts.EventsVisibleRows = if ($eventsOk) { 6 } else { 0 }
    $facts.RefreshedEventsVisibleRows = if ($refreshedEventsOk) { 7 } else { 0 }
    $facts.NewestPublishedReference = if ($refreshedEventsOk) { "BOL-VIEWER-NEW" } else { "Unexpected" }
    $facts.ReadableEventDates = $eventsOk -and $refreshedEventsOk
    $facts.ViewerTabCount = if ($eventsOk) { 3 } else { 0 }
    $facts.ViewerTabCaptions = if ($eventsOk) { "Inventory,Events,ListBox->Table" } else { "Unexpected" }
    $facts.SelectedViewerTab = if ($eventsOk) { "Events" } else { "Unexpected" }
    $facts.RemoveEventsVisible = $eventsOk
    $facts.InternalReservationHidden = $eventsOk
    $facts.ProductionInputEventsVisible = $eventsOk
    $facts.ProductionOutputEventsVisible = $eventsOk
    $facts.EventsReadOnly = $eventsOk
    $facts.RollingDateFilters = $dateFiltersOk
    $facts.RememberedRangeAfterReopen = $rememberedRangeOk
    $facts.InvalidRememberedRangeFallsBackToAll = $invalidRememberedFallbackOk
    $facts.SnapshotHashUnchanged = $snapshotUnchanged
    $facts.NewPublicationChangedSnapshot = $snapshotAdvanced
    if (-not $listBoxExportOk) { $facts.ListBoxTableExportReport = $listBoxExportReport }
    if (-not $eventsOk) { $facts.InitialEventsReport = $eventsReport }
    if (-not $refreshedEventsOk) { $facts.RefreshedEventsReport = $refreshedEventsReport }
    if (-not $dateFiltersOk) {
        $facts.DateFilterReports = "Day=$dayEventsReport; Week=$weekEventsReport; Month=$monthEventsReport; Custom=$customEventsReport; All=$allEventsReport"
    }
    if (-not $rememberedRangeOk) {
        $facts.RememberedRangeReports = "Setup=$rememberedSetupReport; Reopened=$reopenedRememberedReport"
    }

    $passed = $configLoaded -and $authLoaded -and
        $targetResult.StartsWith("OK|") -and $targetPathsSet -and
        $signInResult.StartsWith("OK|") -and $snapshotCreated -and
        $firstOk -and $secondReused -and $filterOk -and $listBoxExportOk -and $eventsOk -and
        $refreshedEventsOk -and $dateFiltersOk -and $rememberedRangeOk -and
        $invalidRememberedFallbackOk -and
        $snapshotAdvanced -and $snapshotUnchanged
    $detail = if ($passed) {
        "The public Operations Viewer action exported the displayed ListBox to a new unsaved worksheet table, loaded readable Receipt, Production input/output, and Shipping Remove events, excluded the internal SHIP_RESERVE fixture from the operator-action log, refreshed the already-open Events page to show a newly published receipt first, applied All/Day/Week/Month/custom rolling-day filters, restored custom 14 days after form close/reopen, kept Events read-only, and left the new snapshot byte-for-byte unchanged."
    } else {
        "The packaged Viewer contract failed at step '$step'."
    }
}
catch {
    $detail = "Harness exception at step '$step': $($_.Exception.Message)"
    $facts.HarnessException = $_.Exception.Message
}
finally {
    $lines = @(
        "# Inventory Viewer Packaged Results", "",
        "- Status: **$(if ($passed) { 'PASS' } else { 'FAIL' })**",
        "- Runtime: isolated generated test warehouse"
    )
    foreach ($entry in $facts.GetEnumerator()) {
        $lines += "- $($entry.Key): $($entry.Value)"
    }
    $lines += ""
    $lines += "## Observed result"
    $lines += ""
    $lines += $detail
    [IO.File]::WriteAllText($resultPath, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))

    foreach ($wb in $opened) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    if ($preferenceExisted) {
        New-Item -Path $preferenceRegistryPath -Force | Out-Null
        Set-ItemProperty -LiteralPath $preferenceRegistryPath -Name $preferenceName `
            -Value $preferenceBefore -Type String
    } elseif (Test-Path -LiteralPath $preferenceRegistryPath) {
        Remove-ItemProperty -LiteralPath $preferenceRegistryPath -Name $preferenceName `
            -ErrorAction SilentlyContinue
    }
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    $resolvedRuntime = [IO.Path]::GetFullPath($runtimeRoot)
    if ($resolvedRuntime.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedRuntime) -like "invsys-inventory-viewer-*") {
        Remove-Item -LiteralPath $resolvedRuntime -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host $detail
Write-Host "Evidence: $resultPath"
if (-not $passed) { exit 1 }
