[CmdletBinding()]
param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path

function Read-Source {
    param([string]$RelativePath)
    Get-Content -LiteralPath (Join-Path $repo $RelativePath) -Raw
}

function Add-Check {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Name,
        [bool]$Passed,
        [string]$Detail
    )
    $Rows.Add([pscustomobject]@{
        Name = $Name
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

$bridge = Read-Source "src/Core/Modules/modOperationsPrimitiveBridge.bas"
$bootstrap = Read-Source "src/Core/Modules/modWarehouseBootstrap.bas"
$roleSurfaces = Read-Source "src/Core/Modules/modRoleWorkbookSurfaces.bas"
$receiving = Read-Source "src/Receiving/Modules/modTS_Received.bas"
$receivingForm = Read-Source "src/Receiving/Forms/frmReceiving.frm"
$production = Read-Source "src/Production/Modules/mProduction.bas"
$shipping = Read-Source "src/Shipping/Modules/modTS_Shipments.bas"
$shippingContext = Read-Source "src/Shipping/Modules/modShippingFormContext.bas"
$receivingEvents = Read-Source "src/Receiving/ClassModules/cReceivingAppEvents.cls"
$productionEvents = Read-Source "src/Production/ClassModules/cProductionAppEvents.cls"
$shippingEvents = Read-Source "src/Shipping/ClassModules/cShippingAppEvents.cls"
$nasValidator = Read-Source "tools/validate_plan022_nas_runtime.ps1"

$rows = New-Object System.Collections.Generic.List[object]

Add-Check $rows "Core primitive resolves eligible role workbook names" `
    ($bridge -match '(?im)^Public Function ResolveEligibleRoleOperatorWorkbookName\(') `
    "Cross-XLAM launcher resolution must use primitive workbook names."

Add-Check $rows "Core primitive provisions current Receiving operator workbook" `
    ($bridge -match '(?im)^Public Function OpenOrCreateCurrentReceivingOperatorWorkbook\(') `
    "Receiving launcher requires one Core-owned open/create primitive."

Add-Check $rows "Core primitive provisions every current role operator workbook" `
    (($bridge -match '(?im)^Public Function OpenOrCreateCurrentRoleOperatorWorkbook\(') -and
     ($bootstrap -match '(?im)^Public Function OpenOrCreateRoleOperatorWorkbookForCurrentTarget\(')) `
    "Receiving, Production, and Shipping must share one Core-owned station-local open/create boundary."

Add-Check $rows "Bootstrap supports isolated operator-root automation" `
    (($bootstrap -match '(?im)^Public Function SetLocalOperatorRootOverrideForAutomation\(') -and
     ($bootstrap -match '(?im)^Public Function OpenOrCreateReceivingOperatorWorkbookForCurrentTarget\(')) `
    "Focused tests must not create station-local workbooks in the real user profile."

Add-Check $rows "Receiving launcher uses Core provisioning primitive" `
    ($receiving -match 'OpenOrCreateCurrentReceivingOperatorWorkbook') `
    "The public Receiving callback must provision/reuse the station-local workbook."

Add-Check $rows "Receiving launcher retains one reusable form instance" `
    (($receiving -match '(?im)^Private mReceivingLauncherForm As frmReceiving') -and
     ($receiving -match 'mReceivingLauncherForm\.SetOperatorWorkbook')) `
    "Repeated clicks must reuse one valid modeless form binding."

Add-Check $rows "Receiving launcher rejects a disappeared form reference" `
    (($receiving -match '(?im)^Private Function IsReceivingLauncherFormReusable\(') -and
     ($receiving -match 'If Not IsReceivingLauncherFormReusable\(wb\) Then') -and
     ($receiving -match 'IsReceivingLauncherFormReusable = visibleState') -and
     ($receiving -match '(?s)If mReceivingLauncherForm\.Visible Then Unload mReceivingLauncherForm.*?Set mReceivingLauncherForm = Nothing')) `
    "A closed/disappeared modeless form must be recreated against the same captured workbook."

Add-Check $rows "Receiving form termination invalidates its launcher cache" `
    (($receiving -match '(?im)^Public Sub NotifyReceivingLauncherFormTerminating\(') -and
     ($receivingForm -match '(?s)Private Sub UserForm_QueryClose\(Cancel As Integer, CloseMode As Integer\).*?modTS_Received\.NotifyReceivingLauncherFormTerminating Me')) `
    "The form must invalidate its cached launcher reference before Excel destroys the modeless proxy."

Add-Check $rows "Operations capability rendering uses the signed-in cache" `
    ($bridge -match 'modRoleUiAccess\.CanCurrentUserPerformCapabilityCached') `
    "Opening a role form must not reload and save canonical config/auth workbooks."

Add-Check $rows "Production resolver has no Operations XLAM fallback" `
    ($production -notmatch 'Set ResolveProductionWorkbook = ThisWorkbook') `
    "ThisWorkbook is the Operations XLAM, never operator authority."

Add-Check $rows "Production launcher uses eligible-role primitive" `
    ($production -match 'OpenOrCreateCurrentRoleOperatorWorkbook') `
    "Production must reuse or self-provision the station-local role workbook before form initialization."

Add-Check $rows "Production launcher uses signed-in capability cache" `
    ($production -match '(?s)Public Sub BtnOpenProductionForm\(\).*?RequireCurrentUserCapabilityCached\(\"PROD_POST\"\)') `
    "The public launcher must not reload and save canonical config/auth workbooks."

Add-Check $rows "Shipping resolver has no Operations XLAM fallback" `
    ($shipping -notmatch 'Set ResolveShippingWorkbook = ThisWorkbook') `
    "ThisWorkbook is the Operations XLAM, never operator authority."

Add-Check $rows "Shipping launcher uses eligible-role primitive" `
    ($shipping -match 'OpenOrCreateCurrentRoleOperatorWorkbook') `
    "Shipping must reuse or self-provision the station-local role workbook before form initialization."

Add-Check $rows "Shipping launcher uses signed-in capability cache" `
    ($shipping -match '(?s)Public Sub BtnOpenShipmentsForm\(\).*?RequireCurrentUserCapabilityCached\(\"SHIP_POST\"\)') `
    "The public launcher must not reload and save canonical config/auth workbooks."

Add-Check $rows "Shipping launcher retains one reusable form instance" `
    (($shipping -match '(?im)^Private mShipmentsLauncherForm As frmShipmentsTally') -and
     ($shipping -match '(?s)Public Sub BtnOpenShipmentsForm\(\).*?If Not modShippingFormContext.CanReuse\(mShipmentsLauncherForm\).*?Set mShipmentsLauncherForm = modShippingFormContext.CreateBoundForm\(') -and
     ($shippingContext -match '(?s)Public Function CanReuse\(ByVal form As frmShipmentsTally\).*?CanReuse = form.HasCurrentContext\(\)') -and
     ($shippingContext -match '(?s)Public Function CreateBoundForm\(ByVal operatorWb As Workbook,.*?Set fresh = New frmShipmentsTally.*?fresh.SetOperatorWorkbook operatorWb')) `
    "Repeated Shipping clicks reuse a current binding; explicit stale-form recovery uses the typed bound factory."

Add-Check $rows "Shipping launcher crosses Core with primitive workbook name" `
    (($shipping -match 'modOperationsPrimitiveBridge\.BeginQuietUiForWorkbook\(wb\.Name') -and
     ($shipping -match 'modOperationsPrimitiveBridge\.EnsureShippingWorkbookSurface\(wb\.Name')) `
    "Shipping must not pass a Workbook object across the Core/Operations XLAM boundary."

Add-Check $rows "Generated Shipping surface retains a visible landing sheet" `
    ($roleSurfaces -match '(?s)Public Function EnsureShippingWorkbookSurface.*?EnsureWorksheetSurface\(wb, "invSys Shipping"\).*?landing\.Visible = xlSheetVisible') `
    "A newly generated Shipping operator workbook must retain one visible operator landing sheet."

Add-Check $rows "Receiving closes stale form binding with captured workbook" `
    (($receiving -match '(?im)^Public Sub HandleReceivingOperatorWorkbookClosing\(') -and
     ($receivingEvents -match 'modTS_Received\.HandleReceivingOperatorWorkbookClosing Wb')) `
    "A closed captured workbook must not leave a stale Receiving form binding."

Add-Check $rows "Production closes stale form binding with captured workbook" `
    (($production -match '(?im)^Public Sub HandleProductionOperatorWorkbookClosing\(') -and
     ($productionEvents -match 'mProduction\.HandleProductionOperatorWorkbookClosing Wb')) `
    "A closed captured workbook must not leave a stale Production form binding."

Add-Check $rows "Shipping closes stale form binding with captured workbook" `
    (($shipping -match '(?im)^Public Sub HandleShippingOperatorWorkbookClosing\(') -and
     ($shippingEvents -match 'modTS_Shipments\.HandleShippingOperatorWorkbookClosing Wb')) `
    "A closed captured workbook must not leave a stale Shipping form binding."

Add-Check $rows "UAT PIN preparation uses masked confirmed input" `
    (($nasValidator -match '\[switch\]\$PrepareUserAcceptancePin') -and
     ($nasValidator -match '(?s)Read-ConfirmedUserAcceptancePin.*?Read-Host.*?-AsSecureString.*?Read-Host.*?-AsSecureString') -and
     ($nasValidator -match 'ZeroFreeBSTR')) `
    "The dedicated UAT credential must be entered locally, confirmed, and cleared."

Add-Check $rows "UAT PIN preparation never writes the PIN" `
    ($nasValidator -notmatch '(?im)Write-(?:Output|Host|Verbose|Debug|Warning).*?\$uatPin') `
    "The temporary UAT PIN must not appear in console or generated evidence."

Add-Check $rows "Automated NAS validation cannot overwrite the UAT user PIN" `
    (($nasValidator -match '\$uatUser\s*=\s*if\s*\(') -and
     ($nasValidator -match '\$testUser\s*=\s*\"plan022-auto-\"') -and
     ($nasValidator -match '(?s)if \(\$PrepareUserAcceptancePin\).*?Set-TestUserPinHash.*?-UserId \$uatUser') -and
     ($nasValidator -match '(?s)for \(\$sessionNumber = 1;.*?Set-TestUserPinHash.*?-UserId \$testUser')) `
    "Automation must use a dedicated account so rerunning NAS validation preserves the human UAT credential hash."

$failed = @($rows | Where-Object { -not $_.Passed })
foreach ($row in $rows) {
    $status = if ($row.Passed) { "PASS" } else { "RED" }
    Write-Output ("{0} {1}: {2}" -f $status, $row.Name, $row.Detail)
}
Write-Output ("PASS={0} RED={1} TOTAL={2}" -f ($rows.Count - $failed.Count), $failed.Count, $rows.Count)
if ($failed.Count -gt 0) {
    exit 1
}
