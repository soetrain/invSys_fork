[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$receivingRoot = Join-Path $repo "src/Receiving"
$modulePath = Join-Path $receivingRoot "Modules/modTS_Received.bas"
$servicePath = Join-Path $receivingRoot "Modules/modReceivingPostingService.bas"
$statePath = Join-Path $receivingRoot "ClassModules/cReceivingWorkflowState.cls"
$formPath = Join-Path $receivingRoot "Forms/frmReceiving.frm"
$legacyFormPath = Join-Path $receivingRoot "Forms/frmReceivedTally.frm"
$legacyCreatorPath = Join-Path $receivingRoot "Modules/modReceivingEventCreator.bas"
$buildPath = Join-Path $repo "tools/build-xlam.ps1"

$moduleText = Get-Content -Raw -LiteralPath $modulePath
$formText = Get-Content -Raw -LiteralPath $formPath
$actionText = Get-Content -Raw -LiteralPath (Join-Path $receivingRoot 'Modules/modReceivingActivityAction.bas')
$buildText = Get-Content -Raw -LiteralPath $buildPath
$serviceText = if (Test-Path -LiteralPath $servicePath) {
    Get-Content -Raw -LiteralPath $servicePath
} else { "" }
$stateText = if (Test-Path -LiteralPath $statePath) {
    Get-Content -Raw -LiteralPath $statePath
} else { "" }
$runtimeText = (
    Get-ChildItem -LiteralPath $receivingRoot -Recurse -File |
        Where-Object { $_.Extension -in @(".bas", ".cls", ".frm") } |
        Sort-Object FullName |
        ForEach-Object { Get-Content -Raw -LiteralPath $_.FullName }
) -join "`n"

$results = [System.Collections.Generic.List[object]]::new()
function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Detail)
    $results.Add([pscustomobject]@{
        Check = $Name
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

Add-Check "Receiving.Workflow.TypedStateAndService" `
    ((Test-Path -LiteralPath $statePath -PathType Leaf) -and
     (Test-Path -LiteralPath $servicePath -PathType Leaf) -and
     ($stateText -match '(?i)STATE_STAGED') -and
     ($stateText -match '(?i)STATE_VALIDATED') -and
     ($stateText -match '(?i)STATE_SUBMITTED') -and
     ($stateText -match '(?i)STATE_PROCESSOR_APPLIED') -and
     ($stateText -match '(?i)STATE_SNAPSHOT_REFRESHED') -and
     ($stateText -match '(?i)STATE_READY')) `
    "Receiving needs typed workflow state for staged -> validated -> submitted -> processor applied -> snapshot refreshed -> ready."

Add-Check "Receiving.Identity.NoManagedRowLiteral" `
    ($runtimeText -notmatch '(?i)"ROW"') `
    "Receiving runtime code must not declare, resolve, serialize, display, or restore the retired ROW identity."

$legacyAuthorityPattern = '(?i)\b(LookupInvSysByROW|FindInvRowByROW|ResolveInvRowForReceiveLog|' +
                          'NormalizeReceivingInventoryRowDisplay|NormalizeInventoryRowForWriteReceiving|' +
                          'ReceivingDemoRowForSku|invRow|rowValue)\b'
Add-Check "Receiving.Identity.NoNumericRowAuthority" `
    ($runtimeText -notmatch $legacyAuthorityPattern) `
    "Receiving must select, stage, aggregate, log, and post by immutable System_Key rather than worksheet position."

Add-Check "Receiving.Identity.SystemKeyStagingAndPosting" `
    (($moduleText -match '(?i)"System_Key"') -and
     ($serviceText -match '(?i)"System_Key"') -and
     ($formText -match '(?i)System_Key')) `
    "The form, staging facade, and posting service must carry System_Key end to end."

$confirmBody = ""
$confirmMatch = [regex]::Match(
    $formText,
    '(?is)Private\s+Sub\s+mBtnConfirm_Click\s*\(\s*\)(?<body>.*?)End\s+Sub'
)
if ($confirmMatch.Success) { $confirmBody = $confirmMatch.Groups["body"].Value }
$formActionBody=[regex]::Match($actionText,'(?is)Public\s+Function\s+ConfirmWrites\b(?<body>.*?)End\s+Function').Groups['body'].Value
$sharedConfirmBody=[regex]::Match($actionText,'(?is)Private\s+Function\s+ConfirmAction\b(?<body>.*?)End\s+Function').Groups['body'].Value
Add-Check "Receiving.Form.RealActionUsesTypedService" `
    (($confirmBody -match '(?i)modReceivingActivityAction\.ConfirmWrites') -and
     ([regex]::Matches($formActionBody,'(?i)\bConfirmAction\s*\(').Count -eq 1) -and
     ($formActionBody -notmatch '(?i)modReceivingPostingService\.ExecuteConfirmWrites') -and
     ([regex]::Matches($sharedConfirmBody,'(?i)modReceivingPostingService\.ExecuteConfirmWrites').Count -eq 1) -and
     ($confirmBody -match '(?i)mOperatorWorkbook') -and
     ($confirmBody -notmatch '(?i)modTS_Received\.ConfirmWrites') -and
     ($confirmBody -notmatch '(?i)ClearReceivingFormStaging')) `
    "The operator handler must reach the typed service once through the context/observation controller; the owning service retains staging cleanup."

Add-Check "Receiving.Form.ModelessCapturedContext" `
    (($moduleText -match '(?i)(?:frm|mReceivingLauncherForm)\.Show\s+vbModeless') -and
     ($formText -match '(?i)Private\s+mOperatorWorkbook\s+As\s+Workbook') -and
     ($formText -notmatch '(?i)\bApplication\.ActiveWorkbook\b') -and
     ($formText -notmatch '(?i)\.Activate\b')) `
    "The modeless Receiving form must stay bound to its captured operator workbook without activating or recapturing ActiveWorkbook."

Add-Check "Receiving.Form.PurchasingStub" `
    (($formText -match '(?i)MSForms\.(TabStrip|MultiPage)') -and
     ($formText -match '(?i)Purchasing') -and
     ($formText -match '(?i)not yet operational') -and
     ($formText -match '(?i)TestPurchasingTabContract')) `
    "The main Receiving shell must expose a selectable, visibly non-operational Purchasing tab with a packaged navigation probe."

$receivingProjectMatch = [regex]::Match(
    $buildText,
    '(?is)Key\s*=\s*"Receiving".*?(?=\n\s*@\{\s*\n\s*Key\s*=|\n\)\s*$)'
)
$receivingProjectText = if ($receivingProjectMatch.Success) {
    $receivingProjectMatch.Value
} else { "" }
Add-Check "Receiving.Ribbon.NoPurchasingLaunchSurface" `
    ($buildText -notmatch '(?i)(Id|Label)\s*=\s*"[^"]*Purchas') `
    "Purchasing remains a form-only stub and must not gain a ribbon button, group, or launch surface."

Add-Check "Receiving.Legacy.RedundantPostingRetired" `
    ((-not (Test-Path -LiteralPath $legacyCreatorPath)) -and
     (-not (Test-Path -LiteralPath $legacyFormPath)) -and
     ($runtimeText -notmatch '(?i)\b(UndoInvDeltas|RecordInvDelta|UpdateInventory)\b')) `
    "The duplicate event creator and direct worksheet-mutation form/undo paths must be retired."

Add-Check "Receiving.Service.EventIdentityAndClearGate" `
    (($serviceText -match '(?i)EnsureEventIdentities') -and
     ($serviceText -match '(?i)RunBatchAndRefreshOperatorWorkbook') -and
     ($serviceText -match '(?i)MarkProcessorApplied') -and
     ($serviceText -match '(?i)MarkSnapshotRefreshed') -and
     ($serviceText -match '(?i)ClearReceivingStaging')) `
    "Stable event identity must survive retries, and staging may clear only after processor application plus snapshot refresh."

$results | Format-Table -AutoSize
$failed = @($results | Where-Object { -not $_.Passed })
Write-Host ("Slice 10 Receiving stabilization contract: {0} passed, {1} failed" -f
    ($results.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
