[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$docs = (Resolve-Path (Join-Path $repo "..\invSys_docs")).Path

function Read-Source([string]$path) {
    Get-Content -Raw -LiteralPath $path
}

$form = Read-Source (Join-Path $repo "src\Admin\Forms\frmCreateDeleteUser.frm")
$spec = Read-Source (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Source (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$failures = [System.Collections.Generic.List[string]]::new()
$passes = [System.Collections.Generic.List[string]]::new()

function Check([string]$name, [bool]$passed, [string]$contract) {
    if ($passed) {
        $passes.Add($name)
        Write-Host "PASS $name"
    } else {
        $failures.Add("${name}: ${contract}")
        Write-Host "FAIL $name - $contract"
    }
}

$packet = [regex]::Match($form, '(?ms)^Private Function BuildAccountClipboardTextForm\(\) As String.*?^End Function').Value
$copyHandler = [regex]::Match($form, '(?ms)^Private Sub mBtnCopyPin_Click\(\).*?^End Sub').Value

Check "OnboardingPacket.FormActionUsesPacket" (
    $copyHandler.Contains('BuildAccountClipboardTextForm()') -and
    $copyHandler.Contains('CopyTextToClipboardForm(accountText)')
) "The visible Copy Account & Setup form action must use the protected generated packet rather than a separate clipboard path."

Check "OnboardingPacket.CopyActionIsLabeled" (
    $form.Contains('"Copy Account & Setup"')
) "The existing copy action must identify that it produces both account and station-setup guidance."

Check "OnboardingPacket.StatesStationSetupCommand" (
    $packet.Contains('NAS_STATION_SETUP_COMMAND_FORM') -and
    $packet.Contains('Install-invSys-Station.cmd')
) "The packet must use the configured NAS StationSetup command, not an individual XLAM or GitHub download."

Check "OnboardingPacket.BindingSequence" (
    $packet.Contains('Server Sign In') -and
    $packet.Contains('select the warehouse target') -and
    $packet.Contains('invSys Sign In')
) "The packet must state the binding install -> server -> target -> invSys user sequence."

Check "OnboardingPacket.SeparatesNetworkAndAuthority" (
    $packet.Contains('authorized NAS/Tailscale access') -and
    $packet.Contains('does not grant an invSys role')
) "The packet must distinguish NAS access from invSys authorization."

Check "OnboardingPacket.NoNasCredentialExposure" (
    -not $packet.Contains('mTxtNasPassword') -and
    -not $packet.Contains('mTxtNasUser')
) "The packet must not copy NAS credentials."

Check "OnboardingPacket.ContractDocumented" (
    $spec.Contains('**Admin user-onboarding packet:**') -and
    $plan.Contains('**Slice 4bk -- generated user onboarding packet: approved; implementation in progress.**')
) "Architecture and Plan 022 must define the generated onboarding-packet contract before its implementation."

Write-Host "RESULT passed=$($passes.Count) failed=$($failures.Count)"
if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "  $_" }
    exit 1
}
