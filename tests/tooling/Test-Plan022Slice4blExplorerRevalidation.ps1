[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$docs = (Resolve-Path (Join-Path $repo "..\invSys_docs")).Path

function Read-Source([string]$path) {
    Get-Content -Raw -LiteralPath $path
}

$nas = Read-Source (Join-Path $repo "src\Core\Modules\modNasConnection.bas")
$form = Read-Source (Join-Path $repo "src\Core\Forms\frmWarehouseConnection.frm")
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

$revalidate = [regex]::Match($nas, '(?ms)^Public Function TryRevalidateRememberedRoot.*?^End Function').Value
$folderExists = [regex]::Match($nas, '(?ms)^Private Function FolderExistsNas.*?^End Function').Value
$normalizer = [regex]::Match($nas, '(?ms)^Private Function NormalizeFolderNas.*?^End Function').Value
$discover = [regex]::Match($nas, '(?ms)^Public Function DiscoverVisibleNasRoots.*?^End Function').Value
$scanHandler = [regex]::Match($form, '(?ms)^Private Sub ScanRootCandidates\(\).*?^End Sub').Value
$connectHandler = [regex]::Match($form, '(?ms)^Private Sub mBtnConnect_Click\(\).*?^End Sub').Value

Check "ExplorerRevalidation.DirectValidationBeforeReconnect" (
    $revalidate.IndexOf('FolderExistsNas(rootPath)') -ge 0 -and
    $revalidate.IndexOf('FolderExistsNas(rootPath)') -lt $revalidate.IndexOf('TryReconnectRememberedShareNas(rootPath, shareRoot)')
) "Core must validate an existing UNC folder before a WNet reconnect can reject it."

Check "ExplorerRevalidation.FilesystemFallback" (
    $folderExists.Contains('GetAttr(folderPath)') -and
    $folderExists.Contains('CreateObject("Scripting.FileSystemObject")') -and
    $folderExists.Contains('FolderExists(folderPath)')
) "Folder validation must fall back to the Windows filesystem provider when GetAttr cannot observe an Explorer-established SMB session."

Check "ExplorerRevalidation.FailureRemainsFailClosed" (
    $revalidate.Contains('TryReconnectRememberedShareNas(rootPath, shareRoot)') -and
    $nas.Contains('Remembered NAS root is unreachable. Windows error')
) "An unavailable root must retain the existing reconnect and fail-closed status path."

Check "ExplorerRevalidation.CanonicalizesExtraLeadingSeparators" (
    $normalizer.Contains('Do While Left$(NormalizeFolderNas, 3)') -and
    $normalizer.Contains('Mid$(NormalizeFolderNas, 4)')
) "A user-entered UNC root with extra leading separators must normalize to exactly two before validation."

Check "ExplorerRevalidation.ScanDiscoversExistingWindowsSharesBeforeTextboxFallback" (
    $nas.Contains('WScript.Shell') -and
    $nas.Contains('net use') -and
    $nas.Contains('net view') -and
    $scanHandler.Contains('DiscoverVisibleNasRoots') -and
    -not $scanHandler.Contains('Trim$(CStr(mTxtRoot.Value))')
) "The public Scan Roots handler must enumerate visible Windows SMB server/share candidates without consulting a manually entered root."

Check "ExplorerRevalidation.ConnectScansOnlySelectedConnectedRoot" (
    $connectHandler.Contains('ConnectNasRootWithCredentials') -and
    $connectHandler.Contains('ScanConnectedRoot') -and
    $form.Contains('Private Sub mLstRoots_Change') -and
    $form.Contains('Set mLstRoots = AddListBox')
) "Selecting a discovered root, connecting with credentials, and scanning runtimes must remain distinct public form actions."

Check "ExplorerRevalidation.DiscoveryRemainsReadOnly" (
    -not $discover.Contains('WNetAddConnection2') -and
    -not $discover.Contains('SaveSetting') -and
    -not $discover.Contains('RememberRoot')
) "SMB-share discovery must not reconnect, persist state, or alter an existing target."

Check "ExplorerRevalidation.ContractDocumented" (
    $spec.Contains('**Explorer-compatible NAS revalidation:**') -and
    $spec.Contains('**Scan Roots** action') -and
    $plan.Contains('**Slice 4bl -- Explorer-compatible NAS revalidation: implemented; visible UAT pending.**')
) "Architecture v4.11 and Plan 022 must define this compatibility behavior before Core implementation."

Write-Host "RESULT passed=$($passes.Count) failed=$($failures.Count)"
if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "  $_" }
    exit 1
}
