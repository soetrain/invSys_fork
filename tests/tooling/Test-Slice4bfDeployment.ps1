[CmdletBinding()]
param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$requiredTools = @(
    "tools/publish_invsys_release.ps1",
    "tools/update_invsys_station.ps1",
    "tools/rollback_invsys_station_release.ps1",
    "tools/register_invsys_update_task.ps1"
)
$expectedPackages = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam",
    "invSys.Admin.xlam"
)

$results = [System.Collections.Generic.List[object]]::new()
function Add-Check([string]$Name, [bool]$Passed, [string]$Detail) {
    $results.Add([pscustomobject]@{ Check = $Name; Passed = $Passed; Detail = $Detail }) | Out-Null
}

foreach ($relativePath in $requiredTools) {
    Add-Check ("Slice4bf.Tool.Exists.$([IO.Path]::GetFileNameWithoutExtension($relativePath))") `
        (Test-Path -LiteralPath (Join-Path $repo $relativePath) -PathType Leaf) `
        "D16 requires the $relativePath entry point."
}

$allToolText = ($requiredTools | ForEach-Object {
    $path = Join-Path $repo $_
    if (Test-Path -LiteralPath $path) { Get-Content -LiteralPath $path -Raw }
}) -join "`n"
Add-Check "Slice4bf.Tooling.FivePackageContract" `
    (@($expectedPackages | Where-Object { $allToolText -notmatch [regex]::Escape($_) }).Count -eq 0) `
    "D16 deployment tools must name every normative XLAM."
Add-Check "Slice4bf.Tooling.HashVerification" `
    (($allToolText -match "Get-FileHash") -and ($allToolText -match "SHA256")) `
    "A station must verify the release manifest hashes before registration."
Add-Check "Slice4bf.Tooling.ExcelDeferral" `
    ($allToolText -match "Get-Process" -and $allToolText -match "EXCEL") `
    "Automatic update must defer while Excel is open."
Add-Check "Slice4bf.Tooling.LocalAdminRollback" `
    ($allToolText -match "WindowsBuiltInRole" -and $allToolText -match "Administrator") `
    "Rollback must require a local Windows administrator."
Add-Check "Slice4bf.Tooling.ScheduledCadence" `
    ($allToolText -match "AtLogOn" -and $allToolText -match "Minutes 15") `
    "The updater task must be registered at logon and every fifteen minutes."
Add-Check "Slice4bf.Tooling.TaskActionUsesLocalAgent" `
    ((Get-Content -LiteralPath (Join-Path $repo "tools/register_invsys_update_task.ps1") -Raw) -match 'New-ScheduledTaskAction[\s\S]*?-f \$agentUpdater') `
    "The registered scheduled-task action must target the local agent, not the repository updater."

if (@($results | Where-Object { -not $_.Passed }).Count -eq 0) {
    $scratch = Join-Path ([IO.Path]::GetTempPath()) ("invSys-Slice4bf-" + [guid]::NewGuid().ToString("N"))
    $registryPath = "HKCU:\Software\invSysTest\Slice4bf\" + [guid]::NewGuid().ToString("N")
    try {
        $source = Join-Path $repo "deploy/current"
        $feed = Join-Path $scratch "feed"
        $cache = Join-Path $scratch "cache"
        $warehouse = Join-Path $scratch "warehouse"
        New-Item -ItemType Directory -Path $warehouse -Force | Out-Null
        $protectedWorkbook = Join-Path $warehouse "WH.TEST.invSys.Data.Inventory.xlsb"
        [IO.File]::WriteAllText($protectedWorkbook, "unrelated warehouse authority")
        $beforeHash = (Get-FileHash -LiteralPath $protectedWorkbook -Algorithm SHA256).Hash

        & (Join-Path $repo "tools/publish_invsys_release.ps1") -SourceRoot $source -ReleaseRoot $feed -ReleaseId "test-r1" -GitCommit "test-commit" | Out-Null
        $manifest = Get-Content -LiteralPath (Join-Path $feed "Releases/test-r1/release-manifest.json") -Raw | ConvertFrom-Json
        $manifestValid = ($manifest.releaseId -eq "test-r1" -and @($manifest.packages).Count -eq 5)
        Add-Check "Slice4bf.Publish.ImmutableManifest" $manifestValid "Publisher creates a five-package R1-5 release manifest."

        New-Item -Path $registryPath -Force | Out-Null
        Set-ItemProperty -Path $registryPath -Name "OPEN" -Value '"C:\\ThirdParty.xlam"' -Type String
        & (Join-Path $repo "tools/update_invsys_station.ps1") -ReleaseRoot $feed -CacheRoot $cache -ExcelProcessName "powershell" -ExcelOptionsKey $registryPath -AddinManagerKey ($registryPath + "\\Manager") | Out-Null
        $deferred = (Get-Content -LiteralPath (Join-Path $cache "update-status.json") -Raw | ConvertFrom-Json).status -eq "DEFERRED_EXCEL_OPEN"
        Add-Check "Slice4bf.Update.DefersWithoutRegistration" $deferred "Updater defers while the supplied Excel process probe is present."

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo "tools/update_invsys_station.ps1") -ReleaseRoot $feed -CacheRoot $cache -ExcelOptionsKey $registryPath -AddinManagerKey ($registryPath + "\\Manager") | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Direct PowerShell station updater invocation failed." }
        $registered = Get-ItemProperty -Path $registryPath
        $openValues = @($registered.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } | ForEach-Object { [string]$_.Value })
        $stationApplied = ((Get-Content -LiteralPath (Join-Path $cache "current-release.json") -Raw | ConvertFrom-Json).releaseId -eq "test-r1")
        $nonMutation = ((Get-FileHash -LiteralPath $protectedWorkbook -Algorithm SHA256).Hash -eq $beforeHash)
        Add-Check "Slice4bf.Update.HashVerifiedRegistration" ($stationApplied -and (@($openValues | Where-Object { $_ -match 'invSys\.Operations\.xlam' }).Count -eq 1) -and (@($openValues | Where-Object { $_ -match 'invSys\.Admin\.xlam' }).Count -eq 1)) "Updater stages and registers only the two leaf add-ins after hash verification."
        Add-Check "Slice4bf.Update.PreservesAuthorityAndThirdPartyAddins" ($nonMutation -and (@($openValues | Where-Object { $_ -match 'ThirdParty\.xlam' }).Count -eq 1)) "Updater neither touches warehouse authority nor removes unrelated Excel add-ins."

        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $rollbackRejected = $false
        try { & (Join-Path $repo "tools/rollback_invsys_station_release.ps1") -ReleaseId "test-r1" -CacheRoot $cache -ExcelOptionsKey $registryPath -AddinManagerKey ($registryPath + "\\Manager") -ReasonCode "ApprovedCorrectiveAction" -ConfirmRollback | Out-Null }
        catch { $rollbackRejected = $true }
        Add-Check "Slice4bf.Rollback.WindowsAdministratorGate" (($isAdmin -and -not $rollbackRejected) -or ((-not $isAdmin) -and $rollbackRejected)) "Rollback executes only for a local Windows administrator and only against a cached verified release."

        & (Join-Path $repo "tools/publish_invsys_release.ps1") -SourceRoot $source -ReleaseRoot $feed -ReleaseId "test-r2" -GitCommit "test-commit" | Out-Null
        Add-Content -LiteralPath (Join-Path $feed "Releases/test-r2/invSys.Admin.xlam") -Value "corrupt"
        $rejected = $false
        try { & (Join-Path $repo "tools/update_invsys_station.ps1") -ReleaseRoot $feed -CacheRoot $cache -ExcelOptionsKey $registryPath -AddinManagerKey ($registryPath + "\\Manager") | Out-Null }
        catch { $rejected = $true }
        $stillTestR1 = ((Get-Content -LiteralPath (Join-Path $cache "current-release.json") -Raw | ConvertFrom-Json).releaseId -eq "test-r1")
        Add-Check "Slice4bf.Update.RejectsTamperedReleaseAndKeepsKnownGood" ($rejected -and $stillTestR1) "A bad remote hash cannot replace the registered known-good release."

        $taskPreview = & (Join-Path $repo "tools/register_invsys_update_task.ps1") -ReleaseRoot $feed -CacheRoot $cache
        $taskUsesLocalAgent = (($taskPreview -join "`n") -match [regex]::Escape((Join-Path (Split-Path -Parent $cache) "Deployment\\Agents"))) -and (($taskPreview -join "`n") -notmatch [regex]::Escape($repo))
        Add-Check "Slice4bf.Task.LocalAgentNoGitDependency" $taskUsesLocalAgent "Task preview must target a local hash-verified station agent, never a repository checkout."

        & (Join-Path $repo "tools/publish_invsys_release.ps1") -SourceRoot $source -ReleaseRoot $feed -ReleaseId "test-r3" -GitCommit "test-commit" | Out-Null
        $failedRegistration = $false
        try { & (Join-Path $repo "tools/update_invsys_station.ps1") -ReleaseRoot $feed -CacheRoot $cache -RegisterScriptPath (Join-Path $repo "tests/fixtures/Slice4bfFailingRegistration.ps1") -ExcelOptionsKey $registryPath -AddinManagerKey ($registryPath + "\\Manager") | Out-Null }
        catch { $failedRegistration = $true }
        $restoredPointer = ((Get-Content -LiteralPath (Join-Path $cache "current-release.json") -Raw | ConvertFrom-Json).releaseId -eq "test-r1")
        $restoredOpenValues = @((Get-ItemProperty -Path $registryPath).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } | ForEach-Object { [string]$_.Value })
        $restoredRegistration = (@($restoredOpenValues | Where-Object { $_ -match 'invSys\.Operations\.xlam' -and $_ -notmatch '\\Broken\\' }).Count -eq 1)
        Add-Check "Slice4bf.Update.RegistrationFailureRestoresKnownGood" ($failedRegistration -and $restoredPointer -and $restoredRegistration) "A registration failure restores the prior leaf registration and known-good pointer."
    }
    finally {
        if (Test-Path -LiteralPath $registryPath) { Remove-Item -LiteralPath $registryPath -Recurse -Force }
        if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Recurse -Force }
    }
}

$results | Format-Table -AutoSize
$failed = @($results | Where-Object { -not $_.Passed })
Write-Host ("Slice 4bf D16 deployment tooling: {0} passed, {1} failed" -f ($results.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
