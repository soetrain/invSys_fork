[CmdletBinding()]
param(
    [string]$RepoRoot = '',
    [string]$DeployRoot = 'deploy/current',
    [ValidateSet('RED','GREEN')][string]$Phase = 'GREEN',
    [switch]$CaptureEvidence,
    [switch]$CheckActivityEvidence,
    [switch]$CheckActivityFoundation,
    [switch]$CheckShippingActivity,
    [switch]$CheckReceivingActivity,
    [switch]$CheckReceivingStagingActivity,
    [switch]$CheckReceivingLocalActivity,
    [switch]$CheckReceivingLifecycleActivity,
    [switch]$CheckReceivingNavigationActivity,
    [switch]$ReceivingNavigationOnly,
    [switch]$CheckReceivingSurfaceCoverage,
    [switch]$ReceivingSurfaceOnly,
    [switch]$CheckReceivingNativeSurface,
    [switch]$CheckReceivingWorksheetActivity,
    [switch]$CheckReceivingWorksheetScenarios,
    [switch]$CheckReceivingWorksheetGuards,
    [switch]$CheckReceivingLauncherDenial,
    [switch]$ReceivingLauncherDenialOnly,
    [switch]$CaptureDenialDialogs,
    [switch]$ReceivingLifecycleOnly,
    [ValidateSet('None','SkipTerminationEvidence','KeepLauncherReference')]
    [string]$LifecycleDiagnostic = 'None'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($RepoRoot)) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deploy = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Close Excel before isolated packaged validation.' }
$runRoot = Join-Path ([IO.Path]::GetTempPath()) ('invsys-config-command-' + [guid]::NewGuid().ToString('N'))
$reportRoot = Join-Path $repo 'reports/runtime/config-commands'
if ($CheckActivityEvidence) {
    $reportRoot = Join-Path $repo 'reports/runtime/slice4be-activity'
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    if ($CheckActivityFoundation) { . (Join-Path $PSScriptRoot 'Slice4beActivityFoundation.ps1') }
}
if ($CheckActivityFoundation -and -not $CheckActivityEvidence) { throw 'Foundation checks require activity evidence mode.' }
if ($CheckShippingActivity -and (-not $CheckActivityFoundation -or $CheckReceivingActivity)) { throw 'Shipping activity requires the foundation and a separate run from Receiving.' }
if ($CheckShippingActivity) {
    . (Join-Path $PSScriptRoot 'Slice4beShippingActivity.ps1')
    $reportRoot = Join-Path $repo ('reports/runtime/slice4be-shipping-activity/'+[guid]::NewGuid().ToString('N'))
}
if ($CheckReceivingStagingActivity -and -not $CheckReceivingActivity) { throw 'Staging coverage requires Receiving activity mode.' }
if ($CheckReceivingLocalActivity -and -not $CheckReceivingStagingActivity) { throw 'Local-action coverage requires staging activity mode.' }
if ($CheckReceivingLifecycleActivity -and -not $CheckReceivingLocalActivity) { throw 'Lifecycle coverage requires the preserved local-action baseline.' }
if ($CheckReceivingNavigationActivity -and (-not $CheckReceivingLifecycleActivity -or $ReceivingLifecycleOnly)) { throw 'Navigation coverage requires the full preserved lifecycle baseline.' }
if ($ReceivingNavigationOnly -and -not $CheckReceivingNavigationActivity) { throw 'Navigation-only diagnosis requires navigation coverage.' }
if ($CheckReceivingSurfaceCoverage -and -not $CheckReceivingNavigationActivity) { throw 'Surface coverage requires the preserved navigation baseline.' }
if ($ReceivingSurfaceOnly -and (-not $CheckReceivingSurfaceCoverage -or $ReceivingNavigationOnly)) { throw 'Surface-only diagnosis requires surface coverage without another diagnostic-only mode.' }
if ($CheckReceivingNativeSurface -and -not $ReceivingSurfaceOnly) { throw 'Native worksheet discovery requires the separate surface-only run.' }
if ($CheckReceivingWorksheetActivity -and -not $CheckReceivingNativeSurface) { throw 'Worksheet activity requires calibrated native surface coverage.' }
if ($CheckReceivingWorksheetScenarios -and -not $CheckReceivingWorksheetActivity) { throw 'Worksheet scenarios require the protected native activity baseline.' }
if ($CheckReceivingWorksheetGuards -and -not $CheckReceivingWorksheetActivity) { throw 'Worksheet guards require the protected native activity baseline.' }
if ($CheckReceivingLauncherDenial -and -not $CheckReceivingNavigationActivity) { throw 'Launcher denial coverage requires the preserved navigation baseline.' }
if ($ReceivingLauncherDenialOnly -and (-not $CheckReceivingLauncherDenial -or $ReceivingSurfaceOnly -or $ReceivingNavigationOnly)) { throw 'Launcher-denial-only diagnosis requires its coverage without another diagnostic-only mode.' }
if ($CaptureDenialDialogs -and -not $ReceivingLauncherDenialOnly) { throw 'Native denial dialog evidence uses the separate focused run.' }
if ($ReceivingLifecycleOnly -and -not $CheckReceivingLifecycleActivity) { throw 'Lifecycle-only diagnosis requires lifecycle coverage.' }
if ($LifecycleDiagnostic -ne 'None' -and -not $ReceivingLifecycleOnly) { throw 'Mutation diagnostics require the separate lifecycle-only report.' }
if ($LifecycleDiagnostic -ne 'None' -and $Phase -ne 'RED') { throw 'Diagnostic mutations cannot be run as acceptance GREEN.' }
if ($CheckReceivingActivity) {
    $reportRoot = Join-Path $repo 'reports/runtime/slice4be-receiving-activity'
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingActivity.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingReferences.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingRetry.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingStaging.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingLocal.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingFreshness.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingLifecycle.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingNavigation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingSurface.ps1')
    if ($CheckReceivingNativeSurface) {
        . (Join-Path $PSScriptRoot 'Slice4beReceivingNativeSurface.ps1')
        if ($CheckReceivingWorksheetScenarios) { . (Join-Path $PSScriptRoot 'Slice4beReceivingWorksheetScenarios.ps1') }
        if ($CheckReceivingWorksheetGuards) { . (Join-Path $PSScriptRoot 'Slice4beReceivingWorksheetGuards.ps1') }
        $reportRoot=Join-Path $reportRoot ('native-surface-'+[guid]::NewGuid().ToString('N'))
        if ($CheckReceivingWorksheetActivity) { $reportRoot=Join-Path $reportRoot 'worksheet-activity' }
    }
    . (Join-Path $PSScriptRoot 'Slice4beReceivingLauncherDenial.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingDenialDialogs.ps1')
    if (-not $CheckActivityFoundation) { . (Join-Path $PSScriptRoot 'Slice4beActivityFoundation.ps1') }
}
New-Item -ItemType Directory -Path $runRoot,$reportRoot -Force | Out-Null
$results = [Collections.Generic.List[object]]::new()
$excel = $null
$step = 'startup'
$settingsRoot = 'HKCU:\Software\VB and VBA Program Settings\invSys'
$registryBefore = @{}
if (Test-Path -LiteralPath $settingsRoot) {
    foreach ($key in @(Get-Item -LiteralPath $settingsRoot) + @(Get-ChildItem -LiteralPath $settingsRoot -Recurse)) {
        $values = @{}
        foreach ($name in $key.GetValueNames()) { $values[$name] = @($key.GetValue($name),$key.GetValueKind($name)) }
        $registryBefore[$key.Name] = $values
    }
}
function Check([string]$Name,[bool]$Passed) {
    $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output ("{0}: {1}" -f $Name, $(if($Passed){'PASS'}else{'FAIL'}))
}
function CaptureFormEvidence([string]$Title,[string]$FileName) {
    if (-not ('InvSysSettingsCapture' -as [type])) {
    Add-Type -ReferencedAssemblies System.Drawing @'
using System; using System.Drawing; using System.Runtime.InteropServices;
public static class InvSysSettingsCapture {
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
    public static void Save(string title, string path) {
        var hwnd=FindWindow(null,title); Rect r;
        if(hwnd==IntPtr.Zero || !GetWindowRect(hwnd,out r)) throw new Exception("Requested form window unavailable.");
        using(var bitmap=new Bitmap(r.Right-r.Left,r.Bottom-r.Top)) {
            using(var graphics=Graphics.FromImage(bitmap)) {
                var hdc=graphics.GetHdc(); bool captured;
                try { captured=PrintWindow(hwnd,hdc,2); } finally { graphics.ReleaseHdc(hdc); }
                if(!captured) graphics.CopyFromScreen(r.Left,r.Top,0,0,bitmap.Size);
            }
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
    }
}
'@
    }
    [InvSysSettingsCapture]::Save($Title,(Join-Path $reportRoot $FileName))
}
function Run([string]$Package,[string]$Macro,[object[]]$Values=@()) {
    $name="'$Package'!$Macro"
    switch($Values.Count) {
        0 { $excel.Run($name) }
        1 { $excel.Run($name,$Values[0]) }
        2 { $excel.Run($name,$Values[0],$Values[1]) }
        3 { $excel.Run($name,$Values[0],$Values[1],$Values[2]) }
        4 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3]) }
        5 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3],$Values[4]) }
        6 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3],$Values[4],$Values[5]) }
        default { throw 'Unsupported macro argument count' }
    }
}
function Table($Workbook,[string]$Name) {
    foreach ($sheet in $Workbook.Worksheets) {
        foreach ($candidate in $sheet.ListObjects) { if($candidate.Name -eq $Name){return $candidate} }
    }
    throw "Fixture table missing: $Name"
}
function CredentialHash([string]$Secret) {
    [double]$acc=5381
    for($i=0;$i -lt $Secret.Length;$i++) { $acc=($acc*33 + [int][char]$Secret[$i] + $i+1) % 2147483647 }
    '{0:X8}' -f [int]$acc
}
function SelectTarget($Fixture,[string]$User='config-admin') {
    [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($Fixture.Root))
    $selected = [string](Run 'invSys.Core.xlam' 'modNasConnection.SelectWarehouseTargetForAutomation' @($Fixture.Root,$Fixture.Root,'S1',$false))
    if(-not $selected.StartsWith('OK|')){throw 'Fixture target selection failed.'}
    [void](Run 'invSys.Core.xlam' 'modNasConnection.SetCurrentTargetPathsForTest' @('\\fixture-host\config-command',$Fixture.Root))
    $signed = [string](Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @($User,$Fixture.Secret,''))
    if(-not $signed.StartsWith('OK|')){
        $code = if ($signed -match '^FAIL\|([0-9]+|NO_TARGET|ERROR)(?:\||$)') { $Matches[1] } else { 'UNAVAILABLE' }
        throw ('Fixture sign-in failed; status code='+$code+'.')
    }
}
function NewFixture([string]$Suffix) {
    $wh='WHD5'+[guid]::NewGuid().ToString('N').Substring(0,6).ToUpperInvariant()
    $root=Join-Path $runRoot $Suffix
    $share=Join-Path $runRoot ($Suffix+'-share')
    New-Item -ItemType Directory -Path $share -Force | Out-Null
    [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($root))
    $created=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.BootstrapWarehouseLocalAdmin' @($wh,'Config command fixture','S1','config-admin',$root,$share))
    if(-not $created){throw 'Admin Generate Warehouse fixture failed.'}
    $secret=[guid]::NewGuid().ToString('N')
    $auth=$excel.Workbooks.Open((Join-Path $root ($wh+'.invSys.Auth.xlsb')),0,$false)
    $users=Table $auth 'tblUsers'
    $caps=Table $auth 'tblCapabilities'
    # Fixture credentials are text. Excel must not coerce a randomly generated
    # all-numeric value or drop its leading zero; never emit either value.
    $users.ListColumns.Item('PinHash').Range.NumberFormat='@'
    foreach($row in $users.ListRows) {
        if($row.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -eq 'config-admin') {
            $row.Range.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2=CredentialHash $secret
            if ([string]$row.Range.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2 -cne (CredentialHash $secret)) { throw 'Fixture credential text did not round-trip.' }
        }
    }
    foreach($identity in @('config-reader','config-producer')) {
        $row=$users.ListRows.Add()
        foreach($pair in @{UserId=$identity;DisplayName='Config fixture';PinHash=(CredentialHash $secret);Status='Active'}.GetEnumerator()) {
            $row.Range.Cells.Item(1,$users.ListColumns.Item($pair.Key).Index).Value2=$pair.Value
        }
        if ([string]$row.Range.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2 -cne (CredentialHash $secret)) { throw 'Fixture credential text did not round-trip.' }
        $row=$caps.ListRows.Add()
        $cap=if($identity -eq 'config-producer'){'PROD_POST'}else{'RECEIVE_POST'}
        foreach($pair in @{UserId=$identity;Capability=$cap;WarehouseId=$wh;StationId='S1';Status='Active'}.GetEnumerator()) {
            $row.Range.Cells.Item(1,$caps.ListColumns.Item($pair.Key).Index).Value2=$pair.Value
        }
    }
    $auth.Save(); $auth.Close($false)
    [pscustomobject]@{Root=$root;Warehouse=$wh;Secret=$secret;Config=(Join-Path $root ($wh+'.invSys.Config.xlsb'))}
}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    $step='load packages'
    $packages=@{}
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')) {
        $packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name),0,$true)
    }
    if ($CheckActivityEvidence) {
        Check 'Activity.ProducingPackageIdentity' (Test-Slice4bePackageIdentity $deploy)
    }
    # Test-only instrumentation in the unsaved Admin project: invokes the exact
    # existing form selection/save handlers without changing their implementation.
    $formCode=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmAdminSettings').CodeModule
    $formCode.AddFromString(@'
Public Function D5TestSave(ByVal keyName As String, ByVal valueText As String) As String
    Dim i As Long
    For i = 0 To mLstConfig.ListCount - 1
        If StrComp(CStr(mLstConfig.List(i, 0)), keyName, vbTextCompare) = 0 Then
            mLstConfig.ListIndex = i
            mLstConfig_Click
            mTxtConfigValue.Value = valueText
            mBtnSaveConfig_Click
            D5TestSave = mLblStatus.Caption
            Exit Function
        End If
    Next i
    Err.Raise 5, , "Fixture key missing"
End Function
'@)
    $testModule=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Add(1)
    $testModule.Name='TestD5Commands'
    $testModule.CodeModule.AddFromString(@'
Option Explicit
Private mForm As frmAdminSettings
Private mLastStatus As String
Public Sub OpenSettings()
    Set mForm = New frmAdminSettings
End Sub
Public Sub ShowSettings()
    mForm.Show vbModeless
    mForm.Repaint
End Sub
Public Sub CloseSettings()
    If Not mForm Is Nothing Then Unload mForm
    Set mForm = Nothing
End Sub
Public Function SaveSettings(ByVal key As String, ByVal value As String) As Boolean
    mLastStatus = mForm.D5TestSave(key, value)
    SaveSettings = (InStr(1, mLastStatus, "saved", vbTextCompare) > 0)
End Function
Public Function LastStatus() As String
    LastStatus = mLastStatus
End Function
Public Function SaveDirect(ByVal key As String, ByVal value As String, ByVal wh As String, ByVal st As String) As Boolean
    Dim report As String
    SaveDirect = modConfig.UpdateConfigValue(key, value, report, wh, st)
End Function
Public Function PublishUom() As Boolean
    Dim rows As Variant, report As String
    rows = modUomSettings.GetUomCatalogRows()
    rows(3, 6) = Not CBool(rows(3, 6))
    PublishUom = modUomSettings.PublishUomCatalogRows(rows, report)
End Function
'@)
    $productionCode=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    $productionCode.AddFromString(@'
Public Function D5UomRoundTrip(ByVal workbookName As String) As Boolean
    Dim prior As Workbook, lo As ListObject, version As Long
    On Error GoTo Done
    If Not mBuilt Then BuildLayout
    Set prior = mOperatorWorkbook
    Set mOperatorWorkbook = Application.Workbooks(workbookName)
    version = modConfig.GetLong("UomConversionCatalogVersion", 1)
    mBtnUomCatalogSend_Click
    Set lo = mOperatorWorkbook.Worksheets("invSys UOM Catalog").ListObjects("tblInvSysUomCatalog")
    lo.DataBodyRange.Cells(3, 6).Value2 = Not CBool(lo.DataBodyRange.Cells(3, 6).Value2)
    lo.Parent.Activate
    lo.DataBodyRange.Cells(1, 1).Select
    mBtnUomCatalogRetrieve_Click
    D5UomRoundTrip = (mOperatorWorkbook.Worksheets("invSys UOM Catalog").ListObjects.Count = 0 _
        And modConfig.GetLong("UomConversionCatalogVersion", 1) = version + 1)
Done:
    Set mOperatorWorkbook = prior
End Function
'@)
    $productionTest=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Add(1)
    $productionTest.Name='TestD5Uom'
    $productionTest.CodeModule.AddFromString(@'
Private mHeld As frmProduction
Public Sub HoldForm()
    Set mHeld = New frmProduction
    Load mHeld
End Sub
Public Function HeldRoundTrip(ByVal workbookName As String) As Boolean
    HeldRoundTrip = mHeld.D5UomRoundTrip(workbookName)
    Unload mHeld
    Set mHeld = Nothing
End Function
Public Function RoundTrip(ByVal workbookName As String) As Boolean
    RoundTrip = frmProduction.D5UomRoundTrip(workbookName)
    Unload frmProduction
End Function
'@)
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $repo 'deploy/current/templates')))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    $step='Admin-generated fixtures'; Write-Output $step
    $a=NewFixture 'a'; $b=NewFixture 'b'
    SelectTarget $a
    $step='unauthenticated command'
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','777',$a.Warehouse,'S1'))
    Check 'Command.SignedOutDenied' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    SelectTarget $a 'config-reader'
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','778',$a.Warehouse,'S1'))
    Check 'Command.MissingCapabilityDenied' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    SelectTarget $a
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    if ($CheckActivityEvidence) { $activityBefore = @(Get-Slice4beActivityFiles $a) }
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','601'))
    Check 'Settings.RealSaveHandler' ($ok -and [long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('BatchSize',0)) -eq 601)
    if ($CheckActivityEvidence) {
        if (-not $ok) { throw 'Activity fixture Settings action failed.' }
        Test-Slice4beObservedAction $a $activityBefore 'ADMIN_SETTINGS_SAVE_VALUE' `
            'CONFIG_SAVE_REQUESTED' 'CONFIG_SAVE_COMPLETED' 'Changed' 'Info' `
            'config-admin' 'Activity.AdminSettings'
        Test-Slice4beUnavailableStore $a
    }
    if($CaptureEvidence){
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
        Start-Sleep -Milliseconds 300
        CaptureFormEvidence 'invSys Settings' 'settings-save.png'
    }
    SelectTarget $b
    $beforeA=(Get-FileHash -LiteralPath $a.Config).Hash; $beforeB=(Get-FileHash -LiteralPath $b.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','602'))
    Check 'Settings.StaleCapturedTargetDenied' (-not $ok -and $beforeA -eq (Get-FileHash -LiteralPath $a.Config).Hash -and $beforeB -eq (Get-FileHash -LiteralPath $b.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    SelectTarget $a
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','not-a-number',$a.Warehouse,'S1'))
    Check 'Command.InvalidTypeNoWrite' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('WarehouseId','different',$a.Warehouse,'S1'))
    Check 'Command.IdentityImmutable' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    $step='read non-mutation'
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $column=$table.ListColumns.Add(); $column.Name='Operator Extra'; $column.DataBodyRange.Value2='preserve'
    $table.ListColumns.Item('Timezone').Delete()
    $cfg.Save(); $count=$table.ListColumns.Count
    $ok=[bool](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($a.Warehouse,'S1'))
    Check 'Read.OptionalHeaderNotRepaired' ($ok -and $table.ListColumns.Count -eq $count -and $cfg.Saved)
    Check 'Read.UnknownColumnPreserved' ($table.ListColumns.Item('Operator Extra').DataBodyRange.Cells.Item(1,1).Value2 -eq 'preserve')
    $cfg.Close($false)
    SelectTarget $a 'config-producer'
    $version=[long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('UomConversionCatalogVersion',1))
    if ($CheckActivityEvidence) { $activityBefore = @(Get-Slice4beActivityFiles $a) }
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PublishUom')
    Check 'Production.ValidatedUomRouteRetained' ($ok -and [long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('UomConversionCatalogVersion',1)) -eq $version+1)
    if ($CheckActivityEvidence) {
        $activityAfter = @(Get-Slice4beActivityFiles $a)
        Check 'Activity.DirectServiceIsNotUserControl' (@($activityAfter | Where-Object { $_ -notin $activityBefore }).Count -eq 0)
    }
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','603',$a.Warehouse,'S1'))
    Check 'Production.ArbitraryConfigDenied' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    $stage=$excel.Workbooks.Add()
    if ($CheckActivityEvidence) { $activityBefore = @(Get-Slice4beActivityFiles $a) }
    $ok=[bool](Run 'invSys.Operations.xlam' 'TestD5Uom.RoundTrip' @($stage.Name))
    Check 'Production.RealUomRetrieveHandler' $ok
    if ($CheckActivityEvidence) {
        if (-not $ok) { throw 'Activity fixture Production action failed.' }
        Test-Slice4beObservedAction $a $activityBefore 'PRODUCTION_UOM_RETRIEVE' `
            'UOM_RETRIEVE_REQUESTED' 'UOM_RETRIEVE_COMPLETED' 'Changed' 'Info' `
            'config-producer' 'Activity.ProductionRetrieve'
    }
    $stage.Close($false)
    SelectTarget $a 'config-reader'
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $stage=$excel.Workbooks.Add()
    if ($CheckActivityEvidence) { $activityBefore = @(Get-Slice4beActivityFiles $a) }
    $ok=[bool](Run 'invSys.Operations.xlam' 'TestD5Uom.RoundTrip' @($stage.Name))
    Check 'Production.DeniedRetrievePreservesStaging' (-not $ok -and $stage.Worksheets.Item('invSys UOM Catalog').ListObjects.Count -eq 1 -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    if ($CheckActivityEvidence) {
        if ($ok) { throw 'Activity fixture denied action unexpectedly succeeded.' }
        Test-Slice4beObservedAction $a $activityBefore 'PRODUCTION_UOM_RETRIEVE' `
            'UOM_RETRIEVE_REQUESTED' 'UOM_RETRIEVE_DENIED' 'Unchanged' 'Blocked' `
            'config-reader' 'Activity.ProductionDenied'
    }
    $stage.Close($false)
    SelectTarget $a
    $cfg=$excel.Workbooks.Open($a.Config,0,$true)
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','604',$a.Warehouse,'S1'))
    Check 'Command.ReadOnlyDenied' (-not $ok -and $cfg.Saved)
    $cfg.Close($false)
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $extra=$table.ListColumns.Item('Operator Extra').DataBodyRange.Cells.Item(1,1)
    $extra.Value2='unsaved user edit'
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','605',$a.Warehouse,'S1'))
    Check 'Command.DirtyWorkbookPreserved' (-not $ok -and -not $cfg.Saved -and $extra.Value2 -eq 'unsaved user edit')
    $cfg.Close($false)
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($a.Warehouse,'S1'))
    Check 'Read.ClosedWorkbookBytesPreserved' ($ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    if ($CheckActivityFoundation) { Test-Slice4beActivityFoundation $a $b }
    if ($CheckShippingActivity) {
        $step='Shipping activity through packaged form handlers'
        Test-Slice4beShippingActivity
        SelectTarget $a
    }
    if ($CheckReceivingActivity) {
        $step='Receiving activity through packaged form handlers'
        Test-Slice4beReceivingActivity
        SelectTarget $a
    }
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $table.ListColumns.Item('WarehouseName').Delete()
    $cfg.Save(); $cfg.Close($false)
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','607',$a.Warehouse,'S1'))
    Check 'Command.OtherRequiredHeaderDenied' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $column=$table.ListColumns.Add(); $column.Name='WarehouseName'; $column.DataBodyRange.Value2='Config command fixture'
    $cfg.Save(); $cfg.Close($false)
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $table.ListColumns.Item('WarehouseId').Delete()
    $cfg.Save(); $count=$table.ListColumns.Count
    $ok=[bool](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($a.Warehouse,'S1'))
    Check 'Read.RequiredHeaderNotRepaired' (-not $ok -and $table.ListColumns.Count -eq $count -and $cfg.Saved)
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','606',$a.Warehouse,'S1'))
    Check 'Command.MissingIdentityDenied' (-not $ok -and $table.ListColumns.Count -eq $count -and $cfg.Saved)
    $cfg.Close($false)
}
catch {
    Check ('Harness.Exception.'+$step) $false
    $message=$_.Exception.Message -replace '(?i)[A-Z]:\\[^\r\n"'']+', '<path>'
    Write-Output $message
}
finally {
    if($null -ne $excel) {
        try { [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings') } catch {}
        try { foreach($book in @($excel.Workbooks)){ $book.Close($false) }; $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if(Test-Path -LiteralPath $settingsRoot) {
        foreach($key in @(Get-Item -LiteralPath $settingsRoot) + @(Get-ChildItem -LiteralPath $settingsRoot -Recurse)) {
            $path='Registry::'+$key.Name
            foreach($name in $key.GetValueNames()) {
                if(-not $registryBefore.ContainsKey($key.Name) -or -not $registryBefore[$key.Name].ContainsKey($name)) {
                    Remove-ItemProperty -LiteralPath $path -Name $name
                }
            }
        }
    }
    foreach($path in $registryBefore.Keys) {
        foreach($name in $registryBefore[$path].Keys) {
            $saved=$registryBefore[$path][$name]
            New-ItemProperty -LiteralPath ('Registry::'+$path) -Name $name -Value $saved[0] -PropertyType $saved[1] -Force | Out-Null
        }
    }
    # Runtime credentials stay only in disposable generated authority fixtures.
    $resolved=[IO.Path]::GetFullPath($runRoot)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'invsys-config-command-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
    $reportName=$Phase.ToLowerInvariant()+'.json'
    if ($CheckReceivingNavigationActivity) { $reportName='navigation-'+$reportName }
    if ($ReceivingNavigationOnly) { $reportName='diagnostic-'+$reportName }
    if ($CheckReceivingSurfaceCoverage) { $reportName='surface-'+$Phase.ToLowerInvariant()+'.json' }
    if ($ReceivingSurfaceOnly) { $reportName='diagnostic-'+$reportName }
    if ($CheckReceivingLauncherDenial) { $reportName='launcher-denial-'+$Phase.ToLowerInvariant()+'.json' }
    if ($ReceivingLauncherDenialOnly) { $reportName='diagnostic-'+$reportName }
    if ($ReceivingLifecycleOnly) { $reportName='lifecycle-only-'+$reportName }
    if ($LifecycleDiagnostic -ne 'None') { $reportName='diagnostic-'+$LifecycleDiagnostic.ToLowerInvariant()+'.json' }
    $results | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $reportRoot $reportName)
}
$failed=@($results | Where-Object { -not $_.Passed }).Count
Write-Output ("$Phase : {0} passed, {1} failed" -f ($results.Count-$failed),$failed)
if($failed -gt 0){exit 1}
