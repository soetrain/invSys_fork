[CmdletBinding()]
param(
    [string]$DeployRoot = 'deploy/validation-production-next-activity-01',
    [ValidateSet('RED','GREEN')][string]$Phase = 'RED',
    [switch]$CaptureEvidence,
    [switch]$CheckLayout
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Close Excel before this isolated packaged test.' }
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$deploy = (Resolve-Path -LiteralPath $DeployRoot).Path
$root = Join-Path 'reports/runtime/warehouse-purpose' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null
$root = (Resolve-Path -LiteralPath $root).Path
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('invSys-purpose-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture | Out-Null
$pins = @(Get-ChildItem -LiteralPath $deploy -Filter '*.xlam' -File | ForEach-Object {
    [pscustomobject]@{ File = $_.FullName; Hash = (Get-FileHash -LiteralPath $_.FullName).Hash }
})
if ($pins.Count -ne 5) { throw 'Five frozen packages required.' }
$pins | ConvertTo-Json | Set-Content (Join-Path $root 'package-pins.json')
$snapshot = Get-InvSysTestSettingsSnapshot
$checks = [Collections.Generic.List[object]]::new()
$start = [DateTimeOffset]::UtcNow
$excel = $null; $books = @{}; $harnessComplete = $false; $failure = $null
$ownedProcessId = 0
$lastMacro = ''
if ($CaptureEvidence) {
    $reportRoot = $root
    $TraceGuideResourcesForTest = $false
    $GuideCaptureVisibleExcelForTest = $false
    $GuideCaptureSavedWorkbookForTest = $false
    $tokens = $null; $parseErrors = $null
    $tree = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1'), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw 'Existing capture helper must parse.' }
    foreach ($name in @('Initialize-SettingsCapture','CaptureFormEvidence','CaptureOwnedFormEvidence','CaptureOwnedFormByCaptionEvidence')) {
        $fn = $tree.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq $name}, $true)
        if ($null -eq $fn) { throw 'Existing capture helper missing.' }
        Invoke-Expression $fn.Extent.Text
    }
}
function Check([string]$Name, [bool]$Pass) {
    $checks.Add([pscustomobject]@{ Name = $Name; Pass = $Pass })
}
function Run([string]$Book, [string]$Macro, [object[]]$Arguments = @()) {
    $script:lastMacro = $Book + '|' + $Macro
    $entry = "'$Book'!$Macro"
    switch ($Arguments.Count) {
        0 { return $excel.Run($entry) }
        1 { return $excel.Run($entry, $Arguments[0]) }
        3 { return $excel.Run($entry, $Arguments[0], $Arguments[1], $Arguments[2]) }
        default { throw 'Unsupported test argument count.' }
    }
}
function Read-Purpose([string]$Path) {
    $book = $excel.Workbooks.Open($Path, 0, $true)
    try {
        $table = $book.Worksheets.Item('WarehouseConfig').ListObjects.Item('tblWarehouseConfig')
        $purpose = ''; $purposeColumns = 0
        foreach ($column in $table.ListColumns) {
            if ($column.Name.Trim() -ieq 'WarehousePurpose') {
                $purposeColumns++
                $purpose = [string]$table.DataBodyRange.Cells.Item(1, $column.Index).Value2
            }
        }
        [pscustomobject]@{ Purpose = $purpose; Columns = $purposeColumns }
    } finally { $book.Close($false) }
}
function Check-PurposeLayout([string]$CaseName) {
    $geometry = [string](Run 'invSys.Admin.xlam' 'modPurposeTest.Geometry')
    $geometry | Set-Content (Join-Path $root ($CaseName+'-geometry.txt'))
    $rects = @{}
    foreach ($line in $geometry.Split("`n")) {
        $parts = $line.Trim().Split('|')
        $rects[$parts[0]] = [pscustomobject]@{Left=[double]$parts[1];Top=[double]$parts[2];Width=[double]$parts[3];Height=[double]$parts[4]}
    }
    $summary = $rects['lblSummary']; $bounds = $rects['Form']
    $clear = $true
    foreach ($name in @('txtPathSharePoint','lblPathSharePointError','chkPublishInitial','cboWarehousePurpose','lblWarehousePurpose')) {
        if ($rects.ContainsKey($name)) { $r=$rects[$name]; $clear=$clear -and ($summary.Top -ge $r.Top+$r.Height+4) }
    }
    Check "$CaseName.SummaryBelowInputs" $clear
    Check "$CaseName.FooterBelowSummary" ($rects['btnOK'].Top -ge $summary.Top+$summary.Height+4 -and $rects['btnCancel'].Top -ge $summary.Top+$summary.Height+4)
    $within = $true
    foreach ($name in $rects.Keys) {
        $r=$rects[$name];$within=$within -and ($r.Left -ge 0 -and $r.Top -ge 0 -and $r.Left+$r.Width -le $bounds.Width+1 -and $r.Top+$r.Height -le $bounds.Height+1)
    }
    Check "$CaseName.ControlsWithinForm" $within
}
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PurposeTestProcess {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
Write-Output ('Warehouse purpose evidence: ' + $root)
try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$processIdValue = 0
    [void][PurposeTestProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$processIdValue)
    $ownedProcessId = $processIdValue
    $excel.Visible = $true; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    foreach ($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')) {
        $books[$name] = $excel.Workbooks.Open((Join-Path $deploy $name), 0, $true)
        foreach ($reference in $books[$name].VBProject.References) {
            if ($reference.IsBroken) { throw 'Broken package reference is a harness failure.' }
            if ($reference.Name -like 'invSys_*' -and (Split-Path -Parent $reference.FullPath) -ine $deploy) {
                throw 'Package dependency escaped the frozen candidate.'
            }
        }
    }
    if ($CaptureEvidence) {
        $displayBook = $excel.Workbooks.Add()
        $displayBook.SaveAs((Join-Path $fixture 'purpose-display.xlsx'), 51)
        $displayBook.Activate()
    }
    # In-memory test adapters only: actual packaged operator handlers do the work.
    $form = $books['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmCreateWarehouse').CodeModule
    $form.AddFromString(@'
Public Function PurposePrepareForTest(ByVal warehouseId As String, ByVal rootPath As String, ByVal purpose As String) As Boolean
    Dim ctl As Object
    Me.txtWarehouseId.Value = warehouseId
    Me.txtWarehouseName.Value = "Training purpose fixture"
    Me.txtStationId.Value = "S1"
    Me.txtAdminUser.Value = "purpose-tester"
    Me.txtPathLocal.Value = rootPath
    Me.txtPathSharePoint.Value = ""
    Me.chkPublishInitial.Value = False
    For Each ctl In Me.Controls
        If ctl.Name = "cboWarehousePurpose" Then
            If purpose <> "" Then ctl.Value = purpose
            PurposePrepareForTest = True
        End If
    Next ctl
End Function
Public Function PurposeFactsForTest() As String
    Dim ctl As Object
    PurposeFactsForTest = "MISSING"
    For Each ctl In Me.Controls
        If ctl.Name = "cboWarehousePurpose" Then
            PurposeFactsForTest = CStr(ctl.Value) & "|" & CStr(ctl.ListCount) & "|" & CStr(ctl.Style)
        End If
    Next ctl
End Function
Public Function PurposeCreateForTest() As Boolean
    btnOK_Click
    PurposeCreateForTest = (Me.Tag = "COMPLETE")
End Function
Public Sub PurposeCancelForTest()
    btnCancel_Click
End Sub
Public Function PurposeGeometryForTest() As String
    Dim ctl As Object
    PurposeGeometryForTest = "Form|0|0|" & CStr(Me.InsideWidth) & "|" & CStr(Me.InsideHeight)
    For Each ctl In Me.Controls
        If ctl.Visible Then PurposeGeometryForTest = PurposeGeometryForTest & vbLf & ctl.Name & "|" & CStr(ctl.Left) & "|" & CStr(ctl.Top) & "|" & CStr(ctl.Width) & "|" & CStr(ctl.Height)
    Next ctl
End Function
'@)
    $adapter = $books['invSys.Admin.xlam'].VBProject.VBComponents.Add(1)
    $adapter.Name = 'modPurposeTest'
    $adapter.CodeModule.AddFromString(@'
Option Explicit
Public Function Prepare(ByVal warehouseId As String, ByVal rootPath As String, ByVal purpose As String) As Boolean
    Load frmCreateWarehouse
    Prepare = frmCreateWarehouse.PurposePrepareForTest(warehouseId, rootPath, purpose)
    frmCreateWarehouse.Show vbModeless
End Function
Public Function Facts() As String
    Facts = frmCreateWarehouse.PurposeFactsForTest()
End Function
Public Function Create() As Boolean
    Create = frmCreateWarehouse.PurposeCreateForTest()
End Function
Public Sub CloseForm()
    frmCreateWarehouse.PurposeCancelForTest
End Sub
Public Function Geometry() As String
    Geometry = frmCreateWarehouse.PurposeGeometryForTest()
End Function
Public Sub Expand()
    frmCreateWarehouse.Width = 800
    frmCreateWarehouse.Height = 700
End Sub
'@)
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $deploy 'templates')))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $fixture 'operators')))
    foreach ($case in @('Default','Training')) {
        $warehouseId = 'WP' + [guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
        $target = Join-Path $fixture $case
        $choice = if ($case -eq 'Training') { 'Training' } else { '' }
        $expected = if ($case -eq 'Training') { 'Training' } else { 'Operational' }
        $present = [bool](Run 'invSys.Admin.xlam' 'modPurposeTest.Prepare' @($warehouseId,$target,$choice))
        Check "$case.PurposeChoicePresent" $present
        Check "$case.PurposeChoiceClosedList" ((Run 'invSys.Admin.xlam' 'modPurposeTest.Facts') -ceq ($expected + '|2|2'))
        if ($CheckLayout) {
            Check-PurposeLayout ($case+'.Minimum')
            [void](Run 'invSys.Admin.xlam' 'modPurposeTest.Expand')
            Check-PurposeLayout ($case+'.Expanded')
        }
        if ($CaptureEvidence) {
            [string](Run 'invSys.Admin.xlam' 'modPurposeTest.Geometry') | Set-Content (Join-Path $root ($case.ToLowerInvariant()+'-geometry.txt'))
            Initialize-SettingsCapture
            CaptureOwnedFormByCaptionEvidence 'Create Warehouse' ($case.ToLowerInvariant()+'.png')
        }
        $created = [bool](Run 'invSys.Admin.xlam' 'modPurposeTest.Create')
        Check "$case.ActualCreateHandlerCompleted" $created
        # Fixture creation must succeed even on RED; no missing-fixture RED allowed.
        if (-not $created) { throw 'Actual creation failed; this is not purpose-contract RED.' }
        [void](Run 'invSys.Admin.xlam' 'modPurposeTest.CloseForm')
        $configPath = Join-Path $target ($warehouseId + '.invSys.Config.xlsb')
        if (-not (Test-Path -LiteralPath $configPath)) { throw 'Missing generated Config fixture.' }
        $saved = Read-Purpose $configPath
        Check "$case.OneManagedPurposeColumn" ($saved.Columns -eq 1)
        Check "$case.PurposeSurvivesReopen" ($saved.Purpose -ceq $expected)
        Check "$case.OwnInventory" (Test-Path -LiteralPath (Join-Path $target ($warehouseId + '.invSys.Data.Inventory.xlsb')))
        Check "$case.OwnInbox" (Test-Path -LiteralPath (Join-Path $target 'inbox'))
        Check "$case.OwnOutbox" (Test-Path -LiteralPath (Join-Path $target 'outbox'))
        $operatorPath = [string](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.GetLastWarehouseOperatorWorkbookPath')
        Check "$case.OwnOperatorWorkbook" ($operatorPath.StartsWith($fixture, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $operatorPath))
        # Re-enter through a NEW real form, selecting the opposite purpose.
        $before = (Get-FileHash -LiteralPath $configPath).Hash
        $opposite = if ($case -eq 'Training') { 'Operational' } else { 'Training' }
        [void](Run 'invSys.Admin.xlam' 'modPurposeTest.Prepare' @($warehouseId,$target,$opposite))
        Check "$case.ExistingRuntimeRefused" (-not [bool](Run 'invSys.Admin.xlam' 'modPurposeTest.Create'))
        [void](Run 'invSys.Admin.xlam' 'modPurposeTest.CloseForm')
        Check "$case.ExistingConfigBytePreserved" ((Get-FileHash -LiteralPath $configPath).Hash -ceq $before)
    }
    $cancelRoot = Join-Path $fixture 'cancel'
    [void](Run 'invSys.Admin.xlam' 'modPurposeTest.Prepare' @(('WC'+[guid]::NewGuid().ToString('N').Substring(0,8)),$cancelRoot,'Training'))
    [void](Run 'invSys.Admin.xlam' 'modPurposeTest.CloseForm')
    Check 'Cancel.DoesNotProvision' (-not (Test-Path -LiteralPath $cancelRoot))
    $harnessComplete = $true
} catch {
    $failure = Get-RecordingFailureFacts $_ $lastMacro
    $failure | Add-Member NoteProperty TestLine $_.InvocationInfo.ScriptLineNumber
} finally {
    if ($null -ne $excel) {
        try { [void](Run 'invSys.Admin.xlam' 'modPurposeTest.CloseForm') } catch {}
        try { foreach ($book in @($excel.Workbooks)) { $book.Close($false) }; $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    $books = $null; $form = $null; $adapter = $null; $book = $null; $reference = $null; $displayBook = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($ownedProcessId -gt 0) {
        $owned = Get-Process -Id $ownedProcessId -ErrorAction SilentlyContinue
        if ($null -ne $owned -and -not $owned.WaitForExit(15000)) { Wait-RecordingCleanup -Creator $null -Worker $null }
    }
    $restored = Restore-InvSysTestSettingsSnapshot $snapshot
    $preserved = @($pins | Where-Object { (Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash }).Count -eq 0
    $failed = @($checks | Where-Object { -not $_.Pass }).Count
    [pscustomobject]@{
        Phase=$Phase; StartUTC=$start.ToString('o'); EndUTC=[DateTimeOffset]::UtcNow.ToString('o')
        HarnessComplete=$harnessComplete; Failure=$failure; Pass=$checks.Count-$failed; Fail=$failed; Checks=$checks
        SettingsRestored=$restored; PackagesPreserved=$preserved; ExcelClosed=(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0)
        B0Accepted=$false; ReleaseAccepted=$false
    } | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
    if (-not $restored -or -not $preserved) { throw 'Fixture preservation failed.' }
}
Write-Output ('Purpose checks: {0} PASS / {1} FAIL; harness complete: {2}' -f ($checks.Count-$failed),$failed,$harnessComplete)
if (-not $harnessComplete) { throw 'Harness failed; inspect sanitized local receipt.' }
if ($failed -gt 0) { exit 1 }
exit 0
