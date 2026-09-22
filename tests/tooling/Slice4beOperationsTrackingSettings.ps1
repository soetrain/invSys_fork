# D18 Operations entry through the packaged Viewer callback, with no Admin XLAM.
function Test-OperationsTrackingSettings($Fixture) {
    Check 'OpsSettings.AdminAbsentFromPackageDirectory' (-not (Test-Path (Join-Path $operationsSettingsDeploy 'invSys.Admin.xlam')) -and @(Get-ChildItem -LiteralPath $operationsSettingsDeploy -Filter '*.xlam').Count -eq 4)
    $rolePackages = @()
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')) {
        $book = $excel.Workbooks.Item($name)
        $rolePackages += [pscustomobject]@{Name=$book.Name;FullName=$book.FullName}
    }
    $expectedDirectory = [IO.Path]::GetFullPath($operationsSettingsDeploy)
    $matchingPaths = @($rolePackages | Where-Object { [IO.Path]::GetFullPath((Split-Path $_.FullName -Parent)) -ieq $expectedDirectory }).Count
    [pscustomobject]@{XlamCount=$rolePackages.Count;CopiedPackagePaths=$matchingPaths;AdminLoaded=(Test-LoadedPackage 'invSys.Admin.xlam')} | ConvertTo-Json | Set-Content (Join-Path $reportRoot 'operations-package-location.json')
    Check 'OpsSettings.FourPackagesLoadedFromIsolatedDirectory' ($rolePackages.Count -eq 4 -and $matchingPaths -eq 4 -and -not (Test-LoadedPackage 'invSys.Admin.xlam'))
    SelectTarget $Fixture 'config-reader'
    $before = (Get-FileHash -LiteralPath $Fixture.Config).Hash
    $module = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $module.AddFromString(@'
Public Function OperationsSettingsEntryObserved() As Boolean
    Dim control As Object
    If mInventoryViewer Is Nothing Then Exit Function
    For Each control In mInventoryViewer.Controls
        If TypeName(control) = "CommandButton" Then
            If control.Name = "btnSettings" And control.Caption = "Settings" Then
                OperationsSettingsEntryObserved = control.Visible And control.Enabled
                Exit Function
            End If
        End If
    Next control
End Function
Public Function OperationsViewerObserved() As Boolean
    If Not mInventoryViewer Is Nothing Then OperationsViewerObserved = mInventoryViewer.Visible
End Function
Public Function OperationsSettingsLoadedFormsForTest() As Long
    OperationsSettingsLoadedFormsForTest = VBA.UserForms.Count
End Function
Public Function OperationsViewerCaptionForTest() As String
    OperationsViewerCaptionForTest = mInventoryViewer.Caption
End Function
Public Function OperationsSettingsCaptionForTest() As String
    OperationsSettingsCaptionForTest = OperationsSettingsFormForTest().Caption
End Function
Public Function OperationsViewerWindowForTest() As Double
    mInventoryViewer.Repaint
    DoEvents
    OperationsViewerWindowForTest = CDbl(modUserFormResizeWin.GetUserFormWindowHandle(mInventoryViewer))
End Function
Public Function OperationsSettingsWindowForTest() As Double
    Dim form As Object
    Set form = OperationsSettingsFormForTest()
    form.Repaint
    DoEvents
    OperationsSettingsWindowForTest = CDbl(modUserFormResizeWin.GetUserFormWindowHandle(form))
End Function
Public Sub CloseOperationsViewerForTest()
    If Not mInventoryViewer Is Nothing Then Unload mInventoryViewer
End Sub
Private Function OperationsSettingsFormForTest() As Object
    Dim instance As Object
    For Each instance In VBA.UserForms
        If TypeName(instance) = "frmEventTrackingSettings" Then Set OperationsSettingsFormForTest = instance: Exit Function
    Next instance
End Function
Private Function OperationsSettingsControlForTest(ByVal parent As Object, ByVal name As String) As Object
    Dim control As Object, page As Object, found As Object
    If parent Is Nothing Then Exit Function
    For Each control In parent.Controls
        If control.Name = name Then Set OperationsSettingsControlForTest = control: Exit Function
        If TypeName(control) = "MultiPage" Then
            For Each page In control.Pages
                Set found = OperationsSettingsControlForTest(page, name)
                If Not found Is Nothing Then Set OperationsSettingsControlForTest = found: Exit Function
            Next page
        End If
    Next control
End Function
Private Function OperationsSettingsFitForTest(ByVal parent As Object) As Boolean
    Dim control As Object, page As Object
    If parent Is Nothing Then Exit Function
    For Each control In parent.Controls
        If control.Visible Then
            If control.Left < 0 Or control.Top < 0 Or control.Left + control.Width > parent.InsideWidth + 1 Or _
               control.Top + control.Height > parent.InsideHeight + 1 Then Exit Function
        End If
        If TypeName(control) = "MultiPage" Then
            For Each page In control.Pages
                If page.Index = control.Value Then
                    If Not OperationsSettingsFitForTest(page) Then Exit Function
                End If
            Next page
        End If
    Next control
    OperationsSettingsFitForTest = True
End Function
Public Function OperationsSettingsSurfaceForTest() As String
    Dim form As Object, rows As Object, choice As Object, index As Long, choices As String, actions As Boolean
    Set form = OperationsSettingsFormForTest()
    If form Is Nothing Then OperationsSettingsSurfaceForTest = "False|False|False|False": Exit Function
    Set rows = OperationsSettingsControlForTest(form, "lstReadOnlyTrackingPolicy")
    Set choice = OperationsSettingsControlForTest(form, "cmbPreferredActionPathView")
    If Not choice Is Nothing Then
        For index = 0 To choice.ListCount - 1
            choices = choices & "|" & choice.List(index)
        Next index
    End If
    actions = (choices = "|Use warehouse default|How-To|Diagnostic|Compare both") And _
        Not OperationsSettingsControlForTest(form, "btnSaveMyPreference") Is Nothing And _
        Not OperationsSettingsControlForTest(form, "btnResetMyPreference") Is Nothing And _
        Not OperationsSettingsControlForTest(form, "btnReloadMyPreference") Is Nothing
    OperationsSettingsSurfaceForTest = CStr(form.Visible) & "|"
    If rows Is Nothing Then
        OperationsSettingsSurfaceForTest = OperationsSettingsSurfaceForTest & "False"
    Else
        OperationsSettingsSurfaceForTest = OperationsSettingsSurfaceForTest & CStr(rows.Locked And rows.ListCount > 0)
    End If
    OperationsSettingsSurfaceForTest = OperationsSettingsSurfaceForTest & "|" & CStr(actions) & "|" & CStr(OperationsSettingsFitForTest(form))
End Function
'@)
    $viewerCode = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    if($viewerCode.Lines(1,$viewerCode.CountOfLines) -match 'Private Sub mBtnSettings_Click\(') {
        $viewerCode.AddFromString(@'
Private mOpsRememberedState As String
Public Sub OperationsPressSettingsForTest()
    mBtnSettings_Click
End Sub
Private Function OpsViewerStateForTest() As String
    Dim row As Long, column As Long, value As String
    OpsViewerStateForTest = CStr(mTabs.Value) & "|" & CStr(mTxtSearch.Value) & "|" & CStr(mLstInventory.ListIndex) & "|" & CStr(mColumnCount) & "|" & mWarehouseId
    If IsEmpty(mRows) Then Exit Function
    For row = LBound(mRows, 1) To UBound(mRows, 1)
        For column = LBound(mRows, 2) To UBound(mRows, 2)
            value = CStr(mRows(row, column))
            OpsViewerStateForTest = OpsViewerStateForTest & "|" & CStr(Len(value)) & ":" & value
        Next column
    Next row
End Function
Public Sub OpsRememberViewerStateForTest(ByVal tabIndex As Long)
    mTabs.Value = tabIndex
    If mTxtSearch.Visible Then mTxtSearch.Value = "settings preservation"
    If mLstInventory.ListCount > 0 Then mLstInventory.ListIndex = 0
    mOpsRememberedState = OpsViewerStateForTest()
End Sub
Public Function OpsViewerStatePreservedForTest() As Boolean
    OpsViewerStatePreservedForTest = (mOpsRememberedState = OpsViewerStateForTest())
End Function
'@)
        $module.AddFromString(@'
Public Sub OperationsPressSettingsForTest()
    mInventoryViewer.OperationsPressSettingsForTest
End Sub
Public Sub OpsRememberViewerStateForTest(ByVal tabIndex As Long)
    mInventoryViewer.OpsRememberViewerStateForTest tabIndex
End Sub
Public Function OpsViewerStatePreservedForTest() As Boolean
    OpsViewerStatePreservedForTest = mInventoryViewer.OpsViewerStatePreservedForTest()
End Function
'@)
    }
    $settingsCode = $null
    foreach($component in $packages['invSys.Operations.xlam'].VBProject.VBComponents) {
        if($component.Name -eq 'frmEventTrackingSettings') { $settingsCode = $component.CodeModule; break }
    }
    if($null -ne $settingsCode) {
        $settingsCode.AddFromString(@'
Private mOpsTestSaveEntries As Long
Public Sub OpsTestChoose(ByVal choice As String)
    mChoice.Value = choice
End Sub
Public Function OpsTestChoice() As String
    OpsTestChoice = CStr(mChoice.Value)
End Function
Public Function OpsTestEntries() As Long
    OpsTestEntries = mOpsTestSaveEntries
End Function
Public Sub OpsTestReset()
    mReset_Click
End Sub
Public Sub OpsTestReload()
    mReload_Click
End Sub
Public Sub OpsTestClose()
    mClose_Click
End Sub
Public Sub OpsTestResize(ByVal width As Double, ByVal height As Double)
    Me.Width = width: Me.Height = height
    mLayout.ApplyAnchoredLayout
End Sub
Public Function OpsTestPolicyDefaults() As Boolean
    Dim index As Long, command As Boolean, navigation As Boolean
    For index = 0 To mPolicyRows.ListCount - 1
        If mPolicyRows.List(index, 6) <> "Built-in defaults" Then Exit Function
        If mPolicyRows.List(index, 0) = "RECEIVING_CONFIRM_WRITES" Then command = (mPolicyRows.List(index, 3) = "Yes")
        If mPolicyRows.List(index, 0) = "RECEIVING_PAGE_RECEIPTS" Then navigation = (mPolicyRows.List(index, 3) = "No")
    Next index
    OpsTestPolicyDefaults = command And navigation And InStr(1, mPolicyStatus.Caption, "Built-in tracking defaults", vbBinaryCompare) > 0 And _
        InStr(1, mEffective.Caption, "policy version 0", vbBinaryCompare) > 0 And InStr(1, mEvidence.Caption, "capture is off", vbBinaryCompare) > 0
End Function
Public Function OpsTestPolicyHeaders() As Boolean
    Dim index As Long, widths As Variant, captions As Variant, x As Double, header As Object
    widths = Split(mPolicyRows.ColumnWidths, ";")
    captions = Array("Operation", "Control", "Collect", "Viewer", "Sequence", "Availability")
    x = mPolicyRows.Left + 3
    On Error GoTo Missing
    For index = 1 To 6
        Set header = mPolicyRows.Parent.Controls("lblPolicyColumn" & CStr(index))
        If header.Caption <> captions(index - 1) Or Abs(header.Left - x) > 1 Or Abs(header.Width - Val(widths(index))) > 1 Then Exit Function
        x = x + Val(widths(index))
    Next index
    ' MSForms quantizes the requested 140 points to 139.95 on this display.
    OpsTestPolicyHeaders = (Abs(Val(widths(1)) - 140) <= 1)
Missing:
End Function
'@)
        $settingsCode.InsertLines($settingsCode.ProcBodyLine('SaveMyPreference',0)+1, '    mOpsTestSaveEntries = mOpsTestSaveEntries + 1')
        $manager = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modOperationsTrackingSettings').CodeModule
        $manager.AddFromString(@'
Public Sub OpsTestChoose(ByVal choice As String)
    mSettings.OpsTestChoose choice
End Sub
Public Function OpsTestChoice() As String
    OpsTestChoice = mSettings.OpsTestChoice()
End Function
Public Function OpsTestEntries() As Long
    OpsTestEntries = mSettings.OpsTestEntries()
End Function
Public Sub OpsTestReset()
    mSettings.OpsTestReset
End Sub
Public Sub OpsTestReload()
    mSettings.OpsTestReload
End Sub
Public Function OpsTestSave() As Boolean
    OpsTestSave = mSettings.SaveMyPreference()
End Function
Public Sub OpsTestClose()
    mSettings.OpsTestClose
End Sub
Public Function OpsTestIsClosed() As Boolean
    OpsTestIsClosed = (mSettings Is Nothing)
End Function
Public Function OpsTestContext(ByVal context As String) As Boolean
    If Not mSettings Is Nothing Then OpsTestContext = mSettings.HasContext(context)
End Function
Public Sub OpsTestResize(ByVal width As Double, ByVal height As Double)
    mSettings.OpsTestResize width, height
End Sub
Public Function OpsTestPolicyDefaults() As Boolean
    OpsTestPolicyDefaults = mSettings.OpsTestPolicyDefaults()
End Function
Public Function OpsTestPolicyHeaders() As Boolean
    OpsTestPolicyHeaders = mSettings.OpsTestPolicyHeaders()
End Function
'@)
        $core = $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActionPathPreference').CodeModule
        $core.AddFromString(@'
Public Function OpsSettingsTestKey(ByVal context As String) As String
    Dim key As String, target As WarehouseTarget, report As String
    If PreferenceKey(context, key, target, report) Then OpsSettingsTestKey = key
End Function
Public Function OpsSettingsTestWarehouseWrite(ByVal context As String) As Boolean
    Dim report As String
    OpsSettingsTestWarehouseWrite = modTrackingPolicySettings.SavePolicy(context, 0, modTrackingPolicySettings.DefaultRequest(), report)
End Function
'@)
    }
    if($CompileEvaluationProbesForTest){
        . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
        Compile-Slice4beEvaluationProbes -PackageNames @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam') -PackageRoot $operationsSettingsDeploy -CheckPrefix 'Harness.OperationsSettingsCompile.'
        $noForms=[long](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsSettingsLoadedFormsForTest') -eq 0
        Check 'Harness.OperationsSettingsProbesInstalledBeforeForms' $noForms
        if(-not $noForms){throw 'Operations Settings probes must precede forms; not product RED.'}
    }
    $visible = $excel.Visible
    try {
        $excel.Visible = $true
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
        Check 'OpsSettings.PublicViewerCallbackOpensForNonAdmin' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsViewerObserved'))
        $entry = [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsSettingsEntryObserved')
        Check 'OpsSettings.SettingsEntryReachable' $entry
        if($entry) { [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsPressSettingsForTest') }
        $flags = ([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsSettingsSurfaceForTest')).Split('|')
        if($flags.Count -ne 4 -or @($flags | Where-Object { $_ -cnotin @('True','False') }).Count) { throw 'Invalid Operations surface facts.' }
        $names = @('SettingsFormOpened','ReadOnlyPolicyDisplayed','PersonalChoicesAndActions','SettingsLayoutFits')
        for($index=0;$index -lt $names.Count;$index++) { Check ('OpsSettings.'+$names[$index]) ($flags[$index] -ceq 'True') }
        if($flags[0] -ceq 'True') { Check 'OpsSettings.ApprovedFormTitle' ([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsSettingsCaptionForTest') -ceq 'Event Tracking Settings') }
        if(@($flags | Where-Object { $_ -ceq 'False' }).Count -eq 0) { Test-OperationsPreferenceActions $Fixture }
        Check 'OpsSettings.ViewerOpenPreservesConfig' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        # Observe the existing native windows after all actual-handler checks.
        # Capture must not dispatch another VBA Repaint/DoEvents callback.
        if($CaptureEvidence) { CaptureOwnedFormByCaptionEvidence ('Viewer - '+$Fixture.Warehouse) 'operations-viewer-settings-entry.png' }
        if($CaptureEvidence -and $flags[0] -ceq 'True') { CaptureOwnedFormByCaptionEvidence 'Event Tracking Settings' 'operations-event-tracking-settings.png' }
    } finally {
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseOperationsViewerForTest')
        $excel.Visible = $visible
    }
}

function Test-OperationsPreferenceActions($Fixture) {
    $context = [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $prior = [string](Run 'invSys.Core.xlam' 'PreferenceRestartProbe.ReadChoice' @($context))
    Check 'OpsSettings.PolicyDefaultsAndPreferenceShareVersion' ([bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestPolicyDefaults'))
    Check 'OpsSettings.PolicyHeadersAlignedAndReadable' ([bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestPolicyHeaders'))
    foreach($tab in @(0,1,2)) {
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpsRememberViewerStateForTest' @($tab))
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsPressSettingsForTest')
        Check ('OpsSettings.OpenPreservesViewerTab'+$tab) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpsViewerStatePreservedForTest'))
    }
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpsRememberViewerStateForTest' @(0))
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoose' @('Diagnostic'))
    Check 'OpsSettings.SelectionStagesOnly' ([string](Run 'invSys.Core.xlam' 'PreferenceRestartProbe.ReadChoice' @($context)) -ceq $prior)
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsPressSettingsForTest')
    Check 'OpsSettings.RepeatedOpenRetainsStaging' ([string](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoice') -ceq 'Diagnostic')
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestReload')
    Check 'OpsSettings.ReloadRestoresSavedChoice' ([string](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoice') -ceq $prior)
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoose' @('Compare both'))
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestReset')
    Check 'OpsSettings.ResetStagesDefaultOnly' ([string](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoice') -ceq 'Use warehouse default' -and [string](Run 'invSys.Core.xlam' 'PreferenceRestartProbe.ReadChoice' @($context)) -ceq $prior)
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoose' @('Diagnostic'))
    $entries = [int](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestEntries')
    $ok = [bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestSave')
    Check 'OpsSettings.RealSaveHandlerEntered' ([int](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestEntries') -eq $entries+1)
    Check 'OpsSettings.PersonalSavePersistsForNonAdmin' ($ok -and [string](Run 'invSys.Core.xlam' 'PreferenceRestartProbe.ReadChoice' @($context)) -ceq 'Diagnostic')
    if(-not $ok) { return }
    Test-OperationsPreferenceGuards $Fixture $context
}

function Test-OperationsPreferenceGuards($Fixture,$Context) {
    $before = (Get-FileHash -LiteralPath $Fixture.Config).Hash
    $key = [string](Run 'invSys.Core.xlam' 'modActionPathPreference.OpsSettingsTestKey' @($Context))
    $ok = [bool](Run 'invSys.Core.xlam' 'modActionPathPreference.OpsSettingsTestWarehouseWrite' @($Context))
    Check 'OpsSettings.PersonalAccessDoesNotGrantPolicyWrite' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    foreach($size in @(@(860,700),@(744,640))) {
        [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestResize' $size)
        $flags = ([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsSettingsSurfaceForTest')).Split('|')
        Check ('OpsSettings.ResizeFits'+$size[0]) ($flags[3] -ceq 'True')
    }
    Test-OperationsSettingsNativeLayout
    $sentinel = $excel.Workbooks.Add()
    try {
        $sentinel.Worksheets.Item(1).Range('B2').Value2 = 'preserve unrelated workbook'
        $sentinel.Activate()
        [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoose' @('How-To'))
        $ok = [bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestSave')
        Check 'OpsSettings.UnrelatedWorkbookDoesNotRedirectSave' ($ok -and (Read-PreferenceRegistry $key) -ceq 'How-To' -and $sentinel.Worksheets.Item(1).Range('B2').Value2 -ceq 'preserve unrelated workbook' -and -not $sentinel.Saved)
    } finally { $sentinel.Close($false) }
    Check 'OpsSettings.SavePreservesViewerState' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpsViewerStatePreservedForTest'))
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoose' @('Compare both'))
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestClose')
    Check 'OpsSettings.CloseReleasesCapturedInstance' ([bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestIsClosed'))
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsPressSettingsForTest')
    Check 'OpsSettings.CloseDiscardsUnsavedChoice' ([string](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoice') -ceq 'How-To')
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoose' @('Compare both'))
    SelectTarget $Fixture 'config-producer'
    $otherContext = [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $otherKey = [string](Run 'invSys.Core.xlam' 'modActionPathPreference.OpsSettingsTestKey' @($otherContext))
    $ok = [bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestSave')
    Check 'OpsSettings.ChangedSessionCannotSaveOrLoseStaging' (-not $ok -and (Read-PreferenceRegistry $key) -ceq 'How-To' -and $null -eq (Read-PreferenceRegistry $otherKey) -and [string](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoice') -ceq 'Compare both')
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsPressSettingsForTest')
    Check 'OpsSettings.ReboundViewerCannotRetargetOpenSettings' ([bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestContext' @($Context)) -and -not [bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestContext' @($otherContext)))
    [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestClose')
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsPressSettingsForTest')
    Check 'OpsSettings.CloseReopenBindsCurrentContext' ([bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestContext' @($otherContext)) -and [string](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoice') -ceq 'Use warehouse default')
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    $ok = [bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestSave')
    Check 'OpsSettings.SignedOutSaveDenied' (-not $ok -and $null -eq (Read-PreferenceRegistry $otherKey))
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseOperationsViewerForTest')
    Check 'OpsSettings.ViewerCloseClosesSettings' ([bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestIsClosed'))
    SelectTarget $Fixture 'config-reader'
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsPressSettingsForTest')
    Check 'OpsSettings.ContextAndConfigRemainIntact' ([string](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestChoice') -ceq 'How-To' -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and -not (Test-LoadedPackage 'invSys.Admin.xlam'))
}

function Test-OperationsSettingsNativeLayout {
    if(-not ('OpsSettingsNativeLayout' -as [type])) {
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class OpsSettingsNativeLayout {
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int command);
    [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr h);
}
'@
    }
    $handle = [long](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsSettingsWindowForTest')
    if(-not $handle) { throw 'Settings has no verified native window.' }
    foreach($command in @(3,9)) {
        [void][OpsSettingsNativeLayout]::ShowWindow([IntPtr]$handle,$command)
        Start-Sleep -Milliseconds 200
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsSettingsWindowForTest')
        $flags = ([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.OperationsSettingsSurfaceForTest')).Split('|')
        $maximized = [OpsSettingsNativeLayout]::IsZoomed([IntPtr]$handle)
        $name = if($command -eq 3){'Maximize'}else{'Restore'}
        Check ('OpsSettings.Native'+$name+'Fits') ($maximized -eq ($command -eq 3) -and $flags[3] -ceq 'True' -and [bool](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.OpsTestPolicyHeaders'))
        if($CaptureEvidence) { CaptureOwnedFormEvidence 'Event Tracking Settings' ('operations-settings-'+$name.ToLowerInvariant()+'.png') $handle }
    }
}
