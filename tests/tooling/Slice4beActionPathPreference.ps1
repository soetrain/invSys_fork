# D18 personal preferences: inspect the real packaged Settings surface.
# Registry/configuration values and runtime identities remain in the fixture.
function Install-Slice4beActionPathPreferenceProbe($TestModule,$FormCode) {
    $script:preferenceSaveVerified = $false
    $TestModule.CodeModule.AddFromString(@'
Public Function PreferenceSurface() As String
    Dim page As Object, sections As Object, child As Object, views As Object
    Set page = TrackingPage(mForm, "Event Tracking")
    Set sections = TrackingControl(page, "MultiPage", "", "mpEventTracking")
    If sections Is Nothing Then Err.Raise 5
    TrackingSettingsSelectPage "Event Tracking"
    For Each child In sections.Pages
        If child.Caption = "Action Paths" Then
            sections.Value = child.Index
            Set page = child
            Exit For
        End If
    Next child
    mForm.Show vbModeless
    mForm.Repaint
    Set views = TrackingControl(page, "ComboBox", "", "cmbPreferredActionPathView")
    PreferenceSurface = CStr(TrackingViewChoices(views)) & "|" & _
        CStr(Not TrackingControl(page, "CommandButton", "Save My Preference", "btnSaveActionPathPreference") Is Nothing And _
             Not TrackingControl(page, "CommandButton", "Reload", "btnReloadActionPathPreference") Is Nothing And _
             Not TrackingControl(page, "CommandButton", "Reset to Default", "btnResetActionPathPreference") Is Nothing) & "|" & _
        CStr(Not TrackingControl(page, "Label", "", "lblActionPathEffectiveView") Is Nothing And _
             Not TrackingControl(page, "Label", "", "lblActionPathEvidence") Is Nothing)
End Function
'@)
    $class = $null
    foreach($component in $packages['invSys.Admin.xlam'].VBProject.VBComponents) {
        if($component.Name -eq 'cAdminActionPathPreference') { $class = $component.CodeModule; break }
    }
    if($null -eq $class) { return }
    $TestModule.CodeModule.AddFromString(@'
Private mPreferenceInitStage As Long
Public Sub NotePreferenceInitStage(ByVal stage As Long)
    mPreferenceInitStage = stage
End Sub
'@)
    $openLine = $TestModule.CodeModule.ProcBodyLine('OpenSettings',0)
    $TestModule.CodeModule.InsertLines($openLine+1, '    On Error GoTo PreferenceOpenFailed')
    $openEnd = $openLine + 2
    while($TestModule.CodeModule.Lines($openEnd,1).Trim() -ne 'End Sub') { $openEnd++ }
    $TestModule.CodeModule.InsertLines($openEnd, "    Exit Sub`r`nPreferenceOpenFailed:`r`n    Err.Raise 5, , ""Settings initialization error "" & CStr(Err.Number) & "" at preference stage "" & CStr(mPreferenceInitStage)")
    $observedCode = $class.Lines(1,$class.CountOfLines)
    $observedCode = $observedCode.Replace('    mContext = context', "    mContext = context`r`n    TestD5Commands.NotePreferenceInitStage 1")
    $observedCode = $observedCode.Replace('    mChoice.Style = fmStyleDropDownList', "    mChoice.Style = fmStyleDropDownList`r`n    TestD5Commands.NotePreferenceInitStage 2")
    $observedCode = $observedCode.Replace('    Next choice', "    Next choice`r`n    TestD5Commands.NotePreferenceInitStage 3")
    $observedCode = $observedCode.Replace('    ReloadPreference', "    TestD5Commands.NotePreferenceInitStage 4`r`n    ReloadPreference")
    $observedCode = $observedCode.Replace('    mLoading = True', "    TestD5Commands.NotePreferenceInitStage 5`r`n    mLoading = True")
    $observedCode = $observedCode.Replace('    mLoading = False', "    TestD5Commands.NotePreferenceInitStage 6`r`n    mLoading = False")
    $observedCode = $observedCode.Replace('    mStatus.Caption = report', "    mStatus.Caption = report`r`n    TestD5Commands.NotePreferenceInitStage 7")
    $class.DeleteLines(1,$class.CountOfLines)
    $class.AddFromString($observedCode)
    $class.AddFromString(@'
Private mPreferenceTestEntries As Long
Public Function PreferenceTestChoice() As String
    PreferenceTestChoice = CStr(mChoice.Value)
End Function
Public Sub PreferenceTestChoose(ByVal value As String)
    mChoice.Value = value
End Sub
Public Function PreferenceTestEntries() As Long
    PreferenceTestEntries = mPreferenceTestEntries
End Function
Public Function PreferenceTestEvidenceOff() As Boolean
    PreferenceTestEvidenceOff = (InStr(1, mEvidence.Caption, "capture is off", vbBinaryCompare) > 0)
End Function
Public Function PreferenceTestEffective(ByVal expected As String) As Boolean
    PreferenceTestEffective = (InStr(1, mEffective.Caption, "Effective view: " & expected, vbBinaryCompare) = 1)
End Function
'@)
    $line = $class.ProcBodyLine('SavePreference',0)
    $class.InsertLines($line+2, '    mPreferenceTestEntries = mPreferenceTestEntries + 1')
    $FormCode.AddFromString(@'
Public Function PreferenceTestChoice() As String
    PreferenceTestChoice = mPreference.PreferenceTestChoice()
End Function
Public Sub PreferenceTestChoose(ByVal value As String)
    mPreference.PreferenceTestChoose value
End Sub
Public Function PreferenceTestSave() As Boolean
    PreferenceTestSave = mPreference.SavePreference()
End Function
Public Sub PreferenceTestReload()
    mPreference.ReloadPreference
End Sub
Public Sub PreferenceTestReset()
    mPreference.ResetPreference
End Sub
Public Function PreferenceTestEntries() As Long
    PreferenceTestEntries = mPreference.PreferenceTestEntries()
End Function
Public Function PreferenceTestEvidenceOff() As Boolean
    PreferenceTestEvidenceOff = mPreference.PreferenceTestEvidenceOff()
End Function
Public Function PreferenceTestEffective(ByVal expected As String) As Boolean
    PreferenceTestEffective = mPreference.PreferenceTestEffective(expected)
End Function
'@)
    $TestModule.CodeModule.AddFromString(@'
Public Function PreferenceChoice() As String
    PreferenceChoice = mForm.PreferenceTestChoice()
End Function
Public Sub PreferenceChoose(ByVal value As String)
    mForm.PreferenceTestChoose value
End Sub
Public Function PreferenceSave() As Boolean
    PreferenceSave = mForm.PreferenceTestSave()
End Function
Public Sub PreferenceReload()
    mForm.PreferenceTestReload
End Sub
Public Sub PreferenceReset()
    mForm.PreferenceTestReset
End Sub
Public Function PreferenceEntries() As Long
    PreferenceEntries = mForm.PreferenceTestEntries()
End Function
Public Function PreferenceEvidenceOff() As Boolean
    PreferenceEvidenceOff = mForm.PreferenceTestEvidenceOff()
End Function
Public Function PreferenceEffective(ByVal expected As String) As Boolean
    PreferenceEffective = mForm.PreferenceTestEffective(expected)
End Function
Public Function PreferenceSaveDirect(ByVal context As String, ByVal choice As String) As Boolean
    Dim report As String
    PreferenceSaveDirect = modActionPathPreference.SavePreference(context, choice, report)
End Function
Public Function PreferenceReadDirect(ByVal context As String) As String
    Dim choice As String, effective As String, evidence As String, report As String
    If modActionPathPreference.ReadPreference(context, choice, effective, evidence, report) Then PreferenceReadDirect = choice
End Function
'@)
    $core = $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActionPathPreference').CodeModule
    $core.AddFromString(@'
Public Function PreferenceTestKey(ByVal context As String) As String
    Dim target As WarehouseTarget, report As String, key As String
    If PreferenceKey(context, key, target, report) Then PreferenceTestKey = key
End Function
'@)
}

function Test-Slice4beActionPathPreference($Fixture,$Other) {
    $before = (Get-FileHash -LiteralPath $Fixture.Config).Hash
    $visible = $excel.Visible
    try {
        $excel.Visible = $true
        $observed = [string](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSurface')
        $flags = $observed.Split('|')
        if($flags.Count -ne 3 -or @($flags | Where-Object { $_ -cnotin @('True','False') }).Count) { throw 'Invalid preference surface observations.' }
        Check 'Preference.FourViewChoices' ($flags[0] -ceq 'True')
        Check 'Preference.SeparatePersonalActions' ($flags[1] -ceq 'True')
        Check 'Preference.EffectiveViewAndEvidenceStatus' ($flags[2] -ceq 'True')
        Check 'Preference.SectionFits' ([bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingSettingsLayoutFits'))
        Check 'Preference.OpenDoesNotWriteConfig' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        if(@($flags | Where-Object { $_ -ceq 'False' }).Count -eq 0) { Test-ActionPathPreferenceActions $Fixture $Other }
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSurface')
        if($CaptureEvidence) { CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'action-path-preference.png' }
    } finally { $excel.Visible = $visible }
}

function Get-PreferenceChoice { [string](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoice') }
function Read-PreferenceRegistry($Key) {
    $path = Join-Path $settingsRoot 'ActionPathPreferencesV1'
    if(-not (Test-Path -LiteralPath $path)) { return $null }
    (Get-Item -LiteralPath $path).GetValue($Key,$null)
}
function Test-ActionPathPreferenceActions($Fixture,$Other) {
    $before = (Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherBefore = (Get-FileHash -LiteralPath $Other.Config).Hash
    $context = [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $key = [string](Run 'invSys.Core.xlam' 'modActionPathPreference.PreferenceTestKey' @($context))
    if($key -notmatch '^U[0-9A-F]+W[0-9A-F]+$') { throw 'Invalid isolated preference key.' }
    Check 'Preference.InitialWarehouseDefault' ((Get-PreferenceChoice) -ceq 'Use warehouse default' -and $null -eq (Read-PreferenceRegistry $key))
    $default = (Get-TrackingPolicyRequest | ConvertFrom-Json).DefaultView
    Check 'Preference.EffectiveWarehousePolicy' ([bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceEffective' @($default)))
    Check 'Preference.CaptureOffShowsUnavailable' ([bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceEvidenceOff'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @('Diagnostic'))
    Check 'Preference.SelectionStagesOnly' ((Get-PreferenceChoice) -ceq 'Diagnostic' -and $null -eq (Read-PreferenceRegistry $key) -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceReload')
    Check 'Preference.ReloadDiscardsStaging' ((Get-PreferenceChoice) -ceq 'Use warehouse default' -and $null -eq (Read-PreferenceRegistry $key))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @('Compare both'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceReset')
    Check 'Preference.ResetStagesDefaultOnly' ((Get-PreferenceChoice) -ceq 'Use warehouse default' -and $null -eq (Read-PreferenceRegistry $key) -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    foreach($bad in @('','Unknown','diagnostic','Diagnostic ')) {
        $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSaveDirect' @($context,$bad))
        Check ('Preference.InvalidChoice'+[array]::IndexOf(@('','Unknown','diagnostic','Diagnostic '),$bad)) (-not $ok -and $null -eq (Read-PreferenceRegistry $key))
    }
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @('Diagnostic'))
    $entries = [int](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceEntries')
    $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
    Check 'Preference.RealSaveActionEntered' ([int](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceEntries') -eq $entries+1)
    $saved = $ok -and (Read-PreferenceRegistry $key) -ceq 'Diagnostic'
    Check 'Preference.SavePersistsLocalChoice' $saved
    if(-not $saved) { return }
    $script:preferenceSaveVerified = $true
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @('How-To'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceReload')
    Check 'Preference.ReloadRestoresSavedChoice' ((Get-PreferenceChoice) -ceq 'Diagnostic')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @('Compare both'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    Check 'Preference.CloseDiscardsUnsavedChoice' ((Get-PreferenceChoice) -ceq 'Diagnostic' -and (Read-PreferenceRegistry $key) -ceq 'Diagnostic')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceReset')
    Check 'Preference.ResetRetainsSavedChoiceUntilSave' ((Get-PreferenceChoice) -ceq 'Use warehouse default' -and (Read-PreferenceRegistry $key) -ceq 'Diagnostic')
    $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
    Check 'Preference.SaveDefaultRemovesOverrideEffect' ($ok -and (Get-PreferenceChoice) -ceq 'Use warehouse default' -and (Read-PreferenceRegistry $key) -ceq 'Use warehouse default' -and [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceEffective' @($default)))
    foreach($choice in @('How-To','Diagnostic','Compare both')) {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @($choice))
        $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
        Check ('Preference.SaveChoice'+[array]::IndexOf(@('How-To','Diagnostic','Compare both'),$choice)) ($ok -and (Read-PreferenceRegistry $key) -ceq $choice -and [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceEffective' @($choice)))
    }
    $path = Join-Path $settingsRoot 'ActionPathPreferencesV1'
    Set-ItemProperty -LiteralPath $path -Name $key -Value 'invalid isolated preference'
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceReload')
    Check 'Preference.InvalidSavedChoiceFallsBackWithoutRepair' ((Get-PreferenceChoice) -ceq 'Use warehouse default' -and (Read-PreferenceRegistry $key) -ceq 'invalid isolated preference' -and [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceEffective' @($default)))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @('Diagnostic'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
    Check 'Preference.SignedOutSaveDenied' (-not $ok -and (Read-PreferenceRegistry $key) -ceq 'Diagnostic')
    SelectTarget $Other
    $otherContext = [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $otherKey = [string](Run 'invSys.Core.xlam' 'modActionPathPreference.PreferenceTestKey' @($otherContext))
    $otherChoice = [string](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceReadDirect' @($otherContext))
    $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
    Check 'Preference.WarehouseIsolationAndCapturedTarget' (-not $ok -and $otherKey -cne $key -and $otherChoice -ceq 'Use warehouse default' -and $null -eq (Read-PreferenceRegistry $otherKey) -and (Read-PreferenceRegistry $key) -ceq 'Diagnostic')
    SelectTarget $Fixture 'config-reader'
    $readerContext = [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $readerKey = [string](Run 'invSys.Core.xlam' 'modActionPathPreference.PreferenceTestKey' @($readerContext))
    $readerChoice = [string](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceReadDirect' @($readerContext))
    $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSaveDirect' @($readerContext,'How-To'))
    Check 'Preference.UserIsolationWithoutAdminCapability' ($ok -and $readerKey -cne $key -and $readerChoice -ceq 'Use warehouse default' -and (Read-PreferenceRegistry $readerKey) -ceq 'How-To' -and (Read-PreferenceRegistry $key) -ceq 'Diagnostic')
    SelectTarget $Fixture
    $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
    Check 'Preference.StaleSessionSaveDenied' (-not $ok -and (Read-PreferenceRegistry $key) -ceq 'Diagnostic')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    Check 'Preference.ReopenRestoresIdentityChoice' ((Get-PreferenceChoice) -ceq 'Diagnostic')
    foreach($readOnly in @($true,$false)) {
        $book = $excel.Workbooks.Open($Fixture.Config,0,$readOnly)
        try {
            $table = Table $book 'tblWarehouseConfig'
            $cell = $table.ListColumns.Item('WarehouseName').DataBodyRange.Cells.Item(1,1)
            if(-not $readOnly) { $cell.Value2 = 'unsaved personal preference fixture' }
            $value = $cell.Value2
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @('Compare both'))
            $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
            $name = if($readOnly){'ReadOnly'}else{'Dirty'}
            $preserved = $ok -and $cell.Value2 -ceq $value -and ($readOnly -or -not $book.Saved)
            if(-not $readOnly) {
                $unavailable = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceEffective' @('Unavailable'))
            }
        } finally { $book.Close($false) }
        Check ('Preference.LocalSavePreserves'+$name+'Config') ($preserved -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        if(-not $readOnly) { Check 'Preference.UnreadablePolicyNeverInventsEffectiveView' $unavailable }
    }
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceReload')
    Check 'Preference.AllActionsPreserveConfigScope' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and $otherBefore -ceq (Get-FileHash -LiteralPath $Other.Config).Hash)
}

function Test-ActionPathPreferenceRestart($Fixture) {
    SelectTarget $fixture
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceChoose' @('Compare both'))
    $ok = [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PreferenceSave')
    Check 'Preference.RestartFixtureSavedThroughHandler' $ok
    if(-not $ok) { return }
    $before = (Get-FileHash -LiteralPath $fixture.Config).Hash
    $owned = @(Get-Process EXCEL -ErrorAction Stop)
    if($owned.Count -ne 1) { throw 'Restart requires one isolated Excel process.' }
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    foreach($book in @($excel.Workbooks)) { $book.Close($false) }
    if($excel.Workbooks.Count -ne 0) { throw 'Restart fixture workbooks did not close.' }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $script:excel = $null
    $script:packages = @{}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if(-not $owned[0].WaitForExit(5000)) { Stop-Process -Id $owned[0].Id; [void]$owned[0].WaitForExit(5000) }
    if(Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Another Excel process prevents isolated restart.' }
    $script:excel = New-Object -ComObject Excel.Application
    # Failure evidence after this intentional restart must identify the new host.
    $script:initialExcelWindow = [long]$excel.Hwnd
    $script:initialExcelProcessIds = @(Get-Process EXCEL -ErrorAction Stop | Select-Object -ExpandProperty Id)
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    $restartDeploy = $deploy
    if($CheckOperationsTrackingSettings) { $restartDeploy = $operationsSettingsDeploy }
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')) {
        $packages[$name] = $excel.Workbooks.Open((Join-Path $restartDeploy $name),0,$true)
    }
    $probe = $packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $probe.Name = 'PreferenceRestartProbe'
    $probe.CodeModule.AddFromString(@'
Public Function ReadChoice(ByVal context As String) As String
    Dim choice As String, effective As String, evidence As String, report As String
    If modActionPathPreference.ReadPreference(context, choice, effective, evidence, report) Then ReadChoice = choice
End Function
'@)
    SelectTarget $fixture
    $context = [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $choice = [string](Run 'invSys.Core.xlam' 'PreferenceRestartProbe.ReadChoice' @($context))
    $current = @(Get-Process EXCEL -ErrorAction Stop)
    Check 'Preference.NewExcelProcessRestoresSavedChoice' ($current.Count -eq 1 -and $current[0].Id -ne $owned[0].Id -and $choice -ceq 'Compare both')
    Check 'Preference.CoreReadHasNoAdminPackageDependency' (-not (Test-LoadedPackage 'invSys.Admin.xlam') -and $choice -ceq 'Compare both')
    Check 'Preference.RestartReadPreservesConfig' ($before -ceq (Get-FileHash -LiteralPath $fixture.Config).Hash)
    if($CheckOperationsTrackingSettings) { Test-OperationsTrackingSettings $fixture }
}
