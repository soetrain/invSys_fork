# D18: actual generated Ribbon dispatch, unchanged authorization owner and
# count/boolean-only unsaved instrumentation. No credential or audit payloads.
function Install-ReceivingLauncherDenialCoreSeam($Project,$Gate) {
    $Gate.AddFromString(@'
Public DenialGuardCalls As Long
Public DenialAuthCalls As Long
Public DenialGuardMessage As String
Public DenialSignOutAtNotice As Boolean
Public DenialNativeNotice As Boolean
Public Sub ResetDenial()
    DenialGuardCalls = 0: DenialAuthCalls = 0: DenialGuardMessage = ""
    DenialSignOutAtNotice = False
    DenialNativeNotice = False
    LauncherCalls = 0
End Sub
Public Sub EnableDenialNativeNotice()
    DenialNativeNotice = True
End Sub
Public Sub InterruptDenialAtNotice()
    DenialSignOutAtNotice = True
End Sub
Public Function DenialGuardState() As String
    DenialGuardState = CStr(DenialGuardCalls) & "|" & CStr(DenialAuthCalls) & "|" & _
        CStr(DenialGuardMessage = "Current user does not have RECEIVE_POST for this warehouse/station.")
End Function
'@)
    $guard=$Project.VBComponents.Item('modRoleUiAccess').CodeModule
    $start=$guard.ProcStartLine('RequireCurrentUserCapabilityCached',0)
    $count=$guard.ProcCountLines('RequireCurrentUserCapabilityCached',0)
    $source=$guard.Lines($start,$count)
    $changed=$source.Replace('    RequireCurrentUserCapabilityCached = CanCurrentUserPerformCapabilityCached',"    TestReceivingActivityGate.DenialGuardCalls = TestReceivingActivityGate.DenialGuardCalls + 1`r`n    RequireCurrentUserCapabilityCached = CanCurrentUserPerformCapabilityCached")
    $changed=$changed.Replace('If deniedMessage <> "" Then MsgBox deniedMessage, vbExclamation',"If deniedMessage <> `"`" Then TestReceivingActivityGate.DenialGuardMessage = deniedMessage`r`n    If TestReceivingActivityGate.DenialNativeNotice Then MsgBox deniedMessage, vbExclamation`r`n    If TestReceivingActivityGate.DenialSignOutAtNotice Then modAuth.SignOut")
    if($changed -eq $source){throw 'Actual cached guard notice seam unavailable.'}
    $guard.DeleteLines($start,$count); $guard.InsertLines($start,$changed)
    $auth=$Project.VBComponents.Item('modAuth').CodeModule
    $start=$auth.ProcStartLine('LogDecision',0); $count=$auth.ProcCountLines('LogDecision',0)
    $source=$auth.Lines($start,$count)
    $changed=$source.Replace('    Debug.Print Format$(Now,',"    If result = `"DENY`" And capability = `"RECEIVE_POST`" Then TestReceivingActivityGate.DenialAuthCalls = TestReceivingActivityGate.DenialAuthCalls + 1`r`n    Debug.Print Format`$(Now,")
    if($changed -eq $source){throw 'Existing authorization decision observation seam unavailable.'}
    $auth.DeleteLines($start,$count); $auth.InsertLines($start,$changed)
}

function Install-ReceivingLauncherDenialOperationsSeam($Project) {
    $notice=$Project.VBComponents.Item('TestReceivingLifecycleNotice').CodeModule
    $notice.AddFromString(@'
Public NativeEvidence As Boolean
Public Sub EnableNativeEvidence()
    NativeEvidence = True
End Sub
'@)
    $start=$notice.ProcStartLine('Reset',0); $count=$notice.ProcCountLines('Reset',0)
    $source=$notice.Lines($start,$count).Replace('    LastMessage = ""','    LastMessage = "": NativeEvidence = False')
    $notice.DeleteLines($start,$count); $notice.InsertLines($start,$source)
    $launcher=$Project.VBComponents.Item('modTS_Received').CodeModule
    $start=$launcher.ProcStartLine('ShowReceivingMessage',0); $count=$launcher.ProcCountLines('ShowReceivingMessage',0)
    $source=$launcher.Lines($start,$count).Replace('    TestReceivingLifecycleNotice.LastMessage = messageText',"    TestReceivingLifecycleNotice.LastMessage = messageText`r`n    If TestReceivingLifecycleNotice.NativeEvidence Then modReceivingActivityAction.ShowMessage messageText, style")
    $launcher.DeleteLines($start,$count); $launcher.InsertLines($start,$source)
    $Project.VBComponents.Item('modTS_Received').CodeModule.AddFromString(@'
Public Function ActivityTestReceivingRibbonEnabled() As Boolean
    Dim control As New TestReceivingRibbonControl, value As Variant
    modRibbonGenerated.RibbonRequiredCapabilityGetEnabledOperations control, value
    ActivityTestReceivingRibbonEnabled = CBool(value)
End Function
Public Function ActivityTestReceivingDirectGuard() As Boolean
    ActivityTestReceivingDirectGuard = modRoleUiAccess.RequireCurrentUserCapabilityCached( _
        "RECEIVE_POST", "Current user does not have RECEIVE_POST for this warehouse/station.")
End Function
'@)
}

function Reset-ReceivingLauncherDenial {
    [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.ResetDenial')
    [void](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Reset')
}

function Test-ReceivingLauncherDeniedOwner([string]$Label,[bool]$SignedIn=$true) {
    $guard=([string](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.DenialGuardState')).Split('|')
    Check ($Label+'.ExistingGuardOnceAndNoticePreserved') ($guard[0] -ceq '1' -and $guard[2] -ceq 'True')
    if($SignedIn){Check ($Label+'.AuthorizationDecisionRetained') ([int]$guard[1] -ge 1)}
    Check ($Label+'.NoProvisioningOrFormOpen') ([int](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.LauncherCallCount') -eq 0 -and [string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState') -eq '')
}

function Test-ReceivingLauncherDenial($Fixture) {
    SelectTarget $Fixture 'config-reader'
    $before=@(Get-Slice4beActivityFiles $Fixture)
    Check 'LauncherDenial.AllowedEnablement' ([bool](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestReceivingRibbonEnabled'))
    Check 'LauncherDenial.AllowedPollingIsNotAnAction' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
    $other=$excel.Workbooks.Add(); $other.Worksheets.Item(1).Cells.Item(1,1).Value2='denial unrelated sentinel'
    $other.SaveAs((Join-Path $runRoot 'denial-other.xlsm'),52)
    $otherHash=Get-ReceivingFixtureHash $other.FullName
    $operator=$null
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    try {
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
        $state=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState')
        if(-not $state.EndsWith('|True')){throw 'Denial fixture could not establish the accepted Receiving launcher.'}
        $operator=$excel.Workbooks.Item($state.Split('|')[0])
        if(-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name))){throw 'Denial staging fixture unavailable.'}
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @('Internal'))
        $extra=(Table $operator 'ReceivedTally').ListColumns.Add(); $extra.Name='Denial Extra'; $extra.DataBodyRange.Value2='preserve denial extension'
        $staged=@(Get-ReceivingFixtureRows (Table $operator 'ReceivedTally')) | ConvertTo-Json -Depth 5 -Compress
        $operator.Save(); $operatorHash=Get-ReceivingFixtureHash $operator.FullName
        SelectTarget $Fixture 'config-producer'; $other.Activate()
        $before=@(Get-Slice4beActivityFiles $Fixture)
        Check 'LauncherDenial.DeniedEnablement' (-not [bool](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestReceivingRibbonEnabled'))
        Check 'LauncherDenial.DeniedPollingIsNotAnAction' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        $authority=Get-ReceivingAuthorityHashes $Fixture
        $repeatedBefore=@(Get-Slice4beActivityFiles $Fixture)
        foreach($case in @('First','Repeated')) {
            Reset-ReceivingLauncherDenial
            $before=@(Get-Slice4beActivityFiles $Fixture)
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
            Test-ReceivingLauncherDeniedOwner ('LauncherDenial.'+$case)
            Test-ReceivingControlRecords $Fixture $before 'RECEIVING_OPEN' 'RECEIVING_WORKFLOW' 'RECEIVE_OPEN_' 'DENIED' 1 'Unchanged' @() ('LauncherDenial.'+$case) 'Blocked' 5 'config-producer'
        }
        Test-ReceivingControlRecords $Fixture $repeatedBefore 'RECEIVING_OPEN' 'RECEIVING_WORKFLOW' 'RECEIVE_OPEN_' 'DENIED' 2 'Unchanged' @() 'LauncherDenial.RepeatedDistinct' 'Blocked' 5 'config-producer'
        Reset-ReceivingLauncherDenial
        $before=@(Get-Slice4beActivityFiles $Fixture)
        Check 'LauncherDenial.DirectGuardDenied' (-not [bool](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestReceivingDirectGuard'))
        Check 'LauncherDenial.DirectGuardIsNotAnAction' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        Reset-ReceivingLauncherDenial
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
        Test-ReceivingLauncherDeniedOwner 'LauncherDenial.SignedOut' $false
        Check 'LauncherDenial.SignedOutIsNotAttributed' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        SelectTarget $Fixture 'config-producer'; $other.Activate()
        Reset-ReceivingLauncherDenial
        [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.InterruptDenialAtNotice')
        $interruptedBefore=@(Get-Slice4beActivityFiles $Fixture)
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
        Test-ReceivingLauncherDeniedOwner 'LauncherDenial.Interrupted'
        $partial=@(Get-Slice4beActivityFiles $Fixture | Where-Object {$_ -notin $interruptedBefore} | ForEach-Object {Get-Content -LiteralPath $_ -Raw | ConvertFrom-Json})
        Check 'LauncherDenial.Interrupted.OnlyOriginalAttempt' ($partial.Count -eq 1 -and $partial[0].ControlId -ceq 'RECEIVING_OPEN' -and $partial[0].OutcomeCode -ceq 'REQUESTED' -and $partial[0].UserId -ceq 'config-producer' -and $partial[0].WarehouseId -ceq $Fixture.Warehouse -and @($partial[0].SourceEventRefs).Count -eq 0)
        Check 'LauncherDenial.Interrupted.IncompleteTrackingVisible' (([string](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Message')).Contains('Tracking unavailable'))
        SelectTarget $Fixture 'config-reader'
        Check 'LauncherDenial.Interrupted.NoNewActorCompletion' (@(Get-Slice4beActivityFiles $Fixture).Count -eq ($interruptedBefore.Count+$partial.Count))
        SelectTarget $Fixture 'config-producer'; $other.Activate()
        $before=@(Get-Slice4beActivityFiles $Fixture)
        Reset-ReceivingLauncherDenial
        $held=Hold-ReceivingLocalActivityStore $Fixture
        $dialogJob=$null
        try {
            if($CaptureDenialDialogs) {
                $dialogJob=Start-ReceivingDenialDialogEvidence
                [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.EnableDenialNativeNotice')
                [void](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.EnableNativeEvidence')
            }
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
            Test-ReceivingLauncherDeniedOwner 'LauncherDenial.StoreFailure'
            Check 'LauncherDenial.StoreFailure.VisibleTrackingNotice' (([string](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Message')).Contains('Tracking unavailable'))
        } finally {
            try {if($null -ne $dialogJob){Complete-ReceivingDenialDialogEvidence $dialogJob}}
            finally {Remove-Item -LiteralPath $held[0]; Move-Item -LiteralPath $held[1] -Destination $held[0]}
        }
        Check 'LauncherDenial.StoreFailure.NoInventedRecords' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        Check 'LauncherDenial.AuthorityBytesPreserved' (Test-ReceivingAuthorityHashes $Fixture $authority)
        Set-ReceivingNavigationPolicy $Fixture $false
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        try {
            $controls=Table $cfg 'tblEventTrackingControls'
            foreach($row in $controls.ListRows){if($row.Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -ceq 'RECEIVING_OPEN'){$row.Range.Cells.Item(1,$controls.ListColumns.Item('Collect').Index).Value2=$false}}
            $cfg.Save()
        } finally {$cfg.Close($false)}
        Reset-ReceivingLauncherDenial
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
        Test-ReceivingLauncherDeniedOwner 'LauncherDenial.PolicyOff'
        Check 'LauncherDenial.PolicyOff.NoOptionalRecordOrFailureNotice' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count -and [string](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Message') -eq '')
        Check 'LauncherDenial.CapturedStagingAndIdentityPreserved' ((@(Get-ReceivingFixtureRows (Table $operator 'ReceivedTally')) | ConvertTo-Json -Depth 5 -Compress) -ceq $staged -and $operatorHash -ceq (Get-ReceivingFixtureHash $operator.FullName))
        Check 'LauncherDenial.UnrelatedWorkbookAndActivationPreserved' ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).ListObjects.Count -eq 0 -and $excel.ActiveWorkbook.Name -ceq $other.Name)
    } finally {
        if($null -ne $operator){$operator.Close($false)}; $other.Close($false)
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        SelectTarget $Fixture 'config-reader'
    }
}
