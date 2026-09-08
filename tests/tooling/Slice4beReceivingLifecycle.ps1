# D18: invoke the real generated Ribbon callback using Office's control interface,
# and actual form Close/QueryClose handlers. All seams remain unsaved in fixtures.
function Install-ReceivingLifecycleCoreSeam($Bridge,$Gate) {
    $Gate.AddFromString(@'
Public LauncherFault As Boolean
Public LauncherCalls As Long
Public Sub SetLauncherFault(ByVal value As Boolean)
    LauncherFault = value
    LauncherCalls = 0
End Sub
Public Function LauncherCallCount() As Long
    LauncherCallCount = LauncherCalls
End Function
'@)
    $start=$Bridge.ProcStartLine('OpenOrCreateCurrentReceivingOperatorWorkbook',0)
    $count=$Bridge.ProcCountLines('OpenOrCreateCurrentReceivingOperatorWorkbook',0)
    $source=$Bridge.Lines($start,$count)
    $changed=$source.Replace('    OpenOrCreateCurrentReceivingOperatorWorkbook = _', @'
    TestReceivingActivityGate.LauncherCalls = TestReceivingActivityGate.LauncherCalls + 1
    If TestReceivingActivityGate.LauncherFault Then report = "Fixture Receiving launch withheld.": Exit Function
    OpenOrCreateCurrentReceivingOperatorWorkbook = _
'@)
    if ($changed -eq $source) { throw 'Receiving launcher owner fault seam is unavailable.' }
    $Bridge.DeleteLines($start,$count); $Bridge.InsertLines($start,$changed)
}

function Install-ReceivingLifecycleSeams($Project,$Form) {
    $control=$Project.VBComponents.Add(2); $control.Name='TestReceivingRibbonControl'
    $control.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnOperationsReceivingForm"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $Form.AddFromString(@'
Public Sub ActivityTestDismiss(ByVal windowClose As Boolean)
    Dim cancelled As Integer
    If windowClose Then
        UserForm_QueryClose cancelled, vbFormControlMenu
        If cancelled = 0 Then Unload Me
    Else
        mBtnClose_Click
    End If
End Sub
Public Function ActivityTestContextCurrent() As Boolean
    ActivityTestContextCurrent = mActivityContext <> "" And mActivityContext = modActivity.CaptureContext()
End Function
Public Function ActivityTestBoundTo(ByVal wb As Workbook) As Boolean
    Dim captured As Workbook
    Set captured = ResolveOperatorWorkbook()
    If Not captured Is Nothing Then ActivityTestBoundTo = (captured Is wb)
End Function
'@)
    $launcher=$Project.VBComponents.Item('modTS_Received').CodeModule
    $launcher.AddFromString(@'
Public Sub ActivityTestRibbonOpen()
    Dim control As New TestReceivingRibbonControl
    modRibbonGenerated.RibbonOnActionOperations control
End Sub
Public Function ActivityTestLauncherState() As String
    If mReceivingLauncherFormTerminated Then Exit Function
    If mReceivingLauncherForm Is Nothing Then Exit Function
    If Not mReceivingLauncherForm.Visible Then Exit Function
    ActivityTestLauncherState = mReceivingLauncherWorkbookName & "|" & CStr(mReceivingLauncherForm.ActivityTestContextCurrent())
End Function
Public Function ActivityTestLauncherStatus() As String
    If mReceivingLauncherFormTerminated Then Exit Function
    If mReceivingLauncherForm Is Nothing Then Exit Function
    ActivityTestLauncherStatus = CStr(mReceivingLauncherForm.Controls("txtStatus").Value)
End Function
Public Function ActivityTestLauncherBoundTo(ByVal workbookName As String) As Boolean
    If mReceivingLauncherFormTerminated Then Exit Function
    If mReceivingLauncherForm Is Nothing Then Exit Function
    ActivityTestLauncherBoundTo = mReceivingLauncherForm.ActivityTestBoundTo(Application.Workbooks(workbookName))
End Function
Public Sub ActivityTestLauncherDismiss(ByVal mode As String)
    Select Case mode
        Case "Button": mReceivingLauncherForm.ActivityTestDismiss False
        Case "Window": mReceivingLauncherForm.ActivityTestDismiss True
        Case "Internal": Unload mReceivingLauncherForm
    End Select
End Sub
'@)
    # Intercept only the existing message presentation. Keep its exact text for
    # boolean cause checks, never raw console/report output or a fixture MsgBox.
    $helper=$Project.VBComponents.Add(1); $helper.Name='TestReceivingLifecycleNotice'
    $helper.CodeModule.AddFromString(@'
Option Explicit
Public LastMessage As String
Public Function Message() As String
    Message = LastMessage
End Function
Public Sub Reset()
    LastMessage = ""
End Sub
'@)
    $start=$launcher.ProcStartLine('ShowReceivingMessage',0)
    $count=$launcher.ProcCountLines('ShowReceivingMessage',0)
    $launcher.DeleteLines($start,$count)
    $launcher.InsertLines($start,@'
Private Sub ShowReceivingMessage(ByVal messageText As String, ByVal style As VbMsgBoxStyle)
    TestReceivingLifecycleNotice.LastMessage = messageText
End Sub
'@)
}

function Test-ReceivingLifecycleRecords($Fixture,$Before,[string]$Control,[string]$Outcome,[string]$Label) {
    $prefix=if($Control -eq 'RECEIVING_OPEN'){'RECEIVE_OPEN_'}else{'RECEIVE_CLOSE_'}
    $effect=if($Outcome -eq 'CLOSED'){'Unchanged'}else{'Unknown'}
    $severity=if($Outcome -eq 'FAILED'){'Error'}else{'Info'}
    Test-ReceivingControlRecords $Fixture $Before $Control 'RECEIVING_WORKFLOW' $prefix $Outcome 1 $effect @() $Label $severity 5
}

function Test-ReceivingLifecycleActivity($Fixture) {
    SelectTarget $Fixture 'config-reader'
    $other=$excel.Workbooks.Add()
    $other.Worksheets.Item(1).Cells.Item(1,1).Value2='lifecycle unrelated sentinel'
    $other.SaveAs((Join-Path $runRoot 'lifecycle-other.xlsm'),52)
    $otherHash=Get-ReceivingFixtureHash $other.FullName
    $before=@(Get-Slice4beActivityFiles $Fixture)
    [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
    $state=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState')
    if (-not $state.EndsWith('|True')) { throw 'Actual Ribbon callback did not establish a current visible Receiving fixture.' }
    $operator=$excel.Workbooks.Item($state.Split('|')[0])
    Check 'Lifecycle.Open.VisibleCapturedWorkbook' (-not $operator.IsAddin -and $operator.Name -cne $other.Name)
    if ($CaptureEvidence) { CaptureFormEvidence 'Receiving' 'coverage-lifecycle-open.png' }
    Test-ReceivingLifecycleRecords $Fixture $before 'RECEIVING_OPEN' 'OPENED' 'Lifecycle.Open'
    $authority=Get-ReceivingAuthorityHashes $Fixture
    $extra=(Table $operator 'ReceivedTally').ListColumns.Add(); $extra.Name='Lifecycle Extra'
    if (-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name))) { throw 'Lifecycle staged-row fixture failed.' }
    [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
    $extra.DataBodyRange.Value2='preserved lifecycle extension'
    $stagingBefore=@(Get-ReceivingFixtureRows (Table $operator 'ReceivedTally')) | ConvertTo-Json -Depth 5 -Compress
    $operator.Save()
    $operatorHash=Get-ReceivingFixtureHash $operator.FullName
    foreach($mode in @('Reuse','SessionChanged','Button','Window','Internal')) {
        $other.Activate()
        $before=@(Get-Slice4beActivityFiles $Fixture)
        if($mode -eq 'SessionChanged') {
            [void](Run 'invSys.Core.xlam' 'modAuth.SignOut'); SelectTarget $Fixture 'config-reader'
        }
        if($mode -in @('Reuse','SessionChanged')) {
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
            $current=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState')
            Check "Lifecycle.$mode.CurrentCapturedBinding" ($current -ceq $state)
            if ($CaptureEvidence) { CaptureFormEvidence 'Receiving' ('coverage-lifecycle-'+$mode.ToLowerInvariant()+'.png') }
            $outcome=if($mode -eq 'Reuse'){'REUSED'}else{'OPENED'}
            Test-ReceivingLifecycleRecords $Fixture $before 'RECEIVING_OPEN' $outcome ('Lifecycle.'+$mode)
            $new=@(Get-Slice4beActivityFiles $Fixture | Where-Object { $_ -notin $before })
            Check "Lifecycle.$mode.NoInternalCloseOrInitializationClicks" ($new.Count -le 2)
        } else {
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @($mode))
            Check "Lifecycle.$mode.Dismissed" ([string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState') -eq '')
            if($mode -eq 'Internal') {
                Check 'Lifecycle.Internal.NotAUserClose' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
            } else { Test-ReceivingLifecycleRecords $Fixture $before 'RECEIVING_CLOSE' 'CLOSED' ('Lifecycle.'+$mode) }
            $beforeDirect=@(Get-Slice4beActivityFiles $Fixture)
            $operator.Activate()
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ShowReceivingForm')
            Check "Lifecycle.$mode.DirectMacroNotAClick" (@(Get-Slice4beActivityFiles $Fixture).Count -eq $beforeDirect.Count)
        }
        Check "Lifecycle.$mode.AuthorityBytesPreserved" (Test-ReceivingAuthorityHashes $Fixture $authority)
        Check "Lifecycle.$mode.UnknownColumnAndSavedWorkbookPreserved" ((Table $operator 'ReceivedTally').ListColumns.Item('Lifecycle Extra').Name -ceq 'Lifecycle Extra' -and $operatorHash -ceq (Get-ReceivingFixtureHash $operator.FullName))
        $stagingAfter=@(Get-ReceivingFixtureRows (Table $operator 'ReceivedTally')) | ConvertTo-Json -Depth 5 -Compress
        Check "Lifecycle.$mode.StagedIdentitiesAndUnknownValuesPreserved" ($stagingAfter -ceq $stagingBefore)
    }
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut'); SelectTarget $Fixture 'config-reader'
    $before=@(Get-Slice4beActivityFiles $Fixture)
    [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @('Button'))
    Check 'Lifecycle.StaleClose.DismissedWithoutNewSessionAttribution' ([string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState') -eq '' -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
    $before=@(Get-Slice4beActivityFiles $Fixture)
    [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetLauncherFault' @($true))
    [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
    Check 'Lifecycle.Failed.OwnerCalledOnceAndCausePreserved' ([int](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.LauncherCallCount') -eq 1 -and [string](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Message') -like '*Fixture Receiving launch withheld.*')
    Test-ReceivingLifecycleRecords $Fixture $before 'RECEIVING_OPEN' 'FAILED' 'Lifecycle.Failed'
    [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetLauncherFault' @($false))
    foreach($mode in @('Button','Window')) {
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $held=Hold-ReceivingLocalActivityStore $Fixture
        try {
            [void](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Reset')
            [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetLauncherFault' @($false))
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
            $status=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherStatus')
            $notice=[string](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Message')
            Check "Lifecycle.StoreFailure.$mode.OpenedOnce" ([string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState') -ceq $state -and [int](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.LauncherCallCount') -eq 1)
            Check "Lifecycle.StoreFailure.$mode.OpenNoticeVisible" (($status+' '+$notice).Contains('Tracking unavailable'))
            [void](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Reset')
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @($mode))
            Check "Lifecycle.StoreFailure.$mode.Dismissed" ([string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState') -eq '')
            Check "Lifecycle.StoreFailure.$mode.CloseNoticeVisible" (([string](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Message')).Contains('Tracking unavailable'))
        } finally {
            Remove-Item -LiteralPath $held[0]
            Move-Item -LiteralPath $held[1] -Destination $held[0]
        }
        Check "Lifecycle.StoreFailure.$mode.NoInventedRecords" (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
    }
    $operator.Activate()
    [void](Run 'invSys.Operations.xlam' 'modTS_Received.ShowReceivingForm')
    $operatorName=$operator.Name
    $operator.Close($false)
    $other.Activate()
    $before=@(Get-Slice4beActivityFiles $Fixture)
    [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
    $operator=$excel.Workbooks.Item($operatorName)
    Check 'Lifecycle.ReopenedWorkbook.ExactLiveObjectBound' ([bool](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherBoundTo' @($operator.Name)))
    Test-ReceivingLifecycleRecords $Fixture $before 'RECEIVING_OPEN' 'OPENED' 'Lifecycle.ReopenedWorkbook'
    [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @('Internal'))
    Check 'Lifecycle.Final.AuthorityBytesPreserved' (Test-ReceivingAuthorityHashes $Fixture $authority)
    Check 'Lifecycle.OtherWorkbookUnchanged' ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).ListObjects.Count -eq 0 -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'lifecycle unrelated sentinel')
    $operator.Close($false); $other.Close($false)
}

function Test-ReceivingLifecycleOlderPolicy($Fixture,[string]$RecordId,[string]$Context) {
    $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
    $meta=Table $cfg 'tblEventTrackingPolicies'; $rows=Table $cfg 'tblEventTrackingControls'
    $meta.ListColumns.Item('CatalogVersion').DataBodyRange.Cells.Item(1,1).Value2=4.0
    foreach($control in @('RECEIVING_CONFIRM_WRITES','RECEIVING_ADD_SELECTED','DISPOSITION_ADD_SELECTED','DISPOSITION_CONFIRM','RECEIVING_REFRESH','RECEIVING_CLEAR')) {
        $row=$rows.ListRows.Add()
        foreach($field in @('Collect','Visible','SequenceEligible')) { $row.Range.Cells.Item(1,$rows.ListColumns.Item($field).Index).Value2=$true }
        $row.Range.Cells.Item(1,$rows.ListColumns.Item('PolicyVersion').Index).Value2=1.0
        $row.Range.Cells.Item(1,$rows.ListColumns.Item('ControlId').Index).Value2=$control
    }
    $cfg.Save(); $cfg.Close($false)
    try {
        Check 'Lifecycle.Policy.CatalogFourRemainsValid' ((Get-ActivityRead $RecordId).StartsWith('OK|'))
        foreach($control in @('RECEIVING_OPEN','RECEIVING_CLOSE')) {
            $count=@(Get-Slice4beActivityFiles $Fixture).Count
            $id=[string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @($control,$Context))
            Check ('Lifecycle.Policy.OlderCatalogDoesNotEnable.'+$control) ($id -eq '' -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $count)
        }
    } finally {
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $rows=Table $cfg 'tblEventTrackingControls'
        while($rows.ListRows.Count -gt 2) { $rows.ListRows.Item($rows.ListRows.Count).Delete() }
        (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Cells.Item(1,1).Value2=1.0
        $cfg.Save(); $cfg.Close($false)
    }
}
