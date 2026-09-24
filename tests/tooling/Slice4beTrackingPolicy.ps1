# Serialized policy requests stay in memory; reports contain check identities,
# Boolean outcomes and version/count observations, never Config values.
function Install-Slice4beTrackingPolicyProbe($TestModule,$FormCode) {
    # Observe a real Excel cancellation, not a substituted persistence method.
    # Install before any fixtures/forms exist so VBE edits cannot reset live state.
    $observer=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Add(2)
    $observer.Name='cTrackingPolicySaveObserver'
    $observer.CodeModule.AddFromString(@'
Option Explicit
Private WithEvents mExcel As Excel.Application
Private mPath As String
Private mCancelled As Long
Public Sub Arm(ByVal path As String)
    mPath = path
    mCancelled = 0
    Set mExcel = Application
End Sub
Public Sub Disarm()
    Set mExcel = Nothing
End Sub
Public Property Get CancelledCount() As Long
    CancelledCount = mCancelled
End Property
Private Sub mExcel_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If StrComp(Wb.FullName, mPath, vbTextCompare) <> 0 Then Exit Sub
    mCancelled = mCancelled + 1
    Cancel = True
End Sub
'@)
    $class=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('cAdminTrackingPolicy').CodeModule
    $class.AddFromString(@'
Private mPolicyTestEntries As Long
Public Function PolicyTestRequest() As String
    PolicyTestRequest = mRequest
End Function
Public Function PolicyTestEntries() As Long
    PolicyTestEntries = mPolicyTestEntries
End Function
Public Sub PolicyTestCapture(ByVal enabled As Boolean)
    mLoading = True
    mCapture.Value = enabled
    mLoading = False
    mCapture_Click
End Sub
Public Sub PolicyTestAdminVisible(ByVal enabled As Boolean)
    mLoading = True
    mAdminVisible.Value = enabled
    mLoading = False
    mAdminVisible_Click
End Sub
Public Sub PolicyTestView(ByVal value As String)
    mLoading = True
    mView.Value = value
    mLoading = False
    mView_Change
End Sub
Public Sub PolicyTestControl(ByVal controlId As String, ByVal enabled As Boolean)
    Dim index As Long
    For index = 0 To mRows.ListCount - 1
        If mRows.List(index, 0) = controlId Then
            mLoading = True
            mRows.ListIndex = index
            mLoading = False
            mRows_Click
            mLoading = True
            mCollect.Value = enabled
            mLoading = False
            mCollect_Click
            mLoading = True
            mVisible.Value = enabled
            mLoading = False
            mVisible_Click
            mLoading = True
            mSequence.Value = enabled
            mLoading = False
            mSequence_Click
            Exit Sub
        End If
    Next index
    Err.Raise 5
End Sub
Public Function PolicyTestShowsSaved() As Boolean
    PolicyTestShowsSaved = (InStr(1, mStatus.Caption, " saved.", vbBinaryCompare) > 0)
End Function
'@)
    $line=$class.ProcBodyLine('SavePolicy',0)
    $class.InsertLines($line+2,'    mPolicyTestEntries = mPolicyTestEntries + 1')
    $FormCode.AddFromString(@'
Public Function PolicyTestRequest() As String
    PolicyTestRequest = mTracking.PolicyTestRequest()
End Function
Public Function PolicyTestSave() As Boolean
    PolicyTestSave = mTracking.SavePolicy()
End Function
Public Function PolicyTestEntries() As Long
    PolicyTestEntries = mTracking.PolicyTestEntries()
End Function
Public Sub PolicyTestCapture(ByVal enabled As Boolean)
    mTracking.PolicyTestCapture enabled
End Sub
Public Sub PolicyTestAdminVisible(ByVal enabled As Boolean)
    mTracking.PolicyTestAdminVisible enabled
End Sub
Public Sub PolicyTestView(ByVal value As String)
    mTracking.PolicyTestView value
End Sub
Public Sub PolicyTestReset()
    mTracking.ResetPolicy
End Sub
Public Sub PolicyTestReload()
    mTracking.ReloadPolicy
End Sub
Public Sub PolicyTestControl(ByVal controlId As String, ByVal enabled As Boolean)
    mTracking.PolicyTestControl controlId, enabled
End Sub
Public Function PolicyTestShowsSaved() As Boolean
    PolicyTestShowsSaved = mTracking.PolicyTestShowsSaved()
End Function
Public Sub PolicyTestCloseAction()
    mBtnClose_Click
End Sub
'@)
    $TestModule.CodeModule.AddFromString(@'
Private mPolicySaveObserver As cTrackingPolicySaveObserver
Public Sub TrackingPolicyCancelSave(ByVal path As String)
    Set mPolicySaveObserver = New cTrackingPolicySaveObserver
    mPolicySaveObserver.Arm path
End Sub
Public Function TrackingPolicyCancelledCount() As Long
    TrackingPolicyCancelledCount = mPolicySaveObserver.CancelledCount
End Function
Public Sub TrackingPolicyStopCancelling()
    mPolicySaveObserver.Disarm
    Set mPolicySaveObserver = Nothing
End Sub
Public Sub TrackingPolicyControl(ByVal controlId As String, ByVal enabled As Boolean)
    mForm.PolicyTestControl controlId, enabled
End Sub
Public Function TrackingPolicyShowsSaved() As Boolean
    TrackingPolicyShowsSaved = mForm.PolicyTestShowsSaved()
End Function
Public Sub TrackingPolicyCloseAction()
    mForm.PolicyTestCloseAction
    Set mForm = Nothing
End Sub
Public Function TrackingPolicyRequest() As String
    TrackingPolicyRequest = mForm.PolicyTestRequest()
End Function
Public Function TrackingPolicySave() As Boolean
    TrackingPolicySave = mForm.PolicyTestSave()
End Function
Public Function TrackingPolicyEntries() As Long
    TrackingPolicyEntries = mForm.PolicyTestEntries()
End Function
Public Sub TrackingPolicyCapture(ByVal enabled As Boolean)
    mForm.PolicyTestCapture enabled
End Sub
Public Sub TrackingPolicyAdminVisible(ByVal enabled As Boolean)
    mForm.PolicyTestAdminVisible enabled
End Sub
Public Sub TrackingPolicyView(ByVal value As String)
    mForm.PolicyTestView value
End Sub
Public Sub TrackingPolicyReset()
    mForm.PolicyTestReset
End Sub
Public Sub TrackingPolicyReload()
    mForm.PolicyTestReload
End Sub
Public Function TrackingPolicyContext() As String
    TrackingPolicyContext = modActivity.CaptureContext()
End Function
Public Function TrackingPolicySaveDirect(ByVal context As String, ByVal version As Long, ByVal request As String) As Boolean
    Dim report As String
    TrackingPolicySaveDirect = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
'@)
}
function Get-TrackingPolicyRequest {
    [string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyRequest')
}
function Get-TrackingPolicyVersion($Fixture) {
    $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
    try {
        $latest=0
        foreach($sheet in $book.Worksheets){ foreach($table in $sheet.ListObjects){
            if($table.Name -eq 'tblEventTrackingPolicies'){
                foreach($row in $table.ListRows){$latest=[Math]::Max($latest,[int]$row.Range.Cells.Item(1,$table.ListColumns.Item('PolicyVersion').Index).Value2)}
            }
        }}
        return $latest
    } finally {$book.Close($false)}
}
function Test-Slice4beTrackingPolicy($Fixture,$Other) {
    $request=Get-TrackingPolicyRequest
    $model=$request|ConvertFrom-Json
    $loaded=($model.SchemaVersion -eq 1 -and $model.CatalogVersion -eq 12 -and $model.Controls.Count -eq 68)
    Check 'TrackingPolicy.EditorLoaded' $loaded
    if(-not $loaded){throw 'Tracking policy editor fixture did not load.'}
    Check 'TrackingPolicy.BuiltInDefaults' (-not $model.ViewerActionPathCaptureEnabled -and $model.AdminViewerEventLoggingEnabled -and $model.DefaultView -ceq 'How-To' -and
        @($model.Controls|Where-Object {$_.ControlId -eq 'RECEIVING_CONFIRM_WRITES' -and $_.Collect}).Count -eq 1 -and
        @($model.Controls|Where-Object {$_.ControlId -eq 'RECEIVING_PAGE_RECEIPTS' -and -not $_.Collect}).Count -eq 1)
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($true))
    Check 'TrackingPolicy.CaptureStagesOnly' ((Get-TrackingPolicyRequest|ConvertFrom-Json).ViewerActionPathCaptureEnabled -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyReset')
    Check 'TrackingPolicy.ResetStagesOnly' (-not (Get-TrackingPolicyRequest|ConvertFrom-Json).ViewerActionPathCaptureEnabled -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    $context=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyContext')
    foreach($case in @('UnknownField','InvalidBoolean','InvalidView','UnknownControl','MissingControl','DuplicateControl')){
        $bad=$request|ConvertFrom-Json
        switch($case){
            UnknownField {$bad|Add-Member NoteProperty Unrecognized $true}
            InvalidBoolean {$bad.ViewerActionPathCaptureEnabled='true'}
            InvalidView {$bad.DefaultView='not-a-view'}
            UnknownControl {$bad.Controls[0].ControlId='CANONICAL_INVENTORY_COLLECTION'}
            MissingControl {$bad.Controls=@($bad.Controls|Select-Object -Skip 1)}
            DuplicateControl {$bad.Controls[1].ControlId=$bad.Controls[0].ControlId}
        }
        $badText=$bad|ConvertTo-Json -Depth 8 -Compress
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,0,$badText))
        Check ('TrackingPolicy.Reject'+$case) (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    }
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,1,$request))
    Check 'TrackingPolicy.WrongExpectedVersionDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,0,$request))
    Check 'TrackingPolicy.SignedOutDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    SelectTarget $Other
    $otherBefore=(Get-FileHash -LiteralPath $Other.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySave')
    Check 'TrackingPolicy.CapturedTargetDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and $otherBefore -ceq (Get-FileHash -LiteralPath $Other.Config).Hash)
    SelectTarget $Fixture 'config-reader'
    $context=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyContext')
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,0,$request))
    Check 'TrackingPolicy.MissingCapabilityDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    SelectTarget $Fixture
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySave')
    Check 'TrackingPolicy.StaleSessionDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    $context=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyContext')
    foreach($readOnly in @($true,$false)){
        $book=$excel.Workbooks.Open($Fixture.Config,0,$readOnly)
        $table=Table $book 'tblWarehouseConfig'
        $cell=$table.ListColumns.Item('WarehouseName').DataBodyRange.Cells.Item(1,1)
        if(-not $readOnly){$cell.Value2='unsaved fixture marker'}
        $value=$cell.Value2
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,0,$request))
        $name=if($readOnly){'ReadOnly'}else{'Dirty'}
        $memoryPreserved=(-not $ok -and $cell.Value2 -ceq $value -and ($readOnly -or -not $book.Saved))
        $book.Close($false)
        Check ('TrackingPolicy.'+$name+'ConfigPreserved') ($memoryPreserved -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    }
    $controlId='RECEIVING_CONFIRM_WRITES'
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyControl' @($controlId,$false))
    $selected=(Get-TrackingPolicyRequest|ConvertFrom-Json).Controls|Where-Object ControlId -CEQ $controlId
    Check 'TrackingPolicy.PerControlHandlersStageOnly' (-not $selected.Collect -and -not $selected.Visible -and -not $selected.SequenceEligible -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyReload')
    $selected=(Get-TrackingPolicyRequest|ConvertFrom-Json).Controls|Where-Object ControlId -CEQ $controlId
    Check 'TrackingPolicy.ReloadDiscardsStagedControlEdits' ($selected.Collect -and $selected.Visible -and $selected.SequenceEligible -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($true))
    $staged=Get-TrackingPolicyRequest
    $eventsEnabled=$excel.EnableEvents
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCancelSave' @($Fixture.Config))
    try {
        $excel.EnableEvents=$true
        $cancelledSave=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySave')
        $cancelCount=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCancelledCount')
    } finally {
        $excel.EnableEvents=$eventsEnabled
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyStopCancelling')
    }
    Check 'TrackingPolicy.ExcelBeforeSaveActuallyCancelled' ($cancelCount -eq 1)
    Check 'TrackingPolicy.CancelledSavePreservesConfigBytes' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and (Get-TrackingPolicyVersion $Fixture) -eq 0)
    Check 'TrackingPolicy.CancelledSaveNeverReportsSuccess' (-not $cancelledSave -and -not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyShowsSaved'))
    Check 'TrackingPolicy.CancelledSaveRetainsStagedEdits' ($staged -ceq (Get-TrackingPolicyRequest))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyControl' @($controlId,$false))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($true))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyAdminVisible' @($false))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyView' @('Compare both'))
    foreach($boxingId in @('BOXING_MAKE','BOXING_UNBOX')) {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyControl' @($boxingId,$false))
        $boxingRows=@((Get-TrackingPolicyRequest|ConvertFrom-Json).Controls|Where-Object ControlId -CEQ $boxingId)
        $boxingOff=$boxingRows.Count -eq 1
        if($boxingOff){foreach($flag in @('Collect','Visible','SequenceEligible')){$boxingOff=$boxingOff -and $boxingRows[0].$flag -is [bool] -and -not $boxingRows[0].$flag}}
        Check ('TrackingPolicy.Boxing.'+$boxingId+'.ActualHandlersStageOnly') ($boxingOff -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    }
    $entries=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyEntries')
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySave')
    Check 'TrackingPolicy.RealSaveActionEntered' ([int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyEntries') -eq $entries+1)
    Check 'TrackingPolicy.AuthorizedSavePublishesVersion' ($ok -and (Get-TrackingPolicyVersion $Fixture) -eq 1)
    if(-not $ok){return}
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyReload')
    Check 'TrackingPolicy.SavedCaptureReloads' (Get-TrackingPolicyRequest|ConvertFrom-Json).ViewerActionPathCaptureEnabled
    Check 'TrackingPolicy.CompatibilityFlagsReloadTogether' (-not (Get-TrackingPolicyRequest|ConvertFrom-Json).AdminViewerEventLoggingEnabled)
    Check 'TrackingPolicy.WarehouseViewReloads' ((Get-TrackingPolicyRequest|ConvertFrom-Json).DefaultView -ceq 'Compare both')
    $selected=(Get-TrackingPolicyRequest|ConvertFrom-Json).Controls|Where-Object ControlId -CEQ $controlId
    Check 'TrackingPolicy.PerControlFlagsPersistTogether' (-not $selected.Collect -and -not $selected.Visible -and -not $selected.SequenceEligible)
    foreach($boxingId in @('BOXING_MAKE','BOXING_UNBOX')) {
        $boxingRows=@((Get-TrackingPolicyRequest|ConvertFrom-Json).Controls|Where-Object ControlId -CEQ $boxingId)
        $boxingOff=$boxingRows.Count -eq 1
        if($boxingOff){foreach($flag in @('Collect','Visible','SequenceEligible')){$boxingOff=$boxingOff -and $boxingRows[0].$flag -is [bool] -and -not $boxingRows[0].$flag}}
        Check ('TrackingPolicy.Boxing.'+$boxingId+'.SavedFlagsReloadTogether') $boxingOff
    }
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,0,$request))
    Check 'TrackingPolicy.StaleVersionCannotAppend' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
    $headers=Table $book 'tblEventTrackingPolicies'; $controls=Table $book 'tblEventTrackingControls'
    $headerColumn=$headers.ListColumns.Add(1); $headerColumn.Name='Operator Extra'; $headerColumn.DataBodyRange.Value2='header fixture value'
    $controlColumn=$controls.ListColumns.Add(1); $controlColumn.Name='Operator Extra'; $controlColumn.DataBodyRange.Value2='control fixture value'
    $priorHeaders=$headers.DataBodyRange.Value2; $priorControls=$controls.DataBodyRange.Value2
    $controlCount=$controls.ListRows.Count
    $book.Save(); $book.Close($false)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($false))
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySave')
    Check 'TrackingPolicy.SecondSaveAppendsVersion' ($ok -and (Get-TrackingPolicyVersion $Fixture) -eq 2)
    $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
    try {
        $headers=Table $book 'tblEventTrackingPolicies'; $controls=Table $book 'tblEventTrackingControls'
        $retained=($headers.ListRows.Count -eq 2 -and $controls.ListRows.Count -eq $controlCount*2)
        for($c=1;$c -le $headers.ListColumns.Count;$c++){$retained=$retained -and $headers.DataBodyRange.Cells.Item(1,$c).Value2 -ceq $priorHeaders.GetValue(1,$c)}
        for($r=1;$r -le $controlCount;$r++){for($c=1;$c -le $controls.ListColumns.Count;$c++){$retained=$retained -and $controls.DataBodyRange.Cells.Item($r,$c).Value2 -ceq $priorControls.GetValue($r,$c)}}
        Check 'TrackingPolicy.EarlierRowsAndUnknownColumnsPreserved' $retained
    } finally {$book.Close($false)}
    if($CaptureEvidence){
        $wasVisible=$excel.Visible
        try {
            $excel.Visible=$true
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingSettingsSelectPage' @('Event Tracking'))
            Start-Sleep -Milliseconds 300
            CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'tracking-policy-saved.png'
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingSettingsSelectPage' @('General'))
        } finally {$excel.Visible=$wasVisible}
    }
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($true))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
    Check 'TrackingPolicy.CloseActionDoesNotSave' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    Check 'TrackingPolicy.CloseDiscardsStagedEdits' (-not (Get-TrackingPolicyRequest|ConvertFrom-Json).ViewerActionPathCaptureEnabled -and (Get-TrackingPolicyVersion $Fixture) -eq 2)
}
