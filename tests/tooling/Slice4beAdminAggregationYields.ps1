# Named returns of the actual operator handlers; only external waits/dialogs
# are adapted in disposable packages. Core authorization and persistence stay real.
function Install-AggregationYieldProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $probe=$admin.VBComponents.Item('TestAggregation').CodeModule
    $probe.AddFromString(@'
Private mYieldPoint As String, mConnectionResult As Boolean, mYieldReturned As Boolean
Private mPriorFaultInjected As Boolean
Private mClearSelectionOnReturn As Boolean
Private mYieldFault As String, mYieldSource As String, mYieldAlternate As String, mYieldFeed As String
Public Sub ArmYield(ByVal point As String, ByVal connectionResult As Boolean, Optional ByVal fault As String = "", Optional ByVal source As String = "", Optional ByVal alternate As String = "", Optional ByVal feed As String = "")
    mYieldPoint = point: mConnectionResult = connectionResult: mYieldReturned = False
    mPriorFaultInjected = False
    mClearSelectionOnReturn = False
    mYieldFault = fault: mYieldSource = source: mYieldAlternate = alternate: mYieldFeed = feed
End Sub
Public Sub ReturnYield(ByVal point As String)
    If mYieldPoint <> point Or mYieldReturned Then Exit Sub
    If mClearSelectionOnReturn Then mForm.SelectSourceForTest ""
    mYieldReturned = True
    mPriorFaultInjected = CBool(Application.Run("'invSys.Core.xlam'!TestAggregationBoundary.Injected"))
    Application.Run "'invSys.Core.xlam'!TestAggregationBoundary.Arm", mYieldFault, point, mYieldSource, mYieldAlternate, mYieldFeed
    Application.Run "'invSys.Core.xlam'!TestAggregationBoundary.AtPoint", point
End Sub
Public Function ConnectionReturn() As Boolean
    ReturnYield "CONNECTION"
    ConnectionReturn = mConnectionResult
End Function
Public Function YieldReturned() As Boolean
    YieldReturned = mYieldReturned
End Function
Public Function PriorFaultInjected() As Boolean
    PriorFaultInjected = mPriorFaultInjected
End Function
Public Sub ClearSelectionOnReturn(ByVal clear As Boolean)
    mClearSelectionOnReturn = clear
End Sub
Public Sub Connect()
    If Not mForm Is Nothing Then mForm.ConnectForTest
End Sub
Public Function NeedsReopen() As Boolean
    If Not mForm Is Nothing Then NeedsReopen = mForm.NeedsReopenForTest()
End Function
Public Function ClaimsCancellation() As Boolean
    If Not mForm Is Nothing Then ClaimsCancellation = mForm.ClaimsCancellationForTest()
End Function
Public Function ClaimsSavedComplete() As Boolean
    If Not mForm Is Nothing Then ClaimsSavedComplete = mForm.ClaimsSavedCompleteForTest()
End Function
Public Function DirectReviewRefusalAccurate() As Boolean
    If Not mForm Is Nothing Then DirectReviewRefusalAccurate = mForm.DirectReviewRefusalAccurateForTest()
End Function
'@)
    $form=$admin.VBComponents.Item('frmAggregationSources').CodeModule
    $form.AddFromString(@'
Public Sub ConnectForTest()
    mTxtRoot.Value = ""
    mBtnConnect_Click
End Sub
Public Function NeedsReopenForTest() As Boolean
    NeedsReopenForTest = InStr(1, mLblStatus.Caption, "Reopen Aggregation", vbTextCompare) > 0
End Function
Public Function ClaimsCancellationForTest() As Boolean
    ClaimsCancellationForTest = InStr(1, mLblStatus.Caption, "cancelled", vbTextCompare) > 0
End Function
Public Function ClaimsSavedCompleteForTest() As Boolean
    ClaimsSavedCompleteForTest = InStr(1, mLblStatus.Caption, "complete output was saved", vbTextCompare) > 0 Or _
        InStr(1, mLblStatus.Caption, "complete advisory output saved", vbTextCompare) > 0
End Function
Public Function DirectReviewRefusalAccurateForTest() As Boolean
    Dim report As String, resolved As Boolean
    resolved = modAggregationCommands.ResolveRetry(mSourceContext, SelectedPaths(), True, report)
    DirectReviewRefusalAccurateForTest = Not resolved And InStr(1, report, "cancelled", vbTextCompare) = 0 And _
        InStr(1, report, "Reopen Aggregation", vbTextCompare) > 0
End Function
'@)
    $first=$form.ProcStartLine('MsgBox',0);$count=$form.ProcCountLines('MsgBox',0)
    $body=$form.Lines($first,$count)
    $body=$body.Replace('    End If',@'
    End If
    If (Buttons And 7) = vbYesNo Then
        TestAggregation.ReturnYield "CONFIRMATION"
    Else
        TestAggregation.ReturnYield "RESULT"
    End If
'@)
    $form.DeleteLines($first,$count);$form.InsertLines($first,$body)
    $connection=$admin.VBComponents.Item('modAdminConsole').CodeModule
    $first=$connection.ProcStartLine('ConnectAggregationServerForAdmin',0);$count=$connection.ProcCountLines('ConnectAggregationServerForAdmin',0)
    # Deterministic Windows connection result only: no credentials or native
    # network effects are introduced. The actual form's before/after guards stay.
    $connection.DeleteLines($first,$count)
    $connection.InsertLines($first,@'
Public Function ConnectAggregationServerForAdmin(ByVal serverRoot As String, ByVal userName As String, ByVal windowsPassword As String, Optional ByRef report As String = "") As Boolean
    report = "Disposable connection return."
    ConnectAggregationServerForAdmin = TestAggregation.ConnectionReturn()
End Function
'@)
    if($AdminAggregationSuppressPhaseRechecksForTest){
        # D13 equivalent seam: preserve captured ownership/entry authorization,
        # suppress only subsequent joint freshness checks in disposable copies.
        $command=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modAggregationCommands').CodeModule
        $first=$command.ProcStartLine('CheckPhase',0);$count=$command.ProcCountLines('CheckPhase',0)
        $body=$command.Lines($first,$count)
        $joint='    CheckPhase = modAggregationAuthority.Joint(context, mSources)'
        if(-not $body.Contains($joint)){throw 'Scoped phase-revalidation seam unavailable.'}
        $body=$body.Replace($joint,"    mSuppressedPhaseChecksForTest = mSuppressedPhaseChecksForTest + 1`r`n    CheckPhase = True")
        $command.DeleteLines($first,$count);$command.InsertLines($first,$body)
        $command.AddFromString(@'
Private mSuppressedPhaseChecksForTest As Long
Public Function SuppressedPhaseChecksForTest() As Long
    SuppressedPhaseChecksForTest = mSuppressedPhaseChecksForTest
End Function
'@)
    }
}

function Test-AggregationYields($Fixture,$Other,$HostBook) {
    if($AdminAggregationYieldGroup -eq 'FeedReopen'){
        Test-AggregationFeedReopen $Fixture $Other $HostBook
        return
    }
    function Call([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestAggregation.'+$Name) $Values}
    function Core([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAggregationBoundary.'+$Name) $Values}
    function SavedHash([string]$Path){(Get-FileHash -LiteralPath $Path).Hash}
    function Receipts { ,@(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -File -Filter 'invSys.Aggregation.*.retry.json' -ErrorAction SilentlyContinue) }
    function OpenActual {
        if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Actual yield handler unavailable; not behavioral RED.'}
    }
    $sourceBytes=[IO.File]::ReadAllBytes($Fixture.Snapshot)
    $sha=[Security.Cryptography.SHA256]::Create()
    try{$sourceHash=([BitConverter]::ToString($sha.ComputeHash($sourceBytes))).Replace('-','')}finally{$sha.Dispose()}
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath)
    $changedFeed=Join-Path $runRoot 'yield-changed-feed'
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($changedFeed).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Yield fault destination escapes disposable scope.'}
    $groups=if($AdminAggregationYieldGroup -eq 'RedFocus'){@('ConnectionFailed','Confirmation','ConfirmationCancelled','ResultFailed','ResultCompletedSelectionChanged')}elseif($AdminAggregationYieldGroup -eq 'PhaseRedFocus'){@('BeforeSave','AfterSave','AfterPromotion')}else{@($AdminAggregationYieldGroup)}
    try {
      foreach($yieldGroup in $groups){
        if($yieldGroup -like 'Connection*'){
            SelectTarget $Fixture
            [void](Core 'Arm' @('','','','',''))
            [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
            if(-not [bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Fixture.Snapshot,$Fixture.Output))){throw 'Connection control feed unavailable.'}
            $before=SavedHash $Fixture.Output
            OpenActual
            [void](Call 'ArmYield' @('CONNECTION',($yieldGroup -eq 'ConnectionSucceeded')))
            [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.ResetError')
            [void](Call 'Connect')
            Check ('Aggregation.Yield.'+$yieldGroup+'.Control.ActualReturnReached') ([bool](Call 'YieldReturned'))
            Check ('Aggregation.Yield.'+$yieldGroup+'.Control.ValidContextRemainsUsable') (-not [bool](Call 'NeedsReopen'))
            Check ('Aggregation.Yield.'+$yieldGroup+'.Control.FeedPreserved') ((SavedHash $Fixture.Output) -ceq $before)
            [void](Call 'CloseForm');[void](Call 'ArmYield' @('',$false))
        }
        $faults=if($AdminAggregationYieldGroup -eq 'RedFocus'){@('SOURCE_CHANGED')}elseif($AdminAggregationYieldGroup -eq 'PhaseRedFocus'){@('SIGN_OUT')}else{@('SOURCE_CHANGED','SOURCE_MISSING','FEED_CHANGED','SIGN_OUT','RIGHTS_REVOKED')}
        foreach($fault in $faults){
            $name='Aggregation.Yield.'+$yieldGroup+'.'+$fault
            SelectTarget $Fixture
            [void](Call 'ArmYield' @('',$false));[void](Core 'Arm' @('','','','',''))
            [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
            if(-not [bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Fixture.Snapshot,$Fixture.Output))){throw 'Known feed prerequisite failed; not behavioral RED.'}
            $before=SavedHash $Fixture.Output
            OpenActual
            $reviewPins=@{}
            if($yieldGroup -in @('Confirmation','ConfirmationCancelled')){
                [void](Core 'Arm' @('WRITE_FAILURE','BEFORE_SAVE',$Fixture.Snapshot,$Other.Snapshot,$changedFeed))
                [void](Call 'Aggregate')
                if(-not [bool](Core 'Injected') -or (Receipts).Count -eq 0){throw 'Pending attempt prerequisite failed; not behavioral RED.'}
                foreach($receipt in (Receipts)){$reviewPins[$receipt.FullName]=SavedHash $receipt.FullName}
                [void](Call 'CloseForm');SelectTarget $Fixture;OpenActual
                [void](Call 'ConfirmRetryReview' @($yieldGroup -eq 'Confirmation'))
            }
            $point=switch($yieldGroup){
                'ConnectionFailed'{'CONNECTION'} 'ConnectionSucceeded'{'CONNECTION'}
                'Confirmation'{'CONFIRMATION'} 'ConfirmationCancelled'{'CONFIRMATION'}
                'ResultCompleted'{'RESULT'} 'ResultCompletedSelectionChanged'{'RESULT'} 'ResultFailed'{'RESULT'}
                'BeforeSave'{'BEFORE_SAVE'} 'AfterSave'{'AFTER_SAVE'} 'AfterPromotion'{'AFTER_PROMOTION'}
            }
            [void](Call 'ArmYield' @($point,($yieldGroup -eq 'ConnectionSucceeded'),$fault,$Fixture.Snapshot,$Other.Snapshot,$changedFeed))
            if($yieldGroup -eq 'ResultCompletedSelectionChanged'){[void](Call 'ClearSelectionOnReturn' @($true))}
            if($AdminAggregationSuppressPhaseRechecksForTest){$phaseBefore=[int](Run 'invSys.Core.xlam' 'modAggregationCommands.SuppressedPhaseChecksForTest')}
            if($yieldGroup -eq 'ResultFailed'){
                [void](Core 'Arm' @('WRITE_FAILURE','BEFORE_SAVE',$Fixture.Snapshot,$Other.Snapshot,$changedFeed))
            }elseif($yieldGroup -in @('BeforeSave','AfterSave','AfterPromotion')){
                [void](Core 'Arm' @($fault,$point,$Fixture.Snapshot,$Other.Snapshot,$changedFeed))
            }else{[void](Core 'Arm' @('','','','',''))}
            if($yieldGroup -like 'Connection*'){
                [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.ResetError')
                [void](Call 'Connect')
            }else{[void](Call 'Aggregate')}
            $returned=if($point -in @('CONNECTION','CONFIRMATION','RESULT')){[bool](Call 'YieldReturned')}else{$true}
            $injected=[bool](Core 'Injected') -and $returned
            Check ($name+'.ActualFaultReturnReached') $injected
            if(-not $injected){throw 'Named yield fault was not injected; not behavioral RED.'}
            if($AdminAggregationSuppressPhaseRechecksForTest){
                $seamReached=[int](Run 'invSys.Core.xlam' 'modAggregationCommands.SuppressedPhaseChecksForTest') -gt $phaseBefore
                Check ($name+'.ScopedFreshnessSeamReached') $seamReached
                if(-not $seamReached){throw 'Phase revalidation seam was not reached; not behavioral RED.'}
            }
            if($yieldGroup -eq 'ResultFailed'){
                $failedSave=[bool](Call 'PriorFaultInjected') -and [int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 1
                Check ($name+'.PrecedingWriteFailureReached') $failedSave
                if(-not $failedSave){throw 'Result did not follow the injected write failure; not behavioral RED.'}
            }
            $complete=$yieldGroup -in @('ResultCompleted','ResultCompletedSelectionChanged','AfterPromotion')
            if($complete -or ($AdminAggregationSuppressPhaseRechecksForTest -and [bool](Call 'OwnerSuccess'))){
                # Compare every exact source identity/quantity, without emitting rows.
                # SOURCE_CHANGED overwrites only the selected fixture; expected rows
                # therefore come from its captured bytes, restored below before read.
                $capture=Join-Path $runRoot 'yield-original-selected.xlsb'
                [IO.File]::WriteAllBytes($capture,$sourceBytes)
                $expected=ReadEntities $capture 'tblInventorySnapshot'
                $actual=ReadEntities $Fixture.Output 'tblGlobalInventorySnapshot'
                $exact=$expected.Count -eq $actual.Count
                foreach($key in $expected.Keys){$exact=$exact -and $actual.ContainsKey($key) -and $actual[$key] -eq $expected[$key]}
                Check ($name+'.CompleteSelectedSetPersisted') $exact
                if(-not $exact){throw 'Unexpected persisted output is unverified; independent case preparation stopped.'}
                Check ($name+'.KnownPersistenceAcknowledged') ([bool](Call 'ClaimsSavedComplete'))
                if($yieldGroup -eq 'AfterPromotion'){Check ($name+'.NoContinuedSuccessAfterContextChange') (-not [bool](Call 'OwnerSuccess'))}
            }
            if(-not $complete){
                Check ($name+'.ExistingFeedPreserved') ((SavedHash $Fixture.Output) -ceq $before)
                Check ($name+'.NoFalseCompletion') (-not [bool](Call 'OwnerSuccess'))
            }
            if($yieldGroup -in @('ConnectionFailed','ConnectionSucceeded','Confirmation','ConfirmationCancelled','ResultCompleted','ResultCompletedSelectionChanged','ResultFailed')){
                Check ($name+'.RequiresReopeningImmediately') ([bool](Call 'NeedsReopen'))
            }
            if($yieldGroup -eq 'Confirmation'){
                Check ($name+'.ConfirmedReviewNotMisreportedCancelled') (-not [bool](Call 'ClaimsCancellation'))
                Check ($name+'.DirectCoreReviewRefusalAccurate') ([bool](Call 'DirectReviewRefusalAccurate'))
            }
            if($yieldGroup -like 'Connection*' -or $yieldGroup -in @('Confirmation','ConfirmationCancelled')){
                Check ($name+'.NoWriterAfterReturn') ([int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
            }
            $pending=Receipts
            if($yieldGroup -in @('Confirmation','ConfirmationCancelled','ResultFailed','BeforeSave','AfterSave','AfterPromotion')){
                Check ($name+'.RetryEvidenceRetained') ($pending.Count -gt 0)
                $bound=$pending.Count -gt 0
                foreach($receipt in $pending){
                    $record=[IO.File]::ReadAllText($receipt.FullName)|ConvertFrom-Json
                    $matching=@($record.Sources|Where-Object {$_.Path -ceq $Fixture.Snapshot -and $_.Hash -ieq $sourceHash})
                    $bound=$bound -and $record.Destination -ceq $Fixture.Output -and $matching.Count -eq 1 -and [bool]$record.Phase -and [bool]$record.Persistence
                }
                Check ($name+'.RetryBindsCapturedDestinationSourceAndPhase') $bound
            }
            if($reviewPins.Count){
                $same=$pending.Count -eq $reviewPins.Count
                foreach($path in $reviewPins.Keys){$same=$same -and (Test-Path $path) -and (SavedHash $path) -ceq $reviewPins[$path]}
                Check ($name+'.InvalidReviewPreservesExactReceipt') $same
            }
            Check ($name+'.NoRedirectedFeed') (-not(Test-Path -LiteralPath (Join-Path $changedFeed 'Global/invSys.Global.InventorySnapshot.xlsb')))
            Check ($name+'.FeedbackExcludesEnteredPaths') ([bool](Call 'FeedbackRedacted' @($Fixture.Root,$changedFeed)))
            [void](Call 'CloseForm');[void](Call 'ArmYield' @('',$false));[void](Core 'Arm' @('','','','',''))
            [IO.File]::WriteAllBytes($Fixture.Snapshot,$sourceBytes)
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            [IO.File]::WriteAllBytes($authPath,$authBytes)
            if($pending.Count){
                SelectTarget $Fixture
                $context=[string](Core 'CaptureDirect' @($HostBook.Name,$Fixture.Warehouse,$Fixture.Config,$Fixture.Snapshot))
                $resolved=[bool](Core 'ResolveKnown' @($context,$Fixture.Snapshot))
                Check ($name+'.ExplicitFreshReviewResolvedKnownState') $resolved
                if(-not $resolved){throw 'Known retry state could not be resolved; independent case preparation stopped.'}
            }
        }
      }
    }finally{
        [void](Call 'CloseForm');[void](Call 'ArmYield' @('',$false));[void](Core 'Arm' @('','','','',''))
        [IO.File]::WriteAllBytes($Fixture.Snapshot,$sourceBytes)
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        [IO.File]::WriteAllBytes($authPath,$authBytes)
        SelectTarget $Fixture
    }
}
