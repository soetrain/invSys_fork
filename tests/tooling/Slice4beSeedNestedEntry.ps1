# Disposable builds only. Admin calls back synchronously before returning its real result.
function Install-SeedNestedEntry($runner,$owner) {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1)
    $probe.Name='TestSeedNested'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Runner As Object, InGesture As Boolean
Public Attempts As Long, AtGuard As Long, PastGuard As Long
Public ControlResult As String, FormStable As Boolean, BindingStable As Boolean
Public Function Report() As String
    Report = Attempts & "|" & AtGuard & "|" & PastGuard & "|" & ControlResult & "|" & FormStable & "|" & BindingStable
End Function
'@)
    $core.VBComponents.Item('modExecutionRun').CodeModule.AddFromString(@'
Private Function NestedBindingForTest(ByVal token As String, ByVal stepId As String) As String
    Dim header As Object, records As Collection, pendingStep As Object
    If mToken <> token Or mPending < 1 Or Not mSeedClaimed Or mStop Or mWriteFailed Then Exit Function
    If mContext <> modActivity.CaptureContext() Or mRun("State") <> "Running" Then Exit Function
    Set pendingStep = mProfile("Steps")(mPending)
    If pendingStep("StepId") <> stepId Or pendingStep("ControlId") <> "ADMIN_DEMO_SEED" Then Exit Function
    If Not modRecordingSession.ActiveFor(mContext) Then Exit Function
    If Not modRecordingSession.ExecutionEvidence(mContext, header, records) Then Exit Function
    If header("SequenceId") <> mRun("Recording")("SequenceId") Or records.Count <> mBefore + 2 Then Exit Function
    NestedBindingForTest = Join(Array(mToken, mContext, mKey, mGuide("RecordId"), mProfile("RecordId"), _
        mTarget.RuntimeRoot, mRun("RunId"), mHash, CStr(mPending), stepId, header("SequenceId"), _
        records(records.Count)("RecordId")), vbTab)
End Function

Public Sub NestedOwnerReturningForTest(ByVal token As String, ByVal stepId As String)
    Dim original As Object, binding As String
    If TestSeedNested.Runner Is Nothing Then Exit Sub
    Set original = TestSeedNested.Runner
    Set TestSeedNested.Runner = Nothing
    binding = NestedBindingForTest(token, stepId)
    TestSeedNested.InGesture = True
    original.AttemptNestedEntryForTest
    TestSeedNested.InGesture = False
    TestSeedNested.BindingStable = (binding <> "" And binding = NestedBindingForTest(token, stepId))
End Sub
'@)
    $apply='ApplyBuiltIn = modAdminSeedObservation.Seed(target.WarehouseId, target.StationId, actor, target.RuntimeRoot, "", report, context)'
    Replace-ProbeLine $owner $apply @($apply,'If ApplyBuiltIn Then modExecutionRun.NestedOwnerReturningForTest token, stepId')

    $library=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmActionPathLibrary').CodeModule
    $library.AddFromString(@'
Public Function OwnsRunnerForTest(ByVal instance As Object) As Boolean
    OwnsRunnerForTest = (mRunner Is instance)
End Function
'@)
    if($AdminSeedReplayNested -ceq 'Start'){
        $entryModule=$runner
        $guard='If mExecuting Or (mNeedsEntity And mEntities.ListIndex < 0) Then Exit Sub'
        $gesture=@'
    TestSeedNested.ControlResult = modInventoryViewer.GuideDraftControlForTest("frmActionPathRun", "btnStartRun", "Click")
    ' A disabled button does not fire Click. Exercise the actual handler separately.
    If TestSeedNested.ControlResult = "DISABLED" Then mStart_Click
'@
    }else{
        $entryModule=$library
        $guard='If Not mRunner Is Nothing Then Exit Sub'
        $gesture='    TestSeedNested.ControlResult = modInventoryViewer.GuideDraftControlForTest("frmActionPathLibrary", "btnRunHowTo", "Click")'
    }
    Replace-ProbeLine $entryModule $guard @(
        'If TestSeedNested.InGesture Then TestSeedNested.AtGuard = TestSeedNested.AtGuard + 1',
        $guard,
        'If TestSeedNested.InGesture Then TestSeedNested.PastGuard = TestSeedNested.PastGuard + 1')
    $callback=@'
Private Function OriginalBusyFormForTest() As Boolean
    If Not mExecuting Or Not Me.Visible Or mCloseRequested Or mOwner Is Nothing Then Exit Function
    If Not mOwner.OwnsRunnerForTest(Me) Then Exit Function
    OriginalBusyFormForTest = (modInventoryViewer.GuideDraftControlForTest("frmActionPathRun", "", "Count") = "1")
End Function

Public Sub AttemptNestedEntryForTest()
    Dim busy As Boolean
    SnapshotSeedBoundary
    busy = OriginalBusyFormForTest()
    TestSeedNested.Attempts = TestSeedNested.Attempts + 1
@GESTURE@
    TestSeedNested.FormStable = (busy And OriginalBusyFormForTest())
End Sub
'@
    $runner.AddFromString($callback.Replace('@GESTURE@',$gesture))
}
