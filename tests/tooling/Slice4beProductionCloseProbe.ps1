# Unsaved adapters expose existing form actions; no dismissal handler is replaced.
function Install-ProductionCloseProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('mProduction').CodeModule
    $start=$owner.ProcStartLine('HandleProductionOperatorWorkbookClosing',0)
    $count=$owner.ProcCountLines('HandleProductionOperatorWorkbookClosing',0)
    $lines=$owner.Lines($start,$count) -split '\r?\n'
    $entry=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Trim() -ceq 'If operatorWb Is Nothing Then Exit Sub'){$start+$i}})
    if($entry.Count -ne 1){throw 'Unique workbook-close observation boundary unavailable.'}
    $owner.InsertLines($entry[0],'    TestProductionDesigner.CloseWorkbookEntriesForTest = TestProductionDesigner.CloseWorkbookEntriesForTest + 1')
    $owner.AddFromString(@'
Public Function CloseBindingForTest(ByVal workbookName As String) As Boolean
    On Error Resume Next
    If Not mProductionOperatorWorkbook Is Nothing Then CloseBindingForTest = (mProductionOperatorWorkbook Is Application.Workbooks(workbookName))
End Function
'@)
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub CloseButtonForTest()
    mBtnClose_Click
End Sub
Public Function CloseBoundWorkbookForTest() As String
    On Error Resume Next
    If Not mOperatorWorkbook Is Nothing Then CloseBoundWorkbookForTest = mOperatorWorkbook.Name
End Function
Public Function CloseCaptionForTest() As String
    CloseCaptionForTest = mBtnClose.Caption
End Function
Public Sub ClosePrepareWorkbenchForTest()
    mBtnUomCatalogSend_Click
End Sub
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.InsertLines(1,'Public CloseWorkbookEntriesForTest As Long')
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function CloseWorkbookEntryCountForTest() As Long
    CloseWorkbookEntryCountForTest = CloseWorkbookEntriesForTest
End Function
Public Function CloseLoadedFormsForTest() As Long
    Dim loaded As Object
    For Each loaded In VBA.UserForms
        If TypeName(loaded) = "frmProduction" Then CloseLoadedFormsForTest = CloseLoadedFormsForTest + 1
    Next loaded
End Function
Public Function CloseBindPublicForTest() As String
    Dim loaded As Object
    If CloseLoadedFormsForTest() <> 1 Then Err.Raise 5, , "One launched fixture form required."
    Set mForm = Nothing
    For Each loaded In VBA.UserForms
        If TypeName(loaded) = "frmProduction" Then
            Set mForm = loaded
            CloseBindPublicForTest = mForm.CloseBoundWorkbookForTest()
            Exit Function
        End If
    Next loaded
End Function
Public Sub CloseForgetForTest()
    Set mForm = Nothing
End Sub
Public Sub CloseShowForTest()
    mForm.Show vbModeless
End Sub
Public Sub CloseButtonForTest()
    mForm.CloseButtonForTest
End Sub
Public Function CloseCaptionForTest() As String
    CloseCaptionForTest = mForm.CloseCaptionForTest()
End Function
Public Sub ClosePrepareWorkbenchForTest()
    mForm.ClosePrepareWorkbenchForTest
End Sub
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function ClosePolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    ClosePolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function CloseTerminalForTest(ByVal wire As String) As Boolean
    CloseTerminalForTest = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
Public Function CloseCanProduceForTest() As Boolean
    CloseCanProduceForTest = modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST")
End Function
'@)
}

function Close-ProductionNativeFixture {
    if(-not ('ProductionCloseWindow' -as [type])){
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class ProductionCloseWindow {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint id);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h,uint message,IntPtr w,IntPtr l);
}
'@
    }
    $window=[InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd)
    [uint32]$owner=0;[uint32]$actual=0
    [void][ProductionCloseWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$owner)
    [void][ProductionCloseWindow]::GetWindowThreadProcessId($window,[ref]$actual)
    if(-not $owner -or $owner -ne $actual -or -not [ProductionCloseWindow]::IsWindowVisible($window)){throw 'Owned visible Production fixture unavailable; not product RED.'}
    if(-not [ProductionCloseWindow]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)){throw 'Native fixture close delivery failed; not product RED.'}
    $until=[DateTime]::UtcNow.AddSeconds(10)
    while([ProductionCloseWindow]::IsWindow($window) -and [DateTime]::UtcNow -lt $until){Start-Sleep -Milliseconds 100}
    return -not [ProductionCloseWindow]::IsWindow($window)
}

function Invoke-ProductionCloseWithNotice {
    param([ValidateSet('Button','Native')][string]$Mode)
    # The standard observer only dismisses single-OK informational dialogs owned
    # by this disposable Excel process. Raw dialog text stays in worker memory.
    $ready=Join-Path $runRoot ('close-notice-'+[guid]::NewGuid().ToString('N')+'.ready')
    $observer=Join-Path $repo 'tools/plan022-dialog-observer.ps1'
    $handle=[long]$excel.Hwnd
    $job=Start-Job -ArgumentList $observer,$handle,$ready -ScriptBlock {
        param($observer,$handle,$ready)
        $ErrorActionPreference='Stop'
        . $observer
        Invoke-Plan022NativeDialogObservation -ProcessId 0 -TimeoutSeconds 0
        [uint32]$owner=0
        [void][Plan022NativeDialogs]::GetWindowThreadProcessId([IntPtr]$handle,[ref]$owner)
        if(-not $owner){throw 'Owned notice process unavailable.'}
        $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        [IO.File]::WriteAllText($ready,'Ready')
        $matched=$false;$until=[DateTime]::UtcNow.AddSeconds(12)
        while([DateTime]::UtcNow -lt $until){
            if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Notice process identity changed.'}
            $texts=@([Plan022NativeDialogs]::Poll($owner))
            if(@($texts|Where-Object{$_ -like '*ControlType.Text*Tracking unavailable*'}).Count){$matched=$true}
            Start-Sleep -Milliseconds 100
        }
        [pscustomobject]@{TrackingNoticeVisible=$matched}
    }
    try {
        for($i=0;$i -lt 100 -and -not(Test-Path -LiteralPath $ready);$i++){Start-Sleep -Milliseconds 100}
        if(-not(Test-Path -LiteralPath $ready)){throw 'Notice observer not ready; not product RED.'}
        $dismissed=$true
        if($Mode -ceq 'Native'){$dismissed=Close-ProductionNativeFixture}
        else{[void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.CloseButtonForTest')}
        [void](Wait-Job $job -Timeout 15)
        $result=Receive-Job $job -ErrorAction Stop
        if($null -eq $result){throw 'Notice observer result unavailable.'}
        [pscustomobject]@{Dismissed=$dismissed;TrackingNoticeVisible=$result.TrackingNoticeVisible}
    }finally{if($job.State -eq 'Running'){Stop-Job $job};Remove-Job $job -Force}
}
