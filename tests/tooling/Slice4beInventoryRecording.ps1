# Real Ribbon entry and form handlers in disposable compiled packages. The
# adapter supplies gestures and suppresses only fixed tracking notices.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
. (Join-Path $PSScriptRoot 'Slice4beInventoryRecordingBoundaries.ps1')
function Install-InventoryRecordingProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestInventoryRecordingPolicy'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function SetPolicy(ByVal mode As String) As Boolean
    Dim context As String, version As Long, wire As String, notice As String
    Dim model As Object, row As Variant, user As Object
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, wire, notice) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    model("ViewerActionPathCaptureEnabled") = (mode = "Capture")
    For Each row In model("Controls")
        If Left$(row("ControlId"), 16) = "ADMIN_INVENTORY_" Then row("Collect") = (mode <> "Disabled" And mode <> "ControlOff")
        If mode = "Capture" Then row("SequenceEligible") = (Left$(row("ControlId"), 16) = "ADMIN_INVENTORY_")
    Next row
    If mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", "config-admin": user.Add "Record", False
        model("Users").Add user
    End If
    SetPolicy = modTrackingPolicySettings.SavePolicy(context, version, modTrainingJson.EncodeObject(model), notice)
End Function
Public Function Definition(ByVal id As String, ByVal version As Long) As String
    Dim model As Object
    Set model = modActivityCatalog.Control(id, version)
    If Not model Is Nothing Then Definition = modTrainingJson.EncodeObject(model)
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(43)
        Set before = modActivityCatalog.Control(CStr(id), 43)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    PriorDefinitionsPreserved = True
End Function
'@)
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $probe=$admin.VBComponents.Add(1);$probe.Name='TestInventoryRecording'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mDriving As Boolean, mMode As String, mGestures As Long, mNotices As Long
Public Sub Execute(ByVal host As String, ByVal mode As String)
    mDriving = True: mMode = mode: mGestures = 0: mNotices = 0
    TestInventoryEntry.Execute host, "Recording", "", "", "", "", ""
    mDriving = False
End Sub
Public Function Driving() As Boolean
    Driving = mDriving
End Function
Public Sub Drive(ByVal form As frmAddInventoryItem)
    If mGestures <> 0 Then Exit Sub
    mGestures = 1
    If mMode = "SignedOut" Then modAuth.SignOut
    form.RecordingNavigationForTest mMode
End Sub
Public Function Gestures() As Long
    Gestures = mGestures
End Function
Public Function Notices() As Long
    Notices = mNotices
End Function
Public Function TrackingNotice(ByVal text As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "") As VbMsgBoxResult
    If Left$(text, 21) = "Tracking unavailable:" Then mNotices = mNotices + 1
    TrackingNotice = vbOK
End Function
'@)
    $driver=$admin.VBComponents.Item('TestInventoryEntry').CodeModule
    $first=$driver.ProcStartLine('Drive',0)
    $driver.InsertLines($first+2,'    If TestInventoryRecording.Driving Then TestInventoryRecording.Drive form: Exit Sub')
    $form=$admin.VBComponents.Item('frmAddInventoryItem').CodeModule
    $form.AddFromString(@'
Public Sub RecordingNavigationForTest(ByVal mode As String)
    Dim cancel As Integer
    mBtnEditMode_Click
    mBtnAddMode_Click
    mBtnEditMode_Click
    If mode = "InterruptBeforeClose" Then modAuth.SignOut
    If mode = "WindowClose" Then
        UserForm_QueryClose cancel, 0
    Else
        mBtnCancel_Click
    End If
End Sub
'@)
    $activity=$admin.VBComponents.Item('cProvisioningActivity').CodeModule
    $text=$activity.Lines(1,$activity.CountOfLines)
    $noticePattern='(?i)MsgBox\s+notice,\s+vbExclamation,\s+"invSys Admin"'
    if([regex]::Matches($text,$noticePattern).Count -ne 1){throw 'Fixed tracking notice boundary unavailable; not product RED.'}
    $text=[regex]::Replace($text,$noticePattern,'TestInventoryRecording.TrackingNotice notice, vbExclamation, "invSys Admin"')
    $activity.DeleteLines(1,$activity.CountOfLines);$activity.AddFromString($text)
    Install-InventoryRecordingBoundaries $probe $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestInventoryRecordingPolicy')
}

function Test-InventoryRecording($Fixture,$Other) {
    SelectTarget $Fixture
    $hostBook=$excel.Workbooks.Add()
    $hostBook.SaveAs((Join-Path $runRoot 'inventory-recording-host.xlsx'),51)
    $source=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $hash=Get-AdminInventoryEntryHash $source
    $ids=@('OPEN','ADD_MODE','EDIT_MODE','CLOSE')
    Check 'InventoryRecording.Prior186DefinitionsPreserved' ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.PriorDefinitionsPreserved'))
    foreach($id in $ids){
        $wire=[string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @(('ADMIN_INVENTORY_'+$id),44))
        $model=if($wire){$wire|ConvertFrom-Json}else{$null}
        Check ('InventoryRecording.Catalog.'+$id) ($null -ne $model -and $model.OwnerId -ceq 'ADMIN_INVENTORY' -and $model.Capability -ceq 'ADMIN_MAINT' -and $model.Class -ceq 'Navigation')
        Check ('InventoryRecording.Version43Excludes.'+$id) ([string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @(('ADMIN_INVENTORY_'+$id),43)) -ceq '')
    }
    try {
        if($AdminInventoryWorkbookBindingOnly){
            Test-InventoryRecordingBoundaries $Fixture $Other $hostBook @('ReplacedWorkbook','ClosedWorkbook')
            return
        }
        foreach($mode in @('Enabled','WindowClose','Disabled','ControlOff','UserOff','SignedOut')){
            SelectTarget $Fixture
            $policy=if($mode -cin @('Disabled','ControlOff','UserOff')){$mode}else{'Enabled'}
            $ready=[bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.SetPolicy' @($policy))
            Check ('InventoryRecording.'+$mode+'.PolicyReady') $ready
            if(-not $ready){throw 'Inventory recording policy fixture unavailable; not product RED.'}
            $before=@(Get-Slice4beActivityFiles $Fixture)
            [void](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Execute' @($hostBook.Name,$mode))
            Check ('InventoryRecording.'+$mode+'.ActualHandlersReached') ([long](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Gestures') -eq 1)
            $rows=@();$raws=@()
            foreach($file in @(Get-Slice4beActivityFiles $Fixture)){
                if($file -cin $before){continue}
                $raw=[IO.File]::ReadAllText($file);$row=$raw|ConvertFrom-Json
                if($row.ControlId -like 'ADMIN_INVENTORY_*'){$rows+=$row;$raws+=$raw}
            }
            if($mode -cin @('Disabled','ControlOff','UserOff')){
                Check ('InventoryRecording.'+$mode+'.Suppressed') ($rows.Count -eq 0)
            }else{
                $expected=if($mode -ceq 'SignedOut'){@('OPEN')}else{$ids}
                Check ('InventoryRecording.'+$mode+'.ExactRecordCount') ($rows.Count -eq $(if($mode -ceq 'SignedOut'){2}else{10}))
                foreach($id in $expected){
                    $suffix=switch($id){'OPEN'{'OPENED'} 'CLOSE'{'CLOSED'} default{'SELECTED'}}
                    $attempts=@($rows|Where-Object {$_.ControlId -ceq ('ADMIN_INVENTORY_'+$id) -and $_.OutcomeCode -ceq 'REQUESTED'})
                    $ends=@($rows|Where-Object {$_.ControlId -ceq ('ADMIN_INVENTORY_'+$id) -and $_.OutcomeCode -ceq $suffix})
                    $count=if($id -ceq 'EDIT_MODE'){2}else{1}
                    Check ('InventoryRecording.'+$mode+'.'+$id+'.AttemptsAndResults') ($attempts.Count -eq $count -and $ends.Count -eq $count)
                    $valid=$attempts.Count -eq $count -and $ends.Count -eq $count
                    foreach($attempt in $attempts){
                        $end=@($ends|Where-Object ActivityId -CEQ $attempt.ActivityId)
                        $valid=$valid -and $end.Count -eq 1 -and $attempt.DataEffect -ceq 'Unknown'
                        if($end.Count -eq 1){$valid=$valid -and $end[0].RecordId -cne $attempt.RecordId -and $end[0].DataEffect -ceq 'Unchanged'}
                    }
                    Check ('InventoryRecording.'+$mode+'.'+$id+'.DistinctCorrelatedOutcomes') $valid
                }
                $context=$rows.Count -gt 0;$redacted=$context
                foreach($row in $rows){$context=$context -and $row.OwnerId -ceq 'ADMIN_INVENTORY' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.UserId -ceq 'config-admin'}
                foreach($raw in $raws){foreach($forbidden in @($Fixture.Root,$Fixture.Secret,'mBtn','inventory-recording-host.xlsx')){if($raw.Contains($forbidden)){$redacted=$false}}}
                Check ('InventoryRecording.'+$mode+'.CapturedOwnerContext') $context
                Check ('InventoryRecording.'+$mode+'.Redacted') $redacted
            }
            Check ('InventoryRecording.'+$mode+'.InventoryBytesPreserved') ((Get-AdminInventoryEntryHash $source) -ceq $hash)
        }
        Test-InventoryRecordingBoundaries $Fixture $Other $hostBook
    }finally{
        foreach($book in @($excel.Workbooks)){
            if($book.FullName -ceq (Join-Path $runRoot 'inventory-recording-host.xlsx')){$book.Close($false)}
        }
    }
}
