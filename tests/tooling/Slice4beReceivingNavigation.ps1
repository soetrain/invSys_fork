# D18 navigation tests use native messages delivered only to the isolated Excel
# process's focused control. Unsaved seams expose state and trace event names;
# they never manufacture activity or export selected values.
function Get-ReceivingNavigationCases {
    @(
        @('RECEIVING_PAGE_RECEIPTS','tabsReceiving',1,0,'Receiving'),
        @('RECEIVING_PAGE_RETURNS','tabsReceiving',0,1,'Returns'),
        @('RECEIVING_PAGE_PURCHASING','tabsReceiving',1,2,'Purchasing'),
        @('RECEIVING_SELECT_ITEM','lstReceiveItems',0,1,'Receive Item Results'),
        @('DISPOSITION_SELECT_ITEM','lstReceiveItems',1,1,'Return Item Results'),
        @('RECEIVING_SELECT_AGGREGATE','lstAggregate',0,1,'Aggregate Received'),
        @('DISPOSITION_SELECT_AGGREGATE','lstAggregate',1,1,'Aggregate Returns'),
        @('RECEIVING_SELECT_HISTORY','lstInventory',0,1,'Receiving Entries History'),
        @('DISPOSITION_SELECT_HISTORY','lstInventory',1,1,'Return Entries History'),
        @('RECEIVING_SELECT_STAGED','lstStaged',0,1,'Received Tally'),
        @('DISPOSITION_SELECT_STAGED','lstStaged',1,1,'Return Tally'),
        @('RECEIVING_SELECT_CONDITION','cboCondition',0,1,'Condition *'),
        @('DISPOSITION_SELECT_KIND','cboDisposition',1,1,'Disposition *')
    )
}

function Install-ReceivingNavigationSeams($Form,$Helper,$Gate) {
    $Form.AddFromString(@'
Public NavigationTrace As String
Public Sub ActivityTestNavigationPrepare(ByVal pageIndex As Long, ByVal name As String)
    Dim item As Object, i As Long
    mTabs.Value = pageIndex
    ApplyReceivingTab
    If name <> "tabsReceiving" Then
        Set item = Me.Controls(name)
        If name = "lstInventory" Or name = "lstStaged" Or name = "lstAggregate" Then
            item.Clear
            For i = 1 To 3
                item.AddItem "NAVIGATION-PRIVATE-PROJECTION-" & CStr(i)
            Next i
        End If
        If item.ListCount < 2 Then Err.Raise 5, , "Navigation fixture needs two choices."
        item.ListIndex = 0
    End If
    If Not Me.Visible Then Me.Show vbModeless
    Me.Controls(name).SetFocus
    mTxtStatus.Value = "Navigation fixture ready."
    NavigationTrace = ""
End Sub
Public Sub ActivityTestNavigationProgrammatic(ByVal name As String, ByVal index As Long)
    If name = "tabsReceiving" Then Me.Controls(name).Value = index Else Me.Controls(name).ListIndex = index
End Sub
Public Function ActivityTestNavigationIndex(ByVal name As String) As Long
    If name = "tabsReceiving" Then ActivityTestNavigationIndex = mTabs.Value Else ActivityTestNavigationIndex = Me.Controls(name).ListIndex
End Function
Public Function ActivityTestNavigationStatus() As String
    ActivityTestNavigationStatus = CStr(InStr(1, mTxtStatus.Value, "Tracking unavailable", vbTextCompare) > 0) & "|" & _
        CStr(InStr(1, mTxtStatus.Value, "Session or warehouse changed", vbTextCompare) > 0)
End Function
Public Function ActivityTestNavigationGeometry(ByVal name As String) As String
    Dim item As Object
    Set item = Me.Controls(name)
    ActivityTestNavigationGeometry = CStr(item.Left) & "|" & CStr(item.Top) & "|" & CStr(item.Width) & "|" & _
        CStr(item.Height) & "|" & CStr(Me.InsideWidth) & "|" & CStr(Me.InsideHeight)
End Function
' This deliberately bypasses the input boundary to protect internal callers.
Public Sub ActivityTestNavigationInternal()
    LoadSelectedReceiveItemDetails
    ShowSelectedAggregateReferences
    RefreshAllViews
End Sub
' Only fixed event names, no key/button/selection values, reach diagnostics.
Public Sub ActivityTestNavigationTrace(ByVal name As String)
    NavigationTrace = NavigationTrace & name & ";"
End Sub
'@)
    # AddFromString inserts declarations before existing procedures; explicit
    # event instrumentation keeps owner bodies and native dispatch intact.
    foreach($field in @('mTabs','mLstReceiveItems','mLstAggregate','mLstInventory','mLstStaged','mCboCondition','mCboDisposition')) {
        $events=@('Change','Click','MouseDown','MouseUp','KeyDown','KeyUp')
        if($field.StartsWith('mCbo')) { $events+='DropButtonClick' }
        foreach($event in $events) {
            $name=$field+'_'+$event
            $parameters=''
            if($event -in @('MouseDown','MouseUp')) {
                $parameters='ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single'
                if($field -eq 'mTabs') { $parameters='ByVal Index As Long, '+$parameters }
            } elseif($event -in @('KeyDown','KeyUp')) {
                $parameters='ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer'
            } elseif($event -eq 'Click' -and $field -eq 'mTabs') { $parameters='ByVal Index As Long' }
            $start=0
            try { $start=$Form.ProcStartLine($name,0) } catch { }
            $trace='    ActivityTestNavigationTrace "'+$name+'"'
            if($event -eq 'MouseUp' -and $field.StartsWith('mCbo')) {
                $trace+="`r`n"+'    ActivityTestNavigationTrace "ChoiceCommittedAtMouseUp." & CStr('+ $field +'.ListIndex = 1)'
            }
            if($start -gt 0) {
                $body=$Form.ProcBodyLine($name,0)
                while($Form.Lines($body,1).TrimEnd().EndsWith('_')) { $body++ }
                $Form.InsertLines($body+1,$trace)
            } else {
                $Form.AddFromString("Private Sub $name($parameters)`r`n$trace`r`nEnd Sub")
            }
        }
    }
    foreach($owner in @('LoadSelectedReceiveItemDetails','ShowSelectedAggregateReferences','ApplyReceivingTab')) {
        $body=$Form.ProcBodyLine($owner,0)
        $Form.InsertLines($body+1,('    ActivityTestNavigationTrace "Owner.'+$owner+'"'))
    }
    $Helper.AddFromString(@'
Public Sub NavigationPrepare(ByVal pageIndex As Long, ByVal name As String)
    mForm.ActivityTestNavigationPrepare pageIndex, name
End Sub
Public Sub NavigationProgrammatic(ByVal name As String, ByVal index As Long)
    mForm.ActivityTestNavigationProgrammatic name, index
End Sub
Public Function NavigationIndex(ByVal name As String) As Long
    NavigationIndex = mForm.ActivityTestNavigationIndex(name)
End Function
Public Function NavigationTrace() As String
    NavigationTrace = mForm.NavigationTrace
End Function
Public Function NavigationStatus() As String
    NavigationStatus = mForm.ActivityTestNavigationStatus()
End Function
Public Function NavigationGeometry(ByVal name As String) As String
    NavigationGeometry = mForm.ActivityTestNavigationGeometry(name)
End Function
Public Function NavigationHandle() As Double
    NavigationHandle = modReceivingFormWindow.ActivityTestWindowHandle(mForm)
End Function
Public Sub NavigationInternal()
    mForm.ActivityTestNavigationInternal
End Sub
Public Sub NavigationInitialize()
    mForm.InitializeFromReceiving
End Sub
'@)
    $Gate.AddFromString(@'
Public Function NavigationDefinition(ByVal controlId As String) As String
    Dim item As Object
    Set item = modActivityCatalog.Control(controlId)
    If item Is Nothing Then Exit Function
    NavigationDefinition = item("Class") & "|" & item("OwnerId") & "|" & item("Caption") & "|" & item("Capability")
End Function
'@)
}

function Initialize-ReceivingNavigationInput {
    if('ReceivingNavigationInput' -as [type]) { return }
    Add-Type @'
using System; using System.Runtime.InteropServices;
public static class ReceivingNavigationInput {
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
    [StructLayout(LayoutKind.Sequential)] public struct Gui { public int Size,Flags; public IntPtr Active,Focus,Capture,Menu,Move,Caret; public Rect CaretRect; }
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
    [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint t,ref Gui g);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr h,out Rect r);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h,ref Point p);
    [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr h,ref Point p);
    [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point p);
    static uint Owner(IntPtr h) { uint p; GetWindowThreadProcessId(h,out p); return p; }
    public static void Key(IntPtr form,IntPtr excel,int key) {
        uint p; uint t=GetWindowThreadProcessId(form,out p); var g=new Gui(); g.Size=Marshal.SizeOf(g);
        if(p==0 || p!=Owner(excel) || !IsWindowVisible(form) || !GetGUIThreadInfo(t,ref g) || g.Focus==IntPtr.Zero || Owner(g.Focus)!=p)
            throw new Exception("Owned focused navigation control unavailable.");
        if(!PostMessage(g.Focus,0x100,(IntPtr)key,(IntPtr)1) || !PostMessage(g.Focus,0x101,(IntPtr)key,(IntPtr)0xC0000001))
            throw new Exception("Navigation key delivery failed.");
    }
    public static void Mouse(IntPtr form,IntPtr excel,double x,double y,double width,double height) {
        Rect r; if(Owner(form)==0 || Owner(form)!=Owner(excel) || !IsWindowVisible(form) || !GetClientRect(form,out r))
            throw new Exception("Owned navigation form unavailable.");
        var p=new Point { X=(int)(x*(r.Right-r.Left)/width), Y=(int)(y*(r.Bottom-r.Top)/height) };
        ClientToScreen(form,ref p); var target=WindowFromPoint(p);
        if(Owner(target)!=Owner(form)) throw new Exception("Navigation point is not on the owned Excel form.");
        ScreenToClient(target,ref p); var lp=(IntPtr)((p.Y<<16)|(p.X&65535));
        if(!PostMessage(target,0x201,(IntPtr)1,lp) || !PostMessage(target,0x202,IntPtr.Zero,lp))
            throw new Exception("Navigation mouse delivery failed.");
    }
}
'@
}

function Invoke-ReceivingNavigationInput($Case,[string]$Mode='Keyboard',[switch]$MayReject) {
    $handle=[IntPtr][long][double](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationHandle')
    if($Mode -eq 'Keyboard') {
        $key=40
        if($Case[1] -eq 'tabsReceiving') { $key=if($Case[3] -lt $Case[2]) {37}else{39} }
        [ReceivingNavigationInput]::Key($handle,[IntPtr]$excel.Hwnd,$key)
    } else {
        $geometry=([string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationGeometry' @($Case[1]))).Split('|')
        $x=[double]$geometry[0]+30; $y=[double]$geometry[1]+6+13*[int]$Case[3]
        if($Case[1] -eq 'tabsReceiving') {
            $x=[double]$geometry[0]+@(22,62,103)[[int]$Case[3]]
            $y=[double]$geometry[1]+7
        } elseif($Case[1].StartsWith('cbo')) {
            [ReceivingNavigationInput]::Mouse($handle,[IntPtr]$excel.Hwnd,([double]$geometry[0]+[double]$geometry[2]-5),([double]$geometry[1]+[double]$geometry[3]/2),[double]$geometry[4],[double]$geometry[5])
            Start-Sleep -Milliseconds 200
            $y=[double]$geometry[1]+[double]$geometry[3]+6+13*[int]$Case[3]
        }
        [ReceivingNavigationInput]::Mouse($handle,[IntPtr]$excel.Hwnd,$x,$y,[double]$geometry[4],[double]$geometry[5])
    }
    Start-Sleep -Milliseconds 250
    $index=[int](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationIndex' @($Case[1]))
    $trace=[string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationTrace')
    $traceLine='Navigation native trace '+$Case[0]+' '+$Mode+': '+$trace
    Write-Output $traceLine
    Add-Content -LiteralPath $navigationTracePath -Value $traceLine -Encoding UTF8
    $expectedEvent=if($Mode -eq 'Keyboard') {'_KeyDown;'}else{'_MouseDown;'}
    if((-not $MayReject -and $index -ne [int]$Case[3]) -or -not $trace.Contains($expectedEvent)) { throw ('Navigation input did not select its intended control: '+$Case[0]+' '+$Mode+' index='+$index) }
}

function Set-ReceivingNavigationPolicy($Fixture,[bool]$Collect,[int]$Catalog=6,[bool]$Capture=$false) {
    $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
    try {
        $meta=Add-ActivityFixtureTable $cfg 'tblEventTrackingPolicies' @('PolicyVersion','SchemaVersion','CatalogVersion','CreatedAtUTC','CreatedByUserId','DefaultView','ViewerActionPathCaptureEnabled','AdminViewerEventLoggingEnabled','Operator Extra') @(
            ,@(1.0,1.0,[double]$Catalog,'2026-09-08T12:00:00.000Z','config-admin','How-To',$Capture,$true,'preserve navigation extension'))
        $ids=@('ADMIN_SETTINGS_SAVE_VALUE','PRODUCTION_UOM_RETRIEVE','RECEIVING_CONFIRM_WRITES','RECEIVING_ADD_SELECTED','DISPOSITION_ADD_SELECTED','DISPOSITION_CONFIRM','RECEIVING_REFRESH','RECEIVING_CLEAR','RECEIVING_OPEN','RECEIVING_CLOSE')
        $rows=@(); foreach($id in $ids) { $rows+=,@(1.0,$id,$true,$true,$true,'preserve navigation extension') }
        if($Catalog -eq 6) { foreach($case in Get-ReceivingNavigationCases) { $rows+=,@(1.0,$case[0],$Collect,$true,$true,'preserve navigation extension') } }
        $table=Add-ActivityFixtureTable $cfg 'tblEventTrackingControls' @('PolicyVersion','ControlId','Collect','Visible','SequenceEligible','Operator Extra') $rows
        $cfg.Save()
    } finally { $cfg.Close($false) }
}

function Test-ReceivingNavigationRecords($Fixture,$Before,$Case,[string]$Label) {
    $records=@(); $raws=@()
    foreach($path in @(Get-Slice4beActivityFiles $Fixture)) {
        if($path -in $Before) { continue }
        $raw=[IO.File]::ReadAllText($path); $record=$raw|ConvertFrom-Json
        $records+=,$record; $raws+=,$raw
    }
    $attempt=@($records|Where-Object {$_.ControlId -ceq $Case[0] -and $_.OutcomeCode -ceq 'REQUESTED'})
    $outcome=@($records|Where-Object {$_.ControlId -ceq $Case[0] -and $_.OutcomeCode -ceq 'SELECTED'})
    $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
    Check ($Label+'.ExactlyOnePair') $pair
    $truth=$false; $redacted=$pair; $readable=$pair
    if($pair) {
        $truth=$attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $attempt[0].RecordId -cne $outcome[0].RecordId -and
            $attempt[0].DataEffect -ceq 'Unknown' -and $outcome[0].DataEffect -ceq 'Unchanged' -and
            $outcome[0].Severity -ceq 'Info' -and $outcome[0].OwnerId -ceq 'RECEIVING_NAVIGATION' -and
            $outcome[0].EventCode -ceq ($Case[0]+'_SELECTED') -and $outcome[0].Caption -ceq $Case[4]
    }
    foreach($record in $records) {
        $truth=$truth -and @($record.SourceEventRefs).Count -eq 0 -and $record.CatalogVersion -eq 6
        $readable=$readable -and (Get-ActivityRead $record.RecordId).StartsWith('OK|')
    }
    foreach($raw in $raws) { foreach($forbidden in @('NAVIGATION-PRIVATE','ACTIVITY-PRIVATE',$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'ListIndex','KeyCode','mLst','mTabs')) { if($raw.Contains($forbidden)) { $redacted=$false } } }
    Check ($Label+'.CorrelatedUiOnlyFixedCaption') $truth
    Check ($Label+'.NoSelectedValues') $redacted
    Check ($Label+'.ValidatedRead') $readable
}

function Test-ReceivingNavigationActivity($Fixture) {
    Initialize-ReceivingNavigationInput
    $traceName='navigation-'+$Phase.ToLowerInvariant()+'-input-trace.txt'
    if($ReceivingNavigationOnly) { $traceName='diagnostic-'+$traceName }
    $navigationTracePath=Join-Path $reportRoot $traceName
    [IO.File]::WriteAllText($navigationTracePath,'',[Text.UTF8Encoding]::new($false))
    SelectTarget $Fixture 'config-reader'
    $operator=$excel.Workbooks.Add(); $operator.SaveAs((Join-Path $runRoot 'navigation-operator.xlsm'),52)
    $other=$excel.Workbooks.Add(); $other.Worksheets.Item(1).Cells.Item(1,1).Value2='navigation unrelated sentinel'
    $other.SaveAs((Join-Path $runRoot 'navigation-other.xlsm'),52)
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $operatorClosed=$false
    try {
        if(-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name))) { throw 'Navigation fixture actual staging failed.' }
        $extra=(Table $operator 'ReceivedTally').ListColumns.Add(); $extra.Name='Navigation Extra'; $extra.DataBodyRange.Value2='preserve navigation staging'
        $stagedBefore=@(Get-ReceivingFixtureRows (Table $operator 'ReceivedTally')) | ConvertTo-Json -Depth 5 -Compress
        $operator.Save(); $operatorHash=Get-ReceivingFixtureHash $operator.FullName; $otherHash=Get-ReceivingFixtureHash $other.FullName
        foreach($case in Get-ReceivingNavigationCases) {
            $definition=[string](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.NavigationDefinition' @($case[0]))
            Check ('Navigation.Catalog.'+$case[0]) ($definition -ceq ('Navigation|RECEIVING_NAVIGATION|'+$case[4]+'|RECEIVE_POST'))
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationPrepare' @($case[2],$case[1]))
            $before=@(Get-Slice4beActivityFiles $Fixture)
            Invoke-ReceivingNavigationInput $case
            Check ('Navigation.DefaultOff.'+$case[0]) (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        }
        Set-ReceivingNavigationPolicy $Fixture $true
        $authority=Get-ReceivingAuthorityHashes $Fixture
        foreach($case in Get-ReceivingNavigationCases) {
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationPrepare' @($case[2],$case[1]))
            $before=@(Get-Slice4beActivityFiles $Fixture)
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationProgrammatic' @($case[1],$case[3]))
            Check ('Navigation.ProgrammaticExcluded.'+$case[0]) (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationPrepare' @($case[2],$case[1]))
            $other.Activate()
            $before=@(Get-Slice4beActivityFiles $Fixture)
            Invoke-ReceivingNavigationInput $case
            Test-ReceivingNavigationRecords $Fixture $before $case ('Navigation.Keyboard.'+$case[0])
        }
        foreach($case in Get-ReceivingNavigationCases) {
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationPrepare' @($case[2],$case[1]))
            $before=@(Get-Slice4beActivityFiles $Fixture)
            Invoke-ReceivingNavigationInput $case 'Mouse'
            Test-ReceivingNavigationRecords $Fixture $before $case ('Navigation.Mouse.'+$case[0])
        }
        foreach($case in Get-ReceivingNavigationCases) {
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationPrepare' @($case[2],$case[1]))
            $same=$case.Clone(); $same[3]=if($case[1] -eq 'tabsReceiving'){$case[2]}else{0}
            Invoke-ReceivingNavigationInput $same 'Mouse'
            $before=@(Get-Slice4beActivityFiles $Fixture)
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationProgrammatic' @($case[1],$case[3]))
            Check ('Navigation.MouseNoChangeDoesNotArmProgrammatic.'+$case[0]) (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        }
        $case=@(Get-ReceivingNavigationCases)[3]
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationPrepare' @($case[2],$case[1]))
        $before=@(Get-Slice4beActivityFiles $Fixture)
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationInternal')
        Check 'Navigation.InternalRefreshAndDetailsExcluded' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        Check 'Navigation.AuthorityBytes' (Test-ReceivingAuthorityHashes $Fixture $authority)
        $stagedAfter=@(Get-ReceivingFixtureRows (Table $operator 'ReceivedTally')) | ConvertTo-Json -Depth 5 -Compress
        Check 'Navigation.ExactStagingAndUnknownColumn' ($stagedAfter -ceq $stagedBefore)
        Check 'Navigation.OperatorBytes' ($operatorHash -ceq (Get-ReceivingFixtureHash $operator.FullName))
        Check 'Navigation.UnrelatedWorkbook' ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
        if($CaptureEvidence) { CaptureFormEvidence 'Receiving' 'coverage-navigation-receiving.png' }
        foreach($mode in @('DisabledCaptureOn','OlderPolicy','MalformedPolicy','StoreFailure','StaleSession')) {
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            [void](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($Fixture.Warehouse,'S1'))
            if($mode -eq 'OlderPolicy') { Set-ReceivingNavigationPolicy $Fixture $true 5 }
            elseif($mode -eq 'DisabledCaptureOn') { Set-ReceivingNavigationPolicy $Fixture $false 6 $true }
            else { Set-ReceivingNavigationPolicy $Fixture $true }
            if($mode -eq 'MalformedPolicy') {
                $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Cells.Item(1,1).Value2=999.0
                $cfg.Save(); $cfg.Close($false)
            }
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationPrepare' @($case[2],$case[1]))
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $held=$null
            try {
                if($mode -eq 'StoreFailure') { $held=Hold-ReceivingLocalActivityStore $Fixture }
                if($mode -eq 'StaleSession') { [void](Run 'invSys.Core.xlam' 'modAuth.SignOut'); SelectTarget $Fixture 'config-reader' }
                Invoke-ReceivingNavigationInput $case -MayReject:($mode -eq 'StaleSession')
                $status=[string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationStatus')
                if($mode -eq 'StaleSession') {
                    $trace=[string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationTrace')
                    Check 'Navigation.StaleSession.NoDetailOwnerEntry' (-not $trace.Contains('Owner.LoadSelectedReceiveItemDetails'))
                    Check 'Navigation.StaleSession.VisibleReopenRequired' ($status.EndsWith('|True'))
                } elseif($mode -ne 'DisabledCaptureOn') {
                    Check ('Navigation.'+$mode+'.VisibleTrackingUnavailable') ($status.StartsWith('True|'))
                }
            } finally {
                if($null -ne $held) { Remove-Item -LiteralPath $held[0]; Move-Item -LiteralPath $held[1] -Destination $held[0] }
            }
            Check ('Navigation.'+$mode+'.NoInventedRecords') (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        }
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        [void](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($Fixture.Warehouse,'S1'))
        Set-ReceivingNavigationPolicy $Fixture $true
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Reopen' @($operator.Name))
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationInitialize')
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationPrepare' @($case[2],$case[1]))
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $operator.Close($false); $operatorClosed=$true; $other.Activate()
        Invoke-ReceivingNavigationInput $case -MayReject
        $trace=[string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationTrace')
        $status=[string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationStatus')
        Check 'Navigation.ClosedWorkbook.NoDetailOwnerEntry' (-not $trace.Contains('Owner.LoadSelectedReceiveItemDetails'))
        Check 'Navigation.ClosedWorkbook.VisibleReopenRequired' ($status.EndsWith('|True'))
        Check 'Navigation.ClosedWorkbook.NoActivity' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        Check 'Navigation.ClosedWorkbook.NoActiveWorkbookRedirection' ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
    } finally {
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
        if(-not $operatorClosed) { $operator.Close($false) }; $other.Close($false)
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        [void](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($Fixture.Warehouse,'S1'))
    }
}
