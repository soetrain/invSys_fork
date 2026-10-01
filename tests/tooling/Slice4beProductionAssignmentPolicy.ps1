# Additional unsaved seams. Existing handlers remain the measured entry points.
function Test-ProductionAssignmentContract {
    $positive=[ordered]@{REFRESH='REFRESHED';PROCESS='PRESENTED';REQUIREMENT='SELECTED';ADD='STAGED';REMOVE='STAGED';CLEAR='STAGED';SAVE='CONFIRMED';PROCESS_SELECT='PRESENTED';REQUIREMENT_SELECT='SELECTED'}
    $submitted='[{"WarehouseId":"CATALOG_TEST","SourceKind":"Designs","EventId":"Source_A","SubmissionState":"Submitted"}]'
    $unknown=$submitted.Replace('Submitted','Unknown')
    $codes=@('REQUESTED','DENIED','REJECTED','FAILED','REFRESHED','PRESENTED','SELECTED','STAGED','CONFIRMED','PENDING','APPLIED','COMPLETED','VALIDATED','CANCELLED')
    foreach($action in $positive.Keys){
        $id='PRODUCTION_ASSIGNMENT_'+$action;$label='AssignmentContract.'+$action
        $outcomes=[ordered]@{REQUESTED=@('Info','Unknown');DENIED=@('Blocked','Unchanged');REJECTED=@('Warning','Unchanged');FAILED=@('Error','Unknown')}
        $outcomes[$positive[$action]]=@('Info',$(if($action -ceq 'SAVE'){'Unknown'}else{'Unchanged'}))
        if($action -ceq 'SAVE'){$outcomes.PENDING=@('Notice','Unknown')}
        foreach($code in $codes){
            $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$code))
            $value=if($wire){$wire|ConvertFrom-Json}else{$null};$supported=$outcomes.Contains($code)
            $correct=if($supported){$null -ne $value -and $value.EventCode -ceq ($id+'_'+$code) -and $value.OutcomeCode -ceq $code -and $value.Severity -ceq $outcomes[$code][0] -and $value.DataEffect -ceq $outcomes[$code][1] -and $value.UserMessage -ne ''}else{$wire -ceq ''}
            Check ($label+'.Outcome.'+$code) $correct
            $record=@{ControlId=$id;OwnerId='PRODUCTION_ASSIGNMENT';CatalogVersion=23;OutcomeCode=$code}|ConvertTo-Json -Compress
            Check ($label+'.Terminal.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @($record)) -eq ($code -ceq $positive[$action]))
            $acceptEmpty=$supported -and $code -cnotin @('CONFIRMED','PENDING')
            Check ($label+'.EmptyReferences.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,'[]')) -eq $acceptEmpty)
        }
        foreach($code in @('REQUESTED','DENIED','REJECTED','FAILED',$positive[$action],'PENDING')|Select-Object -Unique){
            foreach($state in @('Submitted','Unknown')){
                $refs=if($state -ceq 'Submitted'){$submitted}else{$unknown}
                $accept=$action -ceq 'SAVE' -and ($code -ceq 'FAILED' -or ($code -cin @('CONFIRMED','PENDING') -and $state -ceq 'Submitted'))
                Check ($label+'.'+$state+'References.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,$refs)) -eq $accept)
            }
        }
        foreach($change in @('Owner','Catalog')){
            $record=@{ControlId=$id;OwnerId='PRODUCTION_ASSIGNMENT';CatalogVersion=23;OutcomeCode=$positive[$action]}
            if($change -ceq 'Owner'){$record.OwnerId='PRODUCTION_DESIGNER'}else{$record.CatalogVersion=22}
            Check ($label+'.TerminalRejectsWrong'+$change) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($record|ConvertTo-Json -Compress))))
        }
    }
    $invalid=[ordered]@{
        Duplicate=$submitted.Substring(0,$submitted.Length-1)+','+$submitted.Substring(1)
        CrossWarehouse=$submitted.Replace('CATALOG_TEST','OTHER_TEST')
        Inventory=$submitted.Replace('Designs','Inventory')
        InvalidIdentity=$submitted.Replace('Source_A','Source A')
        OversizedIdentity=$submitted.Replace('Source_A',('A'*129))
        MissingIdentity=$submitted.Replace('"EventId":"Source_A",','')
        NumericIdentity=$submitted.Replace('"Source_A"','123')
        ExtraField=$submitted.Replace('"EventId":','"Extra":"forbidden","EventId":')
        UnknownState=$submitted.Replace('Submitted','Applied')
        NonArray='{}'
        Malformed='['
    }
    foreach($case in $invalid.Keys){Check ('AssignmentContract.SAVE.Reject.'+$case) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('PRODUCTION_ASSIGNMENT_SAVE','FAILED',$invalid[$case])))}
}

function Install-ProductionAssignmentPolicyProbe {
    function Seam($Module,[string]$Procedure,[string]$Anchor,[string]$Text,[bool]$After=$false){
        $start=$Module.ProcStartLine($Procedure,0);$count=$Module.ProcCountLines($Procedure,0)
        $lines=@(for($i=$start;$i -lt $start+$count;$i++){if($Module.Lines($i,1).Trim() -ieq $Anchor){$i}})
        if($lines.Count -ne 1){throw ('Assignment fixture seam changed: '+$Procedure+'; not behavioral RED.')}
        $Module.InsertLines($lines[0]+[int]$After,$Text)
    }
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,@'
Private mAssignmentNestedArmed As Boolean, mAssignmentNestedReached As Boolean
'@)
    Seam $form 'SubmitDesignerAction' 'RefreshReusableDesignLists' @'
    If mAssignmentNestedArmed Then
        mAssignmentNestedArmed = False: mAssignmentNestedReached = True
        Call AssignmentActForTest("SAVE")
    End If
'@
    $form.AddFromString(@'
Public Sub AssignmentArmNestedForTest()
    mAssignmentNestedArmed = True: mAssignmentNestedReached = False
End Sub
Public Function AssignmentNestedReachedForTest() As Boolean
    AssignmentNestedReachedForTest = mAssignmentNestedReached
End Function
Public Function AssignmentGuardsRestoredForTest() As Boolean
    AssignmentGuardsRestoredForTest = Not mLoading And Not mDesignerActionInProgress
End Function
Public Sub AssignmentShowForTest()
    Me.Show vbModeless
End Sub
Public Function AssignmentClosedSaveForTest() As String
    TestProductionDesigner.AssignmentClosedEnteredForTest = True
    AssignmentClosedSaveForTest = AssignmentActForTest("SAVE")
End Function
'@)
    $adapter=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Public AssignmentClosedEnteredForTest As Boolean
Private mAssignmentClosedError As Long
Private mAssignmentCapturedBook As Workbook, mAssignmentCapturedContext As String
'@)
    $adapter.AddFromString(@'
Public Sub AssignmentArmNested()
    mForm.AssignmentArmNestedForTest
End Sub
Public Function AssignmentNestedReached() As Boolean
    AssignmentNestedReached = mForm.AssignmentNestedReachedForTest()
End Function
Public Function AssignmentGuardsRestored() As Boolean
    AssignmentGuardsRestored = mForm.AssignmentGuardsRestoredForTest()
End Function
Public Sub AssignmentShowAndCapture(ByVal workbookName As String)
    Set mAssignmentCapturedBook = Application.Workbooks(workbookName)
    mAssignmentCapturedContext = modActivity.CaptureContext()
    mForm.AssignmentShowForTest
End Sub
Public Function AssignmentClosedSave() As String
    On Error GoTo Failed
    AssignmentClosedEnteredForTest = False: mAssignmentClosedError = 0
    AssignmentClosedSave = mForm.AssignmentClosedSaveForTest()
    Exit Function
Failed:
    mAssignmentClosedError = Err.Number
End Function
Public Function AssignmentClosedStatus() As String
    AssignmentClosedStatus = CStr(AssignmentClosedEnteredForTest) & "|" & CStr(mAssignmentClosedError)
End Function
Public Function AssignmentClosedBindingCurrent() As Boolean
    AssignmentClosedBindingCurrent = modProductionDesignerActions.ContextIsCurrent(mAssignmentCapturedContext, mAssignmentCapturedBook)
End Function
Public Sub AssignmentSafeClose()
    On Error Resume Next
    Unload mForm: Set mForm = Nothing
    Set mAssignmentCapturedBook = Nothing
End Sub
'@)
    $writer=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modRoleEventWriter').CodeModule
    $writer.InsertLines($writer.CountOfDeclarationLines+1,'Private mAssignmentSignoutReached As Boolean')
    Seam $writer 'AppendInboxRowToLocalStagingRole' 'AppendInboxRowToLocalStagingRole = True' @'
    If mLifecycleTestMode = "SignOutAfterAppend" Then
        mAssignmentSignoutReached = True
        modAuth.SignOut
    End If
'@ $true
    $writer.AddFromString(@'
Public Function AssignmentSignoutReachedForTest() As Boolean
    AssignmentSignoutReachedForTest = mAssignmentSignoutReached
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function AssignmentDefaultForTest(ByVal id As String) As String
    Dim model As Object, row As Variant
    Set model = modTrackingPolicyModel.Defaults(True)
    For Each row In model("Controls")
        If CStr(row("ControlId")) = id Then
            AssignmentDefaultForTest = CStr(row("Collect")) & "|" & CStr(row("Visible")) & "|" & CStr(row("SequenceEligible"))
            Exit Function
        End If
    Next row
End Function
'@)
}

function Test-ProductionAssignmentPolicy($Fixture,$Other,$Book,$Sheet,[string]$Canary){
    # Probe, Files, Hash and SourceRows come from the companion's test scope.
    function Arm([string]$Mode='Observe'){
        $queue=Join-Path $runRoot ('assignment-policy-queue-'+[guid]::NewGuid().ToString('N'))
        [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @($queue,$Mode,$Fixture.Warehouse))
        [void](Probe 'LifecycleFaultMode' @('Observe'))
    }
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Actor,[string]$Label){
        $rows=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
        $first=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($rows|Where-Object OutcomeCode -CEQ $Outcome)
        $ok=$rows.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        if($ok){$ok=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $last[0].ControlId -ceq ('PRODUCTION_ASSIGNMENT_'+$Action) -and $last[0].UserId -ceq $Actor -and $last[0].OwnerId -ceq 'PRODUCTION_ASSIGNMENT' -and $last[0].WarehouseId -ceq $Fixture.Warehouse}
        Check ($Label+'.OneExactPair') $ok
        if($Outcome -ceq 'DENIED'){Check ($Label+'.PreMutationDenial') ($ok -and $last[0].Severity -ceq 'Blocked' -and $last[0].DataEffect -ceq 'Unchanged' -and @($last[0].SourceEventRefs).Count -eq 0)}
        if($Outcome -in @('FAILED','REJECTED')){
            $severity=if($Outcome -ceq 'FAILED'){'Error'}else{'Warning'};$effect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            Check ($Label+'.LocalPreparationNotRollback') ($ok -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and @($last[0].SourceEventRefs).Count -eq 0)
        }
        if($Outcome -ceq 'CONFIRMED'){
            $writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
            $exact=$ok
            if($ok){$refs=@($last[0].SourceEventRefs);$exact=$writer.Count -eq 5 -and $writer[1] -ceq '1' -and $refs.Count -eq 1 -and $refs[0].EventId -ceq $writer[0] -and $refs[0].SourceKind -ceq 'Designs' -and $refs[0].SubmissionState -ceq 'Submitted' -and $last[0].DataEffect -ceq 'Unknown'}
            Check ($Label+'.OneExactSubmittedReference') $exact
        }
    }
    $actions=@('REFRESH','PROCESS','REQUIREMENT','ADD','REMOVE','CLEAR','SAVE','PROCESS_SELECT','REQUIREMENT_SELECT')
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $guardBook=$null
    try{
        SelectTarget $Fixture 'config-reader';[void](Probe 'ReadReopen' @($Book.Name))
        foreach($action in $actions){
            [void](Probe 'AssignmentStage' @('Normal'));$state=[string](Probe 'AssignmentState');$before=@(Files);Arm
            $notice=[string](Probe 'AssignmentAct' @($action));$writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
            $label='AssignmentPolicy.Denied.'+$action
            Check ($label+'.NoReadsOrMutation') ([string](Probe 'AssignmentState') -ceq $state -and [int](Probe 'ReadCalls') -eq 0)
            Check ($label+'.NoOwningWrite') ($writer.Count -eq 5 -and $writer[1] -ceq '0')
            Pair $before $action 'DENIED' 'config-reader' $label
        }
        foreach($mode in @('Off','Older','NavigationOff','Unavailable')){
            $blocked=Join-Path (Join-Path $Fixture.Root 'Training\Activity') $Fixture.Warehouse;$held=$blocked+'-assignment-held';$moved=$false
            foreach($path in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($path).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Policy fixture escaped disposable root.'}}
            if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
            try{
                SelectTarget $Fixture
                if($mode -in @('Off','NavigationOff')){
                    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.AssignmentPolicyForTest' @(($mode -cne 'Off'),$false))){throw 'Authorized policy fixture unavailable.'}
                }
                if($mode -ceq 'Older'){
                    $olderIds=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(22))).Split([char]10)|Where-Object{$_})
                    if($olderIds.Count -ne 109 -or @($olderIds|Sort-Object -Unique).Count -ne 109){throw 'Catalog22 fixture definitions unavailable; not product RED.'}
                    $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                    try{
                        $headers=Table $cfg 'tblEventTrackingPolicies'
                        $headers.ListColumns.Item('CatalogVersion').DataBodyRange.Value2=22.0
                        $controls=Table $cfg 'tblEventTrackingControls'
                        $removed=0
                        # Later catalogs can add other families; retain the exact declared catalog.
                        for($i=$controls.ListRows.Count;$i -ge 1;$i--){
                            $id=[string]$controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2
                            if($id -cnotin $olderIds){$controls.ListRows.Item($i).Delete();$removed++}
                        }
                        $byVersion=@{}
                        foreach($row in $headers.ListRows){
                            $version=[string]$row.Range.Cells.Item(1,$headers.ListColumns.Item('PolicyVersion').Index).Value2
                            $byVersion[$version]=[Collections.Generic.List[string]]::new()
                        }
                        foreach($row in $controls.ListRows){
                            $version=[string]$row.Range.Cells.Item(1,$controls.ListColumns.Item('PolicyVersion').Index).Value2
                            if(-not $byVersion.ContainsKey($version)){throw 'Older fixture has an unknown policy version; not product RED.'}
                            $byVersion[$version].Add([string]$row.Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2)
                        }
                        foreach($ids in $byVersion.Values){
                            if($ids.Count -ne 109 -or @($ids|Sort-Object -Unique).Count -ne 109 -or @($ids|Where-Object{$_ -cnotin $olderIds}).Count){throw 'Older fixture catalog membership incomplete; not product RED.'}
                        }
                        $cfg.Save()
                        [pscustomobject]@{CatalogVersion=22;RegisteredControls=109;PolicyVersions=$byVersion.Count;RetainedRows=$controls.ListRows.Count;RemovedUnsupportedRows=$removed;ExactMembershipPerPolicy=$true}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'assignment-older-policy-fixture.json')
                    }finally{$cfg.Close($false)}
                }
                SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($Book.Name))
                if($mode -ceq 'Older'){Check 'AssignmentPolicy.Older.ExistingControlReadable' (([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_CLOSE'))).StartsWith('True|True|'))}
                if($mode -ceq 'Unavailable'){
                    if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
                    [IO.File]::WriteAllText($blocked,'Blocked disposable Assignment activity path')
                }
                $selected=if($mode -ceq 'NavigationOff'){@('PROCESS_SELECT','REQUIREMENT_SELECT')}else{$actions}
                foreach($action in $selected){
                    [void](Probe 'AssignmentStage' @('Normal'));Arm;$before=@(Files);$notice=[string](Probe 'AssignmentAct' @($action))
                    $label='AssignmentPolicy.'+$mode+'.'+$action
                    Check ($label+'.AuthorizedActionContinues') ([bool](Probe 'AssignmentPreserved' @($action,'Normal')))
                    Check ($label+'.NoActivityOrFallback') (@(Files).Count -eq $before.Count)
                    if($mode -ceq 'Unavailable'){Check ($label+'.NoticeVisible') $notice.Contains('Tracking unavailable')}
                }
            }finally{
                if($mode -ceq 'Unavailable' -and (Test-Path -LiteralPath $blocked -PathType Leaf)){Remove-Item -LiteralPath $blocked}
                if($moved){Move-Item -LiteralPath $held -Destination $blocked}
                [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            }
        }
        SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($Book.Name))
        foreach($action in @('PROCESS_SELECT','REQUIREMENT_SELECT')){
            Check ('AssignmentPolicy.NavigationDefault.'+$action) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.AssignmentDefaultForTest' @('PRODUCTION_ASSIGNMENT_'+$action)) -ceq 'False|True|True')
        }
        $before=@(Files);[void](Probe 'AssignmentStage' @('Normal'))
        Check 'AssignmentPolicy.ProgrammaticSetupNotUserActions' (@(Files).Count -eq $before.Count)
        foreach($mode in @('PartialFailure','EmptyArray')){
            [void](Probe 'AssignmentStage' @($mode));$state=[string](Probe 'AssignmentState');Arm;$before=@(Files);$sourceBefore=@(SourceRows|ForEach-Object EventID)
            $notice=[string](Probe 'AssignmentAct' @('SAVE'));$writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
            $label='AssignmentPolicy.Save.'+$mode
            Check ($label+'.ExistingLocalPreparationRemains') ([string](Probe 'AssignmentState') -cne $state)
            Check ($label+'.NoOwningWrite') ($writer.Count -eq 5 -and $writer[1] -ceq '0')
            Check ($label+'.SourceEventsPreserved') ((@(SourceRows|ForEach-Object EventID) -join '|') -ceq ($sourceBefore -join '|'))
            Check ($label+'.GuardsRestored') ([bool](Probe 'AssignmentGuardsRestored'))
            $expected=if($mode -ceq 'PartialFailure'){'FAILED'}else{'REJECTED'}
            Pair $before 'SAVE' $expected 'config-producer' $label
        }
        foreach($guard in @('Loading','Busy')){
            [void](Probe 'AssignmentStage' @('Normal'));$state=[string](Probe 'AssignmentState');$before=@(Files);Arm
            [void](Probe 'AssignmentGuard' @('SAVE',$guard));$writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
            Check ('AssignmentPolicy.Save.'+$guard+'.NoReadsOrMutation') ([string](Probe 'AssignmentState') -ceq $state -and [int](Probe 'ReadCalls') -eq 0)
            Check ('AssignmentPolicy.Save.'+$guard+'.NoWriteOrActivity') ($writer.Count -eq 5 -and $writer[1] -ceq '0' -and @(Files).Count -eq $before.Count)
            Check ('AssignmentPolicy.Save.'+$guard+'.GuardRestored') ([bool](Probe 'AssignmentGuardsRestored'))
        }
        [void](Probe 'AssignmentStage' @('Normal'));Arm;$before=@(Files);[void](Probe 'AssignmentArmNested')
        [void](Probe 'AssignmentAct' @('SAVE'));$writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
        Check 'AssignmentPolicy.Save.Nested.ActualHandlerEntered' ([bool](Probe 'AssignmentNestedReached'))
        Check 'AssignmentPolicy.Save.Nested.ExactlyOneWrite' ($writer.Count -eq 5 -and $writer[1] -ceq '1')
        Check 'AssignmentPolicy.Save.Nested.GuardsRestored' ([bool](Probe 'AssignmentGuardsRestored'))
        Pair $before 'SAVE' 'CONFIRMED' 'config-producer' 'AssignmentPolicy.Save.Nested'

        # The queue boundary signs out only after the real append returned success.
        [void](Probe 'AssignmentStage' @('Normal'));Arm 'SignOutAfterAppend';$before=@(Files)
        $notice=[string](Probe 'AssignmentAct' @('SAVE'));$writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
        Check 'AssignmentPolicy.Save.Yield.SignOutBoundaryReached' ([bool](Run 'invSys.Core.xlam' 'modRoleEventWriter.AssignmentSignoutReachedForTest'))
        Check 'AssignmentPolicy.Save.Yield.OneActualAppend' ($writer.Count -eq 5 -and $writer[1] -ceq '1' -and $writer[2] -ceq 'True')
        Check 'AssignmentPolicy.Save.Yield.NoProcessorContinuation' (([string](Probe 'LifecycleFaultEvidence')).Split('|')[0] -ceq '')
        Check 'AssignmentPolicy.Save.Yield.NoSubsequentDesignRefresh' ([int](Probe 'ReadCalls') -eq 1)
        Check 'AssignmentPolicy.Save.Yield.RefusalVisible' ($notice -match 'Reopen|reopen')
        $rows=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
        Check 'AssignmentPolicy.Save.Yield.NoOutcomeInReplacementContext' ($rows.Count -eq 1 -and $rows[0].ControlId -ceq 'PRODUCTION_ASSIGNMENT_SAVE' -and $rows[0].OutcomeCode -ceq 'REQUESTED' -and $rows[0].UserId -ceq 'config-producer')
        Check 'AssignmentPolicy.Save.Yield.GuardsRestored' ([bool](Probe 'AssignmentGuardsRestored'))

        Test-ProductionAssignmentReadYield $Fixture $Book

        SelectTarget $Fixture 'config-producer'
        $guardBook=$excel.Workbooks.Add();$guardBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Preserved Guard Draft'
        $guardPath=Join-Path $runRoot 'assignment-save-guards.xlsb';$guardBook.SaveAs($guardPath,50);$guardPin=Hash $guardPath
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($guardBook.Name));[void](Probe 'AssignmentStage' @('Normal'))
            $state=[string](Probe 'AssignmentState');Arm;$before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){
                Initialize-SettingsCapture
                [void](Probe 'AssignmentShowAndCapture' @($guardBook.Name))
                $capturedName=$guardBook.Name;$otherName=$Book.Name
                $boundary=[ordered]@{BeforeUTC=[DateTimeOffset]::UtcNow.ToString('o');VisibleBefore=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero);WorkbooksBefore=$excel.Workbooks.Count}
                $Book.Activate();$guardBook.Close($false);$guardBook=$null
                $names=@(foreach($openBook in $excel.Workbooks){[string]$openBook.Name})
                $boundary.AfterUTC=[DateTimeOffset]::UtcNow.ToString('o')
                $boundary.VisibleAfter=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
                $boundary.WorkbooksAfter=$excel.Workbooks.Count
                $boundary.CapturedStillOpen=($capturedName -cin $names);$boundary.OtherStillOpen=($otherName -cin $names)
                if(-not $boundary.VisibleBefore -or $boundary.CapturedStillOpen -or -not $boundary.OtherStillOpen){throw 'Closed Assignment operator fixture unavailable; not product RED.'}
                $boundary.HandlerInvoked=[bool]$boundary.VisibleAfter;$protected=-not $boundary.VisibleAfter
                if($boundary.VisibleAfter){
                    $notice=[string](Probe 'AssignmentClosedSave');$entry=([string](Probe 'AssignmentClosedStatus')).Split('|')
                    $boundary.HandlerEntered=($entry[0] -ceq 'True');$boundary.AdapterError=[long]$entry[1]
                    if($entry[0] -cne 'True'){throw 'Visible Assignment fixture did not enter its handler; not product RED.'}
                    $protected=[long]$entry[1] -eq 0 -and $notice -match 'Reopen|reopen' -and [string](Probe 'AssignmentState') -ceq $state -and [int](Probe 'ReadCalls') -eq 0
                }
                $label='AssignmentPolicy.Save.ClosedBoundary'
                Check ($label+'.SurfaceLifetimeEstablished') ($boundary.WorkbooksBefore -eq $boundary.WorkbooksAfter+1)
                Check ($label+'.UserActionProtected') $protected
                Check ($label+'.BindingGuardRejectsClosedBook') (-not [bool](Probe 'AssignmentClosedBindingCurrent'))
                Check ($label+'.SavedBytesPreserved') ((Hash $guardPath) -ceq $guardPin)
                $writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
                Check ($label+'.NoOwningWrite') ($writer.Count -eq 5 -and $writer[1] -ceq '0')
                Check ($label+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
                $boundary|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'assignment-closed-boundary.json')
                [void](Probe 'AssignmentSafeClose')
                continue
            }
            $notice=[string](Probe 'AssignmentAct' @('SAVE'));$writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
            $label='AssignmentPolicy.Save.Guard.'+$guard
            Check ($label+'.NoReadsOrMutation') ([string](Probe 'AssignmentState') -ceq $state -and [int](Probe 'ReadCalls') -eq 0)
            Check ($label+'.NoOwningWrite') ($writer.Count -eq 5 -and $writer[1] -ceq '0')
            Check ($label+'.RefusalVisible') ($notice -match 'Reopen|reopen')
            Check ($label+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        Check 'AssignmentPolicy.Save.Guard.SavedBytesPreserved' ((Hash $guardPath) -ceq $guardPin)
        Check 'AssignmentPolicy.UnknownValuesPreserved' ($Sheet.Cells.Item(2,1).Value2 -ceq $Canary -and $Sheet.Cells.Item(2,2).Formula -ceq '=1+2')
    }finally{
        if($null -ne $guardBook){$guardBook.Close($false)}
        [void](Probe 'LifecycleFaultMode' @(''))
        [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @('','',''))
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes);SelectTarget $Fixture 'config-producer'
    }
}
