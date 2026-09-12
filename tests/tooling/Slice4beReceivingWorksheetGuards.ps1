# D18 supplemental native guard proofs. All VBA seams are unsaved, installed
# before fixture/form creation, and expose counts/booleans rather than payloads.
function Install-ReceivingWorksheetGuardSeams($Gate) {
    $Gate.AddFromString(@'
Public WorksheetMode As String
Public WorksheetBook As String
Public WorksheetInterrupts As Long
Public WorksheetTargetRoot As String
Private WorksheetTargetWarehouse As String
Private WorksheetTargetSecret As String
Private WorksheetTargetReady As Boolean
Public Sub ArmWorksheet(ByVal mode As String, ByVal workbookName As String)
    WorksheetMode = mode: WorksheetBook = workbookName: WorksheetInterrupts = 0
    WorksheetTargetReady = False
    If mode = "" Then WorksheetTargetSecret = ""
End Sub
Public Sub ConfigureWorksheetTarget(ByVal root As String, ByVal warehouseId As String, ByVal secretValue As String)
    WorksheetTargetRoot = root: WorksheetTargetWarehouse = warehouseId: WorksheetTargetSecret = secretValue
End Sub
Public Function WorksheetTargetIsReady() As Boolean
    Dim target As WarehouseTarget
    If Not WorksheetTargetReady Or modActivity.CaptureContext() = "" Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    WorksheetTargetIsReady = (StrComp(target.WarehouseId, WorksheetTargetWarehouse, vbBinaryCompare) = 0) And _
        (StrComp(target.RuntimeRoot, WorksheetTargetRoot, vbTextCompare) = 0) And _
        (modAuth.GetCurrentUserId() = "config-reader")
End Function
Public Function WorksheetInterruptCount() As Long
    WorksheetInterruptCount = WorksheetInterrupts
End Function
Public Sub AfterWorksheetBegin(ByVal controlId As String)
    Dim mode As String, result As String
    If controlId <> "RECEIVING_WORKSHEET_CONFIRM" Or WorksheetMode = "" Then Exit Sub
    mode = WorksheetMode: WorksheetMode = "": WorksheetInterrupts = WorksheetInterrupts + 1
    Select Case mode
        Case "Switch": Application.Workbooks(WorksheetBook).Activate
        Case "Close": Application.Workbooks(WorksheetBook).Close SaveChanges:=False
        Case "SignOut": modAuth.SignOut
        Case "Target"
            modRuntimeWorkbooks.SetCoreDataRootOverride WorksheetTargetRoot
            result = modNasConnection.SelectWarehouseTargetForAutomation(WorksheetTargetRoot, WorksheetTargetRoot, "S1", False)
            If Left$(result, 3) <> "OK|" Then Err.Raise 5, , "Fixture target selection failed."
            modNasConnection.SetCurrentTargetPathsForTest "\\fixture-host\config-command", WorksheetTargetRoot
            result = modAuth.SignInCurrentTargetForAutomation("config-reader", WorksheetTargetSecret, "")
            WorksheetTargetSecret = ""
            If Left$(result, 3) <> "OK|" Then Err.Raise 5, , "Fixture target sign-in failed."
            WorksheetTargetReady = True
    End Select
End Sub
'@)
    $activity=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActivity').CodeModule
    $start=$activity.ProcStartLine('BeginAction',0);$count=$activity.ProcCountLines('BeginAction',0)
    $source=$activity.Lines($start,$count)
    $pattern='(?im)^[ \t]*mBusy = False[ \t]*\r?$'
    if([regex]::Matches($source,$pattern).Count -ne 1){throw 'Unique optional-tracking return seam unavailable.'}
    $changed=[regex]::Replace($source,$pattern,"    mBusy = False`r`n    TestReceivingActivityGate.AfterWorksheetBegin controlId")
    $activity.DeleteLines($start,$count);$activity.InsertLines($start,$changed)
    $project=$packages['invSys.Operations.xlam'].VBProject
    $observer=$project.VBComponents.Add(1);$observer.Name='TestWorksheetOwner'
    $observer.CodeModule.AddFromString(@'
Option Explicit
Public Calls As Long
Public ExpectedUsed As Boolean
Public ExpectedWorkbook As Workbook
Public Sub Reset(ByVal workbookName As String)
    Calls = 0: ExpectedUsed = False
    Set ExpectedWorkbook = Application.Workbooks(workbookName)
End Sub
Public Function State() As String
    State = CStr(Calls) & "|" & CStr(ExpectedUsed)
End Function
'@)
    $owner=$project.VBComponents.Item('modReceivingPostingService').CodeModule
    $start=$owner.ProcStartLine('ExecuteConfirmWrites',0);$count=$owner.ProcCountLines('ExecuteConfirmWrites',0)
    $source=$owner.Lines($start,$count)
    $pattern='(?im)^    outcome = "REJECTED": sourceEventIds = "": submissionState = "Unknown"\r?$'
    if([regex]::Matches($source,$pattern).Count -ne 1){throw 'Unique existing posting-owner observation seam unavailable.'}
    $changed=[regex]::Replace($source,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m)
        "    TestWorksheetOwner.Calls = TestWorksheetOwner.Calls + 1`r`n    TestWorksheetOwner.ExpectedUsed = (operatorWb Is TestWorksheetOwner.ExpectedWorkbook)`r`n"+$m.Value
    })
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$changed)
}

function Test-ReceivingWorksheetGuards($Fixture) {
    $targetFixture=NewFixture 'worksheet-guard-target'
    SelectTarget $targetFixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($targetFixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Second generated guard warehouse was not seeded.'}
    foreach($case in @('Denied','SignedOut','SwitchWorkbook','SignOutDuringTracking','CloseDuringTracking','TargetDuringTracking')) {
        $operator=$null;$other=$null;$operatorName='';$stepName='provision and stage'
        $label='WorksheetGuard.'+$case
        try {
            SelectTarget $Fixture 'config-reader'
            $other=$excel.Workbooks.Add();$other.Worksheets.Item(1).Cells.Item(1,1).Value2='guard unrelated sentinel'
            $other.SaveAs((Join-Path $runRoot ('worksheet-guard-other-'+$case+'.xlsm')),52)
            $otherHash=Get-ReceivingFixtureHash $other.FullName
            [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot ('worksheet-guard-operators-'+$case))))
            $other.Activate();[void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
            $state=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState')
            if(-not $state.EndsWith('|True')){throw 'Guard fixture public launcher did not establish a current form.'}
            $operatorName=$state.Split('|')[0];$operator=$excel.Workbooks.Item($operatorName)
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @('Button'))
            if(-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operatorName,$false,$other.Name))){throw 'Guard fixture did not stage two entries through the real form.'}
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
            $sheet=$operator.Worksheets.Item('ReceivedTally');$button=$sheet.Shapes.Item('btnConfirmWrites')
            $staging=Table $operator 'ReceivedTally';$extra=$staging.ListColumns.Add();$extra.Name='Guard Extra'
            $extra.DataBodyRange.Value2='preserve guard extension'
            $expected=@(Get-ReceivingFixtureRows $staging);$beforeRows=$expected | ConvertTo-Json -Depth 5 -Compress
            $sheet.Visible=-1;foreach($ws in $operator.Worksheets){if($ws.Name -cne 'ReceivedTally'){$ws.Visible=2}}
            $operator.Save();$operatorPath=$operator.FullName;$operatorHash=Get-ReceivingFixtureHash $operatorPath
            $actor='config-reader'
            if($case -eq 'Denied'){$actor='config-producer';SelectTarget $Fixture $actor}
            if($case -eq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            $mode=@{SwitchWorkbook='Switch';SignOutDuringTracking='SignOut';CloseDuringTracking='Close';TargetDuringTracking='Target'}[$case]
            if($null -eq $mode){$mode=''}
            $switchBook=if($case -eq 'SwitchWorkbook'){$other.Name}else{$operatorName}
            [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.ArmWorksheet' @($mode,$switchBook))
            if($case -eq 'TargetDuringTracking'){
                [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.ConfigureWorksheetTarget' @($targetFixture.Root,$targetFixture.Warehouse,$targetFixture.Secret))
                $targetAuthority=Get-ReceivingAuthorityHashes $targetFixture
                $targetActivity=@(Get-Slice4beActivityFiles $targetFixture)
            }
            [void](Run 'invSys.Operations.xlam' 'TestWorksheetOwner.Reset' @($operatorName))
            [void](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Reset')
            $before=@(Get-Slice4beActivityFiles $Fixture);$authority=Get-ReceivingAuthorityHashes $Fixture
            $stepName='native input'
            $native=@(Invoke-ReceivingNativeSurface $operator $sheet $button ('Guard'+$case))
            $native | Where-Object {$_ -is [string]} | Write-Output
            if(-not ($native[-1].Entered -and $native[-1].ShapeCaller)){throw 'Native guard owner/caller was not established.'}
            $ownerState=([string](Run 'invSys.Operations.xlam' 'TestWorksheetOwner.State')).Split('|')
            $calls=[int]$ownerState[0];$succeeded=[bool](Run 'invSys.Operations.xlam' 'modTS_Received.LastConfirmWritesSucceeded')
            $status=[string](Run 'invSys.Operations.xlam' 'modTS_Received.LastConfirmWritesStatus')
            $interrupts=[int](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.WorksheetInterruptCount')
            Check ($label+'.InterruptionDelivered') ($interrupts -eq $(if($mode -eq ''){0}else{1}))
            Check ($label+'.NoOwnerRetryOrWorkbookRedirection') ($calls -le 1 -and ($calls -eq 0 -or $ownerState[1] -ceq 'True'))
            $stepName='owning evidence'
            if($case -eq 'SwitchWorkbook') {
                $data=Open-ReceivingEvidenceBook (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb'))
                $rows=@(Get-ReceivingFixtureRows (Table $data 'tblInventoryLog'))
                $applied=$succeeded -and $calls -eq 1 -and $staging.ListRows.Count -eq 0
                foreach($item in $expected){$applied=$applied -and @($rows | Where-Object {$_.EventID -ceq $item.EventId -and $_.System_Key -ceq $item.System_Key -and [double]$_.QtyDelta -eq [double]$item.QUANTITY}).Count -eq 1}
                Check ($label+'.EveryCapturedEntityApplied') $applied
                Test-ReceivingControlRecords $Fixture $before 'RECEIVING_WORKSHEET_CONFIRM' 'RECEIVING_WORKFLOW' 'RECEIVE_WORKSHEET_CONFIRM_' 'CONFIRMED' 1 'Unknown' $expected $label 'Info' 7
            } else {
                Check ($label+'.NoBusinessMutation') (-not $succeeded -and (Test-ReceivingAuthorityHashes $Fixture $authority))
                Check ($label+'.SavedStagingPreserved') ($operatorHash -ceq (Get-ReceivingFixtureHash $operatorPath))
                if($case -ne 'CloseDuringTracking') {Check ($label+'.LiveStagingAndUnknownValuesPreserved') (($beforeRows -ceq (@(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress)))}
                $paths=@(Get-Slice4beActivityFiles $Fixture | Where-Object {$_ -notin $before})
                if($case -eq 'SignedOut') {Check ($label+'.PreSignInNotAttributed') ($paths.Count -eq 0)}
                elseif($case -in @('SignOutDuringTracking','TargetDuringTracking')) {
                    Check ($label+'.ChangedContextStopsBeforeOwner') ($calls -eq 0 -and $status.Contains('Session or warehouse changed'))
                    if($case -eq 'TargetDuringTracking') {
                        $ready=[bool](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.WorksheetTargetIsReady')
                        Check ($label+'.DifferentValidSignedInTargetReached') ($ready -and $targetFixture.Warehouse -cne $Fixture.Warehouse)
                        if(-not $ready){throw 'Target interruption did not establish a valid authenticated context.'}
                        Check ($label+'.NewTargetAuthorityAndActivityPreserved') ((Test-ReceivingAuthorityHashes $targetFixture $targetAuthority) -and @((Get-Slice4beActivityFiles $targetFixture) | Where-Object {$_ -notin $targetActivity}).Count -eq 0)
                    }
                    $records=@($paths | ForEach-Object {Get-Content -LiteralPath $_ -Raw | ConvertFrom-Json})
                    $incomplete=$records.Count -eq 1
                    if($incomplete){$incomplete=$records[0].ControlId -ceq 'RECEIVING_WORKSHEET_CONFIRM' -and $records[0].OutcomeCode -ceq 'REQUESTED' -and $records[0].UserId -ceq $actor -and $records[0].WarehouseId -ceq $Fixture.Warehouse -and @($records[0].SourceEventRefs).Count -eq 0}
                    Check ($label+'.OnlyOriginalAttemptNoInventedConclusion') $incomplete
                } else {
                    $outcome=if($case -eq 'Denied'){'DENIED'}else{'REJECTED'}
                    $effect=if($case -eq 'Denied'){'Unchanged'}else{'Unknown'}
                    $severity=if($case -eq 'Denied'){'Blocked'}else{'Warning'}
                    Test-ReceivingControlRecords $Fixture $before 'RECEIVING_WORKSHEET_CONFIRM' 'RECEIVING_WORKFLOW' 'RECEIVE_WORKSHEET_CONFIRM_' $outcome 1 $effect @() $label $severity 7 $actor
                }
            }
            Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'guard unrelated sentinel' -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
        } catch {throw ('Worksheet guard failed at '+$stepName+'; test line='+$_.InvocationInfo.ScriptLineNumber+'; HRESULT='+$_.Exception.HResult)}
        finally {
            [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.ArmWorksheet' @('',''))
            foreach($book in $receivingEvidenceOpened){$book.Close($false)};$receivingEvidenceOpened.Clear()
            if($operatorName -ne ''){foreach($book in @($excel.Workbooks)){if($book.Name -ceq $operatorName){$book.Close($false)}}}
            if($null -ne $other){$other.Close($false)}
        }
    }
    SelectTarget $Fixture 'config-reader'
}
