# D18 actual worksheet submissions; independent inbox/Domain proof, no raw values in reports.
function Test-ReceivingWorksheetScenarios($Fixture) {
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    foreach($case in @('Applied','Pending','UnknownSubmission','StoreFailure','OlderPolicy')) {
        $operator=$null;$other=$null;$held='';$blocked=''
        $label='WorksheetScenario.'+$case
        $caseStep='create public fixture'
        try {
            SelectTarget $Fixture 'config-reader'
            $other=$excel.Workbooks.Add()
            $other.Worksheets.Item(1).Cells.Item(1,1).Value2='worksheet unrelated sentinel'
            $other.SaveAs((Join-Path $runRoot ('worksheet-other-'+$case+'.xlsm')),52)
            $otherHash=Get-ReceivingFixtureHash $other.FullName
            [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot ('worksheet-operators-'+$case))))
            $other.Activate()
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
            $state=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState')
            if(-not $state.EndsWith('|True')){throw 'Worksheet fixture public launcher did not establish a current form.'}
            $operator=$excel.Workbooks.Item($state.Split('|')[0])
            Check ($label+'.PublicLauncherCaptured') ([bool](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherBoundTo' @($operator.Name)))
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @('Button'))
            if(-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name,$false,$other.Name))){throw 'Worksheet fixture did not stage two entries through the real form.'}
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
            $sheet=$operator.Worksheets.Item('ReceivedTally')
            $button=$sheet.Shapes.Item('btnConfirmWrites')
            $staging=Table $operator 'ReceivedTally'
            $expected=@(Get-ReceivingFixtureRows $staging)
            $ids=@($expected | ForEach-Object EventId)
            Check ($label+'.ExactCreatedIdentities') ($expected.Count -eq 2 -and @($ids | Select-Object -Unique).Count -eq 2 -and @($expected | ForEach-Object System_Key | Select-Object -Unique).Count -eq 2)
            $caseStep='add unknown column'
            $extra=$staging.ListColumns.Add();$extra.Name='Worksheet Extra'
            $caseStep='set custom values'
            $extra.DataBodyRange.Cells.Item(1,1).Value2='preserve worksheet first'
            $extra.DataBodyRange.Cells.Item(2,1).Value2='preserve worksheet second'
            $sheet.Visible=-1
            $caseStep='prepare worksheet visibility'
            foreach($ws in $operator.Worksheets){if($ws.Name -cne 'ReceivedTally'){$ws.Visible=2}}
            $operator.Save()
            $caseStep='configure pending and acknowledgement'
            [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetPending' @($case -eq 'Pending'))
            [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetLoseAcknowledgement' @($case -eq 'UnknownSubmission'))
            if($case -eq 'OlderPolicy'){Set-ReceivingNavigationPolicy $Fixture $false 6}
            $caseStep='capture policy and prior activity'
            $configBefore=Get-ReceivingFixtureHash $Fixture.Config
            $before=@(Get-Slice4beActivityFiles $Fixture)
            if($case -eq 'StoreFailure') {
                $blocked=Join-Path $Fixture.Root ('Training\Activity\'+$Fixture.Warehouse)
                $held=$blocked+'-worksheet-held'
                $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
                foreach($path in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Worksheet tracking fault escaped fixture root.'}}
                Move-Item -LiteralPath $blocked -Destination $held
                [IO.File]::WriteAllText($blocked,'Blocked disposable activity path')
            }
            [void](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Reset')
            $caseStep='native input'
            Write-Output ('Worksheet fixture before input: windows='+$operator.Windows.Count+'; sheets='+$operator.Worksheets.Count+'; addin='+$operator.IsAddin+'; open='+(@($excel.Workbooks | Where-Object { $_.Name -ceq $operator.Name }).Count -eq 1))
            $native=@(Invoke-ReceivingNativeSurface $operator $sheet $button $case)
            $native | Where-Object {$_ -is [string]} | Write-Output
            if(-not ($native[-1].Entered -and $native[-1].ShapeCaller)){throw 'Native owner and caller not established for worksheet scenario.'}
            $succeeded=[bool](Run 'invSys.Operations.xlam' 'modTS_Received.LastConfirmWritesSucceeded')
            $caseStep='independent owner evidence'
            $status=[string](Run 'invSys.Operations.xlam' 'modTS_Received.LastConfirmWritesStatus')
            if($held -ne ''){Remove-Item -LiteralPath $blocked;Move-Item -LiteralPath $held -Destination $blocked;$held='';$blocked=''}
            $inboxPath=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ResolveInboxWorkbookPath' @('RECEIVE',$Fixture.Warehouse,'S1',''))
            if(-not [IO.Path]::GetFullPath($inboxPath).StartsWith($Fixture.Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Worksheet inbox escaped fixture root.'}
            $inbox=Open-ReceivingEvidenceBook $inboxPath
            $authority=Open-ReceivingEvidenceBook (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb'))
            $queued=@(Get-ReceivingFixtureRows (Table $inbox 'tblInboxReceive') | Where-Object {$_.EventID -cin $ids})
            $applied=@(Get-ReceivingFixtureRows (Table $authority 'tblAppliedEvents') | Where-Object {$_.EventID -cin $ids})
            $logged=@(Get-ReceivingFixtureRows (Table $authority 'tblInventoryLog') | Where-Object {$_.EventID -cin $ids})
            $pending=$case -in @('Pending','UnknownSubmission')
            $business=$queued.Count -eq 2
            if($pending){$business=$business -and -not $succeeded -and $applied.Count -eq 0 -and $logged.Count -eq 0 -and $staging.ListRows.Count -eq 2}
            else {
                $business=$business -and $succeeded -and $applied.Count -eq 2 -and $logged.Count -eq 2 -and $staging.ListRows.Count -eq 0
                foreach($item in $expected){$business=$business -and @($logged | Where-Object {$_.EventID -ceq $item.EventId -and $_.System_Key -ceq $item.System_Key -and [double]$_.QtyDelta -eq [double]$item.QUANTITY}).Count -eq 1}
            }
            Check ($label+'.IndependentSubmissionAndDomainEvidence') $business
            if(-not $business){throw 'Worksheet owning evidence did not establish the intended scenario.'}
            Check ($label+'.UnknownHeaderPreserved') ($staging.ListColumns.Item('Worksheet Extra').Name -ceq 'Worksheet Extra')
            if($pending) {
                $current=@(Get-ReceivingFixtureRows $staging)
                $same=$current.Count -eq $expected.Count
                for($i=0;$i -lt $current.Count;$i++){$same=$same -and $current[$i].System_Key -ceq $expected[$i].System_Key -and $current[$i].EventId -ceq $expected[$i].EventId}
                Check ($label+'.StagedIdentityAndUnknownValuesPreserved') ($same -and $extra.DataBodyRange.Cells.Item(1,1).Value2 -ceq 'preserve worksheet first' -and $extra.DataBodyRange.Cells.Item(2,1).Value2 -ceq 'preserve worksheet second')
            }
            if($case -in @('StoreFailure','OlderPolicy')) {
                Check ($label+'.NoInventedActivity') (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
                Check ($label+'.TrackingFailureVisible') ($status.Contains('Tracking unavailable') -and ([string](Run 'invSys.Operations.xlam' 'TestReceivingLifecycleNotice.Message')).Contains('Tracking unavailable'))
            } else {
                $outcome=@{Applied='CONFIRMED';Pending='PENDING';UnknownSubmission='FAILED'}[$case]
                $severity=@{Applied='Info';Pending='Warning';UnknownSubmission='Error'}[$case]
                $referenceState=if($case -eq 'UnknownSubmission'){'Unknown'}else{'Submitted'}
                Test-ReceivingControlRecords $Fixture $before 'RECEIVING_WORKSHEET_CONFIRM' 'RECEIVING_WORKFLOW' 'RECEIVE_WORKSHEET_CONFIRM_' $outcome 1 'Unknown' $expected $label $severity 7 'config-reader' $referenceState
                Check ($label+'.OnlyNativeControlAttributed') (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count+2)
            }
            Check ($label+'.ConfigBytesPreserved') ($configBefore -ceq (Get-ReceivingFixtureHash $Fixture.Config))
            Check ($label+'.UnrelatedWorkbookPreserved') ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'worksheet unrelated sentinel' -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
        } catch {
            throw ('Worksheet scenario failed at '+$caseStep+'; test line='+$_.InvocationInfo.ScriptLineNumber+'; HRESULT='+$_.Exception.HResult)
        } finally {
            if($held -ne ''){if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked};Move-Item -LiteralPath $held -Destination $blocked}
            foreach($book in $receivingEvidenceOpened){$book.Close($false)};$receivingEvidenceOpened.Clear()
            if($null -ne $operator){$operator.Close($false)}
            if($null -ne $other){$other.Close($false)}
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetPending' @($false))
            [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetLoseAcknowledgement' @($false))
        }
    }
}
