# Unsaved faults at existing public submission returns. Real authorization,
# writes, form handlers and Shipping owners remain active. Reports omit raw values.
function Install-ShippingSubmissionProbe($Module) {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modRoleEventWriter').CodeModule
    $source=$core.Lines(1,$core.CountOfLines)
    $source=$source.Replace('Option Explicit',@'
Option Explicit
Private ActivitySubmitMode As Long
Private ActivityServerCalls As Long, ActivityCurrentCalls As Long
Private ActivityServerAccepted As Boolean, ActivityCurrentAccepted As Boolean
Private ActivityServerId As String, ActivityCurrentInput As String, ActivityCurrentOutput As String
Private ActivitySubmissionRoot As String, ActivityServerPath As String, ActivityCurrentPath As String
Private ActivityPrewriteRefusals As Long, ActivityServerWriteEntries As Long, ActivityLocalWriteEntries As Long
'@)
    foreach($name in @('QueuePayloadEventServer','QueuePayloadEventCurrent')) {
        $pattern='(?ms)^Public Function '+$name+'\(.*?^End Function'
        $matches=[regex]::Matches($source,$pattern)
        if($matches.Count -ne 1){throw 'Shipping public submission observer anchor unavailable.'}
        $block=$matches[0].Value
        if($name -eq 'QueuePayloadEventServer') {
            $anchor='    QueuePayloadEventServer = QueueEventCore('
            $replacement=@'
    If ActivitySubmitMode <> 0 Then ActivityServerCalls = ActivityServerCalls + 1
    If ActivitySubmitMode = 1 Then
        errorMessage = "Fixture server submission unavailable."
        Exit Function
    End If
    QueuePayloadEventServer = QueueEventCore(
'@
            $block=$block.Replace($anchor,$replacement.TrimEnd("`r","`n"))
            $block=$block.Replace('End Function',@'
    If ActivitySubmitMode <> 0 Then
        ActivityServerAccepted = QueuePayloadEventServer
        ActivityServerId = eventIdOut
        If ActivitySubmitMode = 4 And QueuePayloadEventServer Then Err.Raise vbObjectError + 263, , "Fixture server acknowledgment unavailable."
        If ActivitySubmitMode = 2 Or ActivitySubmitMode = 3 Then
            QueuePayloadEventServer = False
            errorMessage = "Fixture server acknowledgment unavailable."
        End If
    End If
End Function
'@)
        } else {
            $anchor='    QueuePayloadEventCurrent = QueueEventCore('
            $replacement=@'
    If ActivitySubmitMode <> 0 Then
        ActivityCurrentCalls = ActivityCurrentCalls + 1
        ActivityCurrentInput = eventIdOut
        If ActivitySubmitMode = 3 Then
            ActivityCurrentOutput = eventIdOut
            errorMessage = "Fixture fallback submission unavailable."
            Exit Function
        End If
    End If
    QueuePayloadEventCurrent = QueueEventCore(
'@
            $block=$block.Replace($anchor,$replacement.TrimEnd("`r","`n"))
            $block=$block.Replace('End Function',@'
    If ActivitySubmitMode <> 0 Then
        ActivityCurrentAccepted = QueuePayloadEventCurrent
        ActivityCurrentOutput = eventIdOut
    End If
End Function
'@)
        }
        if($block -ceq $matches[0].Value){throw 'Shipping submission probe did not install.'}
        $source=$source.Replace($matches[0].Value,$block)
    }
    $pattern='(?ms)^Private Function QueueEventCore\(.*?^End Function'
    $matches=[regex]::Matches($source,$pattern)
    if($matches.Count -ne 1){throw 'Shipping persistence observer anchor unavailable.'}
    $block=$matches[0].Value.Replace('    SaveWorkbookRole wbInbox',"    SaveWorkbookRole wbInbox`r`n    If ActivitySubmitMode <> 0 Then ActivityServerPath = wbInbox.FullName")
    $anchor='    If localStageOnly Then'
    if([regex]::Matches($block,[regex]::Escape($anchor)).Count -ne 1){throw 'Allocated-before-write refusal anchor unavailable.'}
    $block=$block.Replace($anchor,@'
    If ActivitySubmitMode = 5 Then
        If eventIdOut = "" Then Err.Raise 5, , "Fixture expected Core to allocate before this boundary."
        ActivityPrewriteRefusals = ActivityPrewriteRefusals + 1
        errorMessage = "Fixture refused before either inventory submission write."
        GoTo CleanExit
    End If
    If localStageOnly Then
'@)
    $anchor='(?im)^        If Not AppendInboxRowToLocalStagingRole\(rowValues, stagingPath, errorMessage(?:, writeAttemptedOut)?\) Then GoTo CleanExit'
    if([regex]::Matches($block,$anchor).Count -ne 1){throw 'Local submission persistence observer anchor unavailable.'}
    $block=[regex]::Replace($block,$anchor,'$0'+"`r`n        If ActivitySubmitMode <> 0 Then ActivityCurrentPath = stagingPath")
    $source=$source.Replace($matches[0].Value,$block)
    foreach($entry in @(
        @('AppendInboxRowToLocalStagingRole','    On Error GoTo FailAppend','ActivityLocalWriteEntries'),
        @('WriteInboxRowValuesRole','    rowIndex = lo.ListRows.Add.Index','ActivityServerWriteEntries')
    )){
        $pattern='(?ims)^Private (?:Function|Sub) '+$entry[0]+'\(.*?^End (?:Function|Sub)'
        $matches=[regex]::Matches($source,$pattern)
        $anchorPattern='(?im)^'+[regex]::Escape($entry[1])
        if($matches.Count -ne 1 -or [regex]::Matches($matches[0].Value,$anchorPattern).Count -ne 1){throw ('Actual write entry observer anchor unavailable: '+$entry[0])}
        $block=[regex]::Replace($matches[0].Value,$anchorPattern,('    If ActivitySubmitMode <> 0 Then '+$entry[2]+' = '+$entry[2]+" + 1`r`n"+'$0'))
        $source=$source.Replace($matches[0].Value,$block)
    }
    $pattern='(?ms)^Private Function LocalStagingRootRole\(\) As String.*?^End Function'
    $matches=[regex]::Matches($source,$pattern)
    if($matches.Count -ne 1){throw 'Shipping staging fixture root anchor unavailable.'}
    $block=$matches[0].Value.Replace('    Dim rootPath As String',"    Dim rootPath As String`r`n    If ActivitySubmissionRoot <> """" Then LocalStagingRootRole = ActivitySubmissionRoot: Exit Function")
    $source=$source.Replace($matches[0].Value,$block)
    $core.DeleteLines(1,$core.CountOfLines);$core.AddFromString($source)
    $core.AddFromString(@'
Public Sub ActivityShippingSubmissionMode(ByVal mode As Long)
    ActivitySubmitMode = mode
    ActivityServerCalls = 0: ActivityCurrentCalls = 0
    ActivityServerAccepted = False: ActivityCurrentAccepted = False
    ActivityServerId = "": ActivityCurrentInput = "": ActivityCurrentOutput = ""
    ActivityServerPath = "": ActivityCurrentPath = ""
    ActivityPrewriteRefusals = 0: ActivityServerWriteEntries = 0: ActivityLocalWriteEntries = 0
End Sub
Public Function ActivityShippingWriteEntryState() As String
    ActivityShippingWriteEntryState = CStr(ActivityPrewriteRefusals) & "|" & CStr(ActivityServerWriteEntries) & "|" & CStr(ActivityLocalWriteEntries)
End Function
Public Sub ActivityShippingSubmissionFixtureRoot(ByVal root As String)
    ActivitySubmissionRoot = root
End Sub
Public Function ActivityShippingSubmissionStoragePath(ByVal server As Boolean) As String
    If server Then ActivityShippingSubmissionStoragePath = ActivityServerPath Else ActivityShippingSubmissionStoragePath = ActivityCurrentPath
End Function
Public Function ActivityShippingSubmissionState() As String
    ActivityShippingSubmissionState = CStr(ActivityServerCalls) & "|" & CStr(ActivityCurrentCalls) & "|" & _
        CStr(ActivityServerAccepted) & "|" & CStr(ActivityCurrentAccepted) & "|" & _
        CStr(ActivityServerId <> "" And ActivityServerId = ActivityCurrentInput And ActivityServerId = ActivityCurrentOutput) & "|" & _
        CStr(ActivityCurrentInput = "")
End Function
Public Function ActivityShippingSubmissionId() As String
    ActivityShippingSubmissionId = ActivityCurrentOutput
End Function
'@)
    $source=$Module.Lines(1,$Module.CountOfLines)
    $source=$source.Replace('Option Explicit',"Option Explicit`r`nPrivate ActivitySubmitOwnerCompleted As Boolean, ActivitySubmitOwnerResult As Boolean`r`nPrivate ActivitySubmitOwnerId As String")
    $pattern='(?ms)^Public Function ShipmentsFormCommitLine\(.*?^End Function'
    $matches=[regex]::Matches($source,$pattern)
    if($matches.Count -ne 1){throw 'Shipping commit outcome observer anchor unavailable.'}
    $block=$matches[0].Value.Replace("CleanExit:","CleanExit:`r`n    ActivitySubmitOwnerCompleted = True`r`n    ActivitySubmitOwnerResult = ShipmentsFormCommitLine`r`n    ActivitySubmitOwnerId = reserveEventId")
    $source=$source.Replace($matches[0].Value,$block)
    $Module.DeleteLines(1,$Module.CountOfLines);$Module.AddFromString($source)
    $Module.AddFromString(@'
Public Sub ActivityShippingSubmissionResetOwner()
    ActivitySubmitOwnerCompleted = False: ActivitySubmitOwnerResult = False: ActivitySubmitOwnerId = ""
End Sub
Public Function ActivityShippingSubmissionOwnerState(ByVal expectedId As String) As String
    ActivityShippingSubmissionOwnerState = CStr(ActivitySubmitOwnerCompleted) & "|" & _
        CStr(ActivitySubmitOwnerResult) & "|" & CStr(expectedId <> "" And ActivitySubmitOwnerId = expectedId)
End Function
'@)
}

function Test-Slice4beShippingSubmission($Module) {
    . (Join-Path $repo 'tools/plan022-dialog-observer.ps1')
    if(-not ('Plan022NativeDialogs' -as [type])){Invoke-Plan022NativeDialogObservation -ProcessId 0 -TimeoutSeconds 0}
    Install-ShippingSubmissionProbe $Module
    $template=Join-Path $repo 'deploy/current/templates/invSys.Data.Inventory.template.xlsb'
    if(-not (Test-Path -LiteralPath $template)){throw 'Accepted inventory template unavailable.'}
    $templateHash=Get-ShippingActivityHash $template
    if($PrepareShippingFixturesBeforeProbesForTest){$templateHash=$preparedShippingBoundaries['shipping-submission'].TemplateHash}
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Split-Path $template -Parent)))
    $operatorRoot=Join-Path $runRoot 'shipping-submission-operators'
    $bounded=[bool](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @($operatorRoot))
    Check 'Shipping.Submission.ExplicitOperatorRootSet' $bounded
    if(-not $bounded){throw 'Shipping fixture operator root unavailable.'}
    [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionFixtureRoot' @((Join-Path $runRoot 'shipping-submission-staging')))
    $fixture=NewFixture 'shipping-submission'
    $authPath=Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Auth.xlsb')
    $auth=$excel.Workbooks.Open($authPath,0,$false)
    $caps=Table $auth 'tblCapabilities';$row=$caps.ListRows.Add()
    foreach($pair in @{UserId='config-reader';Capability='SHIP_POST';WarehouseId=$fixture.Warehouse;StationId='S1';Status='Active'}.GetEnumerator()){$row.Range.Cells.Item(1,$caps.ListColumns.Item($pair.Key).Index).Value2=$pair.Value}
    $auth.Save();$auth.Close($false)
    SelectTarget $fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Shipping submission fixture seed failed.'}
    SelectTarget $fixture 'config-reader'
    $other=$excel.Workbooks.Add();$other.Worksheets.Item(1).Cells.Item(1,1).Value2='shipping submission sentinel'
    $other.SaveAs((Join-Path $runRoot 'shipping-submission-unrelated.xlsm'),52)
    $otherHash=Get-ShippingActivityHash $other.FullName
    $authHash=Get-ShippingActivityHash $authPath;$configHash=Get-ShippingActivityHash $fixture.Config
    $operator=$null;$dialogJob=$null;$stop=Join-Path $runRoot 'shipping-submission-dialog.stop'
    try {
        $other.Activate()
        $name=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingOpen')
        if(-not $name){throw 'Shipping submission launcher unavailable.'}
        $operator=$excel.Workbooks.Item($name)
        $bounded=[IO.Path]::GetFullPath($operator.FullName).StartsWith([IO.Path]::GetFullPath($operatorRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)
        Check 'Shipping.Submission.OperatorWithinFixtureRoot' $bounded
        if(-not $bounded){throw 'Shipping operator escaped its generated fixture root.'}
        [uint32]$owned=0
        [void][Plan022NativeDialogs]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$owned)
        if($owned -eq 0){throw 'Shipping submission fixture process unavailable.'}
        $dialogJob=Start-Job -ArgumentList (Join-Path $repo 'tools/plan022-dialog-observer.ps1'),$owned,$stop -ScriptBlock {
            param($path,$owned,$stop)
            . $path
            Invoke-Plan022NativeDialogObservation -ProcessId $owned -TimeoutSeconds 180 -StopPath $stop | Out-Null
        }
        $created=[bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingCreateBox')
        Check 'Shipping.Submission.BoxFixtureThroughHandlers' $created
        if(-not $created){throw 'Shipping submission box fixture unavailable.'}
        $ship=Table $operator 'ShipmentsTally';$hold=Table $operator 'NotShipped'
        foreach($table in @($ship,$hold)){$column=$table.ListColumns.Add();$column.Name='Shipping Extra'}
        $operator.Save()
        foreach($mode in @(1,2,4,3)) {
            $label='Shipping.Submission.'+@{1='ServerUnavailable';2='LostAcknowledgment';4='ExceptionalAcknowledgment';3='UncertainAcceptance'}[$mode]
            $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @('Add'))
            if(-not $key){throw 'Shipping submission Add selection unavailable.'}
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSubmissionResetOwner')
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionMode' @($mode))
            $activityBefore=@(Get-Slice4beActivityFiles $fixture)
            $other.Activate()
            [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @('Add'))
            $state=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionState')).Split('|')
            $id=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionId')
            $owner=([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSubmissionOwnerState' @($id))).Split('|')
            Check ($label+'.BothPublicSubmissionBoundariesEnteredOnce') ($state[0] -eq '1' -and $state[1] -eq '1')
            Check ($label+'.ActualServerAcceptanceObserved') ($state[2] -eq ([string]($mode -ne 1)))
            Check ($label+'.ActualFallbackResultObserved') ($state[3] -eq ([string]($mode -ne 3)))
            Check ($label+'.ExactIdSurvivesFallback') ($id -ne '' -and $(if($mode -eq 1){$state[5] -eq 'True'}else{$state[4] -eq 'True'}))
            Check ($label+'.ActualOwnerResultObserved') ($owner[0] -eq 'True' -and $owner[1] -eq ([string]($mode -ne 3)))
            $entries=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingWriteEntryState')).Split('|')
            $calibrated=$entries[0] -ceq '0' -and $entries[1] -ceq $(if($mode -eq 1){'0'}else{'1'}) -and $entries[2] -ceq $(if($mode -eq 3){'0'}else{'1'})
            Check ($label+'.ActualWriteEntriesObserved') $calibrated
            if(-not $calibrated){throw 'Actual write observer calibration failed; not a product RED.'}
            Check ($label+'.ExactIdRetainedByOwner') ($owner[2] -eq 'True')
            $expectedOutcome=if($mode -eq 3){'FAILED'}else{'PENDING'}
            $expectedState=if($mode -eq 3){'Unknown'}else{'Submitted'}
            Test-ShippingActivityPair $fixture $activityBefore 'Add' $label @($id) 'SHIPPING_ADD' $expectedOutcome $expectedState
            $serverPath=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionStoragePath' @($true))
            $localPath=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionStoragePath' @($false))
            $bounded=$true
            foreach($path in @($serverPath,$localPath)){if($path){$bounded=$bounded -and [IO.Path]::GetFullPath($path).StartsWith([IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)}}
            Check ($label+'.SubmissionFilesWithinFixtureRoot') $bounded
            if(-not $bounded){throw 'Submission storage escaped its generated fixture root.'}
            $sourceRows=@();$localRows=@();$readPreserved=$true
            if($serverPath) {
                $hash=Get-ShippingActivityHash $serverPath
                $book=$excel.Workbooks.Open($serverPath,0,$true)
                try{$sourceRows=@(Get-ShippingActivityRows (Table $book 'tblInboxShip') | Where-Object EventID -CEQ $id)}finally{$book.Close($false)}
                $readPreserved=$readPreserved -and $hash -ceq (Get-ShippingActivityHash $serverPath)
            }
            if($localPath) {
                $hash=Get-ShippingActivityHash $localPath
                $localRows=@(Get-Content -LiteralPath $localPath | ForEach-Object {$_ | ConvertFrom-Json} | Where-Object EventID -CEQ $id)
                $readPreserved=$readPreserved -and $hash -ceq (Get-ShippingActivityHash $localPath)
            }
            Check ($label+'.DurableServerIdentity') ($sourceRows.Count -eq $(if($mode -eq 1){0}else{1}) -and @($sourceRows | Where-Object {$_.WarehouseId -cne $fixture.Warehouse -or $_.EventType -cne 'SHIP_RESERVE'}).Count -eq 0)
            Check ($label+'.DurableLocalIdentity') ($localRows.Count -eq $(if($mode -eq 3){0}else{1}) -and @($localRows | Where-Object {$_.WarehouseId -cne $fixture.Warehouse -or $_.EventType -cne 'SHIP_RESERVE'}).Count -eq 0)
            Check ($label+'.ReadPreservesSubmissionBytes') $readPreserved
            $inventory=Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
            $beforeRead=Get-ShippingActivityHash $inventory
            $applied=@(Get-ShippingActivityLog $fixture | Where-Object EventID -CEQ $id)
            Check ($label+'.SubmissionNotYetDomainApplication') ($id -ne '' -and $applied.Count -eq 0)
            Check ($label+'.ReadPreservesAuthorityBytes') ($beforeRead -ceq (Get-ShippingActivityHash $inventory))
            Check ($label+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($name)))
            Check ($label+'.UnknownColumnsPreserved') ($null -ne $ship.ListColumns.Item('Shipping Extra') -and $null -ne $hold.ListColumns.Item('Shipping Extra'))
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionMode' @(0))
            if($mode -ne 3) {
                $remove=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @('Remove'))
                if(-not $remove){throw 'Shipping submission cleanup selection unavailable.'}
                [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @('Remove'))
                Check ($label+'.ActualRemoveClearsStaging') (@(Get-ShippingActivityRows $ship | Where-Object {[double]$_.QUANTITY -gt 0}).Count -eq 0)
            }
        }
        . (Join-Path $PSScriptRoot 'Slice4beShippingPrewriteRefusal.ps1')
        Test-Slice4beShippingPrewriteRefusal $fixture $operator $other $ship $hold
        Check 'Shipping.Submission.AuthBytesPreserved' ($authHash -ceq (Get-ShippingActivityHash $authPath))
        Check 'Shipping.Submission.ConfigBytesPreserved' ($configHash -ceq (Get-ShippingActivityHash $fixture.Config))
        Check 'Shipping.Submission.UnrelatedWorkbookPreserved' ($otherHash -ceq (Get-ShippingActivityHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping submission sentinel')
        Check 'Shipping.Submission.AcceptedTemplateBytesPreserved' ($templateHash -ceq (Get-ShippingActivityHash $template))
    } finally {
        [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionMode' @(0))
        [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.ActivityShippingSubmissionFixtureRoot' @(''))
        if($null -ne $dialogJob){
            [IO.File]::WriteAllText($stop,'stop');[void](Wait-Job $dialogJob -Timeout 5)
            if($dialogJob.State -eq 'Running'){Stop-Job $dialogJob}
            $failed=$dialogJob.State -eq 'Failed';Remove-Job $dialogJob
            if($failed){Check 'Shipping.Submission.DialogObserver' $false}
        }
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClose')
        if($null -ne $operator){$operator.Close($false)}
        $other.Close($false)
    }
}
