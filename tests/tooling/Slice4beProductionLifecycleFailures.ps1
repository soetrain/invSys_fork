# Unsaved, one-boundary fault seams. Actual handlers, queue writes and processing stay real.
function Install-ProductionLifecycleFailureProbe {
    function InsertFault($Module,[string]$Procedure,[string]$Anchor,[string]$Text,[bool]$After=$false) {
        $start=$Module.ProcStartLine($Procedure,0);$count=$Module.ProcCountLines($Procedure,0)
        $found=@(for($line=$start;$line -lt $start+$count;$line++){if($Module.Lines($line,1).Trim() -ceq $Anchor){$line}})
        if($found.Count -ne 1){throw ('Lifecycle fault anchor is not unique: '+$Procedure)}
        $Module.InsertLines(($found[0]+[int]$After),$Text)
    }
    $writer=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modRoleEventWriter').CodeModule
    $writer.InsertLines(($writer.CountOfDeclarationLines+1),@'
Private mLifecycleTestRoot As String, mLifecycleTestMode As String, mLifecycleTestWarehouse As String
Private mLifecycleTestId As String, mLifecycleTestPath As String, mLifecycleTestType As String
Private mLifecycleTestPrinted As Boolean, mLifecycleTestFault As Boolean, mLifecycleTestCalls As Long
'@)
    InsertFault $writer 'LocalStagingRootRole' 'rootPath = Trim$(Environ$("LOCALAPPDATA"))' @'
    If mLifecycleTestRoot <> "" Then LocalStagingRootRole = mLifecycleTestRoot: Exit Function
'@
    InsertFault $writer 'AppendInboxRowToLocalStagingRole' 'writeAttemptedOut = True' @'
    If mLifecycleTestMode <> "" Then
        If CStr(rowValues("WarehouseId")) <> mLifecycleTestWarehouse Then Err.Raise 5, , "Fault fixture warehouse mismatch."
        mLifecycleTestCalls = mLifecycleTestCalls + 1
        mLifecycleTestId = CStr(rowValues("EventID")): mLifecycleTestType = CStr(rowValues("EventType"))
        mLifecycleTestPath = stagingPath
        If mLifecycleTestMode = "BeforeAppend" Then
            mLifecycleTestFault = True
            Err.Raise vbObjectError + 7850, , "Injected lifecycle fixture write failure."
        End If
    End If
'@
    InsertFault $writer 'AppendInboxRowToLocalStagingRole' 'Print #fileNum, serializedRow' @'
    If mLifecycleTestMode <> "" Then
        mLifecycleTestPrinted = True
        If mLifecycleTestMode = "AfterAppend" Then
            mLifecycleTestFault = True
            Err.Raise vbObjectError + 7851, , "Injected lifecycle fixture acknowledgment failure."
        End If
    End If
'@ $true
    $writer.AddFromString(@'
Public Sub LifecycleWriterFixtureForTest(ByVal root As String, ByVal mode As String, ByVal warehouse As String)
    mLifecycleTestRoot = root: mLifecycleTestMode = mode: mLifecycleTestWarehouse = warehouse
    mLifecycleTestId = "": mLifecycleTestPath = "": mLifecycleTestType = ""
    mLifecycleTestPrinted = False: mLifecycleTestFault = False: mLifecycleTestCalls = 0
End Sub
Public Function LifecycleWriterEvidenceForTest() As String
    Dim rows As Collection, row As Object, report As String, found As Boolean
    If mLifecycleTestPath <> "" Then
        Set rows = ReadStagedInboxRowsRole(mLifecycleTestPath, report)
        If Not rows Is Nothing Then
            For Each row In rows
                If CStr(row("EventID")) = mLifecycleTestId And CStr(row("WarehouseId")) = mLifecycleTestWarehouse And _
                   CStr(row("EventType")) = mLifecycleTestType Then found = True
            Next row
        End If
    End If
    LifecycleWriterEvidenceForTest = mLifecycleTestId & "|" & CStr(mLifecycleTestCalls) & "|" & _
        CStr(mLifecycleTestPrinted) & "|" & CStr(mLifecycleTestFault) & "|" & CStr(found)
End Function
'@)
    $owner=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modProductionReusableDesigns').CodeModule
    $owner.InsertLines(($owner.CountOfDeclarationLines+1),@'
Public LifecycleFaultModeForTest As String, LifecycleSubmissionIdForTest As String
Public LifecycleProcessingFaultForTest As Boolean, LifecycleRefreshFaultForTest As Boolean
'@)
    InsertFault $owner 'SubmitReusableDesignEvent' 'warehouseId = Trim$(modConfig.GetWarehouseId())' @'
    If LifecycleFaultModeForTest <> "" Then
        LifecycleSubmissionIdForTest = eventId
        If LifecycleFaultModeForTest = "Pending" Then
            LifecycleProcessingFaultForTest = True
            Err.Raise vbObjectError + 7852, , "Injected lifecycle fixture processing failure."
        End If
    End If
'@
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    $text=$form.Lines(1,$form.CountOfLines)
    $procedures=if($text.Contains('Private Function SubmitDesignerAction(')){@('SubmitDesignerAction')}else{@('SubmitProcessAction','SubmitRecipeAction')}
    foreach($procedure in $procedures){InsertFault $form $procedure 'RefreshReusableDesignLists' @'
    If modProductionReusableDesigns.LifecycleFaultModeForTest = "Refresh" Then
        modProductionReusableDesigns.LifecycleRefreshFaultForTest = True
        Err.Raise vbObjectError + 7853, , "Injected lifecycle fixture refresh failure."
    End If
'@}
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub LifecycleFaultMode(ByVal mode As String)
    modProductionReusableDesigns.LifecycleFaultModeForTest = mode
    modProductionReusableDesigns.LifecycleSubmissionIdForTest = ""
    modProductionReusableDesigns.LifecycleProcessingFaultForTest = False
    modProductionReusableDesigns.LifecycleRefreshFaultForTest = False
End Sub
Public Function LifecycleFaultEvidence() As String
    LifecycleFaultEvidence = modProductionReusableDesigns.LifecycleSubmissionIdForTest & "|" & _
        CStr(modProductionReusableDesigns.LifecycleProcessingFaultForTest) & "|" & _
        CStr(modProductionReusableDesigns.LifecycleRefreshFaultForTest)
End Function
'@)
    . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycleSafety.ps1')
    Install-ProductionLifecycleSafetyProbe
}

function Test-ProductionLifecycleFailures($Fixture,$Book,$Sheet,[string]$Canary) {
    # Every mode gets a fresh queue root inside this disposable fixture. Pending
    # records cannot be accidentally processed by a later case or ordinary workflow.
    $faultRoot=Join-Path $runRoot 'lifecycle-fault-queues'
    New-Item -ItemType Directory -Path $faultRoot -Force|Out-Null
    try {
        [void](Probe 'LifecycleFaultMode' @(''))
        [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @((Join-Path $faultRoot 'setup'),'',$Fixture.Warehouse))
        [void](Probe 'RefreshDesigners')
        $ready=[string](Probe 'LifecyclePrepare' @($Canary))
        if($ready -cnotmatch '^READY\|([^|]+)\|([^|]+)$'){throw 'Lifecycle failure Process fixture unavailable; not behavioral RED.'}
        $processId=$Matches[1];$processVersion=$Matches[2]
        $actions=@(@('Process','Save','DRAFT'),@('Process','Release','RELEASED'),@('Recipe','Save','DRAFT'),@('Recipe','Release','RELEASED'),@('Recipe','Obsolete','OBSOLETE'),@('Process','Obsolete','OBSOLETE'))
        foreach($action in $actions){
            $designer=$action[0];$verb=$action[1];$id='PRODUCTION_'+$designer.ToUpperInvariant()+'_'+$verb.ToUpperInvariant()
            if($designer -ceq 'Recipe' -and $verb -ceq 'Save'){
                if(-not [bool](Probe 'ReleasedRecipe' @($processId,$processVersion,$Canary))){throw 'Released Process prerequisite unavailable for fault cases; not RED.'}
            }
            foreach($mode in @('BeforeAppend','AfterAppend','Pending','Refresh')){
                $queueRoot=Join-Path $faultRoot ([guid]::NewGuid().ToString('N'))
                [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @($queueRoot,$mode,$Fixture.Warehouse))
                [void](Probe 'LifecycleFaultMode' @($mode))
                $before=@(Get-Slice4beActivityFiles $Fixture);$sourceBefore=@(SourceRows)
                $returned=[string](Probe 'LifecycleAct' @($designer,$verb))
                $writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
                $owner=([string](Probe 'LifecycleFaultEvidence')).Split('|')
                $case='ProductionLifecycle.Fault.'+$designer+'.'+$verb+'.'+$mode
                $reached=$writer.Count -eq 5 -and $writer[0] -cmatch '^[A-Za-z0-9_-]+$' -and $writer[1] -ceq '1' -and $owner.Count -eq 3
                if($reached){$reached=switch($mode){'BeforeAppend'{$writer[2] -ceq 'False' -and $writer[3] -ceq 'True'} 'AfterAppend'{$writer[2] -ceq 'True' -and $writer[3] -ceq 'True'} 'Pending'{$writer[2] -ceq 'True' -and $owner[0] -ceq $writer[0] -and $owner[1] -ceq 'True'} 'Refresh'{$writer[2] -ceq 'True' -and $owner[0] -ceq $writer[0] -and $owner[2] -ceq 'True'}}}
                if(-not $reached){throw ('Fault seam was not reached: '+$case+'; not behavioral RED.')}
                Check ($case+'.ActualOwnerFaultReached') ($reached -and $returned -ceq 'RETURNED')
                $source=@(SourceRows|Where-Object {$_.EventID -cnotin @($sourceBefore|ForEach-Object EventID)})
                $stored=switch($mode){'BeforeAppend'{$writer[4] -ceq 'False' -and $source.Count -eq 0} 'Refresh'{$source.Count -eq 1 -and $source[0].EventID -ceq $writer[0] -and $source[0].WarehouseId -ceq $Fixture.Warehouse -and $source[0].EventType -ceq ($designer.ToUpperInvariant()+'_'+$verb.ToUpperInvariant()) -and $source[0].AppliedAtUTC -ne ''} default{$writer[4] -ceq 'True' -and $source.Count -eq 0}}
                Check ($case+'.ActualWriteAndApplicationBoundary') $stored
                $expected=if($mode -in @('Pending','Refresh')){'PENDING'}else{'FAILED'}
                $state=if($mode -ceq 'BeforeAppend'){''}elseif($mode -ceq 'AfterAppend'){'Unknown'}else{'Submitted'}
                $raw=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {[IO.File]::ReadAllText($_)})
                $rows=@($raw|ForEach-Object {$_|ConvertFrom-Json});$attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcome=@($rows|Where-Object OutcomeCode -CEQ $expected)
                $pair=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
                Check ($case+'.ExplicitUncertainOutcome') $pair
                $refsOk=$pair;$correlated=$pair;$effect=$pair;$safe=$pair
                if($pair){
                    $refs=@($outcome[0].SourceEventRefs)
                    $refsOk=if($state -ceq ''){$refs.Count -eq 0}else{$refs.Count -eq 1 -and $refs[0].EventId -ceq $writer[0] -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SourceKind -ceq 'Designs' -and $refs[0].SubmissionState -ceq $state}
                    $refsOk=$refsOk -and @($attempt[0].SourceEventRefs).Count -eq 0
                    $correlated=$attempt[0].ActivityId -cne '' -and $attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $attempt[0].RecordId -cne $outcome[0].RecordId
                    $effect=$outcome[0].DataEffect -ceq 'Unknown' -and $outcome[0].ControlId -ceq $id -and $outcome[0].Severity -ceq $(if($expected -ceq 'PENDING'){'Notice'}else{'Error'}) -and $outcome[0].EventCode -ceq ($id+'_'+$expected)
                }
                foreach($text in $raw){foreach($value in @($Canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,$queueRoot,'Injected lifecycle','mBtn')){if($text.Contains($value)){$safe=$false}}}
                Check ($case+'.ExactOwnerReferenceOrNoUnusedId') $refsOk
                Check ($case+'.DistinctCorrelatedRecords') $correlated
                Check ($case+'.NeverClaimsAppliedOrRolledBack') $effect
                Check ($case+'.NoEnteredDataOrRawFailure') $safe
                Check ($case+'.UnknownOperatorColumnPreserved') ($Sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $Sheet.Cells.Item(2,1).Value2 -ceq $Canary)
                [void](Probe 'LifecycleFaultMode' @(''))
            }
            [void](Probe 'RefreshDesigners')
            if([string](Probe 'LifecycleStatus' @($designer)) -cne $action[2]){throw 'Successful owner effect unavailable after injected refresh failure; fixture cannot advance.'}
        }
    } finally {
        [void](Probe 'LifecycleFaultMode' @(''))
        [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @('','',''))
    }
}
