# Sign out only after an actual owning read returns; retain its payload/parser.
function Install-ProductionAssignmentYieldProbe {
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    foreach($entry in @(
        @{Name='DesignReadPayloadForTest';Anchor='DesignReadPayloadForTest = modOperationsPrimitiveBridge.GetProcessVersion(identity, version)'},
        @{Name='DesignReadListForTest';Anchor='DesignReadListForTest = modOperationsPrimitiveBridge.ListProcesses(status)'}
    )){
        $start=$form.ProcStartLine($entry.Name,0);$count=$form.ProcCountLines($entry.Name,0)
        $matches=@(for($i=$start;$i -lt $start+$count;$i++){if($form.Lines($i,1).Trim() -ieq $entry.Anchor){$i}})
        if($matches.Count -ne 1){throw 'Assignment read-yield seam changed; not product RED.'}
        $form.InsertLines($matches[0]+1,@'
    If mReadModeForTest = "AssignmentSignOutAfterRead" Then
        TestProductionDesigner.AssignmentReadSignoutReachedForTest = True
        modAuth.SignOut
    End If
'@)
    }
    $adapter=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Public AssignmentReadSignoutReachedForTest As Boolean')
    $adapter.AddFromString(@'
Public Sub AssignmentReadSignoutReset()
    AssignmentReadSignoutReachedForTest = False
End Sub
Public Function AssignmentReadSignoutReached() As Boolean
    AssignmentReadSignoutReached = AssignmentReadSignoutReachedForTest
End Function
'@)
}

function Test-ProductionAssignmentReadYield($Fixture,$Book){
    # The surrounding safety scope supplies the real adapter, file and owner reads.
    foreach($action in @('REFRESH','PROCESS','PROCESS_SELECT','SAVE')){
        SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($Book.Name))
        [void](Probe 'AssignmentStage' @('AssignmentSignOutAfterRead'))
        [void](Probe 'AssignmentReadSignoutReset')
        $queue=Join-Path $runRoot ('assignment-read-yield-'+[guid]::NewGuid().ToString('N'))
        [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @($queue,'Observe',$Fixture.Warehouse))
        [void](Probe 'LifecycleFaultMode' @('Observe'))
        $state=[string](Probe 'AssignmentState');$before=@(Files)
        $notice=[string](Probe 'AssignmentAct' @($action))
        $writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
        $label='AssignmentYield.Read.'+$action
        $reached=[bool](Probe 'AssignmentReadSignoutReached')
        Check ($label+'.RealReadReturnedBeforeSignOut') $reached
        if(-not $reached){throw 'Assignment sign-out read boundary not reached; not product RED.'}
        Check ($label+'.NoSubsequentReads') ([int](Probe 'ReadCalls') -eq 1)
        Check ($label+'.LocalStatePreserved') ([string](Probe 'AssignmentState') -ceq $state)
        Check ($label+'.RefusalVisible') ($notice -match 'Reopen|reopen')
        Check ($label+'.NoOwningWrite') ($writer.Count -eq 5 -and $writer[1] -ceq '0')
        Check ($label+'.GuardsRestored') ([bool](Probe 'AssignmentGuardsRestored'))
        $rows=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
        Check ($label+'.OriginalAttemptWithoutMisattributedOutcome') ($rows.Count -eq 1 -and $rows[0].ControlId -ceq ('PRODUCTION_ASSIGNMENT_'+$action) -and $rows[0].OutcomeCode -ceq 'REQUESTED' -and $rows[0].UserId -ceq 'config-producer')
    }
    SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($Book.Name))
    [void](Probe 'AssignmentStage' @('Normal'))
}
