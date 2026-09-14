# Explicit evaluation diagnosis only. Fixed phase labels; never log runtime values.
function Install-Slice4beEvaluationNativeTrace {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $trace=$project.VBComponents.Add(1);$trace.Name='TestSeedTrace'
    $path=(Join-Path $reportRoot 'bootstrap-phases.log').Replace('"','""')
    $trace.CodeModule.AddFromString((@'
Public Sub Mark(ByVal phase As String)
    Dim handle As Integer
    handle = FreeFile
    Open "__TRACE_FILE__" For Append As #handle
    Print #handle, phase
    Close #handle
End Sub
'@).Replace('__TRACE_FILE__',$path))
    $groups=@(
        @{Package='Admin';Module='modAdminConsole';Procedure='SeedDemoInventoryForAutomation';Phases=@(,
            @('If modAdminInventorySeed.SeedDemoInventoryForWarehouse(', 'SEED_CALL'))},
        @{Package='Admin';Module='modAdminInventorySeed';Procedure='SeedDemoInventoryForWarehouse';Phases=@(
            @('On Error GoTo FailSeed', 'SEED_ENTER'),
            @('Set activeGroups = BuildActiveDemoGroupIndex', 'SEED_INDEX'),
            @('Set payloadItems = BuildDemoInventoryPayload', 'SEED_PAYLOAD'),
            @('SeedDemoInventoryForWarehouse = QueueDemoCreateAndProcess(', 'SEED_QUEUE'))},
        @{Package='Admin';Module='modAdminInventorySeed';Procedure='BuildActiveDemoGroupIndex';Phases=@(
            @('Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge', 'SEED_INDEX_RESOLVE'),
            @('Set entityTable = FindTableByNameSeed', 'SEED_INDEX_TABLE'),
            @('Set BuildActiveDemoGroupIndex = groups', 'SEED_INDEX_DONE'))},
        @{Package='Admin';Module='modAdminInventorySeed';Procedure='QueueDemoCreateAndProcess';Phases=@(
            @('On Error GoTo FailQueue', 'SEED_QUEUE_ENTER'),
            @('If Not EnsureDemoStationInboxes', 'SEED_INBOXES'),
            @('payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection', 'SEED_SERIALIZE'),
            @('productionInboxPath = modRoleEventWriter.ResolveInboxWorkbookPath(', 'SEED_INBOX_PATH'),
            @('Set productionInbox = modRoleEventWriter.OpenInboxWorkbook(', 'SEED_INBOX_OPEN'),
            @('If Not modRoleEventWriter.QueueInventoryCreateEvent(', 'SEED_EVENT_QUEUE'),
            @('processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport)', 'SEED_PROCESS'),
            @('processedCount = modProcessor.RunBatch(warehouseId, 0, retryReport)', 'SEED_PROCESS_RETRY'),
            @('QueueDemoCreateAndProcess = True', 'SEED_PROCESS_DONE'),
            @('If Not productionInboxWasOpen Then productionInbox.Close', 'SEED_INBOX_CLOSE'))},
        @{Package='Core';Module='modProcessor';Procedure='RunBatch';Phases=@(
            @('If Not EnsurePhase2Context', 'PROCESS_CONTEXT'),
            @('If Not modAuth.HasProvisionedCapabilityForSystem("INBOX_PROCESS"', 'PROCESS_AUTH'),
            @('localStagingOk = modRoleEventWriter.SyncLocalStagedInboxRows', 'PROCESS_LOCAL_STAGING'),
            @('Set inventoryWb = ResolveInventoryWorkbookBridge', 'PROCESS_INVENTORY_RESOLVE'),
            @('If Not modLockManager.AcquireLock', 'PROCESS_LOCK'),
            @('Set inboxTargets = ResolveInboxTargets', 'PROCESS_INBOX_TARGETS'),
            @('Set evt = BuildInboxEvent', 'PROCESS_EVENT_READ'),
            @('eventApplied = ApplyInventoryEventBridge', 'PROCESS_INVENTORY_APPLY'),
            @('If eventApplied Then', 'PROCESS_APPLY_RETURNED'),
            @('SaveWorkbookProcessor inventoryWb', 'PROCESS_INVENTORY_SAVE'),
            @('If AppendEventsToOutboxBatch', 'PROCESS_OUTBOX'),
            @('SaveWorkbookProcessor target("Workbook")', 'PROCESS_INBOX_SAVE'),
            @('If Not GenerateWarehouseSnapshot', 'PROCESS_SNAPSHOT'),
            @('If Not PublishWarehouseArtifactsToSharePoint', 'PROCESS_PUBLISH'),
            @('If RunBatch > 0 Then modInventoryDomainBridge.ScheduleSourceWorkbookSyncBridge', 'PROCESS_SCHEDULE'),
            @('If Not inboxTargets Is Nothing Then', 'PROCESS_CLEANUP'),
            @('If lockHeld Then Call modLockManager.ReleaseLock', 'PROCESS_UNLOCK'),
            @('If inventoryOpenedTransient Then CloseTransientProcessorWorkbook', 'PROCESS_INVENTORY_CLOSE'),
            @('If perfOwned Then PerfEndSafeProcessor', 'PROCESS_DONE'))}
    )
    foreach($group in $groups){
        $module=$packages['invSys.'+$group.Package+'.xlam'].VBProject.VBComponents.Item($group.Module).CodeModule
        $start=$module.ProcStartLine($group.Procedure,0);$count=$module.ProcCountLines($group.Procedure,0)
        $source=$module.Lines($start,$count)
        $writer=if($group.Package -ceq 'Core'){'TestBootstrapTrace'}else{'TestSeedTrace'}
        foreach($phase in $group.Phases){
            if($phase -isnot [array] -or $phase.Count -ne 2 -or $phase[1] -cnotmatch '^[A-Z_]+$'){
                throw 'Invalid fixed-label evaluation trace definition.'
            }
            $pattern='(?im)^([ \t]*)'+[regex]::Escape($phase[0])
            if([regex]::Matches($source,$pattern).Count -ne 1){throw ('Evaluation trace anchor unavailable: '+$phase[1])}
            $source=[regex]::Replace($source,$pattern,('$1'+$writer+'.Mark "'+$phase[1]+'"'+"`r`n"+'$0'))
        }
        $module.DeleteLines($start,$count);$module.InsertLines($start,$source)
    }
}

function Compile-Slice4beEvaluationProbes {
    $vbeVisible=$excel.VBE.MainWindow.Visible
    if($null -eq $vbeVisible){throw 'VBE visibility is unavailable.'}
    try {
        foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){
            $book=$packages[$name]
            if(-not [string]::Equals($book.FullName,(Join-Path $deploy $name),[StringComparison]::OrdinalIgnoreCase)){
                throw 'Instrumented workbook is outside the candidate package directory.'
            }
            $project=$book.VBProject
            foreach($reference in $project.References){
                if($reference.IsBroken){throw ('Broken instrumented reference: '+$name)}
                if($reference.Name -like 'invSys_*' -and -not [string]::Equals(
                    (Split-Path -Parent $reference.FullPath),$deploy,[StringComparison]::OrdinalIgnoreCase)){
                    throw 'Instrumented package reference is outside the candidate directory.'
                }
            }
            $excel.VBE.ActiveVBProject=$project
            foreach($component in $project.VBComponents){if($component.Type -eq 1){$component.CodeModule.CodePane.Show();break}}
            Start-Sleep -Milliseconds 300
            $control=$excel.VBE.CommandBars.FindControl(1,578)
            if($null -eq $control -or $null -eq $control.Enabled){throw 'Instrumented compile command unavailable.'}
            if($control.Enabled){$control.Execute()}
            Start-Sleep -Milliseconds 500
            $control=$excel.VBE.CommandBars.FindControl(1,578)
            if($null -eq $control -or $null -eq $control.Enabled -or $control.Enabled){
                throw ('Instrumented compilation did not finish: '+$name)
            }
            Check ('Harness.InstrumentedCompile.'+$name) $true
        }
    } finally {
        $excel.VBE.MainWindow.Visible=[bool]$vbeVisible
    }
}
