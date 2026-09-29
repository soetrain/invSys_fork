# Diagnostic instrumentation only: fixed stage names, no values or credentials.
# Install before forms, compile, arm at projection deletion, never save packages.
function Get-ProjectionBoundaryTracePlan {
    $groups=@(
        @('invSys.Core.xlam','modProcessor','RunBatch',@(
            'If Not EnsurePhase2Context(',
            'warehouseId = modConfig.GetString(',
            'localStagingOk = modRoleEventWriter.SyncLocalStagedInboxRows(',
            'CloseHiddenReadOnlyInventoryWorkbookProcessor warehouseId',
            'Set inventoryWb = ResolveInventoryWorkbookBridge(',
            'If Not modLockManager.AcquireLock(',
            'Set inboxTargets = ResolveInboxTargets(',
            'If Not EnsureInboxTargetSchema(',
            'eventApplied = ApplyInventoryEventBridge(',
            'SaveWorkbookProcessor inventoryWb',
            'If Not PublishWarehouseArtifactsToSharePoint(')),
        @('invSys.Inventory.Domain.xlam','modInventoryApply','ApplyEvent',@(
            'Set wb = ResolveInventoryWorkbook(',
            'If Not modInventorySchema.EnsureInventorySchema(',
            'Set loLog = FindListObjectByNameApply(',
            'RebuildInventoryProjections wb',
            'RefreshLedgerStatus wb,')),
        @('invSys.Inventory.Domain.xlam','modInventoryApply','RebuildInventoryProjections',@(
            'Set loLog = FindListObjectByNameApply(',
            'If loLog Is Nothing Or loEntities Is Nothing',
            'RewriteEntityProjectionTable loEntities,',
            'RewriteSkuProjectionTable loSku,',
            'RewriteLocationProjectionTable loLoc,',
            'End Sub')),
        @('invSys.Inventory.Domain.xlam','modInventorySchema','EnsureInventorySchema',@(
            'If targetWb Is Nothing Then',
            'EnsureTableWithHeaders wb, SHEET_SKU_BALANCE,',
            'EnsureTableWithHeaders wb, SHEET_LOCATION_BALANCE,',
            'RemoveProhibitedRowHeaders wb,',
            'EnsureInventorySchema = True')),
        @('invSys.Inventory.Domain.xlam','modInventorySchema','EnsureTableWithHeaders',@(
            'Set ws = EnsureWorksheet(',
            'EnsureWorksheetEditableSchema ws',
            'Set lo = FindListObjectByName(',
            'Set startCell = GetNextTableStartCell(',
            'Set tableRange = ws.Range(',
            'Set lo = ws.ListObjects.Add(',
            'lo.Name = tableName',
            'RemoveBlankSeedRow lo',
            'StyleProtectedHeaders lo,',
            'End Sub'))
    )
    foreach($group in $groups){
        $index=0
        foreach($anchor in $group[3]){
            $index++
            [pscustomobject]@{Package=$group[0];Module=$group[1];Procedure=$group[2];Anchor=$anchor;Stage=($group[1]+'.'+$group[2]+'.'+$index)}
        }
    }
}
function Install-ProjectionBoundaryTrace {
    param($Excel,[hashtable]$Packages,[string]$PackageRoot)
    $edits=[Collections.Generic.List[object]]::new()
    foreach($step in Get-ProjectionBoundaryTracePlan){
        $module=$Packages[$step.Package].VBProject.VBComponents.Item($step.Module).CodeModule
        $start=$module.ProcStartLine($step.Procedure,0)
        $count=$module.ProcCountLines($step.Procedure,0)
        $found=@(for($line=$start;$line -lt $start+$count;$line++){
            if(([string]$module.Lines($line,1)).Trim().StartsWith($step.Anchor,[StringComparison]::Ordinal)){$line}
        })
        if($found.Count -ne 1){throw ('Trace anchor missing/ambiguous: '+$step.Stage+'; not behavioral RED.')}
        $edits.Add([pscustomobject]@{Package=$step.Package;Module=$step.Module;Line=$found[0];Stage=$step.Stage})
    }
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam')){
        $probe=$Packages[$name].VBProject.VBComponents.Add(1)
        $probe.Name='TestProjectionTrace'
        $probe.CodeModule.AddFromString(@'
Option Explicit
Private mPath As String
Public Sub Arm(ByVal tracePath As String)
    mPath = tracePath
End Sub
Public Sub Mark(ByVal stage As String)
    Dim handle As Integer
    If mPath = "" Then Exit Sub
    handle = FreeFile
    Open mPath For Append As #handle
    Print #handle, stage
    Close #handle
End Sub
'@)
    }
    foreach($edit in ($edits|Sort-Object Package,Module,@{Expression='Line';Descending=$true})){
        $module=$Packages[$edit.Package].VBProject.VBComponents.Item($edit.Module).CodeModule
        $module.InsertLines($edit.Line,('    TestProjectionTrace.Mark "'+$edit.Stage+'"'))
    }
    $visible=$Excel.VBE.MainWindow.Visible
    try {
        foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')){
            $project=$Packages[$name].VBProject
            foreach($ref in $project.References){
                if($ref.IsBroken){throw 'Trace reference is broken; not behavioral RED.'}
                if($ref.Name -like 'invSys_*' -and -not [string]::Equals((Split-Path -Parent $ref.FullPath),$PackageRoot,[StringComparison]::OrdinalIgnoreCase)){
                    throw 'Trace dependency outside pinned package directory.'
                }
            }
            $Excel.VBE.ActiveVBProject=$project
            foreach($component in $project.VBComponents){if($component.Type -eq 1){$component.CodeModule.CodePane.Show();break}}
            $command=$Excel.VBE.CommandBars.FindControl(1,578)
            if($null -eq $command){throw 'Trace compile command unavailable.'}
            if($command.Enabled){$command.Execute()}
            if($Excel.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Trace compile incomplete; not behavioral RED.'}
            Write-Output ('PROJECTION_TRACE_COMPILE_PASS '+$name)
        }
    } finally {$Excel.VBE.MainWindow.Visible=$visible}
    Write-Output ('PROJECTION_TRACE_INSTALLED '+$edits.Count)
}
