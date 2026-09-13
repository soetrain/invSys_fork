# Explicit diagnostic only. Fixed phase names; no values, identities or credentials.
function Install-Slice4beBootstrapTrace {
    $project=$packages['invSys.Core.xlam'].VBProject
    $trace=$project.VBComponents.Add(1);$trace.Name='TestBootstrapTrace'
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
    $module=$project.VBComponents.Item('modWarehouseBootstrap').CodeModule
    $groups=@(
        @{Name='BootstrapWarehouseLocalValues';Phases=@(
            @('spec.WarehouseId = warehouseId','VALUES_ENTER'),
            @('BootstrapWarehouseLocalValues = BootstrapWarehouseLocal(spec)','VALUES_CALL'))},
        @{Name='BootstrapWarehouseLocal';Phases=@(
            @('On Error GoTo FailBootstrap','BOOTSTRAP_ENTER'),
            @('If Not ValidateWarehouseSpec(spec, report)','VALIDATE'),
            @('rootPath = ResolveBootstrapRootPath(spec)','RESOLVE_ROOT'),
            @('modRuntimeWorkbooks.SetCoreDataRootOverride rootPath','SET_ROOT'),
            @('If Not CopyInventoryTemplateBootstrap','COPY_INVENTORY'),
            @('If Not modConfig.EnsureStationConfigEntry(spec.WarehouseId, spec.StationId, spec.AdminUser, rootPath & "\inbox\", "ADMIN"','CONFIG_ADMIN'),
            @('Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime','CONFIG_OPEN'),
            @('If Not StampBootstrapConfigWorkbook','CONFIG_STAMP'),
            @('If Not modConfig.EnsureStationConfigEntry(spec.WarehouseId, spec.StationId, spec.AdminUser, rootPath & "\inbox\", "RECEIVE"','CONFIG_RECEIVE'),
            @('If Not modConfig.EnsureStationInbox','RECEIVE_INBOX'),
            @('If Not modAuth.EnsureStationRoleAuth(spec.WarehouseId, spec.StationId, spec.AdminUser, spec.AdminUser, "ADMIN"','AUTH_ADMIN'),
            @('If Not modAuth.EnsureStationRoleAuth(spec.WarehouseId, spec.StationId, spec.AdminUser, spec.AdminUser, "RECEIVE"','AUTH_RECEIVE'),
            @('Set wbInventory = ResolveInventoryWorkbookBridge','INVENTORY_RESOLVE'),
            @('If Not GenerateWarehouseSnapshot','SNAPSHOT'),
            @('Set wbOutbox = ResolveOutboxWorkbook','OUTBOX'),
            @('If Not modConfig.LoadConfig','CONFIG_LOAD'),
            @('If Not modAuth.LoadAuth','AUTH_LOAD'),
            @('If Not SeedBootstrapDemoInventory','SEED'),
            @('operatorPath = BuildReceivingOperatorPathBootstrap','OPERATOR_PATH'),
            @('If Not PrepareReceivingOperatorRuntimeFilesBootstrap','OPERATOR_FILES'),
            @('If Not CreateOrVerifyReceivingOperatorWorkbookBootstrap','OPERATOR_CREATE'),
            @('BootstrapWarehouseLocal = True','BOOTSTRAP_SUCCESS'),
            @('BootstrapWarehouseLocal = False','BOOTSTRAP_FAILED'),
            @('RestoreCoreRootOverrideBootstrap priorRootOverride','RESTORE_ROOT'))},
        @{Name='CreateOrVerifyReceivingOperatorWorkbookBootstrap';Phases=@(
            @('On Error GoTo FailCreate','OPERATOR_ENTER'),
            @('Set wb = FindOpenWorkbookByPathBootstrap','OPERATOR_FIND'),
            @('Set wb = Application.Workbooks.Open','OPERATOR_OPEN'),
            @('Set wb = Application.Workbooks.Add','OPERATOR_ADD'),
            @('If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface','OPERATOR_SURFACE'),
            @('RemoveNonReceivingOperatorSheetsBootstrap','OPERATOR_STRIP_SHEETS'),
            @('Call modOperatorReadModel.RefreshInventoryReadModelForWorkbook','OPERATOR_REFRESH'),
            @('wb.SaveAs Filename:=operatorPath','OPERATOR_SAVE_AS',2),
            @('wb.Save','OPERATOR_SAVE',3),
            @('mLastBootstrapOperatorWorkbookPath = operatorPath','OPERATOR_SAVED'),
            @('CloseWorkbookIfOpenBootstrap wb','OPERATOR_CLOSE'))}
    )
    foreach($group in $groups){
        $start=$module.ProcStartLine($group.Name,0);$count=$module.ProcCountLines($group.Name,0)
        $source=$module.Lines($start,$count)
        foreach($phase in $group.Phases){
            $pattern='(?im)^([ \t]*)'+[regex]::Escape($phase[0])
            $expected=if($phase.Count -gt 2){$phase[2]}else{1}
            if([regex]::Matches($source,$pattern).Count -ne $expected){throw ('Bootstrap phase anchor unavailable: '+$phase[1])}
            $replacement='$1TestBootstrapTrace.Mark "'+$phase[1]+'"'+"`r`n"+'$0'
            $source=[regex]::Replace($source,$pattern,$replacement)
        }
        $module.DeleteLines($start,$count);$module.InsertLines($start,$source)
    }
    [void](Run 'invSys.Core.xlam' 'TestBootstrapTrace.Mark' @('TRACE_READY'))
}
