# Current source runs through the existing operator handlers in unsaved test
# copies. This does not build, save, deploy or qualify a packaged XLAM.
function Install-InventorySourceOverlay {
    $sources=@(
        @('invSys.Core.xlam','Core/ClassModules/cInventoryTableInput.cls',2),
        @('invSys.Core.xlam','Core/ClassModules/cInventoryCommandInputs.cls',2),
        @('invSys.Core.xlam','Core/ClassModules/cInventoryCreatePlan.cls',2),
        @('invSys.Core.xlam','Core/ClassModules/cInventoryRetirementPlan.cls',2),
        @('invSys.Core.xlam','Core/ClassModules/cInventoryEditRow.cls',2),
        @('invSys.Core.xlam','Core/Modules/modInventoryRetirementCommands.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryCreateReceipts.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryCreateWire.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryCatalogWire.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryEditWire.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryEditPlan.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryEditProof.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryEditCommands.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryDomainBridge.bas',1),
        @('invSys.Inventory.Domain.xlam','InventoryDomain/Modules/modInventoryEditReceipts.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryCommands.bas',1),
        @('invSys.Core.xlam','Core/Modules/modInventoryQuantityCommands.bas',1),
        @('invSys.Core.xlam','Core/Modules/modAggregationFiles.bas',1),
        @('invSys.Admin.xlam','Admin/Modules/modAdmin.bas',1),
        @('invSys.Admin.xlam','Admin/Modules/modAdminInventoryCreate.bas',1),
        @('invSys.Admin.xlam','Admin/Modules/modAdminInventoryWorksheet.bas',1),
        @('invSys.Admin.xlam','Admin/Modules/modInventoryUiText.bas',1),
        @('invSys.Admin.xlam','Admin/Forms/frmAddInventoryItem.frm',3)
    )
    if($AdminInventoryTesterBaselineOverlayForTest){
        # Explicit dependencies changed since the frozen tester-recording-02
        # baseline (546b029b). Existing form designers are retained; only code
        # modules are replaced, without saving or qualifying these packages.
        $sources+=@(
            @('invSys.Admin.xlam','Admin/Forms/frmAggregationSources.frm',3),
            @('invSys.Admin.xlam','Admin/Modules/modAdminConsole.bas',1),
            @('invSys.Core.xlam','Core/ClassModules/cCapturedFormStore.cls',2),
            @('invSys.Core.xlam','Core/Modules/modActivityCatalog.bas',1),
            @('invSys.Core.xlam','Core/Modules/modAggregationActivityCodes.bas',1),
            @('invSys.Core.xlam','Core/Modules/modAggregationAuthority.bas',1),
            @('invSys.Core.xlam','Core/Modules/modAggregationCommands.bas',1),
            @('invSys.Core.xlam','Core/Modules/modAggregationOutput.bas',1),
            @('invSys.Core.xlam','Core/Modules/modAggregationRetry.bas',1),
            @('invSys.Core.xlam','Core/Modules/modAggregationSources.bas',1),
            @('invSys.Core.xlam','Core/Modules/modEvaluationMatches.bas',1),
            @('invSys.Core.xlam','Core/Modules/modExpectationDraft.bas',1),
            @('invSys.Core.xlam','Core/Modules/modHqAggregator.bas',1),
            @('invSys.Core.xlam','Core/Modules/modInventoryCommandReceipts.bas',1),
            @('invSys.Core.xlam','Core/Modules/modProcessor.bas',1),
            @('invSys.Core.xlam','Core/Modules/modProcessorPersistence.bas',1),
            @('invSys.Core.xlam','Core/Modules/modRoleEventJson.bas',1),
            @('invSys.Core.xlam','Core/Modules/modRoleEventWriter.bas',1),
            @('invSys.Core.xlam','Core/Modules/modTesterAuthority.bas',1),
            @('invSys.Core.xlam','Core/Modules/modTrainingWire.bas',1)
        )
    }
    $sourceRoot=Join-Path $repo 'src';$export=$null
    if($AdminInventoryOverlaySourceRootForTest){
        $sourceRoot=(Resolve-Path -LiteralPath $AdminInventoryOverlaySourceRootForTest).Path
        $owned=[IO.Path]::GetFullPath((Join-Path $repo 'reports/runtime')).TrimEnd('\')+'\'
        if(-not $sourceRoot.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Exported test source must be in the owned runtime report tree.'}
        $export=Get-Content (Join-Path $sourceRoot 'source-export-manifest.json') -Raw|ConvertFrom-Json
        if(-not $export.ReadOnlySourceExport -or $export.XlamBuilt -or $export.SourceCommit -notmatch '^[a-f0-9]{8,40}$' -or $export.Sources.Count -ne $sources.Count){throw 'Explicit source export manifest is invalid.'}
    }
    $pins=@()
    foreach($source in $sources){
        $path=Join-Path $sourceRoot $source[1]
        if($null -ne $export){
            $expected=@($export.Sources|Where-Object Source -CEQ $source[1])
            if($expected.Count -ne 1 -or (Get-FileHash $path).Hash -cne $expected[0].Hash){throw 'Exported source hash changed.'}
        }
        $text=Get-Content $path -Raw
        $offset=$text.IndexOf('Option Explicit')
        if($offset -lt 0){throw 'Inventory overlay source code boundary unavailable.'}
        $name=[IO.Path]::GetFileNameWithoutExtension($path)
        $project=$packages[$source[0]].VBProject
        $component=$null
        foreach($candidate in $project.VBComponents){if($candidate.Name -ceq $name){$component=$candidate;break}}
        if($null -eq $component){
            if($source[2] -eq 3){throw 'Inventory form designer unavailable.'}
            $component=$project.VBComponents.Add($source[2]);$component.Name=$name
        }
        $module=$component.CodeModule
        if($module.CountOfLines){$module.DeleteLines(1,$module.CountOfLines)}
        # CodeModule accepts executable code, not exported hidden attributes.
        $code=[regex]::Replace($text.Substring($offset),'(?m)^Attribute [^\r\n]*\r?\n','')
        $module.AddFromString($code)
        $pins+=[pscustomobject]@{Source=$source[1];Hash=(Get-FileHash $path).Hash}
    }
    $probe=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $probe.Name='TestInventoryAdmission'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function Facts(ByVal host As String, ByVal warehouse As String) As String
    Dim inputs As cInventoryCommandInputs, sourceBound As Boolean, key As String
    Set inputs = New cInventoryCommandInputs
    sourceBound = inputs.CaptureSource(warehouse)
    Set inputs = Nothing
    Application.Workbooks(host).Activate
    key = modInventoryCommands.Capture(host)
    Facts = CStr(modAuth.IsSignedIn()) & "|" & CStr(sourceBound) & "|" & CStr(key <> "")
    If key <> "" Then modInventoryCommands.Release key
End Function
'@)
    [pscustomobject]@{SourceRoot=$sourceRoot;ExportedSourceCommit=if($null -ne $export){$export.SourceCommit}else{''};SourceOverlay=$true;ActualOperatorHandlers=$true;PackagedAction=$false;
        XlamBuilt=$false;XlamSaved=$false;TesterBaselineDependencies=[bool]$AdminInventoryTesterBaselineOverlayForTest;Sources=$pins}|ConvertTo-Json -Depth 4|
        Set-Content (Join-Path $reportRoot 'inventory-source-overlay-scope.json')
}
