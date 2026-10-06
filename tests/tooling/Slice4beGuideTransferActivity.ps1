# Probes expose existing catalog/evaluator facts and handler entry state only.
function Install-GuideTransferActivityProbe {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $core.Name='TestGuideTransferActivity'
    $core.CodeModule.AddFromString(@'
Option Explicit
Public Function Definition(ByVal id As String, ByVal version As Long) As String
    Dim value As Object
    Set value = modActivityCatalog.Control(id, version)
    If Not value Is Nothing Then Definition = modTrainingJson.EncodeObject(value)
End Function
Public Function Terminal(ByVal serialized As String) As Boolean
    Terminal = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(serialized))
End Function
'@)
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmActionPathLibrary').CodeModule
    $form.AddFromString(@'
Public Sub TransferEntryStateForTest(ByVal mode As String, ByVal state As String)
    On Error GoTo Cleanup
    mLoading = (state = "Loading"): mTransferring = (state = "Busy")
    If mode = "Export" Then mExport_Click Else mImport_Click
Cleanup:
    mLoading = False: mTransferring = False
End Sub
'@)
    $ops=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $ops.AddFromString(@'
Public Function TransferEntryStateForTest(ByVal mode As String, ByVal state As String) As Boolean
    Dim form As Object
    For Each form In VBA.UserForms
        If TypeName(form) = "frmActionPathLibrary" Then
            form.TransferEntryStateForTest mode, state
            TransferEntryStateForTest = True: Exit Function
        End If
    Next form
End Function
'@)
}

function Test-GuideTransferActivity($Fixture,$Other,$Guide) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $files=Join-Path $runRoot 'transfer-activity';New-Item -ItemType Directory -Path $files|Out-Null
    function Observe($Target,[string]$Mode,[string]$Expected,[string]$Case,[string]$Path='', [string]$State='',[bool]$SignOut=$false){
        $root=Join-Path $Target.Root ('Training/Activity/'+$Target.Warehouse)
        $before=BoundPins $root
        TransferSetFile $Path $SignOut
        if($State -ne ''){$delivered=[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.TransferEntryStateForTest' @($Mode,$State))}
        else{$delivered=(BoundControl ('btn'+$Mode+'Guide') 'Click') -ceq 'DELIVERED'}
        $after=BoundPins $root;$raw=@($after.Keys|Where-Object {-not $before.ContainsKey($_)}|ForEach-Object {[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object {$_|ConvertFrom-Json});$id='VIEWER_GUIDE_'+$Mode.ToUpperInvariant();$label='GuideTransferActivity.'+$Case
        Check ($label+'.PriorRecordsImmutable') (TransferRetained $before $after)
        if($Expected -ceq 'NONE'){Check ($label+'.NoInventedActivity') ($delivered -and $records.Count -eq 0);return}
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Expected)
        $wanted=2;if($Expected -ceq 'INTERRUPTED'){$wanted=1}
        $pair=$delivered -and $records.Count -eq $wanted -and $first.Count -eq 1 -and ($wanted -eq 1 -or $last.Count -eq 1)
        $facts=$pair;$safe=$pair;$integrity=$pair;$terminal=$pair
        foreach($r in $records){
            $facts=$facts -and $r.ControlId -ceq $id -and $r.OwnerId -ceq 'CORE_GUIDE_TRANSFER' -and $r.CatalogVersion -eq 30 -and $r.SourceRole -ceq 'Viewer' -and $r.UserId -ceq 'config-admin' -and $r.WarehouseId -ceq $Target.Warehouse -and $r.StationId -ceq 'S1' -and $r.SequenceId -ceq ''
            $safe=$safe -and @($r.SourceEventRefs).Count -eq 0
            $isTerminal=[bool](Run 'invSys.Core.xlam' 'TestGuideTransferActivity.Terminal' @(($r|ConvertTo-Json -Depth 12 -Compress)))
            $terminal=$terminal -and $isTerminal -eq ($r.OutcomeCode -ceq 'COMPLETED')
        }
        foreach($text in $raw){
            foreach($hidden in @($Target.Root,$Path,[string]$Guide.Name,[string]$Guide.Instructions,'Transfer fixture diagnostic content')){
                if($hidden -ne '' -and ($text.Contains($hidden) -or $text.Contains((TransferJson $hidden).Trim('"')))){$safe=$false}
            }
            $match=[regex]::Match($text,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false}else{$integrity=$integrity -and (TransferHash ($text.Substring(0,$match.Index)+'}')) -ceq $match.Groups[1].Value}
        }
        if($pair){
            $facts=$facts -and $first[0].EventCode -ceq ($id+'_REQUESTED') -and $first[0].DataEffect -ceq 'Unknown' -and $seen.Add([string]$first[0].ActivityId)
            if($wanted -eq 2){
                $severity=switch($Expected){'CANCELLED'{'Notice'} 'DENIED'{'Blocked'} 'REJECTED'{'Warning'} 'FAILED'{'Error'} default{'Info'}}
                $effect=switch($Expected){'COMPLETED'{'Changed'} 'FAILED'{'Unknown'} default{'Unchanged'}}
                $facts=$facts -and $last[0].ActivityId -ceq $first[0].ActivityId -and $last[0].RecordId -cne $first[0].RecordId -and $last[0].EventCode -ceq ($id+'_'+$Expected) -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect
            }
        }
        Check ($label+'.ExactAttemptAndOwnerResult') $pair
        Check ($label+'.IdentityContextAndOutcomeFacts') $facts
        Check ($label+'.RedactedIntegrityAndNoSources') ($safe -and $integrity)
        Check ($label+'.ExplicitTerminalMap') $terminal
    }
    try {
        TransferOpen $Fixture;BoundSelect $Guide
        foreach($mode in @('Export','Import')){
            $id='VIEWER_GUIDE_'+$mode.ToUpperInvariant()
            $definition=[string](Run 'invSys.Core.xlam' 'TestGuideTransferActivity.Definition' @($id,30))
            $model=$null;if($definition){$model=$definition|ConvertFrom-Json}
            Check ('GuideTransferActivity.Catalog.'+$mode) ($null -ne $model -and $model.Capability -ceq 'ACTION_PATH_MAINT' -and $model.OwnerId -ceq 'CORE_GUIDE_TRANSFER' -and $model.Class -ceq 'Command' -and $model.Role -ceq 'Viewer')
            Check ('GuideTransferActivity.OlderCatalog.'+$mode) ([string](Run 'invSys.Core.xlam' 'TestGuideTransferActivity.Definition' @($id,29)) -ceq '')
            Observe $Fixture $mode 'CANCELLED' ($mode+'Cancel')
            Observe $Fixture $mode 'NONE' ($mode+'Loading') '' 'Loading'
            Observe $Fixture $mode 'NONE' ($mode+'Busy') '' 'Busy'
        }
        $export=Join-Path $files 'export.json'
        Observe $Fixture 'Export' 'COMPLETED' 'ExportSuccess' $export
        Check 'GuideTransferActivity.Export.OwnerFileCreated' (Test-Path -LiteralPath $export)
        if(-not(Test-Path -LiteralPath $export)){throw 'Existing transfer owner failed; not observation RED.'}
        $hash=(Get-FileHash -LiteralPath $export).Hash
        Observe $Fixture 'Export' 'REJECTED' 'ExportExisting' $export
        Observe $Fixture 'Export' 'FAILED' 'ExportPickerFailure' 'TRANSFER_PICKER_FAULT'
        TransferOpen $Other
        Observe $Other 'Import' 'COMPLETED' 'ImportSuccess' $export
        $guides=Join-Path $Other.Root ('Training/ActionPaths/'+$Other.Warehouse+'/Guides')
        Check 'GuideTransferActivity.Import.OwnerGuideCreated' ((BoundPins $guides).Count -eq 1)
        $invalid=Join-Path $files 'invalid.json';[IO.File]::WriteAllText($invalid,'{}',[Text.Encoding]::ASCII)
        Observe $Other 'Import' 'REJECTED' 'ImportInvalid' $invalid
        Observe $Other 'Import' 'INTERRUPTED' 'ImportContextLoss' $export '' $true
        Check 'GuideTransferActivity.SourcePackagePreserved' ((Get-FileHash -LiteralPath $export).Hash -ceq $hash)
    } finally {TransferSetFile '';CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
