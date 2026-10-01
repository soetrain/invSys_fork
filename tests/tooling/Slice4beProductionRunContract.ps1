# Supplemental headless Core checks; retain the separate packaged actual-handler baselines.
function Install-ProductionRunContractProbe {
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function RunDefaultForTest(ByVal id As String, ByVal enabled As Boolean) As String
    Dim model As Object, row As Variant
    Set model = modTrackingPolicyModel.Defaults(enabled)
    For Each row In model("Controls")
        If CStr(row("ControlId")) = id Then
            RunDefaultForTest = CStr(row("Collect")) & "|" & CStr(row("Visible")) & "|" & CStr(row("SequenceEligible"))
            Exit Function
        End If
    Next row
End Function
'@)
}
function Test-ProductionRunContract {
    $positive=[ordered]@{SCALE='STAGED';CLEAR='STAGED';LOAD='STAGED';LOADER_REFRESH='REFRESHED';MANAGER_REFRESH='REFRESHED';ALLOCATE='STAGED';TREE_ALLOCATE='STAGED';TREE_EXPAND='PRESENTED';TREE_COLLAPSE='PRESENTED'}
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(23))).Split("`n")|Where-Object{$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(24))).Split("`n")|Where-Object{$_})
    Check 'RunContract.CatalogExtends23' ($old.Count -eq 118 -and $new.Count -eq 127 -and @($new|Select-Object -Unique).Count -eq 127 -and @($old|Where-Object{$_ -cnotin $new}).Count -eq 0)
    foreach($id in $old){
        $prior=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,23))
        $latest=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,24))
        Check ('RunContract.Preserve23.'+$id) ($prior -cne '' -and $prior -ceq $latest)
    }
    $codes=@('REQUESTED','DENIED','REJECTED','FAILED','REFRESHED','PRESENTED','SELECTED','STAGED','CONFIRMED','PENDING','APPLIED','COMPLETED','VALIDATED','CANCELLED')
    foreach($action in $positive.Keys){
        $id='PRODUCTION_RUN_'+$action;$label='RunContract.'+$action
        $excluded=$true
        foreach($version in 1..23){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
        Check ($label+'.OlderVersionsExclude') $excluded
        $navigation=$action -cin @('TREE_EXPAND','TREE_COLLAPSE')
        $default=if($navigation){'False|True|True'}else{'True|True|True'}
        Check ($label+'.EnabledDefault') ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunDefaultForTest' @($id,$true)) -ceq $default)
        Check ($label+'.DisabledDefault') ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunDefaultForTest' @($id,$false)) -ceq 'False|False|False')
        $outcomes=[ordered]@{REQUESTED=@('Info','Unknown');DENIED=@('Blocked','Unchanged');REJECTED=@('Warning','Unchanged');FAILED=@('Error','Unknown')}
        $outcomes[$positive[$action]]=@('Info','Unchanged')
        foreach($code in $codes){
            $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$code))
            $value=if($wire){$wire|ConvertFrom-Json}else{$null};$supported=$outcomes.Contains($code)
            $correct=if($supported){$null -ne $value -and $value.EventCode -ceq ($id+'_'+$code) -and $value.OutcomeCode -ceq $code -and $value.Severity -ceq $outcomes[$code][0] -and $value.DataEffect -ceq $outcomes[$code][1] -and $value.UserMessage -ne '' -and $null -ne $value.PSObject.Properties['NextStep']}else{$wire -ceq ''}
            Check ($label+'.Outcome.'+$code) $correct
            $record=@{ControlId=$id;OwnerId='PRODUCTION_RUN_LOCAL';CatalogVersion=24;OutcomeCode=$code}|ConvertTo-Json -Compress
            Check ($label+'.Terminal.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @($record)) -eq ($code -ceq $positive[$action]))
            Check ($label+'.EmptyReferences.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,'[]')) -eq $supported)
        }
        foreach($code in @('REQUESTED','DENIED','REJECTED','FAILED',$positive[$action])){
            foreach($kind in @('Inventory','Designs')){
                foreach($state in @('Submitted','Unknown')){
                    $refs=ConvertTo-Json -InputObject @(@{WarehouseId='CATALOG_TEST';SourceKind=$kind;EventId='Source_A';SubmissionState=$state}) -Compress
                    Check ($label+'.RejectSource.'+$code+'.'+$kind+'.'+$state) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,$refs)))
                }
            }
        }
        foreach($change in @('Owner','Catalog')){
            $record=@{ControlId=$id;OwnerId='PRODUCTION_RUN_LOCAL';CatalogVersion=24;OutcomeCode=$positive[$action]}
            if($change -ceq 'Owner'){$record.OwnerId='PRODUCTION_ASSIGNMENT'}else{$record.CatalogVersion=23}
            Check ($label+'.TerminalRejectsWrong'+$change) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($record|ConvertTo-Json -Compress))))
        }
    }
}
