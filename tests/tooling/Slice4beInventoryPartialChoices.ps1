# Packaged expectation-editor choices only; no source overlay or catalog mutation.
function Test-InventoryPartialChoices {
    function PartialControl([string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value 'frmActionPathExpectation'
    }
    function PartialDelivered([string]$Result) {
        if($Result -cne 'DELIVERED'){throw 'PARTIAL editor gesture unavailable; not behavioral RED.'}
    }
    function PartialValues([string]$Name) {
        @((PartialControl $Name 'Values') -split "`n" | ForEach-Object {$_.TrimEnd("`r")} | Where-Object {$_ -cne ''})
    }
    function PartialSameSet([string[]]$Actual,[string[]]$Expected) {
        ($Actual.Count -eq $Expected.Count -and
            @($Actual | Select-Object -Unique).Count -eq $Actual.Count -and
            @($Actual | Where-Object {$_ -cnotin $Expected}).Count -eq 0 -and
            @($Expected | Where-Object {$_ -cnotin $Actual}).Count -eq 0)
    }
    $controls=@(PartialValues 'cboExpectedControl')
    foreach($id in @('ADMIN_INVENTORY_CREATE_TABLE','ADMIN_INVENTORY_UPLOAD_TABLE','ADMIN_INVENTORY_UOM_ADD')){
        if(@($controls | Where-Object {$_ -ceq $id}).Count -ne 1){
            throw 'Catalog47 editor controls prerequisite unavailable; not behavioral RED.'
        }
    }
    foreach($suffix in @('CREATE_TABLE','UPLOAD_TABLE')){
        $id='ADMIN_INVENTORY_'+$suffix
        $definition=[string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @($id,47))
        if(-not $definition){throw 'Catalog47 worksheet definition unavailable; not behavioral RED.'}
        PartialDelivered (PartialControl 'cboExpectedControl' 'Write' $id)
        $choices=@(PartialValues 'cboExpectedOutcome')
        if($choices.Count -eq 0){throw 'Bound outcome dropdown unavailable; not behavioral RED.'}
        $expected=if($suffix -ceq 'CREATE_TABLE'){
            @('REQUESTED','SAVED','REFUSED','PARTIAL','FAILED')
        }else{
            @('REQUESTED','CONFIRMED','EMPTY','REJECTED','REFUSED','PARTIAL','FAILED')
        }
        Check ('InventoryPartialChoices.'+$suffix+'.ExactOutcomesOffered') (PartialSameSet $choices $expected) | Out-Host
        $selected=$false
        if(@($choices | Where-Object {$_ -ceq 'PARTIAL'}).Count -eq 1){
            PartialDelivered (PartialControl 'cboExpectedOutcome' 'Write' 'PARTIAL')
            $selected=(PartialControl 'cboExpectedOutcome' 'Selected') -ceq 'PARTIAL'
        }
        Check ('InventoryPartialChoices.'+$suffix+'.PartialSelectable') $selected | Out-Host
    }
    PartialDelivered (PartialControl 'cboExpectedControl' 'Write' 'ADMIN_INVENTORY_UOM_ADD')
    $uom=@(PartialValues 'cboExpectedOutcome')
    if($uom.Count -eq 0){throw 'UOM outcome dropdown unavailable; not behavioral RED.'}
    Check 'InventoryPartialChoices.UOM.ExactOutcomesExcludePartial' (PartialSameSet $uom @('REQUESTED','STAGED','CANCELLED','REFUSED','FAILED')) | Out-Host
    PartialDelivered (PartialControl 'cboExpectedOutcome' 'Write' 'STAGED')
    Check 'InventoryPartialChoices.UOM.RetainedSelectionWorks' ((PartialControl 'cboExpectedOutcome' 'Selected') -ceq 'STAGED') | Out-Host
}
