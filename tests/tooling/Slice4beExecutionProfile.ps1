# Actual profile editor handlers only; no profile is fabricated by this fixture.
function Test-ReceivingExecutionProfile($Guide,$Staging,[string]$JournalRoot,[string]$InventoryPath,$Fixture,$Other) {
    function ProfileControl([string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value 'frmActionPathExecution'
    }
    $root=Join-Path $JournalRoot 'ExecutionProfiles'
    $before=BoundPins $root;$activity=ActivityPins
    $inventoryHash=(Get-FileHash -LiteralPath $InventoryPath).Hash
    $editor=(ProfileControl '' 'Count') -ceq '1'
    $save=(ProfileControl 'btnSaveExecutionProfile' 'Click') -ceq 'DELIVERED'
    Check 'ExecutionProfile.MissingInputsRefuseSave' ($editor -and $save -and (BoundSame $before (BoundPins $root)))
    $steps=@((ProfileControl 'lstExecutionSteps' 'Values') -split "`n"|Where-Object {$_ -cne ''})
    $expected=@($Guide.Steps|ForEach-Object {[string]$_.StepId})
    Check 'ExecutionProfile.ExactGuideStepOrder' ($editor -and ($steps -join '|') -ceq ($expected -join '|'))
    $inputs=@(
        @('RECEIVING_SELECT_ITEM','SourceEntity','Prompt','RECEIVING_SOURCE_ENTITY'),
        @('RECEIVING_ADD_SELECTED','Reference','Literal','B0-REPLAY-REFERENCE'),
        @('RECEIVING_ADD_SELECTED','Quantity','Literal','2.5'),
        @('RECEIVING_ADD_SELECTED','Location','Literal','B0-REPLAY-LOCATION'),
        @('RECEIVING_ADD_SELECTED','LotNumber','Literal','B0-REPLAY-LOT'),
        @('RECEIVING_ADD_SELECTED','Condition','Literal','GOOD')
    )
    $delivered=$editor
    foreach($input in $inputs){
        $step=@($Guide.Steps|Where-Object ControlId -CEQ $input[0])
        if($step.Count -ne 1){throw 'B0 profile source requires one occurrence of each input-bearing action.'}
        $index=[Array]::IndexOf($steps,[string]$step[0].StepId)
        if($index -lt 0){$delivered=$false;continue}
        $selected=(ProfileControl 'lstExecutionSteps' 'Select' ([string]$index)) -ceq 'SELECTED'
        $named=(ProfileControl 'cboExecutionInput' 'Write' $input[1]) -ceq 'DELIVERED'
        $kind=(ProfileControl 'cboExecutionBinding' 'Write' $input[2]) -ceq 'DELIVERED'
        $value=(ProfileControl 'txtExecutionValue' 'Write' $input[3]) -ceq 'DELIVERED'
        $applied=(ProfileControl 'btnApplyExecutionInput' 'Click') -ceq 'DELIVERED'
        $delivered=$delivered -and $selected -and $named -and $kind -and $value -and $applied
    }
    Check 'ExecutionProfile.ActualInputHandlers' $delivered
    $save=(ProfileControl 'btnSaveExecutionProfile' 'Click') -ceq 'DELIVERED'
    $created=@()
    if(Test-Path -LiteralPath $root){$created=@(Get-ChildItem -LiteralPath $root -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)})}
    Check 'ExecutionProfile.SaveCreatesOneVersion' ($delivered -and $save -and $created.Count -eq 1)
    $profile=$null;$valid=$false
    if($created.Count -eq 1){
        $text=[IO.File]::ReadAllText($created[0].FullName)
        $profile=$text|ConvertFrom-Json
        $marker=[regex]::Match($text,',"ContentSha256":"([0-9a-f]{64})"\}$')
        if($marker.Success -and $text -notmatch '[^\x00-\x7f]'){
            $body=$text.Substring(0,$marker.Index)+'}';$sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $valid=$profile.SchemaVersion -eq 1 -and $profile.RecordKind -ceq 'ExecutionProfile' -and @($profile.PSObject.Properties).Count -eq 17 -and $profile.ContentSha256 -ceq $hash -and $text.Length -le 1048576
        }
    }
    Check 'ExecutionProfile.ImmutableWireIntegrity' $valid
    $bound=$false;$typed=$false
    if($valid){
        $bound=$profile.Guide.ActionPathId -ceq $Guide.ActionPathId -and $profile.Guide.Version -eq $Guide.Version -and $profile.Guide.RecordId -ceq $Guide.RecordId -and $profile.Guide.ContentSha256 -ceq $Guide.ContentSha256 -and $profile.WarehouseId -ceq $Guide.WarehouseId -and (@($profile.Steps|ForEach-Object StepId) -join '|') -ceq ($expected -join '|') -and ($profile.ExpectedConclusion|ConvertTo-Json -Depth 12 -Compress) -ceq ($Guide.ExpectedConclusion|ConvertTo-Json -Depth 12 -Compress)
        $typed=$true
        foreach($input in $inputs){
            $entry=@($profile.Steps|Where-Object ControlId -CEQ $input[0]|ForEach-Object Inputs|Where-Object Name -CEQ $input[1])
            if($entry.Count -ne 1){$typed=$false;continue}
            if($entry[0].Binding.Kind -cne $input[2]){$typed=$false;continue}
            $field=if($input[2] -ceq 'Prompt'){'PromptId'}else{'Value'}
            if($entry[0].Binding.$field -cne $input[3]){$typed=$false}
        }
    }
    Check 'ExecutionProfile.BindsExactGuideAndConclusion' $bound
    Check 'ExecutionProfile.ReviewedInputsSeparateFromLogs' ($typed -and (PinsRetained $activity) -and (ActivityPins).Count -eq $activity.Count)
    [void](ProfileControl 'btnCloseExecution' 'Click')
    [void](BoundControl 'btnConfigureExecution' 'Click')
    $caption=ProfileControl 'lblExecutionProfile' 'Label'
    $reopened=$valid -and $caption.Contains([string]$profile.ProfileId) -and $caption.Contains([string]$profile.ContentSha256)
    Check 'ExecutionProfile.ReopenExactSavedVersion' $reopened
    Check 'ExecutionProfile.EditingDoesNotExecute' ($editor -and (Get-FileHash -LiteralPath $InventoryPath).Hash -ceq $inventoryHash -and $Staging.ListRows.Count -eq 0 -and (PinsRetained $activity) -and (ActivityPins).Count -eq $activity.Count)
    [void](ProfileControl 'btnCloseExecution' 'Click')
    if($valid){
        . (Join-Path $PSScriptRoot 'Slice4beExecutionProfileSafety.ps1')
        Test-ExecutionProfileSafety $Guide $profile $created[0].FullName $Fixture $Other
    }
}
