function Test-GuideTransferInvalidFiles($Source,$Destination,[string]$Files,[string]$Body,[string]$GuideRaw,[string]$Header) {
    $valid=TransferSeal $Body
    $guide=$GuideRaw|ConvertFrom-Json
    $cases=[ordered]@{}
    $cases.BadOuterHash=$valid.Substring(0,$valid.Length-66)+('0'*64)+'"}'
    $cases.UnknownField=TransferSeal ($Body.Substring(0,$Body.Length-1)+',"Unexpected":true}')
    $cases.DuplicateField=TransferSeal ('{"SchemaVersion":1,'+$Body.Substring(1))
    $cases.WrongSchemaType=TransferSeal ($Body -replace '^\{"SchemaVersion":1','{"SchemaVersion":"1"')
    $v2=$Body -replace '^\{"SchemaVersion":1','{"SchemaVersion":2'
    $cases.UnsupportedExecutableVersion=TransferSeal ($v2.Substring(0,$v2.Length-1)+',"Execution":{}}')
    $cases.SourceWarehouseMismatch=TransferSeal ($Body.Replace('"SourceWarehouseId":"'+$Source.Warehouse+'"','"SourceWarehouseId":"'+$Destination.Warehouse+'"'))
    $cases.NestedGuideHash=TransferSeal ($Body.Replace([string]$guide.ContentSha256,('0'*64)))
    $guide.PSObject.Properties.Remove('ContentSha256')
    $guide.OriginWarehouseId=$Destination.Warehouse
    $badGuide=TransferSeal (TransferJson $guide)
    $cases.ForeignNativeOrigin=TransferSeal ($Header.Substring(0,$Header.Length-1)+',"Guide":'+$badGuide+'}')
    $guide.OriginWarehouseId=$Source.Warehouse
    $steps=@($guide.Steps)
    for($index=2;$index -lt 257;$index++){
        $extra=(TransferJson $steps[0])|ConvertFrom-Json
        $extra.StepId=[guid]::NewGuid().ToString();$steps+=$extra
    }
    $guide.Steps=$steps
    $badGuide=TransferSeal (TransferJson $guide)
    $cases.StepLimit=TransferSeal ($Header.Substring(0,$Header.Length-1)+',"Guide":'+$badGuide+'}')
    $guide=$GuideRaw|ConvertFrom-Json;$guide.PSObject.Properties.Remove('ContentSha256')
    $guide.Instructions='x'*1048576
    $badGuide=TransferSeal (TransferJson $guide)
    $cases.SizeLimit=TransferSeal ($Header.Substring(0,$Header.Length-1)+',"Guide":'+$badGuide+'}')
    foreach($name in $cases.Keys){
        $file=Join-Path $Files ($name+'.json')
        [IO.File]::WriteAllText($file,$cases[$name],[Text.Encoding]::ASCII)
        TransferImportAttempt $name $file $Source $Destination
    }
}

function Test-GuideTransferEditAndRepeat($Source,$Destination,$Original,$Imported,[string]$InputPath,[string]$Files,[string]$GuideFields,[string]$TransferFields) {
    $root=Join-Path $Destination.Root ('Training/ActionPaths/'+$Destination.Warehouse+'/Guides')
    $before=TransferPins $Destination.Root
    $opened=(BoundControl 'btnEditPublishedGuide' 'Click') -ceq 'DELIVERED'
    Check 'GuideTransfer.Edit.OriginLabelInEditor' ($opened -and (BoundControl '' 'Labels' '' 'frmActionPathGuide').Contains('Imported origin evidence; not locally observed'))
    foreach($layout in @('Minimum','Default','Larger','Restored')){
        Check ('GuideTransfer.Edit.Layout.'+$layout) ($opened -and (BoundControl '' 'Fit' $layout 'frmActionPathGuide') -ceq 'True')
        if($CaptureEvidence -and $opened){CaptureOwnedFormByCaptionEvidence 'Action Path guide' ('guide-transfer-editor-'+$layout.ToLowerInvariant()+'.png')}
    }
    $instructions=[string]$Imported.Instructions+"`r`nLocal authored revision."
    [void](BoundControl 'txtGuideInstructions' 'Write' $instructions 'frmActionPathGuide')
    [void](BoundControl 'btnSaveGuide' 'Click' '' 'frmActionPathGuide')
    $revision=TransferRead (Join-Path $root ($Imported.ActionPathId+'.2.json')) 'Guide' 2 $GuideFields
    $saved=$null -ne $revision
    Check 'GuideTransfer.Edit.AppendsLocalVersionAndRetainsOrigin' ($opened -and $saved -and $revision.PreviousRecordId -ceq $Imported.RecordId -and $revision.PreviousSha256 -ceq $Imported.ContentSha256 -and $revision.Instructions -ceq $instructions -and (TransferJson $revision.TransferOrigin) -ceq (TransferJson $Imported.TransferOrigin) -and (TransferJson $revision.Observations) -ceq (TransferJson $Original.Observations) -and (TransferRetained $before (TransferPins $Destination.Root)))
    [void](BoundControl 'btnCancelGuide' 'Click' '' 'frmActionPathGuide')
    [void](BoundControl 'btnRefreshGuides' 'Click')
    if($saved){BoundSelect $revision}
    foreach($layout in @('Minimum','Default','Larger','Restored')){
        Check ('GuideTransfer.Import.Layout.'+$layout) ((BoundControl '' 'Fit' $layout) -ceq 'True')
        if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Published guides' ('guide-transfer-imported-'+$layout.ToLowerInvariant()+'.png')}
        if($layout -ceq 'Minimum'){
            $text=BoundControl 'txtPublishedInstructions' 'Text'
            Check 'GuideTransfer.Import.LongInstructionsBottomReachable' ((BoundControl 'txtPublishedInstructions' 'ViewportBottom') -ceq 'DELIVERED')
            if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Published guides' 'guide-transfer-imported-multiline-bottom.png'}
            Check 'GuideTransfer.Import.ScrollingPreservesCompleteMultilineText' ((BoundControl 'txtPublishedInstructions' 'Text') -ceq $text -and $text.Contains("Second line: "+[char]0x03a9+"`r`nFinal line marker.`r`nLocal authored revision."))
            [void](BoundControl 'txtPublishedInstructions' 'ViewportTop')
        }
    }
    $reexport=Join-Path $Files 'reexport.json';TransferSetFile $reexport
    $clicked=(BoundControl 'btnExportGuide' 'Click') -ceq 'DELIVERED'
    $package=TransferRead $reexport 'GuideTransfer' 1 $TransferFields
    Check 'GuideTransfer.Reexport.UsesLocalVersionAndRetainsForeignEvidence' ($clicked -and $saved -and $null -ne $package -and $package.SourceWarehouseId -ceq $Destination.Warehouse -and $package.Guide.ContentSha256 -ceq $revision.ContentSha256 -and $package.Guide.OriginWarehouseId -ceq $Source.Warehouse -and (TransferJson $package.Guide.TransferOrigin) -ceq (TransferJson $Imported.TransferOrigin))
    $before=BoundPins $root;TransferSetFile $InputPath
    $clicked=(BoundControl 'btnImportGuide' 'Click') -ceq 'DELIVERED'
    $after=BoundPins $root
    $fresh=@($after.Keys|Where-Object {-not $before.ContainsKey($_)})
    $again=$null;if($fresh.Count -eq 1){$again=TransferRead $fresh[0] 'Guide' 2 $GuideFields}
    Check 'GuideTransfer.Import.RepeatedImportCreatesDistinctIdentity' ($clicked -and $null -ne $again -and $again.ActionPathId -cne $Imported.ActionPathId -and $again.ActionPathId -cne $Original.ActionPathId -and $again.Version -eq 1 -and (TransferRetained $before $after))
    if($null -ne $package){
        TransferOpen $Source
        $sourceRoot=Join-Path $Source.Root ('Training/ActionPaths/'+$Source.Warehouse+'/Guides')
        $before=BoundPins $sourceRoot;TransferSetFile $reexport
        $clicked=(BoundControl 'btnImportGuide' 'Click') -ceq 'DELIVERED'
        $after=BoundPins $sourceRoot;$fresh=@($after.Keys|Where-Object {-not $before.ContainsKey($_)})
        $returned=$null;if($fresh.Count -eq 1){$returned=TransferRead $fresh[0] 'Guide' 2 $GuideFields}
        Check 'GuideTransfer.Import.ReexportCanReturnWithoutMerging' ($clicked -and $null -ne $returned -and $returned.ActionPathId -cne $Original.ActionPathId -and $returned.WarehouseId -ceq $Source.Warehouse -and $returned.TransferOrigin.SourceWarehouseId -ceq $Destination.Warehouse -and (TransferJson $returned.Observations) -ceq (TransferJson $Original.Observations) -and (TransferRetained $before $after))
    }
}
