# Mutations are limited to the test-created dummy profile, restored byte-for-byte.
# No fixture body or row value is emitted in evidence.
function Test-ExecutionProfileSafety($Guide,$Profile,[string]$ProfilePath,$Fixture,$Other) {
    $root=Split-Path -Parent $ProfilePath
    $firstPins=BoundPins $root
    [void](BoundControl 'btnConfigureExecution' 'Click')
    $opened=(BoundControl '' 'Count' '' 'frmActionPathExecution') -ceq '1'
    $saved=(BoundControl 'btnSaveExecutionProfile' 'Click' '' 'frmActionPathExecution') -ceq 'DELIVERED'
    $added=@(Get-ChildItem -LiteralPath $root -File -Filter '*.json'|Where-Object {-not $firstPins.ContainsKey($_.FullName)})
    $revision=$null
    if($added.Count -eq 1){$revision=Get-Content -Raw -LiteralPath $added[0].FullName|ConvertFrom-Json}
    $chained=$opened -and $saved -and $null -ne $revision -and $revision.ProfileId -ceq $Profile.ProfileId -and $revision.Version -eq 2 -and $revision.RecordId -cne $Profile.RecordId -and $revision.PreviousRecordId -ceq $Profile.RecordId -and $revision.PreviousSha256 -ceq $Profile.ContentSha256 -and (Get-FileHash -LiteralPath $ProfilePath).Hash -ceq $firstPins[$ProfilePath]
    Check 'ExecutionProfile.NewVersionPreservesFirst' $chained
    [void](BoundControl 'btnCloseExecution' 'Click' '' 'frmActionPathExecution')
    Check 'ExecutionProfile.CloseUnloadsEditor' ((BoundControl '' 'Count' '' 'frmActionPathExecution') -ceq '0')
    if(-not $chained){return}
    $path=$added[0].FullName;$original=[IO.File]::ReadAllBytes($path)
    $pins=BoundPins $root
    try {
        foreach($case in @('UnknownField','DuplicateField','AdapterVersion','OriginalEntityLiteral','QuantityExpression','WrongGuideHash')){
            $model=[Text.Encoding]::UTF8.GetString($original)|ConvertFrom-Json
            $model.PSObject.Properties.Remove('ContentSha256')
            switch($case){
                'UnknownField' {Add-Member -InputObject $model -NotePropertyName Unexpected -NotePropertyValue 'reject'}
                'AdapterVersion' {$model.Steps[0].AdapterVersion=2}
                'OriginalEntityLiteral' {
                    $entry=@($model.Steps|Where-Object ControlId -CEQ 'RECEIVING_SELECT_ITEM')[0].Inputs[0]
                    $entry.Binding=[pscustomobject]@{Kind='Literal';Value=[guid]::NewGuid().ToString()}
                }
                'QuantityExpression' {@($model.Steps|Where-Object ControlId -CEQ 'RECEIVING_ADD_SELECTED')[0].Inputs[1].Binding.Value='1e3'}
                'WrongGuideHash' {$model.Guide.ContentSha256=('0'*64)}
            }
            $body=$model|ConvertTo-Json -Depth 20 -Compress
            if($case -ceq 'DuplicateField'){$body='{"SchemaVersion":1,'+$body.Substring(1)}
            if($body -match '[^\x00-\x7f]'){throw 'Synthetic profile validation fixture must be ASCII.'}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $wire=$body.Substring(0,$body.Length-1)+',"ContentSha256":"'+$hash+'"}'
            [IO.File]::WriteAllBytes($path,[Text.Encoding]::ASCII.GetBytes($wire))
            $before=BoundPins $root
            $delivery=BoundControl 'btnConfigureExecution' 'Click'
            $count=BoundControl '' 'Count' '' 'frmActionPathExecution'
            $delivered=$delivery -ceq 'DELIVERED';$refused=$count -ceq '0'
            $preserved=BoundSame $before (BoundPins $root)
            Check ('ExecutionProfile.Rejects.'+$case) ($delivered -and $refused -and $preserved)
            if(-not ($delivered -and $refused -and $preserved)){
                $label=BoundControl 'lblExecutionProfile' 'Label' '' 'frmActionPathExecution'
                $status=BoundControl 'lblPublishedGuideStatus' 'Label'
                [pscustomobject]@{Case=$case;Delivery=$delivery;EditorCount=$count;FilesPreserved=$preserved;ShowsFirstVersion=$label.Contains([string]$Profile.ContentSha256);ShowsSecondVersion=$label.Contains([string]$revision.ContentSha256);ShowsUnsavedProfile=($label -ceq 'Execution profile not saved.');LabelEmpty=($label -ceq '');LibraryShowsInvalidProfile=$status.Contains('profiles are invalid or ambiguous')}|ConvertTo-Json -Compress|Write-Output
            }
            [void](BoundControl 'btnCloseExecution' 'Click' '' 'frmActionPathExecution')
            [IO.File]::WriteAllBytes($path,$original)
        }
    } finally {[IO.File]::WriteAllBytes($path,$original)}
    Check 'ExecutionProfile.InvalidReadFixturesRestored' (BoundSame $pins (BoundPins $root))
    [void](BoundControl 'btnConfigureExecution' 'Click')
    $opened=(BoundControl '' 'Count' '' 'frmActionPathExecution') -ceq '1'
    foreach($size in @('Minimum','Default','Larger')){
        Check ('ExecutionProfile.Layout.'+$size) ($opened -and (BoundControl '' 'Fit' $size 'frmActionPathExecution') -ceq 'True')
        if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Configure execution' ('b0-profile-'+$size.ToLowerInvariant()+'.png')}
    }
    $otherRoot=Join-Path $Other.Root ('Training/ActionPaths/'+$Other.Warehouse+'/ExecutionProfiles')
    $otherPins=BoundPins $otherRoot
    SelectTarget $Other 'config-admin'
    [void](BoundControl 'btnSaveExecutionProfile' 'Click' '' 'frmActionPathExecution')
    $invalidated=(BoundControl '' 'Count' '' 'frmActionPathExecution') -ceq '0' -or (BoundControl 'btnSaveExecutionProfile' 'State' '' 'frmActionPathExecution') -ceq 'True|False'
    Check 'ExecutionProfile.ContextChangeRefusesSave' ($opened -and $invalidated -and (BoundSame $pins (BoundPins $root)) -and (BoundSame $otherPins (BoundPins $otherRoot)))
    CloseRecordingViewer
    SelectTarget $Fixture 'config-reader'
    OpenRecordingViewer
    if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Existing recording library unavailable to role fixture.'}
    BoundOpen;BoundSelect $Guide
    Check 'ExecutionProfile.ReaderCannotConfigure' ((BoundControl 'btnConfigureExecution' 'State') -ceq 'True|False' -and (BoundControl 'btnConfigureExecution' 'Click') -ceq 'DISABLED' -and (BoundControl '' 'Count' '' 'frmActionPathExecution') -ceq '0' -and (BoundSame $pins (BoundPins $root)))
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
}
