# Use the real native and imported guide bodies produced by this disposable round trip.
function Test-GuideLegacyVersion($Schema,$Example,$Source,$Destination,$Files,$TransferFields) {
    $prefix="GuideCourse.LegacySchema$Schema"
    $sourceRoot=Join-Path $Source.Root ('Training/ActionPaths/'+$Source.Warehouse+'/Guides')
    $destinationRoot=Join-Path $Destination.Root ('Training/ActionPaths/'+$Destination.Warehouse+'/Guides')
    $body=[ordered]@{}
    foreach($property in $Example.PSObject.Properties){
        if($property.Name -notin @('Subject','ContentSha256')){$body[$property.Name]=$property.Value}
    }
    $body.SchemaVersion=$Schema
    $body.ActionPathId=[guid]::NewGuid().ToString()
    $body.RecordId=[guid]::NewGuid().ToString()
    $body.Version=1
    $body.PreviousRecordId=''
    $body.PreviousSha256=''
    $path=Join-Path $sourceRoot ($body.ActionPathId+'.1.json')
    $raw=TransferSeal (TransferJson $body)
    [IO.File]::WriteAllText($path,$raw,[Text.Encoding]::ASCII)
    $legacy=ReadGuideExpectationRecord $path
    if($null -eq $legacy){throw 'Legacy fixture failed independent field and digest checks.'}

    TransferOpen $Source
    $key=$legacy.ActionPathId+'|1|'+$legacy.ContentSha256
    $listed=$key -cin ((BoundControl 'lstPublishedGuides' 'Values') -split "`n")
    Check "$prefix.ListsExactLegacyVersion" $listed
    if(-not $listed){return}
    BoundSelect $legacy
    $text=BoundControl 'txtPublishedInstructions' 'Text'
    Check "$prefix.ReadOriginalTitleAndUnassignedSubject" ($text.Contains('Subject: Unassigned') -and $text.Contains('Title: '+$legacy.Name))
    $export=Join-Path $Files ("legacy-schema-$Schema.json")
    TransferSetFile $export
    [void](BoundControl 'btnExportGuide' 'Click')
    $package=TransferRead $export 'GuideTransfer' 1 $TransferFields
    $exported=$null -ne $package -and [IO.File]::ReadAllText($export).Contains('"Guide":'+$raw)
    Check "$prefix.ExportExactOriginal" $exported
    if(-not $exported){return}

    [void](BoundControl 'btnEditPublishedGuide' 'Click')
    [void](BoundControl 'txtGuideSubject' 'Write' 'Legacy guide qualification' 'frmActionPathGuide')
    [void](BoundControl 'btnSaveGuide' 'Click' '' 'frmActionPathGuide')
    $revision=ReadGuideExpectationRecord (Join-Path $sourceRoot ($legacy.ActionPathId+'.2.json'))
    $linked=$null -ne $revision -and $revision.SchemaVersion -eq 3 -and $revision.Version -eq 2 -and
        $revision.ActionPathId -ceq $legacy.ActionPathId -and $revision.PreviousRecordId -ceq $legacy.RecordId -and
        $revision.PreviousSha256 -ceq $legacy.ContentSha256 -and $revision.Subject -ceq 'Legacy guide qualification' -and
        $revision.Name -ceq $legacy.Name -and (TransferJson $revision.Steps) -ceq (TransferJson $legacy.Steps) -and
        (TransferJson $revision.Observations) -ceq (TransferJson $legacy.Observations)
    if($linked -and $Schema -eq 2){$linked=(TransferJson $revision.TransferOrigin) -ceq (TransferJson $legacy.TransferOrigin)}
    Check "$prefix.EditAppendsLinkedSchema3" $linked
    Check "$prefix.OriginalBytesPreserved" ([IO.File]::ReadAllText($path) -ceq $raw)
    [void](BoundControl 'btnCancelGuide' 'Click' '' 'frmActionPathGuide')

    TransferOpen $Destination
    $before=BoundPins $destinationRoot
    TransferSetFile $export
    [void](BoundControl 'btnImportGuide' 'Click')
    $after=BoundPins $destinationRoot
    $added=@($after.Keys|Where-Object {-not $before.ContainsKey($_)})
    $local=$null
    if($added.Count -eq 1){$local=ReadGuideExpectationRecord $added[0]}
    $imported=$null -ne $local -and $local.SchemaVersion -eq 3 -and $local.Version -eq 1 -and
        $local.WarehouseId -ceq $Destination.Warehouse -and $local.ActionPathId -cne $legacy.ActionPathId -and
        $local.Subject -ceq '' -and $local.Name -ceq $legacy.Name -and
        $local.TransferOrigin.ContentSha256 -ceq $legacy.ContentSha256 -and $local.TransferOrigin.TransferSha256 -ceq $package.ContentSha256 -and
        (TransferJson $local.Observations) -ceq (TransferJson $legacy.Observations)
    Check "$prefix.ImportUnassignedWithBoundOrigin" ($imported -and (TransferRetained $before $after))
    if(-not $imported){return}
    [void](BoundControl 'btnRefreshGuides' 'Click')
    BoundSelect $local
    [void](BoundControl 'btnEditPublishedGuide' 'Click')
    [void](BoundControl 'btnSaveGuide' 'Click' '' 'frmActionPathGuide')
    $next=Join-Path $destinationRoot ($local.ActionPathId+'.2.json')
    Check "$prefix.ImportedEditRequiresSubject" (-not (Test-Path -LiteralPath $next) -and (BoundControl '' 'Count' '' 'frmActionPathGuide') -ceq '1')
    [void](BoundControl 'txtGuideSubject' 'Write' 'Assigned after import' 'frmActionPathGuide')
    [void](BoundControl 'btnSaveGuide' 'Click' '' 'frmActionPathGuide')
    $assigned=ReadGuideExpectationRecord $next
    Check "$prefix.AssignmentPreservesImportedVersion" ($null -ne $assigned -and $assigned.Subject -ceq 'Assigned after import' -and
        $assigned.PreviousSha256 -ceq $local.ContentSha256 -and (TransferRetained $after (BoundPins $destinationRoot)))
    [void](BoundControl 'btnCancelGuide' 'Click' '' 'frmActionPathGuide')
}
