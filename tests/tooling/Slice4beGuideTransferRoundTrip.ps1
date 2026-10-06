# D18 actual form handlers. The dialog seam selects only disposable fixture paths.
function Test-GuideTransferRoundTrip($Fixture,$Other,$Guide) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    $fields='SchemaVersion|RecordKind|TransferId|ExportedAtUTC|ExportedByUserId|SourceWarehouseId|Guide|ContentSha256'
    $guideFields='SchemaVersion|RecordKind|ActionPathId|Version|RecordId|PreviousRecordId|PreviousSha256|WarehouseId|OriginWarehouseId|CreatedByUserId|CreatedAtUTC|Lifecycle|Name|Tags|Instructions|CatalogVersion|PackageSetVersion|BuildIdentity|PolicyVersion|Steps|Observations|SourceRun|ExpectedConclusion|ContentSha256|TransferOrigin'
    $files=Join-Path $runRoot 'guide-transfers';New-Item -ItemType Directory -Path $files|Out-Null
    $long=('Training instruction with retained detail. '*80)+"`r`nSecond line: "+[char]0x03a9+"`r`nFinal line marker."
    $guideRoot=Join-Path $journalRoot 'Guides'
    try {
        TransferOpen $Fixture;BoundSelect $Guide
        if((BoundControl 'btnEditPublishedGuide' 'Click') -cne 'DELIVERED'){throw 'Existing guide editor fixture unavailable.'}
        [void](BoundControl 'txtGuideInstructions' 'Write' $long 'frmActionPathGuide')
        [void](BoundControl 'btnSaveGuide' 'Click' '' 'frmActionPathGuide')
        $path=Join-Path $guideRoot ($Guide.ActionPathId+'.3.json')
        $source=ReadGuideExpectationRecord $path
        if($null -eq $source -or $source.Instructions -cne $long){throw 'Actual long/multiline guide save prerequisite failed.'}
        Check 'GuideTransfer.Fixture.ActualLongMultilineGuideSaved' $true
        [void](BoundControl 'btnCancelGuide' 'Click' '' 'frmActionPathGuide')
        [void](BoundControl 'btnRefreshGuides' 'Click');BoundSelect $source
        $raw=[IO.File]::ReadAllText($path)
        $sourcePins=TransferPins $Fixture.Root;$otherPins=TransferPins $Other.Root
        $sourceActivity=BoundPins (Join-Path $Fixture.Root 'Training/Activity')
        $otherActivity=BoundPins (Join-Path $Other.Root 'Training/Activity')
        $publicationBefore=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
        if($publicationBefore -isnot [int]){throw 'Existing publication counter unavailable.'}
        $export=Join-Path $files 'actual-export.json'
        TransferSetFile $export
        $exportClicked=(BoundControl 'btnExportGuide' 'Click') -ceq 'DELIVERED'
        $package=TransferRead $export 'GuideTransfer' 1 $fields
        $exported=$null -ne $package
        Check 'GuideTransfer.Export.ActualHandlerWritesSealedExactVersion' ($exportClicked -and $exported -and $package.Guide.ContentSha256 -ceq $source.ContentSha256 -and [IO.File]::ReadAllText($export).Contains('"Guide":'+$raw))
        Check 'GuideTransfer.Export.PreservesSourceAndOtherWarehouse' ((BoundSame $sourcePins (TransferPins $Fixture.Root)) -and (BoundSame $otherPins (TransferPins $Other.Root)))
        $known=Join-Path $files 'known-valid-guide.json'
        $header=[ordered]@{SchemaVersion=1;RecordKind='GuideTransfer';TransferId=[guid]::NewGuid().ToString();ExportedAtUTC=[DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ');ExportedByUserId='config-admin';SourceWarehouseId=$Fixture.Warehouse}
        $head=TransferJson $header
        $body=$head.Substring(0,$head.Length-1)+',"Guide":'+$raw+'}'
        $knownRaw=TransferSeal $body
        [IO.File]::WriteAllText($known,$knownRaw,[Text.Encoding]::ASCII)
        if($null -eq (TransferRead $known 'GuideTransfer' 1 $fields)){throw 'Independent transfer fixture integrity failed.'}
        # This independent valid file keeps Import testable while Export is RED.
        # Once Export exists, the real output must be used for the round trip.
        $inputPath=$known;if($exported){$inputPath=$export}
        $inputPackage=TransferRead $inputPath 'GuideTransfer' 1 $fields
        $inputHash=(Get-FileHash -LiteralPath $inputPath).Hash
        $sentinel=Join-Path $files 'existing.json';[IO.File]::WriteAllText($sentinel,'preserve existing destination',[Text.Encoding]::ASCII)
        $sentinelHash=(Get-FileHash -LiteralPath $sentinel).Hash
        TransferSetFile $sentinel
        $clicked=(BoundControl 'btnExportGuide' 'Click') -ceq 'DELIVERED'
        Check 'GuideTransfer.Export.ExistingDestinationRefused' ($clicked -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match '(?i)(exists|unavailable|refus|cannot|could not)' -and (Get-FileHash -LiteralPath $sentinel).Hash -ceq $sentinelHash)
        TransferSetFile ''
        TransferOpen $Other
        $destinationRoot=Join-Path $Other.Root ('Training/ActionPaths/'+$Other.Warehouse+'/Guides')
        if((BoundPins $destinationRoot).Count -ne 0){throw 'Destination guide library is not empty.'}
        Check 'GuideTransfer.Import.EmptyLibraryAllowsAuthorizedEntry' ((BoundControl 'btnImportGuide' 'State') -ceq 'True|True')
        . (Join-Path $PSScriptRoot 'Slice4beGuideTransferGuards.ps1')
        Test-GuideTransferInvalidFiles $Fixture $Other $files $body $raw $head
        TransferSetFile $inputPath
        $clicked=(BoundControl 'btnImportGuide' 'Click') -ceq 'DELIVERED'
        $importFiles=@();if(Test-Path -LiteralPath $destinationRoot){$importFiles=@(Get-ChildItem -LiteralPath $destinationRoot -File -Filter '*.json')}
        $imported=$null
        if($importFiles.Count -eq 1){$imported=TransferRead $importFiles[0].FullName 'Guide' 2 $guideFields}
        $importOK=$null -ne $imported
        Check 'GuideTransfer.Import.ActualHandlerCreatesOneLocalGuide' ($clicked -and $importOK)
        Check 'GuideTransfer.RoundTrip.UsesActualExport' ($exported -and $importOK -and $inputPath -ceq $export)
        Check 'GuideTransfer.Import.NewLocalIdentityAndRetainedOrigin' ($importOK -and $imported.ActionPathId -cne $source.ActionPathId -and $imported.RecordId -cne $source.RecordId -and $imported.Version -eq 1 -and $imported.PreviousRecordId -ceq '' -and $imported.PreviousSha256 -ceq '' -and $imported.WarehouseId -ceq $Other.Warehouse -and $imported.OriginWarehouseId -ceq $Fixture.Warehouse)
        $originOK=$false;$contentOK=$false
        if($importOK){
            $origin=$imported.TransferOrigin
            $originOK=(@($origin.PSObject.Properties.Name|Sort-Object) -join '|') -ceq 'ActionPathId|ContentSha256|RecordId|SourceWarehouseId|TransferId|TransferSha256|Version' -and $origin.TransferId -ceq $inputPackage.TransferId -and $origin.TransferSha256 -ceq $inputPackage.ContentSha256 -and $origin.SourceWarehouseId -ceq $Fixture.Warehouse -and $origin.ActionPathId -ceq $source.ActionPathId -and $origin.Version -eq 3 -and $origin.RecordId -ceq $source.RecordId -and $origin.ContentSha256 -ceq $source.ContentSha256
            $contentOK=$true
            foreach($field in @('Name','Tags','Instructions','Steps','Observations','SourceRun','ExpectedConclusion')){if((TransferJson $imported.$field) -cne (TransferJson $source.$field)){$contentOK=$false}}
            [void](BoundControl 'btnRefreshGuides' 'Click');BoundSelect $imported
        }
        Check 'GuideTransfer.Import.ExactTransferProvenance' $originOK
        Check 'GuideTransfer.Import.PreservesLongTextStepsAndOriginalObservations' $contentOK
        $labels=BoundControl '' 'Labels'
        Check 'GuideTransfer.Import.VisibleOriginAndLongInstructions' ($importOK -and $labels.Contains('Imported origin evidence; not locally observed') -and $labels.Contains([string]$Fixture.Warehouse) -and $labels.Contains([string]$source.ContentSha256) -and (BoundControl 'txtPublishedInstructions' 'Text').Contains($long))
        Check 'GuideTransfer.Import.DoesNotSelectOrInventLocalRun' ($importOK -and (BoundControl 'btnUseGuideForRun' 'State') -ceq 'True|False' -and (BoundPins (Split-Path $destinationRoot)).Count -eq 1)
        Check 'GuideTransfer.Import.SourceFileAndPriorWarehouseBytesPreserved' ((Get-FileHash -LiteralPath $inputPath).Hash -ceq $inputHash -and (BoundSame $sourcePins (TransferPins $Fixture.Root)) -and (TransferRetained $otherPins (TransferPins $Other.Root)))
        if($importOK){
            Test-GuideTransferEditAndRepeat $Fixture $Other $source $imported $inputPath $files $guideFields $fields
        }
        . (Join-Path $PSScriptRoot 'Slice4beGuideTransferPermissions.ps1')
        Test-GuideTransferPermissions $Fixture $Other $source $inputPath $files
        Check 'GuideTransfer.OriginalActivityRetained' ((TransferRetained $sourceActivity (BoundPins (Join-Path $Fixture.Root 'Training/Activity'))) -and (TransferRetained $otherActivity (BoundPins (Join-Path $Other.Root 'Training/Activity'))))
        TransferOpen $Fixture;BoundSelect $Guide
        Check 'GuideTransfer.NativeEarlierVersionStillReadable' ((BoundControl 'lblPublishedGuideSource' 'Label').Contains([string]$Guide.ContentSha256))
        TransferSetFile (Join-Path $files 'lost-context.json') $true
        $lostClick=(BoundControl 'btnExportGuide' 'Click') -ceq 'DELIVERED'
        Check 'GuideTransfer.Export.ContextLossAfterSelectionRefused' ($lostClick -and -not (Test-Path (Join-Path $files 'lost-context.json')) -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match '(?i)(changed|unavailable|signed|requires)')
        TransferSetFile '';TransferOpen $Other
        TransferImportAttempt 'ContextLostAfterSelection' $inputPath $Fixture $Other $true
        Check 'GuideTransfer.NoBusinessPublication' ((Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publicationBefore)
    } finally {TransferSetFile '';CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
