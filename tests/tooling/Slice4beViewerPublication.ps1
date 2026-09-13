# D18 publication must exist on disk, not be simulated by reader-only clipping.
# The existing owning publication command executes unchanged. Only its source
# table read receives a disposable published-fixture copy; canonical rows stay
# untouched. Reports contain fixed checks/Booleans, never source rows or paths.
function Test-Slice4beViewerPublication($Fixture,$OtherFixture) {
    function PublicationSourceHash([string]$Path) {
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        $hash=[Security.Cryptography.SHA256]::Create()
        try {([BitConverter]::ToString($hash.ComputeHash($stream))).Replace('-','')} finally {$hash.Dispose();$stream.Dispose()}
    }
    SelectTarget $Fixture 'config-admin'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $deploy 'invSys.Core.xlam'))
    try {
        $reader=[IO.StreamReader]::new($archive.GetEntry('docProps/custom.xml').Open())
        try {[xml]$metadata=$reader.ReadToEnd()} finally {$reader.Dispose()}
        $version=@($metadata.Properties.property|Where-Object name -eq 'invSysPackageSetVersion')
        $build=@($metadata.Properties.property|Where-Object name -eq 'invSysBuildIdentity')
        if($version.Count -ne 1 -or $build.Count -ne 1){throw 'Publisher package identity metadata is missing.'}
        $publisherVersion=[string]$version[0].InnerText;$publisherBuild=[string]$build[0].InnerText
    } finally {$archive.Dispose()}
    $activityRoot=Join-Path $Fixture.Root ('Training\Activity\'+$Fixture.Warehouse)
    $beforeActivity=@()
    if(Test-Path -LiteralPath $activityRoot){$beforeActivity=@(Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File | Select-Object -ExpandProperty FullName)}
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        $saved=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','601'))
        if(-not $saved){throw 'Actual Settings Save did not prepare publication activity.'}
    } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    $activityRecords=@();$activityPins=@{}
    if(Test-Path -LiteralPath $activityRoot){
        foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File){
            $activityPins[$file.FullName]=PublicationSourceHash $file.FullName
            if($file.FullName -notin $beforeActivity){$activityRecords+=([IO.File]::ReadAllText($file.FullName)|ConvertFrom-Json)}
        }
    }
    $attempts=@($activityRecords|Where-Object {$_.ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and $_.OutcomeCode -ceq 'REQUESTED'})
    $resultsForAction=@($activityRecords|Where-Object {$_.ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and $_.OutcomeCode -cne 'REQUESTED'})
    $prepared=$attempts.Count -eq 1 -and $resultsForAction.Count -eq 1
    if($prepared){$prepared=$attempts[0].ActivityId -ceq $resultsForAction[0].ActivityId -and $attempts[0].RecordId -cne $resultsForAction[0].RecordId}
    Check 'ViewerPublication.RealSettingsHandlerPreparedActivityPair' $prepared
    if(-not $prepared){throw 'Publication activity fixture lacks the actual correlated attempt/result pair.'}
    . (Join-Path $PSScriptRoot 'Slice4beDesignsPublicationSource.ps1')
    Test-Slice4beDesignsPublicationSource $Fixture $OtherFixture
    $snapshot=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Inventory.xlsb')
    $source=Join-Path $runRoot 'viewer-publication-source.xlsb'
    Copy-Item -LiteralPath $snapshot -Destination $source
    $sourceHash=PublicationSourceHash $source
    . (Join-Path $PSScriptRoot 'Slice4beShippingPublicationFixture.ps1')
    $shippingFixture=$null
    New-Slice4beShippingPublicationFixture $Fixture ([ref]$shippingFixture)
    try {
    # Real Boxing/Shipping handlers may publish inventory and append activity.
    # The volume source is already copied; pin all activity after owner setup.
    $allActivity=@();$activityPins=@{}
    foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File){
        $activityPins[$file.FullName]=PublicationSourceHash $file.FullName
        $allActivity+=([IO.File]::ReadAllText($file.FullName)|ConvertFrom-Json)
    }
    $activityIds=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($record in $allActivity){[void]$activityIds.Add([string]$record.ActivityId)}
    $activityGroupCount=$activityIds.Count
    $sourceBook=$excel.Workbooks.Open($source,0,$true)
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modWarehouseSync').CodeModule
    $first=$core.ProcBodyLine('WriteSnapshotEventRows',0)
    $last=$core.ProcStartLine('WriteSnapshotEventRows',0)+$core.ProcCountLines('WriteSnapshotEventRows',0)-1
    $hooked=$false
    for($line=$first;$line -le $last;$line++) {
        if($core.Lines($line,1).Trim() -eq 'Set sourceTable = FindListObjectByNameSync(inventoryWb, "tblInventoryLog")') {
            $core.InsertLines($line+1,@'
    If ViewerPublicationSourceForTest <> "" Then
        Set sourceTable = Application.Workbooks(ViewerPublicationSourceForTest).Worksheets("InventoryEvents").ListObjects("tblInventoryEvents")
        ViewerPublicationReadsForTest = ViewerPublicationReadsForTest + 1
        ViewerPublicationReadOnlyForTest = inventoryWb.ReadOnly
    End If
'@)
            $hooked=$true;break
        }
    }
    if(-not $hooked){throw 'Publication source-fixture hook was not installed.'}
    $core.AddFromString(@'
Private ViewerPublicationSourceForTest As String
Private ViewerPublicationReadsForTest As Long
Private ViewerPublicationReadOnlyForTest As Boolean
Public Sub SetViewerPublicationSourceForTest(ByVal workbookName As String)
    ViewerPublicationSourceForTest = workbookName
    ViewerPublicationReadsForTest = 0
End Sub
Public Function ViewerPublicationSourceReadForTest() As Boolean
    ViewerPublicationSourceReadForTest = (ViewerPublicationReadsForTest = 1)
End Function
Public Function ViewerPublicationSourceReadOnlyForTest() As Boolean
    ViewerPublicationSourceReadOnlyForTest = ViewerPublicationReadOnlyForTest
End Function
'@)
    $admin=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdminConsole').CodeModule
    $admin.AddFromString(@'
Private PublicationReportForTest As String
Public Function PublishViewerGroupsForTest() As Boolean
    PublicationReportForTest = ""
    PublishViewerGroupsForTest = GenerateInventorySnapshot("config-admin", "", Nothing, "", Nothing, PublicationReportForTest)
End Function
Public Function PublicationNoticeForTest(ByVal expected As String) As Boolean
    PublicationNoticeForTest = (InStr(1, PublicationReportForTest, vbCrLf & expected, vbBinaryCompare) > 0)
End Function
'@)
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Filter '*.xlsb') {
        if($file.FullName -ine $snapshot){$pins[$file.FullName]=PublicationSourceHash $file.FullName}
    }
    try {
        SelectTarget $Fixture 'config-admin'
        [void](Run 'invSys.Core.xlam' 'modWarehouseSync.SetViewerPublicationSourceForTest' @($sourceBook.Name))
        $started=[DateTimeOffset]::UtcNow
        $published=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishViewerGroupsForTest')
        $finished=[DateTimeOffset]::UtcNow
        [pscustomobject]@{CommandSeconds=($finished-$started).TotalSeconds} | ConvertTo-Json | Set-Content (Join-Path $reportRoot 'publication-timing.json')
        if(-not $published){throw 'Existing Admin inventory-snapshot command did not complete the publication fixture.'}
        Check 'ViewerPublication.ActualAdminSnapshotCommandSucceeded' $published
        if(-not [bool](Run 'invSys.Core.xlam' 'modWarehouseSync.ViewerPublicationSourceReadForTest')){throw 'Publication did not consume the controlled source fixture.'}
        Check 'ViewerPublication.TransientInventorySourceReadOnly' ([bool](Run 'invSys.Core.xlam' 'modWarehouseSync.ViewerPublicationSourceReadOnlyForTest'))
    } finally {
        [void](Run 'invSys.Core.xlam' 'modWarehouseSync.SetViewerPublicationSourceForTest' @(''))
        $sourceBook.Close($false)
    }
    $inventoryPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $sourceOpen=$false
    foreach($book in $excel.Workbooks){if($book.FullName -ieq $inventoryPath){$sourceOpen=$true}}
    Check 'ViewerPublication.TransientInventorySourceReleased' (-not $sourceOpen)
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $artifact=$null;$raw=''
    if(Test-Path -LiteralPath $path){
        $raw=[IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
        try {$artifact=$raw | ConvertFrom-Json} catch {$artifact=$null}
    }
    Check 'ViewerPublication.SeparateEventsArtifactCreated' (Test-Path -LiteralPath $path)
    $valid=$null -ne $artifact
    if($valid){$valid=($artifact.SchemaVersion -eq 1 -and $artifact.WarehouseId -ceq $Fixture.Warehouse -and $artifact.PublicationId -match '^[0-9a-f-]{36}$')}
    Check 'ViewerPublication.SchemaWarehouseAndPublicationIdentity' $valid
    $utcValid=$false
    if($null -ne $artifact){
        $stamp=[DateTimeOffset]::MinValue
        $utcValid=$artifact.PublishedAtUTC -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$'
        if($utcValid){$utcValid=[DateTimeOffset]::TryParse($artifact.PublishedAtUTC,[ref]$stamp) -and $stamp -ge $started.AddSeconds(-1) -and $stamp -le $finished.AddSeconds(1)}
    }
    Check 'ViewerPublication.VerifiedUtcPublicationTime' $utcValid
    $groups=@()
    if($null -ne $artifact){$groups=@($artifact.Groups)}
    Check 'ViewerPublication.Persists5000CompleteDurableGroups' ($groups.Count -eq 5000)
    $boundary=@($groups | Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq 'EVT-GROUP-00100'})
    $lines=@()
    if($boundary.Count -eq 1){$lines=@($boundary[0].Lines)}
    $retained=$lines.Count -eq 3 -and @($lines | Where-Object {$_.System_Key -ceq 'SYS-GROUP-A'}).Count -eq 2 -and @($lines | Where-Object {$_.System_Key -ceq 'SYS-GROUP-B'}).Count -eq 1
    Check 'ViewerPublication.BoundaryRetainsRepeatedExactKeysAndUnlikeUnits' ($retained -and @($lines | Where-Object {$_.Uom -ceq 'LB'}).Count -eq 1 -and @($lines | Where-Object {$_.Uom -ceq 'EA'}).Count -eq 2)
    Check 'ViewerPublication.NewestRetainedAndOldestOmitted' (@($groups | Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq 'EVT-GROUP-00001'}).Count -eq 1 -and @($groups | Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq 'EVT-GROUP-05001'}).Count -eq 0)
    $coverage=$false
    if($null -ne $artifact){$coverage=$null -ne $artifact.Coverage -and $null -ne $artifact.Coverage.Sources -and @($artifact.Coverage.Sources).Count -ge 4}
    Check 'ViewerPublication.ExpectedSourceCoverageIsExplicit' $coverage
    $activityGroup=@($groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $attempts[0].ActivityId})
    $publishedRecords=@();$publishedOutcomes=@()
    if($activityGroup.Count -eq 1){$publishedRecords=@($activityGroup[0].Lines);$publishedOutcomes=@($activityGroup[0].Outcomes)}
    $exactPair=$publishedRecords.Count -eq 2
    foreach($record in $activityRecords){$exactPair=$exactPair -and @($publishedRecords|Where-Object {$_.RecordId -ceq $record.RecordId -and $_.ActivityId -ceq $record.ActivityId -and $_.OccurredAtUTC -ceq $record.OccurredAtUTC}).Count -eq 1}
    Check 'ViewerPublication.ActivityGroupRetainsEveryExactRecord' ($activityGroup.Count -eq 1 -and $exactPair)
    $truthful=$publishedOutcomes.Count -eq 1
    if($truthful){$truthful=$publishedOutcomes[0].RecordId -ceq $resultsForAction[0].RecordId -and $publishedOutcomes[0].OutcomeCode -ceq $resultsForAction[0].OutcomeCode -and $publishedOutcomes[0].DataEffect -ceq $resultsForAction[0].DataEffect}
    Check 'ViewerPublication.ActivityOutcomeIsObservedResult' $truthful
    $chronology=$activityGroup.Count -eq 1
    if($chronology){$chronology=$activityGroup[0].RecordedAt -ceq $attempts[0].OccurredAtUTC -and $activityGroup[0].SourceKind -ceq 'User activity'}
    Check 'ViewerPublication.ActivityGroupUsesEarliestRecordedTime' $chronology
    $sources=@();if($null -ne $artifact){$sources=@($artifact.Coverage.Sources)}
    $sources | Select-Object Source,Availability,Scope,Explanation,AvailableGroups,IncludedGroups,OmittedGroups,AvailableLines,IncludedLines,OmittedLines | ConvertTo-Json | Set-Content (Join-Path $reportRoot 'publication-coverage.json')
    $named=$sources.Count -eq 5
    foreach($name in @('Inventory','Designs','Activity','ShippingBOM','ShippingHolds')){$named=$named -and @($sources|Where-Object {$_.Source -ceq $name -and $_.Availability -ne '' -and $_.Scope -ne ''}).Count -eq 1}
    Check 'ViewerPublication.EveryExpectedSourceHasNamedCoverage' $named
    $counts=$true
    $inventoryIncluded=5000-2-$activityGroupCount
    $inventoryOmitted=5001-$inventoryIncluded
    foreach($expect in @(@('Inventory',5001,$inventoryIncluded,$inventoryOmitted,5003,($inventoryIncluded+2),$inventoryOmitted),@('Activity',$activityGroupCount,$activityGroupCount,0,$allActivity.Count,$allActivity.Count,0),@('Designs',2,2,0,2,2,0))){
        $entry=@($sources|Where-Object {$_.Source -ceq $expect[0]})
        if($entry.Count -ne 1){$counts=$false;continue}
        $i=1;foreach($field in @('AvailableGroups','IncludedGroups','OmittedGroups','AvailableLines','IncludedLines','OmittedLines')){$counts=$counts -and $entry[0].$field -eq $expect[$i];$i++}
    }
    Check 'ViewerPublication.MixedSourceCountsReconcileGlobalBound' $counts
    Test-Slice4beShippingPublication $artifact $shippingFixture
    $provenance=$null -ne $artifact
    if($provenance){$provenance=$artifact.PackageSetVersion -ceq $publisherVersion -and $artifact.BuildIdentity -ceq $publisherBuild -and $artifact.PolicyVersion -eq $attempts[0].PolicyVersion}
    Check 'ViewerPublication.PackageAndPolicyProvenanceMatchesOwner' $provenance
    $hashValid=$false
    $marker=$raw.LastIndexOf(',"ContentSha256":"',[StringComparison]::Ordinal)
    if($marker -gt 0 -and $raw -match ',"ContentSha256":"([0-9a-f]{64})"}$'){
        $expected=$Matches[1]
        $body=$raw.Substring(0,$marker)+'}'
        $hash=[Security.Cryptography.SHA256]::Create()
        try {$actual=([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($body)))).Replace('-','').ToLowerInvariant()} finally {$hash.Dispose()}
        $hashValid=$actual -ceq $expected
    }
    Check 'ViewerPublication.ExactBodyIntegrityMatches' $hashValid
    Check 'ViewerPublication.UnknownSourceColumnExcluded' ($raw.Length -gt 0 -and -not $raw.Contains('DO-NOT-DISPLAY') -and -not $raw.Contains('User group sentinel'))
    # Fault only the generated destination. The same public Admin command must
    # retain its inventory success while reporting the independent Events result.
    if($null -ne $artifact) {
        $priorHash=PublicationSourceHash $path
        $destinationLock=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        try {
            Check 'ViewerPublication.LockedDestinationKeepsInventorySuccess' ([bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishViewerGroupsForTest'))
            Check 'ViewerPublication.LockedDestinationReportsEventsFailure' ([bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublicationNoticeForTest' @('Events publication failed; the prior complete Events artifact was retained.')))
            Check 'ViewerPublication.LockedDestinationRetainsExactPriorBytes' ((PublicationSourceHash $path) -ceq $priorHash)
            Check 'ViewerPublication.FailedReplacementLeavesNoPendingFile' (@(Get-ChildItem -LiteralPath $Fixture.Root -Filter '.events-*.pending' -File).Count -eq 0)
        } finally {$destinationLock.Dispose()}
        Check 'ViewerPublication.ReplacementAfterUnlockSucceeds' ([bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishViewerGroupsForTest'))
        Check 'ViewerPublication.ReplacementReportsEventsSuccess' ([bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublicationNoticeForTest' @('Events publication complete.')))
        $replacement=[IO.File]::ReadAllText($path)|ConvertFrom-Json
        Check 'ViewerPublication.SuccessfulReplacementHasNewIdentity' ($replacement.PublicationId -cne $artifact.PublicationId)
        . (Join-Path $PSScriptRoot 'Slice4bePublicationSourceFailures.ps1')
        Test-Slice4bePublicationSourceFailures $Fixture $path
    } else {
        foreach($name in @('LockedDestinationKeepsInventorySuccess','LockedDestinationReportsEventsFailure','LockedDestinationRetainsExactPriorBytes','FailedReplacementLeavesNoPendingFile','ReplacementAfterUnlockSucceeds','ReplacementReportsEventsSuccess','SuccessfulReplacementHasNewIdentity',
            'MissingBom.InventoryCommandSucceeds','MissingBom.CountsUnavailableNotZero','MissingBom.NoInventedCurrentState','MissingBom.NotRecreated',
            'DirtyBom.InventoryCommandSucceeds','DirtyBom.CountsUnavailableNotZero','DirtyBom.NoInventedCurrentState','DirtyBom.BorrowedWithoutSaveOrClose',
            'RestoredBom.InventoryCommandSucceeds','RestoredBom.AvailableWithOriginalBytes')){Check ('ViewerPublication.'+$name) $false}
    }
    $unchanged=(PublicationSourceHash $source) -ceq $sourceHash
    $sourceFacts=[Collections.Generic.List[object]]::new()
    $sourceFacts.Add([pscustomobject]@{Source='PublishedFixtureCopy';Unchanged=$unchanged})
    foreach($file in $pins.Keys){
        $same=(PublicationSourceHash $file) -ceq $pins[$file]
        if(-not $same){$unchanged=$false}
        $kind=([IO.Path]::GetFileName($file)).Substring($Fixture.Warehouse.Length+1)
        $sourceFacts.Add([pscustomobject]@{Source=$kind;Unchanged=$same})
    }
    $sourceFacts | ConvertTo-Json | Set-Content (Join-Path $reportRoot 'publication-source-preservation.json')
    Check 'ViewerPublication.SourceCopyAndCanonicalWorkbookBytesUnchanged' $unchanged
    $activityUnchanged=$activityPins.Count -eq @(Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File).Count
    foreach($path in $activityPins.Keys){$activityUnchanged=$activityUnchanged -and (PublicationSourceHash $path) -ceq $activityPins[$path]}
    Check 'ViewerPublication.PublicationDoesNotRewriteActivityRecords' $activityUnchanged
    } finally {Remove-Slice4beShippingPublicationFiles $shippingFixture.LocalFiles}
}
