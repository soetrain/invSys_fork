# D18 publication must exist on disk, not be simulated by reader-only clipping.
# The existing owning publication command executes unchanged. Only its source
# table read receives a disposable published-fixture copy; canonical rows stay
# untouched. Reports contain fixed checks/Booleans, never source rows or paths.
function Test-Slice4beViewerPublication($Fixture) {
    function PublicationSourceHash([string]$Path) {
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        $hash=[Security.Cryptography.SHA256]::Create()
        try {([BitConverter]::ToString($hash.ComputeHash($stream))).Replace('-','')} finally {$hash.Dispose();$stream.Dispose()}
    }
    SelectTarget $Fixture 'config-admin'
    $snapshot=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Inventory.xlsb')
    $source=Join-Path $runRoot 'viewer-publication-source.xlsb'
    Copy-Item -LiteralPath $snapshot -Destination $source
    $sourceHash=PublicationSourceHash $source
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
Public Function PublishViewerGroupsForTest() As Boolean
    Dim report As String
    PublishViewerGroupsForTest = GenerateInventorySnapshot("config-admin", "", Nothing, "", Nothing, report)
End Function
'@)
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Filter '*.xlsb') {
        if($file.FullName -ine $snapshot){$pins[$file.FullName]=PublicationSourceHash $file.FullName}
    }
    try {
        [void](Run 'invSys.Core.xlam' 'modWarehouseSync.SetViewerPublicationSourceForTest' @($sourceBook.Name))
        $started=[DateTimeOffset]::UtcNow
        $published=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishViewerGroupsForTest')
        $finished=[DateTimeOffset]::UtcNow
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
}
