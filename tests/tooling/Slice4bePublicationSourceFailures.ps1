# D18 faults apply only to the generated warehouse. The public Admin command
# must publish truthful unavailable coverage without recreating/saving sources.
function Test-Slice4bePublicationSourceFailures($Fixture,[string]$EventsPath) {
    function CheckUnavailable([string]$Label) {
        Check ('ViewerPublication.'+$Label+'.InventoryCommandSucceeds') ([bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishViewerGroupsForTest'))
        $publication=[IO.File]::ReadAllText($EventsPath)|ConvertFrom-Json
        $coverage=@($publication.Coverage.Sources|Where-Object Source -CEQ 'ShippingBOM')
        $unavailable=$coverage.Count -eq 1
        if($unavailable){
            $unavailable=$coverage[0].Availability -ceq 'Unavailable'
            foreach($field in @('AvailableGroups','IncludedGroups','OmittedGroups','AvailableLines','IncludedLines','OmittedLines')){$unavailable=$unavailable -and $coverage[0].$field -ceq 'Unavailable'}
        }
        Check ('ViewerPublication.'+$Label+'.CountsUnavailableNotZero') $unavailable
        Check ('ViewerPublication.'+$Label+'.NoInventedCurrentState') (@($publication.CurrentState|Where-Object Source -CEQ 'ShippingBOM').Count -eq 0)
    }
    if($Fixture.Warehouse -cnotmatch '^WHD5[A-F0-9]{6}$'){throw 'Source-failure fixture identity is not disposable.'}
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $bom=[IO.Path]::GetFullPath((Join-Path $root ($Fixture.Warehouse+'.invSys.Data.ShippingBOM.xlsb')))
    $held=$bom+'.publication-unavailable'
    if(-not $bom.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or -not $held.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $held)){throw 'Source-failure fixture paths are not isolated.'}
    $hash=PublicationSourceHash $bom
    Move-Item -LiteralPath $bom -Destination $held
    try {
        CheckUnavailable 'MissingBom'
        Check 'ViewerPublication.MissingBom.NotRecreated' (-not (Test-Path -LiteralPath $bom))
    } finally {
        if(Test-Path -LiteralPath $bom){throw 'Publication unexpectedly recreated the missing source; fixture backup retained.'}
        Move-Item -LiteralPath $held -Destination $bom
    }
    $book=$excel.Workbooks.Open($bom,0,$false)
    try {
        if(-not [bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublicationShippingDirtyForTest' @($book.Name))){throw 'Dirty-source fixture was not calibrated.'}
        CheckUnavailable 'DirtyBom'
        Check 'ViewerPublication.DirtyBom.BorrowedWithoutSaveOrClose' (-not $book.Saved -and @($excel.Workbooks|Where-Object FullName -IEQ $bom).Count -eq 1 -and (PublicationSourceHash $bom) -ceq $hash)
    } finally {$book.Close($false)}
    Check 'ViewerPublication.RestoredBom.InventoryCommandSucceeds' ([bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishViewerGroupsForTest'))
    $restored=[IO.File]::ReadAllText($EventsPath)|ConvertFrom-Json
    $coverage=@($restored.Coverage.Sources|Where-Object Source -CEQ 'ShippingBOM')
    Check 'ViewerPublication.RestoredBom.AvailableWithOriginalBytes' ($coverage.Count -eq 1 -and $coverage[0].Availability -ceq 'Available' -and $coverage[0].IncludedLines -eq 2 -and (PublicationSourceHash $bom) -ceq $hash)
}
