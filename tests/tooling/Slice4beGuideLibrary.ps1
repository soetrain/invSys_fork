# D18 reader tests consume versions saved by actual product controls.
# Disposable source corruption is restored byte-for-byte; no guide is fabricated.
function GuideReader([string]$Name,[string]$Action,[string]$Value='') {
    GuideSaveControl $Name $Action $Value 'frmActionPathLibrary'
}
function OpenGuideReader {
    GuideSaveControl 'btnPublishedGuides' 'Click' '' 'frmActionPaths'
}
function GuideVersionKey($Record) {
    [string]$Record.ActionPathId+'|'+[string]$Record.Version+'|'+[string]$Record.ContentSha256
}
function SelectGuideVersion($Record) {
    $values=@((GuideReader 'lstPublishedGuides' 'Values') -split "`n" | Where-Object {$_ -cne ''})
    $key=GuideVersionKey $Record
    $index=[Array]::IndexOf($values,$key)
    if($index -lt 0){return $false}
    (GuideReader 'lstPublishedGuides' 'Select' ([string]$index)) -ceq 'SELECTED'
}
function GuideReaderUnavailable {
    (GuideReader 'txtPublishedInstructions' 'Text') -cin @('','MISSING') -and
        (GuideReader 'txtPublishedObservations' 'Text') -cin @('','MISSING') -and
        (GuideReader 'lblPublishedGuideStatus' 'Label') -match '(?i)unavailable|incomplete|changed'
}
function Test-GuideOrdinaryReader($Record) {
    $opened=(OpenGuideReader) -ceq 'DELIVERED'
    $selected=SelectGuideVersion $Record
    Check 'GuideLibrary.ReaderNeedsNoMaintenanceCapability' ($opened -and $selected -and
        (GuideReader 'txtPublishedInstructions' 'Text').Contains([string]$Record.Name))
    [void](GuideReader 'btnCloseGuides' 'Click')
}
function Test-GuideLibrary($Fixture,$Other,$First,$Second) {
    $sourceBefore=SourcePins; $guidesBefore=GuidePins; $activityBefore=ActivityPins
    $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $warehouseBefore=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){$warehouseBefore[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    $authorityBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest')
    $publishBefore=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
    $present=(GuideSaveControl 'btnPublishedGuides' 'State' '' 'frmActionPaths') -ceq 'True|True'
    Check 'GuideLibrary.PublishedEntryAvailable' $present
    $opened=(OpenGuideReader) -ceq 'DELIVERED'
    Check 'GuideLibrary.ActualEntryOpensOwnedReader' ($opened -and (GuideReader '' 'Count') -ceq '1' -and (GuideReader '' 'Caption') -ceq 'Published guides')
    [void](OpenGuideReader)
    Check 'GuideLibrary.RepeatedEntryReusesReader' ($opened -and (GuideReader '' 'Count') -ceq '1')
    $expected=(GuideVersionKey $First)+"`n"+(GuideVersionKey $Second)+"`n"
    Check 'GuideLibrary.ListContainsBothExactVersionsInNumericOrder' ($opened -and (GuideReader 'lstPublishedGuides' 'Values') -ceq $expected)
    $selected=SelectGuideVersion $First
    $instructions=GuideReader 'txtPublishedInstructions' 'Text'
    $observations=GuideReader 'txtPublishedObservations' 'Text'
    $provenance=GuideReader 'lblPublishedGuideSource' 'Label'
    Check 'GuideLibrary.ExactFirstVersionSelected' ($selected -and (GuideReader 'lstPublishedGuides' 'Selected') -ceq (GuideVersionKey $First) -and $instructions.Contains([string]$First.Name) -and -not $instructions.Contains([string]$Second.Name))
    Check 'GuideLibrary.AuthoredTextAndStepOrderPreserved' ($selected -and $instructions.Contains([string]$First.Instructions) -and
        $instructions.Contains([string]$First.Steps[0].Instruction) -and $instructions.Contains([string]$First.Steps[1].Instruction) -and
        $instructions.IndexOf([string]$First.Steps[0].Instruction) -lt $instructions.IndexOf([string]$First.Steps[1].Instruction))
    $originalOrder=$selected; $position=-1
    foreach($record in $First.Observations){
        $next=$observations.IndexOf([string]$record.ActivityId,$position+1,[StringComparison]::Ordinal)
        if($next -lt 0){$originalOrder=$false;break};$position=$next
        if(-not $observations.Contains([string]$record.Caption) -or -not $observations.Contains([string]$record.OutcomeCode)){$originalOrder=$false}
    }
    Check 'GuideLibrary.OriginalObservationsAndCaptionsStayInOrder' $originalOrder
    Check 'GuideLibrary.AuthoredAndObservedProvenanceDistinct' ($selected -and $instructions -cmatch 'Authored instruction' -and
        $observations -cmatch 'Observed control' -and $observations -notmatch '(?i)conclusion observed' -and
        -not $observations.Contains([string]$First.Steps[0].Instruction))
    Check 'GuideLibrary.ExactGuideAndSourceProvenanceVisible' ($selected -and $provenance.Contains([string]$First.ActionPathId) -and
        $provenance.Contains([string]$First.ContentSha256) -and $provenance -match '(?i)version\s*:?\s*1\b' -and
        $provenance.Contains([string]$First.SourceRun.SequenceId))
    Check 'GuideLibrary.BothPanesAreReadOnly' ($opened -and (GuideReader 'txtPublishedInstructions' 'Locked') -ceq 'True' -and (GuideReader 'txtPublishedObservations' 'Locked') -ceq 'True')
    [void](GuideReader 'btnRefreshGuides' 'Click')
    Check 'GuideLibrary.RefreshRetainsExactVersionAndText' ($selected -and (GuideReader 'lstPublishedGuides' 'Selected') -ceq (GuideVersionKey $First) -and
        (GuideReader 'txtPublishedInstructions' 'Text') -ceq $instructions -and (GuideReader 'txtPublishedObservations' 'Text') -ceq $observations)
    $secondSelected=SelectGuideVersion $Second
    $revised=GuideReader 'txtPublishedInstructions' 'Text'
    Check 'GuideLibrary.SecondVersionUsesAuthoredReorderOnly' ($secondSelected -and $revised.Contains([string]$Second.Name) -and
        $revised.Contains([string]$Second.Steps[0].Instruction) -and $revised.Contains([string]$Second.Steps[1].Instruction) -and
        $revised.IndexOf([string]$Second.Steps[0].Instruction) -lt $revised.IndexOf([string]$Second.Steps[1].Instruction) -and
        (GuideReader 'txtPublishedObservations' 'Text') -ceq $observations)
    foreach($case in @(@('Name','revised',1),@('Tag','TRAINING',2),@('Id',[string]$First.ActionPathId,2),@('Absent','no guide matches this text',0))){
        [void](GuideReader 'txtGuideSearch' 'Write' ([string]$case[1]))
        Check ('GuideLibrary.Search.'+$case[0]) ($opened -and (GuideReader 'lstPublishedGuides' 'Rows') -ceq [string]$case[2])
    }
    [void](GuideReader 'txtGuideSearch' 'Write' '')
    [void](SelectGuideVersion $First)
    foreach($size in @('Minimum','Default','Larger','Restored')){
        Check ('GuideLibrary.Layout.'+$size) ($opened -and (GuideReader '' 'Fit' $size) -ceq 'True')
    }
    if($CaptureEvidence -and $opened){
        Initialize-SettingsCapture
        $handle=[InvSysSettingsCapture]::OwnedVisibleForm('Published guides',[IntPtr]$excel.Hwnd).ToInt64()
        CaptureOwnedFormEvidence 'Published guides' 'published-guide-reader.png' $handle
        Check 'GuideLibrary.VisibleCapture' $true
    }
    Check 'GuideLibrary.SearchSelectionRefreshAndLayoutPreserveConfigBytes' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore)
    $warehouseAfter=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){$warehouseAfter[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    Check 'GuideLibrary.ReadOnlyActionsPreserveEveryWarehouseFile' (SamePins $warehouseBefore $warehouseAfter)
    Check 'GuideLibrary.ReadOnlyActionsAvoidShippingAuthority' ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest') -eq $authorityBefore)
    try {
        SaveGuideVisibility $false
        [void](GuideReader 'btnRefreshGuides' 'Click')
        [void](SelectGuideVersion $First)
        $hidden=GuideReader 'txtPublishedObservations' 'Text'
        $hiddenInstructions=GuideReader 'txtPublishedInstructions' 'Text'
        Check 'GuideLibrary.CurrentPolicyWithholdsRetainedEvidenceAndStepInstructions' ($opened -and
            -not $hidden.Contains([string]$First.Observations[0].ActivityId) -and
            -not $hiddenInstructions.Contains([string]$First.Steps[0].Instruction) -and
            ((GuideReader 'lblPublishedGuideStatus' 'Label')+' '+$hidden+' '+$hiddenInstructions) -cmatch 'Hidden by policy')
    } finally {SaveGuideVisibility $true}
    [void](GuideReader 'btnRefreshGuides' 'Click')
    [void](SelectGuideVersion $First)
    $firstPath=[IO.Path]::GetFullPath((Join-Path $guideRoot ($First.ActionPathId+'.1.json')))
    $held=$firstPath+'.withheld'
    $safeRoot=[IO.Path]::GetFullPath($guideRoot)+[IO.Path]::DirectorySeparatorChar
    foreach($path in @($firstPath,$held)){if(-not $path.StartsWith($safeRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Reader corruption fixture escapes generated guide root.'}}
    if(-not (Test-Path -LiteralPath $firstPath -PathType Leaf) -or (Test-Path -LiteralPath $held)){throw 'Reader corruption fixture unavailable.'}
    $original=[IO.File]::ReadAllBytes($firstPath)
    try {
        $changed=[byte[]]$original.Clone();$changed[0]=[byte][char]'['
        [IO.File]::WriteAllBytes($firstPath,$changed)
        $corruptSelected=SelectGuideVersion $First
        Check 'GuideLibrary.CorruptSelectedVersionClearsRetainedContent' ($opened -and $corruptSelected -and (GuideReaderUnavailable))
        $dependentSelected=SelectGuideVersion $Second
        Check 'GuideLibrary.CorruptPredecessorRejectsOtherwiseValidRevision' ($opened -and $dependentSelected -and (GuideReaderUnavailable))
    } finally {[IO.File]::WriteAllBytes($firstPath,$original)}
    [void](GuideReader 'btnRefreshGuides' 'Click')
    [void](SelectGuideVersion $Second)
    Move-Item -LiteralPath $firstPath -Destination $held
    try {
        $missingSelected=SelectGuideVersion $Second
        Check 'GuideLibrary.MissingPredecessorIsNotRepairedOrRebased' ($opened -and $missingSelected -and (GuideReaderUnavailable) -and -not (Test-Path -LiteralPath $firstPath))
    } finally {
        if(Test-Path -LiteralPath $firstPath){throw 'Guide reader recreated missing predecessor.'}
        Move-Item -LiteralPath $held -Destination $firstPath
    }
    [void](GuideReader 'btnRefreshGuides' 'Click')
    $restored=SelectGuideVersion $First
    Check 'GuideLibrary.RestoredExactBytesReadAgain' ($restored -and (GuideReader 'txtPublishedInstructions' 'Text') -ceq $instructions)
    [void](GuideReader 'btnCloseGuides' 'Click')
    Check 'GuideLibrary.CloseDisposesReader' ($opened -and (GuideReader '' 'Count') -ceq '0')
    [void](OpenGuideReader)
    [void](SelectGuideVersion $First)
    $otherGuideRoot=Join-Path $Other.Root ('Training/ActionPaths/'+$Other.Warehouse+'/Guides')
    if(Test-Path -LiteralPath $otherGuideRoot){throw 'Other warehouse guide-read fixture must start absent.'}
    SelectTarget $Other 'config-admin'
    [void](GuideReader 'btnRefreshGuides' 'Click')
    Check 'GuideLibrary.TargetChangeClearsRetainedContentWithoutRetargeting' ($opened -and
        (GuideReader 'txtPublishedInstructions' 'Text') -cin @('','MISSING') -and
        (GuideReader 'txtPublishedObservations' 'Text') -cin @('','MISSING') -and
        -not (Test-Path -LiteralPath $otherGuideRoot))
    CloseRecordingViewer
    OpenRecordingViewer
    [void](GuideSaveLibrary 'Open')
    $emptyOpened=(OpenGuideReader) -ceq 'DELIVERED'
    Check 'GuideLibrary.EmptyWarehouseReadCreatesNoTrainingFolder' ($emptyOpened -and
        (GuideReader 'lstPublishedGuides' 'Rows') -ceq '0' -and -not (Test-Path -LiteralPath $otherGuideRoot))
    CloseRecordingViewer
    Check 'GuideLibrary.ViewerClosureDisposesReader' ($opened -and (GuideReader '' 'Count') -ceq '0')
    SelectTarget $Fixture 'config-admin'
    OpenRecordingViewer
    [void](GuideSaveLibrary 'Open')
    [void](GuideSaveLibrary 'Select' $source.ActionPathId)
    Check 'GuideLibrary.ReadsPreserveGuideSourceAndActivityBytes' ((SamePins $guidesBefore (GuidePins)) -and (SamePins $sourceBefore (SourcePins)) -and
        (PinsRetained $activityBefore) -and (ActivityPins).Count -eq $activityBefore.Count)
    # Policy changes above are explicit commands; assert no extra changes after their restoration.
    Check 'GuideLibrary.NoBusinessPublication' ([long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishBefore)
}
