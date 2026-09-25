# Supplemental pure evidence checks; the lifecycle path gate supplies actual owner/UI proof.
function Install-DesignsSourceEvidenceProbe {
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function SourceEvidenceForTest(ByVal json As String) As String
    Dim fixture As Object, result As Object, sources As New Collection
    Set fixture = modTrainingJson.DecodeObject(json)
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "WarehouseId", "EVIDENCE_TEST"
    result.Add "TerminalSources", sources
    modEvaluationSources.Retain result, fixture("Terminal"), fixture("Publication")
    result.Add "SavedValid", modEvaluationSources.ValidSavedSources(sources, "EVIDENCE_TEST")
    SourceEvidenceForTest = modTrainingJson.EncodeObject(result)
End Function
Public Function SavedSourcesForTest(ByVal json As String) As Boolean
    Dim wrapper As Object
    Set wrapper = modTrainingJson.DecodeObject(json)
    SavedSourcesForTest = modEvaluationSources.ValidSavedSources(wrapper("Sources"), "EVIDENCE_TEST")
End Function
'@)
}

function Test-DesignsSourceEvidence {
    $cases=@('CompleteText','CompleteInteger','CompleteUnknown','MissingSubmitted','MissingUnknown','MissingOmittedGroups','MissingOmittedLines','MissingUnavailable','MissingCoverage','PresentUnavailable','WrongGroupKind','WrongEventId','WrongWarehouse','MissingWarehouse','MissingAppliedAt','EmptyAppliedAt','ZeroSeq','NegativeSeq','FractionalSeq','SpaceSeq','ExponentSeq','EmptySeq','MissingSeq','BooleanSeq','MismatchOutcome','MissingOutcome','EmptyGroup','SecondLineInvalid','InventoryComplete','InventoryNoWarehouse','InventoryMissingKey')
    foreach($name in $cases){
        $reference=[ordered]@{WarehouseId='EVIDENCE_TEST';SourceKind='Designs';EventId='EVENT_TEST';SubmissionState='Submitted'}
        $lines=@([ordered]@{EventID='EVENT_TEST';WarehouseId='EVIDENCE_TEST';AppliedAtUTC='2026-09-25T00:00:00Z';AppliedSeq='1'},[ordered]@{EventID='EVENT_TEST';WarehouseId='EVIDENCE_TEST';AppliedAtUTC='2026-09-25T00:00:01Z';AppliedSeq='2'})
        $source=[ordered]@{Source='Designs';Availability='Available';OmittedGroups=0;OmittedLines=0}
        $group=[ordered]@{Source='Designs';SourceId='EVENT_TEST';SourceKind='Business event';Lines=$lines;Outcomes=@()}
        $publication=[ordered]@{Groups=@($group);Coverage=[ordered]@{Sources=@($source)}}
        $expected='Unavailable';$lineCount=0;$keys=@()
        if($name -cin @('CompleteText','CompleteInteger','CompleteUnknown','InventoryComplete','InventoryNoWarehouse')){$expected='Applied';$lineCount=2}
        if($name -ceq 'MissingSubmitted'){$expected='Awaiting'}
        if($name.StartsWith('Missing',[StringComparison]::Ordinal) -and $name -cnotin @('MissingWarehouse','MissingAppliedAt','MissingSeq','MissingOutcome')){$publication.Groups=@()}
        switch($name){
            'CompleteInteger' {$lines[0].AppliedSeq=1;$lines[1].AppliedSeq=2}
            'CompleteUnknown' {$reference.SubmissionState='Unknown'}
            'MissingUnknown' {$reference.SubmissionState='Unknown'}
            'MissingOmittedGroups' {$source.OmittedGroups=1}
            'MissingOmittedLines' {$source.OmittedLines=1}
            'MissingUnavailable' {$source.Availability='Unavailable'}
            'MissingCoverage' {$publication.Coverage.Sources=@()}
            'PresentUnavailable' {$source.Availability='Unavailable'}
            'WrongGroupKind' {$group.SourceKind='Activity'}
            'WrongEventId' {$lines[0].EventID='OTHER_EVENT'}
            'WrongWarehouse' {$lines[0].WarehouseId='OTHER_WAREHOUSE'}
            'MissingWarehouse' {$lines[0].Remove('WarehouseId')}
            'MissingAppliedAt' {$lines[0].Remove('AppliedAtUTC')}
            'EmptyAppliedAt' {$lines[0].AppliedAtUTC=''}
            'ZeroSeq' {$lines[0].AppliedSeq='0'}
            'NegativeSeq' {$lines[0].AppliedSeq='-1'}
            'FractionalSeq' {$lines[0].AppliedSeq='1.5'}
            'SpaceSeq' {$lines[0].AppliedSeq=' 1'}
            'ExponentSeq' {$lines[0].AppliedSeq='1E1'}
            'EmptySeq' {$lines[0].AppliedSeq=''}
            'MissingSeq' {$lines[0].Remove('AppliedSeq')}
            'BooleanSeq' {$lines[0].AppliedSeq=$true}
            'EmptyGroup' {$group.Lines=@()}
            'SecondLineInvalid' {$lines[1].AppliedSeq='0'}
        }
        if($name.StartsWith('Inventory',[StringComparison]::Ordinal)){
            $reference.SourceKind='Inventory';$group.Source='Inventory';$source.Source='Inventory'
            $lines[0].System_Key='KEY_A';$lines[1].System_Key='KEY_B';$keys=@('KEY_A','KEY_B')
            if($name -ceq 'InventoryNoWarehouse'){$lines[0].Remove('WarehouseId');$lines[1].Remove('WarehouseId')}
            if($name -ceq 'InventoryMissingKey'){$lines[0].Remove('System_Key')}
        }
        $group.Outcomes=@($group.Lines|ForEach-Object {$_|ConvertTo-Json -Compress|ConvertFrom-Json})
        if($name -ceq 'MismatchOutcome'){$group.Outcomes[0].AppliedSeq='9'}
        if($name -ceq 'MissingOutcome'){$group.Outcomes=@($group.Outcomes[0])}
        $fixture=[ordered]@{Terminal=[ordered]@{SourceEventRefs=@($reference)};Publication=$publication}
        $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.SourceEvidenceForTest' @(($fixture|ConvertTo-Json -Depth 12 -Compress)))
        $result=$wire|ConvertFrom-Json;$evidence=$result.TerminalSources[0]
        Check ('DesignsEvidence.'+$name+'.Classification') ($evidence.OwnerStatus -ceq $expected -and $evidence.LineCount -eq $lineCount)
        $retained=$evidence.WarehouseId -ceq $reference.WarehouseId -and $evidence.EventId -ceq $reference.EventId -and $evidence.SourceKind -ceq $reference.SourceKind -and $evidence.SubmissionState -ceq $reference.SubmissionState
        if($expected -ceq 'Applied'){
            $body=[ordered]@{Lines=@($group.Lines)}|ConvertTo-Json -Depth 12 -Compress
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $retained=$retained -and $evidence.LinesSha256 -ceq $hash -and (@($evidence.SystemKeys) -join '|') -ceq ($keys -join '|')
        }else{$retained=$retained -and $evidence.LinesSha256 -ceq '' -and @($evidence.SystemKeys).Count -eq 0}
        Check ('DesignsEvidence.'+$name+'.ExactRetainedEvidence') $retained
        Check ('DesignsEvidence.'+$name+'.SavedEvidenceValid') ([bool]$result.SavedValid)
    }
    $applied=[ordered]@{WarehouseId='EVIDENCE_TEST';SourceKind='Designs';EventId='EVENT_TEST';SubmissionState='Submitted';OwnerStatus='Applied';LineCount=2;LinesSha256=('a'*64);SystemKeys=@()}
    foreach($name in @('ValidDesigns','InventedInventoryKey','MissingLines','MissingHash','CrossWarehouse')){
        $row=$applied|ConvertTo-Json -Compress|ConvertFrom-Json
        switch($name){'InventedInventoryKey'{$row.SystemKeys=@('INVENTED')} 'MissingLines'{$row.LineCount=0} 'MissingHash'{$row.LinesSha256=''} 'CrossWarehouse'{$row.WarehouseId='OTHER'}}
        $wire=[ordered]@{Sources=@($row)}|ConvertTo-Json -Depth 8 -Compress
        Check ('DesignsEvidence.Saved.'+$name) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.SavedSourcesForTest' @($wire)) -eq ($name -ceq 'ValidDesigns'))
    }
}
