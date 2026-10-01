# D18 catalog25: invoke the packaged operator handler; observe real journal files.
# Owner facts are checked independently of tracking and of displayed status text.
function Install-ProductionCheckInActivityProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub CheckActivityUnselectForTest()
    mLoading = True
    mCmbRunProcess.ListIndex = 0: mCmbTreeRunProcess.ListIndex = 0
    mLoading = False
End Sub
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub CheckActivityUnselect()
    mForm.CheckActivityUnselectForTest
End Sub
'@)
}

function Test-ProductionCheckInActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function Pair([string[]]$Before,[string]$Outcome,[string]$Case){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$safe=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq 'PRODUCTION_RUN_CHECK_IN' -and $r.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $r.UserId -ceq 'config-producer' -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 25
            $safe=$safe -and @($r.SourceEventRefs).Count -eq 0
        }
        foreach($value in $raw){
            foreach($secret in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','DEMO-RAW-BLACK-TEA')+$keys){
                $encoded=ConvertTo-Json -InputObject $secret -Compress
                if($value.Contains($secret) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$safe=$false}
            }
            $match=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
        }
        if($paired){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId -and $activityIds.Add([string]$first[0].ActivityId)
            $severity=if($Outcome -ceq 'REJECTED'){'Warning'}elseif($Outcome -ceq 'FAILED'){'Error'}else{'Info'}
            $effect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            $facts=$first[0].DataEffect -ceq 'Unknown' -and $first[0].Severity -ceq 'Info' -and $last[0].DataEffect -ceq $effect -and $last[0].Severity -ceq $severity -and $first[0].EventCode -ceq 'PRODUCTION_RUN_CHECK_IN_REQUESTED' -and $last[0].EventCode -ceq ('PRODUCTION_RUN_CHECK_IN_'+$Outcome)
            $attempt=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
            $terminal=-not $attempt -and $completed -eq ($Outcome -ceq 'STAGED')
        }
        Check ($Case+'.AttemptAndOutcome') $paired
        Check ($Case+'.ExactCapturedContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $safe
        Check ($Case+'.Integrity') $integrity
        Check ($Case+'.DistinctLinkedAttempt') $linked
        Check ($Case+'.OwnerOutcomeFacts') $facts
        Check ($Case+'.ExactCommandTerminal') $terminal
    }
    $canary='CHECKACT'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{}
    $activityIds=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin Seed unavailable; not product RED.'}
    if([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockPrepareForTest' @($Fixture.Warehouse)) -cne 'READY'){throw 'Real Receiving fixture unavailable; not product RED.'}
    $keys=([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockKeysForTest')).Split("`t")
    if($keys.Count -ne 2 -or $keys[0] -ceq $keys[1]){throw 'Two owner-generated exact identities required.'}
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Authorized tracking policy unavailable; not product RED.'}
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'check-in-activity.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        if(-not [bool](Probe 'ReadPrepare' @($canary))){throw 'Released definitions unavailable; not product RED.'}
        [void](Probe 'RunLocalRememberFixture')
        foreach($root in @($Fixture.Root,$Other.Root)){
            foreach($file in Get-ChildItem -LiteralPath $root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        }
        foreach($case in @(
            @{Name='ReusableSuccess';Mode='Selected';Outcome='STAGED';Checked=$true},
            @{Name='NoProcess';Mode='NoProcess';Outcome='REJECTED';Checked=$false},
            @{Name='Nested';Mode='Selected';Outcome='STAGED';Checked=$true}
        )){
            if(-not [bool](Probe 'CheckBaselineReusableStage' @($case.Mode))){throw 'Reusable fixture unavailable; not product RED.'}
            if($case.Name -ceq 'Nested'){[void](Probe 'CheckBaselineArmNested')}
            $before=Files;$label='CheckInActivity.'+$case.Name
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
            Check ($label+'.IndependentOwnerResult') ([bool](Probe 'CheckBaselineReusableResult' @($case.Checked)))
            Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
            Pair $before $case.Outcome $label
            if($case.Name -ceq 'Nested'){Check ($label+'.OnlyOneOwnerEntry') ([int](Probe 'CheckBaselineOwnerEntries') -eq 1)}
            if($case.Name -ceq 'ReusableSuccess'){
                [void](Probe 'CheckActivityUnselect')
                $state=[string](Probe 'RunLocalOwnerState');$before=Files
                Check 'CheckInActivity.PriorSuccess.RefusedHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
                Check 'CheckInActivity.PriorSuccess.CheckedStateRemainsTrue' ([bool](Probe 'CheckBaselineReusableResult' @($true)))
                Check 'CheckInActivity.PriorSuccess.OwnerStatePreserved' ([string](Probe 'RunLocalOwnerState') -ceq $state)
                Pair $before 'REJECTED' 'CheckInActivity.PriorSuccess'
            }
        }
        foreach($guard in @('Loading','Busy')){
            if(-not [bool](Probe 'CheckBaselineReusableStage' @('Selected'))){throw 'Guard fixture unavailable; not product RED.'}
            $before=Files;[void](Probe 'CheckBaselineResetOwnerEntries')
            Check ('CheckInActivity.'+$guard+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @($guard)))
            Check ('CheckInActivity.'+$guard+'.NoOwnerEntry') ([int](Probe 'CheckBaselineOwnerEntries') -eq 0)
            Check ('CheckInActivity.'+$guard+'.NoRecords') (@(Files|Where-Object{$_ -cnotin $before}).Count -eq 0)
            Check ('CheckInActivity.'+$guard+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
        }
        if(-not [bool](Probe 'RunWorksheetPrepare')){throw 'Worksheet fixture unavailable; not product RED.'}
        $firstKey=[string](Probe 'RunWorksheetKey')
        if($firstKey -cnotin $keys){throw 'Worksheet identity is not from the real Receiving fixture.'}
        $selectedKey=@($keys|Where-Object{$_ -cne $firstKey})[0]
        if(-not [bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))){throw 'Exact-key worksheet fixture unavailable; not product RED.'}
        $before=Files
        Check 'CheckInActivity.Worksheet.ActualHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
        foreach($fact in @('ReachedCheckRows','ExactSelectedKey','CustomValue','CustomFormula','PalettePreserved','HeadersPreserved','DisplayColumns')){
            Check ('CheckInActivity.Worksheet.'+$fact) ([bool](Probe 'CheckBaselineWorksheetFact' @($fact)))
        }
        Pair $before 'STAGED' 'CheckInActivity.Worksheet'
        if(-not [bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))){throw 'Write refusal fixture unavailable; not product RED.'}
        [void](Probe 'CheckBaselineWorksheetInvalid' @('MissingUsed'))
        $state=[string](Probe 'CheckBaselineWorksheetState');$before=Files
        Check 'CheckInActivity.MissingColumns.ActualHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
        Check 'CheckInActivity.MissingColumns.StagingPreserved' ([string](Probe 'CheckBaselineWorksheetState') -ceq $state)
        Pair $before 'FAILED' 'CheckInActivity.MissingColumns'
        Check 'CheckInActivity.CanonicalEntitiesPreserved' ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockSourcePreservedForTest'))
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
        Check 'CheckInActivity.SavedAuthorityPreserved' $same
        Check 'CheckInActivity.CapturedBookExtraValues' ($sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        Check 'CheckInActivity.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
    }finally{
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
