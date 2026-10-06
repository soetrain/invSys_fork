# D18 catalog25: invoke the packaged operator handler; observe real journal files.
# Owner facts are checked independently of tracking and of displayed status text.
. (Join-Path $PSScriptRoot 'Slice4beProductionCheckInPolicy.ps1')
. (Join-Path $PSScriptRoot 'Slice4beProductionCheckInTerminal.ps1')
function Install-ProductionCheckInActivityProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub CheckActivityUnselectForTest()
    mLoading = True
    mCmbRunProcess.ListIndex = 0: mCmbTreeRunProcess.ListIndex = 0
    mLoading = False
End Sub
Public Function CheckActivitySuccessAndNoticeForTest() As Boolean
    CheckActivitySuccessAndNoticeForTest = InStr(1, mTxtStatus.Text, "Checked in ", vbBinaryCompare) = 1 And _
        InStr(1, mTxtStatus.Text, "Tracking unavailable:", vbBinaryCompare) > 1
End Function
Public Function CheckActivityPermissionTextForTest(ByVal withNotice As Boolean) As Boolean
    Dim expected As String
    expected = "Production permission changed. Reopen Production before continuing."
    If withNotice Then expected = expected & " Tracking unavailable: configuration could not be validated."
    CheckActivityPermissionTextForTest = (StrComp(mTxtStatus.Text, expected, vbBinaryCompare) = 0)
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub CheckActivityUnselect()
    mForm.CheckActivityUnselectForTest
End Sub
Public Function CheckActivitySuccessAndNotice() As Boolean
    CheckActivitySuccessAndNotice = mForm.CheckActivitySuccessAndNoticeForTest()
End Function
Public Function CheckActivityPermissionText(ByVal withNotice As Boolean) As Boolean
    CheckActivityPermissionText = mForm.CheckActivityPermissionTextForTest(withNotice)
End Function
'@)
    $core=$packages['invSys.Core.xlam'].VBProject
    $adapter=$core.VBComponents.Item('TestShippingCatalog').CodeModule
    $adapter.InsertLines(1,'Private mCheckPolicyUnavailable As Boolean, mCheckPolicyHits As Long')
    $adapter.AddFromString(@'
Public Sub CheckPolicyArm(ByVal armed As Boolean)
    mCheckPolicyUnavailable = armed
    If armed Then mCheckPolicyHits = 0
End Sub
Public Function CheckPolicyUnavailable() As Boolean
    CheckPolicyUnavailable = mCheckPolicyUnavailable
    If mCheckPolicyUnavailable Then mCheckPolicyHits = mCheckPolicyHits + 1
End Function
Public Function CheckPolicyHits() As Long
    CheckPolicyHits = mCheckPolicyHits
End Function
'@)
    $policy=$core.VBComponents.Item('modActivityPolicy').CodeModule
    $start=$policy.ProcStartLine('ReadPolicy',0);$end=$start+$policy.ProcCountLines('ReadPolicy',0)
    # The VBE normalizes identifier casing in the compiled project.
    $hits=@(for($i=$start;$i -lt $end;$i++){if($policy.Lines($i,1).Trim() -ieq 'If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then Exit Function'){$i}})
    if($hits.Count -ne 1){throw 'Tracking policy read boundary changed; not product RED.'}
    $policy.InsertLines($hits[0],@'
    If controlId = "PRODUCTION_RUN_CHECK_IN" Then
        If TestShippingCatalog.CheckPolicyUnavailable() Then Exit Function
    End If
'@)
    Install-ProductionCheckInPolicyProbe
    Install-ProductionCheckInTerminalProbe
}

function Test-ProductionCheckInActivity($Fixture,$Other) {
    Test-ProductionCheckInCatalog
    $catalogVersion=[int](Run 'invSys.Core.xlam' 'TestShippingCatalog.DeclaredCatalogVersionForTest')
    if($catalogVersion -lt 25){throw 'Check In catalog prerequisite unavailable; not product RED.'}
    [pscustomobject]@{DeclaredCatalogVersion=$catalogVersion;HistoricalCatalogVersion=25}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'check-in-catalog-version.json')
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function Pair([string[]]$Before,[string]$Outcome,[string]$Case,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$safe=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq 'PRODUCTION_RUN_CHECK_IN' -and $r.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq $catalogVersion
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
            $severity=if($Outcome -ceq 'REJECTED'){'Warning'}elseif($Outcome -ceq 'FAILED'){'Error'}elseif($Outcome -ceq 'DENIED'){'Blocked'}else{'Info'}
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
            @{Name='Nested';Mode='Selected';Outcome='STAGED';Checked=$true},
            @{Name='Insufficient';Mode='Insufficient';Outcome='FAILED';Checked=$false}
        )){
            if(-not [bool](Probe 'CheckBaselineReusableStage' @($case.Mode))){throw 'Reusable fixture unavailable; not product RED.'}
            if($case.Name -ceq 'Nested'){[void](Probe 'CheckBaselineArmNested')}
            if($CaptureEvidence -and $case.Name -ceq 'Insufficient'){[void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))}
            $before=Files;$label='CheckInActivity.'+$case.Name
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
            Check ($label+'.IndependentOwnerResult') ([bool](Probe 'CheckBaselineReusableResult' @($case.Checked)))
            Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
            Pair $before $case.Outcome $label
            if($CaptureEvidence -and $case.Name -ceq 'Insufficient'){CaptureOwnedFormByCaptionEvidence 'Production' 'check-in-insufficient-refusal.png'}
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
        if(-not [bool](Probe 'CheckBaselineReusableStage' @('Selected'))){throw 'Tracking failure owner fixture unavailable; not product RED.'}
        if($CaptureEvidence){[void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))}
        $before=Files
        [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckPolicyArm' @($true))
        try{
            Check 'CheckInActivity.PolicyUnavailable.ActualHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
            Check 'CheckInActivity.PolicyUnavailable.OwnerSucceeded' ([bool](Probe 'CheckBaselineReusableResult' @($true)))
            Check 'CheckInActivity.PolicyUnavailable.SuccessMessageAndNotice' ([bool](Probe 'CheckActivitySuccessAndNotice'))
            Check 'CheckInActivity.PolicyUnavailable.GuardsRestored' ([bool](Probe 'CheckBaselineGuards'))
            Check 'CheckInActivity.PolicyUnavailable.PolicyReadReached' ([int](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckPolicyHits') -eq 1)
            Check 'CheckInActivity.PolicyUnavailable.NoFalseRecords' (@(Files|Where-Object{$_ -cnotin $before}).Count -eq 0)
            if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Production' 'check-in-tracking-unavailable.png'}
        }finally{[void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckPolicyArm' @($false))}
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
        foreach($mode in @('MissingKey','UnknownKey','MissingIdentityHeader')){
            if(-not [bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))){throw 'Owner-refusal worksheet prerequisite unavailable; not product RED.'}
            [void](Probe 'CheckBaselineWorksheetInvalid' @($mode))
            $state=[string](Probe 'CheckBaselineWorksheetState');$before=Files;$label='CheckInActivity.Refusal.'+$mode
            [void](Probe 'CheckBaselineResetOwnerEntries')
            if($CaptureEvidence){[void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))}
            $decoy.Activate()
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
            Check ($label+'.OwnerEnteredOnce') ([int](Probe 'CheckBaselineOwnerEntries') -eq 1)
            Check ($label+'.StagingPreserved') ([string](Probe 'CheckBaselineWorksheetState') -ceq $state)
            Check ($label+'.NoSuccessStatus') ([bool](Probe 'CheckBaselineWorksheetFact' @('NoSuccessStatus')))
            Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
            Check ($label+'.CanonicalEntitiesPreserved') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockSourcePreservedForTest'))
            $outcome=if($mode -ceq 'MissingIdentityHeader'){'FAILED'}else{'REJECTED'}
            Pair $before $outcome $label
            if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Production' ('check-in-refusal-'+$mode.ToLowerInvariant()+'.png')}
        }
        SelectTarget $Fixture 'config-reader'
        [void](Probe 'CheckBaselineReopen' @($book.Name))
        foreach($mode in @('Reusable','Worksheet')){
            $ready=if($mode -ceq 'Reusable'){[bool](Probe 'CheckBaselineReusableStage' @('Selected'))}else{[bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))}
            if(-not $ready){throw 'Permission observation fixture unavailable; not product RED.'}
            if($CaptureEvidence -and $mode -ceq 'Reusable'){[void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))}
            $ownerBefore=[string](Probe 'RunLocalOwnerState')
            $projectionBefore=[string](Probe 'RunLocalState')+'|'+[string](Probe 'CheckBaselineWorksheetState')
            [void](Probe 'CheckBaselineResetOwnerEntries')
            $before=Files;$otherBefore=@(Get-Slice4beActivityFiles $Other);$label='CheckInActivity.Permission.'+$mode
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
            Check ($label+'.OwnerNotEntered') ([int](Probe 'CheckBaselineOwnerEntries') -eq 0)
            Check ($label+'.OwnerStatePreserved') ([string](Probe 'RunLocalOwnerState') -ceq $ownerBefore)
            Check ($label+'.ProjectionPreserved') (([string](Probe 'RunLocalState')+'|'+[string](Probe 'CheckBaselineWorksheetState')) -ceq $projectionBefore)
            Check ($label+'.ExactPermissionMessage') ([bool](Probe 'CheckActivityPermissionText' @($false)))
            Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
            Check ($label+'.NoRedirectedRecords') ((@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
            Pair $before 'DENIED' $label 'config-reader'
            if($CaptureEvidence -and $mode -ceq 'Reusable'){CaptureOwnedFormByCaptionEvidence 'Production' 'check-in-permission-denied.png'}
        }
        if(-not [bool](Probe 'CheckBaselineReusableStage' @('Selected'))){throw 'Permission/policy fault fixture unavailable; not product RED.'}
        if($CaptureEvidence){[void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))}
        $ownerBefore=[string](Probe 'RunLocalOwnerState')
        $projectionBefore=[string](Probe 'RunLocalState')+'|'+[string](Probe 'CheckBaselineWorksheetState')
        [void](Probe 'CheckBaselineResetOwnerEntries')
        $before=Files;$otherBefore=@(Get-Slice4beActivityFiles $Other);$label='CheckInActivity.Permission.PolicyUnavailable'
        [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckPolicyArm' @($true))
        try{
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
            Check ($label+'.OwnerNotEntered') ([int](Probe 'CheckBaselineOwnerEntries') -eq 0)
            Check ($label+'.OwnerStatePreserved') ([string](Probe 'RunLocalOwnerState') -ceq $ownerBefore)
            Check ($label+'.ProjectionPreserved') (([string](Probe 'RunLocalState')+'|'+[string](Probe 'CheckBaselineWorksheetState')) -ceq $projectionBefore)
            Check ($label+'.ExactPermissionMessageAndNotice') ([bool](Probe 'CheckActivityPermissionText' @($true)))
            Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
            Check ($label+'.PolicyReadReached') ([int](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckPolicyHits') -eq 1)
            Check ($label+'.NoFalseOrRedirectedRecords') (((Files) -join '|') -ceq ($before -join '|') -and (@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
            if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Production' 'check-in-permission-tracking-unavailable.png'}
        }finally{[void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckPolicyArm' @($false))}
        SelectTarget $Fixture 'config-producer'
        Test-ProductionCheckInPolicy $Fixture $Other $book $decoy $selectedKey $canary $keys
        Check 'CheckInActivity.CanonicalEntitiesPreserved' ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockSourcePreservedForTest'))
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
        Check 'CheckInActivity.SavedAuthorityPreserved' $same
        Check 'CheckInActivity.CapturedBookExtraValues' ($sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        Test-ProductionCheckInTerminal $Fixture $Other $book $decoy $selectedKey $canary $keys
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        Check 'CheckInActivity.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
    }finally{
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}

# Supplemental Core semantics; these never substitute for the actual-handler gate.
function Test-ProductionCheckInCatalog {
    $id='PRODUCTION_RUN_CHECK_IN';$label='CheckInCatalog'
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(24))).Split("`n")|Where-Object{$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(25))).Split("`n")|Where-Object{$_})
    Check ($label+'.Extends24Exactly') ($old.Count -eq 127 -and $new.Count -eq 128 -and @($new|Select-Object -Unique).Count -eq 128 -and ($new[0..126] -join '|') -ceq ($old -join '|') -and $new[-1] -ceq $id)
    $same=$true
    foreach($prior in $old){
        $before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($prior,24))
        $after=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($prior,25))
        $same=$same -and $before -cne '' -and $after -ceq $before
    }
    Check ($label+'.All127PriorDefinitionsPreserved') $same
    $excluded=$true
    foreach($version in 1..24){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
    Check ($label+'.ExcludedFromHistoricalVersions') $excluded
    $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,25))
    $record=if($wire){$wire|ConvertFrom-Json}else{$null}
    Check ($label+'.ExactControlContract') ($null -ne $record -and $record.ControlId -ceq $id -and $record.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $record.Class -ceq 'Command' -and $record.Role -ceq 'Production' -and $record.Caption -ceq 'Check In' -and $record.Surface -ceq 'Operations > Production > Production Run - List' -and $record.Capability -ceq 'PROD_POST' -and $record.CodePrefix -ceq ($id+'_'))
    $supported=[ordered]@{REQUESTED=@('Info','Unknown');DENIED=@('Blocked','Unchanged');REJECTED=@('Warning','Unchanged');FAILED=@('Error','Unknown');STAGED=@('Info','Unchanged')}
    foreach($code in @('REQUESTED','DENIED','REJECTED','FAILED','STAGED','REFRESHED','PRESENTED','SELECTED','CONFIRMED','PENDING','APPLIED','COMPLETED','VALIDATED','CANCELLED')){
        $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$code))
        $record=if($wire){$wire|ConvertFrom-Json}else{$null}
        $correct=if($supported.Contains($code)){$null -ne $record -and $record.EventCode -ceq ($id+'_'+$code) -and $record.OutcomeCode -ceq $code -and $record.Severity -ceq $supported[$code][0] -and $record.DataEffect -ceq $supported[$code][1] -and $record.UserMessage -ne '' -and $null -ne $record.PSObject.Properties['NextStep']}else{$wire -ceq ''}
        Check ($label+'.Outcome.'+$code) $correct
        $json=@{ControlId=$id;OwnerId='PRODUCTION_RUN_LOCAL';CatalogVersion=25;OutcomeCode=$code}|ConvertTo-Json -Compress
        Check ($label+'.Terminal.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @($json)) -eq ($code -ceq 'STAGED'))
        Check ($label+'.EmptyReferences.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,'[]')) -eq $supported.Contains($code))
    }
    foreach($code in $supported.Keys){
        foreach($kind in @('Inventory','Designs')){
            foreach($state in @('Submitted','Unknown')){
                $json=ConvertTo-Json -InputObject @(@{WarehouseId='CATALOG_TEST';SourceKind=$kind;EventId='Source_A';SubmissionState=$state}) -Compress
                Check ($label+'.RejectSource.'+$code+'.'+$kind+'.'+$state) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,$json)))
            }
        }
    }
    foreach($change in @('Owner','Catalog')){
        $record=@{ControlId=$id;OwnerId='PRODUCTION_RUN_LOCAL';CatalogVersion=25;OutcomeCode='STAGED'}
        if($change -ceq 'Owner'){$record.OwnerId='PRODUCTION_ASSIGNMENT'}else{$record.CatalogVersion=24}
        Check ($label+'.RejectWrong'+$change) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($record|ConvertTo-Json -Compress))))
    }
}
