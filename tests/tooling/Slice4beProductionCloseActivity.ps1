# D13 actual dismissal, native teardown and public reopen. Fixture values remain
# in memory; reports contain only named checks, counts and boolean evidence.
function Test-ProductionCloseActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function InventoryPartHashes {
        Add-Type -AssemblyName System.IO.Compression
        $stream=[IO.File]::Open($inventoryPath,'Open','Read','ReadWrite')
        $archive=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Read,$false)
        $parts=@{}
        try{foreach($entry in $archive.Entries){$part=$entry.Open();try{$parts[$entry.FullName]=(Get-FileHash -InputStream $part).Hash}finally{$part.Dispose()}}}finally{$archive.Dispose()}
        return $parts
    }
    function AuthorityCheckpoint([string]$Stage){
        $index=0
        foreach($file in @($pins.Keys|Sort-Object)){
            $index++;$kind=switch -Regex ([IO.Path]::GetFileName($file)){'\.Config\.'{'Config';break} '\.Auth\.'{'Auth';break} '\.Inventory\.'{'Inventory';break} '\.Designs\.'{'Designs';break} default{'Other generated workbook'}}
            $actual=Hash $file
            $authorityTrace.Add([pscustomobject]@{Stage=$Stage;Index=$index;Kind=$kind;ExpectedHash=$pins[$file];ActualHash=$actual;Preserved=$actual -ceq $pins[$file]})
        }
        $authorityTrace|ConvertTo-Json -Depth 4|Set-Content (Join-Path $reportRoot 'close-authority-checkpoints.json')
        $parts=InventoryPartHashes
        $changed=@($parts.Keys|Where-Object{-not $inventoryParts.ContainsKey($_) -or $parts[$_] -cne $inventoryParts[$_]}|Sort-Object)
        $inventoryTrace.Add([pscustomobject]@{Stage=$Stage;ChangedParts=$changed;RemovedParts=@($inventoryParts.Keys|Where-Object{-not $parts.ContainsKey($_)}|Sort-Object)})
        $inventoryTrace|ConvertTo-Json -Depth 5|Set-Content (Join-Path $reportRoot 'close-inventory-part-checkpoints.json')
    }
    function SetPolicy([bool]$Enabled){
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ClosePolicyForTest' @($Enabled))){throw 'Authorized Close policy fixture unavailable; not product RED.'}
        SelectTarget $Fixture 'config-producer'
    }
    function Dismissed {
        ([long](Probe 'CloseLoadedFormsForTest') -eq 0 -and [InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -eq [IntPtr]::Zero)
    }
    function OpenPrivate {
        $decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'CloseShowForTest')
        if([long](Probe 'CloseLoadedFormsForTest') -ne 1){throw 'One visible fixture form required; not product RED.'}
    }
    function Pair([string[]]$Before,[string]$Label) {
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $rows=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$closed=@($rows|Where-Object OutcomeCode -CEQ 'CLOSED')
        $paired=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $closed.Count -eq 1
        $linked=$false;$facts=$paired;$safe=$paired;$integrity=$paired;$terminal=$false
        foreach($row in $rows){
            $facts=$facts -and $row.ControlId -ceq 'PRODUCTION_CLOSE' -and $row.OwnerId -ceq 'PRODUCTION_WORKFLOW' -and $row.UserId -ceq 'config-producer' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.StationId -ceq 'S1' -and $row.CatalogVersion -eq 21 -and @($row.SourceEventRefs).Count -eq 0
        }
        foreach($text in $raw){
            $decoded=($text|ConvertFrom-Json)|ConvertTo-Json -Depth 25 -Compress
            foreach($value in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','PayloadJson')){
                $encoded=ConvertTo-Json -InputObject $value -Compress
                if($decoded.Contains($value) -or $decoded.Contains($encoded.Substring(1,$encoded.Length-2))){$safe=$false}
            }
            $match=[regex]::Match($text,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$digest=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $digest -ceq $match.Groups[1].Value
        }
        if($paired){
            $linked=$attempt[0].ActivityId -cne '' -and $attempt[0].ActivityId -ceq $closed[0].ActivityId -and $attempt[0].RecordId -cne $closed[0].RecordId
            $facts=$facts -and $attempt[0].Severity -ceq 'Info' -and $attempt[0].DataEffect -ceq 'Unknown' -and $attempt[0].EventCode -ceq 'PRODUCTION_CLOSE_REQUESTED' -and $closed[0].Severity -ceq 'Info' -and $closed[0].DataEffect -ceq 'Unchanged' -and $closed[0].EventCode -ceq 'PRODUCTION_CLOSE_CLOSED' -and $closed[0].UserMessage -ceq 'Production form dismissed; no save, posting or Domain application is asserted.'
            $terminal=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CloseTerminalForTest' @(($attempt[0]|ConvertTo-Json -Depth 20 -Compress))) -and [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CloseTerminalForTest' @(($closed[0]|ConvertTo-Json -Depth 20 -Compress)))
        }
        Check ($Label+'.OneAttemptAndClosed') $paired
        Check ($Label+'.DistinctCorrelatedRecords') $linked
        Check ($Label+'.ExactContextAndDismissalFacts') $facts
        Check ($Label+'.NoEnteredDataOrSecrets') $safe
        Check ($Label+'.ContentIntegrity') $integrity
        Check ($Label+'.OnlyDismissalConcludes') $terminal
    }
    $canary='CLOSE'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$publicBook=$null;$pins=@{};$recordPins=@{}
    $authorityTrace=[Collections.Generic.List[object]]::new()
    $inventoryTrace=[Collections.Generic.List[object]]::new()
    $inventoryPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $inventoryParts=InventoryPartHashes
    SetPolicy $true
    try {
        $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(20))).Split([char]10)|Where-Object{$_})
        $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(21))).Split([char]10)|Where-Object{$_})
        Check 'ProductionClose.Catalog21Extends20' ($old.Count -eq 105 -and $new.Count -eq 106 -and @($new|Sort-Object -Unique).Count -eq 106 -and (@($new|Where-Object{$_ -cnotin $old}) -join '|') -ceq 'PRODUCTION_CLOSE')
        foreach($id in $old){$prior=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,20));Check ('ProductionClose.Catalog.Preserve.'+$id) ($prior -cne '' -and $prior -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,21)))}
        Check 'ProductionClose.Catalog20DoesNotContainClose' ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @('PRODUCTION_CLOSE',20)) -ceq '')
        foreach($code in @('CONFIRMED','APPLIED','VALIDATED','STAGED','COMPLETED','DENIED')){Check ('ProductionClose.Unsupported.'+$code) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @('PRODUCTION_CLOSE',$code)) -ceq '')}
        foreach($code in @('REQUESTED','CLOSED','FAILED')){
            $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @('PRODUCTION_CLOSE',$code))
            $value=if($wire){$wire|ConvertFrom-Json}else{$null}
            Check ('ProductionClose.Supported.'+$code) ($null -ne $value -and $value.OutcomeCode -ceq $code -and $value.EventCode -ceq ('PRODUCTION_CLOSE_'+$code))
        }
        foreach($code in @('REQUESTED','CLOSED','FAILED','CONFIRMED','APPLIED','VALIDATED','STAGED','COMPLETED')){
            $wire=@{ControlId='PRODUCTION_CLOSE';OwnerId='PRODUCTION_WORKFLOW';CatalogVersion=21;OutcomeCode=$code}|ConvertTo-Json -Compress
            Check ('ProductionClose.Terminal.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CloseTerminalForTest' @($wire)) -eq ($code -ceq 'CLOSED'))
        }
        $wrongOwner=@{ControlId='PRODUCTION_CLOSE';OwnerId='RECEIVING_WORKFLOW';CatalogVersion=21;OutcomeCode='CLOSED'}|ConvertTo-Json -Compress
        Check 'ProductionClose.Terminal.WrongOwnerRejected' (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CloseTerminalForTest' @($wrongOwner)))
        $invented='[{"WarehouseId":"CATALOG_TEST","SourceKind":"Inventory","EventId":"'+[guid]::NewGuid().ToString()+'","SubmissionState":"Submitted"}]'
        Check 'ProductionClose.InventedSourcesRejected' (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('PRODUCTION_CLOSE','CLOSED',$invented)))
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Range('A1').Value2='System_Key';$sheet.Range('B1').Value2='Operator Extra'
        $sheet.Range('A2').Value2=[guid]::NewGuid().ToString();$sheet.Range('B2').Value2=$canary
        $table=$sheet.ListObjects.Add(1,$sheet.Range('A1:B2'),$null,1);$table.Name='CloseFixture'
        $path=Join-Path $runRoot 'close-operator.xlsb';$book.SaveAs($path,50);$bookPin=Hash $path
        $cells=$sheet.UsedRange.Formula|ConvertTo-Json -Depth 5 -Compress
        $decoy=$excel.Workbooks.Add();$decoy.Worksheets.Item(1).Range('A1').Value2=$canary
        foreach($target in @($Fixture,$Other)){foreach($file in Get-ChildItem -LiteralPath $target.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}}
        foreach($file in Files){$recordPins[$file]=Hash $file}
        $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
        try {
            $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
            try{
                (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Value2=20.0
                $controls=Table $cfg 'tblEventTrackingControls'
                for($i=$controls.ListRows.Count;$i -ge 1;$i--){if($controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -ceq 'PRODUCTION_CLOSE'){$controls.ListRows.Item($i).Delete()}}
                $cfg.Save()
            }finally{$cfg.Close($false)}
            Check 'ProductionClose.OlderPolicy.ExistingControlStillReadable' (([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_OUTPUT_REGULATION_APPLY'))).StartsWith('True|True|'))
            OpenPrivate;$before=@(Files);[void](Probe 'CloseButtonForTest')
            Check 'ProductionClose.OlderPolicy.NoImplicitEnablement' ((Dismissed) -and @(Files).Count -eq $before.Count)
            [void](Probe 'CloseForgetForTest')
        }finally{[IO.File]::WriteAllBytes($Fixture.Config,$configBytes)}
        foreach($mode in @('Button','Native','Internal')){
            $before=@(Files);OpenPrivate
            Check ('ProductionClose.'+$mode+'.InitializationNotAction') (@(Files).Count -eq $before.Count)
            if($mode -ceq 'Button'){
                $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @('PRODUCTION_CLOSE',21));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
                Check 'ProductionClose.FixedMetadata' ($null -ne $definition -and $definition.Caption -ceq [string](Probe 'CloseCaptionForTest') -and $definition.OwnerId -ceq 'PRODUCTION_WORKFLOW' -and $definition.Class -ceq 'Command' -and $definition.Role -ceq 'Production' -and $definition.Capability -ceq 'PROD_POST' -and $definition.Surface -ceq 'Operations > Production')
            }
            if($mode -cne 'Internal'){CaptureOwnedFormByCaptionEvidence 'Production' ('production-close-'+$mode.ToLowerInvariant()+'.png')}
            $decoy.Activate();$before=@(Files)
            if($mode -ceq 'Native'){Check 'ProductionClose.Native.WindowDestroyed' (Close-ProductionNativeFixture)}
            elseif($mode -ceq 'Button'){[void](Probe 'CloseButtonForTest')}
            else{[void](Probe 'CloseDesigner')}
            Check ('ProductionClose.'+$mode+'.Disposed') (Dismissed)
            if($mode -ceq 'Internal'){Check 'ProductionClose.Internal.NoUserActivity' (@(Files).Count -eq $before.Count)}else{Pair $before ('ProductionClose.'+$mode)}
            [void](Probe 'CloseForgetForTest')
            Check ('ProductionClose.'+$mode+'.WorkbookStagingAndUnknownColumnsPreserved') (($sheet.UsedRange.Formula|ConvertTo-Json -Depth 5 -Compress) -ceq $cells -and (Hash $path) -ceq $bookPin -and $book.Saved)
        }
        foreach($guard in @('Session','SignedOut','Target','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';OpenPrivate
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            [void](Probe 'CloseButtonForTest')
            Check ('ProductionClose.Guard.'+$guard+'.DismissalAllowed') (Dismissed)
            Check ('ProductionClose.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
            [void](Probe 'CloseForgetForTest')
        }
        SelectTarget $Fixture 'config-producer';$book=$excel.Workbooks.Open($path,0,$false)
        OpenPrivate
        $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CloseCanProduceForTest')){throw 'Initial Production permission unavailable.'}
        $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb');$authBytes=[IO.File]::ReadAllBytes($authPath)
        try {
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            try{
                $caps=Table $auth 'tblCapabilities';$changed=0
                foreach($row in $caps.ListRows){if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-producer' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'PROD_POST'){$row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$changed++}}
                if($changed -ne 1){throw 'Unique Production permission fixture required.'};$auth.Save()
            }finally{$auth.Close($false)}
            $lost=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CloseCanProduceForTest')
            $same=$context -ceq [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
            Check 'ProductionClose.PermissionLoss.SameContextAndActualDenial' ($lost -and $same)
            if(-not($lost -and $same)){throw 'Permission loss was not isolated from context loss.'}
            $before=@(Files);[void](Probe 'CloseButtonForTest')
            Check 'ProductionClose.PermissionLoss.DismissalAllowed' (Dismissed);Pair $before 'ProductionClose.PermissionLoss'
            [void](Probe 'CloseForgetForTest')
        }finally{[IO.File]::WriteAllBytes($authPath,$authBytes);SelectTarget $Fixture 'config-producer'}
        AuthorityCheckpoint 'AfterPermissionRestoration'
        SetPolicy $false;$pins[$Fixture.Config]=Hash $Fixture.Config
        OpenPrivate;$before=@(Files);[void](Probe 'CloseButtonForTest')
        Check 'ProductionClose.TrackingOff.DismissalWithoutRecords' ((Dismissed) -and @(Files).Count -eq $before.Count);[void](Probe 'CloseForgetForTest')
        SetPolicy $true;$pins[$Fixture.Config]=Hash $Fixture.Config
        AuthorityCheckpoint 'AfterPolicyCommands'
        $blocked=Join-Path (Join-Path $Fixture.Root 'Training\Activity') $Fixture.Warehouse;$held=$blocked+'-close-held'
        foreach($item in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fault escaped fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        foreach($mode in @('Button','Native')){
            OpenPrivate;$before=@(Files);$moved=$false
            try{
                if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
                [IO.File]::WriteAllText($blocked,'Blocked disposable Close activity path')
                $result=Invoke-ProductionCloseWithNotice $mode
                Check ('ProductionClose.Unavailable.'+$mode+'.DismissalAllowed') ($result.Dismissed -and (Dismissed))
                Check ('ProductionClose.Unavailable.'+$mode+'.ActualTrackingNotice') $result.TrackingNoticeVisible
            }finally{if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked};if($moved){Move-Item -LiteralPath $held -Destination $blocked}}
            Check ('ProductionClose.Unavailable.'+$mode+'.NoFallbackActivity') (@(Files).Count -eq $before.Count)
            [void](Probe 'CloseForgetForTest')
        }
        AuthorityCheckpoint 'BeforePublicLauncher'
        Test-ProductionClosePublic $Fixture $canary
        AuthorityCheckpoint 'AfterPublicLauncher'
        Check 'ProductionClose.SavedOperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
        Check 'ProductionClose.DecoyPreserved' ($decoy.Worksheets.Count -eq 1 -and $decoy.Worksheets.Item(1).Range('A1').Value2 -ceq $canary)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]};Check 'ProductionClose.SavedAuthorityPreserved' $same
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]};Check 'ProductionClose.PriorRecordsImmutable' $same
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}

function Test-ProductionClosePublic($Fixture,[string]$Canary) {
    function Probe([string]$Method){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method)}
    $operatorRoot=Join-Path $runRoot 'close-public-operators';$book=$null;$priorEvents=[bool]$excel.EnableEvents
    if(-not [IO.Path]::GetFullPath($operatorRoot).StartsWith([IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Public operator root escaped fixture.'}
    if(-not [bool](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @($operatorRoot))){throw 'Isolated public root unavailable.'}
    [void](Run 'invSys.Operations.xlam' 'modOperationsInit.Auto_Open')
    # The common fixture intentionally disables Excel events. This public native
    # lifecycle case must enable real WorkbookBeforeClose dispatch explicitly.
    $excel.EnableEvents=$true
    $eventsAtEntry=[bool]$excel.EnableEvents
    try{
        [void](Run 'invSys.Operations.xlam' 'mProduction.BtnOpenProductionForm')
        AuthorityCheckpoint 'AfterPublicInitialLaunch'
        $name=[string](Probe 'CloseBindPublicForTest');$book=$excel.Workbooks.Item($name);$path=$book.FullName
        $owned=[IO.Path]::GetFullPath($path).StartsWith([IO.Path]::GetFullPath($operatorRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)
        Check 'ProductionClose.Public.IsolatedOwner' $owned
        if(-not $owned){throw 'Public fixture escaped owned root.'}
        [void](Probe 'ClosePrepareWorkbenchForTest')
        $sheet=$book.Worksheets.Item('invSys UOM Catalog');$table=$sheet.ListObjects.Item('tblInvSysUomCatalog')
        $extra=$table.ListColumns.Add();$extra.Name='Operator Annotation';$extra.DataBodyRange.Value2=$Canary
        $table.ListColumns.Item('Notes').DataBodyRange.Cells.Item(1,1).Value2=$Canary;$book.Save()
        AuthorityCheckpoint 'AfterPublicWorkbenchSetup'
        $before=@(Get-Slice4beActivityFiles $Fixture)
        [void](Probe 'CloseButtonForTest');Check 'ProductionClose.Public.ButtonDisposes' ([long](Probe 'CloseLoadedFormsForTest') -eq 0)
        # Pair is defined by the caller and verifies this real public form as well.
        Pair $before 'ProductionClose.Public.Button';[void](Probe 'CloseForgetForTest')
        AuthorityCheckpoint 'AfterPublicButtonDismissal'
        [void](Run 'invSys.Operations.xlam' 'mProduction.BtnOpenProductionForm')
        AuthorityCheckpoint 'AfterPublicReopen'
        $name=[string](Probe 'CloseBindPublicForTest');$reopened=$excel.Workbooks.Item($name)
        $retained=$reopened.Worksheets.Item('invSys UOM Catalog').ListObjects.Item('tblInvSysUomCatalog')
        Check 'ProductionClose.Public.ReopensSameOwnerWithUnknownValues' ($reopened -eq $book -and $reopened.FullName -ceq $path -and $retained.ListColumns.Item('Operator Annotation').DataBodyRange.Cells.Item(1,1).Value2 -ceq $Canary -and $retained.ListColumns.Item('Notes').DataBodyRange.Cells.Item(1,1).Value2 -ceq $Canary)
        CaptureOwnedFormByCaptionEvidence 'Production' 'production-close-public-reopened.png'
        $eventsBeforeClose=[bool]$excel.EnableEvents
        $bound=[bool](Run 'invSys.Operations.xlam' 'mProduction.CloseBindingForTest' @($book.Name))
        $entries=[long](Probe 'CloseWorkbookEntryCountForTest')
        $before=@(Get-Slice4beActivityFiles $Fixture);$book.Close($false);$book=$null
        [pscustomobject]@{EventsAtEntry=$eventsAtEntry;EventsBeforeClose=$eventsBeforeClose;OwnerBoundBeforeClose=$bound;WorkbookCloseEntries=([long](Probe 'CloseWorkbookEntryCountForTest')-$entries);LoadedFormsAfterClose=[long](Probe 'CloseLoadedFormsForTest')}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'close-public-lifetime.json')
        Check 'ProductionClose.Public.WorkbookShutdownDisposes' ([long](Probe 'CloseLoadedFormsForTest') -eq 0)
        Check 'ProductionClose.Public.WorkbookShutdownNotUserClose' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        [void](Probe 'CloseForgetForTest')
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};$excel.EnableEvents=$priorEvents}
}
