# Additional protection on the same generated fixtures and actual UOM handlers.
function Test-AdminUomCatalog {
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(9))).Split("`n")|Where-Object {$_ -ne ''})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(10))).Split("`n")|Where-Object {$_ -ne ''})
    Check 'AdminUom.Catalog.TenExtendsNine' ($old.Count -eq 33 -and $new.Count -eq 36 -and @($old|Where-Object {$_ -cnotin $new}).Count -eq 0 -and @($new|Select-Object -Unique).Count -eq 36)
    foreach($id in $old){
        $prior=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,9))
        Check ('AdminUom.Catalog.Preserve.'+$id) ($prior -ne '' -and $prior -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,10)))
    }
    foreach($action in @('Add','Remove','Reset')){
        $id='ADMIN_UOM_'+$action.ToUpperInvariant()
        $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,10))
        $d=if($json){$json|ConvertFrom-Json}else{$null}
        Check ('AdminUom.Catalog.Definition.'+$action) ($null -ne $d -and $d.ControlId -ceq $id -and $d.OwnerId -ceq 'CORE_CONFIGURATION' -and $d.Class -ceq 'Command' -and $d.Role -ceq 'Admin' -and $d.Caption -ceq $action -and $d.Surface -ceq 'Admin > Settings > Recipe UOM Catalog' -and $d.Capability -ceq 'ADMIN_MAINT' -and $d.CodePrefix -ceq ($id+'_'))
        $excluded=$true
        foreach($version in 1..9){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
        Check ('AdminUom.Catalog.OlderVersionsExclude.'+$action) $excluded
        $cancel=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,'CANCELLED'))
        Check ('AdminUom.Catalog.CancellationOnlyReset.'+$action) $(if($action -ceq 'Reset'){$cancel -ne '' -and ($cancel|ConvertFrom-Json).DataEffect -ceq 'Unchanged'}else{$cancel -ceq ''})
    }
}

function Test-AdminUomPublication($Fixture) {
    $pins=@{};$records=@()
    foreach($file in @(Get-Slice4beActivityFiles $Fixture)){
        $pins[$file]=(Get-FileHash -LiteralPath $file).Hash
        $record=Get-Content -LiteralPath $file -Raw|ConvertFrom-Json
        if($record.ControlId -cin @('ADMIN_UOM_ADD','ADMIN_UOM_REMOVE','ADMIN_UOM_RESET')){$records+=$record}
    }
    # Eleven asserted outcomes plus the actual Add that prepares Reset No.
    # Publish all twelve actions; never discard that real fixture observation.
    if($records.Count -ne 24){
        $counts=@($records|Group-Object ControlId|ForEach-Object {$_.Name+':'+$_.Count}) -join ', '
        throw ('UOM publication requires twelve actual command pairs; observed count='+$records.Count+'; '+$counts)
    }
    if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.UomPublishForTest')){throw 'Actual Admin publication fixture failed.'}
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $model=Get-Content -LiteralPath $path -Raw|ConvertFrom-Json
    $groups=@($model.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -cin $records.ActivityId})
    Check 'AdminUom.Publication.AllTwelveActualActions' ($groups.Count -eq 12 -and @($groups.SourceId|Select-Object -Unique).Count -eq 12)
    $lines=@($groups|ForEach-Object {$_.Lines})
    $same=$lines.Count -eq 24
    foreach($record in $records){
        $line=@($lines|Where-Object RecordId -CEQ $record.RecordId)
        $same=$same -and $line.Count -eq 1
        if($line.Count -eq 1){foreach($field in @('ControlId','Caption','OutcomeCode','Severity','DataEffect','CatalogVersion','ActivityId')){$same=$same -and $line[0].$field -ceq $record.$field}}
    }
    Check 'AdminUom.Publication.ExactOriginalIdentityCaptionAndOutcome' $same
    $unchanged=@(Get-Slice4beActivityFiles $Fixture).Count -eq $pins.Count
    foreach($file in $pins.Keys){$unchanged=$unchanged -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
    Check 'AdminUom.Publication.OriginalActivityBytesPreserved' $unchanged
}

function Test-AdminUomTracking($Fixture) {
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $leaf=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
    $held=$leaf+'-uom-tracking-test-held'
    foreach($path in @($leaf,$held,$Fixture.Config)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'UOM tracking fixture escaped generated root.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Preserve existing fixture hold directory.'}
    $original=[IO.File]::ReadAllBytes($Fixture.Config)
    $originalPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    function PolicySnapshot {
        $book=$null
        try {
            $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
            $values=@(foreach($sheet in $book.Worksheets){foreach($table in $sheet.ListObjects){
                if($table.Name -cin @('tblEventTrackingPolicies','tblEventTrackingControls')){
                    [pscustomobject]@{Name=$table.Name;Values=$table.Range.Value2}
                }
            }})
            return ConvertTo-Json -InputObject $values -Depth 8 -Compress
        }finally{if($null -ne $book){$book.Close($false)}}
    }
    foreach($mode in @('UnavailableStore','OlderPolicy','DisabledPolicy')){
        $history=@{};foreach($path in @(Get-Slice4beActivityFiles $Fixture)){$history[$path]=(Get-FileHash -LiteralPath $path).Hash}
        $moved=$false;$blocked=$false;$book=$null
        try {
            SelectTarget $Fixture
            if($mode -ceq 'UnavailableStore'){
                if(Test-Path -LiteralPath $leaf){Move-Item -LiteralPath $leaf -Destination $held;$moved=$true}
                [IO.File]::WriteAllText($leaf,'blocked generated UOM activity path');$blocked=$true
            }else{
                $version=if($mode -ceq 'OlderPolicy'){9}else{10}
                $collect=$mode -ceq 'OlderPolicy'
                $ids=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @($version))).Split("`n")|Where-Object {$_ -ne ''})
                $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
                [void](Add-ActivityFixtureTable $book 'tblEventTrackingPolicies' @('PolicyVersion','SchemaVersion','CatalogVersion','CreatedAtUTC','CreatedByUserId','DefaultView','ViewerActionPathCaptureEnabled','AdminViewerEventLoggingEnabled','Operator Extra') @(,@(1.0,1.0,[double]$version,'2026-09-22T12:00:00.000Z','config-admin','How-To',$false,$true,'preserve')))
                $rows=@(foreach($id in $ids){,@(1.0,$id,$collect,$true,$false,'preserve')})
                [void](Add-ActivityFixtureTable $book 'tblEventTrackingControls' @('PolicyVersion','ControlId','Collect','Visible','SequenceEligible','Operator Extra') $rows)
                $book.Save();$book.Close($false);$book=$null
                $known=if($collect){'True|True|True|1'}else{'True|False|True|1'}
                Check ('AdminUom.Tracking.'+$mode+'.WholePolicyValid') ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('ADMIN_SETTINGS_SAVE_VALUE')) -ceq $known)
                foreach($id in @('ADMIN_UOM_ADD','ADMIN_UOM_REMOVE','ADMIN_UOM_RESET')){
                    $expected=if($collect){'False|False|False|0'}else{'True|False|True|1'}
                    Check ('AdminUom.Tracking.'+$mode+'.ExplicitPolicy.'+$id) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @($id)) -ceq $expected)
                }
            }
            $policyBefore=PolicySnapshot
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
            $canary='UOMTRACK'+[guid]::NewGuid().ToString('N').Substring(0,10).ToUpperInvariant()
            foreach($action in @('Add','Remove','ResetNo','ResetYes')){
                if($action -ceq 'ResetNo'){
                    [void](Run 'invSys.Core.xlam' 'modUomSettings.AddConfiguredUom' @($canary))
                    $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
                }
                if($action.StartsWith('Reset',[StringComparison]::Ordinal)){
                    $choice=if($action -ceq 'ResetNo'){'No'}else{'Yes'}
                    $status=Invoke-AdminUomResetChoice $choice ('uom-tracking-'+$mode.ToLowerInvariant()+'-'+$choice.ToLowerInvariant()+'.png')
                }else{$status=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.UomActivityAction' @($action,$canary))}
                $values=([string](Run 'invSys.Core.xlam' 'modUomSettings.GetConfiguredUomsPackedText')).Split('|')
                $changed=if($action -cin @('Add','ResetNo')){$canary -cin $values}else{$canary -cnotin $values}
                if($action -ceq 'ResetNo'){$changed=$changed -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin}
                Check ('AdminUom.Tracking.'+$mode+'.'+$action+'.OwnerResultPreserved') $changed
                Check ('AdminUom.Tracking.'+$mode+'.'+$action+'.TrackingNotice') ($status.Contains('Tracking unavailable') -eq ($mode -cne 'DisabledPolicy'))
                if($blocked){$untouched=(Test-Path -LiteralPath $leaf -PathType Leaf) -and [IO.File]::ReadAllText($leaf) -ceq 'blocked generated UOM activity path'}
                else{$untouched=@(Get-Slice4beActivityFiles $Fixture|Where-Object {-not $history.ContainsKey($_)}).Count -eq 0}
                Check ('AdminUom.Tracking.'+$mode+'.'+$action+'.NoImplicitActivity') $untouched
            }
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
            Check ('AdminUom.Tracking.'+$mode+'.PolicyAndUnknownColumnsUnchanged') ((PolicySnapshot) -ceq $policyBefore)
        }finally{
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
            if($null -ne $book){$book.Close($false)}
            if($blocked -and (Test-Path -LiteralPath $leaf -PathType Leaf)){Remove-Item -LiteralPath $leaf}
            if($moved){Move-Item -LiteralPath $held -Destination $leaf}
            [IO.File]::WriteAllBytes($Fixture.Config,$original)
            [void](Run 'invSys.Core.xlam' 'modConfig.Reload')
        }
        $after=@(Get-Slice4beActivityFiles $Fixture);$same=$after.Count -eq $history.Count
        foreach($path in $after){$same=$same -and $history.ContainsKey($path) -and (Get-FileHash -LiteralPath $path).Hash -ceq $history[$path]}
        Check ('AdminUom.Tracking.'+$mode+'.OriginalActivityBytesPreserved') $same
        Check ('AdminUom.Tracking.'+$mode+'.OriginalConfigRestored') ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $originalPin)
    }
}
