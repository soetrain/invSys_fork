# D8-INV-CONFIG-01 retained shared/automation contracts. This complements, rather
# than replaces, the named owning-caller matrix and packaged operator regressions.
function Install-InventoryConfigRetainedProbe {
    $project=$packages['invSys.Core.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryConfigRetained'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function EnsureSharedInbox(ByVal warehouse As String, ByVal station As String, _
        ByVal role As String, ByVal config As String) As String
    Dim path As String, report As String
    If modConfig.EnsureStationInbox(warehouse, station, role, config, path, report) Then
        EnsureSharedInbox = "OK|" & path
    Else
        EnsureSharedInbox = "FAIL"
    End If
End Function
'@)
}

function Test-InventoryConfigRetainedShared($Fixture) {
    $paths=@{};$inboxValues=@{}
    $required=@('EventID','ParentEventId','UndoOfEventId','EventType','CreatedAtUTC','WarehouseId','StationId',
        'UserId','MigrationSourceId','System_Key','SKU','Qty','Location','Condition','AttributesJson',
        'DesignId','DesignVersion','Note','PayloadJson','Status','RetryCount','ErrorCode','ErrorMessage','FailedAtUTC')
    foreach($entry in @('Shared','Automation','Packed')){
        foreach($role in @('RECEIVE','SHIP','PRODUCTION')){
            $id='InventoryConfigRetained.'+$entry+'.'+$role
            SelectTarget $Fixture
            Set-InventoryConfigPresentation $Fixture 71 'WarehouseConfig'
            $before=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $inventoryPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
            $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
            $inventoryHash=Get-AdminInventoryEntryHash $inventoryPath;$authHash=Get-AdminInventoryEntryHash $authPath
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            $arguments=@($Fixture.Warehouse,'S1',$role,$Fixture.Config)
            $macro=switch($entry){
                'Shared' {'TestInventoryConfigRetained.EnsureSharedInbox'}
                'Automation' {'modConfig.EnsureStationInboxForAutomation'}
                'Packed' {'modConfig.EnsureStationInboxPackedForAutomation'}
            }
            if($entry -ceq 'Packed'){$arguments=@($arguments -join '|')}
            $result=[string](Run 'invSys.Core.xlam' $macro $arguments)
            Check ($id+'.Succeeded') ($result.StartsWith('OK|',[StringComparison]::Ordinal))
            if(-not $result.StartsWith('OK|',[StringComparison]::Ordinal)){continue}
            $path=$result.Substring(3)
            $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
            $owned=[IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)
            Check ($id+'.InboxWithinOwnedFixture') ($owned -and (Test-Path -LiteralPath $path -PathType Leaf))
            if(-not $owned -or -not (Test-Path -LiteralPath $path -PathType Leaf)){throw 'Retained inbox is outside the owned test fixture.'}
            if($entry -ceq 'Shared'){$paths[$role]=$path}
            Check ($id+'.SameRoleDestinationAcrossEntries') ($paths.ContainsKey($role) -and $paths[$role] -ceq $path)
            $after=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $changes=@(@($before.Cells.Keys)+@($after.Cells.Keys)|Select-Object -Unique|Where-Object {
                -not $before.Cells.ContainsKey($_) -or -not $after.Cells.ContainsKey($_) -or $before.Cells[$_] -cne $after.Cells[$_]
            })
            Check ($id+'.ConfigValuesAndShapePreserved') ($changes.Count -eq 0 -and $before.Shapes -ceq $after.Shapes)
            Check ($id+'.InventoryAndAuthBytesPreserved') ($inventoryHash -ceq (Get-AdminInventoryEntryHash $inventoryPath) -and $authHash -ceq (Get-AdminInventoryEntryHash $authPath))
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            Check ($id+'.RetainedConfigSave') ($saveFacts.Count -eq 3 -and [int]$saveFacts[0] -eq 1 -and [int]$saveFacts[2] -eq 0)
            $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
            $sheets=$null;$sheet=$null;$columns=$null;$column=$null;$rows=$null;$row=$null;$font=$null
            try{
                $sheets=$book.Worksheets;$sheet=$sheets.Item('WarehouseConfig');$columns=$sheet.Columns;$column=$columns.Item(1)
                $rows=$sheet.Rows;$row=$rows.Item(1);$font=$row.Font
                Check ($id+'.RetainedAutoFit') ([double]$column.ColumnWidth -ne 71)
                Check ($id+'.RetainedHeaderBold') ($font.Bold -eq $true)
            }finally{
                foreach($ownedCom in @($font,$row,$rows,$column,$columns,$sheet,$sheets)){if($null -ne $ownedCom){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ownedCom)}}
                $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
            }
            Check ($id+'.InspectionPreservedConfigBytes') ($after.ReadOnlyProof -and $after.FileHash -ceq (Get-AdminInventoryEntryHash $Fixture.Config))
            $inboxHash=Get-AdminInventoryEntryHash $path
            $book=$excel.Workbooks.Open($path,0,$true);$table=$null;$range=$null
            try{
                $name=switch($role){'RECEIVE'{'tblInboxReceive'} 'SHIP'{'tblInboxShip'} 'PRODUCTION'{'tblInboxProd'}}
                $table=Table $book $name
                Check ($id+'.ExpectedInboxTable') ($null -ne $table)
                if($null -ne $table){
                    $range=$table.Range;$values=$range.Value2;$headers=@()
                    for($columnIndex=1;$columnIndex -le $values.GetLength(1);$columnIndex++){$headers+=[string]$values.GetValue(1,$columnIndex)}
                    Check ($id+'.RequiredInboxSchema') (@($required|Where-Object {$_ -cnotin $headers}).Count -eq 0 -and 'ROW' -cnotin $headers)
                    # Complete row values stay in memory; only comparisons leave here.
                    $serialized=[Management.Automation.PSSerializer]::Serialize($values)
                    if($entry -ceq 'Shared'){$inboxValues[$role]=$serialized}
                    Check ($id+'.SameInboxValuesAcrossEntries') ($inboxValues.ContainsKey($role) -and $inboxValues[$role] -ceq $serialized)
                }
            }finally{
                if($null -ne $range){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)}
                if($null -ne $table){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($table)}
                $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
            }
            Check ($id+'.InspectionPreservedInboxBytes') ($inboxHash -ceq (Get-AdminInventoryEntryHash $path))
        }
    }
}
