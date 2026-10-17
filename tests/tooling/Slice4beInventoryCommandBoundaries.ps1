# Actual packaged actions on Admin-generated disposable fixtures. Values stay in memory.
function Get-InventoryConfigSemanticSnapshot($Book) {
    $groups=@{Workbook=@{};Names=@{};Tables=@{};Views=@{};Sheets=@{};Connections=@{}}
    $unavailable=@{Workbook=0;Names=0;Tables=0;Views=0;Sheets=0;Connections=0}
    $missingWorkbookProperties=@()
    foreach($name in @('Date1904','PrecisionAsDisplayed','ProtectStructure','ProtectWindows','ReadOnlyRecommended','SaveLinkValues','UpdateRemoteReferences','CheckCompatibility','ForceFullCalculation','FullCalculationOnLoad','CalculationVersion')){
        try{$groups.Workbook[$name]=[string]$Book.$name}catch{$unavailable.Workbook++;$missingWorkbookProperties+=$name}
    }
    try{
        $groups.Names['Count']=[string]$Book.Names.Count
        foreach($binding in $Book.Names){
            $groups.Names[$binding.Name+'|RefersTo']=[string]$binding.RefersTo
            $groups.Names[$binding.Name+'|Visible']=[string]$binding.Visible
        }
    }catch{$unavailable.Names++}
    foreach($sheet in $Book.Worksheets){
        foreach($name in @('Visible','EnableCalculation','ProtectContents','ProtectDrawingObjects','ProtectScenarios','CodeName')){
            try{$groups.Sheets[$sheet.Name+'|'+$name]=[string]$sheet.$name}catch{$unavailable.Sheets++}
        }
        foreach($table in $sheet.ListObjects){
            try{
                $groups.Tables[$table.Name+'|Sheet']=[string]$sheet.Name
                $groups.Tables[$table.Name+'|Range']=[string]$table.Range.Address($false,$false,1,$false)
                $groups.Tables[$table.Name+'|SourceType']=[string]$table.SourceType
                $groups.Tables[$table.Name+'|DisplayName']=[string]$table.DisplayName
                $groups.Tables[$table.Name+'|ShowTotals']=[string]$table.ShowTotals
            }catch{$unavailable.Tables++}
        }
    }
    try{
        $groups.Views['Count']=[string]$Book.Windows.Count
        $index=0
        foreach($window in $Book.Windows){
            $index++
            foreach($name in @('Visible','DisplayWorkbookTabs','DisplayHorizontalScrollBar','DisplayVerticalScrollBar','TabRatio','WindowState','Zoom','ScrollRow','ScrollColumn','Left','Top','Width','Height')){
                try{$groups.Views[$index.ToString()+'|'+$name]=[string]$window.$name}catch{$unavailable.Views++}
            }
            try{$groups.Views[$index.ToString()+'|ActiveSheet']=[string]$window.ActiveSheet.Name}catch{$unavailable.Views++}
        }
    }catch{$unavailable.Views++}
    # Do not read connection strings, credentials, document authors or raw properties.
    try{$groups.Connections['Count']=[string]$Book.Connections.Count}catch{$unavailable.Connections++}
    [pscustomobject]@{Groups=$groups;Unavailable=$unavailable;MissingWorkbookProperties=$missingWorkbookProperties}
}

function Compare-InventoryConfigSemanticSnapshot($Before,$After) {
    $result=[ordered]@{};$complete=$true
    foreach($group in @('Workbook','Names','Tables','Views','Sheets','Connections')){
        $a=$Before.Groups[$group];$b=$After.Groups[$group];$changed=0
        foreach($key in @($a.Keys)+@($b.Keys)|Select-Object -Unique){
            if(-not $a.ContainsKey($key) -or -not $b.ContainsKey($key) -or $a[$key] -cne $b[$key]){$changed++}
        }
        $missing=[int]$Before.Unavailable[$group]+[int]$After.Unavailable[$group]
        if($missing){$complete=$false}
        $result[$group]=[pscustomobject]@{BeforeCount=$a.Keys.Count;AfterCount=$b.Keys.Count;ChangedCount=$changed;UnavailableReads=$missing}
    }
    # Keys, formulas, values, author identities and paths are never emitted.
    $missing=@()
    if($Before.PSObject.Properties['MissingWorkbookProperties'] -and $After.PSObject.Properties['MissingWorkbookProperties']){
        $missing=@(@($Before.MissingWorkbookProperties)+@($After.MissingWorkbookProperties)|Select-Object -Unique)
    }
    [pscustomobject]@{CompleteForObservedProperties=$complete;Groups=[pscustomobject]$result;UnavailableWorkbookProperties=$missing;AllWorkbookSemanticsProven=$false}
}

# MS-XLSB 2.1.4 record framing; payloads/hashes remain memory-only.
# https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/acc8aa92-1f02-4167-99f5-84f9f676b95a
function Get-InventoryConfigBinaryRecords([byte[]]$Bytes) {
    $position=0;$records=@{};$sequence=[Collections.Generic.List[string]]::new();$ordinals=@{}
    $sha=[Security.Cryptography.SHA256]::Create()
    try{
        while($position -lt $Bytes.Length){
            $start=$position;$first=[int]$Bytes[$position++];$type=$first -band 127
            if($first -band 128){
                if($position -ge $Bytes.Length){throw 'Truncated record type.'}
                $next=[int]$Bytes[$position++];if($next -band 128){throw 'Invalid record type.'}
                $type=$type -bor (($next -band 127) -shl 7)
                if($type -lt 128){throw 'Noncanonical record type.'}
            }
            $size=0;$done=$false
            for($shift=0;$shift -lt 4;$shift++){
                if($position -ge $Bytes.Length){throw 'Truncated record size.'}
                $next=[int]$Bytes[$position++];$size=$size -bor (($next -band 127) -shl (7*$shift))
                if(-not($next -band 128) -or $shift -eq 3){$done=$true;break}
            }
            if(-not $done -or $size -gt $Bytes.Length-$position){throw 'Truncated record data.'}
            $ordinal=if($ordinals.ContainsKey($type)){[int]$ordinals[$type]+1}else{1};$ordinals[$type]=$ordinal
            $key=[string]$type+':'+[string]$ordinal;$data=[byte[]]::new($size)
            [Array]::Copy($Bytes,$position,$data,0,$size)
            $records[$key]=[pscustomobject]@{Type=$type;Data=$data;Hash=[Convert]::ToBase64String($sha.ComputeHash($data));HeaderHash=[Convert]::ToBase64String($sha.ComputeHash($Bytes,$start,$position-$start))}
            $sequence.Add($key);$position+=$size
        }
        [pscustomobject]@{Parsed=$true;Records=$records;Sequence=($sequence -join '|')}
    }catch{[pscustomobject]@{Parsed=$false;Records=@{};Sequence=''}}finally{$sha.Dispose()}
}

function Compare-InventoryConfigBinaryRecords($Before,$After) {
    if($null -eq $Before -or $null -eq $After -or -not $Before.Parsed -or -not $After.Parsed){return [pscustomobject]@{Parsed=$false}}
    $types=@{};$viewFields=@{};$headersSame=$true;$unknownView=$false
    foreach($key in @($Before.Records.Keys)+@($After.Records.Keys)|Select-Object -Unique){
        if(-not $Before.Records.ContainsKey($key) -or -not $After.Records.ContainsKey($key)){$types[[int]($key -split ':')[0]]=$true;$headersSame=$false;continue}
        $a=$Before.Records[$key];$b=$After.Records[$key]
        if($a.HeaderHash -cne $b.HeaderHash){$headersSame=$false}
        if($a.Hash -ceq $b.Hash){continue}
        $types[[int]$a.Type]=$true
        # MS-XLSB BrtBookView is type158; only its fixed29-byte layout is decoded.
        if($a.Type -eq 158){
            if($a.Data.Length -ne 29 -or $b.Data.Length -ne 29){$unknownView=$true;continue}
            $names=@('WindowLeft','WindowTop','WindowWidth','WindowHeight','TabRatio','FirstSheetIndex','ActiveSheetIndex')
            for($i=0;$i -lt 7;$i++){if([BitConverter]::ToUInt32($a.Data,4*$i) -ne [BitConverter]::ToUInt32($b.Data,4*$i)){$viewFields[$names[$i]]=$true}}
            if($a.Data[28] -ne $b.Data[28]){$viewFields['ViewFlags']=$true;if(($a.Data[28] -bxor $b.Data[28]) -band 128){$unknownView=$true}}
        }
    }
    [pscustomobject]@{Parsed=$true;BeforeRecordCount=$Before.Records.Keys.Count;AfterRecordCount=$After.Records.Keys.Count;RecordOrderPreserved=($Before.Sequence -ceq $After.Sequence);RecordHeadersPreserved=$headersSame;ChangedRecordTypes=@($types.Keys|Sort-Object);ChangedBookViewFields=@($viewFields.Keys|Sort-Object);UnknownBookViewChange=$unknownView;AllWorkbookSemanticsProven=$false}
}

function Get-InventoryConfigDiagnosticSnapshot($Fixture) {
    $path=[IO.Path]::GetFullPath($Fixture.Config)
    $expected=[IO.Path]::GetFullPath((Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Config.xlsb')))
    $ownedRoot=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if($path -ine $expected -or -not $path.StartsWith($ownedRoot,[StringComparison]::OrdinalIgnoreCase)){
        throw 'Config diagnosis requires this run owned generated fixture.'
    }
    $beforeHash=Get-AdminInventoryEntryHash $path
    Add-Type -AssemblyName System.IO.Compression
    $parts=@{};$binaryRecords=$null
    $stream=[IO.FileStream]::new($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    try {
        $zip=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Read,$true)
        try {
            foreach($entry in $zip.Entries){
                $part=$entry.Open()
                try{$parts[$entry.FullName]=(Get-FileHash -InputStream $part).Hash}finally{$part.Dispose()}
                if($entry.FullName -ceq 'xl/workbook.bin'){
                    $part=$entry.Open();$memory=[IO.MemoryStream]::new()
                    try{$part.CopyTo($memory);$binaryRecords=Get-InventoryConfigBinaryRecords $memory.ToArray()}finally{$memory.Dispose();$part.Dispose()}
                }
            }
        }finally{$zip.Dispose()}
    }finally{$stream.Dispose()}
    $book=$null;$opened=$false;$prior=$null
    try{$prior=$excel.ActiveWorkbook}catch{}
    foreach($candidate in $excel.Workbooks){if($candidate.FullName -ieq $path){$book=$candidate;break}}
    if($null -eq $book){$book=$excel.Workbooks.Open($path,0,$true);$opened=$true}
    $saved=$book.Saved;$cells=@{};$shapes=@()
    try {
        foreach($sheet in $book.Worksheets){foreach($table in $sheet.ListObjects){
            $values=$table.Range.Value2
            $rows=$table.Range.Rows.Count;$columns=$table.Range.Columns.Count
            $shapes+=($table.Name+'|'+$rows+'|'+$columns)
            for($r=1;$r -le $rows;$r++){for($c=1;$c -le $columns;$c++){
                $v=if($values -is [Array]){$values.GetValue($r,$c)}else{$values}
                $cells[$table.Name+'|'+$r+'|'+$c]=if($null -eq $v){'<EMPTY>'}else{$v.GetType().FullName+'|'+[string]$v}
            }}
        }}
        $semantics=Get-InventoryConfigSemanticSnapshot $book
        $memoryPreserved=($book.Saved -eq $saved)
    }finally{
        if($opened){$book.Close($false)}
        if($null -ne $prior){$prior.Activate()}
    }
    # Values and part hashes are deliberately memory-only; export comparisons below.
    [pscustomobject]@{Cells=$cells;Shapes=($shapes -join "`n");Parts=$parts;Semantics=$semantics;BinaryRecords=$binaryRecords;ReadOnlyProof=($memoryPreserved -and $beforeHash -ceq (Get-AdminInventoryEntryHash $path))}
}

function Write-InventoryConfigDiagnosticComparison($Before,$After,[string]$Case,$SavePhase=$null) {
    $changed=0
    foreach($key in @($Before.Cells.Keys)+@($After.Cells.Keys)|Select-Object -Unique){
        if(-not $Before.Cells.ContainsKey($key) -or -not $After.Cells.ContainsKey($key) -or $Before.Cells[$key] -cne $After.Cells[$key]){$changed++}
    }
    $categories=@{Metadata=0;Worksheet=0;Styles=0;SharedStrings=0;Other=0}
    $standard=@('docProps/core.xml','docProps/app.xml','xl/workbook.bin','xl/calcChain.bin','xl/_rels/workbook.bin.rels','[Content_Types].xml','_rels/.rels')
    $known=@();$unidentified=0
    foreach($key in @($Before.Parts.Keys)+@($After.Parts.Keys)|Select-Object -Unique){
        if($Before.Parts[$key] -cne $After.Parts[$key]){
            $category=if($key -like 'docProps/*'){'Metadata'}elseif($key -like 'xl/worksheets/*'){'Worksheet'}elseif($key -like 'xl/styles*'){'Styles'}elseif($key -like 'xl/sharedStrings*'){'SharedStrings'}else{'Other'}
            $categories[$category]++
            if($key -cin $standard){$known+=$key}else{$unidentified++}
        }
    }
    $safe=$Before.ReadOnlyProof -and $After.ReadOnlyProof
    $semantics=$null
    if($Before.PSObject.Properties['Semantics'] -and $After.PSObject.Properties['Semantics']){
        $semantics=Compare-InventoryConfigSemanticSnapshot $Before.Semantics $After.Semantics
    }
    $binary=$null
    if($Before.PSObject.Properties['BinaryRecords'] -and $After.PSObject.Properties['BinaryRecords']){
        $binary=Compare-InventoryConfigBinaryRecords $Before.BinaryRecords $After.BinaryRecords
    }
    [void]$script:InventoryConfigReadChecks.Add([pscustomobject]@{Case=$Case;Passed=$safe})
    [pscustomobject]@{Case=$Case;ReadOnlyExtractionPreserved=$safe;ChangedTableCells=$changed;TableShapesPreserved=($Before.Shapes -ceq $After.Shapes);ChangedPartCategories=$categories;ChangedStandardParts=$known;UnidentifiedChangedPartCount=$unidentified;ObservedConfigSavePhase=$SavePhase;ObservedSemanticComparison=$semantics;ObservedBinaryComparison=$binary;CauseEstablished=$false}|
        ConvertTo-Json -Depth 6 -Compress|Add-Content (Join-Path $reportRoot 'inventory-config-diagnosis.jsonl')
}

function Test-InventoryConfigDiagnosticPreservation {
    foreach($proof in $script:InventoryConfigReadChecks){Check ('InventoryConfigDiagnosis.'+$proof.Case+'.ReadOnlyExtractionPreserved') $proof.Passed}
}

function Install-InventoryCommandBoundaryProbe {
    $coreProject=$packages['invSys.Core.xlam'].VBProject
    $saveProbe=$coreProject.VBComponents.Add(1);$saveProbe.Name='TestInventoryConfigSave'
    $saveProbe.CodeModule.AddFromString(@'
Option Explicit
Private mCalls As Long, mChanges As Long, mUnavailable As Long, mBefore As String
Public Sub Reset()
    mCalls = 0: mChanges = 0: mUnavailable = 0: mBefore = ""
End Sub
Public Sub BeforeSave(ByVal book As Workbook)
    If Not LCase$(book.Name) Like "*.invsys.config.xls*" Then Exit Sub
    mCalls = mCalls + 1
    mBefore = modTrainingWire.FileSha256(book.FullName)
End Sub
Public Sub AfterSave(ByVal book As Workbook)
    Dim hash As String
    If Not LCase$(book.Name) Like "*.invsys.config.xls*" Then Exit Sub
    hash = modTrainingWire.FileSha256(book.FullName)
    If hash = "" Or mBefore = "" Then
        mUnavailable = mUnavailable + 1
    ElseIf hash <> mBefore Then
        mChanges = mChanges + 1
    End If
    mBefore = ""
End Sub
Public Function Facts() As String
    Facts = CStr(mCalls) & "|" & CStr(mChanges) & "|" & CStr(mUnavailable)
End Function
'@)
    $config=$coreProject.VBComponents.Item('modConfig').CodeModule
    $start=$config.ProcStartLine('SaveConfigWorkbookIfWritable',0);$count=$config.ProcCountLines('SaveConfigWorkbookIfWritable',0)
    $saveBody=$config.Lines($start,$count);$saveLine='    wb.Save'
    if([regex]::Matches($saveBody,[regex]::Escape($saveLine)).Count -ne 1){throw 'Config save observation boundary changed; not product RED.'}
    $saveBody=$saveBody.Replace($saveLine,('    TestInventoryConfigSave.BeforeSave wb'+"`r`n"+$saveLine+"`r`n"+'    TestInventoryConfigSave.AfterSave wb'))
    $config.DeleteLines($start,$count);$config.InsertLines($start,$saveBody)
    $project=$packages['invSys.Admin.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryCommands'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mDriving As Boolean, mAction As String, mCase As String, mPhase As String
Private mRoot As String, mWh As String, mSecret As String, mOtherRoot As String
Private mHost As String, mSku As String, mBefore As String
Private mShows As Long, mHandler As Long, mYield As Boolean, mPrepared As Boolean
Private mBound As Boolean, mSignatureReady As Boolean
Private mInjectedEvents As Long
Private mQuantityTarget As Double
Private mCustomHeader As String, mCustomAccepted As Boolean
Private mHeaderRefused As Boolean
Private mLastNotice As String
Public Sub RecordNotice(ByVal text As String)
    mLastNotice = text
End Sub
Public Function PartialQueueOutcomeReported() As Boolean
    PartialQueueOutcomeReported = (InStr(1, mLastNotice, "Event:", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "unconfirmed", vbTextCompare) > 0)
End Function
Public Function IsDriving() As Boolean
    IsDriving = mDriving
End Function
Public Function PrepareBoundaryTable(ByVal host As String) As Boolean
    Dim name As String, report As String, wb As Workbook, lo As ListObject, ws As Worksheet
    Set wb = Application.Workbooks(host)
    If Not modAdminInventoryWorksheet.CreateInventoryWorksheetTable(wb, name, report) Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = name Then
                lo.ListColumns("Action").DataBodyRange.Cells(1, 1).Value2 = "ADD"
                lo.ListColumns("Item Name").DataBodyRange.Cells(1, 1).Value2 = "Zero worksheet fixture"
                lo.ListColumns("UOM").DataBodyRange.Cells(1, 1).Value2 = "LB"
                lo.ListColumns("Qty Mode").DataBodyRange.Cells(1, 1).Value2 = "COUNTED"
                lo.ListColumns("Quantity").DataBodyRange.Cells(1, 1).Value2 = 0
                lo.ListColumns("Default Location").DataBodyRange.Cells(1, 1).Value2 = "CLEARVIEW"
                wb.Save
                lo.DataBodyRange.Cells(1, 1).Select
                PrepareBoundaryTable = True
                Exit Function
            End If
        Next lo
    Next ws
End Function
Public Sub SetSelection(ByVal otherRoot As String, ByVal sku As String)
    mOtherRoot = otherRoot: mSku = sku
End Sub
Public Sub SetQuantityTarget(ByVal target As Double)
    mQuantityTarget = target
End Sub
Public Sub SetCustomHeader(ByVal header As String)
    mCustomHeader = header
End Sub
Public Function CustomFieldAccepted() As Boolean
    CustomFieldAccepted = mCustomAccepted
End Function
Public Function ManagedHeaderRefused() As Boolean
    ManagedHeaderRefused = mHeaderRefused
End Function
Public Function AddManagedFixtureColumn(ByVal host As String, ByVal header As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, col As ListColumn
    Set wb = Application.Workbooks(host)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If Left$(lo.Name, 17) = "invSys_Inventory_" Then
                Set col = lo.ListColumns.Add
                col.Name = header
                col.DataBodyRange.Cells(1, 1).Value2 = IIf(header = "Condition", "GOOD", "Fixture field value")
                wb.Save
                lo.DataBodyRange.Cells(1, 1).Select
                AddManagedFixtureColumn = True
                Exit Function
            End If
        Next lo
    Next ws
End Function
Public Function PrepareExactQuantityFixture(ByVal warehouse As String, ByVal csv As String) As Boolean
    Dim report As String
    PrepareExactQuantityFixture = modAdminInventorySeed.UploadDemoInventoryForWarehouse( _
        warehouse, "S1", "config-admin", csv, report)
End Function
Public Sub Execute(ByVal host As String, ByVal action As String, ByVal caseName As String, _
                   ByVal phase As String, ByVal root As String, ByVal wh As String, _
                   ByVal secret As String)
    Dim control As New TestInventoryEntryRibbon
    mDriving = True: mHost = host: mAction = action: mCase = caseName: mPhase = phase
    mRoot = root: mWh = wh: mSecret = secret
    mShows = 0: mHandler = 0: mYield = False: mPrepared = False: mBound = False
    mBefore = "": mSignatureReady = False
    mInjectedEvents = 0
    mCustomAccepted = False
    mHeaderRefused = False
    mLastNotice = ""
    Application.Workbooks(host).Activate
    On Error GoTo Failed
    RibbonOnActionAdmin control
    mDriving = False: mSecret = ""
    Exit Sub
Failed:
    mDriving = False: mSecret = ""
    Err.Raise vbObjectError + 2101, , "Inventory action adapter failed."
End Sub
Public Sub Drive(ByVal form As frmAddInventoryItem)
    mShows = mShows + 1
    If mShows > 1 Then form.InventoryEntryCancelForTest: Exit Sub
    mBound = form.InventoryEntryOwnerBoundForTest(mHost)
    mPrepared = form.InventoryCommandPrepareForTest(mAction, mSku)
    If Not mPrepared Then form.InventoryEntryCancelForTest: Exit Sub
    If mCase = "SetQuantity" Then form.InventoryCommandTargetForTest mQuantityTarget
    If mPhase = "ModalReturn" Then YieldAt mPhase
    If mPhase = "UomPrompt" Then form.InventoryCommandUomPromptForTest
    mHandler = mHandler + 1
    If mAction = "CustomField" Then
        mCustomAccepted = form.InventoryCommandCustomFieldForTest(mCustomHeader)
    Else
        form.InventoryCommandInvokeForTest mAction
    End If
    If mAction = "Upload" Then mHeaderRefused = form.InventoryCommandHeaderRefusalForTest()
    If form.Visible Then form.InventoryEntryCancelForTest
End Sub
Public Sub YieldAt(ByVal phase As String)
    If Not mDriving Or mYield Or phase <> mPhase Then Exit Sub
    Dim signed As String, selected As String, fields As Object, report As String
    Select Case mCase
        Case "Authorized", "Zero", "Utility", "SetQuantity"
            mYield = True
        Case "SignedOut"
            modAuth.SignOut
            mYield = Not modAuth.IsSignedIn()
        Case "ReplacementSession", "Reader", "OtherWarehouseAdmin"
            modAuth.SignOut
            If mCase = "OtherWarehouseAdmin" Then
                modRuntimeWorkbooks.SetCoreDataRootOverride mOtherRoot
                selected = modNasConnection.SelectWarehouseTargetForAutomation(mOtherRoot, mOtherRoot, "S1", False)
                modNasConnection.SetCurrentTargetPathsForTest "\\fixture-host\config-command", mOtherRoot
            End If
            signed = modAuth.SignInCurrentTargetForAutomation(IIf(mCase = "Reader", "config-reader", "config-admin"), mSecret, "")
            mYield = (Left$(signed, 3) = "OK|")
        Case "CatalogChanged"
            Set fields = CreateObject("Scripting.Dictionary")
            mYield = modAdmin.UpdateInventoryItemCatalogForWarehouse(mWh, mSku, _
                "Externally changed fixture catalog", "LB", "CLEARVIEW", , , , , , , fields, report, "", "config-admin", "S1")
        Case "SourceFeedChanged"
            modRuntimeWorkbooks.SetCoreDataRootOverride mOtherRoot
            mYield = True
        Case "CatalogReplaced", "TableReplaced"
            mYield = ReplaceFixtureTable(mCase = "CatalogReplaced")
        Case "EntityChanged"
            mYield = ChangeExactFixtureEntity()
    End Select
    If mYield Then
        mBefore = InventorySignature()
        mSignatureReady = True
    End If
End Sub
Public Function DialogAnswer(ByVal prompt As String) As String
    If InStr(1, prompt, "New UOM:", vbTextCompare) > 0 Then
        YieldAt "UomPrompt"
        DialogAnswer = "FIXTUREUOM"
        Exit Function
    End If
    If InStr(1, prompt, "Why the edit", vbTextCompare) > 0 Then
        YieldAt "EditReason"
    ElseIf InStr(1, prompt, "deleted from managed", vbTextCompare) > 0 Then
        YieldAt "DeleteReason"
    End If
    DialogAnswer = "Disposable fixture reason"
End Function
Public Function InjectedEvents() As Long
    InjectedEvents = mInjectedEvents
End Function
Private Function ReplaceFixtureTable(ByVal catalog As Boolean) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, selected As ListObject
    Dim cells As Range, tableName As String
    If catalog Then
        Set wb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(mWh)
    Else
        Set wb = Application.Workbooks(mHost)
    End If
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If (catalog And lo.Name = "tblSkuCatalog") Or _
               (Not catalog And Left$(lo.Name, 17) = "invSys_Inventory_") Then
                Set selected = lo: Exit For
            End If
        Next lo
        If Not selected Is Nothing Then Exit For
    Next ws
    If selected Is Nothing Then Exit Function
    Set cells = selected.Range: tableName = selected.Name
    modAdmin.InventoryBoundaryProtectionForTest selected.Parent, False
    selected.Unlist
    Set lo = cells.Parent.ListObjects.Add(xlSrcRange, cells, , xlYes)
    lo.Name = tableName
    If catalog Then modAdmin.InventoryBoundaryProtectionForTest lo.Parent, True
    wb.Save
    If Not catalog Then wb.Activate: lo.DataBodyRange.Cells(1, 1).Select
    ReplaceFixtureTable = True
End Function
Private Function ChangeExactFixtureEntity() As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, r As Long
    Dim item As Object, key As String, location As String, condition As String
    Dim payload As String, eventId As String, report As String, confirmed As Boolean
    Set wb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(mWh)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblInventoryEntities" Then
                For r = 1 To lo.ListRows.Count
                    If CStr(lo.DataBodyRange.Cells(r, lo.ListColumns("SKU").Index).Value2) = mSku And _
                       CStr(lo.DataBodyRange.Cells(r, lo.ListColumns("InventoryState").Index).Value2) = "ACTIVE" Then
                        key = CStr(lo.DataBodyRange.Cells(r, lo.ListColumns("System_Key").Index).Value2)
                        location = CStr(lo.DataBodyRange.Cells(r, lo.ListColumns("Location").Index).Value2)
                        condition = CStr(lo.DataBodyRange.Cells(r, lo.ListColumns("Condition").Index).Value2)
                        Exit For
                    End If
                Next r
            End If
        Next lo
    Next ws
    If key = "" Then Exit Function
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(key, mSku, 1, location, condition)
    item("IoType") = "ADJUST"
    payload = modRoleEventWriter.BuildPayloadJson(item)
    If Not modRoleEventWriter.QueueAdminInventoryAdjustEvent(mWh, "S1", "config-admin", _
            payload, "External fixture entity change", 0, eventId, report) Then Exit Function
    modProcessor.RunBatch mWh, 0, report, eventId, confirmed
    If Not confirmed Then Exit Function
    mInjectedEvents = 1
    ChangeExactFixtureEntity = True
End Function
Public Function DialogChoice(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle) As VbMsgBoxResult
    If (buttons And vbYesNo) = vbYesNo Then
        YieldAt "DeleteConfirmation"
        DialogChoice = vbYes
    Else
        DialogChoice = vbOK
    End If
End Function
Private Function InventorySignature() As String
    Dim wb As Workbook, candidate As Workbook, prior As Workbook, opened As Boolean
    Dim ws As Worksheet, lo As ListObject, cell As Range, result As String, path As String
    Set prior = Application.ActiveWorkbook
    path = mRoot & "\" & mWh & ".invSys.Data.Inventory.xlsb"
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, path, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
        opened = True
    End If
    On Error GoTo Failed
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            result = result & lo.Name & "|" & CStr(lo.ListRows.Count)
            For Each cell In lo.Range.Cells
                If IsError(cell.Value2) Then
                    result = result & "|Error"
                Else
                    result = result & "|" & CStr(Len(CStr(cell.Value2))) & ":" & CStr(cell.Value2)
                End If
            Next cell
        Next lo
    Next ws
    InventorySignature = result
CleanExit:
    If opened Then wb.Close SaveChanges:=False
    If Not prior Is Nothing Then prior.Activate
    Exit Function
Failed:
    If opened Then wb.Close SaveChanges:=False
    If Not prior Is Nothing Then prior.Activate
    Err.Raise vbObjectError + 2102, , "Inventory observation unavailable."
End Function
Public Function Facts() As String
    Dim unchanged As Boolean
    If mSignatureReady Then unchanged = (mBefore = InventorySignature())
    Facts = CStr(mPrepared) & "|" & CStr(mYield) & "|" & CStr(mHandler) & "|" & _
        CStr(mBound) & "|" & CStr(mSignatureReady And unchanged)
End Function
'@)
    if($AdminInventoryQuantityRetryOnly -or -not($AdminInventoryExactQuantityOnly -or $AdminInventoryManagedFieldsOnly)){
        $core=$packages['invSys.Core.xlam'].VBProject
        $quantity=$core.VBComponents.Item('modInventoryQuantityCommands').CodeModule
        $text=$quantity.Lines(1,$quantity.CountOfLines)
        $queueTail='(?m)^[\t ]*report = "Inventory adjustment could not be queued: " & queueError: Exit Function\r?\n[\t ]*End If'
        if([regex]::Matches($text,$queueTail).Count -ne 1){throw 'Exact quantity queue fault boundary changed; not product RED.'}
        $text=[regex]::Replace($text,$queueTail,('$0'+"`r`n"+'    TestInventoryQuantity.AfterQueue eventId'))
        $quantity.DeleteLines(1,$quantity.CountOfLines);$quantity.AddFromString($text)
        $quantityProbe=$core.VBComponents.Add(1);$quantityProbe.Name='TestInventoryQuantity'
        $quantityProbe.CodeModule.AddFromString(@'
Option Explicit
Private mFail As Boolean, mQueued As Long, mEvent As String
Public Sub Arm()
    mFail = True: mQueued = 0: mEvent = ""
End Sub
Public Sub AfterQueue(ByVal eventId As String)
    mQueued = mQueued + 1: mEvent = eventId
    If mFail Then
        mFail = False
        Err.Raise vbObjectError + 2213, , "Injected after confirmed queue and before processing."
    End If
End Sub
Public Function StageUnrelated(ByVal warehouse As String) As Boolean
    Dim report As String, ok As Boolean
    Arm
    ok = modInventoryQuantityCommands.SetSingleEntityQuantity(warehouse, "S1", "config-admin", _
        "DEMO-UNRELATED-QUANTITY", 11, "Unrelated pending fixture", report)
    StageUnrelated = (Not ok And mQueued = 1 And mEvent <> "")
End Function
Public Function QueuedCount() As Long
    QueuedCount = mQueued
End Function
Public Function LastEvent() As String
    LastEvent = mEvent
End Function
'@)
    }
    $form=$project.VBComponents.Item('frmAddInventoryItem').CodeModule
    $form.AddFromString(@'
Public Function InventoryCommandPrepareForTest(ByVal action As String, ByVal sku As String) As Boolean
    Dim i As Long
    If action = "Edit" Or action = "Delete" Then
        mBtnEditMode_Click
        For i = 0 To mCmbEditItem.ListCount - 1
            If CStr(mCmbEditItem.List(i, 1)) = sku Then
                mLoading = True: mCmbEditItem.ListIndex = i: mLoading = False
                mCmbEditItem_Change
                InventoryCommandPrepareForTest = (mSelectedEditSku = sku)
                Exit Function
            End If
        Next i
        Exit Function
    End If
    If action = "Add" Then
        mTxtItemName.Value = "Boundary fixture item"
        mCmbUom.Value = "LB": mCmbLocation.Value = "CLEARVIEW"
        mTxtQty.Value = "0"
    End If
    InventoryCommandPrepareForTest = True
End Function
Public Sub InventoryCommandInvokeForTest(ByVal action As String)
    Select Case action
        Case "Add", "Edit": mBtnOK_Click
        Case "Delete": mBtnDeleteItem_Click
        Case "Create": mBtnCreateInventoryTable_Click
        Case "Upload": mBtnUploadInventoryTable_Click
    End Select
End Sub
Public Sub InventoryCommandUomPromptForTest()
    mCmbUom.Value = ADD_UOM_OPTION
End Sub
Public Sub InventoryCommandTargetForTest(ByVal target As Double)
    mTxtQty.Value = CStr(target)
End Sub
Public Function InventoryCommandCustomFieldForTest(ByVal header As String) As Boolean
    Dim before As Long
    before = mLstCustomFields.ListCount
    mTxtCustomName.Value = header
    mTxtCustomValue.Value = "Fixture field value"
    mBtnAddField_Click
    InventoryCommandCustomFieldForTest = (mLstCustomFields.ListCount = before + 1)
End Function
Public Function InventoryCommandHeaderRefusalForTest() As Boolean
    InventoryCommandHeaderRefusalForTest = (InStr(1, mLastInventoryWorksheetReport, _
        "prohibited header", vbTextCompare) > 0)
End Function
Private Function InputBox(ByVal Prompt As String, Optional ByVal Title As String = "", _
                          Optional ByVal Default As String = "") As String
    InputBox = TestInventoryCommands.DialogAnswer(Prompt)
End Function
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                        Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestInventoryCommands.DialogChoice(Prompt, Buttons)
End Function
'@)
    $project.VBComponents.Item('modAdmin').CodeModule.AddFromString(@'
Public Sub InventoryBoundaryProtectionForTest(ByVal ws As Worksheet, ByVal protectAfter As Boolean)
    SetSheetProtectionAdminLocal ws, protectAfter
End Sub
'@)
}

function Get-InventoryBoundarySku($Fixture,[bool]$CountedOnly=$false) {
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$null;$owned=$false
    foreach($b in $excel.Workbooks){if($b.FullName -ieq $path){$book=$b;break}}
    if($null -eq $book){$book=$excel.Workbooks.Open($path,0,$true);$owned=$true}
    try {
        $catalog=Table $book 'tblSkuCatalog'
        foreach($row in $catalog.ListRows){
            $state=[string]$row.Range.Cells.Item(1,$catalog.ListColumns.Item('CATALOG_STATE').Index).Value2
            $sku=[string]$row.Range.Cells.Item(1,$catalog.ListColumns.Item('ITEM_CODE').Index).Value2
            if($CountedOnly -and [string]$row.Range.Cells.Item(1,$catalog.ListColumns.Item('TRACK_QTY').Index).Value2 -in @('FALSE','False','0')){continue}
            if($sku -and $state -cne 'RETIRED'){return $sku}
        }
        throw 'No generated active catalog item is available; not product RED.'
    }finally{if($owned){$book.Close($false)}}
}

function Test-InventoryCommandBoundaries($Fixture,$Other,[switch]$InputCasesOnly) {
    $cases=if($InputCasesOnly){@('Authorized','CatalogChanged')}else{@('Authorized','SignedOut','ReplacementSession','Reader','OtherWarehouseAdmin','CatalogChanged')}
    foreach($action in @('Add','Edit','Create','Upload','Delete')){
        foreach($case in $cases){
            SelectTarget $Fixture
            $hostBook=$excel.Workbooks.Add()
            $hostBook.SaveAs((Join-Path $runRoot ('inventory-boundary-'+$action+'-'+$case+'.xlsx')),51)
            try {
                $sku=Get-InventoryBoundarySku $Fixture
                if($action -eq 'Upload'){
                    # Fixture preparation uses the public staging helper; measured action uses the real form handler.
                    $ready=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
                    if(-not $ready){throw 'Upload staging preparation unavailable; not product RED.'}
                }
                $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
                $hostHash=Get-AdminInventoryEntryHash $hostBook.FullName
                $phase=if($action -eq 'Edit'){'EditReason'}elseif($action -eq 'Delete'){'DeleteConfirmation'}else{'ModalReturn'}
                # The other warehouse uses its own ephemeral fixture credential.
                $secret=if($case -ceq 'OtherWarehouseAdmin'){$Other.Secret}else{$Fixture.Secret}
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,$action,$case,$phase,$Fixture.Root,$Fixture.Warehouse,$secret))
                $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
                $id='InventoryCommand.'+$action+'.'+$phase+'.'+$case
                $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
                Check ($id+'.ActualHandlerAndYieldReached') $ready
                if(-not $ready){throw 'Actual Inventory action/yield fixture unavailable; not product RED.'}
                Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
                $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
                $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
                if($case -ceq 'Authorized'){
                    if($action -in @('Add','Upload')){
                        Check ($id+'.ExactZeroEntityPersisted') ($newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq 0 -and $after.AppliedRows -eq $before.AppliedRows+1)
                    }elseif($action -eq 'Create'){
                        Check ($id+'.OriginalOperatorStagingCreated') ([long](Run 'invSys.Admin.xlam' 'modAdminInventoryWorksheet.CountInventoryWorksheetTables' @($hostBook)) -eq 1)
                    }else{
                        Check ($id+'.OwningCommandChangesObserved') ($facts[4] -ceq 'False')
                    }
                }else{
                    Check ($id+'.NoInventoryMutationAfterYield') ($facts[4] -ceq 'True')
                    Check ($id+'.NoNewPersistedEvents') ($after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows)
                    Check ($id+'.NoOperatorSaveAfterYield') ($hostHash -ceq (Get-AdminInventoryEntryHash $hostBook.FullName))
                }
            }finally{$hostBook.Close($false)}
        }
    }
}

function Get-InventoryForeignFixtureFiles($Fixture) {
    $files=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File -Force){
        $files[$file.FullName.ToUpperInvariant()]=Get-AdminInventoryEntryHash $file.FullName
    }
    return $files
}

function Test-InventoryReplacementBoundaries($Fixture,$Other) {
    $specs=@(
        @('Add','ModalReturn','CatalogReplaced'),@('Edit','EditReason','CatalogReplaced'),
        @('Delete','DeleteReason','CatalogReplaced'),@('Edit','EditReason','EntityChanged'),
        @('Delete','DeleteReason','EntityChanged'),@('Add','ModalReturn','SourceFeedChanged'),
        @('Create','ModalReturn','SourceFeedChanged'),@('Upload','ModalReturn','TableReplaced'),
        @('Add','UomPrompt','SignedOut'),@('Add','UomPrompt','ReplacementSession'),
        @('Delete','DeleteReason','SignedOut'),@('Delete','DeleteReason','ReplacementSession'))
    foreach($spec in $specs){
        $action,$phase,$case=$spec
        SelectTarget $Fixture
        $hostBook=$excel.Workbooks.Add()
        $hostBook.SaveAs((Join-Path $runRoot ('replacement-'+$action+'-'+$phase+'-'+$case+'.xlsx')),51)
        try{
            $sku=Get-InventoryBoundarySku $Fixture ($case -ceq 'EntityChanged')
            if($action -ceq 'Upload'){
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))){throw 'Replacement table fixture unavailable; not product RED.'}
            }
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $hostHash=Get-AdminInventoryEntryHash $hostBook.FullName
            $foreignBefore=if($case -ceq 'SourceFeedChanged'){Get-InventoryForeignFixtureFiles $Other}else{$null}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,$action,$case,$phase,$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $injected=[long](Run 'invSys.Admin.xlam' 'TestInventoryCommands.InjectedEvents')
            $id='InventoryReplacement.'+$action+'.'+$phase+'.'+$case
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualHandlerAndYieldReached') $ready
            if(-not $ready){throw 'Replacement action or injected fixture change unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            Check ($id+'.NoOwnerInventoryMutation') ($facts[4] -ceq 'True')
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            Check ($id+'.NoAdditionalPersistedEvents') ($after.LogRows -eq $before.LogRows+$injected -and $after.AppliedRows -eq $before.AppliedRows+$injected)
            if($case -cne 'TableReplaced'){Check ($id+'.NoOperatorSaveAfterYield') ($hostHash -ceq (Get-AdminInventoryEntryHash $hostBook.FullName))}
            if($case -ceq 'SourceFeedChanged'){
                $foreignAfter=Get-InventoryForeignFixtureFiles $Other
                $changed=@(@($foreignBefore.Keys)+@($foreignAfter.Keys)|Select-Object -Unique|Where-Object {$foreignBefore[$_] -cne $foreignAfter[$_]})
                Check ($id+'.ForeignRuntimeFilesPreserved') ($changed.Count -eq 0)
            }
        }finally{
            [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($Fixture.Root))
            $hostBook.Close($false)
        }
    }
}

function Test-InventoryExactQuantity($Fixture,$Other,[switch]$RetryQueued) {
    $prefix=if($RetryQueued){'InventoryQuantityRetry'}else{'InventoryExactQuantity'}
    SelectTarget $Fixture
    $sku=if($RetryQueued){'DEMO-EXACT-RETRY'}else{'DEMO-EXACT-QUANTITY'}
    $csv=Join-Path $runRoot ($prefix+'-fixture.csv')
    $fixtureRows=@('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION',($sku+',Exact quantity fixture,LB,CLEARVIEW,10,GOOD'))
    if($RetryQueued){$fixtureRows+='DEMO-UNRELATED-QUANTITY,Unrelated pending fixture,LB,CLEARVIEW,10,GOOD'}
    $fixtureRows|Set-Content -LiteralPath $csv -Encoding ASCII
    $created=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))
    Check ($prefix+'.GeneratedSingleEntityFixtureApplied') $created
    if(-not $created){throw 'Owning single-entity fixture creation failed; not product RED.'}
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$excel.Workbooks.Open($path,0,$true)
    try{
        $table=Table $book 'tblInventoryEntities';$choices=@()
        foreach($row in $table.ListRows){
            if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index).Value2 -ceq $sku -and [string]$row.Range.Cells.Item(1,$table.ListColumns.Item('InventoryState').Index).Value2 -ceq 'ACTIVE'){
                $choices+=[pscustomobject]@{Key=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2;Quantity=[double]$row.Range.Cells.Item(1,$table.ListColumns.Item('QtyOnHand').Index).Value2}
            }
        }
    }finally{$book.Close($false)}
    if($choices.Count -ne 1){throw 'Exact quantity test requires one generated active counted entity; not product RED.'}
    $choice=$choices[0];$target=$choice.Quantity+3
    $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ($prefix+'-host.xlsx')),51)
    try{
        $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
        [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
        [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetQuantityTarget' @($target))
        if($RetryQueued){
            $unrelated=[bool](Run 'invSys.Core.xlam' 'TestInventoryQuantity.StageUnrelated' @($Fixture.Warehouse))
            $unrelatedEvent=[string](Run 'invSys.Core.xlam' 'TestInventoryQuantity.LastEvent')
            Check ($prefix+'.UnrelatedAdjustmentQueued') ($unrelated -and $unrelatedEvent -ne '')
            if(-not $unrelated -or $unrelatedEvent -eq ''){throw 'Unrelated pending fixture unavailable; not product RED.'}
            [void](Run 'invSys.Core.xlam' 'TestInventoryQuantity.Arm')
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Edit','SetQuantity','EditReason',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $first=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $firstReady=$first.Count -eq 5 -and $first[0] -ceq 'True' -and $first[1] -ceq 'True' -and $first[2] -ceq '1'
            Check ($prefix+'.FirstEditAndReasonReached') $firstReady
            if(-not $firstReady){throw 'Queued retry first actual action unavailable; not product RED.'}
            Check ($prefix+'.FirstOperatorCaptured') ($first[3] -ceq 'True')
            $firstEvent=[string](Run 'invSys.Core.xlam' 'TestInventoryQuantity.LastEvent')
            $description=[string](Run 'invSys.Core.xlam' 'modRoleEventWriter.DescribeLocalStagedInboxRows' @('ADMIN_INVENTORY_ADJUST',$Fixture.Warehouse,'S1'))
            if($description -notmatch '^Path=(.*?); Files='){throw 'Owned queue path unavailable; not product RED.'}
            $queueRoot=$Matches[1]
            if((Split-Path $queueRoot -Leaf) -cne 'S1' -or (Split-Path (Split-Path $queueRoot -Parent) -Leaf) -cne $Fixture.Warehouse){throw 'Queue observation escaped the generated fixture identity.'}
            $found=0
            foreach($file in @(Get-ChildItem -LiteralPath $queueRoot -File -Filter '*.staging.jsonl' -ErrorAction SilentlyContinue)){
                foreach($line in Get-Content -LiteralPath $file.FullName){
                    if($line.Trim()){$queued=$line|ConvertFrom-Json;if($queued.EventID -ceq $firstEvent){$found++}}
                }
            }
            Check ($prefix+'.FirstExactQueuedEventSaved') ($firstEvent -ne '' -and $found -eq 1)
            if($firstEvent -eq '' -or $found -ne 1){throw 'Post-queue fault did not reach saved event; not product RED.'}
            $partial=Get-AdminInventoryEntryFacts $Fixture -Persisted
            Check ($prefix+'.FirstQueueRemainsUnapplied') ($partial.EntityQuantities[$choice.Key] -eq $choice.Quantity -and $partial.AppliedRows -eq $before.AppliedRows)
            Check ($prefix+'.PartialOutcomeReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PartialQueueOutcomeReported'))
            SelectTarget $Fixture
            $hostBook.Activate()
        }
        [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Edit','SetQuantity','EditReason',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
        $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
        $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
        Check ($prefix+'.ActualEditAndReasonReached') $ready
        if(-not $ready){throw 'Exact quantity modal fixture unavailable; not product RED.'}
        Check ($prefix+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
        $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
        Check ($prefix+'.ExactSelectedEntityTargetSaved') ($after.EntityQuantities.ContainsKey($choice.Key) -and $after.EntityQuantities[$choice.Key] -eq $target)
        $keys=@(@($before.EntityQuantities.Keys)+@($after.EntityQuantities.Keys)|Select-Object -Unique)
        Check ($prefix+'.EntityIdentitySetPreserved') ($keys.Count -eq $before.EntityQuantities.Count -and $keys.Count -eq $after.EntityQuantities.Count)
        $otherChanged=@($keys|Where-Object {$_ -cne $choice.Key -and $before.EntityQuantities[$_] -ne $after.EntityQuantities[$_]})
        Check ($prefix+'.OtherEntityQuantitiesPreserved') ($otherChanged.Count -eq 0)
        if($RetryQueued){
            Check ($prefix+'.OnlyOneAdjustmentQueued') ([int](Run 'invSys.Core.xlam' 'TestInventoryQuantity.QueuedCount') -eq 1)
            Check ($prefix+'.OnlyOneAdjustmentApplied') ($after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1)
            Check ($prefix+'.OriginalExactQueuedEventApplied') ([bool](Run 'invSys.Core.xlam' 'modInventoryDomainBridge.InventoryEventAppliedBridge' @($Fixture.Warehouse,$firstEvent)))
            Check ($prefix+'.UnrelatedEventRemainsUnapplied') (-not [bool](Run 'invSys.Core.xlam' 'modInventoryDomainBridge.InventoryEventAppliedBridge' @($Fixture.Warehouse,$unrelatedEvent)))
        }
    }finally{$hostBook.Close($false)}
}

function Test-InventoryManagedFields($Fixture,$Other) {
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    Check 'InventoryManagedField.GeneratedFixtureSeedApplied' $seed.StartsWith('OK|')
    if(-not $seed.StartsWith('OK|')){throw 'Managed field fixture unavailable; not product RED.'}
    # Seed's owned open workbook may retain unsaved bookkeeping. Admission must
    # still refuse dirty books. Verify persisted business fixture data, discard
    # only that owned in-memory copy, and never save on the owner's behalf.
    $memory=Get-AdminInventoryEntryFacts $Fixture
    $cold=Get-AdminInventoryEntryFacts $Fixture -Persisted
    $same=$memory.LogRows -eq $cold.LogRows -and $memory.AppliedRows -eq $cold.AppliedRows -and $memory.EntityQuantities.Count -eq $cold.EntityQuantities.Count -and $memory.FileHash -ceq $cold.FileHash
    foreach($key in $memory.EntityQuantities.Keys){
        if(-not $cold.EntityQuantities.ContainsKey($key) -or $memory.EntityQuantities[$key] -ne $cold.EntityQuantities[$key]){$same=$false}
    }
    [pscustomobject]@{GeneratedFixtureOnly=$true;MemoryAndPersistedEntitiesMatch=$same;ClosedWithoutSave=$true;FileBytesPreserved=($memory.FileHash -ceq $cold.FileHash)}|
        ConvertTo-Json -Compress|Set-Content (Join-Path $reportRoot 'managed-fixture-cold-proof.json')
    if(-not $same){throw 'Generated managed-field seed is not independently persisted; not product RED.'}
    foreach($spec in @(@('SystemKey','System_Key',$true),@('Condition','Condition',$true),@('LegacyRow','ROW',$true),@('Ordinary','Fixture note',$false))){
        $label,$header,$reserved=$spec
        SelectTarget $Fixture
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('managed-modal-'+$label+'.xlsx')),51)
        try{
            if($AdminInventorySourceOverlayForTest -and $AdminInventoryManagedFieldsOnly){
                $books=@()
                foreach($kind in @('Config','Auth','Data.Inventory')){
                    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.'+$kind+'.xlsb')
                    $found=$null
                    foreach($candidate in $excel.Workbooks){if($candidate.FullName -ieq $path){$found=$candidate;break}}
                    $books+=[pscustomobject]@{Kind=$kind;Open=($null -ne $found);Saved=if($null -ne $found){[bool]$found.Saved}else{$null};ReadOnly=if($null -ne $found){[bool]$found.ReadOnly}else{$null}}
                }
                $admission=([string](Run 'invSys.Core.xlam' 'TestInventoryAdmission.Facts' @($hostBook.Name,$Fixture.Warehouse))).Split('|')
                if($admission.Count -ne 3){throw 'Fixed admission facts unavailable.'}
                [pscustomobject]@{Case=$label;HostSaved=[bool]$hostBook.Saved;SignedIn=($admission[0] -ceq 'True');InputSourceBound=($admission[1] -ceq 'True');CommandCaptured=($admission[2] -ceq 'True');Books=$books}|
                    ConvertTo-Json -Depth 4 -Compress|Add-Content (Join-Path $reportRoot 'managed-admission-facts.jsonl')
            }
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetCustomHeader' @($header))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'CustomField','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $accepted=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.CustomFieldAccepted')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            $id='InventoryManagedField.Modal.'+$label
            Check ($id+'.ActualFieldHandlerReached') $ready
            if(-not $ready){throw 'Managed field modal action unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            Check ($id+'.ExpectedFieldDisposition') ($accepted -eq (-not $reserved))
            Check ($id+'.CancelledWithoutInventoryMutation') ($facts[4] -ceq 'True')
        }finally{$hostBook.Close($false)}
    }
    foreach($header in @('System_Key','Condition')){
        SelectTarget $Fixture
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('managed-worksheet-'+$header+'.xlsx')),51)
        try{
            $ready=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
            if($ready){$ready=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.AddManagedFixtureColumn' @($hostBook.Name,$header))}
            if(-not $ready){throw 'Managed header staging fixture unavailable; not product RED.'}
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            $id='InventoryManagedField.Worksheet.'+$header
            Check ($id+'.ActualUploadHandlerReached') $ready
            if(-not $ready){throw 'Managed header Upload unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            Check ($id+'.HeaderRejectedAtPreflight') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.ManagedHeaderRefused'))
            Check ($id+'.NoInventoryMutation') ($facts[4] -ceq 'True')
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            Check ($id+'.NoPersistedEvents') ($after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows)
        }finally{$hostBook.Close($false)}
    }
}
