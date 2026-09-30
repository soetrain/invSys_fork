# D13: invoke existing form handlers and observe actual queue-return facts.
# All adapters are installed only into unsaved disposable package instances.
function Install-ProcessWorksheetActivityProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub WorksheetActivityStageForTest(ByVal value As String)
    ClearProcessDraft True
    mTxtProcessName.Text = value
    mTxtProcessDescription.Text = value
End Sub
Public Sub WorksheetActivityActForTest(ByVal action As String)
    Select Case action
        Case "SEND": mBtnProcessWorksheetCreate_Click
        Case "ADD_ITEM": mBtnProcessWorksheetAddAlternative_Click
        Case "RETRIEVE": mBtnProcessWorksheetRetrieve_Click
        Case Else: Err.Raise 5, , "Unsupported test action."
    End Select
End Sub
Public Function WorksheetActivityCaptionForTest(ByVal action As String) As String
    Select Case action
        Case "SEND": WorksheetActivityCaptionForTest = mBtnProcessWorksheetCreate.Caption
        Case "ADD_ITEM": WorksheetActivityCaptionForTest = mBtnProcessWorksheetAddAlternative.Caption
        Case "RETRIEVE": WorksheetActivityCaptionForTest = mBtnProcessWorksheetRetrieve.Caption
    End Select
End Function
'@)
    $module=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $module.InsertLines(1,'Private mWorksheetSubmissionFacts As String')
    $module.AddFromString(@'
Public Sub WorksheetActivityStage(ByVal value As String)
    mForm.WorksheetActivityStageForTest value
End Sub
Public Sub WorksheetActivityAct(ByVal action As String)
    mWorksheetSubmissionFacts = ""
    mForm.WorksheetActivityActForTest action
End Sub
Public Function WorksheetActivityCaption(ByVal action As String) As String
    WorksheetActivityCaption = mForm.WorksheetActivityCaptionForTest(action)
End Function
Public Sub WorksheetObserveSubmissionForTest(ByVal id As String, ByVal submitted As Boolean, ByVal attempted As Boolean)
    If Not submitted And Not attempted Then Exit Sub
    If mWorksheetSubmissionFacts <> "" Then mWorksheetSubmissionFacts = mWorksheetSubmissionFacts & vbLf
    mWorksheetSubmissionFacts = mWorksheetSubmissionFacts & id & "|" & IIf(submitted, "Submitted", "Unknown")
End Sub
Public Function WorksheetSubmissionFactsForTest() As String
    WorksheetSubmissionFactsForTest = mWorksheetSubmissionFacts
End Function
Public Function WorksheetActivitySelect(ByVal workbookName As String, ByVal first As Long, ByVal last As Long, ByVal fill As Boolean) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, selection As Range
    Dim index As Long, report As String
    Set wb = Application.Workbooks(workbookName)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If Left$(lo.Name, 15) = "invSys_Process_" Then
                index = index + 1
                If index >= first And index <= last Then
                    If fill Then
                        If Not modProductionProcessWorksheet.PopulateFormulationExampleForTest(wb, lo.Name, True, report) Then Exit Function
                    End If
                    If selection Is Nothing Then Set selection = lo.DataBodyRange.Cells(1, 1) Else Set selection = Application.Union(selection, lo.DataBodyRange.Cells(1, 1))
                End If
            End If
        Next lo
    Next ws
    If selection Is Nothing Then Exit Function
    wb.Activate: selection.Worksheet.Activate: selection.Select
    WorksheetActivitySelect = True
End Function
'@)
    # Observe real queue-return IDs, including before runtime passes typed facts.
    # No substitution, fabricated ID, report parsing or error-handler interception.
    $owner=$project.VBComponents.Item('modProductionReusableDesigns').CodeModule
    $start=$owner.ProcStartLine('SubmitReusableDesignEvent',0);$count=$owner.ProcCountLines('SubmitReusableDesignEvent',0)
    $lines=$owner.Lines($start,$count) -split '\r?\n'
    $anchors=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Trim() -ceq 'If Not facts Is Nothing Then facts.ObserveSubmission eventId, submitted, writeAttempted'){$start+$i}})
    if($anchors.Count -ne 2){throw 'Actual submission observation boundary unavailable; not product RED.'}
    $owner.InsertLines($anchors[0]+1,'    TestProductionDesigner.WorksheetObserveSubmissionForTest eventId, submitted, writeAttempted')
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function WorksheetPolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    WorksheetPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function WorksheetTerminalForTest(ByVal wire As String) As Boolean
    WorksheetTerminalForTest = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
'@)
}

function Test-ProcessWorksheetActivity($Fixture) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function Tables {@(foreach($sheet in $book.Worksheets){foreach($table in $sheet.ListObjects){if($table.Name -like 'invSys_Process_*'){$table}}})}
    function InventoryState {
        $source=$null;$owned=$false;$path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        foreach($open in $excel.Workbooks){if([string]$open.FullName -ieq $path){$source=$open;break}}
        if($null -eq $source){$source=$excel.Workbooks.Open($path,0,$true);$owned=$true}
        try{
            $values=@(foreach($sheet in $source.Worksheets){foreach($table in $sheet.ListObjects){
                if([string]$table.Name -cin @('tblInventoryLog','tblAppliedEvents','tblInventoryEntities','tblSkuBalance','tblLocationBalance','tblSkuCatalog')){[pscustomobject]@{Name=[string]$table.Name;Values=$table.Range.Value2}}
            }})
            if($values.Count -ne 6){throw 'Six Inventory business tables required.'}
            return ($values|Sort-Object Name|ConvertTo-Json -Depth 8 -Compress)
        }finally{if($owned){$source.Close($false)}}
    }
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Label,[int]$SourceCount=0){
        $id='PRODUCTION_PROCESS_WORKSHEET_'+$Action
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $rows=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $attempt=@($rows|Where-Object OutcomeCode -CEQ REQUESTED);$result=@($rows|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $result.Count -eq 1
        $context=$paired;$safe=$paired;$integrity=$paired;$linked=$false;$effect=$false;$sources=$false;$terminal=$false
        foreach($row in $rows){$context=$context -and $row.ControlId -ceq $id -and $row.OwnerId -ceq 'PRODUCTION_PROCESS_WORKSHEET' -and $row.UserId -ceq 'config-producer' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.StationId -ceq 'S1' -and $row.CatalogVersion -eq 22}
        foreach($text in $raw){
            $decoded=($text|ConvertFrom-Json)|ConvertTo-Json -Depth 25 -Compress
            foreach($value in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','PayloadJson')){if($decoded.Contains($value) -or $decoded.Contains(($value|ConvertTo-Json -Compress).Trim('"'))){$safe=$false}}
            $match=[regex]::Match($text,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$digest=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $digest -ceq $match.Groups[1].Value
        }
        $observed=@(([string](Probe 'WorksheetSubmissionFactsForTest')).Split([char]10)|Where-Object{$_})
        Check ($Label+'.ActualOwnerSubmissionCount') ($observed.Count -eq $SourceCount)
        if($paired){
            $linked=$attempt[0].ActivityId -cne '' -and $attempt[0].ActivityId -ceq $result[0].ActivityId -and $attempt[0].RecordId -cne $result[0].RecordId
            $severity=if($Outcome -ceq 'REJECTED'){'Warning'}else{'Info'};$dataEffect=if($Outcome -ceq 'CONFIRMED'){'Unknown'}else{'Unchanged'}
            $effect=$attempt[0].EventCode -ceq ($id+'_REQUESTED') -and $attempt[0].Severity -ceq 'Info' -and $attempt[0].DataEffect -ceq 'Unknown' -and $result[0].EventCode -ceq ($id+'_'+$Outcome) -and $result[0].Severity -ceq $severity -and $result[0].DataEffect -ceq $dataEffect
            $actual=@($result[0].SourceEventRefs|ForEach-Object{$_.EventId+'|'+$_.SubmissionState})
            $sources=@($attempt[0].SourceEventRefs).Count -eq 0 -and $actual.Count -eq $SourceCount -and ($actual -join "`n") -ceq ($observed -join "`n")
            foreach($reference in $result[0].SourceEventRefs){$sources=$sources -and $reference.SourceKind -ceq 'Designs' -and $reference.WarehouseId -ceq $Fixture.Warehouse -and @($reference.PSObject.Properties).Count -eq 4}
            $terminal=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.WorksheetTerminalForTest' @(($attempt[0]|ConvertTo-Json -Depth 20 -Compress))) -and ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.WorksheetTerminalForTest' @(($result[0]|ConvertTo-Json -Depth 20 -Compress))) -eq ($Outcome -in @('STAGED','CONFIRMED')))
        }
        Check ($Label+'.OneAttemptAndResult') $paired
        Check ($Label+'.ExactCapturedContext') $context
        Check ($Label+'.CorrelatedDistinctRecords') $linked
        Check ($Label+'.ExactOutcomeMeaning') $effect
        Check ($Label+'.ExactOwnerReferencesOnly') $sources
        Check ($Label+'.NoEnteredDataOrSecrets') $safe
        Check ($Label+'.ContentIntegrity') $integrity
        Check ($Label+'.ExplicitTerminalFact') $terminal
    }
    $canary='WORKSHEET'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.WorksheetPolicyForTest' @($true))){throw 'Authorized tracking policy setup unavailable; not product RED.'}
    SelectTarget $Fixture 'config-producer'
    $pins=@{};foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -File -Filter '*.xlsb'){$pins[$file.FullName]=Hash $file.FullName}
    $inventoryBefore=InventoryState
    try{
        $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(21))).Split([char]10)|Where-Object{$_})
        $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(22))).Split([char]10)|Where-Object{$_})
        $ids=@('SEND','ADD_ITEM','RETRIEVE')|ForEach-Object{'PRODUCTION_PROCESS_WORKSHEET_'+$_}
        Check 'ProcessWorksheet.Catalog22Extends21' ($old.Count -eq 106 -and $new.Count -eq 109 -and @($new|Sort-Object -Unique).Count -eq 109 -and (@($new|Where-Object{$_ -cnotin $old}) -join '|') -ceq ($ids -join '|'))
        foreach($id in $old){$definition=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,21));Check ('ProcessWorksheet.Catalog.Preserve.'+$id) ($definition -cne '' -and $definition -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,22)))}
        $book=$excel.Workbooks.Add();$book.Worksheets.Item(1).Range('A1').Value2=$canary
        $path=Join-Path $runRoot 'worksheet-activity-operator.xlsb';$book.SaveAs($path,50)
        $decoy=$excel.Workbooks.Add();$decoy.Worksheets.Item(1).Range('A1').Value2=$canary
        $before=@(Files);[void](Probe 'OpenDesigner' @($book.Name))
        Check 'ProcessWorksheet.Initialize.NoAction' (@(Files).Count -eq $before.Count)
        foreach($action in @('SEND','ADD_ITEM','RETRIEVE')){
            $id='PRODUCTION_PROCESS_WORKSHEET_'+$action
            $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,22));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
            Check ('ProcessWorksheet.Metadata.'+$action) ($null -ne $definition -and $definition.Caption -ceq [string](Probe 'WorksheetActivityCaption' @($action)) -and $definition.Surface -ceq 'Operations > Production > Process Designer' -and $definition.OwnerId -ceq 'PRODUCTION_PROCESS_WORKSHEET' -and $definition.Class -ceq 'Command' -and $definition.Role -ceq 'Production' -and $definition.Capability -ceq 'PROD_POST')
            Check ('ProcessWorksheet.OlderCatalog.Excludes.'+$action) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,21)) -ceq '')
        }
        for($index=1;$index -le 2;$index++){
            [void](Probe 'WorksheetActivityStage' @($canary));$before=@(Files);$decoy.Activate();[void](Probe 'WorksheetActivityAct' @('SEND'))
            Check ('ProcessWorksheet.Send'+$index+'.CapturedLocalSave') (@(Tables).Count -eq $index -and $book.Saved -and $decoy.Worksheets.Count -eq 1 -and $decoy.Worksheets.Item(1).ListObjects.Count -eq 0)
            Pair $before 'SEND' 'STAGED' ('ProcessWorksheet.Send'+$index)
        }
        if(-not [bool](Probe 'WorksheetActivitySelect' @($book.Name,1,1,$false))){throw 'Selected table fixture unavailable.'}
        $table=@(Tables)[0];$columns=$table.ListColumns.Count;$before=@(Files);[void](Probe 'WorksheetActivityAct' @('ADD_ITEM'))
        Check 'ProcessWorksheet.Add.ActualPairAndSave' ($table.ListColumns.Count -eq $columns+2 -and $book.Saved)
        Pair $before 'ADD_ITEM' 'STAGED' 'ProcessWorksheet.Add'
        $before=@(Files);[void](Probe 'WorksheetActivityAct' @('RETRIEVE'))
        Check 'ProcessWorksheet.Rejected.TablesRemain' (@(Tables).Count -eq 2)
        Pair $before 'RETRIEVE' 'REJECTED' 'ProcessWorksheet.Rejected'
        Check 'ProcessWorksheet.LocalActions.CanonicalAuthorityPreserved' (@($pins.Keys|Where-Object{(Hash $_) -cne $pins[$_]}).Count -eq 0)
        if(-not [bool](Probe 'WorksheetActivitySelect' @($book.Name,1,1,$true))){throw 'Valid selected table fixture unavailable.'}
        $before=@(Files);[void](Probe 'WorksheetActivityAct' @('RETRIEVE'))
        Check 'ProcessWorksheet.Single.OnlySelectedRemoved' (@(Tables).Count -eq 1 -and $book.Saved)
        Pair $before 'RETRIEVE' 'CONFIRMED' 'ProcessWorksheet.Single' 1
        [void](Probe 'WorksheetActivityStage' @($canary));[void](Probe 'WorksheetActivityAct' @('SEND'))
        if(-not [bool](Probe 'WorksheetActivitySelect' @($book.Name,1,2,$true))){throw 'Multiple selected tables fixture unavailable.'}
        $before=@(Files);[void](Probe 'WorksheetActivityAct' @('RETRIEVE'))
        Check 'ProcessWorksheet.Multiple.AllSelectedRemoved' (@(Tables).Count -eq 0 -and $book.Saved)
        Pair $before 'RETRIEVE' 'CONFIRMED' 'ProcessWorksheet.Multiple' 2
        Check 'ProcessWorksheet.UnknownWorkbookDataPreserved' ($book.Worksheets.Item(1).Range('A1').Value2 -ceq $canary -and $decoy.Worksheets.Item(1).Range('A1').Value2 -ceq $canary)
        Check 'ProcessWorksheet.InventoryBusinessStatePreserved' ((InventoryState) -ceq $inventoryBefore)
        Check 'ProcessWorksheet.AuthConfigBytesPreserved' (@($pins.Keys|Where-Object{($_ -like '*.Auth.xlsb' -or $_ -like '*.Config.xlsb') -and (Hash $_) -cne $pins[$_]}).Count -eq 0)
        [void](Probe 'WorksheetHeadersShow');CaptureOwnedFormByCaptionEvidence 'Production' 'process-worksheet-actual-retrieve.png'
    }finally{
        try{[void](Probe 'CloseDesigner')}catch{}
        if($null -ne $decoy){$decoy.Close($false)}
        if($null -ne $book){$book.Close($false)}
    }
}
