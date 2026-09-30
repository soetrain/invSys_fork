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
Public Function WorksheetActivityGuardForTest(ByVal action As String, ByVal guard As String) As Boolean
    Dim priorLoading As Boolean, priorBusy As Boolean
    TestProductionDesigner.WorksheetGuardEnteredForTest = True
    priorLoading = mLoading: priorBusy = mDesignerActionInProgress
    On Error GoTo Restore
    If guard = "Loading" Then mLoading = True
    If guard = "Nested" Then mDesignerActionInProgress = True
    WorksheetActivityActForTest action
    WorksheetActivityGuardForTest = True
Restore:
    mLoading = priorLoading: mDesignerActionInProgress = priorBusy
End Function
Public Function WorksheetActivityNoticeForTest() As Boolean
    WorksheetActivityNoticeForTest = (InStr(1, mTxtStatus.Text, "Tracking unavailable", vbTextCompare) > 0)
End Function
'@)
    $module=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $module.InsertLines(1,'Private mWorksheetSubmissionFacts As String')
    $module.InsertLines(1,"Public WorksheetGuardEnteredForTest As Boolean`r`nPrivate mWorksheetGuardError As Long")
    $module.InsertLines(1,"Private mWorksheetCapturedBook As Workbook`r`nPrivate mWorksheetCapturedContext As String")
    $module.InsertLines(1,"Private mWorksheetFault As String`r`nPrivate mWorksheetQueueCount As Long`r`nPrivate mWorksheetDeleteCount As Long`r`nPrivate mWorksheetFaultHits As Long")
    $module.AddFromString(@'
Public Sub WorksheetActivityStage(ByVal value As String)
    mForm.WorksheetActivityStageForTest value
End Sub
Public Sub WorksheetActivityAct(ByVal action As String)
    mWorksheetSubmissionFacts = ""
    mWorksheetQueueCount = 0: mWorksheetDeleteCount = 0: mWorksheetFaultHits = 0
    mForm.WorksheetActivityActForTest action
End Sub
Public Function WorksheetActivityGuard(ByVal action As String, ByVal guard As String) As Boolean
    On Error GoTo Failed
    mWorksheetSubmissionFacts = ""
    WorksheetGuardEnteredForTest = False: mWorksheetGuardError = 0
    WorksheetActivityGuard = mForm.WorksheetActivityGuardForTest(action, guard)
    Exit Function
Failed:
    mWorksheetGuardError = Err.Number
End Function
Public Function WorksheetGuardStatusForTest() As String
    WorksheetGuardStatusForTest = CStr(WorksheetGuardEnteredForTest) & "|" & CStr(mWorksheetGuardError)
End Function
Public Sub WorksheetSafeCloseForTest()
    On Error Resume Next
    Unload mForm: Set mForm = Nothing
    Set mWorksheetCapturedBook = Nothing
End Sub
Public Sub WorksheetCaptureBindingForTest(ByVal workbookName As String)
    Set mWorksheetCapturedBook = Application.Workbooks(workbookName)
    mWorksheetCapturedContext = modActivity.CaptureContext()
    mWorksheetSubmissionFacts = ""
End Sub
Public Function WorksheetBindingIsCurrentForTest() As Boolean
    WorksheetBindingIsCurrentForTest = modProductionDesignerActions.ContextIsCurrent(mWorksheetCapturedContext, mWorksheetCapturedBook)
End Function
Public Function WorksheetActivityNotice() As Boolean
    WorksheetActivityNotice = mForm.WorksheetActivityNoticeForTest()
End Function
Public Sub WorksheetFaultForTest(ByVal mode As String)
    mWorksheetFault = mode
End Sub
Public Function WorksheetFaultHitsForTest() As Long
    WorksheetFaultHitsForTest = mWorksheetFaultHits
End Function
Public Function WorksheetUncertainAckForTest() As Boolean
    mWorksheetQueueCount = mWorksheetQueueCount + 1
    If mWorksheetFault = "SignOutAfterFirst" And mWorksheetQueueCount = 1 Then
        mWorksheetFaultHits = mWorksheetFaultHits + 1
        modAuth.SignOut
    End If
    If mWorksheetFault = "UncertainSecond" And mWorksheetQueueCount = 2 Then
        WorksheetUncertainAckForTest = True
        mWorksheetFaultHits = mWorksheetFaultHits + 1
    End If
End Function
Public Function WorksheetQueueReturnsForTest() As Long
    WorksheetQueueReturnsForTest = mWorksheetQueueCount
End Function
Public Function WorksheetRemovalFailureForTest(ByVal afterDelete As Boolean) As Boolean
    If Not afterDelete Then mWorksheetDeleteCount = mWorksheetDeleteCount + 1
    If mWorksheetDeleteCount <> 2 Then Exit Function
    If (mWorksheetFault = "RemoveSecond" And Not afterDelete) Or _
       (mWorksheetFault = "SaveSecond" And afterDelete) Then
        WorksheetRemovalFailureForTest = True
        mWorksheetFaultHits = mWorksheetFaultHits + 1
    End If
End Function
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
    # The actual write occurs first. Simulate loss of its acknowledgment, not a
    # fabricated event or an unhandled Domain Err.Raise/VBE interruption.
    $owner.InsertLines($anchors[0],'    If TestProductionDesigner.WorksheetUncertainAckForTest() Then submitted = False')
    $owner=$project.VBComponents.Item('modProductionProcessWorksheet').CodeModule
    $start=$owner.ProcStartLine('DeleteProcessWorksheetTable',0);$count=$owner.ProcCountLines('DeleteProcessWorksheetTable',0)
    $lines=$owner.Lines($start,$count) -split '\r?\n'
    $before=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Trim() -ceq 'Set lo = FindProcessTableByName(wb, tableName)'){$start+$i}})
    $after=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Trim() -ceq 'wb.Save'){$start+$i}})
    if($before.Count -ne 1 -or $after.Count -ne 1){throw 'Actual removal boundaries unavailable; not product RED.'}
    # Use the existing failure return at two separate stages. SaveSecond reaches
    # real lo.Delete, then simulates a save refusal; no real disk fault is claimed.
    $owner.InsertLines($after[0],'    If TestProductionDesigner.WorksheetRemovalFailureForTest(True) Then GoTo Failed')
    $owner.InsertLines($before[0],'    If TestProductionDesigner.WorksheetRemovalFailureForTest(False) Then GoTo Failed')
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
Public Function WorksheetCanProduceForTest() As Boolean
    Dim produce As Boolean, maintain As Boolean
    produce = modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST")
    maintain = modRoleUiAccess.CanCurrentUserPerformCapability("ADMIN_MAINT")
    WorksheetCanProduceForTest = produce Or maintain
End Function
'@)
}

function Test-ProcessWorksheetCatalog {
    $submitted='[{"WarehouseId":"CATALOG_TEST","SourceKind":"Designs","EventId":"Source_A","SubmissionState":"Submitted"}]'
    $unknown=$submitted.Replace('Submitted','Unknown')
    $mixed=$submitted.Substring(0,$submitted.Length-1)+','+$unknown.Substring(1).Replace('Source_A','Source_B')
    $invalid=[ordered]@{
        Duplicate=$submitted.Substring(0,$submitted.Length-1)+','+$submitted.Substring(1)
        CrossWarehouse=$submitted.Replace('CATALOG_TEST','ANOTHER_TEST')
        Inventory=$submitted.Replace('Designs','Inventory')
        InvalidIdentity=$submitted.Replace('Source_A','Source A')
        OversizedIdentity=$submitted.Replace('Source_A',('A'*129))
        MissingIdentity=$submitted.Replace('"EventId":"Source_A",','')
        NumericIdentity=$submitted.Replace('"Source_A"','123')
        ExtraField=$submitted.Replace('"EventId":','"Extra":"forbidden","EventId":')
        UnknownState=$submitted.Replace('Submitted','Applied')
    }
    foreach($action in @('SEND','ADD_ITEM','RETRIEVE')){
        $id='PRODUCTION_PROCESS_WORKSHEET_'+$action;$prefix='ProcessWorksheet.Wire.'+$action
        $outcomes=[ordered]@{REQUESTED=@('Info','Unknown');DENIED=@('Blocked','Unchanged');REJECTED=@('Warning','Unchanged');FAILED=@('Error','Unknown')}
        if($action -ceq 'RETRIEVE'){$outcomes.CONFIRMED=@('Info','Unknown')}else{$outcomes.STAGED=@('Info','Unchanged')}
        foreach($code in @('REQUESTED','DENIED','REJECTED','FAILED','STAGED','CONFIRMED','PENDING','APPLIED','COMPLETED','VALIDATED','CANCELLED')){
            $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$code));$value=if($wire){$wire|ConvertFrom-Json}else{$null}
            $supported=$outcomes.Contains($code)
            $correct=if($supported){$null -ne $value -and $value.EventCode -ceq ($id+'_'+$code) -and $value.OutcomeCode -ceq $code -and $value.Severity -ceq $outcomes[$code][0] -and $value.DataEffect -ceq $outcomes[$code][1] -and $value.UserMessage -ne ''}else{$wire -ceq ''}
            Check ($prefix+'.Outcome.'+$code) $correct
            $acceptEmpty=$supported -and $code -cne 'CONFIRMED'
            Check ($prefix+'.EmptyReferences.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,'[]')) -eq $acceptEmpty)
            foreach($state in @('Submitted','Unknown')){
                $reference=if($state -ceq 'Submitted'){$submitted}else{$unknown}
                $accept=$action -ceq 'RETRIEVE' -and ($code -ceq 'FAILED' -or ($code -ceq 'CONFIRMED' -and $state -ceq 'Submitted'))
                Check ($prefix+'.'+$state+'References.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,$reference)) -eq $accept)
            }
        }
    }
    foreach($case in $invalid.Keys){Check ('ProcessWorksheet.Wire.RETRIEVE.Reject.'+$case) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('PRODUCTION_PROCESS_WORKSHEET_RETRIEVE','FAILED',$invalid[$case])))}
    Check 'ProcessWorksheet.Wire.RETRIEVE.FailedMixed' ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('PRODUCTION_PROCESS_WORKSHEET_RETRIEVE','FAILED',$mixed)))
    Check 'ProcessWorksheet.Wire.RETRIEVE.ConfirmedRejectsMixed' (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('PRODUCTION_PROCESS_WORKSHEET_RETRIEVE','CONFIRMED',$mixed)))
    $recipeCaptions=[ordered]@{PRODUCTION_RECIPE_MOVE_UP='Move Up';PRODUCTION_RECIPE_MOVE_DOWN='Move Down';PRODUCTION_RECIPE_AUTO_ORDER='Auto Order'}
    foreach($id in $recipeCaptions.Keys){
        $definition=([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,17)))|ConvertFrom-Json
        $exact=$null -ne $definition -and @($definition.PSObject.Properties).Count -eq 8 -and $definition.ControlId -ceq $id -and $definition.OwnerId -ceq 'PRODUCTION_DESIGNER' -and $definition.Class -ceq 'Command' -and $definition.Role -ceq 'Production' -and $definition.Caption -ceq $recipeCaptions[$id] -and $definition.Surface -ceq 'Operations > Production > Recipe Designer' -and $definition.Capability -ceq 'PROD_POST' -and $definition.CodePrefix -ceq ($id+'_')
        Check ('ProcessWorksheet.Wire.PreserveRecipeOrder.'+$id) $exact
    }
}

function Test-ProcessWorksheetActivity($Fixture,$Other) {
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
            $severity=switch($Outcome){'REJECTED'{'Warning'} 'DENIED'{'Blocked'} 'FAILED'{'Error'} default{'Info'}};$dataEffect=if($Outcome -in @('CONFIRMED','FAILED')){'Unknown'}else{'Unchanged'}
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
        if(-not $ProcessWorksheetClosedDiagnostic){Test-ProcessWorksheetCatalog}
        if($ProcessWorksheetCatalogOnly){return}
        if($ProcessWorksheetClosedDiagnostic){
            $book=$excel.Workbooks.Add();$closedPath=Join-Path $runRoot 'closed-diagnostic.xlsb';$book.SaveAs($closedPath,50)
            $decoy=$excel.Workbooks.Add()
            [void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'WorksheetActivityStage' @($canary));[void](Probe 'WorksheetActivityAct' @('SEND'))
            $book.Save();[void](Probe 'WorksheetHeadersShow')
            $before=@(Files);$disk=Hash $closedPath
            Write-Output 'Closed diagnostic: visible form prepared; closing captured workbook.'
            $book.Close($false);$book=$null
            Write-Output 'Closed diagnostic: captured workbook closed; invoking guarded adapter.'
            $returned=[bool](Probe 'WorksheetActivityGuard' @('SEND','ClosedWorkbook'))
            $status=([string](Probe 'WorksheetGuardStatusForTest')).Split('|')
            [pscustomobject]@{HandlerEntered=($status[0] -ceq 'True');Returned=$returned;AdapterError=[long]$status[1];WorkbookBytesPreserved=((Hash $closedPath) -ceq $disk);NoActivity=(@(Files).Count -eq $before.Count);DiagnosticOnly=$true}|
                ConvertTo-Json|Set-Content (Join-Path $reportRoot 'closed-boundary-diagnostic.json')
            [void](Probe 'WorksheetSafeCloseForTest')
            return
        }
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
        # Each failure still exercises real submissions/imports. All exact source
        # identities are taken from the owner return, including catch-up runs.
        foreach($fault in @('UncertainSecond','RemoveSecond','SaveSecond')){
            [void](Probe 'CloseDesigner');$book.Close($false);$book=$excel.Workbooks.Add()
            $book.SaveAs((Join-Path $runRoot ($fault+'.xlsb')),50)
            [void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'WorksheetFaultForTest' @(''))
            for($i=0;$i -lt 2;$i++){[void](Probe 'WorksheetActivityStage' @($canary));[void](Probe 'WorksheetActivityAct' @('SEND'))}
            if(-not [bool](Probe 'WorksheetActivitySelect' @($book.Name,1,2,$true))){throw 'Partial retrieval fixture unavailable.'}
            [void](Probe 'WorksheetFaultForTest' @($fault));$before=@(Files);[void](Probe 'WorksheetActivityAct' @('RETRIEVE'))
            Check ('ProcessWorksheet.Partial.'+$fault+'.InjectedOnce') ([long](Probe 'WorksheetFaultHitsForTest') -eq 1)
            $expectedTables=if($fault -ceq 'SaveSecond'){0}else{1}
            Check ('ProcessWorksheet.Partial.'+$fault+'.ActualPartialTableState') (@(Tables).Count -eq $expectedTables)
            if($fault -ceq 'SaveSecond'){Check 'ProcessWorksheet.Partial.SaveSecond.UnsavedDeletionRetained' (-not $book.Saved)}
            $observed=@(([string](Probe 'WorksheetSubmissionFactsForTest')).Split([char]10)|Where-Object{$_})
            $states=@($observed|ForEach-Object{($_ -split '\|')[1]})
            $expectedStates=if($fault -ceq 'UncertainSecond'){'Submitted|Unknown'}else{'Submitted|Submitted'}
            Check ('ProcessWorksheet.Partial.'+$fault+'.ActualSubmissionStates') (($states -join '|') -ceq $expectedStates)
            Pair $before 'RETRIEVE' 'FAILED' ('ProcessWorksheet.Partial.'+$fault) 2
            [void](Probe 'WorksheetFaultForTest' @(''))
        }
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        foreach($guard in @('Loading','Nested','Session','Target','SignedOut','ClosedWorkbook')){
            foreach($action in @('SEND','ADD_ITEM','RETRIEVE')){
                SelectTarget $Fixture 'config-producer'
                $book=$excel.Workbooks.Add();$guardPath=Join-Path $runRoot ('guard-'+$guard+'-'+$action+'.xlsb');$book.SaveAs($guardPath,50)
                [void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'WorksheetActivityStage' @($canary));[void](Probe 'WorksheetActivityAct' @('SEND'))
                if(-not [bool](Probe 'WorksheetActivitySelect' @($book.Name,1,1,$true))){throw 'Guard fixture unavailable.'}
                $book.Save();$table=@(Tables)[0];$tableName=[string]$table.Name;$cells=$table.Range.Formula|ConvertTo-Json -Depth 6 -Compress;$diskPin=Hash $guardPath
                $draft=[string](Probe 'State' @('Process'))
                if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
                if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
                if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                $closedBoundary=$null
                if($guard -ceq 'ClosedWorkbook'){
                    [void](Probe 'WorksheetHeadersShow')
                    [void](Probe 'WorksheetCaptureBindingForTest' @($book.Name))
                    $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
                    $capturedName=$book.Name;$decoyName=$decoy.Name
                    $closedBoundary=[ordered]@{Action=$action;BeforeUTC=[DateTimeOffset]::UtcNow.ToString('o');VisibleBefore=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero);WorkbooksBefore=$excel.Workbooks.Count}
                    $book.Close($false);$book=$null
                    $openNames=@(foreach($openBook in $excel.Workbooks){[string]$openBook.Name})
                    $closedBoundary.AfterUTC=[DateTimeOffset]::UtcNow.ToString('o')
                    $closedBoundary.VisibleAfter=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
                    $closedBoundary.WorkbooksAfter=$excel.Workbooks.Count
                    $closedBoundary.CapturedBookStillOpen=($capturedName -cin $openNames)
                    $closedBoundary.DecoyStillOpen=($decoyName -cin $openNames)
                    $closedBoundary.DiagnosticOnly=$true
                    $closedBoundary|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('closed-sequence-'+$action+'.json'))
                    if(-not $closedBoundary.VisibleBefore -or $closedBoundary.CapturedBookStillOpen -or -not $closedBoundary.DecoyStillOpen){throw 'Closed-workbook operator fixture unavailable; not product RED.'}
                    if(-not $closedBoundary.VisibleAfter){
                        # Workbook shutdown can dismiss the native form. Its
                        # disconnected reference is not an operator callback.
                        $label='ProcessWorksheet.ClosedDismissal.'+$action
                        Check ($label+'.NativeSurfaceDismissed') (-not $closedBoundary.VisibleAfter)
                        Check ($label+'.BindingGuardRejectsClosedBook') (-not [bool](Probe 'WorksheetBindingIsCurrentForTest'))
                        Check ($label+'.NoWorkbookSave') ((Hash $guardPath) -ceq $diskPin)
                        Check ($label+'.NoSubmission') ([string](Probe 'WorksheetSubmissionFactsForTest') -ceq '')
                        Check ($label+'.NoActivityOrRetarget') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
                        $closedBoundary.HandlerInvoked=$false
                        $closedBoundary|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('closed-sequence-'+$action+'.json'))
                        [void](Probe 'WorksheetSafeCloseForTest')
                        continue
                    }
                }
                $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
                $returned=[bool](Probe 'WorksheetActivityGuard' @($action,$guard))
                $entry=([string](Probe 'WorksheetGuardStatusForTest')).Split('|')
                if($null -ne $closedBoundary){
                    $closedBoundary.HandlerInvoked=$true
                    $closedBoundary.HandlerEntered=($entry[0] -ceq 'True');$closedBoundary.Returned=$returned;$closedBoundary.AdapterError=[long]$entry[1]
                    $closedBoundary.VisibleAfterCall=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
                    $closedBoundary|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('closed-sequence-'+$action+'.json'))
                }
                if($entry[0] -cne 'True'){throw 'Guard fixture did not enter the form adapter; not product RED.'}
                $same=$true
                if($null -ne $book){
                    $remaining=@(Tables);$same=$remaining.Count -eq 1
                    if($same){$same=[string]$remaining[0].Name -ceq $tableName -and ($remaining[0].Range.Formula|ConvertTo-Json -Depth 6 -Compress) -ceq $cells}
                }
                $label='ProcessWorksheet.Guard.'+$guard+'.'+$action
                Check ($label+'.NoUnhandledError') $returned
                Check ($label+'.ActualHandlerEntered') ($entry[0] -ceq 'True' -and [long]$entry[1] -eq 0)
                if($guard -cne 'ClosedWorkbook'){Check ($label+'.WorksheetUnchanged') $same}
                Check ($label+'.FormDraftUnchanged') ([string](Probe 'State' @('Process')) -ceq $draft)
                Check ($label+'.NoWorkbookSave') ((Hash $guardPath) -ceq $diskPin)
                Check ($label+'.NoSubmission') ([string](Probe 'WorksheetSubmissionFactsForTest') -ceq '')
                Check ($label+'.NoActivityOrRetarget') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
                if($guard -ceq 'ClosedWorkbook'){[void](Probe 'WorksheetSafeCloseForTest')}else{[void](Probe 'CloseDesigner')};if($null -ne $book){$book.Close($false)};$book=$null
            }
        }
        SelectTarget $Fixture 'config-producer'
        Check 'ProcessWorksheet.Extended.InventoryBusinessStatePreserved' ((InventoryState) -ceq $inventoryBefore)
        # Lose capability without replacing the captured session. Auth bytes stay
        # only in memory for restoration; no credential material is emitted.
        foreach($action in @('SEND','ADD_ITEM','RETRIEVE')){
            SelectTarget $Fixture 'config-producer';$book=$excel.Workbooks.Add()
            $deniedPath=Join-Path $runRoot ('denied-'+$action+'.xlsb');$book.SaveAs($deniedPath,50)
            [void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'WorksheetActivityStage' @($canary));[void](Probe 'WorksheetActivityAct' @('SEND'))
            if(-not [bool](Probe 'WorksheetActivitySelect' @($book.Name,1,1,$true))){throw 'Permission fixture unavailable.'}
            $book.Save();$table=@(Tables)[0];$cells=$table.Range.Formula|ConvertTo-Json -Depth 6 -Compress;$draft=[string](Probe 'State' @('Process'));$diskPin=Hash $deniedPath
            $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
            if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.WorksheetCanProduceForTest')){throw 'Initial capability unavailable.'}
            $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb');$authBytes=[IO.File]::ReadAllBytes($authPath)
            try{
                $auth=$excel.Workbooks.Open($authPath,0,$false)
                try{
                    $caps=Table $auth 'tblCapabilities';$changed=0
                    foreach($row in $caps.ListRows){if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-producer' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'PROD_POST'){$row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$changed++}}
                    if($changed -ne 1){throw 'Unique capability fixture required.'};$auth.Save()
                }finally{$auth.Close($false)}
                $lost=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.WorksheetCanProduceForTest')
                $sameContext=$context -ceq [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
                Check ('ProcessWorksheet.Denied.'+$action+'.ActualLossSameContext') ($lost -and $sameContext)
                if(-not($lost -and $sameContext)){throw 'Permission loss not isolated from context loss.'}
                $before=@(Files);[void](Probe 'WorksheetActivityAct' @($action))
                $remaining=@(Tables);$same=$remaining.Count -eq 1
                if($same){$same=($remaining[0].Range.Formula|ConvertTo-Json -Depth 6 -Compress) -ceq $cells}
                Check ('ProcessWorksheet.Denied.'+$action+'.WorksheetAndDraftUnchanged') ($same -and [string](Probe 'State' @('Process')) -ceq $draft)
                Check ('ProcessWorksheet.Denied.'+$action+'.NoWorkbookSave') ((Hash $deniedPath) -ceq $diskPin)
                Pair $before $action 'DENIED' ('ProcessWorksheet.Denied.'+$action)
            }finally{[IO.File]::WriteAllBytes($authPath,$authBytes);[void](Probe 'CloseDesigner');$book.Close($false);$book=$null;SelectTarget $Fixture 'config-producer'}
        }
        $recordPins=@{};foreach($file in Files){$recordPins[$file]=Hash $file}
        $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
        foreach($mode in @('Off','Older','Unavailable')){
            $blocked=Join-Path (Join-Path $Fixture.Root 'Training\Activity') $Fixture.Warehouse;$held=$blocked+'-worksheet-held';$moved=$false
            foreach($pathCheck in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($pathCheck).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fixture escaped owned runtime.'}}
            if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
            try{
                if($mode -ceq 'Off'){
                    SelectTarget $Fixture
                    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.WorksheetPolicyForTest' @($false))){throw 'Tracking-off fixture unavailable.'}
                }
                if($mode -ceq 'Older'){
                    $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                    try{
                        (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Value2=21.0
                        $controls=Table $cfg 'tblEventTrackingControls'
                        for($i=$controls.ListRows.Count;$i -ge 1;$i--){if($controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -cin $ids){$controls.ListRows.Item($i).Delete()}}
                        $cfg.Save()
                    }finally{$cfg.Close($false)}
                }
                SelectTarget $Fixture 'config-producer'
                if($mode -ceq 'Older'){Check 'ProcessWorksheet.OlderPolicy.ExistingControlReadable' (([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_CLOSE'))).StartsWith('True|True|'))}
                if($mode -ceq 'Unavailable'){
                    if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
                    [IO.File]::WriteAllText($blocked,'Blocked disposable worksheet activity path')
                }
                foreach($action in @('SEND','ADD_ITEM','RETRIEVE')){
                    $book=$excel.Workbooks.Add();$book.SaveAs((Join-Path $runRoot ($mode+'-'+$action+'.xlsb')),50)
                    [void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'WorksheetActivityStage' @($canary));[void](Probe 'WorksheetActivityAct' @('SEND'))
                    if(-not [bool](Probe 'WorksheetActivitySelect' @($book.Name,1,1,$true))){throw 'Optional tracking fixture unavailable.'}
                    $columns=@(Tables)[0].ListColumns.Count;$before=@(Files);[void](Probe 'WorksheetActivityAct' @($action))
                    $remaining=@(Tables);$worked=switch($action){'SEND'{$remaining.Count -eq 2} 'ADD_ITEM'{$remaining.Count -eq 1 -and $remaining[0].ListColumns.Count -eq $columns+2} 'RETRIEVE'{$remaining.Count -eq 0}}
                    Check ('ProcessWorksheet.Tracking.'+$mode+'.'+$action+'.ActionAndSaveContinue') ($worked -and $book.Saved)
                    Check ('ProcessWorksheet.Tracking.'+$mode+'.'+$action+'.NoActivity') (@(Files).Count -eq $before.Count)
                    if($mode -ceq 'Unavailable'){Check ('ProcessWorksheet.Tracking.'+$mode+'.'+$action+'.NoticeVisible') ([bool](Probe 'WorksheetActivityNotice'))}
                    [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
                }
            }finally{
                if($mode -ceq 'Unavailable' -and (Test-Path -LiteralPath $blocked -PathType Leaf)){Remove-Item -LiteralPath $blocked}
                if($moved){Move-Item -LiteralPath $held -Destination $blocked}
                [IO.File]::WriteAllBytes($Fixture.Config,$configBytes);SelectTarget $Fixture 'config-producer'
            }
        }
        # Revoke the session after the first actual queue return. The next
        # submission/removal boundary must refuse further work in this action.
        SelectTarget $Fixture 'config-producer';$book=$excel.Workbooks.Add()
        $yieldPath=Join-Path $runRoot 'yield-context-loss.xlsb';$book.SaveAs($yieldPath,50)
        [void](Probe 'OpenDesigner' @($book.Name))
        for($i=0;$i -lt 2;$i++){[void](Probe 'WorksheetActivityStage' @($canary));[void](Probe 'WorksheetActivityAct' @('SEND'))}
        if(-not [bool](Probe 'WorksheetActivitySelect' @($book.Name,1,2,$true))){throw 'Yield-context fixture unavailable.'}
        $book.Save();$yieldPin=Hash $yieldPath;$before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
        try{
            [void](Probe 'WorksheetFaultForTest' @('SignOutAfterFirst'));[void](Probe 'WorksheetActivityAct' @('RETRIEVE'))
            Check 'ProcessWorksheet.Yield.InjectedOnce' ([long](Probe 'WorksheetFaultHitsForTest') -eq 1)
            Check 'ProcessWorksheet.Yield.SessionLost' ([string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq '')
            Check 'ProcessWorksheet.Yield.StopsBeforeSecondQueueCall' ([long](Probe 'WorksheetQueueReturnsForTest') -eq 1)
            Check 'ProcessWorksheet.Yield.NoTableRemovalOrSave' (@(Tables).Count -eq 2 -and (Hash $yieldPath) -ceq $yieldPin)
            $source=@(([string](Probe 'WorksheetSubmissionFactsForTest')).Split([char]10)|Where-Object{$_})
            Check 'ProcessWorksheet.Yield.OneActualSubmittedReference' ($source.Count -eq 1 -and ($source[0] -split '\|')[1] -ceq 'Submitted')
            $newFiles=@(Files|Where-Object{$_ -cnotin $before});$originalAttempt=$newFiles.Count -eq 1
            if($originalAttempt){$record=Get-Content -LiteralPath $newFiles[0] -Raw|ConvertFrom-Json;$originalAttempt=$record.ControlId -ceq 'PRODUCTION_PROCESS_WORKSHEET_RETRIEVE' -and $record.OutcomeCode -ceq 'REQUESTED' -and $record.WarehouseId -ceq $Fixture.Warehouse}
            Check 'ProcessWorksheet.Yield.OnlyOriginalAttemptNoNewContextResult' $originalAttempt
            Check 'ProcessWorksheet.Yield.NoOtherTargetActivity' (@(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }finally{[void](Probe 'WorksheetFaultForTest' @(''));[void](Probe 'CloseDesigner');$book.Close($false);$book=$null;SelectTarget $Fixture 'config-producer'}
        Check 'ProcessWorksheet.PriorActivityImmutable' (@($recordPins.Keys|Where-Object{(Hash $_) -cne $recordPins[$_]}).Count -eq 0)
        Check 'ProcessWorksheet.Final.InventoryBusinessStatePreserved' ((InventoryState) -ceq $inventoryBefore)
        Check 'ProcessWorksheet.Final.AuthConfigBytesPreserved' (@($pins.Keys|Where-Object{($_ -like '*.Auth.xlsb' -or $_ -like '*.Config.xlsb') -and (Hash $_) -cne $pins[$_]}).Count -eq 0)
    }finally{
        try{[void](Probe 'CloseDesigner')}catch{}
        if($null -ne $decoy){$decoy.Close($false)}
        if($null -ne $book){$book.Close($false)}
    }
}
