# Unsaved adapters invoke actual owning form handlers; never replace a handler.
function Install-ProductionLifecycleProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function LifecyclePrepareForTest(ByVal fixtureName As String) As String
    mBtnProcessNew_Click
    mTxtProcessName.Text = fixtureName
    mTxtProcessOutputId.Text = "A01": mTxtProcessOutputName.Text = fixtureName
    mTxtProcessOutputItemCode.Text = "LIFECYCLE-FIXTURE": mTxtProcessOutputQty.Text = "1"
    RefreshProcessOutputUomCatalog "EA"
    mBtnProcessOutputAdd_Click
    If mLstProcessOutputs.ListCount <> 1 Then Exit Function
    LifecyclePrepareForTest = "READY|" & mTxtProcessId.Text & "|" & mTxtProcessVersion.Text
End Function
Public Function LifecycleActForTest(ByVal designer As String, ByVal action As String) As String
    Dim previousTest As Boolean
    previousTest = mReusableActionTestInProgress
    On Error GoTo Failed
    mReusableActionTestInProgress = True
    If designer = "Process" Then
        Select Case action
            Case "Save": mBtnProcessSave_Click
            Case "Release": mBtnProcessRelease_Click
            Case "Obsolete": mBtnProcessObsolete_Click
            Case Else: Err.Raise 5
        End Select
    ElseIf designer = "Recipe" Then
        Select Case action
            Case "Save": mBtnRecipeSave_Click
            Case "Release": mBtnRecipeRelease_Click
            Case "Obsolete": mBtnRecipeObsolete_Click
            Case Else: Err.Raise 5
        End Select
    Else
        Err.Raise 5
    End If
    LifecycleActForTest = "RETURNED"
Done:
    mReusableActionTestInProgress = previousTest
    Exit Function
Failed:
    LifecycleActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
    Resume Done
End Function
Public Function LifecycleStatusForTest(ByVal designer As String) As String
    Dim json As String, records As Collection, record As Object, report As String
    If designer = "Process" Then
        json = modOperationsPrimitiveBridge.GetProcessVersion(mTxtProcessId.Text, mTxtProcessVersion.Text)
    Else
        json = modOperationsPrimitiveBridge.GetRecipeGraph(mTxtReusableRecipeId.Text, mTxtReusableRecipeVersion.Text)
    End If
    Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(json, report)
    If records Is Nothing Then Exit Function
    For Each record In records
        If modProductionReusableDesigns.ReusableRecordText(record, "RecordType") = UCase$(designer) Then
            LifecycleStatusForTest = modProductionReusableDesigns.ReusableRecordText(record, "Status")
            Exit Function
        End If
    Next record
End Function
Public Function LifecycleCaptionForTest(ByVal designer As String, ByVal action As String) As String
    If designer = "Process" Then
        Select Case action
            Case "Save": LifecycleCaptionForTest = mBtnProcessSave.Caption
            Case "Release": LifecycleCaptionForTest = mBtnProcessRelease.Caption
            Case "Obsolete": LifecycleCaptionForTest = mBtnProcessObsolete.Caption
        End Select
    Else
        Select Case action
            Case "Save": LifecycleCaptionForTest = mBtnRecipeSave.Caption
            Case "Release": LifecycleCaptionForTest = mBtnRecipeRelease.Caption
            Case "Obsolete": LifecycleCaptionForTest = mBtnRecipeObsolete.Caption
        End Select
    End If
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function LifecyclePrepare(ByVal fixtureName As String) As String
    LifecyclePrepare = mForm.LifecyclePrepareForTest(fixtureName)
End Function
Public Function LifecycleAct(ByVal designer As String, ByVal action As String) As String
    LifecycleAct = mForm.LifecycleActForTest(designer, action)
End Function
Public Function LifecycleStatus(ByVal designer As String) As String
    LifecycleStatus = mForm.LifecycleStatusForTest(designer)
End Function
Public Function LifecycleCaption(ByVal designer As String, ByVal action As String) As String
    LifecycleCaption = mForm.LifecycleCaptionForTest(designer, action)
End Function
Public Function LifecycleNotice() As String
    LifecycleNotice = mForm.TestStatusText()
End Function
'@)
    . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycleFailures.ps1')
    Install-ProductionLifecycleFailureProbe
}

function Test-ProductionLifecycle($Fixture,$Other) {
    function Probe([string]$method,[object[]]$Values=@()){
        if($method -cnotmatch '^[A-Za-z]+$'){throw 'Unexpected lifecycle adapter identifier.'}
        # Boundary identifiers only: no arguments, draft text, credentials or paths.
        [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Method=$method}|
            ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'lifecycle-boundaries.jsonl')
        Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$method) $Values
    }
    function SourceRows([bool]$AllowMissing=$false) {
        $wire=[string](Run 'invSys.Designs.Domain.xlam' 'modDesignsBridgeApi.ReadDesignsQueryBridgeResult' @('PUBLICATION_EVENTS',$Fixture.Warehouse,$Fixture.Root))
        $lines=@($wire -split '\r?\n')
        if($AllowMissing -and $lines[0] -match '^EVTSRC1\tDesigns\tUnavailable\t' -and $lines[0].EndsWith("`tMissingSource")){return @()}
        if($lines.Count -lt 2 -or $lines[0] -notmatch '^EVTSRC1\tDesigns\tAvailable\t'){throw 'Owning Designs source unavailable; not behavioral RED.'}
        $headers=$lines[1].Split("`t")
        @(foreach($line in $lines|Select-Object -Skip 2){if($line -ne ''){$fields=$line.Split("`t");if($fields.Count -ne $headers.Count){throw 'Designs source field count mismatch.'};$row=@{};for($i=0;$i -lt $headers.Count;$i++){$row[$headers[$i]]=$fields[$i]};[pscustomobject]$row}})
    }
    $canary='LIFECYCLE'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    try {
        SelectTarget $Fixture
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('DesignsEnabled','TRUE'))){throw 'Explicit Designs setup failed; not behavioral RED.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'lifecycle-operator.xlsb';$book.SaveAs($path,50)
        $decoy=$excel.Workbooks.Add();[void](Probe 'OpenDesigner' @($book.Name))
        $ready=[string](Probe 'LifecyclePrepare' @($canary))
        if($ready -cnotmatch '^READY\|([^|]+)\|([^|]+)$'){throw 'Process form fixture could not be staged; not behavioral RED.'}
        $processId=$Matches[1];$processVersion=$Matches[2]
        # The prerequisite Process is released before constructing its Recipe.
        $actions=@(@('Process','Save','DRAFT'),@('Process','Release','RELEASED'),@('Recipe','Save','DRAFT'),@('Recipe','Release','RELEASED'),@('Recipe','Obsolete','OBSOLETE'),@('Process','Obsolete','OBSOLETE'))
        foreach($action in $actions){
            $designer=$action[0];$verb=$action[1];$id='PRODUCTION_'+$designer.ToUpperInvariant()+'_'+$verb.ToUpperInvariant();$case='ProductionLifecycle.'+$designer+'.'+$verb
            if($designer -ceq 'Recipe' -and $verb -ceq 'Save'){
                if(-not [bool](Probe 'ReleasedRecipe' @($processId,$processVersion,$canary))){throw 'Real released Process prerequisite unavailable; not behavioral RED.'}
            }
            $old=@(Get-Slice4beActivityFiles $Fixture);$sourceBefore=@(SourceRows $true)
            $decoy.Activate();$returned=[string](Probe 'LifecycleAct' @($designer,$verb))
            Check ($case+'.ActualHandlerState') ($returned -ceq 'RETURNED' -and [string](Probe 'LifecycleStatus' @($designer)) -ceq $action[2])
            $source=@(SourceRows|Where-Object {$_.EventID -cnotin @($sourceBefore|ForEach-Object EventID)})
            Check ($case+'.OneOwningDesignsEvent') ($source.Count -eq 1 -and $source[0].EventType -ceq ($designer.ToUpperInvariant()+'_'+$verb.ToUpperInvariant()) -and $source[0].WarehouseId -ceq $Fixture.Warehouse -and $source[0].AppliedAtUTC -ne '')
            $raw=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $old}|ForEach-Object {[IO.File]::ReadAllText($_)})
            $rows=@($raw|ForEach-Object {$_|ConvertFrom-Json});$attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcome=@($rows|Where-Object OutcomeCode -CEQ 'CONFIRMED')
            $pair=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
            Check ($case+'.AttemptAndOutcome') $pair
            $definition=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,13));$metadata=if($definition){$definition|ConvertFrom-Json}else{$null}
            Check ($case+'.RegisteredVisibleControl') ($null -ne $metadata -and $metadata.Caption -ceq [string](Probe 'LifecycleCaption' @($designer,$verb)) -and $metadata.OwnerId -ceq 'PRODUCTION_DESIGN_LIFECYCLE' -and $metadata.Surface -ceq ('Operations > Production > '+$designer+' Designer'))
            $exact=$pair;$linked=$pair;$safe=$pair;$unknown=$pair
            if($pair){
                $refs=@($outcome[0].SourceEventRefs)
                $exact=$source.Count -eq 1 -and $refs.Count -eq 1 -and $refs[0].EventId -ceq $source[0].EventID -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SourceKind -ceq 'Designs' -and $refs[0].SubmissionState -ceq 'Submitted' -and @($attempt[0].SourceEventRefs).Count -eq 0
                $linked=$attempt[0].ActivityId -cne '' -and $attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $attempt[0].RecordId -cne $outcome[0].RecordId
                foreach($row in $rows){$unknown=$unknown -and $row.DataEffect -ceq 'Unknown' -and $row.ControlId -ceq $id -and $row.OwnerId -ceq 'PRODUCTION_DESIGN_LIFECYCLE' -and $row.UserId -ceq 'config-producer' -and $row.CatalogVersion -eq 13}
            }
            foreach($text in $raw){foreach($value in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PayloadJson')){if($text.Contains($value)){$safe=$false}}}
            Check ($case+'.ExactDesignsReference') $exact
            Check ($case+'.DistinctCorrelatedRecords') $linked
            Check ($case+'.NoAppliedInference') $unknown
            Check ($case+'.NoEnteredDataOrSecrets') $safe
            Check ($case+'.UnknownOperatorColumnPreserved') ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        }
        $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(12))).Split("`n")|Where-Object {$_})
        $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(13))).Split("`n")|Where-Object {$_})
        Check 'ProductionLifecycle.Catalog.ThirteenExtendsTwelve' ($old.Count -eq 68 -and $new.Count -eq 74 -and @($new|Select-Object -Unique).Count -eq 74 -and @($old|Where-Object {$_ -cnotin $new}).Count -eq 0)
        foreach($id in $old){Check ('ProductionLifecycle.Catalog.Preserve.'+$id) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,12)) -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,13)))}
        $reference='{"WarehouseId":"CATALOG_TEST","SourceKind":"Designs","EventId":"DESIGNS_TEST_1","SubmissionState":"Submitted"}'
        $lifecycleIds=@(foreach($designer in @('PROCESS','RECIPE')){foreach($verb in @('SAVE','RELEASE','OBSOLETE')){'PRODUCTION_'+$designer+'_'+$verb}})
        foreach($id in $lifecycleIds){
            foreach($code in @('CONFIRMED','PENDING','FAILED')){
                Check ('ProductionLifecycle.References.'+$id+'.'+$code+'.Submitted') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,('['+$reference+']'))))
            }
            $cases=@(
                @('FAILED',('['+$reference.Replace('Submitted','Unknown')+']'),$true,'UncertainWrite'),
                @('CONFIRMED','[]',$false,'Missing'),
                @('PENDING',('['+$reference.Replace('Submitted','Unknown')+']'),$false,'UncertainPending'),
                @('FAILED',('['+$reference+','+$reference+']'),$false,'Duplicate'),
                @('CONFIRMED',('['+$reference.Replace('CATALOG_TEST','OTHER')+']'),$false,'Warehouse'),
                @('CONFIRMED',('['+$reference.Replace('Designs','Inventory')+']'),$false,'Inventory'),
                @('CONFIRMED',('['+$reference.Replace('Submitted','Applied')+']'),$false,'State'),
                @('CONFIRMED',('['+$reference.Replace('"SubmissionState"','"Extra":"invalid","SubmissionState"')+']'),$false,'ExtraField')
            )
            foreach($code in @('REQUESTED','DENIED','REJECTED','CANCELLED')){$cases+=,@($code,('['+$reference+']'),$false,($code+'Prewrite'))}
            foreach($case in $cases){Check ('ProductionLifecycle.References.'+$id+'.'+$case[3]) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$case[0],$case[1])) -eq $case[2])}
        }
        Check 'ProductionLifecycle.References.InventoryControlRejectsDesigns' (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('RECEIVING_CONFIRM_WRITES','CONFIRMED',('['+$reference+']'))))
        foreach($actor in @('config-reader','config-producer')){
            [void](Probe 'CloseDesigner');SelectTarget $Fixture $actor;[void](Probe 'OpenDesigner' @($book.Name))
            $verbs=if($actor -ceq 'config-reader'){@('Save','Release','Obsolete')}else{@('Save')}
            $expected=if($actor -ceq 'config-reader'){'DENIED'}else{'REJECTED'}
            foreach($designer in @('Process','Recipe')){foreach($verb in $verbs){
                [void](Probe 'Stage' @($designer,$canary))
                $draft=[string](Probe 'State' @($designer));$sourcesBefore=@(SourceRows|ForEach-Object EventID);$before=@(Get-Slice4beActivityFiles $Fixture)
                [void](Probe 'LifecycleAct' @($designer,$verb))
                $case='ProductionLifecycle.'+$expected+'.'+$designer+'.'+$verb
                $after=@(SourceRows|ForEach-Object EventID)
                Check ($case+'.NoOwningSubmission') (@(Compare-Object $sourcesBefore $after -CaseSensitive).Count -eq 0)
                Check ($case+'.DraftPreserved') ($draft -ceq [string](Probe 'State' @($designer)))
                $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {[IO.File]::ReadAllText($_)|ConvertFrom-Json})
                $attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcome=@($rows|Where-Object OutcomeCode -CEQ $expected)
                $valid=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
                if($valid){
                    $valid=$attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $outcome[0].DataEffect -ceq 'Unchanged' -and $outcome[0].ControlId -ceq ('PRODUCTION_'+$designer.ToUpperInvariant()+'_'+$verb.ToUpperInvariant()) -and $outcome[0].UserId -ceq $actor -and @($outcome[0].SourceEventRefs).Count -eq 0
                }
                Check ($case+'.ExplicitPrewriteObservation') $valid
            }}
        }
        [void](Probe 'CloseDesigner');SelectTarget $Fixture 'config-producer';[void](Probe 'OpenDesigner' @($book.Name))
        . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycleFailures.ps1')
        Test-ProductionLifecycleFailures $Fixture $book $sheet $canary
        . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycleSafety.ps1')
        Test-ProductionLifecycleClosedWorkbook $Fixture $Other $canary $decoy
        Test-ProductionLifecycleSafety $Fixture $book $sheet $canary
        foreach($change in @('Target','Session')){
            [void](Probe 'CloseDesigner');SelectTarget $Fixture 'config-producer';[void](Probe 'OpenDesigner' @($book.Name))
            if($change -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($change -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($designer in @('Process','Recipe')){foreach($verb in @('Save','Release','Obsolete')){
                [void](Probe 'Stage' @($designer,$canary));$state=[string](Probe 'State' @($designer))
                [void](Probe 'LifecycleAct' @($designer,$verb))
                Check ('ProductionLifecycle.Guard.'+$change+'.'+$designer+'.'+$verb) ($state -ceq [string](Probe 'State' @($designer)) -and [string](Probe 'LifecycleNotice') -match 'Reopen|reopen')
            }}
            Check ('ProductionLifecycle.Guard.'+$change+'.NoRedirectedActivity') (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
    } finally {
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
